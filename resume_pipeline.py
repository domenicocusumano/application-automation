"""
Resume Pipeline — Dom Cusumano
Takes top jobs from job_scraper.py output, fetches each job description,
scores fit with Claude, and builds a tailored .docx resume for any role scoring 7+.
Uploads finished resumes to Google Drive automatically.

SETUP:
  pip3 install anthropic gspread google-auth google-api-python-client playwright
  npm install docx   (in the project directory, NOT -g)

USAGE:
  python3 resume_pipeline.py              # standalone with pasted jobs
  python3 job_scraper.py --resume         # integrated: scraper feeds into pipeline
"""

import math
import warnings
warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", message=".*urllib3.*")

import os
import re
import json
import time
import base64
import subprocess
import tempfile
import anthropic
from pathlib import Path
from urllib.parse import urlparse
from dotenv import load_dotenv

from playwright.sync_api import sync_playwright
from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from google.oauth2.credentials import Credentials as OAuthCredentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

load_dotenv()

# ── CONFIG ─────────────────────────────────────────────────────────────────────

ANTHROPIC_API_KEY   = os.getenv("ANTHROPIC_API_KEY",  "YOUR_ANTHROPIC_API_KEY")
GOOGLE_CREDS_FILE   = os.getenv("GOOGLE_CREDS_FILE",   "google_credentials.json")
GDRIVE_FOLDER_ID    = os.getenv("GDRIVE_FOLDER_ID",    "")

# OAuth 2.0 credentials for Drive uploads (personal account — service accounts
# have no storage quota and cannot upload files to personal My Drive).
# oauth_credentials.json  → downloaded from Google Cloud Console (Desktop app type)
# gdrive_token.json       → auto-created on first run, reused on every run after
GDRIVE_OAUTH_CREDS  = os.getenv("GDRIVE_OAUTH_CREDS",  "oauth_credentials.json")
GDRIVE_TOKEN_FILE   = os.getenv("GDRIVE_TOKEN_FILE",   "gdrive_token.json")
GDRIVE_OAUTH_SCOPES = ["https://www.googleapis.com/auth/drive.file"]

SCRIPT_DIR          = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE         = os.path.join(SCRIPT_DIR, "config.json")


def _response_text(response) -> str:
    """
    Return the concatenated text from a Messages API response, skipping any
    non-text blocks (e.g. ThinkingBlock).

    Claude Sonnet 5 and Opus 5 run adaptive thinking by default when
    `thinking` is omitted from the request — unlike Sonnet 4.6/Opus 4.8,
    where thinking stayed off unless explicitly enabled. That moved the
    thinking block to content[0] and pushed the text block to content[1],
    so every `response.content[0].text` in this codebase started raising
    AttributeError on ThinkingBlock. Always locate the text block by type.
    """
    return "".join(
        block.text for block in (response.content or []) if block.type == "text"
    )


# Models that accept `thinking: {"type": "adaptive"}` + `output_config.effort`.
# Haiku 4.5 (this project's default scoring_model) and other pre-4.6 models
# reject the adaptive form outright with a 400 ("adaptive thinking is not
# supported on this model") — they take the older enabled/budget_tokens shape
# or no thinking config at all. Settings lets either scoring_model or
# resume_model be swapped independently, so both call sites must check the
# model actually in use rather than assume the project default.
_ADAPTIVE_THINKING_MODELS = (
    "claude-sonnet-5", "claude-opus-5",
    "claude-fable-5", "claude-mythos-5",
    "claude-opus-4-6", "claude-opus-4-7", "claude-opus-4-8",
    "claude-sonnet-4-6",
)


def _adaptive_thinking_kwargs(model: str) -> dict:
    """kwargs to add adaptive thinking capped at a given effort, or {} if unsupported."""
    if any(model.startswith(prefix) for prefix in _ADAPTIVE_THINKING_MODELS):
        return {"thinking": {"type": "adaptive"}, "output_config": {"effort": "medium"}}
    return {}


def _load_pipeline_config():
    """Reads scoring/filter settings and applicant info from config.json."""
    defaults = {
        "first_name":          "",
        "last_name":           "",
        "score_threshold":     6.5,
        "salary_minimum":      0,
        "preferred_locations": ["remote"],
        "gdrive_folder_id":    "",
        "scoring_model":        "claude-sonnet-5",
        "resume_model":         "claude-sonnet-5",
        "resume_output_format": "docx",
    }
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
            result = {k: cfg.get(k, v) for k, v in defaults.items()}
            # gdrive_folder_id can also come from the top-level env var
            if not result["gdrive_folder_id"]:
                result["gdrive_folder_id"] = GDRIVE_FOLDER_ID
            return result
    except (FileNotFoundError, ValueError):
        return defaults


score_threshold = _load_pipeline_config()["score_threshold"]

# ── DOM'S BACKGROUND ───────────────────────────────────────────────────────────

def _load_prompt_file(filename):
    path = os.path.join(SCRIPT_DIR, filename)
    with open(path, "r", encoding="utf-8") as f:
        return f.read()

BACKGROUND_PROMPT = _load_prompt_file("background_prompt.txt")


# ── GOOGLE SERVICES ───────────────────────────────────────────────────────────

def _get_sheet_id() -> str:
    """Returns the Google Sheet ID from config.json (preferred) or .env."""
    try:
        import json as _json, re as _re
        cfg_path = Path(__file__).parent / "config.json"
        cfg = _json.loads(cfg_path.read_text())
        url_or_id = cfg.get("google_sheet_url", "").strip()
        if url_or_id:
            m = _re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url_or_id)
            if m:
                return m.group(1)
            if _re.match(r'^[a-zA-Z0-9_-]+$', url_or_id):
                return url_or_id
    except Exception:
        pass
    return os.getenv("GOOGLE_SHEET_ID", "1IhPY7ukaZh5CAV2ZILWLwEDnrg_KXJcs-blNIqKuCpU")


SHEET_ID = _get_sheet_id()
_SHEET_GID_CACHE: dict = {}  # tab name -> numeric sheetId, memoized per process


def _get_sheet_gid(service, tab_name: str) -> int:
    """Numeric sheetId for a tab name — needed for batchUpdate (cell formatting),
    which addresses sheets by gid, not name. Cached: this only changes if a tab
    is renamed/recreated mid-run, which doesn't happen in normal use."""
    if tab_name not in _SHEET_GID_CACHE:
        meta = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
        for sheet in meta["sheets"]:
            props = sheet["properties"]
            _SHEET_GID_CACHE[props["title"]] = props["sheetId"]
    return _SHEET_GID_CACHE[tab_name]


def get_drive_service():
    """
    Returns an authenticated Drive service using OAuth 2.0 (personal account).
    Service accounts have no storage quota and cannot upload to personal My Drive.

    First run: opens a browser tab to authorize — takes ~10 seconds.
    All subsequent runs: uses the saved token silently, no browser needed.

    Setup (one-time):
      1. Google Cloud Console → APIs & Services → Credentials
      2. Create Credentials → OAuth 2.0 Client ID → Desktop app
      3. Download JSON → save as oauth_credentials.json in this directory
      4. Run the pipeline — a browser tab opens, click Allow, done.
    """
    token_path  = os.path.join(SCRIPT_DIR, GDRIVE_TOKEN_FILE)
    oauth_path  = os.path.join(SCRIPT_DIR, GDRIVE_OAUTH_CREDS)

    creds = None

    # Load saved token if it exists
    if os.path.exists(token_path):
        creds = OAuthCredentials.from_authorized_user_file(token_path, GDRIVE_OAUTH_SCOPES)

    # Refresh or run the full OAuth flow if needed
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(oauth_path):
                raise FileNotFoundError(
                    f"OAuth credentials file not found: {oauth_path}\n"
                    "See the setup instructions in get_drive_service() above."
                )
            flow  = InstalledAppFlow.from_client_secrets_file(oauth_path, GDRIVE_OAUTH_SCOPES)
            creds = flow.run_local_server(port=0)

        # Save token for next run
        with open(token_path, "w") as f:
            f.write(creds.to_json())

    return build("drive", "v3", credentials=creds)


def get_sheets_service():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds  = ServiceAccountCredentials.from_service_account_file(GOOGLE_CREDS_FILE, scopes=scopes)
    return build("sheets", "v4", credentials=creds)


def normalize_url(url: str) -> str:
    if not url:
        return ""
    url = str(url).strip().lower()
    url = re.sub(r'#.*$', '', url)
    url = re.sub(r'\?.*$', '', url)
    url = url.rstrip('/')
    url = re.sub(r'^(https?://)www\.', r'\1', url)
    return url


def _builtin_job_id(url: str) -> str:
    """Extract numeric ID from a Built-in URL: /job/some-slug/12345 → '12345'. Returns '' if not found."""
    m = re.search(r'builtin\.com/job/[^/?#]+/(\d+)', (url or "").lower())
    return m.group(1) if m else ""


def load_applied_from_sheet(service):
    """
    Reads the Applications + Skips tabs fresh from the Sheet and returns
    (applied_pairs, applied_urls) — used to dedup right before scoring/building,
    independent of whatever snapshot the calling scraper used.
    """
    applied_pairs: set = set()
    applied_urls: set  = set()

    for tab_name in ["Applications", "Skips"]:
        try:
            response = service.spreadsheets().values().get(
                spreadsheetId=SHEET_ID,
                range=f"{tab_name}!A:ZZ",
            ).execute()
            all_values = response.get("values", [])
            if not all_values:
                continue

            headers = [h.strip() for h in all_values[0]]
            for row_vals in all_values[1:]:
                while len(row_vals) < len(headers):
                    row_vals.append("")
                if not any(v.strip() for v in row_vals):
                    continue
                row = {headers[i]: row_vals[i] for i in range(len(headers))}

                company = str(row.get("Company", "")).strip().lower()
                title   = str(row.get("Position Title", "")).strip().lower()
                if company or title:
                    applied_pairs.add((company, title))

                for col in ["URL", "Linked In URL"]:
                    val  = row.get(col, "")
                    norm = normalize_url(val)
                    if norm and norm.startswith("http"):
                        applied_urls.add(norm)
                    bid = _builtin_job_id(val)
                    if bid:
                        applied_urls.add(f"builtin-id:{bid}")
        except Exception as e:
            print(f"      Could not load {tab_name} tab for dedup check: {e}")

    return applied_pairs, applied_urls


def is_already_logged(job: dict, applied_pairs: set, applied_urls: set) -> bool:
    """Returns True if this job's URL or Linked In URL already appears in the sheet."""
    for field in ["url", "linkedin_url"]:
        val  = job.get(field, "")
        norm = normalize_url(val)
        if norm and norm in applied_urls:
            return True
        bid = _builtin_job_id(val)
        if bid and f"builtin-id:{bid}" in applied_urls:
            return True

    company = job.get("company", "").lower().strip()
    title   = job.get("title", "").lower().strip()
    return (company, title) in applied_pairs


def _col_letter(index: int) -> str:
    """0-based column index -> spreadsheet column letter (0 -> A, 25 -> Z, 26 -> AA, ...)."""
    letters = ""
    index += 1
    while index > 0:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def log_job_to_sheet(job, score, assessment, tab_name="Applications", drive_link=None, bold_company=False):
    """
    Logs a job to the specified Google Sheet tab.
    Reads the header row to find correct column indices, so it's robust to column reordering.
    Fills: Company, Position Title, URL, Linked In URL, Date, Location, Claude Score,
           Claude notes, and Resume Treatment (Drive URL, Applications tab only).

    bold_company: bolds the Company cell (column A) of the written row - used as a
    visual flag for rows logged without a job description, so they are easy to spot
    and revisit by eye instead of reading every note column.
    """
    try:
        service = get_sheets_service()

        # Read header row to find column indices
        header_range = f"{tab_name}!1:1"
        header_response = service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range=header_range
        ).execute()
        headers = header_response.get('values', [[]])[0]

        # Map field names to column indices (0-based)
        col_map = {}
        for idx, header in enumerate(headers):
            col_map[header.strip()] = idx

        # Check if we have all required columns
        required = ["Company", "Position Title", "URL", "Linked In URL", "Date", "Location", "Claude Score", "Claude notes"]
        missing = [col for col in required if col not in col_map]
        if missing:
            print(f"      ✗ Missing columns in {tab_name}: {missing}")
            return

        # Format location: "Remote" if remote, "Miami" if Miami-based
        location = job.get("location", "")
        if "remote" in location.lower():
            location_formatted = "Remote"
        elif "miami" in location.lower():
            location_formatted = "Miami"
        else:
            location_formatted = location

        # Format date as MM/DD/YYYY
        from datetime import datetime
        today = datetime.now().strftime("%m/%d/%Y")

        # Build a sparse row with empty strings for all columns, then fill our values
        num_cols = len(headers)
        row = [""] * num_cols

        # Fill only the columns we care about
        row[col_map["Company"]] = job.get("company", "")
        row[col_map["Position Title"]] = job.get("title", "")
        row[col_map["URL"]] = job.get("url", "")
        row[col_map["Linked In URL"]] = job.get("linkedin_url", "")  # May be empty
        row[col_map["Date"]] = today
        row[col_map["Location"]] = location_formatted
        row[col_map["Claude Score"]] = score
        row[col_map["Claude notes"]] = assessment

        # Write the Drive URL to "Resume Treatment" if the column exists and we have a link
        if drive_link and "Resume Treatment" in col_map:
            row[col_map["Resume Treatment"]] = drive_link

        # Write to the next empty row — via update() on an explicit A1 range,
        # never values().append(). append() infers where to place a new row
        # by heuristically detecting a "table" boundary from whatever
        # non-empty cells it finds across the whole queried range; on this
        # sheet (20 columns, many of them sparse — interview rounds,
        # recordings, etc. — filled in by hand well after a row is first
        # logged) that heuristic misjudged the table's start and wrote new
        # rows starting several columns right of A instead of at A. Reading
        # the real bottom of the sheet ourselves and writing to a pinned
        # "A{row}:{last_col}{row}" range removes the ambiguity: the write
        # can only ever land at column A of that exact row.
        #
        # Scan the full width (A:ZZ), not just A:{last_col}, so a stray
        # value sitting further right than our own tracked columns still
        # counts toward "the next empty row" — we must never pick a row
        # number that already has data anywhere in it.
        existing = service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID, range=f"{tab_name}!A:ZZ"
        ).execute().get("values", [])
        next_row = len(existing) + 1  # 1-indexed; row 1 is the header

        last_col = _col_letter(num_cols - 1)
        range_name = f"{tab_name}!A{next_row}:{last_col}{next_row}"
        body = {"values": [row]}
        service.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=range_name,
            valueInputOption="USER_ENTERED",
            body=body
        ).execute()

        if bold_company:
            gid = _get_sheet_gid(service, tab_name)
            service.spreadsheets().batchUpdate(spreadsheetId=SHEET_ID, body={
                "requests": [{
                    "repeatCell": {
                        "range": {
                            "sheetId": gid,
                            "startRowIndex": next_row - 1, "endRowIndex": next_row,
                            "startColumnIndex": col_map["Company"], "endColumnIndex": col_map["Company"] + 1,
                        },
                        "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
                        "fields": "userEnteredFormat.textFormat.bold",
                    }
                }]
            }).execute()

        print(f"      ✓ Logged to {tab_name} tab (row {next_row})")

    except Exception as e:
        import traceback
        print(f"      ✗ Failed to log to {tab_name} tab: {e}")
        print(f"      Error details: {traceback.format_exc()}")


def upload_to_drive(filepath, filename, folder_id=None):
    """Uploads a file to the configured Google Drive folder.
    Uses OAuth 2.0 (personal account) so the file lands in your own Drive quota."""
    try:
        _folder = folder_id or _load_pipeline_config().get("gdrive_folder_id") or GDRIVE_FOLDER_ID
        service  = get_drive_service()
        metadata = {"name": filename, "parents": [_folder] if _folder else []}
        media    = MediaFileUpload(
            filepath,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        file = service.files().create(
            body=metadata,
            media_body=media,
            fields="id,webViewLink",
        ).execute()
        return file.get("webViewLink", "uploaded")
    except Exception as e:
        print(f"      Drive upload failed: {e}")
        return None


# ── JOB DESCRIPTION FETCHER ────────────────────────────────────────────────────

SESSION_FILE = os.path.join(SCRIPT_DIR, "linkedin_session.json")


def _looks_like_signup_gate(text: str) -> bool:
    """True when scraped "description" text is actually an account/application
    signup form rather than a job description. A form like this can easily
    clear the >200-char length check (a country-name <select> alone runs to
    thousands of characters) while containing none of the actual role content.
    Kept intentionally narrow (an exact phrase, not a general heuristic) to
    avoid ever discarding a real, unusually-formatted JD."""
    return "enter your email to continue" in text.lower()


# Exact phrases checked live against 10/10 real closed postings pulled from a
# single Dynamite run (Rippling, Workable, Greenhouse, Lever, Ashby, and a
# company's own career-site redirect) — every one of them hit at least one of
# these. This is a job aggregator's index lagging behind the employer's own
# ATS closing the req, not a scraper bug, and it will keep happening
# periodically regardless of source. Kept as exact phrases (not a loose
# heuristic) to avoid ever discarding a real JD that happens to mention
# hiring status in passing.
_CLOSED_POSTING_PHRASES = (
    "no longer available",
    "no longer open",
    "no longer accepting",
    "job not found",
    "requested was not found",
    "couldn't find anything here",
    "might have closed",
    "page not found",
    "position has been filled",
    "this job has been closed",
    "posting is no longer accepting",
)


def _looks_like_closed_posting(text: str) -> bool:
    t = text.lower()
    return any(phrase in t for phrase in _CLOSED_POSTING_PHRASES)


def _page_body_text(page) -> str:
    """Raw visible body text, independent of fetch_job_description's own
    selector-based extraction. Necessary because a closed-posting page is
    often SHORT enough (a two-line "job not found" notice) that none of
    fetch_job_description's >200-char selector matches ever fire — checked
    live, jd came back completely empty for every Ashby/Lever closed
    posting tested, which would silently skip a check gated on `jd` being
    truthy even though the page loaded fine and says exactly what's wrong."""
    try:
        return page.evaluate("() => document.body.innerText") or ""
    except Exception:
        return ""


def _job_id_from_url(url: str) -> str:
    """Last non-empty path segment — for an ATS job posting this is almost
    always the unique job identifier (UUID, numeric ID, or slug)."""
    try:
        path = urlparse(url).path.rstrip("/")
    except Exception:
        return ""
    segments = [s for s in path.split("/") if s]
    return segments[-1].lower() if segments else ""


def _closed_by_silent_redirect(requested_url: str, final_url: str) -> bool:
    """True when the page navigated away from the specific job to a generic
    listing with no explicit "closed" text at all — observed live on
    Rippling, which 302s a closed job straight to the bare company job list.
    Only trusted when the dropped id is long/opaque enough (>=8 chars) that
    its absence from the final URL isn't just an unrelated formatting
    difference — a real UUID or ATS numeric ID clears this easily; a short
    generic path segment does not."""
    job_id = _job_id_from_url(requested_url)
    if len(job_id) < 8:
        return False
    return job_id not in (final_url or "").lower()


def fetch_job_description(page, url):
    """Fetches the full job description text from a job URL.
    Uses the shared page object (with LinkedIn session loaded).
    Falls back to requests library if Playwright navigation fails (e.g., download trigger)."""
    try:
        # Primary: Try Playwright navigation (works for most sites)
        try:
            page.context.set_default_timeout(20000)
            page.goto(url, wait_until="domcontentloaded", timeout=20000)
            time.sleep(2)
        except Exception as e:
            # Fallback: Try requests library if Playwright fails (download trigger, etc.)
            if "Download" in str(e):
                print(f"      Playwright blocked by download, trying requests fallback...")
                try:
                    import requests
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
                    }
                    response = requests.get(url, headers=headers, timeout=10)
                    print(f"      Requests status: {response.status_code}, Content-Type: {response.headers.get('Content-Type', 'unknown')}")
                    if response.status_code == 200 and "text/html" in response.headers.get("Content-Type", ""):
                        page.set_content(response.text)
                        time.sleep(1)
                        print(f"      ✓ Requests fallback succeeded")
                    else:
                        print(f"      ✗ Requests fallback failed: not HTML or bad status")
                        raise Exception("Requests fallback failed")
                except Exception as fallback_error:
                    print(f"      ✗ Requests fallback error: {fallback_error}")
                    raise
            else:
                raise

        desc = ""

        # LinkedIn job pages: locate the "About the job" section by its heading
        # text rather than CSS classes. LinkedIn's job-view layout uses
        # build-hashed class names (e.g. "_4e88d095") that change between
        # deployments, so the old .jobs-description__content-style selectors
        # below no longer match at all on that page. The "…more" truncation
        # there is CSS-only (line-clamp) — the full text is already present
        # in the DOM, so no click is needed to read it in full.
        if "linkedin.com" in url.lower():
            try:
                desc = (page.evaluate("""
                    () => {
                        const heading = Array.from(document.querySelectorAll('h1,h2,h3,h4,span,div'))
                            .find(el => el.textContent.trim() === 'About the job');
                        if (!heading) return null;
                        const container = heading.parentElement?.parentElement || heading.parentElement;
                        return container ? container.innerText : null;
                    }
                """) or "").strip()
            except Exception:
                desc = ""

        if desc and len(desc) > 200:
            return desc[:8000]

        # Expand "Show more" / "more" / "Read full description" accordions
        # (generic company career sites where truncation actually removes DOM text)
        for more_sel in [
            "button.show-more-less-html__button--more",
            "button[aria-label='Click to see more description']",
            "a.show-more-less-html__button--more",
            "button:has-text('Show more')",
            "a:has-text('more')",
        ]:
            try:
                btn = page.query_selector(more_sel)
                if btn and btn.is_visible():
                    btn.click()
                    time.sleep(1)
                    break
            except Exception:
                pass

        # Extract job description — LinkedIn selectors first, then generic fallbacks
        for selector in [
            # LinkedIn "About the job" section (classic layout)
            ".jobs-description__content",
            ".jobs-box__html-content",
            ".description__text",
            ".show-more-less-html__markup",
            "div.jobs-description",
            # Generic
            "[class*='description']",
            "article",
            "main",
        ]:
            el = page.query_selector(selector)
            if el:
                desc = el.inner_text().strip()
                if len(desc) > 200:
                    break

        return desc[:8000] if desc else ""

    except Exception as e:
        print(f"      Could not fetch job description: {e}")
        return ""


# ── CLAUDE SCORING ─────────────────────────────────────────────────────────────

def score_job(client, job, job_description, salary_minimum=0, preferred_locations=None, model="claude-sonnet-5"):
    """Scores a job against Dom's background. Returns dict with score + assessment +
    hard disqualification flags for salary and location."""

    if preferred_locations is None:
        preferred_locations = ["remote", "miami"]

    # We do NOT ask the model to decide salary_disqualify itself — an LLM
    # judgment call is the wrong tool for a hard yes/no business rule, and in
    # practice it sometimes disqualified jobs that stated no salary at all
    # ("can't assess it, so disqualify" reasoning creeping in despite being
    # told not to). Instead we ask only for the plain fact — the maximum
    # salary figure actually printed in the JD, or null if none is stated —
    # and compute the disqualification ourselves in Python below. That makes
    # "no salary listed" structurally incapable of disqualifying a job,
    # regardless of anything the model writes in its reasoning.
    salary_clause = ""
    if salary_minimum and salary_minimum > 0:
        salary_clause = f"""
SALARY: Report the maximum total compensation figure (top of the stated range)
explicitly printed in the job description, as a plain number in "salary_max_usd".
- If the JD states a salary range or figure, put the TOP of that range/figure there.
- If the JD does not state any salary figure at all, set salary_max_usd to null.
  Do not guess, infer, or estimate a figure — null means "not stated", not "low".
Dom's minimum acceptable total compensation is ${salary_minimum:,.0f}/year, for context only —
do not use this to decide anything; just report what the JD states.
"""

    loc_list = ", ".join(str(l).title() for l in preferred_locations)
    location_clause = f"""
LOCATION DISQUALIFIER: Acceptable work arrangements are: {loc_list}.
Read the full job description carefully for any location requirement hidden in the text:
- If it says the candidate MUST live near / within commuting distance of a specific office or hub
  that is NOT in {loc_list} (e.g., "must be within 50 miles of our NYC office", "required to be
  on-site in San Francisco"), set location_disqualify: true.
- If it says "occasional travel" (roughly once a month or less), or lists office hubs as optional
  / for those who prefer in-person, set location_disqualify: false — that is acceptable.
- If fully remote with no proximity requirement, set location_disqualify: false.
"""

    prompt = f"""---

JOB TO EVALUATE:
Title: {job.get('title')}
Company: {job.get('company')}
Location: {job.get('location')}
URL: {job.get('url')}

JOB DESCRIPTION:
{job_description if job_description else "[Could not fetch — use title/company to infer]"}

---
{salary_clause}{location_clause}
---

Score this role for Dom on a scale of 0–10 based on the scoring criteria above. Use the "Top scoring roles from prior sessions" as calibration — compare this role to those examples.

IMPORTANT: Prioritize role-type alignment (30% weight) over industry fit (20% weight). If the role type is a direct match (e.g., agentic AI role + Dom's agent framework, or protocol PM + EON/L3 experience), score highly even if the industry is outside crypto/fintech. Industry matters less when the actual work is a perfect fit.

Be direct and honest. If it's a poor fit, say so clearly (e.g., "hard pass", "would be doing you a disservice"). If there are hard gaps — specific must-have requirements Dom cannot claim — list them as bullets under "Hard gaps:" in your assessment.

Structure your assessment like this:
- Opening line with frank verdict (e.g., "Strong fit", "Worth applying despite X", "Hard pass")
- If there are specific hard gaps (domain expertise, certifications, industry background Dom doesn't have), list them as bullets under "Hard gaps:"
- 1-2 lines explaining why the gaps matter or why the fit works
- Flag compensation if it's notably low (under $160K)

Return ONLY a JSON object:
{{
  "score": 8.5,
  "assessment": "Your full assessment here with hard gaps as bullets if applicable",
  "role_type": "blockchain|fintech|ai|consumer|compliance|marketplace|other",
  "title_for_file": "Senior_PM_Crypto",
  "salary_max_usd": null,
  "location_disqualify": false,
  "disqualify_reason": ""
}}

No other text. Just the JSON."""

    response = client.messages.create(
        model=model,
        max_tokens=3000,
        system=[{"type": "text", "text": BACKGROUND_PROMPT, "cache_control": {"type": "ephemeral"}}],
        messages=[{"role": "user", "content": prompt}],
        **_adaptive_thinking_kwargs(model),
    )
    usage = response.usage
    if getattr(usage, "cache_read_input_tokens", 0):
        print(f"      [cache] {usage.cache_read_input_tokens:,} tokens read from cache")

    raw = _response_text(response).strip()
    raw = re.sub(r'^```json\s*', '', raw)
    raw = re.sub(r'^```\s*',     '', raw)
    raw = re.sub(r'\s*```$',     '', raw)

    try:
        data = json.loads(raw)
    except Exception:
        return {"score": 0, "assessment": "Could not parse score", "role_type": "other", "title_for_file": "Senior_PM"}

    # The JSON can parse fine while "score" itself is still unusable — the
    # model sometimes decides it can't score (e.g. a too-short/truncated JD)
    # and returns "score": null with an explanation instead of a number,
    # even though it was told to always return a number. That crashed the
    # pipeline downstream at `score >= score_threshold` (None vs float).
    # Treat an unscoreable job as score 0 (skip it, matching the "N/A"/
    # unfetchable-JD path) but keep the model's own explanation instead of
    # discarding it, since it's usually a real diagnostic (JD too short/stub).
    score_raw = data.get("score")
    if isinstance(score_raw, bool) or not isinstance(score_raw, (int, float)) or not math.isfinite(score_raw):
        print(f"      ⚠️  Model returned a non-numeric score ({score_raw!r}) — treating as 0/10.")
        data["score"] = 0
        if not data.get("assessment"):
            data["assessment"] = "Model did not return a numeric score."
    else:
        data["score"] = float(score_raw)

    # Compute salary_disqualify ourselves from the reported figure — never
    # from a model-decided boolean. salary_max_usd must be a real positive
    # number for disqualification to even be possible; null/missing/zero/
    # non-numeric (i.e. "the JD didn't state a salary") always means False.
    # See the comment on salary_clause above for why this moved out of the
    # model's hands.
    salary_max = data.get("salary_max_usd")
    data["salary_disqualify"] = bool(
        salary_minimum and salary_minimum > 0
        and isinstance(salary_max, (int, float)) and not isinstance(salary_max, bool)
        and salary_max > 0
        and salary_max < salary_minimum
    )
    if data["salary_disqualify"]:
        data["disqualify_reason"] = (
            f"Max stated salary ${salary_max:,.0f} is below your ${salary_minimum:,.0f} minimum."
        )

    return data


# ── RESUME BUILDER ─────────────────────────────────────────────────────────────

def build_resume_content(client, job, job_description, score_data, model="claude-sonnet-5"):
    """Asks Claude to produce the full tailored resume content as structured JSON."""

    role_type   = score_data.get("role_type", "other")
    score       = score_data.get("score", 0)
    assessment  = score_data.get("assessment", "")

    resume_path = None
    base_resume_dir = os.path.join(SCRIPT_DIR, "base_resume")
    for ext in (".pdf", ".docx"):
        import glob as _glob
        matches = _glob.glob(os.path.join(base_resume_dir, f"*{ext}"))
        if matches:
            resume_path = matches[0]
            break
    if resume_path is None:
        raise FileNotFoundError(
            "No base resume found in base_resume/. "
            "Upload your resume (.pdf or .docx) via the web UI."
        )

    if resume_path.endswith(".pdf"):
        with open(resume_path, "rb") as f:
            resume_content_block = {
                "type": "document",
                "source": {
                    "type": "base64",
                    "media_type": "application/pdf",
                    "data": base64.standard_b64encode(f.read()).decode("utf-8"),
                },
            }
    else:
        from docx import Document
        doc = Document(resume_path)
        resume_text = "\n".join(p.text for p in doc.paragraphs)
        resume_content_block = {"type": "text", "text": resume_text}

    text_prompt = f"""═══════════════════════════════════════════════════
YOUR TASK: BUILD A TAILORED RESUME
═══════════════════════════════════════════════════
Role:       {job.get('title')}
Company:    {job.get('company')}
Location:   {job.get('location')}
Role type:  {role_type}
Score:      {score}/10
Assessment: {assessment}

JOB DESCRIPTION:
{job_description if job_description else "[Use title/company/role type to tailor]"}

═══════════════════════════════════════════════════
BEFORE WRITING: DO THESE THREE STEPS INTERNALLY
═══════════════════════════════════════════════════
STEP 1 — IDENTIFY THE TOP 3 SIGNALS
  What are the 3 strongest signals in Dom's background for this specific role?
  These are not the most impressive things overall — they are the most
  relevant to what this hiring team is actually testing for.

STEP 2 — IDENTIFY THE HONEST GAPS
  What 1-2 requirements in this JD does Dom genuinely not have?
  Do not paper over them. Do not claim them. Work around them.

STEP 3 — DECIDE THE LEAD BULLET
  The first bullet of the most relevant role block must be the single best
  signal for this hiring team. Not the Vela bullet by default. Not the $200M
  migration by default. The right bullet for THIS role.

═══════════════════════════════════════════════════
OUTPUT FORMAT
═══════════════════════════════════════════════════
Return ONLY a JSON object. No markdown. No extra text. No explanation.

{{
  "name": "YOUR FULL NAME",
  "contact": "City, State  •  Phone  •  Email  •  LinkedIn  •  Work Authorization",
  "linkedin_url": "https://www.linkedin.com/in/yourprofile/",
  "summary": "Strategic Product Leader with 15+ years of experience...",
  "competencies": [
    {{ "category": "Category Name", "skills": "Skill 1, Skill 2, Skill 3..." }}
  ],
  "experience": [
    {{
      "title": "Senior Product Manager",
      "company": "Horizen Labs",
      "location": "Remote",
      "dates": "Jan 2024 – Mar 2026",
      "bullets": [
        "First bullet — most relevant signal for this role...",
        "Second bullet..."
      ]
    }}
  ],
  "entrepreneurship": {{
    "header": "Founder, Product & Technical Lead | Disci.io | Miami, FL  Ongoing",
    "bullets": [
      "Sole technical founder...",
      "Second bullet..."
    ]
  }},
  "education": [
    "MIT Sloan School of Management — Blockchain: Business Innovation & Application (2021)",
    "Queens College, CUNY — B.A., Computer Science"
  ]
}}"""

    messages = [{
        "role": "user",
        "content": [
            {
                "type": "text",
                "text": "This is Dom's existing resume. Use the bullet points and wording from this document as your source of truth. Reorder bullets to lead with the most relevant signal for this role. Reframe where needed for the JD. Do not invent new bullets, do not expand existing bullets beyond what is written here, and do not add claims that are not already present in this document.",
            },
            resume_content_block,
            {
                "type": "text",
                "text": text_prompt,
            },
        ],
    }]

    # An unparseable response has been observed intermittently. The parse
    # error alone ("Expecting value: line 1 column 1 (char 0)") is what
    # json.loads raises for ANY string that doesn't start with valid JSON —
    # not just an empty one — so previously logging only the exception gave
    # no way to tell whether the response was truly empty, had a code-fence
    # variant our old regex missed, or had prose before/after the JSON.
    # Now: log the actual raw text on failure, and parse with raw_decode()
    # starting at the first '{' so leading/trailing prose around a valid
    # JSON object no longer fails the whole parse.
    last_raw = ""
    for attempt in range(3):
        response = client.messages.create(
            model=model,
            max_tokens=12000,
            system=[{"type": "text", "text": BACKGROUND_PROMPT, "cache_control": {"type": "ephemeral"}}],
            messages=messages,
            **_adaptive_thinking_kwargs(model),
        )
        usage = response.usage
        if getattr(usage, "cache_read_input_tokens", 0):
            print(f"      [cache] {usage.cache_read_input_tokens:,} tokens read from cache")

        raw = _response_text(response).strip()
        last_raw = raw
        brace_idx = raw.find("{")

        if brace_idx == -1:
            print(f"   ⚠️  Resume response had no '{{' (attempt {attempt + 1}/3) "
                  f"| stop_reason={response.stop_reason} | len={len(raw)} chars "
                  f"| text={raw[:300]!r}")
            continue

        try:
            data, _ = json.JSONDecoder().raw_decode(raw, brace_idx)
            return data
        except Exception as e:
            print(f"   ⚠️  Could not parse resume JSON (attempt {attempt + 1}/3): {e} "
                  f"| stop_reason={response.stop_reason} | len={len(raw)} chars "
                  f"| text={raw[:300]!r}")

    print(f"   ⚠️  Giving up on resume JSON after 3 attempts. Last raw response (first 1500 chars):\n{last_raw[:1500]!r}")
    return None


def build_docx(resume_data, output_path):
    """Generates a .docx file from resume JSON using docx-js via Node."""

    js_code = r"""
const { Document, Packer, Paragraph, TextRun, AlignmentType,
        LevelFormat, ExternalHyperlink, BorderStyle, HeadingLevel,
        UnderlineType } = require('docx');
const fs = require('fs');

const data = JSON.parse(fs.readFileSync(process.argv[3], 'utf8'));

const COLORS = {
  navy: '1F4E79',    // name, section headings, borders
  black: '000000',   // body text, bullets
  gray: '595959',    // contact line, role metadata, skill lists
  link: '1155CC',    // LinkedIn hyperlink
};
const FONT = 'Arial';

function sectionHeading(text) {
  return new Paragraph({
    children: [new TextRun({ text: text.toUpperCase(), bold: true, size: 22, font: FONT, color: COLORS.navy })],
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: COLORS.navy, space: 1 } },
    spacing: { before: 140, after: 40 }
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: 'bullets', level: 0 },
    children: [new TextRun({ text, size: 20, font: FONT, color: COLORS.black })],
    spacing: { before: 20, after: 20 }
  });
}

// Role header: bold black title  |  gray company  |  gray location  |  gray dates
function roleHeader(title, company, location, dates) {
  const meta = [company, location, dates].filter(s => s).join('  |  ');
  return new Paragraph({
    children: [
      new TextRun({ text: title, bold: true, size: 20, font: FONT, color: COLORS.black }),
      new TextRun({ text: '  |  ' + meta, size: 20, font: FONT, color: COLORS.gray }),
    ],
    spacing: { before: 120, after: 30 }
  });
}

const children = [];

// ── NAME ──
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: data.name, bold: true, size: 32, font: FONT, color: COLORS.navy })],
  spacing: { before: 0, after: 40 }
}));

// ── CONTACT LINE with hyperlinked LinkedIn ──
const contactParts = data.contact.split('LinkedIn');
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [
    new TextRun({ text: contactParts[0], size: 18, font: FONT, color: COLORS.gray }),
    new ExternalHyperlink({
      link: data.linkedin_url,
      children: [new TextRun({ text: 'LinkedIn', size: 18, font: FONT, color: COLORS.link, underline: { type: UnderlineType.SINGLE } })]
    }),
    new TextRun({ text: contactParts[1] || '', size: 18, font: FONT, color: COLORS.gray }),
  ],
  spacing: { before: 0, after: 60 }
}));

// ── SUMMARY ──
children.push(sectionHeading('Professional Summary'));
children.push(new Paragraph({
  children: [new TextRun({ text: data.summary, size: 20, font: FONT, color: COLORS.black })],
  spacing: { before: 60, after: 80 }
}));

// ── CORE COMPETENCIES ──
children.push(sectionHeading('Core Competencies'));
data.competencies.forEach(c => {
  children.push(new Paragraph({
    children: [
      new TextRun({ text: c.category + ': ', bold: true, size: 19, font: FONT, color: COLORS.black }),
      new TextRun({ text: c.skills, size: 19, font: FONT, color: COLORS.gray })
    ],
    spacing: { before: 30, after: 30 }
  }));
});

// ── EXPERIENCE ──
children.push(sectionHeading('Professional Experience'));
data.experience.forEach(role => {
  children.push(roleHeader(role.title, role.company, role.location, role.dates));
  role.bullets.forEach(b => children.push(bullet(b)));
});

// ── ENTREPRENEURSHIP ──
if (data.entrepreneurship) {
  children.push(sectionHeading('Entrepreneurship'));
  // Parse header: expect "Title | Company | Location | Dates" or just a string
  const hdr = data.entrepreneurship.header || '';
  const hdrParts = hdr.split('|').map(s => s.trim());
  if (hdrParts.length >= 4) {
    children.push(roleHeader(hdrParts[0], hdrParts[1], hdrParts[2], hdrParts[3]));
  } else if (hdrParts.length === 3) {
    // 3 parts: "Title | Company | Location Dates"
    children.push(roleHeader(hdrParts[0], hdrParts[1], hdrParts[2], ''));
  } else if (hdrParts.length === 2) {
    children.push(new Paragraph({
      children: [
        new TextRun({ text: hdrParts[0], bold: true, size: 20, font: FONT, color: COLORS.black }),
        new TextRun({ text: '  |  ' + hdrParts[1], size: 20, font: FONT, color: COLORS.gray }),
      ],
      spacing: { before: 80, after: 30 }
    }));
  } else {
    // Fallback: bold title, rest gray
    children.push(new Paragraph({
      children: [new TextRun({ text: hdr, bold: true, size: 20, font: FONT, color: COLORS.black })],
      spacing: { before: 80, after: 30 }
    }));
  }
  data.entrepreneurship.bullets.forEach(b => children.push(bullet(b)));
}

// ── EDUCATION ──
children.push(sectionHeading('Education & Certifications'));
data.education.forEach((e, i) => {
  children.push(new Paragraph({
    children: [new TextRun({ text: e, size: 20, font: FONT, color: COLORS.black })],
    spacing: { before: i === 0 ? 60 : 0, after: i === 0 ? 20 : 0 }
  }));
});

// ── BUILD DOC ──
const doc = new Document({
  numbering: {
    config: [{
      reference: 'bullets',
      levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 480, hanging: 240 } },
          run: { font: FONT, size: 20, color: COLORS.black } } }]
    }]
  },
  styles: {
    default: { document: { run: { font: FONT, size: 20, color: COLORS.black } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 720, right: 900, bottom: 720, left: 900 }
      }
    },
    children
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(process.argv[2], buf);
  console.log('ok');
});
"""

    with tempfile.TemporaryDirectory() as tmp:
        # Write resume data JSON
        data_path = os.path.join(tmp, "resume_data.json")
        with open(data_path, "w") as f:
            json.dump(resume_data, f)

        # Write JS builder
        js_path = os.path.join(tmp, "build.js")
        with open(js_path, "w") as f:
            f.write(js_code)

        # Run Node from the PROJECT directory so require('docx') resolves
        # to the project's node_modules, while reading data from the temp dir
        result = subprocess.run(
            ["node", js_path, output_path, data_path],
            cwd=SCRIPT_DIR,
            env={**os.environ, "NODE_PATH": os.path.join(SCRIPT_DIR, "node_modules")},
            capture_output=True, text=True
        )

        if result.returncode != 0 or "ok" not in result.stdout:
            print(f"      docx build error: {result.stderr[:300]}")
            return False

        return True


# ── MAIN PIPELINE ──────────────────────────────────────────────────────────────

def run_pipeline(jobs, test_scoring_only=False):
    """
    Main pipeline. Accepts a list of job dicts with keys: title, company, location, url.
    Scores each, builds resumes for score >= score_threshold, uploads to Drive.

    Args:
        jobs: List of job dicts
        test_scoring_only: If True, skips resume building and only tests scoring + sheet logging
    """
    # Re-read config at run time so settings changes take effect without restart
    pipeline_cfg       = _load_pipeline_config()
    score_threshold    = pipeline_cfg["score_threshold"]
    salary_minimum     = pipeline_cfg["salary_minimum"]
    preferred_locs     = pipeline_cfg["preferred_locations"]
    gdrive_folder      = pipeline_cfg.get("gdrive_folder_id") or GDRIVE_FOLDER_ID
    scoring_model      = pipeline_cfg.get("scoring_model", "claude-sonnet-5")
    resume_model        = pipeline_cfg.get("resume_model",         "claude-sonnet-5")
    resume_output_fmt   = pipeline_cfg.get("resume_output_format", "docx").lower()

    print(f"\n{'='*65}")
    mode = "TEST MODE - Scoring Only" if test_scoring_only else "Starting"
    print(f"  RESUME PIPELINE -- {mode}")
    print(f"  Processing {len(jobs)} jobs  |  Score threshold: {score_threshold}/10")
    if salary_minimum:
        print(f"  Salary minimum:    ${salary_minimum:,.0f}/yr  (from Settings)")
    print(f"  Location filter:   {preferred_locs}  (from Settings)")
    if test_scoring_only:
        print("  (Resume building DISABLED for testing)")
    print(f"{'='*65}\n")

    client   = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    results  = []
    built    = 0

    sheets_service = get_sheets_service()
    print("Checking Applications + Skips tabs for already-logged jobs...")
    applied_pairs, applied_urls = load_applied_from_sheet(sheets_service)
    print(f"  {len(applied_pairs)} job pairs, {len(applied_urls)} URLs loaded\n")

    # Open one browser session with LinkedIn cookies for all JD fetching
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox",
                "--disable-blink-features=AutomationControlled"
            ]
        )

        # Create context with realistic browser fingerprint to avoid bot detection
        context_options = {
            "viewport": {"width": 1920, "height": 1080},
            "user_agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
        }

        if os.path.exists(SESSION_FILE):
            context = browser.new_context(storage_state=SESSION_FILE, **context_options)
        else:
            context = browser.new_context(**context_options)

        page = context.new_page()

        for i, job in enumerate(jobs, 1):
            title   = job.get("title", "Unknown")
            company = job.get("company", "Unknown")
            url     = job.get("url", "")
            linkedin_url = job.get("linkedin_url", "")

            print(f"[{i}/{len(jobs)}]  {title} @ {company}")

            if is_already_logged(job, applied_pairs, applied_urls):
                print("      Already in Applications/Skips (matched by URL/Linked In URL) — skipping")
                continue

            # 1. Fetch job description
            #    Primary:    apply URL (external apply link always takes precedence;
            #                for Easy Apply jobs this URL is the LinkedIn listing itself)
            #    Fallback 1: listing page URL (LinkedIn "About the job" / BuiltIn page) —
            #                also the effective source for Easy Apply jobs above
            #    Fallback 2: pre-scraped listing text (BuiltIn "The Role") — last resort
            #    No JD:      log to Skips (Company bolded as a visual flag) and skip — don't create an application
            print("      Fetching job description...")
            jd = ""
            closed = False

            if url:
                jd = fetch_job_description(page, url)
                # Checked independently of `jd` — a closed-posting notice is
                # often short enough ("Job not found") that none of
                # fetch_job_description's own >200-char selectors ever match,
                # so jd comes back empty even though the page loaded fine and
                # plainly says what's wrong. A job board's search index can
                # lag behind the employer's own ATS closing the req — checked
                # live, 10/10 candidates from one Dynamite run turned out to
                # be closed postings (Rippling/Workable/Greenhouse/Lever/
                # Ashby), each burning a Claude scoring call and a built
                # resume for nothing. Catch it here, before either is spent.
                if _looks_like_closed_posting(_page_body_text(page)) or _closed_by_silent_redirect(url, page.url):
                    print(f"      Apply URL shows this posting is closed/expired — not scoring, logging to Skips")
                    jd, closed = "", True
                elif jd and _looks_like_signup_gate(jd):
                    # Seen on Dynamite's Apply Now intermediary for jobs without an
                    # external ATS (dynamitejobs.com/apply/<id>): "Enter your email
                    # to continue" plus a full country-name <select> dropdown, which
                    # is >200 chars of real text — passes the length check below but
                    # is not a job description. Treat it as a failed fetch so the
                    # real content (listing page, then the pre-scraped description)
                    # gets a chance instead of Claude scoring an email/country form.
                    print(f"      Apply URL returned a signup-gate page, not a JD ({len(jd)} chars) — discarding")
                    jd = ""
                elif jd:
                    print(f"      Got description from apply URL ({len(jd)} chars)")

            if not closed and not jd and linkedin_url and linkedin_url != url:
                jd = fetch_job_description(page, linkedin_url)
                if _looks_like_closed_posting(_page_body_text(page)) or _closed_by_silent_redirect(linkedin_url, page.url):
                    print(f"      Listing page shows this posting is closed/expired — not scoring, logging to Skips")
                    jd, closed = "", True
                elif jd and _looks_like_signup_gate(jd):
                    print(f"      Listing page returned a signup-gate page, not a JD ({len(jd)} chars) — discarding")
                    jd = ""
                elif jd:
                    print(f"      Got description from listing page ({len(jd)} chars)")

            if closed:
                # Deliberately skips the pre-scraped-description fallback below —
                # that text was captured when the posting still looked open, so
                # using it here would just re-introduce the exact stale data
                # that got us here, and spend the Claude call anyway.
                log_job_to_sheet(
                    job, "N/A",
                    "Position closed / no longer accepting applications — not scored",
                    tab_name="Skips",
                    bold_company=True,
                )
                continue

            if not jd:
                scraped_desc = job.get("description", "")
                if scraped_desc and len(scraped_desc) > 200:
                    jd = scraped_desc
                    print(f"      Using pre-scraped listing description ({len(jd)} chars)")

            if not jd:
                print("      Could not fetch job description by any means — logging to Skips")
                log_job_to_sheet(
                    job, "N/A",
                    "Could not fetch job description — no application created",
                    tab_name="Skips",
                    bold_company=True,
                )
                continue

            # 2. Score + hard disqualification checks
            print("      Scoring with Claude...")
            score_data        = score_job(client, job, jd, salary_minimum, preferred_locs, model=scoring_model)
            score             = score_data.get("score", 0)
            assessment        = score_data.get("assessment", "")
            salary_disqualify = score_data.get("salary_disqualify", False)
            loc_disqualify    = score_data.get("location_disqualify", False)
            disq_reason       = score_data.get("disqualify_reason", "")

            print(f"      Score: {score}/10 -- {assessment}")

            # Hard disqualifiers → Skips tab regardless of score
            if salary_disqualify or loc_disqualify:
                tag = "Salary below minimum" if salary_disqualify else "Location mismatch"
                reason = disq_reason or tag
                print(f"      ✗ DISQUALIFIED ({tag}): {reason}")
                log_job_to_sheet(job, score, f"[DISQUALIFIED — {tag}] {assessment}", tab_name="Skips")
                continue

            result = {**job, "score": score, "assessment": assessment, "resume_built": False, "drive_link": None}

            # 3. Build resume if score meets threshold (unless in test mode)
            if score >= score_threshold:
                # Re-check fresh against the sheet right before committing to a build —
                # closes the race window where a concurrent/overlapping run logged this
                # same job while this one was fetching the JD and scoring with Claude.
                if not test_scoring_only:
                    applied_pairs, applied_urls = load_applied_from_sheet(sheets_service)
                    if is_already_logged(job, applied_pairs, applied_urls):
                        print("      Already logged by another run just now (matched by URL/Linked In URL) — skipping build")
                        results.append({**job, "score": score, "assessment": assessment, "resume_built": False, "drive_link": None})
                        print()
                        continue

                if test_scoring_only:
                    print(f"      Score >= {score_threshold} -- would build resume (skipped in test mode)")
                    log_job_to_sheet(job, score, assessment, tab_name="Applications")
                else:
                    print(f"      Score >= {score_threshold} -- building resume...")

                    resume_data = build_resume_content(client, job, jd, score_data, model=resume_model)
                    drive_link  = None

                    if resume_data:
                        first_name   = pipeline_cfg.get("first_name", "").strip()
                        last_name    = pipeline_cfg.get("last_name", "").strip()
                        name_prefix  = f"{first_name}_{last_name}" if first_name or last_name else "Resume"
                        title_slug   = score_data.get("title_for_file", "Role").replace(" ", "_")
                        company_slug = company.replace(" ", "_")
                        # Save to a resumes/ folder in the project directory
                        resumes_dir = os.path.join(SCRIPT_DIR, "resumes")
                        os.makedirs(resumes_dir, exist_ok=True)

                        docx_filename = f"{name_prefix}_{company_slug}_{title_slug}.docx"
                        docx_path     = os.path.join(resumes_dir, docx_filename)
                        success       = build_docx(resume_data, docx_path)

                        if success and resume_output_fmt == "pdf":
                            pdf_filename = docx_filename.replace(".docx", ".pdf")
                            pdf_path     = os.path.join(resumes_dir, pdf_filename)
                            try:
                                conv = subprocess.run(
                                    ["soffice", "--headless", "--convert-to", "pdf",
                                     "--outdir", resumes_dir, docx_path],
                                    capture_output=True, text=True
                                )
                                pdf_ok = conv.returncode == 0 and os.path.exists(pdf_path)
                                if not pdf_ok:
                                    print(f"      PDF conversion failed, keeping .docx: {conv.stderr[:200]}")
                            except FileNotFoundError:
                                print("      LibreOffice (soffice) not found — keeping .docx. Install with: brew install --cask libreoffice")
                                pdf_ok = False
                            if pdf_ok:
                                os.remove(docx_path)
                                output_path = pdf_path
                                filename    = pdf_filename
                            else:
                                output_path = docx_path
                                filename    = docx_filename
                        else:
                            output_path = docx_path
                            filename    = docx_filename

                        if success:
                            print(f"      Uploading to Google Drive...")
                            drive_link = upload_to_drive(output_path, filename, folder_id=gdrive_folder)
                            result["resume_built"] = True
                            result["drive_link"]   = drive_link
                            result["filename"]     = filename
                            built += 1
                            if drive_link:
                                print(f"      Drive link: {drive_link}")
                        else:
                            print("      Resume build failed")

                    # Log to Applications tab with the Drive link (if the build succeeded)
                    log_job_to_sheet(job, score, assessment, tab_name="Applications", drive_link=drive_link)
            else:
                print(f"      Score below {score_threshold} -- skipping")
                # Log to Skips tab
                print(f"      Attempting to log to Skips tab...")
                log_job_to_sheet(job, score, assessment, tab_name="Skips")

            results.append(result)
            print()
            time.sleep(1)

        browser.close()

    # ── SUMMARY ──
    print(f"\n{'='*65}")
    print(f"  PIPELINE COMPLETE -- {built} resumes built and uploaded")
    print(f"{'='*65}")
    print(f"\n  {'ROLE':<45} {'SCORE':>6}  {'RESUME'}")
    print(f"  {'-'*58}")
    for r in sorted(results, key=lambda x: x['score'] if isinstance(x['score'], (int, float)) else -1, reverse=True):
        status = f"OK {r.get('filename','')}" if r['resume_built'] else "--"
        label  = f"{r['title'][:30]} @ {r['company'][:14]}"
        score_str = f"{r['score']:>5.1f}" if isinstance(r['score'], (int, float)) else f"{r['score']:>5}"
        print(f"  {label:<45} {score_str}  {status}")
    print()


# ── ENTRY POINT ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    # Check for flags
    test_mode = "--test-scoring" in sys.argv
    if test_mode:
        sys.argv.remove("--test-scoring")

    # Check for a JSON file passed as argument (from job_scraper integration)
    if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
        with open(sys.argv[1]) as f:
            jobs = json.load(f)
        print(f"Loaded {len(jobs)} jobs from {sys.argv[1]}")
        run_pipeline(jobs, test_scoring_only=test_mode)
    else:
        print("Usage: python3 resume_pipeline.py <jobs.json> [--test-scoring]")
        print("       or import run_pipeline() from job_scraper.py")
        print()
        print("The jobs.json file should be an array of objects with:")
        print('  title, company, location, url, linkedin_url')
        print()
        print("Options:")
        print("  --test-scoring  Skip resume building, only test scoring and sheet logging")
