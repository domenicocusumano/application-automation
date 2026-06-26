"""
LinkedIn Job Scraper — Senior PM Roles
Scrapes LinkedIn (logged-in session), filters against your Google Sheet,
ranks via Claude, prints top 10.

SETUP (one-time):
  pip install playwright requests gspread google-auth anthropic python-dotenv
  playwright install chromium

USAGE:
  python job_scraper.py          # first run: opens browser for you to log in
  python job_scraper.py          # subsequent runs: reuses saved session

REQUIRED — set in .env file:
  1. ANTHROPIC_API_KEY     — get from https://console.anthropic.com
  2. GOOGLE_SHEET_ID       — the long ID from your sheet's URL
  3. GOOGLE_CREDS_FILE     — path to your Google service account JSON
"""

import warnings
warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", message=".*urllib3.*")

import json
import time
import random
import re
import os
from pathlib import Path
from urllib.parse import quote_plus
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright

load_dotenv()

# ── CONFIG ─────────────────────────────────────────────────────────────────────

GOOGLE_SHEET_ID    = os.getenv("GOOGLE_SHEET_ID",   "YOUR_GOOGLE_SHEET_ID")
GOOGLE_CREDS_FILE  = os.getenv("GOOGLE_CREDS_FILE",  "google_credentials.json")

COL_COMPANY      = "Company"
COL_TITLE        = "Position Title"
COL_URL          = "URL"
COL_LINKEDIN_URL = "Linked In URL"

# Saved browser session (created on first login)
SCRIPT_DIR   = Path(__file__).parent
SESSION_FILE = SCRIPT_DIR / "linkedin_session.json"

# ── SEARCH CONFIG ─────────────────────────────────────────────────────────────
# Single broad search — LinkedIn personalizes results when you're logged in.
# Claude handles all the industry/seniority prioritization.
SEARCH_KEYWORDS = "Product Manager"
SEARCH_LOCATION = "United States"

_DEFAULT_EXCLUDED_TITLES = [
    "program manager", "project manager", "program management",
]

_DEFAULT_EXCLUDED_TITLE_WORDS = [
    "engineer", "engineering", "architect", "developer", "scientist",
    "consultant", "devops", "sre",
]

_DEFAULT_TITLE_KEYWORDS = [
    "product manager", "product management", "product owner",
    "head of product", "vp of product", "vp, product",
    "director of product", "director, product",
    "product lead", "product director", "chief product",
]


def _load_title_filters(cfg: dict) -> tuple:
    """Returns (title_keywords, excluded_titles, excluded_title_words) from config."""
    return (
        [k.lower() for k in cfg.get("title_keywords",        _DEFAULT_TITLE_KEYWORDS)],
        [k.lower() for k in cfg.get("excluded_titles",       _DEFAULT_EXCLUDED_TITLES)],
        [k.lower() for k in cfg.get("excluded_title_words",  _DEFAULT_EXCLUDED_TITLE_WORDS)],
    )

# ── GOOGLE SHEETS ──────────────────────────────────────────────────────────────

def normalize_url(url):
    """
    Normalize a URL for deduplication comparison.
    Strips query params (?...), URL fragments (#...), trailing slashes, and lowercases.
    Returns an empty string if the input is empty or not a real URL.
    """
    if not url:
        return ""
    url = str(url).strip().lower()
    url = re.sub(r'#.*$', '', url)   # strip #fragment (must be before ?-strip)
    url = re.sub(r'\?.*$', '', url)  # strip ?query_params
    url = url.rstrip('/')
    # Treat www and non-www as identical
    url = re.sub(r'^(https?://)www\.', r'\1', url)
    return url


def load_applied_jobs():
    """Returns a set of (company_lower, title_lower) and a set of urls already applied to.
    Loads from both 'Applications' and 'Skips' tabs to avoid re-scoring."""
    print("Loading applied jobs and skips from Google Sheets...")
    try:
        scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
        creds  = Credentials.from_service_account_file(GOOGLE_CREDS_FILE, scopes=scopes)
        gc = gspread.authorize(creds)
        # Sheet ID: config.json takes priority over .env
        sheet_id = GOOGLE_SHEET_ID
        try:
            cfg_raw = json.loads((SCRIPT_DIR / "config.json").read_text())
            url_or_id = cfg_raw.get("google_sheet_url", "").strip()
            if url_or_id:
                m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url_or_id)
                if m:
                    sheet_id = m.group(1)
                elif re.match(r'^[a-zA-Z0-9_-]+$', url_or_id):
                    sheet_id = url_or_id
        except Exception:
            pass
        workbook = gc.open_by_key(sheet_id)

        applied_pairs = set()
        applied_urls  = set()

        # Load both Applications and Skips tabs
        for tab_name in ["Applications", "Skips"]:
            try:
                sheet = workbook.worksheet(tab_name)

                # get_all_records() stops at the first blank row — use get_all_values()
                # instead so we read every row in the sheet regardless of gaps.
                all_values = sheet.get_all_values()
                if not all_values:
                    print(f"   0 entries from {tab_name} tab (empty)")
                    continue

                headers = [h.strip() for h in all_values[0]]
                rows = []
                for row_vals in all_values[1:]:
                    # Pad rows that are shorter than the header
                    while len(row_vals) < len(headers):
                        row_vals.append("")
                    # Skip rows that are entirely blank
                    if not any(v.strip() for v in row_vals):
                        continue
                    rows.append({headers[i]: row_vals[i] for i in range(len(headers))})

                for row in rows:
                    company = str(row.get(COL_COMPANY, "")).strip().lower()
                    title   = str(row.get(COL_TITLE,   "")).strip().lower()
                    if company or title:
                        applied_pairs.add((company, title))
                    # Normalize and store both the apply URL and the LinkedIn URL so
                    # we can match against whichever one the scraper has at check time
                    for col in [COL_URL, COL_LINKEDIN_URL]:
                        val = row.get(col, "")
                        norm = normalize_url(val)
                        if norm and norm.startswith("http"):
                            applied_urls.add(norm)
                        # Store Built-in numeric job ID so slug changes don't defeat dedup
                        m = re.search(r'builtin\.com/job/[^/?#]+/(\d+)', (val or "").lower())
                        if m:
                            applied_urls.add(f"builtin-id:{m.group(1)}")

                print(f"   {len(rows)} entries from {tab_name} tab")
            except Exception as e:
                print(f"   Could not load {tab_name} tab: {e}")

        print(f"   Total: {len(applied_pairs)} job pairs, {len(applied_urls)} URLs loaded")
        return applied_pairs, applied_urls

    except FileNotFoundError:
        print(f"   Could not find {GOOGLE_CREDS_FILE} — skipping dedup filter")
        return set(), set()
    except Exception as e:
        print(f"   Google Sheets error: {e}")
        print( "   Continuing without dedup filter...")
        return set(), set()


def already_applied(job, applied_pairs, applied_urls):
    """Returns True if this job appears in the applied sheet.
    Checks both the apply URL and the LinkedIn URL so we catch a match regardless
    of which URL format the sheet stored or the scraper currently has."""
    for field in ["url", "linkedin_url"]:
        norm = normalize_url(job.get(field, ""))
        if norm and norm in applied_urls:
            return True

    company = job.get("company", "").lower().strip()
    title   = job.get("title",   "").lower().strip()

    if (company, title) in applied_pairs:
        return True

    for (ac, at) in applied_pairs:
        if ac and company and (ac in company or company in ac):
            t_words  = set(title.split())
            at_words = set(at.split())
            if len(t_words & at_words) >= 2:
                return True

    return False


# ── LINKEDIN SCRAPER ───────────────────────────────────────────────────────────

def build_search_url(keywords, location, start=0):
    """Builds a LinkedIn jobs search URL for logged-in users."""
    params = {
        "keywords": keywords,
        "location": location,
        "f_TPR":    "r604800",   # posted in last 7 days
        "f_E":      "4,5,6",     # seniority: Director, Executive, Mid-Senior
        "sortBy":   "DD",        # date descending
        "start":    str(start),  # pagination offset (0, 25, 50, ...)
    }
    base = "https://www.linkedin.com/jobs/search/?"
    return base + "&".join(f"{k}={quote_plus(str(v))}" for k, v in params.items())


def ensure_login(playwright):
    """
    Returns a browser context with a valid LinkedIn session.
    - If a saved session exists, loads it (headless).
    - Otherwise, opens a visible browser for manual login, then saves the session.
    """
    if SESSION_FILE.exists():
        print("Reusing saved LinkedIn session...")
        browser = playwright.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox",
                "--disable-blink-features=AutomationControlled",
                "--disable-infobars",
                "--disable-dev-shm-usage",
                "--window-size=1280,900",
            ]
        )
        context = browser.new_context(
            storage_state=str(SESSION_FILE),
            viewport={"width": 1280, "height": 900},
            locale="en-US",
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            extra_http_headers={"Accept-Language": "en-US,en;q=0.9"},
        )
        return browser, context

    # First run — open visible browser for manual login
    print("\n" + "=" * 60)
    print("  FIRST RUN — LinkedIn login required")
    print("  A browser window will open. Log into LinkedIn.")
    print("  When you see your LinkedIn feed, come back here")
    print("  and press ENTER to continue.")
    print("=" * 60 + "\n")

    browser = playwright.chromium.launch(
        headless=False,
        args=["--no-sandbox", "--disable-blink-features=AutomationControlled"]
    )
    context = browser.new_context(
        viewport={"width": 1280, "height": 900},
        locale="en-US",
    )
    page = context.new_page()
    page.goto("https://www.linkedin.com/login", wait_until="domcontentloaded")

    input("\n>>> Press ENTER after you've logged into LinkedIn... ")

    # Save the session for future runs
    context.storage_state(path=str(SESSION_FILE))
    print(f"   Session saved to {SESSION_FILE}")

    return browser, context


def check_linkedin_session():
    """
    Checks the saved LinkedIn session by inspecting the li_at cookie expiry.
    Avoids making any HTTP requests (which can trigger LinkedIn bot detection
    and invalidate the very session we're trying to preserve).
    Returns (valid: bool, reason: str).
    """
    if not SESSION_FILE.exists():
        return False, "No session file found"
    try:
        data = json.loads(SESSION_FILE.read_text())
        cookies = {c["name"]: c for c in data.get("cookies", [])}
        li_at = cookies.get("li_at")
        if not li_at:
            return False, "li_at cookie missing — please re-login"
        exp = li_at.get("expires", -1)
        if exp != -1 and exp < time.time():
            return False, "li_at cookie expired — please re-login"
        return True, "Session active"
    except Exception as e:
        return False, f"Could not read session file: {e}"


def relogin_linkedin():
    """
    Opens a visible browser, navigates to LinkedIn login, polls until the user
    reaches the feed, then saves the session. Yields status lines for streaming.
    """
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=["--no-sandbox", "--disable-blink-features=AutomationControlled"]
        )
        context = browser.new_context(
            viewport={"width": 1280, "height": 900},
            locale="en-US",
        )
        page = context.new_page()
        page.goto("https://www.linkedin.com/login", wait_until="domcontentloaded")
        yield "Browser opened — please log into LinkedIn in the window that appeared."

        LOGIN_PAGES = ["/login", "/signup", "/uas/", "/authwall", "/checkpoint", "accounts.google.com"]

        deadline = time.time() + 180  # 3 minute timeout
        while time.time() < deadline:
            time.sleep(2)
            try:
                url = page.url
                on_linkedin = "linkedin.com" in url
                on_auth_page = any(x in url for x in LOGIN_PAGES)
                if on_linkedin and not on_auth_page:
                    break
            except Exception:
                pass
            yield "Waiting for login..."
        else:
            yield "ERROR: Timed out waiting for login (3 min). Please try again."
            browser.close()
            return

        context.storage_state(path=str(SESSION_FILE))
        browser.close()
        yield "Login confirmed — session saved."


def dismiss_modals(page):
    """Dismiss any modal dialogs LinkedIn opens automatically (auth walls, Easy Apply prompts, etc.)."""
    page.keyboard.press("Escape")
    time.sleep(0.4)
    for sel in [
        "button[aria-label='Dismiss']",
        "button.artdeco-modal__dismiss",
        "button[data-test-modal-close-btn]",
        "button.modal__dismiss",
        "button[aria-label='Dismiss welcome message']",
    ]:
        try:
            btn = page.query_selector(sel)
            if btn and btn.is_visible():
                btn.click()
                time.sleep(0.3)
        except Exception:
            pass


def extract_apply_url(page, context):
    """
    On a /jobs/view/ page, extracts the external apply URL.
    Primary method: reads the URL from LinkedIn's embedded JSON data in <code> tags.
    Fallback: clicks the Apply button and captures the new tab.
    Returns (apply_url, is_external):
      - External apply: the company career site URL, True
      - Easy Apply / fallback: None, False
    """
    time.sleep(1)

    # ── Method 1: Extract URL from embedded page data (most reliable) ──
    try:
        result = page.evaluate("""
            () => {
                // LinkedIn embeds job data as JSON inside <code> elements
                const codes = document.querySelectorAll('code');
                for (const code of codes) {
                    const text = code.textContent || '';
                    // Look for companyApplyUrl or applyUrl fields
                    for (const key of ['companyApplyUrl', 'applyUrl', 'applyMethod']) {
                        const pattern = new RegExp('"' + key + '"\\\\s*:\\\\s*"([^"]+)"');
                        const match = text.match(pattern);
                        if (match && match[1] && !match[1].includes('linkedin.com')) {
                            return { url: match[1], method: key };
                        }
                    }
                    // Also check for offsite apply indicator
                    if (text.includes('"applyMethod"') && text.includes('"OFF_SITE"')) {
                        // It's external — try to find the URL
                        const urlMatch = text.match(/"companyApplyUrl"\\s*:\\s*"([^"]+)"/);
                        if (urlMatch) return { url: urlMatch[1], method: 'OFF_SITE' };
                    }
                }

                // Also check <script type="application/ld+json"> for apply links
                const ldScripts = document.querySelectorAll('script[type="application/ld+json"]');
                for (const script of ldScripts) {
                    try {
                        const data = JSON.parse(script.textContent);
                        if (data.directApply === false && data.url) return { url: data.url, method: 'ld+json' };
                    } catch(e) {}
                }

                // Check for Easy Apply indicator
                const applyBtn = document.querySelector('.jobs-apply-button, button[aria-label*="Apply"]');
                if (applyBtn) {
                    const text = applyBtn.innerText?.trim().toLowerCase() || '';
                    if (text.includes('easy apply')) return { url: null, method: 'easy_apply' };
                }

                return null;
            }
        """)

        if result:
            if result.get("method") == "easy_apply":
                print("      [apply] Easy Apply — keeping LinkedIn URL")
                return None, False
            url = result.get("url")
            if url and "linkedin.com" not in url:
                # Unescape any JSON-escaped characters
                url = url.replace("\\u002F", "/").replace("\\u003A", ":").replace("\\u0026", "&")
                print(f"      [apply] From page data ({result['method']}): {url[:80]}")
                return url.strip(), True
    except Exception as e:
        print(f"      [apply] Page data extraction error: {e}")

    # ── Method 2: Click the apply button and capture the new tab ──
    apply_el = None
    for selector in [
        "button.jobs-apply-button", ".jobs-apply-button",
        "button[aria-label*='Apply']", ".jobs-s-apply button",
    ]:
        try:
            el = page.query_selector(selector)
            if el and "apply" in el.inner_text().strip().lower():
                apply_el = el
                break
        except Exception:
            continue

    if not apply_el:
        # JS fallback to find any visible Apply button
        try:
            handle = page.evaluate_handle("""
                () => [...document.querySelectorAll('button, a')].find(el =>
                    el.offsetParent !== null &&
                    /^(apply|easy apply)/i.test(el.innerText?.trim() || '')
                ) || null
            """)
            if handle and handle.as_element():
                apply_el = handle.as_element()
        except Exception:
            pass

    if not apply_el:
        print("      [apply] No apply button found — may need re-login")
        return None, False

    btn_text = ""
    try:
        btn_text = apply_el.inner_text().strip().lower()
    except Exception:
        pass

    if "easy apply" in btn_text:
        print("      [apply] Easy Apply — keeping LinkedIn URL")
        return None, False

    # Click and capture the new tab
    try:
        with context.expect_page(timeout=10000) as new_page_info:
            apply_el.click()
        new_page = new_page_info.value
        new_page.wait_for_load_state("load", timeout=15000)
        time.sleep(2)
        external_url = new_page.url
        new_page.close()
        if external_url and "linkedin.com" not in external_url:
            print(f"      [apply] External URL (click): {external_url[:80]}")
            return external_url.strip(), True
    except Exception as e:
        print(f"      [apply] Click capture failed: {e}")

    return None, False


def _count_job_cards(page):
    """Returns the highest card count across all known LinkedIn job card selectors."""
    count_js = """
    () => {
        const sels = [
            'li.scaffold-layout__list-item',
            'li[data-occludable-job-id]',
            '[data-entity-urn*="jobPosting"]',
            'li.jobs-search-results__list-item',
            'div.job-card-container',
            '[data-job-id]',
        ];
        let best = 0;
        for (const s of sels) {
            try { const n = document.querySelectorAll(s).length; if (n > best) best = n; } catch(e) {}
        }
        return best;
    }
    """
    try:
        return page.evaluate(count_js)
    except Exception:
        return 0


def _container_center(page):
    """Returns (x, y) of the job list panel so mouse wheel lands in the right place."""
    pos = page.evaluate("""
    () => {
        const c = document.querySelector('.scaffold-layout__list > div')
               || document.querySelector('.scaffold-layout__list')
               || document.querySelector('.jobs-search-results-list');
        if (c) {
            const r = c.getBoundingClientRect();
            return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2)};
        }
        return {x: Math.round(window.innerWidth * 0.25), y: Math.round(window.innerHeight * 0.5)};
    }
    """)
    return pos.get("x", 400), pos.get("y", 400)


def scroll_job_list(page, min_cards=200, slow=False):
    """
    Scrolls the LinkedIn job list until min_cards are loaded or no more appear.
    slow=True: uses real mouse-wheel events (needed for collection pages whose
               IntersectionObserver doesn't fire on programmatic scrollTop changes).
    slow=False: fast JS-only scroll used for paginated search pages.
    """
    # JS scroll — always done; triggers the panel's scrollTop
    scroll_js = """
    () => {
        const candidates = [
            document.querySelector('.scaffold-layout__list > div'),
            document.querySelector('.scaffold-layout__list'),
            document.querySelector('.jobs-search-results-list'),
            document.querySelector('[class*="scaffold-layout__list"]'),
        ];
        const container = candidates.find(c => c && c.scrollHeight > c.clientHeight);
        if (container) {
            container.scrollTop = container.scrollHeight;
            container.dispatchEvent(new Event('scroll', {bubbles: true}));
            return 'container';
        }
        window.scrollTo(0, document.body.scrollHeight);
        window.dispatchEvent(new Event('scroll'));
        return 'page';
    }
    """

    max_stale  = 8 if slow else 3
    sleep_lo   = 2.5 if slow else 0.8
    sleep_hi   = 4.0 if slow else 1.4
    prev_count = 0
    stale_rounds = 0
    scroll_type  = None

    if slow:
        # Position the mouse over the job list panel once before starting
        try:
            cx, cy = _container_center(page)
            page.mouse.move(cx, cy)
        except Exception:
            pass

    while stale_rounds < max_stale:
        try:
            scroll_type = page.evaluate(scroll_js)
        except Exception:
            try:
                page.wait_for_load_state("load", timeout=10000)
                scroll_type = page.evaluate(scroll_js)
            except Exception:
                break

        if slow:
            # Real mouse-wheel event fires LinkedIn's IntersectionObserver,
            # which programmatic scrollTop alone does not reliably trigger.
            try:
                page.mouse.wheel(0, 3000)
            except Exception:
                pass

        time.sleep(random.uniform(sleep_lo, sleep_hi))

        count = _count_job_cards(page)
        if count >= min_cards:
            break
        if count == prev_count:
            stale_rounds += 1
        else:
            stale_rounds = 0
        prev_count = count

    print(f"      (scrolled via {scroll_type or 'none'}, {prev_count} elements in DOM)")


def collect_cards_on_page(page, min_cards=200, slow=False):
    """Scrolls the job list and collects all job cards using the best-matching selector."""
    scroll_job_list(page, min_cards=min_cards, slow=slow)

    candidate_selectors = [
        # Prefer the :has() variant first — skips occluded wrappers (no inner content)
        "li.scaffold-layout__list-item:has(a[href*='/jobs/view/'])",
        "li.scaffold-layout__list-item",
        "li[data-occludable-job-id]",
        "li.jobs-search-results__list-item",
        "li.ember-view.jobs-search-results__list-item",
        "div.job-card-container",
        "div.job-card-list__entity-lockup",
        "[data-entity-urn*='jobPosting']",
        "div.base-card",
        "li.result-card",
    ]

    best_cards = []
    best_selector = None
    for selector in candidate_selectors:
        try:
            found = page.query_selector_all(selector)
        except Exception:
            continue
        if len(found) > len(best_cards):
            best_cards = found
            best_selector = selector

    if best_selector:
        print(f"      (matched {len(best_cards)} cards via: {best_selector})")

    return best_cards


def extract_jobs_from_page(page, seen_ids, min_cards=200, slow=False, title_filters=None):
    """Collects cards on the current page, extracts job data, and reports filter stats."""
    jobs = []
    cards = collect_cards_on_page(page, min_cards=min_cards, slow=slow)

    stats = {"no_data": 0, "excluded": 0, "not_matched": 0, "dup": 0, "kept": 0}

    for card in cards:
        try:
            result = extract_job_from_card(card, return_reason=True, title_filters=title_filters)
            if isinstance(result, str):
                if result == "no_data":
                    stats["no_data"] += 1
                elif result == "excluded":
                    stats["excluded"] += 1
                elif result == "not_matched":
                    stats["not_matched"] += 1
                continue

            job = result
            job_id = extract_job_id(job["url"])
            if job_id and job_id in seen_ids:
                stats["dup"] += 1
                continue
            if job_id:
                seen_ids.add(job_id)

            job["linkedin_url"] = job["url"]
            jobs.append(job)
            stats["kept"] += 1
        except Exception:
            stats["no_data"] += 1
            continue

    print(f"      Kept {stats['kept']} | No data: {stats['no_data']} | "
          f"Excluded title: {stats['excluded']} | Not matched: {stats['not_matched']} | Dup: {stats['dup']}")

    return jobs


MIN_CANDIDATES = 50  # fallback threshold for Phase 2 search pagination


def _load_config():
    config_path = SCRIPT_DIR / "config.json"
    try:
        return json.loads(config_path.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        return {}


def scrape_linkedin(keywords, location, applied_pairs, applied_urls):
    """
    Scrapes jobs from up to three sources in one browser session:
      0. Top Applicant (LinkedIn Premium) — optional, prioritized when enabled
      1. Recommended jobs (personalized feed) — primary
      2. Keyword search (paginated) — fallback
    Filters (location + already-applied) run after each page so we stop early
    once we have MIN_CANDIDATES ready-to-go jobs.
    """
    config         = _load_config()
    top_applicant  = config.get("top_applicant", False)
    preferred_locs = config.get("preferred_locations", ["remote", "miami"])
    title_filters  = _load_title_filters(config)

    candidates = []  # jobs that passed ALL filters
    seen_ids = set()

    def filter_and_keep(raw_jobs):
        """Applies location + already-applied filters, appends survivors to candidates."""
        for job in raw_jobs:
            if not is_valid_location(job.get("location", ""), preferred_locs):
                continue
            if already_applied(job, applied_pairs, applied_urls):
                continue
            candidates.append(job)

    with sync_playwright() as p:
        browser, context = ensure_login(p)
        page = context.new_page()

        CARD_WAIT_SELECTOR = (
            '[data-entity-urn*="jobPosting"], li[data-occludable-job-id], '
            'li.scaffold-layout__list-item, div.job-card-container, [data-job-id]'
        )

        def render_all_cards(page):
            """
            Scroll the left-panel job list in small increments so LinkedIn's React
            virtualizer renders every card's inner HTML (title, company, link).
            Without this, ~half the cards per page are empty shells because they
            sit below the viewport and React hasn't filled them in yet.
            """
            step_js = """
            (step) => {
                const c = document.querySelector('.scaffold-layout__list > div')
                       || document.querySelector('.scaffold-layout__list')
                       || document.querySelector('.jobs-search-results-list');
                if (!c) { window.scrollBy(0, step); return document.body.scrollHeight; }
                c.scrollTop += step;
                return c.scrollHeight;
            }
            """
            reset_js = """
            () => {
                const c = document.querySelector('.scaffold-layout__list > div')
                       || document.querySelector('.scaffold-layout__list')
                       || document.querySelector('.jobs-search-results-list');
                if (c) c.scrollTop = 0;
            }
            """
            try:
                total = 0
                # Scroll down in 150px steps — small enough that each card enters
                # the viewport individually and gets rendered by React
                for _ in range(30):
                    scroll_height = page.evaluate(step_js, 150)
                    time.sleep(0.12)
                    total += 150
                    if scroll_height and total >= scroll_height:
                        break
                # Scroll back to top so Next-button click shows the right page
                page.evaluate(reset_js)
                time.sleep(0.3)
            except Exception:
                pass

        # Selectors for LinkedIn's "Next page" pagination button on collection pages
        NEXT_BTN_SELECTORS = [
            "button[aria-label='Next']",
            "button[aria-label='View next page']",
            ".artdeco-pagination__button--next",
            "li.artdeco-pagination__indicator--next button",
            "button.artdeco-pagination__button[data-test-pagination-next-btn]",
        ]

        def click_next_page():
            """Click the Next pagination button. Returns True if clicked, False if not found/disabled."""
            for sel in NEXT_BTN_SELECTORS:
                try:
                    btn = page.query_selector(sel)
                    if btn and btn.is_visible() and btn.is_enabled():
                        btn.click()
                        return True
                except Exception:
                    continue
            return False

        def scrape_collection(url, label, max_pages=20):
            """
            Paginate through a LinkedIn collection page (Top Applicant, Recommended)
            by clicking the Next button on each page, exhausting all available pages.
            """
            print(f"\n   Loading {label}...")
            try:
                page.goto(url, wait_until="load", timeout=30000)
                time.sleep(random.uniform(2, 4))
                dismiss_modals(page)
                try:
                    page.wait_for_selector(CARD_WAIT_SELECTOR, timeout=10000)
                except Exception:
                    print(f"   Warning: no cards appeared on {label} — skipping")
                    return

                for pg in range(1, max_pages + 1):
                    print(f"   --- {label} page {pg} ---")
                    render_all_cards(page)
                    jobs = extract_jobs_from_page(page, seen_ids, min_cards=1, title_filters=title_filters)
                    filter_and_keep(jobs)
                    print(f"      {len(jobs)} raw | Candidates so far: {len(candidates)}")

                    if not click_next_page():
                        print(f"   No more pages on {label} (exhausted after {pg} page(s))")
                        break

                    time.sleep(random.uniform(2, 4))
                    dismiss_modals(page)
                    try:
                        page.wait_for_selector(CARD_WAIT_SELECTOR, timeout=10000)
                    except Exception:
                        print(f"   Cards didn't appear after Next — stopping {label}")
                        break

            except Exception as e:
                print(f"   {label} error: {e}")

        # ── Phase 0: Top Applicant jobs (LinkedIn Premium, optional) ──
        # Always exhaust this source fully before moving on.
        if top_applicant:
            print("\n[LinkedIn — Phase 0] Top Applicant feed (LinkedIn Premium)...")
            scrape_collection(
                "https://www.linkedin.com/jobs/collections/top-applicant/",
                "LinkedIn: Top Applicant"
            )

        # ── Phase 1: Recommended jobs — always exhaust before falling back to search ──
        print("\n[LinkedIn — Phase 1] Recommended jobs feed...")
        scrape_collection(
            "https://www.linkedin.com/jobs/collections/recommended/?discover=recommended",
            "LinkedIn: Recommended"
        )

        # ── Phase 2: Search results — paginate until MIN_CANDIDATES reached or pages exhausted ──
        print(f"\n[LinkedIn — Phase 2] Keyword search: '{keywords}' | {location}")
        max_pages = 10

        for page_num in range(max_pages):
            if len(candidates) >= MIN_CANDIDATES:
                print(f"\n   Reached {len(candidates)} candidates — done scraping")
                break

            start_offset = page_num * 25
            url = build_search_url(keywords, location, start=start_offset)

            try:
                print(f"\n   --- LinkedIn: Search page {page_num + 1} (offset {start_offset}) ---")
                page.goto(url, wait_until="domcontentloaded", timeout=30000)
                time.sleep(random.uniform(3, 5))

                # Search uses ?start= pagination — 25 items load on page load, no scroll needed.
                page_jobs = extract_jobs_from_page(page, seen_ids, min_cards=1)
                filter_and_keep(page_jobs)
                print(f"   Candidates so far: {len(candidates)}")

                if not page_jobs:
                    print("   No new jobs on this page — stopping search")
                    break

            except Exception as e:
                print(f"   Error on search page {page_num + 1}: {e}")
                break

            if page_num < max_pages - 1:
                time.sleep(random.uniform(2, 4))

        browser.close()

    print(f"\nTotal filtered candidates: {len(candidates)}")
    return candidates


def fetch_apply_urls(jobs):
    """Navigates to each job's LinkedIn page and extracts the actual apply URL.
    Only called on jobs that survived all filters — keeps it fast."""
    if not jobs:
        return jobs

    print(f"\nFetching apply URLs for {len(jobs)} filtered jobs...")

    with sync_playwright() as p:
        browser, context = ensure_login(p)
        page = context.new_page()

        for i, job in enumerate(jobs):
            linkedin_url = job.get("linkedin_url", job["url"])
            try:
                page.goto(linkedin_url, wait_until="load", timeout=25000)
                time.sleep(random.uniform(2, 3))
                dismiss_modals(page)

                # Check for session redirect before trying to extract URL
                current_url = page.url
                if any(x in current_url for x in ["/login", "/checkpoint", "/authwall"]):
                    print("      [apply] ⚠️  Redirected to login — session may need refresh")
                    job["apply_type"] = "session_expired"
                    continue

                apply_url, is_external = extract_apply_url(page, context)
                if apply_url:
                    job["url"] = apply_url
                    job["apply_type"] = "external"
                else:
                    job["apply_type"] = "easy_apply"

                # Dismiss any leftover modals before next job
                for dismiss_sel in [
                    "button[aria-label='Dismiss']",
                    "button.artdeco-modal__dismiss",
                    "button.artdeco-toast-item__dismiss",
                ]:
                    dismiss_btn = page.query_selector(dismiss_sel)
                    if dismiss_btn:
                        try:
                            dismiss_btn.click()
                            time.sleep(0.3)
                        except Exception:
                            pass

            except Exception as e:
                job["apply_type"] = "unknown"
                print(f"   [{i+1}/{len(jobs)}] Error: {e}")
                continue

            print(f"   [{i+1}/{len(jobs)}] {job['title'][:45]} @ {job['company'][:25]} -> {job['apply_type']} | {job['url'][:70]}")
            time.sleep(random.uniform(0.5, 1.0))

        browser.close()

    return jobs


def extract_job_from_card(card, return_reason=False, title_filters=None):
    """Extracts title, company, location, url from a LinkedIn job card.
    Uses CSS selectors first, then a JS fallback for unknown card layouts.
    If return_reason=True, returns a string reason instead of None on rejection.
    title_filters: tuple of (title_keywords, excluded_titles, excluded_title_words) from config."""
    if title_filters:
        title_keywords, excl_titles, excl_words = title_filters
    else:
        title_keywords = _DEFAULT_TITLE_KEYWORDS
        excl_titles    = _DEFAULT_EXCLUDED_TITLES
        excl_words     = _DEFAULT_EXCLUDED_TITLE_WORDS

    def text(selector):
        el = card.query_selector(selector)
        return el.inner_text().strip() if el else ""

    def attr(selector, attribute):
        el = card.query_selector(selector)
        val = el.get_attribute(attribute) if el else None
        return val.strip() if val else ""

    def reject(reason):
        return reason if return_reason else None

    # ── Try CSS selectors first ──
    title = (text("a.job-card-list__title")
             or text("a.job-card-container__link")
             or text("a.job-card-list__title--link")
             or text("[class*='job-card'] strong")
             or text("a[class*='job-card-list__title']")
             or text(".artdeco-entity-lockup__title a")
             or text(".artdeco-entity-lockup__title")
             or text("h3.base-search-card__title"))

    company = (text("span.job-card-container__primary-description")
               or text("a.job-card-container__company-name")
               or text(".artdeco-entity-lockup__subtitle span")
               or text(".artdeco-entity-lockup__subtitle")
               or text("h4.base-search-card__subtitle")
               or text("a.hidden-nested-link"))

    location = (text("li.job-card-container__metadata-item")
                or text("span.job-card-container__metadata-wrapper")
                or text(".artdeco-entity-lockup__caption span")
                or text(".artdeco-entity-lockup__caption")
                or text("span.job-search-card__location"))

    url = (attr("a.job-card-list__title", "href")
           or attr("a.job-card-container__link", "href")
           or attr("a.job-card-list__title--link", "href")
           or attr("a[href*='/jobs/view/']", "href")
           or attr(".artdeco-entity-lockup__title a", "href")
           or attr("a.base-card__full-link", "href"))

    # ── JS fallback: extract from any card structure ──
    if not title or not url:
        try:
            data = card.evaluate("""
                (el) => {
                    // Find the job link (any <a> pointing to /jobs/view/)
                    const link = el.querySelector('a[href*="/jobs/view/"]');
                    if (!link) return null;

                    const url = link.href || link.getAttribute('href') || '';
                    // Title is usually the link text, or a nested strong/span
                    let title = (link.querySelector('strong') || link.querySelector('span') || link).innerText?.trim() || '';

                    // Get all distinct text lines in the card for company/location
                    const allText = el.innerText || '';
                    const lines = allText.split('\\n').map(l => l.trim()).filter(l => l && l !== title);

                    return { title, url, company: lines[0] || '', location: lines[1] || '' };
                }
            """)
            if data and data.get("url"):
                if not title:
                    title = data["title"]
                if not url:
                    url = data["url"]
                if not company:
                    company = data.get("company", "")
                if not location:
                    location = data.get("location", "")
        except Exception:
            pass

    # ── Normalise URL ──
    if url:
        if url.startswith("/"):
            url = "https://www.linkedin.com" + url
        url = re.sub(r'\?.*$', '', url).rstrip('/')

    if not title or not url:
        return reject("no_data")

    # ── Title filters ──
    title_lower = title.lower()
    if any(exc in title_lower for exc in excl_titles):
        return reject("excluded")

    if any(re.search(r'\b' + re.escape(word) + r'\b', title_lower) for word in excl_words):
        return reject("excluded")

    if title_keywords and not any(term in title_lower for term in title_keywords):
        return reject("not_matched")

    return {
        "title":    title,
        "company":  company,
        "location": location,
        "url":      url,
    }


def extract_job_id(url):
    """Pulls the numeric job ID from a LinkedIn URL."""
    match = re.search(r'/jobs/view/(\d+)', url or "")
    return match.group(1) if match else None


# ── PROGRAMMATIC RANKING ───────────────────────────────────────────────────────

_DEFAULT_SENIORITY_TIERS = [
    "vp", "vice president", "cpo", "head of", "director",
    "principal", "staff", "lead", "senior", "group",
]


def programmatic_rank(jobs, config):
    """
    Scores and sorts jobs by seniority tier (configured in Settings).
    Seniority is a scoring bonus — NOT a hard filter. Jobs without any
    seniority keyword receive a base score of 3.0 and still proceed.
    Returns jobs sorted best-to-worst with 'rank', 'score', and 'reason' fields.
    """
    print("\nRanking jobs by configured seniority tiers...")

    tiers = config.get("seniority_tiers", _DEFAULT_SENIORITY_TIERS)
    n = max(len(tiers), 1)

    def tier_score(title):
        t = title.lower()
        for i, tier in enumerate(tiers):
            if tier.lower() in t:
                return round(3.0 + (n - i) / n * 7.0, 1), tier
        return 3.0, None

    for job in jobs:
        score, matched = tier_score(job.get("title", ""))
        job["score"]  = score
        job["reason"] = f"Seniority: {matched}" if matched else "No seniority tier matched"

    ranked = sorted(jobs, key=lambda j: j["score"], reverse=True)
    for i, job in enumerate(ranked, 1):
        job["rank"] = i

    print(f"  Ranked {len(ranked)} jobs. Top score: {ranked[0]['score'] if ranked else 'N/A'}")
    return ranked


# ── OUTPUT ─────────────────────────────────────────────────────────────────────

def print_results(ranked_jobs):
    """Prints the top 10 to terminal in a clean format."""
    print("\n" + "=" * 65)
    print("  TOP 10 SENIOR PM ROLES FOR YOU")
    print("=" * 65)

    if not ranked_jobs:
        print("  No results to display.")
        return

    for job in ranked_jobs[:10]:
        rank    = job.get("rank",     "?")
        title   = job.get("title",    "Unknown Title")
        company = job.get("company",  "Unknown Company")
        location= job.get("location", "Unknown Location")
        url     = job.get("url",      "No URL")
        score   = job.get("score",    "?")
        reason  = job.get("reason",   "")

        print(f"\n  #{rank}  {title}")
        print(f"       {company}")
        print(f"       {location}")
        print(f"       Score: {score}/10  --  {reason}")
        print(f"       {url}")

    print("\n" + "=" * 65)
    print(f"  Ran at: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 65 + "\n")


# ── LOCATION FILTER ───────────────────────────────────────────────────────────

def is_valid_location(location, preferred_locs=None):
    """Returns True if the job location matches any configured preferred location."""
    if preferred_locs is None:
        preferred_locs = ["remote", "miami"]
    loc = location.lower().strip()
    if not loc:
        return True  # unknown — don't filter out
    if "anywhere" in loc:
        return True
    for pref in preferred_locs:
        if pref.lower() in loc:
            return True
    # "United States" with no specific major city — likely remote/open
    if "united states" in loc and not any(
        city in loc for city in [
            "new york", "san francisco", "los angeles", "chicago", "seattle",
            "boston", "austin", "denver", "dallas", "atlanta", "houston",
            "portland", "phoenix", "philadelphia", "washington", "detroit",
            "minneapolis", "san diego", "san jose", "charlotte", "nashville",
            "raleigh", "salt lake", "pittsburgh", "columbus", "indianapolis",
        ]
    ):
        return True
    return False


# ── MAIN ───────────────────────────────────────────────────────────────────────

def main():
    cfg = {}
    try:
        cfg = json.loads((SCRIPT_DIR / "config.json").read_text())
    except Exception:
        pass

    preferred_locs  = cfg.get("preferred_locations", ["remote", "miami"])
    seniority_tiers = cfg.get("seniority_tiers", _DEFAULT_SENIORITY_TIERS)

    print("[Pipeline Settings — applied to all scrapers]")
    print(f"  Preferred locations : {preferred_locs}")
    print(f"  Seniority tiers     : {seniority_tiers}")
    print(f"  Scoring             : programmatic (no Claude)")
    print(f"  Already-applied     : checked against Applications + Skips tabs\n")
    print("LinkedIn Pipeline — Starting\n")

    linkedin_enabled     = cfg.get("linkedin_enabled", True)
    linkedin_search_term = cfg.get("linkedin_search_term", SEARCH_KEYWORDS).strip() or SEARCH_KEYWORDS

    # 1. Load already-applied jobs from Google Sheets
    applied_pairs, applied_urls = load_applied_jobs()

    # 2. Scrape + filter in one pass — stops once we have 10 candidates
    if not linkedin_enabled:
        print("LinkedIn scraping is disabled in Settings — skipping.")
        candidates = []
    else:
        candidates = scrape_linkedin(
            linkedin_search_term, SEARCH_LOCATION, applied_pairs, applied_urls
        )

    if not candidates:
        print("No matching jobs found.")
        print("If this is your first run, make sure you completed the LinkedIn login.")
        print("If the session expired, delete linkedin_session.json and run again.")
        return

    # 3. Fetch actual apply URLs (only for the ~10 survivors)
    candidates = fetch_apply_urls(candidates)

    # 4. Second dedup pass — now that we have external URLs, check against the sheet again
    #    (The first pass only had LinkedIn URLs; the sheet has external apply URLs)
    before = len(candidates)
    candidates = [j for j in candidates if not already_applied(j, applied_pairs, applied_urls)]
    if before - len(candidates) > 0:
        print(f"Filtered out {before - len(candidates)} already-applied (by apply URL) -> {len(candidates)} remaining")

    if not candidates:
        print("All remaining jobs were already applied to. Try again later.")
        return

    # 5. Rank programmatically by seniority tier
    ranked = programmatic_rank(candidates, cfg)

    # 5.5. Final dedup pass - remove internal duplicates from ranked list
    #      (LinkedIn pagination can show same job multiple times)
    print("\nFinal deduplication check...")
    seen_in_ranked = set()
    deduped_ranked = []
    duplicates_removed = 0

    for job in ranked:
        # Create a unique key from URL (primary) or company+title (fallback)
        url_key = normalize_url(job.get("url", ""))
        title_key = (job.get("company", "").lower(), job.get("title", "").lower())

        # Check if we've seen this job already in the ranked list
        if url_key and url_key in seen_in_ranked:
            duplicates_removed += 1
            print(f"   Removing duplicate: {job.get('title', 'Unknown')} @ {job.get('company', 'Unknown')}")
            continue
        elif title_key in seen_in_ranked:
            duplicates_removed += 1
            print(f"   Removing duplicate: {job.get('title', 'Unknown')} @ {job.get('company', 'Unknown')}")
            continue

        seen_in_ranked.add(url_key)
        seen_in_ranked.add(title_key)
        deduped_ranked.append(job)

    if duplicates_removed > 0:
        print(f"Removed {duplicates_removed} duplicate(s) from ranked list -> {len(deduped_ranked)} remaining")
    else:
        print(f"No duplicates found in ranked list")

    ranked = deduped_ranked

    # 6. Print results
    print_results(ranked)

    # 7. Optionally feed into resume pipeline
    if "--resume" in sys.argv and ranked:
        from resume_pipeline import run_pipeline
        run_pipeline(ranked)

    return ranked


if __name__ == "__main__":
    import sys
    main()
