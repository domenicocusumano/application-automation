#!/usr/bin/env python3
"""
Built-in.com job scraper.
Mirrors the LinkedIn scraper's candidate-collection style:
  - Filters jobs by title (PM role, excluded words) on the list page
  - Visits detail page only for jobs that pass the title filter
  - Fetches the actual apply URL from the detail page
  - Checks already-applied against Google Sheets before adding
  - Filters by configured preferred locations
  - Stops once MAX_CANDIDATES are collected
  - Scores each candidate by configurable seniority tiers (not a hard filter)
  - Prints a summary sorted by seniority score
"""

import asyncio
import json
import os
import re
import sys
import time
from pathlib import Path
from typing import Optional, List
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse, urljoin

from dotenv import load_dotenv
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

load_dotenv()

SCRIPT_DIR     = Path(__file__).parent
CONFIG_FILE    = SCRIPT_DIR / "config.json"
SESSION_FILE   = SCRIPT_DIR / "builtin_session.json"
MAX_PAGES      = 20
NAV_TIMEOUT    = 30_000
DETAIL_TIMEOUT = 20_000
MAX_CANDIDATES = 10

# Built-in sits behind Cloudflare, which rate-limits by IP on request volume
# (all subresources count, not just the document). Past ~8 detail pages
# fetched back-to-back it starts answering with HTTP 429 and the "Just a
# moment..." interstitial instead of the job page. DETAIL_PAUSE paces the
# detail visits to stay under that ceiling; DETAIL_BACKOFF is the retry
# schedule when we trip it anyway. Measured: the challenge does NOT solve
# itself (still up after 30s of waiting), but a plain re-navigation ~10s
# later comes back 200 with the real page.
DETAIL_PAUSE    = 1.5
DETAIL_BACKOFF  = (10, 20, 40)

# Individual rate-limit hits are already handled by DETAIL_BACKOFF's retry —
# this is the escape hatch for when the whole run is stuck, not just one
# job. If every (or nearly every) page in a row eats at least one backoff,
# that's Cloudflare rate-limiting the session as a whole, not a one-off
# blip — continuing to grind through MAX_PAGES at 10-40s per hit just burns
# time for a handful of extra candidates. Once RATE_LIMIT_PAGE_THRESHOLD
# consecutive pages each see >=1 rate-limit event, stop scraping and hand
# off whatever candidates were already collected to the resume pipeline.
RATE_LIMIT_PAGE_THRESHOLD = 3

GOOGLE_CREDS_FILE = Path(__file__).parent / os.getenv("GOOGLE_CREDS_FILE", "google_credentials.json")

# ── TITLE FILTERS ──────────────────────────────────────────────────────────────

_DEFAULT_EXCLUDED_TITLES: set = set()

_DEFAULT_EXCLUDED_TITLE_WORDS: set = set()

_DEFAULT_TITLE_KEYWORDS: list = []

_DEFAULT_SENIORITY_TIERS = [
    "vp", "vice president", "head of", "director",
    "principal", "staff", "lead", "senior", "group",
]


def is_excluded_title(title: str, excl_titles=None, excl_words=None) -> bool:
    t = title.lower()
    _excl_titles = set(excl_titles) if excl_titles is not None else _DEFAULT_EXCLUDED_TITLES
    _excl_words  = set(excl_words)  if excl_words  is not None else _DEFAULT_EXCLUDED_TITLE_WORDS
    if any(ex in t for ex in _excl_titles):
        return True
    words = set(re.split(r"[\s,/\-]+", t))
    return bool(words & _excl_words)


def is_matching_title(title: str, keywords=None) -> bool:
    """Returns True if the title contains at least one configured role keyword."""
    t = title.lower()
    _keywords = keywords if keywords is not None else _DEFAULT_TITLE_KEYWORDS
    if not _keywords:
        return True  # no keyword filter set — accept all titles
    return any(kw in t for kw in _keywords)


def is_valid_location(location: str, preferred_locs: Optional[List[str]] = None) -> bool:
    if preferred_locs is None:
        preferred_locs = ["remote", "miami"]
    if not location:
        return True  # unknown — don't filter out
    loc = location.lower()
    if "anywhere" in loc:
        return True
    for pref in preferred_locs:
        if pref.lower() in loc:
            return True
    return False


def seniority_score(title: str, tiers: Optional[List[str]] = None) -> tuple:
    """Returns (score: float, matched_tier: str|None)."""
    if tiers is None:
        tiers = _DEFAULT_SENIORITY_TIERS
    n = max(len(tiers), 1)
    t = title.lower()
    for i, tier in enumerate(tiers):
        if tier.lower() in t:
            return round(3.0 + (n - i) / n * 7.0, 1), tier
    return 3.0, None


# ── HELPERS ────────────────────────────────────────────────────────────────────

def log(msg: str):
    print(msg, flush=True)


def load_config() -> dict:
    try:
        return json.loads(CONFIG_FILE.read_text())
    except Exception:
        return {}


def normalize_url(url: str) -> str:
    if not url:
        return ""
    url = str(url).strip().lower()
    url = re.sub(r'#.*$', '', url)
    url = re.sub(r'\?.*$', '', url)
    url = url.rstrip('/')
    # Treat www and non-www as identical
    url = re.sub(r'^(https?://)www\.', r'\1', url)
    return url


def _builtin_job_id(url: str) -> str:
    """Extract numeric ID from a Built-in URL: /job/some-slug/12345 → '12345'. Returns '' if not found."""
    m = re.search(r'builtin\.com/job/[^/?#]+/(\d+)', (url or "").lower())
    return m.group(1) if m else ""


def load_applied_jobs():
    """Loads applied/skipped jobs from Google Sheets. Returns (applied_pairs, applied_urls)."""
    log("Loading applied jobs and skips from Google Sheets...")
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        scopes   = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
        creds    = Credentials.from_service_account_file(str(GOOGLE_CREDS_FILE), scopes=scopes)
        gc       = gspread.authorize(creds)

        # Sheet ID from config.json first, then env
        sheet_id = os.getenv("GOOGLE_SHEET_ID", "YOUR_GOOGLE_SHEET_ID")
        try:
            cfg      = json.loads(CONFIG_FILE.read_text())
            url_or_id = cfg.get("google_sheet_url", "").strip()
            if url_or_id:
                m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url_or_id)
                if m:
                    sheet_id = m.group(1)
                elif re.match(r'^[a-zA-Z0-9_-]+$', url_or_id):
                    sheet_id = url_or_id
        except Exception:
            pass

        workbook = gc.open_by_key(sheet_id)

        applied_pairs: set = set()
        applied_urls: set  = set()

        for tab_name in ["Applications", "Skips"]:
            try:
                sheet      = workbook.worksheet(tab_name)
                all_values = sheet.get_all_values()
                if not all_values:
                    log(f"   0 entries from {tab_name} tab (empty)")
                    continue

                headers = [h.strip() for h in all_values[0]]
                rows    = []
                for row_vals in all_values[1:]:
                    while len(row_vals) < len(headers):
                        row_vals.append("")
                    if not any(v.strip() for v in row_vals):
                        continue
                    rows.append({headers[i]: row_vals[i] for i in range(len(headers))})

                for row in rows:
                    company = str(row.get("Company", "")).strip().lower()
                    title   = str(row.get("Position Title", "")).strip().lower()
                    if company or title:
                        applied_pairs.add((company, title))
                    for col in ["URL", "Linked In URL"]:
                        val = row.get(col, "")
                        norm = normalize_url(val)
                        if norm and norm.startswith("http"):
                            applied_urls.add(norm)
                        # Store Built-in numeric job ID as a canonical fallback key
                        bid = _builtin_job_id(val)
                        if bid:
                            applied_urls.add(f"builtin-id:{bid}")

                log(f"   {len(rows)} entries from {tab_name} tab")
            except Exception as e:
                log(f"   Could not load {tab_name} tab: {e}")

        log(f"   Total: {len(applied_pairs)} job pairs, {len(applied_urls)} URLs loaded")
        return applied_pairs, applied_urls

    except FileNotFoundError:
        log(f"   Could not find {GOOGLE_CREDS_FILE} — skipping dedup filter")
        return set(), set()
    except Exception as e:
        log(f"   Google Sheets error: {e} — continuing without dedup filter")
        return set(), set()


def already_applied(job: dict, applied_pairs: set, applied_urls: set) -> bool:
    """Returns True if this job is already in the applied/skips sheet."""
    for field in ["apply_url", "url"]:
        val = job.get(field, "")
        norm = normalize_url(val)
        if norm and norm in applied_urls:
            return True
        # Fallback: match by Built-in numeric job ID even if slug differs
        bid = _builtin_job_id(val)
        if bid and f"builtin-id:{bid}" in applied_urls:
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


def check_builtin_session():
    """
    Checks the saved Built-in session by inspecting cookies for the builtin.com
    domain. Mirrors job_scraper.check_linkedin_session() — no HTTP request,
    just a local expiry check. Returns (valid: bool, reason: str).
    """
    if not SESSION_FILE.exists():
        return False, "No session file found"
    try:
        data = json.loads(SESSION_FILE.read_text())
        cookies = [c for c in data.get("cookies", []) if "builtin.com" in c.get("domain", "")]
        if not cookies:
            return False, "No builtin.com cookies — please re-login"
        now = time.time()
        if all(c.get("expires", -1) != -1 and c["expires"] < now for c in cookies):
            return False, "Session cookies expired — please re-login"
        return True, "Session active"
    except Exception as e:
        return False, f"Could not read session file: {e}"


# ── DETAIL PAGE ────────────────────────────────────────────────────────────────

def _extract_job_post_init(html: str) -> Optional[dict]:
    """
    Pull the authoritative job payload out of a Built-in detail page.

    Every Built-in job page server-renders a bootstrap call:

        Builtin.jobPostInit({"job":{"id":9099544,
                                    "howToApply":"https://jobs.ashbyhq.com/atob/7e6b...",
                                    "companyName":"AtoB",
                                    "title":"Senior Product Manager",
                                    "isEasyApply":false, ...}, ...})

    `howToApply` is the real external apply destination — the same URL the
    Apply button sends a logged-in user to. It is present in the raw HTML
    (no JS execution, no login required), which is why this is the primary
    source instead of any DOM selector. See the comment in scrape_job_detail
    for why the DOM-based approaches kept failing.

    Returns the inner "job" dict, or None if the payload isn't present.
    """
    marker = re.search(r"Builtin\.jobPostInit\(\s*\{", html)
    if not marker:
        return None

    # Brace-match from the opening "{" so nested objects/strings survive.
    start = html.index("{", marker.start())
    depth, in_str, escaped, end = 0, False, False, -1
    for i in range(start, len(html)):
        c = html[i]
        if in_str:
            if escaped:
                escaped = False
            elif c == "\\":
                escaped = True
            elif c == '"':
                in_str = False
        elif c == '"':
            in_str = True
        elif c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                end = i + 1
                break
    if end == -1:
        return None

    try:
        payload = json.loads(html[start:end])
    except json.JSONDecodeError:
        return None

    job = payload.get("job")
    return job if isinstance(job, dict) else None


# Ad-click wrappers Built-in routes some apply links through. The real
# destination is appended after the tracker's own query separator, e.g.
#   https://ad.doubleclick.net/ddm/clk/628601142;435308584;f?https://www.example.com/job/1
_AD_REDIRECT_HOSTS = (
    "ad.doubleclick.net", "adclick.g.doubleclick.net", "googleads.g.doubleclick.net",
)


def _external_apply_url(raw: str, base: str = "") -> str:
    """
    Normalise an apply link and return it only if it actually leaves Built-in.

    Returns "" for empty/relative-to-Built-in/on-site links.

    Host is compared by *hostname*, never by substring: Built-in appends its
    own attribution params to outbound links, so a perfectly good external URL
    can carry "utm_source=builtin.com" in its query string. The old substring
    check rejected those, which silently dropped the apply URL and left the
    Built-in listing URL in the sheet.
    """
    url = (raw or "").strip()
    if not url:
        return ""
    if url.startswith("//"):
        url = "https:" + url
    elif url.startswith("/") and base:
        url = urljoin(base, url)
    if not url.startswith("http"):
        return ""

    host = (urlparse(url).hostname or "").lower()

    # Unwrap ad-click trackers so the sheet gets the employer's real ATS link
    # rather than a tracker that may expire or be blocked.
    if host in _AD_REDIRECT_HOSTS and "?" in url:
        inner = url.split("?", 1)[1]
        if inner.startswith("http://") or inner.startswith("https://"):
            url = inner
            host = (urlparse(url).hostname or "").lower()

    if host == "builtin.com" or host.endswith(".builtin.com"):
        return ""
    return url


async def _is_rate_limited(page, response) -> bool:
    """
    True when Cloudflare answered with its challenge page instead of the job
    page. Three independent tells, because only the first is present on a
    hard 429 and only the last survives a client-side challenge redirect:
      - 429/403/503 status on the document response
      - "__cf_chl" query param Cloudflare appends when it bounces the request
      - "Just a moment..." <title> on the interstitial itself
    """
    if response is not None and response.status in (403, 429, 503):
        return True
    if "__cf_chl" in page.url:
        return True
    try:
        return (await page.title()).strip().lower().startswith("just a moment")
    except Exception:
        return False


class RateLimitTracker:
    """Counts rate-limit events across the whole run so the main page loop
    can tell an isolated blip (one job, quickly recovered) apart from
    Cloudflare rate-limiting the session as a whole (see
    RATE_LIMIT_PAGE_THRESHOLD)."""
    def __init__(self):
        self.events = 0

    def bump(self):
        self.events += 1


async def _goto_detail(page, job_url: str, tracker: "RateLimitTracker") -> bool:
    """
    Navigate to a job detail page, retrying through Cloudflare rate limits.

    Returns True once the real page is loaded, False if every attempt came
    back challenged. Callers must treat False as "no data" — scraping the
    interstitial yields a blank company and no apply URL, which is exactly
    how blank Company cells and Built-in listing URLs ended up in the sheet.
    """
    response = await page.goto(job_url, wait_until="domcontentloaded", timeout=DETAIL_TIMEOUT)
    if not await _is_rate_limited(page, response):
        return True

    for delay in DETAIL_BACKOFF:
        tracker.bump()
        log(f"    [rate-limited] Cloudflare challenge — backing off {delay}s and retrying")
        await asyncio.sleep(delay)
        response = await page.goto(job_url, wait_until="domcontentloaded", timeout=DETAIL_TIMEOUT)
        if not await _is_rate_limited(page, response):
            return True
    return False


async def scrape_job_detail(page, job_url: str, tracker: "RateLimitTracker") -> dict:
    """
    Fetch company, location, job description, Easy Apply status, and actual apply URL
    from a Built-in job detail page.
    """
    result = {
        "company": "", "location": "", "description": "",
        "easy_apply": False, "apply_url": job_url, "login_required": False,
        "blocked": False,
    }
    # True once the apply URL / Easy Apply flag come from the embedded
    # jobPostInit payload, which outranks every DOM heuristic below.
    trusted_apply = False
    try:
        if not await _goto_detail(page, job_url, tracker):
            result["blocked"] = True
            log(f"    [blocked] Cloudflare kept challenging {job_url} — skipping this job")
            return result

        # ── Authoritative source: the embedded jobPostInit payload ──
        # This runs BEFORE any DOM scraping and, when present, settles
        # company / easy_apply / apply_url outright.
        #
        # Why this exists (the bug that kept coming back): the scraper
        # browses Built-in logged OUT. For a logged-out visitor Built-in
        # never renders the external apply anchor at all — the sticky bar
        # renders "Log In to apply" / "Sign up" links instead. So
        # `a#applyButton`, `a[aria-label='Apply to job']`, `a:has-text('Apply')`
        # and friends match nothing (or match the login link), every fallback
        # falls through, and apply_url silently keeps its `job_url` default —
        # which is exactly how the Built-in listing URL ended up in the sheet.
        # No selector tweak can fix that, because the element being hunted for
        # does not exist in the page the scraper is looking at. The URL is
        # only ever in this JSON payload, which is server-rendered for
        # everyone. Verified across Ashby/Greenhouse/Workday/Oracle/Dover
        # postings, Easy Apply and not.
        #
        # Newer postings drop `howToApply` from the payload entirely and only
        # supply `applyUrl`, a same-site "/job/...?handler=ApplyRedirect" path.
        # Logged out, that handler 302s to "?applyRequired=true" (a login
        # wall) instead of the company site — see the applyUrl handling below.
        try:
            job_data = _extract_job_post_init(await page.content())
        except Exception:
            job_data = None

        if job_data:
            company_name = (job_data.get("companyName") or "").strip()
            if company_name:
                result["company"] = company_name

            result["easy_apply"] = bool(job_data.get("isEasyApply"))

            external = _external_apply_url(job_data.get("howToApply") or "")

            # No howToApply — try the applyUrl redirect handler instead. Follow
            # it through the browser context's request API (shares the loaded
            # builtin_session.json cookies, if any) rather than navigating the
            # page, so the JD/company scraping below is unaffected either way.
            if not external and not result["easy_apply"]:
                apply_path = (job_data.get("applyUrl") or "").strip()
                if apply_path:
                    redirect_url = urljoin(job_url, apply_path)
                    # This hit counts against the same Cloudflare rate limit as
                    # the page load, and a challenged response carries no
                    # Location header at all — indistinguishable from "no apply
                    # URL" unless we check the status. Retry on the same backoff
                    # schedule as the detail navigation.
                    for attempt, delay in enumerate((0,) + DETAIL_BACKOFF):
                        if delay:
                            tracker.bump()
                            log(f"    [rate-limited] apply redirect challenged — backing off {delay}s and retrying")
                            await asyncio.sleep(delay)
                        try:
                            resp = await page.context.request.get(
                                redirect_url, max_redirects=0, timeout=10_000
                            )
                        except Exception:
                            break
                        if resp.status in (403, 429, 503):
                            continue

                        location = resp.headers.get("location", "")
                        resolved = _external_apply_url(location, job_url)
                        if resolved:
                            external = resolved
                        elif "applyrequired" in location.lower():
                            # Built-in only reveals this job's real apply URL to a
                            # logged-in account — without a valid session there's
                            # nothing more to try (the DOM fallback below won't
                            # find anything either; see comment above).
                            result["login_required"] = True
                        break

            # Easy Apply jobs deliberately keep the Built-in URL so the
            # application filler drives Built-in's own apply flow.
            if external and not result["easy_apply"]:
                result["apply_url"] = external
            trusted_apply = True

        # BuiltIn renders the company name and apply button via Alpine.js
        # (x-if templates) that clone content into the DOM only after the JS
        # bundle hydrates — they don't exist yet at "domcontentloaded". A
        # flat 1s sleep isn't reliably long enough under real network/CPU
        # load, which was silently producing blank company + apply_url
        # falling back to the BuiltIn listing URL. Wait for either apply
        # flow's markup (external redirect button or native Easy Apply
        # button) to actually appear before scraping starts.
        try:
            await page.wait_for_selector(
                "a#applyButton, [aria-label='Easy Apply to job']", timeout=8000
            )
        except PWTimeout:
            pass
        await page.wait_for_timeout(500)

        # Company name: h2 in the job card header is the most reliable on BuiltIn,
        # followed by breadcrumb first-link, then various data-testid/class selectors.
        # Skipped entirely when jobPostInit already supplied the name.
        for sel in [] if result["company"] else [
            # BuiltIn renders company name as <a href="/company/..."><h2>Name</h2></a>
            # — the h2 is a direct child of the company anchor, NOT the other way around
            "a[href*='/company/'] > h2",
            "a[href*='/companies/'] > h2",
            # Breadcrumb (first link = company name)
            "nav[aria-label*='breadcrumb'] a:first-child",
            "[data-testid='breadcrumb'] a:first-child",
            "[class*='breadcrumb'] a:first-child",
            # Explicit data-testid
            "[data-testid='employer-name']", "[data-testid='company-title']",
            "[data-testid='company-name']",
            # Class-based
            ".company-title", ".employer-name",
            "a[href*='/companies/']", "a[href*='/company/']",
            "[class*='CompanyName']", "[class*='company-name']",
        ]:
            try:
                el = await page.query_selector(sel)
                if el:
                    text = (await el.inner_text()).strip()
                    if text and len(text) < 120:
                        result["company"] = text
                        break
            except Exception:
                pass

        # JS fallback: h2 near the job title h1, then breadcrumb, then page <title>
        if not result["company"]:
            try:
                result["company"] = await page.evaluate("""() => {
                    // BuiltIn structure: <a href="/company/..."><h2>Company Name</h2></a>
                    // Select h2 that is a direct child of a company link
                    for (const a of document.querySelectorAll('a[href*="/company/"], a[href*="/companies/"]')) {
                        const h2 = a.querySelector(':scope > h2');
                        if (h2) {
                            const t = (h2.innerText || '').trim();
                            if (t && t.length < 120) return t;
                        }
                    }
                    // Breadcrumb first link
                    const bc = document.querySelector('[class*="breadcrumb"], nav[aria-label*="breadcrumb"]');
                    if (bc) {
                        const link = bc.querySelector('a');
                        if (link) {
                            const t = (link.innerText || '').trim();
                            if (t && t.length < 120) return t;
                        }
                    }
                    // Parse "Job Title at Company | Built In" from <title>
                    const m = (document.title || '').match(/\\bat\\s+(.+?)\\s*(?:\\||$)/i);
                    return m ? m[1].trim() : '';
                }""") or ""
            except Exception:
                pass

        for sel in [
            "[data-testid='job-location']", "[data-testid='location']",
            "[data-testid*='remote']", "[data-testid*='work-model']",
            ".job-location", "[class*='Location']",
            "[class*='location']", "[class*='JobLocation']",
            "[class*='remote']", "[class*='workModel']", "[class*='work-model']",
        ]:
            try:
                el = await page.query_selector(sel)
                if el:
                    text = (await el.inner_text()).strip()
                    if text and len(text) < 120:
                        result["location"] = text
                        break
            except Exception:
                pass

        # JS fallback: scan the page for the first line that matches a work-model keyword
        if not result["location"]:
            try:
                result["location"] = await page.evaluate("""() => {
                    const KEYWORDS = ['remote', 'hybrid', 'in-office', 'in office', 'on-site', 'onsite'];
                    for (const el of document.querySelectorAll('li, span, p, div')) {
                        if (el.children.length > 3) continue;
                        const raw = (el.innerText || '');
                        for (const line of raw.split('\\n')) {
                            const t = line.trim();
                            if (!t || t.length > 60) continue;
                            if (KEYWORDS.some(k => t.toLowerCase().includes(k))) return t;
                        }
                    }
                    return '';
                }""") or ""
            except Exception:
                pass

        # Expand "Read Full Description" accordion before extracting
        try:
            for expand_sel in [
                "button:has-text('Read Full Description')",
                "a:has-text('Read Full Description')",
                "button:has-text('Read full description')",
                "a:has-text('Read full description')",
                "[class*='read-full']", "[class*='ReadFull']",
            ]:
                el = await page.query_selector(expand_sel)
                if el and await el.is_visible():
                    await el.click()
                    await page.wait_for_timeout(2000)
                    break
        except Exception:
            pass

        # Primary: JS extractor walks from the "The Role" heading — most reliable on BuiltIn
        try:
            jd_from_role = await page.evaluate("""() => {
                // Find any element whose sole text content is "The Role"
                for (const el of document.querySelectorAll('div, section, h1, h2, h3, h4')) {
                    if ((el.innerText || '').trim() !== 'The Role') continue;
                    // Gather text from siblings that follow, stopping at the next section heading
                    let text = '';
                    let sibling = el.nextElementSibling;
                    while (sibling) {
                        const sibText = (sibling.innerText || '').trim();
                        // Stop at sibling headings like "The Company", "Benefits", etc.
                        if (sibling.matches('h1,h2,h3,h4') && sibText && sibText !== 'The Role') break;
                        text += sibText + '\\n';
                        sibling = sibling.nextElementSibling;
                        if (text.length > 7000) break;
                    }
                    // If siblings gave nothing, grab the parent's text (role content is a child)
                    if (text.trim().length < 100) {
                        const parent = el.parentElement;
                        if (parent) text = (parent.innerText || '').trim();
                    }
                    if (text.trim().length > 100) return text.trim().slice(0, 7000);
                }
                return '';
            }""") or ""
            if jd_from_role:
                result["description"] = jd_from_role
        except Exception:
            pass

        # CSS selector fallbacks if JS walk found nothing
        if not result["description"]:
            for sel in [
                "[data-testid='job-description']", ".job-description",
                "#job-description", "[class*='JobDescription']",
                "[class*='job-description']",
                "[class*='job-details']", "[class*='JobDetails']",
                "section[class*='job']", "[class*='the-role']", "[class*='TheRole']",
                "main article", "article",
            ]:
                try:
                    el = await page.query_selector(sel)
                    if el:
                        text = (await el.inner_text()).strip()
                        if len(text) > 200:
                            result["description"] = text[:7000]
                            break
                except Exception:
                    pass

        # ── Easy Apply detection ──
        # Only runs when jobPostInit was unavailable — `isEasyApply` from the
        # payload is exact, while these selectors can be tripped by the
        # "Easy Apply" badges on the similar-jobs cards further down the page.
        # Check button/link elements directly — avoid scanning broad containers
        # (sticky divs can contain "Easy Apply" text in unrelated parts of the page)
        for sel in [] if trusted_apply else [
            "button:has-text('Easy Apply')", "a:has-text('Easy Apply')",
            "[data-testid*='easy-apply']", ".easy-apply",
        ]:
            try:
                el = await page.query_selector(sel)
                if el and await el.is_visible():
                    result["easy_apply"] = True
                    break
            except Exception:
                pass

        if not trusted_apply and not result["easy_apply"]:
            try:
                result["easy_apply"] = await page.evaluate("""() => {
                    // Only flag Easy Apply when a visible <a> or <button> has exactly
                    // that text — not just any container that mentions "easy apply"
                    for (const el of document.querySelectorAll('a, button')) {
                        const text = (el.innerText || '').trim().toLowerCase();
                        if (text === 'easy apply') return true;
                    }
                    return false;
                }""") or False
            except Exception:
                pass

        # ── External apply URL (legacy DOM fallbacks) ──
        # Only reached if jobPostInit was missing — e.g. Built-in changes the
        # bootstrap call's name. If Easy Apply, the apply URL stays as the
        # Built-in listing URL. Otherwise, try to find the external company
        # apply link.
        if not trusted_apply and not result["easy_apply"]:
            for sel in [
                # BuiltIn's real external-apply button is consistently
                # <a id="applyButton" aria-label="Apply to job" href="...">;
                # the native Easy Apply button never has this id/aria-label
                # (it's aria-label="Easy Apply to job" with no id). Checked
                # live across Workday/Workable/Greenhouse/appcast postings.
                "a#applyButton[href]",
                "a[aria-label='Apply to job'][href]",
                "a[data-testid='apply-button'][href]",
                "a[href*='apply'][target='_blank']",
                "a:has-text('Apply Now')[href]",
                "a:has-text('Apply on')[href]",
                "a:has-text('Apply')[href]",
                "[class*='apply'] a[href]",
                "[class*='Apply'] a[href]",
            ]:
                try:
                    el = await page.query_selector(sel)
                    if el:
                        href = _external_apply_url(await el.get_attribute("href") or "", job_url)
                        if href:
                            result["apply_url"] = href
                            break
                except Exception:
                    pass

            # JS fallback: any <a> whose text normalises to "APPLY" or "APPLY NOW"
            # Strip non-alpha chars first so the pencil-icon prefix ("✎ APPLY") doesn't
            # prevent a match. Also handles protocol-relative href ("//greenhouse.io/...")
            if result["apply_url"] == job_url:
                try:
                    href = await page.evaluate("""() => {
                        for (const a of document.querySelectorAll('a[href]')) {
                            const raw = (a.innerText || a.textContent || '').trim();
                            const text = raw.replace(/[^A-Za-z\\s]/g, '').trim().toUpperCase();
                            let href = a.getAttribute('href') || '';
                            if (href.startsWith('//')) href = 'https:' + href;
                            if ((text === 'APPLY' || text === 'APPLY NOW') &&
                                (href.startsWith('http')) &&
                                !href.toLowerCase().includes('builtin.com')) {
                                return href;
                            }
                        }
                        return '';
                    }""") or ""
                    if href:
                        result["apply_url"] = href
                except Exception:
                    pass

            # Click-and-capture fallback: open Apply button in new tab, capture URL
            if result["apply_url"] == job_url:
                try:
                    for btn_sel in [
                        "a:has-text('Apply')", "button:has-text('Apply')",
                        "[class*='apply-btn']", "[class*='ApplyBtn']",
                    ]:
                        btn = await page.query_selector(btn_sel)
                        if not btn or not await btn.is_visible():
                            continue
                        # Never click Built-in's own auth links. Logged out,
                        # the sticky bar shows "Log In to apply" / "Sign up to
                        # apply", both of which match a:has-text('Apply') and
                        # navigate this tab away from the job page.
                        btn_href = (await btn.get_attribute("href") or "").lower()
                        btn_text = ((await btn.inner_text()) or "").lower()
                        if "/auth/" in btn_href or "builtin.com/auth" in btn_href:
                            continue
                        if "log in" in btn_text or "sign up" in btn_text:
                            continue

                        async with page.context.expect_page(timeout=10000) as new_page_info:
                            await btn.click()
                        new_tab = await new_page_info.value
                        await new_tab.wait_for_load_state("domcontentloaded", timeout=10000)
                        external_url = new_tab.url
                        await new_tab.close()
                        external_url = _external_apply_url(external_url, job_url)
                        if external_url:
                            result["apply_url"] = external_url
                            break
                except Exception:
                    pass

        # Loud, explicit failure signal. Previously this fell through silently
        # and the Built-in listing URL was written to the sheet as if it were
        # the real apply link.
        if not result["easy_apply"] and result["apply_url"] == job_url:
            if result["login_required"]:
                log(f"    [warn] Built-in requires a logged-in session to reveal this apply URL — falling back to Built-in URL: {job_url}")
            else:
                log(f"    [warn] no external apply URL found — falling back to Built-in URL: {job_url}")

    except PWTimeout:
        log(f"    [timeout] {job_url}")
    except Exception as e:
        log(f"    [error]   {job_url}: {e}")

    return result


# ── LIST PAGE ──────────────────────────────────────────────────────────────────

async def extract_jobs_from_page(list_page) -> List[dict]:
    jobs: List[dict] = []
    seen_hrefs: set = set()
    BAD_FRAGMENTS = ["/company/", "/companies/", "/topic/", "/tech-hub/", "/author/", "/people/"]

    SELECTORS = [
        "a[href*='/job/']",
        "a[href*='/jobs/view/']",
        "[data-testid*='job-card'] a",
        ".job-card a[href]",
        "article a[href]",
    ]

    links = []
    used_sel = None
    for sel in SELECTORS:
        try:
            found = await list_page.query_selector_all(sel)
            if len(found) >= 2:
                links = found
                used_sel = sel
                break
        except Exception:
            continue

    if not links:
        try:
            log(f"  [warn] No job links found. Page title: {await list_page.title()}")
        except Exception:
            log("  [warn] No job links found.")
        return jobs

    log(f"  [selector] {used_sel!r} — {len(links)} raw links")

    # Build href → work-model map for the whole page in one JS call.
    # Walks up to 8 levels from each job link, finds the first sibling/cousin
    # element whose text matches a work-model keyword, returns only the matching line.
    href_to_workmodel: dict = {}
    try:
        href_to_workmodel = await list_page.evaluate("""() => {
            const KEYWORDS = ['remote', 'hybrid', 'in-office', 'in office', 'on-site', 'onsite'];
            function matchLine(text) {
                for (const line of text.split('\\n')) {
                    const t = line.trim();
                    if (!t || t.length > 60) continue;
                    if (KEYWORDS.some(k => t.toLowerCase().includes(k))) return t;
                }
                return '';
            }
            const map = {};
            for (const link of document.querySelectorAll('a[href*="/job/"]')) {
                const raw = link.getAttribute('href') || '';
                const href = raw.startsWith('http') ? raw : 'https://builtin.com' + raw;
                let node = link;
                found: for (let i = 0; i < 8; i++) {
                    node = node.parentElement;
                    if (!node) break;
                    for (const child of node.querySelectorAll('span, div, p, li')) {
                        if (child.contains(link)) continue;
                        const matched = matchLine(child.innerText || '');
                        if (matched) { map[href] = matched; break found; }
                    }
                }
            }
            return map;
        }""") or {}
    except Exception:
        pass

    for link in links:
        try:
            href = (await link.get_attribute("href") or "").strip()
            text = (await link.inner_text()).strip()
            if not href or not text or len(text) < 3:
                continue
            if href.startswith("/"):
                href = "https://builtin.com" + href
            if href in seen_hrefs:
                continue
            if any(bad in href for bad in BAD_FRAGMENTS):
                continue
            seen_hrefs.add(href)
            work_model = href_to_workmodel.get(href, "")
            jobs.append({"title": text, "url": href, "location": work_model})
        except Exception:
            continue

    return jobs


async def find_next_url(list_page, page_num: int, current_url: str) -> Optional[str]:
    NEXT_SELECTORS = [
        "a[aria-label*='Next']", "a[rel='next']",
        "button[aria-label*='Next']", "a:has-text('Next')",
        "[data-testid*='pagination'] a:last-child",
        "nav[aria-label*='agination'] a:last-child",
    ]
    for sel in NEXT_SELECTORS:
        try:
            el = await list_page.query_selector(sel)
            if el:
                href = await el.get_attribute("href")
                if href:
                    return href if href.startswith("http") else f"https://builtin.com{href}"
                await el.click()
                await list_page.wait_for_load_state("networkidle", timeout=15_000)
                new_url = list_page.url
                return new_url if new_url != current_url else None
        except Exception:
            continue

    parsed = urlparse(current_url)
    qs = parse_qs(parsed.query, keep_blank_values=True)
    for param in ("page", "p"):
        if param in qs:
            try:
                qs[param] = [str(int(qs[param][0]) + 1)]
                return urlunparse(parsed._replace(query=urlencode(qs, doseq=True)))
            except (ValueError, IndexError):
                pass

    if page_num == 1 and "page=" not in current_url:
        sep = "&" if "?" in current_url else "?"
        return f"{current_url}{sep}page=2"

    return None


# ── MAIN SCRAPE ────────────────────────────────────────────────────────────────

async def scrape_builtin(start_url: str, config: dict) -> List[dict]:
    candidates: List[dict] = []
    seen_urls: set         = set()
    visited_detail         = 0  # detail pages hit this run — drives DETAIL_PAUSE pacing
    seen_fingerprints: set = set()  # (company.lower(), title.lower()) — within-run dedup

    preferred_locs  = config.get("preferred_locations",   ["remote", "miami"])
    seniority_tiers = config.get("seniority_tiers",       _DEFAULT_SENIORITY_TIERS)
    title_keywords  = [k.lower() for k in config.get("title_keywords",       _DEFAULT_TITLE_KEYWORDS)]
    excl_titles     = [k.lower() for k in config.get("excluded_titles",      list(_DEFAULT_EXCLUDED_TITLES))]
    excl_words      = [k.lower() for k in config.get("excluded_title_words", list(_DEFAULT_EXCLUDED_TITLE_WORDS))]
    max_candidates  = int(config.get("max_candidates") or MAX_CANDIDATES)

    log(f"[Pipeline Settings — applied to all scrapers]")
    log(f"  Preferred locations : {preferred_locs}")
    log(f"  Seniority tiers     : {seniority_tiers}")
    log(f"  Title keywords      : {title_keywords}")
    log(f"  Scoring             : programmatic (no Claude)")
    log(f"  Already-applied     : checked against Applications + Skips tabs\n")
    log(f"[Built-in] Starting — target: {max_candidates} candidates")
    log(f"[Built-in] URL: {start_url}\n")

    # Load already-applied jobs before starting the browser
    applied_pairs, applied_urls = load_applied_jobs()

    if SESSION_FILE.exists():
        log("[Built-in] Reusing saved Built-in session...")
    else:
        log("[Built-in] No saved session — apply URLs gated behind a Built-in "
            "login will fall back to the Built-in listing URL. Use the "
            "'Re-login Built-in' button to fix this.")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context_kwargs = dict(
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1280, "height": 900},
        )
        if SESSION_FILE.exists():
            context_kwargs["storage_state"] = str(SESSION_FILE)
        context = await browser.new_context(**context_kwargs)

        # Cloudflare's rate limit counts every request, not just documents, and
        # a Built-in job page pulls a long tail of images/fonts/tracking pixels
        # we never read. Dropping them keeps far more headroom under the limit
        # (stylesheets stay, so is_visible() checks still see real layout).
        async def _block_heavy_assets(route):
            if route.request.resource_type in ("image", "font", "media"):
                await route.abort()
            else:
                await route.continue_()

        await context.route("**/*", _block_heavy_assets)

        list_page   = await context.new_page()
        detail_page = await context.new_page()

        current_url = start_url
        page_num    = 1
        tracker     = RateLimitTracker()
        consecutive_limited_pages = 0

        while page_num <= MAX_PAGES and len(candidates) < max_candidates:
            log(f"[Built-in] ── Page {page_num} ──────────────────────────────────────")
            log(f"[Built-in] {current_url}")
            events_before_page = tracker.events

            try:
                await list_page.goto(current_url, wait_until="networkidle", timeout=NAV_TIMEOUT)
            except PWTimeout:
                log("[Built-in] networkidle timeout — proceeding with what loaded.")
            except Exception as e:
                log(f"[Built-in] Failed to load page: {e}")
                break

            page_jobs = await extract_jobs_from_page(list_page)

            kept         = 0
            not_pm       = 0
            excluded     = 0
            bad_location = 0
            already_app  = 0
            dup          = 0
            blocked      = 0

            for job in page_jobs:
                if len(candidates) >= max_candidates:
                    break

                title = job["title"]
                url   = job["url"]

                if url in seen_urls:
                    dup += 1
                    continue

                if is_excluded_title(title, excl_titles, excl_words):
                    excluded += 1
                    continue

                if not is_matching_title(title, title_keywords):
                    not_pm += 1
                    continue

                # Pre-check: skip detail page entirely if URL already known
                _pre_norm = normalize_url(url)
                _pre_bid  = _builtin_job_id(url)
                if (_pre_norm and _pre_norm in applied_urls) or \
                   (_pre_bid and f"builtin-id:{_pre_bid}" in applied_urls):
                    already_app += 1
                    continue

                seen_urls.add(url)

                # Location pre-filter using the list-card value (avoids unnecessary detail visits)
                list_location = job.get("location", "")
                if list_location and not is_valid_location(list_location, preferred_locs):
                    bad_location += 1
                    continue

                # Title (and list-page location) passed — visit detail page.
                # Pace the visits: Cloudflare rate-limits on request volume and
                # answers with a challenge page once we go too fast (see
                # DETAIL_PAUSE).
                if visited_detail:
                    await asyncio.sleep(DETAIL_PAUSE)
                visited_detail += 1

                detail = await scrape_job_detail(detail_page, url, tracker)

                # Challenged on every attempt — the page we saw was Cloudflare's
                # interstitial, not the job. Its company/apply URL would both be
                # wrong, so drop the job rather than log a bad row.
                if detail["blocked"]:
                    blocked += 1
                    continue

                job["company"]     = detail["company"]
                job["description"] = detail["description"]
                job["easy_apply"]  = detail["easy_apply"]
                job["apply_url"]   = detail["apply_url"]
                # Prefer list-page work model (Remote/Hybrid) over detail-page location text
                job["location"]    = list_location or detail["location"]

                if not is_valid_location(job["location"], preferred_locs):
                    bad_location += 1
                    continue

                # Within-run duplicate: same company+title found via a different URL
                _fp = (job.get("company", "").lower().strip(), title.lower().strip())
                if _fp[1] and _fp in seen_fingerprints:
                    dup += 1
                    continue
                if _fp[1]:
                    seen_fingerprints.add(_fp)

                # Already-applied check (uses actual apply URL + Built-in ID + company/title pairs)
                if already_applied(job, applied_pairs, applied_urls):
                    already_app += 1
                    continue

                score, matched_tier = seniority_score(title, seniority_tiers)
                job["score"]  = score
                job["reason"] = f"Seniority: {matched_tier}" if matched_tier else "No seniority tier matched"

                kept += 1
                candidates.append(job)
                log(f"  ✓ Candidate #{len(candidates)} found")

            log(
                f"\n  Kept {kept} | Not PM: {not_pm} | Excluded: {excluded} "
                f"| Bad location: {bad_location} | Already applied: {already_app} | Dup: {dup}"
                f"{f' | Rate-limited: {blocked}' if blocked else ''}"
            )
            log(f"  Candidates so far: {len(candidates)} / {max_candidates}")

            # Individual rate-limit hits are already retried (DETAIL_BACKOFF)
            # and don't stop anything on their own — this is for when nearly
            # every page in a row eats at least one, which means Cloudflare
            # is rate-limiting the session as a whole, not one unlucky job.
            # Grinding through MAX_PAGES at 10-40s per hit for a handful more
            # candidates isn't worth it — stop and hand off what we have.
            if tracker.events > events_before_page:
                consecutive_limited_pages += 1
            else:
                consecutive_limited_pages = 0

            if consecutive_limited_pages >= RATE_LIMIT_PAGE_THRESHOLD:
                log(
                    f"\n[Built-in] Rate-limited on {consecutive_limited_pages} pages in a row — "
                    f"stopping early with {len(candidates)} candidate(s) and moving on to the resume pipeline."
                )
                break

            if len(candidates) >= max_candidates:
                log(f"\n[Built-in] Reached {max_candidates} candidates — done scraping.")
                break

            next_url = await find_next_url(list_page, page_num, current_url)
            if next_url and next_url != current_url:
                current_url = next_url
                page_num   += 1
                await asyncio.sleep(1.5)
            else:
                log("\n[Built-in] No next page — done.")
                break

        await browser.close()

    # Sort by seniority score before printing
    candidates.sort(key=lambda j: j["score"], reverse=True)

    # ── Summary ────────────────────────────────────────────────────────────────
    log(f"\n[Built-in] {'═' * 52}")
    log(f"[Built-in] Candidates collected: {len(candidates)} / {max_candidates}")

    if candidates:
        log("\n[Built-in] Candidate list (sorted by seniority score):")
        for i, job in enumerate(candidates, 1):
            ea       = "Easy Apply" if job.get("easy_apply") else "Apply via site"
            apply_url = job.get("apply_url") or job.get("url", "")
            log(f"\n  {i}. {job['title']}  [{job['score']}/10]")
            log(f"     Company:   {job.get('company') or '?'}")
            log(f"     Location:  {job.get('location') or '?'}")
            log(f"     Apply:     {ea}")
            log(f"     Apply URL: {apply_url}")
    else:
        log("[Built-in] No candidates found. Check your search URL or loosen the filters.")

    return candidates


# ── ENTRY POINT ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    _config = load_config()
    _url = _config.get("builtin_url", "").strip()
    if not _url:
        log("[Built-in] ERROR: No Built-in URL set. Go to Settings and enter your search URL.")
        sys.exit(1)

    # Run async scraping first — event loop is fully closed before resume_pipeline
    # (which uses sync Playwright) is called.
    _candidates = asyncio.run(scrape_builtin(_url, _config))

    if not _candidates:
        log("\n[Built-in] No candidates to score — exiting.")
        sys.exit(0)

    # Remap fields so resume_pipeline gets the right URLs in the right columns:
    #   url         → actual apply URL  (written to "URL" column in sheet)
    #   linkedin_url → Built-in listing URL (written to "Linked In URL" column)
    for _job in _candidates:
        _builtin_url = _job.get("url", "")
        _apply_url   = _job.get("apply_url") or _builtin_url
        _job["url"]          = _apply_url
        _job["linkedin_url"] = _builtin_url

    log("\n[Built-in] Handing off to resume pipeline...\n")
    from resume_pipeline import run_pipeline
    run_pipeline(_candidates)
