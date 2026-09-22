#!/usr/bin/env python3
"""
Indeed scraper — scrapes candidates, then hands off to resume_pipeline.py
for Claude scoring, resume building, and Google Sheet logging, same as
dynamite_scraper.py and builtin_scraper.py.

ARCHITECTURE — why this scrapes rendered HTML (unlike Dynamite's public API):
A bare HTTP request to indeed.com gets an immediate 403 "Security Check"
wall — checked live. But a real headless Playwright session with a normal
user-agent gets served real results cleanly (200, no CAPTCHA) — Indeed's bot
defense is clearly aimed at non-browser clients, not headless browsers as
such. So this scrapes like builtin_scraper.py: a real browser, DOM
selectors, no public API to call directly.

Indeed's search results page is a single-page list+detail UI: clicking a
job's title updates the URL with `&vjk=<job key>` and refreshes a detail
pane in place, no full navigation. That vjk-bearing URL becomes this
scraper's "listing" URL (the "Linked In URL" column) — it's exactly what
you see in your own browser when you select a job yourself.

Apply URL, per job: two buttons, checked live across several postings.
  - "Apply on company site" is real per-job, but resolving it requires a
    logged-in Indeed session — 5/5 unauthenticated clicks tested hit
    Indeed's own sign-in wall (secure.indeed.com/auth...thirdpartysignin)
    instead of the employer's site. Without a session, resolve_apply_url
    reports login_required and the listing URL is kept, same pattern as
    Built-in's session-gated postings.
  - "Apply with Indeed" (native, button id `indeedApplyButton`) has no
    external destination — Indeed hosts the application itself, so
    apply_url is just the listing URL (easy_apply=True), per instruction.

Resolution is DEFERRED until after Claude scoring (resolve_apply_url, passed
to resume_pipeline.run_pipeline as a callback) and only done for jobs that
clear the score threshold. The /applystart redirect is Indeed's most
Cloudflare-sensitive endpoint; resolving it eagerly for every candidate —
most of which score below threshold and are discarded — was 10 jobs x up to
4 attempts = ~40 hits per run, which got this machine's residential IP
hard-blocked ("Additional Verification Required", no challenge widget at
all, even in a real logged-in Chrome). Deferring cuts that to 1-3 hits.

Browser fingerprint: real Google Chrome (channel="chrome") with its GENUINE
user-agent, never an override. A spoofed UA string is contradicted by the
Sec-CH-UA client-hint headers Chromium sends with the real build number —
checked live: Playwright's bundled Chromium 148 claiming "Chrome/131" made
Cloudflare re-challenge in a loop even after a human solved the Turnstile
(solve passes, post-solve fingerprint integrity check fails, re-challenge).
Playwright's new-headless mode of real Chrome reports a normal "Chrome"
UA, not "HeadlessChrome". Falls back to the bundled Chromium only if Chrome
isn't installed.

Honeypot: Indeed injects a hidden decoy job-title link into the list DOM —
a duplicate title, zero rendered size, with a deterministic/fake job key
(observed: "123456789abcdef0") — a classic anti-scraping trap for scripts
that blindly harvest every link. Filtered by Playwright's `is_visible()`,
which correctly reports the 0x0 element as not visible even though its CSS
`visibility` computed style says "visible".
"""

import asyncio
import json
import os
import random
import re
import sys
import time
from pathlib import Path
from typing import Optional, List
from urllib.parse import urlparse, parse_qs

from dotenv import load_dotenv
from playwright.async_api import async_playwright

from dedup_common import pairs_match

load_dotenv()

SCRIPT_DIR     = Path(__file__).parent
CONFIG_FILE    = SCRIPT_DIR / "config.json"
SESSION_FILE   = SCRIPT_DIR / "indeed_session.json"
MAX_PAGES      = 20
MAX_CANDIDATES = 10
NAV_TIMEOUT    = 30_000

# Randomized pause between per-job detail-pane clicks. Each click fires
# Indeed XHRs that count toward Cloudflare's per-IP volume score; a fixed
# sub-second cadence reads as a script.
DETAIL_PAUSE_RANGE = (2.0, 4.0)

# One retry only for a challenged /applystart. Every attempt against a
# challenged endpoint is another strike on the IP's reputation — the old
# 10/20/40s triple retry was part of what got the IP hard-blocked.
APPLY_RETRY_DELAY = 8

# See module docstring: real Chrome, genuine UA. AutomationControlled off so
# navigator.webdriver isn't set — the one automation tell that is safe to
# remove without creating a new inconsistency.
BROWSER_ARGS = ["--no-sandbox", "--disable-blink-features=AutomationControlled"]

GOOGLE_CREDS_FILE = Path(__file__).parent / os.getenv("GOOGLE_CREDS_FILE", "google_credentials.json")


# Headless Chrome announces itself as "HeadlessChrome/153.0.0.0" in the UA
# string (checked live, Chrome 153 via Playwright's new-headless) while its
# Sec-CH-UA brands say "Google Chrome/153" — Cloudflare scores that token
# hard. The ONLY safe correction is to take the browser's own UA and drop the
# "Headless" token, so version, platform and client hints all still agree.
# Never substitute a hand-written UA string (see module docstring).
def _unheadless(ua: str) -> str:
    return ua.replace("HeadlessChrome/", "Chrome/")


async def launch_browser(pw, headless: bool = True):
    """Real Chrome if installed (see module docstring), else bundled Chromium.
    Returns (browser, user_agent) — the UA to pass to new_context()."""
    try:
        browser = await pw.chromium.launch(channel="chrome", headless=headless, args=BROWSER_ARGS)
    except Exception as e:
        log(f"[Indeed] Google Chrome not available ({str(e).splitlines()[0][:80]}) — "
            "falling back to Playwright's bundled Chromium (more likely to be challenged).")
        browser = await pw.chromium.launch(headless=headless, args=BROWSER_ARGS)
    probe = await browser.new_context()
    try:
        ua = await (await probe.new_page()).evaluate("() => navigator.userAgent")
    finally:
        await probe.close()
    return browser, _unheadless(ua)


def launch_browser_sync(pw, headless: bool = True):
    """Sync twin of launch_browser, for resolve_apply_url (runs inside
    resume_pipeline's sync Playwright)."""
    try:
        browser = pw.chromium.launch(channel="chrome", headless=headless, args=BROWSER_ARGS)
    except Exception as e:
        log(f"[Indeed] Google Chrome not available ({str(e).splitlines()[0][:80]}) — "
            "falling back to Playwright's bundled Chromium (more likely to be challenged).")
        browser = pw.chromium.launch(headless=headless, args=BROWSER_ARGS)
    probe = browser.new_context()
    try:
        ua = probe.new_page().evaluate("() => navigator.userAgent")
    finally:
        probe.close()
    return browser, _unheadless(ua)


def context_kwargs(user_agent: Optional[str] = None) -> dict:
    """user_agent must come from launch_browser*() — never hand-written."""
    kw = dict(viewport={"width": 1280, "height": 900}, locale="en-US")
    if user_agent:
        kw["user_agent"] = user_agent
    if SESSION_FILE.exists():
        kw["storage_state"] = str(SESSION_FILE)
    return kw


def check_indeed_session():
    """
    Checks the saved Indeed session by inspecting cookies for the indeed.com
    domain. Mirrors builtin_scraper.check_builtin_session() — no HTTP
    request, just a local expiry check. Returns (valid: bool, reason: str).
    """
    if not SESSION_FILE.exists():
        return False, "No session file found"
    try:
        data = json.loads(SESSION_FILE.read_text())
        cookies = [c for c in data.get("cookies", []) if "indeed.com" in c.get("domain", "")]
        if not cookies:
            return False, "No indeed.com cookies — please re-login"
        now = time.time()
        if all(c.get("expires", -1) != -1 and c["expires"] < now for c in cookies):
            return False, "Session cookies expired — please re-login"
        return True, "Session active"
    except Exception as e:
        return False, f"Could not read session file: {e}"

# ── TITLE / LOCATION / SENIORITY FILTERS (mirrors builtin_scraper.py) ──────────

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
    t = title.lower()
    _keywords = keywords if keywords is not None else _DEFAULT_TITLE_KEYWORDS
    if not _keywords:
        return True
    return any(kw in t for kw in _keywords)


def is_valid_location(location: str, preferred_locs: Optional[List[str]] = None) -> bool:
    if preferred_locs is None:
        preferred_locs = ["remote", "miami"]
    if not location:
        return True
    loc = location.lower()
    if "anywhere" in loc:
        return True
    for pref in preferred_locs:
        if pref.lower() in loc:
            return True
    return False


def seniority_score(title: str, tiers: Optional[List[str]] = None) -> tuple:
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


def _indeed_job_key(url: str) -> str:
    """The `jk` query param — Indeed's actual unique job identifier. Needed
    because job identity here lives entirely in the query string (?...&vjk=
    or ?...&jk=), unlike Built-in/Dynamite where it's in the URL path — the
    normalize_url() convention used elsewhere (strips query strings for
    dedup) would collapse every Indeed job on the same search to one
    identical URL and silently break deduplication."""
    try:
        qs = parse_qs(urlparse(url).query)
    except Exception:
        return ""
    return (qs.get("vjk", qs.get("jk", [""]))[0] or "").strip().lower()


def normalize_url(url: str) -> str:
    if not url:
        return ""
    url = str(url).strip().lower()
    url = re.sub(r'#.*$', '', url)
    url = re.sub(r'\?.*$', '', url)
    url = url.rstrip('/')
    url = re.sub(r'^(https?://)www\.', r'\1', url)
    return url


def load_applied_jobs():
    """Loads applied/skipped jobs from Google Sheets. Returns (applied_pairs, applied_urls)."""
    log("Loading applied jobs and skips from Google Sheets...")
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
        creds  = Credentials.from_service_account_file(str(GOOGLE_CREDS_FILE), scopes=scopes)
        gc     = gspread.authorize(creds)

        sheet_id = os.getenv("GOOGLE_SHEET_ID", "YOUR_GOOGLE_SHEET_ID")
        try:
            cfg = json.loads(CONFIG_FILE.read_text())
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
                        jk = _indeed_job_key(val)
                        if jk:
                            # Indeed URL: dedup by job key ONLY. normalize_url
                            # strips the query string, and an Indeed listing's
                            # identity lives entirely there (?...&vjk=<jk>), so
                            # every Indeed URL collapses to a bare
                            # "indeed.com/jobs" — adding that would poison the
                            # set and mark every future Indeed candidate as
                            # already-applied. (This is the exact failure
                            # _indeed_job_key's own docstring warns about.)
                            applied_urls.add(f"indeed-jk:{jk}")
                            continue
                        norm = normalize_url(val)
                        if norm and norm.startswith("http"):
                            applied_urls.add(norm)

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
    # Checks both "url" and "linkedin_url" — at this point both are the Indeed
    # listing URL (resolve_apply_url runs later, post-scoring). For an Indeed
    # URL, match by job key ONLY, never normalize_url: it strips the query
    # string where an Indeed listing's identity lives, so every Indeed URL
    # collapses to a bare "indeed.com/jobs" and the normalize check would
    # match every candidate against any one Indeed row in the sheet. See
    # load_applied_jobs for the mirror of this on the loading side.
    for field in ["url", "linkedin_url"]:
        val = job.get(field, "")
        jk = _indeed_job_key(val)
        if jk:
            if f"indeed-jk:{jk}" in applied_urls:
                return True
            continue
        norm = normalize_url(val)
        if norm and norm in applied_urls:
            return True

    company = job.get("company", "").lower().strip()
    title   = job.get("title", "").lower().strip()
    return pairs_match(company, title, applied_pairs)


# ── LIST PAGE ──────────────────────────────────────────────────────────────────

# Badge lines Indeed injects into a card's text ABOVE the company name. Any
# of these present shifts every positional assumption by one line — checked
# live, "Easily apply" cards parsed as company="Easily apply",
# location="<company>", which then failed the location filter. The fallback
# parser below skips them; the primary path avoids the problem entirely by
# reading the dedicated elements.
_CARD_BADGES = {
    "easily apply", "urgently hiring", "responsive employer",
    "hiring multiple candidates", "new", "just posted", "sponsored",
}


def _is_salary_line(line: str) -> bool:
    low = line.lower()
    return "$" in line or "an hour" in low or "a year" in low or "a month" in low


def _parse_card_text(text: str) -> tuple:
    """Fallback only: company/location/salary from a card's innerText when
    the data-testid elements aren't found. Line 0 is the title (read
    separately from the title element); after skipping badge lines the
    next is the company; salary is wherever the "$"/"a year" line is."""
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    rest = [l for l in lines[1:] if l.lower() not in _CARD_BADGES]
    company  = rest[0] if rest else ""
    location = ""
    salary   = ""
    for line in rest[1:]:
        if _is_salary_line(line):
            salary = line
        elif not location:
            location = line
    return company, location, salary


async def _card_fields(card) -> tuple:
    """Company/location/salary via the card's dedicated elements (checked
    live: data-testid="company-name" / "text-location", and the salary
    snippet container). Falls back to positional text parsing per field
    only when an element is missing, so a badge line can't shift the rest."""
    company = location = salary = ""
    try:
        company, location, salary = await card.evaluate("""el => {
            const q = sel => { const n = el.querySelector(sel); return n ? n.innerText.trim() : ""; };
            return [q('[data-testid="company-name"]'), q('[data-testid="text-location"]'),
                    q('.salary-snippet-container') || q('[data-testid="attribute_snippet_testid"]')];
        }""")
    except Exception:
        pass
    if not company or not location:
        fb_company, fb_location, fb_salary = _parse_card_text((await card.inner_text()).strip())
        company  = company  or fb_company
        location = location or fb_location
        salary   = salary   or fb_salary
    if salary and not _is_salary_line(salary):
        salary = ""  # attribute_snippet can be a non-pay attribute ("Full-time")
    return company, location, salary


async def extract_jobs_from_page(page) -> List[dict]:
    """Cheap, no-click pass over the list DOM — title/company/location/
    salary/jk for every real (visible) card. Detail-only fields (JD, apply
    button kind) are filled in later, only for candidates that pass the
    title/location filters, via _select_and_scrape_detail."""
    jobs: List[dict] = []
    cards = await page.query_selector_all("td.resultContent")
    for card in cards:
        try:
            if not await card.is_visible():
                continue  # honeypot / not-yet-hydrated placeholder
            link = await card.query_selector("a[data-jk]")
            if not link:
                continue
            jk = (await link.get_attribute("data-jk") or "").strip()
            if not jk:
                continue
            title = (await link.inner_text()).strip()
            if not title:
                continue
            company, location, salary = await _card_fields(card)
            jobs.append({
                "title": title, "jk": jk, "company": company,
                "location": location, "salary": salary,
            })
        except Exception:
            continue
    return jobs


async def find_next_url(page) -> Optional[str]:
    """Prefer the pagination nav's own "Next Page" link — reliable and
    avoids guessing the `start=` increment ourselves (checked live: it's
    10 per page, not the 15 actually shown, so guessing would drift)."""
    try:
        el = await page.query_selector("a[aria-label='Next Page']")
        if el:
            href = await el.get_attribute("href")
            if href:
                return href if href.startswith("http") else f"https://www.indeed.com{href}"
    except Exception:
        pass
    return None


# ── DETAIL (same-page click, not a separate navigation) ────────────────────────

def _classify_apply_dest(dest: str) -> str:
    """Where an apply-redirect actually landed. "resolved" only when we
    genuinely left indeed.com — checked live, a Cloudflare challenge on
    /applystart leaves the URL on www.indeed.com/applystart itself (with a
    __cf_chl_rt_tk param appended, no further redirect) rather than
    bouncing to the secure.indeed.com/auth sign-in wall, so checking only
    for that one specific login-wall URL let this exact case silently
    pass as "resolved" while the apply_url was still a raw Indeed URL.
    Any hostname still ending in indeed.com is a failure to resolve,
    regardless of which specific indeed.com page it is."""
    try:
        host = (urlparse(dest).hostname or "").lower()
    except Exception:
        return "blocked"
    if not host.endswith("indeed.com"):
        return "resolved"
    if host == "secure.indeed.com" and "/auth" in dest:
        return "login_required"
    return "blocked"


_EASY_APPLY_SELECTORS = (
    "#indeedApplyButton",
    "[data-testid='indeedApplyButton-test']",
    "[data-testid='viewjob-indeed-apply']",
)
_EXTERNAL_APPLY_SELECTORS = (
    "[data-testid='viewjob-apply']",
    "button:has-text('Apply on company site')",
)

# JD lives in the split-pane detail view. Checked live 2026-09-20: the old
# "#jobDescriptionText" id is gone from Indeed's current markup — the JD is
# now a React-rendered ".react-native-html-content simple-job-description-*"
# block, with "[data-testid='viewjob-job-content']" (JD + a little chrome) as
# a broader fallback. The pane populates ~1s AFTER the card click, so
# _read_pane_jd polls rather than reading once.
_JD_SELECTORS = (
    "[class*='simple-job-description']",
    "[data-testid='viewjob-job-content']",
)


async def _read_pane_jd(page, timeout_ms: int = 6000) -> str:
    """Polls the detail pane for the JD text after a card click."""
    deadline = timeout_ms
    while deadline > 0:
        for sel in _JD_SELECTORS:
            try:
                el = await page.query_selector(sel)
                if el:
                    txt = (await el.inner_text()).strip()
                    if len(txt) > 200:
                        return txt[:8000]
            except Exception:
                pass
        await page.wait_for_timeout(400)
        deadline -= 400
    return ""


async def _select_and_scrape_detail(page, jk: str) -> dict:
    """Clicks this job's title link (selecting it in the split-pane UI,
    which updates the URL to include &vjk=<jk> without a full page
    navigation) and reads back the detail pane: JD text plus which kind of
    apply button it has. Does NOT resolve the external apply URL — that's
    deferred to resolve_apply_url after scoring (see module docstring).

    result["apply_outcome"]:
      "easy_apply"      — native "Apply with Indeed"; listing URL is final.
      "deferred"        — external button found; resolve later if scored in.
      "no_button_found" — neither button visible.
    result["apply_href"] is the /applystart href when logged in (the <a>
    form of the button); empty when logged out (a <button> with a JS click
    handler only). Checked live: the element changes tag entirely with auth
    state, data-testid is the one stable hook across both.
    """
    result = {
        "description": "", "listing_url": "", "apply_href": "",
        "easy_apply": False, "apply_outcome": "no_button_found",
    }
    link = await page.query_selector(f"a[data-jk='{jk}']")
    if not link or not await link.is_visible():
        return result
    try:
        await link.click()
        await page.wait_for_timeout(1800)
    except Exception:
        return result

    result["listing_url"] = page.url
    result["description"] = await _read_pane_jd(page)

    # Logged out it's <button id="indeedApplyButton">; logged in that id is
    # gone and it's data-testid="viewjob-indeed-apply". Check all forms.
    try:
        for sel in _EASY_APPLY_SELECTORS:
            el = await page.query_selector(sel)
            if el and await el.is_visible():
                result["easy_apply"] = True
                result["apply_outcome"] = "easy_apply"
                return result
    except Exception:
        pass

    try:
        for sel in _EXTERNAL_APPLY_SELECTORS:
            el = await page.query_selector(sel)
            if el and await el.is_visible():
                result["apply_outcome"] = "deferred"
                result["apply_href"] = (await el.get_attribute("href")) or ""
                break
    except Exception:
        pass

    return result


# ── DEFERRED APPLY-URL RESOLUTION (called by resume_pipeline after scoring) ─────

def resolve_apply_url(job: dict, pw) -> Optional[str]:
    """resume_pipeline.run_pipeline() callback: for a job that cleared the
    score threshold, follow Indeed's /applystart redirect to the employer's
    real apply URL. Returns it, or None to keep the listing URL.

    Runs inside resume_pipeline's sync Playwright (`pw` is its playwright
    object) in its own short-lived real-Chrome browser — the scraper's async
    browser is long closed by the time scoring finishes, and 1-3 calls per
    run doesn't justify keeping one open. Loads the same indeed_session.json
    so the session cookies (and any cf_clearance earned in the same browser
    build) carry over.

    One attempt + one retry: see APPLY_RETRY_DELAY.
    """
    if job.get("_apply_outcome") != "deferred":
        return None

    href        = job.get("_apply_href", "")
    listing_url = job.get("linkedin_url", "")
    if not href and not listing_url:
        return None

    def _block_heavy(route):
        if route.request.resource_type in ("image", "font", "media"):
            route.abort()
        else:
            route.continue_()

    browser, ua = launch_browser_sync(pw, headless=True)
    try:
        context = browser.new_context(**context_kwargs(ua))
        context.route("**/*", _block_heavy)
        pg = context.new_page()

        dest = ""
        outcome = "blocked"
        for attempt in range(2):
            if attempt:
                print(f"      [Indeed] apply redirect challenged — one retry in {APPLY_RETRY_DELAY}s")
                time.sleep(APPLY_RETRY_DELAY)
            try:
                if href:
                    pg.goto(href, wait_until="domcontentloaded", timeout=20000)
                    dest = pg.url
                else:
                    # Logged out: no href, only a JS click handler. Open the
                    # listing and click the button; the destination opens in
                    # a new tab.
                    pg.goto(listing_url, wait_until="domcontentloaded", timeout=20000)
                    pg.wait_for_timeout(1500)
                    btn = None
                    for sel in _EXTERNAL_APPLY_SELECTORS:
                        btn = pg.query_selector(sel)
                        if btn and btn.is_visible():
                            break
                        btn = None
                    if not btn:
                        outcome = "no_button_found"
                        break
                    with context.expect_page(timeout=8000) as new_page_info:
                        btn.click()
                    new_pg = new_page_info.value
                    new_pg.wait_for_load_state("domcontentloaded", timeout=15000)
                    dest = new_pg.url
                    new_pg.close()
            except Exception as e:
                print(f"      [Indeed] apply redirect navigation failed: {str(e).splitlines()[0][:100]}")
                continue
            outcome = _classify_apply_dest(dest)
            if outcome != "blocked":
                break

        if outcome == "resolved":
            print(f"      [Indeed] Resolved apply URL: {dest}")
            return dest
        if outcome == "login_required":
            print("      [Indeed] Apply URL needs a logged-in Indeed session — keeping listing URL "
                  "(use 'Re-login Indeed')")
        elif outcome == "blocked":
            print("      [Indeed] Apply redirect still challenged by Cloudflare after retry — keeping listing URL")
        else:
            print(f"      [Indeed] Could not resolve apply URL ({outcome}) — keeping listing URL")
        return None
    finally:
        browser.close()


# ── MAIN SCRAPE ────────────────────────────────────────────────────────────────

async def scrape_indeed(start_url: str, config: dict) -> List[dict]:
    candidates: List[dict] = []
    seen_jks: set           = set()
    seen_fingerprints: set  = set()

    preferred_locs  = config.get("preferred_locations",   ["remote", "miami"])
    seniority_tiers = config.get("seniority_tiers",       _DEFAULT_SENIORITY_TIERS)
    title_keywords  = [k.lower() for k in config.get("title_keywords",       _DEFAULT_TITLE_KEYWORDS)]
    excl_titles     = [k.lower() for k in config.get("excluded_titles",      list(_DEFAULT_EXCLUDED_TITLES))]
    excl_words      = [k.lower() for k in config.get("excluded_title_words", list(_DEFAULT_EXCLUDED_TITLE_WORDS))]
    max_candidates  = int(config.get("max_candidates") or MAX_CANDIDATES)

    log("[Pipeline Settings — applied to all scrapers]")
    log(f"  Preferred locations : {preferred_locs}")
    log(f"  Seniority tiers     : {seniority_tiers}")
    log(f"  Title keywords      : {title_keywords}")
    log("")
    log(f"[Indeed] Starting — target: {max_candidates} candidates")
    log(f"[Indeed] URL: {start_url}\n")

    applied_pairs, applied_urls = load_applied_jobs()

    kept = not_pm = excluded = bad_location = already_app = dup = 0

    if SESSION_FILE.exists():
        log("[Indeed] Reusing saved Indeed session...")
    else:
        log("[Indeed] No saved session — Indeed caps anonymous visitors to page 1 of "
            "results and gates 'Apply on company site' behind its own sign-in wall. "
            "Use the 'Re-login Indeed' button to fix both.")

    async with async_playwright() as p:
        browser, ua = await launch_browser(p, headless=True)
        context = await browser.new_context(**context_kwargs(ua))

        # Same mitigation as builtin_scraper.py for the identical Cloudflare
        # challenge signature: every subresource counts against the limit,
        # and a results page pulls a long tail of images/fonts/tracking
        # pixels we never read.
        async def _block_heavy_assets(route):
            if route.request.resource_type in ("image", "font", "media"):
                await route.abort()
            else:
                await route.continue_()

        await context.route("**/*", _block_heavy_assets)

        page = await context.new_page()

        current_url = start_url
        page_num    = 1
        visited_detail = 0

        while page_num <= MAX_PAGES and len(candidates) < max_candidates:
            log(f"[Indeed] ── Page {page_num} ──────────────────────────────────────")
            log(f"[Indeed] {current_url}")

            try:
                await page.goto(current_url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
                await page.wait_for_timeout(2500)
            except Exception as e:
                log(f"[Indeed] Failed to load page: {e}")
                break

            # Anonymous visitors get hard-capped to page 1 — checked live,
            # any attempt to go further bounces to secure.indeed.com/auth
            # with "To see more than one page of jobs, create an account or
            # sign in." Report this honestly instead of quietly logging
            # "0 real results" / "no next page", which reads as "exhausted
            # all available jobs" when the true reason is a login wall.
            if "secure.indeed.com" in page.url or "onboarding.indeed.com" in page.url:
                log(
                    "[Indeed] Hit Indeed's login wall — without a session, only page 1 "
                    "is visible. Use 'Re-login Indeed' to see further pages. Stopping here."
                )
                break

            page_jobs = await extract_jobs_from_page(page)
            log(f"  {len(page_jobs)} real result(s) on this page")

            kept_this_page = 0
            for pj in page_jobs:
                if len(candidates) >= max_candidates:
                    break

                title, jk = pj["title"], pj["jk"]

                if jk in seen_jks:
                    dup += 1
                    continue

                if is_excluded_title(title, excl_titles, excl_words):
                    excluded += 1
                    continue
                if not is_matching_title(title, title_keywords):
                    not_pm += 1
                    continue

                # Pre-check against the sheet before spending a click+apply-button
                # interaction — same shape as Built-in's numeric-ID pre-check.
                if f"indeed-jk:{jk}" in applied_urls:
                    already_app += 1
                    continue

                if pj["location"] and not is_valid_location(pj["location"], preferred_locs):
                    bad_location += 1
                    continue

                seen_jks.add(jk)

                if visited_detail:
                    await asyncio.sleep(random.uniform(*DETAIL_PAUSE_RANGE))
                visited_detail += 1

                detail = await _select_and_scrape_detail(page, jk)
                listing_url = detail["listing_url"] or current_url

                # url == listing URL for now; resolve_apply_url swaps in the
                # employer's real URL later, only if this job scores in.
                # _use_scraped_description: the JD was just read off the live
                # detail pane, so resume_pipeline must not re-fetch it from
                # indeed.com — that's another Cloudflare-scored hit per job,
                # from a browser without this session's cookies.
                job = {
                    "title": title,
                    "company": pj["company"],
                    "location": pj["location"] or "Remote",
                    "salary": pj.get("salary", ""),
                    "description": detail["description"],
                    "url": listing_url,
                    "linkedin_url": listing_url,
                    "easy_apply": detail["easy_apply"],
                    "_apply_outcome": detail["apply_outcome"],
                    "_apply_href": detail["apply_href"],
                    "_use_scraped_description": bool(detail["description"]),
                }

                if not is_valid_location(job["location"], preferred_locs):
                    bad_location += 1
                    continue

                fp = (job["company"].lower().strip(), title.lower().strip())
                if fp[1] and fp in seen_fingerprints:
                    dup += 1
                    continue
                if fp[1]:
                    seen_fingerprints.add(fp)

                if already_applied(job, applied_pairs, applied_urls):
                    already_app += 1
                    continue

                score, matched_tier = seniority_score(title, seniority_tiers)
                job["score"]  = score
                job["reason"] = f"Seniority: {matched_tier}" if matched_tier else "No seniority tier matched"

                kept += 1
                kept_this_page += 1
                candidates.append(job)
                log(f"  ✓ Candidate #{len(candidates)} found: {title} @ {job['company']}")

            log(
                f"\n  Kept {kept} | Not PM: {not_pm} | Excluded: {excluded} "
                f"| Bad location: {bad_location} | Already applied: {already_app} | Dup: {dup}"
            )
            log(f"  Candidates so far: {len(candidates)} / {max_candidates}")

            if len(candidates) >= max_candidates:
                log(f"\n[Indeed] Reached {max_candidates} candidates — done scraping.")
                break

            next_url = await find_next_url(page)
            if next_url and next_url != current_url:
                current_url = next_url
                page_num   += 1
                await asyncio.sleep(1.5)
            else:
                log("\n[Indeed] No next page — done.")
                break

        await browser.close()

    candidates.sort(key=lambda j: j["score"], reverse=True)

    log(f"\n[Indeed] {'═' * 52}")
    log(f"[Indeed] Candidates collected: {len(candidates)} / {max_candidates}")

    if candidates:
        log("\n[Indeed] Candidate list (sorted by seniority score):")
        for i, job in enumerate(candidates, 1):
            outcome = job.get("_apply_outcome")
            if outcome == "easy_apply":
                apply_kind = "Apply with Indeed (native) — listing URL"
            elif outcome == "deferred":
                apply_kind = "Apply on company site — resolved after scoring, only if it clears the threshold"
            else:
                apply_kind = "No apply button found — using listing URL"
            log(f"\n  {i}. {job['title']}  [{job['score']}/10]")
            log(f"     Company:    {job.get('company') or '?'}")
            log(f"     Location:   {job.get('location') or '?'}")
            if job.get("salary"):
                log(f"     Salary:     {job['salary']}")
            log(f"     Apply:      {apply_kind}")
            log(f"     Listing:    {job.get('linkedin_url', '')}")
    else:
        log("[Indeed] No candidates found. Check your search URL or loosen the filters.")

    return candidates


# ── ENTRY POINT ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    _config = load_config()
    _url = _config.get("indeed_url", "").strip()
    if not _url:
        log("[Indeed] ERROR: No Indeed URL set. Go to Settings and enter your search URL.")
        sys.exit(1)

    # Run async scraping first — event loop is fully closed before resume_pipeline
    # (which uses sync Playwright) is called. Same ordering as builtin_scraper.py.
    _candidates = asyncio.run(scrape_indeed(_url, _config))

    if not _candidates:
        log("\n[Indeed] No candidates to score — exiting.")
        sys.exit(0)

    # No field remap needed here (unlike dynamite_scraper.py / builtin_scraper.py):
    # url/linkedin_url are both the listing URL at this point; resolve_apply_url
    # upgrades url to the employer's real link for jobs that clear the threshold.
    log("\n[Indeed] Handing off to resume pipeline...\n")
    from resume_pipeline import run_pipeline
    run_pipeline(_candidates, resolve_apply_url=resolve_apply_url)
