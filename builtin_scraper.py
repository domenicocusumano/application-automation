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
from pathlib import Path
from typing import Optional, List
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

from dotenv import load_dotenv
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

load_dotenv()

CONFIG_FILE    = Path(__file__).parent / "config.json"
MAX_PAGES      = 20
NAV_TIMEOUT    = 30_000
DETAIL_TIMEOUT = 20_000
MAX_CANDIDATES = 10

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


# ── DETAIL PAGE ────────────────────────────────────────────────────────────────

async def scrape_job_detail(page, job_url: str) -> dict:
    """
    Fetch company, location, job description, Easy Apply status, and actual apply URL
    from a Built-in job detail page.
    """
    result = {
        "company": "", "location": "", "description": "",
        "easy_apply": False, "apply_url": job_url,
    }
    try:
        await page.goto(job_url, wait_until="domcontentloaded", timeout=DETAIL_TIMEOUT)
        await page.wait_for_timeout(1000)

        # Breadcrumb first-link is the most reliable company source on BuiltIn
        for sel in [
            "nav[aria-label*='breadcrumb'] a:first-child",
            "[data-testid='breadcrumb'] a:first-child",
            "[class*='breadcrumb'] a:first-child",
            "[data-testid='employer-name']", "[data-testid='company-title']",
            "[data-testid='company-name']",
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

        # JS fallback: breadcrumb first link, then parse "at [Company]" from page title
        if not result["company"]:
            try:
                result["company"] = await page.evaluate("""() => {
                    const bc = document.querySelector(
                        '[class*="breadcrumb"], nav[aria-label*="breadcrumb"]'
                    );
                    if (bc) {
                        const link = bc.querySelector('a');
                        if (link) {
                            const t = (link.innerText || '').trim();
                            if (t && t.length < 120) return t;
                        }
                    }
                    // Fallback: parse "Job Title at Company | Built In" from <title>
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
                    await page.wait_for_timeout(1500)
                    break
        except Exception:
            pass

        for sel in [
            "[data-testid='job-description']", ".job-description",
            "#job-description", "[class*='JobDescription']",
            "[class*='job-description']",
            # BuiltIn "The Role" section containers
            "[class*='job-details']", "[class*='JobDetails']",
            "section[class*='job']", "[class*='the-role']", "[class*='TheRole']",
            "main article", "main", "article",
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
        for sel in [
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

        if not result["easy_apply"]:
            for sticky_sel in ["[class*='sticky']", "[class*='fixed']", "footer"]:
                try:
                    el = await page.query_selector(sticky_sel)
                    if el and "easy apply" in (await el.inner_text()).lower():
                        result["easy_apply"] = True
                        break
                except Exception:
                    pass

        # ── External apply URL ──
        # If Easy Apply, the apply URL stays as the Built-in listing URL.
        # Otherwise, try to find the external company apply link.
        if not result["easy_apply"]:
            for sel in [
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
                        href = (await el.get_attribute("href") or "").strip()
                        if href and href.startswith("http") and "builtin.com" not in href.lower():
                            result["apply_url"] = href
                            break
                except Exception:
                    pass

            # JS fallback: any <a> whose text is exactly "Apply" / "Apply Now"
            # with an external href (catches sticky footer buttons missed by CSS selectors)
            if result["apply_url"] == job_url:
                try:
                    href = await page.evaluate("""() => {
                        for (const a of document.querySelectorAll('a[href]')) {
                            const text = (a.innerText || a.textContent || '').trim().toUpperCase();
                            const href = a.getAttribute('href') || '';
                            if ((text === 'APPLY' || text === 'APPLY NOW') &&
                                href.startsWith('http') &&
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
                        if btn and await btn.is_visible():
                            async with page.context.expect_page(timeout=10000) as new_page_info:
                                await btn.click()
                            new_tab = await new_page_info.value
                            await new_tab.wait_for_load_state("domcontentloaded", timeout=10000)
                            external_url = new_tab.url
                            await new_tab.close()
                            if external_url and "builtin.com" not in external_url.lower():
                                result["apply_url"] = external_url
                                break
                except Exception:
                    pass

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
    seen_fingerprints: set = set()  # (company.lower(), title.lower()) — within-run dedup

    preferred_locs  = config.get("preferred_locations",   ["remote", "miami"])
    seniority_tiers = config.get("seniority_tiers",       _DEFAULT_SENIORITY_TIERS)
    title_keywords  = [k.lower() for k in config.get("title_keywords",       _DEFAULT_TITLE_KEYWORDS)]
    excl_titles     = [k.lower() for k in config.get("excluded_titles",      list(_DEFAULT_EXCLUDED_TITLES))]
    excl_words      = [k.lower() for k in config.get("excluded_title_words", list(_DEFAULT_EXCLUDED_TITLE_WORDS))]

    log(f"[Pipeline Settings — applied to all scrapers]")
    log(f"  Preferred locations : {preferred_locs}")
    log(f"  Seniority tiers     : {seniority_tiers}")
    log(f"  Title keywords      : {title_keywords}")
    log(f"  Scoring             : programmatic (no Claude)")
    log(f"  Already-applied     : checked against Applications + Skips tabs\n")
    log(f"[Built-in] Starting — target: {MAX_CANDIDATES} candidates")
    log(f"[Built-in] URL: {start_url}\n")

    # Load already-applied jobs before starting the browser
    applied_pairs, applied_urls = load_applied_jobs()

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1280, "height": 900},
        )
        list_page   = await context.new_page()
        detail_page = await context.new_page()

        current_url = start_url
        page_num    = 1

        while page_num <= MAX_PAGES and len(candidates) < MAX_CANDIDATES:
            log(f"[Built-in] ── Page {page_num} ──────────────────────────────────────")
            log(f"[Built-in] {current_url}")

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

            for job in page_jobs:
                if len(candidates) >= MAX_CANDIDATES:
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

                # Title (and list-page location) passed — visit detail page
                detail = await scrape_job_detail(detail_page, url)
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
            )
            log(f"  Candidates so far: {len(candidates)} / {MAX_CANDIDATES}")

            if len(candidates) >= MAX_CANDIDATES:
                log(f"\n[Built-in] Reached {MAX_CANDIDATES} candidates — done scraping.")
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
    log(f"[Built-in] Candidates collected: {len(candidates)} / {MAX_CANDIDATES}")

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
