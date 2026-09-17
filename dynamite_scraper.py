#!/usr/bin/env python3
"""
Dynamite Jobs (dynamitejobs.com) scraper.

Mirrors builtin_scraper.py's candidate-collection style, then hands the
collected candidates to resume_pipeline for AI scoring, resume building,
and Google Sheet logging — same pipeline every other scraper uses.

ARCHITECTURE — why this doesn't use Playwright at all:
Dynamite's own search box calls Algolia directly from the browser, using a
public search-only API key (Algolia search keys are *meant* to be exposed
client-side — access is scoped server-side by Algolia's ACL to read-only
search on this one index; this isn't a secret we're extracting). Calling
that same endpoint ourselves — verified with a bare `requests.post`, no
cookies, no browser — returns the identical result count as a real browser
session, and hands back everything we need per job in ONE response: title,
company, the full plain-text description, applyType/applyLink, salary, and
locationSlugs. That also sidesteps the login wall on /my-jobs entirely
(confirmed: that path 302s to a signup page for an anonymous request, but
the search API behind it is public regardless) — no session file, no
relogin flow needed, unlike LinkedIn/Built-in.

Per-job apply URL: `applyType` is "link" (→ `applyLink`, the real external
ATS URL — Ashby/Greenhouse/Lever/Workable, all seen live) or "platform-form"
(the employer hasn't listed elsewhere, so Dynamite hosts its own lead-
capture application at https://dynamitejobs.com/apply/<objectID> — matches
exactly what a real "Apply Now" click opens for those jobs).
"""

import json
import os
import re
import sys
import time
from pathlib import Path
from typing import Optional, List
from urllib.parse import urlparse, parse_qs

import requests
from dotenv import load_dotenv

load_dotenv()

SCRIPT_DIR  = Path(__file__).parent
CONFIG_FILE = SCRIPT_DIR / "config.json"

BASE_URL = "https://dynamitejobs.com"

ALGOLIA_APP_ID  = "49HKL9G3SB"
ALGOLIA_API_KEY = "578864e6f2d8bc38a05a8f3302d5a9ac"
ALGOLIA_URL     = (
    f"https://{ALGOLIA_APP_ID.lower()}-dsn.algolia.net/1/indexes/*/queries"
    f"?x-algolia-agent=Algolia+for+JavaScript&x-algolia-api-key={ALGOLIA_API_KEY}"
    f"&x-algolia-application-id={ALGOLIA_APP_ID}"
)
HITS_PER_PAGE    = 16  # matches the site's own page size
MAX_PAGES        = 20
MAX_CANDIDATES   = 10
REQUEST_TIMEOUT  = 15
REQUEST_PAUSE    = 0.5  # light pacing between Algolia calls — good citizenship

GOOGLE_CREDS_FILE = Path(__file__).parent / os.getenv("GOOGLE_CREDS_FILE", "google_credentials.json")

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
    if "anywhere" in loc or "worldwide" in loc:
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


# ── HELPERS ──────────────────────────────────────────────────────────────────

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
                        norm = normalize_url(row.get(col, ""))
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
    """Returns True if this job is already in the applied/skips sheet."""
    for field in ["apply_url", "url"]:
        norm = normalize_url(job.get(field, ""))
        if norm and norm in applied_urls:
            return True

    company = job.get("company", "").lower().strip()
    title   = job.get("title", "").lower().strip()

    if (company, title) in applied_pairs:
        return True

    for (ac, at) in applied_pairs:
        if ac and company and (ac in company or company in ac):
            t_words  = set(title.split())
            at_words = set(at.split())
            if len(t_words & at_words) >= 2:
                return True
    return False


# ── ALGOLIA SEARCH ─────────────────────────────────────────────────────────────

_LOCATION_LABELS = {
    "us": "United States", "ca": "Canada", "gb": "United Kingdom", "pt": "Portugal",
    "ro": "Romania", "mx": "Mexico", "br": "Brazil", "in": "India", "au": "Australia",
    "de": "Germany", "fr": "France", "es": "Spain", "nl": "Netherlands", "ie": "Ireland",
    "europe": "Europe", "latinamerica": "Latin America", "northamerica": "North America",
    "southamerica": "South America", "westernasia": "Middle East", "asia": "Asia",
    "africa": "Africa", "oceania": "Oceania", "worldwide": "Worldwide",
    "remote": "Remote", "anywhere": "Anywhere",
}


def _format_location(slugs) -> str:
    """Builds a readable location string from Dynamite's locationSlugs array.
    Coarse — good enough for the pre-filter and the phase-1 summary printout;
    precise eligibility (e.g. "must be UTC+3 to UTC-5") lives in the JD text
    itself and is a job for the AI scoring stage added in a later phase."""
    if not slugs:
        return "Remote"
    seen: set = set()
    ordered = []
    for s in slugs:
        key = str(s).strip().lower()
        label = _LOCATION_LABELS.get(key, str(s).strip().replace("-", " ").title())
        if label not in seen:
            seen.add(label)
            ordered.append(label)
    return "Remote — " + ", ".join(ordered)


def _extract_search_text(url: str) -> str:
    """Pulls the `text` query param out of whatever Dynamite search URL was
    configured (works for /my-jobs, /remote-jobs, or any other path — we
    never actually navigate to it, just read the search term back out)."""
    try:
        qs = parse_qs(urlparse(url).query)
        return (qs.get("text", [""])[0] or "").replace("+", " ").strip()
    except Exception:
        return ""


def _extract_category_filters(url: str) -> list:
    """Pulls the Category/Subcategory filter out of the URL PATH, e.g.
    /remote-jobs/product/product-management -> Algolia facetFilters
    ["categories.category.slug:product", "categories.subcategory.slug:product-management"].

    Dynamite's own "Category" dropdown doesn't add a query param — clicking
    it rewrites the path itself (verified live: selecting Product > Product
    Management sends the browser to exactly that URL and adds those two
    facetFilters to its own Algolia request). Reading it back out of the
    path means picking a category in your browser and pasting the resulting
    URL into Settings just works, the same way the `text=` search term does
    — no separate category field needed. Cuts a huge amount of noise: an
    unfiltered "senior product manager" search returns ~11k loosely-matched
    jobs; scoped to Product > Product Management it's ~1.2k, and every
    result on page 1 is an actual PM title.
    """
    try:
        path = urlparse(url).path
    except Exception:
        return []
    parts = [p for p in path.split("/") if p]
    if parts and parts[0] in ("remote-jobs", "my-jobs"):
        parts = parts[1:]
    filters = []
    if len(parts) >= 1 and parts[0]:
        filters.append(f"categories.category.slug:{parts[0]}")
    if len(parts) >= 2 and parts[1]:
        filters.append(f"categories.subcategory.slug:{parts[1]}")
    return filters


def _algolia_search(query: str, page: int, category_filters: Optional[list] = None) -> dict:
    facet_filters = [
        ["flags.isVisible:true", "flags.isFinished:true",
         "flags.isExpired:true", "flags.isFulfilled:true"],
        "flags.isBlocked:false",
    ]
    if category_filters:
        facet_filters.extend(category_filters)
    body = {
        "requests": [{
            "indexName": "prod_jobs",
            "query": query,
            "page": page,
            "hitsPerPage": HITS_PER_PAGE,
            "facetFilters": facet_filters,
            "optionalFilters": [],
            "disableExactOnAttributes": ["description"],
            "removeWordsIfNoResults": "lastWords",
        }]
    }
    resp = requests.post(ALGOLIA_URL, json=body, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    return resp.json()["results"][0]


def _job_from_hit(hit: dict) -> dict:
    company     = hit.get("company") or {}
    apply_type  = hit.get("applyType", "")
    apply_link  = hit.get("applyLink", "")
    doc_id      = hit.get("objectID", "")

    apply_url = apply_link if (apply_type == "link" and apply_link) else (
        f"{BASE_URL}/apply/{doc_id}" if doc_id else ""
    )

    company_username = company.get("username", "")
    slug = hit.get("slug", "")
    permalink = (
        f"{BASE_URL}/company/{company_username}/remote-job/{slug}"
        if company_username and slug else ""
    )

    salary = hit.get("salary") or {}
    salary_bits = []
    if salary.get("from"):
        salary_bits.append(f"${salary['from']:,}")
    if salary.get("to"):
        salary_bits.append(f"${salary['to']:,}")
    salary_str = "-".join(salary_bits) + (f"/{salary['type']}" if salary.get("type") else "") if salary_bits else ""

    description = (hit.get("description") or "")[:8000]
    # Dynamite's salary figures come from structured Algolia data, not from
    # the description prose itself — the JD text alone often never states a
    # number at all. score_job()'s salary_max_usd extraction only reads
    # whatever text we hand it, so without this the salary-disqualify check
    # (deliberately hardened to never fire on an unstated salary — see
    # resume_pipeline.py) would silently never fire for a *real* stated
    # Dynamite salary either. Label it by Algolia's own "public" flag:
    # `public: true` is the employer's own stated figure — presented plainly,
    # so it's eligible for disqualification like any other stated salary.
    # `public: false`/missing is Dynamite's internal estimate, not something
    # the employer published — labelled "(estimated)" so score_job's own
    # instruction ("do not guess/infer/estimate") correctly excludes it from
    # disqualification, exactly as an unstated salary would be.
    if salary_bits:
        label = "Salary" if salary.get("public") else "Salary (estimated, not stated by employer)"
        description = f"{label}: {salary_str}\n\n{description}"

    return {
        "title":       hit.get("title", ""),
        "company":     company.get("name", ""),
        "url":         permalink,   # Dynamite listing permalink -> "Linked In URL" column
        "apply_url":   apply_url,   # actual apply destination     -> "URL" column
        "location":    _format_location(hit.get("locationSlugs")),
        "description": description,
        "salary":      salary_str,
        "easy_apply":  False,
        "_apply_type": apply_type,
    }


# ── MAIN SCRAPE ────────────────────────────────────────────────────────────────

def scrape_dynamite(start_url: str, config: dict) -> List[dict]:
    candidates: List[dict] = []
    seen_urls: set          = set()
    seen_fingerprints: set  = set()

    preferred_locs  = config.get("preferred_locations",   ["remote", "miami"])
    seniority_tiers = config.get("seniority_tiers",       _DEFAULT_SENIORITY_TIERS)
    title_keywords  = [k.lower() for k in config.get("title_keywords",       _DEFAULT_TITLE_KEYWORDS)]
    excl_titles     = [k.lower() for k in config.get("excluded_titles",      list(_DEFAULT_EXCLUDED_TITLES))]
    excl_words      = [k.lower() for k in config.get("excluded_title_words", list(_DEFAULT_EXCLUDED_TITLE_WORDS))]
    max_candidates  = int(config.get("max_candidates") or MAX_CANDIDATES)

    query            = _extract_search_text(start_url)
    category_filters = _extract_category_filters(start_url)

    log("[Pipeline Settings — applied to all scrapers]")
    log(f"  Preferred locations : {preferred_locs}")
    log(f"  Seniority tiers     : {seniority_tiers}")
    log(f"  Title keywords      : {title_keywords}")
    log("")
    log(f"[Dynamite] Starting — target: {max_candidates} candidates")
    log(f"[Dynamite] Search URL: {start_url}")
    if query:
        log(f"[Dynamite] Search query: {query!r}")
    else:
        log("[Dynamite] WARNING: no 'text=' param found in that URL — searching with an empty query (all jobs).")
    if category_filters:
        log(f"[Dynamite] Category filter: {category_filters}")
    log("")

    applied_pairs, applied_urls = load_applied_jobs()

    kept = not_pm = excluded = bad_location = already_app = dup = 0

    for page_num in range(MAX_PAGES):
        log(f"\n[Dynamite] ── Page {page_num + 1} ──────────────────────────────────────")
        try:
            result = _algolia_search(query, page_num, category_filters)
        except Exception as e:
            log(f"[Dynamite] ERROR fetching page {page_num + 1}: {e}")
            break

        hits = result.get("hits", [])
        log(f"[Dynamite] {len(hits)} results (nbHits={result.get('nbHits')}, nbPages={result.get('nbPages')})")
        if not hits:
            log("[Dynamite] No more results — done.")
            break

        for hit in hits:
            if len(candidates) >= max_candidates:
                break

            title = hit.get("title", "")
            job = _job_from_hit(hit)

            if not job["url"] or job["url"] in seen_urls:
                dup += 1
                continue

            if is_excluded_title(title, excl_titles, excl_words):
                excluded += 1
                continue
            if not is_matching_title(title, title_keywords):
                not_pm += 1
                continue
            if not is_valid_location(job["location"], preferred_locs):
                bad_location += 1
                continue

            seen_urls.add(job["url"])

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
            candidates.append(job)
            log(f"  ✓ Candidate #{len(candidates)} found: {title} @ {job['company']}")

        log(
            f"\n  Kept {kept} | Not PM: {not_pm} | Excluded: {excluded} "
            f"| Bad location: {bad_location} | Already applied: {already_app} | Dup: {dup}"
        )
        log(f"  Candidates so far: {len(candidates)} / {max_candidates}")

        if len(candidates) >= max_candidates:
            log(f"\n[Dynamite] Reached {max_candidates} candidates — done scraping.")
            break

        if page_num + 1 >= result.get("nbPages", MAX_PAGES):
            log("\n[Dynamite] No more pages available — done.")
            break

        time.sleep(REQUEST_PAUSE)

    candidates.sort(key=lambda j: j["score"], reverse=True)

    log(f"\n[Dynamite] {'═' * 52}")
    log(f"[Dynamite] Candidates collected: {len(candidates)} / {max_candidates}")

    if candidates:
        log("\n[Dynamite] Candidate list (sorted by seniority score):")
        for i, job in enumerate(candidates, 1):
            apply_kind = "Dynamite-hosted form" if job.get("_apply_type") != "link" else "External ATS"
            log(f"\n  {i}. {job['title']}  [{job['score']}/10]")
            log(f"     Company:    {job.get('company') or '?'}")
            log(f"     Location:   {job.get('location') or '?'}")
            if job.get("salary"):
                log(f"     Salary:     {job['salary']}")
            log(f"     Apply:      {apply_kind}")
            log(f"     Apply URL:  {job.get('apply_url', '')}")
            log(f"     Listing:    {job.get('url', '')}")
    else:
        log("[Dynamite] No candidates found. Check your search URL or loosen the filters.")

    return candidates


# ── ENTRY POINT ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    _config = load_config()
    _url = _config.get("dynamite_url", "").strip()
    if not _url:
        log("[Dynamite] ERROR: No Dynamite Jobs URL set. Go to Settings and enter your search URL.")
        sys.exit(1)

    _candidates = scrape_dynamite(_url, _config)

    if not _candidates:
        log("\n[Dynamite] No candidates to score — exiting.")
        sys.exit(0)

    # Remap fields so resume_pipeline gets the right URLs in the right columns:
    #   url          -> actual apply URL   (written to "URL" column in the sheet)
    #   linkedin_url -> Dynamite listing permalink (written to "Linked In URL" column)
    # scrape_dynamite/_job_from_hit build candidates the other way around
    # (url = listing permalink, apply_url = real apply destination) because
    # that's the natural shape while scraping; this mirrors the identical
    # remap builtin_scraper.py does at its own hand-off.
    for _job in _candidates:
        _listing_url = _job.get("url", "")
        _apply_url   = _job.get("apply_url") or _listing_url
        _job["url"]          = _apply_url
        _job["linkedin_url"] = _listing_url

    log("\n[Dynamite] Handing off to resume pipeline...\n")
    from resume_pipeline import run_pipeline
    run_pipeline(_candidates)
