#!/usr/bin/env python3
"""
Manual job upload — analyze a single job posting URL you found yourself (or
a friend sent you) and run it through the exact same scoring, threshold
check, resume building, and sheet logging as the automated scrapers.

Unlike the scrapers, there's no search results page to pull title/company
from, so this does one page load first to extract them, then hands off to
resume_pipeline.run_pipeline() exactly like builtin_scraper.py /
dynamite_scraper.py do at their own entry points — same pipeline, same
Applications/Skips outcome, just a list of one job instead of many.

USAGE:
  python3 manual_scraper.py <job posting URL>
"""

import os
import re
import sys
from urllib.parse import urlparse

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv()


def log(msg: str):
    print(msg, flush=True)


# (hostname_regex, path_segment_index) for ATS platforms that put the
# employer's slug in a fixed, known position — either the URL path or, for
# Workday, the hostname's subdomain. Checked live: neither Ashby nor
# Greenhouse sets og:site_name at all, and og:title on both is just the
# clean role title with no company — so for these platforms the URL is
# actually the most reliable signal available, more reliable than any meta
# tag or <title> text heuristic.
_ATS_COMPANY_FROM_PATH = [
    (re.compile(r"^jobs\.ashbyhq\.com$"), 0),
    (re.compile(r"^(job-boards|boards)\.greenhouse\.io$"), 0),
    (re.compile(r"^jobs\.lever\.co$"), 0),
    (re.compile(r"^apply\.workable\.com$"), 0),
]
_WORKDAY_HOST_RE = re.compile(r"^([a-z0-9-]+)\.wd\d+\.myworkdayjobs\.com$", re.I)


def _clean_title(raw_title: str) -> str:
    """Job posting <title> tags are almost universally "Job Title - Company"
    or "Job Title | Company" (Greenhouse, Lever, Ashby, Workday, and plain
    company career pages all follow this convention) — take the first
    segment, which is the role title."""
    if not raw_title:
        return ""
    parts = [p.strip() for p in re.split(r"\s*[|–—\-·:@]\s*", raw_title.strip()) if p.strip()]
    return parts[0] if parts else raw_title.strip()


def _company_from_url(url: str) -> str:
    """Employer slug from the URL itself, for ATS platforms where that slug
    is in a fixed, known position. Humanizing a slug ("form-health" ->
    "Form Health") only works when the slug has separators — a concatenated
    slug ("formhealth") comes back as one capitalized word, which is an
    imperfect but reasonable label (this only feeds a sheet column for your
    own tracking, not anything that drives a filtering decision)."""
    parsed = urlparse(url)
    host = (parsed.hostname or "").lower()

    m = _WORKDAY_HOST_RE.match(host)
    if m:
        return m.group(1).replace("-", " ").title()

    for host_re, idx in _ATS_COMPANY_FROM_PATH:
        if host_re.match(host):
            segments = [s for s in parsed.path.split("/") if s]
            if len(segments) > idx:
                return segments[idx].replace("-", " ").title()

    return ""


def _looks_like_company(segment: str, company: str) -> bool:
    """Case-insensitive fuzzy match — handles a URL-slug-derived company
    name ("Trustarc") differing only in internal capitalization from the
    real one as it appears in page text ("TrustArc")."""
    s, c = segment.strip().lower(), company.strip().lower()
    return bool(s and c and (s == c or s in c or c in s))


def _extract_title_and_company(page, url: str) -> tuple:
    """Pulls a job title and company name from a job posting page.

    Company is resolved FIRST, in priority order: og:site_name (rare, but
    most reliable when present) -> a "Role at Company" pattern in
    og:title/<title> -> a known ATS URL slug (_company_from_url).

    Title then uses whichever company was found to correctly split a
    two-part title regardless of which side the company is on — checked
    live, there's no single positional convention: Lever's <title> is
    "Company - Role" (company first), while Ashby's/Greenhouse's are
    "Role @ Company" / "Job Application for Role at Company" (role
    first). Falls through to "first segment" only when no company signal
    is available to disambiguate.
    """
    raw_title = (page.title() or "").strip()

    try:
        og_title = (page.eval_on_selector(
            "meta[property='og:title']", "el => el.content"
        ) or "").strip()
    except Exception:
        og_title = ""

    company = ""
    try:
        company = (page.eval_on_selector(
            "meta[property='og:site_name']", "el => el.content"
        ) or "").strip()
    except Exception:
        pass

    if not company:
        for text in (og_title, raw_title):
            m = re.search(r"\bat\s+(.+)$", text, re.I)
            if m:
                company = m.group(1).strip()
                break

    if not company:
        company = _company_from_url(url)

    title_source = og_title or raw_title
    segments = [p.strip() for p in re.split(r"\s*[|–—\-·:@]\s*", title_source) if p.strip()]

    title = ""
    if company and len(segments) >= 2:
        non_company = [s for s in segments if not _looks_like_company(s, company)]
        # Only trust the split when removing company-matching segment(s)
        # leaves exactly one segment behind — if 0 or 2+ remain, the company
        # match was ambiguous (matched nothing, or matched more than one
        # segment) and it's safer to fall through to the regex/first-segment
        # heuristics below than guess wrong.
        if len(non_company) == len(segments) - 1:
            title = non_company[0] if len(non_company) == 1 else " - ".join(non_company)

    if not title:
        # "Role at Company" pattern, checked against title_source itself
        # only — not a second, independent attempt on raw_title once
        # og_title already gave an unambiguous (if unsplit) answer. That
        # would otherwise re-match Greenhouse's noisier raw <title>
        # ("Job Application for Role at Company") even when the clean
        # og:title ("Role") was already sitting right there.
        m = re.search(r"^(.*?)\s+\bat\s+.+$", title_source, re.I)
        if m:
            title = m.group(1).strip()

    if not title:
        title = segments[0] if segments else title_source

    if not company:
        if len(segments) >= 2:
            company = segments[-1]
        else:
            host = urlparse(url).hostname or ""
            host = re.sub(r"^(www\.|jobs\.|careers\.|apply\.)", "", host)
            company = host.split(".")[0].replace("-", " ").title()

    return title, company


def _extract_location(page) -> str:
    """Same work-model keyword scan used by builtin_scraper.py — scans
    visible short lines of text for a remote/hybrid/on-site mention."""
    try:
        return page.evaluate("""() => {
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
        return ""


if __name__ == "__main__":
    if len(sys.argv) < 2 or not sys.argv[1].strip():
        log("[Manual] ERROR: No URL provided.")
        sys.exit(1)

    url = sys.argv[1].strip()
    log(f"[Manual] Analyzing: {url}")

    from resume_pipeline import SESSION_FILE

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-blink-features=AutomationControlled"],
        )
        context_options = {
            "viewport": {"width": 1920, "height": 1080},
            "user_agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        }
        if os.path.exists(SESSION_FILE):
            context = browser.new_context(storage_state=SESSION_FILE, **context_options)
        else:
            context = browser.new_context(**context_options)
        page = context.new_page()

        try:
            page.goto(url, wait_until="domcontentloaded", timeout=20000)
            page.wait_for_timeout(1500)
        except Exception as e:
            log(f"[Manual] ERROR: Could not load page: {e}")
            browser.close()
            sys.exit(1)

        title, company = _extract_title_and_company(page, url)
        location = _extract_location(page)
        browser.close()

    if not title:
        log("[Manual] ERROR: Could not determine a job title from the page — aborting.")
        sys.exit(1)

    log(f"[Manual] Title:    {title}")
    log(f"[Manual] Company:  {company or '(could not determine)'}")
    log(f"[Manual] Location: {location or '(unknown)'}")

    # url/linkedin_url both get the same value: the other scrapers use url
    # for the resolved apply URL and linkedin_url for the listing URL as a
    # stable dedup key (see resume_pipeline.is_already_logged), but a manual
    # entry only ever has the one URL the user gave, so it fills both roles.
    job = {
        "title": title,
        "company": company,
        "url": url,
        "linkedin_url": url,
        "location": location,
    }

    log("\n[Manual] Handing off to resume pipeline...\n")
    from resume_pipeline import run_pipeline
    run_pipeline([job])
