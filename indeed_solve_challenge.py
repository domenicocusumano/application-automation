"""
Optional helper: opens a VISIBLE browser using the saved Indeed session,
navigates to your configured search URL, and waits for you to clear any
Cloudflare challenge by hand (checkbox / "Verify you are human"). Once the
real results render, it also opens the first job and clicks "Apply on
company site" so you can clear the same challenge on the /applystart
redirect if it shows up there too. Saves storage_state back to
indeed_session.json at the end either way, so headless runs pick up
whatever clearance cookies got issued.

Rarely needed since the scraper moved to real Chrome with a genuine UA and
defers apply-URL resolution until after scoring (see indeed_scraper.py's
docstring) — that removed the traffic that was triggering the challenges.
Only useful when a challenge WITH a widget actually appears. If the page
says "Additional Verification Required" with only a "Return home" link,
that's an IP-level block with nothing to solve; wait a few hours instead.

The browser here must match indeed_scraper.py's exactly (real Chrome,
genuine UA): Cloudflare binds cf_clearance to the fingerprint that earned
it. This is precisely why the previous version of this script looped —
its bundled-Chromium-148-claiming-Chrome-131 fingerprint failed the
post-solve integrity check every time.

Run manually: python3 indeed_solve_challenge.py
"""
import json
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright

SCRIPT_DIR   = Path(__file__).parent
SESSION_FILE = SCRIPT_DIR / "indeed_session.json"
CONFIG_FILE  = SCRIPT_DIR / "config.json"
BROWSER_ARGS = ["--no-sandbox", "--disable-blink-features=AutomationControlled"]


def log(msg):
    print(msg, flush=True)


def is_challenge(page) -> bool:
    try:
        title = page.title()
    except Exception:
        return False
    return "just a moment" in title.lower() or "verification required" in (page.content() or "").lower()


def wait_until_clear(page, label: str, timeout_s: int = 180):
    log(f"  If a Cloudflare check appears for {label}, solve it in the window now "
        f"(waiting up to {timeout_s}s)...")
    for _ in range(timeout_s):
        page.wait_for_timeout(1000)
        if not is_challenge(page):
            log(f"  {label}: clear.")
            return True
    log(f"  {label}: still challenged after {timeout_s}s — giving up on this step.")
    return False


def main():
    cfg = json.loads(CONFIG_FILE.read_text()) if CONFIG_FILE.exists() else {}
    search_url = cfg.get("indeed_url", "").strip() or "https://www.indeed.com/jobs?q=product+manager&l=Remote"

    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(channel="chrome", headless=False, args=BROWSER_ARGS)
        except Exception:
            log("Google Chrome not found — falling back to Playwright's bundled Chromium.")
            browser = p.chromium.launch(headless=False, args=BROWSER_ARGS)
        context_kwargs = dict(
            viewport={"width": 1280, "height": 900},
            locale="en-US",
        )
        if SESSION_FILE.exists():
            context_kwargs["storage_state"] = str(SESSION_FILE)
        context = browser.new_context(**context_kwargs)
        page = context.new_page()

        log(f"Opening search: {search_url}")
        page.goto(search_url, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(1500)
        wait_until_clear(page, "search page")

        # Try to click into the first real job and hit "Apply on company site"
        # so the /applystart redirect (the thing actually rate-limiting the
        # scraper) gets a chance to be solved too.
        try:
            link = page.query_selector("a[data-jk]")
            if link and link.is_visible():
                link.click()
                page.wait_for_timeout(1800)
                apply_el = (
                    page.query_selector("[data-testid='viewjob-apply']")
                    or page.query_selector("button:has-text('Apply on company site')")
                )
                if apply_el and apply_el.is_visible():
                    log("Clicking 'Apply on company site' on the first job...")
                    with context.expect_page(timeout=8000) as new_page_info:
                        apply_el.click()
                    new_page = new_page_info.value
                    new_page.wait_for_load_state("domcontentloaded", timeout=15000)
                    wait_until_clear(new_page, "apply redirect")
                    new_page.close()
                else:
                    log("No 'Apply on company site' button found on the first job — skipping that step.")
            else:
                log("No job cards found to test the apply flow — skipping that step.")
        except Exception as e:
            log(f"Apply-flow check skipped ({e}).")

        log("Saving session...")
        context.storage_state(path=str(SESSION_FILE))
        browser.close()
        log("Done. Session saved — try the pipeline again.")


if __name__ == "__main__":
    main()
