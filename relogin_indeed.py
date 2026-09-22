"""
Standalone re-login script. Run by the web UI via /relogin-indeed SSE endpoint.
Opens a visible Chromium window, waits for the user to log into Indeed,
then saves the session and exits.

Note: this only proves you're logged in — it does NOT by itself clear
Indeed's Cloudflare challenge on www.indeed.com/jobs or /applystart, since
it never navigates there (only secure.indeed.com's login flow). A valid
login can still see "0 real results" or apply-redirect rate-limiting from
Cloudflare; that's what indeed_solve_challenge.py (and indeed_scraper.py's
own automatic manual-solve fallback) are for.
"""
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright

SCRIPT_DIR   = Path(__file__).parent
SESSION_FILE = SCRIPT_DIR / "indeed_session.json"
BROWSER_ARGS = ["--no-sandbox", "--disable-blink-features=AutomationControlled"]

# Still mid-login while the URL matches any of these — secure.indeed.com hosts
# both the sign-in form and the post-signup onboarding flow.
LOGIN_MARKERS = ["secure.indeed.com", "onboarding.indeed.com", "accounts.google.com"]

def log(msg):
    print(msg, flush=True)

def main():
    log("Opening browser — log into Indeed in the window that appears...")
    with sync_playwright() as p:
        # Must match indeed_scraper.py's browser exactly (real Chrome, genuine
        # UA) — Cloudflare binds cf_clearance to the fingerprint that earned
        # it, so a session saved from a different build/UA is re-challenged
        # the moment the scraper loads it. See indeed_scraper's docstring for
        # why the UA is never overridden.
        try:
            browser = p.chromium.launch(channel="chrome", headless=False, args=BROWSER_ARGS)
        except Exception:
            log("Google Chrome not found — falling back to Playwright's bundled Chromium.")
            browser = p.chromium.launch(headless=False, args=BROWSER_ARGS)
        context = browser.new_context(
            viewport={"width": 1280, "height": 900},
            locale="en-US",
        )
        page = context.new_page()
        page.goto("https://secure.indeed.com/account/login", wait_until="domcontentloaded")

        log("Waiting for you to log in...")
        try:
            page.wait_for_url(
                lambda url: "indeed.com" in url and not any(x in url for x in LOGIN_MARKERS),
                timeout=300_000,  # 5 minutes
            )
        except Exception:
            log("ERROR: Timed out — please try again.")
            browser.close()
            sys.exit(1)

        log("Logged in — saving session...")
        context.storage_state(path=str(SESSION_FILE))
        browser.close()
        log("Session saved. You can close this and run the pipeline.")
        sys.exit(0)

if __name__ == "__main__":
    main()
