"""
Standalone re-login script. Run by the web UI via /relogin SSE endpoint.
Opens a visible Chromium window, waits for the user to log into LinkedIn,
then saves the session and exits.
"""
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright

SCRIPT_DIR   = Path(__file__).parent
SESSION_FILE = SCRIPT_DIR / "linkedin_session.json"

LOGIN_PAGES = ["/login", "/signup", "/uas/", "/authwall", "/checkpoint", "accounts.google.com"]

def log(msg):
    print(msg, flush=True)

def main():
    log("Opening browser — log into LinkedIn in the window that appears...")
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=["--no-sandbox", "--disable-blink-features=AutomationControlled"]
        )
        context = browser.new_context(
            viewport={"width": 1280, "height": 900},
            locale="en-US",
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            extra_http_headers={"Accept-Language": "en-US,en;q=0.9"},
        )
        page = context.new_page()
        page.goto("https://www.linkedin.com/login", wait_until="domcontentloaded")

        log("Waiting for you to log in...")
        try:
            page.wait_for_url(
                lambda url: "linkedin.com" in url and not any(x in url for x in LOGIN_PAGES),
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
