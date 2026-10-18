"""
Standalone Built-in phone-verification helper. Run by the web UI via the
/verify-builtin SSE endpoint (and auto-opened when the scraper reports that
Built-in now gates apply URLs behind phone verification).

Built-in added a phone-verification gate AFTER login (checked live 2026-09-30):
a logged-in but phone-unverified session's apply redirect lands on
"?verifyPhone=true" and the real employer apply URL is withheld until the
account completes phone verification. This is a one-time account action, not a
per-run one — once verified, every future apply redirect resolves.

Opens a VISIBLE browser with the saved session, lands on a real job's apply
flow so the "verify your phone" prompt appears, and waits for you to enter
your number + the SMS code. It detects success by re-checking the apply
redirect (which stops returning verifyPhone once the account is verified),
then saves the refreshed session.
"""
import json
import sys
import time
from pathlib import Path
from urllib.parse import urljoin, urlparse
from playwright.sync_api import sync_playwright

SCRIPT_DIR   = Path(__file__).parent
SESSION_FILE = SCRIPT_DIR / "builtin_session.json"
CONFIG_FILE  = SCRIPT_DIR / "config.json"
POLL_TIMEOUT = 300  # 5 minutes, same budget as relogin


def log(msg):
    print(msg, flush=True)


def _first_job_url(page, search_url: str) -> str:
    """Open the configured search and return the first real /job/ listing URL."""
    try:
        page.goto(search_url, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(2000)
        el = page.query_selector("a[href*='/job/']")
        if el:
            href = el.get_attribute("href") or ""
            return href if href.startswith("http") else urljoin("https://builtin.com", href)
    except Exception:
        pass
    return ""


def _redirect_still_gated(page, job_url: str) -> bool:
    """True while the apply redirect is still withheld behind the phone gate.
    Shares the live browser cookies, so it reflects verification completed in
    the same window moments earlier."""
    path = urlparse(job_url).path
    redirect = urljoin(job_url, f"{path}?handler=ApplyRedirect")
    try:
        resp = page.context.request.get(redirect, max_redirects=0, timeout=10000)
    except Exception:
        return True  # can't tell — treat as still gated, keep waiting
    loc = (resp.headers.get("location", "") or "").lower()
    # Gated while the redirect bounces to verifyPhone (or, if the session
    # somehow dropped, applyRequired). Anything else — an off-site 302, or no
    # gate param — means verification cleared it.
    return "verifyphone" in loc or "applyrequired" in loc


def main():
    if not SESSION_FILE.exists():
        log("ERROR: No Built-in session found. Click 'Re-login Built-in' first, then verify.")
        sys.exit(1)

    try:
        cfg = json.loads(CONFIG_FILE.read_text())
    except Exception:
        cfg = {}
    search_url = (cfg.get("builtin_url") or "").strip() or "https://builtin.com/jobs"

    log("Opening browser — complete Built-in phone verification in the window that appears...")
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=["--no-sandbox", "--disable-blink-features=AutomationControlled"],
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
        page = context.new_page()

        job_url = _first_job_url(page, search_url)
        if not job_url:
            log("Couldn't find a job to open — navigate to any Built-in job and click Apply "
                "to trigger the phone prompt. Waiting...")
        else:
            try:
                # Open the job with ?verifyPhone=true so Built-in surfaces the
                # phone-entry step directly. Also click Apply as a fallback
                # trigger — for a logged-in unverified account that pops the
                # "verify your phone" flow either way.
                sep = "&" if "?" in job_url else "?"
                page.goto(f"{job_url}{sep}verifyPhone=true", wait_until="domcontentloaded", timeout=30000)
                page.wait_for_timeout(1500)
                btn = page.query_selector("a[href*='handler=ApplyRedirect'], a:has-text('Apply'), button:has-text('Apply')")
                if btn:
                    try:
                        btn.click(timeout=4000)
                    except Exception:
                        pass
            except Exception:
                pass

        log("→ In the window: click Apply if needed, enter your phone number, then the")
        log("  SMS code Built-in texts you. This is a one-time account step.")
        log(f"Waiting up to {POLL_TIMEOUT // 60} minutes for verification to complete...")

        if not job_url:
            # No job to poll against — fall back to waiting for the user to
            # finish, then save whatever session state resulted.
            time.sleep(POLL_TIMEOUT)
            context.storage_state(path=str(SESSION_FILE))
            browser.close()
            log("Session saved. If apply URLs still don't resolve, re-run this after verifying.")
            sys.exit(0)

        deadline = time.time() + POLL_TIMEOUT
        verified = False
        while time.time() < deadline:
            time.sleep(4)
            if not _redirect_still_gated(page, job_url):
                verified = True
                break

        if verified:
            log("Phone verification detected — apply URLs now resolve. Saving session...")
            context.storage_state(path=str(SESSION_FILE))
            browser.close()
            # Machine-readable success marker the UI watches for to update the
            # Verify button (the session cookie check alone can't tell verified
            # from merely-logged-in).
            log("[VERIFIED]")
            log("Session saved. Re-run the pipeline and Built-in apply URLs will resolve.")
            sys.exit(0)
        else:
            log("Timed out waiting for verification. Saving session anyway — if you did verify, "
                "re-run the pipeline; if not, run this again.")
            context.storage_state(path=str(SESSION_FILE))
            browser.close()
            sys.exit(1)


if __name__ == "__main__":
    main()
