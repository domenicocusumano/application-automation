#!/usr/bin/env python3
"""
application_filler.py
Intelligent job application form filler using Playwright + Claude

Takes a list of job application URLs and tailored resumes, then automatically
fills out and submits each application form by having Claude read the page HTML
and map fields to values.

Usage:
    python3 application_filler.py                    # Auto mode (reads pipeline_output.json)
    python3 application_filler.py --manual           # Manual mode (uses MANUAL_JOBS list)
    python3 application_filler.py --dry-run          # Fill but don't submit
    python3 application_filler.py --no-pause         # Don't pause before submit
"""

import os
import re
import json
import time
from dotenv import load_dotenv

import anthropic
from playwright.sync_api import sync_playwright

load_dotenv()

# ── CONFIG ─────────────────────────────────────────────────────────────────────

ANTHROPIC_API_KEY  = os.getenv("ANTHROPIC_API_KEY", "YOUR_ANTHROPIC_API_KEY")

PIPELINE_OUTPUT    = "pipeline_output.json"   # Auto mode input file
HEADLESS           = False    # False = see the browser, True = run hidden
PAUSE_BEFORE_SUBMIT = True    # True = ask for confirmation before each submit
DRY_RUN            = True     # True = fill but do not submit (safe testing mode)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ── STEALTH INIT SCRIPT ────────────────────────────────────────────────────────
# Injected into every page context to suppress the automation signals that
# Cloudflare, Greenhouse, and similar bot-detection systems check for.

STEALTH_JS = """
// 1. Hide navigator.webdriver — the single strongest automation signal
Object.defineProperty(navigator, 'webdriver', { get: () => undefined });

// 2. Restore a realistic plugins array (empty in headless/automation browsers)
Object.defineProperty(navigator, 'plugins', { get: () => [
  { name: 'Chrome PDF Plugin',  filename: 'internal-pdf-viewer',          description: 'Portable Document Format' },
  { name: 'Chrome PDF Viewer',  filename: 'mhjfbmdgcfjbbpaeojofohoefgiehjai', description: '' },
  { name: 'Native Client',      filename: 'internal-nacl-plugin',         description: '' }
]});

// 3. Real Chrome always reports languages
Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en'] });

// 4. Restore window.chrome — absent in Playwright Chromium by default
if (!window.chrome) {
  window.chrome = {
    runtime:    {},
    app:        {},
    loadTimes:  function() {},
    csi:        function() {}
  };
}

// 5. Fix Permissions API (some detectors probe this)
try {
  const _origQuery = window.navigator.permissions.query.bind(navigator.permissions);
  window.navigator.permissions.query = (p) =>
    p.name === 'notifications'
      ? Promise.resolve({ state: Notification.permission })
      : _origQuery(p);
} catch(e) {}
"""

# ── APPLICANT DETAILS ──────────────────────────────────────────────────────────
# Loaded from config.json → "applicant" section at runtime.
# See config.example.json for the full list of supported fields.

def _load_applicant() -> dict:
    _defaults = {
        "first_name": "Jane", "preferred_name": "Jane", "last_name": "Doe",
        "email": "jane.doe@example.com", "phone": "555-000-0000",
        "location": "City, ST", "city": "City", "state": "State", "state_abbr": "ST",
        "zip": "", "country": "United States",
        "linkedin": "https://www.linkedin.com/in/yourprofile/",
        "website": "", "github": "", "citizenship": "US Citizen",
        "work_auth": "Yes", "requires_visa": "No", "years_exp": "10",
        "current_title": "Product Manager", "current_company": "",
        "salary_min": "0", "salary_mid": "0", "salary_max": "0",
        "notice_period": "2 weeks", "remote_pref": "Remote",
        "gender": "Decline to state", "ethnicity": "Decline to state",
        "race": "Decline to state", "veteran": "No", "disability": "No",
        "sexual_orientation": "Decline to state / Do not disclose",
        "transgender": "No / Decline to state",
        "requires_sponsorship": "No", "authorized_to_work": "Yes",
        "can_work_location": "Yes",
    }
    try:
        cfg_path = os.path.join(SCRIPT_DIR, "config.json")
        with open(cfg_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        applicant = cfg.get("applicant", {})
        return {**_defaults, **applicant}
    except Exception:
        return _defaults

APPLICANT = _load_applicant()

# ── MANUAL TEST JOBS ───────────────────────────────────────────────────────────
# Used when running with --manual flag for testing

MANUAL_JOBS = [
    # Add jobs here when testing with --manual flag
    # {
    #     "title": "Senior Product Manager",
    #     "company": "Example Corp",
    #     "apply_url": "https://example.com/careers/apply",
    #     "resume_path": os.path.join(SCRIPT_DIR, "resumes/My_Resume_Senior_PM_Example_Corp.docx"),
    #     "location": "Remote"
    # }
]

# ── CLAUDE API ─────────────────────────────────────────────────────────────────

def get_field_mapping(html_content, job_title, company_name, client):
    """
    Send page HTML to Claude and get back a field mapping: selector → value
    """
    applicant_name = f"{APPLICANT.get('first_name', '')} {APPLICANT.get('last_name', '')}".strip()
    prompt = f"""You are filling out a job application form for {applicant_name}.

Here is the full HTML of the current page:
{html_content[:50000]}

Here are Dom's details:
{json.dumps(APPLICANT, indent=2)}

This is for the role: {job_title} at {company_name}

Analyze every visible form field on this page. For each field:
- Identify the best CSS selector (prefer id, then name, then aria-label, then placeholder)
- Determine the correct value from Dom's details
- Identify the field type: text, email, tel, select, radio, checkbox, textarea, file, autocomplete

Return ONLY a JSON array. No other text:
[
  {{"selector": "#first_name", "type": "text",         "value": "Domenico"}},
  {{"selector": "#resume",     "type": "file",         "value": "__RESUME__"}},
  {{"selector": "#gender",     "type": "select",       "value": "Decline to state"}},
  {{"selector": "#city",       "type": "autocomplete", "value": "Miami"}},
  {{"selector": "#cover",      "type": "textarea",     "value": "Generated cover letter text here..."}}
]

Rules:
- Use "__RESUME__" as the value for any file upload field that is asking for a resume or CV
- For dropdowns (<select> elements), return the exact option text as it appears in the HTML
- Use type "autocomplete" for any city, location, address, or country field that shows a live
  typeahead / suggestions dropdown as you type (common in Greenhouse, Lever, Workday, iCIMS).
  These are plain <input> elements that are NOT a <select> but trigger a dropdown on keystroke.
  For a city/location autocomplete, set value to "Miami" so the dropdown can be triggered and
  the correct Miami, FL option selected.
- Country fields come in two distinct forms — identify which one it is from the surrounding HTML:
  1. Standalone country field (labelled "Country", "Country of residence", etc., NOT adjacent to a
     phone input): use "United States" — or the exact option text in the HTML that matches it.
  2. Phone dial-code / country-code selector (a dropdown that sits immediately before or beside a
     phone number input, or whose options are formatted as "+1 United States", "+44 United Kingdom",
     etc.): use "+1" — or the exact option text in the HTML for the United States dial code.
  Never confuse the two: check the label and the surrounding elements before deciding.
- Skip fields that are hidden, disabled, or clearly not relevant (honeypot fields)
- If a field is ambiguous, make your best judgment — do not skip it
- If a cover letter textarea is present, write a 3-sentence cover letter for this role: {job_title} at {company_name}. Keep it direct, confident, no fluff. Format: why this role, one proof point from the resume, one forward-looking sentence.

Always apply these answers exactly — match the closest available option in the HTML:
- Work authorization / legally authorized to work: Yes
- Require visa sponsorship / need sponsorship now or in future: No
- Can you work in [location] / able to work onsite: Yes
- US citizen / authorized: Yes
- Veteran status: No / I am not a veteran / Decline to self-identify
- Disability status: No / I do not have a disability / Decline to self-identify
- Gender: Decline to state / Prefer not to say / Do not wish to disclose
- Race / ethnicity: Decline to state / Prefer not to say / Do not wish to disclose
- Sexual orientation: Decline to state / Prefer not to say / Do not wish to disclose
- Transgender / gender identity: No / Decline to state / Prefer not to say
- For any EEO or demographic field not listed above: always pick the "decline" or "prefer not to answer" option
"""

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}]
    )

    raw = response.content[0].text.strip()
    # Strip markdown if present
    raw = re.sub(r'^```json\s*', '', raw)
    raw = re.sub(r'^```\s*', '', raw)
    raw = re.sub(r'\s*```$', '', raw)

    try:
        return json.loads(raw)
    except Exception as e:
        # Claude sometimes appends explanation text after the JSON — extract just the array
        try:
            decoder = json.JSONDecoder()
            parsed, _ = decoder.raw_decode(raw)
            return parsed if isinstance(parsed, list) else []
        except Exception:
            print(f"      ✗ Failed to parse Claude field mapping: {e}")
            print(f"      Raw response: {raw[:500]}")
            return []


def review_filled_form(html_content, client):
    """
    Send filled form HTML to Claude for pre-submit review
    Returns: {"ready": bool, "issues": [], "missed_fields": []}
    """
    prompt = f"""Review this filled job application form before submission.

Here is the HTML of the filled form:
{html_content[:50000]}

Check:
- Are all required fields filled?
- Are any fields obviously wrong?
- Are there any fields that were missed?

Return ONLY JSON:
{{
  "ready": true,
  "issues": [],
  "missed_fields": []
}}

If ready is false, list the issues. Do not mark as ready if there are obvious problems.
"""

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1000,
        messages=[{"role": "user", "content": prompt}]
    )

    raw = response.content[0].text.strip()
    raw = re.sub(r'^```json\s*', '', raw)
    raw = re.sub(r'^```\s*', '', raw)
    raw = re.sub(r'\s*```$', '', raw)

    try:
        return json.loads(raw)
    except Exception:
        try:
            decoder = json.JSONDecoder()
            parsed, _ = decoder.raw_decode(raw)
            return parsed if isinstance(parsed, dict) else {"ready": False, "issues": ["Unexpected response format"], "missed_fields": []}
        except Exception:
            return {"ready": False, "issues": ["Failed to parse review response"], "missed_fields": []}


# ── FORM FILLING ───────────────────────────────────────────────────────────────

def detect_captcha(page):
    """Check if page has a CAPTCHA"""
    captcha_selectors = [
        'iframe[src*="recaptcha"]',
        'iframe[src*="hcaptcha"]',
        '.cf-challenge',
        '[data-sitekey]'
    ]
    for sel in captcha_selectors:
        if page.query_selector(sel):
            return True
    return False


def fill_field(page, field, resume_path, main_page=None):
    """
    Fill a single field based on its type.
    main_page: the top-level Page object (may differ from `page` when the form
               lives inside an iframe). Used by the autocomplete handler to search
               for dropdown options that ATS platforms inject into the main document
               body rather than inside the iframe (e.g. Google Places .pac-item).
    Returns True if successful, False if failed.
    """
    selector = field.get("selector")
    field_type = field.get("type")
    value = field.get("value")

    try:
        # Wait for element to be visible
        page.wait_for_selector(selector, timeout=3000, state="visible")

        if field_type == "file":
            if value == "__RESUME__" and resume_path:
                page.set_input_files(selector, resume_path)
                return True
            else:
                return False

        elif field_type in ["text", "email", "tel", "number", "textarea"]:
            page.fill(selector, str(value))
            return True

        elif field_type == "autocomplete":
            # Type character-by-character to fire input/keydown events that trigger
            # the live-search dropdown, then click the best matching visible option.
            page.click(selector)
            page.fill(selector, "")                       # clear first
            page.type(selector, str(value), delay=60)     # slow-type to fire events

            # Candidate selectors for dropdown option elements, most-specific first.
            # Covers Google Places, Greenhouse, Lever, Workday, React-Select, etc.
            option_selectors = [
                '[role="option"]',
                '[role="listbox"] li',
                '[role="listbox"] [role="option"]',
                '.pac-item',                  # Google Places (injected into main page body)
                '.pac-item .pac-item-query',
                '[class*="suggestion"]',
                '[class*="autocomplete-item"]',
                '[class*="dropdown-item"]',
                '[class*="select__option"]',  # React Select
                '[class*="typeahead"] li',
                '.tt-suggestion',
                '[data-testid*="option"]',
            ]

            # Search both the form frame AND the main page.
            # Google Places and some ATS platforms inject dropdown options into the
            # main document body rather than inside the iframe where the form lives.
            search_contexts = [page]
            if main_page is not None and main_page is not page:
                search_contexts.append(main_page)

            value_lower = str(value).lower()

            def try_find_and_click(wait_seconds):
                """Wait, then search all contexts for a visible dropdown option."""
                time.sleep(wait_seconds)
                for ctx in search_contexts:
                    for opt_sel in option_selectors:
                        try:
                            all_opts = ctx.query_selector_all(opt_sel)
                            visible   = [o for o in all_opts if o.is_visible()]
                            if not visible:
                                continue
                            # Prefer an option whose text contains our search string
                            best = None
                            for opt in visible:
                                try:
                                    opt_text = opt.inner_text().strip().lower()
                                    if value_lower in opt_text or opt_text.startswith(value_lower[:4]):
                                        best = opt
                                        break
                                except Exception:
                                    continue
                            if best is None:
                                best = visible[0]   # fallback: first visible option
                            best.click()
                            time.sleep(0.5)
                            print(f"         ✓ Autocomplete: selected '{best.inner_text().strip()}'")
                            return True
                        except Exception:
                            continue
                return False

            # First attempt after 2 seconds, then one retry after another 2 seconds
            if try_find_and_click(2.0):
                return True
            print(f"         ⚠  Autocomplete: no dropdown on first attempt, retrying...")
            if try_find_and_click(2.0):
                return True

            # No dropdown appeared after both attempts — typed value is kept as-is
            print(f"         ⚠  Autocomplete: no dropdown found, kept typed value")
            return True

        elif field_type == "select":
            # Try exact match first, then partial match
            try:
                page.select_option(selector, label=value)
            except Exception:
                # Try value attribute instead of label
                try:
                    page.select_option(selector, value=value)
                except Exception:
                    # Try partial match on label
                    options = page.query_selector_all(f"{selector} option")
                    for opt in options:
                        if value.lower() in opt.inner_text().lower():
                            page.select_option(selector, label=opt.inner_text())
                            break
            return True

        elif field_type == "radio":
            # Try clicking radio by value
            try:
                page.click(f'{selector}[value="{value}"]')
            except Exception:
                # Try finding by label text
                labels = page.query_selector_all("label")
                for label in labels:
                    if value.lower() in label.inner_text().lower():
                        label.click()
                        break
            return True

        elif field_type == "checkbox":
            if value in ["Yes", "True", "true", True]:
                page.check(selector)
            else:
                page.uncheck(selector)
            return True

        else:
            print(f"         Unknown field type: {field_type}")
            return False

    except Exception as e:
        print(f"         ✗ Failed to fill {selector}: {e}")
        return False


def fill_form_page(page, job, client, resume_path):
    """
    Fill all fields on the current page (or iframe containing the form).
    Returns: number of fields filled successfully.
    """
    frame = get_active_form_frame(page)
    html  = frame.content()
    field_map = get_field_mapping(html, job.get("title"), job.get("company"), client)

    if not field_map:
        print("      ✗ No fields identified by Claude")
        return 0

    # Re-acquire the active frame after the Claude API call.
    # The API call takes 3-5 seconds — long enough for a Greenhouse, Workday,
    # or other ATS iframe to finish its own internal navigation and replace the
    # frame we captured above. Re-querying here ensures we fill into the live frame.
    time.sleep(1)
    frame = get_active_form_frame(page)

    filled_count = 0
    for field in field_map:
        # Pass the top-level page as main_page so the autocomplete handler can
        # search the main document body for dropdowns (e.g. Google Places .pac-item)
        # that are injected outside the iframe where the form lives.
        if fill_field(frame, field, resume_path, main_page=page):
            filled_count += 1
        time.sleep(0.3)

    return filled_count


def is_multi_step_form(page):
    """
    Check if there's a Next/Continue button (indicates multi-step form)
    """
    next_buttons = [
        'button:has-text("Next")',
        'button:has-text("Continue")',
        'button:has-text("Proceed")',
        'input[value="Next"]',
        'input[value="Continue"]'
    ]
    for sel in next_buttons:
        if page.query_selector(sel):
            return True
    return False


def is_submission_confirmed(page):
    """
    Check if we're on a confirmation/thank you page
    """
    content = page.content().lower()
    confirmation_keywords = [
        "submitted",
        "received",
        "thank you",
        "confirmation",
        "application complete",
        "we'll be in touch",
        "successfully applied"
    ]
    return any(keyword in content for keyword in confirmation_keywords)


# ── FORM DETECTION & NAVIGATION ────────────────────────────────────────────────

# Field selectors that only appear on real application forms (not job overviews)
APPLICATION_FIELD_SIGNALS = [
    'input[name*="first_name"]', 'input[name*="firstName"]', 'input[id*="first_name"]',
    'input[name*="last_name"]',  'input[name*="lastName"]',  'input[id*="last_name"]',
    'input[name*="email"]',      'input[id*="email"]',       'input[type="email"]',
    'input[type="file"]',        'input[name="resume"]',     'input[id*="resume"]',
    'input[name="name"][type="text"]',  # Lever pattern
]


def is_application_form(page):
    """
    Returns True if the current page (or any iframe on it) looks like an application form.
    Checks for field signals that only appear on real forms, not job overview pages.
    """
    # Check main page
    for sel in APPLICATION_FIELD_SIGNALS:
        if page.query_selector(sel):
            return True

    # Check iframes — some ATS platforms (Workday, some Greenhouse embeds) render
    # the form inside an iframe on the overview page
    try:
        frames = page.frames
        for frame in frames[1:]:  # skip main frame
            for sel in APPLICATION_FIELD_SIGNALS:
                try:
                    if frame.query_selector(sel):
                        return True
                except Exception:
                    continue
    except Exception:
        pass

    return False


def get_active_form_frame(page):
    """
    Returns the frame (main page or iframe) that contains the application form.
    Used so fill_form_page can work on the right frame.
    """
    for sel in APPLICATION_FIELD_SIGNALS:
        if page.query_selector(sel):
            return page

    try:
        for frame in page.frames[1:]:
            for sel in APPLICATION_FIELD_SIGNALS:
                try:
                    if frame.query_selector(sel):
                        return frame
                except Exception:
                    continue
    except Exception:
        pass

    return page  # fallback to main page


def find_apply_link(page):
    """
    Scans all links on the page and returns the href of the one most likely
    to be the prominent Apply CTA for this specific job (not a nav/header link).
    Platform-agnostic — works for any ATS or custom career page.
    Returns None if nothing found.
    """
    try:
        # Collect links with ancestor context so we can exclude nav/header elements
        links = page.evaluate("""() => {
            const NAV_TAGS = new Set(['NAV', 'HEADER']);
            const NAV_CLASSES = ['nav', 'navbar', 'navigation', 'header', 'topbar', 'menu'];

            function isInNav(el) {
                let node = el.parentElement;
                while (node) {
                    if (NAV_TAGS.has(node.tagName)) return true;
                    const cls = (node.className || '').toLowerCase();
                    if (NAV_CLASSES.some(c => cls.includes(c))) return true;
                    node = node.parentElement;
                }
                return false;
            }

            function isInMain(el) {
                let node = el.parentElement;
                while (node) {
                    const tag = node.tagName;
                    const cls = (node.className || '').toLowerCase();
                    const id  = (node.id || '').toLowerCase();
                    if (['MAIN', 'ARTICLE', 'SECTION'].includes(tag)) return true;
                    if (['main', 'content', 'body', 'job', 'posting'].some(c => cls.includes(c) || id.includes(c))) return true;
                    node = node.parentElement;
                }
                return false;
            }

            return Array.from(document.querySelectorAll('a[href], button')).map(e => ({
                href:   e.href || '',
                text:   e.innerText.trim(),
                cls:    (e.className || '').toLowerCase(),
                inNav:  isInNav(e),
                inMain: isInMain(e),
            }));
        }""")
    except Exception:
        return None

    apply_keywords  = ["apply now", "apply for this job", "apply for this position",
                       "start application", "apply today", "submit application", "apply here"]
    apply_path_hint = ["apply", "application", "job-apply", "careers/apply"]

    def valid_href(href):
        return href and not href.startswith("javascript") and href != "#" and not href.endswith("#")

    # Pass 1: strong text match, NOT in nav, preferring main content
    for in_main_required in [True, False]:
        for link in links:
            if link.get("inNav"):
                continue
            if in_main_required and not link.get("inMain"):
                continue
            text = link.get("text", "").lower().strip()
            href = link.get("href", "")
            if not valid_href(href):
                continue
            if any(kw in text for kw in apply_keywords):
                return href

    # Pass 2: href path hint + text contains "apply", not in nav
    for link in links:
        if link.get("inNav"):
            continue
        text = link.get("text", "").lower().strip()
        href = link.get("href", "")
        if not valid_href(href):
            continue
        if any(seg in href.lower() for seg in apply_path_hint) and "apply" in text:
            return href

    # Pass 3: class contains "apply", short text, not in nav
    for link in links:
        if link.get("inNav"):
            continue
        text = link.get("text", "").lower().strip()
        href = link.get("href", "")
        cls  = link.get("cls", "")
        if not valid_href(href):
            continue
        if "apply" in cls and len(text) < 30:
            return href

    return None


def advance_to_application_form(page, client):
    """
    Universal entry point. Ensures we are on the actual application form before
    field filling begins. Works across all ATS platforms and custom career pages.

    Strategy (in order):
      1. Already on a form? Done.
      2. Link scan — find the Apply CTA href and navigate directly.
      3. Claude — ask it to identify the Apply button selector or href.
      4. Give up gracefully and let fill_form_page handle whatever is on screen.
    """
    if is_application_form(page):
        print(f"  → Already on application form")
        return True

    # Step 2: link scan
    print(f"  → Not a form page — scanning for Apply link...")
    href = find_apply_link(page)
    if href:
        print(f"  → Found Apply link: {href}")
        try:
            page.goto(href, wait_until="load", timeout=30000)
            try:
                page.wait_for_load_state("networkidle", timeout=10000)
            except Exception:
                pass  # networkidle timeout is non-fatal
        except Exception:
            pass
        time.sleep(3)  # give iframe-based forms time to fully initialise
        if is_application_form(page):
            print(f"  → Now on application form: {page.url}")
            return True
        print(f"  → Navigated but still not on form ({page.url}) — trying Claude...")

    # Step 3: Claude
    print(f"  → Asking Claude to locate Apply button...")
    html = page.content()
    prompt = f"""This is a job listing page — NOT an application form. Find the single element that a user would click to begin filling out their application.

It might be a button saying "Apply Now", "Apply for this job", "Apply for this position", "Start Application", or similar. It might also be a link (<a href="...">) pointing to an external ATS like Greenhouse, Lever, Ashby, Workday, SmartRecruiters, iCIMS, BambooHR, Jobvite, etc.

HTML:
{html[:30000]}

Rules:
- ONLY return the element that starts the application process for THIS specific job.
- Do NOT return navigation links, job search links, "Back to jobs", or any other UI element.
- Prefer returning the href directly if the element is an <a> tag — navigating is safer than clicking.

Return ONLY JSON:
{{"selector": "a.apply-btn", "href": "https://boards.greenhouse.io/company/jobs/123", "text": "Apply Now"}}

If selector cannot be determined, set it to null. If href cannot be determined, set it to null.
No other text."""

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=300,
        messages=[{"role": "user", "content": prompt}]
    )

    raw = response.content[0].text.strip()
    raw = re.sub(r'^```json\s*', '', raw)
    raw = re.sub(r'^```\s*', '', raw)
    raw = re.sub(r'\s*```$', '', raw)

    try:
        result = json.loads(raw)
    except Exception:
        try:
            decoder = json.JSONDecoder()
            result, _ = decoder.raw_decode(raw)
        except Exception:
            result = {}

    # Navigate via href if Claude found one
    if result.get("href") and result["href"].startswith("http"):
        print(f"  → Claude found Apply URL: {result['href']}")
        page.goto(result["href"], wait_until="load", timeout=30000)
        try:
            page.wait_for_load_state("networkidle", timeout=10000)
        except Exception:
            pass
        time.sleep(3)
        if is_application_form(page):
            print(f"  → Now on application form: {page.url}")
            return True

    # Click via selector
    if result.get("selector"):
        print(f"  → Claude found Apply button: {result.get('text', result['selector'])}")
        try:
            page.click(result["selector"], timeout=5000)
            page.wait_for_load_state("load", timeout=15000)
            try:
                page.wait_for_load_state("networkidle", timeout=10000)
            except Exception:
                pass
            time.sleep(3)
            if is_application_form(page):
                print(f"  → Now on application form: {page.url}")
                return True
        except Exception as e:
            print(f"  ⚠️  Click failed: {e}")

    print(f"  ⚠️  Could not reach application form — proceeding with current page")
    return False


# ── MAIN APPLICATION FLOW ──────────────────────────────────────────────────────

def apply_to_job(job, client, browser):
    """
    Apply to a single job. Returns status: "Applied", "Needs Review", "Failed"
    """
    title = job.get("title", "Unknown")
    company = job.get("company", "Unknown")
    url = job.get("apply_url", "")
    resume_path = job.get("resume_path", "")

    print(f"\n{'━'*70}")
    print(f"  Applying: {title} @ {company}")
    print(f"  URL: {url}")
    print(f"  Resume: {os.path.basename(resume_path)}")
    print(f"{'━'*70}")

    if not os.path.exists(resume_path):
        print(f"  ✗ Resume file not found: {resume_path}")
        return "Failed — Resume not found"

    try:
        context = browser.new_context(
            viewport={"width": 1440, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/136.0.0.0 Safari/537.36"
            ),
            locale="en-US",
            timezone_id="America/New_York",
        )
        # Inject stealth patches before any page script runs
        context.add_init_script(STEALTH_JS)
        page = context.new_page()

        # Navigate to application URL
        print(f"  → Opening application page...")
        page.goto(url, wait_until="load", timeout=30000)
        try:
            page.wait_for_load_state("networkidle", timeout=10000)
        except Exception:
            pass
        time.sleep(2)

        # If we landed on a job overview page, click through to the actual form
        advance_to_application_form(page, client)

        # Guard: some sites call window.close() when navigated to directly,
        # which kills the page object before we can fill anything.
        if page.is_closed():
            print(f"  ✗ Page was closed by the site after navigation")
            return "Failed — Page closed by site"

        # Check for CAPTCHA
        if detect_captcha(page):
            print(f"  ⚠️  CAPTCHA detected — please solve it in the browser window")
            input("     Press Enter after solving CAPTCHA...")

        page_num = 1
        total_filled = 0

        # Fill pages until we reach confirmation or submission
        while True:
            print(f"  → Page {page_num}: Analyzing fields...")

            filled_count = fill_form_page(page, job, client, resume_path)
            total_filled += filled_count
            print(f"  → Page {page_num}: {filled_count} fields filled")

            # Check if multi-step
            if is_multi_step_form(page):
                print(f"  → Multi-step form detected, clicking Next...")
                # Click Next button
                next_selectors = [
                    'button:has-text("Next")',
                    'button:has-text("Continue")',
                    'input[value="Next"]'
                ]
                clicked = False
                for sel in next_selectors:
                    try:
                        page.click(sel, timeout=2000)
                        page.wait_for_load_state("networkidle", timeout=5000)
                        clicked = True
                        break
                    except Exception:
                        continue

                if clicked:
                    page_num += 1
                    time.sleep(1)
                    continue
                else:
                    print(f"  ✗ Could not find Next button")
                    break
            else:
                # Single page or final page — ready to submit
                break

        # Pre-submit review
        print(f"  → Pre-submit review...")
        html = page.content()
        review = review_filled_form(html, client)

        if not review.get("ready"):
            print(f"  ✗ Pre-submit review failed:")
            for issue in review.get("issues", []):
                print(f"     - {issue}")
            for missed in review.get("missed_fields", []):
                print(f"     - Missed field: {missed}")
            context.close()
            return "Needs Review"

        print(f"  → Pre-submit review: ✅ Ready")

        # Check for dry run
        if DRY_RUN:
            print(f"  → DRY RUN mode — skipping submission")
            context.close()
            return "Dry Run (not submitted)"

        # Pause before submit if enabled
        if PAUSE_BEFORE_SUBMIT:
            input(f"\n  Review the filled form in the browser, then press Enter to submit (or Ctrl+C to cancel)...")

        # Find and click submit button
        submit_selectors = [
            'button:has-text("Submit")',
            'button:has-text("Apply")',
            'button:has-text("Send")',
            'input[type="submit"]',
            'button[type="submit"]'
        ]

        submitted = False
        for sel in submit_selectors:
            try:
                page.click(sel, timeout=2000)
                page.wait_for_load_state("networkidle", timeout=10000)
                submitted = True
                break
            except Exception:
                continue

        if not submitted:
            print(f"  ✗ Could not find submit button")
            context.close()
            return "Failed — No submit button found"

        # Check for confirmation
        time.sleep(2)
        if is_submission_confirmed(page):
            print(f"  → Submitted ✅")
            context.close()
            return "Applied"
        else:
            print(f"  ⚠️  Submission status unclear — please verify manually")
            context.close()
            return "Needs Review — Verify submission"

    except Exception as e:
        print(f"  ✗ Application failed: {e}")
        return f"Failed — {str(e)[:100]}"


def run_applications(jobs):
    """
    Process all applications
    """
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    stats = {
        "applied": 0,
        "needs_review": 0,
        "failed": 0
    }

    with sync_playwright() as p:
        # Prefer the user's real Chrome install — it has a genuine fingerprint,
        # all real plugins, and passes most bot-detection checks automatically.
        # Fall back to Playwright's bundled Chromium if Chrome is not installed.
        _launch_kwargs = dict(
            headless=HEADLESS,
            args=["--no-sandbox", "--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],   # remove the automation banner flag
        )
        try:
            browser = p.chromium.launch(channel="chrome", **_launch_kwargs)
            print("  Browser: real Chrome")
        except Exception:
            browser = p.chromium.launch(**_launch_kwargs)
            print("  Browser: Playwright Chromium (Chrome not found)")

        for job in jobs:
            status = apply_to_job(job, client, browser)

            # Update stats
            if "Applied" in status:
                stats["applied"] += 1
            elif "Needs Review" in status or "Verify" in status:
                stats["needs_review"] += 1
            else:
                stats["failed"] += 1

            time.sleep(2)  # Pause between applications

        browser.close()

    # Print summary
    print(f"\n{'━'*70}")
    print(f"  SUMMARY")
    print(f"  Applied:       {stats['applied']}")
    print(f"  Needs review:  {stats['needs_review']}")
    print(f"  Failed:        {stats['failed']}")
    print(f"{'━'*70}\n")


# ── ENTRY POINT ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    # Check for flags
    manual_mode = "--manual" in sys.argv
    if "--dry-run" in sys.argv:
        DRY_RUN = True
    if "--no-pause" in sys.argv:
        PAUSE_BEFORE_SUBMIT = False

    print(f"\n{'='*70}")
    print(f"  APPLICATION FILLER")
    print(f"  Mode: {'MANUAL' if manual_mode else 'AUTO'}")
    print(f"  Dry Run: {DRY_RUN}")
    print(f"  Pause Before Submit: {PAUSE_BEFORE_SUBMIT}")
    print(f"  Headless: {HEADLESS}")
    print(f"{'='*70}\n")

    if manual_mode:
        print(f"Using manual test jobs from MANUAL_JOBS list")
        jobs = MANUAL_JOBS
    else:
        # Load from pipeline output
        if not os.path.exists(PIPELINE_OUTPUT):
            print(f"Error: {PIPELINE_OUTPUT} not found")
            print(f"Run this in auto mode after resume_pipeline.py completes")
            print(f"Or use --manual flag to test with MANUAL_JOBS")
            sys.exit(1)

        with open(PIPELINE_OUTPUT) as f:
            jobs = json.load(f)
        print(f"Loaded {len(jobs)} jobs from {PIPELINE_OUTPUT}")

    if not jobs:
        print("No jobs to process")
        sys.exit(0)

    run_applications(jobs)
