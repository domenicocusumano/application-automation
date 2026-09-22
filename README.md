# Job Application Automation

A self-hosted pipeline that scrapes job listings from LinkedIn, Built-in.com, Dynamite Jobs, or Indeed, scores each one against your background with Claude, and builds a tailored `.docx` or `.pdf` resume for every strong match — uploading it to Google Drive and logging everything to a Google Sheet. A **Manual Upload** mode lets you run a single job URL you found by hand through the same scoring/resume/logging pipeline. All settings are controlled from a local web UI.

The pipeline is **fully role-agnostic** — there are no baked-in defaults for any specific role type. All search terms, title filters, seniority tiers, and the scoring prompt are configured through the UI. It works equally well for engineering, design, data science, marketing, operations, or any other role.

> **Note on the form filler:** `application_filler.py` is included but is **experimental and not fully reliable** — application form layouts vary too widely across ATS platforms for fully automated filling to work consistently. The core value of this project is in candidate discovery and resume generation. The form filler is provided as a starting point for anyone who wants to build on it, and should be used at your own risk.

---

## How the pipeline fits together

```
Web UI (app.py + ui.html)
  └── Reads/writes config.json (all settings)
        └── Launches exactly one of these per run (priority order:
            Indeed > Dynamite > Built-in > LinkedIn), or Manual Upload
            for a single hand-picked URL:

job_scraper.py       builtin_scraper.py    dynamite_scraper.py   indeed_scraper.py
(LinkedIn)            (Built-in.com)        (Dynamite Jobs)       (Indeed)
  │                     │                     │                    │
  │ Phase 0/1/2:         Detail-page apply     Algolia search API   Real Chrome, genuine
  │ Top Applicant +      URL extraction +      (no browser needed   UA; session for full
  │ Recommended +        Cloudflare backoff     for search) +       results; apply URL
  │ keyword search        session (optional)    no login needed     resolved LATER, only
  │                                                                  for jobs that score
  │                                                                  above threshold
  └──────────────┬──────────────┴──────────────┴────────────────────┘
                 · title keywords / exclusions (configurable)
                 · location filter
                 · already-applied sheet check (pre-filter, before Claude)
                 · programmatic seniority scoring (no Claude yet)
                        └── feeds into ↓

              manual_scraper.py <url>  (Manual Upload — same hand-off, one URL at a time)
                        │
                        ▼
resume_pipeline.py
  └── Fetches full job description
        └── Scores fit with Claude (0–10)
              └── Hard disqualifiers: salary below minimum, wrong location
                    └── Builds tailored .docx resume for score ≥ threshold
                          └── Uploads to Google Drive
                                └── Logs to Google Sheet (Applications / Skips tab)

application_filler.py  ⚠ experimental, and currently only usable via --manual (see below)
  └── Opens each application URL in a real browser
        └── Claude reads the page HTML and maps every form field
              └── Playwright fills fields — use at your own risk
```

---

## Prerequisites

### Accounts and API keys

| Service | What it's for | Where to get it |
|---|---|---|
| Anthropic API | Claude scores jobs and writes resumes | console.anthropic.com |
| Google Service Account | Read/write Google Sheet, Google Drive upload | Google Cloud Console → IAM → Service Accounts |
| Google OAuth 2.0 credentials | Drive uploads using your personal account (service accounts have no storage quota) | Google Cloud Console → Credentials → Desktop app |
| LinkedIn account | Scraping job listings (uses your saved browser session) | linkedin.com |
| Built-in account *(optional)* | Reveals the real external apply URL for postings that gate it behind login | builtin.com |
| Indeed account *(optional but recommended)* | Without it, Indeed caps you to page 1 of results and can't resolve real apply URLs — see the Indeed scraper section below | indeed.com |

### Software

- **Python 3.9+**
- **Node.js 18+** — used by the resume builder to generate `.docx` files
- **Google Chrome** — the Indeed scraper (and its login/challenge helpers) and the form filler use your real Chrome install so the browser fingerprint is genuine; both fall back to Playwright's bundled Chromium if Chrome is not found, which is far more likely to get Cloudflare-challenged
- **LibreOffice** *(optional)* — required only if you enable **PDF output** in Settings. Install with:
  ```bash
  brew install --cask libreoffice
  ```
  If LibreOffice is not installed and PDF is selected, the pipeline falls back to `.docx` automatically.

### Python packages

```bash
pip3 install anthropic playwright gspread google-auth google-api-python-client \
             google-auth-oauthlib python-dotenv fastapi uvicorn python-docx pypdf
playwright install chromium
```

### Node packages

```bash
npm install   # installs docx (defined in package.json)
```

---

## Setup

### 1. `.env` file

Create a `.env` file in the project root:

```env
ANTHROPIC_API_KEY=sk-ant-...
GOOGLE_CREDS_FILE=google_credentials.json
GDRIVE_FOLDER_ID=your_google_drive_folder_id
```

`GOOGLE_SHEET_ID` is no longer required in `.env` — it is entered via the web UI and saved to `config.json`.

### 2. Google Service Account (`google_credentials.json`)

1. Go to **Google Cloud Console → APIs & Services → Credentials**
2. Create a Service Account → download the JSON key → save it as `google_credentials.json` in the project root
3. Enable the **Google Sheets API** for your project
4. Share your Google Sheet with the service account's `client_email` (found inside `google_credentials.json`) — give it **Editor** access

### 3. Google OAuth credentials (`oauth_credentials.json`)

Used only for Google Drive uploads (service accounts cannot upload to a personal Drive).

1. Go to **Google Cloud Console → APIs & Services → Credentials**
2. Create an OAuth 2.0 Client ID → Desktop app type → download JSON → save as `oauth_credentials.json`
3. Enable the **Google Drive API** for your project
4. On first run, a browser window opens to authorize access — a `gdrive_token.json` is saved and reused on every subsequent run

### 4. Google Sheet structure

#### Quickstart template

Make a copy of this public template and use it as your tracker:

> **File → Make a copy** of the sheet, then paste the URL of your copy into the web UI under Settings → Google Sheet URL.

If you prefer to create the sheet from scratch, follow the instructions below exactly — column names and tab names are case-sensitive and must be spelled exactly as shown.

#### Tabs

The sheet must have exactly two tabs:

| Tab name | Purpose |
|---|---|
| `Applications` | Jobs for which a resume was built (score ≥ threshold) |
| `Skips` | Jobs below the score threshold or hard-disqualified (salary / location) |

Both tabs are checked during deduplication — a job already in either tab will not be surfaced as a candidate again.

#### Column headers (both tabs — same structure)

Column order does not matter. All nine columns must exist on both tabs. Names are **case-sensitive** and must be spelled exactly as shown:

| Column name | Written by | Notes |
|---|---|---|
| `Company` | `resume_pipeline.py` | Company name extracted from the scraper |
| `Position Title` | `resume_pipeline.py` | Job title |
| `URL` | `resume_pipeline.py` | The actual apply URL (company career site). You can replace this with the real URL after applying — dedup will still work via `Linked In URL`. |
| `Linked In URL` | `resume_pipeline.py` | The scraper source URL (LinkedIn listing or Built-in listing). Never changes — used as a stable dedup key. |
| `Date` | `resume_pipeline.py` | Date the row was written (YYYY-MM-DD) |
| `Location` | `resume_pipeline.py` | Work location extracted from the job |
| `Claude Score` | `resume_pipeline.py` | Claude's 0–10 fit score |
| `Claude notes` | `resume_pipeline.py` | Claude's assessment / disqualification reason |
| `* Applied` | You (manually) | Mark with any non-empty value once you have actually submitted the application. Not read by the pipeline — for your own tracking. |

#### How to set it up from scratch

1. Create a new Google Sheet at sheets.google.com
2. Rename the first tab to `Applications` (double-click the tab name)
3. Add a second tab named `Skips`
4. On **both** tabs, paste this row into row 1 exactly as written:

```
Company	Position Title	URL	Linked In URL	Date	Location	Claude Score	Claude notes	* Applied
```

   The easiest way is to copy the line above and paste it into cell A1 — Google Sheets will split it across columns automatically if you paste with **Ctrl+Shift+V** (paste values only) or use **Data → Split text to columns** with Tab as the delimiter.

5. Share the sheet with your service account email (found in `google_credentials.json` → `client_email`) and give it **Editor** access
6. Copy the sheet URL and paste it into the web UI under **Settings → Google Sheet URL**

### 5. Start the web UI

```bash
uvicorn app:app --reload --port 8000
```

Then open **http://localhost:8000** in your browser. All settings are configured here, including the Google Sheet URL.

---

## Configuration

All settings are managed through the web UI and persisted to `config.json`. You do not need to edit any script files for normal operation.

| Setting | Description |
|---|---|
| **First name / Last name** | Used to name generated resume files: `FirstName_LastName_Company_Role.docx` |
| **Score threshold** | Minimum Claude score (0–10) to trigger a resume build |
| **Scoring model** | Claude model used to evaluate and score each job. Haiku is ~20× cheaper and accurate enough for pass/fail decisions. |
| **Resume build model** | Claude model used to write the tailored resume content. Sonnet or Opus recommended for output quality. |
| **Resume output format** | `DOCX` (default, editable in Word) or `PDF` (requires LibreOffice — see Prerequisites) |
| **Top Applicant feed** | Also scrape LinkedIn's "Top Applicant" feed (requires LinkedIn Premium) |
| **LinkedIn enabled** | Run the LinkedIn scraper |
| **LinkedIn search term** | Keyword used for the Phase 2 LinkedIn keyword search (e.g. `Software Engineer`, `Data Scientist`, `UX Designer`) |
| **Built-in enabled** | Run the Built-in.com scraper |
| **Built-in URL** | The Built-in.com search results URL to paginate through — build it by searching on the site |
| **Dynamite enabled** | Run the Dynamite Jobs scraper |
| **Dynamite URL** | A Dynamite Jobs search results URL — the scraper extracts the query text and category filters from it and queries their Algolia search API directly (no browser needed) |
| **Indeed enabled** | Run the Indeed scraper |
| **Indeed URL** | An indeed.com job search results URL (e.g. `https://www.indeed.com/jobs?q=...&l=Remote`) |
| **Role keywords (must match)** | A job title must contain at least one of these phrases to be considered. Change these to match your target role. One phrase per line. |
| **Excluded titles** | Titles containing any of these exact phrases are rejected, even if they match a role keyword (e.g. block "program manager" while searching for "product manager"). One phrase per line. |
| **Excluded title words** | Individual words that, if found as a whole word in a title, cause rejection. Useful for blocking adjacent professions. Comma-separated. |
| **Seniority tiers** | Ordered list of title keywords for programmatic scoring — earlier = higher score. Not a hard filter; unmatched titles score 3.0/10. |
| **Preferred locations** | Comma-separated list (e.g. `remote, new york`). Jobs not matching are filtered out, and Claude hard-disqualifies roles requiring office presence outside these locations. |
| **Salary minimum** | If the JD states a max salary below this number, the job is skipped. Set to 0 to disable. |
| **Max candidates** | How many candidates a single scraper run collects before stopping (per-scraper cap, not a global one) |
| **Google Drive folder ID** | Optional — upload built resumes into a specific Drive folder instead of the root of your Drive |
| **Google Sheet URL** | Paste your full sheet URL — the ID is extracted automatically |

Only one scraper runs per pipeline execution. If more than one of LinkedIn/Built-in/Dynamite/Indeed is enabled, the server picks in this priority order: **Indeed > Dynamite > Built-in > LinkedIn**. The UI's Run button is meant to keep only one enabled at a time.

Settings take effect immediately on the next run. No restart needed.

### Targeting your role

Update these settings in the UI for your target role:

| Setting | What to enter |
|---|---|
| LinkedIn search term | The job title you want LinkedIn to search — e.g. `Software Engineer`, `Data Scientist`, `UX Designer` |
| Role keywords | Phrases a matching title must contain. One per line. Leave blank to accept all titles. |
| Excluded titles | Exact phrases that disqualify a title even if it contains a role keyword. One per line. |
| Excluded title words | Individual words that reject a title when found as a whole word. Leave blank to skip word-level filtering. |
| Built-in URL | Build by searching builtin.com with your filters (role, location, remote, etc.) and paste the results URL |
| Dynamite URL | Build by searching dynamitejobs.com with your filters and paste the results URL |
| Indeed URL | Build by searching indeed.com with your filters and paste the results URL |

Also edit `background_prompt.txt` to reflect your actual experience and scoring criteria — this is the primary input Claude uses when scoring and writing resumes.

---

## Component reference

### Web UI — `app.py` + `ui.html`

The local FastAPI server that hosts the control panel. From the UI you can:

- Start a scraping run and watch the live output stream
- Configure all pipeline settings
- Edit the background prompt (your resume context sent to Claude)
- Upload your base resume (`.pdf` or `.docx`)
- Re-authenticate your LinkedIn, Built-in, or Indeed session
- Run a single job URL through Manual Upload

Run with:
```bash
uvicorn app:app --reload --port 8000
```

---

### LinkedIn scraper — `job_scraper.py`

Scrapes LinkedIn using your saved browser session. Three phases:

- **Phase 0** *(optional)*: LinkedIn's "Top Applicant" collection (requires Premium)
- **Phase 1**: Your personalized "Recommended" jobs feed — exhausted fully before Phase 2
- **Phase 2**: Keyword search using your configured **LinkedIn search term** — paginated until the candidate target is reached

The search term (set in Settings → LinkedIn → Search term) drives Phase 2. Phases 0 and 1 use LinkedIn's personalization, so they surface roles matching your profile regardless of the search term.

For each job that passes all title and location filters, it navigates to the LinkedIn job page and extracts the **actual external apply URL** (the company career site link, not just the LinkedIn URL), then runs a second dedup pass against the sheet.

After collecting candidates, they are **ranked programmatically** by seniority tier score — no Claude API call at this stage.

**First run:** A browser window opens for manual LinkedIn login. After login, press Enter in the terminal. The session is saved to `linkedin_session.json` and reused on every run.

**Session expiry:** Delete `linkedin_session.json` and run again (or use the Re-authenticate button in the UI).

Run standalone:
```bash
python3 job_scraper.py           # scrape and print ranked list
python3 job_scraper.py --resume  # scrape, rank, and immediately build resumes
```

---

### Built-in scraper — `builtin_scraper.py`

Scrapes Built-in.com via a headless browser. The search is driven entirely by the **Built-in URL** you configure — build the URL by doing a search on builtin.com with your filters (role, location, remote, etc.) and paste the results page URL into Settings. Workflow per job:

1. Extracts title and location from the list page card (JS-evaluated to get the work model — Remote/Hybrid/In-office)
2. Filters by title (using your configured role keywords and exclusions) and location
3. **Pre-checks against the Google Sheet** by Built-in URL and numeric job ID — skips the detail page entirely if already applied
4. Visits the detail page to extract: company, full job description, and the **actual external apply URL** (the button that takes you to the company's career site). Many postings now gate this URL behind a Built-in login — see below.
5. Checks within-run fingerprint (company + title) to catch the same job appearing on multiple pages with different URLs
6. Scores by seniority tier and adds to the candidate list

After collecting up to **Max candidates** (default 10, set in Settings), automatically calls `resume_pipeline.run_pipeline()`.

**Built-in login (optional, recommended):** Built-in no longer exposes the real external apply URL for most postings to logged-out visitors — without a session, the scraper falls back to the Built-in listing URL for those jobs. Click **Re-login Built-in** in the web UI (or run `python3 relogin_builtin.py`) once to open a browser, log in, and save the session to `builtin_session.json`. Reused on every run; delete the file or click Re-login again if it expires.

Runs standalone (and is the default when Built-in is enabled in the UI):
```bash
python3 builtin_scraper.py
```

#### URL deduplication (all scrapers)

URLs are normalized before any comparison: query parameters, URL fragments, trailing slashes, `www.` prefix, and case are all stripped. The numeric job ID is also extracted from Built-in URLs (and the `vjk`/`jk` job-key param from Indeed URLs) and stored as a fallback key — so slug/tracking-param changes in the URL don't defeat dedup. Both the `URL` column and `Linked In URL` column from the sheet are checked, as are both the apply URL and the listing URL from the scraper. A fuzzy company+title fallback also catches the same posting appearing on a different board entirely, with no URL in common.

Each scraper also does its own dedup pre-check (against a snapshot of the sheet taken at the start of the run) before spending time resolving a detail page or apply URL — `resume_pipeline.py` then re-checks fresh right before committing to a Claude call and resume build, closing the race window if another run logged the same job in the meantime.

---

### Dynamite Jobs scraper — `dynamite_scraper.py`

Unlike the other scrapers, this one doesn't drive a browser at all — Dynamite Jobs' search results are backed by an Algolia index, so the scraper extracts the search text and category filters straight out of your configured **Dynamite URL** and queries that Algolia API directly. Faster and less fragile than DOM scraping, and there's no login/session to manage since Dynamite doesn't gate any of its listing data.

Every listing on the site is remote by definition, so location filtering is mostly a formality here.

After collecting candidates it calls `resume_pipeline.run_pipeline()` directly, same as the other scrapers.

Runs standalone (and is the default when Dynamite is enabled in the UI):
```bash
python3 dynamite_scraper.py
```

---

### Indeed scraper — `indeed_scraper.py`

Scrapes Indeed via headless **real Google Chrome** (Playwright's `channel="chrome"`, falling back to bundled Chromium only if Chrome isn't installed) — a plain HTTP request gets an immediate Cloudflare block, and Playwright's bundled Chromium with a hand-written user-agent gets fingerprinted (see Troubleshooting). Workflow per job:

1. Extracts title/company/location/salary from the search results list
2. Filters by title and location, and pre-checks the already-applied sheet (by Indeed's `jk` job key) before spending a click on the detail pane
3. Clicks the job's title, which updates the URL in place (`&vjk=<jk>`) rather than a full navigation, and reads the JD off the detail pane (2–4s randomized pacing between clicks). That JD is passed straight to Claude — `resume_pipeline.py` does not re-fetch it
4. Notes whether the job is "Apply with Indeed" (native) or "Apply on company site" (external) — but does **not** resolve the external URL yet
5. Scores by seniority tier and hands off to `resume_pipeline.run_pipeline()`
6. **Only for jobs that clear the Claude score threshold**, resolves the real "Apply on company site" URL via Indeed's `/applystart` redirect, right before the resume build, and writes that to the sheet's `URL` column. Everything else keeps the Indeed listing URL.

Step 6 is deliberately last. `/applystart` is Indeed's most Cloudflare-sensitive endpoint, and resolving it eagerly for every candidate (most of which score below threshold and get thrown away) was ~40 hits per run — enough to get a residential IP hard-blocked for hours. Deferring it cuts that to 1–3 hits per run. Resolution gets one retry; if it's still challenged, the listing URL is kept and you click "Apply on company site" yourself when applying.

**Indeed login (recommended):** Click **Re-login Indeed** in the web UI (or run `python3 relogin_indeed.py`) to open a browser, log in, and save the session to `indeed_session.json`. Indeed gates pagination past page 1 (~15 jobs) and the real apply URL behind login. The login browser is the same real Chrome the scraper uses, on purpose — Cloudflare binds its clearance cookie to the browser fingerprint that earned it.

If a Cloudflare check **with a widget** ("Verify you are human") ever shows up, `python3 indeed_solve_challenge.py` opens a visible browser to clear it by hand. If the page instead says "Additional Verification Required" with only a "Return home" link, that's an IP-level block with nothing to solve — wait a few hours.

Runs standalone (and is the default when Indeed is enabled in the UI):
```bash
python3 indeed_scraper.py
```

---

### Manual Upload — `manual_scraper.py`

For a single job posting you found by hand (not via any scraper). Extracts the title and company from the URL/page, then runs it through the exact same Claude scoring, resume building, and sheet logging as the automated scrapers. Available from the **Manual** tab in the web UI, or standalone:
```bash
python3 manual_scraper.py https://company.com/careers/some-job-posting
```

---

### Resume pipeline — `resume_pipeline.py`

Processes a list of job candidates end-to-end:

1. **Fetches the full job description** in a headless browser. Primary source is always the apply URL. Falls back to pre-scraped listing text, then to the listing page URL if the primary fetch fails.

2. **Scores the job with Claude** (0–10 scale). The scoring prompt reads your background from `background_prompt.txt` and considers: role-type alignment, seniority match, industry fit, and location.

3. **Hard disqualification checks** (applied before the score threshold):
   - If the JD states a max salary below your configured minimum → Skips tab
   - If the JD requires living near an office city not in your preferred locations → Skips tab
   - Occasional travel (≤ once/month) is not disqualifying

4. **Builds a tailored resume** (`.docx` or `.pdf`, configurable) for any job scoring ≥ the score threshold (requires a base resume in `base_resume/`):
   - Experience bullets reordered to lead with the most relevant stories for this specific role
   - Competency categories reordered to match the JD's priorities
   - Tailored summary written for the role
   - Correct role headers, hyperlinked LinkedIn, two-page hard limit

5. **Uploads the resume to Google Drive** using your personal OAuth credentials

6. **Logs to the Google Sheet**: Applications tab if a resume was built, Skips tab if below threshold or disqualified

Run standalone with a JSON file:
```bash
python3 resume_pipeline.py jobs.json
```

Where `jobs.json` is an array:
```json
[
  {
    "title": "Senior Software Engineer",
    "company": "Acme Corp",
    "location": "Remote",
    "url": "https://acme.com/careers/sse-role",
    "linkedin_url": "https://www.linkedin.com/jobs/view/1234567890"
  }
]
```

---

### Background prompt — `background_prompt.txt`

Your full resume context, candidate facts, style rules, and scoring calibration anchors. This is the primary source of truth Claude uses when scoring and writing resumes. Edit it from the **Background Prompt** tab in the web UI or directly in the file.

---

### Form filler — `application_filler.py` ⚠ experimental

> **This component is not fully reliable.** Application form layouts vary widely across ATS platforms (Greenhouse, Lever, Workday, iCIMS, custom systems) and change frequently. Automated field detection works well for simple forms but fails unpredictably on multi-step, iframe-based, or heavily dynamic forms. Use it as a starting point or build on it — but always review before submitting and don't rely on it for important applications.

Attempts to automatically fill and submit job application forms:

1. Opens a browser (real Chrome if installed for genuine fingerprint, Playwright Chromium as fallback) with stealth patches applied to suppress bot-detection signals
2. Navigates to the application URL; if on a job overview page, finds the actual form via link scan or Claude fallback
3. Sends the page HTML to Claude, which returns a field-by-field mapping (CSS selector → value → input type)
4. Fills every field: text, email, phone, `<select>` dropdowns, radio buttons, checkboxes, file uploads (resume), textareas, and autocomplete/typeahead fields
5. For multi-step forms, clicks Next and repeats on each page
6. Runs a pre-submit review with Claude to verify all required fields are filled
7. In **dry-run mode** (default): stops before submitting for manual review
8. In **live mode**: submits and checks for a confirmation page

**`DRY_RUN = True` is the default.** You must explicitly set `DRY_RUN = False` in the file to actually submit.

> **Auto mode is currently non-functional.** `python3 application_filler.py` (no flags) reads jobs from `pipeline_output.json`, but nothing in this codebase writes that file anymore — `resume_pipeline.py` now logs straight to the Google Sheet/Drive instead of a local JSON hand-off. Only `--manual` mode (below, using the `MANUAL_JOBS` list you edit directly in the file) currently works.

Run:
```bash
python3 application_filler.py --manual     # uses MANUAL_JOBS list in the file (the only working mode right now)
python3 application_filler.py --dry-run    # fill but do not submit
```

---

## Running the full end-to-end pipeline

```bash
# 1. Start the web UI
uvicorn app:app --reload --port 8000

# 2. Configure settings at http://localhost:8000:
#    - Paste your Google Sheet URL
#    - Set preferred locations, seniority tiers, score threshold
#    - Upload your base resume
#    - Edit the background prompt if needed

# 3. Click "Run" in the UI to start a scraping run
#    OR run a scraper directly from the terminal (each includes the resume pipeline):
python3 job_scraper.py --resume     # LinkedIn
python3 builtin_scraper.py          # Built-in
python3 dynamite_scraper.py         # Dynamite Jobs
python3 indeed_scraper.py           # Indeed
python3 manual_scraper.py <url>     # Manual Upload — a single job URL

# 4. Review the resumes/ folder and Google Drive

# 5. Test-fill an application (dry run, visible browser)
python3 application_filler.py --manual --dry-run

# 6. Submit for real (set DRY_RUN = False in application_filler.py)
python3 application_filler.py --manual
```

---

## File structure

```
application-automation/
├── app.py                      # FastAPI server — hosts the web UI
├── ui.html                     # Web UI (settings, live log, prompt editor)
│
├── job_scraper.py              # LinkedIn scraper + programmatic ranker
├── builtin_scraper.py          # Built-in.com scraper + programmatic ranker
├── dynamite_scraper.py         # Dynamite Jobs scraper (Algolia API, no browser) + ranker
├── indeed_scraper.py           # Indeed scraper + programmatic ranker
├── manual_scraper.py           # Manual Upload — runs one hand-picked URL through the pipeline
├── dedup_common.py             # Shared fuzzy company/title match, used by all four scrapers' already_applied()
├── resume_pipeline.py          # Claude scorer + .docx builder + Drive uploader
├── application_filler.py       # Browser-based form filler + submitter (⚠ experimental, --manual only)
├── relogin.py                  # LinkedIn session re-authentication helper
├── relogin_builtin.py          # Built-in session re-authentication helper
├── relogin_indeed.py           # Indeed session re-authentication helper
├── indeed_solve_challenge.py   # Optional: clear an interactive Cloudflare check for Indeed by hand (rarely needed)
│
├── background_prompt.txt       # Your resume context and scoring rules for Claude
├── config.json                 # All pipeline settings (managed via UI)
│
├── base_resume/                # Drop your base resume here (gitignored — never committed)
│   └── Your_Resume.docx        #   or upload it via the web UI
│
├── resumes/                    # Generated tailored resumes (.docx or .pdf, gitignored)
│
├── .env                        # API keys (git-ignored)
├── google_credentials.json     # Google service account key (git-ignored)
├── oauth_credentials.json      # Google OAuth desktop app credentials (git-ignored)
├── gdrive_token.json           # Auto-created Drive OAuth token (git-ignored)
├── linkedin_session.json       # Saved LinkedIn browser session (git-ignored)
├── builtin_session.json        # Saved Built-in browser session (git-ignored)
├── indeed_session.json         # Saved Indeed browser session (git-ignored)
│
├── package.json                # Node dependency: docx
└── node_modules/               # Node packages
```

---

## Sensitive files — never commit these

| File | Why it's sensitive |
|---|---|
| `.env` | Contains your Anthropic API key |
| `google_credentials.json` | Google service account private key — full read/write access to your Sheet |
| `oauth_credentials.json` | Google OAuth client secret |
| `gdrive_token.json` | Live Drive access token — grants upload access to your personal Drive |
| `linkedin_session.json` | Saved browser cookies — anyone with this file can act as you on LinkedIn |
| `builtin_session.json` | Saved browser cookies — anyone with this file can act as you on Built-in |
| `indeed_session.json` | Saved browser cookies — anyone with this file can act as you on Indeed |

All seven are already in `.gitignore`. Verify before pushing:
```bash
git status --short | grep -E "\.env|credentials|token|session"
```

---

## Troubleshooting

**LinkedIn session expired**
Delete `linkedin_session.json` and run `job_scraper.py` again, or click **Re-authenticate** in the web UI.

**No jobs found by the LinkedIn scraper**
LinkedIn's DOM changes regularly. If no cards are being extracted, check `extract_job_from_card()` in `job_scraper.py` — the CSS selectors may need updating.

**Built-in scraper keeps logging "requires a logged-in session to reveal this apply URL"**
Delete `builtin_session.json` and click **Re-login Built-in** in the web UI (or run `python3 relogin_builtin.py`) to refresh the session. Without a valid session, Built-in only reveals the Built-in listing URL for postings that gate their real apply link behind an account — this is expected, not a scraper bug.

**Built-in scraper finds 0 jobs**
Check the `extract_jobs_from_page()` selector list in `builtin_scraper.py`. The `[selector]` log line shows which selector matched and how many links were found — if it shows 0, Built-in's markup has changed.

**Indeed scraper finds 0 results, even on page 1**
Almost always Cloudflare showing a "Just a moment... / Additional Verification Required" page instead of real results — not a selector or code problem. Two distinct cases:
- **A widget to click** ("Verify you are human"): run `python3 indeed_solve_challenge.py` to clear it in a visible browser.
- **No widget, just "Return home"**: an IP-level block. Nothing to solve — it will hit your own Chrome too. Common triggers are a VPN (turn it off) or too much automated traffic to Indeed in a short window. Stop running Indeed for a few hours and it clears on its own; every retry while blocked extends it.

**Indeed logs `apply redirect challenged` / `still challenged by Cloudflare after retry`**
The `/applystart` redirect has a stricter Cloudflare check than the search page. This now only runs for jobs that score above threshold (1–3 per run) and retries once, so it should be rare. When it happens, the job keeps its Indeed listing URL in the sheet — click "Apply on company site" yourself when applying. If it's happening on every job, you're probably in the IP-block state above.

**Cloudflare challenge loops — solve it, page reloads, asked again**
The browser's fingerprint is inconsistent, so Cloudflare rejects the clearance after the solve. The classic cause is a hand-written user-agent string (e.g. `Chrome/131`) on a browser that's actually a different build — Chromium sends its real version in `Sec-CH-UA` headers regardless, and the contradiction fails the post-solve check every time. The Indeed scraper and its helpers avoid this by using real Chrome and only ever stripping the `Headless` token from the browser's own UA; if you change the browser setup, keep it that way. Solving harder won't help — and each failed loop worsens the IP's reputation.

**Indeed scraper caps out at ~15 jobs / page 1 only**
No valid session. Click **Re-login Indeed** in the web UI (or run `python3 relogin_indeed.py`). After switching browser setups (e.g. this repo's move to real Chrome), re-login once — the old session's Cloudflare clearance was bound to the old fingerprint.

**Resume build fails with a Node error**
Make sure `npm install` was run in the project directory (not with `-g`). The `require('docx')` call resolves from the local `node_modules/`.

**Google Sheets write fails**
Verify the service account email (from `google_credentials.json` → `client_email`) has Editor access on the sheet. Both the Google Sheets API and (for Drive) Google Drive API must be enabled in your Google Cloud project.

**Google Drive upload fails on first run**
A browser window will open to complete the OAuth consent flow. After authorizing, `gdrive_token.json` is saved automatically and reused on every run after.

**Score threshold setting not taking effect**
The pipeline re-reads `config.json` at the start of each run. Saving settings in the UI writes `config.json` immediately — no restart needed.

**Duplicate jobs appearing despite being in the sheet**
Both the `URL` and `Linked In URL` columns are checked from both the Applications and Skips tabs. Built-in jobs are also matched by the numeric job ID extracted from the Built-in URL, and Indeed jobs by the `vjk`/`jk` job-key param, so slug/tracking-param changes don't defeat dedup. If duplicates still appear, check that the service account has read access to the sheet and that the Google Sheet URL in Settings is correct. Note that each scraper's own pre-check runs against a snapshot of the sheet taken at the start of that run — if another run is logging to the sheet concurrently, a job can slip past the scraper's pre-filter and only get caught by `resume_pipeline.py`'s later re-check (visible in the log as "Already in Applications/Skips — skipping" with no Claude score printed above it) — this is expected, not a bug.

**PDF output falls back to DOCX**
LibreOffice is not installed or not on your PATH. Install it with `brew install --cask libreoffice`, then re-run. The pipeline prints a message confirming it fell back and which file was saved.

**Form filler gets blocked by Cloudflare**
The script detects a CAPTCHA challenge page and pauses for manual solving before continuing. Make sure `HEADLESS = False` (default) so the browser is visible.
