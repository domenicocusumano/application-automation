# Job Application Automation

A self-hosted pipeline that scrapes PM roles from LinkedIn and Built-in.com, scores each one against your background with Claude, builds a tailored `.docx` resume for every strong match, uploads it to Google Drive, and logs everything to a Google Sheet — all controlled from a local web UI.

---

## How the pipeline fits together

```
Web UI (app.py + ui.html)
  └── Reads/writes config.json (all settings)
        └── Launches one of two scrapers:

job_scraper.py  (LinkedIn)          builtin_scraper.py  (Built-in.com)
  └── Filters candidates by title,      └── Same filters + fetches actual
      location, seniority, already-         apply URL from detail page
      applied sheet check                   └── Calls resume_pipeline below
        └── Programmatic seniority
            scoring (no Claude)
              └── (optional) feeds into ↓

resume_pipeline.py
  └── Fetches full job description
        └── Scores fit with Claude (0–10)
              └── Hard disqualifiers: salary below minimum, wrong location
                    └── Builds tailored .docx resume for score ≥ threshold
                          └── Uploads to Google Drive
                                └── Logs to Google Sheet (Applications / Skips tab)

application_filler.py
  └── Opens each application URL in a real browser
        └── Claude reads the page HTML and maps every form field
              └── Playwright fills and submits
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

### Software

- **Python 3.9+**
- **Node.js 18+** — used by the resume builder to generate `.docx` files
- **Google Chrome** — the form filler uses your real Chrome install to avoid bot detection; falls back to Playwright's bundled Chromium if Chrome is not found

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
| **Score threshold** | Minimum Claude score (0–10) to trigger a resume build |
| **Top Applicant feed** | Also scrape LinkedIn's "Top Applicant" feed (requires LinkedIn Premium) |
| **LinkedIn enabled** | Run the LinkedIn scraper |
| **LinkedIn search term** | Keyword used for the Phase 2 LinkedIn search |
| **Built-in enabled** | Run the Built-in.com scraper |
| **Built-in URL** | The Built-in.com search URL to paginate through |
| **Seniority tiers** | Ordered list of title keywords — earlier = higher score. Used for programmatic candidate ranking (no Claude). |
| **Preferred locations** | Comma-separated list (e.g. `remote, miami`). Jobs not matching are filtered out of candidates AND hard-disqualified by Claude if the JD requires proximity to another city. |
| **Salary minimum** | If the JD states a max salary below this number, the job goes to Skips. Set to 0 to disable. |
| **Google Sheet URL** | Paste your full sheet URL — the ID is extracted automatically |

Settings take effect immediately on the next run. No restart needed.

---

## Component reference

### Web UI — `app.py` + `ui.html`

The local FastAPI server that hosts the control panel. From the UI you can:

- Start a scraping run and watch the live output stream
- Configure all pipeline settings
- Edit the background prompt (your resume context sent to Claude)
- Upload your base resume (`.pdf` or `.docx`)
- Re-authenticate your LinkedIn session

Run with:
```bash
uvicorn app:app --reload --port 8000
```

---

### LinkedIn scraper — `job_scraper.py`

Scrapes LinkedIn using your saved browser session. Three phases:

- **Phase 0** *(optional)*: LinkedIn's "Top Applicant" collection (requires Premium)
- **Phase 1**: Your personalized "Recommended" jobs feed — exhausted fully before Phase 2
- **Phase 2**: Keyword search (`"Product Manager"` across United States by default) — paginated until the candidate target is reached

For each job that passes title and location filters, it navigates to the LinkedIn job page, extracts the **actual external apply URL** (the company career site link, not just the LinkedIn URL), then runs a second dedup pass against the sheet with those real URLs.

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

Scrapes Built-in.com via a headless browser (no login required). Workflow per job:

1. Extracts title and location from the list page card (JS-evaluated to get the work model — Remote/Hybrid/In-office)
2. Filters by title (must be a PM role, not program/project/engineering) and location
3. **Pre-checks against the Google Sheet** by Built-in URL and numeric job ID — skips the detail page entirely if already applied
4. Visits the detail page to extract: company, full job description, and the **actual external apply URL** (the button that takes you to the company's career site)
5. Checks within-run fingerprint (company + title) to catch the same job appearing on multiple pages with different URLs
6. Scores by seniority tier and adds to the candidate list

After collecting up to 10 candidates, automatically calls `resume_pipeline.run_pipeline()`.

Runs standalone (and is the default when Built-in is enabled in the UI):
```bash
python3 builtin_scraper.py
```

#### URL deduplication (both scrapers)

URLs are normalized before any comparison: query parameters, URL fragments, trailing slashes, `www.` prefix, and case are all stripped. The numeric job ID is also extracted from Built-in URLs and stored as a fallback key — so slug changes in the URL don't defeat dedup. Both the `URL` column and `Linked In URL` column from the sheet are checked, as are both the apply URL and the listing URL from the scraper.

---

### Resume pipeline — `resume_pipeline.py`

Processes a list of job candidates end-to-end:

1. **Fetches the full job description** in a headless browser. For Built-in jobs the description is already pre-scraped and reused directly (no second fetch). Falls back to the LinkedIn URL if the primary URL fails.

2. **Scores the job with Claude** (0–10 scale). The scoring prompt reads your background from `background_prompt.txt` and considers: role-type alignment, seniority match, industry fit, and location.

3. **Hard disqualification checks** (applied before the score threshold):
   - If the JD states a max salary below your configured minimum → Skips tab
   - If the JD requires living near an office city not in your preferred locations → Skips tab
   - Occasional travel (≤ once/month) is not disqualifying

4. **Builds a tailored `.docx` resume** for any job scoring ≥ the score threshold:
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
    "title": "Senior Product Manager",
    "company": "Acme Corp",
    "location": "Remote",
    "url": "https://acme.com/careers/spm-role",
    "linkedin_url": "https://www.linkedin.com/jobs/view/1234567890"
  }
]
```

---

### Background prompt — `background_prompt.txt`

Your full resume context, candidate facts, style rules, and scoring calibration anchors. This is the primary source of truth Claude uses when scoring and writing resumes. Edit it from the **Background Prompt** tab in the web UI or directly in the file.

---

### Form filler — `application_filler.py`

Automatically fills and submits job application forms:

1. Opens a browser (real Chrome if installed for genuine fingerprint, Playwright Chromium as fallback) with stealth patches applied to suppress bot-detection signals
2. Navigates to the application URL; if on a job overview page, finds the actual form via link scan or Claude fallback
3. Sends the page HTML to Claude, which returns a field-by-field mapping (CSS selector → value → input type)
4. Fills every field: text, email, phone, `<select>` dropdowns, radio buttons, checkboxes, file uploads (resume), textareas, and autocomplete/typeahead fields
5. For multi-step forms, clicks Next and repeats on each page
6. Runs a pre-submit review with Claude to verify all required fields are filled
7. In **dry-run mode** (default): stops before submitting for manual review
8. In **live mode**: submits and checks for a confirmation page

**`DRY_RUN = True` is the default.** You must explicitly set `DRY_RUN = False` in the file to actually submit.

Run:
```bash
python3 application_filler.py              # reads pipeline_output.json
python3 application_filler.py --manual     # uses MANUAL_JOBS list in the file
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
#    OR run a scraper directly from the terminal:
python3 builtin_scraper.py          # Built-in scraper (includes resume pipeline)
python3 job_scraper.py --resume     # LinkedIn scraper (includes resume pipeline)

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
├── resume_pipeline.py          # Claude scorer + .docx builder + Drive uploader
├── application_filler.py       # Browser-based form filler + submitter
├── relogin.py                  # LinkedIn session re-authentication helper
│
├── background_prompt.txt       # Your resume context and scoring rules for Claude
├── config.json                 # All pipeline settings (managed via UI)
│
├── resumes/                    # Generated .docx resumes (local copies)
│
├── Dom_Cusumano_General_Resume.docx   # Base resume uploaded via UI
│
├── .env                        # API keys (git-ignored)
├── google_credentials.json     # Google service account key (git-ignored)
├── oauth_credentials.json      # Google OAuth desktop app credentials (git-ignored)
├── gdrive_token.json           # Auto-created Drive OAuth token (git-ignored)
├── linkedin_session.json       # Saved LinkedIn browser session (git-ignored)
│
├── pipeline_output.json        # Written by the pipeline, read by the form filler
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

All five are already in `.gitignore`. Verify before pushing:
```bash
git status --short | grep -E "\.env|credentials|token|session"
```

---

## Troubleshooting

**LinkedIn session expired**
Delete `linkedin_session.json` and run `job_scraper.py` again, or click **Re-authenticate** in the web UI.

**No jobs found by the LinkedIn scraper**
LinkedIn's DOM changes regularly. If no cards are being extracted, check `extract_job_from_card()` in `job_scraper.py` — the CSS selectors may need updating.

**Built-in scraper finds 0 jobs**
Check the `extract_jobs_from_page()` selector list in `builtin_scraper.py`. The `[selector]` log line shows which selector matched and how many links were found — if it shows 0, Built-in's markup has changed.

**Resume build fails with a Node error**
Make sure `npm install` was run in the project directory (not with `-g`). The `require('docx')` call resolves from the local `node_modules/`.

**Google Sheets write fails**
Verify the service account email (from `google_credentials.json` → `client_email`) has Editor access on the sheet. Both the Google Sheets API and (for Drive) Google Drive API must be enabled in your Google Cloud project.

**Google Drive upload fails on first run**
A browser window will open to complete the OAuth consent flow. After authorizing, `gdrive_token.json` is saved automatically and reused on every run after.

**Score threshold setting not taking effect**
The pipeline re-reads `config.json` at the start of each run. Saving settings in the UI writes `config.json` immediately — no restart needed.

**Duplicate jobs appearing despite being in the sheet**
Both the `URL` and `Linked In URL` columns are checked from both the Applications and Skips tabs. Built-in jobs are also matched by the numeric job ID extracted from the Built-in URL (so slug changes don't defeat dedup). If duplicates still appear, check that the service account has read access to the sheet and that the Google Sheet URL in Settings is correct.

**Form filler gets blocked by Cloudflare**
The script detects a CAPTCHA challenge page and pauses for manual solving before continuing. Make sure `HEADLESS = False` (default) so the browser is visible.
