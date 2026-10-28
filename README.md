# ECOS: Enterprise Collaboration Optimization System

Python 3.9+ backend (standard library plus SQLite) and a React/Tailwind frontend, built on your Excel workbook.

## Run it

```bash
cd backend
pip install -r requirements.txt     # openpyxl (Excel import) and pypdf (reading PDF reports)
python import_excel.py              # builds ecos.db from ../data/*.xlsx (re-run any time to reset the data)
python server.py                    # http://localhost:8000  (PORT=9000 python server.py to change)
python test_api.py                  # 42 end-to-end tests against a throwaway database
```

Open http://localhost:8000. The server serves the frontend and the API from the same address. The browser needs internet for the React, Tailwind and Babel CDN scripts.

Environment variables: `PORT`, `HOST` (default 127.0.0.1), `ECOS_DB`, `ECOS_FRONTEND`, `ECOS_TODAY` (fixes "today" for deadline checks, e.g. `2026-10-22`; defaults to the real date).

## Layout

| Path | Purpose |
|---|---|
| `backend/schema.py` | Table definitions, enums, ranges and foreign keys (column names equal the workbook headers) |
| `backend/import_excel.py` | Workbook to SQLite, with foreign keys and indexes, and an integrity check |
| `backend/server.py` | REST API, CSP checker, heuristic optimizer, reports (dashboard, analytics, employee report), static file server |
| `backend/summarizer.py` | Report text extraction (PDF/DOCX/TXT) and summarization (AI or built-in NLP) |
| `backend/.env.example` | Template for your API key and settings |
| `samples/` | A sample report as TXT, DOCX and PDF to try the summarizer |
| `backend/test_api.py` | Test suite |
| `frontend/index.html` | The ECOS UI; loads everything from `/api/bootstrap` |

## API

All responses are JSON. Errors look like `{"error":{"code","message","details"}}` with status 400, 404, 405, 409, 413, 422 or 500.

Resources: `employees`, `departments`, `projects`, `tasks`, `assignments`, `constraints`, `communications`, `meetings`, `resources`, `dependencies`.

| Method and path | What it does |
|---|---|
| `GET /api/{resource}` | List. Filters: any column (`?Status=Pending,Blocked`), `q=` text search, `sort=`, `order=asc/desc`, `page=`, `size=` (max 1000) |
| `GET /api/{resource}/{id}` | One record |
| `POST /api/{resource}` | Create. The id is generated when omitted (e.g. `E151`). Returns 201 |
| `PUT` or `PATCH /api/{resource}/{id}` | Partial update. The primary key cannot change |
| `DELETE /api/{resource}/{id}` | Delete. Returns 409 if other records reference it |
| `GET /api/bootstrap` | Every table plus `today` (what the UI loads) |
| `POST /api/optimize` `{task_id, options?}` | Ranks every employee not yet on the task, then checks the hard constraints. Returns `options` (default 3 free employees to choose from), plus `required`, `remaining` open slots, `free` (how many employees are free) and the wider `candidates` ranking |
| `POST /api/assignments/apply` `{task_id, employee_id}` | Assigns the chosen employee. `{task_id, auto:true}` assigns the single best free one. 409 if the employee is busy, a duplicate, breaks a hard constraint, or the task is full; `force:true` overrides |
| `GET /api/constraints` | Constraints with a computed `Status` of Satisfied or Violated |
| `GET /api/constraints/validate?task_id=T0001` | Constraint check for one task |
| `GET /api/dashboard`, `/api/analytics`, `/api/recommendations` | KPIs, chart data, recommendation lists |
| `POST /api/meetings/{id}/summary` | Rule-based meeting summary and action items |
| `GET /api/employees/{id}/report` | Full report for one employee: profile, summary counts, open/completed/suggested tasks with constraint status, projects, comparison with department and organization, communication, meetings and rule-based insights. 404 for an unknown id |
| `GET /api/health` | Status and row counts |

Validation covers types, enums, numeric ranges, date formats, foreign keys, task dependency rules and meeting participants.

## Report Summarizer

Open **Report Summarizer** in the sidebar, drop in a PDF, DOCX or TXT report and press *Summarize report*. You get an overview, key points, important findings, statistics, conclusions and recommendations, with Copy and Download buttons. Try the files in `samples/` first.

### Setup

1. `pip install -r requirements.txt` (this installs `pypdf`, needed for PDFs; DOCX and TXT need nothing extra).
2. That's all for the **built-in NLP summarizer**. It needs no key, runs entirely on your server and works offline.
3. Optional, for **AI summaries (Claude)**: create an API key at https://console.anthropic.com/ , then give it to the server in one of these ways before `python server.py`:
   - Copy `backend/.env.example` to `backend/.env` and put your key after `ANTHROPIC_API_KEY=`, or
   - Windows (cmd): `set ANTHROPIC_API_KEY=sk-ant-...`, PowerShell: `$env:ANTHROPIC_API_KEY="sk-ant-..."`, Mac/Linux: `export ANTHROPIC_API_KEY=sk-ant-...`
4. Restart the server. The page then offers Automatic (AI) / Built-in NLP only / AI only. AI calls are billed to your Anthropic account.

The key stays on the server: it is never sent to the browser, never returned by any endpoint and never logged.

### How it behaves

- **Accepted files:** `.pdf`, `.docx`, `.txt`. The file's real content must match its extension (a renamed file is rejected). Legacy `.doc`, password-protected PDFs, corrupt files, and scanned PDFs with no text layer return a clear message.
- **Limits:** 10 MB per file (`ECOS_MAX_UPLOAD_MB`), 500 PDF pages (`ECOS_MAX_PAGES`), 10 summaries per minute per client (`ECOS_RATE_LIMIT`), 3 summaries at once (`ECOS_SUMMARY_CONCURRENCY`). The size is checked before the upload is read, and again in the browser.
- **Privacy:** files are processed in memory only; nothing is saved to disk or the database, and file contents are never written to logs. In AI mode the text is sent to Anthropic after emails, phone numbers, card numbers and ID numbers are masked (`ECOS_REDACT=0` turns masking off). Choose *Built-in NLP only* to keep every report on your server. Do not use AI mode for documents your organisation forbids sending to third-party services.
- **Accuracy guards:** the built-in summarizer only quotes sentences that exist in the report. In AI mode the model is told to use only the report text and to treat it as untrusted data (so instructions hidden in a report are ignored); any bullet whose figures do not appear in the report is removed and you are told how many.
- **Long reports:** the built-in summarizer reads the whole report. For AI mode, reports over 120,000 characters (`ECOS_AI_MAX_CHARS`) are condensed to their most important passages first, and the page says so.
- **Fallback:** in Automatic mode, if the AI service is unreachable the built-in summary is shown with a note.
- **Model:** `ECOS_AI_MODEL` (default `claude-sonnet-5-5`). If the provider says the model name is not found, set this to a model you have access to.
- **Languages:** the built-in summarizer is tuned for English; the AI mode handles many languages.

### API

| Method and path | What it does |
|---|---|
| `GET /api/reports/config` | Limits, accepted formats and whether AI mode is available (never the key) |
| `POST /api/reports/summarize?mode=auto\|local\|ai` | Body is the raw file bytes; send the name in an `X-File-Name` header (URL-encoded). Returns `overview`, `key_points`, `findings`, `statistics`, `conclusions`, `recommendations` and `meta` |

Errors: 400 empty file or bad mode, 413 too large, 415 wrong type, 422 unreadable or no text, 429 rate limit, 502 AI provider problem, 503 busy. Example: `curl -X POST --data-binary @report.pdf -H "X-File-Name: report.pdf" "http://localhost:8000/api/reports/summarize?mode=local"`

## Employee Report

Open **Employee Report** in the sidebar (or press *Generate full report* in an employee's side panel on the Employees page), type a name or ID, and press *Generate report*. The report shows the profile and skills, headline numbers (open, completed and overdue tasks, tasks due within 14 days, average assignment score), rule-based insights, open tasks with their hard-constraint status, a comparison with the department and the whole organisation, completed and suggested tasks, projects, meetings and recent messages.

- **Export:** *Print / Save as PDF* (uses the browser's print dialog; the sidebar is hidden), *Download .txt* and *Download .csv*.
- **Definitions:** an *open task* is an Assigned record on a task that is not Completed (the same rule the optimizer uses for "free"). *Overdue* means an open task whose deadline is before today (`ECOS_TODAY` or the real date).
- **Data:** read-only. Nothing is written to the database. The insights are produced by rules, no language model is involved.

## How the AI parts work

- Heuristic search: score = 40% skill match + 25% availability (hours/day ÷ 9) + 20% workload headroom (100 − 2 × workload%) + 15% performance.
- CSP: a constraint is evaluated against the task's currently Assigned employees (candidates when optimizing). Rules the workbook does not define are interpreted as follows: Availability needs at least 5 h/day; Resource Availability fails only if every resource on the project is Unavailable; Team Capacity counts the employee's Assigned records.
- Free employee: at least 5 h/day available and no other open task (an Assigned record on a task that is not Completed). One person takes one open task at a time; a task that needs several people is staffed one person per assignment, and the page offers a fresh set of options each time.
- Meeting summaries and the AI Recommendations page are produced by rules over stored data. Only the Report Summarizer can use a language model, and only when you add an API key.

## Limits

No authentication (add it before exposing the server beyond localhost), and the built-in HTTP server is meant for development and demos. Edits are saved in `ecos.db`. SQLite serializes writes, which is fine at this scale.
