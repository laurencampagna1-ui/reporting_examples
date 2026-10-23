# Reporting Demo

A minimal demo of the spreadsheet-to-deck workflow used in the investor
report generator &mdash; same pattern, deliberately generic content (teams
and projects, not investors or real estate) so it's safe to show around.

**[Download the spreadsheet template](report_template.xlsx)** to get started.

## How it works

1. Upload a `.xlsx`/`.csv` file: one row per team per project.
2. The app groups rows by `team_name`, splitting each team's projects into
   "in progress" and "completed".
3. For each team, it clones `templates/base_template.pptx`, fills in the
   cover slide and the two summary tables (adding/removing table rows to
   match how many projects that team has), and converts to PDF if
   requested.
4. All the decks come back as a single `reports.zip`, one folder per team.

## Spreadsheet format

One row per team per project. Required: `team_name`, `status`
(`in_progress` or `completed`), `project_name`.

| Column | Used for | Notes |
|---|---|---|
| team_name | all | required |
| status | all | required: `in_progress` or `completed` |
| project_name | all | required |
| owner | all | |
| start_date | all | |
| target_date | in_progress | |
| percent_complete | in_progress | e.g. `0.65` or `65` or `65%` |
| notes | in_progress | |
| completed_date | completed | |
| outcome | completed | |

A sample file is at `sample_data/sample_projects.csv`.

## Template merge fields

Slides are found by content, not a fixed position:

- **Cover slide**: a shape whose text is exactly `{team_name}`, plus
  `{report_period}` (set per-run from the form field on the upload page).
- **In Progress Projects table**: identified by a header row containing
  "% Complete". Row 2 is the merge-field template (`{project_name}`,
  `{owner}`, `{start_date}`, `{target_date}`, `{percent_complete}`,
  `{notes}`) &mdash; cloned once per in-progress project.
- **Completed Projects table**: identified by a header row containing
  "Outcome". Same idea, with `{completed_date}`, `{outcome}`.

## Running locally

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
# LibreOffice is required for PDF export:
#   macOS: brew install --cask libreoffice
#   Debian/Ubuntu: sudo apt-get install libreoffice

uvicorn app.main:app --reload
```

Then open http://localhost:8000.

## Deploying

The included `Dockerfile` installs LibreOffice and the app, and runs
`uvicorn` on `$PORT`. Works as-is on Render, Fly.io, Railway, or any
container host:

```bash
docker build -t reporting-demo .
docker run -p 8000:8000 reporting-demo
```
