# Auto-DTS

Auto-DTS is a headless Python CLI that submits Daily Time Sheets to BSS by calling backend HTTP endpoints directly.

It avoids browser/UI automation and performs:

- Login with CSRF token scraping
- Date-based timesheet page fetches
- Per-date CSRF extraction
- Form-urlencoded payload submission with Symfony-style bracket keys
- Safe skipping for approved/locked timesheets

## Requirements

- Python 3.8+
- Valid HR portal account

Install dependencies:

```bash
pip install -r requirements.txt
```

## Input Excel Format

Default file name: timesheet_data.xlsx

Required columns:

- Date
- Type
- Activity Name
- Regular Hours
- OT Hours

Notes:

- Date values are normalized to YYYY-MM-DD
- OT Hours can be blank and defaults to 0
- Rows are grouped by Date and submitted once per day

## Activity Mapping

The script loads activity mappings from JSON files:

- [direct_activities.json](direct_activities.json): direct activity ID mapping
- [indirect_activities.json](indirect_activities.json): category/activity mapping

Update these files to match the latest BSS master data.

You can run mapping preflight without logging in/submitting:

```bash
python app.py --validate-mappings
```

This checks:

- Mapping JSON structure and value types
- Unknown activity names in Excel that are not covered by mappings

## Run

Basic run:

```bash
python app.py
```

With options:

```bash
python app.py --excel timesheet_data.xlsx --username your.user --project 484 --workorder S26-000-12V-00
```

Dry run (build/validate payloads only, no submission):

```bash
python app.py --dry-run
```

Set credentials by environment variables:

```bash
set AUTO_DTS_USERNAME=your.user
set AUTO_DTS_PASSWORD=your_password
python app.py
```

## CLI Options

- --excel: Path to Excel file
- --username: Portal username (fallback: AUTO_DTS_USERNAME)
- --password: Portal password (fallback: AUTO_DTS_PASSWORD)
- --project: Project ID used for direct activities
- --workorder: Workorder used for direct activities
- --delay: Delay in seconds between day submissions (default 1.5)
- --dry-run: Do not submit updates
- --direct-map: Path to direct activity mapping JSON
- --indirect-map: Path to indirect activity mapping JSON
- --validate-mappings: Validate mappings and Excel activity coverage, then exit

## Submission Flow

1. GET /login and parse _csrf_token
2. POST /login_check with _username, _password, _csrf_token
3. For each date:
4. GET /dtstimesheet/?dtsdate=YYYY-MM-DD and parse oss_dtsbundle_timesheet[_token]
5. Skip when approved/locked flag is present
6. POST /dtstimesheet/update with _method=PUT and form-urlencoded payload

## Error Handling

- Automatic retries for 500/502/503/504
- Data validation for required columns and numeric hours
- Skips unknown activity mappings with warnings
- Run summary with submitted/skipped/failed counts
   