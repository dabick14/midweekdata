# Midweek Data Automation

This repository includes an automated updater that builds a fresh midweek workbook from the Synago GraphQL API.

Generated detail sheets use this column layout: Governorship, Governor, No. Of Bacentas, Att, Income(GHS), No. Of Services, Services/Bacentas, Services Not Held, Comment.

## Files

- `update_midweek_data.py`: Main script that logs in, queries stream governorships, creates a new workbook, creates one sheet per council name, creates a Summary sheet, and writes `Midweek_Data_updated.xlsx`.
- `.github/workflows/update_midweek.yml`: Weekly GitHub Actions workflow (Wednesdays, 08:00 UTC) plus manual trigger.
- `.github/workflows/update_midweek.yml`: Weekly GitHub Actions workflow (Saturdays, 18:00 UTC) plus manual trigger.

## Prerequisites

- Python 3.10+
- `requests`
- `openpyxl`

## Environment Variables

Set these before running:

- `FLC_EMAIL`: API login email
- `FLC_PASSWORD`: API login password
- `STREAM_ID_1`: Stream ID for Colossians (default in script if unset)
- `STREAM_ID_2`: Stream ID for Galatians (default in script if unset)
- `STREAM_ID_3`: Stream ID for Jesus Night (default in script if unset)

For local runs, a `.env` file is supported automatically (no extra package needed):

1. Copy `.env.example` to `.env` (already created in this repo)
2. Fill your real values in `.env`
3. Run the script normally

The script loads `.env` first, then reads environment variables from the shell.

## Local Run

Install dependencies:

```bash
pip install requests openpyxl
```

Run the updater:

```bash
python update_midweek_data.py --input Midweek_Data.xlsx --output Midweek_Data_updated.xlsx
```

`--input` is kept only for backward compatibility and is ignored.

If successful, the script prints row counts per sheet and writes a new workbook file.

Note: GraphQL requests are sent with `Authorization: Bearer <accessToken>`.

## GitHub Actions Setup

1. Add repository secrets:
   - `FLC_EMAIL`
   - `FLC_PASSWORD`
   - `STREAM_ID_1`
   - `STREAM_ID_2`
   - `STREAM_ID_3`
   - `GOOGLE_SERVICE_ACCOUNT_JSON`
   - `GOOGLE_DRIVE_FOLDER_ID`
2. Set `GOOGLE_SERVICE_ACCOUNT_JSON` to the full JSON contents of your Google service account key.
3. Set `GOOGLE_DRIVE_FOLDER_ID` to the destination Google Drive folder ID.
4. Use a folder in a Shared Drive and share it with the service account email as Content manager (or Editor).
5. Service accounts do not have personal Drive storage quota, so uploads to personal "My Drive" folders can fail with `storageQuotaExceeded`.
6. Enable actions and run the workflow manually once via `workflow_dispatch` to validate.

The scheduled workflow runs every Saturday at 18:00 UTC, uploads a dated workbook copy to the configured Google Drive folder, and commits workbook updates back to the repository when changes are detected.
