# GitHub Actions Setup

## 1) Repository settings
- Repository: `hykim-1010/monthly_report_tool_DW`
- Default branch: `main`

## 2) Required GitHub secret
- Name: `GOOGLE_SERVICE_ACCOUNT_JSON`
- Value: GA4 service account JSON file content (entire JSON text)

## 3) Required GitHub variable
- Name: `CLIENT_NAME`
- Value example: `대상웰라이프`

## 4) Workflow behavior
- File: `.github/workflows/monthly-report.yml`
- Schedule: every month, 26th at 11:00 KST (`0 2 26 * *` UTC)
- Manual run: `workflow_dispatch` (optional `start`, `end`, `report_month`)

## 5) Output
- Report files are uploaded as Actions artifact:
  - `monthly-ppt-YYYYMM`
  - `monthly-summary-<CLIENT_NAME>-YYYYMM`

## 6) Annual baseline (repo-managed)
- Path: `config/annual_baseline/<client_name>/<year>.json`
- Data shape:
  - `users_total_monthly`: 12 monthly totals (ko+en+cn)
  - `pageviews_total_monthly`: 12 monthly totals (ko+en+cn)
- Baseline is managed as a repository file (PR update, typically once per year).
- Workflow tries to restore previous month summary artifact and passes it to next run.

## 7) Security notes
- Do not commit `.env` or GA4 key files.
- Keep `config/ga4-automation-*.json` ignored.
