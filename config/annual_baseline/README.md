# Annual Baseline

Store previous-year monthly totals for each client.

Path format:
- `config/annual_baseline/<client_name>/<year>.json`

JSON schema:
```json
{
  "users_total_monthly": [12 monthly totals],
  "pageviews_total_monthly": [12 monthly totals]
}
```

Notes:
- Values must be total monthly sums (ko+en+cn), not per-language values.
- Keep exactly 12 integer values in each array.
- Update once per year via PR.
