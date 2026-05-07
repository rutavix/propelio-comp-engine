# Real Estate Comp Engine

Neighborhood-aware real estate comp and ARV engine that pulls property data from Propelio, scores comps with explainable confidence logic, and exports a client-ready Excel report.

## What This Tool Does

- Resolves a subject property from a full address.
- Pulls candidate comps from Propelio.
- Scores comps using lot fit, living-area fit, neighborhood match, and boundary context.
- Selects top comps and calculates ARV.
- Exports a multi-sheet Excel workbook.

## Quick Start

### 1) Prerequisites

- Python 3.10+
- Propelio account credentials
- Dependencies from `requirements.txt`

### 2) Install

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

### 3) Run

```bash
python main.py "4044 Williamsburg Rd, Dallas, TX" --username "YOUR_EMAIL" --password "YOUR_PASSWORD"
```

If `output/comps_report.xlsx` is locked, the app automatically writes a fallback file:

- `output/comps_report_YYYYMMDD_HHMMSS.xlsx`

## Common Commands

Standard run:

```bash
python main.py "4044 Williamsburg Rd, Dallas, TX" --username "YOUR_EMAIL" --password "YOUR_PASSWORD"
```

Expand to nearby neighborhoods:

```bash
python main.py "4044 Williamsburg Rd, Dallas, TX" --username "YOUR_EMAIL" --password "YOUR_PASSWORD" --expand-radius
```

Force subject lot size:

```bash
python main.py "3761 Dunhaven Rd, Dallas, TX" --username "YOUR_EMAIL" --password "YOUR_PASSWORD" --lot-size 7414
```

Restrict to new-build comps:

```bash
python main.py "4044 Williamsburg Rd, Dallas, TX" --username "YOUR_EMAIL" --password "YOUR_PASSWORD" --new-builds
```

## CLI Options

| Option | Purpose |
|---|---|
| `--username` | Propelio login username |
| `--password` | Propelio login password |
| `--expand-radius` | Include nearby neighborhoods when exact matches are too few |
| `--lot-size <sqft>` | Override subject lot size |
| `--new-builds` | Restrict comp pool to new-build properties |

## Output Workbook

The workbook is written to `output/` and includes:

- `Subject Property`: Core subject details and valuation context.
- `Top Comps`: Selected comps, confidence, boundary flags, and formula explanation.
- `Neighborhood Summary`: Neighborhood-level comp rollup.
- `ARV Analysis`: ARV value, range, and disclosure fields.

## Scoring and Selection Behavior

- Living-area outlier filter: excludes comps with living-area variance above 40%.
- Neighborhood handling:
  - Exact match preferred.
  - Normalized similar names treated as similar match.
  - Optional neighborhood expansion via `--expand-radius`.
- Boundary flags and penalties are included in confidence logic where applicable.
- Missing living area is disclosed in output rather than silently hidden.

## Edge Cases

| Case | Expected Behavior |
|---|---|
| Invalid / unmatched address | Run fails with clear parcel-match error and non-zero exit code |
| Missing credentials | Run fails with credentials-not-configured error |
| No qualifying comps | Workbook still writes with header-only comp sections |
| Locked Excel output | Falls back to timestamped output filename |
| Missing comp living area | Comp is disclosed as missing sqft; ARV notes include missing-sqft indicator |

## Troubleshooting

- If address lookup fails, use full address format: street number, street name, city, state.
- If no comps qualify, remove restrictive flags and retry.
- If exact subdivision coverage is thin, retry with `--expand-radius`.
- If lot size from source is unreliable, run again with `--lot-size`.

## Operational Workflow

1. Run the command with full address and credentials.
2. Open the generated workbook in `output/`.
3. Review `Top Comps` and `ARV Analysis` first.
4. Validate boundary flags, confidence formula, and disclosures.
5. Re-run with adjusted flags if coverage is thin.

## Security Note

For production usage, prefer storing credentials in environment variables or a secure secret manager instead of plain command history.
