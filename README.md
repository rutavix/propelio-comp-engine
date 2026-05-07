# Real Estate Comp Engine

Pulls nearby property records from Propelio for a subject address, scores them
by lot-size similarity, and writes a 3-sheet Excel report.

## Project layout

```
real-estate-comps/
├── config.py        # credentials, proxy/HTTP settings, scoring tolerances
├── scraper.py       # Propelio API client (login, CMA, lead details)
├── comp_engine.py   # confidence scoring, filtering, neighborhood rollup
├── output.py        # Excel writer (subject / comps / neighborhood / ARV)
├── main.py          # CLI entry point
├── requirements.txt
└── output/          # generated reports land here
```

## Setup

1. Install Python 3.10+.
2. Create a virtualenv and install dependencies:

   ```bash
   python -m venv .venv
   .venv\Scripts\activate          # Windows
   pip install -r requirements.txt
   ```

3. Open `config.py` and fill in your Propelio credentials (or pass
   `--username` / `--password` on the command line).

## Usage

```bash
python main.py "123 Main St, Austin, TX"

# custom radius and output location
python main.py "123 Main St, Austin, TX" --radius 0.75 --output ./output/report.xlsx

# subject lot size when the API does not return it (sqft)
python main.py "3761 Dunhaven Rd, Dallas, TX" --lot-size 7414

# verbose logging
python main.py "123 Main St, Austin, TX" -v
```

The console prints a quick summary; the full report lands at
`output/comps_report.xlsx` by default (or a timestamped sibling if that file
is open in Excel).

## Tunables (in `config.py`)

| Setting                | Default | Meaning                                         |
|------------------------|---------|-------------------------------------------------|
| `RADIUS_MILES`         | `0.5`   | Search radius around the subject address.       |
| `LOT_SIZE_TOLERANCE`   | `0.33`  | Hard cutoff for lot-size deviation.             |
| `CONFIDENCE_THRESHOLD` | `0.25`  | Minimum confidence score required to keep comp. |
| `NEW_BUILD_YEARS`      | `5`     | Properties newer than this are excluded.        |
| `TOP_COMP_COUNT`       | `3`     | Number of comps returned per run.               |

## Confidence scoring

| Lot-size deviation | Confidence              |
|--------------------|-------------------------|
| ≤ 25%              | 1.0                     |
| 25% – 33%          | 0.9 → 0.5 (linear)      |
| > 33%              | 0.0 (filtered out)      |

## Notes

- The scraper talks to Propelio's HTTP API (`/login`, `/parcels/v1/...`,
  `/legacy/leads/withaddress`, `/legacy/cma`). If responses change, check
  logs for payload keys and adjust parsing in `scraper.py`.
- The outbound proxy lives in `config.PROPELIO_PROXIES`. Set it to `{}` to
  disable.
