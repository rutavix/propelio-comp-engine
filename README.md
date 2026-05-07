# Real Estate Comp Engine

## Overview
A Python tool that takes any US property address, pulls live data from the Propelio API, scores comparable properties, and outputs a professional Excel report with ARV calculation.

## How to Install & Setup

### Step 1 — Clone the Repository
```bash
git clone https://github.com/rutavix/propelio-comp-engine.git
cd propelio-comp-engine
```

### Step 2 — Open Terminal in the Project Folder
- Press `Win + R` → type `cmd` → press Enter
- Type: `F:`
- Type: `cd \real-estate-comps`

### Step 3 — Activate Virtual Environment
```bash
.venv\Scripts\activate
```
You will see `(.venv)` at the start of the line when ready.

### Step 4 — Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 5 — Setup Credentials
- Copy `.env.example` and rename it to `.env`
- Fill in your Propelio email, password, and proxy URL

---

## Use Cases

### Use Case 1 — Standard Run
**Purpose:** Run a comp analysis on any property address.

**Command:**
```bash
python main.py "3761 Dunhaven Rd, Dallas, TX"
```

**Output:** Excel report saved to `output/comps_report.xlsx` with ARV, top 3 comps, and neighborhood summary.

**Video:** https://www.loom.com/share/1fe17ff756224a0f96e55d7320d247b0

---

### Use Case 2 — Expand Radius
**Purpose:** When fewer than 3 exact subdivision matches are found, automatically pulls comps from nearby neighborhoods with a 0.8x confidence penalty applied.

**Command:**
```bash
python main.py "4044 Williamsburg Rd, Dallas, TX" --expand-radius
```

**Output:** Wider comp pool, cross-subdivision penalty visible in confidence formula column.

**Video:** https://www.loom.com/share/d28b3922d08e45b4a7ea33093c64025e

---

### Use Case 3 — New Builds Filter
**Purpose:** Filter comps to new construction only (built 2015 or later). Useful when the subject property is a new build.

**Command:**
```bash
python main.py "3761 Dunhaven Rd, Dallas, TX" --new-builds
```

**Output:** Only comps with `year_built >= 2015` included in scoring pool.

**Video:** https://www.loom.com/share/16e47ea956d24972a1bc9c30027aa456

---

## Excel Report — 4 Sheets

| Sheet | Contents |
|---|---|
| Subject Property | Address, neighborhood, lot size, living area, year built, valuation |
| Top Comps | Top 3 comps with full confidence formula per comp |
| Neighborhood Summary | Avg price and comp count per subdivision |
| ARV Analysis | Final ARV, low/high range, sold comps used |

## Dependencies
```
requests==2.31.0
pandas==2.0.3
openpyxl==3.1.2
python-dotenv
```
