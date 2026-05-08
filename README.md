# Real Estate Comp Engine

## Overview
A Python tool that takes any US property address, pulls live data from the 
Propelio API, scores comparable properties, and outputs a professional Excel 
report with ARV calculation.

---

## Before You Start — Install These First

### 1. Install Python
Download: https://www.python.org/downloads/
- Download Python 3.10 or higher
- During install: ✅ check "Add Python to PATH"
- YouTube guide: https://www.youtube.com/watch?v=YYXdXT2l-Gg

### 2. Install Git
Download: https://git-scm.com/downloads
- Download Git for Windows
- Use all default settings during install
- Video guide: https://www.loom.com/share/6847c474d4234f478b95e09857b208b0

---

## Setup Walkthrough (Video)
https://www.loom.com/share/46c66207a45148a4b60482faaa2dab94

---

## Step by Step Setup

### Step 1 — Clone the Repository
```bash
git clone https://github.com/rutavix/propelio-comp-engine.git
cd propelio-comp-engine
```

### Step 2 — Create Virtual Environment
```bash
python -m venv .venv
```

### Step 3 — Activate Virtual Environment
```bash
.venv\Scripts\activate
```
You will see (.venv) at the start of the line — this means you are ready.

### Step 4 — Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 5 — Setup Credentials
```bash
copy .env.example .env
notepad .env
```
Fill in your credentials:
```env
PROPELIO_EMAIL=your_email_here
PROPELIO_PASSWORD=your_password_here
PROXY_URL=http://user:pass@host:port
```
Save and close the file.

---

## Use Cases

### Use Case 1 — Standard Run
**Purpose:** Run a comp analysis on any property address.

**Command:**
```bash
python main.py "3761 Dunhaven Rd, Dallas, TX"
```

**Output:** Excel report saved to output/comps_report.xlsx with ARV, 
top 3 comps, and neighborhood summary.

**Video:** https://www.loom.com/share/e3ef18c3ece4426c952191d1f3f6980b

---

### Use Case 2 — Expand Radius
**Purpose:** When fewer than 3 exact subdivision matches are found, 
automatically pulls comps from nearby neighborhoods with a 0.8x 
confidence penalty applied.

**Command:**
```bash
python main.py "4044 Williamsburg Rd, Dallas, TX" --expand-radius
```

**Output:** Wider comp pool, cross-subdivision penalty visible in 
confidence formula column.

**Video:** https://www.loom.com/share/6c0fe12caaa34e359e7183b4152fa672

---

### Use Case 3 — New Builds Filter
**Purpose:** Filter comps to new construction only (built 2015 or later). 
Useful when the subject property is a new build.

**Command:**
```bash
python main.py "3761 Dunhaven Rd, Dallas, TX" --new-builds
```

**Output:** Only comps with year_built >= 2015 included in scoring pool.

**Video:** https://www.loom.com/share/6ff3d202a8004f3abbe142f5e04d1b15

---

## Excel Report — 4 Sheets

| Sheet | Contents |
|---|---|
| Subject Property | Address, neighborhood, lot size, living area, year built, valuation |
| Top Comps | Top 3 comps with full confidence formula per comp |
| Neighborhood Summary | Avg price and comp count per subdivision |
| ARV Analysis | Final ARV, low/high range, sold comps used |

---

## Dependencies
```txt
requests==2.31.0
pandas==2.0.3
openpyxl==3.1.2
python-dotenv
```
