# Real Estate Comp Engine — Implementation Summary

## What This Tool Does
An automated Python tool that takes any US property address, pulls live data 
from the Propelio API, scores comparable properties, and outputs a professional 
Excel report with ARV (After Repair Value) calculation.

## How To Run
pip install -r requirements.txt
python main.py "3761 Dunhaven Rd, Dallas, TX"
python main.py "4044 Williamsburg Rd, Dallas, TX" --expand-radius
python main.py "3761 Dunhaven Rd, Dallas, TX" --new-builds

## Project Structure
main.py          — CLI entry point, orchestrates the full pipeline
scraper.py       — All Propelio API calls (login, parcel lookup, CMA data)
comp_engine.py   — Scoring logic (lot size, sqft, neighborhood, confidence)
output.py        — Excel report generation (4 sheets)
config.py        — Credentials, proxy settings, scoring tolerances
requirements.txt — Dependencies: requests, pandas, openpyxl

## Pipeline (Step by Step)
1. Address input → Propelio parcel lookup (exact/close/fuzzy match)
2. Parcel data → lot size, living area, subdivision, lat/lon, valuation
3. Lead creation → Propelio CMA pull (up to 42 nearby comps)
4. Comp scoring → each comp scored on:
   - Lot size confidence (±25% = 1.0, ±25-33% = 0.5-0.9, >33% = 0)
   - Living area filter (>40% variance = excluded)
   - Neighborhood match (exact=1.0, similar=0.9, same_city=0.8, unknown=0.8)
   - Boundary flag (major_street_possible if distance > 0.3mi)
5. Top 3 comps selected by final confidence (guarantees at least 1 sold comp)
6. ARV = average of sold comps in top pool
7. Excel report saved to output/comps_report.xlsx

## Excel Output (4 Sheets)
Sheet 1 - Subject Property   : address, neighborhood, lot size, living area, 
                                year built, valuation estimate, last sale price
Sheet 2 - Top Comps          : all scoring details, full confidence formula, 
                                boundary flag, distance, living area filter status
Sheet 3 - Neighborhood Summary: comp count, avg price per subdivision
Sheet 4 - ARV Analysis       : ARV value, low/high range, sold comps used, 
                                missing sqft flags, source pool explanation

## Key Features
- --expand-radius flag: auto-expands to nearby subdivisions when <3 exact 
  matches found; applies 0.8x cross-subdivision confidence penalty
- --new-builds flag: filters comps to year_built >= 2015
- Full confidence formula visible in every comp row (auditable)
- Missing living area explicitly flagged in ARV sheet
- Boundary detection: major_street_possible vs no_boundary_penalty
- Proxy support for all API calls

## Test Results
Address: 4044 Williamsburg Rd, Dallas, TX
- Subdivision: Glenridge Estates 3
- Lot Size: 10,139 sqft | Living Area: 1,413 sqft | Year Built: 1952
- Comps pool: 9 (after sqft filter from 42 raw)
- Expansion triggered: 0 exact subdivision matches found
- Top 3 comps: all from Glenridge Estates (similar match, conf=0.90)
- ARV: $723,333 (avg of 3 sold comps)
- ARV Range: $620,000 — $885,000

Address: 3761 Dunhaven Rd, Dallas, TX
- Lead ID: 7938989
- Pipeline runs end-to-end, Excel report generated successfully

## Dependencies
requests==2.31.0
pandas==2.0.3
openpyxl==3.1.2
