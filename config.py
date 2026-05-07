"""
Central configuration for the comp engine.

Everything that you might want to tweak without touching real code lives
here: Propelio account credentials, the proxy used for outbound API
calls, HTTP timeouts/headers, the comp scoring tolerances, and the
default Excel output location. All other modules import from this file
instead of holding their own constants.

Secrets (credentials, proxy URL) are loaded from a local ``.env`` file
via ``python-dotenv`` so they never need to be committed to source
control. See ``.env.example`` for the expected variables.
"""

import os

from dotenv import load_dotenv

load_dotenv()

# --- Propelio API credentials ---------------------------------------------
EMAIL = os.getenv("PROPELIO_EMAIL")
PASSWORD = os.getenv("PROPELIO_PASSWORD")

# Backwards-compatible aliases used by the rest of the codebase.
PROPELIO_USERNAME = EMAIL
PROPELIO_PASSWORD = PASSWORD

# --- HTTP / network knobs (used by scraper.py) ----------------------------
# Outbound proxy so Propelio sees a residential-looking source IP.
# Set PROXY_URL to an empty string in .env to disable.
PROXY_URL = os.getenv("PROXY_URL")
PROPELIO_PROXIES = (
    {"http": PROXY_URL, "https": PROXY_URL} if PROXY_URL else {}
)
HTTP_TIMEOUT_SECONDS = 30
HTTP_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36"
)

# --- Comp filtering / scoring (used by comp_engine.py) --------------------
RADIUS_MILES = 0.5
RADIUS_EXPANSION_STEPS = (0.25, 0.5, 0.75, 1.0, 1.5)
MIN_POOL_BEFORE_RADIUS_EXPAND = 5
LOT_SIZE_TOLERANCE = 0.33
LIVING_AREA_TOLERANCE = 0.40
CONFIDENCE_THRESHOLD = 0.25
CROSS_SUBDIVISION_CONF_MULT = 0.8
MAJOR_STREET_BOUNDARY_MULT = 0.9
MAJOR_STREET_MIN_DISTANCE_FT = 40
MAJOR_STREET_MAX_DISTANCE_FT = 900
NEW_BUILD_YEARS = 5
NEW_BUILDS_MIN_YEAR = 2015     # only used when --new-builds is set on the CLI
TOP_COMP_COUNT = 3
EXPAND_NEIGHBORHOOD_CONF_MULT = 0.6
EXPAND_NEIGHBORHOOD_DISTANCE_FLAG_MI = 0.3

# --- Output ---------------------------------------------------------------
OUTPUT_FILE = "output/comps_report.xlsx"
