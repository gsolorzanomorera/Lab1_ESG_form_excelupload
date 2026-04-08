# ══════════════════════════════════════════════════════════════════════════════
#  GHG CARBON INVENTORY — McKinsey Style
#  Streamlit web application
#  Based on: GHG Protocol Corporate Standard
# ══════════════════════════════════════════════════════════════════════════════

# ── IMPORTS ───────────────────────────────────────────────────────────────────
# streamlit is the web-app framework — every widget, layout, and page element
# comes from this library. We alias it "st" so we can write st.button() etc.
import streamlit as st

# pandas is a data-analysis library. We use it here mainly for its DataFrame
# type (a table) when building the CSV export.
import pandas as pd

# json is Python's built-in module for reading and writing JSON files.
# We use it to save/restore all the user's inputs as a .json file.
import json

# datetime gives us the current date and time so we can stamp the report.
from datetime import datetime


# ── PAGE CONFIGURATION ────────────────────────────────────────────────────────
# st.set_page_config() MUST be the very first Streamlit call in the script.
# It controls the browser tab title, the favicon icon, the page layout,
# and whether the sidebar starts open or collapsed.
st.set_page_config(
    page_title="GHG Carbon Inventory",   # text shown in the browser tab
    page_icon="◈",                        # favicon shown in the browser tab
    layout="wide",                        # "wide" uses the full browser width
    initial_sidebar_state="expanded",     # sidebar is open when the app loads
)


# ══════════════════════════════════════════════════════════════════════════════
#  McKINSEY-STYLE CSS
#  Streamlit lets us inject raw HTML/CSS using st.markdown() with
#  unsafe_allow_html=True. Everything inside <style>...</style> is pure CSS
#  that overrides Streamlit's default look and feel.
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>

/* ── GOOGLE FONTS ────────────────────────────────────────────────────────────
   We import three typefaces from Google Fonts:
   - Playfair Display  → editorial serif, used for headings and KPI numbers
   - Source Sans 3     → clean humanist sans-serif, used for body text
   - Source Code Pro   → monospace, used for numbers and emission-factor chips
   The "wght@400;600;700" syntax requests specific font weights only,
   so the browser downloads smaller font files. */
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600;700&family=Source+Sans+3:wght@300;400;500;600&family=Source+Code+Pro:wght@400;500&display=swap');

/* ── CSS CUSTOM PROPERTIES (design tokens) ───────────────────────────────────
   :root defines variables available everywhere in the stylesheet.
   Using variables means changing one value here updates every place it's used.
   This is the McKinsey palette: white backgrounds, deep navy, one blue accent. */
:root {
    --white:     #FFFFFF;   /* pure page background                          */
    --off-white: #F7F7F5;   /* very slightly warm white — used in inputs     */
    --rule:      #E0DDD8;   /* light warm grey — horizontal divider lines    */
    --light-grey:#F0EEE9;   /* slightly darker warm grey — table total rows  */
    --mid-grey:  #9A9591;   /* medium grey — labels and captions             */
    --body:      #1A1A1A;   /* near-black — main body text                   */
    --navy:      #002B5B;   /* McKinsey deep navy — headings, borders, sidebar*/
    --accent:    #005EB8;   /* McKinsey blue — links, chips, progress bars   */
    --accent-lt: #E8F0FA;   /* very light blue — insight box background      */
    --teal:      #00857C;   /* teal green — used for positive/down deltas    */
    --amber:     #C9600A;   /* amber orange — used for negative/up deltas    */
    --rule-h:    2px solid #002B5B; /* the thick navy rule under h1 headings */
}

/* ── GLOBAL PAGE BACKGROUND ──────────────────────────────────────────────────
   Streamlit wraps content in several nested div elements. We target all of
   them to ensure the background is white everywhere — not Streamlit's default
   grey. !important is required to override Streamlit's own inline styles. */
html, body,
[data-testid="stAppViewContainer"],
[data-testid="stMain"] {
    background: var(--white) !important;
    color: var(--body);                    /* default text colour             */
    font-family: 'Source Sans 3', sans-serif; /* default body font           */
    font-size: 15px;                       /* base font size                  */
}

/* ── SIDEBAR ─────────────────────────────────────────────────────────────────
   The sidebar uses McKinsey's navy as its background — a strong contrast
   with the white main area. The selector targets Streamlit's sidebar element
   by its internal data-testid attribute. */
[data-testid="stSidebar"] {
    background: var(--navy) !important;
    border-right: none;   /* remove Streamlit's default sidebar border line  */
}

/* All text elements inside the sidebar inherit a light blue-grey colour
   so they're readable against the dark navy background. */
[data-testid="stSidebar"] * { color: #C8D8EE !important; }

/* The horizontal rule (divider) inside the sidebar uses a semi-transparent
   white so it's subtle rather than stark. */
[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.15) !important; }

/* st.metric() produces a value (stMetricValue) and a label (stMetricLabel).
   In the sidebar we make values bright white and monospace (numbers need
   fixed-width font so they align when they change), and labels a lighter blue. */
[data-testid="stSidebar"] [data-testid="stMetricValue"] {
    color: #FFFFFF !important;
    font-family: 'Source Code Pro', monospace !important;
    font-size: 1.05rem !important;
}
[data-testid="stSidebar"] [data-testid="stMetricLabel"] {
    color: #7DA3CC !important;   /* lighter blue for label text               */
    font-size: 11px !important;
}

/* ── HEADINGS ────────────────────────────────────────────────────────────────
   h1 is used for page titles. The signature McKinsey treatment:
   - Playfair Display serif for authority and elegance
   - Deep navy colour
   - A 2px navy bottom border (like a McKinsey exhibit title line)
   - Slight negative letter-spacing makes large type look more refined */
h1 {
    font-family: 'Playfair Display', serif;
    font-weight: 700;
    font-size: 2rem;
    color: var(--navy);
    border-bottom: var(--rule-h);   /* the thick navy underline              */
    padding-bottom: 12px;
    margin-bottom: 4px;
    letter-spacing: -0.02em;        /* tighten tracking for large display type*/
}

/* h2 is used for section sub-headings within a page.
   Lighter weight than h1, with a thinner grey rule underneath. */
h2 {
    font-family: 'Playfair Display', serif;
    font-weight: 600;
    font-size: 1.2rem;
    color: var(--navy);
    margin-top: 28px;
    margin-bottom: 4px;
    border-bottom: 1px solid var(--rule);  /* thin grey rule                 */
    padding-bottom: 6px;
}

/* ── EYEBROW LABEL ───────────────────────────────────────────────────────────
   An "eyebrow" is a small uppercase label that sits above a large heading.
   Common in McKinsey presentations: e.g. "STEP 01" above the page title.
   display:block ensures it occupies its own line even though it's a <span>. */
.eyebrow {
    font-family: 'Source Sans 3', sans-serif;
    font-size: 11px;
    font-weight: 600;
    text-transform: uppercase;     /* force all-caps regardless of input text */
    letter-spacing: 0.15em;        /* wide tracking is standard for eyebrows  */
    color: var(--accent);          /* McKinsey blue                            */
    margin-bottom: 4px;
    display: block;
}

/* Subtitle / caption text beneath a page title. Smaller, grey, short. */
.page-caption {
    font-size: 13px;
    color: var(--mid-grey);
    margin-top: 2px;
    margin-bottom: 20px;
}

/* ── KPI CARD ────────────────────────────────────────────────────────────────
   KPI cards show a single metric with label, large number, and unit.
   McKinsey style: NO box, NO shadow, NO background — just a thick top border.
   The 3px solid navy top-border is the only visual frame. */
.mck-kpi {
    border-top: 3px solid var(--navy);
    padding: 16px 0 12px 0;
}

/* Small uppercase label above the number (e.g. "SCOPE 1 — DIRECT") */
.mck-kpi-label {
    font-size: 11px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.12em;
    color: var(--mid-grey);
    margin-bottom: 6px;
}

/* The large number itself — Playfair Display makes it feel weighty and precise */
.mck-kpi-value {
    font-family: 'Playfair Display', serif;
    font-size: 2.3rem;
    font-weight: 700;
    color: var(--navy);
    line-height: 1;     /* tight line height so the number sits close to label*/
    margin-bottom: 4px;
}

/* Small unit label below the number (e.g. "tCO₂e") */
.mck-kpi-unit { font-size: 12px; color: var(--mid-grey); }

/* Year-on-year delta line below the unit (e.g. "↓ 8.2% vs prior year") */
.mck-kpi-delta { font-size: 12px; font-weight: 600; margin-top: 4px; }

/* A decrease (↓) in emissions is GOOD — coloured teal */
.delta-neg { color: var(--teal); }

/* An increase (↑) in emissions is BAD — coloured amber */
.delta-pos { color: var(--amber); }

/* ── INSIGHT BOX ─────────────────────────────────────────────────────────────
   A highlighted callout box used to summarise a key finding.
   The left blue border is the McKinsey "sidebar quote" styling.
   Light blue background ensures it's readable without being distracting. */
.insight-box {
    background: var(--accent-lt);       /* very light blue background         */
    border-left: 4px solid var(--accent); /* solid blue left border           */
    padding: 12px 16px;
    margin: 12px 0;
    font-size: 13px;
    line-height: 1.5;
}
/* Bold text inside insight boxes uses navy so it stands out clearly */
.insight-box strong { color: var(--navy); }

/* ── DATA TABLE ──────────────────────────────────────────────────────────────
   McKinsey exhibit tables use:
   - Double top border (2px navy) to frame the header
   - Single bottom borders on each row
   - Double bottom border (2px navy) on the last row to close the table
   - No vertical lines (border-collapse:collapse removes them)
   - Uppercase column headers at 11px                                        */
.mck-table { width:100%; border-collapse:collapse; font-size:13px; margin-top:8px; }

/* Header row: thick navy above, thinner navy below */
.mck-table thead tr { border-top:2px solid var(--navy); border-bottom:1px solid var(--navy); }

/* Column header cells: small, uppercase, navy, wide tracking */
.mck-table thead th {
    padding:8px 12px; text-align:left;
    font-size:11px; font-weight:600; text-transform:uppercase;
    letter-spacing:0.08em; color:var(--navy);
}

/* Numeric columns align right — financial table convention */
.mck-table thead th.num { text-align:right; }

/* Each data row has a light grey bottom rule between it and the next row */
.mck-table tbody tr { border-bottom:1px solid var(--rule); }

/* Last row has a closing 2px navy rule — mirrors the top header rule */
.mck-table tbody tr:last-child { border-bottom:2px solid var(--navy); }

/* Standard cell padding */
.mck-table tbody td { padding:8px 12px; }

/* Numeric data cells: right-aligned and monospace so digits stack vertically */
.mck-table tbody td.num {
    text-align:right;
    font-family:'Source Code Pro', monospace;
    font-size:12px;
}

/* The "Total" row gets a light grey background to distinguish it */
.mck-table tr.total-row { background:var(--light-grey); }

/* Total row text uses navy and bold weight */
.mck-table tr.total-row td { color:var(--navy); font-weight:600; }

/* ── EMISSION FACTOR CHIP ────────────────────────────────────────────────────
   A small pill-shaped badge that shows an emission factor next to an input.
   E.g. "53.06 kgCO₂/MMBtu". Monospace font keeps the numbers aligned.
   Light blue background and border match the accent colour family. */
.ef-chip {
    display:inline-block;
    font-family:'Source Code Pro', monospace;
    font-size:11px;
    color:var(--accent);
    background:var(--accent-lt);
    padding:2px 7px;
    border-radius:2px;          /* slightly rounded corners — very subtle     */
    border:1px solid #C0D4EF;   /* light blue border                          */
}

/* ── PROGRESS BAR ────────────────────────────────────────────────────────────
   We build our own progress bar in HTML because Streamlit's default one
   has styles we can't easily override. Three parts:
   - mck-progress-label: the text row with left label and right percentage
   - mck-progress-track: the grey background rail
   - mck-progress-fill:  the blue filled portion (width set inline per value) */
.mck-progress-label {
    display:flex;                  /* puts label and percentage on one line     */
    justify-content:space-between; /* label left, percentage right             */
    font-size:12px;
    color:var(--mid-grey);
    margin-bottom:4px;
}
.mck-progress-track { background:var(--light-grey); height:6px; width:100%; }
.mck-progress-fill  { background:var(--accent); height:6px; }

/* ── INPUT FIELDS ────────────────────────────────────────────────────────────
   Override Streamlit's default input styling so number inputs and text inputs
   match our McKinsey palette: off-white background, warm grey border, navy
   text, and a subtle square border-radius (2px instead of Streamlit's 8px). */
[data-testid="stNumberInput"] input,
[data-testid="stTextInput"] input {
    background: var(--off-white) !important;
    border: 1px solid var(--rule) !important;
    border-radius: 2px !important;
    color: var(--body) !important;
    font-family: 'Source Code Pro', monospace !important; /* numbers in mono  */
    font-size: 13px !important;
}

/* When the user clicks into an input field, the border turns blue (accent).
   box-shadow:none removes Streamlit's default glow ring on focus. */
[data-testid="stNumberInput"] input:focus,
[data-testid="stTextInput"] input:focus {
    border-color: var(--accent) !important;
    box-shadow:none !important;
}

/* Selectbox (dropdown) gets the same off-white / grey-border treatment */
[data-testid="stSelectbox"] > div > div {
    background: var(--off-white) !important;
    border: 1px solid var(--rule) !important;
    border-radius: 2px !important;
}

/* All input labels: slightly smaller than body text, dark grey, medium weight */
label { font-size:13px !important; color:#444 !important; font-weight:500 !important; }

/* ── BUTTONS ─────────────────────────────────────────────────────────────────
   Primary action buttons: solid navy background, white text, square corners,
   uppercase text with wide tracking — classic McKinsey CTA style.
   On hover they switch to the lighter blue accent. */
.stButton > button {
    background: var(--navy) !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 2px !important;
    font-family: 'Source Sans 3', sans-serif !important;
    font-size:13px !important;
    font-weight:600 !important;
    text-transform:uppercase !important;
    letter-spacing:0.06em !important;
    padding:8px 20px !important;
}
.stButton > button:hover { background: var(--accent) !important; }

/* ── EXPANDER ────────────────────────────────────────────────────────────────
   st.expander() creates a collapsible section. We give it a thin grey border
   and off-white background to match our input field style.
   The summary (the clickable header) is given navy bold text. */
[data-testid="stExpander"] {
    border: 1px solid var(--rule) !important;
    border-radius: 2px !important;
    background: var(--off-white) !important;
}
[data-testid="stExpander"] summary {
    font-weight: 600 !important;
    font-size: 13px !important;
    color: var(--navy) !important;
}

/* ── DOWNLOAD BUTTONS ────────────────────────────────────────────────────────
   Download buttons use an "outlined" style: transparent background with a
   navy border and navy text. On hover the colours invert (navy bg, white text).
   This distinguishes download buttons from primary action buttons. */
[data-testid="stDownloadButton"] button {
    background: transparent !important;
    border: 1.5px solid var(--navy) !important;
    color: var(--navy) !important;
    border-radius: 2px !important;
    font-size:12px !important;
    font-weight:600 !important;
    text-transform:uppercase !important;
    letter-spacing:0.06em !important;
}
[data-testid="stDownloadButton"] button:hover {
    background: var(--navy) !important;
    color: #FFFFFF !important;
}

/* ── DIVIDER LINE ────────────────────────────────────────────────────────────
   st.divider() renders an <hr> element. We make it a thin warm-grey line
   with generous vertical margins to create breathing room between sections. */
hr { border:none; border-top:1px solid var(--rule) !important; margin:20px 0 !important; }

/* Hide Streamlit's default footer ("Made with Streamlit") and hamburger menu */
footer { display:none; }
#MainMenu { display:none; }

</style>
""", unsafe_allow_html=True)
# The unsafe_allow_html=True argument is required any time we inject raw HTML.
# Without it, Streamlit would escape the tags and print them as plain text.


# ══════════════════════════════════════════════════════════════════════════════
#  SESSION STATE INITIALISATION
#
#  Streamlit reruns the entire script from top to bottom on every user
#  interaction (e.g., typing in an input, clicking a button).
#  st.session_state is a dictionary that persists values across reruns.
#  Without it, every rerun would reset all values to zero.
#
#  init_state() sets a default value for every variable ONCE (the first time
#  the app loads). On subsequent reruns the "if k not in st.session_state"
#  check prevents overwriting values the user has already entered.
# ══════════════════════════════════════════════════════════════════════════════
def init_state():
    # defaults is a plain Python dictionary of all the variables we need.
    # Keys are string names; values are the starting defaults.
    defaults = {
        # ── Company / reporting context ────────────────────────────────────
        "company_name":   "Your Company",  # shown in headers and report
        "industry":       "Technology",    # used for context notes
        "country":        "United States", # used to suggest a grid EF
        "reporting_year": 2025,            # the year being reported on

        # ── Financial metrics (for intensity calculations) ─────────────────
        "revenue_musd": 0.0,  # revenue in USD millions (e.g. $2.4B = 2400)
        "employees":    0,    # full-time equivalent headcount

        # ── GWP factors (Global Warming Potential) — IPCC AR6 defaults ─────
        # GWP converts non-CO₂ gases to CO₂ equivalents.
        # These are the AR6 (2021) 100-year values. Users can override them.
        "gwp_ch4_fossil": 29.8,    # methane from fossil fuels
        "gwp_n2o":        273.0,   # nitrous oxide
        "gwp_hfc134a":    1526.0,  # HFC-134a refrigerant
        "gwp_sf6":        25200.0, # SF₆ electrical equipment

        # ── Scope 1 inputs: Stationary Combustion ─────────────────────────
        "s1_natgas_mmbtu":   0.0,  # natural gas burned, in MMBtu
        "s1_diesel_litres":  0.0,  # diesel burned in boilers etc., litres
        "s1_lpg_litres":     0.0,  # LPG burned, litres
        "s1_coal_shorttons": 0.0,  # coal burned, short tons

        # ── Scope 1 inputs: Mobile Combustion (fleet) ─────────────────────
        "s1_gasoline_litres":      0.0,  # gasoline used by company cars
        "s1_diesel_fleet_litres":  0.0,  # diesel used by trucks/vans
        "s1_jet_litres":           0.0,  # aviation fuel for company aircraft

        # ── Scope 1 inputs: Fugitive emissions (refrigerant leaks) ─────────
        "s1_hfc134a_kg": 0.0,  # HFC-134a leaked, kg
        "s1_hfc410a_kg": 0.0,  # HFC-410A leaked, kg
        "s1_sf6_kg":     0.0,  # SF₆ leaked from electrical switchgear, kg

        # ── Scope 2 inputs: Purchased electricity ─────────────────────────
        "s2_elec_mwh":   0.0,    # total electricity consumed, MWh
        "s2_grid_ef":    386.0,  # location-based grid EF, kgCO₂e/MWh (US default)
        "s2_market_ef":  0.0,    # market-based supplier EF, kgCO₂e/MWh
        "s2_recs_mwh":   0.0,    # MWh covered by RECs or PPAs (offsets MB)
        "s2_steam_gj":   0.0,    # purchased steam/heat/cooling, GJ

        # ── Scope 3 inputs: Value chain categories ─────────────────────────
        "s3_cat1_spend":      0.0,   # Cat 1: annual spend in $000s USD
        "s3_cat1_ef":         0.35,  # Cat 1: EEIO factor kgCO₂e per $
        "s3_cat3_elec_mwh":   0.0,   # Cat 3: electricity for T&D loss calc
        "s3_cat6_air_km":     0.0,   # Cat 6: total air travel km
        "s3_cat6_rail_km":    0.0,   # Cat 6: total rail travel km
        "s3_cat7_km_per_emp": 0.0,   # Cat 7: avg commute km per employee/year
        "s3_cat11_units":     0.0,   # Cat 11: product-years in use
        "s3_cat11_ef":        0.0,   # Cat 11: lifecycle EF kgCO₂e per unit

        # ── Prior year actuals (for year-on-year comparison) ───────────────
        "prior_s1":   0.0,  # last year's Scope 1 total
        "prior_s2mb": 0.0,  # last year's Scope 2 market-based total
        "prior_s3":   0.0,  # last year's Scope 3 total

        # ── Reduction target settings ──────────────────────────────────────
        "target_year":          2030,   # year by which target must be hit
        "target_reduction_pct": 50.0,   # % reduction vs baseline
        "target_baseline":      0.0,    # baseline emissions in tCO₂e

        # ── Industry benchmark (for peer comparison on Dashboard) ──────────
        "benchmark_revenue_intensity": 0.0,  # tCO₂e per $M revenue
    }

    # Loop over every key-value pair in defaults.
    # Only set the value if it doesn't already exist in session_state.
    # This means the function is safe to call on every rerun — it won't
    # overwrite values the user has already entered.
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

# Call init_state() immediately so all variables exist before any widget runs.
init_state()


# ══════════════════════════════════════════════════════════════════════════════
#  EMISSION FACTORS — LOCKED REFERENCE TABLE
#
#  These are pre-loaded EPA / DEFRA / IEA 2023 emission factors.
#  They are stored in a plain Python dictionary (not in session_state)
#  because they should NOT be editable by the user — they are scientific
#  constants sourced from official databases.
#
#  Naming convention:
#    {fuel}_{gas}  → e.g. natgas_co2 = natural gas CO₂ factor
#  Units:
#    _co2  → kgCO₂ per unit of fuel
#    _ch4  → gCH₄  per unit of fuel  (converted to tCO₂e using GWP)
#    _n2o  → gN₂O  per unit of fuel  (converted to tCO₂e using GWP)
#    _gwp  → kgCO₂e per kg (already a GWP-weighted factor for refrigerants)
# ══════════════════════════════════════════════════════════════════════════════
EF = {
    # Stationary combustion — EPA CCCL 2023
    "natgas_co2":  53.06,  # kgCO₂  per MMBtu of natural gas
    "natgas_ch4":  1.0,    # gCH₄   per MMBtu of natural gas
    "natgas_n2o":  0.1,    # gN₂O   per MMBtu of natural gas

    "diesel_co2":  2.663,  # kgCO₂  per litre of diesel / fuel oil
    "diesel_ch4":  0.139,  # gCH₄   per litre of diesel
    "diesel_n2o":  0.014,  # gN₂O   per litre of diesel

    "lpg_co2":     1.555,  # kgCO₂  per litre of LPG
    "lpg_ch4":     0.0617, # gCH₄   per litre of LPG
    "lpg_n2o":     0.0062, # gN₂O   per litre of LPG

    "coal_co2":    2325.0, # kgCO₂  per short ton of bituminous coal
    "coal_ch4":    274.0,  # gCH₄   per short ton of coal
    "coal_n2o":    40.0,   # gN₂O   per short ton of coal

    # Mobile combustion — DEFRA 2023 / EPA 2023
    "gasoline_co2":  2.312, # kgCO₂  per litre of gasoline
    "gasoline_ch4":  0.339, # gCH₄   per litre of gasoline
    "gasoline_n2o":  0.033, # gN₂O   per litre of gasoline

    "jet_co2":       2.553, # kgCO₂  per litre of aviation fuel (Jet-A)

    # Fugitive — GWP-based (already kgCO₂e per kg leaked)
    "hfc410a_gwp":   2088,  # HFC-410A GWP (IPCC AR5; AR6 not yet finalised)

    # Scope 3 factors
    "s3_cat3_td_loss": 0.05,  # 5% transmission & distribution loss factor
    "s3_cat6_air":     0.255, # kgCO₂e per km — economy air travel (DEFRA 2023)
    "s3_cat6_rail":    0.041, # kgCO₂e per km — average rail travel
    "s3_cat7_car":     0.170, # kgCO₂e per km — average car commute

    # Scope 2 steam / heat
    "s2_steam_ef":     66.4,  # kgCO₂e per GJ of purchased steam or heat
}


# ── COUNTRY GRID EMISSION FACTORS ─────────────────────────────────────────────
# Suggested location-based grid EFs by country (kgCO₂e/MWh), IEA 2023.
# When the user picks their country in Assumptions, the Scope 2 page
# pre-fills the grid EF input with the matching value as a starting suggestion.
COUNTRY_EF = {
    "United States":  386,
    "United Kingdom": 207,
    "Germany":        380,
    "France":          85,  # high nuclear share → very low grid EF
    "China":          581,
    "India":          713,
    "Japan":          471,
    "Australia":      680,
    "Brazil":         119,  # high hydro share → low grid EF
    "Canada":         130,
    "Other":          386,  # defaults to US average if unknown
}


# ══════════════════════════════════════════════════════════════════════════════
#  CALCULATION FUNCTIONS
#
#  Each scope has its own pure calculation function.
#  They read from st.session_state (user inputs) and EF (locked factors),
#  then return the computed results as a dictionary and/or total.
#
#  Keeping calculations separate from UI code makes them easier to test,
#  debug, and modify without touching the display logic.
# ══════════════════════════════════════════════════════════════════════════════

def calc_scope1():
    """
    Calculate Scope 1 — Direct GHG Emissions.

    Returns:
        r     (dict)  — per-fuel breakdown as tuples (CO₂ tCO₂e, CH₄ tCO₂e, N₂O tCO₂e)
        total (float) — grand total of all Scope 1 sources in tCO₂e
    """
    s = st.session_state  # local alias so we can write s["key"] instead of st.session_state["key"]

    # Pull the user-editable GWP values from session state.
    # These are used inside the inner tco2e() function below.
    gwp_ch4 = s["gwp_ch4_fossil"]  # e.g. 29.8 for IPCC AR6
    gwp_n2o = s["gwp_n2o"]         # e.g. 273.0 for IPCC AR6

    # ── Inner helper function: convert activity data to tCO₂e ─────────────
    # This function takes the three gas factors for one fuel type and the
    # quantity consumed, then returns a 3-tuple: (CO₂ tonnes, CH₄ CO₂e, N₂O CO₂e)
    #
    # Formula:
    #   CO₂ tonnes    = co2_factor (kgCO₂/unit) × quantity / 1000   → converts kg to tonnes
    #   CH₄ CO₂e     = ch4_factor (gCH₄/unit) × quantity / 1,000,000 × GWP_CH4
    #                   (÷1e6 converts grams to tonnes; ×GWP gives CO₂ equivalents)
    #   N₂O CO₂e     = same structure as CH₄ but using GWP_N₂O
    def tco2e(co2, ch4, n2o, qty):
        return (
            co2 * qty / 1000,         # CO₂ in tCO₂
            ch4 * qty / 1e6 * gwp_ch4, # CH₄ in tCO₂e
            n2o * qty / 1e6 * gwp_n2o  # N₂O in tCO₂e
        )

    # r (results) is a dictionary where each key is a fuel type and
    # each value is a 3-tuple (co2, ch4, n2o) in tCO₂e.
    r = {}

    # ── Part A: Stationary Combustion ──────────────────────────────────────
    # Each line calls tco2e() with the relevant emission factors from EF
    # and the user-entered quantity from session_state.
    r["natgas"]       = tco2e(EF["natgas_co2"],   EF["natgas_ch4"],   EF["natgas_n2o"],   s["s1_natgas_mmbtu"])
    r["diesel"]       = tco2e(EF["diesel_co2"],   EF["diesel_ch4"],   EF["diesel_n2o"],   s["s1_diesel_litres"])
    r["lpg"]          = tco2e(EF["lpg_co2"],      EF["lpg_ch4"],      EF["lpg_n2o"],      s["s1_lpg_litres"])
    r["coal"]         = tco2e(EF["coal_co2"],     EF["coal_ch4"],     EF["coal_n2o"],     s["s1_coal_shorttons"])

    # ── Part B: Mobile Combustion (fleet vehicles) ─────────────────────────
    # Gasoline and diesel fleet use the same emission factors as stationary
    # diesel/gasoline — the fuel chemistry is the same; only the source differs.
    r["gasoline"]     = tco2e(EF["gasoline_co2"], EF["gasoline_ch4"], EF["gasoline_n2o"], s["s1_gasoline_litres"])
    r["diesel_fleet"] = tco2e(EF["diesel_co2"],   EF["diesel_ch4"],   EF["diesel_n2o"],   s["s1_diesel_fleet_litres"])

    # Aviation fuel: DEFRA 2023 only publishes a CO₂ factor for Jet-A.
    # CH₄ and N₂O are negligible/not reported, so we use 0 for them.
    r["jet"]          = (EF["jet_co2"] * s["s1_jet_litres"] / 1000, 0, 0)

    # ── Part C: Fugitive Emissions (refrigerant leaks) ─────────────────────
    # For refrigerants, the factor is already a GWP-weighted kgCO₂e per kg
    # (i.e., the GWP is embedded in the factor, unlike combustion gases).
    # Formula: kg leaked × GWP / 1000 = tCO₂e
    # We store the result in the same 3-tuple format (total, 0, 0) for consistency.
    fug = (
        s["s1_hfc134a_kg"] * s["gwp_hfc134a"]    / 1000  # user's chosen GWP for HFC-134a
      + s["s1_hfc410a_kg"] * EF["hfc410a_gwp"]   / 1000  # locked AR5 GWP for HFC-410A
      + s["s1_sf6_kg"]     * s["gwp_sf6"]         / 1000  # user's chosen GWP for SF₆
    )
    r["fugitive"] = (fug, 0, 0)  # all fugitive goes in the CO₂ slot; no separate CH₄/N₂O

    # ── Total: sum all gases across all fuel types ─────────────────────────
    # sum(v) sums the 3-tuple for each fuel (co2 + ch4 + n2o)
    # The outer sum() adds all fuel totals together → Scope 1 grand total
    total = sum(sum(v) for v in r.values())

    return r, total  # return both the breakdown dict and the grand total


def calc_scope2():
    """
    Calculate Scope 2 — Purchased Energy.
    GHG Protocol requires BOTH location-based and market-based reporting.

    Returns a dict with keys:
        lb       — location-based total (tCO₂e)
        mb       — market-based total (tCO₂e)  ← this is the PRIMARY figure
        net      — net MWh after subtracting RECs/PPAs (used in market-based calc)
        recs_pct — percentage of electricity covered by renewables
    """
    s = st.session_state

    # Net MWh for market-based = total electricity minus renewable certificates.
    # max(0, ...) prevents a negative result if RECs somehow exceed consumption.
    net = max(0, s["s2_elec_mwh"] - s["s2_recs_mwh"])

    # Steam/heat/cooling: same factor used for both location- and market-based.
    # The emission factor is in kgCO₂e/GJ; dividing by 1000 gives tCO₂e.
    steam = s["s2_steam_gj"] * EF["s2_steam_ef"] / 1000

    # ── Location-Based ────────────────────────────────────────────────────
    # Uses the average grid emission factor for the region (kgCO₂e/MWh).
    # Reflects physical reality of the electricity grid.
    lb = s["s2_elec_mwh"] * s["s2_grid_ef"] / 1000 + steam

    # ── Market-Based ──────────────────────────────────────────────────────
    # Uses the supplier-specific EF and applies it to the NET MWh
    # (i.e., after subtracting renewable energy certificates).
    # Reflects the company's actual energy procurement decisions.
    mb = net * s["s2_market_ef"] / 1000 + steam

    # Renewable coverage as a percentage of total electricity consumed.
    # max(1, ...) in the denominator prevents division-by-zero if elec = 0.
    recs_pct = s["s2_recs_mwh"] / max(1, s["s2_elec_mwh"]) * 100

    return {"lb": lb, "mb": mb, "net": net, "recs_pct": recs_pct}


def calc_scope3():
    """
    Calculate Scope 3 — Value Chain Emissions.
    We cover 5 of the 15 GHG Protocol categories.

    Returns a dict with individual category results and a 'total'.
    """
    s = st.session_state

    # Cat 1 — Purchased Goods & Services (spend-based method)
    # Formula: spend ($000s) × EEIO factor (kgCO₂e per $) / 1000 → tCO₂e
    # The spend is in thousands of dollars; factor is per single dollar.
    # So spend_thousands × 1000 × factor_per_dollar / 1000 simplifies to spend × factor.
    cat1 = s["s3_cat1_spend"] * s["s3_cat1_ef"] / 1000

    # Cat 3 — Fuel & Energy-Related (upstream T&D losses)
    # Represents emissions from transmitting electricity to the company's buildings.
    # Formula: MWh × grid EF × 5% loss rate / 1000 → tCO₂e
    cat3 = s["s3_cat3_elec_mwh"] * s["s2_grid_ef"] * EF["s3_cat3_td_loss"] / 1000

    # Cat 6 — Business Travel
    # Air + rail travel. Both are in km; factors are kgCO₂e per km.
    cat6 = (
        s["s3_cat6_air_km"]  * EF["s3_cat6_air"] +
        s["s3_cat6_rail_km"] * EF["s3_cat6_rail"]
    ) / 1000  # ÷1000 converts kgCO₂e to tCO₂e

    # Cat 7 — Employee Commuting
    # Average commute distance per employee × number of employees × car EF.
    # This assumes all commuting is by car (conservative estimate).
    cat7 = s["s3_cat7_km_per_emp"] * s["employees"] * EF["s3_cat7_car"] / 1000

    # Cat 11 — Use of Sold Products
    # Total product-years in use × lifecycle emission factor per unit.
    # The EF here must be provided by the user from product lifecycle analysis.
    cat11 = s["s3_cat11_units"] * s["s3_cat11_ef"] / 1000

    # Return all category values plus the grand total
    return {
        "cat1": cat1, "cat3": cat3, "cat6": cat6,
        "cat7": cat7, "cat11": cat11,
        "total": cat1 + cat3 + cat6 + cat7 + cat11
    }


# ── FORMATTING HELPER FUNCTIONS ───────────────────────────────────────────────
# These tiny functions format numbers consistently throughout the app
# so we don't repeat the same f-string format code in 50 different places.

def fi(v):
    """Format as integer with thousands comma. e.g. 15542600 → '15,542,600'"""
    return f"{v:,.0f}"

def ff(v):
    """Format with 1 decimal place and thousands comma. e.g. 1234.5 → '1,234.5'"""
    return f"{v:,.1f}"

def dpct(c, p):
    """
    Calculate year-on-year percentage change.
    Returns None (not a number) if prior year is zero, to avoid division by zero.
    c = current value, p = prior year value
    """
    return (c - p) / p * 100 if p > 0 else None


# ── UI HELPER: KPI CARD ───────────────────────────────────────────────────────
def kpi_html(label, value, unit, delta=None):
    """
    Build the HTML string for a McKinsey-style KPI card.

    Args:
        label  — small uppercase text above the number (e.g. "SCOPE 1 — DIRECT")
        value  — the formatted number string (e.g. "143,510")
        unit   — small text below the number (e.g. "tCO₂e")
        delta  — optional float: % change vs prior year. None = don't show delta.

    The delta arrow logic:
        - Negative delta (emissions went DOWN) → teal ↓ (good news)
        - Positive delta (emissions went UP)   → amber ↑ (bad news)
    """
    d = ""  # default: no delta HTML — only added if delta is provided

    if delta is not None:
        # Choose the CSS class and arrow direction based on sign
        cls = "delta-neg" if delta < 0 else "delta-pos"
        arr = "↓" if delta < 0 else "↑"
        # Build the delta div. abs(delta) shows the magnitude without a sign
        # because the arrow already communicates direction.
        d = f'<div class="mck-kpi-delta {cls}">{arr} {abs(delta):.1f}% vs prior year</div>'

    # Return the full KPI card HTML block.
    # CSS classes are defined in the <style> block at the top of the file.
    return f"""<div class="mck-kpi">
      <div class="mck-kpi-label">{label}</div>
      <div class="mck-kpi-value">{value}</div>
      <div class="mck-kpi-unit">{unit}</div>{d}
    </div>"""


# ── UI HELPER: PAGE HEADER ────────────────────────────────────────────────────
def section_head(eyebrow, title, caption=""):
    """
    Render the standard McKinsey page header: eyebrow → h1 title → caption.

    Args:
        eyebrow — small all-caps label above the title (e.g. "STEP 01")
        title   — main page title rendered as h1
        caption — optional grey subtitle below the title
    """
    # The eyebrow is a <span> with the CSS class "eyebrow" defined above.
    # unsafe_allow_html=True lets us inject the raw HTML tag.
    st.markdown(f'<span class="eyebrow">{eyebrow}</span>', unsafe_allow_html=True)

    # The "# " prefix makes Streamlit render this as an h1 heading,
    # which our CSS styles with Playfair Display and the navy underline rule.
    st.markdown(f"# {title}")

    # Caption is optional — only render it if a non-empty string was provided
    if caption:
        st.markdown(f'<p class="page-caption">{caption}</p>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR
#
#  The "with st.sidebar:" block places everything inside it into the sidebar
#  panel on the left side of the screen.
#  The sidebar contains:
#    1. App title / branding
#    2. Navigation (st.radio acts as a page router)
#    3. Live totals — recalculated on every rerun so they always reflect
#       the latest user input
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:

    # ── Branding block ────────────────────────────────────────────────────
    # Raw HTML for the app title and subtitle.
    # We use inline styles here because these are very specific one-off
    # styles that don't need their own CSS class.
    st.markdown("""
    <div style="padding:24px 0 8px 0;">
      <div style="font-family:'Playfair Display',serif;font-size:1.1rem;
                  font-weight:700;color:#FFFFFF;letter-spacing:-0.01em;">GHG Inventory</div>
      <div style="font-size:11px;color:#7DA3CC;text-transform:uppercase;
                  letter-spacing:0.12em;margin-top:2px;">GHG Protocol Framework</div>
    </div>""", unsafe_allow_html=True)

    st.divider()  # thin horizontal line

    # ── Navigation radio buttons ──────────────────────────────────────────
    # st.radio() renders a list of options as clickable radio buttons.
    # The selected option is stored in the variable "page".
    # label_visibility="collapsed" hides the "Nav" label (it's just for
    # accessibility / screen readers — we don't want it visible).
    # Streamlit reruns the script when the user clicks a different page,
    # so "page" will hold the new selection on the next run.
    page = st.radio("Nav", [
        "Assumptions",
        "Scope 1 — Direct",
        "Scope 2 — Purchased Energy",
        "Scope 3 — Value Chain",
        "Dashboard",
        "Export Report",
    ], label_visibility="collapsed")

    st.divider()

    # ── Live Totals ───────────────────────────────────────────────────────
    # We call the calculation functions here (inside the sidebar block)
    # so the totals always reflect the current session_state values.
    # The underscore _ is a Python convention for "I don't need this return value"
    # — calc_scope1() returns (breakdown_dict, total); we only need the total here.
    _, s1t = calc_scope1()   # s1t = Scope 1 total in tCO₂e
    s2     = calc_scope2()   # s2 is a dict with keys "lb", "mb", "net", "recs_pct"
    s3     = calc_scope3()   # s3 is a dict with category values and "total"

    # Grand total = Scope 1 + Scope 2 market-based + Scope 3
    # (GHG Protocol primary method uses market-based for Scope 2)
    grand = s1t + s2["mb"] + s3["total"]

    # Small uppercase label above the metrics
    st.markdown("<div style='font-size:10px;text-transform:uppercase;letter-spacing:0.12em;"
                "color:#7DA3CC;margin-bottom:8px;'>Live Totals</div>", unsafe_allow_html=True)

    # st.metric() renders a labelled number. In the sidebar these are styled
    # via CSS (see the stMetricValue / stMetricLabel rules near the top).
    # fi() formats numbers with comma separators and no decimals.
    st.metric("Scope 1",    f"{fi(s1t)} tCO₂e")
    st.metric("Scope 2 (MB)", f"{fi(s2['mb'])} tCO₂e")
    st.metric("Scope 3",    f"{fi(s3['total'])} tCO₂e")
    st.metric("Total",      f"{fi(grand)} tCO₂e")

    # Company name and year displayed at the very bottom of the sidebar.
    # A semi-transparent white top-border separates it from the metrics.
    st.markdown(
        f"""<div style="margin-top:20px;padding-top:16px;
            border-top:1px solid rgba(255,255,255,0.12);
            font-size:11px;color:#7DA3CC;">
            {st.session_state['company_name']}<br>FY {st.session_state['reporting_year']}
        </div>""",
        unsafe_allow_html=True,
    )


# ══════════════════════════════════════════════════════════════════════════════
#  PAGE ROUTING
#
#  Streamlit has no built-in multi-page routing (beyond its pages/ folder).
#  We simulate it with a chain of if / elif statements that check the value
#  of "page" (set by the sidebar radio buttons above).
#  Only the matching block renders — everything else is skipped.
# ══════════════════════════════════════════════════════════════════════════════


# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 1: ASSUMPTIONS
# ══════════════════════════════════════════════════════════════════════════════
if page == "Assumptions":

    # Render the McKinsey-style page header (eyebrow + h1 + caption)
    section_head(
        "Step 01",
        "Central Assumptions",
        "All calculation sheets reference these inputs. Complete this section before proceeding.",
    )

    # ── Two-column layout: Company Profile | Financial Metrics ────────────
    # st.columns(2, gap="large") creates two equal-width columns with a
    # generous gap between them. "with col1:" places content in the left column.
    col1, col2 = st.columns(2, gap="large")

    with col1:
        st.markdown("## Company Profile")

        # st.text_input() renders a single-line text field.
        # The second argument is the default value shown in the field.
        # We immediately write the result back to session_state so it persists.
        st.session_state["company_name"] = st.text_input(
            "Legal Entity / Company Name",
            st.session_state["company_name"],
        )

        # st.selectbox() renders a dropdown. The list argument provides the options.
        # The returned value (the selected string) is stored in session_state.
        st.session_state["industry"] = st.selectbox("Industry Sector", [
            "Technology", "Manufacturing", "Retail & Consumer", "Financial Services",
            "Healthcare & Pharma", "Energy & Utilities", "Real Estate",
            "Transportation & Logistics", "Food & Beverage", "Other",
        ])

        # list(COUNTRY_EF.keys()) extracts the country names from our EF dict
        # so the dropdown shows exactly the countries we have EF data for.
        st.session_state["country"] = st.selectbox(
            "Primary Country of Operations",
            list(COUNTRY_EF.keys()),
        )

        # st.number_input() renders a numeric input with up/down arrows.
        # min_value and max_value clamp the acceptable range.
        # The third positional arg is the default value.
        st.session_state["reporting_year"] = st.number_input(
            "Reporting Year", 2000, 2030, st.session_state["reporting_year"]
        )

    with col2:
        st.markdown("## Financial Metrics")
        st.caption("Used to calculate carbon intensity ratios.")

        # step=100.0 means clicking the ▲▼ arrows changes the value by 100 at a time.
        # The help= argument shows a tooltip when the user hovers over the (?) icon.
        st.session_state["revenue_musd"] = st.number_input(
            "Annual Revenue (USD Millions)",
            0.0,
            value=st.session_state["revenue_musd"],
            step=100.0,
        )

        st.session_state["employees"] = st.number_input(
            "Full-Time Equivalent Employees (FTE)",
            0,
            value=st.session_state["employees"],
            step=100,
        )

        st.markdown("## Prior Year Actuals")
        st.caption("Used for year-on-year comparison.")

        # Three equal columns inside the right main column — one per scope
        pa, pb, pc = st.columns(3)
        with pa:
            st.session_state["prior_s1"] = st.number_input(
                "Scope 1", 0.0, value=st.session_state["prior_s1"]
            )
        with pb:
            st.session_state["prior_s2mb"] = st.number_input(
                "Scope 2 MB", 0.0, value=st.session_state["prior_s2mb"]
            )
        with pc:
            st.session_state["prior_s3"] = st.number_input(
                "Scope 3", 0.0, value=st.session_state["prior_s3"]
            )

    st.divider()  # visual separator between sections

    # ── GWP Factors ───────────────────────────────────────────────────────
    st.markdown("## GWP Factors — IPCC AR6 (2021)")
    st.caption("Adjust only if your regulatory regime requires AR5 or AR4.")

    # Four equal columns, one per GWP factor
    gc1, gc2, gc3, gc4 = st.columns(4)

    with gc1:
        st.session_state["gwp_ch4_fossil"] = st.number_input(
            "CH₄ fossil GWP", value=st.session_state["gwp_ch4_fossil"]
        )
        # ef-chip is a styled badge showing the AR5 vs AR6 reference values
        # so the user knows what a "normal" value looks like before changing it.
        st.markdown('<span class="ef-chip">AR5: 25 | AR6: 29.8</span>', unsafe_allow_html=True)

    with gc2:
        st.session_state["gwp_n2o"] = st.number_input(
            "N₂O GWP", value=st.session_state["gwp_n2o"]
        )
        st.markdown('<span class="ef-chip">AR5: 298 | AR6: 273</span>', unsafe_allow_html=True)

    with gc3:
        st.session_state["gwp_hfc134a"] = st.number_input(
            "HFC-134a GWP", value=st.session_state["gwp_hfc134a"]
        )
        st.markdown('<span class="ef-chip">AR5: 1430 | AR6: 1526</span>', unsafe_allow_html=True)

    with gc4:
        st.session_state["gwp_sf6"] = st.number_input(
            "SF₆ GWP", value=st.session_state["gwp_sf6"]
        )
        st.markdown('<span class="ef-chip">AR5: 22800 | AR6: 25200</span>', unsafe_allow_html=True)

    st.divider()

    # ── Reduction Target ──────────────────────────────────────────────────
    st.markdown("## Reduction Target")
    ta, tb, tc = st.columns(3)

    with ta:
        st.session_state["target_year"] = st.number_input(
            "Target Year", 2025, 2060, st.session_state["target_year"]
        )
    with tb:
        # st.slider() renders a draggable slider. Arguments: label, min, max, default.
        st.session_state["target_reduction_pct"] = st.slider(
            "Reduction vs Baseline (%)", 0, 100, int(st.session_state["target_reduction_pct"])
        )
    with tc:
        st.session_state["target_baseline"] = st.number_input(
            "Baseline Emissions (tCO₂e)", 0.0, value=st.session_state["target_baseline"]
        )

    # Only show the calculated trajectory if the user has entered a baseline.
    # An insight box appears automatically once baseline > 0.
    if st.session_state["target_baseline"] > 0:
        # Calculate the absolute target value
        target_val = st.session_state["target_baseline"] * (
            1 - st.session_state["target_reduction_pct"] / 100
        )
        # Years remaining until the target year (minimum 1 to avoid division by zero)
        yrs = max(1, st.session_state["target_year"] - st.session_state["reporting_year"])
        # Linear annual reduction required to reach the target
        annual = (st.session_state["target_baseline"] - target_val) / yrs

        # Render the insight box with the computed values embedded in the text
        st.markdown(
            f"""<div class="insight-box">
                <strong>Target trajectory:</strong>
                Reduce to <strong>{fi(target_val)} tCO₂e</strong>
                by {st.session_state['target_year']}, requiring
                <strong>{fi(annual)} tCO₂e/year</strong>
                reduction over {yrs} years.
            </div>""",
            unsafe_allow_html=True,
        )

    # Industry benchmark — optional, used for peer comparison on the Dashboard
    st.session_state["benchmark_revenue_intensity"] = st.number_input(
        "Industry Benchmark — Revenue Intensity (tCO₂e / $M revenue)",
        0.0,
        value=st.session_state["benchmark_revenue_intensity"],
    )


# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 2: SCOPE 1 — DIRECT EMISSIONS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Scope 1 — Direct":

    # Run the Scope 1 calculation to get the current totals
    results, s1t = calc_scope1()

    # Sub-totals for each part — used in the KPI cards at the top
    # sum(results[k]) sums the (co2, ch4, n2o) 3-tuple into a single number
    stationary = sum(sum(results[k]) for k in ["natgas", "diesel", "lpg", "coal"])
    mobile     = sum(sum(results[k]) for k in ["gasoline", "diesel_fleet", "jet"])
    fugitive   = sum(results["fugitive"])

    section_head(
        "Step 02",
        "Scope 1 — Direct GHG Emissions",
        "Sources directly owned or controlled: stationary combustion, mobile combustion, and fugitive releases.",
    )

    # ── Four KPI cards across the top ─────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(kpi_html("Total Scope 1",        fi(s1t),       "tCO₂e"), unsafe_allow_html=True)
    with c2: st.markdown(kpi_html("Stationary Combustion",fi(stationary),"tCO₂e"), unsafe_allow_html=True)
    with c3: st.markdown(kpi_html("Mobile Combustion",    fi(mobile),    "tCO₂e"), unsafe_allow_html=True)
    with c4: st.markdown(kpi_html("Fugitive Emissions",   fi(fugitive),  "tCO₂e"), unsafe_allow_html=True)

    st.divider()

    # ── Part A: Stationary Combustion ─────────────────────────────────────
    st.markdown("## Part A — Stationary Combustion")
    st.caption("Boilers, furnaces, generators, turbines. Source: EPA CCCL 2023.")

    # We loop over a list of tuples. Each tuple describes one fuel type:
    #   (display name, session_state key, unit string, CO₂ EF value, results dict key)
    for label, key, unit, ef_val, rkey in [
        ("Natural Gas",             "s1_natgas_mmbtu",   "MMBtu",      EF["natgas_co2"],    "natgas"),
        ("Diesel / Fuel Oil No.2",  "s1_diesel_litres",  "Litres",     EF["diesel_co2"],    "diesel"),
        ("Liquefied Petroleum Gas", "s1_lpg_litres",     "Litres",     EF["lpg_co2"],       "lpg"),
        ("Coal (Bituminous)",       "s1_coal_shorttons", "Short Tons", EF["coal_co2"]/1000, "coal"),
        # Note: coal_co2 is per short ton, not per litre — we show it per unit
        # in the chip label. The /1000 converts to a "per short ton" display value.
    ]:
        # Calculate total tCO₂e for this fuel (sum of CO₂ + CH₄ + N₂O)
        row_total = sum(results[rkey])

        # st.expander() creates a collapsible section.
        # The header shows the fuel name and its current total — useful at a glance.
        with st.expander(f"{label}   —   {ff(row_total)} tCO₂e"):

            # Three sub-columns: input | EF badge | result metric
            ea, eb, ec_ = st.columns([3, 2, 2])

            with ea:
                # The key= argument gives this widget a unique ID in Streamlit's
                # widget tree. Without it, identical widget types on the same page
                # would conflict with each other.
                st.session_state[key] = st.number_input(
                    f"Activity Quantity ({unit})", 0.0,
                    value=st.session_state[key],
                    key=f"si_{key}",  # "si_" prefix = "scope 1 input"
                    step=1.0 if "ton" in unit.lower() else 100.0,
                )

            with eb:
                # Show the CO₂ emission factor so users can verify it
                st.markdown(
                    f"<br><span class='ef-chip'>CO₂ EF: {ef_val:.4f} kgCO₂/{unit}</span>",
                    unsafe_allow_html=True,
                )

            with ec_:
                # st.metric() displays a labelled value. Here we show the result
                # for just this one fuel type so users can see the impact.
                st.metric("Result", f"{ff(row_total)} tCO₂e")

    st.divider()

    # ── Part B: Mobile Combustion ─────────────────────────────────────────
    st.markdown("## Part B — Mobile Combustion")
    st.caption("Company-owned or leased vehicles and aircraft. Source: DEFRA 2023 / EPA 2023.")

    mb1, mb2, mb3 = st.columns(3)
    with mb1:
        st.session_state["s1_gasoline_litres"] = st.number_input(
            "Gasoline / Petrol (Litres)", 0.0,
            value=st.session_state["s1_gasoline_litres"], step=100.0
        )
        st.markdown(f"<span class='ef-chip'>{EF['gasoline_co2']} kgCO₂/L</span>", unsafe_allow_html=True)

    with mb2:
        st.session_state["s1_diesel_fleet_litres"] = st.number_input(
            "Diesel Fleet (Litres)", 0.0,
            value=st.session_state["s1_diesel_fleet_litres"], step=100.0
        )
        st.markdown(f"<span class='ef-chip'>{EF['diesel_co2']} kgCO₂/L</span>", unsafe_allow_html=True)

    with mb3:
        st.session_state["s1_jet_litres"] = st.number_input(
            "Aviation Fuel / Jet-A (Litres)", 0.0,
            value=st.session_state["s1_jet_litres"], step=100.0
        )
        st.markdown(f"<span class='ef-chip'>{EF['jet_co2']} kgCO₂/L</span>", unsafe_allow_html=True)

    st.divider()

    # ── Part C: Fugitive Emissions ────────────────────────────────────────
    st.markdown("## Part C — Fugitive Emissions")
    st.caption("Refrigerant leaks and electrical equipment. GWP factors from IPCC AR6.")

    fc1, fc2, fc3 = st.columns(3)
    with fc1:
        st.session_state["s1_hfc134a_kg"] = st.number_input(
            "HFC-134a leaked (kg)", 0.0, value=st.session_state["s1_hfc134a_kg"]
        )
        # The GWP shown in the chip comes from session_state (user-editable on Assumptions page)
        st.markdown(f"<span class='ef-chip'>GWP: {st.session_state['gwp_hfc134a']}</span>", unsafe_allow_html=True)

    with fc2:
        st.session_state["s1_hfc410a_kg"] = st.number_input(
            "HFC-410A leaked (kg)", 0.0, value=st.session_state["s1_hfc410a_kg"]
        )
        # HFC-410A uses the locked EF dict GWP (AR5 value), not the editable one
        st.markdown(f"<span class='ef-chip'>GWP: {EF['hfc410a_gwp']}</span>", unsafe_allow_html=True)

    with fc3:
        st.session_state["s1_sf6_kg"] = st.number_input(
            "SF₆ leaked (kg)", 0.0, value=st.session_state["s1_sf6_kg"]
        )
        st.markdown(f"<span class='ef-chip'>GWP: {st.session_state['gwp_sf6']}</span>", unsafe_allow_html=True)

    st.divider()

    # ── Summary Table ─────────────────────────────────────────────────────
    st.markdown("## Scope 1 Summary")

    # We build the HTML table as a string rather than using st.dataframe()
    # because st.dataframe() can't be styled with McKinsey's exact table rules.
    # The mck-table CSS class (defined above) handles all the visual formatting.
    rows = [
        ("Natural Gas",       "Stationary", ff(sum(results["natgas"]))),
        ("Diesel / Fuel Oil", "Stationary", ff(sum(results["diesel"]))),
        ("LPG",               "Stationary", ff(sum(results["lpg"]))),
        ("Coal",              "Stationary", ff(sum(results["coal"]))),
        ("Gasoline Fleet",    "Mobile",     ff(sum(results["gasoline"]))),
        ("Diesel Fleet",      "Mobile",     ff(sum(results["diesel_fleet"]))),
        ("Aviation",          "Mobile",     ff(sum(results["jet"]))),
        ("Fugitive",          "Fugitive",   ff(sum(results["fugitive"]))),
    ]

    # Build the table HTML string by concatenating each row
    tbl = ('<table class="mck-table">'
           '<thead><tr><th>Source</th><th>Category</th><th class="num">tCO₂e</th></tr></thead>'
           '<tbody>')
    for name, cat, val in rows:
        tbl += f"<tr><td>{name}</td><td>{cat}</td><td class='num'>{val}</td></tr>"
    # Total row uses the "total-row" CSS class for the grey background + bold text
    tbl += (f'<tr class="total-row">'
            f'<td><strong>Total Scope 1</strong></td><td></td>'
            f'<td class="num"><strong>{ff(s1t)}</strong></td>'
            f'</tr></tbody></table>')

    st.markdown(tbl, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 3: SCOPE 2 — PURCHASED ENERGY
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Scope 2 — Purchased Energy":

    s2 = calc_scope2()  # get all Scope 2 figures from the calculation function

    section_head(
        "Step 03",
        "Scope 2 — Purchased Energy",
        "GHG Protocol requires dual reporting: Market-Based (primary) and Location-Based (supplemental).",
    )

    # ── Three KPI cards ───────────────────────────────────────────────────
    c1, c2, c3 = st.columns(3)
    with c1: st.markdown(kpi_html("Market-Based (Primary)",  fi(s2["mb"]),       "tCO₂e"),              unsafe_allow_html=True)
    with c2: st.markdown(kpi_html("Location-Based (Suppl.)", fi(s2["lb"]),       "tCO₂e"),              unsafe_allow_html=True)
    with c3: st.markdown(kpi_html("Renewable Coverage",      f"{s2['recs_pct']:.1f}%", "of total electricity"), unsafe_allow_html=True)

    st.divider()
    st.markdown("## Part A — Purchased Electricity")

    ea, eb, ec = st.columns(3)

    with ea:
        st.session_state["s2_elec_mwh"] = st.number_input(
            "Total Electricity Consumed (MWh)", 0.0,
            value=st.session_state["s2_elec_mwh"], step=100.0
        )

    with eb:
        # Look up the suggested grid EF for the user's selected country.
        # .get() returns the default value (386) if the country key isn't found.
        suggested = COUNTRY_EF.get(st.session_state["country"], 386)
        st.session_state["s2_grid_ef"] = st.number_input(
            "Location-Based Grid EF (kgCO₂e/MWh)", 0.0,
            value=float(suggested),  # float() ensures the type matches what number_input expects
            step=1.0
        )
        st.markdown(
            f"<span class='ef-chip'>Suggested for {st.session_state['country']}: {suggested}</span>",
            unsafe_allow_html=True,
        )

    with ec:
        st.session_state["s2_market_ef"] = st.number_input(
            "Market-Based Supplier EF (kgCO₂e/MWh)", 0.0,
            value=st.session_state["s2_market_ef"], step=1.0
        )

    # Slider for RECs/PPAs — capped at total electricity consumed.
    # min() prevents the slider default from exceeding the maximum,
    # which would cause a Streamlit error.
    st.session_state["s2_recs_mwh"] = st.slider(
        "RECs / PPAs Covered (MWh)",
        0.0,
        max(st.session_state["s2_elec_mwh"], 1.0),  # max() prevents max < min when elec = 0
        min(st.session_state["s2_recs_mwh"], max(st.session_state["s2_elec_mwh"], 0.0)),
        step=100.0,
    )

    # Insight box explaining the net MWh calculation result
    st.markdown(
        f"""<div class="insight-box">
            <strong>Net metered electricity:</strong>
            {fi(s2['net'])} MWh after {fi(st.session_state['s2_recs_mwh'])} MWh RECs/PPAs.
            Market-based emissions = net MWh × {st.session_state['s2_market_ef']} kgCO₂e/MWh.
        </div>""",
        unsafe_allow_html=True,
    )

    st.divider()
    st.markdown("## Part B — Purchased Steam, Heat & Cooling")

    st.session_state["s2_steam_gj"] = st.number_input(
        "Purchased Steam / Heat / Cooling (GJ)", 0.0,
        value=st.session_state["s2_steam_gj"]
    )
    st.markdown(f"<span class='ef-chip'>EF: {EF['s2_steam_ef']} kgCO₂e/GJ</span>", unsafe_allow_html=True)

    st.divider()

    # Summary table comparing the two methods
    tbl = ('<table class="mck-table"><thead><tr>'
           '<th>Method</th><th>Basis</th><th class="num">tCO₂e</th>'
           '</tr></thead><tbody>')
    tbl += f"<tr><td>Location-Based</td><td>Grid average × total MWh</td><td class='num'>{ff(s2['lb'])}</td></tr>"
    tbl += f"<tr><td>Market-Based</td><td>Supplier EF × net MWh after RECs</td><td class='num'>{ff(s2['mb'])}</td></tr>"
    tbl += (f'<tr class="total-row"><td><strong>Primary (Market-Based)</strong></td>'
            f'<td>Used in total inventory</td>'
            f'<td class="num"><strong>{ff(s2["mb"])}</strong></td>'
            f'</tr></tbody></table>')
    st.markdown(tbl, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 4: SCOPE 3 — VALUE CHAIN
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Scope 3 — Value Chain":

    s3 = calc_scope3()

    section_head(
        "Step 04",
        "Scope 3 — Value Chain Emissions",
        "Indirect emissions from upstream and downstream activities. Typically 70–90% of total inventory.",
    )

    # Single KPI card showing the Scope 3 total
    c1, _ = st.columns([1, 3])  # narrow left column for the KPI; wide right column empty
    with c1:
        st.markdown(kpi_html("Scope 3 Total", fi(s3["total"]), "tCO₂e"), unsafe_allow_html=True)

    st.divider()

    # ── Category expanders ────────────────────────────────────────────────
    # configs is a list of tuples, one per Scope 3 category we've implemented.
    # Each tuple contains everything needed to render the input UI:
    #   (category number, name, input key 1, input key 2 (or None),
    #    label 1, label 2, methodology note, results dict key, has_two_inputs)
    configs = [
        ("Cat. 1",  "Purchased Goods & Services",       "s3_cat1_spend",      "s3_cat1_ef",
         "Annual Spend ($000s USD)", "EEIO Factor (kgCO₂e per $)",
         "Spend-based | US EPA USEEIO v2.0 | Typical: 0.12–0.95",
         "cat1", True),  # True = has two input fields

        ("Cat. 3",  "Fuel & Energy-Related (T&D Losses)","s3_cat3_elec_mwh", None,
         "Electricity consumed (MWh)", None,
         f"5% T&D loss × grid EF ({st.session_state['s2_grid_ef']} kgCO₂e/MWh)",
         "cat3", False),  # False = only one input field

        ("Cat. 6",  "Business Travel",                   "s3_cat6_air_km",    None,
         "Total Air Travel (km)", None,
         f"Economy class | EF: {EF['s3_cat6_air']} kgCO₂e/km | DEFRA 2023",
         "cat6", False),

        ("Cat. 7",  "Employee Commuting",                "s3_cat7_km_per_emp",None,
         "Avg commute per employee/year (km)", None,
         f"Applied to {st.session_state['employees']:,} FTEs | EF: {EF['s3_cat7_car']} kgCO₂e/km",
         "cat7", False),

        ("Cat. 11", "Use of Sold Products",              "s3_cat11_units",    "s3_cat11_ef",
         "Units / product-years in use", "Lifecycle EF (kgCO₂e/unit)",
         "Activity-based | Provide product-specific lifecycle EF",
         "cat11", True),
    ]

    for cat_num, cat_name, k1, k2, lbl1, lbl2, note, rkey, has_two in configs:

        val = s3[rkey]  # the calculated tCO₂e for this category

        # The expander header shows both the category name and its current total
        with st.expander(f"**{cat_num} — {cat_name}**   —   {ff(val)} tCO₂e"):

            if has_two:
                # Two-input categories (e.g. Cat 1: spend AND EEIO factor)
                ca, cb, cc = st.columns([3, 2, 2])
                with ca:
                    st.session_state[k1] = st.number_input(
                        lbl1, 0.0, value=st.session_state[k1],
                        step=1000.0, key=f"s3a_{k1}"  # "s3a_" = scope 3 first input
                    )
                with cb:
                    st.session_state[k2] = st.number_input(
                        lbl2, 0.0, value=st.session_state[k2],
                        step=0.01, key=f"s3b_{k2}"    # "s3b_" = scope 3 second input
                    )
                with cc:
                    st.metric("Result", f"{ff(val)} tCO₂e")
            else:
                # Single-input categories (e.g. Cat 3: just MWh consumed)
                ca, cb = st.columns([3, 2])
                with ca:
                    st.session_state[k1] = st.number_input(
                        lbl1, 0.0, value=st.session_state[k1],
                        step=1000.0, key=f"s3a_{k1}"
                    )
                with cb:
                    st.metric("Result", f"{ff(val)} tCO₂e")

            # Small caption below each expander's inputs explaining the methodology
            st.caption(f"Methodology: {note}")

    st.divider()
    st.markdown("## Category Summary")

    # List of (results_key, display_label, method) for building the summary table
    cat_labels = [
        ("cat1",  "Cat 1 — Purchased Goods & Services", "Spend-based"),
        ("cat3",  "Cat 3 — Fuel & Energy (T&D)",        "Activity-based"),
        ("cat6",  "Cat 6 — Business Travel",             "Activity-based"),
        ("cat7",  "Cat 7 — Employee Commuting",          "Activity-based"),
        ("cat11", "Cat 11 — Use of Sold Products",       "Activity-based"),
    ]

    tbl = ('<table class="mck-table"><thead><tr>'
           '<th>Category</th><th>Method</th>'
           '<th class="num">tCO₂e</th><th class="num">% of S3</th>'
           '</tr></thead><tbody>')
    for k, label, method in cat_labels:
        # Each category's share of the Scope 3 total as a percentage
        pct = s3[k] / max(s3["total"], 1) * 100  # max() prevents division by zero
        tbl += f"<tr><td>{label}</td><td>{method}</td><td class='num'>{ff(s3[k])}</td><td class='num'>{pct:.1f}%</td></tr>"
    tbl += (f'<tr class="total-row">'
            f'<td><strong>Total Scope 3</strong></td><td></td>'
            f'<td class="num"><strong>{ff(s3["total"])}</strong></td>'
            f'<td class="num"><strong>100%</strong></td>'
            f'</tr></tbody></table>')
    st.markdown(tbl, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 5: DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Dashboard":

    # Recalculate all three scopes freshly for this page render
    _, s1t = calc_scope1()
    s2     = calc_scope2()
    s3     = calc_scope3()
    grand  = s1t + s2["mb"] + s3["total"]

    # Prior year grand total (for overall YoY delta on the Total KPI card)
    prior_grand = (st.session_state["prior_s1"]
                 + st.session_state["prior_s2mb"]
                 + st.session_state["prior_s3"])

    s = st.session_state  # local alias for brevity

    section_head(
        "Executive View",
        f"{s['company_name']} — GHG Inventory {s['reporting_year']}",
        "GHG Protocol Corporate Standard  ·  Market-Based Scope 2  ·  Figures in tCO₂e",
    )

    # ── Top KPI row: four cards, one per scope + grand total ──────────────
    c1, c2, c3, c4 = st.columns(4)

    # dpct() returns the % change between current and prior; None if prior = 0.
    # kpi_html() uses the delta to show the ↑/↓ arrow line under the number.
    with c1: st.markdown(kpi_html("Scope 1 — Direct",       fi(s1t),        "tCO₂e", dpct(s1t,        s["prior_s1"])),    unsafe_allow_html=True)
    with c2: st.markdown(kpi_html("Scope 2 — Market-Based", fi(s2["mb"]),   "tCO₂e", dpct(s2["mb"],   s["prior_s2mb"])), unsafe_allow_html=True)
    with c3: st.markdown(kpi_html("Scope 3 — Value Chain",  fi(s3["total"]),"tCO₂e", dpct(s3["total"],s["prior_s3"])),   unsafe_allow_html=True)
    with c4: st.markdown(kpi_html("Total (S1 + S2 + S3)",   fi(grand),      "tCO₂e", dpct(grand, prior_grand)),           unsafe_allow_html=True)

    st.divider()

    # ── Two-column layout: Scope Breakdown (left) | Intensity (right) ─────
    # [3, 2] means the left column is 3/5 of the width; right column is 2/5.
    col_l, col_r = st.columns([3, 2], gap="large")

    with col_l:
        st.markdown("## Scope Breakdown")

        # Each tuple: (scope label, description, current value, prior year value)
        scope_rows = [
            ("Scope 1",      "Direct — combustion, fleet, fugitive",   s1t,        s["prior_s1"]),
            ("Scope 2 (MB)", "Purchased energy — market-based",        s2["mb"],   s["prior_s2mb"]),
            ("Scope 3",      "Value chain — upstream & downstream",    s3["total"],s["prior_s3"]),
            ("Total",        "S1 + S2(MB) + S3",                       grand,      prior_grand),
        ]

        tbl = ('<table class="mck-table"><thead><tr>'
               '<th>Scope</th><th>Description</th>'
               '<th class="num">tCO₂e</th><th class="num">% Share</th><th class="num">YoY</th>'
               '</tr></thead><tbody>')

        for sc, desc, val, prior in scope_rows:
            # % share of grand total (100% for the Total row itself)
            pct = f"{val / max(grand, 1) * 100:.1f}%" if sc != "Total" else "100%"

            # Year-on-year change: formatted as +/- percentage, or "—" if no prior data
            yoy = f"{dpct(val, prior):+.1f}%" if prior > 0 else "—"

            # The Total row gets the mck-table total-row CSS class (grey bg + bold)
            cls = ' class="total-row"' if sc == "Total" else ""

            # Bold tags applied to the Total row only
            b  = "<strong>" if sc == "Total" else ""
            be = "</strong>" if sc == "Total" else ""

            tbl += (f"<tr{cls}><td>{b}{sc}{be}</td><td>{desc}</td>"
                    f"<td class='num'>{b}{fi(val)}{be}</td>"
                    f"<td class='num'>{pct}</td>"
                    f"<td class='num'>{yoy}</td></tr>")

        tbl += "</tbody></table>"
        st.markdown(tbl, unsafe_allow_html=True)

        # Small footnote showing the location-based Scope 2 figure (supplemental)
        st.markdown(
            f"<div style='font-size:11px;color:var(--mid-grey);margin-top:6px;'>"
            f"◦ Scope 2 Location-Based (supplemental): {fi(s2['lb'])} tCO₂e"
            f"</div>",
            unsafe_allow_html=True,
        )

    with col_r:
        st.markdown("## Carbon Intensity")

        rev   = s["revenue_musd"]   # revenue in $M
        emp   = s["employees"]      # headcount
        bench = s["benchmark_revenue_intensity"]  # industry benchmark (optional)

        # Build the intensity table; rows added conditionally only if data exists
        tbl2 = ('<table class="mck-table"><thead><tr>'
                '<th>Metric</th><th class="num">Value</th><th class="num">Unit</th><th class="num">YoY</th>'
                '</tr></thead><tbody>')

        if rev > 0:
            ri = grand / rev  # current revenue intensity
            # Prior year intensity — only calculable if we have prior year emissions
            pri_ri = prior_grand / rev if prior_grand > 0 else None
            tbl2 += (f"<tr><td>Revenue intensity</td>"
                     f"<td class='num'>{ri:.2f}</td><td class='num'>tCO₂e/$M</td>"
                     f"<td class='num'>{f'{dpct(ri, pri_ri):+.1f}%' if pri_ri else '—'}</td></tr>")

            if bench > 0:
                # Show how the company compares to the industry benchmark
                vs = (ri - bench) / bench * 100  # positive = worse than benchmark
                tbl2 += (f"<tr><td>vs Benchmark</td>"
                         f"<td class='num'>{vs:+.1f}%</td>"
                         f"<td class='num'>Bench: {bench:.1f}</td>"
                         f"<td class='num'>—</td></tr>")

        if emp > 0:
            ei = grand / emp  # current employee intensity
            pri_ei = prior_grand / emp if prior_grand > 0 else None
            tbl2 += (f"<tr><td>Employee intensity</td>"
                     f"<td class='num'>{ei:.2f}</td><td class='num'>tCO₂e/FTE</td>"
                     f"<td class='num'>{f'{dpct(ei, pri_ei):+.1f}%' if pri_ei else '—'}</td></tr>")

        tbl2 += "</tbody></table>"

        # Only render the table if there's at least one row of data
        if rev > 0 or emp > 0:
            st.markdown(tbl2, unsafe_allow_html=True)
        else:
            st.caption("Enter revenue and employees in Assumptions to view intensity ratios.")

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Renewable Energy progress bar ──────────────────────────────────
        st.markdown("## Renewable Energy")
        rp = s2["recs_pct"]  # percentage of electricity covered by RECs/PPAs

        # HTML progress bar — Streamlit's built-in progress bar can't be styled
        # with our McKinsey CSS, so we build our own using divs.
        # min(rp, 100) prevents the fill from exceeding 100% width.
        st.markdown(
            f"""<div style="margin:12px 0;">
              <div class="mck-progress-label">
                <span>Renewable Coverage</span><span>{rp:.1f}%</span>
              </div>
              <div class="mck-progress-track">
                <div class="mck-progress-fill" style="width:{min(rp, 100):.1f}%"></div>
              </div>
            </div>""",
            unsafe_allow_html=True,
        )

    # ── Reduction Target Progress (only shown if baseline has been set) ────
    if s["target_baseline"] > 0:
        st.divider()
        st.markdown("## Reduction Target Progress")

        # Target value in absolute tCO₂e terms
        target_val = s["target_baseline"] * (1 - s["target_reduction_pct"] / 100)

        # Progress ratio (0 to 1): how much of the required reduction has been achieved?
        # Clamped to [0, 1] so the progress bar never goes negative or over 100%.
        progress = max(0, min(1,
            (s["target_baseline"] - grand) / max(1, s["target_baseline"] - target_val)
        ))

        yrs = max(1, s["target_year"] - s["reporting_year"])  # years remaining

        # Annual reduction still needed to reach target from current level
        annual_req = max(0, (grand - target_val) / yrs)

        ta, tb, tc, td = st.columns(4)
        with ta: st.markdown(kpi_html("Baseline",              fi(s["target_baseline"]), "tCO₂e"),             unsafe_allow_html=True)
        with tb: st.markdown(kpi_html("Current",               fi(grand),                "tCO₂e"),             unsafe_allow_html=True)
        with tc: st.markdown(kpi_html("Target",                fi(target_val),           f"tCO₂e by {s['target_year']}"), unsafe_allow_html=True)
        with td: st.markdown(kpi_html("Annual Reduction Req.", fi(annual_req),           "tCO₂e/year"),        unsafe_allow_html=True)

        # Progress bar showing % of the way to the target
        st.markdown(
            f"""<div style="margin-top:16px;">
              <div class="mck-progress-label">
                <span>Progress toward {s['target_reduction_pct']}% reduction by {s['target_year']}</span>
                <span>{progress * 100:.0f}%</span>
              </div>
              <div class="mck-progress-track">
                <div class="mck-progress-fill" style="width:{min(progress * 100, 100):.1f}%"></div>
              </div>
            </div>""",
            unsafe_allow_html=True,
        )


# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 6: EXPORT REPORT
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Export Report":

    # Recalculate everything for the export page
    _, s1t = calc_scope1()
    s2     = calc_scope2()
    s3     = calc_scope3()
    grand  = s1t + s2["mb"] + s3["total"]
    s      = st.session_state

    section_head(
        "Step 05",
        "Export Report",
        "Download formatted outputs or save and restore your complete input dataset.",
    )

    # ── Build the report text ──────────────────────────────────────────────
    # This is a multi-line f-string that produces a formatted plain-text report.
    # '═'*65 creates a line of 65 '═' characters as a visual border.
    # The ">13,.1f" format spec: right-align in 13 chars, comma separator, 1 decimal.
    report_text = f"""{'═'*65}
  GHG CARBON FOOTPRINT REPORT
  {s['company_name']}  |  FY {s['reporting_year']}
  Generated: {datetime.now().strftime('%B %d, %Y %H:%M')}
  Framework: GHG Protocol Corporate Standard
{'═'*65}

SECTION 1 — SCOPE TOTALS
{'─'*65}
  Scope 1   Direct Emissions                  {s1t:>13,.1f}  tCO₂e
  Scope 2   Market-Based (Primary)            {s2['mb']:>13,.1f}  tCO₂e
  Scope 2   Location-Based (Supplemental)     {s2['lb']:>13,.1f}  tCO₂e
  Scope 3   Value Chain Emissions             {s3['total']:>13,.1f}  tCO₂e
{'─'*65}
  TOTAL     S1 + S2(MB) + S3                  {grand:>13,.1f}  tCO₂e

SECTION 2 — SCOPE 3 DETAIL
{'─'*65}
  Cat  1   Purchased Goods & Services         {s3['cat1']:>13,.1f}  tCO₂e
  Cat  3   Fuel & Energy (T&D Losses)         {s3['cat3']:>13,.1f}  tCO₂e
  Cat  6   Business Travel                    {s3['cat6']:>13,.1f}  tCO₂e
  Cat  7   Employee Commuting                 {s3['cat7']:>13,.1f}  tCO₂e
  Cat 11   Use of Sold Products               {s3['cat11']:>13,.1f}  tCO₂e

SECTION 3 — CARBON INTENSITY
{'─'*65}"""

    # Intensity lines are only added if the user has entered revenue / employees
    rev, emp = s["revenue_musd"], s["employees"]
    if rev > 0:
        report_text += f"\n  Revenue Intensity          {grand/rev:>10.2f}  tCO₂e / $M revenue"
    if emp > 0:
        report_text += f"\n  Employee Intensity         {grand/emp:>10.2f}  tCO₂e / FTE"

    # Continue building the report with assumptions and methodology sections
    report_text += f"""

SECTION 4 — ASSUMPTIONS
{'─'*65}
  Revenue                    ${rev:>12,.0f}M USD
  Employees                  {emp:>12,} FTE
  Grid EF (location)         {s['s2_grid_ef']:>12} kgCO₂e/MWh
  Market EF                  {s['s2_market_ef']:>12} kgCO₂e/MWh
  GWP CH₄ (fossil)           {s['gwp_ch4_fossil']:>12}  [IPCC AR6]
  GWP N₂O                    {s['gwp_n2o']:>12}  [IPCC AR6]
  GWP HFC-134a               {s['gwp_hfc134a']:>12}  [IPCC AR6]
  GWP SF₆                    {s['gwp_sf6']:>12}  [IPCC AR6]

SECTION 5 — REDUCTION TARGET
{'─'*65}
  Baseline                   {s['target_baseline']:>12,.0f}  tCO₂e
  Target Year                {s['target_year']:>12}
  Reduction Required         {s['target_reduction_pct']:>11.0f}%
  Target Value               {s['target_baseline']*(1-s['target_reduction_pct']/100):>12,.0f}  tCO₂e

METHODOLOGY
{'─'*65}
  Emission Factors: EPA CCCL 2023, DEFRA 2023, IEA 2023,
    US EPA USEEIO v2.0
  GWP Reference: IPCC AR6, 2021
  Scope 2 Primary Method: Market-Based
  Scope 3 Methods: Spend-based (Cat 1), Activity-based (Cat 3, 6, 7, 11)
{'═'*65}
"""

    # ── Three download buttons side by side ───────────────────────────────
    d1, d2, d3 = st.columns(3)

    with d1:
        # st.download_button() creates a button that, when clicked, downloads a file.
        # data= is the content; file_name= is the name the browser saves it as.
        # The company name has spaces replaced with underscores for a clean filename.
        st.download_button(
            "⬇  Report (.txt)",
            data=report_text,
            file_name=f"GHG_Report_{s['company_name'].replace(' ', '_')}_{s['reporting_year']}.txt",
            mime="text/plain",
            use_container_width=True,  # button stretches to fill the column width
        )

    with d2:
        # Build a CSV as a list of rows, then join with newlines.
        # Each row is a list of values joined with commas.
        csv_rows = [
            ["Metric",       "Value",               "Unit"],
            ["Company",      s["company_name"],      ""],
            ["Year",         s["reporting_year"],    ""],
            ["Scope 1",      f"{s1t:.2f}",           "tCO₂e"],
            ["Scope 2 MB",   f"{s2['mb']:.2f}",      "tCO₂e"],
            ["Scope 2 LB",   f"{s2['lb']:.2f}",      "tCO₂e"],
            ["Scope 3",      f"{s3['total']:.2f}",   "tCO₂e"],
            ["Total",        f"{grand:.2f}",          "tCO₂e"],
        ]
        if rev > 0:
            csv_rows.append(["Revenue Intensity", f"{grand/rev:.4f}", "tCO₂e/$M"])
        if emp > 0:
            csv_rows.append(["Employee Intensity", f"{grand/emp:.4f}", "tCO₂e/FTE"])

        # Convert the list of lists into a CSV string:
        # Inner join: each row's values joined by commas
        # Outer join: each row joined by a newline character
        csv_str = "\n".join(",".join(str(c) for c in row) for row in csv_rows)

        st.download_button(
            "⬇  Summary (.csv)",
            data=csv_str,
            file_name=f"GHG_Summary_{s['company_name'].replace(' ', '_')}_{s['reporting_year']}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with d3:
        # Save all session_state inputs as JSON.
        # We exclude widget-internal keys that Streamlit auto-generates
        # (they start with "si_", "s3a_", or "s3b_" — our own prefix convention).
        save_keys = [
            k for k in st.session_state
            if not k.startswith("si_")
            and not k.startswith("s3a_")
            and not k.startswith("s3b_")
        ]

        # json.dumps() converts the dict to a JSON string.
        # indent=2 formats it with 2-space indentation for readability.
        st.download_button(
            "💾  Save Inputs (.json)",
            data=json.dumps({k: st.session_state[k] for k in save_keys}, indent=2),
            file_name=f"GHG_Inputs_{s['company_name'].replace(' ', '_')}.json",
            mime="application/json",
            use_container_width=True,
        )

    st.divider()

    # Collapsible report preview — useful for checking before downloading
    with st.expander("Report Preview"):
        st.text(report_text)  # st.text() renders the string in a monospace pre-formatted block

    st.divider()

    # ── Restore saved inputs ───────────────────────────────────────────────
    st.markdown("## Restore Saved Inputs")
    st.caption("Upload a previously saved .json file to reload all inputs.")

    # st.file_uploader() renders a file picker widget.
    # type=["json"] restricts it to JSON files only.
    # label_visibility="collapsed" hides the widget label (we use st.markdown above instead).
    uploaded = st.file_uploader("Upload JSON", type="json", label_visibility="collapsed")

    if uploaded:
        # json.load() parses the uploaded file object into a Python dictionary.
        # We then write every key-value pair back into session_state.
        for k, v in json.load(uploaded).items():
            st.session_state[k] = v
        # st.success() shows a green banner confirming the operation
        st.success("Inputs restored. Navigate to any section to review.")
