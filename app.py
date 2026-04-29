import io
from typing import Optional

import pandas as pd
import streamlit as st
from pandas.errors import ParserError


APP_CSS = """
    <style>
        /* Sun Life of Canada Theme - Modern Refresh */
        :root {
            --sl-gold: #FFD100;
            --sl-gold-light: #FFE873;
            --sl-navy: #003DA5;
            --sl-navy-dark: #002A73;
            --sl-navy-soft: #E8F0FE;
            --sl-slate: #1E293B;
            --sl-gray-50: #F8FAFC;
            --sl-gray-100: #F1F5F9;
            --sl-gray-200: #E2E8F0;
            --sl-gray-600: #475569;
            --sl-white: #FFFFFF;
            --shadow-sm: 0 1px 3px rgba(0,0,0,0.05), 0 1px 2px rgba(0,0,0,0.03);
            --shadow-md: 0 4px 6px -1px rgba(0,0,0,0.05), 0 2px 4px -1px rgba(0,0,0,0.03);
            --shadow-lg: 0 10px 15px -3px rgba(0,0,0,0.05), 0 4px 6px -2px rgba(0,0,0,0.02);
            --shadow-xl: 0 20px 25px -5px rgba(0,0,0,0.05), 0 8px 10px -6px rgba(0,0,0,0.02);
        }

        /* Base styling */
        html, body, .stApp {
            font-family: 'Inter', system-ui, -apple-system, 'Segoe UI', Roboto, sans-serif;
            background: var(--sl-gray-50);
        }

        /* Main container background */
        .stApp {
            background: linear-gradient(135deg, var(--sl-gray-50) 0%, var(--sl-white) 100%);
        }

        /* Sidebar - Sun Life Navy with gradient */
        div[data-testid="stSidebar"] {
            background: linear-gradient(180deg, var(--sl-navy) 0%, var(--sl-navy-dark) 100%);
            border-right: none;
        }
        
        div[data-testid="stSidebar"] *:not(button) {
            color: var(--sl-white) !important;
        }
        
        div[data-testid="stSidebar"] .stMarkdown h1,
        div[data-testid="stSidebar"] .stMarkdown h2,
        div[data-testid="stSidebar"] .stMarkdown h3 {
            color: var(--sl-gold) !important;
        }
        
        div[data-testid="stSidebar"] .stSelectbox label,
        div[data-testid="stSidebar"] .stMultiSelect label,
        div[data-testid="stSidebar"] .stSlider label,
        div[data-testid="stSidebar"] .stNumberInput label {
            color: var(--sl-gray-200) !important;
            font-weight: 500;
        }
        
        div[data-testid="stSidebar"] .stSelectbox div[data-baseweb="select"] {
            background-color: rgba(255,255,255,0.1);
            border-color: rgba(255,255,255,0.2);
        }

        /* Hero Section - Modern Card */
        .sl-hero {
            background: var(--sl-white);
            border-radius: 24px;
            padding: 24px 28px;
            margin-bottom: 28px;
            border: 1px solid var(--sl-gray-200);
            box-shadow: var(--shadow-lg);
            position: relative;
            overflow: hidden;
            transition: all 0.2s ease;
        }
        
        .sl-hero::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 6px;
            height: 100%;
            background: linear-gradient(135deg, var(--sl-gold) 0%, var(--sl-gold-light) 100%);
        }
        
        .sl-hero-title {
            font-size: 1.75rem;
            font-weight: 700;
            background: linear-gradient(135deg, var(--sl-navy) 0%, var(--sl-navy-dark) 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            margin-bottom: 8px;
            letter-spacing: -0.3px;
        }
        
        .sl-hero-sub {
            color: var(--sl-gray-600);
            font-size: 0.95rem;
            line-height: 1.5;
        }

        /* Modern Metric Cards */
        div[data-testid="stMetric"] {
            background: var(--sl-white);
            border: 1px solid var(--sl-gray-200);
            border-radius: 20px;
            padding: 16px 20px;
            box-shadow: var(--shadow-md);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        
        div[data-testid="stMetric"]:hover {
            transform: translateY(-2px);
            box-shadow: var(--shadow-lg);
        }
        
        div[data-testid="stMetricLabel"] {
            color: var(--sl-navy);
            font-weight: 600;
            font-size: 0.8rem;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        div[data-testid="stMetricValue"] {
            color: var(--sl-slate);
            font-weight: 800;
            font-size: 2rem;
        }
        
        div[data-testid="stMetricDelta"] {
            font-size: 0.85rem;
            font-weight: 600;
        }

        /* Tabs - Sun Life Style */
        div[data-baseweb="tab-list"] {
            background: var(--sl-gray-100);
            border-radius: 60px;
            padding: 4px;
            gap: 4px;
        }
        
        button[role="tab"] {
            border-radius: 60px !important;
            padding: 8px 20px !important;
            font-weight: 600 !important;
            transition: all 0.2s ease;
        }
        
        button[role="tab"][aria-selected="true"] {
            background: var(--sl-gold) !important;
            color: var(--sl-navy-dark) !important;
            box-shadow: var(--shadow-sm);
        }
        
        button[role="tab"]:hover:not([aria-selected="true"]) {
            background: var(--sl-gray-200) !important;
        }

        /* Download Buttons */
        .stDownloadButton button {
            background: linear-gradient(135deg, var(--sl-navy) 0%, var(--sl-navy-dark) 100%);
            color: var(--sl-white);
            border: none;
            border-radius: 40px;
            padding: 10px 24px;
            font-weight: 600;
            transition: all 0.2s ease;
        }
        
        .stDownloadButton button:hover {
            transform: translateY(-1px);
            box-shadow: var(--shadow-md);
            border-color: var(--sl-gold);
        }

        /* Section Headers */
        .sl-section {
            margin: 24px 0 16px 0;
            border-left: 4px solid var(--sl-gold);
            padding: 16px 20px;
            background: var(--sl-white);
            border-radius: 16px;
            box-shadow: var(--shadow-sm);
            border: 1px solid var(--sl-gray-200);
            border-left-width: 4px;
        }
        
        .sl-section-title {
            color: var(--sl-navy);
            font-weight: 700;
            font-size: 1.2rem;
            margin-bottom: 4px;
        }
        
        .sl-section-sub {
            color: var(--sl-gray-600);
            font-size: 0.85rem;
        }

        /* Expander styling */
        .streamlit-expanderHeader {
            background: var(--sl-gray-50);
            border-radius: 12px;
            font-weight: 600;
            color: var(--sl-navy);
        }

        /* Dataframes and tables */
        .stDataFrame {
            border-radius: 16px;
            overflow: hidden;
            box-shadow: var(--shadow-sm);
        }
        
        .stDataFrame div[data-testid="stDataFrameResizable"] {
            border-radius: 16px;
            border: 1px solid var(--sl-gray-200);
        }

        /* Info/Warning/Success boxes */
        .stAlert {
            border-radius: 16px;
            border-left-width: 4px;
        }
        
        .stAlert div[data-testid="stMarkdownContainer"] {
            font-size: 0.9rem;
        }

        /* File uploader */
        div[data-testid="stFileUploader"] {
            background: var(--sl-gray-50);
            border: 2px dashed var(--sl-gray-200);
            border-radius: 20px;
            padding: 20px;
        }
        
        div[data-testid="stFileUploader"]:hover {
            border-color: var(--sl-gold);
            background: var(--sl-white);
        }

        /* Select boxes and inputs */
        .stSelectbox div[data-baseweb="select"] {
            border-radius: 12px;
            border-color: var(--sl-gray-200);
        }
        
        .stSelectbox div[data-baseweb="select"]:focus-within {
            border-color: var(--sl-gold);
            box-shadow: 0 0 0 2px rgba(255, 209, 0, 0.2);
        }
        
        .stNumberInput input {
            border-radius: 12px;
            border-color: var(--sl-gray-200);
        }
        
        .stTextInput input {
            border-radius: 12px;
            border-color: var(--sl-gray-200);
        }

        /* Chart containers */
        .stChart {
            background: var(--sl-white);
            border-radius: 16px;
            padding: 12px;
            border: 1px solid var(--sl-gray-200);
        }

        /* Caption text */
        .stCaption {
            color: var(--sl-gray-600);
            font-size: 0.8rem;
        }

        /* Code blocks */
        .stCodeBlock {
            border-radius: 12px;
        }

        /* Markdown headers */
        h1, h2, h3, h4 {
            color: var(--sl-navy);
        }
        
        h3 {
            font-size: 1.25rem;
            font-weight: 600;
        }

        /* Scrollbar styling */
        ::-webkit-scrollbar {
            width: 8px;
            height: 8px;
        }
        
        ::-webkit-scrollbar-track {
            background: var(--sl-gray-100);
            border-radius: 10px;
        }
        
        ::-webkit-scrollbar-thumb {
            background: var(--sl-gray-600);
            border-radius: 10px;
        }
        
        ::-webkit-scrollbar-thumb:hover {
            background: var(--sl-navy);
        }

        /* Button hover states */
        .stButton button {
            border-radius: 40px;
            font-weight: 600;
            transition: all 0.2s ease;
        }
        
        .stButton button:hover {
            transform: translateY(-1px);
            box-shadow: var(--shadow-md);
        }

        /* Metric container adjustments */
        div[data-testid="column"] {
            gap: 1rem;
        }
        
        .stMetric > div {
            background: transparent !important;
        }
    </style>
"""


def apply_theme() -> None:
    st.markdown(APP_CSS, unsafe_allow_html=True)


def render_hero() -> None:
    st.markdown(
        """
    <div class="sl-hero">
        <div class="sl-hero-title">✨ Manpower Validation System</div>
        <div class="sl-hero-sub">
            Upload an Excel file and analyze production by advisor classification with
            AC, NSC, and Lives across monthly and quarterly views.
        </div>
    </div>
    """,
        unsafe_allow_html=True,
    )


def render_section(title: str, subtitle: str) -> None:
    st.markdown(
        f"""
        <div class="sl-section">
            <div class="sl-section-title">{title}</div>
            <div class="sl-section-sub">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


st.set_page_config(page_title="Manpower Validation System", layout="wide", page_icon="✨")
apply_theme()
render_hero()


def normalize_number(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace(" ", "", regex=False)
        .replace({"": "0", "nan": "0", "None": "0"})
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


def build_period_columns(
    df: pd.DataFrame,
    date_col: Optional[str],
    month_col: Optional[str],
    year_col: Optional[str],
) -> pd.DataFrame:
    out = df.copy()
    parsed_date = pd.Series(pd.NaT, index=out.index, dtype="datetime64[ns]")

    if date_col:
        date_text = out[date_col].astype(str).str.strip()
        # Primary format: "April 1, 2026"
        parsed_date = pd.to_datetime(date_text, format="%B %d, %Y", errors="coerce")
        # Fallback for other valid date shapes.
        parsed_date = parsed_date.fillna(pd.to_datetime(date_text, errors="coerce"))
    elif month_col and year_col:
        month_raw = out[month_col].astype(str).str.strip()
        year_raw = out[year_col].astype(str).str.extract(r"(\d{4})")[0]
        month_num = pd.to_numeric(month_raw, errors="coerce")
        month_from_name = pd.to_datetime(month_raw, format="%B", errors="coerce").dt.month
        month_final = month_num.fillna(month_from_name)
        composed = year_raw.fillna("") + "-" + month_final.fillna(1).astype(int).astype(str) + "-01"
        parsed_date = pd.to_datetime(composed, errors="coerce")

    out["Period Month"] = parsed_date.dt.to_period("M").astype(str)
    out["Period Quarter"] = parsed_date.dt.to_period("Q").astype(str)
    out["Period Month"] = out["Period Month"].replace("NaT", "Unknown")
    out["Period Quarter"] = out["Period Quarter"].replace("NaT", "Unknown")
    return out


def guess_column(columns: list[str], preferred_keywords: list[str]) -> str:
    lowered = [(col, col.lower()) for col in columns]
    for keyword in preferred_keywords:
        for original, lowered_col in lowered:
            if keyword in lowered_col:
                return original
    return columns[0]


def series_is_mostly_numeric(series: pd.Series) -> bool:
    cleaned = series.astype(str).str.strip()
    cleaned = cleaned[~cleaned.isin(["", "nan", "None"])]
    if cleaned.empty:
        return False
    numeric_ratio = pd.to_numeric(cleaned, errors="coerce").notna().mean()
    return numeric_ratio >= 0.8


def sort_period_labels(labels: list[str], freq: str) -> list[str]:
    known = [label for label in labels if label != "Unknown"]
    ordered = sorted(known, key=lambda x: pd.Period(x, freq=freq))
    if "Unknown" in labels:
        ordered.append("Unknown")
    return ordered


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, data in sheets.items():
            data.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return buffer.getvalue()


def performance_status(actual: float, target: float) -> str:
    if target <= 0:
        return "No target set"
    ratio = actual / target
    if ratio >= 1:
        return "On track (target achieved)"
    if ratio >= 0.8:
        return "Needs attention (close to target)"
    return "Off track (below target)"


def format_compact(value: float) -> str:
    abs_val = abs(value)
    if abs_val >= 1_000_000_000:
        return f"{value / 1_000_000_000:.2f}B"
    if abs_val >= 1_000_000:
        return f"{value / 1_000_000:.2f}M"
    if abs_val >= 1_000:
        return f"{value / 1_000:.1f}K"
    return f"{value:,.2f}" if value % 1 else f"{value:,.0f}"


def normalize_flag(series: pd.Series) -> pd.Series:
    truthy = {"1", "y", "yes", "true", "completed", "done", "pass", "passed"}
    return series.astype(str).str.strip().str.lower().isin(truthy)


def evaluate_sunlife_validation(row: pd.Series) -> dict[str, object]:
    cls = str(row.get("Classification", "")).strip().upper()
    tenure_text = str(row.get("Tenure Raw", "")).strip().lower()
    ac_nsc = float(row.get("AC+NSC", 0) or 0)
    coding_q = int(row.get("Coding Quarter", 0) or 0)
    jfw = bool(row.get("JFW Done", False))
    start = bool(row.get("START Done", False))
    pillars = bool(row.get("Pillars Done", False))
    vul = bool(row.get("VUL Advance Done", False))
    mandatory_training = bool(row.get("Mandatory Training Done", False))

    is_rookie = ("year 0" in tenure_text) or ("rookie" in tenure_text) or ("external" in tenure_text)
    is_external_mc = cls == "MC" and "external" in tenure_text
    training_bundle_ok = jfw and start and pillars

    requirement = "No mapped rule"
    passed = False
    reason = "Classification/tenure combination not mapped to a rule."

    if is_external_mc:
        requirement = "External MC: >=45K AC/NSC + JFW + START + 4 Pillars"
        passed = ac_nsc >= 45_000 and training_bundle_ok
        reason = "Meets requirement." if passed else "Needs AC/NSC >=45K and complete JFW/START/4 Pillars."
    elif cls == "A":
        threshold = 15_000 if coding_q == 4 else 45_000
        requirement = f"Rookie A: >={threshold:,.0f} AC/NSC + JFW + START + 4 Pillars"
        passed = ac_nsc >= threshold and training_bundle_ok
        reason = "Meets requirement." if passed else "Needs AC/NSC threshold and complete JFW/START/4 Pillars."
    elif cls == "B":
        requirement = "Tenured B: >=90K AC/NSC + VUL Advance"
        passed = ac_nsc >= 90_000 and vul
        reason = "Meets requirement." if passed else "Needs AC/NSC >=90K and VUL Advance completion."
    elif cls == "C":
        requirement = "Tenured C: >=135K AC/NSC"
        passed = ac_nsc >= 135_000
        reason = "Meets requirement." if passed else "Needs AC/NSC >=135K."
    elif cls in {"D", "E", "MC"}:
        requirement = "Tenured D/E/MC: >=180K AC/NSC"
        passed = ac_nsc >= 180_000
        reason = "Meets requirement." if passed else "Needs AC/NSC >=180K."
    elif cls == "F":
        requirement = "F advisor: optional, counted as VMP if >=180K AC/NSC"
        passed = ac_nsc >= 180_000
        reason = "Counts as VMP via >=180K AC/NSC." if passed else "Below optional 180K VMP threshold."

    vna = is_rookie and passed
    rookie_vmp = is_rookie and (ac_nsc >= 90_000) and training_bundle_ok
    sm_eligible = ac_nsc >= 180_000 and mandatory_training
    vmp = passed or rookie_vmp

    return {
        "Validation Requirement": requirement,
        "Validation Status": "Pass" if passed else "Fail",
        "Validation Reason": reason,
        "VNA Eligible": "Yes" if vna else "No",
        "VMP Eligible": "Yes" if vmp else "No",
        "SM Appointment Eligible": "Yes" if sm_eligible else "No",
    }


uploaded_files = st.file_uploader(
    "Upload Excel file(s)", type=["xlsx", "xlsm", "xls"], accept_multiple_files=True
)

# Streamlit re-runs the script on every widget change. File uploads are usually
# preserved by the widget, but storing bytes in session_state makes this
# behavior explicit and avoids losing them on re-runs.
if uploaded_files:
    file_payloads = [(f.name, f.read()) for f in uploaded_files]
    st.session_state["uploaded_file_payloads"] = file_payloads
else:
    file_payloads = st.session_state.get("uploaded_file_payloads")

if not file_payloads:
    st.info("📂 Upload one or more files to start.")
    st.stop()

excel = pd.ExcelFile(io.BytesIO(file_payloads[0][1]))
preferred_sheets = ["Settled Apps - Details", "Submitted Apps - Details"]
default_sheet_index = 0
for preferred_name in preferred_sheets:
    if preferred_name in excel.sheet_names:
        default_sheet_index = excel.sheet_names.index(preferred_name)
        break

sheet = st.selectbox("Sheet", excel.sheet_names, index=default_sheet_index)
header_row = st.number_input(
    "Header row number in Excel (1-based)",
    min_value=1,
    value=12,
    step=1,
    help="Set this to 12 if your actual column names start on row 12.",
)
column_range = st.text_input(
    "Column range to read (Excel style)",
    value="",
    help="Optional. Example: A:Z or A:AA. Leave blank to read all available columns.",
)
requested_range = column_range.strip()
loaded_frames = []
loaded_file_names = []
skipped_file_names = []
for file_name, file_bytes in file_payloads:
    try:
        frame = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=sheet,
            header=int(header_row) - 1,
        )
        if requested_range:
            try:
                frame = pd.read_excel(
                    io.BytesIO(file_bytes),
                    sheet_name=sheet,
                    header=int(header_row) - 1,
                    usecols=requested_range,
                )
            except ParserError:
                st.warning(
                    f"`{file_name}`: column range `{requested_range}` is wider than this sheet. "
                    "Loaded all available columns instead."
                )
        frame = frame.dropna(how="all")
        frame.columns = [str(c).strip() for c in frame.columns]
        frame = frame.loc[:, ~pd.Index(frame.columns).str.startswith("Unnamed:")]
        if not frame.empty:
            frame["Source File"] = file_name
            loaded_frames.append(frame)
            loaded_file_names.append(file_name)
    except ValueError:
        skipped_file_names.append(file_name)

if skipped_file_names:
    st.warning(
        f"Skipped {len(skipped_file_names)} file(s) that do not have sheet `{sheet}`: "
        + ", ".join(skipped_file_names)
    )

if not loaded_frames:
    st.warning("No valid rows were loaded from uploaded files.")
    st.stop()

raw_df = pd.concat(loaded_frames, ignore_index=True, sort=False)

if raw_df.empty:
    st.warning("Selected sheet is empty across uploaded files.")
    st.stop()

st.caption(
    f"📊 Loaded {len(loaded_file_names)} file(s), {len(raw_df):,} rows. "
    f"Detected columns: {', '.join(raw_df.columns)}"
)
with st.expander("📁 Loaded files summary", expanded=False):
    file_summary = (
        raw_df.groupby("Source File", dropna=False)
        .size()
        .reset_index(name="Rows Loaded")
        .sort_values("Rows Loaded", ascending=False)
    )
    st.dataframe(file_summary, use_container_width=True, hide_index=True)

render_section("1) Data Mapping", "Map source columns to advisor, classification, metrics, and date.")
all_cols = list(raw_df.columns)
none_opt = ["(none)"] + all_cols

advisor_guess = guess_column(all_cols, ["advisor name", "advisor", "agent name", "agent"])
class_guess = guess_column(all_cols, ["classification", "class", "rank", "tier"])
ac_guess = guess_column(all_cols, ["agency credits", "ac", "annualized", "ape"])
nsc_guess = guess_column(all_cols, ["net sales credits", "nsc", "net sales"])
lives_guess = guess_column(all_cols, ["settled apps", "submitted apps", "lives", "life", "lives count"])
date_guess = guess_column(all_cols, ["process date", "date", "issue date", "submitted date"])

# Standard mapping for your provided row-12 format.
if "Advisor Name" in all_cols:
    advisor_guess = "Advisor Name"
if "Class Code" in all_cols:
    class_guess = "Class Code"
if "Agency Credits" in all_cols:
    ac_guess = "Agency Credits"
if "Net Sales Credits" in all_cols:
    nsc_guess = "Net Sales Credits"
if "Process Date" in all_cols:
    date_guess = "Process Date"
if sheet == "Settled Apps - Details" and "Settled Apps" in all_cols:
    lives_guess = "Settled Apps"
elif sheet == "Submitted Apps - Details" and "Submitted Apps" in all_cols:
    lives_guess = "Submitted Apps"

# If class guess is mostly numeric, prefer Unit when available (usually A/B/C bucket).
if "Unit" in all_cols and series_is_mostly_numeric(raw_df[class_guess]):
    unit_values = raw_df["Unit"].astype(str).str.strip()
    unit_values = unit_values[~unit_values.isin(["", "nan", "None"])]
    if not unit_values.empty and not series_is_mostly_numeric(unit_values):
        class_guess = "Unit"

advisor_col = st.selectbox(
    "Advisor name column", all_cols, index=all_cols.index(advisor_guess), key="advisor_col"
)
class_col = st.selectbox(
    "Classification column (A/B/C/etc.)",
    all_cols,
    index=all_cols.index(class_guess),
    key="class_col",
)
lives_help = "Use 'Settled Apps' or 'Submitted Apps' for lives."
lives_col = st.selectbox(
    "Lives column (Settled Apps / Submitted Apps)",
    all_cols,
    index=all_cols.index(lives_guess),
    key="lives_col",
    help=lives_help,
)
ac_col = None
nsc_col = None
if sheet != "Submitted Apps - Details":
    ac_index = none_opt.index(ac_guess) if ac_guess in none_opt else 0
    ac_pick = st.selectbox("AC column", none_opt, index=ac_index, key="ac_col")
    nsc_index = none_opt.index(nsc_guess) if nsc_guess in none_opt else 0
    nsc_pick = st.selectbox("NSC column", none_opt, index=nsc_index, key="nsc_col")
    ac_col = None if ac_pick == "(none)" else ac_pick
    nsc_col = None if nsc_pick == "(none)" else nsc_pick
else:
    st.caption("`Submitted Apps - Details` has no AC/NSC columns. AC and NSC are set to 0.")

st.markdown("**Date mapping (choose either Date OR Month+Year)**")
date_col_pick = st.selectbox(
    "Date column", none_opt, index=none_opt.index(date_guess) if date_guess in none_opt else 0
)
date_col = None if date_col_pick == "(none)" else date_col_pick
month_col = None
year_col = None
if not date_col:
    month_col_pick = st.selectbox("Month column", none_opt, index=0)
    year_col_pick = st.selectbox("Year column", none_opt, index=0)
    month_col = None if month_col_pick == "(none)" else month_col_pick
    year_col = None if year_col_pick == "(none)" else year_col_pick
else:
    st.caption("Using Date column only. Month/Year mapping hidden.")
    month_col = None
    year_col = None

st.markdown("**Validation mapping (optional, for Sun Life standards)**")
tenure_guess = guess_column(all_cols, ["tenure", "segment", "advisor tenure"])
coding_guess = guess_column(all_cols, ["coding date", "coded date", "date coded", "process date"])
jfw_guess = guess_column(all_cols, ["jfw"])
start_guess = guess_column(all_cols, ["start"])
pillars_guess = guess_column(all_cols, ["pillar", "4 pillars"])
vul_guess = guess_column(all_cols, ["vul advance", "vul"])
mandatory_guess = guess_column(all_cols, ["mandatory training", "mandatory"])

tenure_pick = st.selectbox(
    "Tenure/segment column", none_opt, index=none_opt.index(tenure_guess) if tenure_guess in none_opt else 0
)
coding_pick = st.selectbox(
    "Coding date column",
    none_opt,
    index=none_opt.index(coding_guess) if coding_guess in none_opt else 0,
    help="Used for Q4-coded Rookie A rule.",
)
jfw_pick = st.selectbox("JFW completed column", none_opt, index=none_opt.index(jfw_guess) if jfw_guess in none_opt else 0)
start_pick = st.selectbox(
    "START completed column", none_opt, index=none_opt.index(start_guess) if start_guess in none_opt else 0
)
pillars_pick = st.selectbox(
    "4 Pillars completed column",
    none_opt,
    index=none_opt.index(pillars_guess) if pillars_guess in none_opt else 0,
)
vul_pick = st.selectbox(
    "VUL Advance completed column", none_opt, index=none_opt.index(vul_guess) if vul_guess in none_opt else 0
)
mandatory_pick = st.selectbox(
    "Mandatory training completed column",
    none_opt,
    index=none_opt.index(mandatory_guess) if mandatory_guess in none_opt else 0,
)

tenure_col = None if tenure_pick == "(none)" else tenure_pick
coding_validation_col = None if coding_pick == "(none)" else coding_pick
jfw_col = None if jfw_pick == "(none)" else jfw_pick
start_col = None if start_pick == "(none)" else start_pick
pillars_col = None if pillars_pick == "(none)" else pillars_pick
vul_col = None if vul_pick == "(none)" else vul_pick
mandatory_col = None if mandatory_pick == "(none)" else mandatory_pick

required_mapping = [advisor_col, class_col, lives_col]
optional_mapping = [c for c in [ac_col, nsc_col] if c]
core_mapping = required_mapping + optional_mapping
if len(set(core_mapping)) < len(core_mapping):
    st.error(
        "Please map different columns for Advisor, Classification, Lives, and optional AC/NSC. "
        "Right now some mappings are duplicated."
    )
    st.stop()

if not date_col and not (month_col and year_col):
    st.warning("Please choose either a Date column or both Month and Year columns.")
    st.stop()

with st.expander("Preview mapped columns"):
    preview_cols = [advisor_col, class_col, lives_col]
    if ac_col:
        preview_cols.append(ac_col)
    if nsc_col:
        preview_cols.append(nsc_col)
    if tenure_col:
        preview_cols.append(tenure_col)
    if coding_validation_col:
        preview_cols.append(coding_validation_col)
    if date_col:
        preview_cols.append(date_col)
    elif month_col and year_col:
        preview_cols.extend([month_col, year_col])
    # Avoid duplicate columns in preview (e.g., Process Date chosen twice).
    preview_cols = list(dict.fromkeys(preview_cols))
    st.dataframe(raw_df[preview_cols].head(10), use_container_width=True)

df = raw_df.copy()
df["Advisor"] = df[advisor_col].astype(str).str.strip()
df["Classification"] = df[class_col].astype(str).str.strip().str.upper()
df["AC"] = normalize_number(df[ac_col]) if ac_col else 0
df["NSC"] = normalize_number(df[nsc_col]) if nsc_col else 0
df["Lives"] = normalize_number(df[lives_col])
df["Tenure Raw"] = df[tenure_col].astype(str).str.strip() if tenure_col else ""
if coding_validation_col:
    coding_dt = pd.to_datetime(df[coding_validation_col], errors="coerce")
elif date_col:
    coding_dt = pd.to_datetime(df[date_col], errors="coerce")
else:
    coding_dt = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
df["Coding Quarter"] = coding_dt.dt.quarter.fillna(0).astype(int)
df["JFW Done"] = normalize_flag(df[jfw_col]) if jfw_col else False
df["START Done"] = normalize_flag(df[start_col]) if start_col else False
df["Pillars Done"] = normalize_flag(df[pillars_col]) if pillars_col else False
df["VUL Advance Done"] = normalize_flag(df[vul_col]) if vul_col else False
df["Mandatory Training Done"] = normalize_flag(df[mandatory_col]) if mandatory_col else False
if sheet != "Submitted Apps - Details" and (not ac_col or not nsc_col):
    st.info("AC/NSC not available in this sheet; missing values are treated as 0.")
df = build_period_columns(df, date_col, month_col, year_col)
df = df[
    ~df["Advisor"].str.lower().str.contains("total", na=False)
    & ~df["Classification"].str.lower().str.contains("total", na=False)
].copy()

if series_is_mostly_numeric(df["Classification"]):
    st.info(
        "Classification values look numeric codes. If you want A/B/C groups, choose `Unit` "
        "or provide your class-code-to-A/B/C mapping rules."
    )

render_section("2) Analysis Filters", "Narrow down classification and time periods for focused insights.")
classes = sorted({str(c) for c in df["Classification"].dropna() if str(c).strip()})
period_months = sort_period_labels(
    list({str(p) for p in df["Period Month"].fillna("Unknown")}), freq="M"
)
period_quarters = sort_period_labels(
    list({str(p) for p in df["Period Quarter"].fillna("Unknown")}), freq="Q"
)

st.sidebar.header("📊 Dashboard Filters")
with st.sidebar.expander("ℹ️ How to use this dashboard", expanded=False):
    st.markdown(
        "- Upload one or more files with the same layout.\n"
        "- Pick the target detail sheet and map columns once.\n"
        "- Use filters to focus the story by period/classification.\n"
        "- Download full Excel report with charts and summaries."
    )
selected_classes = st.sidebar.multiselect("Classification", classes, default=classes)
selected_months = st.sidebar.multiselect("Months", period_months, default=period_months)
selected_quarters = st.sidebar.multiselect("Quarters", period_quarters, default=period_quarters)
top_n = st.sidebar.slider("Top advisors to show", min_value=5, max_value=30, value=10, step=1)
ranking_metric = st.sidebar.selectbox("Advisor ranking metric", ["AC", "NSC", "Lives"], index=0)
st.sidebar.header("🎯 Performance Targets")
target_ac = st.sidebar.number_input("Target AC", min_value=0.0, value=0.0, step=1000.0)
target_nsc = st.sidebar.number_input("Target NSC", min_value=0.0, value=0.0, step=1000.0)
target_lives = st.sidebar.number_input("Target Lives", min_value=0.0, value=0.0, step=10.0)

filtered = df[
    df["Classification"].isin(selected_classes)
    & df["Period Month"].isin(selected_months)
    & df["Period Quarter"].isin(selected_quarters)
].copy()

if filtered.empty:
    st.warning("No rows match your filters.")
    st.stop()

render_section("3) Results & Story", "Executive summary, drivers, trends, data quality, and exports.")

monthly_summary = (
    filtered.groupby(["Period Month", "Classification"], dropna=False)[["AC", "NSC", "Lives"]]
    .sum()
    .reset_index()
    .sort_values(["Period Month", "Classification"])
)

quarterly_summary = (
    filtered.groupby(["Period Quarter", "Classification"], dropna=False)[["AC", "NSC", "Lives"]]
    .sum()
    .reset_index()
    .sort_values(["Period Quarter", "Classification"])
)

classification_summary = (
    filtered.groupby("Classification", dropna=False)[["AC", "NSC", "Lives"]]
    .sum()
    .reset_index()
    .sort_values("AC", ascending=False)
)

advisor_detail = (
    filtered.groupby(
        ["Period Month", "Period Quarter", "Classification", "Advisor"], dropna=False
    )[["AC", "NSC", "Lives"]]
    .sum()
    .reset_index()
    .sort_values(["Period Month", "Classification", "Advisor"])
)

validation_base = filtered.copy()
validation_base["AC+NSC"] = validation_base["AC"] + validation_base["NSC"]
advisor_validation = (
    validation_base.groupby(["Advisor", "Classification"], dropna=False)
    .agg(
        {
            "Tenure Raw": "first",
            "Coding Quarter": "max",
            "AC": "sum",
            "NSC": "sum",
            "AC+NSC": "sum",
            "Lives": "sum",
            "JFW Done": "max",
            "START Done": "max",
            "Pillars Done": "max",
            "VUL Advance Done": "max",
            "Mandatory Training Done": "max",
        }
    )
    .reset_index()
)
validation_cols = advisor_validation.apply(evaluate_sunlife_validation, axis=1, result_type="expand")
advisor_validation = pd.concat([advisor_validation, validation_cols], axis=1).sort_values(
    ["Validation Status", "AC+NSC", "Advisor"], ascending=[True, False, True]
)

render_section("Story Dashboard", "Executive summary, validation snapshot, drivers, trends, and details.")
total_ac = filtered["AC"].sum()
total_nsc = filtered["NSC"].sum()
total_lives = filtered["Lives"].sum()
active_advisors = filtered["Advisor"].nunique()
record_count = len(filtered)
quality_invalid_dates = int((filtered["Period Month"] == "Unknown").sum())
quality_missing_advisor = int(filtered["Advisor"].astype(str).str.strip().isin(["", "nan", "None"]).sum())
quality_missing_class = int(
    filtered["Classification"].astype(str).str.strip().isin(["", "nan", "None"]).sum()
)
quality_negative_values = int(((filtered["AC"] < 0) | (filtered["NSC"] < 0) | (filtered["Lives"] < 0)).sum())

monthly_kpi = (
    filtered.groupby("Period Month", dropna=False)[["AC", "NSC", "Lives"]]
    .sum()
    .reset_index()
)
monthly_kpi["SortMonth"] = pd.PeriodIndex(
    monthly_kpi["Period Month"].where(monthly_kpi["Period Month"] != "Unknown", "1900-01"),
    freq="M",
)
monthly_kpi = monthly_kpi.sort_values("SortMonth").drop(columns=["SortMonth"])

known_monthly = monthly_kpi[monthly_kpi["Period Month"] != "Unknown"].copy()
latest_month = None
previous_month = None
if len(known_monthly) >= 1:
    latest_month = known_monthly.iloc[-1]
if len(known_monthly) >= 2:
    previous_month = known_monthly.iloc[-2]

delta_ac = None
delta_nsc = None
delta_lives = None
if latest_month is not None and previous_month is not None:
    delta_ac = latest_month["AC"] - previous_month["AC"]
    delta_nsc = latest_month["NSC"] - previous_month["NSC"]
    delta_lives = latest_month["Lives"] - previous_month["Lives"]

kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
kpi1.metric("Total AC", format_compact(total_ac), "" if delta_ac is None else format_compact(delta_ac))
kpi2.metric("Total NSC", format_compact(total_nsc), "" if delta_nsc is None else format_compact(delta_nsc))
kpi3.metric(
    "Total Lives",
    format_compact(total_lives),
    "" if delta_lives is None else format_compact(delta_lives),
)
kpi4.metric("Active Advisors", f"{active_advisors:,}")
kpi5.metric("Records", f"{record_count:,}")

status_ac = performance_status(total_ac, target_ac)
status_nsc = performance_status(total_nsc, target_nsc)
status_lives = performance_status(total_lives, target_lives)


def build_status_row(metric: str, actual: float, target: float, status: str) -> dict[str, object]:
    if target <= 0:
        return {
            "Metric": metric,
            "Actual": actual,
            "Target": target,
            "Achievement %": "",
            "Gap to Target": "",
            "Status": status,
            "What it means": "Add a target to evaluate this metric.",
        }
    achievement_pct = (actual / target) * 100
    gap = actual - target
    if achievement_pct >= 100:
        meaning = "Met or exceeded target."
    elif achievement_pct >= 80:
        meaning = "Near target; improve to close the gap."
    else:
        meaning = "Far from target; needs focused action."
    return {
        "Metric": metric,
        "Actual": actual,
        "Target": target,
        "Achievement %": f"{achievement_pct:.1f}%",
        "Gap to Target": gap,
        "Status": status,
        "What it means": meaning,
    }


status_df = pd.DataFrame(
    [
        build_status_row("AC", total_ac, target_ac, status_ac),
        build_status_row("NSC", total_nsc, target_nsc, status_nsc),
        build_status_row("Lives", total_lives, target_lives, status_lives),
    ]
)

classification_ac_chart = (
    monthly_summary.pivot(index="Period Month", columns="Classification", values="AC")
    .fillna(0)
    .reindex([m for m in period_months if m in set(monthly_summary["Period Month"])])
)

top_advisors = (
    advisor_detail.groupby("Advisor", dropna=False)[["AC", "NSC", "Lives"]]
    .sum()
    .reset_index()
    .sort_values(ranking_metric, ascending=False)
    .head(top_n)
)

top_class = classification_summary.iloc[0] if not classification_summary.empty else None
top_advisor_row = top_advisors.iloc[0] if not top_advisors.empty else None
coverage = (top_class["AC"] / total_ac * 100) if top_class is not None and total_ac else 0

if latest_month is not None:
    headline = (
        f"In **{latest_month['Period Month']}**, production reached "
        f"**AC {latest_month['AC']:,.2f}**, **NSC {latest_month['NSC']:,.2f}**, "
        f"and **Lives {latest_month['Lives']:,.0f}**."
    )
    st.markdown(headline)

insight_lines = []
if top_class is not None:
    insight_lines.append(
        f"- Top classification: **{top_class['Classification']}** contributed "
        f"**AC {top_class['AC']:,.2f}** ({coverage:.1f}% of total AC)."
    )
if top_advisor_row is not None:
    insight_lines.append(
        f"- Leading advisor by {ranking_metric}: **{top_advisor_row['Advisor']}** "
        f"with **{ranking_metric} {top_advisor_row[ranking_metric]:,.2f}**."
    )
if previous_month is not None and latest_month is not None:
    trend_word = "up" if (delta_ac is not None and delta_ac >= 0) else "down"
    insight_lines.append(
        f"- Month-over-month AC is **{trend_word} {abs(delta_ac):,.2f}** "
        f"from {previous_month['Period Month']} to {latest_month['Period Month']}."
    )
if insight_lines:
    st.markdown("### Executive Insights")
    st.markdown("\n".join(insight_lines))

validation_pass_count = int((advisor_validation["Validation Status"] == "Pass").sum())
validation_total = len(advisor_validation)
vna_count = int((advisor_validation["VNA Eligible"] == "Yes").sum())
vmp_count = int((advisor_validation["VMP Eligible"] == "Yes").sum())
sm_eligible_count = int((advisor_validation["SM Appointment Eligible"] == "Yes").sum())
st.markdown(
    f"### Validation Snapshot\n"
    f"- Passed standards: **{validation_pass_count}/{validation_total}** advisors\n"
    f"- VNA eligible: **{vna_count}** | VMP eligible: **{vmp_count}** | SM-appointment eligible: **{sm_eligible_count}**"
)

story_tab, drivers_tab, trends_tab, details_tab = st.tabs(
    ["📖 Story", "🚀 Drivers", "📈 Trends", "🔍 Details"]
)
with story_tab:
    st.markdown("**Performance status vs target**")
    st.caption(
        "Status guide: On track = target achieved; Needs attention = at least 80% of target; "
        "Off track = below 80% of target."
    )
    st.dataframe(status_df, use_container_width=True, hide_index=True)
    st.markdown("**Performance composition by classification**")
    comp_col1, comp_col2 = st.columns(2)
    with comp_col1:
        st.bar_chart(
            classification_summary.set_index("Classification")[["AC", "NSC", "Lives"]],
            height=330,
        )
    with comp_col2:
        st.markdown(f"**Top {top_n} advisors by {ranking_metric}**")
        st.bar_chart(top_advisors.set_index("Advisor")[["AC", "NSC", "Lives"]], height=330)

with drivers_tab:
    st.markdown("**Who drives production?**")
    left_col, right_col = st.columns(2)
    with left_col:
        class_share = classification_summary.copy()
        class_share["AC Share %"] = (
            class_share["AC"] / class_share["AC"].sum() * 100 if class_share["AC"].sum() else 0
        )
        st.dataframe(class_share, use_container_width=True, height=340)
    with right_col:
        advisor_rank = (
            advisor_detail.groupby(["Classification", "Advisor"], dropna=False)[["AC", "NSC", "Lives"]]
            .sum()
            .reset_index()
            .sort_values(ranking_metric, ascending=False)
            .head(top_n)
        )
        st.dataframe(advisor_rank, use_container_width=True, height=340)

with trends_tab:
    trend_col1, trend_col2 = st.columns(2)
    with trend_col1:
        st.markdown("**Monthly momentum (AC / NSC / Lives)**")
        st.line_chart(monthly_kpi.set_index("Period Month")[["AC", "NSC", "Lives"]], height=320)
    with trend_col2:
        st.markdown("**Monthly AC by classification**")
        st.bar_chart(classification_ac_chart, height=320)
    st.markdown("**Quarterly AC trend by classification**")
    quarterly_kpi = quarterly_summary.pivot(
        index="Period Quarter", columns="Classification", values="AC"
    ).fillna(0)
    st.area_chart(quarterly_kpi, height=280)

with details_tab:
    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        ["Monthly Summary", "Quarterly Summary", "Advisor Details", "Validation Results", "Data Quality"]
    )
    with tab1:
        st.dataframe(monthly_summary, use_container_width=True)
    with tab2:
        st.dataframe(quarterly_summary, use_container_width=True)
    with tab3:
        st.dataframe(advisor_detail, use_container_width=True, height=420)
    with tab4:
        st.dataframe(advisor_validation, use_container_width=True, height=420)
    with tab5:
        quality_df = pd.DataFrame(
            [
                {"Check": "Rows with unknown/invalid date", "Count": quality_invalid_dates},
                {"Check": "Rows with missing advisor", "Count": quality_missing_advisor},
                {"Check": "Rows with missing classification", "Count": quality_missing_class},
                {"Check": "Rows with negative AC/NSC/Lives", "Count": quality_negative_values},
            ]
        )
        st.dataframe(quality_df, use_container_width=True, hide_index=True)

quality_df = pd.DataFrame(
    [
        {"Check": "Rows with unknown/invalid date", "Count": quality_invalid_dates},
        {"Check": "Rows with missing advisor", "Count": quality_missing_advisor},
        {"Check": "Rows with missing classification", "Count": quality_missing_class},
        {"Check": "Rows with negative AC/NSC/Lives", "Count": quality_negative_values},
    ]
)
executive_summary = pd.DataFrame(
    [
        {"Metric": "Total AC", "Value": total_ac},
        {"Metric": "Total NSC", "Value": total_nsc},
        {"Metric": "Total Lives", "Value": total_lives},
        {"Metric": "Active Advisors", "Value": active_advisors},
        {"Metric": "Records", "Value": record_count},
    ]
)
export_bytes = to_excel_bytes(
    {
        "Executive Summary": executive_summary,
        "Target Status": status_df,
        "Data Quality": quality_df,
        "Monthly Summary": monthly_summary,
        "Quarterly Summary": quarterly_summary,
        "Advisor Details": advisor_detail,
        "Validation Results": advisor_validation,
        "Filtered Raw Data": filtered,
    }
)

st.markdown("**📥 Download**")
download_col1, download_col2 = st.columns(2)
with download_col1:
    st.download_button(
        "Download full report (Excel)",
        data=export_bytes,
        file_name="manpower_validation_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with download_col2:
    st.download_button(
        "Download advisor details (CSV)",
        data=advisor_detail.to_csv(index=False).encode("utf-8"),
        file_name="advisor_details.csv",
        mime="text/csv",
    )