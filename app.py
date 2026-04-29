import io
from typing import Optional

import pandas as pd
import streamlit as st
from pandas.errors import ParserError


st.set_page_config(page_title="Manpower Validation System", layout="wide")
st.markdown(
    """
    <style>
        :root {
            --sl-yellow: #ffd100;
            --sl-yellow-soft: #fff3b3;
            --sl-blue: #003da5;
            --sl-blue-dark: #002a73;
            --sl-gray: #f7f9fc;
            --sl-text: #1b1f24;
        }

        .stApp {
            background: linear-gradient(180deg, #fffef8 0%, #ffffff 220px);
            color: var(--sl-text);
        }

        div[data-testid="stSidebar"] {
            background: linear-gradient(180deg, var(--sl-blue) 0%, var(--sl-blue-dark) 100%);
        }
        div[data-testid="stSidebar"] * {
            color: #ffffff !important;
        }
        div[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h1,
        div[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2,
        div[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3 {
            color: var(--sl-yellow) !important;
        }

        .sl-hero {
            padding: 14px 18px;
            border-radius: 14px;
            border-left: 8px solid var(--sl-yellow);
            background: linear-gradient(90deg, #ffffff 0%, #fffbe6 100%);
            box-shadow: 0 4px 14px rgba(0, 61, 165, 0.12);
            margin-bottom: 10px;
        }
        .sl-hero-title {
            color: var(--sl-blue);
            font-weight: 700;
            font-size: 1.2rem;
            margin-bottom: 3px;
        }
        .sl-hero-sub {
            color: #334155;
            font-size: 0.95rem;
        }

        div[data-testid="stMetric"] {
            background: #ffffff;
            border: 1px solid #cfd8ea;
            border-radius: 12px;
            padding: 12px 14px;
            box-shadow: 0 4px 12px rgba(0, 61, 165, 0.14);
        }
        div[data-testid="stMetricLabel"] {
            color: var(--sl-blue-dark);
            font-weight: 700;
            font-size: 0.95rem;
        }
        div[data-testid="stMetricValue"] {
            color: #0b132b;
            font-weight: 800;
            font-size: 2rem;
            line-height: 1.2;
            white-space: normal !important;
            overflow: visible !important;
            text-overflow: clip !important;
            word-break: break-word;
        }
        div[data-testid="stMetricDelta"] {
            font-size: 0.9rem;
            font-weight: 700;
        }

        div[data-baseweb="tab-list"] {
            background: var(--sl-gray);
            border-radius: 10px;
            padding: 6px;
        }
        button[role="tab"][aria-selected="true"] {
            background: var(--sl-yellow) !important;
            color: #111827 !important;
            border-radius: 8px !important;
        }

        .stDownloadButton button {
            background: var(--sl-blue);
            color: #ffffff;
            border: 1px solid var(--sl-blue);
            border-radius: 8px;
        }
        .stDownloadButton button:hover {
            border-color: var(--sl-yellow);
            box-shadow: 0 0 0 2px rgba(255, 209, 0, 0.25);
        }
    </style>
    """,
    unsafe_allow_html=True,
)
st.title("Manpower Validation System")
st.markdown(
    """
    <div class="sl-hero">
        <div class="sl-hero-title">Manpower Validation System</div>
        <div class="sl-hero-sub">
            Upload an Excel file and analyze production by advisor classification with
            AC, NSC, and Lives across monthly and quarterly views.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


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


uploaded_files = st.file_uploader(
    "Upload Excel file(s)", type=["xlsx", "xlsm", "xls"], accept_multiple_files=True
)

if not uploaded_files:
    st.info("Upload one or more files to start.")
    st.stop()

file_payloads = [(f.name, f.read()) for f in uploaded_files]
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
    f"Loaded {len(loaded_file_names)} file(s), {len(raw_df):,} rows. "
    f"Detected columns: {', '.join(raw_df.columns)}"
)

st.subheader("1) Map your columns")
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
    if date_col:
        preview_cols.append(date_col)
    elif month_col and year_col:
        preview_cols.extend([month_col, year_col])
    st.dataframe(raw_df[preview_cols].head(10), use_container_width=True)

df = raw_df.copy()
df["Advisor"] = df[advisor_col].astype(str).str.strip()
df["Classification"] = df[class_col].astype(str).str.strip().str.upper()
df["AC"] = normalize_number(df[ac_col]) if ac_col else 0
df["NSC"] = normalize_number(df[nsc_col]) if nsc_col else 0
df["Lives"] = normalize_number(df[lives_col])
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

st.subheader("2) Filters")
classes = sorted({str(c) for c in df["Classification"].dropna() if str(c).strip()})
period_months = sort_period_labels(
    list({str(p) for p in df["Period Month"].fillna("Unknown")}), freq="M"
)
period_quarters = sort_period_labels(
    list({str(p) for p in df["Period Quarter"].fillna("Unknown")}), freq="Q"
)

st.sidebar.header("Dashboard Filters")
selected_classes = st.sidebar.multiselect("Classification", classes, default=classes)
selected_months = st.sidebar.multiselect("Months", period_months, default=period_months)
selected_quarters = st.sidebar.multiselect("Quarters", period_quarters, default=period_quarters)
top_n = st.sidebar.slider("Top advisors to show", min_value=5, max_value=30, value=10, step=1)
ranking_metric = st.sidebar.selectbox("Advisor ranking metric", ["AC", "NSC", "Lives"], index=0)
st.sidebar.header("Performance Targets")
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

st.subheader("3) Results")

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

st.subheader("Story Dashboard")
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

story_tab, drivers_tab, trends_tab, details_tab = st.tabs(
    ["Story", "Drivers", "Trends", "Details"]
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
    tab1, tab2, tab3, tab4 = st.tabs(
        ["Monthly Summary", "Quarterly Summary", "Advisor Details", "Data Quality"]
    )
    with tab1:
        st.dataframe(monthly_summary, use_container_width=True)
    with tab2:
        st.dataframe(quarterly_summary, use_container_width=True)
    with tab3:
        st.dataframe(advisor_detail, use_container_width=True, height=420)
    with tab4:
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
        "Filtered Raw Data": filtered,
    }
)

st.markdown("**Download**")
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

