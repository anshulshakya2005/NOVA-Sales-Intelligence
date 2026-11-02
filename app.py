
import io
import json
import textwrap
import warnings
from typing import Optional

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

warnings.filterwarnings("ignore")

# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="NOVA | Sales Intelligence",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# CUSTOM CSS - PREMIUM UI
# ============================================================

st.markdown("""
<style>

/* Main application */
.stApp {
    background:
        radial-gradient(circle at 85% 5%, #17223d 0%, transparent 30%),
        linear-gradient(135deg, #080b18 0%, #0d1020 55%, #111329 100%);
    color: #f5f6ff;
}

/* Remove default header */
header[data-testid="stHeader"] {
    background: rgba(8, 11, 24, 0.8);
}

/* Main content spacing */
.block-container {
    padding-top: 2rem;
    padding-bottom: 3rem;
    max-width: 1500px;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #11162b, #0b1020);
    border-right: 1px solid #252b48;
}

section[data-testid="stSidebar"] * {
    color: #e7eaff;
}

/* Sidebar divider */
section[data-testid="stSidebar"] hr {
    border-color: #343b5c;
}

/* Hero section */
.hero {
    position: relative;
    overflow: hidden;
    padding: 48px 52px;
    border-radius: 25px;
    background:
        radial-gradient(circle at 85% 20%, rgba(120, 75, 255, .28), transparent 35%),
        linear-gradient(120deg, #201b48, #122a3d 65%, #241a3b);
    border: 1px solid rgba(151, 137, 255, .35);
    box-shadow: 0 20px 65px rgba(0, 0, 0, .3);
    margin-bottom: 25px;
}

.hero:after {
    content: "✦";
    position: absolute;
    right: 65px;
    top: 15px;
    font-size: 150px;
    color: rgba(255,255,255,.045);
}

.eyebrow {
    color: #b6a8ff;
    font-size: 12px;
    letter-spacing: 3px;
    font-weight: 700;
    margin-bottom: 22px;
}

.hero h1 {
    font-size: clamp(32px, 4vw, 53px);
    line-height: 1.12;
    color: #ffffff;
    font-weight: 800;
    margin: 0 0 18px 0;
}

.hero p {
    font-size: 16px;
    color: #c2c8e5;
    max-width: 550px;
    line-height: 1.8;
    margin-bottom: 0;
}

/* KPI cards */
.kpi-card {
    background: linear-gradient(145deg, #1b2240, #12182d);
    border: 1px solid #333d65;
    border-radius: 18px;
    padding: 22px 20px;
    min-height: 145px;
    box-shadow: 0 10px 30px rgba(0,0,0,.18);
    transition: all .25s ease;
}

.kpi-card:hover {
    transform: translateY(-5px);
    border-color: #8b79ff;
    box-shadow: 0 12px 35px rgba(104, 86, 255, .15);
}

.kpi-label {
    font-size: 13px;
    color: #aeb8d8;
    margin-bottom: 15px;
    font-weight: 600;
}

.kpi-value {
    color: #ffffff;
    font-size: clamp(24px, 2.5vw, 34px);
    font-weight: 800;
    letter-spacing: -1px;
}

.kpi-note {
    color: #8795bd;
    font-size: 11px;
    margin-top: 9px;
}

/* Section headings */
.section-title {
    font-size: 23px;
    font-weight: 750;
    color: #f6f7ff;
    margin-top: 28px;
    margin-bottom: 5px;
}

.section-subtitle {
    font-size: 13px;
    color: #929fbe;
    margin-bottom: 20px;
}

/* General metric styling */
[data-testid="stMetric"] {
    background: linear-gradient(145deg, #1b2240, #12182d);
    border: 1px solid #333d65;
    padding: 17px;
    border-radius: 15px;
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    gap: 12px;
    border-bottom: 1px solid #303858;
}

.stTabs [data-baseweb="tab"] {
    color: #aab4d3;
    font-weight: 600;
}

.stTabs [aria-selected="true"] {
    color: #b9aaff !important;
}

/* Buttons */
.stButton > button {
    background: linear-gradient(90deg, #725cff, #a05cff);
    color: white;
    border: none;
    border-radius: 10px;
    font-weight: 700;
    padding: 10px 20px;
    transition: .2s;
}

.stButton > button:hover {
    border: 1px solid #bcaaff;
    box-shadow: 0 0 20px rgba(137, 100, 255, .3);
}

/* File uploader */
[data-testid="stFileUploader"] {
    background: rgba(28, 36, 65, .55);
    border: 1px dashed #7065bd;
    border-radius: 15px;
    padding: 10px;
}

/* Dataframe */
[data-testid="stDataFrame"] {
    border: 1px solid #303858;
    border-radius: 12px;
}

/* Select boxes and inputs */
.stSelectbox > div > div,
.stMultiSelect > div > div,
.stDateInput > div > div {
    background-color: #171d33;
    border-color: #343d61;
    border-radius: 10px;
}

/* Footer */
.footer {
    text-align: center;
    color: #7f8aa9;
    font-size: 12px;
    padding: 35px 0 10px;
    letter-spacing: 1px;
}

</style>
""", unsafe_allow_html=True)


# ============================================================
# REUSABLE DATA FUNCTIONS
# ============================================================

REQUIRED_METRICS = ["Sales", "Profit", "Quantity"]

DATE_CANDIDATES = ["Order Date", "Ship Date"]

NUMERIC_CANDIDATES = [
    "Sales",
    "Profit",
    "Quantity",
    "Discount",
    "Price",
    "Cost",
    "Revenue",
    "Unit Price",
    "Total"
]


def read_uploaded_file(uploaded_file):
    """Read CSV, XLSX or JSON into a pandas DataFrame."""

    extension = uploaded_file.name.lower().rsplit(".", 1)[-1]

    if extension == "csv":
        try:
            return pd.read_csv(uploaded_file)
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding="latin-1")

    elif extension == "xlsx":
        return pd.read_excel(uploaded_file, engine="openpyxl")

    elif extension == "json":
        raw = json.load(uploaded_file)

        if isinstance(raw, list):
            return pd.DataFrame(raw)

        if isinstance(raw, dict):
            if "data" in raw and isinstance(raw["data"], list):
                return pd.DataFrame(raw["data"])

            try:
                return pd.DataFrame(raw)
            except ValueError:
                return pd.json_normalize(raw)

        raise ValueError("Invalid JSON structure.")

    else:
        raise ValueError("Only CSV, XLSX and JSON files are supported.")


def normalize_column_names(df):
    """Clean column names."""

    result = df.copy()
    result.columns = [str(col).strip() for col in result.columns]

    return result


def validation_report(df):
    """Generate dataset validation information."""

    return {
        "rows": int(df.shape[0]),
        "columns": int(df.shape[1]),
        "missing_by_column": df.isna().sum().sort_values(
            ascending=False
        ),
        "duplicate_rows": int(df.duplicated().sum()),
        "dtypes": df.dtypes.astype(str)
    }


def preprocess_data(
    source,
    numeric_missing="Median",
    remove_duplicates=True
):
    """
    Clean missing values, duplicate records,
    text inconsistencies and date formats.
    """

    df = normalize_column_names(source)

    before_rows = len(df)

    duplicate_count = int(df.duplicated().sum())

    # Standardize text
    text_cols = df.select_dtypes(
        include=["object", "string"]
    ).columns

    for col in text_cols:

        df[col] = df[col].astype("string").str.strip()

        df[col] = df[col].replace({
            "": pd.NA,
            "nan": pd.NA,
            "None": pd.NA,
            "NULL": pd.NA,
            "null": pd.NA
        })

    # Standardize inconsistent capitalization
    for col in df.select_dtypes(
        include=["object", "string"]
    ).columns:

        if col not in DATE_CANDIDATES:
            df[col] = df[col].apply(
                lambda x: x.title() if isinstance(x, str) else x
            )

    # Remove duplicates
    if remove_duplicates:
        df = df.drop_duplicates().copy()

    # Convert numeric columns
    for col in df.columns:

        if col in NUMERIC_CANDIDATES:

            if df[col].dtype == "object" or str(df[col].dtype) == "string":

                cleaned = (
                    df[col]
                    .astype("string")
                    .str.replace(r"[$,₹]", "", regex=True)
                    .str.replace("%", "", regex=False)
                    .str.strip()
                )

                df[col] = pd.to_numeric(
                    cleaned,
                    errors="coerce"
                )

        elif pd.api.types.is_numeric_dtype(df[col]):
            continue

    # Convert date columns
    for col in df.columns:

        if (
            col in DATE_CANDIDATES
            or "date" in col.lower()
        ):

            df[col] = pd.to_datetime(
                df[col],
                errors="coerce",
                dayfirst=False
            )

    # Handle numeric missing values
    numeric_cols = df.select_dtypes(
        include=np.number
    ).columns

    for col in numeric_cols:

        if df[col].isna().any():

            if numeric_missing == "Median":

                fill_value = df[col].median()

            elif numeric_missing == "Mean":

                fill_value = df[col].mean()

            elif numeric_missing == "Zero":

                fill_value = 0

            else:
                fill_value = df[col].median()

            if pd.notna(fill_value):
                df[col] = df[col].fillna(fill_value)

    # Handle categorical missing values
    categorical_cols = df.select_dtypes(
        include=["object", "string", "category"]
    ).columns

    for col in categorical_cols:

        if df[col].isna().any():

            mode = df[col].mode(dropna=True)

            fill_value = (
                mode.iloc[0]
                if not mode.empty
                else "Unknown"
            )

            df[col] = df[col].fillna(fill_value)

    audit = {
        "rows_before": before_rows,
        "rows_after": len(df),
        "duplicates_found": duplicate_count,
        "duplicates_removed": (
            duplicate_count if remove_duplicates else 0
        ),
        "missing_before": int(source.isna().sum().sum()),
        "missing_after": int(df.isna().sum().sum()),
    }

    return df, audit


def get_sales_summary(df):
    """Generate reusable statistical summary."""

    numeric_df = df.select_dtypes(include=np.number)

    if numeric_df.empty:
        return pd.DataFrame()

    return numeric_df.describe().T


def get_column(df, possible_names):
    """Find a column using case-insensitive matching."""

    lookup = {
        str(col).strip().lower(): col
        for col in df.columns
    }

    for name in possible_names:

        found = lookup.get(name.lower())

        if found is not None:
            return found

    return None


def money_format(value):
    """Format monetary values."""

    if pd.isna(value):
        return "₹0"

    if abs(value) >= 1_000_000:
        return f"₹{value / 1_000_000:.2f}M"

    if abs(value) >= 1_000:
        return f"₹{value / 1_000:.2f}K"

    return f"₹{value:,.2f}"


def create_kpi(label, value, note, icon):
    """Render an attractive KPI card."""

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{icon} &nbsp; {label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def chart_layout(fig, height=400):
    """Apply a consistent dark theme to charts."""

    fig.update_layout(
        template="plotly_dark",
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(
            family="Arial",
            color="#e8eaff"
        ),
        margin=dict(l=15, r=15, t=45, b=15),
        legend=dict(
            bgcolor="rgba(0,0,0,0)"
        ),
        hoverlabel=dict(
            bgcolor="#202947",
            font_color="white"
        )
    )

    return fig


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown(
        """
        <div style="padding:12px 0 22px">
            <div style="
                font-size:23px;
                font-weight:900;
                letter-spacing:3px;
                color:#ffffff;">
                NOVA <span style="color:#a78bfa">/ INSIGHTS</span>
            </div>
            <div style="
                color:#9ba9d2;
                font-size:11px;
                letter-spacing:1.5px;
                margin-top:9px;">
                SALES INTELLIGENCE PLATFORM
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.divider()

    st.markdown("### ◉ Data source")

    uploaded_file = st.file_uploader(
        "Drop your dataset here",
        type=["csv", "xlsx", "json"],
        help="Supported formats: CSV, Excel and JSON"
    )

    st.caption("Supported formats · CSV / XLSX / JSON")

    st.divider()

    st.markdown("### ⚙ Data preparation")

    numeric_missing = st.selectbox(
        "Missing numeric values",
        ["Median", "Mean", "Zero"]
    )

    remove_duplicates = st.toggle(
        "Remove duplicate rows",
        value=True
    )

    st.divider()

    st.markdown("### ✦ About NOVA")

    st.caption(
        "An interactive workspace for exploring, "
        "cleaning and understanding business data."
    )


# ============================================================
# HERO SECTION
# ============================================================

st.markdown(
    textwrap.dedent("""
    <div class="hero">
        <div class="eyebrow">
            ✦ BUSINESS PERFORMANCE / LIVE WORKSPACE
        </div>
        <h1>
            See the story behind<br>
            your sales data.
        </h1>
        <p>
            A smarter workspace to validate data,
            uncover trends and turn raw records
            into confident decisions.
        </p>
    </div>
    """),
    unsafe_allow_html=True
)


# ============================================================
# EMPTY STATE
# ============================================================

if uploaded_file is None:

    st.markdown(
        """
        <div style="
            text-align:center;
            padding:70px 25px;
            background:linear-gradient(145deg,#171e38,#11172a);
            border:1px solid #303858;
            border-radius:20px;
        ">
            <div style="font-size:65px;">📂</div>
            <h2 style="color:white;">Your data journey starts here</h2>
            <p style="color:#aab4d3;">
                Upload a CSV, Excel or JSON dataset using
                the sidebar to unlock your analytics dashboard.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.stop()


# ============================================================
# LOAD DATA
# ============================================================

try:

    raw_df = read_uploaded_file(uploaded_file)

    if raw_df.empty:
        st.error("The uploaded dataset is empty.")
        st.stop()

    raw_df = normalize_column_names(raw_df)

    st.success(
        f"Successfully loaded: {uploaded_file.name}"
    )

except Exception as e:

    st.error(f"Unable to read the uploaded file: {e}")
    st.stop()


# ============================================================
# DATA VALIDATION
# ============================================================

report = validation_report(raw_df)

st.markdown(
    f"""
    <div style="margin:20px 0 15px;">
        <span style="color:#9ba9d2;font-size:12px;">
            CURRENT DATASET &nbsp; /
        </span>
        <b style="color:#ffffff;">
            {uploaded_file.name}
        </b>
        <span style="color:#aab4d3;">
            &nbsp; · {report['rows']} rows
            &nbsp; · {report['columns']} columns
        </span>
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# PREPROCESS DATA
# ============================================================

df, audit = preprocess_data(
    raw_df,
    numeric_missing=numeric_missing,
    remove_duplicates=remove_duplicates
)


# ============================================================
# NAVIGATION TABS
# ============================================================

tab_overview, tab_analytics, tab_quality, tab_explorer = st.tabs(
    [
        "✦ Overview",
        "▥ Analytics",
        "◇ Data Quality",
        "▤ Data Explorer"
    ]
)


# ============================================================
# SIDEBAR FILTERS
# ============================================================

st.sidebar.divider()

st.sidebar.markdown("### ◈ Interactive filters")

filtered_df = df.copy()

# Date filter
date_col = get_column(
    df,
    ["Order Date", "Date", "Transaction Date", "Invoice Date"]
)

if date_col is not None:

    valid_dates = df[date_col].dropna()

    if not valid_dates.empty:

        min_date = valid_dates.min().date()
        max_date = valid_dates.max().date()

        selected_dates = st.sidebar.date_input(
            "Date range",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date
        )

        if isinstance(selected_dates, tuple) and len(selected_dates) == 2:

            start_date, end_date = selected_dates

            filtered_df = filtered_df[
                filtered_df[date_col].dt.date.between(
                    start_date,
                    end_date
                )
            ]

# Region filter
region_col = get_column(
    df,
    ["Region", "State", "Country", "Market"]
)

if region_col is not None:

    region_values = sorted(
        df[region_col].dropna().astype(str).unique().tolist()
    )

    selected_regions = st.sidebar.multiselect(
        "Select region",
        region_values,
        default=region_values
    )

    filtered_df = filtered_df[
        filtered_df[region_col].astype(str).isin(selected_regions)
    ]

# Category filter
category_col = get_column(
    df,
    ["Category", "Product Category", "Sub-Category"]
)

if category_col is not None:

    category_values = sorted(
        df[category_col].dropna().astype(str).unique().tolist()
    )

    selected_categories = st.sidebar.multiselect(
        "Select category",
        category_values,
        default=category_values
    )

    filtered_df = filtered_df[
        filtered_df[category_col].astype(str).isin(selected_categories)
    ]

# Segment filter
segment_col = get_column(
    df,
    ["Segment", "Customer Segment"]
)

if segment_col is not None:

    segment_values = sorted(
        df[segment_col].dropna().astype(str).unique().tolist()
    )

    selected_segments = st.sidebar.multiselect(
        "Select segment",
        segment_values,
        default=segment_values
    )

    filtered_df = filtered_df[
        filtered_df[segment_col].astype(str).isin(selected_segments)
    ]


# ============================================================
# KPI CALCULATIONS
# ============================================================

sales_col = get_column(df, ["Sales", "Revenue", "Total Sales"])
profit_col = get_column(df, ["Profit", "Net Profit"])
quantity_col = get_column(df, ["Quantity", "Qty", "Units Sold"])
order_col = get_column(df, ["Order ID", "Order Id", "Order Number"])

total_sales = (
    filtered_df[sales_col].sum()
    if sales_col is not None else 0
)

total_profit = (
    filtered_df[profit_col].sum()
    if profit_col is not None else 0
)

total_quantity = (
    filtered_df[quantity_col].sum()
    if quantity_col is not None else 0
)

total_orders = (
    filtered_df[order_col].nunique()
    if order_col is not None
    else len(filtered_df)
)

profit_margin = (
    (total_profit / total_sales) * 100
    if total_sales != 0 else 0
)


# ============================================================
# OVERVIEW TAB
# ============================================================

with tab_overview:

    st.markdown(
        '<div class="section-title">Dataset snapshot</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-subtitle">A quick read of the file you uploaded.</div>',
        unsafe_allow_html=True
    )

    # Raw dataset KPIs
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        create_kpi(
            "Records",
            f"{report['rows']:,}",
            "Total uploaded rows",
            "▤"
        )

    with c2:
        create_kpi(
            "Attributes",
            f"{report['columns']:,}",
            "Total dataset columns",
            "▦"
        )

    with c3:
        create_kpi(
            "Missing cells",
            f"{int(raw_df.isna().sum().sum()):,}",
            "Before preprocessing",
            "◈"
        )

    with c4:
        create_kpi(
            "Duplicate rows",
            f"{report['duplicate_rows']:,}",
            "Before preprocessing",
            "⧉"
        )

    st.markdown(
        '<div class="section-title">Business performance</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-subtitle">Live KPIs based on your selected filters.</div>',
        unsafe_allow_html=True
    )

    # Five dynamically updating KPIs
    k1, k2, k3, k4, k5 = st.columns(5)

    with k1:
        create_kpi(
            "Total Sales",
            money_format(total_sales),
            "Filtered revenue",
            "💰"
        )

    with k2:
        create_kpi(
            "Total Profit",
            money_format(total_profit),
            "Filtered profit",
            "📈"
        )

    with k3:
        create_kpi(
            "Total Orders",
            f"{total_orders:,}",
            "Unique order IDs",
            "🛒"
        )

    with k4:
        create_kpi(
            "Quantity Sold",
            f"{total_quantity:,.0f}",
            "Total units",
            "📦"
        )

    with k5:
        create_kpi(
            "Profit Margin",
            f"{profit_margin:.2f}%",
            "Profit / Sales × 100",
            "🎯"
        )

    # Quick visualizations
    st.markdown(
        '<div class="section-title">Performance at a glance</div>',
        unsafe_allow_html=True
    )

    chart1, chart2 = st.columns(2)

    if sales_col is not None and category_col is not None:

        category_sales = (
            filtered_df.groupby(category_col)[sales_col]
            .sum()
            .reset_index()
            .sort_values(sales_col, ascending=False)
        )

        with chart1:

            fig = px.bar(
                category_sales,
                x=category_col,
                y=sales_col,
                color=sales_col,
                title="Sales by category",
                color_continuous_scale="Purples",
                text_auto=".2s"
            )

            st.plotly_chart(
                chart_layout(fig),
                use_container_width=True
            )

    if profit_col is not None and region_col is not None:

        region_profit = (
            filtered_df.groupby(region_col)[profit_col]
            .sum()
            .reset_index()
        )

        with chart2:

            fig = px.pie(
                region_profit,
                names=region_col,
                values=profit_col,
                title="Profit distribution by region",
                hole=.55,
                color_discrete_sequence=px.colors.qualitative.Vivid
            )

            st.plotly_chart(
                chart_layout(fig),
                use_container_width=True
            )


# ============================================================
# ANALYTICS TAB
# ============================================================

with tab_analytics:

    st.markdown(
        '<div class="section-title">Sales intelligence</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-subtitle">Explore trends, relationships and business performance.</div>',
        unsafe_allow_html=True
    )

    # Chart 1: Sales trend
    if date_col is not None and sales_col is not None:

        trend_data = filtered_df.dropna(
            subset=[date_col]
        ).copy()

        if not trend_data.empty:

            trend_data["Month"] = (
                trend_data[date_col].dt.to_period("M").astype(str)
            )

            monthly_sales = (
                trend_data.groupby("Month")[sales_col]
                .sum()
                .reset_index()
            )

            fig = px.line(
                monthly_sales,
                x="Month",
                y=sales_col,
                markers=True,
                title="Monthly sales trend",
                color_discrete_sequence=["#9d8aff"]
            )

            fig.update_traces(
                line_width=3,
                marker_size=8
            )

            st.plotly_chart(
                chart_layout(fig, 420),
                use_container_width=True
            )

    col1, col2 = st.columns(2)

    # Chart 2: Category analysis
    if category_col is not None and sales_col is not None:

        category_analysis = (
            filtered_df.groupby(category_col)[sales_col]
            .sum()
            .reset_index()
            .sort_values(sales_col, ascending=False)
        )

        with col1:

            fig = px.bar(
                category_analysis,
                x=category_col,
                y=sales_col,
                color=category_col,
                title="Revenue by category",
                text_auto=".2s"
            )

            st.plotly_chart(
                chart_layout(fig),
                use_container_width=True
            )

    # Chart 3: Profit analysis
    if region_col is not None and profit_col is not None:

        profit_analysis = (
            filtered_df.groupby(region_col)[profit_col]
            .sum()
            .reset_index()
        )

        with col2:

            fig = px.pie(
                profit_analysis,
                names=region_col,
                values=profit_col,
                hole=.5,
                title="Profit contribution",
                color_discrete_sequence=px.colors.qualitative.Bold
            )

            st.plotly_chart(
                chart_layout(fig),
                use_container_width=True
            )

    # Chart 4: Sales vs Profit
    if sales_col is not None and profit_col is not None:

        scatter_cols = [sales_col, profit_col]

        if category_col is not None:
            scatter_cols.append(category_col)

        scatter_data = filtered_df[scatter_cols].dropna()

        if not scatter_data.empty:

            fig = px.scatter(
                scatter_data,
                x=sales_col,
                y=profit_col,
                color=category_col if category_col else None,
                title="Sales vs Profit relationship",
                opacity=.75
            )

            st.plotly_chart(
                chart_layout(fig, 430),
                use_container_width=True
            )

    # Chart 5: Quantity analysis
    if category_col is not None and quantity_col is not None:

        quantity_analysis = (
            filtered_df.groupby(category_col)[quantity_col]
            .sum()
            .reset_index()
        )

        fig = px.bar(
            quantity_analysis,
            x=category_col,
            y=quantity_col,
            title="Quantity sold by category",
            color=quantity_col,
            color_continuous_scale="Teal"
        )

        st.plotly_chart(
            chart_layout(fig),
            use_container_width=True
        )

    # Fallback message
    if (
        date_col is None
        and category_col is None
        and region_col is None
    ):

        st.info(
            "Your dataset does not contain standard date, category "
            "or region columns. Check Data Explorer for available fields."
        )


# ============================================================
# DATA QUALITY TAB
# ============================================================

with tab_quality:

    st.markdown(
        '<div class="section-title">Data quality centre</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-subtitle">Understand and monitor your dataset health.</div>',
        unsafe_allow_html=True
    )

    q1, q2, q3 = st.columns(3)

    with q1:
        st.metric(
            "Rows before cleaning",
            f"{audit['rows_before']:,}"
        )

    with q2:
        st.metric(
            "Rows after cleaning",
            f"{audit['rows_after']:,}"
        )

    with q3:
        st.metric(
            "Duplicates removed",
            f"{audit['duplicates_removed']:,}"
        )

    st.markdown("### Missing values by column")

    missing_data = pd.DataFrame({
        "Column": raw_df.columns,
        "Missing values": raw_df.isna().sum().values,
        "Missing percentage": (
            raw_df.isna().mean().values * 100
        ).round(2)
    })

    missing_data = missing_data.sort_values(
        "Missing values",
        ascending=False
    )

    fig = px.bar(
        missing_data,
        x="Column",
        y="Missing values",
        color="Missing percentage",
        title="Missing values before cleaning",
        color_continuous_scale="Reds",
        hover_data=["Missing percentage"]
    )

    st.plotly_chart(
        chart_layout(fig),
        use_container_width=True
    )

    st.markdown("### Column data types")

    dtype_df = pd.DataFrame({
        "Column": raw_df.columns,
        "Data Type": raw_df.dtypes.astype(str).values,
        "Non-null values": raw_df.notna().sum().values,
        "Missing values": raw_df.isna().sum().values
    })

    st.dataframe(
        dtype_df,
        use_container_width=True,
        hide_index=True
    )

    st.markdown("### Preprocessing audit")

    audit_df = pd.DataFrame({
        "Metric": [
            "Rows before cleaning",
            "Rows after cleaning",
            "Duplicate rows found",
            "Duplicate rows removed",
            "Missing values before",
            "Missing values after"
        ],
        "Value": [
            audit["rows_before"],
            audit["rows_after"],
            audit["duplicates_found"],
            audit["duplicates_removed"],
            audit["missing_before"],
            audit["missing_after"]
        ]
    })

    st.dataframe(
        audit_df,
        use_container_width=True,
        hide_index=True
    )

    st.markdown("### Statistical summary")

    summary = get_sales_summary(filtered_df)

    if not summary.empty:

        st.dataframe(
            summary.round(2),
            use_container_width=True
        )

    else:
        st.info("No numeric columns available for statistical analysis.")


# ============================================================
# DATA EXPLORER TAB
# ============================================================

with tab_explorer:

    st.markdown(
        '<div class="section-title">Explore your records</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-subtitle">Inspect cleaned data and download your results.</div>',
        unsafe_allow_html=True
    )

    search_text = st.text_input(
        "⌕ Search records",
        placeholder="Search for a customer, product, region..."
    )

    display_df = filtered_df.copy()

    if search_text:

        mask = display_df.astype(str).apply(
            lambda col: col.str.contains(
                search_text,
                case=False,
                na=False
            )
        ).any(axis=1)

        display_df = display_df[mask]

    st.caption(
        f"Showing {len(display_df):,} records "
        f"out of {len(filtered_df):,} filtered records"
    )

    st.dataframe(
        display_df,
        use_container_width=True,
        height=450,
        hide_index=True
    )

    st.markdown("### Download your cleaned dataset")

    csv_data = display_df.to_csv(index=False).encode("utf-8")

    st.download_button(
        label="⬇ Download cleaned CSV",
        data=csv_data,
        file_name="nova_cleaned_sales_data.csv",
        mime="text/csv"
    )

    excel_buffer = io.BytesIO()

    with pd.ExcelWriter(
        excel_buffer,
        engine="openpyxl"
    ) as writer:

        display_df.to_excel(
            writer,
            index=False,
            sheet_name="Cleaned Data"
        )

        if not summary.empty:
            summary.to_excel(
                writer,
                sheet_name="Statistics"
            )

    st.download_button(
        label="⬇ Download Excel report",
        data=excel_buffer.getvalue(),
        file_name="nova_sales_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    """
    <div class="footer">
        NOVA / SALES INTELLIGENCE PLATFORM
        <br><br>
        BUILT WITH PYTHON · STREAMLIT · PANDAS · PLOTLY
        <br>
        TURNING RAW DATA INTO SMART DECISIONS ✦
    </div>
    """,
    unsafe_allow_html=True
)
