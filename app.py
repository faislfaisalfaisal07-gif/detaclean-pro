"""
DataClean Pro
Professional AI-powered data cleaning application.
Features:
- CSV and Excel file upload
- Sample data loader for testing
- AI-powered data quality analysis using Groq
- Duplicate row detection/removal
- Missing value handling
- Date format normalization
- Numerical outlier detection
- Text formatting normalization
- Detailed cleaning report
- Before/After comparison
- Savings calculator
- Cleaned CSV download
- Professional SaaS-style Streamlit interface
Requirements:
    pip install streamlit pandas openpyxl groq
Run:
    streamlit run app.py
Environment variable:
    GROQ_API_KEY= 
"""
import os
import io
import re
import json
import datetime
import pandas as pd
import streamlit as st
from groq import Groq

# ============================================================
# PAGE CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="DataClean Pro",
    page_icon="🧹",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# CUSTOM CSS
# ============================================================
st.markdown(
    """
    <style>
    .stApp { background-color: #f8fafc; }
    .block-container { max-width: 1400px; padding-top: 2rem; padding-bottom: 3rem; }
    .main-header { padding: 1rem 0 2rem 0; }
    .brand { font-size: 2.2rem; font-weight: 800; color: #111827; letter-spacing: -1px; }
    .brand span { color: #2563eb; }
    .subtitle { color: #64748b; font-size: 1rem; margin-top: 0.3rem; }
    .metric-card {
        background: white; border: 1px solid #e5e7eb; border-radius: 14px;
        padding: 1.2rem; box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04);
    }
    .metric-title { color: #64748b; font-size: 0.85rem; font-weight: 600; }
    .metric-value { color: #111827; font-size: 1.8rem; font-weight: 750; margin-top: 0.2rem; }
    .section-title {
        font-size: 1.35rem; font-weight: 750; color: #111827;
        margin-top: 1.5rem; margin-bottom: 0.8rem;
    }
    .success-badge {
        display: inline-block; padding: 0.35rem 0.75rem; border-radius: 999px;
        background: #dcfce7; color: #166534; font-size: 0.8rem; font-weight: 700;
    }
    .warning-badge {
        display: inline-block; padding: 0.35rem 0.75rem; border-radius: 999px;
        background: #fef3c7; color: #92400e; font-size: 0.8rem; font-weight: 700;
    }
    .info-box {
        background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 12px;
        padding: 1rem; color: #1e40af;
    }
    .savings-box {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        border-radius: 14px; padding: 1.5rem; color: white;
        box-shadow: 0 4px 14px rgba(16, 185, 129, 0.3);
    }
    .savings-title { font-size: 0.9rem; font-weight: 600; opacity: 0.9; }
    .savings-value { font-size: 2.2rem; font-weight: 800; margin-top: 0.3rem; }
    .savings-subtitle { font-size: 0.8rem; opacity: 0.85; margin-top: 0.3rem; }
    section[data-testid="stSidebar"] {
        background-color: #ffffff; border-right: 1px solid #e5e7eb;
    }
    .stButton > button { border-radius: 9px; font-weight: 650; }
    [data-testid="stDataFrame"] { border-radius: 12px; overflow: hidden; }
    </style>
    """,
    unsafe_allow_html=True
)

# ============================================================
# SESSION STATE
# ============================================================
if "cleaned_df" not in st.session_state:
    st.session_state.cleaned_df = None
if "original_df" not in st.session_state:
    st.session_state.original_df = None
if "cleaning_report" not in st.session_state:
    st.session_state.cleaning_report = []
    if "ai_analysis" not in st.session_state:
            st.session_state.ai_analysis = None

# ============================================================
# TRACKING
# ============================================================
try:
    ref = st.query_params.get("ref", "direct")
except:
    ref = "direct"

st.session_state.referrer = ref
    
# SAMPLE DATA
# ============================================================
def get_sample_data():
    """
    Return a sample dirty dataset for testing purposes.
    """
    sample_data = {
        "name": [
            "Ali Rezaei", "Ali Rezaei", "Sara Ahmadi",
            "", "Mohammad Karimi", "Fatima Noori", "asdfgh",
            "Hassan Ahmadi", "Zahra Karimi", "Reza Mohammadi"
        ],
        "age": [25, 25, 300, 30, -5, 28, 999, 35, 42, 19],
        "email": [
            "ali@gmail.com", "ali@gmail.com", "sara@yahoo",
            "", "mohammad@", "", "test@test",
            "hassan@gmail.com", "zahra@yahoo.com", "reza@"
        ],
        "signup_date": [
            "2024/01/05", "2024/01/05", "05-01-2024",
            "January 5 2024", "2024-01-05", "", "invalid_date",
            "2024/02/15", "03-20-2024", "2024-04-10"
        ],
        "city": [
            "Tehran", "Tehran", "tehran",
            "", "Kabul", "kabul", "",
            "  Tehran  ", "HERAT", "Kabul"
        ]
    }
    return pd.DataFrame(sample_data)

# ============================================================
# HELPER FUNCTIONS
# ============================================================
def get_groq_client():
    """Create a Groq client using the GROQ_API_KEY environment variable."""
    api_key = os.getenv("GROQ_API_KEY")
    if not api_key:
        return None
    return Groq(api_key=api_key)

def load_file(uploaded_file):
    """Load CSV or Excel files into a pandas DataFrame."""
    file_name = uploaded_file.name.lower()
    if file_name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
        return pd.read_excel(uploaded_file)
    raise ValueError("Unsupported file format. Please upload a CSV or Excel file.")

def generate_data_profile(df):
    """Generate a compact profile of the dataset for AI analysis."""
    profile = {"rows": len(df), "columns": len(df.columns), "columns_info": []}
    for column in df.columns:
        series = df[column]
        profile["columns_info"].append({
            "name": str(column),
            "dtype": str(series.dtype),
            "missing": int(series.isna().sum()),
            "unique": int(series.nunique(dropna=True)),
            "sample": series.dropna().astype(str).head(5).tolist()
        })
    return profile

def ask_groq_for_analysis(df):
    """Send a compact dataset profile to Groq and request data-quality recommendations."""
    client = get_groq_client()
    if client is None:
        return {
            "status": "unavailable",
            "message": "Groq API key was not found. Set GROQ_API_KEY to enable AI analysis."
        }
    profile = generate_data_profile(df)
    prompt = f"""
You are a senior data quality analyst.
Analyze this dataset profile and identify potential data-quality issues.
Dataset profile:
{json.dumps(profile, indent=2)}
Look specifically for:
1. Missing values
2. Duplicate records
3. Incorrect data types
4. Potential date columns
5. Suspicious numerical values or outliers
6. Inconsistent text formatting
7. Columns that may need normalization
Return ONLY valid JSON using this structure:
{{
    "overall_quality": "Excellent|Good|Fair|Poor",
    "summary": "Short professional summary",
    "issues": [
        {{"column": "column name", "issue": "description", "severity": "Low|Medium|High"}}
    ],
    "recommendations": ["recommendation 1", "recommendation 2"]
}}
Do not invent columns that do not exist.
"""
    try:
        response = client.chat.completions.create(
            model="openai/gpt-oss-20b",
            messages=[
                {"role": "system", "content": "You are an expert data quality analyst. Return concise, accurate JSON."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
            max_tokens=1500
        )
        content = response.choices[0].message.content.strip()
        content = re.sub(r"^```json\s*", "", content)
        content = re.sub(r"\s*```$", "", content)
        return {"status": "success", "data": json.loads(content)}
    except json.JSONDecodeError:
        return {"status": "error", "message": "Groq returned an invalid JSON response."}
    except Exception as e:
        return {"status": "error", "message": f"AI analysis failed: {str(e)}"}

def clean_duplicates(df, report):
    """Remove completely duplicated rows."""
    duplicate_count = int(df.duplicated().sum())
    if duplicate_count > 0:
        df = df.drop_duplicates().reset_index(drop=True)
        report.append({
            "category": "Duplicates",
            "column": "All columns",
            "count": duplicate_count,
            "action": f"Removed {duplicate_count} duplicate rows."
        })
    return df

def clean_missing_values(df, report):
    """Fill missing values using sensible column-specific rules."""
    for column in df.columns:
        missing_count = int(df[column].isna().sum())
        if missing_count == 0:
            continue
        series = df[column]
        if pd.api.types.is_numeric_dtype(series):
            median_value = series.median()
            if pd.notna(median_value):
                df[column] = series.fillna(median_value)
                report.append({
                    "category": "Missing Values",
                    "column": str(column),
                    "count": missing_count,
                    "action": f"Filled missing numeric values with median ({median_value})."
                })
        elif pd.api.types.is_datetime64_any_dtype(series):
            mode = series.mode()
            if not mode.empty:
                df[column] = series.fillna(mode.iloc[0])
                report.append({
                    "category": "Missing Values",
                    "column": str(column),
                    "count": missing_count,
                    "action": "Filled missing dates with the most common date."
                })
        else:
            mode = series.mode()
            if not mode.empty:
                replacement = mode.iloc[0]
            else:
                replacement = "Unknown"
            df[column] = series.fillna(replacement)
            report.append({
                "category": "Missing Values",
                "column": str(column),
                "count": missing_count,
                "action": f'Filled missing text values with "{replacement}".'
            })
    return df

def detect_and_fix_dates(df, report):
    """Detect columns that appear to contain dates and normalize them."""
    for column in df.columns:
        series = df[column]
        if pd.api.types.is_datetime64_any_dtype(series):
            continue
        if not (pd.api.types.is_object_dtype(series) or pd.api.types.is_string_dtype(series)):
            continue
        non_null = series.dropna()
        if len(non_null) == 0:
            continue
        parsed = pd.to_datetime(series, errors="coerce")
        valid_ratio = parsed.notna().sum() / len(non_null)
        if valid_ratio >= 0.80:
            before_invalid = int(series.notna().sum() - parsed.notna().sum())
            df[column] = parsed
            report.append({
                "category": "Date Formatting",
                "column": str(column),
                "count": int(parsed.notna().sum()),
                "action": "Converted date values to a standardized datetime format."
            })
            if before_invalid > 0:
                report.append({
                    "category": "Date Formatting",
                    "column": str(column),
                    "count": before_invalid,
                    "action": "Some invalid date values could not be parsed and were converted to missing values."
                })
    return df

def clean_text_formatting(df, report):
    """Normalize text formatting: trim, collapse spaces, standardize empties."""
    for column in df.columns:
        if not (pd.api.types.is_object_dtype(df[column]) or pd.api.types.is_string_dtype(df[column])):
            continue
        original = df[column].copy()
        cleaned = (
            df[column].astype("string").str.strip().str.replace(r"\s+", " ", regex=True)
        )
        cleaned = cleaned.replace("", pd.NA)
        changes = int(
            (original.fillna("__NULL__").astype(str) != cleaned.fillna("__NULL__").astype(str)).sum()
        )
        if changes > 0:
            df[column] = cleaned
            report.append({
                "category": "Text Formatting",
                "column": str(column),
                "count": changes,
                "action": "Trimmed whitespace and normalized repeated spaces."
            })
    return df

def detect_and_fix_outliers(df, report):
    """Detect numerical outliers using the IQR method and cap them."""
    for column in df.select_dtypes(include="number").columns:
        series = df[column]
        if series.dropna().empty:
            continue
        q1 = series.quantile(0.25)
        q3 = series.quantile(0.75)
        iqr = q3 - q1
        if pd.isna(iqr) or iqr == 0:
            continue
        lower_bound = q1 - 1.5 * iqr
        upper_bound = q3 + 1.5 * iqr
        outlier_mask = (series < lower_bound) | (series > upper_bound)
        outlier_count = int(outlier_mask.sum())
        if outlier_count == 0:
            continue
        column_lower = str(column).lower()
        if "age" in column_lower or column_lower == "years":
            logical_lower = 0
            logical_upper = 120
            final_lower = max(lower_bound, logical_lower)
            final_upper = min(upper_bound, logical_upper)
            df[column] = df[column].clip(lower=final_lower, upper=final_upper)
            report.append({
                "category": "Outliers",
                "column": str(column),
                "count": outlier_count,
                "action": f"Detected {outlier_count} suspicious age values and capped them to a logical range ({final_lower:.1f}–{final_upper:.1f})."
            })
        else:
            df[column] = df[column].clip(lower=lower_bound, upper=upper_bound)
            report.append({
                "category": "Outliers",
                "column": str(column),
                "count": outlier_count,
                "action": f"Capped {outlier_count} IQR outliers to the statistical boundaries ({lower_bound:.2f}–{upper_bound:.2f})."
            })
    return df

def run_cleaning_pipeline(df, progress_callback=None):
    """Execute all cleaning operations in sequence."""
    report = []
    total_steps = 5
    if progress_callback:
        progress_callback(1 / total_steps, "Removing duplicate rows...")
    df = clean_duplicates(df, report)
    if progress_callback:
        progress_callback(2 / total_steps, "Handling missing values...")
    df = clean_missing_values(df, report)
    if progress_callback:
        progress_callback(3 / total_steps, "Standardizing date formats...")
    df = detect_and_fix_dates(df, report)
    if progress_callback:
        progress_callback(4 / total_steps, "Normalizing text formatting...")
    df = clean_text_formatting(df, report)
    if progress_callback:
        progress_callback(5 / total_steps, "Detecting numerical outliers...")
    df = detect_and_fix_outliers(df, report)
    if progress_callback:
        progress_callback(1.0, "Cleaning complete.")
    return df, report

def dataframe_to_csv(df):
    """Convert DataFrame to downloadable CSV bytes."""
    return df.to_csv(index=False).encode("utf-8")

def calculate_savings(original_df, report):
    """
    Calculate estimated time and money saved by cleaning.
    """
    total_changes = sum(item["count"] for item in report)
    rows_affected = len(original_df)
    
    # Estimate: 30 seconds saved per issue on average
    # For a human analyst: ~$40/hour
    seconds_saved = total_changes * 30 + rows_affected * 0.5
    hours_saved = seconds_saved / 3600
    money_saved = hours_saved * 40
    
    return {
        "total_changes": total_changes,
        "rows_affected": rows_affected,
        "hours_saved": hours_saved,
        "money_saved": money_saved
    }

# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.markdown(
        """
        <div style="padding: 0.5rem 0 1.5rem 0;">
            <div style="font-size:1.5rem;font-weight:800;">
                🧹 DataClean <span style="color:#2563eb;">Pro</span>
            </div>
            <div style="color:#64748b;font-size:0.85rem;">
                AI-powered data quality platform
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
    st.markdown("### Processing")
    st.markdown(
        """
        **Pipeline**
        1. Upload dataset (or load sample)
        2. AI quality analysis
        3. Detect data issues
        4. Automatically clean
        5. Review report & savings
        6. Download clean data
        """
    )
    st.divider()
    st.markdown("### AI Configuration")
    api_status = os.getenv("GROQ_API_KEY")
    if api_status:
        st.success("Groq API connected")
    else:
        st.warning("Groq API key not configured")
        st.caption("AI model: openai/gpt-oss-20b")
        st.write(f"📍 Ref: {st.session_state.referrer}")
    st.divider()
    st.caption("DataClean Pro")
    st.caption("Professional data cleaning for modern businesses.")

# ============================================================
# MAIN HEADER
# ============================================================
st.markdown(
    """
    <div class="main-header">
        <div class="brand">DataClean <span>Pro</span></div>
        <div class="subtitle">
            AI-powered data cleaning and quality analysis for business data.
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

# ============================================================
# DATA INPUT SECTION
# ============================================================
st.markdown('<div class="section-title">Get Started</div>', unsafe_allow_html=True)

input_col1, input_col2 = st.columns(2)

with input_col1:
    st.markdown("#### 📁 Upload Your File")
    uploaded_file = st.file_uploader(
        "Choose a CSV or Excel file",
        type=["csv", "xlsx", "xls"],
        help="Supported formats: CSV, XLSX and XLS"
    )

with input_col2:
    st.markdown("#### 📊 Try Sample Data")
    st.caption("No file? Load a sample dirty dataset to test the pipeline.")
    load_sample = st.button(
        "Load Sample Data",
        type="secondary",
        use_container_width=True
    )

# ============================================================
# HANDLE DATA LOADING
# ============================================================
df = None

if load_sample:
    df = get_sample_data()
    st.session_state.original_df = df.copy()
    st.session_state.cleaned_df = None
    st.session_state.cleaning_report = []
    st.session_state.ai_analysis = None
    st.success("✅ Sample data loaded! Scroll down to analyze and clean it.")

if uploaded_file is not None:
    try:
        with st.spinner("Loading dataset..."):
            df = load_file(uploaded_file)
        if df.empty:
            st.error("The uploaded file contains no data.")
            st.stop()
        st.session_state.original_df = df.copy()
        st.session_state.cleaned_df = None
        st.session_state.cleaning_report = []
        st.session_state.ai_analysis = None
    except Exception as e:
        st.error(f"Unable to read the uploaded file: {str(e)}")
        st.stop()

# Use stored dataframe if available
if df is None and st.session_state.original_df is not None:
    df = st.session_state.original_df

if df is None:
    st.markdown(
        """
        <div class="info-box">
            <strong>Get started</strong><br>
            Upload a CSV/Excel file or click <b>Load Sample Data</b> to test the pipeline.
        </div>
        """,
        unsafe_allow_html=True
    )
    st.stop()

# ============================================================
# DATASET OVERVIEW
# ============================================================
st.markdown('<div class="section-title">Dataset Overview</div>', unsafe_allow_html=True)

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-title">ROWS</div>
            <div class="metric-value">{len(df):,}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
with col2:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-title">COLUMNS</div>
            <div class="metric-value">{len(df.columns):,}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
with col3:
    missing_total = int(df.isna().sum().sum())
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-title">MISSING VALUES</div>
            <div class="metric-value">{missing_total:,}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
with col4:
    duplicates = int(df.duplicated().sum())
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-title">DUPLICATES</div>
            <div class="metric-value">{duplicates:,}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# ============================================================
# DATA PREVIEW
# ============================================================
st.markdown('<div class="section-title">Data Preview</div>', unsafe_allow_html=True)
st.dataframe(df.head(100), use_container_width=True, height=400)

# ============================================================
# AI ANALYSIS
# ============================================================
st.markdown('<div class="section-title">AI Data Quality Analysis</div>', unsafe_allow_html=True)

analyze_col1, analyze_col2 = st.columns([1, 3])
with analyze_col1:
    analyze_clicked = st.button("Run AI Analysis", type="primary", use_container_width=True)
with analyze_col2:
    st.caption("DataClean Pro uses Groq's  openai/gpt-oss-20b,model to identify potential quality issues.")

if analyze_clicked:
    with st.spinner("AI is analyzing your dataset..."):
        result = ask_groq_for_analysis(df)
    st.session_state.ai_analysis = result

if st.session_state.ai_analysis:
    result = st.session_state.ai_analysis
    if result["status"] == "success":
        analysis = result["data"]
        quality = analysis.get("overall_quality", "Unknown")
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-title">AI QUALITY ASSESSMENT</div>
                <div style="font-size:1.5rem;font-weight:750;margin-top:0.4rem;">{quality}</div>
                <div style="color:#64748b;margin-top:0.4rem;">{analysis.get("summary", "")}</div>
            </div>
            """,
            unsafe_allow_html=True
        )
        issues = analysis.get("issues", [])
        if issues:
            st.markdown("#### Detected Issues")
            issue_df = pd.DataFrame(issues)
            st.dataframe(issue_df, use_container_width=True, hide_index=True)
        recommendations = analysis.get("recommendations", [])
        if recommendations:
            st.markdown("#### AI Recommendations")
            for recommendation in recommendations:
                st.markdown(f"- {recommendation}")
    elif result["status"] == "unavailable":
        st.info(result["message"])
    else:
        st.error(result["message"])

# ============================================================
# CLEAN DATA
# ============================================================
st.markdown('<div class="section-title">Automatic Data Cleaning</div>', unsafe_allow_html=True)
st.markdown("DataClean Pro will automatically process your dataset using a deterministic pandas-based cleaning pipeline.")

clean_clicked = st.button("Clean Dataset", type="primary", use_container_width=True)

if clean_clicked:
    progress_bar = st.progress(0)
    progress_text = st.empty()
    def update_progress(value, message):
        progress_bar.progress(min(max(value, 0.0), 1.0))
        progress_text.caption(message)
    try:
        cleaned_df, report = run_cleaning_pipeline(df.copy(), progress_callback=update_progress)
        st.session_state.cleaned_df = cleaned_df
        st.session_state.cleaning_report = report
        progress_bar.progress(1.0)
        progress_text.success("Dataset cleaned successfully.")
    except Exception as e:
        progress_bar.empty()
        progress_text.empty()
        st.error(f"Cleaning failed: {str(e)}")

# ============================================================
# SAVINGS CALCULATOR
# ============================================================
if st.session_state.cleaned_df is not None and st.session_state.cleaning_report:
    st.markdown('<div class="section-title">💰 Estimated Savings</div>', unsafe_allow_html=True)
    
    savings = calculate_savings(
        st.session_state.original_df,
        st.session_state.cleaning_report
    )
    
    sav_col1, sav_col2, sav_col3 = st.columns(3)
    with sav_col1:
        st.markdown(
            f"""
            <div class="savings-box">
                <div class="savings-title">TIME SAVED</div>
                <div class="savings-value">{savings['hours_saved']:.1f}h</div>
                <div class="savings-subtitle">vs. manual cleaning</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    with sav_col2:
        st.markdown(
            f"""
            <div class="savings-box">
                <div class="savings-title">MONEY SAVED</div>
                <div class="savings-value">${savings['money_saved']:.0f}</div>
                <div class="savings-subtitle">at $40/hour analyst rate</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    with sav_col3:
        st.markdown(
            f"""
            <div class="savings-box">
                <div class="savings-title">ISSUES FIXED</div>
                <div class="savings-value">{savings['total_changes']:,}</div>
                <div class="savings-subtitle">across {savings['rows_affected']:,} rows</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    
    st.caption("💡 Estimated savings based on industry-standard analyst rates ($40/hour) and average manual cleaning time.")

# ============================================================
# CLEANING REPORT
# ============================================================
if st.session_state.cleaned_df is not None:
    cleaned_df = st.session_state.cleaned_df
    report = st.session_state.cleaning_report
    
    st.markdown('<div class="section-title">Cleaning Report</div>', unsafe_allow_html=True)
    
    total_changes = sum(item["count"] for item in report)
    
    report_col1, report_col2, report_col3 = st.columns(3)
    with report_col1:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-title">CLEANING ACTIONS</div>
                <div class="metric-value">{len(report):,}</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    with report_col2:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-title">VALUES / ROWS AFFECTED</div>
                <div class="metric-value">{total_changes:,}</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    with report_col3:
        remaining_missing = int(cleaned_df.isna().sum().sum())
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-title">REMAINING MISSING VALUES</div>
                <div class="metric-value">{remaining_missing:,}</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    
    if report:
        report_df = pd.DataFrame(report)
        st.dataframe(report_df, use_container_width=True, hide_index=True)
    else:
        st.success("No significant data-quality issues were detected.")

# ============================================================
# BEFORE / AFTER COMPARISON
# ============================================================
if st.session_state.cleaned_df is not None:
    st.markdown('<div class="section-title">🔄 Before & After Comparison</div>', unsafe_allow_html=True)
    
    original = st.session_state.original_df
    cleaned = st.session_state.cleaned_df
    
    comp_col1, comp_col2, comp_col3 = st.columns(3)
    with comp_col1:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-title">ROWS BEFORE</div>
                <div class="metric-value">{len(original):,}</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    with comp_col2:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-title">ROWS AFTER</div>
                <div class="metric-value">{len(cleaned):,}</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    with comp_col3:
        removed = len(original) - len(cleaned)
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-title">DUPLICATES REMOVED</div>
                <div class="metric-value">{removed:,}</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    
    before_col, after_col = st.columns(2)
    with before_col:
        st.markdown("#### ❌ Original (Dirty)")
        st.dataframe(original.head(50), use_container_width=True, height=350)
    with after_col:
        st.markdown("#### ✅ Cleaned")
        st.dataframe(cleaned.head(50), use_container_width=True, height=350)

# ============================================================
# DOWNLOAD
# ============================================================
if st.session_state.cleaned_df is not None:
    st.markdown('<div class="section-title">Export Clean Data</div>', unsafe_allow_html=True)
    
    cleaned_csv = dataframe_to_csv(st.session_state.cleaned_df)
    
    st.download_button(
        label="⬇️ Download Cleaned CSV",
        data=cleaned_csv,
        file_name="dataclean_pro_cleaned.csv",
        mime="text/csv",
        type="primary",
        use_container_width=True
    )
    st.caption("Your cleaned dataset is exported as a UTF-8 CSV file.")

# ============================================================
# FOOTER
# ============================================================
st.divider()
st.markdown(
    """
    <div style="text-align:center;color:#94a3b8;font-size:0.85rem;">
        DataClean Pro · AI-powered data quality for modern businesses
    </div>
    """,
    unsafe_allow_html=True
)
