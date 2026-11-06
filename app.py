import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import numpy as np
import math
import re
import docx
from docx.shared import Inches, Pt, Mm
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime

# =========================
# PAGE CONFIG
# =========================
st.set_page_config(
    page_title="SPC Color Dashboard",
    page_icon="📊",
    layout="wide"
)

st.markdown(
    """
    <style>
    .stApp {
        background: linear-gradient(270deg, #ffffff, #f0f9ff, #e0f2fe, #fef3c7, #ecfeff);
        background-size: 800% 800%;
        animation: gradientBG 20s ease infinite;
    }
    @keyframes gradientBG {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }
    </style>
    """,
    unsafe_allow_html=True
)

if st.button("🔄 Refresh data"):
    st.cache_data.clear()
    st.rerun()

st.markdown("""<style>[data-testid="stSidebar"] {background-color: #f6f8fa;}</style>""", unsafe_allow_html=True)

# =========================
# FILE UPLOAD DATA SOURCE
# =========================
st.sidebar.markdown("### 📂 Data Upload")
# Diagnostics show session memory use to distinguish large data from Python exceptions.
with st.sidebar.expander("🩺 App diagnostics", expanded=False):
    try:
        import resource
        _rss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
        # On Streamlit Cloud/Linux, ru_maxrss is reported in KiB.
        st.caption(f"Peak process memory (Linux): {_rss / 1024:.0f} MB")
    except Exception:
        st.caption("Memory metric unavailable on this server")
    st.caption("If the server restarts, check Streamlit Cloud logs for OOM / exit code.")
st.sidebar.caption("Upload the 2 source files used by the SPC app.")

uploaded_data_file = st.sidebar.file_uploader(
    "1️⃣ Production / SPC Data",
    type=["csv", "xlsx", "xls"],
    key="upload_spc_data",
    help="Main SPC dataset containing Time, 製造批號, 塗料編號, LAB and LINE color data."
)

uploaded_limit_file = st.sidebar.file_uploader(
    "2️⃣ Control Limit Data",
    type=["csv", "xlsx", "xls"],
    key="upload_limit_data",
    help="Control-limit table containing Color_code, Control_batch and LAB/LINE LCL-UCL columns."
)

@st.cache_data(show_spinner=False, max_entries=2, ttl=1800)
def read_uploaded_table(file_bytes, file_name):
    """Read an uploaded CSV/Excel file from bytes."""
    name = str(file_name).lower()
    bio = io.BytesIO(file_bytes)

    if name.endswith(".csv"):
        # Try common encodings used by exported Google Sheets / Taiwanese Windows files.
        last_error = None
        for enc in ["utf-8-sig", "utf-8", "cp950", "big5"]:
            try:
                bio.seek(0)
                return pd.read_csv(bio, encoding=enc)
            except UnicodeDecodeError as e:
                last_error = e
        raise last_error

    if name.endswith(".xlsx"):
        bio.seek(0)
        try:
            return pd.read_excel(bio, engine="openpyxl")
        except ImportError as e:
            raise ImportError(
                "Missing Excel reader 'openpyxl'. Please add openpyxl to requirements.txt and redeploy the app."
            ) from e

    if name.endswith(".xls"):
        bio.seek(0)
        try:
            return pd.read_excel(bio, engine="xlrd")
        except ImportError as e:
            raise ImportError(
                "Missing Excel reader 'xlrd'. Please add xlrd to requirements.txt and redeploy the app."
            ) from e

    raise ValueError("Unsupported file format. Please upload CSV, XLSX, or XLS.")


def fmt_num_trim(value, digits=2):
    """Format numeric values with up to 2 decimals; drop trailing zeros and decimal point."""
    if value is None:
        return "—"
    try:
        if pd.isna(value):
            return "—"
    except Exception:
        pass
    try:
        num = float(value)
    except Exception:
        return value
    if abs(num) < 10 ** (-(digits + 1)):
        num = 0.0
    rounded_int = round(num)
    if abs(num - rounded_int) < 10 ** (-(digits + 1)):
        return f"{int(rounded_int)}"
    out = f"{num:.{digits}f}".rstrip("0").rstrip(".")
    return "0" if out in ("-0", "-0.0") else out


def fmt_percent_trim(value, digits=2):
    base = fmt_num_trim(value, digits=digits)
    return base if base == "—" else f"{base}%"


# Stop here until both files have been uploaded.
if uploaded_data_file is None or uploaded_limit_file is None:
    st.title("📊 SPC Color Dashboard")
    st.info(
        "Please upload both source files in the left sidebar:\n\n"
        "**1. Production / SPC Data**  \n"
        "**2. Control Limit Data**"
    )
    st.stop()

try:
    df_raw = read_uploaded_table(uploaded_data_file.getvalue(), uploaded_data_file.name)
    limit_df = read_uploaded_table(uploaded_limit_file.getvalue(), uploaded_limit_file.name)
except Exception as e:
    st.error(f"❌ Unable to read uploaded file: {e}")
    st.stop()

# =========================
# BASIC FILE VALIDATION
# =========================
def clean_column_names(df):
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.replace("\r\n", " ", regex=False)
        .str.replace("\n", " ", regex=False)
        .str.replace("　", " ", regex=False)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    return df

df_raw = clean_column_names(df_raw)
limit_df = clean_column_names(limit_df)

# Validate ALL columns required by calculate_batch_averages() before any View runs.
# Never fill missing color measurements with zero: that would corrupt SPC.
required_data_cols = [
    "Time", "製造批號", "塗料編號",
    "入料檢測 ΔL 正面", "入料檢測 Δa 正面", "入料檢測 Δb 正面",
    "正-北 ΔL", "正-南 ΔL", "正-北 Δa", "正-南 Δa",
    "正-北 Δb", "正-南 Δb",
]
missing_data_cols = [c for c in required_data_cols if c not in df_raw.columns]
if missing_data_cols:
    st.error(
        "❌ File 1 (Production / SPC Data) is missing required columns: "
        + ", ".join(missing_data_cols)
    )
    st.caption("Available columns in Production / SPC Data:")
    st.code(" | ".join(map(str, df_raw.columns)))
    st.stop()

# Ambiguous or duplicate column headers can also produce incorrect calculations.
for source_name, frame in [("Production / SPC Data", df_raw), ("Control Limit Data", limit_df)]:
    duplicate_names = frame.columns[frame.columns.duplicated()].tolist()
    if duplicate_names:
        st.error(f"❌ Duplicate column names in {source_name}: {duplicate_names}")
        st.stop()

required_limit_cols = ["Color_code"]
missing_limit_cols = [c for c in required_limit_cols if c not in limit_df.columns]
if missing_limit_cols:
    st.error(
        "❌ File 2 (Control Limit Data) is missing required columns: "
        + ", ".join(missing_limit_cols)
    )
    st.stop()

# Parse time after upload, including mixed timestamp styles.
# Preserve the original Time column semantics; discard only invalid dates.
df_raw["Time"] = pd.to_datetime(df_raw["Time"], errors="coerce", format="mixed")
invalid_time_rows = int(df_raw["Time"].isna().sum())
df_raw = df_raw.dropna(subset=["Time"]).copy()
if df_raw.empty:
    st.error("❌ No valid Time values found in the uploaded Production / SPC Data.")
    st.stop()
if invalid_time_rows:
    st.sidebar.warning(f"⚠ {invalid_time_rows:,} rows removed: invalid Time values.")

st.sidebar.success(
    f"✅ Data loaded\n\n"
    f"File 1: {uploaded_data_file.name}\n\n"
    f"File 2: {uploaded_limit_file.name}"
)

# CRITICAL FIX: Remove NA, NaN, and Empty spaces from Batch Number globally
df_raw = df_raw.dropna(subset=["製造批號"])
df_raw["製造批號"] = df_raw["製造批號"].astype(str).str.strip()
df_raw = df_raw[df_raw["製造批號"] != ""]
df_raw = df_raw[df_raw["製造批號"].str.lower() != "nan"].copy()
if df_raw.empty:
    st.error("❌ No valid 製造批號 values remain after cleaning the uploaded data.")
    st.stop()

# Blank paint codes would make the color selector fail.
df_raw["塗料編號"] = df_raw["塗料編號"].astype("string").str.strip()
df_raw["塗料編號"] = df_raw["塗料編號"].replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
if df_raw["塗料編號"].dropna().empty:
    st.error("❌ No valid 塗料編號 (color codes) in the uploaded data.")
    st.stop()

numeric_columns = [
    "入料檢測 ΔL 正面", "入料檢測 Δa 正面", "入料檢測 Δb 正面",
    "正-北 ΔL", "正-南 ΔL", "正-北 Δa", "正-南 Δa", "正-北 Δb", "正-南 Δb",
    "Avergage Thickness", "Average value ΔE 正面", "Average value ΔL 正面", 
    "Average value Δa 正面", "Average value Δb 正面"
]
for col in numeric_columns:
    if col in df_raw.columns:
        df_raw[col] = pd.to_numeric(df_raw[col], errors='coerce')

# =========================
# DISPLAY LABELS (CIELAB notation)
# Keep raw/internal keys unchanged so source-file column names still work.
# =========================
DELTA_DISPLAY = {"ΔL": "ΔL*", "Δa": "Δa*", "Δb": "Δb*", "ΔE": "ΔE*"}

def delta_label(value):
    return DELTA_DISPLAY.get(value, value)

# =========================
# HELPER FUNCTIONS
# =========================
def safe_get_limit(c_code, src, fac):
    try:
        c_search = str(c_code).strip().upper()
        match_idx = -1
        if "Color_code" in limit_df.columns:
            for idx, val in limit_df["Color_code"].items():
                if str(val).strip().upper() == c_search:
                    match_idx = idx
                    break
        if match_idx == -1: return None, None
        
        row = limit_df.iloc[[match_idx]]
        src_up = str(src).strip().upper()
        fac1 = str(fac).strip().upper()
        fac2 = fac1.replace("Δ", "DELTA ")
        
        lcl, ucl = None, None
        for col in row.columns:
            cup = str(col).strip().upper()
            if src_up in cup and (fac1 in cup or fac2 in cup):
                val = row[col].values[0]
                try:
                    num = float(val)
                    if not pd.isna(num):
                        if "LCL" in cup: lcl = num
                        if "UCL" in cup: ucl = num
                except: pass
        return lcl, ucl
    except:
        return None, None

def get_kpi_control_limit_exact(c_code, src, fac):
    """
    KPI-report-only control-limit resolver.
    Existing tasks continue using safe_get_limit() unchanged.
    """
    try:
        if "Color_code" not in limit_df.columns:
            return None, None, None, None

        c_search = str(c_code).strip().upper()
        row = limit_df[
            limit_df["Color_code"].astype(str).str.strip().str.upper() == c_search
        ]
        if row.empty:
            return None, None, None, None

        src_up = str(src).strip().upper()
        fac_up = str(fac).strip().upper()
        factor_aliases = {
            "ΔL": ["ΔL", "DELTA L", "DELTA_L", "DELTAL"],
            "ΔA": ["ΔA", "DELTA A", "DELTA_A", "DELTAA"],
            "ΔB": ["ΔB", "DELTA B", "DELTA_B", "DELTAB"],
        }
        aliases = factor_aliases.get(fac_up, [fac_up, fac_up.replace("Δ", "DELTA ")])

        def normalize(s):
            s = str(s).upper().replace("　", " ")
            s = re.sub(r"[\-_:/()\[\]]+", " ", s)
            s = re.sub(r"\s+", " ", s).strip()
            return s

        def has_token(norm_name, token):
            return re.search(rf"(?:^|\s){re.escape(token)}(?:$|\s)", norm_name) is not None

        candidates = {"LCL": [], "UCL": []}
        for col in row.columns:
            norm = normalize(col)
            if not has_token(norm, src_up):
                continue
            if not any(normalize(a) in norm for a in aliases):
                continue
            for kind in ["LCL", "UCL"]:
                if has_token(norm, kind):
                    try:
                        val = float(row[col].iloc[0])
                        if not pd.isna(val):
                            candidates[kind].append((col, val, len(norm)))
                    except Exception:
                        pass

        def choose(items):
            if not items:
                return None, None
            items = sorted(items, key=lambda x: (x[2], str(x[0])))
            return items[0][1], items[0][0]

        lcl, lcl_col = choose(candidates["LCL"])
        ucl, ucl_col = choose(candidates["UCL"])
        return lcl, ucl, lcl_col, ucl_col
    except Exception:
        return None, None, None, None

def get_control_batch(color):
    c_search = str(color).strip().upper()
    mask = limit_df["Color_code"].astype(str).str.strip().str.upper() == c_search
    row = limit_df[mask]
    if row.empty: return None
    if "Control_batch" not in row.columns:
        return None  # Missing control-start metadata: leave Phase II undefined.
    value = row["Control_batch"].values[0]
    if pd.isna(value): return None
    if isinstance(value, str):
        m = re.search(r"\d+", value)
        if m: return int(m.group())
    try: return int(float(value))
    except: return None

def get_control_batch_code(df_unfiltered, control_batch):
    if control_batch is None or df_unfiltered.empty: return None
    batch_order = df_unfiltered.sort_values("Time")["製造批號"].drop_duplicates().tolist()
    if 1 <= control_batch <= len(batch_order): 
        return batch_order[control_batch - 1]
    return None

def detect_out_of_control(spc_df, lcl, ucl):
    mean = spc_df["value"].mean()
    std = spc_df["value"].std()
    result = spc_df.copy()
    result["Rule_CL"] = False
    result["Rule_3Sigma"] = False
    if lcl is not None and ucl is not None:
        result["Rule_CL"] = ((result["value"] < lcl) | (result["value"] > ucl))
    if std > 0:
        result["Rule_3Sigma"] = ((result["value"] > mean + 3 * std) | (result["value"] < mean - 3 * std))
    result["Out_of_Control"] = (result["Rule_CL"] | result["Rule_3Sigma"])
    return result[result["Out_of_Control"]]

SPC_INPUT_COLUMNS = ["製造批號", "Time"] + [
    col for factor in ("ΔL", "Δa", "Δb")
    for col in (f"正-北 {factor}", f"正-南 {factor}", f"入料檢測 {factor} 正面")
]

@st.cache_data(show_spinner=False, max_entries=3, ttl=1800)
def calculate_batch_averages(df_filtered_color):
    """Compute paired LAB/LINE batch means, preserving original SPC mathematics.

    Cached across Streamlit reruns for identical filtered data. Input and output
    contain only the essential columns; no repeated whole-frame copies per factor.
    """
    res = {}
    required = ["製造批號", "Time"] + [
        c for f in ("ΔL", "Δa", "Δb")
        for c in (f"正-北 {f}", f"正-南 {f}", f"入料檢測 {f} 正面")
    ]
    missing = [c for c in required if c not in df_filtered_color.columns]
    if missing:
        raise ValueError(f"Missing columns required for SPC batch calculation: {missing}")

    # Reuse one small numeric frame rather than copying the complete file 3 times.
    tmp = df_filtered_color.loc[:, required].copy()
    for c in required[2:]:
        tmp[c] = pd.to_numeric(tmp[c], errors="coerce")

    for f in ("ΔL", "Δa", "Δb"):
        col_n, col_s = f"正-北 {f}", f"正-南 {f}"
        col_lab = f"入料檢測 {f} 正面"
        # Same rules as before: NORTH/SOUTH row mean; group by batch; paired-only.
        row_avg = tmp[[col_n, col_s]].mean(axis=1)
        line_b = (
            pd.DataFrame({"製造批號": tmp["製造批號"], "Time": tmp["Time"], "row_avg": row_avg})
            .groupby("製造批號", as_index=False)
            .agg({"Time": "min", "row_avg": "mean"})
            .rename(columns={"row_avg": "value"})
            .sort_values("Time").dropna()
        )
        lab_b = (
            tmp.groupby("製造批號", as_index=False)
            .agg({"Time": "min", col_lab: "mean"})
            .rename(columns={col_lab: "value"})
            .sort_values("Time").dropna()
        )
        common_batches = set(line_b["製造批號"]).intersection(set(lab_b["製造批號"]))
        res[f] = {
            "line": line_b[line_b["製造批號"].isin(common_batches)],
            "lab": lab_b[lab_b["製造批號"].isin(common_batches)],
        }
    return res

# =========================
# SIDEBAR – NAVIGATION
# =========================
st.sidebar.markdown("### 📊 View Mode")
app_mode = st.sidebar.radio(
    "Select View Mode",
    [
        "🚀 Main Dashboard", 
        "📋 Limit Status Summary", 
        "🎛️ Control Limit Calculator",
        "🔬 Lab vs Line Scale-up",
        "📄 AI OOC Word Report",
        "📈 I-MR Chart (Coil-Level)" # THÊM DÒNG NÀY VÀO MENU
        ,"🎯 KPI Achievement Report"
    ],
    label_visibility="collapsed"
)

st.sidebar.divider()
st.sidebar.title("🎨 Filter")
color = st.sidebar.selectbox("Color code", sorted(df_raw["塗料編號"].dropna().unique()), key="sidebar_color")

df_color = df_raw[df_raw["塗料編號"] == color].copy()

# Synchronized Timeline
global_batch_order = df_color.sort_values("Time")["製造批號"].drop_duplicates().tolist()
control_batch = get_control_batch(color)
control_batch_code = get_control_batch_code(df_color, control_batch)
control_time = df_color[df_color["製造批號"] == control_batch_code]["Time"].min() if control_batch_code else None

all_years = sorted(df_color["Time"].dt.year.dropna().astype(int).unique())
selected_years = st.sidebar.multiselect("📅 Year (Leave empty for ALL)", options=all_years, default=[], key="sidebar_year")

all_months = sorted(df_color["Time"].dt.month.dropna().astype(int).unique())
selected_months = st.sidebar.multiselect("📅 Month (optional)", options=all_months, default=[], key="sidebar_month")

df = df_color.copy()
if len(selected_years) > 0: df = df[df["Time"].dt.year.isin(selected_years)]
if len(selected_months) > 0: df = df[df["Time"].dt.month.isin(selected_months)]

spc_data = calculate_batch_averages(df[SPC_INPUT_COLUMNS])

# =========================================================
# PLOTTING CORE ENGINES
# =========================================================
def spc_combined(lab, line, title, lab_lim, line_lim, control_batch_code, glb_order):
    fig, ax = plt.subplots(figsize=(12, 5))
    ax.set_facecolor('#f2f2f2')

    valid_batches = set(lab["製造批號"]).intersection(set(line["製造批號"]))
    local_order = [b for b in glb_order if b in valid_batches]

    lab_dict = dict(zip(lab["製造批號"], lab["value"]))
    line_dict = dict(zip(line["製造批號"], line["value"]))

    lab_x, lab_y = [], []
    line_x, line_y = [], []

    for i, b in enumerate(local_order):
        if b in lab_dict and b in line_dict:
            lab_x.append(i)
            lab_y.append(lab_dict[b])
            line_x.append(i)
            line_y.append(line_dict[b])

    if lab_x: ax.plot(lab_x, lab_y, marker="^", color="#548235", linestyle="-", linewidth=1.5, markersize=8, label="LAB Input")
    if line_x: ax.plot(line_x, line_y, marker="o", color="#ffc000", linestyle="-", linewidth=1.5, markersize=8, markerfacecolor="white", markeredgewidth=2, label="LINE Output")

    import matplotlib.transforms as transforms
    trans = transforms.blended_transform_factory(ax.transData, ax.transAxes)

    if control_batch_code is not None and control_batch_code in glb_order:
        global_ctrl_idx = glb_order.index(control_batch_code)
        ctrl_x = None
        for i, b in enumerate(local_order):
            if glb_order.index(b) >= global_ctrl_idx:
                ctrl_x = i
                break
        
        if ctrl_x is not None:
            ax.axvline(x=ctrl_x, color="#000000", linestyle=(0, (3, 3)), linewidth=1.5)
            ax.text(ctrl_x, 1.02, "  After Control", color="#0070c0", fontsize=14, ha="left", va="bottom", transform=trans)
            ax.text(ctrl_x, 1.02, "Before Control  ", color="#0070c0", fontsize=14, ha="right", va="bottom", transform=trans)

    if lab_lim[0] is not None and lab_lim[1] is not None and len(lab_y) > 0:
        lab_y_arr = np.array(lab_y); lab_x_arr = np.array(lab_x)
        out_lab = (lab_y_arr > lab_lim[1]) | (lab_y_arr < lab_lim[0])
        if out_lab.any(): ax.plot(lab_x_arr[out_lab], lab_y_arr[out_lab], marker="^", color="red", linestyle="None", markersize=10, zorder=5)
    
    if line_lim[0] is not None and line_lim[1] is not None and len(line_y) > 0:
        line_y_arr = np.array(line_y); line_x_arr = np.array(line_x)
        out_line = (line_y_arr > line_lim[1]) | (line_y_arr < line_lim[0])
        if out_line.any(): ax.plot(line_x_arr[out_line], line_y_arr[out_line], marker="o", color="red", linestyle="None", markersize=10, zorder=5)

    if lab_lim[0] is not None: 
        ax.axhline(lab_lim[0], color="#7030a0", linestyle="-", linewidth=2, label="LAB LCL")
        ax.axhline(lab_lim[1], color="#7030a0", linestyle="-", linewidth=2, label="LAB UCL")
    if line_lim[0] is not None: 
        ax.axhline(line_lim[0], color="#ff0000", linestyle="--", linewidth=2, label="LINE LCL")
        ax.axhline(line_lim[1], color="#ff0000", linestyle="--", linewidth=2, label="LINE UCL")

    if local_order:
        ax.set_xticks(range(len(local_order)))
        ax.set_xticklabels(local_order, rotation=45)

    ax.set_title(title, fontsize=15, fontweight="bold", pad=25)
    ax.grid(axis="y", color="#cccccc", linestyle="-", linewidth=1)
    ax.grid(axis="x", visible=False)
    ax.legend(bbox_to_anchor=(1.02, 1), loc="upper left", frameon=True, edgecolor="black")
    fig.subplots_adjust(right=0.8, top=0.8) 
    return fig

def spc_combined_phase2(lab, line, title, lab_lim, line_lim, control_batch_code, ctrl_time, glb_order):
    if ctrl_time is None: return None
    lab2 = lab[lab["Time"] >= ctrl_time]; line2 = line[line["Time"] >= ctrl_time]
    if lab2.empty and line2.empty: return None
    
    fig, ax = plt.subplots(figsize=(12, 5))
    ax.set_facecolor('#f2f2f2')

    valid_batches = set(lab2["製造批號"]).intersection(set(line2["製造批號"]))
    local_order = [b for b in glb_order if b in valid_batches]

    lab_dict = dict(zip(lab2["製造批號"], lab2["value"]))
    line_dict = dict(zip(line2["製造批號"], line2["value"]))

    lab_x, lab_y = [], []
    line_x, line_y = [], []

    for i, b in enumerate(local_order):
        if b in lab_dict and b in line_dict:
            lab_x.append(i)
            lab_y.append(lab_dict[b])
            line_x.append(i)
            line_y.append(line_dict[b])

    if lab_x: ax.plot(lab_x, lab_y, marker="^", color="#548235", linestyle="-", linewidth=1.5, markersize=8, label="LAB Input")
    if line_x: ax.plot(line_x, line_y, marker="o", color="#ffc000", linestyle="-", linewidth=1.5, markersize=8, markerfacecolor="white", markeredgewidth=2, label="LINE Output")

    import matplotlib.transforms as transforms
    trans = transforms.blended_transform_factory(ax.transData, ax.transAxes)
    
    if control_batch_code is not None and control_batch_code in glb_order:
        global_ctrl_idx = glb_order.index(control_batch_code)
        ctrl_x = None
        for i, b in enumerate(local_order):
            if glb_order.index(b) >= global_ctrl_idx:
                ctrl_x = i
                break
        if ctrl_x is not None:
            ax.axvline(x=ctrl_x, color="#000000", linestyle=(0, (3, 3)), linewidth=1.5)
            ax.text(ctrl_x, 1.02, "  After Control (Phase II)", color="#0070c0", fontsize=14, ha="left", va="bottom", transform=trans)

    if lab_lim[0] is not None and lab_lim[1] is not None and len(lab_y) > 0:
        lab_y_arr = np.array(lab_y); lab_x_arr = np.array(lab_x)
        out_lab = (lab_y_arr > lab_lim[1]) | (lab_y_arr < lab_lim[0])
        if out_lab.any(): ax.plot(lab_x_arr[out_lab], lab_y_arr[out_lab], marker="^", color="red", linestyle="None", markersize=10, zorder=6)
        
    if line_lim[0] is not None and line_lim[1] is not None and len(line_y) > 0:
        line_y_arr = np.array(line_y); line_x_arr = np.array(line_x)
        out_line = (line_y_arr > line_lim[1]) | (line_y_arr < line_lim[0])
        if out_line.any(): ax.plot(line_x_arr[out_line], line_y_arr[out_line], marker="o", color="red", linestyle="None", markersize=10, zorder=6)

    if lab_lim[0] is not None: 
        ax.axhline(lab_lim[0], color="#7030a0", linestyle="-", linewidth=2, label="LAB LCL")
        ax.axhline(lab_lim[1], color="#7030a0", linestyle="-", linewidth=2, label="LAB UCL")
    if line_lim[0] is not None: 
        ax.axhline(line_lim[0], color="#ff0000", linestyle="--", linewidth=2, label="LINE LCL")
        ax.axhline(line_lim[1], color="#ff0000", linestyle="--", linewidth=2, label="LINE UCL")

    if local_order:
        ax.set_xticks(range(len(local_order)))
        ax.set_xticklabels(local_order, rotation=45)

    ax.set_title(title, fontsize=15, fontweight="bold", pad=25)
    ax.grid(axis="y", color="#cccccc", linestyle="-", linewidth=1)
    ax.grid(axis="x", visible=False)
    ax.legend(bbox_to_anchor=(1.02, 1), loc="upper left", frameon=True, edgecolor="black")
    fig.subplots_adjust(right=0.8, top=0.8)
    return fig

def download(fig, name):
    buf = io.BytesIO(); fig.savefig(buf, format="png", dpi=200, bbox_inches="tight"); buf.seek(0)
    st.download_button("📥 Download PNG", buf, name, "image/png", key=f"dl_{name}")


# =========================================================
# TASK 3 — constrained LAB box candidate optimization.
# Existing IQR / Standard Method calculations are unchanged.
# This deterministic multi-start heuristic is NOT a proof of
# global mathematical optimality; report it as a candidate.
# =========================================================
@st.cache_data(show_spinner=False, max_entries=12, ttl=1800)
def task3_optimize_box(lab_values, original_lo, original_hi, lab_limit=0.50):
    points = np.asarray(lab_values, dtype=float)
    low, high = np.asarray(original_lo, float), np.asarray(original_hi, float)
    tol = 1e-10
    initial_ok = np.all((points >= low - tol) & (points <= high + tol), axis=1)
    initial_ok &= np.linalg.norm(points, axis=1) <= lab_limit + tol
    pts = points[initial_ok]
    original_n = len(pts)
    if original_n == 0:
        return None
    if float(np.linalg.norm(np.maximum(abs(low), abs(high)))) <= lab_limit + tol:
        return {'lo': low, 'hi': high, 'retained': original_n, 'candidate_pool': original_n}

    # Choose up to 96 diverse seeds and 128 distinct expansion candidates.
    # These caps keep Streamlit responsive on large production datasets.
    unique = np.unique(np.round(pts, decimals=9), axis=0)
    if len(unique) > 128:
        # Deterministic directional representatives + evenly spaced stable subset.
        center = np.median(unique, axis=0)
        directions = np.vstack([np.eye(3), -np.eye(3), np.ones((1, 3)), -np.ones((1, 3))])
        ids = {int(np.argmax((unique - center) @ d)) for d in directions}
        ids.update(np.linspace(0, len(unique) - 1, 128, dtype=int).tolist())
        unique = unique[sorted(ids)[:144]]
    if len(unique) > 40:
        seeds = unique[np.linspace(0, len(unique) - 1, 40, dtype=int)]
    else:
        seeds = unique

    best = None
    best_score = (-1, -1.0)
    for seed in seeds:
        lo = seed.copy()
        hi = seed.copy()
        chosen = np.zeros(len(unique), dtype=bool)
        # At most 45 inclusion steps; every step chooses a feasible new box.
        for _ in range(45):
            trial_lo = np.minimum(lo[None, :], unique)
            trial_hi = np.maximum(hi[None, :], unique)
            safe = np.linalg.norm(np.maximum(abs(trial_lo), abs(trial_hi)), axis=1) <= lab_limit + tol
            safe &= ~chosen
            if not np.any(safe):
                break
            candidate_ids = np.flatnonzero(safe)
            scores = []
            for i in candidate_ids:
                mask = np.all((pts >= trial_lo[i] - tol) & (pts <= trial_hi[i] + tol), axis=1)
                scores.append((int(mask.sum()), float(np.prod(np.maximum(trial_hi[i] - trial_lo[i], 1e-8)))))
            winner_pos = max(range(len(scores)), key=lambda j: scores[j])
            winner = candidate_ids[winner_pos]
            new_lo, new_hi = trial_lo[winner], trial_hi[winner]
            if np.array_equal(new_lo, lo) and np.array_equal(new_hi, hi):
                chosen[winner] = True
                continue
            lo, hi = new_lo, new_hi
            chosen[winner] = True
        count = int(np.all((pts >= lo - tol) & (pts <= hi + tol), axis=1).sum())
        score = (count, float(np.prod(np.maximum(hi - lo, 1e-8))))
        if score > best_score:
            best_score = score
            best = (lo.copy(), hi.copy())

    if best is None:
        return None
    lo, hi = best
    # Secondary objective: widen the box within original IQR without losing
    # the proven worst-corner LAB safety property. Coordinate-wise bisection.
    # Widen tighter sides first; no historic batch is lost by widening.
    for _ in range(10):
        prior = (lo.copy(), hi.copy())
        proposals = []
        for j in range(3):
            for side in ('lo', 'hi'):
                a, b = (low[j], lo[j]) if side == 'lo' else (hi[j], high[j])
                if b - a < 1e-12:
                    continue
                current = lo if side == 'lo' else hi
                left, right = a, b
                # For lower endpoints seek smallest feasible; for upper seek largest.
                for _ in range(35):
                    mid = (left + right) / 2
                    l2, h2 = lo.copy(), hi.copy()
                    if side == 'lo': l2[j] = mid
                    else: h2[j] = mid
                    safe = np.linalg.norm(np.maximum(abs(l2), abs(h2))) <= lab_limit + 1e-11
                    if side == 'lo':
                        if safe: right = mid
                        else: left = mid
                    else:
                        if safe: left = mid
                        else: right = mid
                new = right if side == 'lo' else left
                improvement = (lo[j] - new) if side == 'lo' else (new - hi[j])
                proposals.append((improvement, j, side, new))
        if not proposals:
            break
        improvement, j, side, new = max(proposals)
        if improvement < 1e-8:
            break
        if side == 'lo': lo[j] = new
        else: hi[j] = new
    lo = np.maximum(lo, low)
    hi = np.minimum(hi, high)
    if np.linalg.norm(np.maximum(abs(lo), abs(hi))) > lab_limit + 1e-8:
        return None
    retained = int(np.all((pts >= lo - tol) & (pts <= hi + tol), axis=1).sum())
    return {'lo': lo, 'hi': hi, 'retained': retained, 'candidate_pool': len(unique)}


def task3_simulate(lab_bounds_lo, lab_bounds_hi, shifts, method, seed=20261106, n=10000, lab_limit=0.50, line_limit=1.00):
    rng = np.random.default_rng(seed)
    shifts = np.asarray(shifts, dtype=float)
    if len(shifts) < 3:
        return {'error': '有效歷史 Shift 少於 3 批，無法執行情境模擬。'}
    batches = []
    count = 0
    for _ in range(70):
        k = max(3000, min(60000, (n - count) * 5))
        samples = rng.uniform(lab_bounds_lo, lab_bounds_hi, size=(k, 3))
        if method == 'A':
            samples = samples[np.linalg.norm(samples, axis=1) <= lab_limit + 1e-12]
        if len(samples):
            batches.append(samples[:n - count])
            count += min(len(samples), n - count)
        if count >= n:
            break
    if count < n:
        return {'error': '有效 LAB 情境抽樣不足；請檢查 IQR 與 ΔE* 條件。'}
    lab_samples = np.concatenate(batches, axis=0)
    ids = rng.integers(0, len(shifts), size=n)
    sim_line = lab_samples + shifts[ids]
    de = np.linalg.norm(sim_line, axis=1)
    return {'n': n, 'bad': int(np.sum(de > line_limit + 1e-12)),
            'p95': float(np.quantile(de, 0.95)), 'lab_samples': lab_samples,
            'shift_ids': ids, 'shift_samples': shifts[ids], 'line_de': de}

# =========================================================
# TASK 3: LINE-risk-constrained candidate LAB box
# Use entire paired, in-spec LAB history as calibration shift scenarios.
# A historical scenario stress test is NOT a future guarantee.
# =========================================================
def task3_shift_envelope_peak(lo, hi, shift_min, shift_max):
    """Conservative LINE ΔE* bound for an independent Min/Max Shift envelope.

    Every component's shift can vary inside its observed marginal range;
    this includes combinations never jointly observed in one batch.
    """
    lo, hi = np.asarray(lo, float), np.asarray(hi, float)
    s0, s1 = np.asarray(shift_min, float), np.asarray(shift_max, float)
    endpoints = np.stack([lo + s0, lo + s1, hi + s0, hi + s1])
    return float(np.linalg.norm(np.max(np.abs(endpoints), axis=0)))


def task3_line_risk_box(lab_points, shift_vectors, lab_limit, line_limit):
    """Simple backwards-LAB limits using observed Shift Min/Max.

    Steps:
      1. Shift Min/Max from complete paired historical batches with LAB in-spec.
      2. Control center = -(Shift Min + Shift Max)/2, clipped into the
         observed in-spec LAB component range. This compensates the
         mid-shift (not an arbitrary shrink percentage).
      3. Expand towards the observed LAB min/max with one scale t in [0,1].
         Binary search for maximum t satisfying both LAB ΔE* and the
         independent-component Shift-envelope LINE ΔE* limits.
    
    Feasible only relative to recorded Shift extrema, NOT future guarantee.
    """
    pts = np.asarray(lab_points, float)
    shifts = np.asarray(shift_vectors, float)
    pts = pts[np.isfinite(pts).all(axis=1)] if pts.ndim == 2 and pts.shape[1] == 3 else np.empty((0,3))
    shifts = shifts[np.isfinite(shifts).all(axis=1)] if shifts.ndim == 2 and shifts.shape[1] == 3 else np.empty((0,3))
    if len(pts) < 3 or len(shifts) < 3:
        return {'error':'LAB 與 LINE 有效配對少於 3 批，資料不足，暫不提出候選界限。'}
    low, high = pts.min(axis=0), pts.max(axis=0)
    shift_min, shift_max = shifts.min(axis=0), shifts.max(axis=0)
    span = high - low
    if np.any(span <= 1e-12):
        return {'error':'至少一項 LAB 因子沒有有效歷史範圍，無法推導三因子界限。'}
    mid_shift = (shift_min + shift_max)/2
    center = np.clip(-mid_shift, low, high)

    def make_box(t):
        return center + t*(low-center), center + t*(high-center)

    def check(t):
        lo, hi = make_box(t)
        lab_peak = float(np.linalg.norm(np.maximum(np.abs(lo), np.abs(hi))))
        line_peak = task3_shift_envelope_peak(lo, hi, shift_min, shift_max)
        return lab_peak, line_peak

    base_lab, base_line = check(0.0)
    if base_lab > lab_limit + 1e-10 or base_line > line_limit + 1e-10:
        return {'error':(
            f'按歷史 Shift 中心補償所得的 LAB 中心點，LAB ΔE*={base_lab:.2f}、'
            f'LINE 保守上界={base_line:.2f}，尚無法同時滿足本色號規格。'
            '此簡易方法不強行輸出界限，建議擴充配對數據或使用進階模型。'),
            'center':center, 'center_lab':base_lab, 'center_line':base_line,
            'shift_min':shift_min,'shift_max':shift_max}
    left, right = 0.0, 1.0
    if any(v > limit + 1e-10 for v,limit in zip(check(1.0),(lab_limit,line_limit))):
        for _ in range(65):
            mid=(left+right)/2
            lab_val,line_val=check(mid)
            if lab_val <= lab_limit and line_val <= line_limit:
                left=mid
            else:
                right=mid
        t=left
    else:
        t=1.0
    lo,hi=make_box(t)
    lab_peak,line_peak=check(t)
    if t <= 1e-8 or np.any((hi-lo) <= 1e-9):
        return {'error':'目前歷史 Shift 變動過大，簡易反推法找不到有實際寬度的三因子界限。'}
    return {'lo':lo,'hi':hi,'lab_worst':lab_peak,'line_worst':line_peak,
            'historical_shift_n':len(shifts),'observed_lo':low,'observed_hi':high,
            'shift_min':shift_min,'shift_max':shift_max,'shift_mid':mid_shift,
            'center':center,'scale':t,
            'method':'Shift Min/Max → 補償中心 → 由實測 LAB 範圍擴大；二分搜尋滿足 LAB 與 LINE 規格的最大擴大係數 t。'
                     ' LINE 採各項 Shift Min/Max 獨立組合之保守上界，非未來保證。'}

# =========================================================
# TASK 3 AUDIT: exact historical-shift stress test for a LAB box.
# A finite historical shift set cannot establish future risk probability.
# =========================================================
def task3_stress_audit(lo, hi, shift_df, line_spec):
    if lo is None or hi is None or shift_df is None or shift_df.empty:
        return pd.DataFrame(), None
    lo = np.asarray(lo, dtype=float)
    hi = np.asarray(hi, dtype=float)
    rows = []
    for _, r in shift_df.iterrows():
        shift = np.array([r['Shift_ΔL'], r['Shift_Δa'], r['Shift_Δb']], dtype=float)
        # Axis-aligned box: each coordinate chooses the worse endpoint.
        left, right = lo + shift, hi + shift
        worst_corner = np.where(np.abs(left) >= np.abs(right), lo, hi)
        worst_line = worst_corner + shift
        peak = float(np.linalg.norm(worst_line))
        rows.append({
            'Shift 來源批號': str(r['製造批號']),
            'Shift ΔL*': shift[0], 'Shift Δa*': shift[1], 'Shift Δb*': shift[2],
            '最不利 LAB ΔL*': worst_corner[0],
            '最不利 LAB Δa*': worst_corner[1],
            '最不利 LAB Δb*': worst_corner[2],
            '預測 LINE ΔL*': worst_line[0],
            '預測 LINE Δa*': worst_line[1],
            '預測 LINE Δb*': worst_line[2],
            '最不利 LINE ΔE*': peak,
            'LINE 規格餘裕': float(line_spec) - peak,
            '判定': '符合歷史情境' if peak <= line_spec + 1e-9 else '歷史情境超規',
        })
    detail = pd.DataFrame(rows).sort_values('最不利 LINE ΔE*', ascending=False).reset_index(drop=True)
    return detail, detail.iloc[0].to_dict() if not detail.empty else None



def task3_worst_case_calculation_tables(case, paired_df, line_spec, box_lo, box_hi):
    """Transparent audit: same-source measured LAB/LINE -> shift -> predicted LINE.

    All calculations use full-precision values. Display rounding happens only in UI.
    """
    if case is None or paired_df is None or paired_df.empty or box_lo is None or box_hi is None:
        return None
    batch = str(case['Shift 來源批號'])
    matches = paired_df.loc[paired_df['製造批號'].astype(str) == batch]
    if len(matches) != 1:
        return None
    source = matches.iloc[0]
    keys = ('ΔL', 'Δa', 'Δb')
    input_l = np.array([float(source[f'LAB_{j}']) for j in keys])
    output_l = np.array([float(source[f'LINE_{j}']) for j in keys])
    shifts = output_l - input_l
    worst_lab = np.array([float(case[f'最不利 LAB {j}*']) for j in keys])
    predicted = worst_lab + shifts
    square = predicted ** 2
    de = float(np.sqrt(square.sum()))
    margin = float(line_spec) - de
    source_table = pd.DataFrame([
        {'色差因子': f'{k}*', '來源 LAB 實測值':input_l[i], '來源 LINE 實測值':output_l[i],
         'Shift = LINE − LAB':shifts[i], '計算公式':f'{output_l[i]:.2f} − ({input_l[i]:.2f})'}
        for i,k in enumerate(keys)
    ])
    prediction_table = pd.DataFrame([
        {'色差因子':f'{k}*', '候選 LAB 最不利點 (A)':worst_lab[i], '來源 Shift (B)':shifts[i],
         'LINE 預測值 (A+B)':predicted[i], '預測 LINE 平方':square[i],
         '計算公式':f'({worst_lab[i]:.2f} + ({shifts[i]:.2f}))²'}
        for i,k in enumerate(keys)
    ])
    totals_table = pd.DataFrame([
        {'驗證項目':'Shift 來源批號', '計算公式 / 資料來源':'配對實測 LAB / LINE → SHIFT = LINE − LAB', '計算結果':batch},
        {'驗證項目':'候選 LAB ΔE*', '計算公式 / 資料來源':'SQRT(' + ' + '.join(f'({x:.2f})²' for x in worst_lab) + ')', '計算結果':float(np.linalg.norm(worst_lab))},
        {'驗證項目':'預測 LINE ΔE*', '計算公式 / 資料來源':'SQRT(' + ' + '.join(f'({x:.2f})²' for x in predicted) + ')', '計算結果':de},
        {'驗證項目':'LINE ΔE* 規格', '計算公式 / 資料來源':'使用者輸入 LINE 規格上限', '計算結果':float(line_spec)},
        {'驗證項目':'LINE 規格餘裕', '計算公式 / 資料來源':f'{float(line_spec):.2f} − ({de:.2f})', '計算結果':margin},
        {'驗證項目':'最不利情境判定', '計算公式 / 資料來源':'全精度 LINE ΔE* ≤ LINE 規格', '計算結果':'符合歷史情境' if de <= float(line_spec)+1e-9 else '歷史情境超規'},
    ])
    # Display all 8 extreme LAB corners for THIS worst historical shift.
    # Global worst comes from comparing all historical shifts x 8 corners.
    from itertools import product
    corners=[]
    for flags in product((0,1),repeat=3):
        lab = np.where(np.asarray(flags,dtype=bool), np.asarray(box_hi,float), np.asarray(box_lo,float))
        line = lab + shifts
        corners.append({'界限組合':''.join('U' if flag else 'L' for flag in flags),
                        'LAB ΔL*':lab[0], 'LAB Δa*':lab[1], 'LAB Δb*':lab[2],
                        'LINE ΔL*':line[0], 'LINE Δa*':line[1], 'LINE Δb*':line[2],
                        'LINE ΔE*':float(np.linalg.norm(line)),
                        'LINE 判定':'OK' if np.linalg.norm(line) <= line_spec+1e-9 else 'NG'})
    return {'source':source_table, 'prediction':prediction_table, 'verification':totals_table,
            'corners':pd.DataFrame(corners), 'batch':batch, 'de':de, 'margin':margin}


def task3_shift_summary(shift_df):
    rows = []
    for factor in ('ΔL', 'Δa', 'Δb'):
        v = pd.to_numeric(shift_df['Shift_' + factor], errors='coerce').dropna()
        if v.empty:
            continue
        rows.append({'因子': factor + '*', '樣本數':len(v),
                     '平均 Shift':float(v.mean()), '最小 Shift':float(v.min()),
                     '最大 Shift':float(v.max()),
                     '變動幅度 (Max-Min)':float(v.max()-v.min()),
                     '標準差 SD':float(v.std(ddof=1)) if len(v)>1 else np.nan})
    return pd.DataFrame(rows)


# =========================================================
# VIEW 1: MAIN DASHBOARD
# =========================================================
if app_mode == "🚀 Main Dashboard":

    st.sidebar.divider()
    if control_batch_code is not None:
        st.sidebar.info(f"🔔 **Control batch**\n\nBatch #{control_batch} → **{control_batch_code}**")
    elif control_batch is not None:
        st.sidebar.warning(f"⚠ Control batch #{control_batch} exceeds available batches")
    st.sidebar.divider()

    def show_limits(factor):
        c_search = str(color).strip().upper()
        mask = limit_df["Color_code"].astype(str).str.strip().str.upper() == c_search
        row = limit_df[mask]
        if row.empty: return
        table = row.filter(like=factor).copy()
        for c in table.columns:
            table[c] = table[c].map(
                lambda x: f"{float(x):.2f}" if pd.notna(x) and pd.to_numeric(x, errors="coerce") == pd.to_numeric(x, errors="coerce") else (str(x) if pd.notna(x) else "")
            )
        st.sidebar.markdown(f"**{factor} Control Limits (Sheet)**")
        st.sidebar.dataframe(table, use_container_width=True, hide_index=True)

    show_limits("LAB")
    show_limits("LINE")

    st.title("室內隔間用途－塗料入料管控專案")
    st.caption("Incoming Paint SPC · LAB / LINE · Phase II Monitoring")
    st.title(f"📊 SPC Color Dashboard — {color}")

    if not df.empty:
        t_min = df["Time"].min().strftime("%Y-%m-%d")
        t_max = df["Time"].max().strftime("%Y-%m-%d")
        n_batch = df["製造批號"].nunique()
        display_year = "ALL" if len(selected_years) == 0 else ", ".join(map(str, selected_years))
        display_month = "ALL" if len(selected_months) == 0 else ", ".join(map(str, selected_months))
    else:
        t_min = t_max = "N/A"; n_batch = 0; display_year = "N/A"; display_month = "N/A"

    st.markdown(f"⏱ **{t_min} → {t_max} | n = {n_batch} batches | Year: {display_year} | Month: {display_month}**")

    # =========================================================
    # EXECUTIVE SUMMARY STATISTICS (OVERALL VS. PHASE II)
    # =========================================================
    def build_comparison_table(source_key):
        rows = []
        for k in ["ΔL", "Δa", "Δb", "ΔE"]:
            try:
                # 1. Fetch Data (Handle normal factors vs ΔE)
                if k != "ΔE":
                    df_overall = spc_data[k][source_key].dropna(subset=["value"]).copy()
                else:
                    df_L = spc_data["ΔL"][source_key][["製造批號", "Time", "value"]].rename(columns={"value": "L"})
                    df_a = spc_data["Δa"][source_key][["製造批號", "value"]].rename(columns={"value": "a"})
                    df_b = spc_data["Δb"][source_key][["製造批號", "value"]].rename(columns={"value": "b"})
                    df_overall = df_L.merge(df_a, on="製造批號").merge(df_b, on="製造批號")
                    df_overall["value"] = np.sqrt(df_overall["L"]**2 + df_overall["a"]**2 + df_overall["b"]**2)
                    df_overall = df_overall.dropna(subset=["value"]).copy()
                
                if df_overall.empty: continue
                
                # 2. Calculate Overall Stats
                m_all = df_overall["value"].mean()
                s_all = df_overall["value"].std()
                min_all = df_overall["value"].min()
                max_all = df_overall["value"].max()
                n_all = len(df_overall)
                
                row = {
                    "Factor": f"**{delta_label(k)}**",
                    "Overall Mean": f"{m_all:.3f}",
                    "Phase II Mean": "-",
                    "Overall Std": f"{s_all:.3f}",
                    "Phase II Std": "-",
                    "Overall Min": f"{min_all:.3f}",
                    "Phase II Min": "-",
                    "Overall Max": f"{max_all:.3f}",
                    "Phase II Max": "-",
                    "Overall (n)": n_all,
                    "Phase II (n)": "-"
                }
                
                # 3. Calculate Phase II Stats (if available) & Improvement
                if control_time is not None:
                    df_p2 = df_overall[df_overall["Time"] >= control_time]
                    if not df_p2.empty:
                        m_p2 = df_p2["value"].mean()
                        s_p2 = df_p2["value"].std()
                        min_p2 = df_p2["value"].min()
                        max_p2 = df_p2["value"].max()
                        n_p2 = len(df_p2)
                        
                        # Format Phase II Mean with Delta for ΔE (Lower is better)
                        mean_str = f"{m_p2:.3f}"
                        if k == "ΔE" and m_all > 0:
                            imp_m = ((m_all - m_p2) / m_all) * 100
                            if imp_m > 0: mean_str += f" (↓ {imp_m:.1f}%)"
                            elif imp_m < 0: mean_str += f" (↑ {-imp_m:.1f}%)"
                        row["Phase II Mean"] = mean_str
                        
                        # Format Phase II Std with Delta for all factors
                        std_str = f"{s_p2:.3f}"
                        if pd.notna(s_all) and pd.notna(s_p2) and s_all > 0:
                            imp_s = ((s_all - s_p2) / s_all) * 100
                            if imp_s > 0: std_str += f" (↓ {imp_s:.1f}%)"
                            elif imp_s < 0: std_str += f" (↑ {-imp_s:.1f}%)"
                        row["Phase II Std"] = std_str
                        
                        row["Phase II Min"] = f"{min_p2:.3f}"
                        row["Phase II Max"] = f"{max_p2:.3f}"
                        row["Phase II (n)"] = n_p2
                        
                rows.append(row)
            except Exception:
                continue
                
        df_out = pd.DataFrame(rows)
        if not df_out.empty:
            cols_order = ["Factor", "Overall Mean", "Phase II Mean", "Overall Std", "Phase II Std", 
                          "Overall Min", "Phase II Min", "Overall Max", "Phase II Max", 
                          "Overall (n)", "Phase II (n)"]
            df_out = df_out[cols_order]
            
        return df_out

    df_summary_line = build_comparison_table("line")
    df_summary_lab = build_comparison_table("lab")

    st.markdown("### 📋 Executive Performance Summary (Overall vs. Phase II)")
    st.info("💡 **How to read:** Arrows (↓) in Phase II indicate a percentage reduction in color drift (Mean ΔE*) or process variation (Std). **Lower is better.**")
    
    col_sum1, col_sum2 = st.columns(2)
    with col_sum1: 
        st.markdown("#### 🏭 LINE (Production)")
        if not df_summary_line.empty:
            st.dataframe(df_summary_line, use_container_width=True, hide_index=True)
        else:
            st.warning("No line data available.")
            
    with col_sum2: 
        st.markdown("#### 🧪 LAB (Input)")
        if not df_summary_lab.empty:
            st.dataframe(df_summary_lab, use_container_width=True, hide_index=True)
        else:
            st.warning("No lab data available.")


    with st.expander("🔎 Batch Summary (Before SPC Aggregation)"):
        if not df.empty:
            batch_summary = (
                df.groupby("製造批號", as_index=False)
                .agg(
                    First_Time=("Time", "min"),
                    LAB_ΔL=("入料檢測 ΔL 正面", "mean"), LAB_Δa=("入料檢測 Δa 正面", "mean"), LAB_Δb=("入料檢測 Δb 正面", "mean"),
                    LINE_ΔL_N=("正-北 ΔL", "mean"), LINE_ΔL_S=("正-南 ΔL", "mean"),
                    LINE_Δa_N=("正-北 Δa", "mean"), LINE_Δa_S=("正-南 Δa", "mean"),
                    LINE_Δb_N=("正-北 Δb", "mean"), LINE_Δb_S=("正-南 Δb", "mean"),
                    Rows_in_Batch=("製造批號", "count")
                ).sort_values("First_Time")
            )
            batch_summary["LINE_ΔL"] = batch_summary[["LINE_ΔL_N", "LINE_ΔL_S"]].mean(axis=1)
            batch_summary["LINE_Δa"] = batch_summary[["LINE_Δa_N", "LINE_Δa_S"]].mean(axis=1)
            batch_summary["LINE_Δb"] = batch_summary[["LINE_Δb_N", "LINE_Δb_S"]].mean(axis=1)
            display_cols = ["製造批號", "First_Time", "LAB_ΔL", "LAB_Δa", "LAB_Δb", "LINE_ΔL", "LINE_Δa", "LINE_Δb", "Rows_in_Batch"]
            batch_summary_display = batch_summary[display_cols].rename(columns={
                "LAB_ΔL": "LAB_ΔL*", "LAB_Δa": "LAB_Δa*", "LAB_Δb": "LAB_Δb*",
                "LINE_ΔL": "LINE_ΔL*", "LINE_Δa": "LINE_Δa*", "LINE_Δb": "LINE_Δb*"
            })
            st.dataframe(batch_summary_display, use_container_width=True, hide_index=True)
        else:
            st.warning("No data after filtering.")

    st.markdown("### 📊 CONTROL CHART: LAB-LINE")
    for k in ["ΔL", "Δa", "Δb"]:
        lab_lim = safe_get_limit(color, "LAB", k)
        line_lim = safe_get_limit(color, "LINE", k)
        fig = spc_combined(spc_data[k]["lab"], spc_data[k]["line"], f"COMBINED {delta_label(k)}", lab_lim, line_lim, control_batch_code, global_batch_order)
        st.pyplot(fig); download(fig, f"COMBINED_{color}_{k}.png"); plt.close(fig)

    st.markdown("---")
    st.subheader("📊 SPC Combined Chart (LAB + LINE) – Phase II")
    for k in ["ΔL", "Δa", "Δb"]:
        lab_lim = safe_get_limit(color, "LAB", k)
        line_lim = safe_get_limit(color, "LINE", k)
        fig = spc_combined_phase2(spc_data[k]["lab"], spc_data[k]["line"], f"{delta_label(k)} – LAB + LINE (Phase II)", lab_lim, line_lim, control_batch_code, control_time, global_batch_order)
        if fig is not None: st.pyplot(fig); download(fig, f"COMBINED_PHASE2_{color}_{k}.png"); plt.close(fig)
        else: st.info(f"{k}: Not enough Phase II data")

    st.markdown("---")
    st.markdown("## 📈 Line Process Distribution Dashboard")
    def normal_pdf(x, mean, std): return (1 / (std * math.sqrt(2 * math.pi))) * np.exp(-0.5 * ((x - mean) / std) ** 2)
    cols = st.columns(3)
    for i, k in enumerate(["ΔL", "Δa", "Δb"]):
        with cols[i]:
            values = spc_data[k]["line"]["value"].dropna()
            if len(values) < 3: st.warning("Not enough data"); continue
            mean, std = values.mean(), values.std(); lcl, ucl = safe_get_limit(color, "LINE", k)
            fig, ax = plt.subplots(figsize=(5, 4))
            bins = np.histogram_bin_edges(values, bins=10)
            counts, _, patches = ax.hist(values, bins=bins, edgecolor="white", color="#4dabf7", alpha=0.85)
            for p, l, r in zip(patches, bins[:-1], bins[1:]):
                center = (l + r) / 2
                if lcl is not None and ucl is not None and (center < lcl or center > ucl): p.set_facecolor("#ff6b6b")
            if std > 0:
                x = np.linspace(mean - 4 * std, mean + 4 * std, 500)
                ax.plot(x, normal_pdf(x, mean, std) * len(values) * (bins[1] - bins[0]), color="black", linewidth=2)
            if lcl is not None: ax.axvline(lcl, color="red", linestyle="--", linewidth=1.5, label="LSL")
            if ucl is not None: ax.axvline(ucl, color="red", linestyle="--", linewidth=1.5, label="USL")
            ax.text(0.02, 0.95, f"N = {len(values)}\nMean = {mean:.3f}\nStd = {std:.3f}", transform=ax.transAxes, va="top", fontsize=9, bbox=dict(facecolor="white", alpha=0.9))
            ax.set_title(f"{delta_label(k)} (LINE)"); ax.grid(axis="y", alpha=0.3); ax.legend(fontsize=8); st.pyplot(fig)
            buf = io.BytesIO(); fig.savefig(buf, format="png", dpi=150, bbox_inches="tight"); buf.seek(0)
            st.download_button("⬇ Download", data=buf, file_name=f"{k}_line_dist.png", mime="image/png", key=f"dl_line_dist_{k}")
            plt.close(fig)

    st.markdown("---")
    st.markdown("## 🧪 LAB Process Distribution Dashboard")
    cols = st.columns(3)
    for i, k in enumerate(["ΔL", "Δa", "Δb"]):
        with cols[i]:
            values = spc_data[k]["lab"]["value"].dropna()
            if len(values) < 3: st.warning("Not enough data"); continue
            mean, std = values.mean(), values.std(); lcl, ucl = safe_get_limit(color, "LAB", k)
            fig, ax = plt.subplots(figsize=(5, 4))
            bins = np.histogram_bin_edges(values, bins=10)
            counts, _, patches = ax.hist(values, bins=bins, edgecolor="white", color="#1f77b4", alpha=0.85)
            for p, l, r in zip(patches, bins[:-1], bins[1:]):
                center = (l + r) / 2
                if lcl is not None and ucl is not None and (center < lcl or center > ucl): p.set_facecolor("#ff6b6b")
            if std > 0:
                x = np.linspace(mean - 4 * std, mean + 4 * std, 500)
                ax.plot(x, normal_pdf(x, mean, std) * len(values) * (bins[1] - bins[0]), color="black", linewidth=2)
            if lcl is not None: ax.axvline(lcl, color="red", linestyle="--", linewidth=1.5, label="LSL")
            if ucl is not None: ax.axvline(ucl, color="red", linestyle="--", linewidth=1.5, label="USL")
            ax.text(0.02, 0.95, f"N = {len(values)}\nMean = {mean:.3f}\nStd = {std:.3f}", transform=ax.transAxes, va="top", fontsize=9, bbox=dict(facecolor="white", alpha=0.9))
            ax.set_title(f"{delta_label(k)} (LAB)"); ax.grid(axis="y", alpha=0.3); ax.legend(fontsize=8); st.pyplot(fig)
            buf = io.BytesIO(); fig.savefig(buf, format="png", dpi=150, bbox_inches="tight"); buf.seek(0)
            st.download_button("⬇ Download", data=buf, file_name=f"{k}_lab_dist.png", mime="image/png", key=f"dl_lab_dist_{k}")
            plt.close(fig)

    # Section: OOC TABLE
    st.markdown("## 🚨 Out-of-Control Batches")
    ooc_rows = []
    
    for k in ["ΔL", "Δa", "Δb"]:
        lcl, ucl = safe_get_limit(color, "LINE", k)
        if control_time is not None:
            line_phase2 = spc_data[k]["line"][spc_data[k]["line"]["Time"] >= control_time]
            ooc_line = detect_out_of_control(line_phase2, lcl, ucl)
            for _, r in ooc_line.iterrows(): 
                ooc_rows.append({"Factor": delta_label(k), "Type": "LINE", "Batch No.": r["製造批號"], "Value": round(r["value"], 2), "Rule_CL": r["Rule_CL"], "Rule_3Sigma": r["Rule_3Sigma"]})
        
        lcl, ucl = safe_get_limit(color, "LAB", k)
        if control_time is not None:
            lab_phase2 = spc_data[k]["lab"][spc_data[k]["lab"]["Time"] >= control_time]
            ooc_lab = detect_out_of_control(lab_phase2, lcl, ucl)
            for _, r in ooc_lab.iterrows(): 
                ooc_rows.append({"Factor": delta_label(k), "Type": "LAB", "Batch No.": r["製造批號"], "Value": round(r["value"], 2), "Rule_CL": r["Rule_CL"], "Rule_3Sigma": r["Rule_3Sigma"]})
    
    if ooc_rows: 
        ooc_df = pd.DataFrame(ooc_rows)
        st.dataframe(ooc_df, use_container_width=True)
    else: 
        st.success("✅ No out-of-control batches detected")

    # Section: THICKNESS CORRELATION
    st.markdown("---")
    st.header("🎨 Thickness – Color Analysis (Per Coil)")
    
    coil_col, time_col, thickness_col = "Coil No.", "Time", "Avergage Thickness"
    dE_col, dL_col, da_col, db_col = "Average value ΔE 正面", "Average value ΔL 正面", "Average value Δa 正面", "Average value Δb 正面"
    required_cols = [coil_col, time_col, thickness_col, dE_col, dL_col, da_col, db_col]
    missing = [c for c in required_cols if c not in df.columns]

    if missing:
        st.error(f"❌ Missing required columns: {missing}")
    else:
        df_plot = df.copy()
        df_plot[time_col] = pd.to_datetime(df_plot[time_col], errors="coerce")
        df_plot = df_plot.dropna(subset=[time_col])
        df_plot["Year"] = df_plot[time_col].dt.year
        df_plot["Month"] = df_plot[time_col].dt.to_period("M").astype(str)

        st.subheader("⏱ Time Filter")
        col1, col2 = st.columns(2)
        
        with col1: 
            filter_mode_bottom = st.radio("Filter by", ["Month", "Year"], horizontal=True, key="bottom_filter_mode")
        
        with col2:
            if filter_mode_bottom == "Month":
                all_months = sorted(df_plot["Month"].unique())
                month_sel = st.multiselect("Select month(s) [Leave empty to show all]", all_months, default=[], key="bottom_month_sel")
                if month_sel: 
                    df_plot = df_plot[df_plot["Month"].isin(month_sel)]
            else:
                all_years = sorted(df_plot["Year"].unique())
                year_sel = st.multiselect("Select year(s) [Leave empty to show all]", all_years, default=[], key="bottom_year_sel")
                if year_sel: 
                    df_plot = df_plot[df_plot["Year"].isin(year_sel)]

        if df_plot.empty:
            st.warning("⚠️ No data available for the selected time period.")
        else:
            st.subheader("📊 Average Thickness vs ΔE* (Each Point = 1 Coil)")
            fig, ax = plt.subplots(figsize=(10, 6))
            ax.scatter(df_plot[thickness_col], df_plot[dE_col], alpha=0.75)
            if len(df_plot) > 0: 
                ax.axhline(df_plot[dE_col].mean(), linestyle="--", linewidth=2, label=f"Mean ΔE* = {df_plot[dE_col].mean():.2f}")
            ax.set_xlabel("Average Thickness")
            ax.set_ylabel("ΔE*")
            ax.set_title("Thickness – Color Relationship per Coil")
            ax.legend()
            ax.grid(True, linestyle="--", alpha=0.4)
            st.pyplot(fig)
            plt.close(fig)

            st.subheader("📈 ΔE* Distribution (Per Coil)")
            fig2, ax2 = plt.subplots(figsize=(10, 4))
            data_de = df_plot[dE_col].dropna()
            if len(data_de) > 0:
                mean_de, std_de = data_de.mean(), data_de.std()
                ax2.hist(data_de, bins=20, density=True, alpha=0.7, edgecolor="black", label="ΔE* Histogram")
                if std_de > 0:
                    x_de = np.linspace(mean_de - 5*std_de, mean_de + 5*std_de, 1000)
                    ax2.plot(x_de, normal_pdf(x_de, mean_de, std_de), linewidth=3, label="Normal Distribution")
                ax2.axvline(mean_de, linestyle="--", linewidth=2, label=f"Mean = {mean_de:.2f}")
                ax2.set_xlabel("ΔE*")
                ax2.set_ylabel("Density")
                ax2.set_title("ΔE* Distribution")
                ax2.legend()
                ax2.grid(True, linestyle="--", alpha=0.4)
                st.pyplot(fig2)
            plt.close(fig2)

            st.subheader("📊 Average Thickness Distribution")
            data = df_plot[thickness_col].dropna()
            if len(data) >= 2 and pd.notna(data.std()) and data.std() > 0:
                mean, std = data.mean(), data.std()
                col1, col2 = st.columns(2)
                with col1: LSL = st.number_input("LSL", value=float(mean - 3 * std), key="bottom_lsl")
                with col2: USL = st.number_input("USL", value=float(mean + 3 * std), key="bottom_usl")
                
                if LSL >= USL: 
                    st.error("❌ LSL must be strictly smaller than USL")
                else:
                    fig, ax = plt.subplots(figsize=(10, 4))
                    ax.hist(data, bins=20, density=True, alpha=0.7, edgecolor="black", label="Thickness Histogram")
                    x = np.linspace(mean - 5 * std, mean + 5 * std, 1000)
                    ax.plot(x, normal_pdf(x, mean, std), linewidth=3, label="Normal Distribution")
                    ax.axvline(mean, linestyle="--", linewidth=2, color="red", label=f"Mean = {mean:.2f}")
                    ax.axvline(LSL, linestyle="--", linewidth=2, color="green", label=f"LSL = {LSL:.2f}")
                    ax.axvline(USL, linestyle="--", linewidth=2, color="green", label=f"USL = {USL:.2f}")
                    ax.axvspan(LSL, USL, alpha=0.15, label="Spec Zone")
                    ax.set_xlim(mean - 5 * std, mean + 5 * std)
                    ax.set_xlabel("Average Thickness")
                    ax.set_ylabel("Density")
                    ax.grid(True, linestyle="--", alpha=0.4)
                    ax.legend()
                    st.pyplot(fig)
                    plt.close(fig)

            with st.expander("📋 Coil Summary Data"):
                st.dataframe(df_plot[[coil_col, thickness_col, dE_col, dL_col, da_col, db_col, time_col]].sort_values(by=dE_col, ascending=False), use_container_width=True)

            st.markdown("---")
            st.header("🔬 PHASE II – THICKNESS CORRELATION")
            if control_time is None:
                st.warning("⚠ Control batch not defined. Phase II cannot be determined.")
            else:
                df_p2 = df[(df["Time"] >= control_time) & (df["塗料編號"] == color)].copy()
                if df_p2.empty:
                    st.warning("⚠ No Phase II data after filtering.")
                else:
                    COLOR_FACTORS = {"ΔL": ["入料檢測 ΔL 正面", "Average value ΔL 正面"], "Δa": ["入料檢測 Δa 正面", "Average value Δa 正面"], "Δb": ["入料檢測 Δb 正面", "Average value Δb 正面"], "ΔE": ["Average value ΔE 正面"]}
                    available_factors = {k: c for k, cols in COLOR_FACTORS.items() for c in cols if c in df_p2.columns}
                    
                    if not available_factors:
                        st.warning("⚠ No color factor columns found in dataset.")
                    else:
                        factor_label = st.selectbox("🎯 Select Color Factor", list(available_factors.keys()), index=0, format_func=delta_label, key="bottom_color_factor")
                        factor_col = available_factors[factor_label]
                        coil_df = df_p2.groupby("Coil No.", as_index=False).agg({"Avergage Thickness": "mean", factor_col: "mean", "製造批號": "min"}).dropna()
                        
                        if coil_df.empty:
                            st.warning("⚠ No valid coil-level data.")
                        else:
                            lcl, ucl = safe_get_limit(color, "LINE", factor_label)
                            ooc_mask = (coil_df[factor_col] < lcl) | (coil_df[factor_col] > ucl) if lcl is not None and ucl is not None else np.zeros(len(coil_df), dtype=bool)
                            x, y = coil_df["Avergage Thickness"].values, coil_df[factor_col].values
                            r2 = None
                            
                            if len(x) >= 2:
                                slope, intercept = np.polyfit(x, y, 1)
                                r2 = 1 - np.sum((y - (slope * x + intercept)) ** 2) / np.sum((y - np.mean(y)) ** 2) if np.sum((y - np.mean(y)) ** 2) != 0 else 0
                            
                            fig, ax = plt.subplots(figsize=(9, 6))
                            ax.scatter(coil_df[~ooc_mask]["Avergage Thickness"], coil_df[~ooc_mask][factor_col], alpha=0.7, label="Normal Coil")
                            if ooc_mask.any(): 
                                ax.scatter(coil_df[ooc_mask]["Avergage Thickness"], coil_df[ooc_mask][factor_col], color="red", s=80, label="OOC Coil")
                            if r2 is not None: 
                                ax.plot(np.linspace(x.min(), x.max(), 100), slope * np.linspace(x.min(), x.max(), 100) + intercept, linestyle="--", linewidth=2, label=f"Regression Line (R² = {r2:.3f})")
                            
                            ax.set_title(f"Phase II – Per Coil Analysis\nThickness vs {delta_label(factor_label)}" + (f" | r = {coil_df['Avergage Thickness'].corr(coil_df[factor_col]):.3f}, R² = {r2:.3f}" if r2 is not None else ""))
                            ax.set_xlabel("Average Thickness (per Coil)")
                            ax.set_ylabel(delta_label(factor_label))
                            ax.legend()
                            ax.grid(True, linestyle="--", alpha=0.4)
                            st.pyplot(fig)
                            plt.close(fig)

                            st.markdown("### 🧠 Interpretation")
                            st.markdown("""
                            **📊 Correlation Levels Reference:**
                            
                            | Level | Correlation Coefficient (\\|R\\|) | Coefficient of Determination (R²) | Interpretation |
                            | :--- | :--- | :--- | :--- |
                            | 🔴 **Strong** | \\|R\\| ≥ 0.77 | R² ≥ 0.60 | Thickness strongly affects and explains most of the color variation. |
                            | 🟠 **Moderate**| 0.55 ≤ \\|R\\| < 0.77 | 0.30 ≤ R² < 0.60 | Thickness partially contributes to color drift. |
                            | 🟢 **Weak/Low** | \\|R\\| < 0.55 | R² < 0.30 | Thickness is unlikely the main driver of color variation. |
                            """)

                            if r2 is not None:
                                if r2 >= 0.6: 
                                    st.error("🔴 Thickness strongly explains color variation (High R²)")
                                elif r2 >= 0.3: 
                                    st.warning("🟠 Thickness may contribute to color drift (Moderate R²)")
                                else: 
                                    st.success("🟢 Thickness unlikely main driver (Low R²)")
                            else: 
                                st.info("ℹ Not enough data for regression analysis.")

                            if ooc_mask.any() and (~ooc_mask).any():
                                thick_col = "Avergage Thickness"
                                ooc_thick = coil_df[ooc_mask][thick_col]
                                norm_thick = coil_df[~ooc_mask][thick_col]
                                
                                mean_ooc = ooc_thick.mean()
                                q1_norm = norm_thick.quantile(0.25)
                                q3_norm = norm_thick.quantile(0.75)
                                
                                if mean_ooc > q3_norm:
                                    st.warning(f"🚨 **Automated Risk Alert:** OOC coils are noticeably clustered in the **HIGH** thickness zone (Mean OOC Thickness: {mean_ooc:.2f} > Normal Q3: {q3_norm:.2f}). Consider tightening the **Upper Specification Limit (USL)** for thickness to mitigate color drift risks.")
                                elif mean_ooc < q1_norm:
                                    st.warning(f"🚨 **Automated Risk Alert:** OOC coils are noticeably clustered in the **LOW** thickness zone (Mean OOC Thickness: {mean_ooc:.2f} < Normal Q1: {q1_norm:.2f}). Consider tightening the **Lower Specification Limit (LSL)** for thickness to mitigate color drift risks.")
                            
                            with st.expander("📋 Phase II – Coil Level Data"): 
                                st.dataframe(coil_df.sort_values("製造批號"), use_container_width=True)

# =========================================================
# VIEW 2: LIMIT STATUS SUMMARY
# =========================================================
elif app_mode == "📋 Limit Status Summary":
    st.title("📋 Limit Status Summary")
    st.markdown("Global overview of all color codes, identifying stable processes and those requiring control limit recalculation based on SPC rules.")

    st.markdown("### ⚙️ Alert Settings")
    col_set1, col_set2 = st.columns(2)
    with col_set1:
        c_th = st.number_input("Consecutive OOC threshold (Rule 4):", min_value=1, max_value=10, value=2, step=1)
    with col_set2:
        t_th = st.number_input("Total (non-consecutive) OOC threshold:", min_value=1, max_value=20, value=5, step=1)

    all_colors = sorted(df_raw["塗料編號"].dropna().unique())
    summary_data = []

    def max_consecutive_true(s):
        if s.empty: return 0
        return (s * (s.groupby((s != s.shift()).cumsum()).cumcount() + 1)).max()

    for c in all_colors:
        c_clean = str(c).strip()
        mask = limit_df["Color_code"].astype(str).str.strip() == c_clean
        row = limit_df[mask]
        
        status = "❌ No"
        if not row.empty:
            limit_cols = [col for col in row.columns if "LCL" in col or "UCL" in col]
            if not row[limit_cols].isna().all().all():
                status = "✅ Yes"

        df_c = df_raw[df_raw["塗料編號"] == c].sort_values("Time")
        total_batches = df_c["製造批號"].nunique()
        can_calc_initial = "✅ Yes" if total_batches >= 3 else "❌ No"
        
        cb = get_control_batch(c)
        cb_code = get_control_batch_code(df_c, cb)
        
        c_time = df_c[df_c["製造批號"] == cb_code]["Time"].min() if cb_code else None
        
        recalc_status = "❌ Not Enough Data"
        phase2_batches = 0
        
        if c_time is not None:
            df_p2 = df_c[df_c["Time"] >= c_time]
            phase2_batches = df_p2["製造批號"].nunique()
            
            if phase2_batches >= 3:
                if status == "✅ Yes":
                    max_consec_any_chart, max_total_any_chart = 0, 0
                    spc_p2 = calculate_batch_averages(df_p2[SPC_INPUT_COLUMNS])
                    
                    for f in ["ΔL", "Δa", "Δb"]:
                        for source in ["line", "lab"]:
                            lcl, ucl = safe_get_limit(c, source, f)
                            if lcl is not None and ucl is not None:
                                vals = spc_p2[f][source]["value"]
                                mask_val = (vals < lcl) | (vals > ucl)
                                consec, total_ooc = max_consecutive_true(mask_val), mask_val.sum()
                                if consec > max_consec_any_chart: max_consec_any_chart = consec
                                if total_ooc > max_total_any_chart: max_total_any_chart = total_ooc
                    
                    if max_consec_any_chart >= c_th: 
                        recalc_status = f"⚠️ Propose Recalc ({max_consec_any_chart} Consec. OOCs)"
                    elif max_total_any_chart >= t_th: 
                        recalc_status = f"⚠️ Propose Recalc ({max_total_any_chart} Total OOCs)"
                    else: 
                        recalc_status = f"✅ Stable (Max Consec: {max_consec_any_chart}, Total: {max_total_any_chart})"
                else: 
                    recalc_status = "❌ Missing Current Limits"
            
        summary_data.append({
            "Color Code": c, 
            "Total Batches": total_batches, 
            "Phase II Batches": phase2_batches,
            "Current Limits": status, 
            "Ready for Calc (Total)": can_calc_initial, 
            "Recommend Recalc (Phase II)": recalc_status
        })

    summary_df = pd.DataFrame(summary_data)
    total_c = len(summary_df)
    has_limit_c = len(summary_df[summary_df["Current Limits"] == "✅ Yes"])
    ready_initial_c = len(summary_df[summary_df["Ready for Calc (Total)"] == "✅ Yes"])
    needs_recalc_c = sum(1 for d in summary_data if "⚠️ Propose Recalc" in d["Recommend Recalc (Phase II)"])

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Colors", total_c)
    col2.metric("Colors Configured", has_limit_c)
    col3.metric("Ready to Calc (Initial)", ready_initial_c)
    col4.metric("Needs Recalculation", needs_recalc_c, delta="Process Shift Alert", delta_color="inverse")

    st.markdown("---")
    st.markdown("### 📊 Comprehensive Status Table")
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### 🚨 Action Required: Missing Limits")
    st.markdown("The following colors do not have configured control limits but have enough data (≥ 3 batches). Please navigate to **🎛️ Control Limit Calculator** to configure them.")
    
    pending_colors = summary_df[
        (summary_df["Current Limits"] == "❌ No") & 
        (summary_df["Ready for Calc (Total)"] == "✅ Yes")
    ]
    
    if not pending_colors.empty:
        st.warning(f"Found **{len(pending_colors)}** color(s) waiting for limit calculation:")
        col_table, col_empty = st.columns([1, 2])
        with col_table:
            st.dataframe(
                pending_colors[["Color Code", "Total Batches"]], 
                hide_index=True,
                use_container_width=True
            )
    else:
        st.success("🎉 All colors with sufficient data already have their control limits configured!")

# =========================================================
# VIEW 3: CONTROL LIMIT CALCULATOR
# GOAL: Define incoming LAB limits that help control LINE output color trend
# =========================================================
elif app_mode == "🎛️ Control Limit Calculator":

    # Task 3 specific clean theme (remove yellow background in this view only)
    st.markdown(
        """
        <style>
        .stApp {
            background: #f6f8fb !important;
            animation: none !important;
        }
        div[data-testid="stNumberInput"] input {
            background-color: #ffffff !important;
            border: 1px solid #d7dee8 !important;
        }
        div[data-testid="stDataFrame"] {
            border: 1px solid #e3e8ef;
            border-radius: 10px;
            overflow: hidden;
            background: white;
        }
        .task3-note {
            padding: 12px 16px;
            background: #eef6ff;
            border: 1px solid #cfe3f7;
            border-radius: 10px;
            margin-bottom: 16px;
            color: #163a5f;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.title("🎛️ Incoming LAB Control Limit → LINE Output Trend")
    st.markdown(
        """
        <div class="task3-note">
        <b>Analysis objective:</b> define suitable <b>incoming LAB control limits</b> for ΔL* / Δa* / Δb*,
        while using the corresponding <b>LINE output</b> from the same production batch to verify the downstream color trend.
        LAB and LINE are therefore analyzed together; no source selection is required.
        </div>
        """,
        unsafe_allow_html=True
    )

    # Per-color specification values. Session state is keyed by color so standards
    # cannot silently carry over between selected color codes.
    st.markdown("### 🎯 色號 LAB / LINE ΔE* 規格設定")
    st.caption("請依當前色號的正式檢驗規格輸入。LAB 與 LINE 分開判定，並同步套用至分析及 Word/Excel 報告。")
    _spec_col_lab, _spec_col_line = st.columns(2)
    with _spec_col_lab:
        lab_de_target = st.number_input("LAB ΔE* Specification (Max)｜進料規格", min_value=0.001,
            value=0.80 if str(color).strip().upper() == "PK2E1230L" else 0.50,
            step=0.01, format="%.2f", key=f"task3_lab_spec_{color}")
    with _spec_col_line:
        line_de_target = st.number_input("LINE ΔE* Specification (Max)｜產出規格", min_value=0.001,
            value=1.00, step=0.01, format="%.2f", key=f"task3_line_spec_{color}")
    st.info(f"目前色號：{color} ｜ LAB ΔE* ≤ {lab_de_target:.2f} ｜ LINE ΔE* ≤ {line_de_target:.2f}。請確認以上數值為正式規格。")

    factors = ["ΔL", "Δa", "Δb"]

    # ---------------------------------------------------------
    # Task 3 summary table: LAB INPUT statistics only
    # Control limits shown here use fixed x̄ ± 1.5σ, matching the report format.
    # LINE output remains available in the paired analysis and LINE-limit section below,
    # but is intentionally excluded from this summary table.
    # ---------------------------------------------------------
    st.markdown("### 📋 LAB 入料色差統計與 ±1.5σ 管制界限")
    st.caption(
        "僅針對 LAB 入料資料，以各製造批號之平均值計算 Min、Max、平均值 (x̄)、標準差 (σ)，"
        "並以 x̄ ± 1.5σ 計算 LCL / UCL。"
    )

    summary_sigma_k = 1.5
    task3_stats_rows = []

    for _f in factors:
        _vals = pd.to_numeric(
            spc_data[_f]["lab"]["value"], errors="coerce"
        ).dropna()

        if len(_vals) > 0:
            _min = _vals.min()
            _max = _vals.max()
            _mean = _vals.mean()
            _std = _vals.std()
            if pd.isna(_std):
                _std = 0.0
            _lcl = _mean - summary_sigma_k * _std
            _ucl = _mean + summary_sigma_k * _std
        else:
            _min = _max = _mean = _std = _lcl = _ucl = np.nan

        task3_stats_rows.append({
            "入料色差走向": f"{delta_label(_f)}入料",
            "Min": _min,
            "Max": _max,
            "平均值 (x̄)": _mean,
            "標準差 (σ)": _std,
            "LCL (-1.5σ)": _lcl,
            "UCL (+1.5σ)": _ucl,
        })

    task3_stats_df = pd.DataFrame(task3_stats_rows)
    st.dataframe(
        task3_stats_df.style.format({
            "Min": "{:.2f}",
            "Max": "{:.2f}",
            "平均值 (x̄)": "{:.2f}",
            "標準差 (σ)": "{:.2f}",
            "LCL (-1.5σ)": "{:.2f}",
            "UCL (+1.5σ)": "{:.2f}",
        }, na_rep="-"),
        hide_index=True,
        use_container_width=True,
    )

    # Download the Task 3 statistical summary table for the selected color
    _safe_color_name = re.sub(r'[^A-Za-z0-9._-]+', '_', str(color)).strip('_') or 'Color'

    _xlsx_buf = io.BytesIO()
    with pd.ExcelWriter(_xlsx_buf, engine="openpyxl") as writer:
        task3_stats_df.to_excel(writer, index=False, sheet_name="LAB_Input_Stats")
        # Display two decimals without rounding the underlying numeric data.
        for sheet in writer.book.worksheets:
            for row in sheet.iter_rows(min_row=2):
                for cell in row:
                    if isinstance(cell.value, (float, np.floating)) and not isinstance(cell.value, bool):
                        cell.number_format = '0.00' 
        ws = writer.book["LAB_Input_Stats"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col_cells in ws.columns:
            max_len = 0
            col_letter = col_cells[0].column_letter
            for cell in col_cells:
                cell_val = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(cell_val))
            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 28)
    _xlsx_buf.seek(0)

    _csv_buf = task3_stats_df.to_csv(index=False, float_format="%.2f").encode("utf-8-sig")

    _dl_col1, _dl_col2, _dl_spacer = st.columns([1, 1, 2])
    with _dl_col1:
        st.download_button(
            label="📥 Download Table (Excel)",
            data=_xlsx_buf,
            file_name=f"{_safe_color_name}_LAB_Input_Stats_1.5Sigma.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"task3_stats_xlsx_{_safe_color_name}",
            use_container_width=True,
        )
    with _dl_col2:
        st.download_button(
            label="📥 Download Table (CSV)",
            data=_csv_buf,
            file_name=f"{_safe_color_name}_LAB_Input_Stats_1.5Sigma.csv",
            mime="text/csv",
            key=f"task3_stats_csv_{_safe_color_name}",
            use_container_width=True,
        )

    st.markdown("---")

    # ---------------------------------------------------------
    # Build paired LAB + LINE batch-level dataset
    # ---------------------------------------------------------
    paired_all = None
    for _f in factors:
        _lab = spc_data[_f]["lab"][["製造批號", "Time", "value"]].copy().rename(
            columns={"Time": f"Time_LAB_{_f}", "value": f"LAB_{_f}"}
        )
        _line = spc_data[_f]["line"][["製造批號", "Time", "value"]].copy().rename(
            columns={"Time": f"Time_LINE_{_f}", "value": f"LINE_{_f}"}
        )
        _pair = _lab.merge(_line, on="製造批號", how="inner")
        paired_all = _pair if paired_all is None else paired_all.merge(_pair, on="製造批號", how="inner")

    st.markdown("### 📋 Paired Batch Data — LAB Input vs LINE Output")

    if paired_all is None or paired_all.empty:
        st.warning("No paired LAB + LINE batch data are available after the current filters.")
        paired_all = pd.DataFrame()
    else:
        _time_cols = [c for c in paired_all.columns if c.startswith("Time_")]
        for _c in _time_cols:
            paired_all[_c] = pd.to_datetime(paired_all[_c], errors="coerce")
        paired_all["Time"] = paired_all[_time_cols].min(axis=1)
        paired_all = paired_all.sort_values(["Time", "製造批號"]).reset_index(drop=True)

        display_paired = paired_all[[
            "製造批號", "Time",
            "LAB_ΔL", "LINE_ΔL",
            "LAB_Δa", "LINE_Δa",
            "LAB_Δb", "LINE_Δb"
        ]].rename(columns={
            "製造批號": "Batch No.",
            "LAB_ΔL": "LAB_ΔL*", "LINE_ΔL": "LINE_ΔL*",
            "LAB_Δa": "LAB_Δa*", "LINE_Δa": "LINE_Δa*",
            "LAB_Δb": "LAB_Δb*", "LINE_Δb": "LINE_Δb*"
        })

        st.caption(
            f"Color Code: {color} | Paired batches: {len(display_paired)} | "
            "Each row contains incoming LAB and production LINE results from the same batch."
        )
        st.dataframe(
            display_paired.style.format({
                "LAB_ΔL*": "{:.2f}", "LINE_ΔL*": "{:.2f}",
                "LAB_Δa*": "{:.2f}", "LINE_Δa*": "{:.2f}",
                "LAB_Δb*": "{:.2f}", "LINE_Δb*": "{:.2f}"
            }, na_rep=""),
            hide_index=True,
            use_container_width=True
        )

    st.markdown("---")

    calc_res = {}
    chart_buffers = {}  # Existing full-size charts used by the app/download
    report_chart_buffers = {}  # Dedicated high-resolution charts for the one-page Word layout

    def build_task3_report_chart(pair_df, factor):
        """Clean, high-resolution trend chart for the Word report.
        Important: keep all Matplotlib text in the original English/CIELAB notation
        so Streamlit Cloud does not depend on a Chinese CJK font package.
        """
        fig_r, ax_r = plt.subplots(figsize=(6.2, 3.45))
        fig_r.patch.set_facecolor("white")
        ax_r.set_facecolor("white")

        x_r = np.arange(len(pair_df))
        ax_r.plot(
            x_r, pair_df["LINE"].values,
            marker="D", linestyle="-", linewidth=2.0, markersize=6.0,
            color="#ed7d31", label="LINE Output"
        )
        ax_r.plot(
            x_r, pair_df["LAB"].values,
            marker="s", linestyle="-", linewidth=2.0, markersize=6.0,
            color="#5b9bd5", label="LAB Input"
        )

        ax_r.set_xticks(x_r)
        ax_r.set_xticklabels([str(i + 1) for i in x_r], fontsize=10)
        ax_r.set_xlabel("Batch sequence", fontsize=11, fontweight="bold", labelpad=7)
        ax_r.set_ylabel(delta_label(factor), fontsize=11, fontweight="bold")
        ax_r.set_title(
            f"{delta_label(factor)} LAB Input vs LINE Output",
            fontsize=13, fontweight="bold", pad=9
        )
        ax_r.tick_params(axis="y", labelsize=10)
        ax_r.grid(axis="y", color="#c9c9c9", linestyle="-", linewidth=0.8)
        ax_r.grid(axis="x", visible=False)

        for spine in ["top", "right", "left", "bottom"]:
            ax_r.spines[spine].set_visible(True)
            ax_r.spines[spine].set_color("#8a8a8a")
            ax_r.spines[spine].set_linewidth(0.9)

        ax_r.legend(
            loc="upper center", bbox_to_anchor=(0.5, -0.20),
            ncol=2, frameon=False, fontsize=10,
            handlelength=2.0, columnspacing=1.8
        )
        fig_r.subplots_adjust(left=0.13, right=0.97, top=0.84, bottom=0.31)
        return fig_r
    dE_std_sq = 0.0
    dE_iqr_sq = 0.0

    for f in factors:
        st.markdown(f"## 📊 {delta_label(f)}: LAB Input vs LINE Output")

        col_sig, col_iqr = st.columns(2)
        with col_sig:
            sig = st.number_input(
                f"Sigma (K) for incoming LAB {delta_label(f)}",
                value=3.0,
                step=0.1,
                key=f"paired_sig_{f}"
            )
        with col_iqr:
            iqr_k = st.number_input(
                f"IQR sensitivity for incoming LAB {delta_label(f)}",
                value=1.5,
                step=0.1,
                key=f"paired_iqr_{f}"
            )

        if paired_all.empty:
            st.warning(f"Not enough paired LAB + LINE data for {f}.")
            st.markdown("---")
            continue

        pair_f = paired_all[["製造批號", "Time", f"LAB_{f}", f"LINE_{f}"]].dropna().copy()
        pair_f = pair_f.rename(columns={f"LAB_{f}": "LAB", f"LINE_{f}": "LINE"})

        if len(pair_f) < 3:
            st.warning(f"Not enough paired data for {f} (minimum 3 batches).")
            st.markdown("---")
            continue

        # Existing sheet limits. LAB = incoming control target, LINE = output validation target.
        lab_lcl_sheet, lab_ucl_sheet = safe_get_limit(color, "LAB", f)
        line_lcl_sheet, line_ucl_sheet = safe_get_limit(color, "LINE", f)

        # Use LINE in-control batches to learn which incoming LAB values are associated
        # with acceptable downstream output. If LINE limits are unavailable, use all pairs.
        if line_lcl_sheet is not None and line_ucl_sheet is not None:
            pair_f["LINE_Status"] = np.where(
                (pair_f["LINE"] >= line_lcl_sheet) & (pair_f["LINE"] <= line_ucl_sheet),
                "In Control",
                "OOC"
            )
            safe_lab = pair_f.loc[pair_f["LINE_Status"] == "In Control", "LAB"].dropna()
            filter_note = (
                f"LAB candidate limits are calculated from paired batches whose LINE output is within "
                f"the current LINE range [{line_lcl_sheet:.2f}, {line_ucl_sheet:.2f}]."
            )
        else:
            pair_f["LINE_Status"] = "No LINE limit"
            safe_lab = pair_f["LAB"].dropna()
            filter_note = (
                "Current LINE control limits are not configured, so all paired batches are used for the LAB calculation."
            )

        if len(safe_lab) < 3:
            # Avoid unstable statistics caused by an excessively small filtered subset.
            safe_lab = pair_f["LAB"].dropna()
            filter_note += " Fewer than 3 LINE-in-control pairs were available, so all paired LAB values were used."

        m = safe_lab.mean()
        s = safe_lab.std()
        q1 = safe_lab.quantile(0.25)
        q3 = safe_lab.quantile(0.75)
        median = safe_lab.median()

        std_lcl = m - sig * s
        std_ucl = m + sig * s
        iqr_value = q3 - q1
        iqr_lcl = q1 - iqr_k * iqr_value
        iqr_ucl = q3 + iqr_k * iqr_value

        # -----------------------------------------------------
        # IQR manual calculation table (step-by-step)
        # -----------------------------------------------------
        st.markdown(f"### 🧮 IQR 手算步驟 — {delta_label(f)} 入料")
        st.caption(
            "依標準 IQR 公式計算：IQR = Q3 − Q1；LL = Q1 − k×IQR；UL = Q3 + k×IQR。"
        )

        _iqr_manual_df = pd.DataFrame([
            {
                "計算項目": "Q1 (25%)",
                f"{delta_label(f)} 入料": f"{q1:.2f}"
            },
            {
                "計算項目": "Q3 (75%)",
                f"{delta_label(f)} 入料": f"{q3:.2f}"
            },
            {
                "計算項目": "IQR (Q3 − Q1)",
                f"{delta_label(f)} 入料": f"{q3:.2f} − {q1:.2f} = {iqr_value:.2f}"
            },
            {
                "計算項目": f"控制範圍 LL (Q1 − {iqr_k:.2f}×IQR)",
                f"{delta_label(f)} 入料": f"{q1:.2f} − ({iqr_k:.2f} × {iqr_value:.2f}) = {iqr_lcl:.2f}"
            },
            {
                "計算項目": f"控制範圍 UL (Q3 + {iqr_k:.2f}×IQR)",
                f"{delta_label(f)} 入料": f"{q3:.2f} + ({iqr_k:.2f} × {iqr_value:.2f}) = {iqr_ucl:.2f}"
            },
        ])

        st.dataframe(
            _iqr_manual_df,
            hide_index=True,
            use_container_width=True,
        )

        st.info(
            f"IQR結果：{delta_label(f)} 入料 LL = {iqr_lcl:.2f}，UL = {iqr_ucl:.2f} （k = {iqr_k:.2f}）"
        )

        rec_lcl = max(std_lcl, iqr_lcl)
        rec_ucl = min(std_ucl, iqr_ucl)
        has_overlap = rec_lcl <= rec_ucl

        r = pair_f["LAB"].corr(pair_f["LINE"]) if len(pair_f) >= 2 else np.nan
        r2 = r ** 2 if pd.notna(r) else np.nan

        calc_res[f] = {
            "pair": pair_f,
            "safe_lab": safe_lab,
            "m": m,
            "s": s,
            "median": median,
            "sig": sig,
            "iqr_k": iqr_k,
            "lab_lcl_sheet": lab_lcl_sheet,
            "lab_ucl_sheet": lab_ucl_sheet,
            "line_lcl_sheet": line_lcl_sheet,
            "line_ucl_sheet": line_ucl_sheet,
            "std_lcl": std_lcl,
            "std_ucl": std_ucl,
            "iqr_lcl": iqr_lcl,
            "iqr_ucl": iqr_ucl,
            "rec_lcl": rec_lcl if has_overlap else None,
            "rec_ucl": rec_ucl if has_overlap else None,
            "r": r,
            "r2": r2,
            "filter_note": filter_note,
        }

        dE_std_sq += max(abs(std_lcl), abs(std_ucl)) ** 2
        dE_iqr_sq += max(abs(iqr_lcl), abs(iqr_ucl)) ** 2

        # -----------------------------------------------------
        # Combined LAB + LINE chart
        # -----------------------------------------------------
        col_chart, col_table = st.columns([2.15, 1])

        with col_chart:
            fig, ax = plt.subplots(figsize=(10.5, 5.0))

            # Excel-like framed chart area
            fig.patch.set_facecolor("white")
            fig.patch.set_edgecolor("#8a8a8a")
            fig.patch.set_linewidth(1.2)
            ax.set_facecolor("white")

            ordered_batches = pair_f["製造批號"].tolist()
            x = np.arange(len(ordered_batches))

            ax.plot(
                x, pair_f["LINE"].values,
                marker="D", linestyle="-", linewidth=1.6, markersize=5.5,
                color="#ed7d31", markerfacecolor="#ed7d31", markeredgewidth=0,
                label="LINE Output"
            )
            ax.plot(
                x, pair_f["LAB"].values,
                marker="s", linestyle="-", linewidth=1.6, markersize=5.5,
                color="#5b9bd5", markerfacecolor="#5b9bd5", markeredgewidth=0,
                label="LAB Input"
            )

            # Current LINE limits show whether the output remains controlled.
            if line_lcl_sheet is not None and line_ucl_sheet is not None:
                ax.axhline(line_lcl_sheet, color="#dc2626", linestyle="--", linewidth=1.2, label="LINE LCL/UCL")
                ax.axhline(line_ucl_sheet, color="#dc2626", linestyle="--", linewidth=1.2)

                ooc_mask = (pair_f["LINE"] < line_lcl_sheet) | (pair_f["LINE"] > line_ucl_sheet)
                if ooc_mask.any():
                    ax.scatter(
                        x[ooc_mask.values], pair_f.loc[ooc_mask, "LINE"],
                        s=75, color="#dc2626", marker="o", zorder=6,
                        label="LINE OOC"
                    )

            # Candidate incoming LAB range is the main output of this task.
            if has_overlap:
                ax.axhspan(rec_lcl, rec_ucl, color="#dbeafe", alpha=0.35, label="Recommended LAB")
                ax.axhline(rec_lcl, color="#2563eb", linestyle=":", linewidth=1.4)
                ax.axhline(rec_ucl, color="#2563eb", linestyle=":", linewidth=1.4)

            if lab_lcl_sheet is not None and lab_ucl_sheet is not None:
                ax.axhline(lab_lcl_sheet, color="#7f7f7f", linestyle="-.", linewidth=1.0, label="LAB LCL/UCL")
                ax.axhline(lab_ucl_sheet, color="#7f7f7f", linestyle="-.", linewidth=1.0)

            ax.set_xticks(x)
            ax.set_xticklabels(ordered_batches, rotation=35, ha="right", fontsize=9)
            ax.set_ylabel(delta_label(f))
            ax.set_xlabel("Batch No.", fontsize=10, labelpad=10)
            ax.set_title(f"Analysis: {delta_label(f)} (LAB vs LINE)", fontsize=14, fontweight="bold", pad=12)
            ax.grid(axis="y", color="#bfbfbf", linestyle="-", linewidth=0.8)
            ax.grid(axis="x", visible=False)

            # Full border around plot area like the user's example
            for spine in ["top", "right", "left", "bottom"]:
                ax.spines[spine].set_visible(True)
                ax.spines[spine].set_color("#7f7f7f")
                ax.spines[spine].set_linewidth(0.9)

            ax.tick_params(axis='both', colors='black', labelsize=9)

            # Deduplicate legend labels and place the legend just below the chart
            handles, labels = ax.get_legend_handles_labels()
            uniq_handles, uniq_labels = [], []
            for h, lab in zip(handles, labels):
                if lab not in uniq_labels:
                    uniq_handles.append(h)
                    uniq_labels.append(lab)
            # Put legend on its own row at the very bottom of the figure.
            # This prevents overlap with rotated Batch No. tick labels and X-axis title.
            fig.legend(
                uniq_handles, uniq_labels,
                loc="lower center", bbox_to_anchor=(0.5, 0.015),
                ncol=min(3, len(uniq_labels)), frameon=False, fontsize=9,
                handlelength=2.0, columnspacing=1.6, handletextpad=0.5
            )

            # Reserve three clearly separated vertical zones:
            # plot -> Batch No. labels/X title -> legend
            fig.subplots_adjust(bottom=0.36, right=0.96)
            st.pyplot(fig)

            # Save exactly the displayed chart for PNG download and Word export.
            chart_png = io.BytesIO()
            fig.savefig(chart_png, format="png", dpi=320, bbox_inches="tight", facecolor="white")
            chart_png.seek(0)
            chart_bytes = chart_png.getvalue()
            chart_buffers[f] = chart_bytes

            # Generate large 500-DPI Word images only when export is requested.
            # Keeping them out of routine Streamlit reruns avoids large RAM spikes.

            factor_filename = {"ΔL": "DeltaL", "Δa": "Deltaa", "Δb": "Deltab"}.get(f, str(f))
            st.download_button(
                label=f"📥 Download {f} Chart (PNG)",
                data=chart_bytes,
                file_name=f"Task3_{color}_{factor_filename}_LAB_vs_LINE.png",
                mime="image/png",
                key=f"task3_png_{color}_{factor_filename}"
            )
            plt.close(fig)

        with col_table:
            current_lab = (
                f"{lab_lcl_sheet:.2f} ~ {lab_ucl_sheet:.2f}"
                if lab_lcl_sheet is not None and lab_ucl_sheet is not None else "Not set"
            )
            current_line = (
                f"{line_lcl_sheet:.2f} ~ {line_ucl_sheet:.2f}"
                if line_lcl_sheet is not None and line_ucl_sheet is not None else "Not set"
            )
            rec_txt = f"{rec_lcl:.2f} ~ {rec_ucl:.2f}" if has_overlap else "No overlap"

            result_table = pd.DataFrame([
                {"Item": "Current LAB limit", "Range / Value": current_lab},
                {"Item": "Current LINE limit", "Range / Value": current_line},
                {"Item": f"Standard LAB ({sig:.2f}σ)", "Range / Value": f"{std_lcl:.2f} ~ {std_ucl:.2f}"},
                {"Item": f"IQR LAB (k={iqr_k:.2f})", "Range / Value": f"{iqr_lcl:.2f} ~ {iqr_ucl:.2f}"},
                {"Item": "Recommended LAB", "Range / Value": rec_txt},
                {"Item": "Paired batches", "Range / Value": str(len(pair_f))},
                {"Item": "LAB data used", "Range / Value": str(len(safe_lab))},
                {"Item": "LAB→LINE r", "Range / Value": f"{r:.2f}" if pd.notna(r) else "N/A"},
                {"Item": "LAB→LINE R²", "Range / Value": f"{r2:.2f}" if pd.notna(r2) else "N/A"},
            ])
            st.dataframe(result_table, hide_index=True, use_container_width=True)
            st.caption(filter_note)

            if pd.notna(r2):
                if r2 >= 0.60:
                    st.error(f"🔴 Strong LAB→LINE relationship (R²={r2:.2f})")
                elif r2 >= 0.30:
                    st.warning(f"🟠 Moderate LAB→LINE relationship (R²={r2:.2f})")
                else:
                    st.success(f"🟢 Weak LAB→LINE relationship (R²={r2:.2f})")

        st.markdown("---")

    # ---------------------------------------------------------
    # Three-factor incoming LAB control-limit summary
    # ---------------------------------------------------------
    if len(calc_res) == 3:
        st.markdown("### 📋 Recommended Incoming LAB Control Limits")

        summary_rows = []
        recommendation_valid = True
        dE_rec_sq = 0.0

        for f in factors:
            res = calc_res[f]
            has_rec = res["rec_lcl"] is not None and res["rec_ucl"] is not None
            if has_rec:
                dE_rec_sq += max(abs(res["rec_lcl"]), abs(res["rec_ucl"])) ** 2
            else:
                recommendation_valid = False

            summary_rows.append({
                "Factor": f,
                "Paired n": len(res["pair"]),
                "LAB n used": len(res["safe_lab"]),
                "LAB Mean": round(res["m"], 2),
                "LAB σ": round(res["s"], 2),
                "Standard LCL": round(res["std_lcl"], 2),
                "Standard UCL": round(res["std_ucl"], 2),
                "IQR LCL": round(res["iqr_lcl"], 2),
                "IQR UCL": round(res["iqr_ucl"], 2),
                "Recommended LAB LCL": round(res["rec_lcl"], 2) if has_rec else None,
                "Recommended LAB UCL": round(res["rec_ucl"], 2) if has_rec else None,
                "LAB→LINE r": round(res["r"], 2) if pd.notna(res["r"]) else None,
                "LAB→LINE R²": round(res["r2"], 2) if pd.notna(res["r2"]) else None,
            })

        incoming_summary = pd.DataFrame(summary_rows)
        # The TWO Recommended LAB columns in THIS summary are editable.
        # All statistics remain read-only reference values.
        _manual_key = f"task3_manual_start_limits_{str(color)}"
        _reset_key = f"task3_manual_editor_version_{str(color)}"
        _recommended_columns = ["Recommended LAB LCL", "Recommended LAB UCL"]
        _defaults = incoming_summary[_recommended_columns].copy()
        for _idx, _factor in enumerate(factors):
            for _kind, _col in (("lcl", _recommended_columns[0]), ("ucl", _recommended_columns[1])):
                if pd.isna(_defaults.at[_idx, _col]):
                    _defaults.at[_idx, _col] = round(float(calc_res[_factor][f"iqr_{_kind}"]), 2)
        # Session-state migration: earlier app versions stored a different
        # DataFrame schema under the same key.  Normalize it before indexing.
        # Always align rows to the current factor order; never trust positional
        # old data from an unrelated color or malformed state.
        def _normalize_saved_limits(saved, defaults, factors):
            result = defaults.copy().reset_index(drop=True)
            if isinstance(saved, pd.DataFrame):
                legacy = saved.copy().reset_index(drop=True)
            elif isinstance(saved, dict):
                try:
                    legacy = pd.DataFrame(saved).reset_index(drop=True)
                except (ValueError, TypeError):
                    return result
            else:
                return result
            # Match old source-column aliases, preserving valid user edits.
            aliases = {
                "Recommended LAB LCL": ["Recommended LAB LCL", "人工確認 LCL", "LCL", "LAB LCL"],
                "Recommended LAB UCL": ["Recommended LAB UCL", "人工確認 UCL", "UCL", "LAB UCL"],
            }
            factor_col = next((c for c in ("Factor", "因子") if c in legacy.columns), None)
            factor_aliases = {"ΔL": "ΔL*", "Δa": "Δa*", "Δb": "Δb*"}
            for dest_col, candidates in aliases.items():
                existing = next((c for c in candidates if c in legacy.columns), None)
                if existing is None:
                    continue
                if factor_col is not None:
                    names = legacy[factor_col].astype(str).str.strip().replace(factor_aliases)
                    for idx, factor in enumerate(factors):
                        positions = names.index[names == factor].tolist()
                        if not positions:
                            positions = names.index[names == factor + "*"].tolist()
                        if positions:
                            value = pd.to_numeric(pd.Series([legacy.at[positions[0], existing]]), errors="coerce").iloc[0]
                            if pd.notna(value) and np.isfinite(value):
                                result.at[idx, dest_col] = float(value)
                elif len(legacy) == len(factors):
                    values = pd.to_numeric(legacy[existing], errors="coerce")
                    for idx, value in enumerate(values):
                        if pd.notna(value) and np.isfinite(value):
                            result.at[idx, dest_col] = float(value)
            return result

        _previous_limits = st.session_state.get(_manual_key)
        _normalized_limits = _normalize_saved_limits(_previous_limits, _defaults, factors)
        # Invalid old LCL/UCL pairs are replaced with fresh calculated defaults.
        for _idx in range(len(factors)):
            _lo = _normalized_limits.at[_idx, _recommended_columns[0]]
            _hi = _normalized_limits.at[_idx, _recommended_columns[1]]
            if not (np.isfinite(_lo) and np.isfinite(_hi) and _lo < _hi):
                _normalized_limits.loc[_idx, _recommended_columns] = _defaults.loc[_idx, _recommended_columns]
        st.session_state[_manual_key] = _normalized_limits
        if _reset_key not in st.session_state:
            st.session_state[_reset_key] = 0

        st.info("✍️ **請直接點選下表最右側的 Recommended LAB LCL / UCL 數字進行修改。**"
                "可填入 IQR、Standard、交集或人工調整值；其餘統計欄位僅供參考。"
                "後續 LINE 風險、Word / Excel 均使用修改後數值。")
        if st.button("↩ 恢復自動建議值", key=f"task3_reset_manual_{color}"):
            st.session_state[_manual_key] = _defaults.copy()
            st.session_state[_reset_key] += 1
            st.rerun()

        _editor_source = incoming_summary.copy()
        _stored = st.session_state[_manual_key]
        for _col in _recommended_columns:
            _editor_source[_col] = pd.to_numeric(_stored[_col], errors="coerce").to_numpy(float)
        _read_only_columns = [c for c in _editor_source.columns if c not in _recommended_columns]
        _edited = st.data_editor(
            _editor_source,
            column_config={
                "Recommended LAB LCL": st.column_config.NumberColumn(
                    "✏️ Recommended LAB LCL", format="%.2f", step=0.01, required=True),
                "Recommended LAB UCL": st.column_config.NumberColumn(
                    "✏️ Recommended LAB UCL", format="%.2f", step=0.01, required=True),
            },
            disabled=_read_only_columns,
            hide_index=True,
            num_rows="fixed",
            use_container_width=True,
            key=f"task3_direct_limits_editor_{color}_{st.session_state[_reset_key]}",
        )
        st.session_state[_manual_key] = _edited[_recommended_columns].copy()
        _start_lcl = pd.to_numeric(_edited[_recommended_columns[0]], errors="coerce").to_numpy(float)
        _start_ucl = pd.to_numeric(_edited[_recommended_columns[1]], errors="coerce").to_numpy(float)
        _manual_valid = (
            len(_start_lcl) == 3 and len(_start_ucl) == 3
            and np.isfinite(_start_lcl).all() and np.isfinite(_start_ucl).all()
            and np.all(_start_lcl < _start_ucl)
        )
        if not _manual_valid:
            st.error("輸入無效：所有 LCL / UCL 必須為數值，且每個因子的 LCL 必須小於 UCL。")
            st.stop()
        _manual_peak = float(np.linalg.norm(np.maximum(np.abs(_start_lcl), np.abs(_start_ucl))))
        st.caption(f"人工確認界限最不利 LAB ΔE* = {_manual_peak:.2f}；本色號規格 ≤ {lab_de_target:.2f}。"
                   "超出規格仍可評估 LINE 風險，但不得視為符合 LAB 規格。")
        _selected_limits_df = pd.DataFrame({
            "因子": ["ΔL*", "Δa*", "Δb*"],
            "人工確認 LCL": _start_lcl,
            "人工確認 UCL": _start_ucl,
        })
        st.caption("原始 IQR / Standard 計算保持不變；本表之✏️欄位為後續計算的人工確認基準。")

        dE_std = math.sqrt(dE_std_sq)
        dE_iqr = math.sqrt(dE_iqr_sq)
        dE_rec = math.sqrt(dE_rec_sq) if recommendation_valid else None

        # Existing project target for incoming LAB color difference.

        st.markdown("### 🎯 Incoming LAB ΔE* Boundary Check")
        c1, c2, c3 = st.columns(3)
        c1.metric("Standard-derived ΔE*", f"{dE_std:.2f}")
        c2.metric("IQR-derived ΔE*", f"{dE_iqr:.2f}")
        c3.metric("Confirmed LAB ΔE*", f"{_manual_peak:.2f}")

        if _manual_peak <= lab_de_target:
            st.success(
                f"✅ Manually confirmed incoming LAB limits produce a derived ΔE* boundary of {_manual_peak:.2f}, "
                f"within the current LAB target ≤ {lab_de_target}."
            )
        else:
            st.warning(
                f"⚠ Manually confirmed incoming LAB limits produce a derived ΔE* boundary of {_manual_peak:.2f}, "
                f"above the current LAB target {lab_de_target}. Further tightening should be reviewed."
            )

        st.markdown("### 💡 How to use this result")
        st.info(
            "The above statistics are exploratory. Before applying limits to incoming LAB, "
            f"also verify LAB ΔE* ≤ {lab_de_target:.2f} and paired historical "
            f"LINE ΔE* ≤ {line_de_target:.2f} using the step-by-step validation below."
        )

    # =========================================================
    # TASK 3: THREE REPORT-READY RESULTS + ONE AUDIT EXPANDER
    # Production target: LINE ΔE* <= color-specific specification.
    # Historical-shift scenarios are a stress test, NOT a guarantee.
    # Existing IQR and Standard Method computations above remain intact.
    # =========================================================
    st.markdown("---")
    st.markdown("## 5. LAB–LINE 產出風險驗證與管制界限最佳化")
    st.caption(f"色號：{color}　｜　LAB ΔE* 規格 ≤ {lab_de_target:.2f}　｜　LINE ΔE* 規格 ≤ {line_de_target:.2f}")
    _factors = ('ΔL', 'Δa', 'Δb')
    _names = ['ΔL*', 'Δa*', 'Δb*']
    _lab_keys = [f'LAB_{f}' for f in _factors]
    _shift_keys = [f'Shift_{f}' for f in _factors]
    _limits = pd.DataFrame()
    _risk = pd.DataFrame()
    _risk_overview = None
    _comparison = None
    _flow = pd.DataFrame()
    _check = pd.DataFrame()
    _summary = pd.DataFrame()
    _shift_stats = pd.DataFrame()
    _shift_range_table = pd.DataFrame()
    _opt = {}
    _case_audit = None
    _stress_initial = pd.DataFrame()
    _stress_proposed = pd.DataFrame()
    _worst_initial_case = None
    _worst_proposed_case = None
    _paired = pd.DataFrame()
    _valid_paired = pd.DataFrame()
    _conclusion = None
    _b_lo = _b_hi = None
    _line_worst = float('nan')
    _new_worst = float('nan')
    _proposed_margin = float('nan')
    _lo = _hi = None
    _valid_iqr = all(
        f in calc_res and np.isfinite(calc_res[f].get('iqr_lcl', np.nan))
        and np.isfinite(calc_res[f].get('iqr_ucl', np.nan))
        and calc_res[f]['iqr_lcl'] < calc_res[f]['iqr_ucl']
        for f in _factors
    )
    if not _valid_iqr:
        st.warning('請先完成三個色差因子的 IQR 管制界限計算。')
    else:
        # Important: the starting bounds are explicitly confirmed by the user,
        # not silently replaced with raw IQR or Standard limits.
        _lo = _start_lcl.copy()
        _hi = _start_ucl.copy()
        _tol = 1e-10
        # LAB acceptance is independent of LINE availability and IQR inclusion.
        _input_columns = ['入料檢測 ΔL 正面', '入料檢測 Δa 正面', '入料檢測 Δb 正面']
        _labs = df.groupby('製造批號', as_index=False)[_input_columns].mean()
        _labs = _labs.rename(columns=dict(zip(_input_columns, _lab_keys)))
        _labs[_lab_keys] = _labs[_lab_keys].apply(pd.to_numeric, errors='coerce')
        _labs = _labs.replace([np.inf, -np.inf], np.nan).dropna(subset=_lab_keys).copy()
        if _labs.empty:
            st.warning('此色號無三因子完整 LAB 批次，無法評估。')
        else:
            _labs['LAB ΔE*'] = np.linalg.norm(_labs[_lab_keys].to_numpy(float), axis=1)
            _labs['LAB 規格合格'] = _labs['LAB ΔE*'] <= lab_de_target + _tol
            _arr = _labs[_lab_keys].to_numpy(float)
            _labs['三因子人工確認界限合格'] = np.all((_arr >= _lo - _tol) & (_arr <= _hi + _tol), axis=1)
            _labs['人工確認界限 + ΔE* 合格'] = _labs['三因子人工確認界限合格'] & _labs['LAB 規格合格']

            # Paired batch averages: NORTH/SOUTH per row, then batch mean.
            _line_parts = []
            for _f in _factors:
                _row_line = df[[f'正-北 {_f}', f'正-南 {_f}']].mean(axis=1)
                _agg = pd.DataFrame({'製造批號':df['製造批號'], 'v':_row_line}).groupby('製造批號')['v'].mean().rename(f'LINE_{_f}')
                _line_parts.append(_agg)
            _line_batch = pd.concat(_line_parts, axis=1).reset_index()
            _paired = _labs.merge(_line_batch, on='製造批號', how='inner', validate='one_to_one')
            _line_keys = [f'LINE_{f}' for f in _factors]
            _paired = _paired.replace([np.inf, -np.inf], np.nan).dropna(subset=_line_keys).copy()
            _paired['LINE ΔE*'] = np.linalg.norm(_paired[_line_keys].to_numpy(float), axis=1)
            _paired['LINE NG'] = _paired['LINE ΔE*'] > line_de_target + _tol
            for _f in _factors:
                _paired[f'Shift_{_f}'] = _paired[f'LINE_{_f}'] - _paired[f'LAB_{_f}']
            # In-spec LAB and complete paired readings, regardless of LINE pass/fail.
            _valid_paired = _paired.loc[_paired['LAB 規格合格']].copy()
            _shift_stats = task3_shift_summary(_valid_paired)
            _shift_matrix = _valid_paired[_shift_keys].to_numpy(float)

            # REPORT-READY THREE STEPS: Shift range -> backwards LAB bounds -> LINE check.
            _shift_min = _shift_matrix.min(axis=0) if len(_shift_matrix) else None
            _shift_max = _shift_matrix.max(axis=0) if len(_shift_matrix) else None
            st.markdown('### 歷史 LAB–LINE Shift 摘要（供分析）')
            if _shift_min is not None:
                _shift_range_table = pd.DataFrame({
                    '因子':_names, 'Shift 最小值':_shift_min, 'Shift 最大值':_shift_max,
                    '變動幅度 (Max−Min)':_shift_max-_shift_min,
                })
                st.dataframe(_shift_range_table.style.format(precision=2), use_container_width=True, hide_index=True)
                st.caption('資料來源：LAB ΔE* 達本色號規格且 LAB / LINE 三因子完整配對的批次。Shift = 同批 LINE − LAB；不事先移除 LINE NG。')
            else:
                _shift_range_table = pd.DataFrame()
                st.warning('沒有可用的 LAB–LINE 完整配對資料。')

            _stress_initial, _worst_initial_case = task3_stress_audit(_lo, _hi, _valid_paired, line_de_target)
            _current_peak = (float(_worst_initial_case['最不利 LINE ΔE*'])
                             if _worst_initial_case is not None else float('nan'))
            _current_envelope = (task3_shift_envelope_peak(_lo,_hi,_shift_min,_shift_max)
                                 if _shift_min is not None else float('nan'))
            _current_margin = line_de_target - _current_envelope
            _risk_overview = pd.DataFrame([
                {'評估項目':'LAB ΔE* 規格', '結果':lab_de_target},
                {'評估項目':'LINE ΔE* 規格', '結果':line_de_target},
                {'評估項目':'人工確認初始界限：LINE 保守上界', '結果':_current_envelope},
                {'評估項目':'LINE 規格餘裕', '結果':_current_margin},
            ])

            st.markdown('### 表一｜LAB 初步建議管制界限')
            _opt = task3_line_risk_box(_valid_paired[_lab_keys].to_numpy(float),
                                       _shift_matrix, lab_de_target, line_de_target)
            if 'error' in _opt:
                st.warning(_opt['error'])
                st.caption('無法確認候選界限時，不以虛構數值取代。')
            else:
                _b_lo=np.asarray(_opt['lo'],float)
                _b_hi=np.asarray(_opt['hi'],float)
                _new_worst=float(_opt['lab_worst'])
                _line_worst=float(_opt['line_worst'])
                _limits=pd.DataFrame({
                    '因子':_names,'人工確認 LCL':_lo,'人工確認 UCL':_hi,
                    '建議 LAB LCL':_b_lo,'建議 LAB UCL':_b_hi,
                })
                st.dataframe(_limits.style.format(precision=2),use_container_width=True,hide_index=True)
                _stress_proposed,_worst_proposed_case=task3_stress_audit(
                    _b_lo,_b_hi,_valid_paired,line_de_target)
                if _worst_proposed_case is not None:
                    _case_audit=task3_worst_case_calculation_tables(
                        _worst_proposed_case,_valid_paired,line_de_target,_b_lo,_b_hi)
                _proposed_margin=line_de_target-_line_worst
                st.caption('方法：Shift 各項 Min/Max → 補償中心 = −(Min+Max)/2（限制在 LAB 合格實測範圍）'
                           ' → 從中心向實測範圍擴展，求得同時符合 LAB 與 LINE 規格的最大係數 t。'
                           ' 係數由二分搜尋計算，非人工設定縮減比例。')
                st.caption(f'補償中心 = ({_opt["center"][0]:.2f}, {_opt["center"][1]:.2f}, {_opt["center"][2]:.2f})；'
                           f'擴展係數 t = {_opt["scale"]:.2f}。')

            st.markdown('### 表二｜改善前後 LINE 品質驗證')
            _comparison = pd.DataFrame([
                {'評估項目':'LINE 保守上界（Shift Min/Max）','人工確認初始界限':_current_envelope,'建議界限':_line_worst},
                {'評估項目':'LINE ΔE* 規格','人工確認初始界限':line_de_target,'建議界限':line_de_target},
                {'評估項目':'LINE 規格餘裕','人工確認初始界限':_current_margin,'建議界限':_proposed_margin},
            ])
            st.dataframe(_comparison.style.format({'人工確認初始界限':'{:.2f}','建議界限':'{:.2f}'},na_rep='—'),
                         use_container_width=True,hide_index=True)
            if _b_lo is None:
                _conclusion = ('按歷史 Shift Min/Max 的保守範圍，目前尚無法用簡易反推法提出可通過雙重規格的 LAB 界限；'
                               '需擴充資料及確認製程變動。')
            else:
                _conclusion = (f'人工確認界限下 LINE 保守上界 {_current_envelope:.2f}，改善後 {_line_worst:.2f}；'
                               f'規格餘裕 {_proposed_margin:.2f}。此為歷史 Shift 範圍的保守測試，'
                               '未包含未來新增的變動與量測不確定性，建議先行試行並持續驗證。')
            st.markdown('**初步結論：** '+_conclusion)
            st.caption('保守上界將三個 Shift 項目的歷史 Min/Max 分別組合，可能包含未曾同時出現的情況；'
                       '屬壓力測試，不是未來 NG 機率或產品保證。')

            # Only one audit expander, rather than multiple technical tables in main report.
            with st.expander('計算依據與逐批資料（供查核）', expanded=False):
                st.markdown('**數據來源｜同批 LAB、LINE 與 Shift = LINE − LAB**')
                st.dataframe(_paired.style.format(precision=2), use_container_width=True, hide_index=True)
                st.markdown('**Shift 分布（已觀察變動，非未來預測區間）**')
                if not _shift_stats.empty:
                    st.dataframe(_shift_stats.style.format(precision=2), use_container_width=True, hide_index=True)
                if _case_audit is not None:
                    st.markdown(f'**同批 Shift 情境來源批號：{_case_audit["batch"]}（非三因子 Min/Max 合成情境）**')
                    st.dataframe(_case_audit['source'].style.format(precision=2), use_container_width=True, hide_index=True)
                    st.markdown('**LINE 預測 = LAB 界限點 + 來源 Shift**')
                    st.dataframe(_case_audit['prediction'].style.format(precision=2), use_container_width=True, hide_index=True)
                    st.dataframe(_case_audit['verification'].style.format(
                        {'計算結果':lambda v: f'{v:.2f}' if isinstance(v,(float,np.floating)) else str(v)}),
                        use_container_width=True, hide_index=True)
                    st.markdown('**同一歷史批次 Shift 之八個 LAB 界限角落（補充查核，非 Min/Max 保守上界）**')
                    st.dataframe(_case_audit['corners'].style.format(precision=2), use_container_width=True, hide_index=True)
                st.latex(r'Shift_{i,j}=LINE_{i,j}-LAB_{i,j}')
                st.latex(r'\Delta E^*_{LINE,upper}=\sqrt{\sum_j\max_{x\in\{LCL_j,UCL_j\},\,s\in\{S_{j,min},S_{j,max}\}}|x+s|^2}')
                st.caption('界限判定使用未四捨五入值；畫面、Word 與 Excel 顯示兩位小數。')

            # Trace tables share the exact same, unrounded numerical inputs with the optimiser.
            # These are supplementary only; the management report has two main result tables.
            _appendix_center = pd.DataFrame()
            _appendix_limits = pd.DataFrame()
            _appendix_t_checks = pd.DataFrame()
            _appendix_corner_audit = pd.DataFrame()
            if isinstance(_opt, dict) and 'scale' in _opt and _b_lo is not None:
                _appendix_center = pd.DataFrame({
                    '色差因子': _names,
                    'Shift Min': _opt['shift_min'], 'Shift Max': _opt['shift_max'],
                    '理論補償中心 = −(Min+Max)/2': -_opt['shift_mid'],
                    'LAB 實測 Min': _opt['observed_lo'],
                    'LAB 實測 Max': _opt['observed_hi'],
                    '採用中心 C (限制於實測範圍)': _opt['center'],
                })
                _appendix_limits = pd.DataFrame({
                    '色差因子': _names, '中心 C': _opt['center'],
                    'LAB 實測 Min': _opt['observed_lo'],
                    'LAB 實測 Max': _opt['observed_hi'],
                    '擴展係數 t': [_opt['scale']]*3,
                    '建議 LCL': _b_lo, '建議 UCL': _b_hi,
                })
                _trace_t = [0.0, float(_opt['scale']), 1.0]
                _trace_rows=[]
                for _trial_t in _trace_t:
                    _trace_lo=_opt['center']+_trial_t*(_opt['observed_lo']-_opt['center'])
                    _trace_hi=_opt['center']+_trial_t*(_opt['observed_hi']-_opt['center'])
                    _trial_lab=float(np.linalg.norm(np.maximum(np.abs(_trace_lo),np.abs(_trace_hi))))
                    _trial_line=task3_shift_envelope_peak(_trace_lo,_trace_hi,_opt['shift_min'],_opt['shift_max'])
                    _trace_rows.append({'係數 t':_trial_t,'LAB ΔE* 最不利值':_trial_lab,
                        'LINE ΔE* 最不利值':_trial_line,
                        'LAB 判定':'符合' if _trial_lab<=lab_de_target+1e-10 else '超規',
                        'LINE 判定':'符合' if _trial_line<=line_de_target+1e-10 else '超規'})
                _appendix_t_checks=pd.DataFrame(_trace_rows)
                _corners=[]
                for _j, _factor in enumerate(_names):
                    _c=float(_opt['center'][_j]); _t=float(_opt['scale'])
                    _mn=float(_opt['observed_lo'][_j]); _mx=float(_opt['observed_hi'][_j])
                    _corners.append({'因子':_factor,'LCL 代入式':f'{_c:.2f} + {_t:.2f} × ({_mn:.2f} − {_c:.2f})',
                                     'LCL 計算結果':float(_b_lo[_j]),
                                     'UCL 代入式':f'{_c:.2f} + {_t:.2f} × ({_mx:.2f} − {_c:.2f})',
                                     'UCL 計算結果':float(_b_hi[_j])})
                _appendix_corner_audit=pd.DataFrame(_corners)

            # Formula-level Excel audit with full numerical precision and 0.00 display.
            _excel_bytes = io.BytesIO()
            with pd.ExcelWriter(_excel_bytes, engine='openpyxl') as _writer:
                _shift_range_table.to_excel(_writer, sheet_name='5_1_Shift_Range', index=False)
                _risk_overview.to_excel(_writer, sheet_name='Confirmed_Start_Risk', index=False)
                _selected_limits_df.to_excel(_writer, sheet_name='Confirmed_Input_Limits', index=False)
                _limits.to_excel(_writer, sheet_name='5_2_LAB_Limits', index=False)
                _comparison.to_excel(_writer, sheet_name='5_3_LINE_Compare', index=False)
                if not _appendix_center.empty:
                    _appendix_center.to_excel(_writer, sheet_name='Appendix_Center', index=False)
                    _appendix_t_checks.to_excel(_writer, sheet_name='Appendix_t_Check', index=False)
                    _appendix_limits.to_excel(_writer, sheet_name='Appendix_Limits', index=False)
                    _appendix_corner_audit.to_excel(_writer, sheet_name='Appendix_Substitution', index=False)
                _labs.to_excel(_writer, sheet_name='All_LAB', index=False)
                _paired.to_excel(_writer, sheet_name='Paired_LAB_LINE', index=False)
                _valid_paired.to_excel(_writer, sheet_name='Shift_Calibration', index=False)
                _shift_stats.to_excel(_writer, sheet_name='Shift_Summary', index=False)
                _stress_initial.to_excel(_writer, sheet_name='Confirmed_Shift_Stress', index=False)
                _stress_proposed.to_excel(_writer, sheet_name='Proposed_Shift_Stress', index=False)
                if _case_audit is not None:
                    for _sheet,_key in [('Source_Shift','source'),('LINE_Calculation','prediction'),
                                         ('Risk_Verification','verification'),('Eight_Corners','corners')]:
                        _case_audit[_key].to_excel(_writer, sheet_name=_sheet, index=False)
                pd.DataFrame([
                    {'計算項目':'歷史 Shift','公式':'Shift_j = LINE_j − LAB_j'},
                    {'計算項目':'LINE Min/Max 保守上界','公式':'SQRT(SUM_j MAX(ABS(LCL_j+ShiftMin_j), ABS(LCL_j+ShiftMax_j), ABS(UCL_j+ShiftMin_j), ABS(UCL_j+ShiftMax_j))^2)'},
                    {'計算項目':'LINE 規格餘裕','公式':'LINE ΔE* 規格 − LINE Min/Max 保守上界'},
                    {'計算項目':'Shift 最小 / 最大','公式':'MIN_i(LINE_ij − LAB_ij) / MAX_i(LINE_ij − LAB_ij)'},
                    {'計算項目':'Shift 補償中心','公式':'CLIP(−(ShiftMin+ShiftMax)/2, historical LAB min, historical LAB max)'},
                    {'計算項目':'LAB 候選界限','公式':'LCL = center+t×(observed min−center); UCL = center+t×(observed max−center)'},
                    {'計算項目':'擴展係數 t','公式':'MAX t in [0,1] satisfying LAB ΔE* and independent Min/Max Shift LINE ΔE* constraints (binary search)'},
                ]).to_excel(_writer, sheet_name='Formulas', index=False)
                for _sheet in _writer.book.worksheets:
                    _sheet.freeze_panes='A2'
                    for _row in _sheet.iter_rows(min_row=2):
                        for _cell in _row:
                            if isinstance(_cell.value,(float,np.floating)) and not isinstance(_cell.value,bool):
                                _cell.number_format='0.00'
            st.download_button('📥 下載分析與逐批計算 Excel', _excel_bytes.getvalue(),
                file_name=f'{re.sub(r"[^A-Za-z0-9_-]+","_",str(color))}_LAB_LINE_Risk_Audit.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key='task3_line_risk_excel')

    # ---------------------------------------------------------
    # ---------------------------------------------------------
    # Task 3 Word Report Export
    # ---------------------------------------------------------
    if calc_res:
        st.markdown("---")
        st.markdown("### 📄 Export Task 3 Word Report")
        st.caption(
            "The Word report retains the original statistical analysis, factor results, full three-chart page and original conclusion; supplementary Task 3 results and appendix follow."
        )

        def build_task3_word_report():
            doc = docx.Document()
            def add_table(frame, cols=None):
                """Append supplementary Task 3 table, preserving original report tables."""
                if frame is None or frame.empty:
                    doc.add_paragraph('本色號資料不足，無法提供有效計算結果。')
                    return
                cols = list(cols or frame.columns)
                table = doc.add_table(rows=1, cols=len(cols))
                table.style = 'Table Grid'
                for k, column in enumerate(cols):
                    table.rows[0].cells[k].text = str(column)
                for _, record in frame.iterrows():
                    cells = table.add_row().cells
                    for k, column in enumerate(cols):
                        value = record[column]
                        if isinstance(value, (int, np.integer)) and not isinstance(value, bool):
                            cells[k].text = str(int(value))
                        elif isinstance(value, (float, np.floating)):
                            cells[k].text = f'{float(value):.2f}' if np.isfinite(value) else '—'
                        else:
                            cells[k].text = '—' if pd.isna(value) else str(value)
                doc.add_paragraph('')


            # A4 portrait for the normal report pages
            section = doc.sections[0]
            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Mm(210)
            section.page_height = Mm(297)
            section.top_margin = Inches(0.55)
            section.bottom_margin = Inches(0.55)
            section.left_margin = Inches(0.60)
            section.right_margin = Inches(0.60)

            # Base font in Word only. This does not change Matplotlib chart fonts.
            normal_style = doc.styles["Normal"]
            normal_style.font.name = "Arial"
            normal_style.font.size = Pt(10)

            title = doc.add_heading("Incoming LAB Control Limit Analysis Report", level=0)
            title.alignment = 1

            p = doc.add_paragraph()
            p.add_run("Color Code: ").bold = True
            p.add_run(str(color))
            p.add_run("\nAnalysis Date: ").bold = True
            p.add_run(datetime.now().strftime("%Y-%m-%d %H:%M"))

            if not df.empty:
                p.add_run("\nData Period: ").bold = True
                p.add_run(
                    f"{df['Time'].min().strftime('%Y-%m-%d')} to {df['Time'].max().strftime('%Y-%m-%d')}"
                )
                p.add_run("\nTotal filtered batches: ").bold = True
                p.add_run(str(df["製造批號"].nunique()))

            doc.add_heading("1. Analysis Objective", level=1)
            doc.add_paragraph(
                "Determine a practical incoming LAB control range for ΔL*, Δa* and Δb*, and use the paired "
                "LINE output from the same production batches to evaluate the downstream color trend. "
                "The recommended LAB range is derived from the overlap between the Standard (Mean ± Kσ) "
                "and IQR methods using LAB data associated with acceptable LINE output whenever current "
                "LINE control limits are available."
            )

            doc.add_heading("2. LAB Input Statistical Summary (±1.5σ)", level=1)
            doc.add_paragraph(
                "This table summarizes LAB input data only. Min, Max, mean (x̄), standard deviation (σ), "
                "and control limits are calculated at batch level. LCL/UCL are defined as x̄ ± 1.5σ."
            )
            stats_table = doc.add_table(rows=1, cols=7)
            stats_table.style = "Table Grid"
            stats_headers = [
                "入料色差走向", "Min", "Max", "平均值 (x̄)", "標準差 (σ)",
                "LCL (-1.5σ)", "UCL (+1.5σ)"
            ]
            for _i, _h in enumerate(stats_headers):
                stats_table.rows[0].cells[_i].text = _h

            for _, _r in task3_stats_df.iterrows():
                _cells = stats_table.add_row().cells
                _cells[0].text = str(_r["入料色差走向"])
                for _j, _col in enumerate(stats_headers[1:], start=1):
                    _val = _r[_col]
                    _cells[_j].text = f"{float(_val):.3f}" if pd.notna(_val) else "-"

            doc.add_heading("3. Analysis Method", level=1)
            method_table = doc.add_table(rows=1, cols=3)
            method_table.style = "Table Grid"
            hdr = method_table.rows[0].cells
            hdr[0].text = "Method"
            hdr[1].text = "Calculation"
            hdr[2].text = "Purpose"
            method_rows = [
                ("Standard", "Mean ± K × σ", "Evaluate LAB distribution using mean and standard deviation"),
                ("IQR", "Q1 − k×IQR to Q3 + k×IQR", "Reduce sensitivity to extreme values"),
                ("Recommended LAB", "Intersection of Standard and IQR", "Candidate incoming LAB control window"),
                ("LAB→LINE", "Pearson r and R²", "Evaluate the relationship between incoming LAB and LINE output"),
            ]
            for m, calc, purpose in method_rows:
                cells = method_table.add_row().cells
                cells[0].text = m
                cells[1].text = calc
                cells[2].text = purpose

            doc.add_heading("4. Factor Results", level=1)

            for factor in factors:
                if factor not in calc_res:
                    continue
                res = calc_res[factor]
                doc.add_heading(f"4.{factors.index(factor)+1} {delta_label(factor)} Analysis", level=2)

                result_table = doc.add_table(rows=1, cols=2)
                result_table.style = "Table Grid"
                h = result_table.rows[0].cells
                h[0].text = "Item"
                h[1].text = "Result"

                lab_sheet = (
                    f"{res['lab_lcl_sheet']:.3f} ~ {res['lab_ucl_sheet']:.3f}"
                    if res['lab_lcl_sheet'] is not None and res['lab_ucl_sheet'] is not None else "Not set"
                )
                line_sheet = (
                    f"{res['line_lcl_sheet']:.3f} ~ {res['line_ucl_sheet']:.3f}"
                    if res['line_lcl_sheet'] is not None and res['line_ucl_sheet'] is not None else "Not set"
                )
                rec = (
                    f"{res['rec_lcl']:.3f} ~ {res['rec_ucl']:.3f}"
                    if res['rec_lcl'] is not None and res['rec_ucl'] is not None else "No overlap"
                )
                rows = [
                    ("Current LAB limit", lab_sheet),
                    ("Current LINE limit", line_sheet),
                    (f"Standard LAB ({res['sig']:.1f}σ)", f"{res['std_lcl']:.3f} ~ {res['std_ucl']:.3f}"),
                    (f"IQR LAB (k={res['iqr_k']:.2f})", f"{res['iqr_lcl']:.3f} ~ {res['iqr_ucl']:.3f}"),
                    ("Recommended LAB", rec),
                    ("Paired batches", str(len(res['pair']))),
                    ("LAB data used", str(len(res['safe_lab']))),
                    ("LAB→LINE r", f"{res['r']:.3f}" if pd.notna(res['r']) else "N/A"),
                    ("LAB→LINE R²", f"{res['r2']:.3f}" if pd.notna(res['r2']) else "N/A"),
                ]
                for item, value in rows:
                    c = result_table.add_row().cells
                    c[0].text = item
                    c[1].text = value

                doc.add_paragraph(res['filter_note'])

                # Charts are placed together on one dedicated A4 landscape page below.

            if len(calc_res) == 3:
                doc.add_heading("5. Recommended Incoming LAB Control Limit Summary", level=1)
                sum_table = doc.add_table(rows=1, cols=8)
                sum_table.style = "Table Grid"
                headers = [
                    "Factor", "Paired n", "LAB n", "Mean", "σ",
                    "Recommended LCL", "Recommended UCL", "R²"
                ]
                for i, htxt in enumerate(headers):
                    sum_table.rows[0].cells[i].text = htxt

                for factor in factors:
                    res = calc_res[factor]
                    row = sum_table.add_row().cells
                    vals = [
                        factor,
                        str(len(res['pair'])),
                        str(len(res['safe_lab'])),
                        f"{res['m']:.3f}",
                        f"{res['s']:.3f}",
                        f"{res['rec_lcl']:.3f}" if res['rec_lcl'] is not None else "N/A",
                        f"{res['rec_ucl']:.3f}" if res['rec_ucl'] is not None else "N/A",
                        f"{res['r2']:.3f}" if pd.notna(res['r2']) else "N/A",
                    ]
                    for i, value in enumerate(vals):
                        row[i].text = value

                # One-page high-resolution chart overview, arranged like the user's reference:
                # ΔL* top-left, Δa* top-right, Δb* centered below.
                if all(f in report_chart_buffers for f in factors):
                    chart_section = doc.add_section(WD_SECTION.NEW_PAGE)
                    chart_section.orientation = WD_ORIENT.LANDSCAPE
                    chart_section.page_width = Mm(297)
                    chart_section.page_height = Mm(210)
                    chart_section.top_margin = Inches(0.28)
                    chart_section.bottom_margin = Inches(0.28)
                    chart_section.left_margin = Inches(0.30)
                    chart_section.right_margin = Inches(0.30)

                    p_title = doc.add_paragraph()
                    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r_title = p_title.add_run(
                        f"Figure 2  ΔL*, Δa* and Δb* LAB Input vs LINE Output Trend ({color})"
                    )
                    r_title.bold = True
                    r_title.font.size = Pt(13)

                    chart_table = doc.add_table(rows=2, cols=2)
                    chart_table.autofit = False
                    tl = chart_table.cell(0, 0)
                    tr = chart_table.cell(0, 1)
                    bot = chart_table.cell(1, 0).merge(chart_table.cell(1, 1))

                    for cell in [tl, tr, bot]:
                        cell.vertical_alignment = 1
                        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                    tl.paragraphs[0].add_run().add_picture(
                        io.BytesIO(report_chart_buffers["ΔL"]), width=Inches(4.85)
                    )
                    tr.paragraphs[0].add_run().add_picture(
                        io.BytesIO(report_chart_buffers["Δa"]), width=Inches(4.85)
                    )
                    bot.paragraphs[0].add_run().add_picture(
                        io.BytesIO(report_chart_buffers["Δb"]), width=Inches(5.45)
                    )

                    note = doc.add_paragraph(
                        "Note: X-axis shows batch sequence in chronological order. "
                        "The blue line is LAB Input and the orange line is LINE Output."
                    )
                    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in note.runs:
                        run.font.size = Pt(9)

                    # Continue the remaining report on A4 portrait pages.
                    portrait_section = doc.add_section(WD_SECTION.NEW_PAGE)
                    portrait_section.orientation = WD_ORIENT.PORTRAIT
                    portrait_section.page_width = Mm(210)
                    portrait_section.page_height = Mm(297)
                    portrait_section.top_margin = Inches(0.55)
                    portrait_section.bottom_margin = Inches(0.55)
                    portrait_section.left_margin = Inches(0.60)
                    portrait_section.right_margin = Inches(0.60)

                # Derived ΔE* results, using the same values already calculated in the app.
                dE_std_report = math.sqrt(dE_std_sq)
                dE_iqr_report = math.sqrt(dE_iqr_sq)
                valid_rec = all(
                    calc_res[f]['rec_lcl'] is not None and calc_res[f]['rec_ucl'] is not None
                    for f in factors
                )
                if valid_rec:
                    dE_rec_report = math.sqrt(sum(
                        max(abs(calc_res[f]['rec_lcl']), abs(calc_res[f]['rec_ucl'])) ** 2
                        for f in factors
                    ))
                else:
                    dE_rec_report = None

                doc.add_heading("6. Derived Incoming LAB ΔE* Boundary", level=1)
                de_table = doc.add_table(rows=1, cols=2)
                de_table.style = "Table Grid"
                de_table.rows[0].cells[0].text = "Item"
                de_table.rows[0].cells[1].text = "Value"
                de_rows = [
                    ("Standard-derived ΔE*", f"{dE_std_report:.3f}"),
                    ("IQR-derived ΔE*", f"{dE_iqr_report:.3f}"),
                    ("Recommended ΔE*", f"{dE_rec_report:.3f}" if dE_rec_report is not None else "N/A"),
                    ("Current LAB ΔE* target", f"≤ {lab_de_target:.2f}"),
                    ("Current LINE ΔE* target", f"≤ {line_de_target:.2f}"),
                ]
                for item, value in de_rows:
                    c = de_table.add_row().cells
                    c[0].text = item
                    c[1].text = value

                doc.add_heading("7. Conclusion", level=1)
                conclusion_parts = []
                for factor in factors:
                    res = calc_res[factor]
                    if res['rec_lcl'] is not None and res['rec_ucl'] is not None:
                        conclusion_parts.append(
                            f"{factor}: recommended incoming LAB range {res['rec_lcl']:.3f} to {res['rec_ucl']:.3f}"
                        )
                if conclusion_parts:
                    doc.add_paragraph(
                        "Based on the current paired historical data, " + "; ".join(conclusion_parts) + ". "
                        "These ranges should be treated as candidate incoming control limits and continuously verified "
                        "against future LINE output stability as more paired batches become available."
                    )
                else:
                    doc.add_paragraph(
                        "A complete recommended incoming LAB range could not be established for all factors from the "
                        "current data. Additional paired LAB and LINE data are required before finalizing the incoming limits."
                    )

            # ADDITION ONLY: the legacy 1–7 sections, original analysis tables,
            # original dedicated three-chart landscape page and conclusion above
            # are deliberately preserved without replacement.
            doc.add_page_break()
            doc.add_heading('8. 補充分析｜LAB 入料管制界限初步建議與 LINE 品質驗證', level=1)
            doc.add_paragraph(f'塗料編號：{color}；LAB ΔE* 規格 ≤ {lab_de_target:.2f}；'
                              f'LINE ΔE* 規格 ≤ {line_de_target:.2f}。')
            doc.add_paragraph('本研究以人工確認之初始界限（可參考 IQR 與標準差法）為比較基準，並依歷史 LAB–LINE 色差偏移，'
                              '初步建立三項入料色差管制範圍；詳細計算及資料來源列於附錄。')
            doc.add_paragraph('初始界限來源：人工確認（可採 IQR、標準差法或調整值）；原始統計結果不受更改。')
            doc.add_paragraph('表｜人工確認初始管制界限')
            add_table(_selected_limits_df)
            doc.add_paragraph('表｜LAB 初步建議管制界限')
            add_table(_limits)
            doc.add_paragraph('表｜改善前後 LINE 品質驗證')
            add_table(_comparison)
            doc.add_paragraph('結論：' + (_conclusion or '目前資料不足，無法提出初步界限。'))
            doc.add_paragraph('註：上述結果為歷史 Shift Min/Max 情境下的初步建議，'
                              '並非已證實之最適界限，亦不代表未來 LINE 零超規。'
                              '所有判定採用未四捨五入數值，表列僅顯示小數點後二位。')

            doc.add_page_break()
            doc.add_heading('附錄｜LAB 建議管制界限計算依據', level=1)
            doc.add_heading('A1. 歷史 LAB–LINE Shift 來源', level=2)
            doc.add_paragraph('資料單位：同一製造批號之 LAB 與 LINE 三項色差完整配對；'
                              'Shift = LINE 實測值 − LAB 實測值。'
                              'Shift 估計使用 LAB ΔE* 符合本色號規格之有效配對批次，未排除 LINE NG。')
            if not _valid_paired.empty:
                _raw_cols = ['製造批號'] + _lab_keys + [f'LINE_{f}' for f in _factors] + _shift_keys
                add_table(_valid_paired[_raw_cols])
            else:
                doc.add_paragraph('無完整有效配對批次。')
            doc.add_heading('A2. Shift 變動與補償中心', level=2)
            doc.add_paragraph('Shift Min = MIN(各批 Shift)；Shift Max = MAX(各批 Shift)；'
                              '理論中心 = −(Shift Min + Shift Max) / 2；'
                              '採用中心 C = MIN(MAX(理論中心, LAB 實測 Min), LAB 實測 Max)。')
            add_table(_appendix_center)

            doc.add_heading('A3. 擴展係數 t 之確認', level=2)
            doc.add_paragraph('從中心 C 向合格 LAB 實測 Min/Max 延伸：'
                              'LCL(t) = C + t × (LAB Min − C)；'
                              'UCL(t) = C + t × (LAB Max − C)，0 ≤ t ≤ 1。'
                              '二分搜尋最大可行 t，須同時滿足 LAB ΔE* 及 LINE ΔE* 規格。')
            doc.add_paragraph('LAB 最不利值 = SQRT[Σ MAX(|LCL|, |UCL|)²]。')
            doc.add_paragraph('LINE 保守最不利值 = SQRT[Σ MAX(|LCL+Shift Min|, |LCL+Shift Max|, '
                              '|UCL+Shift Min|, |UCL+Shift Max|)²]。')
            add_table(_appendix_t_checks)
            if isinstance(_opt, dict) and 'scale' in _opt:
                doc.add_paragraph(f'採用擴展係數 t = {_opt["scale"]:.2f}。'
                                  ' t = 0 及 t = 1 是端點查核；本表不代表二分搜尋每次迭代紀錄。')

            doc.add_heading('A4. LAB 三因子建議 LCL / UCL 逐項代入', level=2)
            add_table(_appendix_corner_audit)
            doc.add_paragraph('上述代入式以兩位小數顯示，計算以未四捨五入之實際數值為準，'
                              '因此顯示數字直接重算可能有末位差異。')

            doc.add_heading('A5. LINE 最不利情境確認', level=2)
            doc.add_paragraph('比較基準為使用者人工確認之 LAB 三因子初始界限，而非自動計算的 IQR 界限。')
            add_table(_comparison)
            doc.add_paragraph('LINE 規格餘裕 = LINE ΔE* 規格 − LINE 保守最不利值。'
                              '保守情境將三因子 Shift 的觀測極值獨立組合，'
                              '可能包含歷史上未曾同時出現的組合。')
            if _case_audit is not None:
                doc.add_paragraph('同批 Shift 核對範例（供追溯，與 Min/Max 獨立組合之保守上界不同）：'
                                  + str(_case_audit['batch']))
                add_table(_case_audit['source'])
                add_table(_case_audit['prediction'])
            doc.add_paragraph('後續需累積更多實際生產數據，再評估界限之適用性與 LINE 產出品質風險。')

            # Footer
            for sec in doc.sections:
                footer = sec.footer.paragraphs[0]
                footer.text = "SPC Color Dashboard - Task 3 Incoming LAB Control Limit Analysis"
                footer.alignment = 1

            report_buf = io.BytesIO()
            doc.save(report_buf)
            report_buf.seek(0)
            return report_buf.getvalue()

        safe_color_name = re.sub(r"[^A-Za-z0-9_-]+", "_", str(color))
        report_state_key = "task3_export_payload"
        report_identity = ("WORD_CHARTS_VERIFIED_V3", tuple(_start_lcl), tuple(_start_ucl), uploaded_data_file.name, uploaded_data_file.size,
                           uploaded_limit_file.name, uploaded_limit_file.size,
                           str(color), tuple(selected_years), tuple(selected_months),
                           float(lab_de_target), float(line_de_target))
        if st.session_state.get("task3_export_identity") != report_identity:
            st.session_state.pop(report_state_key, None)
            st.session_state["task3_export_identity"] = report_identity
        if st.button("📄 Prepare Word Report", key="task3_prepare_word"):
            with st.spinner("Generating high-resolution report charts and Word file..."):
                # Restore the ORIGINAL pre-export chart generation. Without this,
                # report_chart_buffers remains empty and the original Word chart
                # page is silently skipped by its `if all(...)` guard.
                report_chart_buffers.clear()
                missing_report_factors = []
                for _factor in factors:
                    _res = calc_res.get(_factor)
                    _pair = _res.get("pair") if _res is not None else None
                    if _pair is None or _pair.empty:
                        missing_report_factors.append(_factor)
                        continue
                    _report_fig = None
                    try:
                        _report_fig = build_task3_report_chart(_pair, _factor)
                        _report_buf = io.BytesIO()
                        _report_fig.savefig(
                            _report_buf, format="png", dpi=500,
                            bbox_inches="tight", facecolor="white",
                            pad_inches=0.04,
                        )
                        _png_bytes = _report_buf.getvalue()
                        if not _png_bytes:
                            raise ValueError(f"Empty chart PNG for {_factor}")
                        report_chart_buffers[_factor] = _png_bytes
                    finally:
                        if _report_fig is not None:
                            plt.close(_report_fig)

                if missing_report_factors:
                    st.error(
                        "Word chart page cannot be generated: missing paired "
                        "LAB/LINE data for " + ", ".join(missing_report_factors)
                        + ". No incomplete Word report has been generated."
                    )
                    st.session_state.pop(report_state_key, None)
                else:
                    # All three images must be present before saving DOCX.
                    # Never silently ship a Word report without the chart page.
                    missing_images = [f for f in factors if f not in report_chart_buffers]
                    if missing_images:
                        st.error("Chart generation incomplete: " + ", ".join(missing_images))
                        st.session_state.pop(report_state_key, None)
                    else:
                        _word_bytes = build_task3_word_report()
                        # Verify embedded media exist in the actual DOCX, not
                        # just inside the in-memory plot dictionaries.
                        import zipfile
                        with zipfile.ZipFile(io.BytesIO(_word_bytes)) as _docx_zip:
                            _embedded_images = [n for n in _docx_zip.namelist()
                                                if n.startswith("word/media/") and n.lower().endswith(".png")]
                        if len(_embedded_images) < 3:
                            st.error(
                                f"Word chart verification failed: only {len(_embedded_images)} "
                                "embedded PNG images were found (expected 3)."
                            )
                            st.session_state.pop(report_state_key, None)
                        else:
                            st.session_state[report_state_key] = _word_bytes
        if st.session_state.get(report_state_key):
            st.download_button(
                label="📄 Download Word Report",
                data=st.session_state[report_state_key],
                file_name=f"Task3_Incoming_LAB_Control_Limit_{safe_color_name}_{datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"task3_word_{safe_color_name}"
            )

    # ---------------------------------------------------------
    # Manual LAB ΔE calculator
    # ---------------------------------------------------------
    st.markdown("---")
    st.subheader("🧮 Manual Incoming LAB ΔE* Calculator")
    col_ml, col_ma, col_mb = st.columns(3)
    with col_ml:
        man_L = st.number_input("Incoming LAB ΔL*:", value=0.000, step=0.100, format="%.2f")
    with col_ma:
        man_a = st.number_input("Incoming LAB Δa*:", value=0.000, step=0.100, format="%.2f")
    with col_mb:
        man_b = st.number_input("Incoming LAB Δb*:", value=0.000, step=0.100, format="%.2f")

    manual_dE = math.sqrt(man_L ** 2 + man_a ** 2 + man_b ** 2)

    st.markdown("#### Calculation Result")
    if manual_dE <= lab_de_target:
        st.success(f"### 🎯 Incoming LAB ΔE*: **{manual_dE:.2f}** (✅ ≤ {lab_de_target})")
    else:
        st.error(f"### 🎯 Incoming LAB ΔE*: **{manual_dE:.2f}** (⚠️ > {lab_de_target})")

# =========================================================
# VIEW 4: LAB VS LINE SCALE-UP ANALYSIS
# =========================================================
elif app_mode == "🔬 Lab vs Line Scale-up":
    st.title("🔬 Lab to Line Scale-up Analysis")
    st.markdown("""
    Analyze the historical deviation between **Laboratory (LAB)** inputs and **Production (LINE)** outcomes. 
    Use this tool to determine the necessary **Offset Compensation** for color formulation.
    """)

    if df.empty:
        st.warning("⚠️ No data available for analysis.")
    else:
        batch_compare = df.groupby("製造批號", as_index=False).agg({
            "入料檢測 ΔL 正面": "mean", "入料檢測 Δa 正面": "mean", "入料檢測 Δb 正面": "mean",
            "正-北 ΔL": "mean", "正-南 ΔL": "mean",
            "正-北 Δa": "mean", "正-南 Δa": "mean",
            "正-北 Δb": "mean", "正-南 Δb": "mean"
        }).dropna()

        batch_compare["LINE_ΔL"] = batch_compare[["正-北 ΔL", "正-南 ΔL"]].mean(axis=1)
        batch_compare["LINE_Δa"] = batch_compare[["正-北 Δa", "正-南 Δa"]].mean(axis=1)
        batch_compare["LINE_Δb"] = batch_compare[["正-北 Δb", "正-南 Δb"]].mean(axis=1)

        batch_compare = batch_compare.rename(columns={
            "入料檢測 ΔL 正面": "LAB_ΔL",
            "入料檢測 Δa 正面": "LAB_Δa",
            "入料檢測 Δb 正面": "LAB_Δb"
        })

        factors = ["ΔL", "Δa", "Δb"]
        tabs = st.tabs([f"Factor {f}" for f in factors])
        
        for i, f in enumerate(factors):
            with tabs[i]:
                lab_col, line_col = f"LAB_{f}", f"LINE_{f}"
                x, y = batch_compare[lab_col].values, batch_compare[line_col].values
                
                if len(x) < 3:
                    st.info(f"Insufficient paired data for {f}.")
                    continue
                    
                diff = y - x
                mean_bias = np.mean(diff)
                std_dev = np.std(diff)
                slope, intercept = np.polyfit(x, y, 1)
                r2_score = (np.corrcoef(x, y)[0, 1])**2
                
                st.markdown(f"### 📊 Process Metrics: **{f}**")
                m1, m2, m3 = st.columns(3)
                m1.metric("Systematic Bias (Avg)", f"{mean_bias:+.3f}", help="Average deviation between Line and Lab")
                m2.metric("Fluctuation (1σ)", f"±{std_dev:.3f}", help="Standard deviation of the shift")
                m3.metric("Predictability (R²)", f"{r2_score:.3f}", help="Model reliability (closer to 1.0 is better)")

                if f == "ΔL":
                    direction = "LIGHTER" if mean_bias > 0.05 else "DARKER" if mean_bias < -0.05 else "STABLE"
                elif f == "Δa":
                    direction = "REDDER" if mean_bias > 0.05 else "GREENER" if mean_bias < -0.05 else "STABLE"
                else: 
                    direction = "YELLOWER" if mean_bias > 0.05 else "BLUER" if mean_bias < -0.05 else "STABLE"

                if direction != "STABLE":
                    st.warning(f"💡 **Insight:** Production tends to be **{direction}** than Lab samples (Offset: {mean_bias:+.3f}).")
                else:
                    st.success(f"✅ **Insight:** Production results are highly consistent with Lab inputs.")

                col_chart, col_pred = st.columns([2.2, 1])
                
                with col_pred:
                    st.subheader("🔮 Outcome Predictor")
                    user_lab = st.number_input(f"Current LAB {f}:", value=float(x[-1]), step=0.01, format="%.3f", key=f"f5_en_{f}")
                    
                    pred_line = slope * user_lab + intercept
                    ci_95 = 2 * std_dev
                    
                    st.info(f"**Predicted LINE {f}:**\n## {pred_line:.3f}")
                    st.caption(f"Confidence Range (95%):\n**[{pred_line-ci_95:.3f} to {pred_line+ci_95:.3f}]**")
                    st.success(f"🛠 **Lab Suggestion:**\nTo reach 0.000 on LINE, formulate LAB at: **{-mean_bias:+.3f}**")

                with col_chart:
                    fig, ax = plt.subplots(figsize=(8, 6))
                    ax.grid(True, linestyle="--", alpha=0.3, zorder=0)
                    
                    ax.scatter(x, y, alpha=0.6, color="#3498db", edgecolors="white", s=80, label="Historical Data", zorder=3)
                    ax.scatter(user_lab, pred_line, color="#f1c40f", edgecolors="black", s=300, marker="*", label="Prediction ⭐", zorder=5)
                    
                    mn, mx = min(x.min(), y.min(), user_lab, pred_line) - 0.1, max(x.max(), y.max(), user_lab, pred_line) + 0.1
                    ax.plot([mn, mx], [mn, mx], color="#7f8c8d", linestyle="--", alpha=0.6, label="Ideal (LINE = LAB)", zorder=1)
                    ax.plot(np.linspace(mn, mx, 100), slope * np.linspace(mn, mx, 100) + intercept, color="#e74c3c", linewidth=2.5, label="Actual Trend", zorder=2)
                    
                    bbox_style = dict(boxstyle="round,pad=0.3", alpha=0.1, lw=1)
                    if f == "ΔL":
                        ax.annotate("☀️ Lighter", xy=(0.95, 0.95), xycoords='axes fraction', ha='right', bbox=dict(facecolor='yellow', **bbox_style))
                        ax.annotate("🌑 Darker", xy=(0.05, 0.05), xycoords='axes fraction', ha='left', bbox=dict(facecolor='gray', **bbox_style))
                    elif f == "Δa":
                        ax.annotate("🔴 Redder", xy=(0.95, 0.95), xycoords='axes fraction', ha='right', bbox=dict(facecolor='red', **bbox_style))
                        ax.annotate("🟢 Greener", xy=(0.05, 0.05), xycoords='axes fraction', ha='left', bbox=dict(facecolor='green', **bbox_style))
                    elif f == "Δb":
                        ax.annotate("🟡 Yellower", xy=(0.95, 0.95), xycoords='axes fraction', ha='right', bbox=dict(facecolor='orange', **bbox_style))
                        ax.annotate("🔵 Bluer", xy=(0.05, 0.05), xycoords='axes fraction', ha='left', bbox=dict(facecolor='blue', **bbox_style))

                    ax.set_title(f"Lab-to-Line Scale-up: {f}", fontweight='bold')
                    ax.set_xlabel(f"LAB Input ({f})")
                    ax.set_ylabel(f"LINE Actual ({f})")
                    ax.legend(loc='lower right')
                    ax.set_xlim(mn, mx); ax.set_ylim(mn, mx)
                    
                    st.pyplot(fig)
                    plt.close(fig)

# =========================================================
# VIEW 5: AI OOC WORD REPORT GENERATOR
# =========================================================
elif app_mode == "📄 AI OOC Word Report":
    st.title("📄 AI Automated OOC Report Generator")
    st.markdown("The system will scan **ALL** color codes with active control limits (Phase II) to find Out-Of-Control batches. The AI will then automatically package the data, current limits, charts, and corrective recommendations into a comprehensive Word report.")

    if st.button("🚀 Run System Scan & Generate Word Report", type="primary"):
        with st.spinner("🔍 Scanning the entire system and generating analytical insights... This may take 10-30 seconds."):
            all_colors = sorted(df_raw["塗料編號"].dropna().unique())
            master_ooc_rows = []
            report_data_dict = {} 

            for c in all_colors:
                c_clean = str(c).strip()
                mask = limit_df["Color_code"].astype(str).str.strip() == c_clean
                row = limit_df[mask]
                
                if row.empty: continue
                limit_cols = [col for col in row.columns if "LCL" in col or "UCL" in col]
                if row[limit_cols].isna().all().all(): continue

                df_c = df_raw[df_raw["塗料編號"] == c].copy()
                cb = get_control_batch(c)
                cb_code = get_control_batch_code(df_c, cb)
                
                c_time = df_c[df_c["製造批號"] == cb_code]["Time"].min() if cb_code else None

                if c_time is None: continue

                global_batch_order_c = df_c.sort_values("Time")["製造批號"].drop_duplicates().tolist()
                spc_c = calculate_batch_averages(df_c[SPC_INPUT_COLUMNS])
                color_has_ooc = False

                for k in ["ΔL", "Δa", "Δb"]:
                    lcl_line, ucl_line = safe_get_limit(c, "LINE", k)
                    if lcl_line is not None and ucl_line is not None:
                        line_phase2 = spc_c[k]["line"][spc_c[k]["line"]["Time"] >= c_time]
                        ooc_line = detect_out_of_control(line_phase2, lcl_line, ucl_line)
                        for _, r in ooc_line.iterrows():
                            master_ooc_rows.append({"Color": c, "Factor": k, "Type": "LINE", "Batch": r["製造批號"], "Value": round(r["value"], 2), "Rule_CL": r["Rule_CL"], "Rule_3Sigma": r["Rule_3Sigma"]})
                            color_has_ooc = True

                    lcl_lab, ucl_lab = safe_get_limit(c, "LAB", k)
                    if lcl_lab is not None and ucl_lab is not None:
                        lab_phase2 = spc_c[k]["lab"][spc_c[k]["lab"]["Time"] >= c_time]
                        ooc_lab = detect_out_of_control(lab_phase2, lcl_lab, ucl_lab)
                        for _, r in ooc_lab.iterrows():
                            master_ooc_rows.append({"Color": c, "Factor": k, "Type": "LAB", "Batch": r["製造批號"], "Value": round(r["value"], 2), "Rule_CL": r["Rule_CL"], "Rule_3Sigma": r["Rule_3Sigma"]})
                            color_has_ooc = True

                if color_has_ooc:
                    report_data_dict[c] = {"spc": spc_c, "cb_code": cb_code, "global_order": global_batch_order_c}

            if not master_ooc_rows:
                st.success("✅ Excellent! The system did not detect any out-of-control batches.")
            else:
                df_master_ooc = pd.DataFrame(master_ooc_rows)
                st.warning(f"⚠️ Detected **{len(df_master_ooc)}** OOC alerts across **{len(report_data_dict)}** different color codes.")
                st.dataframe(df_master_ooc, use_container_width=True)

                def generate_ai_insights(factor, factor_ooc, line_lim):
                    avg_val = factor_ooc["Value"].astype(float).mean()
                    
                    color_shift = "Unknown"
                    if factor == "ΔL": color_shift = "Lighter (White)" if avg_val > 0 else "Darker (Black)"
                    elif factor == "Δa": color_shift = "Redder" if avg_val > 0 else "Greener"
                    elif factor == "Δb": color_shift = "Yellower" if avg_val > 0 else "Bluer"
                    
                    direction = "upper"
                    if line_lim[1] is not None and avg_val > line_lim[1]: direction = "upper"
                    elif line_lim[0] is not None and avg_val < line_lim[0]: direction = "lower"
                    
                    analysis = f"Statistical analysis indicates that the '{factor}' factor exhibits a significant deviation towards the {direction} control limit, resulting in a '{color_shift}' visual shift on the production line. A total of {len(factor_ooc)} recent batches violated stability rules. "
                    
                    has_3sigma = any(factor_ooc["Rule_3Sigma"] == True)
                    if has_3sigma:
                        analysis += "The presence of 3-Sigma (fluctuation) violations indicates an unstable process variation, likely caused by inconsistent machine parameters or uneven film thickness."
                    else:
                        analysis += "The violations are strictly Control Limit breaches, suggesting a systematic mean shift in the baseline formulation."
                        
                    actions = (
                        f"1. Formulation Adjustment: Request the Lab to apply a counter-offset to the baseline formula to mitigate the '{color_shift}' drift.\n"
                        f"2. Process Audit: Verify film thickness consistency on the production line, as thickness variations strongly correlate with {factor} drifts.\n"
                        f"3. Material Traceability: Isolate and inspect the raw material lots used in the highlighted OOC batches for potential contamination."
                    )
                    return analysis, actions

                doc = docx.Document()
                doc.add_heading('System-Wide Out-of-Control (OOC) Report', 0)
                doc.add_paragraph(f'Report generated on: {datetime.now().strftime("%Y-%m-%d %H:%M")}')
                doc.add_paragraph(f'Total OOC Instances: {len(df_master_ooc)}')
                doc.add_page_break()

                for color_name, data in report_data_dict.items():
                    doc.add_heading(f'Color Code: {color_name}', level=1)
                    color_ooc = df_master_ooc[df_master_ooc["Color"] == color_name]

                    def spc_combined_word(lab, line, title, lab_lim, line_lim, control_batch_code, glb_order):
                        fig, ax = plt.subplots(figsize=(10, 4.5))
                        ax.set_facecolor('#f2f2f2')

                        valid_batches = set(lab["製造批號"]).intersection(set(line["製造批號"]))
                        local_order = [b for b in glb_order if b in valid_batches]

                        lab_dict = dict(zip(lab["製造批號"], lab["value"]))
                        line_dict = dict(zip(line["製造批號"], line["value"]))

                        lab_x, lab_y = [], []
                        line_x, line_y = [], []

                        for i, b in enumerate(local_order):
                            if b in lab_dict and b in line_dict:
                                lab_x.append(i)
                                lab_y.append(lab_dict[b])
                                line_x.append(i)
                                line_y.append(line_dict[b])

                        if lab_x: ax.plot(lab_x, lab_y, marker="^", color="#548235", linestyle="-", linewidth=1.5, markersize=7, label="LAB Input")
                        if line_x: ax.plot(line_x, line_y, marker="o", color="#ffc000", linestyle="-", linewidth=1.5, markersize=7, markerfacecolor="white", markeredgewidth=1.5, label="LINE Output")

                        import matplotlib.transforms as transforms
                        trans = transforms.blended_transform_factory(ax.transData, ax.transAxes)

                        if control_batch_code is not None and control_batch_code in glb_order:
                            global_ctrl_idx = glb_order.index(control_batch_code)
                            ctrl_x = None
                            for i, b in enumerate(local_order):
                                if glb_order.index(b) >= global_ctrl_idx:
                                    ctrl_x = i
                                    break
                            if ctrl_x is not None:
                                ax.axvline(x=ctrl_x, color="#000000", linestyle=(0, (3, 3)), linewidth=1.5)
                                ax.text(ctrl_x, 1.02, "  After Control", color="#0070c0", fontsize=12, ha="left", va="bottom", transform=trans)
                                ax.text(ctrl_x, 1.02, "Before Control  ", color="#0070c0", fontsize=12, ha="right", va="bottom", transform=trans)

                        if lab_lim[0] is not None and lab_lim[1] is not None and len(lab_y) > 0:
                            lab_y_arr = np.array(lab_y); lab_x_arr = np.array(lab_x)
                            out_lab = (lab_y_arr > lab_lim[1]) | (lab_y_arr < lab_lim[0])
                            if out_lab.any(): ax.plot(lab_x_arr[out_lab], lab_y_arr[out_lab], marker="^", color="red", linestyle="None", markersize=9, zorder=5)
                        
                        if line_lim[0] is not None and line_lim[1] is not None and len(line_y) > 0:
                            line_y_arr = np.array(line_y); line_x_arr = np.array(line_x)
                            out_line = (line_y_arr > line_lim[1]) | (line_y_arr < line_lim[0])
                            if out_line.any(): ax.plot(line_x_arr[out_line], line_y_arr[out_line], marker="o", color="red", linestyle="None", markersize=9, zorder=5)

                        if lab_lim[0] is not None: 
                            ax.axhline(lab_lim[0], color="#7030a0", linestyle="-", linewidth=2, label="LAB LCL")
                            ax.axhline(lab_lim[1], color="#7030a0", linestyle="-", linewidth=2, label="LAB UCL")
                        if line_lim[0] is not None: 
                            ax.axhline(line_lim[0], color="#ff0000", linestyle="--", linewidth=2, label="LINE LCL")
                            ax.axhline(line_lim[1], color="#ff0000", linestyle="--", linewidth=2, label="LINE UCL")

                        if local_order:
                            ax.set_xticks(range(len(local_order)))
                            ax.set_xticklabels(local_order, rotation=45)

                        ax.set_title(title, fontsize=13, fontweight="bold", pad=20)
                        ax.grid(axis="y", color="#cccccc", linestyle="-", linewidth=1)
                        ax.grid(axis="x", visible=False)
                        ax.legend(bbox_to_anchor=(1.02, 1), loc="upper left", frameon=True, edgecolor="black")
                        fig.subplots_adjust(right=0.75, top=0.8)
                        return fig

                    for factor in ["ΔL", "Δa", "Δb"]:
                        factor_ooc = color_ooc[color_ooc["Factor"] == factor]
                        if not factor_ooc.empty:
                            doc.add_heading(f'Factor: {factor}', level=2)

                            lab_lim = safe_get_limit(color_name, "LAB", factor)
                            line_lim = safe_get_limit(color_name, "LINE", factor)

                            lab_str = f"[{lab_lim[0]:.2f} ~ {lab_lim[1]:.2f}]" if lab_lim[0] is not None else "Not Configured"
                            line_str = f"[{line_lim[0]:.2f} ~ {line_lim[1]:.2f}]" if line_lim[0] is not None else "Not Configured"
                            
                            limit_p = doc.add_paragraph()
                            limit_p.add_run("📌 Current Control Limits (Spec):\n").bold = True
                            limit_p.add_run(f"   • LAB Target: {lab_str}\n   • LINE Target: {line_str}")

                            table = doc.add_table(rows=1, cols=4)
                            table.style = 'Table Grid'
                            hdr = table.rows[0].cells
                            hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text = 'Type', 'Batch No.', 'Value', 'Violated Rule'

                            for _, r in factor_ooc.iterrows():
                                row_cells = table.add_row().cells
                                row_cells[0].text = str(r["Type"])
                                row_cells[1].text = str(r["Batch"])
                                val = float(r["Value"])
                                row_cells[2].text = str(int(val)) if val.is_integer() else str(val)
                                
                                rules = []
                                if r["Rule_CL"]: rules.append("Control Limit")
                                if r["Rule_3Sigma"]: rules.append("3-Sigma")
                                row_cells[3].text = " + ".join(rules)

                            doc.add_paragraph("\n📈 Control Chart Reference:")

                            cb_code = data["cb_code"]
                            spc_c = data["spc"]
                            glb_order_c = data["global_order"]

                            fig = spc_combined_word(spc_c[factor]["lab"], spc_c[factor]["line"], f"{color_name} - COMBINED {factor}", lab_lim, line_lim, cb_code, glb_order_c)
                            img_buf = io.BytesIO()
                            fig.savefig(img_buf, format='png', dpi=150, bbox_inches='tight')
                            img_buf.seek(0)
                            doc.add_picture(img_buf, width=Inches(6.0))
                            plt.close(fig)

                            analysis_text, action_text = generate_ai_insights(factor, factor_ooc, line_lim)
                            
                            doc.add_heading('🤖 AI Process Analysis', level=3)
                            doc.add_paragraph(analysis_text)
                            
                            doc.add_heading('🛠️ Recommended Corrective Actions', level=3)
                            doc.add_paragraph(action_text)

                    doc.add_page_break()

                report_buf = io.BytesIO()
                doc.save(report_buf)
                report_buf.seek(0)

                st.success("✅ Word file successfully compiled with AI Insights!")
                st.download_button(
                    label="📥 Download System-Wide Word Report",
                    data=report_buf,
                    file_name=f"System_OOC_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
# =========================================================
# VIEW 6: I-MR CHART (COIL-LEVEL)
# =========================================================
elif app_mode == "📈 I-MR Chart (Coil-Level)":
    st.title("📈 Individual - Moving Range (I-MR) Chart")
    st.markdown("Coil-level statistical process control analysis without batch aggregation.")

    st.subheader("⚙️ Configuration")
    col_cfg1, col_cfg2 = st.columns(2)
    with col_cfg1:
        source_opt = st.radio("Select Data Source:", ["LINE", "LAB"], horizontal=True, key="imr_source")
    with col_cfg2:
        factor_opt = st.selectbox("Select Color Factor:", ["ΔL", "Δa", "Δb", "ΔE"], format_func=delta_label, key="imr_factor")

    # Map the correct column based on user selection
    target_col = None
    df_imr_base = df.copy()

    if source_opt == "LINE":
        if factor_opt == "ΔE":
            st.warning("⚠️ ΔE* is not directly pre-calculated for LINE in raw data. Please select ΔL*, Δa*, or Δb*.")
        else:
            col_n, col_s = f"正-北 {factor_opt}", f"正-南 {factor_opt}"
            if col_n in df_imr_base.columns and col_s in df_imr_base.columns:
                df_imr_base["Coil_Value"] = df_imr_base[[col_n, col_s]].mean(axis=1)
                target_col = "Coil_Value"
            else:
                st.error(f"❌ Missing LINE columns for {factor_opt}")
    else: # LAB
        if factor_opt == "ΔE":
            target_col = "Average value ΔE 正面"
        else:
            target_col = f"入料檢測 {factor_opt} 正面"
            if target_col not in df_imr_base.columns:
                target_col = f"Average value {factor_opt} 正面"

    if target_col and target_col in df_imr_base.columns:
        # Strict Coil-Level Filtering: Drop missing values, sort chronologically by time
        df_imr = df_imr_base.dropna(subset=["Coil No.", "Time", target_col]).sort_values("Time").copy()
        
        # Ensure 1 row = 1 unique coil (take the first chronological record per coil)
        df_imr = df_imr.drop_duplicates(subset=["Coil No."], keep="first").reset_index(drop=True)

        if len(df_imr) < 2:
            st.warning("⚠️ Not enough coil data to generate an I-MR chart (minimum 2 coils required).")
        else:
            # ---------------------------------------------------------
            # STATISTICAL CALCULATIONS (I-MR for n=2)
            # ---------------------------------------------------------
            df_imr["MR"] = df_imr[target_col].diff().abs()
            
            mean_x = df_imr[target_col].mean()
            mean_mr = df_imr["MR"].mean()

            # SPC Constants for Moving Range of size n=2
            E2 = 2.66
            D4 = 3.267

            ucl_i = mean_x + E2 * mean_mr
            lcl_i = mean_x - E2 * mean_mr
            ucl_mr = D4 * mean_mr

            # ---------------------------------------------------------
            # ---------------------------------------------------------
            # INDIVIDUAL (I) CHART PLOTTING
            # ---------------------------------------------------------
            # Get date range for the title
            t_min = df_imr["Time"].min().strftime("%Y-%m-%d")
            t_max = df_imr["Time"].max().strftime("%Y-%m-%d")
            date_str = f"{t_min} to {t_max}"

            # Fetch Spec Limits from Google Sheet
            spec_lcl, spec_ucl = safe_get_limit(color, source_opt, factor_opt)

            st.markdown("---")
            st.markdown(f"### 📊 Individual (I) Chart - {factor_opt} ({source_opt})")
            
            # Kích thước rộng hơn để dàn trải data đều và chuyên nghiệp
            fig_i, ax_i = plt.subplots(figsize=(14, 6))
            ax_i.set_facecolor('#f2f2f2')
            
            # 1. Plot actual data points (Base line) - Sử dụng màu Xanh dương đậm (#003366)
            ax_i.plot(df_imr["Coil No."], df_imr[target_col], marker='o', color='#003366', linestyle='-', linewidth=1.5, markersize=6, label="Individual Value")
            
            # 2. Highlight Out-of-Control (3-Sigma) points in RED
            ooc_i_mask = (df_imr[target_col] > ucl_i) | (df_imr[target_col] < lcl_i)
            if ooc_i_mask.any():
                ax_i.scatter(df_imr[ooc_i_mask]["Coil No."], df_imr[ooc_i_mask][target_col], color="red", s=80, zorder=5, label="🚨 Out of Control (3σ)")

            # 3. Plot Statistical Limits (Calculated from Data)
            ax_i.axhline(mean_x, color="DeepSkyBlue", linestyle="-", linewidth=2.5, label=f"Mean (X) = {mean_x:.3f}")
            ax_i.axhline(ucl_i, color="#d62728", linestyle="--", linewidth=1.5, label=f"Stat UCL (3σ) = {ucl_i:.3f}")
            ax_i.axhline(lcl_i, color="#d62728", linestyle="--", linewidth=1.5, label=f"Stat LCL (3σ) = {lcl_i:.3f}")
            
            # 4. Plot Specification Limits (From Google Sheet)
            if spec_lcl is not None:
                ax_i.axhline(spec_lcl, color="#7030a0", linestyle="-", linewidth=2, label=f"Spec LCL = {spec_lcl}")
            if spec_ucl is not None:
                ax_i.axhline(spec_ucl, color="#7030a0", linestyle="-", linewidth=2, label=f"Spec UCL = {spec_ucl}")

            # Smart X-Axis Ticks
            if len(df_imr) > 40:
                ax_i.set_xticks(ax_i.get_xticks()[::len(df_imr)//30])
            ax_i.set_xticklabels(ax_i.get_xticklabels(), rotation=45, ha='right', fontsize=9)

            # Title and Labels
            ax_i.set_title(f"I-Chart: {color} | {factor_opt} ({source_opt})\nTimeframe: {date_str}", fontsize=15, fontweight="bold", pad=15)
            ax_i.set_ylabel("Measured Value", fontsize=11)
            ax_i.grid(axis="y", color="#cccccc", linestyle="-", linewidth=1)
            
            # Legend Configuration
            ax_i.legend(bbox_to_anchor=(1.02, 1), loc="upper left", frameon=True, edgecolor="black")
            fig_i.subplots_adjust(right=0.8, bottom=0.2)
            
            st.pyplot(fig_i)
            plt.close(fig_i)

            # ---------------------------------------------------------
            # MOVING RANGE (MR) CHART PLOTTING
            # ---------------------------------------------------------
            st.markdown(f"### 📉 Moving Range (MR) Chart - {factor_opt} ({source_opt})")
            
            # Kích thước đồng bộ với I-Chart
            fig_mr, ax_mr = plt.subplots(figsize=(14, 5))
            ax_mr.set_facecolor('#f2f2f2')
            
            # 1. Plot actual data points (Base line) - Sử dụng màu Xanh dương đậm (#003366)
            ax_mr.plot(df_imr["Coil No."], df_imr["MR"], marker='s', color='#003366', linestyle='-', linewidth=1.5, markersize=6, label="Moving Range")
            
            # 2. Highlight Out-of-Control points in RED
            ooc_mr_mask = (df_imr["MR"] > ucl_mr)
            if ooc_mr_mask.any():
                ax_mr.scatter(df_imr[ooc_mr_mask]["Coil No."], df_imr[ooc_mr_mask]["MR"], color="red", marker='s', s=80, zorder=5, label="🚨 Out of Control (3σ)")

            # 3. Plot Limits
            ax_mr.axhline(mean_mr, color="DeepSkyBlue", linestyle="-", linewidth=2, label=f"Mean (MR) = {mean_mr:.3f}")
            ax_mr.axhline(ucl_mr, color="#d62728", linestyle="--", linewidth=1.5, label=f"UCL = {ucl_mr:.3f}")
            ax_mr.axhline(0, color="#d62728", linestyle="--", linewidth=1.5, label="LCL = 0.000")

            if len(df_imr) > 40:
                ax_mr.set_xticks(ax_mr.get_xticks()[::len(df_imr)//30])
            ax_mr.set_xticklabels(ax_mr.get_xticklabels(), rotation=45, ha='right', fontsize=9)

            ax_mr.legend(bbox_to_anchor=(1.02, 1), loc="upper left", frameon=True, edgecolor="black")
            ax_mr.grid(axis="y", color="#cccccc", linestyle="-", linewidth=1)
            ax_mr.set_ylabel("Range |X(i) - X(i-1)|", fontsize=11)
            fig_mr.subplots_adjust(right=0.8, bottom=0.2)
            
            st.pyplot(fig_mr)
            plt.close(fig_mr)

            # ---------------------------------------------------------
            # COIL LEVEL DATA TABLE
            # ---------------------------------------------------------

            # ---------------------------------------------------------
            # COIL LEVEL DATA TABLE
            # ---------------------------------------------------------
            with st.expander("📋 Coil-Level Statistical Data"):
                display_cols = ["Time", "製造批號", "Coil No.", target_col, "MR"]
                st.dataframe(
                    df_imr[display_cols].rename(columns={
                        "Time": "Timestamp",
                        "製造批號": "Batch No.",
                        target_col: f"{delta_label(factor_opt)} Value"
                    }).style.format({f"{delta_label(factor_opt)} Value": "{:.3f}", "MR": "{:.3f}"}, na_rep="-"),
                    use_container_width=True,
                    hide_index=True
                )

# =========================================================
# NEW VIEW: KPI ACHIEVEMENT REPORT – CONTROLLED COLORS
# OOC KPI ONLY (NO Cpk – no true LSL/USL specification limits available)
# =========================================================
elif app_mode == "🎯 KPI Achievement Report":
    st.title("🎯 SPC OOC KPI Report – Controlled Colors")
    st.markdown(
        "Evaluate controlled colors using the management target: "
        "**Post-control OOC Rate ≤ target**."
    )

    st.info(
        "**KPI logic:** The **Control Start Batch itself is excluded**; KPI evaluation starts from the "
        "**next batch onward**. A post-control batch is counted as OOC only once even if multiple factors "
        "(ΔL* / Δa* / Δb*) exceed the configured control limits. **Value = LCL or Value = UCL remains IN CONTROL**; "
        "OOC is counted only when **Value < LCL** or **Value > UCL**. The original tasks and their 3-Sigma logic "
        "remain unchanged. This KPI report uses only configured LCL/UCL breaches."
    )

    col_cfg1, col_cfg2 = st.columns(2)
    with col_cfg1:
        report_source = st.radio(
            "Data Source",
            ["LINE", "LAB"],
            horizontal=True,
            index=0,
            key="kpi_report_source"
        )
    with col_cfg2:
        ooc_target = st.number_input(
            "OOC Rate Target (%)",
            min_value=0.0,
            max_value=100.0,
            value=5.0,
            step=0.5,
            key="kpi_ooc_target"
        )

    source_key = report_source.lower()
    factors = ["ΔL", "Δa", "Δb"]

    report_rows = []
    detail_store = {}
    all_colors_report = sorted(df_raw["塗料編號"].dropna().unique())

    for c in all_colors_report:
        df_c = df_raw[df_raw["塗料編號"] == c].sort_values("Time").copy()
        if df_c.empty:
            continue

        cb = get_control_batch(c)
        cb_code = get_control_batch_code(df_c, cb)
        c_time = df_c[df_c["製造批號"] == cb_code]["Time"].min() if cb_code else None

        # Only colors with an actual control start point are treated as controlled colors.
        if cb is None or cb_code is None or c_time is None:
            continue

        # IMPORTANT: the control start batch is only the boundary.
        # KPI counting starts from the NEXT batch.
        batch_order_c = df_c.sort_values("Time")["製造批號"].drop_duplicates().tolist()
        try:
            control_idx = batch_order_c.index(cb_code)
            post_control_batches = batch_order_c[control_idx + 1:]
        except ValueError:
            post_control_batches = []

        df_p2 = df_c[df_c["製造批號"].isin(post_control_batches)].copy()
        if df_p2.empty:
            continue

        spc_p2 = calculate_batch_averages(df_p2[SPC_INPUT_COLUMNS])

        # Use only batches with valid values for all 3 color factors so the
        # denominator is consistent across ΔL / Δa / Δb.
        factor_batch_sets = []
        factor_frames = {}
        for f in factors:
            fdf = spc_p2[f][source_key][["製造批號", "Time", "value"]].dropna(subset=["value"]).copy()
            factor_frames[f] = fdf
            factor_batch_sets.append(set(fdf["製造批號"].tolist()))

        valid_batches = set.intersection(*factor_batch_sets) if factor_batch_sets and all(factor_batch_sets) else set()
        n_valid = len(valid_batches)
        post_control_raw_batches = df_p2["製造批號"].nunique()

        unique_ooc_batches = set()
        cl_complete = True
        factor_ooc_counts = {}
        ooc_trigger_rows = []
        control_limit_columns = {}

        for f in factors:
            lcl, ucl, lcl_col_used, ucl_col_used = get_kpi_control_limit_exact(c, report_source, f)
            control_limit_columns[f] = {
                "LCL column": lcl_col_used,
                "UCL column": ucl_col_used,
                "LCL": lcl,
                "UCL": ucl,
            }

            if lcl is None or ucl is None:
                cl_complete = False

            fdf = factor_frames[f]
            if valid_batches:
                fdf = fdf[fdf["製造批號"].isin(valid_batches)].copy()
            else:
                fdf = fdf.iloc[0:0].copy()

            if not fdf.empty and lcl is not None and ucl is not None:
                vals = pd.to_numeric(fdf["value"], errors="coerce")

                # Boundary rule: equal to LCL/UCL is IN CONTROL.
                # Tiny tolerance only avoids floating-point artifacts.
                tol = 1e-9
                below_lcl = vals < (float(lcl) - tol)
                above_ucl = vals > (float(ucl) + tol)
                ooc_mask_kpi = below_lcl | above_ucl

                ooc_f = fdf[ooc_mask_kpi].copy()
                factor_ooc_counts[f] = ooc_f["製造批號"].nunique()
                unique_ooc_batches.update(ooc_f["製造批號"].astype(str).tolist())

                for _, rr in ooc_f.iterrows():
                    vv = float(rr["value"])
                    side = "Below LCL" if vv < (float(lcl) - tol) else "Above UCL"
                    distance = (float(lcl) - vv) if side == "Below LCL" else (vv - float(ucl))
                    ooc_trigger_rows.append({
                        "Factor": f,
                        "Batch No.": str(rr["製造批號"]),
                        "Time": rr.get("Time", pd.NaT),
                        "Batch Average": vv,
                        "LCL": float(lcl),
                        "UCL": float(ucl),
                        "Trigger": side,
                        "Outside by": distance,
                    })
            else:
                factor_ooc_counts[f] = 0

        n_ooc = len(unique_ooc_batches)
        ooc_rate = (n_ooc / n_valid * 100) if n_valid > 0 else np.nan

        if not cl_complete:
            ooc_status = "⚪ CL Missing"
        elif pd.isna(ooc_rate):
            ooc_status = "⚪ No Data"
        elif ooc_rate <= ooc_target:
            ooc_status = "✅ Achieved"
        else:
            ooc_status = "❌ Not Achieved"

        report_rows.append({
            "Color Code": c,
            "Control Batch #": cb,
            "Control Start Batch": cb_code,
            "Post-Control Raw Batches": post_control_raw_batches,
            "SPC Valid Batches": n_valid,
            "OOC Batches": n_ooc,
            "OOC Rate (%)": round(ooc_rate, 2) if pd.notna(ooc_rate) else np.nan,
            "OOC KPI": ooc_status,
        })

        detail_store[c] = {
            "control_batch": cb,
            "control_batch_code": cb_code,
            "control_time": c_time,
            "raw_post_control_batches": post_control_raw_batches,
            "valid_batches": valid_batches,
            "ooc_batches": unique_ooc_batches,
            "ooc_rate": ooc_rate,
            "ooc_status": ooc_status,
            "factor_ooc_counts": factor_ooc_counts,
            "ooc_trigger_rows": ooc_trigger_rows,
            "control_limit_columns": control_limit_columns,
        }

    report_df = pd.DataFrame(report_rows)

    if report_df.empty:
        st.warning("No controlled colors with post-control data were found.")
    else:
        n_controlled = len(report_df)
        n_ooc_ok = int((report_df["OOC KPI"] == "✅ Achieved").sum())
        n_action = int((report_df["OOC KPI"] == "❌ Not Achieved").sum())
        n_incomplete = n_controlled - n_ooc_ok - n_action

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Controlled Colors", n_controlled)
        m2.metric(f"OOC ≤ {ooc_target:.1f}%", f"{n_ooc_ok}/{n_controlled}")
        m3.metric("Action Required", n_action)
        m4.metric("Incomplete / Missing CL", n_incomplete)

        st.markdown("---")
        st.markdown("### 📋 Controlled Color OOC KPI Summary")

        def _kpi_row_style(row):
            if row.get("OOC KPI") == "✅ Achieved":
                return ["background-color: #e8f5e9"] * len(row)
            if row.get("OOC KPI") == "❌ Not Achieved":
                return ["background-color: #ffebee"] * len(row)
            return ["background-color: #fff8e1"] * len(row)

        display_report_df = report_df.sort_values(
            by=["OOC Rate (%)", "Color Code"],
            ascending=[False, True],
            na_position="last"
        ).reset_index(drop=True)

        st.dataframe(
            display_report_df.style
                .apply(_kpi_row_style, axis=1)
                .format({"OOC Rate (%)": "{:.2f}"}, na_rep="-"),
            use_container_width=True,
            hide_index=True
        )

        csv_buf = display_report_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "📥 Download OOC KPI Summary CSV",
            data=csv_buf,
            file_name=f"SPC_OOC_KPI_{report_source}_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            key="download_kpi_report_csv"
        )

        # ---------------------------------------------------------
        # SHORT WORD REPORT FOR MANAGEMENT
        # Only added to this KPI task; existing tasks/charts unchanged.
        # ---------------------------------------------------------
        def build_ooc_kpi_word_report(summary_df, source_name, target_pct,
                                      controlled_count, achieved_count,
                                      action_count, incomplete_count):
            doc = docx.Document()

            # Compact page layout
            section = doc.sections[0]
            section.top_margin = Inches(0.55)
            section.bottom_margin = Inches(0.55)
            section.left_margin = Inches(0.55)
            section.right_margin = Inches(0.55)

            # Font setup for Traditional Chinese / English
            styles = doc.styles
            styles["Normal"].font.name = "Microsoft JhengHei"
            styles["Normal"].font.size = docx.shared.Pt(10)
            for style_name in ["Title", "Heading 1", "Heading 2"]:
                if style_name in styles:
                    styles[style_name].font.name = "Microsoft JhengHei"

            title = doc.add_heading("SPC 管制成效 KPI 摘要報告", level=0)
            title.alignment = 1

            p = doc.add_paragraph()
            p.add_run("資料來源：").bold = True
            p.add_run(str(source_name))
            p.add_run("    報告日期：").bold = True
            p.add_run(datetime.now().strftime("%Y-%m-%d"))

            # 1. Objective / rule
            doc.add_heading("一、評估目標", level=1)
            p = doc.add_paragraph()
            p.add_run(f"管理目標：管制後 OOC Rate ≤ {target_pct:.1f}% 。").bold = True
            doc.add_paragraph(
                "Control Start Batch 僅作為管制起始點，不納入 KPI 計算；"
                "僅計算其後批次。數值等於 LCL 或 UCL 視為管制內，"
                "僅當數值 < LCL 或 > UCL 時判定為 OOC。"
            )

            # 2. Overall summary
            doc.add_heading("二、整體結果", level=1)
            summary_table = doc.add_table(rows=2, cols=4)
            summary_table.style = "Table Grid"
            headers = ["已管制色號", f"OOC ≤ {target_pct:.1f}%", "需改善", "資料/界限不完整"]
            values = [str(controlled_count), str(achieved_count), str(action_count), str(incomplete_count)]
            for j, h in enumerate(headers):
                summary_table.rows[0].cells[j].text = h
                summary_table.rows[1].cells[j].text = values[j]

            # 3. Color summary table
            doc.add_heading("三、色號 KPI 結果", level=1)
            word_cols = [
                "Color Code", "Control Start Batch", "SPC Valid Batches",
                "OOC Batches", "OOC Rate (%)", "OOC KPI"
            ]
            word_df = summary_df[word_cols].copy()
            word_df = word_df.sort_values(
                by=["OOC Rate (%)", "Color Code"],
                ascending=[False, True],
                na_position="last"
            )

            table = doc.add_table(rows=1, cols=len(word_cols))
            table.style = "Table Grid"
            zh_headers = ["色號", "管制起始批號", "有效批數", "OOC批數", "OOC率(%)", "判定"]
            for j, h in enumerate(zh_headers):
                table.rows[0].cells[j].text = h

            for _, rr in word_df.iterrows():
                cells = table.add_row().cells
                cells[0].text = str(rr["Color Code"])
                cells[1].text = str(rr["Control Start Batch"])
                cells[2].text = str(int(rr["SPC Valid Batches"]))
                cells[3].text = str(int(rr["OOC Batches"]))
                cells[4].text = "-" if pd.isna(rr["OOC Rate (%)"]) else f"{float(rr['OOC Rate (%)']):.2f}"
                status = str(rr["OOC KPI"])
                if "Achieved" in status and "Not" not in status:
                    cells[5].text = "達成"
                elif "Not Achieved" in status:
                    cells[5].text = "未達成"
                elif "CL Missing" in status:
                    cells[5].text = "界限缺失"
                else:
                    cells[5].text = "資料不足"

            # 4. Short management conclusion
            doc.add_heading("四、結論", level=1)
            if action_count == 0 and incomplete_count == 0:
                conclusion = (
                    f"本次共評估 {controlled_count} 個已管制色號，全部符合 "
                    f"OOC Rate ≤ {target_pct:.1f}% 之管理目標，整體管制成效符合預期。"
                )
            else:
                parts = [
                    f"本次共評估 {controlled_count} 個已管制色號，"
                    f"其中 {achieved_count} 個達成 OOC Rate ≤ {target_pct:.1f}% 目標。"
                ]
                if action_count > 0:
                    not_ok = word_df[word_df["OOC KPI"] == "❌ Not Achieved"]
                    action_text = ", ".join(
                        f"{r['Color Code']} ({float(r['OOC Rate (%)']):.2f}%)"
                        for _, r in not_ok.iterrows() if pd.notna(r["OOC Rate (%)"])
                    )
                    if action_text:
                        parts.append(f"未達目標色號：{action_text}，建議列為後續優先追蹤對象。")
                if incomplete_count > 0:
                    parts.append(f"另有 {incomplete_count} 個色號因資料或管制界限不完整，暫無法完成 KPI 判定。")
                conclusion = "".join(parts)

            doc.add_paragraph(conclusion)

            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            return buf

        # Word creation is deferred until requested to avoid rebuilding on every filter rerun.
        if st.button("📄 Prepare Short Word Report for Management", key="prepare_kpi_management_word"):
            with st.spinner("Preparing management report..."):
                st.session_state["kpi_word_payload"] = build_ooc_kpi_word_report(
                    display_report_df, report_source, ooc_target,
                    n_controlled, n_ooc_ok, n_action, n_incomplete
                ).getvalue()
                st.session_state["kpi_word_identity"] = (report_source, ooc_target, len(display_report_df))
        if (st.session_state.get("kpi_word_payload") and
                st.session_state.get("kpi_word_identity") == (report_source, ooc_target, len(display_report_df))):
            st.download_button(
                "📄 Download Short Word Report for Management",
                data=st.session_state["kpi_word_payload"],
                file_name=f"SPC_OOC_KPI_Management_Report_{report_source}_{datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="download_kpi_management_word"
            )

        st.markdown("---")
        st.markdown("### 🔎 Color Detail")
        selected_kpi_color = st.selectbox(
            "Select controlled color",
            display_report_df["Color Code"].tolist(),
            key="kpi_detail_color"
        )

        d = detail_store[selected_kpi_color]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Control Start Batch", str(d["control_batch_code"]))
        c2.metric("Post-Control Valid Batches", len(d["valid_batches"]))
        c3.metric("OOC Batches", len(d["ooc_batches"]))
        c4.metric(
            "OOC Rate",
            f"{d['ooc_rate']:.2f}%" if pd.notna(d["ooc_rate"]) else "N/A",
            delta=d["ooc_status"]
        )

        factor_detail_rows = []
        for f in factors:
            lcl, ucl, lcl_col_used, ucl_col_used = get_kpi_control_limit_exact(
                selected_kpi_color, report_source, f
            )
            factor_detail_rows.append({
                "Factor": f,
                "SPC LCL": lcl,
                "SPC UCL": ucl,
                "LCL Source Column": lcl_col_used,
                "UCL Source Column": ucl_col_used,
                "Factor OOC Batches": d["factor_ooc_counts"].get(f, 0),
            })

        factor_detail_df = pd.DataFrame(factor_detail_rows)
        st.dataframe(
            factor_detail_df.style.format({
                "SPC LCL": "{:.3f}",
                "SPC UCL": "{:.3f}",
            }, na_rep="-"),
            use_container_width=True,
            hide_index=True
        )

        st.markdown("#### 🚨 OOC Trigger Audit (KPI Only)")
        if d.get("ooc_trigger_rows"):
            audit_df = pd.DataFrame(d["ooc_trigger_rows"]).sort_values(["Time", "Factor", "Batch No."])
            st.dataframe(
                audit_df.style.format({
                    "Batch Average": "{:.10f}",
                    "LCL": "{:.10f}",
                    "UCL": "{:.10f}",
                    "Outside by": "{:.10f}",
                }),
                use_container_width=True,
                hide_index=True,
            )
            st.warning(
                "Only batches AFTER the Control Start Batch listed above are counted in the KPI OOC numerator. "
                "If this table is empty, OOC Batches = 0."
            )
        else:
            st.success("✅ No post-control batch is strictly outside the configured LCL/UCL for this source.")

        st.caption(
            "**Interpretation:** OOC Rate = unique post-control OOC batches ÷ valid post-control batches × 100%. "
            "The Control Start Batch itself is excluded. Values exactly equal to LCL or UCL are treated as in control."
        )

