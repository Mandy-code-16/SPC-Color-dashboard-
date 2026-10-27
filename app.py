import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import numpy as np
import math
import re
import docx
from docx.shared import Inches
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

@st.cache_data(show_spinner=False)
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

required_data_cols = ["Time", "製造批號", "塗料編號"]
missing_data_cols = [c for c in required_data_cols if c not in df_raw.columns]
if missing_data_cols:
    st.error(
        "❌ File 1 (Production / SPC Data) is missing required columns: "
        + ", ".join(missing_data_cols)
    )
    st.stop()

required_limit_cols = ["Color_code"]
missing_limit_cols = [c for c in required_limit_cols if c not in limit_df.columns]
if missing_limit_cols:
    st.error(
        "❌ File 2 (Control Limit Data) is missing required columns: "
        + ", ".join(missing_limit_cols)
    )
    st.stop()

# Parse time after upload.
df_raw["Time"] = pd.to_datetime(df_raw["Time"], errors="coerce")
df_raw = df_raw.dropna(subset=["Time"])

st.sidebar.success(
    f"✅ Data loaded\n\n"
    f"File 1: {uploaded_data_file.name}\n\n"
    f"File 2: {uploaded_limit_file.name}"
)

# CRITICAL FIX: Remove NA, NaN, and Empty spaces from Batch Number globally
df_raw = df_raw.dropna(subset=["製造批號"])
df_raw["製造批號"] = df_raw["製造批號"].astype(str).str.strip()
df_raw = df_raw[df_raw["製造批號"] != ""]
df_raw = df_raw[df_raw["製造批號"].str.lower() != "nan"]

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

def calculate_batch_averages(df_filtered_color):
    res = {}
    for f in ["ΔL", "Δa", "Δb"]:
        tmp = df_filtered_color.copy()
        
        # Calculate LINE average
        col_n, col_s = f"正-北 {f}", f"正-南 {f}"
        tmp[col_n] = pd.to_numeric(tmp[col_n], errors='coerce')
        tmp[col_s] = pd.to_numeric(tmp[col_s], errors='coerce')
        tmp["row_avg"] = tmp[[col_n, col_s]].mean(axis=1)
        line_b = tmp.groupby("製造批號", as_index=False).agg({"Time": "min", "row_avg": "mean"}).rename(columns={"row_avg": "value"}).sort_values("Time").dropna()
        
        # Calculate LAB average
        col_lab = f"入料檢測 {f} 正面"
        tmp[col_lab] = pd.to_numeric(tmp[col_lab], errors='coerce')
        lab_b = tmp.groupby("製造批號", as_index=False).agg({"Time": "min", col_lab: "mean"}).rename(columns={col_lab: "value"}).sort_values("Time").dropna()
        
        # CRITICAL FIX: Enforce Strict 1:1 Intersection
        common_batches = set(line_b["製造批號"]).intersection(set(lab_b["製造批號"]))
        line_b = line_b[line_b["製造批號"].isin(common_batches)]
        lab_b = lab_b[lab_b["製造批號"].isin(common_batches)]
        
        res[f] = {"line": line_b, "lab": lab_b}
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

spc_data = calculate_batch_averages(df)

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
        for c in table.columns: table[c] = table[c].map(lambda x: f"{x:.2f}" if pd.notnull(x) else "")
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
        st.pyplot(fig); download(fig, f"COMBINED_{color}_{k}.png")

    st.markdown("---")
    st.subheader("📊 SPC Combined Chart (LAB + LINE) – Phase II")
    for k in ["ΔL", "Δa", "Δb"]:
        lab_lim = safe_get_limit(color, "LAB", k)
        line_lim = safe_get_limit(color, "LINE", k)
        fig = spc_combined_phase2(spc_data[k]["lab"], spc_data[k]["line"], f"{delta_label(k)} – LAB + LINE (Phase II)", lab_lim, line_lim, control_batch_code, control_time, global_batch_order)
        if fig is not None: st.pyplot(fig); download(fig, f"COMBINED_PHASE2_{color}_{k}.png")
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

            st.subheader("📊 Average Thickness Distribution")
            data = df_plot[thickness_col].dropna()
            if len(data) > 0:
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

                            st.markdown("### 🧠 Interpretation")
                            st.markdown("""
                            **📊 Correlation Levels Reference:**
                            
                            | Level | Correlation Coefficient (\|R\|) | Coefficient of Determination (R²) | Interpretation |
                            | :--- | :--- | :--- | :--- |
                            | 🔴 **Strong** | \|R\| ≥ 0.77 | R² ≥ 0.60 | Thickness strongly affects and explains most of the color variation. |
                            | 🟠 **Moderate**| 0.55 ≤ \|R\| < 0.77 | 0.30 ≤ R² < 0.60 | Thickness partially contributes to color drift. |
                            | 🟢 **Weak/Low** | \|R\| < 0.55 | R² < 0.30 | Thickness is unlikely the main driver of color variation. |
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
                    spc_p2 = calculate_batch_averages(df_p2)
                    
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
            "Min": "{:.3f}",
            "Max": "{:.3f}",
            "平均值 (x̄)": "{:.3f}",
            "標準差 (σ)": "{:.3f}",
            "LCL (-1.5σ)": "{:.3f}",
            "UCL (+1.5σ)": "{:.3f}",
        }, na_rep="-"),
        hide_index=True,
        use_container_width=True,
    )

    # Download the Task 3 statistical summary table for the selected color
    _safe_color_name = re.sub(r'[^A-Za-z0-9._-]+', '_', str(color)).strip('_') or 'Color'

    _xlsx_buf = io.BytesIO()
    with pd.ExcelWriter(_xlsx_buf, engine="openpyxl") as writer:
        task3_stats_df.to_excel(writer, index=False, sheet_name="LAB_Input_Stats")
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

    _csv_buf = task3_stats_df.to_csv(index=False).encode("utf-8-sig")

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
                "LAB_ΔL*": "{:.3f}", "LINE_ΔL*": "{:.3f}",
                "LAB_Δa*": "{:.3f}", "LINE_Δa*": "{:.3f}",
                "LAB_Δb*": "{:.3f}", "LINE_Δb*": "{:.3f}"
            }, na_rep=""),
            hide_index=True,
            use_container_width=True
        )

    st.markdown("---")

    calc_res = {}
    chart_buffers = {}  # Store Task 3 chart PNG bytes for download + Word report
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
                f"the current LINE range [{line_lcl_sheet:.3f}, {line_ucl_sheet:.3f}]."
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
                f"{delta_label(f)} 入料": f"{q1:.3f}"
            },
            {
                "計算項目": "Q3 (75%)",
                f"{delta_label(f)} 入料": f"{q3:.3f}"
            },
            {
                "計算項目": "IQR (Q3 − Q1)",
                f"{delta_label(f)} 入料": f"{q3:.3f} − {q1:.3f} = {iqr_value:.3f}"
            },
            {
                "計算項目": f"控制範圍 LL (Q1 − {iqr_k:.2f}×IQR)",
                f"{delta_label(f)} 入料": f"{q1:.3f} − ({iqr_k:.2f} × {iqr_value:.3f}) = {iqr_lcl:.3f}"
            },
            {
                "計算項目": f"控制範圍 UL (Q3 + {iqr_k:.2f}×IQR)",
                f"{delta_label(f)} 入料": f"{q3:.3f} + ({iqr_k:.2f} × {iqr_value:.3f}) = {iqr_ucl:.3f}"
            },
        ])

        st.dataframe(
            _iqr_manual_df,
            hide_index=True,
            use_container_width=True,
        )

        st.info(
            f"IQR結果：{delta_label(f)} 入料 LL = {iqr_lcl:.3f}，UL = {iqr_ucl:.3f} （k = {iqr_k:.2f}）"
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
            fig.savefig(chart_png, format="png", dpi=220, bbox_inches="tight", facecolor="white")
            chart_png.seek(0)
            chart_bytes = chart_png.getvalue()
            chart_buffers[f] = chart_bytes

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
                f"{lab_lcl_sheet:.3f} ~ {lab_ucl_sheet:.3f}"
                if lab_lcl_sheet is not None and lab_ucl_sheet is not None else "Not set"
            )
            current_line = (
                f"{line_lcl_sheet:.3f} ~ {line_ucl_sheet:.3f}"
                if line_lcl_sheet is not None and line_ucl_sheet is not None else "Not set"
            )
            rec_txt = f"{rec_lcl:.3f} ~ {rec_ucl:.3f}" if has_overlap else "No overlap"

            result_table = pd.DataFrame([
                {"Item": "Current LAB limit", "Range / Value": current_lab},
                {"Item": "Current LINE limit", "Range / Value": current_line},
                {"Item": f"Standard LAB ({sig:.1f}σ)", "Range / Value": f"{std_lcl:.3f} ~ {std_ucl:.3f}"},
                {"Item": f"IQR LAB (k={iqr_k:.1f})", "Range / Value": f"{iqr_lcl:.3f} ~ {iqr_ucl:.3f}"},
                {"Item": "Recommended LAB", "Range / Value": rec_txt},
                {"Item": "Paired batches", "Range / Value": str(len(pair_f))},
                {"Item": "LAB data used", "Range / Value": str(len(safe_lab))},
                {"Item": "LAB→LINE r", "Range / Value": f"{r:.3f}" if pd.notna(r) else "N/A"},
                {"Item": "LAB→LINE R²", "Range / Value": f"{r2:.3f}" if pd.notna(r2) else "N/A"},
            ])
            st.dataframe(result_table, hide_index=True, use_container_width=True)
            st.caption(filter_note)

            if pd.notna(r2):
                if r2 >= 0.60:
                    st.error(f"🔴 Strong LAB→LINE relationship (R²={r2:.3f})")
                elif r2 >= 0.30:
                    st.warning(f"🟠 Moderate LAB→LINE relationship (R²={r2:.3f})")
                else:
                    st.success(f"🟢 Weak LAB→LINE relationship (R²={r2:.3f})")

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
                "LAB Mean": round(res["m"], 3),
                "LAB σ": round(res["s"], 3),
                "Standard LCL": round(res["std_lcl"], 3),
                "Standard UCL": round(res["std_ucl"], 3),
                "IQR LCL": round(res["iqr_lcl"], 3),
                "IQR UCL": round(res["iqr_ucl"], 3),
                "Recommended LAB LCL": round(res["rec_lcl"], 3) if has_rec else None,
                "Recommended LAB UCL": round(res["rec_ucl"], 3) if has_rec else None,
                "LAB→LINE r": round(res["r"], 3) if pd.notna(res["r"]) else None,
                "LAB→LINE R²": round(res["r2"], 3) if pd.notna(res["r2"]) else None,
            })

        incoming_summary = pd.DataFrame(summary_rows)
        st.dataframe(incoming_summary, hide_index=True, use_container_width=True)
        st.caption(
            "Recommended LAB limits are the intersection of Mean ± Kσ and IQR limits calculated from "
            "incoming LAB values associated with acceptable LINE outputs whenever current LINE limits are available."
        )

        dE_std = math.sqrt(dE_std_sq)
        dE_iqr = math.sqrt(dE_iqr_sq)
        dE_rec = math.sqrt(dE_rec_sq) if recommendation_valid else None

        # Existing project target for incoming LAB color difference.
        lab_de_target = 0.5

        st.markdown("### 🎯 Incoming LAB ΔE* Boundary Check")
        c1, c2, c3 = st.columns(3)
        c1.metric("Standard-derived ΔE*", f"{dE_std:.3f}")
        c2.metric("IQR-derived ΔE*", f"{dE_iqr:.3f}")
        c3.metric("Recommended ΔE*", f"{dE_rec:.3f}" if dE_rec is not None else "N/A")

        if dE_rec is None:
            st.warning("At least one factor has no overlap between Standard and IQR; review that factor before setting the incoming LAB limit.")
        elif dE_rec <= lab_de_target:
            st.success(
                f"✅ Recommended incoming LAB component limits produce a derived ΔE* boundary of {dE_rec:.3f}, "
                f"within the current LAB target ≤ {lab_de_target}."
            )
        else:
            st.warning(
                f"⚠ Recommended incoming LAB component limits produce a derived ΔE* boundary of {dE_rec:.3f}, "
                f"above the current LAB target {lab_de_target}. Further tightening should be reviewed."
            )

        st.markdown("### 💡 How to use this result")
        st.info(
            "Use the **Recommended LAB LCL/UCL** as the incoming paint control window. "
            "Continue monitoring the paired LINE output trend. If LINE OOC events remain concentrated at particular LAB zones, "
            "the incoming LAB window should be tightened or shifted using the paired historical relationship rather than treating LAB and LINE separately."
        )

    # ---------------------------------------------------------
    # Task 3 Word Report Export
    # ---------------------------------------------------------
    if calc_res:
        st.markdown("---")
        st.markdown("### 📄 Export Task 3 Word Report")
        st.caption(
            "The Word report includes the paired LAB–LINE analysis, calculated incoming LAB limits, "
            "correlation results, summary table, and the current ΔL* / Δa* / Δb* charts."
        )

        def build_task3_word_report():
            doc = docx.Document()

            # Page margins
            section = doc.sections[0]
            section.top_margin = Inches(0.65)
            section.bottom_margin = Inches(0.65)
            section.left_margin = Inches(0.7)
            section.right_margin = Inches(0.7)

            # Base font
            normal_style = doc.styles["Normal"]
            normal_style.font.name = "Arial"
            normal_style.font.size = docx.shared.Pt(10)

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

                if factor in chart_buffers:
                    doc.add_picture(io.BytesIO(chart_buffers[factor]), width=Inches(6.6))
                    cap = doc.add_paragraph(f"Figure: {factor} LAB Input vs LINE Output")
                    cap.alignment = 1

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
                    ("Current LAB ΔE* target", "≤ 0.500"),
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

            # Footer
            for sec in doc.sections:
                footer = sec.footer.paragraphs[0]
                footer.text = "SPC Color Dashboard - Task 3 Incoming LAB Control Limit Analysis"
                footer.alignment = 1

            report_buf = io.BytesIO()
            doc.save(report_buf)
            report_buf.seek(0)
            return report_buf.getvalue()

        task3_word_bytes = build_task3_word_report()
        safe_color_name = re.sub(r"[^A-Za-z0-9_-]+", "_", str(color))
        st.download_button(
            label="📄 Download Word Report",
            data=task3_word_bytes,
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
        man_L = st.number_input("Incoming LAB ΔL*:", value=0.000, step=0.100, format="%.3f")
    with col_ma:
        man_a = st.number_input("Incoming LAB Δa*:", value=0.000, step=0.100, format="%.3f")
    with col_mb:
        man_b = st.number_input("Incoming LAB Δb*:", value=0.000, step=0.100, format="%.3f")

    manual_dE = math.sqrt(man_L ** 2 + man_a ** 2 + man_b ** 2)
    lab_de_target = 0.5

    st.markdown("#### Calculation Result")
    if manual_dE <= lab_de_target:
        st.success(f"### 🎯 Incoming LAB ΔE*: **{manual_dE:.3f}** (✅ ≤ {lab_de_target})")
    else:
        st.error(f"### 🎯 Incoming LAB ΔE*: **{manual_dE:.3f}** (⚠️ > {lab_de_target})")

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
                spc_c = calculate_batch_averages(df_c)
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

        spc_p2 = calculate_batch_averages(df_p2)

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

        word_report_buf = build_ooc_kpi_word_report(
            display_report_df,
            report_source,
            ooc_target,
            n_controlled,
            n_ooc_ok,
            n_action,
            n_incomplete
        )

        st.download_button(
            "📄 Download Short Word Report for Management",
            data=word_report_buf,
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

