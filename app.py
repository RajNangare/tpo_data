import streamlit as st
import pandas as pd
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── helpers ──────────────────────────────────────────────────────────────────

def to_title_case(text: str) -> str:
    """Convert any string to Title Case (e.g. RAJ NANGARE -> Raj Nangare)."""
    s = str(text).strip()
    return ' '.join(word.capitalize() for word in s.split())


def convert_column_type(series: pd.Series, dtype: str) -> pd.Series:
    """Convert a series to int, float, or 10-digit phone number."""
    if dtype == "Integer":
        def to_int(v):
            try:
                return int(float(str(v).strip()))
            except:
                return v
        return series.apply(lambda v: to_int(v) if pd.notna(v) else v)

    elif dtype == "Float":
        def to_float(v):
            try:
                return float(str(v).strip())
            except:
                return v
        return series.apply(lambda v: to_float(v) if pd.notna(v) else v)

    elif dtype == "Phone (last 10 digits)":
        def to_phone(v):
            digits = re.sub(r"\D", "", str(v))  # strip all non-digits
            return digits[-10:] if len(digits) >= 10 else digits
        return series.apply(lambda v: to_phone(v) if pd.notna(v) else v)

    return series


def read_file(uploaded) -> pd.DataFrame:
    name = uploaded.name.lower()
    if name.endswith('.csv'):
        import io as _io
        raw = uploaded.read()
        uploaded.seek(0)

        # Auto-detect separator
        sample = raw[:4096].decode('utf-8', errors='replace')
        sep = ','
        for candidate in [',', ';', '\t', '|']:
            if sample.count(candidate) > sample.count(sep):
                sep = candidate

        # Read header to get real column names
        text = raw.decode('utf-8', errors='replace')
        lines = text.splitlines()
        header = lines[0].split(sep) if lines else []

        # Find max columns across all rows (some CSVs have more data cols than header)
        max_cols = max((len(l.split(sep)) for l in lines if l.strip()), default=len(header))

        # Pad header if needed so pandas doesn't eat col 0 as index
        extra = max_cols - len(header)
        col_names = header + [f'_extra_{i}' for i in range(extra)]

        df = pd.read_csv(
            _io.BytesIO(raw),
            sep=sep,
            on_bad_lines='skip',
            engine='python',
            index_col=False,
            names=col_names,
            skiprows=1,
        )
        # Drop the padded overflow columns
        df = df[[c for c in df.columns if not c.startswith('_extra_')]]

    else:
        df = pd.read_excel(uploaded, index_col=None)

    # Drop only columns that are Unnamed AND completely empty (stray Excel cols)
    unnamed_mask = df.columns.astype(str).str.match(r'^Unnamed')
    all_empty_mask = df.isnull().all(axis=0)
    df = df.loc[:, ~(unnamed_mask & all_empty_mask)]

    # Drop rows that are completely empty
    df = df.dropna(how='all')

    # Reset index cleanly
    df = df.reset_index(drop=True)
    return df


def df_to_xlsx_bytes(df: pd.DataFrame, col_padding: int = 0) -> bytes:
    """Write df to xlsx — Calibri, yellow bold headers, white data, thin borders on table only."""
    wb_out = Workbook()
    ws = wb_out.active

    # ── Fonts ──────────────────────────────────────────────────────────────
    # Heading: Calibri 12 bold, BLACK text
    header_font = Font(bold=True, name="Calibri", size=12, color="000000")
    # Data: Calibri 11, BLACK text
    data_font   = Font(bold=False, name="Calibri", size=11, color="000000")

    # ── Header fill — yellow background ────────────────────────────────────
    header_fill  = PatternFill("solid", fgColor="FFFF00")   # yellow
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # ── Data rows fill — white background ──────────────────────────────────
    row_fill_odd  = PatternFill("solid", fgColor="FFFFFF")  # white
    row_fill_even = PatternFill("solid", fgColor="FFFFFF")  # white

    # ── Borders — thin black on all 4 sides, ONLY on table cells ──────────
    thin = Side(border_style="thin", color="000000")
    cell_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    n_rows = len(df)
    n_cols = len(df.columns)

    # ── Write header row ───────────────────────────────────────────────────
    for ci, col in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = header_align
        cell.border    = cell_border

    # ── Write data rows ────────────────────────────────────────────────────
    for ri in range(n_rows):
        fill = row_fill_odd if ri % 2 == 0 else row_fill_even
        for ci in range(n_cols):
            val = df.iloc[ri, ci]
            if hasattr(val, 'item'):
                val = val.item()
            cell = ws.cell(row=ri + 2, column=ci + 1, value=val)
            cell.font      = data_font
            cell.fill      = fill
            cell.border    = cell_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # ── Column widths ──────────────────────────────────────────────────────
    for ci, col in enumerate(df.columns, start=1):
        col_letter = get_column_letter(ci)
        try:
            data_max = df[col].dropna().astype(str).str.len().max()
            data_max = int(data_max) if pd.notna(data_max) else 0
        except Exception:
            data_max = 0
        max_len = max(len(str(col)), data_max)
        ws.column_dimensions[col_letter].width = max_len + col_padding + 2

    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb_out.save(buf)
    buf.seek(0)
    return buf.read()



def df_to_html_preview(df: pd.DataFrame) -> str:
    """Render df as a styled HTML table matching the Excel output exactly."""
    rows_html = ""
    for ri in range(min(len(df), 20)):
        cells = ""
        for ci in range(len(df.columns)):
            val = df.iloc[ri, ci]
            val = "" if pd.isna(val) else str(val)
            cells += f'<td>{val}</td>'
        rows_html += f"<tr>{cells}</tr>"

    headers = "".join(f"<th>{col}</th>" for col in df.columns)

    return f"""
    <div style="overflow-x:auto; margin-top:8px;">
    <table style="
        border-collapse: collapse;
        font-family: Calibri, sans-serif;
        width: 100%;
        font-size: 13px;
    ">
        <thead>
            <tr style="background:#FFFF00;">
                {headers}
            </tr>
        </thead>
        <tbody>
            {rows_html}
        </tbody>
    </table>
    </div>
    <style>
    table th {{
        background: #FFFF00 !important;
        color: #000000 !important;
        font-weight: bold !important;
        font-size: 14px !important;
        text-align: center !important;
        padding: 7px 10px !important;
        border: 1.5px solid #000000 !important;
        white-space: nowrap;
    }}
    table td {{
        background: #FFFFFF !important;
        color: #000000 !important;
        font-size: 13px !important;
        text-align: center !important;
        padding: 5px 10px !important;
        border: 1px solid #000000 !important;
    }}
    </style>
    """

# ── page config ──────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Sheet Toolkit",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'Space Grotesk', sans-serif;
}
.block-container { padding-top: 1.5rem; }

h1 { font-size: 2rem !important; font-weight: 700 !important; letter-spacing: -0.5px; }
h2 { font-size: 1.35rem !important; font-weight: 600 !important; }
h3 { font-size: 1.1rem !important; font-weight: 600 !important; }

/* tab styling */
.stTabs [data-baseweb="tab-list"] {
    gap: 6px;
    background: #e2e8f0;
    padding: 5px;
    border-radius: 12px;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px;
    padding: 10px 28px;
    font-weight: 700;
    font-size: 0.95rem;
    color: #1e293b !important;
    background: transparent;
}
.stTabs [data-baseweb="tab"]:hover {
    background: #cbd5e1 !important;
    color: #0f172a !important;
}
.stTabs [aria-selected="true"] {
    background: #ffffff !important;
    color: #0f172a !important;
    box-shadow: 0 1px 4px rgba(0,0,0,0.15);
}
/* force all tab label text to be dark and visible */
.stTabs [data-baseweb="tab"] p,
.stTabs [data-baseweb="tab"] span,
.stTabs [data-baseweb="tab"] div {
    color: #1e293b !important;
    font-weight: 700 !important;
}

/* download button — default (structure tab) */
.stDownloadButton > button {
    background: #2E4057 !important;
    color: white !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    padding: 0.5rem 1.5rem !important;
    border: none !important;
}
.stDownloadButton > button:hover {
    background: #1a2d42 !important;
}

/* mapping download button — bright green, larger */
.map-download-btn .stDownloadButton > button {
    background: #16A34A !important;
    color: white !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
    padding: 0.75rem 2.5rem !important;
    border: none !important;
    box-shadow: 0 4px 14px rgba(22,163,74,0.4) !important;
    letter-spacing: 0.3px;
}
.map-download-btn .stDownloadButton > button:hover {
    background: #15803D !important;
    box-shadow: 0 6px 18px rgba(22,163,74,0.55) !important;
    transform: translateY(-1px);
}

/* info boxes */
.tip-box {
    background: #EFF6FF;
    border-left: 4px solid #3B82F6;
    padding: 0.75rem 1rem;
    border-radius: 0 8px 8px 0;
    font-size: 0.88rem;
    color: #1E40AF;
    margin-bottom: 1rem;
}

.section-card {
    background: white;
    border: 1px solid #E5E7EB;
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    margin-bottom: 1.2rem;
}

code { font-family: 'JetBrains Mono', monospace; font-size: 0.82rem; }
</style>
""", unsafe_allow_html=True)

# ── sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## 📊 Sheet Toolkit")
    st.markdown("---")
    st.markdown("""
**Two modes:**

🔧 **Structure** — Clean & format a single file  
🔗 **Mapping** — Map a base file to an expected schema
    """)
    st.markdown("---")
    st.markdown("**Supported formats:** `.csv`, `.xlsx`, `.xls`, `.xlsm`")
    st.markdown("---")
    st.caption("Tip: All downloads are `.xlsx` with formatted headers.")

# ── main ─────────────────────────────────────────────────────────────────────

st.markdown("# 📊 Sheet Toolkit")
st.markdown("Clean, structure, and map your spreadsheet data in seconds.")

tab1, tab2 = st.tabs(["🔧  Structure File", "🔗  Mapping"])


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — STRUCTURE FILE
# ═══════════════════════════════════════════════════════════════════════════════

with tab1:
    st.markdown("### Upload your file")
    uploaded = st.file_uploader(
        "Drag & drop or browse",
        type=["csv", "xlsx", "xls", "xlsm"],
        key="struct_upload",
    )

    if uploaded:
        df = read_file(uploaded)
        original_cols = list(df.columns)

        st.success(f"Loaded **{len(df)} rows × {len(df.columns)} cols**")

        with st.expander("👀 Preview raw data", expanded=False):
            st.dataframe(df.head(10), use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown("### ⚙️ Transformations")

        col_a, col_b = st.columns(2)

        with col_a:
            st.markdown("#### 1️⃣ Title Case headers")
            st.markdown('<div class="tip-box">All column headers will be converted to Title Case (e.g. "FULL NAME" → "Full Name").</div>', unsafe_allow_html=True)
            do_all_camel = st.checkbox("Convert ALL headers to Title Case", value=True, key="all_camel")

            st.markdown("#### 2️⃣ Title Case selected data columns")
            st.markdown('<div class="tip-box">Pick columns whose <em>cell values</em> should be converted to Title Case (e.g. RAJ NANGARE → Raj Nangare).</div>', unsafe_allow_html=True)
            data_camel_cols = st.multiselect(
                "Select columns",
                options=original_cols,
                key="data_camel",
            )

        with col_b:
            st.markdown("#### 3️⃣ Column width padding")
            st.markdown('<div class="tip-box">Extra characters added to each column beyond its longest value. Excel default fits content; adding padding gives breathing room.</div>', unsafe_allow_html=True)
            col_padding = st.slider("Extra padding (chars)", 0, 20, 5, key="col_pad")

            st.markdown("#### 4️⃣ Drop empty rows / columns")
            drop_empty = st.checkbox("Drop fully-empty rows", value=False, key="drop_empty")

        st.markdown("---")
        st.markdown("#### 5️⃣ Convert column data types")
        st.markdown('<div class="tip-box">Select a column and choose how to convert its values. Phone number keeps only the last 10 digits.</div>', unsafe_allow_html=True)

        type_conversions = {}   # col_name -> dtype string
        num_conversions = st.number_input("How many columns to convert?", min_value=0, max_value=20, value=0, step=1, key="num_conv")
        for i in range(int(num_conversions)):
            cc1, cc2 = st.columns([2, 1])
            col_choice = cc1.selectbox(f"Column #{i+1}", options=["— select —"] + original_cols, key=f"conv_col_{i}")
            type_choice = cc2.selectbox(f"Convert to", options=["Integer", "Float", "Phone (last 10 digits)"], key=f"conv_type_{i}")
            if col_choice != "— select —":
                type_conversions[col_choice] = type_choice

        st.markdown("---")
        st.markdown("### 🔄 Preview result")

        result = df.copy()

        if drop_empty:
            # Only drop rows that are ENTIRELY empty — never drop columns
            result = result.dropna(how='all')

        # Apply type conversions using original column names
        for col, dtype in type_conversions.items():
            if col in result.columns:
                result[col] = convert_column_type(result[col], dtype)

        # Apply value transformation FIRST using original column names
        for col in data_camel_cols:
            if col in result.columns:
                result[col] = result[col].apply(
                    lambda v: to_title_case(str(v)) if pd.notna(v) else v
                )

        # THEN rename headers
        if do_all_camel:
            result.columns = [to_title_case(c) for c in result.columns]

        st.markdown(df_to_html_preview(result), unsafe_allow_html=True)

        xlsx_bytes = df_to_xlsx_bytes(result, col_padding=col_padding)
        out_name = uploaded.name.rsplit('.', 1)[0] + "_structured.xlsx"

        st.download_button(
            label="⬇️  Download structured file",
            data=xlsx_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — MAPPING
# ═══════════════════════════════════════════════════════════════════════════════

with tab2:
    st.markdown("### Upload both files")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**📋 Expected / Target schema file**")
        st.caption("Defines the column names you want in the output.")
        exp_file = st.file_uploader(
            "Expected file",
            type=["csv", "xlsx", "xls", "xlsm"],
            key="exp_upload",
            label_visibility="collapsed",
        )

    with c2:
        st.markdown("**📂 Base / Source data file**")
        st.caption("The file whose data will be mapped to the expected schema.")
        base_file = st.file_uploader(
            "Base file",
            type=["csv", "xlsx", "xls", "xlsm"],
            key="base_upload",
            label_visibility="collapsed",
        )

    if exp_file and base_file:
        exp_df  = read_file(exp_file)
        base_df = read_file(base_file)

        exp_cols  = list(exp_df.columns)
        base_cols = list(base_df.columns)

        st.success(
            f"Expected: **{len(exp_cols)} cols** | Base: **{len(base_df)} rows × {len(base_cols)} cols**"
        )

        with st.expander("👀 Preview expected schema (first 5 rows)", expanded=False):
            st.dataframe(exp_df.head(5), use_container_width=True, hide_index=True)

        with st.expander("👀 Preview base data (first 5 rows)", expanded=False):
            st.dataframe(base_df.head(5), use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown("### 🗺️ Column Mapping")
        st.markdown('<div class="tip-box">For each expected column, choose the matching column from your base file — or leave as <code>— (empty column) —</code> to include the column header with no data.</div>', unsafe_allow_html=True)

        mapping = {}   # expected_col -> base_col or None
        EMPTY_OPT = "— (empty column) —"
        base_options = [EMPTY_OPT] + base_cols

        # auto-suggest: case-insensitive exact match
        def auto_match(exp_col):
            for bc in base_cols:
                if bc.strip().lower() == exp_col.strip().lower():
                    return bc
            return EMPTY_OPT

        cols_per_row = 3
        for i in range(0, len(exp_cols), cols_per_row):
            row_cols = exp_cols[i:i + cols_per_row]
            grid = st.columns(len(row_cols))
            for j, exp_col in enumerate(row_cols):
                default = auto_match(exp_col)
                chosen = grid[j].selectbox(
                    f"`{exp_col}`",
                    options=base_options,
                    index=base_options.index(default),
                    key=f"map_{i}_{j}",
                )
                mapping[exp_col] = None if chosen == EMPTY_OPT else chosen

        st.markdown("---")
        st.markdown("### ⚙️ Output options")

        mc1, mc2 = st.columns(2)
        with mc1:
            map_all_camel = st.checkbox("Convert all output headers to Title Case", value=True, key="map_camel")
            map_data_camel_cols = st.multiselect(
                "Also convert values to Title Case in these output columns",
                options=exp_cols,
                key="map_data_camel",
            )
        with mc2:
            map_padding = st.slider("Extra column-width padding (chars)", 0, 20, 5, key="map_pad")
            map_drop_empty = st.checkbox("Drop fully-empty rows", value=False, key="map_drop")

        st.markdown("---")
        st.markdown("#### Convert column data types")
        st.markdown('<div class="tip-box">Select a column and choose how to convert its values. Phone number keeps only the last 10 digits.</div>', unsafe_allow_html=True)

        map_type_conversions = {}
        map_num_conv = st.number_input("How many columns to convert?", min_value=0, max_value=20, value=0, step=1, key="map_num_conv")
        for i in range(int(map_num_conv)):
            mc1b, mc2b = st.columns([2, 1])
            col_choice = mc1b.selectbox(f"Column #{i+1}", options=["— select —"] + exp_cols, key=f"map_conv_col_{i}")
            type_choice = mc2b.selectbox(f"Convert to", options=["Integer", "Float", "Phone (last 10 digits)"], key=f"map_conv_type_{i}")
            if col_choice != "— select —":
                map_type_conversions[col_choice] = type_choice

        st.markdown("---")
        st.markdown("### 🔄 Preview mapped result")

        # build output dataframe
        out_data = {}
        for exp_col, base_col in mapping.items():
            if base_col and base_col in base_df.columns:
                out_data[exp_col] = base_df[base_col].values
            else:
                out_data[exp_col] = [None] * len(base_df)

        out_df = pd.DataFrame(out_data)

        if map_drop_empty:
            out_df = out_df.dropna(how='all')

        # Apply type conversions
        for col, dtype in map_type_conversions.items():
            if col in out_df.columns:
                out_df[col] = convert_column_type(out_df[col], dtype)

        # Apply value transformation FIRST using original col names
        for col in map_data_camel_cols:
            if col in out_df.columns:
                out_df[col] = out_df[col].apply(
                    lambda v: to_title_case(str(v)) if pd.notna(v) else v
                )

        # THEN rename headers
        if map_all_camel:
            out_df.columns = [to_title_case(c) for c in out_df.columns]

        st.markdown(df_to_html_preview(out_df), unsafe_allow_html=True)

        # mapping summary
        with st.expander("📋 Mapping summary"):
            summary_rows = []
            for exp_col, base_col in mapping.items():
                summary_rows.append({
                    "Expected Column": exp_col,
                    "Mapped From (Base)": base_col if base_col else "— empty —",
                    "Status": "✅ Mapped" if base_col else "⬜ Empty",
                })
            st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

        xlsx_bytes = df_to_xlsx_bytes(out_df, col_padding=map_padding)
        out_name = base_file.name.rsplit('.', 1)[0] + "_mapped.xlsx"

        st.markdown('<div class="map-download-btn">', unsafe_allow_html=True)
        st.download_button(
            label="⬇️  Download Mapped File",
            data=xlsx_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.markdown('</div>', unsafe_allow_html=True)

    elif exp_file or base_file:
        st.info("Please upload **both** the expected schema file and the base data file to continue.")