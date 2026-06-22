"""
app.py
──────
Streamlit UI for the Pending Orders Processor.

Accepts either:
  • 1 file  — a pre-built Raw Data export (already has Brand/Depot/SKU Display)
  • 2 files — the two raw JSSOReport exports (original plant + Sonepat)

The correct mode is detected automatically from the sheet names present.
"""

from datetime import datetime
import io

import pandas as pd
import streamlit as st

from processor import process, ProcessingError

# ─────────────────────────────────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Pending Orders Processor",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

CUSTOM_CSS = """
<style>
/* ── Base ── */
.stApp { background-color: #F7FAFC; }

/* ── Hero banner ── */
.hero {
    background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 60%, #4A90C4 100%);
    padding: 2rem 2.4rem;
    border-radius: 14px;
    color: white;
    margin-bottom: 1.6rem;
    box-shadow: 0 6px 20px rgba(31,78,121,.28);
}
.hero h1 { margin: 0 0 .35rem 0; font-size: 1.85rem; letter-spacing: -.3px; }
.hero p  { margin: 0; opacity: .9; font-size: .97rem; line-height: 1.55; }

/* ── Section cards ── */
.section-card {
    background: white;
    border-radius: 12px;
    padding: 1.35rem 1.6rem;
    border: 1px solid #E2EBF3;
    margin-bottom: 1.2rem;
    box-shadow: 0 1px 4px rgba(0,0,0,.05);
}

/* ── Mode-badge pills ── */
.mode-badge {
    display: inline-flex;
    align-items: center;
    gap: .45rem;
    padding: .32rem .85rem;
    border-radius: 999px;
    font-weight: 700;
    font-size: .8rem;
    margin-bottom: .9rem;
}
.mode-badge.raw  { background:#DEEAF1; color:#1F4E79; border:1px solid #BDD7EE; }
.mode-badge.jsso { background:#E2EFDA; color:#375623; border:1px solid #A9D18E; }

/* ── How-to step pills (sidebar) ── */
.step-pill {
    display: inline-block;
    background: #1F4E79;
    color: white;
    font-weight: 700;
    font-size: .72rem;
    width: 1.35rem; height: 1.35rem;
    line-height: 1.35rem;
    text-align: center;
    border-radius: 50%;
    margin-right: .45rem;
}

/* ── Mode-choice info boxes ── */
.mode-box {
    border-radius: 10px;
    padding: .85rem 1.1rem;
    margin-bottom: .7rem;
    border: 1.5px solid;
}
.mode-box.a { background:#F0F7FF; border-color:#BDD7EE; }
.mode-box.b { background:#F4FAF0; border-color:#A9D18E; }
.mode-box h4 { margin: 0 0 .25rem 0; font-size: .88rem; }
.mode-box p  { margin: 0; font-size: .82rem; opacity: .85; }

/* ── Status banners ── */
.overdue-banner {
    background:#FFE8E8; border:1.5px solid #F1A9A9; color:#C00000;
    border-radius:10px; padding:.7rem 1rem; font-weight:600; margin-bottom:.75rem;
}
.clean-banner {
    background:#EBF7E6; border:1.5px solid #A9D18E; color:#2E7D32;
    border-radius:10px; padding:.7rem 1rem; font-weight:600; margin-bottom:.75rem;
}
.warn-banner {
    background:#FFF8E6; border:1.5px solid #FFD966; color:#7D5A00;
    border-radius:10px; padding:.7rem 1rem; font-weight:600; margin-bottom:.75rem;
}

/* ── Metrics ── */
div[data-testid="stMetric"] {
    background: white;
    border: 1.5px solid #E2EBF3;
    border-radius: 10px;
    padding: .85rem 1rem .45rem 1rem;
}

/* ── Buttons ── */
.stButton > button {
    background: #1F4E79; color: white;
    border-radius: 8px; font-weight: 600; border: none;
    padding: .55rem 1.5rem; transition: background .2s;
}
.stButton > button:hover { background: #2E75B6; color: white; }

.stDownloadButton > button {
    background: #1A6B2A; color: white;
    border-radius: 8px; font-weight: 700; border: none;
    padding: .6rem 1.8rem; font-size: .95rem; transition: background .2s;
}
.stDownloadButton > button:hover { background: #145422; color: white; }

/* ── Misc ── */
footer { visibility: hidden; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────
# Hero
# ─────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <h1>📦 Pending Orders Processor</h1>
  <p>Upload your raw sales-order export(s), click <strong>Generate</strong>,
  and download a clean, formatted workbook with Production Planning,
  Order Queue, External Orders and Raw Data — all in one click.
  Supports both a pre-built Raw Data file <em>and</em> the original
  two-file JSSOReport workflow.</p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────
# Sidebar
# ─────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### How to use")
    st.markdown("""
<span class="step-pill">1</span> Choose your input mode (see below) and
upload 1 or 2 files.<br><br>
<span class="step-pill">2</span> Click <strong>Generate Report</strong>.<br><br>
<span class="step-pill">3</span> Review the summary metrics and overdue flags.<br><br>
<span class="step-pill">4</span> Click <strong>Download</strong> to save
the formatted <code>.xlsx</code>.
""", unsafe_allow_html=True)

    st.divider()
    st.markdown("### Input modes")
    st.markdown("""
<div class="mode-box b">
  <h4>📄 Single Raw Data file</h4>
  <p>Upload <strong>one</strong> file that already has a <code>Raw Data</code>
  sheet with Brand / Depot / SKU Display / Order Type pre-filled.
  This is the fastest option if you have yesterday's output on hand.</p>
</div>
<div class="mode-box a">
  <h4>📂 Two JSSOReport files</h4>
  <p>Upload <strong>both</strong> raw system exports (each must have a
  <code>JSSOReport</code> sheet) — the plant file and the Sonepat file.
  The app derives brand, depot, SKU and qty from scratch.</p>
</div>
""", unsafe_allow_html=True)

    st.divider()
    st.markdown("### Output sheets")
    st.markdown(
        "- **Production Planning** — pending qty pivoted by brand × SKU × depot\n"
        "- **Order Queue (FIFO)** — orders by schedule date, overdue flagged ⚠\n"
        "- **External Orders** — Sulphur Powder & Hariyali DF outside Jaishil\n"
        "- **Raw Data** — the full cleaned dataset"
    )
    st.divider()
    st.caption("Nothing you upload is retained — all processing is in-memory for this session only.")


# ─────────────────────────────────────────────────────────────────────────
# Step 1 — Upload
# ─────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown("#### Step 1 &nbsp;·&nbsp; Upload file(s)")
st.markdown(
    "Drag and drop **1 file** *(Raw Data export)* or **2 files** "
    "*(both JSSOReport exports)* — the mode is detected automatically.",
    unsafe_allow_html=True,
)
uploaded_files = st.file_uploader(
    "Drop .xlsx file(s) here or click to browse",
    type=["xlsx"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)
st.markdown("</div>", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────
# Live file preview + mode sniffing
# ─────────────────────────────────────────────────────────────────────────
if uploaded_files:
    n = len(uploaded_files)

    # Show uploaded files
    cols = st.columns(n if n <= 2 else 3)
    for col, f in zip(cols, uploaded_files):
        with col:
            st.success(f"📄 **{f.name}** &nbsp; {f.size/1024:.1f} KB")

    # Quick sheet peek for immediate feedback (no processing yet)
    import pandas as _pd
    def _sheets(f):
        try:
            f.seek(0)
            s = _pd.ExcelFile(f).sheet_names
            f.seek(0)
            return s
        except Exception:
            return []

    sheets_list = [_sheets(f) for f in uploaded_files]
    all_names   = [getattr(f, 'name', '') for f in uploaded_files]

    if n == 1:
        if 'Raw Data' in sheets_list[0]:
            st.markdown(
                '<div class="mode-badge raw">📄 Single Raw Data file detected — one-file mode</div>',
                unsafe_allow_html=True,
            )
        elif 'JSSOReport' in sheets_list[0]:
            st.warning(
                "⚠️ This file has a **JSSOReport** sheet — for two-file mode "
                "please also upload the second export (plant or Sonepat).",
                icon=None,
            )
        else:
            st.error(
                f"**{all_names[0]}** doesn't have a 'Raw Data' or 'JSSOReport' sheet. "
                "Please upload an unmodified export file."
            )
    elif n == 2:
        both_jsso = all('JSSOReport' in s for s in sheets_list)
        any_raw   = any('Raw Data' in s for s in sheets_list)
        if both_jsso:
            st.markdown(
                '<div class="mode-badge jsso">📂 Two JSSOReport exports detected — two-file mode</div>',
                unsafe_allow_html=True,
            )
        elif any_raw:
            st.warning(
                "⚠️ One of your files looks like a Raw Data export — "
                "for single-file mode please upload **only** that one file, "
                "or for two-file mode upload **both** raw JSSOReport exports."
            )
    else:
        st.warning(f"⚠️ You've uploaded {n} files — please upload 1 or 2.")


# ─────────────────────────────────────────────────────────────────────────
# Step 2 — Generate
# ─────────────────────────────────────────────────────────────────────────
if uploaded_files and len(uploaded_files) in (1, 2):
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown("#### Step 2 &nbsp;·&nbsp; Generate the report")
    generate = st.button("🚀  Generate Report", use_container_width=False)
    st.markdown("</div>", unsafe_allow_html=True)

    if generate:
        with st.spinner("Processing orders, mapping depots and building sheets…"):
            try:
                buffer, summary = process(uploaded_files)
                st.session_state["report_buffer"]  = buffer.getvalue()
                st.session_state["summary"]        = summary
                st.session_state["generated_at"]   = datetime.now()
            except ProcessingError as e:
                st.session_state.pop("report_buffer", None)
                st.error(f"⚠️ {e}")
            except Exception as e:
                st.session_state.pop("report_buffer", None)
                st.error(
                    "Something went wrong while processing. "
                    "Make sure the files are unmodified exports."
                )
                with st.expander("Technical details"):
                    st.exception(e)


# ─────────────────────────────────────────────────────────────────────────
# Step 3 — Results
# ─────────────────────────────────────────────────────────────────────────
if "report_buffer" in st.session_state:
    summary = st.session_state["summary"]
    mode    = summary.get("mode", "")

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown("#### Step 3 &nbsp;·&nbsp; Review &amp; download")

    # ── Mode-specific detection panel ────────────────────────────────
    with st.expander("🔎 Detection details", expanded=False):
        if mode == "jssoreport":
            st.write(f"**Mode:** Two JSSOReport files")
            st.write(f"**Plant / Sulphur Powder file:** {summary['new_file_label']}")
            st.write(f"**Sonepat file:** {summary['sonepat_file_label']}")
            if summary.get("ambiguous_detection"):
                st.markdown(
                    '<div class="warn-banner">⚠️ Couldn\'t tell the two files apart by '
                    'Sulphur Powder content — both were combined and processed. '
                    'Double-check the file roles above.</div>',
                    unsafe_allow_html=True,
                )
        else:
            st.write(f"**Mode:** Single Raw Data file")
            st.write(f"**Source file:** {summary.get('source_file', '—')}")
            st.info(
                "Columns Brand / Depot / SKU Display / Order Type were read directly "
                "from the Raw Data sheet (no re-derivation).",
                icon="ℹ️",
            )
        if summary["excluded_a20_rows"]:
            st.caption(
                f"Excluded {summary['excluded_a20_rows']} row(s) belonging to "
                "the internal Sales (A/20) ledger account."
            )

    # ── Key metrics ───────────────────────────────────────────────────
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Inter-depot orders",    summary["jaishil_orders"])
    m2.metric("External order lines",  summary["external_order_lines"])
    m3.metric("⚠ Overdue orders",      summary["overdue_count"])
    m4.metric("Total rows processed",  summary["total_rows_processed"])

    # ── Depots in report ──────────────────────────────────────────────
    if summary["depots_present"]:
        st.caption(
            "Depots in Production Planning: "
            + " · ".join(f"**{d}**" for d in summary["depots_present"])
        )

    # ── Overdue banner ────────────────────────────────────────────────
    if summary["overdue_count"]:
        st.markdown(
            f'<div class="overdue-banner">'
            f'⚠️ {summary["overdue_count"]} order(s) past their schedule date</div>',
            unsafe_allow_html=True,
        )
        with st.expander(f"View {summary['overdue_count']} overdue order number(s)"):
            for so in summary["overdue_orders"]:
                st.write(f"• {so}")
    else:
        st.markdown(
            '<div class="clean-banner">✅ No overdue orders — all on schedule</div>',
            unsafe_allow_html=True,
        )

    # ── Download button ───────────────────────────────────────────────
    ts       = st.session_state["generated_at"]
    filename = f"pending_orders_{ts.strftime('%d%m%Y_%H%M%S')}.xlsx"
    st.download_button(
        label="⬇️  Download Formatted Report",
        data=st.session_state["report_buffer"],
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.markdown("</div>", unsafe_allow_html=True)

    # ── In-browser preview ────────────────────────────────────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown("#### Preview")

    raw_bytes = io.BytesIO(st.session_state["report_buffer"])
    tab_oq, tab_ext, tab_rd = st.tabs(
        ["📋 Order Queue (FIFO)", "🌐 External Orders", "🗄 Raw Data"])

    with tab_oq:
        df_oq = pd.read_excel(raw_bytes, sheet_name="Order Queue (FIFO)")
        st.dataframe(df_oq, use_container_width=True, height=340)

    raw_bytes.seek(0)
    with tab_ext:
        df_ext = pd.read_excel(raw_bytes, sheet_name="External Orders")
        st.dataframe(df_ext, use_container_width=True, height=340)

    raw_bytes.seek(0)
    with tab_rd:
        df_rd = pd.read_excel(raw_bytes, sheet_name="Raw Data")
        st.dataframe(df_rd, use_container_width=True, height=340)

    st.caption(
        "💡 **Production Planning** uses merged cells and a multi-block pivot layout — "
        "open the downloaded file in Excel to view it correctly."
    )
    st.markdown("</div>", unsafe_allow_html=True)
