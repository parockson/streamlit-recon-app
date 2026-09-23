import streamlit as st
import pandas as pd
import io
from src.processor import process_recon_data, coerce_id_phone_columns
from src.report_generator import build_excel_workbook, build_raw_workbook

st.set_page_config(page_title="Reconciliation Dashboard", layout="wide")
st.title("📊 Financial Reconciliation App")

MATCH_SHEET = "Match Results"

uploaded_file = st.file_uploader(
    "Upload Recon Excel Workbook (.xlsx) — must contain a 'Match Results' sheet",
    type=["xlsx"]
)

if uploaded_file:
    # ── Read all sheets from the workbook ──────────────────────────────────
    file_bytes = uploaded_file.read()  # read once; reuse for raw download
    all_sheets: dict[str, pd.DataFrame] = pd.read_excel(
        io.BytesIO(file_bytes), sheet_name=None
    )

    # Preserve ID / phone columns as clean strings across every sheet
    all_sheets = {name: coerce_id_phone_columns(df) for name, df in all_sheets.items()}

    # Validate that Match Results sheet exists
    if MATCH_SHEET not in all_sheets:
        st.error(
            f"❌ The uploaded file does not contain a sheet named **'{MATCH_SHEET}'**. "
            f"Found sheets: {', '.join(all_sheets.keys())}"
        )
        st.stop()

    st.success(f"✅ '{MATCH_SHEET}' sheet found — running reconciliation report.")

    # ── Use only Match Results as the recon data source ───────────────────
    match_df = all_sheets[MATCH_SHEET]

    # Process metrics, reports, and base filename
    summaries, reports, base_name = process_recon_data(match_df)

    # Derive the two download filenames from the same base
    recon_filename = base_name          # e.g. UNIT_01-01-2025_to_31-01-2025_recon.xlsx
    raw_filename = base_name.replace("_recon.xlsx", "_recon-raw.xlsx")  # …_recon-raw.xlsx

    # ── 1. Summaries Display ──────────────────────────────────────────────
    st.header("📈 Summaries")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1. Total Summary")
        st.dataframe(summaries["s1_total"], use_container_width=True)

        st.subheader("3. Xchange Match Summary (Exchange=Y, External=N)")
        st.dataframe(summaries["s3_xchange_only"], use_container_width=True)

    with col2:
        st.subheader("2. Match Summary (Exchange=Y, External=Y)")
        st.dataframe(summaries["s2_match"], use_container_width=True)

        st.subheader("4. External Summary (Exchange=N, External=Y)")
        st.dataframe(summaries["s4_ext_only"], use_container_width=True)

    # ── 2. Detailed Transaction Reports ───────────────────────────────────
    st.divider()
    st.header("📋 Transaction Reports")

    t1, t2, t3 = st.tabs([
        "5. Match Transactions",
        "6. Xchange Match Transactions",
        "7. External Match Transactions"
    ])

    with t1:
        st.dataframe(reports["match_txns"], use_container_width=True)
    with t2:
        st.dataframe(reports["xchange_txns"], use_container_width=True)
    with t3:
        st.dataframe(reports["external_txns"], use_container_width=True)

    # ── 3. Downloads ───────────────────────────────────────────────────────
    st.divider()
    st.subheader("📥 Downloads")

    dl_col1, dl_col2 = st.columns(2)

    with dl_col1:
        recon_data = build_excel_workbook(summaries, reports)
        st.download_button(
            label=f"📊 Download Recon Report",
            help=f"Processed reconciliation report → {recon_filename}",
            data=recon_data,
            file_name=recon_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption(f"`{recon_filename}`")

    with dl_col2:
        raw_data = build_raw_workbook(all_sheets)
        st.download_button(
            label=f"🗂️ Download Raw Workbook",
            help=f"Full original workbook (all sheets) → {raw_filename}",
            data=raw_data,
            file_name=raw_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption(f"`{raw_filename}`")