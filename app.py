import streamlit as st
import pandas as pd
from src.processor import process_recon_data
from src.report_generator import build_excel_workbook

st.set_page_config(page_title="Reconciliation Dashboard", layout="wide")
st.title("📊 Financial Reconciliation App")

uploaded_file = st.file_uploader("Upload Recon Dataset (.xlsx or .csv)", type=["xlsx", "csv"])

if uploaded_file:
    # Read uploaded file
    df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)

    # Process metrics, reports, and filename
    summaries, reports, recon_filename = process_recon_data(df)

    # 1. Summaries Display
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

    # 2. Detailed Transaction Reports Display
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

    # 3. Export Single Excel Download
    excel_data = build_excel_workbook(summaries, reports)
    st.divider()
    
    st.download_button(
        label=f"📥 Download Report ({recon_filename})",
        data=excel_data,
        file_name=recon_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )