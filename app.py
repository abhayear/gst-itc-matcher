"""Streamlit app for GST ITC invoice matching — upload two files, get results."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from matcher.engine import export_to_excel, load_and_match

st.set_page_config(page_title="GST ITC Matcher", page_icon="📊", layout="wide")

st.title("GST ITC Matcher")
st.caption("Upload your Purchase Register and GSTR-2A/2B — matching runs automatically.")

col1, col2 = st.columns(2)

with col1:
    pr_file = st.file_uploader(
        "1. Purchase Register (Excel)",
        type=["xlsx", "xls"],
        key="pr",
    )

with col2:
    gstr_file = st.file_uploader(
        "2. GSTR-2A / GSTR-2B (Excel)",
        type=["xlsx", "xls"],
        key="gstr",
    )

if pr_file and gstr_file:
    file_key = f"{pr_file.name}|{pr_file.size}|{gstr_file.name}|{gstr_file.size}"

    if st.session_state.get("last_file_key") != file_key:
        with st.spinner("Matching invoices automatically..."):
            try:
                pr_file.seek(0)
                gstr_file.seek(0)
                result, summary = load_and_match(pr_file, gstr_file)
                st.session_state["match_result"] = result
                st.session_state["match_summary"] = summary
                st.session_state["last_file_key"] = file_key
            except Exception as exc:
                st.session_state.pop("match_result", None)
                st.session_state.pop("match_summary", None)
                st.error(f"Could not process files: {exc}")

    if "match_result" in st.session_state:
        result: pd.DataFrame = st.session_state["match_result"]
        summary = st.session_state["match_summary"]

        st.success("Done! Review results below or download the ITC Taken Excel.")

        excel_bytes = export_to_excel(result, summary)
        st.download_button(
            label="Download ITC Taken Excel",
            data=excel_bytes,
            file_name="itc_matching_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

        st.subheader("Summary")
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Fully Matched", summary.fully_matched)
        m2.metric("Tax Mismatch", summary.tax_mismatch)
        m3.metric("Not Matched", summary.not_matched)
        m4.metric("Duplicates", summary.duplicate)
        m5.metric("Total ITC Taken", f"₹{summary.total_itc:,.2f}")

        m6, m7, m8, m9 = st.columns(4)
        m6.metric("GSTIN Mismatch", summary.gstin_mismatch)
        m7.metric("Inv No Mismatch", summary.inv_no_mismatch)
        m8.metric("Inv Date Mismatch", summary.inv_date_mismatch)
        m9.metric("Missing in PR", summary.missing_in_pr)

        itc1, itc2, itc3 = st.columns(3)
        itc1.metric("Eligible IGST", f"₹{summary.total_itc_igst:,.2f}")
        itc2.metric("Eligible CGST", f"₹{summary.total_itc_cgst:,.2f}")
        itc3.metric("Eligible SGST", f"₹{summary.total_itc_sgst:,.2f}")

        st.subheader("Matching Results")
        st.dataframe(result, use_container_width=True, height=450)

else:
    st.info("Upload both Excel files above. Matching will start automatically.")

st.divider()
st.caption(
    "Columns are auto-detected from Tally, Busy, SAP, and GST portal exports. "
    "ITC decisions are indicative — verify under GST rules before claiming."
)
