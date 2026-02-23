import streamlit as st
import pandas as pd
import io
from reconciliation_engine import parse_tally, parse_gstr2b, reconcile
from datetime import datetime

st.set_page_config(page_title="GST Reconciliation", layout="wide")

st.title("📊 GST Reconciliation System")
st.caption("Internal Office Tool | ITC Books vs GSTR-2B Comparison")

col1, col2 = st.columns(2)

with col1:
    tally_file = st.file_uploader("📘 Upload Tally Purchase Register", type=["xlsx", "xls", "csv"])

with col2:
    gstr2b_file = st.file_uploader("📗 Upload Structured GSTR-2B", type=["xlsx", "xls", "csv"])

if st.button("🚀 Run Reconciliation", use_container_width=True):
    if tally_file is None or gstr2b_file is None:
        st.error("Please upload both files.")
    else:
        try:
            tally_raw = pd.read_csv(tally_file) if tally_file.name.endswith("csv") else pd.read_excel(tally_file)
            gstr2b_raw = pd.read_csv(gstr2b_file) if gstr2b_file.name.endswith("csv") else pd.read_excel(gstr2b_file)

            tally_clean, no_itc_df, invalid_gstin_df = parse_tally(tally_raw)
            gstr2b_clean = parse_gstr2b(gstr2b_raw)

            results = reconcile(gstr2b_clean, tally_clean)

            results["no_itc"] = no_itc_df
            results["invalid_gstin"] = invalid_gstin_df

            st.session_state["results"] = results
            st.success("Reconciliation Completed Successfully.")

        except Exception as e:
            st.error(str(e))

# ================= DISPLAY ================= #

if "results" in st.session_state:
    results = st.session_state["results"]
    summary = results["summary"]

    # ================= CLEAN EXECUTIVE SUMMARY ================= #

    st.markdown("## 📊 Executive Summary")

    r1c1, r1c2, r1c3 = st.columns(3)
    r1c1.metric("Total Invoices - Books", summary["Total_Invoices_Books"])
    r1c2.metric("Total Invoices - GSTR-2B", summary["Total_Invoices_2B"])
    r1c3.metric("Fully Matched Invoices", summary["Total_Matched"])

    st.divider()

    r2c1, r2c2, r2c3 = st.columns(3)
    r2c1.metric("Total ITC - Books", f"₹{summary['Total_ITC_Books']:,.2f}")
    r2c2.metric("Total ITC - GSTR-2B", f"₹{summary['Total_ITC_2B']:,.2f}")
    r2c3.metric("Overall ITC Difference", f"₹{summary['ITC_Difference']:,.2f}")

    st.divider()

    r3c1, r3c2 = st.columns(2)
    r3c1.metric("Invoices Missing in Books", summary["Total_Missing_Books"])
    r3c2.metric("Invoices Missing in GSTR-2B", summary["Total_Missing_2B"])

    st.divider()

    # ================= INVALID GSTIN SECTION ================= #

    if not results["invalid_gstin"].empty:
        st.warning(f"⚠ {len(results['invalid_gstin'])} invoices skipped due to invalid GSTIN.")
        with st.expander("View Invalid GSTIN Invoices"):
            st.dataframe(results["invalid_gstin"], use_container_width=True)

    # ================= TABS ================= #

    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "✅ Fully Matched", 
        "📕 Missing in Books", 
        "📗 Missing in GSTR-2B",
        "💰 Value Mismatch",
        "🧾 Tax Mismatch",
        "🟡 NO ITC Invoices",
        "📊 Supplier Analysis"
    ])

    with tab1:
        st.dataframe(results["fully_matched"], use_container_width=True)

    with tab2:
        st.dataframe(results["missing_in_books"], use_container_width=True)

    with tab3:
        st.dataframe(results["missing_in_2b"], use_container_width=True)

    with tab4:
        st.dataframe(results["value_mismatch"], use_container_width=True)

    with tab5:
        st.dataframe(results["tax_mismatch"], use_container_width=True)

    with tab6:
        st.dataframe(results["no_itc"], use_container_width=True)

    # ================= SUPPLIER ANALYSIS ================= #

    with tab7:
        st.markdown("### Supplier-wise ITC & Difference Summary")

        combined = results["fully_matched"].copy()

        if not results["value_mismatch"].empty:
            combined = pd.concat([combined, results["value_mismatch"]])

        if not results["tax_mismatch"].empty:
            combined = pd.concat([combined, results["tax_mismatch"]])

        if not combined.empty:

            combined["Supplier"] = combined["Trade_Name_2B"]

            supplier_pivot = combined.groupby("Supplier").agg(
                Total_Invoices=("Invoice_No_2B", "count"),
                ITC_2B=("TOTAL_TAX_2B", "sum"),
                ITC_Books=("TOTAL_TAX_Tally", "sum"),
                Total_Value_Difference=("VALUE_DIFFERENCE", "sum"),
                Total_Tax_Difference=("TAX_DIFFERENCE", "sum")
            ).reset_index()

            supplier_pivot["Absolute_ITC_Difference"] = (
                abs(supplier_pivot["ITC_2B"] - supplier_pivot["ITC_Books"])
            )

            supplier_pivot = supplier_pivot.sort_values(
                "Absolute_ITC_Difference", ascending=False
            )

            st.dataframe(supplier_pivot, use_container_width=True)

        else:
            st.info("No supplier data available.")

    # ================= DOWNLOAD EXCEL ================= #

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        pd.DataFrame([summary]).to_excel(writer, sheet_name="Summary", index=False)

        if not results["fully_matched"].empty:
            results["fully_matched"].to_excel(writer, sheet_name="Fully Matched", index=False)

        if not results["missing_in_books"].empty:
            results["missing_in_books"].to_excel(writer, sheet_name="Missing in Books", index=False)

        if not results["missing_in_2b"].empty:
            results["missing_in_2b"].to_excel(writer, sheet_name="Missing in 2B", index=False)

        if not results["value_mismatch"].empty:
            results["value_mismatch"].to_excel(writer, sheet_name="Value Mismatch", index=False)

        if not results["tax_mismatch"].empty:
            results["tax_mismatch"].to_excel(writer, sheet_name="Tax Mismatch", index=False)

        if not results["no_itc"].empty:
            results["no_itc"].to_excel(writer, sheet_name="NO ITC Invoices", index=False)

        if not results["invalid_gstin"].empty:
            results["invalid_gstin"].to_excel(writer, sheet_name="Invalid GSTIN", index=False)

    output.seek(0)

    st.download_button(
        "⬇ Download Complete Excel Report",
        data=output,
        file_name=f"GST_Reconciliation_Report_{datetime.now().strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
