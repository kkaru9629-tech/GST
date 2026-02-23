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

            # UPDATED RETURN STRUCTURE
            tally_clean, no_itc_df, invalid_gstin_df = parse_tally(tally_raw)
            gstr2b_clean = parse_gstr2b(gstr2b_raw)

            results = reconcile(gstr2b_clean, tally_clean)

            # Store extra data
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

    # ================= INVALID GSTIN SECTION ================= #

    if not results["invalid_gstin"].empty:
        st.warning(f"⚠ {len(results['invalid_gstin'])} invoices skipped due to invalid GSTIN.")

        with st.expander("View Invalid GSTIN Invoices"):
            st.dataframe(results["invalid_gstin"], use_container_width=True)

            csv_invalid = results["invalid_gstin"].to_csv(index=False).encode("utf-8")
            st.download_button(
                "Download Invalid GSTIN Invoices",
                csv_invalid,
                "Invalid_GSTIN_Invoices.csv",
                "text/csv"
            )

    # ================= SUMMARY ================= #

    st.markdown("## 📌 Summary Overview")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("📘 ITC as per Books", f"₹{summary['Total_ITC_Books']:,.2f}")
    col2.metric("📗 ITC as per GSTR-2B", f"₹{summary['Total_ITC_2B']:,.2f}")
    col3.metric("Difference", f"₹{summary['ITC_Difference']:,.2f}")
    col4.metric("Match %", f"{summary['Match_Percentage']} %")

    st.divider()

    # ================= TABS ================= #

    tab1, tab2, tab3, tab4 = st.tabs([
        "✅ Fully Matched",
        "📕 Missing in Books",
        "📗 Missing in GSTR-2B",
        "🟡 NO ITC Invoices"
    ])

    # ===== TAB 1 =====
    with tab1:
        if not results["fully_matched"].empty:
            display = results["fully_matched"][
                ["GSTIN_2B", "Trade_Name_2B",
                 "Invoice_No_2B", "Taxable_Value_2B",
                 "TOTAL_TAX_2B"]
            ].rename(columns={
                "GSTIN_2B": "GSTIN",
                "Trade_Name_2B": "Supplier Name",
                "Invoice_No_2B": "Invoice Number",
                "Taxable_Value_2B": "Taxable Value",
                "TOTAL_TAX_2B": "ITC Amount"
            })
            st.dataframe(display, use_container_width=True)
        else:
            st.info("No fully matched invoices.")

    # ===== TAB 2 =====
    with tab2:
        if not results["missing_in_books"].empty:
            display = results["missing_in_books"][
                ["GSTIN","Trade_Name","Invoice_No","Invoice_Date",
                 "Taxable_Value","TOTAL_TAX"]
            ].rename(columns={
                "Trade_Name":"Supplier Name",
                "Invoice_No":"Invoice Number",
                "Invoice_Date":"Invoice Date",
                "Taxable_Value":"Taxable Value",
                "TOTAL_TAX":"ITC Amount"
            })

            display["Invoice Date"] = pd.to_datetime(display["Invoice Date"]).dt.date
            st.dataframe(display, use_container_width=True)
        else:
            st.success("No invoices missing in Books.")

    # ===== TAB 3 =====
    with tab3:
        if not results["missing_in_2b"].empty:
            display = results["missing_in_2b"][
                ["GSTIN","Trade_Name","Invoice_No","Invoice_Date",
                 "Taxable_Value","TOTAL_TAX"]
            ].rename(columns={
                "Trade_Name":"Supplier Name",
                "Invoice_No":"Invoice Number",
                "Invoice_Date":"Invoice Date",
                "Taxable_Value":"Taxable Value",
                "TOTAL_TAX":"ITC Amount"
            })

            display["Invoice Date"] = pd.to_datetime(display["Invoice Date"]).dt.date
            st.dataframe(display, use_container_width=True)
        else:
            st.success("No invoices missing in GSTR-2B.")

    # ===== TAB 4 - NO ITC =====
    with tab4:
        if not results["no_itc"].empty:

            col1, col2, col3 = st.columns(3)

            col1.metric("Total Invoices", len(results["no_itc"]))
            col2.metric("Total Taxable Value",
                        f"₹{results['no_itc']['Taxable_Value'].sum():,.2f}")
            col3.metric("Total Invoice Value",
                        f"₹{results['no_itc']['Invoice_Value'].sum():,.2f}")

            display_no_itc = results["no_itc"][
                ["GSTIN","Trade_Name","Invoice_No","Invoice_Date",
                 "Taxable_Value","Invoice_Value"]
            ]

            display_no_itc["Invoice_Date"] = pd.to_datetime(display_no_itc["Invoice_Date"]).dt.date

            st.dataframe(display_no_itc, use_container_width=True)

        else:
            st.success("No Zero ITC invoices found.")

    # ================= DOWNLOAD EXCEL ================= #

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

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

    output.seek(0)

    st.download_button(
        "⬇ Download Professional Excel Report",
        data=output,
        file_name=f"GST_Reconciliation_Report_{datetime.now().strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
