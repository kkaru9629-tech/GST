import streamlit as st
import pandas as pd
import io
from reconciliation_engine import parse_tally, parse_gstr2b, reconcile
from datetime import datetime

st.set_page_config(page_title="GST Reconciliation", layout="wide")

st.title("GST Reconciliation System")
st.caption("ITC Books vs GSTR-2B Comparison")

col1, col2 = st.columns(2)

with col1:
    tally_file = st.file_uploader("Upload Tally Purchase Register", type=["xlsx", "xls", "csv"])

with col2:
    gstr2b_file = st.file_uploader("Upload Structured GSTR-2B", type=["xlsx", "xls", "csv"])

if st.button("Run Reconciliation", use_container_width=True):
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
            st.success("Reconciliation Completed Successfully")

        except Exception as e:
            st.error(str(e))


# ================= DISPLAY ================= #

if "results" in st.session_state:

    results = st.session_state["results"]
    summary = results["summary"]

    # ================= SUMMARY ================= #

    st.markdown("## Executive Summary")

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Invoices - Books", summary["Total_Invoices_Books"])
    c2.metric("Total Invoices - GSTR-2B", summary["Total_Invoices_2B"])
    c3.metric("Fully Matched Invoices", summary["Total_Matched"])

    st.divider()

    c4, c5, c6 = st.columns(3)
    c4.metric("Total ITC - Books", summary["Total_ITC_Books"])
    c5.metric("Total ITC - GSTR-2B", summary["Total_ITC_2B"])
    c6.metric("Overall ITC Difference", summary["ITC_Difference"])

    st.divider()

    # ================= TABS ================= #

    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "Fully Matched",
        "Missing in Books",
        "Missing in GSTR-2B",
        "Value Mismatch",
        "Tax Mismatch",
        "NO ITC Invoices",
        "Supplier Analysis"
    ])

    # ---------- FULLY MATCHED ----------
    with tab1:
        if not results["fully_matched"].empty:
            df = results["fully_matched"][[
                "GSTIN_2B","Trade_Name_2B","Invoice_No_2B",
                "Invoice_Date_2B","Taxable_Value_2B","TOTAL_TAX_2B"
            ]].copy()

            df["Invoice_Date_2B"] = pd.to_datetime(df["Invoice_Date_2B"]).dt.date

            df.columns = [
                "GSTIN","Supplier Name","Invoice Number",
                "Invoice Date","Taxable Value","ITC Amount"
            ]

            st.dataframe(df, use_container_width=True)
        else:
            st.info("No fully matched invoices")

    # ---------- MISSING IN BOOKS ----------
    with tab2:
        if not results["missing_in_books"].empty:
            df = results["missing_in_books"][[
                "GSTIN","Trade_Name","Invoice_No",
                "Invoice_Date","Taxable_Value","TOTAL_TAX"
            ]].copy()

            df["Invoice_Date"] = pd.to_datetime(df["Invoice_Date"]).dt.date

            df.columns = [
                "GSTIN","Supplier Name","Invoice Number",
                "Invoice Date","Taxable Value","ITC Amount"
            ]

            st.dataframe(df, use_container_width=True)
        else:
            st.info("No invoices missing in Books")

    # ---------- MISSING IN GSTR-2B ----------
    with tab3:
        if not results["missing_in_2b"].empty:
            df = results["missing_in_2b"][[
                "GSTIN","Trade_Name","Invoice_No",
                "Invoice_Date","Taxable_Value","TOTAL_TAX"
            ]].copy()

            df["Invoice_Date"] = pd.to_datetime(df["Invoice_Date"]).dt.date

            df.columns = [
                "GSTIN","Supplier Name","Invoice Number",
                "Invoice Date","Taxable Value","ITC Amount"
            ]

            st.dataframe(df, use_container_width=True)
        else:
            st.info("No invoices missing in GSTR-2B")

    # ---------- VALUE MISMATCH ----------
    with tab4:
        if not results["value_mismatch"].empty:
            df = results["value_mismatch"][[
                "GSTIN_2B","Trade_Name_2B","Invoice_No_2B",
                "Invoice_Date_2B",
                "Taxable_Value_2B","Taxable_Value_Tally","VALUE_DIFFERENCE"
            ]].copy()

            df["Invoice_Date_2B"] = pd.to_datetime(df["Invoice_Date_2B"]).dt.date

            df.columns = [
                "GSTIN","Supplier Name","Invoice Number",
                "Invoice Date","Value as per GSTR-2B",
                "Value as per Books","Value Difference"
            ]

            st.dataframe(df, use_container_width=True)
        else:
            st.info("No value mismatches found")

    # ---------- TAX MISMATCH ----------
    with tab5:
        if not results["tax_mismatch"].empty:
            df = results["tax_mismatch"][[
                "GSTIN_2B","Trade_Name_2B","Invoice_No_2B",
                "Invoice_Date_2B",
                "TOTAL_TAX_2B","TOTAL_TAX_Tally","TAX_DIFFERENCE"
            ]].copy()

            df["Invoice_Date_2B"] = pd.to_datetime(df["Invoice_Date_2B"]).dt.date

            df.columns = [
                "GSTIN","Supplier Name","Invoice Number",
                "Invoice Date","ITC as per GSTR-2B",
                "ITC as per Books","ITC Difference"
            ]

            st.dataframe(df, use_container_width=True)
        else:
            st.info("No tax mismatches found")

    # ---------- NO ITC ----------
    with tab6:
        if not results["no_itc"].empty:
            df = results["no_itc"][[
                "GSTIN","Trade_Name","Invoice_No",
                "Invoice_Date","Taxable_Value","Invoice_Value"
            ]].copy()

            df["Invoice_Date"] = pd.to_datetime(df["Invoice_Date"]).dt.date

            df.columns = [
                "GSTIN","Supplier Name","Invoice Number",
                "Invoice Date","Taxable Value","Invoice Value"
            ]

            st.dataframe(df, use_container_width=True)
        else:
            st.info("No zero ITC invoices")

    # ---------- SUPPLIER ANALYSIS ----------
    with tab7:
        combined = results["fully_matched"].copy()

        if not results["value_mismatch"].empty:
            combined = pd.concat([combined, results["value_mismatch"]])

        if not results["tax_mismatch"].empty:
            combined = pd.concat([combined, results["tax_mismatch"]])

        if not combined.empty:
            combined["Supplier"] = combined["Trade_Name_2B"]

            pivot = combined.groupby("Supplier").agg(
                Total_Invoices=("Invoice_No_2B", "count"),
                ITC_2B=("TOTAL_TAX_2B", "sum"),
                ITC_Books=("TOTAL_TAX_Tally", "sum"),
            ).reset_index()

            pivot["ITC Difference"] = pivot["ITC_2B"] - pivot["ITC_Books"]

            st.dataframe(pivot, use_container_width=True)
        else:
            st.info("No supplier data available")

    # ================= DOWNLOAD ================= #

    st.divider()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        pd.DataFrame([summary]).to_excel(writer, sheet_name="Summary", index=False)

        for key, name in [
            ("fully_matched", "Fully Matched"),
            ("missing_in_books", "Missing in Books"),
            ("missing_in_2b", "Missing in 2B"),
            ("value_mismatch", "Value Mismatch"),
            ("tax_mismatch", "Tax Mismatch"),
            ("no_itc", "NO ITC Invoices"),
            ("invalid_gstin", "Invalid GSTIN")
        ]:
            if key in results and not results[key].empty:
                results[key].to_excel(writer, sheet_name=name, index=False)

    output.seek(0)

    st.download_button(
        "Download Complete Excel Report",
        data=output,
        file_name=f"GST_Reconciliation_Report_{datetime.now().strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
