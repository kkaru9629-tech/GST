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

    # ===== SUMMARY (UNCHANGED STRUCTURE) =====
    st.markdown("## 📊 Executive Summary")

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Invoices - Books", summary["Total_Invoices_Books"])
    c2.metric("Total Invoices - GSTR-2B", summary["Total_Invoices_2B"])
    c3.metric("Fully Matched Invoices", summary["Total_Matched"])

    st.divider()

    c4, c5, c6 = st.columns(3)
    c4.metric("Total ITC - Books", f"₹{summary['Total_ITC_Books']:,.2f}")
    c5.metric("Total ITC - GSTR-2B", f"₹{summary['Total_ITC_2B']:,.2f}")
    c6.metric("Overall ITC Difference", f"₹{summary['ITC_Difference']:,.2f}")

    st.divider()

    # Helper function for formatting
    def format_currency(df, columns):
        for col in columns:
            df[col] = df[col].apply(lambda x: f"₹{x:,.2f}")
        return df

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

    # FULLY MATCHED
    with tab1:
        if not results["fully_matched"].empty:
            df = results["fully_matched"][[
                "GSTIN_2B","Trade_Name_2B","Invoice_No_2B",
                "Invoice_Date_2B","Taxable_Value_2B","TOTAL_TAX_2B"
            ]].copy()

            df["Invoice_Date_2B"] = pd.to_datetime(df["Invoice_Date_2B"]).dt.date

            df = df.rename(columns={
                "GSTIN_2B":"GSTIN",
                "Trade_Name_2B":"Supplier Name",
                "Invoice_No_2B":"Invoice Number",
                "Invoice_Date_2B":"Invoice Date",
                "Taxable_Value_2B":"Taxable Value",
                "TOTAL_TAX_2B":"ITC Amount"
            })

            df = format_currency(df, ["Taxable Value", "ITC Amount"])

            st.dataframe(df, use_container_width=True)

    # MISSING IN BOOKS
    with tab2:
        if not results["missing_in_books"].empty:
            df = results["missing_in_books"][[
                "GSTIN","Trade_Name","Invoice_No",
                "Invoice_Date","Taxable_Value","TOTAL_TAX"
            ]].copy()

            df["Invoice_Date"] = pd.to_datetime(df["Invoice_Date"]).dt.date

            df = df.rename(columns={
                "Trade_Name":"Supplier Name",
                "Invoice_No":"Invoice Number",
                "Invoice_Date":"Invoice Date",
                "Taxable_Value":"Taxable Value",
                "TOTAL_TAX":"ITC Amount"
            })

            df = format_currency(df, ["Taxable Value", "ITC Amount"])

            st.dataframe(df, use_container_width=True)

    # MISSING IN 2B
    with tab3:
        if not results["missing_in_2b"].empty:
            df = results["missing_in_2b"][[
                "GSTIN","Trade_Name","Invoice_No",
                "Invoice_Date","Taxable_Value","TOTAL_TAX"
            ]].copy()

            df["Invoice_Date"] = pd.to_datetime(df["Invoice_Date"]).dt.date

            df = df.rename(columns={
                "Trade_Name":"Supplier Name",
                "Invoice_No":"Invoice Number",
                "Invoice_Date":"Invoice Date",
                "Taxable_Value":"Taxable Value",
                "TOTAL_TAX":"ITC Amount"
            })

            df = format_currency(df, ["Taxable Value", "ITC Amount"])

            st.dataframe(df, use_container_width=True)

    # VALUE MISMATCH
    with tab4:
        if not results["value_mismatch"].empty:
            df = results["value_mismatch"][[
                "GSTIN_2B","Trade_Name_2B","Invoice_No_2B",
                "Invoice_Date_2B",
                "Taxable_Value_2B","Taxable_Value_Tally","VALUE_DIFFERENCE"
            ]].copy()

            df["Invoice_Date_2B"] = pd.to_datetime(df["Invoice_Date_2B"]).dt.date

            df = df.rename(columns={
                "GSTIN_2B":"GSTIN",
                "Trade_Name_2B":"Supplier Name",
                "Invoice_No_2B":"Invoice Number",
                "Invoice_Date_2B":"Invoice Date",
                "Taxable_Value_2B":"Value as per GSTR-2B",
                "Taxable_Value_Tally":"Value as per Books",
                "VALUE_DIFFERENCE":"Value Difference"
            })

            df = format_currency(df, [
                "Value as per GSTR-2B",
                "Value as per Books",
                "Value Difference"
            ])

            st.dataframe(df, use_container_width=True)

    # TAX MISMATCH
    with tab5:
        if not results["tax_mismatch"].empty:
            df = results["tax_mismatch"][[
                "GSTIN_2B","Trade_Name_2B","Invoice_No_2B",
                "Invoice_Date_2B",
                "TOTAL_TAX_2B","TOTAL_TAX_Tally","TAX_DIFFERENCE"
            ]].copy()

            df["Invoice_Date_2B"] = pd.to_datetime(df["Invoice_Date_2B"]).dt.date

            df = df.rename(columns={
                "GSTIN_2B":"GSTIN",
                "Trade_Name_2B":"Supplier Name",
                "Invoice_No_2B":"Invoice Number",
                "Invoice_Date_2B":"Invoice Date",
                "TOTAL_TAX_2B":"ITC as per GSTR-2B",
                "TOTAL_TAX_Tally":"ITC as per Books",
                "TAX_DIFFERENCE":"ITC Difference"
            })

            df = format_currency(df, [
                "ITC as per GSTR-2B",
                "ITC as per Books",
                "ITC Difference"
            ])

            st.dataframe(df, use_container_width=True)

    # NO ITC
    with tab6:
        if not results["no_itc"].empty:
            df = results["no_itc"][[
                "GSTIN","Trade_Name","Invoice_No",
                "Invoice_Date","Taxable_Value","Invoice_Value"
            ]].copy()

            df["Invoice_Date"] = pd.to_datetime(df["Invoice_Date"]).dt.date

            df = df.rename(columns={
                "Trade_Name":"Supplier Name",
                "Invoice_No":"Invoice Number",
                "Invoice_Date":"Invoice Date"
            })

            df = format_currency(df, ["Taxable Value", "Invoice Value"])

            st.dataframe(df, use_container_width=True)

    # ================= DOWNLOAD EXCEL ================= #

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
        "⬇ Download Complete Excel Report",
        data=output,
        file_name=f"GST_Reconciliation_Report_{datetime.now().strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
