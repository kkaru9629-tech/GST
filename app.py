import streamlit as st
import pandas as pd
import io
from reconciliation_engine import parse_tally, parse_gstr2b, reconcile

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

            tally_clean = parse_tally(tally_raw)
            gstr2b_clean = parse_gstr2b(gstr2b_raw)

            results = reconcile(gstr2b_clean, tally_clean)
            st.session_state["results"] = results
            st.success("Reconciliation Completed Successfully.")

        except Exception as e:
            st.error(str(e))


# ================= DISPLAY ================= #

if "results" in st.session_state:

    results = st.session_state["results"]
    summary = results["summary"]

    st.markdown("## 📌 Summary Overview")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("📘 ITC as per Books", f"₹{summary['Total_ITC_Books']:,.2f}")
    col2.metric("📗 ITC as per GSTR-2B", f"₹{summary['Total_ITC_2B']:,.2f}")
    col3.metric("Difference", f"₹{summary['ITC_Difference']:,.2f}")
    col4.metric("Match %", f"{summary['Match_Percentage']} %")

    st.divider()

    tab1, tab2, tab3 = st.tabs([
        "✅ Fully Matched",
        "❗ Missing / Value Issues",
        "💸 Tax Mismatch"
    ])

    with tab1:
        if not results["fully_matched"].empty:
            display = results["fully_matched"][[
                "GSTIN_2B", "Trade_Name_2B",
                "Invoice_No_2B",
                "Taxable_Value_2B",
                "TOTAL_TAX_2B"
            ]]
            st.dataframe(display)
        else:
            st.info("No fully matched invoices.")

    with tab2:
        combined = pd.concat([
            results["missing_in_books"],
            results["missing_in_2b"],
            results["value_mismatch"]
        ])
        if not combined.empty:
            st.dataframe(combined)
        else:
            st.success("No missing or value mismatch.")

    with tab3:
        if not results["tax_mismatch"].empty:
            simple = results["tax_mismatch"][[
                "GSTIN_2B",
                "Invoice_No_2B",
                "TOTAL_TAX_2B",
                "TOTAL_TAX_Tally",
                "TAX_DIFFERENCE"
            ]]
            st.dataframe(simple)
        else:
            st.success("No tax mismatch.")

    # ================= SIMPLIFIED EXCEL REPORT ================= #

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        
        # Formats
        header_format = workbook.add_format({
            "bold": True,
            "bg_color": "#4472C4",
            "font_color": "white",
            "border": 1
        })
        
        money_format = workbook.add_format({
            "num_format": "₹#,##0.00",
            "border": 1
        })
        
        def write_sheet(df, name, columns_to_show=None):
            """Write dataframe with clean column selection"""
            if df.empty:
                pd.DataFrame({"Message": ["No records found"]}).to_excel(writer, sheet_name=name, index=False)
                return
                
            # Select only relevant columns if specified
            if columns_to_show:
                display_df = df[columns_to_show].copy()
            else:
                display_df = df.copy()
                
            display_df.to_excel(writer, sheet_name=name, index=False)
            worksheet = writer.sheets[name]
            
            # Format headers
            for col_num, column in enumerate(display_df.columns):
                worksheet.write(0, col_num, column, header_format)
                
                # Auto-width
                max_len = max(
                    display_df[column].astype(str).map(len).max(),
                    len(column)
                ) + 2
                worksheet.set_column(col_num, col_num, max_len)
                
                # Format money columns
                if any(x in column.lower() for x in ['value', 'tax', 'itc', 'amount', 'diff']):
                    worksheet.set_column(col_num, col_num, max_len, money_format)
        
        # ===== SHEET 1: EXECUTIVE SUMMARY ===== #
        summary_data = {
            "Particulars": [
                "Total Invoices in GSTR-2B",
                "Total Invoices in Books",
                "✅ Fully Matched Invoices",
                "❌ Missing in Books",
                "❌ Missing in GSTR-2B",
                "⚠️ Value Mismatch",
                "⚠️ Tax Mismatch",
                "",
                "ITC as per GSTR-2B",
                "ITC as per Books",
                "ITC Difference",
                "Match Percentage"
            ],
            "Count": [
                summary["Total_Invoices_2B"],
                summary["Total_Invoices_Books"],
                summary["Total_Matched"],
                summary["Total_Missing_Books"],
                summary["Total_Missing_2B"],
                len(results["value_mismatch"]),
                len(results["tax_mismatch"]),
                "",
                "",
                "",
                "",
                f"{summary['Match_Percentage']}%"
            ],
            "Amount (₹)": [
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                f"₹{summary['Total_ITC_2B']:,.2f}",
                f"₹{summary['Total_ITC_Books']:,.2f}",
                f"₹{summary['ITC_Difference']:,.2f}",
                ""
            ]
        }
        
        pd.DataFrame(summary_data).to_excel(writer, sheet_name="Executive Summary", index=False)
        
        # Format summary sheet
        summary_sheet = writer.sheets["Executive Summary"]
        summary_sheet.set_column("A:A", 35)
        summary_sheet.set_column("B:B", 15)
        summary_sheet.set_column("C:C", 20)
        
        # ===== SHEET 2: FULLY MATCHED (SIMPLIFIED) ===== #
        if not results["fully_matched"].empty:
            simple_matched = results["fully_matched"][[
                "GSTIN_2B", 
                "Trade_Name_2B",
                "Invoice_No_2B",
                "Invoice_Date_2B",
                "Taxable_Value_2B",
                "TOTAL_TAX_2B"
            ]].rename(columns={
                "GSTIN_2B": "GSTIN",
                "Trade_Name_2B": "Supplier Name",
                "Invoice_No_2B": "Invoice Number",
                "Invoice_Date_2B": "Invoice Date",
                "Taxable_Value_2B": "Taxable Value",
                "TOTAL_TAX_2B": "ITC Amount"
            })
            write_sheet(simple_matched, "✅ Fully Matched")
        
        # ===== SHEET 3: MISSING IN BOOKS ===== #
        if not results["missing_in_books"].empty:
            missing_books_simple = results["missing_in_books"][[
                "GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date",
                "Taxable_Value", "TOTAL_TAX"
            ]].rename(columns={
                "Trade_Name": "Supplier Name",
                "Invoice_No": "Invoice Number",
                "Invoice_Date": "Invoice Date",
                "Taxable_Value": "Taxable Value",
                "TOTAL_TAX": "ITC Amount"
            })
            write_sheet(missing_books_simple, "❌ Missing in Books")
        
        # ===== SHEET 4: MISSING IN GSTR-2B ===== #
        if not results["missing_in_2b"].empty:
            missing_2b_simple = results["missing_in_2b"][[
                "GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date",
                "Taxable_Value", "TOTAL_TAX"
            ]].rename(columns={
                "Trade_Name": "Supplier Name",
                "Invoice_No": "Invoice Number",
                "Invoice_Date": "Invoice Date",
                "Taxable_Value": "Taxable Value",
                "TOTAL_TAX": "ITC Amount"
            })
            write_sheet(missing_2b_simple, "❌ Missing in GSTR-2B")
        
        # ===== SHEET 5: VALUE MISMATCH ===== #
        if not results["value_mismatch"].empty:
            value_mismatch_simple = results["value_mismatch"][[
                "GSTIN_2B",
                "Trade_Name_2B",
                "Invoice_No_2B",
                "Invoice_Date_2B",
                "Taxable_Value_2B",
                "Taxable_Value_Tally",
                "VALUE_DIFFERENCE"
            ]].rename(columns={
                "GSTIN_2B": "GSTIN",
                "Trade_Name_2B": "Supplier Name",
                "Invoice_No_2B": "Invoice Number",
                "Invoice_Date_2B": "Invoice Date",
                "Taxable_Value_2B": "Value as per GSTR-2B",
                "Taxable_Value_Tally": "Value as per Books",
                "VALUE_DIFFERENCE": "Difference"
            })
            write_sheet(value_mismatch_simple, "⚠️ Value Mismatch")
        
        # ===== SHEET 6: TAX MISMATCH ===== #
        if not results["tax_mismatch"].empty:
            tax_mismatch_simple = results["tax_mismatch"][[
                "GSTIN_2B",
                "Trade_Name_2B",
                "Invoice_No_2B",
                "Invoice_Date_2B",
                "TOTAL_TAX_2B",
                "TOTAL_TAX_Tally",
                "TAX_DIFFERENCE"
            ]].rename(columns={
                "GSTIN_2B": "GSTIN",
                "Trade_Name_2B": "Supplier Name",
                "Invoice_No_2B": "Invoice Number",
                "Invoice_Date_2B": "Invoice Date",
                "TOTAL_TAX_2B": "ITC as per GSTR-2B",
                "TOTAL_TAX_Tally": "ITC as per Books",
                "TAX_DIFFERENCE": "ITC Difference"
            })
            write_sheet(tax_mismatch_simple, "⚠️ Tax Mismatch")

    output.seek(0)

    st.download_button(
        "⬇ Download Professional Excel Report",
        data=output,
        file_name="GST_Reconciliation_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
