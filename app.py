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

    # ================= INDIAN FORMAT SUMMARY ================= #
    
    st.markdown("## 📌 RECONCILIATION SUMMARY")
    st.markdown("══════════════════════════════")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("### 📋 INVOICE COUNT SUMMARY:")
        st.metric("Total Invoices in GSTR-2B", f"{summary['Total_Invoices_2B']:,}")
        st.metric("Total Invoices in Books", f"{summary['Total_Invoices_Books']:,}")
        st.metric("✅ Fully Matched Invoices", f"{summary['Total_Matched']:,}")
        st.metric("❌ Missing in Books", f"{summary['Total_Missing_Books']:,}")
        st.metric("❌ Missing in GSTR-2B", f"{summary['Total_Missing_2B']:,}")
        if 'value_mismatch' in results:
            st.metric("⚠️ Value Mismatch", f"{len(results['value_mismatch']):,}")
        if 'tax_mismatch' in results:
            st.metric("⚠️ Tax Mismatch", f"{len(results['tax_mismatch']):,}")
    
    with col2:
        st.markdown("### 💰 ITC AMOUNT SUMMARY:")
        st.metric("ITC as per GSTR-2B", f"₹{summary['Total_ITC_2B']:,.2f}")
        st.metric("ITC as per Books", f"₹{summary['Total_ITC_Books']:,.2f}")
        st.metric("ITC Difference", f"₹{summary['ITC_Difference']:,.2f}")
    
    with col3:
        st.markdown("### 📈 MATCH PERCENTAGE:")
        st.metric("Match %", f"{summary['Match_Percentage']} %")
        if summary['Match_Percentage'] >= 90:
            st.success("⭐ Excellent Match")
        elif summary['Match_Percentage'] >= 75:
            st.warning("⚠️ Needs Attention")
        else:
            st.error("🔴 Risk - Review Required")
    
    st.markdown("══════════════════════════════")
    st.divider()

    # ================= INVALID GSTIN SECTION ================= #

    if not results["invalid_gstin"].empty:
        st.warning(f"⚠ {len(results['invalid_gstin'])} invoices skipped due to invalid GSTIN.")
        with st.expander("View Invalid GSTIN Invoices"):
            st.dataframe(results["invalid_gstin"], use_container_width=True)

    # ================= TABS ================= #

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "✅ Fully Matched",
        "📕 Missing in Books",
        "📗 Missing in GSTR-2B",
        "💰 Value Mismatch",
        "🧾 Tax Mismatch",
        "🟡 NO ITC Invoices"
    ])

    # ===== TAB 1: Fully Matched =====
    with tab1:
        if not results["fully_matched"].empty:
            display = results["fully_matched"][
                ["GSTIN_2B", "Trade_Name_2B", "Invoice_No_2B", "Invoice_Date_2B", 
                 "Taxable_Value_2B", "TOTAL_TAX_2B"]
            ].copy()
            
            display["Invoice_Date_2B"] = pd.to_datetime(display["Invoice_Date_2B"]).dt.date
            display["Taxable_Value_2B"] = display["Taxable_Value_2B"].apply(lambda x: f"₹{x:,.2f}")
            display["TOTAL_TAX_2B"] = display["TOTAL_TAX_2B"].apply(lambda x: f"₹{x:,.2f}")
            
            display.columns = [
                "GSTIN", "Supplier Name", "Invoice Number", 
                "Invoice Date", "Taxable Value", "ITC Amount"
            ]
            
            st.dataframe(display, use_container_width=True)
        else:
            st.info("No fully matched invoices.")

    # ===== TAB 2: Missing in Books =====
    with tab2:
        if not results["missing_in_books"].empty:
            display = results["missing_in_books"][
                ["GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date", 
                 "Taxable_Value", "TOTAL_TAX"]
            ].copy()
            
            display["Invoice_Date"] = pd.to_datetime(display["Invoice_Date"]).dt.date
            display["Taxable_Value"] = display["Taxable_Value"].apply(lambda x: f"₹{x:,.2f}")
            display["TOTAL_TAX"] = display["TOTAL_TAX"].apply(lambda x: f"₹{x:,.2f}")
            
            display.columns = [
                "GSTIN", "Supplier Name", "Invoice Number", 
                "Invoice Date", "Taxable Value", "ITC Amount"
            ]
            
            st.dataframe(display, use_container_width=True)
        else:
            st.success("No invoices missing in Books.")

    # ===== TAB 3: Missing in GSTR-2B =====
    with tab3:
        if not results["missing_in_2b"].empty:
            display = results["missing_in_2b"][
                ["GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date", 
                 "Taxable_Value", "TOTAL_TAX"]
            ].copy()
            
            display["Invoice_Date"] = pd.to_datetime(display["Invoice_Date"]).dt.date
            display["Taxable_Value"] = display["Taxable_Value"].apply(lambda x: f"₹{x:,.2f}")
            display["TOTAL_TAX"] = display["TOTAL_TAX"].apply(lambda x: f"₹{x:,.2f}")
            
            display.columns = [
                "GSTIN", "Supplier Name", "Invoice Number", 
                "Invoice Date", "Taxable Value", "ITC Amount"
            ]
            
            st.dataframe(display, use_container_width=True)
        else:
            st.success("No invoices missing in GSTR-2B.")

    # ===== TAB 4: Value Mismatch =====
    with tab4:
        if not results["value_mismatch"].empty:
            display = results["value_mismatch"][
                ["GSTIN_2B", "Trade_Name_2B", "Invoice_No_2B", "Invoice_Date_2B",
                 "Taxable_Value_2B", "Taxable_Value_Tally", "VALUE_DIFFERENCE"]
            ].copy()
            
            display["Invoice_Date_2B"] = pd.to_datetime(display["Invoice_Date_2B"]).dt.date
            display["Taxable_Value_2B"] = display["Taxable_Value_2B"].apply(lambda x: f"₹{x:,.2f}")
            display["Taxable_Value_Tally"] = display["Taxable_Value_Tally"].apply(lambda x: f"₹{x:,.2f}")
            display["VALUE_DIFFERENCE"] = display["VALUE_DIFFERENCE"].apply(lambda x: f"₹{x:,.2f}")
            
            display.columns = [
                "GSTIN", "Supplier Name", "Invoice Number", "Invoice Date",
                "Value (2B)", "Value (Books)", "Difference"
            ]
            
            st.dataframe(display, use_container_width=True)
        else:
            st.success("No value mismatches found.")

    # ===== TAB 5: Tax Mismatch =====
    with tab5:
        if not results["tax_mismatch"].empty:
            display = results["tax_mismatch"][
                ["GSTIN_2B", "Trade_Name_2B", "Invoice_No_2B", "Invoice_Date_2B",
                 "TOTAL_TAX_2B", "TOTAL_TAX_Tally", "TAX_DIFFERENCE"]
            ].copy()
            
            display["Invoice_Date_2B"] = pd.to_datetime(display["Invoice_Date_2B"]).dt.date
            display["TOTAL_TAX_2B"] = display["TOTAL_TAX_2B"].apply(lambda x: f"₹{x:,.2f}")
            display["TOTAL_TAX_Tally"] = display["TOTAL_TAX_Tally"].apply(lambda x: f"₹{x:,.2f}")
            display["TAX_DIFFERENCE"] = display["TAX_DIFFERENCE"].apply(lambda x: f"₹{x:,.2f}")
            
            display.columns = [
                "GSTIN", "Supplier Name", "Invoice Number", "Invoice Date",
                "ITC (2B)", "ITC (Books)", "Difference"
            ]
            
            st.dataframe(display, use_container_width=True)
        else:
            st.success("No tax mismatches found.")

    # ===== TAB 6: NO ITC =====
    with tab6:
        if not results["no_itc"].empty:
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Invoices", f"{len(results['no_itc']):,}")
            col2.metric("Total Taxable Value", f"₹{results['no_itc']['Taxable_Value'].sum():,.2f}")
            col3.metric("Total Invoice Value", f"₹{results['no_itc']['Invoice_Value'].sum():,.2f}")
            
            display = results["no_itc"][
                ["GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date", 
                 "Taxable_Value", "Invoice_Value"]
            ].copy()
            
            display["Invoice_Date"] = pd.to_datetime(display["Invoice_Date"]).dt.date
            display["Taxable_Value"] = display["Taxable_Value"].apply(lambda x: f"₹{x:,.2f}")
            display["Invoice_Value"] = display["Invoice_Value"].apply(lambda x: f"₹{x:,.2f}")
            
            display.columns = [
                "GSTIN", "Supplier Name", "Invoice Number", 
                "Invoice Date", "Taxable Value", "Invoice Value"
            ]
            
            st.dataframe(display, use_container_width=True)
        else:
            st.success("No zero ITC invoices found.")

    # ================= DOWNLOAD EXCEL REPORT ================= #

    st.divider()
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        
        # Formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 16,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D9E1F2',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        text_format = workbook.add_format({'border': 1})
        money_format = workbook.add_format({'num_format': '₹#,##0.00', 'border': 1})
        
        # ===== EXECUTIVE SUMMARY SHEET =====
        summary_sheet = workbook.add_worksheet("Executive Summary")
        
        # Title
        summary_sheet.merge_range("A1:B1", "GST RECONCILIATION SUMMARY REPORT", title_format)
        summary_sheet.set_row(0, 30)
        
        # Summary data
        summary_data = [
            ["══════════════════════════════", ""],
            ["📋 INVOICE COUNT SUMMARY:", ""],
            ["Total Invoices in GSTR-2B", summary["Total_Invoices_2B"]],
            ["Total Invoices in Books", summary["Total_Invoices_Books"]],
            ["✅ Fully Matched Invoices", summary["Total_Matched"]],
            ["❌ Missing in Books", summary["Total_Missing_Books"]],
            ["❌ Missing in GSTR-2B", summary["Total_Missing_2B"]],
            ["⚠️ Value Mismatch", len(results.get("value_mismatch", pd.DataFrame()))],
            ["⚠️ Tax Mismatch", len(results.get("tax_mismatch", pd.DataFrame()))],
            ["", ""],
            ["💰 ITC AMOUNT SUMMARY:", ""],
            ["ITC as per GSTR-2B", summary["Total_ITC_2B"]],
            ["ITC as per Books", summary["Total_ITC_Books"]],
            ["ITC Difference", summary["ITC_Difference"]],
            ["", ""],
            ["📈 Match Percentage", f"{summary['Match_Percentage']} %"],
            ["══════════════════════════════", ""]
        ]
        
        row = 2
        for i, (label, value) in enumerate(summary_data):
            if "═════" in label or label == "":
                summary_sheet.write(row, 0, label, text_format)
                summary_sheet.write(row, 1, value, text_format)
            elif "📋" in label or "💰" in label or "📈" in label:
                summary_sheet.write(row, 0, label, title_format)
            elif "ITC" in label:
                summary_sheet.write(row, 0, label, text_format)
                summary_sheet.write(row, 1, value, money_format)
            else:
                summary_sheet.write(row, 0, label, text_format)
                summary_sheet.write(row, 1, value, text_format)
            row += 1
        
        summary_sheet.set_column("A:A", 40)
        summary_sheet.set_column("B:B", 25)
        
        # ===== FUNCTION TO WRITE SHEETS =====
        def write_formatted_sheet(df, sheet_name, columns_mapping):
            if df.empty:
                return
            
            df_to_write = df[list(columns_mapping.keys())].copy()
            df_to_write.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
            
            worksheet = writer.sheets[sheet_name]
            
            # Write title
            worksheet.write(0, 0, f"{sheet_name} Invoices", title_format)
            worksheet.merge_range(0, 0, 0, len(columns_mapping)-1, f"{sheet_name} Invoices", title_format)
            
            # Write headers
            for col_num, (_, new_name) in enumerate(columns_mapping.items()):
                worksheet.write(1, col_num, new_name, header_format)
            
            # Format columns
            for col_num, (col_name, _) in enumerate(columns_mapping.items()):
                if "Value" in col_name or "ITC" in col_name or "Tax" in col_name or "Amount" in col_name:
                    worksheet.set_column(col_num, col_num, 18, money_format)
                else:
                    worksheet.set_column(col_num, col_num, 20)
            
            worksheet.freeze_panes(2, 0)
        
        # ===== WRITE ALL SHEETS =====
        
        # Fully Matched
        if not results["fully_matched"].empty:
            cols = {
                "GSTIN_2B": "GSTIN",
                "Trade_Name_2B": "Supplier Name",
                "Invoice_No_2B": "Invoice Number",
                "Invoice_Date_2B": "Invoice Date",
                "Taxable_Value_2B": "Taxable Value",
                "TOTAL_TAX_2B": "ITC Amount"
            }
            write_formatted_sheet(results["fully_matched"], "✅ Fully Matched", cols)
        
        # Missing in Books
        if not results["missing_in_books"].empty:
            cols = {
                "GSTIN": "GSTIN",
                "Trade_Name": "Supplier Name",
                "Invoice_No": "Invoice Number",
                "Invoice_Date": "Invoice Date",
                "Taxable_Value": "Taxable Value",
                "TOTAL_TAX": "ITC Amount"
            }
            write_formatted_sheet(results["missing_in_books"], "📕 Missing in Books", cols)
        
        # Missing in 2B
        if not results["missing_in_2b"].empty:
            cols = {
                "GSTIN": "GSTIN",
                "Trade_Name": "Supplier Name",
                "Invoice_No": "Invoice Number",
                "Invoice_Date": "Invoice Date",
                "Taxable_Value": "Taxable Value",
                "TOTAL_TAX": "ITC Amount"
            }
            write_formatted_sheet(results["missing_in_2b"], "📗 Missing in GSTR-2B", cols)
        
        # Value Mismatch
        if not results["value_mismatch"].empty:
            cols = {
                "GSTIN_2B": "GSTIN",
                "Trade_Name_2B": "Supplier Name",
                "Invoice_No_2B": "Invoice Number",
                "Invoice_Date_2B": "Invoice Date",
                "Taxable_Value_2B": "Value (2B)",
                "Taxable_Value_Tally": "Value (Books)",
                "VALUE_DIFFERENCE": "Difference"
            }
            write_formatted_sheet(results["value_mismatch"], "💰 Value Mismatch", cols)
        
        # Tax Mismatch
        if not results["tax_mismatch"].empty:
            cols = {
                "GSTIN_2B": "GSTIN",
                "Trade_Name_2B": "Supplier Name",
                "Invoice_No_2B": "Invoice Number",
                "Invoice_Date_2B": "Invoice Date",
                "TOTAL_TAX_2B": "ITC (2B)",
                "TOTAL_TAX_Tally": "ITC (Books)",
                "TAX_DIFFERENCE": "Difference"
            }
            write_formatted_sheet(results["tax_mismatch"], "🧾 Tax Mismatch", cols)
        
        # NO ITC
        if not results["no_itc"].empty:
            cols = {
                "GSTIN": "GSTIN",
                "Trade_Name": "Supplier Name",
                "Invoice_No": "Invoice Number",
                "Invoice_Date": "Invoice Date",
                "Taxable_Value": "Taxable Value",
                "Invoice_Value": "Invoice Value"
            }
            write_formatted_sheet(results["no_itc"], "🟡 NO ITC Invoices", cols)
        
        # Invalid GSTIN
        if not results["invalid_gstin"].empty:
            cols = {
                "GSTIN": "GSTIN",
                "Trade_Name": "Supplier Name",
                "Invoice_No": "Invoice Number",
                "Invoice_Date": "Invoice Date",
                "Taxable_Value": "Taxable Value",
                "Invoice_Value": "Invoice Value",
                "TOTAL_TAX": "ITC Amount"
            }
            write_formatted_sheet(results["invalid_gstin"], "⚠️ Invalid GSTIN", cols)
    
    output.seek(0)
    
    st.download_button(
        "⬇ Download Excel Report",
        data=output,
        file_name=f"GST_Reconciliation_{datetime.now().strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
