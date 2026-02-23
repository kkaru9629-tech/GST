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
            with st.spinner("Processing files..."):
                tally_raw = pd.read_csv(tally_file) if tally_file.name.endswith("csv") else pd.read_excel(tally_file)
                gstr2b_raw = pd.read_csv(gstr2b_file) if gstr2b_file.name.endswith("csv") else pd.read_excel(gstr2b_file)

                tally_clean = parse_tally(tally_raw)
                gstr2b_clean = parse_gstr2b(gstr2b_raw)
                
                # Check for invalid GSTINs
                if hasattr(tally_clean, 'attrs') and tally_clean.attrs.get('invalid_count', 0) > 0:
                    st.session_state["invalid_gstin"] = tally_clean.attrs['invalid_invoices']
                    st.session_state["invalid_count"] = tally_clean.attrs['invalid_count']
                    st.warning(f"⚠️ Skipped {tally_clean.attrs['invalid_count']} invoices with invalid/missing GSTIN. These won't be reconciled.")
                
                results = reconcile(gstr2b_clean, tally_clean)
                st.session_state["results"] = results
                st.success("✅ Reconciliation Completed Successfully!")

        except Exception as e:
            st.error(f"❌ Error: {str(e)}")


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

    # Add NO ITC metrics
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("🟡 Invoices with NO ITC", f"{summary['Total_No_ITC']}")
    col2.metric("✅ NO ITC - In 2B", f"{summary['No_ITC_with_2B']}")
    col3.metric("❌ NO ITC - Not in 2B", f"{summary['No_ITC_without_2B']}")
    col4.metric("", "")

    st.divider()

    # Updated tabs with NO ITC
    tab1, tab2, tab3, tab4 = st.tabs([
        "✅ Fully Matched",
        "📕 Missing in Books",
        "📗 Missing in GSTR-2B",
        "🟡 NO ITC Invoices"
    ])

    with tab1:
        if not results["fully_matched"].empty:
            display = results["fully_matched"][[
                "GSTIN_2B", "Trade_Name_2B",
                "Invoice_No_2B",
                "Taxable_Value_2B",
                "TOTAL_TAX_2B"
            ]].rename(columns={
                "GSTIN_2B": "GSTIN",
                "Trade_Name_2B": "Supplier Name",
                "Invoice_No_2B": "Invoice Number",
                "Taxable_Value_2B": "Taxable Value",
                "TOTAL_TAX_2B": "ITC Amount"
            })
            st.dataframe(display, use_container_width=True)
        else:
            st.info("No fully matched invoices.")

    with tab2:
        if not results["missing_in_books"].empty:
            display = results["missing_in_books"][[
                "GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date",
                "Taxable_Value", "TOTAL_TAX"
            ]].rename(columns={
                "Trade_Name": "Supplier Name",
                "Invoice_No": "Invoice Number",
                "Invoice_Date": "Invoice Date",
                "Taxable_Value": "Taxable Value",
                "TOTAL_TAX": "ITC Amount"
            })
            display["Invoice Date"] = pd.to_datetime(display["Invoice Date"]).dt.date
            st.dataframe(display, use_container_width=True)
        else:
            st.success("No invoices missing in Books.")

    with tab3:
        if not results["missing_in_2b"].empty:
            display = results["missing_in_2b"][[
                "GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date",
                "Taxable_Value", "TOTAL_TAX"
            ]].rename(columns={
                "Trade_Name": "Supplier Name",
                "Invoice_No": "Invoice Number",
                "Invoice_Date": "Invoice Date",
                "Taxable_Value": "Taxable Value",
                "TOTAL_TAX": "ITC Amount"
            })
            display["Invoice Date"] = pd.to_datetime(display["Invoice Date"]).dt.date
            st.dataframe(display, use_container_width=True)
        else:
            st.success("No invoices missing in GSTR-2B.")

    with tab4:
        if not results["no_itc"].empty:
            st.subheader("📊 NO ITC Invoices Summary")
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("✅ Present in GSTR-2B", f"{summary['No_ITC_with_2B']}")
            with col2:
                st.metric("❌ Not in GSTR-2B", f"{summary['No_ITC_without_2B']}")
            
            st.divider()
            
            # Show all NO ITC invoices
            display = results["no_itc"][[
                "GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date",
                "Taxable_Value", "Invoice_Value"
            ]].rename(columns={
                "Trade_Name": "Supplier Name",
                "Invoice_No": "Invoice Number",
                "Invoice_Date": "Invoice Date",
                "Taxable_Value": "Taxable Value",
                "Invoice_Value": "Invoice Value"
            })
            display["Invoice Date"] = pd.to_datetime(display["Invoice Date"]).dt.date
            
            # Add status column
            def get_status(row):
                key = row["GSTIN"] + "|" + str(row["Invoice Number"])
                no_itc_with_2b_keys = results["no_itc_with_2b"]["KEY"].tolist() if not results["no_itc_with_2b"].empty else []
                if key in no_itc_with_2b_keys:
                    return "✅ In GSTR-2B"
                else:
                    return "❌ Not in GSTR-2B"
            
            display["Status"] = display.apply(get_status, axis=1)
            
            st.dataframe(display, use_container_width=True)
        else:
            st.success("No NO ITC invoices found.")

    # ===== INVALID GSTIN SECTION ===== #
    if "invalid_gstin" in st.session_state and not st.session_state["invalid_gstin"].empty:
        st.divider()
        with st.expander(f"⚠️ {st.session_state['invalid_count']} Invoices Skipped Due to Invalid/Missing GSTIN", expanded=False):
            st.caption("These invoices were not reconciled. Please update GSTIN in Tally and re-upload.")
            
            # Format date for display
            invalid_display = st.session_state["invalid_gstin"].copy()
            if "Invoice_Date" in invalid_display.columns:
                invalid_display["Invoice_Date"] = pd.to_datetime(invalid_display["Invoice_Date"]).dt.date
            
            # Add a download button for invalid invoices
            invalid_output = io.BytesIO()
            with pd.ExcelWriter(invalid_output, engine="xlsxwriter") as writer:
                invalid_display.to_excel(writer, sheet_name="Invalid GSTIN", index=False)
            invalid_output.seek(0)
            
            col1, col2 = st.columns([3, 1])
            with col1:
                st.dataframe(invalid_display, use_container_width=True)
            with col2:
                st.download_button(
                    "📥 Download List",
                    data=invalid_output,
                    file_name=f"Invalid_GSTIN_{datetime.now().strftime('%d%m%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    # ================= ENHANCED EXCEL REPORT ================= #

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        
        # Font configuration
        try:
            workbook.add_format({'font_name': 'Aptos Narrow'})
            font_name = 'Aptos Narrow'
        except:
            font_name = 'Calibri'
        
        # Formats
        summary_header = workbook.add_format({
            "bold": True,
            "font_name": font_name,
            "font_size": 11,
            "bg_color": "#4472C4",
            "font_color": "white",
            "border": 0,
            "align": "center",
            "valign": "vcenter"
        })
        
        regular_header = workbook.add_format({
            "bold": True,
            "font_name": font_name,
            "font_size": 11,
            "bg_color": "#D3D3D3",
            "font_color": "black",
            "border": 0,
            "align": "center",
            "valign": "vcenter"
        })
        
        money_format = workbook.add_format({
            "num_format": "₹#,##0.00",
            "font_name": font_name,
            "font_size": 11,
            "border": 0
        })
        
        integer_format = workbook.add_format({
            "num_format": "#,##0",
            "font_name": font_name,
            "font_size": 11,
            "border": 0
        })
        
        date_format = workbook.add_format({
            "num_format": "dd-mm-yyyy",
            "font_name": font_name,
            "font_size": 11,
            "border": 0,
            "align": "left"
        })
        
        text_format = workbook.add_format({
            "font_name": font_name,
            "font_size": 11,
            "border": 0
        })
        
        def write_sheet(df, name, columns_to_show=None, is_summary=False):
            if df.empty:
                empty_df = pd.DataFrame({"Message": ["No records found"]})
                empty_df.to_excel(writer, sheet_name=name[:31], index=False, header=True, startrow=0)
                worksheet = writer.sheets[name[:31]]
                worksheet.set_column(0, 0, 30, text_format)
                return
                
            if columns_to_show:
                display_df = df[columns_to_show].copy()
            else:
                display_df = df.copy()
                
            date_columns = [col for col in display_df.columns if 'Date' in col or 'date' in col]
            for col in date_columns:
                if col in display_df.columns:
                    display_df[col] = pd.to_datetime(display_df[col], errors='coerce').dt.date
            
            display_df.to_excel(writer, sheet_name=name[:31], index=False, header=True, startrow=0)
            worksheet = writer.sheets[name[:31]]
            
            for col_num, column in enumerate(display_df.columns):
                worksheet.write(0, col_num, column, summary_header if is_summary else regular_header)
            
            for col_num, column in enumerate(display_df.columns):
                if not display_df[column].isna().all():
                    max_len = max(
                        display_df[column].astype(str).map(len).max(),
                        len(column)
                    ) + 2
                else:
                    max_len = len(column) + 2
                max_len = min(max_len, 50)
                
                if 'Date' in column or 'date' in column:
                    worksheet.set_column(col_num, col_num, max_len, date_format)
                elif any(x in column.lower() for x in ['value', 'tax', 'itc', 'amount', 'diff']):
                    worksheet.set_column(col_num, col_num, max_len, money_format)
                elif any(x in column.lower() for x in ['count']):
                    worksheet.set_column(col_num, col_num, max_len, integer_format)
                else:
                    worksheet.set_column(col_num, col_num, max_len, text_format)
        
        # Executive Summary with NO ITC and Invalid GSTIN
        summary_data = {
            "Particulars": [
                "📊 RECONCILIATION SUMMARY",
                "══════════════════════════",
                "",
                "📋 INVOICE COUNT SUMMARY:",
                "   Total Invoices in GSTR-2B",
                "   Total Invoices in Books",
                "   ✅ Fully Matched Invoices",
                "   ❌ Missing in Books",
                "   ❌ Missing in GSTR-2B",
                "   ⚠️ Value Mismatch",
                "   ⚠️ Tax Mismatch",
                "   🟡 NO ITC Invoices",
                "      ├─ ✅ Present in GSTR-2B",
                "      └─ ❌ Not in GSTR-2B",
                "",
                "💰 ITC AMOUNT SUMMARY:",
                "   ITC as per GSTR-2B",
                "   ITC as per Books",
                "   ITC Difference",
                "══════════════════════════",
                "📈 Match Percentage"
            ],
            "Count/Amount": [
                "",
                "",
                "",
                "",
                f"{summary['Total_Invoices_2B']:,.0f}",
                f"{summary['Total_Invoices_Books']:,.0f}",
                f"{summary['Total_Matched']:,.0f}",
                f"{summary['Total_Missing_Books']:,.0f}",
                f"{summary['Total_Missing_2B']:,.0f}",
                f"{len(results['value_mismatch']):,.0f}",
                f"{len(results['tax_mismatch']):,.0f}",
                f"{summary['Total_No_ITC']:,.0f}",
                f"{summary['No_ITC_with_2B']:,.0f}",
                f"{summary['No_ITC_without_2B']:,.0f}",
                "",
                "",
                f"₹{summary['Total_ITC_2B']:,.2f}",
                f"₹{summary['Total_ITC_Books']:,.2f}",
                f"₹{summary['ITC_Difference']:,.2f}",
                "",
                f"{summary['Match_Percentage']}%"
            ],
            "Status": [
                "",
                "",
                "",
                "",
                "✓",
                "✓",
                "✅",
                "❌",
                "❌",
                "⚠️",
                "⚠️",
                "🟡",
                "✅",
                "❌",
                "",
                "",
                "📌",
                "📌",
                "⚠️",
                "",
                "🎯"
            ]
        }
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name="Executive Summary", index=False, header=False, startrow=0)
        summary_sheet = writer.sheets["Executive Summary"]
        
        no_border_format = workbook.add_format({
            "font_name": font_name,
            "font_size": 11,
            "border": 0,
            "valign": "vcenter"
        })
        
        summary_sheet.set_column("A:A", 40, no_border_format)
        summary_sheet.set_column("B:B", 20, no_border_format)
        summary_sheet.set_column("C:C", 10, no_border_format)
        
        # Write all sheets
        if not results["fully_matched"].empty:
            simple_matched = results["fully_matched"][[
                "GSTIN_2B", "Trade_Name_2B", "Invoice_No_2B", "Invoice_Date_2B",
                "Taxable_Value_2B", "TOTAL_TAX_2B"
            ]].rename(columns={
                "GSTIN_2B": "GSTIN", "Trade_Name_2B": "Supplier Name",
                "Invoice_No_2B": "Invoice Number", "Invoice_Date_2B": "Invoice Date",
                "Taxable_Value_2B": "Taxable Value", "TOTAL_TAX_2B": "ITC Amount"
            })
            write_sheet(simple_matched, "✅ Fully Matched")
        
        if not results["missing_in_books"].empty:
            missing_books_simple = results["missing_in_books"][[
                "GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date",
                "Taxable_Value", "TOTAL_TAX"
            ]].rename(columns={
                "Trade_Name": "Supplier Name", "Invoice_No": "Invoice Number",
                "Invoice_Date": "Invoice Date", "Taxable_Value": "Taxable Value",
                "TOTAL_TAX": "ITC Amount"
            })
            write_sheet(missing_books_simple, "❌ Missing in Books")
        
        if not results["missing_in_2b"].empty:
            missing_2b_simple = results["missing_in_2b"][[
                "GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date",
                "Taxable_Value", "TOTAL_TAX"
            ]].rename(columns={
                "Trade_Name": "Supplier Name", "Invoice_No": "Invoice Number",
                "Invoice_Date": "Invoice Date", "Taxable_Value": "Taxable Value",
                "TOTAL_TAX": "ITC Amount"
            })
            write_sheet(missing_2b_simple, "❌ Missing in GSTR-2B")
        
        if not results["no_itc"].empty:
            no_itc_simple = results["no_itc"][[
                "GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date",
                "Taxable_Value", "Invoice_Value"
            ]].rename(columns={
                "Trade_Name": "Supplier Name", "Invoice_No": "Invoice Number",
                "Invoice_Date": "Invoice Date", "Taxable_Value": "Taxable Value",
                "Invoice_Value": "Invoice Value"
            })
            
            # Add status column
            status_list = []
            no_itc_with_2b_keys = results["no_itc_with_2b"]["KEY"].tolist() if not results["no_itc_with_2b"].empty else []
            for _, row in no_itc_simple.iterrows():
                key = row["GSTIN"] + "|" + str(row["Invoice Number"])
                if key in no_itc_with_2b_keys:
                    status_list.append("In GSTR-2B")
                else:
                    status_list.append("Not in GSTR-2B")
            no_itc_simple["Status"] = status_list
            
            write_sheet(no_itc_simple, "🟡 NO ITC Invoices")
        
        if not results["value_mismatch"].empty:
            value_mismatch_simple = results["value_mismatch"][[
                "GSTIN_2B", "Trade_Name_2B", "Invoice_No_2B", "Invoice_Date_2B",
                "Taxable_Value_2B", "Taxable_Value_Tally", "VALUE_DIFFERENCE"
            ]].rename(columns={
                "GSTIN_2B": "GSTIN", "Trade_Name_2B": "Supplier Name",
                "Invoice_No_2B": "Invoice Number", "Invoice_Date_2B": "Invoice Date",
                "Taxable_Value_2B": "Value as per GSTR-2B",
                "Taxable_Value_Tally": "Value as per Books",
                "VALUE_DIFFERENCE": "Difference"
            })
            write_sheet(value_mismatch_simple, "⚠️ Value Mismatch")
        
        if not results["tax_mismatch"].empty:
            tax_mismatch_simple = results["tax_mismatch"][[
                "GSTIN_2B", "Trade_Name_2B", "Invoice_No_2B", "Invoice_Date_2B",
                "TOTAL_TAX_2B", "TOTAL_TAX_Tally", "TAX_DIFFERENCE"
            ]].rename(columns={
                "GSTIN_2B": "GSTIN", "Trade_Name_2B": "Supplier Name",
                "Invoice_No_2B": "Invoice Number", "Invoice_Date_2B": "Invoice Date",
                "TOTAL_TAX_2B": "ITC as per GSTR-2B",
                "TOTAL_TAX_Tally": "ITC as per Books",
                "TAX_DIFFERENCE": "ITC Difference"
            })
            write_sheet(tax_mismatch_simple, "⚠️ Tax Mismatch")
        
        # Add Invalid GSTIN sheet if exists
        if "invalid_gstin" in st.session_state and not st.session_state["invalid_gstin"].empty:
            invalid_sheet = st.session_state["invalid_gstin"].copy()
            if "Invoice_Date" in invalid_sheet.columns:
                invalid_sheet["Invoice_Date"] = pd.to_datetime(invalid_sheet["Invoice_Date"]).dt.date
            write_sheet(invalid_sheet, "⚠️ Invalid GSTIN")

    output.seek(0)

    st.download_button(
        "⬇ Download Professional Excel Report",
        data=output,
        file_name=f"GST_Reconciliation_Report_{datetime.now().strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
