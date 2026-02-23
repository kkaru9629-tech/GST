import streamlit as st
import pandas as pd
import io
from reconciliation_engine import parse_tally, parse_gstr2b, reconcile
from datetime import datetime

st.set_page_config(page_title="GST Reconciliation", layout="wide")

st.title("📊 GST Reconciliation System")
st.caption("ITC Books vs GSTR-2B Comparison")

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

    # ================= SIMPLE SUMMARY ================= #
    
    st.markdown("## 📌 Summary")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Total ITC - Books", f"₹{summary['Total_ITC_Books']:,.2f}")
        st.metric("Missing in 2B (Books - 2B)", f"₹{summary['Total_Missing_2B_value']:,.2f}")
    
    with col2:
        st.metric("Total ITC - GSTR-2B", f"₹{summary['Total_ITC_2B']:,.2f}")
        st.metric("Missing in Books (2B - Books)", f"₹{summary['Total_Missing_Books_value']:,.2f}")
    
    with col3:
        st.metric("Net Difference", f"₹{summary['ITC_Difference']:,.2f}")
    
    st.divider()

    # ================= TABS ================= #

    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "Books Data",
        "GSTR-2B Data",
        "Missing in 2B",
        "Missing in Books",
        "NO ITC Invoices",
        "Invalid GSTIN",
        "Supplier Analysis"
    ])

    # ===== TAB 1: Books Data (NO DATES) =====
    with tab1:
        books_data_list = []
        
        if not results["missing_in_2b"].empty:
            books_data_list.append(results["missing_in_2b"])
        
        if "fully_matched" in results and not results["fully_matched"].empty:
            matched_books = results["fully_matched"][["GSTIN_Tally", "Trade_Name_Tally", "Invoice_No_Tally", 
                                                      "Taxable_Value_Tally", "CGST_Tally", "SGST_Tally", 
                                                      "IGST_Tally", "TOTAL_TAX_Tally"]].copy()
            matched_books.columns = ["GSTIN", "Trade_Name", "Invoice_No", 
                                     "Taxable_Value", "CGST", "SGST", "IGST", "TOTAL_TAX"]
            books_data_list.append(matched_books)
        
        if books_data_list:
            books_data = pd.concat(books_data_list, ignore_index=True)
            
            display = books_data[["Trade_Name", "GSTIN", "Invoice_No", 
                                   "Taxable_Value", "CGST", "SGST", "IGST", "TOTAL_TAX"]].copy()
            
            for col in ["Taxable_Value", "CGST", "SGST", "IGST", "TOTAL_TAX"]:
                display[col] = display[col].apply(lambda x: f"₹{x:,.2f}")
            
            display.columns = ["Supplier Name", "GSTIN", "Invoice No", 
                                "Taxable Value", "CGST", "SGST", "IGST", "Total ITC"]
            
            st.dataframe(display, use_container_width=True, hide_index=True)
        else:
            st.info("No books data available.")

    # ===== TAB 2: GSTR-2B Data (NO DATES) =====
    with tab2:
        gstr_data_list = []
        
        if not results["missing_in_books"].empty:
            gstr_data_list.append(results["missing_in_books"])
        
        if "fully_matched" in results and not results["fully_matched"].empty:
            matched_gstr = results["fully_matched"][["GSTIN_2B", "Trade_Name_2B", "Invoice_No_2B",
                                                      "Taxable_Value_2B", "CGST_2B", "SGST_2B", 
                                                      "IGST_2B", "TOTAL_TAX_2B"]].copy()
            matched_gstr.columns = ["GSTIN", "Trade_Name", "Invoice_No",
                                     "Taxable_Value", "CGST", "SGST", "IGST", "TOTAL_TAX"]
            gstr_data_list.append(matched_gstr)
        
        if gstr_data_list:
            gstr_data = pd.concat(gstr_data_list, ignore_index=True)
            
            display = gstr_data[["Trade_Name", "GSTIN", "Invoice_No",
                                  "Taxable_Value", "CGST", "SGST", "IGST", "TOTAL_TAX"]].copy()
            
            for col in ["Taxable_Value", "CGST", "SGST", "IGST", "TOTAL_TAX"]:
                display[col] = display[col].apply(lambda x: f"₹{x:,.2f}")
            
            display.columns = ["Supplier Name", "GSTIN", "Invoice No",
                                "Taxable Value", "CGST", "SGST", "IGST", "Total ITC"]
            
            st.dataframe(display, use_container_width=True, hide_index=True)
        else:
            st.info("No GSTR-2B data available.")

    # ===== TAB 3: Missing in 2B =====
    with tab3:
        if not results["missing_in_2b"].empty:
            display = results["missing_in_2b"][
                ["Trade_Name", "GSTIN", "Invoice_No", "Taxable_Value", "TOTAL_TAX"]
            ].copy()
            
            display["Taxable_Value"] = display["Taxable_Value"].apply(lambda x: f"₹{x:,.2f}")
            display["TOTAL_TAX"] = display["TOTAL_TAX"].apply(lambda x: f"₹{x:,.2f}")
            
            display.columns = ["Supplier Name", "GSTIN", "Invoice No", "Taxable Value", "ITC Amount"]
            
            st.dataframe(display, use_container_width=True, hide_index=True)
        else:
            st.success("No invoices missing in 2B.")

    # ===== TAB 4: Missing in Books =====
    with tab4:
        if not results["missing_in_books"].empty:
            display = results["missing_in_books"][
                ["Trade_Name", "GSTIN", "Invoice_No", "Taxable_Value", "TOTAL_TAX"]
            ].copy()
            
            display["Taxable_Value"] = display["Taxable_Value"].apply(lambda x: f"₹{x:,.2f}")
            display["TOTAL_TAX"] = display["TOTAL_TAX"].apply(lambda x: f"₹{x:,.2f}")
            
            display.columns = ["Supplier Name", "GSTIN", "Invoice No", "Taxable Value", "ITC Amount"]
            
            st.dataframe(display, use_container_width=True, hide_index=True)
        else:
            st.success("No invoices missing in Books.")

    # ===== TAB 5: NO ITC =====
    with tab5:
        if not results["no_itc"].empty:
            display = results["no_itc"][
                ["Trade_Name", "GSTIN", "Invoice_No", "Taxable_Value", "Invoice_Value"]
            ].copy()
            
            display["Taxable_Value"] = display["Taxable_Value"].apply(lambda x: f"₹{x:,.2f}")
            display["Invoice_Value"] = display["Invoice_Value"].apply(lambda x: f"₹{x:,.2f}")
            
            display.columns = ["Supplier Name", "GSTIN", "Invoice No", "Taxable Value", "Invoice Value"]
            
            st.dataframe(display, use_container_width=True, hide_index=True)
        else:
            st.success("No zero ITC invoices found.")

    # ===== TAB 6: Invalid GSTIN =====
    with tab6:
        if not results["invalid_gstin"].empty:
            display = results["invalid_gstin"][
                ["Trade_Name", "GSTIN", "Invoice_No", "Taxable_Value", "TOTAL_TAX"]
            ].copy()
            
            display["Taxable_Value"] = display["Taxable_Value"].apply(lambda x: f"₹{x:,.2f}")
            display["TOTAL_TAX"] = display["TOTAL_TAX"].apply(lambda x: f"₹{x:,.2f}")
            
            display.columns = ["Supplier Name", "GSTIN", "Invoice No", "Taxable Value", "ITC Amount"]
            
            st.dataframe(display, use_container_width=True, hide_index=True)
        else:
            st.success("No invalid GSTIN invoices.")

    # ===== TAB 7: Supplier Analysis =====
    with tab7:
        st.markdown("### 📊 Supplier-wise Analysis")
        
        # Prepare data
        supplier_data = []
        
        # Get all unique suppliers from missing_in_2b and missing_in_books
        all_suppliers = set()
        
        if not results["missing_in_2b"].empty:
            for _, row in results["missing_in_2b"].iterrows():
                all_suppliers.add((row["GSTIN"], row["Trade_Name"]))
        
        if not results["missing_in_books"].empty:
            for _, row in results["missing_in_books"].iterrows():
                all_suppliers.add((row["GSTIN"], row["Trade_Name"]))
        
        # Build supplier analysis
        for gstin, name in all_suppliers:
            books_itc = 0
            gstr_itc = 0
            missing_2b_count = 0
            missing_2b_itc = 0
            missing_books_count = 0
            missing_books_itc = 0
            total_invoices = 0
            
            # Missing in 2B (in Books but not in 2B)
            if not results["missing_in_2b"].empty:
                mask = results["missing_in_2b"]["GSTIN"] == gstin
                missing_2b_count = mask.sum()
                if missing_2b_count > 0:
                    missing_2b_itc = results["missing_in_2b"].loc[mask, "TOTAL_TAX"].sum()
                    books_itc += missing_2b_itc
                    total_invoices += missing_2b_count
            
            # Missing in Books (in 2B but not in Books)
            if not results["missing_in_books"].empty:
                mask = results["missing_in_books"]["GSTIN"] == gstin
                missing_books_count = mask.sum()
                if missing_books_count > 0:
                    missing_books_itc = results["missing_in_books"].loc[mask, "TOTAL_TAX"].sum()
                    gstr_itc += missing_books_itc
                    total_invoices += missing_books_count
            
            supplier_data.append({
                "Supplier Name": name,
                "GSTIN": gstin,
                "Total Invoices": total_invoices,
                "ITC - Books": books_itc,
                "ITC - GSTR-2B": gstr_itc,
                "Difference": gstr_itc - books_itc,
                "Missing in 2B": missing_2b_count,
                "Missing in 2B (ITC)": missing_2b_itc,
                "Missing in Books": missing_books_count,
                "Missing in Books (ITC)": missing_books_itc
            })
        
        if supplier_data:
            supplier_df = pd.DataFrame(supplier_data)
            
            # Format currency columns for display
            display_df = supplier_df.copy()
            for col in ["ITC - Books", "ITC - GSTR-2B", "Difference", "Missing in 2B (ITC)", "Missing in Books (ITC)"]:
                display_df[col] = display_df[col].apply(lambda x: f"₹{x:,.2f}")
            
            # Sort by absolute difference (highest first)
            display_df = display_df.sort_values("Difference", key=lambda x: abs(float(str(x).replace('₹', '').replace(',', ''))), ascending=False)
            
            st.dataframe(display_df, use_container_width=True, hide_index=True)
            
            # High Risk Suppliers (Top 5 by Difference)
            st.markdown("### ⚠️ High Risk Suppliers")
            high_risk = display_df.head(5)[["Supplier Name", "GSTIN", "Difference", "Missing in 2B", "Missing in Books"]]
            st.dataframe(high_risk, use_container_width=True, hide_index=True)
            
            # High ITC Suppliers (Top 5 by ITC)
            st.markdown("### 💰 High ITC Suppliers")
            # Create a copy with numeric values for sorting
            numeric_df = supplier_df.copy()
            numeric_df["ITC - Books"] = pd.to_numeric(numeric_df["ITC - Books"])
            high_itc_numeric = numeric_df.nlargest(5, "ITC - Books")[["Supplier Name", "GSTIN", "ITC - Books"]]
            high_itc_numeric["ITC - Books"] = high_itc_numeric["ITC - Books"].apply(lambda x: f"₹{x:,.2f}")
            st.dataframe(high_itc_numeric, use_container_width=True, hide_index=True)
        else:
            st.info("No supplier data available for analysis.")

    # ================= DOWNLOAD EXCEL REPORT ================= #

    st.divider()
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        
        # Set default font to Aptos Narrow
        workbook.formats[0].set_font_name('Aptos Narrow')
        
        # Formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'font_name': 'Aptos Narrow',
            'align': 'left'
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'font_name': 'Aptos Narrow',
            'bg_color': '#D9E1F2',
            'border': 1,
            'align': 'center'
        })
        
        text_format = workbook.add_format({'font_name': 'Aptos Narrow'})
        money_format = workbook.add_format({'font_name': 'Aptos Narrow', 'num_format': '₹#,##0.00'})
        
        # ===== 1. SUMMARY SHEET =====
        summary_sheet = workbook.add_worksheet("Summary")
        
        summary_sheet.set_column('A:A', 35)
        summary_sheet.set_column('B:B', 25)
        summary_sheet.set_column('C:C', 15)
        
        # Title
        summary_sheet.write('A1', 'GST RECONCILIATION SUMMARY', title_format)
        
        # Summary Table
        headers = ['Particulars', 'Amount', 'Status']
        summary_sheet.write_row('A3', headers, header_format)
        
        summary_rows = [
            ['Total ITC as per Books', summary['Total_ITC_Books'], ''],
            ['Total ITC as per GSTR-2B', summary['Total_ITC_2B'], ''],
            ['Missing in 2B (Books - 2B)', summary['Total_Missing_2B_value'], '❌' if summary['Total_Missing_2B_value'] > 0 else '✓'],
            ['Missing in Books (2B - Books)', summary['Total_Missing_Books_value'], '❌' if summary['Total_Missing_Books_value'] > 0 else '✓'],
            ['Net Difference', summary['ITC_Difference'], '⚠️' if abs(summary['ITC_Difference']) > 1000 else '✓']
        ]
        
        row = 3
        for i, (label, value, status) in enumerate(summary_rows):
            summary_sheet.write(row, 0, label, text_format)
            summary_sheet.write(row, 1, value, money_format)
            summary_sheet.write(row, 2, status, text_format)
            row += 1
        
        # Conditional Formatting
        summary_sheet.conditional_format(f'B4:B{row-1}', {
            'type': 'cell',
            'criteria': '>',
            'value': 0,
            'format': workbook.add_format({'font_name': 'Aptos Narrow', 'num_format': '₹#,##0.00', 'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        })
        
        summary_sheet.conditional_format(f'B4:B{row-1}', {
            'type': 'cell',
            'criteria': '<',
            'value': 0,
            'format': workbook.add_format({'font_name': 'Aptos Narrow', 'num_format': '₹#,##0.00', 'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        })
        
        summary_sheet.add_table(f'A3:C{row-1}', {
            'columns': [{'header': h} for h in headers],
            'style': 'Table Style Light 1'
        })
        
        # ===== 2. BOOKS DATA SHEET =====
        books_data_list = []
        
        if not results["missing_in_2b"].empty:
            books_data_list.append(results["missing_in_2b"])
        
        if "fully_matched" in results and not results["fully_matched"].empty:
            matched_books = results["fully_matched"][["GSTIN_Tally", "Trade_Name_Tally", "Invoice_No_Tally", 
                                                      "Taxable_Value_Tally", "CGST_Tally", "SGST_Tally", 
                                                      "IGST_Tally", "TOTAL_TAX_Tally"]].copy()
            matched_books.columns = ["GSTIN", "Trade_Name", "Invoice_No", 
                                     "Taxable_Value", "CGST", "SGST", "IGST", "TOTAL_TAX"]
            books_data_list.append(matched_books)
        
        if books_data_list:
            books_data = pd.concat(books_data_list, ignore_index=True)
            books_data = books_data[["Trade_Name", "GSTIN", "Invoice_No", 
                                      "Taxable_Value", "CGST", "SGST", "IGST", "TOTAL_TAX"]]
            books_data.columns = ["Supplier Name", "GSTIN", "Invoice No", 
                                   "Taxable Value", "CGST", "SGST", "IGST", "Total ITC"]
            
            books_data.to_excel(writer, sheet_name="Books Data", index=False, startrow=1)
            
            # Format
            worksheet = writer.sheets["Books Data"]
            worksheet.write(0, 0, "Books Data", title_format)
            
            for col_num, col in enumerate(books_data.columns):
                worksheet.write(1, col_num, col, header_format)
                if "Value" in col or "CGST" in col or "SGST" in col or "IGST" in col or "ITC" in col:
                    worksheet.set_column(col_num, col_num, 18, money_format)
                else:
                    worksheet.set_column(col_num, col_num, 22)
            
            worksheet.freeze_panes(2, 0)
        
        # ===== 3. GSTR-2B DATA SHEET =====
        gstr_data_list = []
        
        if not results["missing_in_books"].empty:
            gstr_data_list.append(results["missing_in_books"])
        
        if "fully_matched" in results and not results["fully_matched"].empty:
            matched_gstr = results["fully_matched"][["GSTIN_2B", "Trade_Name_2B", "Invoice_No_2B",
                                                      "Taxable_Value_2B", "CGST_2B", "SGST_2B", 
                                                      "IGST_2B", "TOTAL_TAX_2B"]].copy()
            matched_gstr.columns = ["GSTIN", "Trade_Name", "Invoice_No",
                                     "Taxable_Value", "CGST", "SGST", "IGST", "TOTAL_TAX"]
            gstr_data_list.append(matched_gstr)
        
        if gstr_data_list:
            gstr_data = pd.concat(gstr_data_list, ignore_index=True)
            gstr_data = gstr_data[["Trade_Name", "GSTIN", "Invoice_No",
                                    "Taxable_Value", "CGST", "SGST", "IGST", "TOTAL_TAX"]]
            gstr_data.columns = ["Supplier Name", "GSTIN", "Invoice No",
                                  "Taxable Value", "CGST", "SGST", "IGST", "Total ITC"]
            
            gstr_data.to_excel(writer, sheet_name="GSTR-2B Data", index=False, startrow=1)
            
            # Format
            worksheet = writer.sheets["GSTR-2B Data"]
            worksheet.write(0, 0, "GSTR-2B Data", title_format)
            
            for col_num, col in enumerate(gstr_data.columns):
                worksheet.write(1, col_num, col, header_format)
                if "Value" in col or "CGST" in col or "SGST" in col or "IGST" in col or "ITC" in col:
                    worksheet.set_column(col_num, col_num, 18, money_format)
                else:
                    worksheet.set_column(col_num, col_num, 22)
            
            worksheet.freeze_panes(2, 0)
        
        # ===== 4. MISSING IN 2B SHEET =====
        if not results["missing_in_2b"].empty:
            missing_2b_sheet = workbook.add_worksheet("Missing in 2B")
            
            missing_2b = results["missing_in_2b"][["Trade_Name", "GSTIN", "Invoice_No", "Taxable_Value", "TOTAL_TAX"]].copy()
            missing_2b = missing_2b.sort_values("Trade_Name")
            
            missing_2b.columns = ["Supplier Name", "GSTIN", "Invoice No", "Taxable Value", "ITC Amount"]
            
            missing_2b.to_excel(writer, sheet_name="Missing in 2B", index=False, startrow=1)
            
            # Format
            worksheet = writer.sheets["Missing in 2B"]
            worksheet.write(0, 0, "Missing in 2B (In Books but not in GSTR-2B)", title_format)
            
            for col_num, col in enumerate(missing_2b.columns):
                worksheet.write(1, col_num, col, header_format)
                if "Value" in col or "ITC" in col:
                    worksheet.set_column(col_num, col_num, 18, money_format)
                else:
                    worksheet.set_column(col_num, col_num, 22)
            
            worksheet.freeze_panes(2, 0)
            
            # Add totals
            last_row = len(missing_2b) + 2
            worksheet.write(last_row, 3, "Total ITC Impact:", text_format)
            worksheet.write(last_row, 4, missing_2b["ITC Amount"].sum(), money_format)
        
        # ===== 5. MISSING IN BOOKS SHEET =====
        if not results["missing_in_books"].empty:
            missing_books_sheet = workbook.add_worksheet("Missing in Books")
            
            missing_books = results["missing_in_books"][["Trade_Name", "GSTIN", "Invoice_No", "Taxable_Value", "TOTAL_TAX"]].copy()
            missing_books = missing_books.sort_values("Trade_Name")
            
            missing_books.columns = ["Supplier Name", "GSTIN", "Invoice No", "Taxable Value", "ITC Amount"]
            
            missing_books.to_excel(writer, sheet_name="Missing in Books", index=False, startrow=1)
            
            # Format
            worksheet = writer.sheets["Missing in Books"]
            worksheet.write(0, 0, "Missing in Books (In GSTR-2B but not in Books)", title_format)
            
            for col_num, col in enumerate(missing_books.columns):
                worksheet.write(1, col_num, col, header_format)
                if "Value" in col or "ITC" in col:
                    worksheet.set_column(col_num, col_num, 18, money_format)
                else:
                    worksheet.set_column(col_num, col_num, 22)
            
            worksheet.freeze_panes(2, 0)
            
            # Add totals
            last_row = len(missing_books) + 2
            worksheet.write(last_row, 3, "Total ITC Impact:", text_format)
            worksheet.write(last_row, 4, missing_books["ITC Amount"].sum(), money_format)
        
        # ===== 6. SUPPLIER ANALYSIS SHEET =====
        if supplier_data:
            analysis_sheet = workbook.add_worksheet("Supplier Analysis")
            
            supplier_df = pd.DataFrame(supplier_data)
            
            # Pivot Table 1: Supplier-wise ITC
            pivot1 = supplier_df[["Supplier Name", "GSTIN", "ITC - Books", "ITC - GSTR-2B", "Difference", "Missing in 2B (ITC)"]].copy()
            pivot1 = pivot1.sort_values("Difference", ascending=False)
            
            pivot1.to_excel(writer, sheet_name="Supplier Analysis", index=False, startrow=1)
            
            # Format
            worksheet = writer.sheets["Supplier Analysis"]
            worksheet.write(0, 0, "Supplier-wise ITC Analysis", title_format)
            
            for col_num, col in enumerate(pivot1.columns):
                worksheet.write(1, col_num, col, header_format)
                if "ITC" in col or "Difference" in col:
                    worksheet.set_column(col_num, col_num, 18, money_format)
                else:
                    worksheet.set_column(col_num, col_num, 22)
            
            # Pivot Table 2: Count Analysis (starting at row len(pivot1)+5)
            start_row = len(pivot1) + 5
            
            worksheet.write(start_row, 0, "Supplier-wise Invoice Count Analysis", title_format)
            
            pivot2 = supplier_df[["Supplier Name", "GSTIN", "Total Invoices", "Missing in 2B", "Missing in Books"]].copy()
            pivot2 = pivot2.sort_values("Total Invoices", ascending=False)
            
            for col_num, col in enumerate(pivot2.columns):
                worksheet.write(start_row + 1, col_num, col, header_format)
            
            for i, (_, row) in enumerate(pivot2.iterrows()):
                for col_num, col in enumerate(pivot2.columns):
                    worksheet.write(start_row + 2 + i, col_num, row[col], text_format)
            
            # Conditional Formatting
            worksheet.conditional_format(f'D{start_row + 2}:E{start_row + 2 + len(pivot2)}', {
                'type': 'cell',
                'criteria': '>',
                'value': 0,
                'format': workbook.add_format({'font_name': 'Aptos Narrow', 'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
            })
            
            worksheet.freeze_panes(2, 0)
        
        # ===== 7. NO ITC SHEET =====
        if not results["no_itc"].empty:
            no_itc_sheet = workbook.add_worksheet("NO ITC")
            
            no_itc = results["no_itc"][["Trade_Name", "GSTIN", "Invoice_No", "Taxable_Value", "Invoice_Value"]].copy()
            no_itc.columns = ["Supplier Name", "GSTIN", "Invoice No", "Taxable Value", "Invoice Value"]
            
            no_itc.to_excel(writer, sheet_name="NO ITC", index=False, startrow=1)
            
            # Format
            worksheet = writer.sheets["NO ITC"]
            worksheet.write(0, 0, "Zero ITC Invoices", title_format)
            
            for col_num, col in enumerate(no_itc.columns):
                worksheet.write(1, col_num, col, header_format)
                if "Value" in col:
                    worksheet.set_column(col_num, col_num, 18, money_format)
                else:
                    worksheet.set_column(col_num, col_num, 22)
            
            worksheet.freeze_panes(2, 0)
        
        # ===== 8. INVALID GSTIN SHEET =====
        if not results["invalid_gstin"].empty:
            invalid_sheet = workbook.add_worksheet("Invalid GSTIN")
            
            invalid_df = results["invalid_gstin"][["Trade_Name", "GSTIN", "Invoice_No", "Taxable_Value", "TOTAL_TAX"]].copy()
            invalid_df.columns = ["Supplier Name", "GSTIN", "Invoice No", "Taxable Value", "ITC Amount"]
            
            invalid_df.to_excel(writer, sheet_name="Invalid GSTIN", index=False, startrow=1)
            
            # Format
            worksheet = writer.sheets["Invalid GSTIN"]
            worksheet.write(0, 0, "Invalid GSTIN Invoices", title_format)
            
            for col_num, col in enumerate(invalid_df.columns):
                worksheet.write(1, col_num, col, header_format)
                if "Value" in col or "ITC" in col:
                    worksheet.set_column(col_num, col_num, 18, money_format)
                else:
                    worksheet.set_column(col_num, col_num, 22)
            
            worksheet.freeze_panes(2, 0)
    
    output.seek(0)
    
    st.download_button(
        "⬇ Download Excel Report",
        data=output,
        file_name=f"GST_Reconciliation_{datetime.now().strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
