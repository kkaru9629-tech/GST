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

    # ================= INVALID GSTIN SECTION ================= #
    
    if not results["invalid_gstin"].empty:
        st.warning(f"⚠ {len(results['invalid_gstin'])} invoices skipped due to invalid GSTIN.")
        with st.expander("View Invalid GSTIN Invoices"):
            st.dataframe(results["invalid_gstin"], use_container_width=True)

    # ================= EXECUTIVE SUMMARY ================= #

    st.markdown("## 📊 Executive Summary")

    col1, col2, col3 = st.columns(3)
    col1.metric("📘 Total Invoices - Books", summary["Total_Invoices_Books"])
    col2.metric("📗 Total Invoices - GSTR-2B", summary["Total_Invoices_2B"])
    col3.metric("✅ Fully Matched Invoices", summary["Total_Matched"])

    st.divider()

    col4, col5, col6 = st.columns(3)
    col4.metric("💰 ITC as per Books", f"₹{summary['Total_ITC_Books']:,.2f}")
    col5.metric("💰 ITC as per GSTR-2B", f"₹{summary['Total_ITC_2B']:,.2f}")
    col6.metric("📊 ITC Difference", f"₹{summary['ITC_Difference']:,.2f}")

    st.divider()

    col7, col8 = st.columns(2)
    col7.metric("📕 Missing in Books", summary["Total_Missing_Books"])
    col8.metric("📗 Missing in GSTR-2B", summary["Total_Missing_2B"])

    st.divider()

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

    # ===== TAB 1: Fully Matched =====
    with tab1:
        if not results["fully_matched"].empty:
            display = results["fully_matched"][
                ["GSTIN_2B", "Trade_Name_2B", "Invoice_No_2B", "Invoice_Date_2B", 
                 "Taxable_Value_2B", "TOTAL_TAX_2B"]
            ].copy()
            
            display["Invoice_Date_2B"] = pd.to_datetime(display["Invoice_Date_2B"]).dt.date
            
            display.columns = [
                "GSTIN", "Supplier Name", "Invoice Number", 
                "Invoice Date", "Taxable Value", "ITC Amount"
            ]
            
            st.dataframe(display, use_container_width=True)
            
            csv = display.to_csv(index=False).encode("utf-8")
            st.download_button(
                "📥 Download Fully Matched",
                csv,
                "Fully_Matched.csv",
                "text/csv"
            )
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
            
            display.columns = [
                "GSTIN", "Supplier Name", "Invoice Number", 
                "Invoice Date", "Taxable Value", "ITC Amount"
            ]
            
            st.dataframe(display, use_container_width=True)
            
            csv = display.to_csv(index=False).encode("utf-8")
            st.download_button(
                "📥 Download Missing in Books",
                csv,
                "Missing_in_Books.csv",
                "text/csv"
            )
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
            
            display.columns = [
                "GSTIN", "Supplier Name", "Invoice Number", 
                "Invoice Date", "Taxable Value", "ITC Amount"
            ]
            
            st.dataframe(display, use_container_width=True)
            
            csv = display.to_csv(index=False).encode("utf-8")
            st.download_button(
                "📥 Download Missing in GSTR-2B",
                csv,
                "Missing_in_GSTR2B.csv",
                "text/csv"
            )
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
            
            display.columns = [
                "GSTIN", "Supplier Name", "Invoice Number", "Invoice Date",
                "Value (2B)", "Value (Books)", "Difference"
            ]
            
            st.dataframe(display, use_container_width=True)
            
            csv = display.to_csv(index=False).encode("utf-8")
            st.download_button(
                "📥 Download Value Mismatches",
                csv,
                "Value_Mismatch.csv",
                "text/csv"
            )
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
            
            display.columns = [
                "GSTIN", "Supplier Name", "Invoice Number", "Invoice Date",
                "ITC (2B)", "ITC (Books)", "Difference"
            ]
            
            st.dataframe(display, use_container_width=True)
            
            csv = display.to_csv(index=False).encode("utf-8")
            st.download_button(
                "📥 Download Tax Mismatches",
                csv,
                "Tax_Mismatch.csv",
                "text/csv"
            )
        else:
            st.success("No tax mismatches found.")

    # ===== TAB 6: NO ITC =====
    with tab6:
        if not results["no_itc"].empty:
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Invoices", len(results["no_itc"]))
            col2.metric("Total Taxable Value", f"₹{results['no_itc']['Taxable_Value'].sum():,.2f}")
            col3.metric("Total Invoice Value", f"₹{results['no_itc']['Invoice_Value'].sum():,.2f}")
            
            display = results["no_itc"][
                ["GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date", 
                 "Taxable_Value", "Invoice_Value"]
            ].copy()
            
            display["Invoice_Date"] = pd.to_datetime(display["Invoice_Date"]).dt.date
            
            display.columns = [
                "GSTIN", "Supplier Name", "Invoice Number", 
                "Invoice Date", "Taxable Value", "Invoice Value"
            ]
            
            st.dataframe(display, use_container_width=True)
            
            csv = display.to_csv(index=False).encode("utf-8")
            st.download_button(
                "📥 Download Zero ITC Invoices",
                csv,
                "Zero_ITC.csv",
                "text/csv"
            )
        else:
            st.success("No zero ITC invoices found.")

    # ===== TAB 7: Supplier Analysis =====
    with tab7:
        st.markdown("### 📊 Supplier-wise ITC Summary")
        
        # Combine data for analysis
        dfs_to_concat = []
        if not results["fully_matched"].empty:
            dfs_to_concat.append(results["fully_matched"])
        if not results["value_mismatch"].empty:
            dfs_to_concat.append(results["value_mismatch"])
        if not results["tax_mismatch"].empty:
            dfs_to_concat.append(results["tax_mismatch"])
        
        if dfs_to_concat:
            combined = pd.concat(dfs_to_concat, ignore_index=True)
            
            # Create supplier summary
            supplier_summary = combined.groupby("Trade_Name_2B").agg({
                "Invoice_No_2B": "count",
                "TOTAL_TAX_2B": "sum",
                "TOTAL_TAX_Tally": "sum",
                "VALUE_DIFFERENCE": "sum",
                "TAX_DIFFERENCE": "sum"
            }).reset_index()
            
            supplier_summary.columns = [
                "Supplier Name", "Total Invoices", "ITC (2B)", 
                "ITC (Books)", "Value Diff", "Tax Diff"
            ]
            
            supplier_summary = supplier_summary.sort_values("ITC (2B)", ascending=False)
            
            st.dataframe(supplier_summary, use_container_width=True)
            
            # Top 5 suppliers by ITC
            st.markdown("### 🏆 Top 5 Suppliers by ITC")
            top5 = supplier_summary.head(5)[["Supplier Name", "ITC (2B)"]]
            st.dataframe(top5, use_container_width=True)
        else:
            st.info("No supplier data available for analysis.")

    # ================= DOWNLOAD EXCEL REPORT ================= #
    
    st.divider()
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        
        # Formats
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D9E1F2',
            'border': 1,
            'align': 'center'
        })
        
        money_format = workbook.add_format({'num_format': '₹#,##0.00'})
        
        # Summary Sheet
        summary_df = pd.DataFrame([{
            "Particular": "Total Invoices - Books",
            "Value": summary["Total_Invoices_Books"]
        }, {
            "Particular": "Total Invoices - GSTR-2B",
            "Value": summary["Total_Invoices_2B"]
        }, {
            "Particular": "Fully Matched",
            "Value": summary["Total_Matched"]
        }, {
            "Particular": "Missing in Books",
            "Value": summary["Total_Missing_Books"]
        }, {
            "Particular": "Missing in GSTR-2B",
            "Value": summary["Total_Missing_2B"]
        }, {
            "Particular": "Match %",
            "Value": summary["Match_Percentage"]
        }, {
            "Particular": "ITC - Books",
            "Value": summary["Total_ITC_Books"]
        }, {
            "Particular": "ITC - GSTR-2B",
            "Value": summary["Total_ITC_2B"]
        }, {
            "Particular": "ITC Difference",
            "Value": summary["ITC_Difference"]
        }])
        
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        
        # Format summary sheet
        worksheet = writer.sheets["Summary"]
        worksheet.set_column("A:A", 30)
        worksheet.set_column("B:B", 20)
        
        # Write detailed sheets
        sheet_mapping = [
            ("fully_matched", "Fully Matched"),
            ("missing_in_books", "Missing in Books"),
            ("missing_in_2b", "Missing in 2B"),
            ("value_mismatch", "Value Mismatch"),
            ("tax_mismatch", "Tax Mismatch"),
            ("no_itc", "NO ITC"),
            ("invalid_gstin", "Invalid GSTIN")
        ]
        
        for key, sheet_name in sheet_mapping:
            if key in results and not results[key].empty:
                df = results[key].copy()
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                worksheet = writer.sheets[sheet_name]
                worksheet.freeze_panes(1, 0)
                
                # Format header
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                # Adjust column widths
                for i, col in enumerate(df.columns):
                    max_len = max(
                        df[col].astype(str).map(len).max() if not df[col].empty else 0,
                        len(col)
                    ) + 2
                    worksheet.set_column(i, i, min(max_len, 50))
    
    output.seek(0)
    
    st.download_button(
        "⬇ Download Complete Excel Report",
        data=output,
        file_name=f"GST_Reconciliation_{datetime.now().strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
