import streamlit as st
import pandas as pd
import io
from reconciliation_engine import parse_tally, parse_gstr2b, reconcile
from datetime import datetime

st.set_page_config(page_title="GST Reconciliation", layout="wide", initial_sidebar_state="collapsed")

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1E3A8A;
        font-weight: 700;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1rem;
        color: #6B7280;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #F9FAFB;
        padding: 1.5rem;
        border-radius: 0.75rem;
        border-left: 4px solid #3B82F6;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .critical {
        background-color: #FEF2F2;
        padding: 0.5rem;
        border-radius: 0.5rem;
        color: #DC2626;
        font-weight: 600;
    }
    .watch {
        background-color: #FFFBEB;
        padding: 0.5rem;
        border-radius: 0.5rem;
        color: #D97706;
        font-weight: 600;
    }
    .good {
        background-color: #F0FDF4;
        padding: 0.5rem;
        border-radius: 0.5rem;
        color: #059669;
        font-weight: 600;
    }
    .insight-box {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.2rem;
        border-radius: 0.75rem;
        color: white;
        margin-bottom: 1rem;
    }
    .stat-box {
        background-color: white;
        padding: 1rem;
        border-radius: 0.5rem;
        border: 1px solid #E5E7EB;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
    }
    .top-five {
        background-color: #F3F4F6;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #8B5CF6;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-header">📊 GST Reconciliation System</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">ITC Books vs GSTR-2B Comparison | Auto-fit Columns | Supplier Analysis</p>', unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    tally_file = st.file_uploader("📘 Upload Tally Purchase Register", type=["xlsx", "xls", "csv"])

with col2:
    gstr2b_file = st.file_uploader("📗 Upload Structured GSTR-2B", type=["xlsx", "xls", "csv"])

if st.button("🚀 Run Reconciliation", use_container_width=True, type="primary"):
    if tally_file is None or gstr2b_file is None:
        st.error("Please upload both files.")
    else:
        try:
            with st.spinner("Processing... Please wait"):
                tally_raw = pd.read_csv(tally_file) if tally_file.name.endswith("csv") else pd.read_excel(tally_file)
                gstr2b_raw = pd.read_csv(gstr2b_file) if gstr2b_file.name.endswith("csv") else pd.read_excel(gstr2b_file)

                tally_clean, no_itc_df, invalid_gstin_df = parse_tally(tally_raw)
                gstr2b_clean = parse_gstr2b(gstr2b_raw)

                results = reconcile(gstr2b_clean, tally_clean)

                results["no_itc"] = no_itc_df
                results["invalid_gstin"] = invalid_gstin_df

                st.session_state["results"] = results
                st.success("✅ Reconciliation Completed Successfully!")

        except Exception as e:
            st.error(f"Error: {str(e)}")

# ================= DISPLAY ================= #

if "results" in st.session_state:
    results = st.session_state["results"]
    summary = results["summary"]

    # ================= ENHANCED SUMMARY DASHBOARD ================= #
    
    st.markdown("## 📊 Reconciliation Dashboard")
    
    # Top Row - Key Metrics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown('<div class="stat-box">', unsafe_allow_html=True)
        st.metric("📘 ITC - Books", f"₹{summary['Total_ITC_Books']:,.2f}")
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="stat-box">', unsafe_allow_html=True)
        st.metric("📗 ITC - GSTR-2B", f"₹{summary['Total_ITC_2B']:,.2f}")
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col3:
        st.markdown('<div class="stat-box">', unsafe_allow_html=True)
        diff_color = "inverse" if summary['ITC_Difference'] < 0 else "normal"
        st.metric("📊 Net Difference", f"₹{summary['ITC_Difference']:,.2f}", delta_color=diff_color)
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col4:
        st.markdown('<div class="stat-box">', unsafe_allow_html=True)
        match_percent = summary['Match_Percentage']
        if match_percent >= 90:
            st.metric("✅ Match %", f"{match_percent}%")
        elif match_percent >= 75:
            st.metric("⚠️ Match %", f"{match_percent}%")
        else:
            st.metric("🔴 Match %", f"{match_percent}%")
        st.markdown('</div>', unsafe_allow_html=True)
    
    st.divider()
    
    # ================= PREPARE SUPPLIER DATA FOR INSIGHTS ================= #
    
    # Build supplier dictionary grouped by GSTIN
    supplier_dict = {}
    
    # Process Missing in 2B (Books data)
    if not results["missing_in_2b"].empty:
        for _, row in results["missing_in_2b"].iterrows():
            gstin = row["GSTIN"]
            name = row["Trade_Name"]
            
            if gstin not in supplier_dict:
                supplier_dict[gstin] = {
                    "names": set(),
                    "books_itc": 0,
                    "gstr_itc": 0,
                    "missing_2b_count": 0,
                    "missing_2b_itc": 0,
                    "missing_books_count": 0,
                    "missing_books_itc": 0,
                    "total_invoices": 0
                }
            
            supplier_dict[gstin]["names"].add(name)
            supplier_dict[gstin]["books_itc"] += row["TOTAL_TAX"]
            supplier_dict[gstin]["missing_2b_count"] += 1
            supplier_dict[gstin]["missing_2b_itc"] += row["TOTAL_TAX"]
            supplier_dict[gstin]["total_invoices"] += 1
    
    # Process Missing in Books (GSTR-2B data)
    if not results["missing_in_books"].empty:
        for _, row in results["missing_in_books"].iterrows():
            gstin = row["GSTIN"]
            name = row["Trade_Name"]
            
            if gstin not in supplier_dict:
                supplier_dict[gstin] = {
                    "names": set(),
                    "books_itc": 0,
                    "gstr_itc": 0,
                    "missing_2b_count": 0,
                    "missing_2b_itc": 0,
                    "missing_books_count": 0,
                    "missing_books_itc": 0,
                    "total_invoices": 0
                }
            
            supplier_dict[gstin]["names"].add(name)
            supplier_dict[gstin]["gstr_itc"] += row["TOTAL_TAX"]
            supplier_dict[gstin]["missing_books_count"] += 1
            supplier_dict[gstin]["missing_books_itc"] += row["TOTAL_TAX"]
            supplier_dict[gstin]["total_invoices"] += 1
    
    # ================= TOP INSIGHTS SECTION ================= #
    
    st.markdown("## 🔍 Top Insights")
    
    if supplier_dict:
        # Convert to list for analysis
        supplier_list = []
        for gstin, data in supplier_dict.items():
            names_list = list(data["names"])
            primary_name = max(names_list, key=len)  # Use longest name as primary
            supplier_list.append({
                "gstin": gstin,
                "name": primary_name,
                "books_itc": data["books_itc"],
                "gstr_itc": data["gstr_itc"],
                "difference": data["gstr_itc"] - data["books_itc"],
                "missing_2b_itc": data["missing_2b_itc"],
                "missing_books_itc": data["missing_books_itc"]
            })
        
        supplier_df = pd.DataFrame(supplier_list)
        
        # Top 3 Insights Cards
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # Highest Positive Difference
            if not supplier_df.empty:
                top_diff = supplier_df.nlargest(1, "difference").iloc[0]
                st.markdown(f'''
                <div class="insight-box">
                    <h3 style="margin:0; font-size:1rem; opacity:0.9;">🏆 HIGHEST ITC DIFFERENCE</h3>
                    <p style="margin:0.5rem 0; font-size:1.5rem; font-weight:700;">{top_diff['name']}</p>
                    <p style="margin:0; font-size:1.2rem;">₹{top_diff['difference']:,.2f}</p>
                    <p style="margin:0.5rem 0 0 0; font-size:0.9rem; opacity:0.8;">GSTIN: {top_diff['gstin']}</p>
                </div>
                ''', unsafe_allow_html=True)
        
        with col2:
            # Highest ITC in Books
            if not supplier_df.empty:
                top_books = supplier_df.nlargest(1, "books_itc").iloc[0]
                st.markdown(f'''
                <div class="insight-box" style="background: linear-gradient(135deg, #059669 0%, #047857 100%);">
                    <h3 style="margin:0; font-size:1rem; opacity:0.9;">📈 HIGHEST ITC IN BOOKS</h3>
                    <p style="margin:0.5rem 0; font-size:1.5rem; font-weight:700;">{top_books['name']}</p>
                    <p style="margin:0; font-size:1.2rem;">₹{top_books['books_itc']:,.2f}</p>
                    <p style="margin:0.5rem 0 0 0; font-size:0.9rem; opacity:0.8;">GSTIN: {top_books['gstin']}</p>
                </div>
                ''', unsafe_allow_html=True)
        
        with col3:
            # Highest Missing in 2B
            if not supplier_df.empty:
                top_missing = supplier_df.nlargest(1, "missing_2b_itc").iloc[0]
                st.markdown(f'''
                <div class="insight-box" style="background: linear-gradient(135deg, #DC2626 0%, #B91C1C 100%);">
                    <h3 style="margin:0; font-size:1rem; opacity:0.9;">📊 HIGHEST MISSING IN 2B</h3>
                    <p style="margin:0.5rem 0; font-size:1.5rem; font-weight:700;">{top_missing['name']}</p>
                    <p style="margin:0; font-size:1.2rem;">₹{top_missing['missing_2b_itc']:,.2f}</p>
                    <p style="margin:0.5rem 0 0 0; font-size:0.9rem; opacity:0.8;">GSTIN: {top_missing['gstin']}</p>
                </div>
                ''', unsafe_allow_html=True)
        
        st.divider()
        
        # ================= QUICK STATS DASHBOARD ================= #
        
        st.markdown("## 📌 Quick Stats")
        
        col1, col2, col3, col4 = st.columns(4)
        
        # Calculate stats
        total_suppliers = len(supplier_df)
        critical_suppliers = len(supplier_df[abs(supplier_df["difference"]) > 100000])
        watch_suppliers = len(supplier_df[(abs(supplier_df["difference"]) > 10000) & (abs(supplier_df["difference"]) <= 100000)])
        healthy_suppliers = len(supplier_df[abs(supplier_df["difference"]) <= 10000])
        total_itc_at_stake = supplier_df["missing_2b_itc"].sum() + supplier_df["missing_books_itc"].sum()
        
        with col1:
            st.markdown(f'''
            <div class="stat-box">
                <p style="color:#6B7280; margin:0; font-size:0.9rem;">Total Suppliers</p>
                <p style="font-size:2rem; font-weight:700; margin:0; color:#1F2937;">{total_suppliers}</p>
            </div>
            ''', unsafe_allow_html=True)
        
        with col2:
            st.markdown(f'''
            <div class="stat-box">
                <p style="color:#6B7280; margin:0; font-size:0.9rem;">🔴 Critical</p>
                <p style="font-size:2rem; font-weight:700; margin:0; color:#DC2626;">{critical_suppliers}</p>
                <p style="color:#6B7280; margin:0; font-size:0.8rem;">Diff > ₹1L</p>
            </div>
            ''', unsafe_allow_html=True)
        
        with col3:
            st.markdown(f'''
            <div class="stat-box">
                <p style="color:#6B7280; margin:0; font-size:0.9rem;">🟡 Watch</p>
                <p style="font-size:2rem; font-weight:700; margin:0; color:#D97706;">{watch_suppliers}</p>
                <p style="color:#6B7280; margin:0; font-size:0.8rem;">Diff > ₹10k</p>
            </div>
            ''', unsafe_allow_html=True)
        
        with col4:
            st.markdown(f'''
            <div class="stat-box">
                <p style="color:#6B7280; margin:0; font-size:0.9rem;">🟢 Healthy</p>
                <p style="font-size:2rem; font-weight:700; margin:0; color:#059669;">{healthy_suppliers}</p>
                <p style="color:#6B7280; margin:0; font-size:0.8rem;">Diff < ₹10k</p>
            </div>
            ''', unsafe_allow_html=True)
        
        st.markdown(f'''
        <div style="background-color:#F3F4F6; padding:1rem; border-radius:0.75rem; margin-top:1rem;">
            <p style="font-size:1.2rem; margin:0; color:#374151;">💰 Total ITC at Stake: <strong>₹{total_itc_at_stake:,.2f}</strong></p>
            <p style="color:#6B7280; margin:0.5rem 0 0 0; font-size:0.9rem;">⚠️ Suppliers needing immediate action: {critical_suppliers + watch_suppliers}</p>
        </div>
        ''', unsafe_allow_html=True)
        
        st.divider()
        
        # ================= TOP 5 LISTS ================= #
        
        st.markdown("## 🏆 Top 5 Lists")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown('<div class="top-five">', unsafe_allow_html=True)
            st.markdown("### 🔥 Top 5 by ITC Difference")
            top5_diff = supplier_df.nlargest(5, "difference")[["name", "difference"]].copy()
            top5_diff["difference"] = top5_diff["difference"].apply(lambda x: f"₹{x:,.2f}")
            top5_diff.columns = ["Supplier", "Difference"]
            st.dataframe(top5_diff, use_container_width=True, hide_index=True)
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col2:
            st.markdown('<div class="top-five">', unsafe_allow_html=True)
            st.markdown("### 📦 Top 5 by Missing ITC")
            top5_missing = supplier_df.nlargest(5, "missing_2b_itc")[["name", "missing_2b_itc"]].copy()
            top5_missing["missing_2b_itc"] = top5_missing["missing_2b_itc"].apply(lambda x: f"₹{x:,.2f}")
            top5_missing.columns = ["Supplier", "Missing ITC"]
            st.dataframe(top5_missing, use_container_width=True, hide_index=True)
            st.markdown('</div>', unsafe_allow_html=True)
        
        st.divider()
    
    # ================= TABS ================= #

    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "📚 Books Data",
        "📖 GSTR-2B Data",
        "❌ Missing in 2B",
        "📕 Missing in Books",
        "🟡 NO ITC Invoices",
        "⚠️ Invalid GSTIN",
        "📊 Supplier Analysis"
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
            
            # Total
            total_itc = results["missing_in_2b"]["TOTAL_TAX"].sum()
            st.markdown(f"**Total ITC Impact: ₹{total_itc:,.2f}**")
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
            
            # Total
            total_itc = results["missing_in_books"]["TOTAL_TAX"].sum()
            st.markdown(f"**Total ITC Impact: ₹{total_itc:,.2f}**")
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

    # ===== TAB 7: Supplier Analysis (Grouped by GSTIN with Name Variations) =====
    with tab7:
        st.markdown("### 📊 Supplier-wise Analysis (Grouped by GSTIN)")
        
        if supplier_dict:
            # Convert to DataFrame with all columns
            supplier_rows = []
            for gstin, data in supplier_dict.items():
                names_list = list(data["names"])
                # Sort names by length (longest first) for better primary name
                names_list.sort(key=len, reverse=True)
                primary_name = names_list[0]
                other_names = ", ".join(names_list[1:]) if len(names_list) > 1 else "-"
                
                supplier_rows.append({
                    "GSTIN": gstin,
                    "Primary Name": primary_name,
                    "Other Names": other_names,
                    "ITC - Books": data["books_itc"],
                    "ITC - GSTR-2B": data["gstr_itc"],
                    "Difference": data["gstr_itc"] - data["books_itc"],
                    "Missing in 2B (ITC)": data["missing_2b_itc"]
                })
            
            supplier_df = pd.DataFrame(supplier_rows)
            
            # Sort by absolute difference
            supplier_df["Abs_Difference"] = abs(supplier_df["Difference"])
            supplier_df = supplier_df.sort_values("Abs_Difference", ascending=False)
            
            # Create display version with formatted currency
            display_df = supplier_df.copy()
            for col in ["ITC - Books", "ITC - GSTR-2B", "Difference", "Missing in 2B (ITC)"]:
                display_df[col] = display_df[col].apply(lambda x: f"₹{x:,.2f}")
            
            display_df = display_df.drop("Abs_Difference", axis=1)
            
            # Add color coding based on difference
            def color_diff(val):
                if isinstance(val, str) and '₹' in val:
                    num_val = float(val.replace('₹', '').replace(',', ''))
                    if num_val > 100000:
                        return '🔴 '
                    elif num_val > 10000:
                        return '🟡 '
                    elif num_val < -100000:
                        return '🔴 '
                    elif num_val < -10000:
                        return '🟡 '
                return ''
            
            display_df["Difference"] = display_df["Difference"].apply(lambda x: color_diff(x) + x)
            
            st.dataframe(display_df, use_container_width=True, hide_index=True)
            
            # Summary stats for this tab
            st.markdown("---")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                pos_diff = supplier_df[supplier_df["Difference"] > 0]["Difference"].sum()
                st.metric("📈 Net ITC in 2B", f"₹{pos_diff:,.2f}")
            
            with col2:
                neg_diff = abs(supplier_df[supplier_df["Difference"] < 0]["Difference"].sum())
                st.metric("📉 Net ITC in Books", f"₹{neg_diff:,.2f}")
            
            with col3:
                st.metric("🏢 Total Suppliers", len(supplier_df))
        else:
            st.info("No supplier data available for analysis.")

    # ================= DOWNLOAD EXCEL REPORT (WITH AUTO-FIT COLUMNS) ================= #

    st.divider()
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        
        # Set default font
        workbook.formats[0].set_font_name('Aptos Narrow')
        
        # Formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'font_name': 'Aptos Narrow',
            'align': 'left',
            'font_color': '#1E3A8A'
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'font_name': 'Aptos Narrow',
            'bg_color': '#D9E1F2',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        text_format = workbook.add_format({'font_name': 'Aptos Narrow'})
        money_format = workbook.add_format({'font_name': 'Aptos Narrow', 'num_format': '₹#,##0.00'})
        
        # ===== 1. DASHBOARD SHEET =====
        dashboard_sheet = workbook.add_worksheet("Dashboard")
        dashboard_sheet.set_column('A:A', 35)
        dashboard_sheet.set_column('B:B', 20)
        dashboard_sheet.set_column('C:C', 20)
        dashboard_sheet.set_column('D:D', 20)
        
        # Title
        dashboard_sheet.merge_range('A1:D1', 'GST RECONCILIATION DASHBOARD', title_format)
        
        # Key Metrics
        dashboard_sheet.write('A3', 'Key Metrics', title_format)
        metrics = [
            ['Total ITC - Books', summary['Total_ITC_Books']],
            ['Total ITC - GSTR-2B', summary['Total_ITC_2B']],
            ['Net Difference', summary['ITC_Difference']],
            ['Match %', f"{summary['Match_Percentage']}%"]
        ]
        
        for i, (label, value) in enumerate(metrics):
            dashboard_sheet.write(i+4, 0, label, header_format)
            if 'ITC' in label:
                dashboard_sheet.write(i+4, 1, value, money_format)
            else:
                dashboard_sheet.write(i+4, 1, value, text_format)
        
        # Top Insights
        if supplier_dict:
            supplier_list = []
            for gstin, data in supplier_dict.items():
                names_list = list(data["names"])
                primary_name = max(names_list, key=len)
                supplier_list.append({
                    "name": primary_name,
                    "difference": data["gstr_itc"] - data["books_itc"],
                    "missing_2b": data["missing_2b_itc"]
                })
            
            insight_df = pd.DataFrame(supplier_list)
            
            if not insight_df.empty:
                row_start = 10
                dashboard_sheet.write(row_start, 0, 'Top Insights', title_format)
                
                # Highest Difference
                top_diff = insight_df.nlargest(1, "difference").iloc[0]
                dashboard_sheet.write(row_start+1, 0, '🏆 Highest ITC Difference:', header_format)
                dashboard_sheet.write(row_start+1, 1, top_diff['name'], text_format)
                dashboard_sheet.write(row_start+1, 2, top_diff['difference'], money_format)
                
                # Highest Missing
                top_missing = insight_df.nlargest(1, "missing_2b").iloc[0]
                dashboard_sheet.write(row_start+2, 0, '📊 Highest Missing in 2B:', header_format)
                dashboard_sheet.write(row_start+2, 1, top_missing['name'], text_format)
                dashboard_sheet.write(row_start+2, 2, top_missing['missing_2b'], money_format)
        
        # Auto-fit columns
        dashboard_sheet.autofit()
        
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
            worksheet.autofit()
        
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
            worksheet.autofit()
        
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
            worksheet.autofit()
        
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
            worksheet.autofit()
        
        # ===== 6. SUPPLIER ANALYSIS SHEET (Grouped by GSTIN) =====
        if supplier_dict:
            analysis_sheet = workbook.add_worksheet("Supplier Analysis")
            
            supplier_rows = []
            for gstin, data in supplier_dict.items():
                names_list = list(data["names"])
                names_list.sort(key=len, reverse=True)
                primary_name = names_list[0]
                other_names = ", ".join(names_list[1:]) if len(names_list) > 1 else "-"
                
                supplier_rows.append({
                    "GSTIN": gstin,
                    "Primary Name": primary_name,
                    "Other Names": other_names,
                    "ITC - Books": data["books_itc"],
                    "ITC - GSTR-2B": data["gstr_itc"],
                    "Difference": data["gstr_itc"] - data["books_itc"],
                    "Missing in 2B (ITC)": data["missing_2b_itc"]
                })
            
            analysis_df = pd.DataFrame(supplier_rows)
            analysis_df = analysis_df.sort_values("Difference", ascending=False)
            
            analysis_df.to_excel(writer, sheet_name="Supplier Analysis", index=False, startrow=1)
            
            # Format
            worksheet = writer.sheets["Supplier Analysis"]
            worksheet.write(0, 0, "Supplier-wise Analysis (Grouped by GSTIN)", title_format)
            
            for col_num, col in enumerate(analysis_df.columns):
                worksheet.write(1, col_num, col, header_format)
                if "ITC" in col or "Difference" in col:
                    worksheet.set_column(col_num, col_num, 18, money_format)
                else:
                    worksheet.set_column(col_num, col_num, 25)
            
            worksheet.freeze_panes(2, 0)
            worksheet.autofit()
        
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
            worksheet.autofit()
        
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
            worksheet.autofit()
    
    output.seek(0)
    
    st.download_button(
        "⬇ Download Excel Report (Auto-fit Columns)",
        data=output,
        file_name=f"GST_Reconciliation_{datetime.now().strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )
