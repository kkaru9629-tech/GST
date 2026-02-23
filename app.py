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

    # ================= EXECUTIVE SUMMARY ================= #

    st.markdown("## 📌 Executive Summary")

    col1, col2, col3 = st.columns(3)

    col1.metric("Total Invoices (Books)", summary["Total_Invoices_Books"])
    col1.metric("Total ITC (Books)", f"₹{summary['Total_ITC_Books']:,.2f}")

    col2.metric("Total Invoices (GSTR-2B)", summary["Total_Invoices_2B"])
    col2.metric("Total ITC (GSTR-2B)", f"₹{summary['Total_ITC_2B']:,.2f}")

    col3.metric("Fully Matched", summary["Total_Matched"])
    col3.metric("Match %", f"{summary['Match_Percentage']} %")

    st.divider()

    col4, col5 = st.columns(2)
    col4.metric("Missing in Books", summary["Total_Missing_Books"])
    col5.metric("Missing in 2B", summary["Total_Missing_2B"])

    st.metric("ITC Difference (2B - Books)", f"₹{summary['ITC_Difference']:,.2f}")

    st.divider()

    # ================= TABS ================= #

    tab1, tab2, tab3, tab4 = st.tabs([
        "Fully Matched",
        "Missing in Books",
        "Missing in 2B",
        "NO ITC Invoices"
    ])

    with tab1:
        st.dataframe(results["fully_matched"], use_container_width=True)

    with tab2:
        st.dataframe(results["missing_in_books"], use_container_width=True)

    with tab3:
        st.dataframe(results["missing_in_2b"], use_container_width=True)

    with tab4:
        st.dataframe(results["no_itc"], use_container_width=True)

    # ================= EXCEL DOWNLOAD ================= #

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        header_format = workbook.add_format({
            'bold': True,
            'border': 1,
            'align': 'center',
            'bg_color': '#D9E1F2'
        })

        money_format = workbook.add_format({
            'num_format': '₹#,##0.00',
            'border': 1
        })

        title_format = workbook.add_format({
            'bold': True,
            'font_size': 18,
            'align': 'center'
        })

        # Executive Summary Sheet
        summary_sheet = workbook.add_worksheet("Executive Summary")
        summary_sheet.merge_range("A1:D1", "GST RECONCILIATION REPORT", title_format)
        summary_sheet.write_row("A3", ["Particulars", "Value"], header_format)

        summary_data = [
            ["Total Invoices - Books", summary["Total_Invoices_Books"]],
            ["Total Invoices - 2B", summary["Total_Invoices_2B"]],
            ["Fully Matched", summary["Total_Matched"]],
            ["Missing in Books", summary["Total_Missing_Books"]],
            ["Missing in 2B", summary["Total_Missing_2B"]],
            ["Match %", summary["Match_Percentage"]],
            ["Total ITC - Books", summary["Total_ITC_Books"]],
            ["Total ITC - 2B", summary["Total_ITC_2B"]],
            ["ITC Difference", summary["ITC_Difference"]],
        ]

        row = 3
        for item in summary_data:
            summary_sheet.write(row, 0, item[0])
            if "ITC" in item[0]:
                summary_sheet.write(row, 1, item[1], money_format)
            else:
                summary_sheet.write(row, 1, item[1])
            row += 1

        summary_sheet.set_column("A:B", 35)

        def write_sheet(df, name):
            if not df.empty:
                df.to_excel(writer, sheet_name=name, index=False)
                worksheet = writer.sheets[name]
                for col_num, col_name in enumerate(df.columns):
                    worksheet.write(0, col_num, col_name, header_format)
                worksheet.freeze_panes(1, 0)

        write_sheet(results["fully_matched"], "Fully Matched")
        write_sheet(results["missing_in_books"], "Missing in Books")
        write_sheet(results["missing_in_2b"], "Missing in 2B")
        write_sheet(results["value_mismatch"], "Value Mismatch")
        write_sheet(results["tax_mismatch"], "Tax Mismatch")
        write_sheet(results["no_itc"], "NO ITC Invoices")

    output.seek(0)

    st.download_button(
        "⬇ Download Professional Excel Report",
        data=output,
        file_name=f"GST_Reconciliation_Report_{datetime.now().strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
