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

    # ================= PROFESSIONAL EXCEL ================= #

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        header_format = workbook.add_format({
            "bold": True,
            "font_name": "Aptos Display",
            "border": 1
        })

        money_format = workbook.add_format({
            "num_format": "#,##0.00",
            "font_name": "Aptos Display"
        })

        def write_sheet(df, name):

            df.to_excel(writer, sheet_name=name, index=False)
            worksheet = writer.sheets[name]

            for col_num, column in enumerate(df.columns):
                worksheet.write(0, col_num, column, header_format)

                # Auto column width (ALT+H+O+I logic)
                max_len = max(
                    df[column].astype(str).map(len).max(),
                    len(column)
                ) + 2
                worksheet.set_column(col_num, col_num, max_len)

                if df[column].dtype in ["float64", "int64"]:
                    worksheet.set_column(col_num, col_num, max_len, money_format)

        summary_df = pd.DataFrame({
            "Metric": [
                "Total Invoices (Books)",
                "Total Invoices (2B)",
                "Matched",
                "ITC as per Books",
                "ITC as per 2B",
                "Difference",
                "Match %"
            ],
            "Value": [
                summary["Total_Invoices_Books"],
                summary["Total_Invoices_2B"],
                summary["Total_Matched"],
                summary["Total_ITC_Books"],
                summary["Total_ITC_2B"],
                summary["ITC_Difference"],
                summary["Match_Percentage"]
            ]
        })

        write_sheet(summary_df, "Summary")
        write_sheet(results["fully_matched"], "Fully Matched")
        write_sheet(results["missing_in_books"], "Missing in Books")
        write_sheet(results["missing_in_2b"], "Missing in 2B")
        write_sheet(results["value_mismatch"], "Value Mismatch")
        write_sheet(results["tax_mismatch"], "Tax Mismatch")

    output.seek(0)

    st.download_button(
        "⬇ Download Professional Excel Report",
        data=output,
        file_name="GST_Reconciliation_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
