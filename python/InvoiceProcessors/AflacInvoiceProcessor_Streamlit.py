import streamlit as st
import pandas as pd

st.title("Aflac Invoice and Support Generator")

# Upload both files
invoice_file = st.file_uploader("Upload Invoice Excel File", type=["xlsx"], key="invoice")
template_file = st.file_uploader("Upload Template Excel File (with Code Map sheet)", type=["xlsx"], key="template")

if invoice_file and template_file:
    # Load data
    invoice_xls = pd.ExcelFile(invoice_file, engine='openpyxl')
    detail_df = pd.read_excel(invoice_xls, sheet_name='Detail', engine='openpyxl')

    template_xls = pd.ExcelFile(template_file, engine='openpyxl')
    code_map_df = pd.read_excel(template_xls, sheet_name='Code Map', engine='openpyxl')

    # Filter Code Map for non-empty Division Description
    code_map_filtered = code_map_df[code_map_df['Division Description'].notna()]

    # Filter Detail sheet to include only Companies with Division Description in Code Map
    valid_companies = code_map_filtered['Invoice Company Code'].unique()
    filtered_detail_df = detail_df[detail_df['Company'].isin(valid_companies)]

    # Table 1: Summary by Company
    premium_summary = filtered_detail_df.groupby('Company')['Monthly Premium'].sum().reset_index()
    premium_summary.columns = ['Row Labels', 'Sum of Monthly Premium']

    result_df = premium_summary.merge(
        code_map_df[['Invoice Company Code', 'Company Description']],
        left_on='Row Labels', right_on='Invoice Company Code', how='left'
    )

    result_df.rename(columns={'Company Description': 'Full Company Name'}, inplace=True)
    final_df = result_df[['Row Labels', 'Sum of Monthly Premium', 'Full Company Name']]

    # Table 2: Hierarchical breakdown by Company and Division
    breakdown_rows = []
    company_totals = filtered_detail_df.groupby('Company')['Monthly Premium'].sum().reset_index()

    for company, comp_premium in company_totals.values:
        breakdown_rows.append({'Label': company, 'Monthly Premium': comp_premium})
        divisions = filtered_detail_df[filtered_detail_df['Company'] == company].groupby('Division')['Monthly Premium'].sum().reset_index()
        for div, div_premium in divisions.values:
            if pd.notna(div):
                breakdown_rows.append({'Label': f"  {div}", 'Monthly Premium': div_premium})

    breakdown_df = pd.DataFrame(breakdown_rows)

    # Display tables in Streamlit
    st.subheader("Summary Table")
    st.dataframe(final_df)

    st.subheader("Company & Division Breakdown")
    st.dataframe(breakdown_df)

    # Export to Excel
    def convert_df_to_excel(df1, df2):
        from io import BytesIO
        from openpyxl import Workbook
        from openpyxl.utils.dataframe import dataframe_to_rows

        output = BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = 'Aflac Invoice and Support'

        # Write Table 1 at A1
        for r in dataframe_to_rows(df1, index=False, header=True):
            ws.append(r)

        # Write Table 2 at E1
        for i, r in enumerate(dataframe_to_rows(df2, index=False, header=True), start=1):
            for j, val in enumerate(r, start=5):  # Column E is index 5
                ws.cell(row=i, column=j, value=val)

        wb.save(output)
        return output.getvalue()

    excel_data = convert_df_to_excel(final_df, breakdown_df)
    st.download_button(
        label="Download Aflac Invoice and Support Excel",
        data=excel_data,
        file_name="Aflac_Invoice_and_Support.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

