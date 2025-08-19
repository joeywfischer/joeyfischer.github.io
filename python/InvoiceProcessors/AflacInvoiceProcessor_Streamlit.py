import streamlit as st
import pandas as pd

st.title("Aflac Invoice and Support Generator")

# Upload both files
invoice_file = st.file_uploader("Upload Invoice Excel File", type=["xlsx"], key="invoice")
template_file = st.file_uploader("Upload Template Excel File (with Code Map sheet)", type=["xlsx"], key="template")

if invoice_file and template_file:
    invoice_xls = pd.ExcelFile(invoice_file, engine='openpyxl')
    detail_df = pd.read_excel(invoice_xls, sheet_name='Detail', engine='openpyxl')

    template_xls = pd.ExcelFile(template_file, engine='openpyxl')
    code_map_df = pd.read_excel(template_xls, sheet_name='Code Map', engine='openpyxl')

    # First table: Summary by Company
    premium_summary = detail_df.groupby('Company')['Monthly Premium'].sum().reset_index()
    premium_summary.columns = ['Row Labels', 'Sum of Monthly Premium']

    result_df = premium_summary.merge(
        code_map_df[['Invoice Company Code', 'Company Description']],
        left_on='Row Labels', right_on='Invoice Company Code', how='left'
    )

    result_df.rename(columns={'Company Description': 'Full Company Name'}, inplace=True)
    final_df = result_df[['Row Labels', 'Sum of Monthly Premium', 'Full Company Name']]

    # Second table: Summary by Division Description grouped by Invoice Company Code
    merged_df = detail_df.merge(
        code_map_df[['Invoice Company Code', 'Division Description']],
        left_on='Company', right_on='Invoice Company Code', how='left'
    )

    division_summary = (
        merged_df[merged_df['Division Description'].notna()]
        .groupby(['Invoice Company Code', 'Division Description'])['Monthly Premium']
        .sum()
        .reset_index()
    )
    division_summary.columns = ['Invoice Company Code', 'Division', 'Sum of Monthly Premium']

    # Third table: Additional Break-Down by Division Code
    if 'Division Code' in code_map_df.columns:
        breakdown_df = detail_df.merge(
            code_map_df[['Invoice Company Code', 'Division Description', 'Division Code']],
            left_on='Company', right_on='Invoice Company Code', how='left'
        )

        breakdown_df = breakdown_df[breakdown_df['Division Description'].notna() & breakdown_df['Division Code'].notna()]
        breakdown_summary = breakdown_df.groupby(['Division Description', 'Division Code'])['Monthly Premium'].sum().reset_index()
        breakdown_summary.columns = ['Division Description', 'Division', 'Cost']
    else:
        breakdown_summary = pd.DataFrame(columns=['Division Description', 'Division', 'Cost'])

    st.dataframe(final_df)
    st.dataframe(division_summary)
    st.dataframe(breakdown_summary)

    def convert_df_to_excel(df1, df2, df3):
        from io import BytesIO
        from openpyxl import Workbook
        from openpyxl.utils.dataframe import dataframe_to_rows

        output = BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = 'Aflac Invoice and Support'

        # Write first table starting at A1
        for r in dataframe_to_rows(df1, index=False, header=True):
            ws.append(r)

        # Write second table starting at E1
        for i, r in enumerate(dataframe_to_rows(df2, index=False, header=True), start=1):
            for j, val in enumerate(r, start=5):  # Column E is index 5
                ws.cell(row=i, column=j, value=val)

        # Write third table below the existing tables
        start_row = max(len(df1), len(df2)) + 3
        ws.cell(row=start_row, column=1, value='Additional Break-Down')
        for i, r in enumerate(dataframe_to_rows(df3, index=False, header=True), start=start_row + 1):
            for j, val in enumerate(r, start=1):
                ws.cell(row=i, column=j, value=val)

        wb.save(output)
        return output.getvalue()

    excel_data = convert_df_to_excel(final_df, division_summary, breakdown_summary)
    st.download_button(
        label="Download Aflac Invoice and Support Excel",
        data=excel_data,
        file_name="Aflac_Invoice_and_Support.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
