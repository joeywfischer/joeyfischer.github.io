import streamlit as st
import pandas as pd

st.title("Aflac Invoice and Support Generator")

# Upload both files
invoice_file = st.file_uploader("Upload Invoice Excel File", type=["xlsx"], key="invoice")
template_file = st.file_uploader("Upload Template Excel File (with Code Map sheet)", type=["xlsx"], key="template")

if invoice_file and template_file:
    # Load the Detail sheet from the invoice file
    invoice_xls = pd.ExcelFile(invoice_file, engine='openpyxl')
    detail_df = pd.read_excel(invoice_xls, sheet_name='Detail', engine='openpyxl')

    # Load the Code Map sheet from the template file
    template_xls = pd.ExcelFile(template_file, engine='openpyxl')
    code_map_df = pd.read_excel(template_xls, sheet_name='Code Map', engine='openpyxl')

    # Summarize Monthly Premium by Company
    premium_summary = detail_df.groupby('Company')['Monthly Premium'].sum().reset_index()
    premium_summary.columns = ['Row Labels', 'Sum of Monthly Premium']

    # Merge with Code Map to get Full Company Name
    result_df = premium_summary.merge(
        code_map_df[['Invoice Company Code', 'Company Description']],
        left_on='Row Labels', right_on='Invoice Company Code', how='left'
    )

    result_df.rename(columns={'Company Description': 'Full Company Name'}, inplace=True)
    final_df = result_df[['Row Labels', 'Sum of Monthly Premium', 'Full Company Name']]

    st.dataframe(final_df)

    def convert_df_to_excel(df):
        from io import BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Aflac Invoice and Support')
        return output.getvalue()

    excel_data = convert_df_to_excel(final_df)
    st.download_button(
        label="Download Aflac Invoice and Support Excel",
        data=excel_data,
        file_name="Aflac_Invoice_and_Support.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

