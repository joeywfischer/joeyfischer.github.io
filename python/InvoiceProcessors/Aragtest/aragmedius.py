import streamlit as st
import pandas as pd
import io

st.title("Aflac Medius Template Generator")

invoice_file = st.file_uploader("Upload Mapped Invoice CSV File", type=["csv"])
template_file = st.file_uploader("Upload Medius Template Excel File", type=["xlsx"])

if invoice_file and template_file:
    try:
        # Load mapped invoice data
        df_invoice = pd.read_csv(invoice_file)

        # Load the original template structure
        df_template = pd.read_excel(template_file, sheet_name=0, engine='openpyxl')

        # Normalize and clean data
        df_invoice['DESC'] = df_invoice['DESC'].astype(str).str.strip()
        df_invoice['Inter-Co'] = df_invoice['Inter-Co'].astype(str).str.strip()
        df_invoice['CC'] = df_invoice['CC'].astype(str).str.strip()
        df_invoice['G/L ACCT'] = df_invoice['G/L ACCT'].astype(str).str.strip()
        df_invoice['Approver'] = df_invoice['Approver'].astype(str).str.strip()
        df_invoice['Monthly Premium'] = pd.to_numeric(df_invoice['Monthly Premium'], errors='coerce')

        # Aggregate totals by DESC, Inter-Co, CC, G/L ACCT, Approver
        df_aggregated = df_invoice.groupby(
            ['DESC', 'Inter-Co', 'CC', 'G/L ACCT', 'Approver'], dropna=False
        )['Monthly Premium'].sum().reset_index()

        # Rename Monthly Premium to NET
        df_aggregated.rename(columns={'Monthly Premium': 'NET'}, inplace=True)

        # Append aggregated rows to the template
        df_result = pd.concat([df_template, df_aggregated], ignore_index=True)

        # Export to Excel
        output = io.BytesIO()
        df_result.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.success("Template updated with aggregated invoice data!")
        st.download_button(
            label="Download Updated Medius Template",
            data=output,
            file_name="Updated_Aflac_Medius_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
