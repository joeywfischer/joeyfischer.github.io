import streamlit as st
import pandas as pd
import io

st.title("Aflac Invoice Processor")

invoice_file = st.file_uploader("Upload Aflac Invoice Excel File", type=["xlsx"])
template_file = st.file_uploader("Upload Medius Template Excel File", type=["xlsx"])

if invoice_file and template_file:
    # Load invoice and template data
    df_invoice = pd.read_excel(invoice_file, sheet_name='Detail', engine='openpyxl')
    df_template = pd.read_excel(template_file, sheet_name=0, engine='openpyxl')

    # Load mapping sheets
    df_code_map = pd.read_excel(template_file, sheet_name='Code Map', engine='openpyxl')
    df_hhi_thc_map = pd.read_excel(template_file, sheet_name='HHI and THC Code Map', engine='openpyxl')

    # Normalize key columns
    df_invoice['Company'] = df_invoice['Company'].astype(str).str.strip().str.upper()
    df_invoice['Division'] = df_invoice['Division'].astype(str).str.strip()
    df_invoice['Monthly Premium'] = pd.to_numeric(df_invoice['Monthly Premium'], errors='coerce')

    # --- Mapping by Template Desc ---
    df_code_map_desc = df_code_map[df_code_map['Template Desc'].notna() & (df_code_map['Template Desc'].astype(str).str.strip() != '')]

    description_totals = {}

    for _, row in df_code_map_desc.iterrows():
        desc = str(row['Template Desc']).strip()
        division_code = str(row.get('Division Code', '')).strip()
        invoice_company_code = str(row.get('Invoice Company Code', '')).strip().upper()

        filtered_df = pd.DataFrame()

        # Use Division Code if available
        if division_code:
            filtered_df = df_invoice[df_invoice['Division'] == division_code]
        # Otherwise use Invoice Company Code
        elif invoice_company_code:
            filtered_df = df_invoice[df_invoice['Company'] == invoice_company_code
