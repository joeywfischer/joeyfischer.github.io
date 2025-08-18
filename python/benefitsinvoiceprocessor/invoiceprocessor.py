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
    df_invoice['Department'] = df_invoice['Department'].astype(str).str.strip()
    df_invoice['Monthly Premium'] = pd.to_numeric(df_invoice['Monthly Premium'], errors='coerce')

   # --- HHI and THC Department Mapping ---
    df_hhi_thc = df_invoice[df_invoice['Company'].isin(['HHI', 'THC'])].copy()
    df_hhi_thc['Department'] = df_hhi_thc['Department'].astype(str).str.strip()

    # Normalize mapping sheet columns
    df_hhi_thc_map['Invoice Department Code'] = df_hhi_thc_map['Invoice Department Code'].astype(str).str.strip()
    df_hhi_thc_map['Template CC'] = df_hhi_thc_map['Template CC'].astype(str).str.strip()

    # Create mapping from Department to Template CC
    dept_to_cc_map = dict(zip(df_hhi_thc_map['Invoice Department Code'], df_hhi_thc_map['Template CC']))
    df_hhi_thc['CC_Code'] = df_hhi_thc['Department'].map(dept_to_cc_map)

    # Group by CC_Code and sum Monthly Premium
    df_cc_totals = df_hhi_thc.groupby('CC_Code')['Monthly Premium'].sum().reset_index()

    # Normalize CC column in template for matching
    df_template['CC'] = df_template['CC'].astype(str).str.strip()

    # Merge and update NET column only
    cc_merged_df = pd.merge(df_template[['CC', 'NET']], df_cc_totals, left_on='CC', right_on='CC_Code', how='left')
    match_indices = cc_merged_df[cc_merged_df['Monthly Premium'].notna()].index
    df_template.loc[match_indices, 'NET'] = cc_merged_df.loc[match_indices, 'Monthly Premium']

    # --- Final Output ---
    df_template['CC'] = df_template['CC'].replace('nan', '', regex=False)

    output = io.BytesIO()
    df_template.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)

    st.success("Processing complete!")
    st.download_button(
        label="Download Updated Medius Template",
        data=output,
        file_name="Updated_Aflac_Medius_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
