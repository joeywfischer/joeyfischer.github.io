
import streamlit as st
import pandas as pd
import io

# Title
st.title("Aflac Invoice Processor")

# File upload
invoice_file = st.file_uploader("Upload Aflac Invoice Excel File", type=["xlsx"])
template_file = st.file_uploader("Upload Medius Template Excel File", type=["xlsx"])

if invoice_file and template_file:
    # Load invoice data
    df_invoice = pd.read_excel(invoice_file, sheet_name='Detail', engine='openpyxl')
    df_template = pd.read_excel(template_file, engine='openpyxl')

    # Group by Company and sum Monthly Premium
    df_relevant = df_invoice[['Company', 'Monthly Premium']].dropna(subset=['Company', 'Monthly Premium'])
    df_company_totals = df_relevant.groupby('Company')['Monthly Premium'].sum().reset_index()

    # Mapping from template Inter-Co to invoice Company
    company_code_mapping = {
        'BARTL': 'BART',
        'CECOC': 'CECO',
        'CFACO': 'CFA',
        'CMSHR': 'CMSHD',
        'DVIES': 'DAVIE',
        'HCGKC': 'HCG',
        'INSIG': 'INVIS',
        'BAPRD': 'MUMSR',
        'NSSTI': 'NSTOK',
        'STRAN': 'NSTRD',
        'PBCRP': 'PB',
        'PTLMI': 'PTL',
        'SJHAM': 'SJH',
        'SARUS': 'SRCLD',
        'SHRED': 'STECH',
        'VERSA': 'VMD'
    }

    # Prepare template for merge
    df_template_for_merge = df_template[['Inter-Co', 'DESC', 'NET']].copy()
    df_template_for_merge['Mapped_Invoice_Company'] = df_template_for_merge['Inter-Co'].replace(company_code_mapping)

    # Merge and update NET column
    merged_df = pd.merge(df_template_for_merge, df_company_totals, left_on='Mapped_Invoice_Company', right_on='Company', how='left')
    new_net_values = merged_df.set_index(df_template.index)['Monthly Premium']
    update_condition = new_net_values.notna() & (df_template['Inter-Co'] != 'CADES')
    df_template.loc[update_condition, 'NET'] = new_net_values[update_condition]

    # Replace 'nan' string in CC column
    df_template['CC'] = df_template['CC'].replace('nan', '', regex=False)

    # Save to buffer
    output = io.BytesIO()
    df_template.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)

    # Download button
    st.success("Processing complete!")
    st.download_button(
        label="Download Updated Medius Template",
        data=output,
        file_name="Complete_Aflac_Medius_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
