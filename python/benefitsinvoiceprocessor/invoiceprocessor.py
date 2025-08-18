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

    # --- Mapping by Template Desc with Division Code ---
    df_code_map_div = df_code_map[
        df_code_map['Template Desc'].notna() &
        (df_code_map['Template Desc'].astype(str).str.strip() != '') &
        (df_code_map['Division Code'].astype(str).str.strip() != '')
    ]

    for _, row in df_code_map_div.iterrows():
        desc = str(row['Template Desc']).strip()
        division_code = str(row['Division Code']).strip()
        filtered_df = df_invoice[df_invoice['Division'] == division_code]
        total = filtered_df['Monthly Premium'].sum()
        if total > 0:
            match_rows = df_template[df_template['DESC'].astype(str).str.strip().str.lower() == desc.lower()].index
            df_template.loc[match_rows, 'NET'] = total

    # --- Mapping by Template Desc with Invoice Company Code (no Division Code) ---
    df_code_map_company = df_code_map[
        df_code_map['Template Desc'].notna() &
        (df_code_map['Template Desc'].astype(str).str.strip() != '') &
        ((df_code_map['Division Code'].isna()) | (df_code_map['Division Code'].astype(str).str.strip() == ''))
    ]

    for _, row in df_code_map_company.iterrows():
        desc = str(row['Template Desc']).strip()
        invoice_company_code = str(row['Invoice Company Code']).strip().upper()
        filtered_df = df_invoice[df_invoice['Company'] == invoice_company_code]
        total = filtered_df['Monthly Premium'].sum()
        if total > 0:
            match_rows = df_template[df_template['DESC'].astype(str).str.strip().str.lower() == desc.lower()].index
            df_template.loc[match_rows, 'NET'] = total

    # --- Mapping by Inter-Co (no Template Desc) ---
    df_code_map_no_desc = df_code_map[
        df_code_map['Template Desc'].isna() |
        (df_code_map['Template Desc'].astype(str).str.strip() == '')
    ]

    company_totals = df_invoice[['Company', 'Monthly Premium']].dropna().groupby('Company')['Monthly Premium'].sum().reset_index()
    interco_mapping = dict(zip(
        df_code_map_no_desc['Invoice Company Code'].astype(str).str.upper(),
        df_code_map_no_desc['Template Inter-Co'].astype(str).str.strip()
    ))

    for invoice_company, interco in interco_mapping.items():
        total = company_totals[company_totals['Company'] == invoice_company]['Monthly Premium'].sum()
        if total > 0:
            match_rows = df_template[df_template['Inter-Co'].astype(str).str.strip() == interco].index
            df_template.loc[match_rows, 'NET'] = total

    # --- Corrected HHI and THC Department Mapping ---
    df_hhi_thc = df_invoice[df_invoice['Company'].isin(['HHI', 'THC'])].copy()
    df_hhi_thc['Department'] = df_hhi_thc['Department'].astype(str).str.strip()

    # Group by Department and sum Monthly Premium
    dept_totals = df_hhi_thc.groupby('Department')['Monthly Premium'].sum().reset_index()

    # Map Department to Invoice Department Code and then to Template CC
    df_hhi_thc_map['Invoice Department Code'] = df_hhi_thc_map['Invoice Department Code'].astype(str).str.strip()
    df_hhi_thc_map['Template CC'] = df_hhi_thc_map['Template CC'].astype(str).str.strip()
    dept_to_cc = dict(zip(df_hhi_thc_map['Invoice Department Code'], df_hhi_thc_map['Template CC']))

    df_template['CC'] = df_template['CC'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

    for _, row in dept_totals.iterrows():
        dept = row['Department']
        total = row['Monthly Premium']
        template_cc = dept_to_cc.get(dept)
        if template_cc:
            match_rows = df_template[df_template['CC'] == template_cc].index
            df_template.loc[match_rows, 'NET'] = total

    # --- Final Output ---
    df_template['CC'] = df_template['CC'].replace('nan', '', regex=False)

    output = io.BytesIO()
    df_template.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)

    st.success("Processing complete!")
    st.download_button(
        label="Download Updated Medius Template",
        data=output,
        file_name="Complete_Aflac_Medius_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
