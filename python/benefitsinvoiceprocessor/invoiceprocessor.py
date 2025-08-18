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

    # --- Mapping by Template Desc ---
    df_code_map_desc = df_code_map[df_code_map['Template Desc'].notna() & (df_code_map['Template Desc'].astype(str).str.strip() != '')]

    description_totals = {}
    for _, row in df_code_map_desc.iterrows():
        desc = str(row['Template Desc']).strip()
        division_code = str(row.get('Division Code', '')).strip()
        company_code = str(row.get('Invoice Company Code', '')).strip()

        filtered_df = pd.DataFrame()

   
    if division_code:
        filtered_df = df_invoice[df_invoice['Division'].astype(str).str.strip() == division_code]
    elif company_code:
        filtered_df = df_invoice[df_invoice['Company'].astype(str).str.strip() == company_code]
    else:
        continue

    total = filtered_df['Monthly Premium'].sum()
    if total > 0:
        description_totals[desc] = total

    for desc, total in description_totals.items():
        match_rows = df_template[df_template['DESC'].astype(str).str.strip().str.lower() == desc.lower()].index
        df_template.loc[match_rows, 'NET'] = total

    # --- Mapping by Inter-Co (no Template Desc) ---
    df_code_map_no_desc = df_code_map[df_code_map['Template Desc'].isna() | (df_code_map['Template Desc'].astype(str).str.strip() == '')]

    company_totals = df_invoice[['Company', 'Monthly Premium']].dropna().groupby('Company')['Monthly Premium'].sum().reset_index()
    interco_mapping = dict(zip(df_code_map_no_desc['Invoice Company Code'], df_code_map_no_desc['Template Inter-Co']))

    for invoice_company, interco in interco_mapping.items():
        total = company_totals[company_totals['Company'] == invoice_company]['Monthly Premium'].sum()
        if total > 0:
            match_rows = df_template[df_template['Inter-Co'].astype(str).str.strip() == str(interco).strip()].index
            df_template.loc[match_rows, 'NET'] = total

    # --- HHI and THC Department Mapping ---
    df_hhi_thc = df_invoice[df_invoice['Company'].isin(['HHI', 'THC'])].copy()
    df_hhi_thc['Department'] = df_hhi_thc['Department'].astype(str)

    dept_map = dict(zip(df_hhi_thc_map['Invoice Department Code'].astype(str), df_hhi_thc_map['Template CC'].astype(str)))
    df_hhi_thc['CC_Code'] = df_hhi_thc['Department'].map(dept_map)

    df_cc_totals = df_hhi_thc.groupby('CC_Code')['Monthly Premium'].sum().reset_index()
    df_template['CC'] = df_template['CC'].astype(str).str.replace(r'\.0$', '', regex=True)

    cc_merged_df = pd.merge(df_template, df_cc_totals, left_on='CC', right_on='CC_Code', how='left')
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
        file_name="Complete_Aflac_Medius_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

