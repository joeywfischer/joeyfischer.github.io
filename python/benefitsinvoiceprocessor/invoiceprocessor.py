import streamlit as st
import pandas as pd
import io

st.title("Aflac Invoice Processor")

invoice_file = st.file_uploader("Upload Aflac Invoice Excel File", type=["xlsx"])
template_file = st.file_uploader("Upload Medius Template Excel File", type=["xlsx"])

if invoice_file and template_file:
    # Load invoice and template data
    df_invoice = pd.read_excel(invoice_file, sheet_name='Detail', engine='openpyxl')
    df_template = pd.read_excel(template_file, sheet_name=0, engine='openpyxl')  # First sheet is the template

    # Load mapping sheets
    df_code_map = pd.read_excel(template_file, sheet_name='Code Map', engine='openpyxl')
    df_hhi_thc_map = pd.read_excel(template_file, sheet_name='HHI and THC Code Map', engine='openpyxl')

    # --- Company Code Mapping ---
    company_code_mapping = dict(zip(df_code_map['Template Inter-Co'], df_code_map['Invoice Company Code']))
    df_template_for_merge = df_template[['Inter-Co', 'DESC', 'NET']].copy()
    df_template_for_merge['Mapped_Invoice_Company'] = df_template_for_merge['Inter-Co'].replace(company_code_mapping)

    df_relevant = df_invoice[['Company', 'Monthly Premium']].dropna()
    df_company_totals = df_relevant.groupby('Company')['Monthly Premium'].sum().reset_index()

    merged_df = pd.merge(df_template_for_merge, df_company_totals, left_on='Mapped_Invoice_Company', right_on='Company', how='left')
    new_net_values = merged_df.set_index(df_template.index)['Monthly Premium']
    update_condition = new_net_values.notna() & (df_template['Inter-Co'] != df_code_map['Template Inter-Co'].unique())
    df_template.loc[update_condition, 'NET'] = new_net_values[update_condition]

    # --- Description Source Mapping ---
    df_code_map_filtered = df_code_map[df_code_map['Template Desc'].notna() & (df_code_map['Template Desc'].astype(str).str.strip() != '')]

    description_totals = {}
    for _, row in df_code_map_filtered.iterrows():
        desc = str(row['Template Desc']).strip()
        company_code = str(row['Invoice Company Code']).strip()
        division_code = str(row.get('Division Code', '')).strip()

        if division_code:  # Use both company and division
            filtered_df = df_invoice[
                (df_invoice['Company'] == company_code) &
                (df_invoice['Division'].astype(str) == division_code)
            ].dropna(subset=['Monthly Premium'])
        else:  # Use only company
            filtered_df = df_invoice[
                df_invoice['Company'] == company_code
            ].dropna(subset=['Monthly Premium'])

        description_totals[desc] = filtered_df['Monthly Premium'].sum()

    for desc, total in description_totals.items():
        rows_to_update = df_template[df_template['DESC'].str.contains(desc, case=False, na=False)].index
        df_template.loc[rows_to_update, 'NET'] = total

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


