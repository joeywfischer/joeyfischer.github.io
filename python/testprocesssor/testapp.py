import streamlit as st
import pandas as pd
import io

st.title("Benefits Invoice Processor")
st.header('Aflac')

invoice_file = st.file_uploader("Upload Invoice Excel File", type=["xlsx"])
template_file = st.file_uploader("Upload Medius Template Excel File", type=["xlsx"])

if invoice_file and template_file:
    df_invoice_detail = pd.read_excel(invoice_file, sheet_name='Detail', engine='openpyxl')
    df_template = pd.read_excel(template_file, engine='openpyxl')

    df_relevant = df_invoice_detail[['Company', 'Monthly Premium']].dropna()
    df_company_totals = df_relevant.groupby('Company')['Monthly Premium'].sum().reset_index()

    company_code_mapping = {
        'BARTL': 'BART', 'CECOC': 'CECO', 'CFACO': 'CFA', 'CMSHR': 'CMSHD',
        'DVIES': 'DAVIE', 'HCGKC': 'HCG', 'INSIG': 'INVIS', 'BAPRD': 'MUMSR',
        'NSSTI': 'NSTOK', 'STRAN': 'NSTRD', 'PBCRP': 'PB', 'PTLMI': 'PTL',
        'SJHAM': 'SJH', 'SARUS': 'SRCLD', 'SHRED': 'STECH', 'VERSA': 'VMD'
    }

    df_template_for_merge = df_template[['Inter-Co', 'DESC', 'NET']].copy()
    df_template_for_merge['Mapped_Invoice_Company'] = df_template_for_merge['Inter-Co'].replace(company_code_mapping)

    merged_df = pd.merge(df_template_for_merge, df_company_totals, left_on='Mapped_Invoice_Company', right_on='Company', how='left')
    new_net_values = merged_df.set_index(df_template.index)['Monthly Premium']
    update_condition = new_net_values.notna() & (df_template['Inter-Co'] != 'CADES')
    df_template.loc[update_condition, 'NET'] = new_net_values[update_condition]

    description_source_mapping = {
        'Ancra Aircraft': {'type': 'division', 'company': 'ANCRA', 'code': '20AC'},
        'Ancra Cargo': {'type': 'division', 'company': 'ANCRA', 'code': '20CG'},
        'Ancra Group': {'type': 'division', 'company': 'ANCRA', 'code': '20AN'},
        'CA Design Exeter': {'type': 'division', 'company': 'CADES', 'code': '3320'},
        'CA Design Raleigh': {'type': 'division', 'company': 'CADES', 'code': '3350'},
        'Corporate': {'type': 'division', 'company': 'DWIRE', 'code': '6000'},
        'Kent': {'type': 'division', 'company': 'DWIRE', 'code': '6003'},
        'Irwindale': {'type': 'division', 'company': 'DWIRE', 'code': '6004'},
        'Neo Alabama': {'type': 'company', 'code': 'NEOAL'},
        'Neo Indiana': {'type': 'company', 'code': 'NEOIN'},
        'Neo Kentucky': {'type': 'company', 'code': 'NEOKY'},
        'Neo Knoxville': {'type': 'company', 'code': 'NEOTN'},
        'Neo Weirton': {'type': 'company', 'code': 'NEOWV'},
        'Wakefield Thermal': {'type': 'company', 'code': 'WVNH'},
        'Wakefield Midwest': {'type': 'company', 'code': 'WVWI'}
    }

    description_totals = {}
    for desc, info in description_source_mapping.items():
        if info['type'] == 'division':
            filtered_df = df_invoice_detail[
                (df_invoice_detail['Company'] == info['company']) &
                (df_invoice_detail['Division'].astype(str) == str(info['code']))
            ].dropna(subset=['Monthly Premium'])
        else:
            filtered_df = df_invoice_detail[
                df_invoice_detail['Company'] == info['code']
            ].dropna(subset=['Monthly Premium'])
        description_totals[desc] = filtered_df['Monthly Premium'].sum()

    for desc, total in description_totals.items():
        rows_to_update = df_template[df_template['DESC'].str.contains(desc, case=False, na=False)].index
        df_template.loc[rows_to_update, 'NET'] = total

    df_hhi_thc = df_invoice_detail[df_invoice_detail['Company'].isin(['HHI', 'THC'])].copy()
    df_hhi_thc['Department'] = df_hhi_thc['Department'].astype(str)
    df_hhi_thc['CC_Code'] = df_hhi_thc['Department'].str[-4:]
    df_cc_totals = df_hhi_thc.groupby('CC_Code')['Monthly Premium'].sum().reset_index()

    df_template['CC'] = df_template['CC'].astype(str).str.replace(r'\.0$', '', regex=True)
    df_cc_totals['CC_Code'] = df_cc_totals['CC_Code'].astype(str)

    cc_merged_df = pd.merge(df_template, df_cc_totals, left_on='CC', right_on='CC_Code', how='left')
    match_indices = cc_merged_df[cc_merged_df['Monthly Premium'].notna()].index
    df_template.loc[match_indices, 'NET'] = cc_merged_df.loc[match_indices, 'Monthly Premium']

    df_template['CC'] = df_template['CC'].replace('nan', '', regex=False)

    # Calculate total of 'NET' column in the updated template
    total_net_template = df_template['NET'].sum()

    # Calculate the total of 'Monthly Premium' in the 'Detail' sheet
    total_invoice_premium = df_invoice_detail['Monthly Premium'].sum()


    # Display the totals for debugging
    st.write(f"Total amount in template: {total_net_template:.2f}")
    st.write(f"Total amount in invoice: {total_invoice_premium:.2f}")


    # Compare the totals and display a message
    if abs(total_net_template - total_invoice_premium) < 0.01: # Use a small tolerance for floating point comparison
        st.success("Processing complete! No Errors Found.")
    else:
        st.warning(f"Processing complete, but the total of the template NET amounts ({total_net_template:.2f}) does not match the total invoice monthly premium amount ({total_invoice_premium:.2f}). Please review the output.")


    output = io.BytesIO()
    df_template.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)

    st.download_button(
        label="Download Updated Medius Template",
        data=output,
        file_name="Complete_Medius_Template.xlsx",
        mime="application/vnd.openxmlformats-owizardssheet"
    )
