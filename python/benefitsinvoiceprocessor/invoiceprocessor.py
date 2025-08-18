import streamlit as st
import pandas as pd
import io

st.title("Aflac Invoice Processor")

invoice_file = st.file_uploader("Upload Aflac Invoice Excel File", type=["xlsx"])
template_file = st.file_uploader("Upload Medius Template Excel File", type=["xlsx"])

if invoice_file and template_file:
    try:
        # Load invoice and template data
        df_invoice = pd.read_excel(invoice_file, sheet_name='Detail', engine='openpyxl')
        df_template = pd.read_excel(template_file, sheet_name=0, engine='openpyxl')
        df_code_map = pd.read_excel(template_file, sheet_name='Code Map', engine='openpyxl')

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

        # --- HHI and THC Department Mapping using stripped Department code ---
        df_hhi_thc = df_invoice[df_invoice['Company'].isin(['HHI', 'THC'])].copy()
        df_hhi_thc['Department'] = df_hhi_thc['Department'].astype(str).str.strip()

        def strip_prefix(dept, company):
            dept = str(dept).strip()
            if company == 'HHI' and dept.startswith('10'):
                return dept[2:]
            elif company == 'THC' and dept.startswith('11'):
                return dept[2:]
            return dept

        df_hhi_thc['CC_Code'] = df_hhi_thc.apply(lambda row: strip_prefix(row['Department'], row['Company']), axis=1)
        df_cc_totals = df_hhi_thc.groupby('CC_Code')['Monthly Premium'].sum().reset_index()

        df_template['CC'] = df_template['CC'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        cc_merged_df = pd.merge(df_template, df_cc_totals, left_on='CC', right_on='CC_Code', how='left')
        match_indices = cc_merged_df[cc_merged_df['Monthly Premium'].notna()].index
        df_template.loc[match_indices, 'NET'] = cc_merged_df.loc[match_indices, 'Monthly Premium']

        # --- Final Output for Medius Template ---
        df_template['CC'] = df_template['CC'].replace('nan', '', regex=False)

        output_template = io.BytesIO()
        df_template.to_excel(output_template, index=False, engine='openpyxl')
        output_template.seek(0)

        # --- Create Aflac Invoice and Support File ---
        df_support = company_totals.rename(columns={'Company': 'Row Labels', 'Monthly Premium': 'Sum of Monthly Premium'})
        df_support['Full Company Name'] = df_support['Row Labels'].map(
            dict(zip(
                df_code_map['Invoice Company Code'].astype(str).str.upper(),
                df_code_map['Company Description'].astype(str).str.strip()
            ))
        )

        output_support = io.BytesIO()
        df_support.to_excel(output_support, index=False, engine='openpyxl')
        output_support.seek(0)

        # --- Streamlit Outputs ---
        st.success("Processing complete!")

        st.download_button(
            label="Download Updated Medius Template",
            data=output_template,
            file_name="Complete_Aflac_Medius_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="Download Aflac Invoice and Support",
            data=output_support,
            file_name="Aflac_Invoice_and_Support.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")


