import streamlit as st
import pandas as pd
import io

st.title("Aflac Medius Template Generator")

invoice_file = st.file_uploader("Upload Aflac Invoice Excel File", type=["xlsx"])
template_file = st.file_uploader("Upload Medius Template Excel File", type=["xlsx"])
approver_name = st.text_input("Enter Approver Name")

if invoice_file and template_file and approver_name:
    try:
        # Load sheets
        df_invoice = pd.read_excel(invoice_file, sheet_name='Detail', engine='openpyxl')
        df_code_map = pd.read_excel(template_file, sheet_name='Code Map', engine='openpyxl')
        df_gl_acct = pd.read_excel(template_file, sheet_name='GL ACCT', engine='openpyxl')
        df_heico_dept = pd.read_excel(template_file, sheet_name='Heico Departments', engine='openpyxl')
        df_template = pd.read_excel(template_file, sheet_name='Medius Excel Template', engine='openpyxl')

        # Normalize invoice data
        df_invoice['Company'] = df_invoice['Company'].astype(str).str.strip().str.upper()
        df_invoice['Division'] = df_invoice['Division'].apply(lambda x: str(x).strip() if pd.notna(x) else '')
        df_invoice['Department'] = df_invoice['Department'].astype(str).str.strip()
        df_invoice['Monthly Premium'] = pd.to_numeric(df_invoice['Monthly Premium'], errors='coerce')

        # === Non-Heico Aggregation ===
        df_non_heico = df_invoice[~df_invoice['Company'].isin(['THC', 'HHI'])].copy()
        df_non_heico['Group'] = 'Non-Heico'
        df_non_heico['G/L ACCT'] = df_non_heico['Group'].map(df_gl_acct.set_index('Group')['G/L ACCT'].to_dict())

        dept_map = df_heico_dept.set_index('Department')['Department Code'].astype(str).str.strip().to_dict()
        df_non_heico['Stripped Dept'] = df_non_heico['Department']
        df_non_heico['CC'] = df_non_heico['Stripped Dept'].map(dept_map)

        df_code_map['Division Code'] = df_code_map['Division Code'].apply(lambda x: str(x).strip() if pd.notna(x) else None)
        desc_map = df_code_map.set_index('Division Code')['Template Desc'].astype(str).str.strip().to_dict()
        interco_map = df_code_map.set_index('Division Code')['Template Inter-Co'].astype(str).str.strip().to_dict()

        df_non_heico['DESC'] = df_non_heico['Division'].map(desc_map)
        df_non_heico['Inter-Co'] = df_non_heico['Division'].map(interco_map)

        fallback_desc = df_code_map[df_code_map['Division Code'].isna()].set_index('Invoice Company Code')['Template Desc'].astype(str).str.strip().to_dict()
        fallback_interco = df_code_map[df_code_map['Division Code'].isna()].set_index('Invoice Company Code')['Template Inter-Co'].astype(str).str.strip().to_dict()

        df_non_heico['DESC'] = df_non_heico.apply(lambda row: row['DESC'] if row['DESC'] else fallback_desc.get(row['Company'], ''), axis=1)
        df_non_heico['Inter-Co'] = df_non_heico.apply(lambda row: row['Inter-Co'] if row['Inter-Co'] else fallback_interco.get(row['Company'], ''), axis=1)
        df_non_heico['Approver'] = approver_name

        df_agg = df_non_heico.groupby(['DESC', 'Inter-Co', 'CC', 'G/L ACCT', 'Approver'], dropna=False)['Monthly Premium'].sum().reset_index()
        df_agg.rename(columns={'Monthly Premium': 'NET'}, inplace=True)

        # === Heico (HHI/THC) Aggregation using merge ===
        df_hhi_thc = df_invoice[df_invoice['Company'].isin(['HHI', 'THC'])].copy()
        df_hhi_thc['Monthly Premium'] = pd.to_numeric(df_hhi_thc['Monthly Premium'], errors='coerce')

        df_dept_sum = df_hhi_thc.groupby('Department')['Monthly Premium'].sum().reset_index()

        # Normalize both sides to string
        df_dept_sum['Department'] = df_dept_sum['Department'].astype(str).str.strip()
        df_heico_dept['Department Code'] = df_heico_dept['Department Code'].astype(str).str.strip()

        # Merge instead of map
        df_merged = pd.merge(
            df_dept_sum,
            df_heico_dept[['Department Code', 'Department', 'Template Code']],
            left_on='Department',
            right_on='Department Code',
            how='left'
        )

        df_merged['G/L ACCT'] = df_gl_acct[df_gl_acct['Group'] == 'Heico']['G/L ACCT'].values[0]
        df_merged['Inter-Co'] = 'HEICO'
        df_merged['Approver'] = approver_name
        df_merged.rename(columns={'Monthly Premium': 'NET', 'Department_y': 'DESC', 'Template Code': 'CC'}, inplace=True)

        df_final = df_merged[['DESC', 'Inter-Co', 'CC', 'G/L ACCT', 'Approver', 'NET']]

        # === Export to Excel ===
        df_result = pd.concat([df_template, df_agg], ignore_index=True).sort_values(by='Inter-Co')

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_result.to_excel(writer, sheet_name='Updated Template', index=False)
            df_final.to_excel(writer, sheet_name='HHI_THC Aggregation', index=False)
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

