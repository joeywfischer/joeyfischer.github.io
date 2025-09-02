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

        # Normalize columns
        df_invoice['Company'] = df_invoice['Company'].astype(str).str.strip().str.upper()
        df_invoice['Division'] = df_invoice['Division'].apply(lambda x: str(x).strip() if pd.notna(x) else '')
        df_invoice['Department'] = df_invoice['Department'].astype(str).str.strip()
        df_invoice['Monthly Premium'] = pd.to_numeric(df_invoice['Monthly Premium'], errors='coerce')
        df_heico_dept['Department Code'] = df_heico_dept['Department Code'].astype(str).str.strip()
        df_heico_dept['Department Name'] = df_heico_dept['Department Name'].astype(str).str.strip()
        df_heico_dept['Template Code'] = df_heico_dept['Template Code'].astype(str).str.strip()

        # === Filter Non-HHI/THC ===
        df_non_heico = df_invoice[~df_invoice['Company'].isin(['THC', 'HHI'])].copy()
        df_non_heico['Group'] = 'Non-Heico'
        gl_map = df_gl_acct.set_index('Group')['G/L ACCT'].to_dict()
        df_non_heico['G/L ACCT'] = df_non_heico['Group'].map(gl_map)

        # Map Department to CC
        dept_map = df_heico_dept.set_index('Department Code')['Template Code'].to_dict()
        df_non_heico['CC'] = df_non_heico['Department'].map(dept_map)

        # Map DESC and Inter-Co
        df_code_map['Division Code'] = df_code_map['Division Code'].apply(lambda x: str(x).strip() if pd.notna(x) else None)
        desc_map = df_code_map.set_index('Division Code')['Template Desc'].astype(str).str.strip().to_dict()
        interco_map = df_code_map.set_index('Division Code')['Template Inter-Co'].astype(str).str.strip().to_dict()

        fallback_desc_map = df_code_map[df_code_map['Division Code'].isna()].set_index('Invoice Company Code')['Template Desc'].astype(str).str.strip().to_dict()
        fallback_interco_map = df_code_map[df_code_map['Division Code'].isna()].set_index('Invoice Company Code')['Template Inter-Co'].astype(str).str.strip().to_dict()

        def map_desc(row):
            return desc_map.get(row['Division'], fallback_desc_map.get(row['Company'], ''))

        def map_interco(row):
            return interco_map.get(row['Division'], fallback_interco_map.get(row['Company'], ''))

        df_non_heico['DESC'] = df_non_heico.apply(map_desc, axis=1)
        df_non_heico['Inter-Co'] = df_non_heico.apply(map_interco, axis=1)
        df_non_heico['Approver'] = approver_name

        # Aggregate
        df_aggregated = df_non_heico.groupby(
            ['DESC', 'Inter-Co', 'CC', 'G/L ACCT', 'Approver'], dropna=False
        )['Monthly Premium'].sum().reset_index().rename(columns={'Monthly Premium': 'NET'})

        # === Clean Template Before Appending ===
        df_template = df_template[df_template['Inter-Co'].notna() & (df_template['Inter-Co'].str.strip() != '')]
        df_aggregated['DESC'] = df_aggregated['DESC'].fillna('').apply(lambda x: '' if str(x).lower() == 'nan' else str(x))
        
        # === Handle HHI/THC ===
        df_heico = df_invoice[df_invoice['Company'].isin(['HHI', 'THC'])].copy()
        df_heico['Monthly Premium'] = pd.to_numeric(df_heico['Monthly Premium'], errors='coerce')
        df_heico['Department'] = df_heico['Department'].astype(str).str.strip()

        df_dept_sum = df_heico.groupby('Department')['Monthly Premium'].sum().reset_index()
        df_dept_sum['Department Code'] = df_dept_sum['Department']
        df_dept_sum = df_dept_sum[df_dept_sum['Department Code'].isin(df_heico_dept['Department Code'])]

        # Map DESC from Department Name
        dept_name_map = df_heico_dept.set_index('Department Code')['Department Name'].to_dict()
        df_dept_sum['DESC'] = df_dept_sum['Department Code'].map(dept_name_map)
        df_dept_sum['CC'] = df_dept_sum['Department Code'].map(dept_map)
        df_dept_sum['G/L ACCT'] = gl_map.get('Heico', '')
        df_dept_sum['Inter-Co'] = ''  # Clear Inter-Co for HHI/THC
        df_dept_sum['Approver'] = approver_name
        df_dept_sum.rename(columns={'Monthly Premium': 'NET'}, inplace=True)

        df_dept_sum = df_dept_sum[['DESC', 'Inter-Co', 'CC', 'G/L ACCT', 'Approver', 'NET']]

        # === Combine and Export ===
        df_result = pd.concat([df_template, df_aggregated, df_dept_sum], ignore_index=True)
        df_result = df_result.sort_values(by='Inter-Co', ascending=True)
        df_result = df_result[~(df_result['Inter-Co'].fillna('').str.strip() == '') & ~(df_result['CC'].fillna('').str.strip() == '')]

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

