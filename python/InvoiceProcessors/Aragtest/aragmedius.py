import streamlit as st
import pandas as pd
import io

st.title("Aflac Medius Template Generator")

invoice_file = st.file_uploader("Upload Aflac Invoice Excel File", type=["xlsx"])
template_file = st.file_uploader("Upload Medius Template Excel File", type=["xlsx"])
approver_name = st.text_input("Enter Approver Name")

if invoice_file and template_file and approver_name:
    try:
        # Load invoice and template data
        df_invoice = pd.read_excel(invoice_file, sheet_name='Detail', engine='openpyxl')
        df_code_map = pd.read_excel(template_file, sheet_name='Code Map', engine='openpyxl')
        df_gl_acct = pd.read_excel(template_file, sheet_name='GL ACCT', engine='openpyxl')
        df_heico_dept = pd.read_excel(template_file, sheet_name='Heico Departments', engine='openpyxl')

        # Normalize invoice data
        df_invoice['Company'] = df_invoice['Company'].astype(str).str.strip().str.upper()
        df_invoice['Division'] = df_invoice['Division'].astype(str).str.strip()
        df_invoice['Department'] = df_invoice['Department'].astype(str).str.strip()
        df_invoice['Monthly Premium'] = pd.to_numeric(df_invoice['Monthly Premium'], errors='coerce')

        # Determine Group (Heico or Non-Heico)
        df_invoice['Group'] = df_invoice['Company'].apply(lambda x: 'Heico' if x in ['HHI', 'THC'] else 'Non-Heico')

        # Map G/L ACCT based on Group
        gl_map = df_gl_acct.set_index('Group')['G/L ACCT'].to_dict()
        df_invoice['G/L ACCT'] = df_invoice['Group'].map(gl_map)

        # Strip Department prefix and map to Organization Code
        def strip_prefix(dept, company):
            if company == 'HHI' and dept.startswith('10'):
                return dept[2:]
            elif company == 'THC' and dept.startswith('11'):
                return dept[2:]
            return dept

        df_invoice['Stripped Dept'] = df_invoice.apply(lambda row: strip_prefix(row['Department'], row['Company']), axis=1)
        dept_map = df_heico_dept.set_index('Department')['Organization Code'].to_dict()
        df_invoice['CC'] = df_invoice['Stripped Dept'].map(dept_map)

        # Map Inter-Co from Code Map
        interco_map = df_code_map.set_index('Invoice Company Code')['Template Inter-Co'].to_dict()
        df_invoice['Inter-Co'] = df_invoice['Company'].map(interco_map)

        # Map DESC from Code Map using Division Code
        desc_map_div = df_code_map[df_code_map['Division Code'].notna()]
        desc_map_div = desc_map_div.set_index('Division Code')['Template Desc'].to_dict()
        df_invoice['DESC'] = df_invoice['Division'].map(desc_map_div)

        # Fill DESC from Code Map using Invoice Company Code if Division Code is not available
        desc_map_company = df_code_map[df_code_map['Division Code'].isna()]
        desc_map_company = desc_map_company.set_index('Invoice Company Code')['Template Desc'].to_dict()
        df_invoice['DESC'] = df_invoice.apply(
            lambda row: row['DESC'] if pd.notna(row['DESC']) else desc_map_company.get(row['Company'], ''),
            axis=1
        )

        # Add Approver
        df_invoice['Approver'] = approver_name

        # Aggregate NET by DESC, Inter-Co, CC, G/L ACCT, Approver
        df_template = df_invoice.groupby(
            ['DESC', 'Inter-Co', 'CC', 'G/L ACCT', 'Approver'], dropna=False
        )['Monthly Premium'].sum().reset_index()

        df_template.rename(columns={'Monthly Premium': 'NET'}, inplace=True)

        # Final Output
        output = io.BytesIO()
        df_template.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.success("Template generation complete!")
        st.download_button(
            label="Download Generated Medius Template",
            data=output,
            file_name="Generated_Aflac_Medius_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
