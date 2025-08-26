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
        df_template = pd.read_excel(template_file, sheet_name='Medius Excel Template', engine='openpyxl')

        # Normalize invoice data
        df_invoice['Company'] = df_invoice['Company'].astype(str).str.strip().str.upper()
        df_invoice['Division'] = df_invoice['Division'].apply(lambda x: str(x).strip() if pd.notna(x) else '')
        df_invoice['Department'] = df_invoice['Department'].astype(str).str.strip()
        df_invoice['Monthly Premium'] = pd.to_numeric(df_invoice['Monthly Premium'], errors='coerce')

        # Remove THC and HHI for now
        df_invoice = df_invoice[~df_invoice['Company'].isin(['THC', 'HHI'])]

        # Determine Group (Heico or Non-Heico)
        df_invoice['Group'] = df_invoice['Company'].apply(lambda x: 'Heico' if x in ['HHI', 'THC'] else 'Non-Heico')

        # Map G/L ACCT based on Group
        gl_map = df_gl_acct.set_index('Group')['G/L ACCT'].to_dict()
        df_invoice['G/L ACCT'] = df_invoice['Group'].map(gl_map)

        # Strip Department prefix and map to Department Code
        def strip_prefix(dept, company):
            if company == 'HHI' and dept.startswith('10'):
                return dept[2:]
            elif company == 'THC' and dept.startswith('11'):
                return dept[2:]
            return dept

        df_invoice['Stripped Dept'] = df_invoice.apply(lambda row: strip_prefix(row['Department'], row['Company']), axis=1)
        dept_map = df_heico_dept.set_index('Department')['Department Code'].astype(str).str.strip().to_dict()
        df_invoice['CC'] = df_invoice['Stripped Dept'].map(dept_map)

        # Prepare Code Map
        df_code_map['Division Code'] = df_code_map['Division Code'].apply(lambda x: str(x).strip() if pd.notna(x) else None)

        # Separate string and numeric division codes
        df_code_map_str = df_code_map[df_code_map['Division Code'].apply(lambda x: not x.isdigit() if isinstance(x, str) else False)]
        df_code_map_num = df_code_map[df_code_map['Division Code'].apply(lambda x: x.isdigit() if isinstance(x, str) else False)]

        # Create mapping dictionaries
        desc_map_str = df_code_map_str.set_index('Division Code')['Template Desc'].astype(str).str.strip().to_dict()
        interco_map_str = df_code_map_str.set_index('Division Code')['Template Inter-Co'].astype(str).str.strip().to_dict()

        desc_map_num = df_code_map_num.set_index('Division Code')['Template Desc'].astype(str).str.strip().to_dict()
        interco_map_num = df_code_map_num.set_index('Division Code')['Template Inter-Co'].astype(str).str.strip().to_dict()

        # Apply mapping based on type of Division
        def map_desc(row):
            division = row['Division']
            if division.isdigit():
                return desc_map_num.get(division, '')
            else:
                return desc_map_str.get(division, '')

        def map_interco(row):
            division = row['Division']
            if division.isdigit():
                return interco_map_num.get(division, '')
            else:
                return interco_map_str.get(division, '')

        df_invoice['DESC'] = df_invoice.apply(map_desc, axis=1)
        df_invoice['Inter-Co'] = df_invoice.apply(map_interco, axis=1)

        # Fallback DESC and Inter-Co using Invoice Company Code if Division Code is missing
        fallback_desc_map = df_code_map[df_code_map['Division Code'].isna()].set_index('Invoice Company Code')['Template Desc'].astype(str).str.strip().to_dict()
        fallback_interco_map = df_code_map[df_code_map['Division Code'].isna()].set_index('Invoice Company Code')['Template Inter-Co'].astype(str).str.strip().to_dict()

        df_invoice['DESC'] = df_invoice.apply(
            lambda row: row['DESC'] if row['DESC'] else fallback_desc_map.get(row['Company'], ''),
            axis=1
        )
        df_invoice['Inter-Co'] = df_invoice.apply(
            lambda row: row['Inter-Co'] if row['Inter-Co'] else fallback_interco_map.get(row['Company'], ''),
            axis=1
        )

        # Replace actual NaN values in DESC with empty strings
        df_invoice['DESC'] = df_invoice['DESC'].fillna('').astype(str)

        # Add Approver
        df_invoice['Approver'] = approver_name

        # Aggregate totals by DESC, Inter-Co, CC, G/L ACCT, Approver
        df_aggregated = df_invoice.groupby(
            ['DESC', 'Inter-Co', 'CC', 'G/L ACCT', 'Approver'], dropna=False
        )['Monthly Premium'].sum().reset_index()

        # Rename Monthly Premium to NET
        df_aggregated.rename(columns={'Monthly Premium': 'NET'}, inplace=True)

        # Remove rows with missing Inter-Co
        df_aggregated = df_aggregated[df_aggregated['Inter-Co'].notna() & (df_aggregated['Inter-Co'].str.strip() != '')]

        # Append aggregated rows to the template
        df_result = pd.concat([df_template, df_aggregated], ignore_index=True)

        # Export to Excel
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
