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
        df_invoice['Division'] = df_invoice['Division'].astype(str).str.strip().replace('', pd.NA)
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

        # Normalize Code Map
        df_code_map['Invoice Company Code'] = df_code_map['Invoice Company Code'].astype(str).str.strip().str.upper()
        df_code_map['Division Code'] = df_code_map['Division Code'].astype(str).str.strip().replace('', pd.NA)

        # Inter-Co Mapping
        interco_map = df_code_map[df_code_map['Division Code'].notna()].drop_duplicates(subset=['Invoice Company Code', 'Division Code'])
        interco_map = interco_map.set_index(['Invoice Company Code', 'Division Code'])['Template Inter-Co'].astype(str).str.strip().to_dict()
        fallback_interco_map = df_code_map[df_code_map['Division Code'].isna()].set_index('Invoice Company Code')['Template Inter-Co'].astype(str).str.strip().to_dict()

        def get_interco(row):
            if pd.isna(row['Division']) or str(row['Division']).strip() == '':
                return fallback_interco_map.get(row['Company'], '')
            key = (row['Company'], row['Division'])
            return interco_map.get(key, fallback_interco_map.get(row['Company'], ''))

        df_invoice['Inter-Co'] = df_invoice.apply(get_interco, axis=1)

        # DESC Mapping
        desc_map_full = df_code_map[df_code_map['Division Code'].notna()].drop_duplicates(subset=['Invoice Company Code', 'Division Code'])
        desc_map_full = desc_map_full.set_index(['Invoice Company Code', 'Division Code'])['Template Desc'].astype(str).str.strip().to_dict()
        desc_map_fallback = df_code_map[df_code_map['Division Code'].isna()].set_index('Invoice Company Code')['Template Desc'].astype(str).str.strip().to_dict()

        def get_desc(row):
            if pd.isna(row['Division']) or str(row['Division']).strip() == '':
                return desc_map_fallback.get(row['Company'], '')
            key = (row['Company'], row['Division'])
            return desc_map_full.get(key, desc_map_fallback.get(row['Company'], ''))

        df_invoice['DESC'] = df_invoice.apply(get_desc, axis=1)
        df_invoice['DESC'] = df_invoice['DESC'].fillna('').astype(str)

        # Remove rows with missing Inter-Co
        df_invoice = df_invoice[(df_invoice['Inter-Co'] != '') & (df_invoice['Inter-Co'].notna())]

        # Add Approver
        df_invoice['Approver'] = approver_name

        # Aggregate totals by DESC, Inter-Co, CC, G/L ACCT, Approver
        df_aggregated = df_invoice.groupby(
            ['DESC', 'Inter-Co', 'CC', 'G/L ACCT', 'Approver'], dropna=False
        )['Monthly Premium'].sum().reset_index()

        # Rename Monthly Premium to NET
        df_aggregated.rename(columns={'Monthly Premium': 'NET'}, inplace=True)

        # Append aggregated rows to the template
        df_result = pd.concat([df_template, df_aggregated], ignore_index=True)

        # Export to Excel
        output = io.BytesIO()
        df_result.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        
        # Debug: Show rows with missing Division
        st.subheader("🔍 Debug: Rows with Missing Division")
        missing_div = df_invoice[df_invoice['Division'].isna()]
        st.write("Rows with missing Division:", len(missing_div))
        st.dataframe(missing_div[['Company', 'Division', 'Inter-Co', 'DESC']].head(10))
        
        # Debug: Show rows with fallback Inter-Co
        st.subheader("🔍 Debug: Fallback Inter-Co Applied")
        fallback_interco_rows = df_invoice[df_invoice['Division'].isna() | ~df_invoice.set_index(['Company', 'Division']).index.isin(interco_map.keys())]
        st.write("Rows using fallback Inter-Co:", len(fallback_interco_rows))
        st.dataframe(fallback_interco_rows[['Company', 'Division', 'Inter-Co']].head(10))
        
        # Debug: Show rows with fallback DESC
        st.subheader("🔍 Debug: Fallback DESC Applied")
        fallback_desc_rows = df_invoice[df_invoice['Division'].isna() | ~df_invoice.set_index(['Company', 'Division']).index.isin(desc_map_full.keys())]
        st.write("Rows using fallback DESC:", len(fallback_desc_rows))
        st.dataframe(fallback_desc_rows[['Company', 'Division', 'DESC']].head(10))
        
        # Debug: Show rows that will be dropped
        st.subheader("⚠️ Debug: Rows Missing Inter-Co (Will Be Dropped)")
        missing_interco = df_invoice[(df_invoice['Inter-Co'] == '') | (df_invoice['Inter-Co'].isna())]
        st.write("Rows missing Inter-Co:", len(missing_interco))
        st.dataframe(missing_interco[['Company', 'Division', 'Inter-Co', 'DESC']].head(10))
        
        st.success("Template updated with aggregated invoice data!")
        st.download_button(
            label="Download Updated Medius Template",
            data=output,
            file_name="Updated_Aflac_Medius_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"An error occurred: {e}")

