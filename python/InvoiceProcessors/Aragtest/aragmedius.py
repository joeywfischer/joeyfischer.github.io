import streamlit as st
import pandas as pd
import io

st.title("Aflac Medius Template Updater")

invoice_file = st.file_uploader("Upload Aflac Invoice Excel File", type=["xlsx"])
template_file = st.file_uploader("Upload Medius Template Excel File", type=["xlsx"])
approver_name = st.text_input("Enter Approver Name")

if invoice_file and template_file and approver_name:
    try:
        # Load invoice and template data
        df_invoice = pd.read_excel(invoice_file, sheet_name='Detail', engine='openpyxl')
        df_template = pd.read_excel(template_file, sheet_name=0, engine='openpyxl')
        df_code_map = pd.read_excel(template_file, sheet_name='Code Map', engine='openpyxl')
        df_gl_acct = pd.read_excel(template_file, sheet_name='GL ACCT', engine='openpyxl')
        df_heico_dept = pd.read_excel(template_file, sheet_name='Heico Departments', engine='openpyxl')

        # Normalize invoice data
        df_invoice['Company'] = df_invoice['Company'].astype(str).str.strip().str.upper()
        df_invoice['Division'] = df_invoice['Division'].astype(str).str.strip()
        df_invoice['Department'] = df_invoice['Department'].astype(str).str.strip()
        df_invoice['Monthly Premium'] = pd.to_numeric(df_invoice['Monthly Premium'], errors='coerce')
        df_invoice['Group'] = df_invoice['Company'].apply(lambda x: 'Heico' if x in ['HHI', 'THC'] else 'Non-Heico')

        # Mappings
        gl_map = df_gl_acct.set_index('Group')['G/L ACCT'].to_dict()
        dept_map = df_heico_dept.set_index('Department')['Organization Code'].to_dict()
        interco_map = df_code_map.set_index('Invoice Company Code')['Template Inter-Co'].to_dict()
        desc_map_div = df_code_map[df_code_map['Division Code'].notna()].set_index('Division Code')['Template Desc'].to_dict()
        desc_map_company = df_code_map[df_code_map['Division Code'].isna()].set_index('Invoice Company Code')['Template Desc'].to_dict()

        def strip_prefix(dept, company):
            if company == 'HHI' and dept.startswith('10'):
                return dept[2:]
            elif company == 'THC' and dept.startswith('11'):
                return dept[2:]
            return dept

        df_updated = df_template.copy()
        df_updated['DESC'] = df_updated['DESC'].astype(str).str.strip()
        df_updated['Inter-Co'] = df_updated['Inter-Co'].astype(str).str.strip()
        df_updated['CC'] = df_updated['CC'].astype(str).str.strip().str.replace(r'\\.0$', '', regex=True)

        for idx, row in df_updated.iterrows():
            desc = str(row['DESC']).strip()
            interco = str(row['Inter-Co']).strip()
            cc = str(row['CC']).strip()

            matching_rows = df_invoice.copy()

            if desc:
                matching_rows['DESC'] = df_invoice['Division'].map(desc_map_div)
                matching_rows['DESC'] = matching_rows.apply(
                    lambda r: r['DESC'] if pd.notna(r['DESC']) else desc_map_company.get(r['Company'], ''),
                    axis=1
                )
                matching_rows = matching_rows[matching_rows['DESC'].str.lower() == desc.lower()]

            if interco:
                matching_rows['Inter-Co'] = matching_rows['Company'].map(interco_map)
                matching_rows = matching_rows[matching_rows['Inter-Co'] == interco]

            if cc:
                matching_rows['Stripped Dept'] = matching_rows.apply(lambda r: strip_prefix(r['Department'], r['Company']), axis=1)
                matching_rows['CC'] = matching_rows['Stripped Dept'].map(dept_map)
                matching_rows = matching_rows[matching_rows['CC'].astype(str).str.strip() == cc]

            total = matching_rows['Monthly Premium'].sum()
            if total > 0:
                df_updated.at[idx, 'NET'] = total
                df_updated.at[idx, 'Approver'] = approver_name
                group = matching_rows['Group'].iloc[0] if not matching_rows.empty else ''
                df_updated.at[idx, 'G/L ACCT'] = gl_map.get(group, '')

        output = io.BytesIO()
        df_updated.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.success("Template update complete!")
        st.download_button(
            label="Download Updated Medius Template",
            data=output,
            file_name="Updated_Aflac_Medius_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")

