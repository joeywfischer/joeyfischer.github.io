import streamlit as st
import pandas as pd
import io

st.title("Aflac Medius Template Generator")

# Upload files and input approver name
invoice_file = st.file_uploader("Upload Aflac Invoice Excel File", type=["xlsx"])
template_file = st.file_uploader("Upload Medius Template Excel File", type=["xlsx"])
approver_name = st.text_input("Enter Approver Name")

if invoice_file and template_file and approver_name:
    try:
        # Load all necessary sheets
        df_invoice = pd.read_excel(invoice_file, sheet_name='Detail', engine='openpyxl')
        df_code_map = pd.read_excel(template_file, sheet_name='Code Map', engine='openpyxl')
        df_gl_acct = pd.read_excel(template_file, sheet_name='GL ACCT', engine='openpyxl')
        df_heico_dept = pd.read_excel(template_file, sheet_name='Heico Departments', engine='openpyxl')
        df_template = pd.read_excel(template_file, sheet_name='Medius Excel Template', engine='openpyxl')

        # === Normalize Invoice Data ===
        df_invoice['Company'] = df_invoice['Company'].astype(str).str.strip().str.upper()
        df_invoice['Division'] = df_invoice['Division'].apply(lambda x: str(x).strip() if pd.notna(x) else '')
        df_invoice['Department'] = df_invoice['Department'].astype(str).str.strip()
        df_invoice['Monthly Premium'] = pd.to_numeric(df_invoice['Monthly Premium'], errors='coerce')

        # === DEBUG: Check Department Matching ===
        st.subheader("Debug: Department Matching")
        
        # Normalize columns
        df_invoice['Department'] = df_invoice['Department'].astype(str).str.strip()
        df_heico_dept['Department Code'] = df_heico_dept['Department Code'].astype(str).str.strip()
        
        # Show unique values
        unique_invoice_departments = sorted(df_invoice['Department'].unique())
        unique_heico_department_codes = sorted(df_heico_dept['Department Code'].unique())
        
        st.write("Unique 'Department' values from Invoice Sheet:")
        st.write(unique_invoice_departments)
        
        st.write("Unique 'Department Code' values from Heico Departments Sheet:")
        st.write(unique_heico_department_codes)
        
        # Show mismatches
        missing_departments = [dept for dept in unique_invoice_departments if dept not in unique_heico_department_codes]
        if missing_departments:
            st.warning("Departments in invoice sheet not found in Heico Departments sheet:")
            st.write(missing_departments)
        else:
            st.success("All departments from invoice sheet are found in Heico Departments sheet.")

        # Remove THC and HHI temporarily
        df_invoice_filtered = df_invoice[~df_invoice['Company'].isin(['THC', 'HHI'])].copy()
        df_invoice_filtered['Group'] = df_invoice_filtered['Company'].apply(lambda x: 'Heico' if x in ['HHI', 'THC'] else 'Non-Heico')

        # Map G/L ACCT
        gl_map = df_gl_acct.set_index('Group')['G/L ACCT'].to_dict()
        df_invoice_filtered['G/L ACCT'] = df_invoice_filtered['Group'].map(gl_map)

        # Strip prefixes and map department codes
        def strip_prefix(dept, company):
            if company == 'HHI' and dept.startswith('10'):
                return dept[2:]
            elif company == 'THC' and dept.startswith('11'):
                return dept[2:]
            return dept

        df_invoice_filtered['Stripped Dept'] = df_invoice_filtered.apply(
            lambda row: strip_prefix(row['Department'], row['Company']), axis=1
        )
        dept_map = df_heico_dept.set_index('Department')['Department Code'].astype(str).str.strip().to_dict()
        df_invoice_filtered['CC'] = df_invoice_filtered['Stripped Dept'].map(dept_map)

        # Prepare Code Map
        df_code_map['Division Code'] = df_code_map['Division Code'].apply(lambda x: str(x).strip() if pd.notna(x) else None)
        df_code_map_str = df_code_map[df_code_map['Division Code'].apply(lambda x: not x.isdigit() if isinstance(x, str) else False)]
        df_code_map_num = df_code_map[df_code_map['Division Code'].apply(lambda x: x.isdigit() if isinstance(x, str) else False)]

        desc_map_str = df_code_map_str.set_index('Division Code')['Template Desc'].astype(str).str.strip().to_dict()
        interco_map_str = df_code_map_str.set_index('Division Code')['Template Inter-Co'].astype(str).str.strip().to_dict()
        desc_map_num = df_code_map_num.set_index('Division Code')['Template Desc'].astype(str).str.strip().to_dict()
        interco_map_num = df_code_map_num.set_index('Division Code')['Template Inter-Co'].astype(str).str.strip().to_dict()

        def map_desc(row):
            return desc_map_num.get(row['Division'], '') if row['Division'].isdigit() else desc_map_str.get(row['Division'], '')

        def map_interco(row):
            return interco_map_num.get(row['Division'], '') if row['Division'].isdigit() else interco_map_str.get(row['Division'], '')

        df_invoice_filtered['DESC'] = df_invoice_filtered.apply(map_desc, axis=1)
        df_invoice_filtered['Inter-Co'] = df_invoice_filtered.apply(map_interco, axis=1)

        # Fallback using Invoice Company Code
        fallback_desc_map = df_code_map[df_code_map['Division Code'].isna()].set_index('Invoice Company Code')['Template Desc'].astype(str).str.strip().to_dict()
        fallback_interco_map = df_code_map[df_code_map['Division Code'].isna()].set_index('Invoice Company Code')['Template Inter-Co'].astype(str).str.strip().to_dict()

        df_invoice_filtered['DESC'] = df_invoice_filtered.apply(
            lambda row: row['DESC'] if row['DESC'] else fallback_desc_map.get(row['Company'], ''),
            axis=1
        )
        df_invoice_filtered['Inter-Co'] = df_invoice_filtered.apply(
            lambda row: row['Inter-Co'] if row['Inter-Co'] else fallback_interco_map.get(row['Company'], ''),
            axis=1
        )

        df_invoice_filtered['Approver'] = approver_name

        # === Aggregate Non-HHI/THC Data ===
        df_aggregated = df_invoice_filtered.groupby(
            ['DESC', 'Inter-Co', 'CC', 'G/L ACCT', 'Approver'], dropna=False
        )['Monthly Premium'].sum().reset_index().rename(columns={'Monthly Premium': 'NET'})

        df_aggregated = df_aggregated[df_aggregated['Inter-Co'].notna() & (df_aggregated['Inter-Co'].str.strip() != '')]

       # === Handle HHI and THC Data ===
        df_hhi_thc = df_invoice[df_invoice['Company'].isin(['HHI', 'THC'])].copy()
        df_hhi_thc['Monthly Premium'] = pd.to_numeric(df_hhi_thc['Monthly Premium'], errors='coerce')
        df_hhi_thc['Department'] = df_hhi_thc['Department'].astype(str).str.strip()
        
        # Normalize Heico Departments sheet
        df_heico_dept['Department Code'] = df_heico_dept['Department Code'].astype(str).str.strip()
        df_heico_dept['Template Code'] = df_heico_dept['Template Code'].astype(str).str.strip()
        
        # Filter only departments that exist in Heico Departments sheet
        df_hhi_thc = df_hhi_thc[df_hhi_thc['Department'].isin(df_heico_dept['Department Code'])]
        
        # Debug: Show matched departments
        st.subheader("Debug: Matched HHI/THC Departments")
        st.write("Matched Departments from Invoice Sheet:")
        st.write(df_hhi_thc['Department'].unique())
        
        # Aggregate by Department
        df_dept_sum = df_hhi_thc.groupby('Department')['Monthly Premium'].sum().reset_index()
        
        # Map to DESC and CC using Heico Departments sheet
        dept_lookup = df_heico_dept.set_index('Department Code')[['Department', 'Template Code']].dropna()
        
        df_dept_sum['DESC'] = df_dept_sum['Department']
        df_dept_sum['CC'] = df_dept_sum['Department'].map(dept_lookup['Template Code'])
        df_dept_sum['G/L ACCT'] = df_gl_acct[df_gl_acct['Group'] == 'Heico']['G/L ACCT'].values[0]
        df_dept_sum['Inter-Co'] = 'HEICO'
        df_dept_sum['Approver'] = approver_name
        df_dept_sum.rename(columns={'Monthly Premium': 'NET'}, inplace=True)

df_dept_sum = df_dept_sum[['DESC', 'Inter-Co', 'CC', 'G/L ACCT', 'Approver', 'NET']]

        # === Combine and Export ===
        df_result = pd.concat([df_template, df_aggregated, df_dept_sum], ignore_index=True)
        df_result = df_result.sort_values(by='Inter-Co', ascending=True)

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

