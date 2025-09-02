import streamlit as st
import pandas as pd
import io

st.title("Aflac Medius Template Generator")

# === Upload Inputs ===
invoice_file = st.file_uploader("Upload Aflac Invoice Excel File", type=["xlsx"])
template_file = st.file_uploader("Upload Medius Template Excel File", type=["xlsx"])
approver_name = st.text_input("Enter Approver Name")

# === Helper Functions ===
def load_sheets(invoice_file, template_file):
    invoice = pd.read_excel(invoice_file, sheet_name='Detail', engine='openpyxl')
    code_map = pd.read_excel(template_file, sheet_name='Code Map', engine='openpyxl')
    gl_acct = pd.read_excel(template_file, sheet_name='GL ACCT', engine='openpyxl')
    heico_dept = pd.read_excel(template_file, sheet_name='Heico Departments', engine='openpyxl')
    template = pd.read_excel(template_file, sheet_name='Medius Excel Template', engine='openpyxl')
    return invoice, code_map, gl_acct, heico_dept, template

def preprocess_invoice(df):
    df['Company'] = df['Company'].astype(str).str.strip().str.upper()
    df['Division'] = df['Division'].apply(lambda x: str(x).strip() if pd.notna(x) else '')
    df['Department'] = pd.to_numeric(df['Department'], errors='coerce').astype('Int64')
    df['Monthly Premium'] = pd.to_numeric(df['Monthly Premium'], errors='coerce')
    return df

def aggregate_non_heico(df, code_map, gl_acct, heico_dept, approver_name):
    df = df[~df['Company'].isin(['THC', 'HHI'])].copy()
    df['Group'] = 'Non-Heico'
    df['G/L ACCT'] = df['Group'].map(gl_acct.set_index('Group')['G/L ACCT'].to_dict())

    dept_map = heico_dept.set_index('Department')['Department Code'].astype(str).str.strip().to_dict()
    df['Stripped Dept'] = df['Department'].astype(str)
    df['CC'] = df['Stripped Dept'].map(dept_map)

    code_map['Division Code'] = code_map['Division Code'].apply(lambda x: str(x).strip() if pd.notna(x) else None)
    desc_map = code_map.set_index('Division Code')['Template Desc'].astype(str).str.strip().to_dict()
    interco_map = code_map.set_index('Division Code')['Template Inter-Co'].astype(str).str.strip().to_dict()

    df['DESC'] = df['Division'].map(desc_map)
    df['Inter-Co'] = df['Division'].map(interco_map)

    fallback_desc = code_map[code_map['Division Code'].isna()].set_index('Invoice Company Code')['Template Desc'].astype(str).str.strip().to_dict()
    fallback_interco = code_map[code_map['Division Code'].isna()].set_index('Invoice Company Code')['Template Inter-Co'].astype(str).str.strip().to_dict()

    df['DESC'] = df.apply(lambda row: row['DESC'] if row['DESC'] else fallback_desc.get(row['Company'], ''), axis=1)
    df['Inter-Co'] = df.apply(lambda row: row['Inter-Co'] if row['Inter-Co'] else fallback_interco.get(row['Company'], ''), axis=1)
    df['Approver'] = approver_name

    df_agg = df.groupby(['DESC', 'Inter-Co', 'CC', 'G/L ACCT', 'Approver'], dropna=False)['Monthly Premium'].sum().reset_index()
    df_agg.rename(columns={'Monthly Premium': 'NET'}, inplace=True)
    return df_agg[df_agg['Inter-Co'].notna() & (df_agg['Inter-Co'].str.strip() != '')]

def aggregate_heico(df, heico_dept, gl_acct, approver_name):
    df = df[df['Company'].isin(['HHI', 'THC'])].copy()
    df['Department'] = pd.to_numeric(df['Department'], errors='coerce').astype('Int64')
    df['Monthly Premium'] = pd.to_numeric(df['Monthly Premium'], errors='coerce')

    df_sum = df.groupby('Department')['Monthly Premium'].sum().reset_index()

    heico_dept['Department Code'] = pd.to_numeric(heico_dept['Department Code'], errors='coerce').astype('Int64')
    lookup = heico_dept.set_index('Department Code')[['Department', 'Template Code']].dropna()

    df_sum['DESC'] = df_sum['Department'].map(lookup['Department'])
    df_sum['CC'] = df_sum['Department'].map(lookup['Template Code'])
    df_sum['G/L ACCT'] = gl_acct[gl_acct['Group'] == 'Heico']['G/L ACCT'].values[0]
    df_sum['Inter-Co'] = 'HEICO'
    df_sum['Approver'] = approver_name
    df_sum.rename(columns={'Monthly Premium': 'NET'}, inplace=True)

    return df_sum[['DESC', 'Inter-Co', 'CC', 'G/L ACCT', 'Approver', 'NET']]

# === Main Logic ===
if invoice_file and template_file and approver_name:
    try:
        df_invoice, df_code_map, df_gl_acct, df_heico_dept, df_template = load_sheets(invoice_file, template_file)
        df_invoice = preprocess_invoice(df_invoice)

        df_non_heico = aggregate_non_heico(df_invoice, df_code_map, df_gl_acct, df_heico_dept, approver_name)
        df_heico = aggregate_heico(df_invoice, df_heico_dept, df_gl_acct, approver_name)

        # Combine with template and export
        df_result = pd.concat([df_template, df_non_heico], ignore_index=True).sort_values(by='Inter-Co')

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_result.to_excel(writer, sheet_name='Updated Template', index=False)
            df_heico.to_excel(writer, sheet_name='HHI_THC Aggregation', index=False)
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

