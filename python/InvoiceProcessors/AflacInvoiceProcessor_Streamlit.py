import streamlit as st
import pandas as pd
import io
import fitz  # PyMuPDF

st.title("Aflac Invoice Processor")

invoice_file = st.file_uploader("Upload Aflac Invoice Excel File", type=["xlsx"])
template_file = st.file_uploader("Upload Medius Template Excel File", type=["xlsx"])

if invoice_file and template_file:
    # Load invoice and template data
    df_invoice = pd.read_excel(invoice_file, sheet_name='Detail', engine='openpyxl')
    df_template = pd.read_excel(template_file, sheet_name=0, engine='openpyxl')
    df_code_map = pd.read_excel(template_file, sheet_name='Code Map', engine='openpyxl')

    # Normalize key columns
    df_invoice['Company'] = df_invoice['Company'].astype(str).str.strip().str.upper()
    df_invoice['Division'] = df_invoice['Division'].astype(str).str.strip()
    df_invoice['Monthly Premium'] = pd.to_numeric(df_invoice['Monthly Premium'], errors='coerce')

    # Mapping by Division Code
    df_code_map_div = df_code_map[
        df_code_map['Template Desc'].notna() &
        (df_code_map['Template Desc'].astype(str).str.strip() != '') &
        (df_code_map['Division Code'].astype(str).str.strip() != '')
    ]
    for _, row in df_code_map_div.iterrows():
        desc = str(row['Template Desc']).strip()
        division_code = str(row['Division Code']).strip()
        total = df_invoice[df_invoice['Division'] == division_code]['Monthly Premium'].sum()
        if total > 0:
            match_rows = df_template[df_template['DESC'].astype(str).str.strip().str.lower() == desc.lower()].index
            df_template.loc[match_rows, 'NET'] = total

    # Mapping by Company Code
    df_code_map_company = df_code_map[
        df_code_map['Template Desc'].notna() &
        (df_code_map['Template Desc'].astype(str).str.strip() != '') &
        ((df_code_map['Division Code'].isna()) | (df_code_map['Division Code'].astype(str).str.strip() == ''))
    ]
    for _, row in df_code_map_company.iterrows():
        desc = str(row['Template Desc']).strip()
        invoice_company_code = str(row['Invoice Company Code']).strip().upper()
        total = df_invoice[df_invoice['Company'] == invoice_company_code]['Monthly Premium'].sum()
        if total > 0:
            match_rows = df_template[df_template['DESC'].astype(str).str.strip().str.lower() == desc.lower()].index
            df_template.loc[match_rows, 'NET'] = total

    # Mapping by Inter-Co
    df_code_map_no_desc = df_code_map[
        df_code_map['Template Desc'].isna() |
        (df_code_map['Template Desc'].astype(str).str.strip() == '')
    ]
    company_totals = df_invoice.groupby('Company')['Monthly Premium'].sum().reset_index()
    interco_mapping = dict(zip(
        df_code_map_no_desc['Invoice Company Code'].astype(str).str.upper(),
        df_code_map_no_desc['Template Inter-Co'].astype(str).str.strip()
    ))
    for invoice_company, interco in interco_mapping.items():
        total = company_totals[company_totals['Company'] == invoice_company]['Monthly Premium'].sum()
        if total > 0:
            match_rows = df_template[df_template['Inter-Co'].astype(str).str.strip() == interco].index
            df_template.loc[match_rows, 'NET'] = total

    # HHI and THC Department Mapping
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
    df_template['CC'] = df_template['CC'].replace('nan', '', regex=False)

    # Create Excel with Summary and Pivot Table
    df_summary = pd.read_excel(invoice_file, sheet_name='Summary', engine='openpyxl')
    pivot_company = df_invoice.pivot_table(index='Company', values='Monthly Premium', aggfunc='sum').reset_index()
    pivot_division = df_invoice.pivot_table(index='Division', values='Monthly Premium', aggfunc='sum').reset_index()
    pivot_department = df_invoice.pivot_table(index='Department', values='Monthly Premium', aggfunc='sum').reset_index()

    excel_output = io.BytesIO()
    with pd.ExcelWriter(excel_output, engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name='Invoice', index=False)
        pivot_company.to_excel(writer, sheet_name='Pivot Table', index=False, startrow=0)
        pivot_division.to_excel(writer, sheet_name='Pivot Table', index=False, startrow=len(pivot_company)+3)
        pivot_department.to_excel(writer, sheet_name='Pivot Table', index=False, startrow=len(pivot_company)+len(pivot_division)+6)
    excel_output.seek(0)

    # Create PDF from Summary
    pdf_output = io.BytesIO()
    doc = fitz.open()
    summary_text = df_summary.to_string(index=False)
    page = doc.new_page()
    text_rect = fitz.Rect(50, 50, 550, 800)
    page.insert_textbox(text_rect, summary_text, fontsize=10, fontname="helv")
    doc.save(pdf_output)
    doc.close()
    pdf_output.seek(0)

    # Final Template Output
    template_output = io.BytesIO()
    df_template.to_excel(template_output, index=False, engine='openpyxl')
    template_output.seek(0)

    st.success("Processing complete!")

    st.download_button("Download Updated Medius Template", template_output, "Complete_Aflac_Medius_Template.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("Download Invoice Summary and Pivot Table", excel_output, "Aflac_Invoice_Summary_Pivot.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("Download Aflac Invoice PDF", pdf_output, "Aflac Invoice.pdf", "application/pdf")

