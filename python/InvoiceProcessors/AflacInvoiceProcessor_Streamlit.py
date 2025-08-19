import streamlit as st
import pandas as pd

st.title("Aflac Invoice and Support Generator with Debugging")

# Upload both files
invoice_file = st.file_uploader("Upload Invoice Excel File", type=["xlsx"], key="invoice")
template_file = st.file_uploader("Upload Template Excel File (with Code Map sheet)", type=["xlsx"], key="template")

if invoice_file and template_file:
    # Load invoice data
    invoice_xls = pd.ExcelFile(invoice_file, engine='openpyxl')
    detail_df = pd.read_excel(invoice_xls, sheet_name='Detail', engine='openpyxl')

    # Load template data and detect first sheet dynamically
    template_xls = pd.ExcelFile(template_file, engine='openpyxl')
    first_sheet_name = template_xls.sheet_names[0]
    template_df = pd.read_excel(template_xls, sheet_name=first_sheet_name, engine='openpyxl')

    # Load Code Map sheet
    code_map_df = pd.read_excel(template_xls, sheet_name='Code Map', engine='openpyxl')

    # Table 1: Summary by Company (grouped properly to avoid duplicates)
    premium_summary = detail_df.groupby('Company')['Monthly Premium'].sum().reset_index()
    premium_summary.columns = ['Row Labels', 'Sum of Monthly Premium']

    result_df = premium_summary.merge(
        code_map_df[['Invoice Company Code', 'Company Description']],
        left_on='Row Labels', right_on='Invoice Company Code', how='left'
    ).drop_duplicates(subset=['Row Labels'])

    result_df.rename(columns={'Company Description': 'Full Company Name'}, inplace=True)
    final_df = result_df[['Row Labels', 'Sum of Monthly Premium', 'Full Company Name']]

    # Table 2: Hierarchical breakdown by Company and Division Description
    code_map_filtered = code_map_df[code_map_df['Division Description'].notna()]
    valid_companies = code_map_filtered['Invoice Company Code'].unique()
    filtered_detail_df = detail_df[detail_df['Company'].isin(valid_companies)]

    breakdown_rows = []
    company_totals = filtered_detail_df.groupby('Company')['Monthly Premium'].sum().reset_index()
    for company, comp_premium in company_totals.values:
        breakdown_rows.append({'Label': company, 'Monthly Premium': comp_premium})
        company_divisions = filtered_detail_df[filtered_detail_df['Company'] == company]
        company_divisions = company_divisions.merge(
            code_map_df[['Invoice Company Code', 'Division Description']],
            left_on='Company', right_on='Invoice Company Code', how='left'
        )
        division_totals = company_divisions.groupby('Division Description')['Monthly Premium'].sum().reset_index()
        for div_desc, div_premium in division_totals.values:
            if pd.notna(div_desc):
                breakdown_rows.append({'Label': f"  {div_desc}", 'Monthly Premium': div_premium})
    breakdown_df = pd.DataFrame(breakdown_rows)

    # Table 3: THC & HHI breakdown using Department column
    thchhi_df = detail_df[detail_df['Company'].isin(['THC', 'HHI'])]
    total_thchhi = thchhi_df['Monthly Premium'].sum()

    department_summary = thchhi_df.groupby('Department')['Monthly Premium'].sum().reset_index()

    # Debug table 1: Raw Department codes
    st.subheader("Debug Table 1: Raw Department Codes from THC & HHI")
    st.dataframe(department_summary)

    # Debug table 2: Stripped Department codes
    department_summary['Stripped Code'] = department_summary['Department'].apply(
        lambda x: x[2:] if isinstance(x, str) and len(x) >= 2 else x
    )
    st.subheader("Debug Table 2: Stripped Department Codes")
    st.dataframe(department_summary[['Department', 'Stripped Code', 'Monthly Premium']])

    # Debug table 3: Mapped DESC values
    cc_desc_map = template_df[['CC', 'DESC']].dropna()
    cc_desc_dict = dict(zip(cc_desc_map['CC'].astype(str), cc_desc_map['DESC']))
    department_summary['Mapped DESC'] = department_summary['Stripped Code'].apply(
        lambda x: cc_desc_dict.get(x, x)
    )
    st.subheader("Debug Table 3: Mapped DESC Values")
    st.dataframe(department_summary[['Stripped Code', 'Mapped DESC', 'Monthly Premium']])

    # Final third table
    department_rows = [{'Department': 'THC & HHI', 'Sum of Monthly Premium': total_thchhi}]
    for _, row in department_summary.iterrows():
        department_rows.append({
            'Department': row['Mapped DESC'],
            'Sum of Monthly Premium': row['Monthly Premium']
        })
    department_df = pd.DataFrame(department_rows)

    # Display final tables
    st.subheader("Summary Table")
    st.dataframe(final_df)

    st.subheader("Company & Division Breakdown (Filtered)")
    st.dataframe(breakdown_df)

    st.subheader("THC & HHI Department Breakdown")
    st.dataframe(department_df)

