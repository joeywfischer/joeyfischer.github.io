import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

st.title("Aflac Invoice Processor")

invoice_file = st.file_uploader("Upload Aflac Invoice Excel File", type=["xlsx"])
template_file = st.file_uploader("Upload Medius Template Excel File", type=["xlsx"])

if invoice_file and template_file:
    try:
        # Load invoice and template data
        df_invoice = pd.read_excel(invoice_file, sheet_name='Detail', engine='openpyxl')
        df_template = pd.read_excel(template_file, sheet_name=0, engine='openpyxl')
        df_code_map = pd.read_excel(template_file, sheet_name='Code Map', engine='openpyxl')

        # Normalize key columns
        df_invoice['Company'] = df_invoice['Company'].astype(str).str.strip().str.upper()
        df_invoice['Division'] = df_invoice['Division'].astype(str).str.strip()
        df_invoice['Monthly Premium'] = pd.to_numeric(df_invoice['Monthly Premium'], errors='coerce')

        # --- Invoice Support Table ---
        df_support = df_invoice[['Company', 'Monthly Premium']].dropna().groupby('Company')['Monthly Premium'].sum().reset_index()
        df_support = df_support.rename(columns={'Company': 'Invoice Company Code', 'Monthly Premium': 'Sum of Monthly Premium'})
        df_support['Full Company Name'] = df_support['Invoice Company Code'].map(
            dict(zip(
                df_code_map['Invoice Company Code'].astype(str).str.upper(),
                df_code_map['Company Description'].astype(str).str.strip()
            ))
        )
        df_support = df_support.dropna()
        df_support = df_support[['Invoice Company Code', 'Sum of Monthly Premium', 'Full Company Name']]
        df_support.loc[len(df_support.index)] = ['Grand Total', df_support['Sum of Monthly Premium'].sum(), '']

        # --- Additional Break-Down Table ---
        valid_divisions = df_code_map['Division Code'].dropna().astype(str).str.strip().unique()
        df_division_filtered = df_invoice[df_invoice['Division'].isin(valid_divisions)]
        df_division_summary = df_division_filtered.groupby('Division')['Monthly Premium'].sum().reset_index()
        df_division_summary = df_division_summary.rename(columns={'Division': 'Division Code', 'Monthly Premium': 'Sum of Monthly Premium'})
        df_division_summary['Division Description'] = df_division_summary['Division Code'].map(
            dict(zip(
                df_code_map['Division Code'].astype(str).str.strip(),
                df_code_map['Division Description'].astype(str).str.strip()
            ))
        )
        df_division_summary = df_division_summary.dropna()
        df_division_summary = df_division_summary[['Division Description', 'Sum of Monthly Premium']]
        df_division_summary.loc[len(df_division_summary.index)] = ['Grand Total', df_division_summary['Sum of Monthly Premium'].sum()]

        # --- Save to one sheet with spacing ---
        output_combined = io.BytesIO()
        with pd.ExcelWriter(output_combined, engine='openpyxl') as writer:
            sheet_name = 'Combined Report'
            workbook = writer.book
            worksheet = workbook.create_sheet(title=sheet_name)

            # Write Invoice Support
            for r_idx, row in enumerate(df_support.itertuples(index=False), start=2):
                for c_idx, value in enumerate(row, start=1):
                    worksheet.cell(row=r_idx, column=c_idx, value=value)

            # Write headers
            for c_idx, col_name in enumerate(df_support.columns, start=1):
                cell = worksheet.cell(row=1, column=c_idx, value=col_name)
                cell.fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
                cell.font = Font(bold=True)

            # Format currency column
            for r in range(2, 2 + len(df_support)):
                worksheet.cell(row=r, column=2).number_format = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'

            # Style Grand Total
            total_row = 1 + len(df_support)
            worksheet.cell(row=total_row, column=1).font = Font(bold=True)
            worksheet.cell(row=total_row, column=1).fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')

            # Leave spacing
            start_row = total_row + 3

            # Write Additional Break-Down
            for r_idx, row in enumerate(df_division_summary.itertuples(index=False), start=start_row + 1):
                for c_idx, value in enumerate(row, start=1):
                    worksheet.cell(row=r_idx, column=c_idx, value=value)

            # Headers
            for c_idx, col_name in enumerate(df_division_summary.columns, start=1):
                cell = worksheet.cell(row=start_row, column=c_idx, value=col_name)
                cell.fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
                cell.font = Font(bold=True)

            # Format currency
            for r in range(start_row + 1, start_row + 1 + len(df_division_summary)):
                worksheet.cell(row=r, column=2).number_format = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'

            # Style Grand Total
            total_row_2 = start_row + len(df_division_summary)
            worksheet.cell(row=total_row_2, column=1).font = Font(bold=True)
            worksheet.cell(row=total_row_2, column=1).fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')

            # Adjust column widths
            for col in worksheet.columns:
                max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
                worksheet.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2

        output_combined.seek(0)

        # --- Streamlit Output ---
        st.success("Combined report generated!")

        st.download_button(
            label="Download Combined Aflac Report",
            data=output_combined,
            file_name="Combined_Aflac_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")

