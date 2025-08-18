import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, PatternFill, numbers
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook

# ... [previous code for processing invoice and template remains unchanged]

        # --- Create Aflac Invoice and Support File ---
        company_totals = df_invoice[['Company', 'Monthly Premium']].dropna().groupby('Company')['Monthly Premium'].sum().reset_index()
        df_support = company_totals.rename(columns={'Company': 'Row Labels', 'Monthly Premium': 'Sum of Monthly Premium'})
        df_support['Full Company Name'] = df_support['Row Labels'].map(
            dict(zip(
                df_code_map['Invoice Company Code'].astype(str).str.upper(),
                df_code_map['Company Description'].astype(str).str.strip()
            ))
        )

        # Drop rows with missing values
        df_support = df_support.dropna()

        # Add Grand Total row
        grand_total = df_support['Sum of Monthly Premium'].sum()
        df_support.loc[len(df_support.index)] = ['Grand Total', grand_total, '']

        # Save with formatting
        output_support = io.BytesIO()
        with pd.ExcelWriter(output_support, engine='openpyxl') as writer:
            df_support.to_excel(writer, index=False, sheet_name='Invoice Support')
            workbook = writer.book
            worksheet = writer.sheets['Invoice Support']

            # Format header row
            header_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
            header_font = Font(bold=True)
            for col_num, col_name in enumerate(df_support.columns, 1):
                cell = worksheet.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = header_font

            # Adjust column widths
            for col in worksheet.columns:
                max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
                worksheet.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2

            # Format currency column
            currency_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            for row in range(2, worksheet.max_row + 1):
                cell = worksheet.cell(row=row, column=2)
                cell.number_format = currency_format

            # Style Grand Total row
            grand_total_row_idx = worksheet.max_row
            grand_total_label_cell = worksheet.cell(row=grand_total_row_idx, column=1)
            grand_total_label_cell.font = Font(bold=True)
            grand_total_label_cell.fill = header_fill

        output_support.seek(0)

        # --- Streamlit Outputs ---
        st.success("Processing complete!")

        st.download_button(
            label="Download Updated Medius Template",
            data=output_template,
            file_name="Complete_Aflac_Medius_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="Download Aflac Invoice and Support",
            data=output_support,
            file_name="Aflac_Invoice_and_Support.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

