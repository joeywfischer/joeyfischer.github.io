import streamlit as st
import pandas as pd

st.title("Debug: Department Mapping")

invoice_file = st.file_uploader("Upload Invoice File", type=["xlsx"])
template_file = st.file_uploader("Upload Template File", type=["xlsx"])

if invoice_file and template_file:
    try:
        # Load sheets
        df_invoice = pd.read_excel(invoice_file, sheet_name='Detail', engine='openpyxl')
        df_heico_dept = pd.read_excel(template_file, sheet_name='Heico Departments', engine='openpyxl')

        # Convert to string for mapping
        df_invoice['Department'] = df_invoice['Department'].astype(str).str.strip()
        df_heico_dept['Department Code'] = df_heico_dept['Department Code'].astype(str).str.strip()

        # Show unique values
        st.subheader("Unique Department values from Invoice File")
        st.write(sorted(df_invoice['Department'].unique()))

        st.subheader("Unique Department Code values from Template File")
        st.write(sorted(df_heico_dept['Department Code'].unique()))

        # Optional: show unmatched values
        unmatched = set(df_invoice['Department'].unique()) - set(df_heico_dept['Department Code'].unique())
        st.subheader("Unmatched Department values")
        st.write(sorted(unmatched))

    except Exception as e:
        st.error(f"Error loading files: {e}")

