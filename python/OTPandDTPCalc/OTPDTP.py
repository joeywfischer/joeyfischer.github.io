import streamlit as st
import pandas as pd
import io

st.title("OTP Calculator")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file is not None:
    df_main = pd.read_excel(uploaded_file, sheet_name=0, engine='openpyxl')
    df_mapping = pd.read_excel(uploaded_file, sheet_name='Code Map', engine='openpyxl')

    contains_codes = df_mapping['Contains Codes'].dropna().tolist()
    dollars_codes = df_mapping['Dollars Codes'].dropna().tolist()
    hours_codes = df_mapping['Hours Codes'].dropna().tolist()

    df_filtered = df_main[df_main['Earnings Code'].isin(contains_codes)]
    relevant_ids = df_filtered['ID'].unique().tolist()
    df_main = df_main[df_main['ID'].isin(relevant_ids)]

    df_grouped = df_main.groupby('ID')

    total_dollars_per_id = {}
    total_hours_per_id = {}

    for id, group in df_grouped:
        total_dollars = group[group['Earnings Code'].isin(dollars_codes)]['Current Amount'].sum()
        total_hours = group[group['Earnings Code'].isin(hours_codes)]['Current Hours'].sum()
        total_dollars_per_id[id] = total_dollars
        total_hours_per_id[id] = total_hours

    adjusted_otp_rate_per_id = {
        id: (total_dollars_per_id[id] / total_hours_per_id[id]) if total_hours_per_id[id] > 0 else 0
        for id in total_dollars_per_id
    }

    difference_in_amount_per_id = {}
    otp_df = df_main[df_main['Earnings Code'] == 'OTP'].copy()

    for id, adjusted_otp_rate in adjusted_otp_rate_per_id.items():
        otp_row = otp_df[otp_df['ID'] == id]
        if not otp_row.empty:
            otp_current_hours = otp_row['Current Hours'].iloc[0]
            original_total_amount = otp_row['Total(Current Amount)'].iloc[0]
            adjusted_amount = adjusted_otp_rate * (otp_current_hours * 0.5)
            difference = adjusted_amount - original_total_amount
            difference_in_amount_per_id[id] = difference

    df_differences = pd.DataFrame(list(difference_in_amount_per_id.items()), columns=['ID', 'Difference in Amount'])
    df_differences['Earnings Code'] = 'OTP'
    df_differences['Difference in Amount'] = df_differences['Difference in Amount'].apply(
        lambda x: f'${x:,.2f}' if pd.notnull(x) else ''
    )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_differences.to_excel(writer, index=False)
    output.seek(0)

    st.download_button(
        label="Download OTP Calculated Excel",
        data=output,
        file_name="OTP_Calculated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
