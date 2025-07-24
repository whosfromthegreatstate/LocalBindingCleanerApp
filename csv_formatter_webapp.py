import re
import pandas as pd
import streamlit as st
from io import StringIO

st.set_page_config(page_title="CSV Formatter", layout="centered")
st.title("📄 CSV Formatter Tool")
st.markdown("Upload your .csv file and download the cleaned version.")

uploaded_file = st.file_uploader("Upload CSV", type=["csv"])

if uploaded_file:
    # Read CSV
    df = pd.read_csv(uploaded_file)

    # Remove Column A (first column)
    df = df.iloc[:, 1:]

    # Fill down the Section/Column
    if 'Section/Column' in df.columns:
        df['Section/Column'] = df['Section/Column'].fillna(method='ffill')

    # Split Name into Name + Quantity
    def split_name_quantity(name):
        if pd.isna(name):
            return pd.Series([None, None])
        match = re.search(r'(\d+)[xX]\s*(.*)', name)
        if match:
            return pd.Series([match.group(2).strip(), int(match.group(1))])
        match = re.search(r'(.*?)(\d+)[xX]', name)
        if match:
            return pd.Series([match.group(1).strip(), int(match.group(2))])
        return pd.Series([name.strip(), 1])  # Default quantity = 1

    if 'Name' in df.columns:
        df[['Name', 'Quantity']] = df['Name'].apply(split_name_quantity)

    # Reorder columns
    cols = df.columns.tolist()
    if 'Quantity' in cols:
        cols.insert(cols.index('Name') + 1, cols.pop(cols.index('Quantity')))
        df = df[cols]

    # Create downloadable CSV
    csv = df.to_csv(index=False)
    st.download_button(
        label="📥 Download Cleaned CSV",
        data=csv,
        file_name="cleaned_output.csv",
        mime="text/csv"
    )
