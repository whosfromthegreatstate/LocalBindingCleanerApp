import re
import pandas as pd
import streamlit as st
from io import StringIO

st.set_page_config(page_title="Local Binding Formatter", layout="centered")
st.title("📄  Local Binding Asana to Production Formatter")
st.markdown("Upload your `.csv` file below. The app will clean the data and show a preview before download.")

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
        match = re.search(r'^(\d+)[xX]\s*(.*)', name)
        if match:
            return pd.Series([match.group(2).strip(), int(match.group(1))])
        match = re.search(r'(.*?)(\d+)[xX]$', name)
        if match:
            return pd.Series([match.group(1).strip(), int(match.group(2))])
        return pd.Series([name.strip(), None])  # Leave quantity blank if no number

    if 'Name' in df.columns:
        df[['Name', 'Quantity']] = df['Name'].apply(split_name_quantity)

    # Reorder columns
    cols = df.columns.tolist()
    if 'Quantity' in cols:
        cols.insert(cols.index('Name') + 1, cols.pop(cols.index('Quantity')))
        df = df[cols]

    # Preview cleaned data
    st.subheader("🔍 Preview of Cleaned Data")
    st.dataframe(df, use_container_width=True)

    # Create downloadable CSV
    csv = df.to_csv(index=False)
    st.download_button(
        label="📥 Download Cleaned CSV",
        data=csv,
        file_name="cleaned_output.csv",
        mime="text/csv"
    )
