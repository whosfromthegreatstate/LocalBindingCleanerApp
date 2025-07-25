import pandas as pd
import streamlit as st
from io import StringIO, BytesIO

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
        parts = str(name).split()
        name_parts = [p for p in parts if not p.isdigit()]
        quantity_parts = [p for p in parts if p.isdigit()]
        cleaned_name = " ".join(name_parts).strip() if name_parts else None
        quantity = int(quantity_parts[0]) if quantity_parts else None
        return pd.Series([cleaned_name, quantity])

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

    # Create downloadable XLSX without using unsupported libraries
    xlsx_output = BytesIO()
    with pd.ExcelWriter(xlsx_output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Formatted Data")
    st.download_button(
        label="📥 Download Excel (.xlsx)",
        data=xlsx_output.getvalue(),
        file_name="cleaned_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
