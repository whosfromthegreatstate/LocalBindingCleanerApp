import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

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

    # Prepare Excel with multiple sheets and formatting
    xlsx_output = BytesIO()
    with pd.ExcelWriter(xlsx_output, engine="openpyxl") as writer:
        def autofit_columns(worksheet):
            for column_cells in worksheet.columns:
                max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
                worksheet.column_dimensions[get_column_letter(column_cells[0].column)].width = max_length + 2

        def style_headers(worksheet):
            for cell in worksheet[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')

        # Sheet 1: Full cleaned data
        df.to_excel(writer, index=False, sheet_name="Formatted Data")
        ws1 = writer.book["Formatted Data"]
        autofit_columns(ws1)
        style_headers(ws1)

        # Sheet 2: Filtered & sorted copy
        filtered_df = df[df['Projects'].isna()] if 'Projects' in df.columns else df.copy()
        filtered_df = filtered_df.sort_values(by='Name', ascending=False)
        filtered_df.to_excel(writer, index=False, sheet_name="Filtered View")
        ws2 = writer.book["Filtered View"]
        autofit_columns(ws2)
        style_headers(ws2)

        # Sheet 3: Pivot summary of sizes (if present)
        size_counts = pd.DataFrame()
        if 'Name' in df.columns:
            keywords = ['small', 'medium', 'large']
            for key in keywords:
                count = df[df['Name'].str.lower().str.contains(key, na=False)]['Quantity'].sum()
                size_counts.at[0, key.title()] = count if pd.notna(count) else 0
            size_counts = size_counts.T.reset_index()
            size_counts.columns = ['Size', 'Total Quantity']
            size_counts.to_excel(writer, index=False, sheet_name="Pivot Summary")
            ws3 = writer.book["Pivot Summary"]
            autofit_columns(ws3)
            style_headers(ws3)

    st.download_button(
        label="📥 Download Excel (.xlsx)",
        data=xlsx_output.getvalue(),
        file_name="cleaned_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
