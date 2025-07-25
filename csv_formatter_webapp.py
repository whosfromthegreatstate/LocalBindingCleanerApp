import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

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
# Replace the inheritance section (lines ~35-65) with this corrected version:

if {'Parent task', 'Name', 'Tags', 'Notes'}.issubset(df.columns):
    # Clean up the data first
    df['Tags'] = df['Tags'].fillna('').astype(str)
    df['Notes'] = df['Notes'].fillna('').astype(str)
    df['Parent task'] = df['Parent task'].fillna('').astype(str)
    
    # Create a mapping of names to their data for easy lookup
    name_to_data = {}
    for _, row in df.iterrows():
        name_to_data[row['Name']] = {
            'tags': row['Tags'],
            'notes': row['Notes'],
            'parent': row['Parent task'] if row['Parent task'] else None
        }
    
    def get_parent_tags_and_notes(child_name, visited=None):
        """Recursively collect all parent tags and notes"""
        if visited is None:
            visited = set()
        
        if child_name in visited or child_name not in name_to_data:
            return [], []
        
        visited.add(child_name)
        child_data = name_to_data[child_name]
        parent_name = child_data['parent']
        
        all_tags = []
        all_notes = []
        
        # If this child has a parent, get parent's inherited values first
        if parent_name and parent_name in name_to_data:
            parent_tags, parent_notes = get_parent_tags_and_notes(parent_name, visited)
            all_tags.extend(parent_tags)
            all_notes.extend(parent_notes)
            
            # Add the immediate parent's own tags and notes
            parent_data = name_to_data[parent_name]
            if parent_data['tags']:
                all_tags.extend([tag.strip() for tag in parent_data['tags'].split(',') if tag.strip()])
            if parent_data['notes']:
                all_notes.append(parent_data['notes'])
        
        return all_tags, all_notes
    
    # Apply inheritance to each child row
    for idx, row in df.iterrows():
        if row['Parent task']:  # This is a child row
            parent_tags, parent_notes = get_parent_tags_and_notes(row['Name'])
            
            # Combine child's existing values with inherited parent values
            child_existing_tags = [tag.strip() for tag in row['Tags'].split(',') if tag.strip()] if row['Tags'] else []
            child_existing_notes = [row['Notes']] if row['Notes'] else []
            
            # Combine all tags (remove duplicates while preserving order)
            all_tags = child_existing_tags + parent_tags
            unique_tags = []
            for tag in all_tags:
                if tag not in unique_tags:
                    unique_tags.append(tag)
            
            # Combine all notes
            all_notes = child_existing_notes + parent_notes
            
            # Update the dataframe
            df.at[idx, 'Tags'] = ', '.join(unique_tags) if unique_tags else ''
            df.at[idx, 'Notes'] = '\n'.join(filter(None, all_notes)) if all_notes else ''
    # Preview cleaned data
    st.subheader("🔍 Preview of Cleaned Data")
    st.dataframe(df, use_container_width=True)

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

        def apply_table_filter(worksheet):
            tab = Table(displayName="FilteredTable", ref=worksheet.dimensions)
            style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                   showLastColumn=False, showRowStripes=True, showColumnStripes=False)
            tab.tableStyleInfo = style
            worksheet.add_table(tab)

        def hide_columns(worksheet, columns_to_hide):
            for col_letter in columns_to_hide:
                worksheet.column_dimensions[col_letter].hidden = True

        # Sheet 1: Full cleaned data
        df.to_excel(writer, index=False, sheet_name="Formatted Data")
        ws1 = writer.book["Formatted Data"]
        autofit_columns(ws1)
        style_headers(ws1)

        # Apply conditional formatting to rows with "Local Binding Shop Orders" in Projects column
        for row in ws1.iter_rows(min_row=2, max_row=ws1.max_row):
            projects_cell = row[df.columns.get_loc('Projects')] if 'Projects' in df.columns else None
            name_cell = row[df.columns.get_loc('Name')] if 'Name' in df.columns else None
            if projects_cell and projects_cell.value == "Local Binding Shop Orders" and name_cell:
                name_cell.font = Font(bold=True, size=14, color="FFFFFF")
                name_cell.fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")

        # Sheet 2: Filtered & sorted copy
        filtered_df = df[df['Projects'].isna()] if 'Projects' in df.columns else df.copy()
        filtered_df = filtered_df.sort_values(by='Name', ascending=False, key=lambda col: col.str.lower())
        filtered_df.to_excel(writer, index=False, sheet_name="Filtered View")
        ws2 = writer.book["Filtered View"]
        autofit_columns(ws2)
        style_headers(ws2)
        apply_table_filter(ws2)
        hide_columns(ws2, ['A','B','G','H','I','J','O','P'])

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
