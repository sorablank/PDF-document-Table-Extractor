import streamlit as st
import pdfplumber
import pandas as pd
import io
import zipfile
import re
import os

# 🔹 Set Streamlit Page Config
st.set_page_config(page_title="Document Extractor", page_icon="📊")

# 🔹 Plymouth Rock Logo & Title
st.image("https://www.answerfinancial.com/ContentResponsive/Assets/images/partners/partners-page/plymouthrock.png", width=300)
st.title("Product Management - Document Extractor 📊")

def extract_tables_from_pdf(pdf_file, selected_pages, merge_tables=True):
    """Extract tables from selected pages of a PDF, optionally merging tables with the same title and columns."""
    extracted_tables = {}

    with pdfplumber.open(pdf_file) as pdf:
        total_pages = len(pdf.pages)

        # Default to all pages if no selection is made
        pages_to_extract = selected_pages if selected_pages else list(range(1, total_pages + 1))

        for page_num in pages_to_extract:
            page = pdf.pages[page_num - 1]  # Convert to zero-based index
            tables = page.extract_tables()

            if tables:
                for table in tables:
                    df = pd.DataFrame(table).dropna(how="all")  # Convert to DataFrame, remove empty rows

                    if len(df) > 1:  # Ensure table has at least two rows
                        # Detect if the first row is a title
                        if df.iloc[0].count() < df.iloc[1].count():
                            table_title = df.iloc[0, 0]  # Use first cell as title
                            df.columns = df.iloc[1]  # Use second row as actual column names
                            df = df[2:].reset_index(drop=True)  # Remove title + headers row
                            
                            # Add the title only if it's the first time merging
                            if merge_tables:
                                title_row = pd.DataFrame([[table_title] + [""] * (len(df.columns) - 1)], columns=df.columns)
                                df = pd.concat([title_row, df], ignore_index=True)
                        else:
                            table_title = None  # No title row detected
                            df.columns = df.iloc[0]  # Use first row as column names
                            df = df[1:].reset_index(drop=True)  # Remove headers row

                        # Compare using (Title, Column Names) if merging is enabled
                        table_key = (table_title, tuple(df.columns)) if merge_tables else None

                        if merge_tables and table_key in extracted_tables:
                            df_without_title = df.iloc[1:] if extracted_tables[table_key].iloc[0, 0] == table_title else df
                            extracted_tables[table_key] = pd.concat(
                                [extracted_tables[table_key], df_without_title], ignore_index=True
                            )
                        else:
                            extracted_tables[table_key or f"Table_{len(extracted_tables) + 1}"] = df  # Store new table

    return extracted_tables, total_pages


def save_excel_files(extracted_tables, pdf_filename, selected_pages, max_sheets):
    """Split extracted tables into multiple Excel files if max_sheets is set and return a list of filenames + file data."""
    file_outputs = []
    table_items = list(extracted_tables.items())

    # Generate base filename
    page_info = f"_pages_{selected_pages[0]}-{selected_pages[-1]}" if selected_pages else ""
    base_filename = os.path.splitext(pdf_filename)[0] + page_info

    # Split into multiple files if needed
    for i in range(0, len(table_items), max_sheets):
        output = io.BytesIO()
        file_name = f"{base_filename}_part_{i//max_sheets + 1}.xlsx"

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for j, (key, df) in enumerate(table_items[i:i+max_sheets]):
                sheet_name = str(key[0])[:31] if isinstance(key, tuple) and key[0] else f"Table_{j + 1}"
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=True)

        output.seek(0)
        file_outputs.append((file_name, output))

    return file_outputs


def create_zip(files, zip_filename):
    """Creates a ZIP file from multiple Excel files."""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file_name, file_data in files:
            zipf.writestr(file_name, file_data.getvalue())  # Add Excel file contents to zip

    zip_buffer.seek(0)
    return zip_buffer

# Upload PDF
pdf_file = st.file_uploader("Upload a PDF", type=["pdf"])

if pdf_file is not None:
    with pdfplumber.open(pdf_file) as pdf:
        total_pages = len(pdf.pages)

    st.write(f"📄 This PDF has **{total_pages} pages**.")

    # Page Range Selection Toggle
    enable_page_selection = st.checkbox("Select specific page range")
    
    # Page Range Input
    selected_pages_input = ""
    if enable_page_selection:
        selected_pages_input = st.text_input(f"Enter page numbers (1-{total_pages}), e.g., 1,3,5-7", value="")
    
    # Validate Page Input
    if selected_pages_input:
        try:
            selected_pages = [p for sublist in [
                list(range(int(x.split("-")[0]), int(x.split("-")[1]) + 1)) if "-" in x else [int(x)]
                for x in re.split(r"[,\s]+", selected_pages_input)
            ] for p in sublist if 1 <= p <= total_pages]
        except:
            st.error("❌ Please enter a valid range of numbers.")
            selected_pages = []
    else:
        selected_pages = []  # Default to all pages

    # Merge Tables Toggle
    merge_option = st.checkbox("Merge tables with the same title and column names", value=True)

    # Enable Excel File Splitting
    enable_splitting = st.checkbox("Split into multiple Excel files based on sheet count")

    # Max Sheets Per File Input
    max_sheets = 100  # Default value
    if enable_splitting:
        max_sheets_input = st.text_input("Enter max number of sheets per file", value="100")

        # Validate Max Sheets Input
        if max_sheets_input.isdigit():
            max_sheets = int(max_sheets_input)
        else:
            st.error("❌ Please enter a valid number for max sheets per file.")

    if st.button("Extract Tables"):
        extracted_tables, total_pages = extract_tables_from_pdf(pdf_file, selected_pages, merge_tables=merge_option)

        if extracted_tables:
            file_outputs = save_excel_files(extracted_tables, pdf_file.name, selected_pages, max_sheets)

            # ✅ "Download All as ZIP" button
            zip_filename = os.path.splitext(pdf_file.name)[0] + "_Extracted_Tables.zip"
            zip_data = create_zip(file_outputs, zip_filename)

            st.download_button(
                label="📥 Download All as ZIP",
                data=zip_data,
                file_name=zip_filename,
                mime="application/zip"
            )

            # ✅ Individual File Download Buttons
            for filename, data in file_outputs:
                st.download_button(
                    label=f"📥 Download {filename}",
                    data=data,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.error("❌ No tables found in the selected pages.")

# 🔹 Help Section
st.markdown(
    """
    ---
    ### **Help & Support**  
    If you encounter any issues, have feature requests, or need assistance,  
    please reach out to:  

    **Anthony Sacco, Data Scientist**  
    📧 **Email:** [asacco@plymouthrock.com](mailto:asacco@plymouthrock.com)  

    **Hamzah Mukadam, Product and Data Science Intern**  
    📧 **Email:** [hmukadam@plymouthrock.com](mailto:hmukadam@plymouthrock.com)  
    """
)
