import streamlit as st
import pdfplumber
import pandas as pd
import io
import zipfile
import re
import os

# Streamlit Page Config
st.set_page_config(page_title="Document Extractor", page_icon="📊")
st.image("https://www.answerfinancial.com/ContentResponsive/Assets/images/partners/partners-page/plymouthrock.png", width=300)
st.title("Product Management - Document Extractor 📊")


# ✅ Cache PDF Loading to Avoid Reprocessing on Every Input Change
@st.cache_resource
def load_pdf(pdf_file):
    """Loads the PDF once and stores it in cache."""
    with pdfplumber.open(pdf_file) as pdf:
        return pdf, len(pdf.pages)

# Function to extract tables from PDF with progress updates
def extract_tables_from_pdf(pdf, selected_pages, merge_tables=True):
    """Extract tables from selected pages of a PDF with real-time progress updates."""
    extracted_tables = {}

    progress_bar = st.progress(0)  # Initialize progress bar
    num_pages = len(selected_pages)

    for idx, page_num in enumerate(selected_pages):
        page = pdf.pages[page_num - 1]  
        tables = page.extract_tables()

        if tables:
            for table in tables:
                df = pd.DataFrame(table).dropna(how="all")  # Convert to DataFrame, remove empty rows

                if len(df) > 1:  
                    if df.iloc[0].count() < df.iloc[1].count():
                        table_title = df.iloc[0, 0]  
                        df.columns = df.iloc[1]  
                        df = df[2:].reset_index(drop=True)  

                        if merge_tables:
                            title_row = pd.DataFrame([[table_title] + [""] * (len(df.columns) - 1)], columns=df.columns)
                            df = pd.concat([title_row, df], ignore_index=True)
                    else:
                        table_title = None  
                        df.columns = df.iloc[0]  
                        df = df[1:].reset_index(drop=True)  

                    table_key = (table_title, tuple(df.columns)) if merge_tables else None

                    if merge_tables and table_key in extracted_tables:
                        df_without_title = df.iloc[1:] if extracted_tables[table_key].iloc[0, 0] == table_title else df
                        extracted_tables[table_key] = pd.concat(
                            [extracted_tables[table_key], df_without_title], ignore_index=True
                        )
                    else:
                        extracted_tables[table_key or f"Table_{len(extracted_tables) + 1}"] = df  

        progress_bar.progress(min((idx + 1) / num_pages, 1.0))  # Ensure within 0-1 range

    return extracted_tables

def sanitize_sheet_name(sheet_name):
    """Removes invalid characters and ensures the sheet name is within 31-character limit."""
    sheet_name = re.sub(r'[\/\\\?\*\[\]\:]', '', sheet_name)  # Remove invalid characters
    return sheet_name[:31] if sheet_name else "Sheet1"  # Ensure name is not empty

def save_excel_files(extracted_tables, pdf_filename, selected_pages, max_sheets, enable_splitting):
    """Save extracted tables into multiple Excel files with progress updates."""
    file_outputs = []
    table_items = list(extracted_tables.items())

    page_info = f"_pages_{selected_pages[0]}-{selected_pages[-1]}" if selected_pages else ""
    base_filename = os.path.splitext(pdf_filename)[0] + page_info

    if not enable_splitting:
        max_sheets = len(table_items)  

    progress_bar = st.progress(0)  
    total_files = max((len(table_items) // max_sheets), 1)  

    for i in range(0, len(table_items), max_sheets):
        output = io.BytesIO()
        file_name = f"{base_filename}_part_{i//max_sheets + 1}.xlsx" if enable_splitting else f"{base_filename}.xlsx"

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for j, (key, df) in enumerate(table_items[i:i+max_sheets]):
                sheet_name = sanitize_sheet_name(str(key[0])) if isinstance(key, tuple) and key[0] else f"Table_{j + 1}"
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=True)

        output.seek(0)
        file_outputs.append((file_name, output))

        progress_bar.progress(min((i + 1) / total_files, 1.0))  # Ensure within 0-1 range

    return file_outputs

def create_zip(files, zip_filename):
    """Creates a ZIP file from multiple Excel files."""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file_name, file_data in files:
            zipf.writestr(file_name, file_data.getvalue())  

    zip_buffer.seek(0)
    return zip_buffer

# Upload PDF
pdf_file = st.file_uploader("Upload a PDF", type=["pdf"])

if pdf_file is not None:
    pdf, total_pages = load_pdf(pdf_file)
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
        selected_pages = list(range(1, total_pages + 1))  # Default to all pages

    # Merge Tables Toggle
    merge_option = st.checkbox("Merge tables with the same title and column names", value=True)

    # Enable Excel File Splitting (Default: ON)
    enable_splitting = st.checkbox("Split into multiple Excel files based on sheet count", value=True)

    # Max Sheets Per File Input (Default: 25)
    max_sheets = 25  
    if enable_splitting:
        max_sheets_input = st.text_input("Enter max number of sheets per file", value="25")

        if max_sheets_input.isdigit():
            max_sheets = int(max_sheets_input)
        else:
            st.error("❌ Please enter a valid number for max sheets per file.")

    # ✅ Extract Tables Button with Status Updates
    if st.button("Extract Tables"):
        status = st.status("🔄 Extracting tables... Please wait.")

        extracted_tables = extract_tables_from_pdf(pdf, selected_pages, merge_tables=merge_option)

        if extracted_tables:
            status.update(label="✅ Tables extracted successfully! Now saving to Excel...", state="running")
            
            file_outputs = save_excel_files(extracted_tables, pdf_file.name, selected_pages, max_sheets, enable_splitting)

            # ✅ Show "Download All as ZIP" first
            zip_filename = os.path.splitext(pdf_file.name)[0] + "_Extracted_Tables.zip"
            zip_data = create_zip(file_outputs, zip_filename)

            st.download_button("📥 Download All as ZIP", data=zip_data, file_name=zip_filename, mime="application/zip")

            # ✅ Show individual file downloads
            for filename, data in file_outputs:
                st.download_button(f"📥 Download {filename}", data=data, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            status.update(label="✅ All files ready for download!", state="complete")

#Help Section
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
