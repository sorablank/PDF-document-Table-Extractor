#Imports (Have been installed in environment)
import streamlit as st
import pdfplumber
import pandas as pd
import io

#Function to extract tables from a PDF and return a DataFrame

def table_from_pdf(pdf_file):
	df = []
	with pdfplumber.open(pdf_file) as pdf:
		for page in pdf.pages:
			table = page.extract_table()
			if table:
				 df.append(pd.DataFrame(table))
	return df

#Convert the Dataframe to Excel

def to_excel(df):
	output = io.BytesIO()
	with pd.ExcelWriter(output, engine="openpyxl") as writer:
		for i, df in enumerate(df):
			df.to_excel(writer, sheet_name=f"Table_{i+1}", index = False, header = False)
	output.seek(0)
	return output

# Streamlit UI
st.set_page_config(page_title="Document Extractor", page_icon="📊")
st.image("https://www.answerfinancial.com/ContentResponsive/Assets/images/partners/partners-page/plymouthrock.png", width=300)
st.title("Product Management - Document Extractor 📊")
#st.write("Upload the PDF document that needs to be extracted")

pdf_file = st.file_uploader("Upload a PDF", type=["pdf"])

if pdf_file is not None:
	st.success("✅ File uploaded successfully!")

	#Extract Tables from the uploaded file
	tables = table_from_pdf(pdf_file)

	if tables:
		st.write(f"✅ Extracted {len(tables)} tables from the PDF.")

		#Show preview of the first table (if there are multiple)
		st.write("Exctracted Table preview:")
		st.dataframe(tables[0])
		
		
		#Download Button
		excel_data = to_excel(tables)

		st.download_button(
			label="📥 Download Excel File",
            data=excel_data,
            file_name="extracted_tables.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
		)
	else:
		st.error("❌ No tables found in the uploaded PDF.")