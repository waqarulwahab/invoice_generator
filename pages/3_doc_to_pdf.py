
import os
import subprocess
import streamlit as st

st.sidebar.page_link('pages/0_generate_invoice_DD.py',     label="Generate Invoice",               icon="🏡")
st.sidebar.page_link('pages/1_list_of_clients_projects.py',label="List of Clients / Projects List",icon="📓")    
st.sidebar.page_link('pages/2_add_new_client_project.py',  label="Add New Client/Project record",  icon="✒️")
st.sidebar.page_link('pages/3_doc_to_pdf.py',              label="Convert To PDF",                 icon="🖨️")



import os
import pypandoc
import streamlit as st

# Ensure Pandoc is available
pypandoc.download_pandoc()

# Function to convert DOCX to PDF using pypandoc
def convert_to_pdf(docx_file_path):
    try:
        # Generate the PDF file name
        pdf_path = os.path.splitext(docx_file_path)[0] + ".pdf"
        
        # Convert the DOCX file to PDF using pypandoc
        output = pypandoc.convert_file(docx_file_path, 'pdf', outputfile=pdf_path)
        assert output == ""

        return pdf_path
    except Exception as e:
        st.error(f"Conversion failed: {e}")
        return None


st.title("DOCX to PDF Converter")

# File uploader
uploaded_file = st.file_uploader("Upload a DOCX file", type=["docx"])

if uploaded_file is not None:
    st.info("File successfully uploaded!")

    # Check if the button is clicked
    if st.button("Convert to PDF"):
        # Save uploaded file to a temporary location
        with open(uploaded_file.name, 'wb') as f:
            f.write(uploaded_file.getbuffer())

        # Convert the DOCX file to PDF
        pdf_path = convert_to_pdf(uploaded_file.name)

        if pdf_path:
            with open(pdf_path, "rb") as pdf_file:
                st.download_button(label="Download PDF", data=pdf_file, file_name=os.path.basename(pdf_path))
            st.success(f"PDF file created: {pdf_path}")
        else:
            st.error("Failed to convert DOCX to PDF. Please check the file and try again.")









