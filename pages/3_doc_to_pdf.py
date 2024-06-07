
import streamlit as st
st.sidebar.page_link('pages/0_generate_invoice_DD.py',     label="Generate Invoice",               icon="🏡")
st.sidebar.page_link('pages/1_list_of_clients_projects.py',label="List of Clients / Projects List",icon="📓")    
st.sidebar.page_link('pages/2_add_new_client_project.py',  label="Add New Client/Project record",  icon="✒️")
st.sidebar.page_link('pages/3_doc_to_pdf.py',              label="Convert To PDF",                 icon="🖨️")

import os
import subprocess
import streamlit as st

# Function to convert DOCX to PDF using unoconv
def convert_to_pdf(docx_file):
    # Get the absolute path of the DOCX file
    docx_path = os.path.abspath(docx_file)
    # Generate the PDF file name
    pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
    
    try:
        # Use unoconv to convert DOCX to PDF
        cmd = ['unoconv', '-f', 'pdf', '-o', os.path.dirname(pdf_path), docx_path]
        subprocess.call(cmd)
        
        # Return the path to the PDF
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
        with open(uploaded_file.name, 'wb') as f:
            f.write(uploaded_file.getbuffer())

        # Convert the DOCX file to PDF
        pdf_path = convert_to_pdf(uploaded_file.name)

        if pdf_path:
            st.success(f"PDF file created: [Download PDF]({pdf_path})")
        else:
            st.error("Failed to convert DOCX to PDF. Please check the file and try again.")







