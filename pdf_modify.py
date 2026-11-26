import PyPDF2
import streamlit as st
from io import BytesIO
import zipfile

# Function to remove the first page from a PDF
def remove_first_page(input_pdf):
    try:
        # Create a PdfReader from the in-memory file
        reader = PyPDF2.PdfReader(input_pdf)

        # Create a PdfWriter object to save the new PDF
        writer = PyPDF2.PdfWriter()

        # Add all pages except the first one
        for page_num in range(1, len(reader.pages)):
            writer.add_page(reader.pages[page_num])

        # Write the output PDF to a BytesIO object
        output_pdf = BytesIO()
        writer.write(output_pdf)
        output_pdf.seek(0)
        return output_pdf
    
    except Exception as e:
        st.error(f"Error processing the PDF: {e}")
        return None

# Main function to run Streamlit UI
def pdf_modify_run():
    st.title("PDF First Page Remove")
    with st.spinner("Process.."):
        with st.sidebar:
            # Multiple file uploader
            uploaded_files = st.file_uploader("Choose multiple PDF files", type="pdf", accept_multiple_files=True)

        if uploaded_files:
            modified_pdfs = []  # List to hold modified PDFs
            filenames = []  # List to store the original filenames

            # Process each uploaded file in-memory
            for uploaded_file in uploaded_files:
                output_pdf = remove_first_page(uploaded_file)

                if output_pdf:
                    # Store the modified PDFs and their names
                    modified_pdfs.append(output_pdf)
                    filenames.append(f"{uploaded_file.name}")

            # Zip the modified PDFs into one file
            if modified_pdfs:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for i, pdf in enumerate(modified_pdfs):
                        # Add the PDF to the zip file with its filename
                        zip_file.writestr(filenames[i], pdf.getvalue())

                zip_buffer.seek(0)  # Rewind the buffer to the beginning

                # Provide a single download button for the zip file
                st.download_button(
                    label="Download",
                    data=zip_buffer,
                    file_name="HRD PDF.zip",
                    mime="application/zip"
                )

if __name__ == "__main__":

    pdf_modify_run()
