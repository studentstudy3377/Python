import os
import PyPDF2
from tkinter import Tk,messagebox
from tkinter.filedialog import askopenfilenames

def remove_first_page(input_pdf_path, output_pdf_path):
    try:
        # Open the PDF file
        with open(input_pdf_path, "rb") as input_pdf_file:
            reader = PyPDF2.PdfReader(input_pdf_file)
            
            # Create a PdfWriter object to save the new PDF
            writer = PyPDF2.PdfWriter()

            # Add all pages except the first one
            for page_num in range(1, len(reader.pages)):
                writer.add_page(reader.pages[page_num])

            # Write the output PDF
            with open(output_pdf_path, "wb") as output_pdf_file:
                writer.write(output_pdf_file)

    except Exception as e:
        print(f"Error processing the PDF: {e}")

def select_multiple_pdfs_and_process():
    # Hide the root window of Tkinter (since we only need the file dialog)
    Tk().withdraw()

    # Open file dialog to select multiple PDF files
    input_pdf_paths = askopenfilenames(title="Select PDF files", filetypes=[("PDF files", "*.pdf")])

    if input_pdf_paths:
        for input_pdf_path in input_pdf_paths:
            # Set the output PDF file name (same name)
            output_pdf_path = os.path.join(os.path.dirname(input_pdf_path), "" + os.path.basename(input_pdf_path))

            # Call the function to remove the first page
            remove_first_page(input_pdf_path, output_pdf_path)
    
        messagebox.showinfo("","Done !")

    else:
        print("No files selected.")

# Run the function
select_multiple_pdfs_and_process()