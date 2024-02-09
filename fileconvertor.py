import os
import win32com.client
from docx import Document
from fpdf import FPDF
import comtypes.client
from pdf2docx import converter

# Define conversion functions for DOC files

def convert_docx_to_pdf(docx_file):
    # Input and output file paths
    docx_path = os.path.abspath(docx_file)
    pdf_path = os.path.abspath(docx_file.replace(".docx", ".pdf"))

    #Create a Word Application
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False  # Set to True if you want to see the Word application

    try:
        # Open the DOCX file
        doc = word.Documents.Open(docx_path)

        # Specify PDF format (17 is for PDF)
        pdf_format = 17

        # Save the document as PDF 
        doc.SaveAs(pdf_path, FileFormat=pdf_format)
        doc.Close()

    except Exception as e:
        print(f"Error: {e}")
    finally:
        # Quit the Word application
        word.Quit()

# Menu Ddisplay
def display_menu():
    print("Select a file type to convert:")
    print("1. DOCX")
    choice = input("Enter your choice: ")
    return choice

def main():
    choice = display_menu()
    file_types = {1: "docx", 2: "pdf"}

    file_type = file_types.get(int(choice))
    if file_type is None:
        print("Invalid choice!")
        return

    file_path = input("Enter the path of the file: ")

    if not os.path.exists(file_path):
        print("File not found!")
        return

    # Define conversion options for each file type
    conversions = {
        "docx": {"pdf": convert_docx_to_pdf},

    }

    # Retrieve conversion options for the selected file type
    conversion_options = conversions.get(file_type)
    if conversion_options is None:
        print("Conversion options not available for this file type!")
        return

    # Display available conversion options for the selected file type
    print(f"Conversion options for {file_type.upper()}:")
    for target_format in conversion_options.keys():
        print(f"- {target_format.upper()}")

    # Get user's target format choice
    target_format = input("Enter the target format: ").lower()

    # Retrieve conversion function for the target format
    conversion_function = conversion_options.get(target_format)
    if conversion_function is None:
        print("Invalid target format!")
        return

    # Perform conversion
    conversion_function(file_path)
    print("Conversion complete!")

if __name__ == "__main__":
    main()