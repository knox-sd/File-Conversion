import os
import win32com.client
from docx import Document
from fpdf import FPDF

# Define conversion functions for DOC files
def convert_doc_to_docx(doc_file):
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(doc_file)
    doc.SaveAs(doc_file.replace(".doc", ".docx"), 12)  # 12 for .docx format
    doc.Close()
    word.Quit()
    
def convert_doc_to_html(doc_file):
    pass

def convert_doc_to_pdf(doc_file):
    # Load content from DOC file
    doc = Document(doc_file)

    # Create a new PDF document
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Convert each paragraph to PDF text
    for paragraph in doc.paragraphs:
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0, 10, paragraph.text)

    # Save the PDF document
    pdf.output(doc_file.replace(".doc", ".pdf"))

# Define conversion functions for DOCX files
def convert_docx_to_doc(docx_file):
    pass

def convert_docx_to_html(docx_file):
    pass

def convert_docx_to_pdf(docx_file):
    pass

# Define conversion functions for XLSX files
def convert_xlsx_to_xls(xlsx_file):
    pass

def convert_xlsx_to_pdf(xlsx_file):
    pass

# Define conversion functions for XLS files
def convert_xls_to_xlsx(xls_file):
    pass

def convert_xls_to_pdf(xls_file):
    pass

# Define conversion functions for RT files
def convert_rt_to_docx(rt_file):
    pass

def convert_rt_to_html(rt_file):
    pass

def convert_rt_to_pdf(rt_file):
    pass

# Define conversion functions for HTML files
def convert_html_to_docx(html_file):
    pass

def convert_html_to_pdf(html_file):
    pass

# Define conversion functions for PDF files
def convert_pdf_to_docx(pdf_file):
    pass

def convert_pdf_to_xlsx(pdf_file):
    pass

def convert_pdf_to_pptx(pdf_file):
    pass

def convert_pdf_to_html(pdf_file):
    pass

# Define conversion functions for PPTX files
def convert_pptx_to_pdf(pptx_file):
    pass

def convert_pptx_to_ppt(pptx_file):
    pass

# Menu Ddisplay
def display_menu():
    print("Select a file type to convert:")
    print("1. DOC")
    print("2. DOCX")
    print("3. XLSX")
    print("4. XLS")
    print("5. RT")
    print("6. HTML")
    print("7. PDF")
    print("8. PPTX")
    print("9. PPT")
    choice = input("Enter your choice (1-9): ")
    return choice

def main():
    choice = display_menu()
    file_types = {1: "doc", 2: "docx", 3: "xlsx", 4: "xls", 5: "rt", 6: "html", 7: "pdf", 8: "pptx", 9: "ppt"}

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
        "doc": {"docx": convert_doc_to_docx, "html": convert_doc_to_html, "pdf": convert_doc_to_pdf},
        "docx": {"doc": convert_docx_to_doc, "html": convert_docx_to_html, "pdf": convert_docx_to_pdf},
        "xlsx": {"xls": convert_xlsx_to_xls, "pdf": convert_xlsx_to_pdf},
        "xls": {"xlsx": convert_xls_to_xlsx, "pdf": convert_xls_to_pdf},
        "rt": {"docx": convert_rt_to_docx, "html": convert_rt_to_html, "pdf": convert_rt_to_pdf},
        "html": {"docx": convert_html_to_docx, "pdf": convert_html_to_pdf},
        "pdf": {"docx": convert_pdf_to_docx, "xlsx": convert_pdf_to_xlsx, "pptx": convert_pdf_to_pptx, "html": convert_pdf_to_html},
        "pptx": {"pdf": convert_pptx_to_pdf, "ppt": convert_pptx_to_ppt},
        "ppt": {"pptx": convert_pptx_to_ppt},
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