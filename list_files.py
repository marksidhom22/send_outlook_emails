import os
import csv
import fitz

from pdf_processor import PDFProcessor
def list_pdfs_with_grandfather_folder(directory):
    pdf_files = []

    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.pdf'):
                # Get the full path of the file
                full_path = os.path.join(root, file)
                # Get the directory two levels up
                parent_folder = os.path.basename(full_path)

                grandfather_folder = os.path.basename(os.path.dirname(os.path.dirname(full_path)))
                pdf_files.append([file,parent_folder, grandfather_folder]+process_page(full_path))
                

    return pdf_files

def write_to_csv(file_list, csv_filename):
    with open(csv_filename, 'w', newline='') as csvfile:
        csv_writer = csv.writer(csvfile)
        csv_writer.writerow(['PDF File', 'Parent Folder', 'Grandfather Folder', 'Name', 'Program', 'Tag', 'Degree'])  # Header
        for item in file_list:
            csv_writer.writerow(item)

def get_pages(pdf_path):
    doc = fitz.open(pdf_path)
    pages=[]
    for page in doc:
        pages.append(page)
    return pages

def process_page(pdf_path):
    try:
        pages = get_pages(pdf_path)
        if not pages:
            raise ValueError("No pages found in the PDF.")

        # Initialize variables outside the loop
        name, program, tag, degree = "", "", "", ""

        for page in pages:
            try:
                name = PDFProcessor.extract_checklist_name(page)
                program = PDFProcessor.extract_cehcklist_program(page)
                tag = PDFProcessor.extract_cehcklist_tag(page)
                degree = PDFProcessor.extract_cehcklist_degree(page)

                # # Check if the extracted values are valid
                # if not all([name, program, tag, degree]):
                #     return ""
                #     raise ValueError("One or more extracted values are empty.")

                # # Exit the loop once valid values are found
                # break

            except Exception as e:
                print(f"Error processing page: {e}")
                return ["","","",""]

        # Ensure all extracted values are lists and have at least one element
        # if not all(isinstance(i, list) and i for i in [name, program, tag, degree]):
            # raise ValueError("One or more extracted values are invalid.")

        return [name, program, tag, degree]

    except Exception as e:
        print(f"Error processing PDF: {e}")
        return ["","","",""]
        


# Example usage
directory_path = 'P:/GradPublic/LeMark (Graduation)/Grad Files/202440 grad files/1 - Needs PC Form'
csv_filename = 'pdf_files_with_grandfather_folders_1.csv'
pdf_files = list_pdfs_with_grandfather_folder(directory_path)
write_to_csv(pdf_files, csv_filename)

print(f'PDF files and their grandfather folders have been written to {csv_filename}.')
