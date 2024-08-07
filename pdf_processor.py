from PIL import Image
import pytesseract
import fitz  # PyMuPDF
import os
import re
import cv2
import numpy as np
# from config_loader import ConfigLoader
# from logger_setup import LoggerSetup
# from directory_manager import DirectoryManager

# Set the Tesseract OCR path
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\marin\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

class PDFProcessor:
    # def __init__(self, config_path):
    #     self.config_path = config_path
    #     self.output_dir = DirectoryManager.create_output_directory()
    #     self.logger = LoggerSetup.setup_logging(self.output_dir)
    #     self.config = ConfigLoader.load_config(self.config_path)
    #     self.target_root = self.config["target_root"]
    #     self.downloaded_pcfrom_path = self.config["downloaded_pcfrom_path"]
    #     self.total_pages_all_pdfs = 0
    #     self.total_pages_copied = 0
    #     self.total_pages_moved = 0
    #     self.tag_to_pdf = {}
    #     self.tag_to_degree = {}
    #     self.tag_to_hours = {}
    #     self.tag_signature = {}
    #     self.tag_paper_requirements = {}
    #     self.tag_graduation_status = {}

    
    def reset_dic(self):
        self.tag_to_pdf = {}
        self.tag_to_degree = {}
        self.tag_to_hours = {}
        self.tag_signature = {}
        self.tag_paper_requirements = {}
        self.tag_graduation_status = {}

    def process_all_pdfs(self):
        pdf_files = [f for f in os.listdir(self.downloaded_pcfrom_path) if f.endswith('.pdf')]
        for pdf_file in pdf_files:
            self.reset_dic()
            pdf_path = os.path.join(self.downloaded_pcfrom_path, pdf_file)
            number_pages = self.split_pdf_by_tag(pdf_path)
            self.logger.info(f"Total pages found in {pdf_file}: {number_pages}")

            for tag, pdf_filename in self.tag_to_pdf.items():
                target_dir = ""
                target_dir = self.find_target_directory(tag)
                if target_dir:

                    pdf_full_path= DirectoryManager.copy_file_to_directory(pdf_filename, target_dir, self.logger, tag)
                    # base_name = os.path.splitext(os.path.basename(pdf_filename))[0]
                    # pdf_full_path = os.path.join(target_dir, base_name + '.pdf')
                    if self.open_pdf(pdf_full_path):
                        self.total_pages_copied+=1

                    if self.is_info_complete(tag, pdf_file):
                        # Log the selected status and requirement
                        status = self.tag_graduation_status.get(tag, [None, None])[1]
                        requirement = self.tag_paper_requirements.get(tag, [None, None])[1]
                        if requirement:
                            self.logger.info(f"for tag {tag}: '{requirement}' is selected in PDF {pdf_file}")                        
                        if status:
                            self.logger.info(f"for tag {tag}: '{status}' is selected in PDF {pdf_file}")


                        if self.handle_folder_move(tag, target_dir):
                            self.total_pages_moved+=1
                else:
                    self.logger.warning(f"Student Folder not found for tag {tag} in PDF {pdf_file}")

        self.logger.info(f"Total number of pages in all PDFs: {self.total_pages_all_pdfs}")
        self.logger.info(f"Total number of copied: {self.total_pages_copied}")
        self.logger.info(f"Total number of moved: {self.total_pages_moved}")
        self.logger.info(f"Process all PDFs completed")

    def is_info_complete(self, tag, pdf_file):
        info_complete = True
        if not self.tag_signature[tag]:
            self.logger.warning(f"Missing signature for tag: {tag} in PDF {pdf_file}")
            info_complete = False

        if "master" in self.tag_to_degree[tag].lower() and tag not in self.tag_to_hours:
            self.logger.warning(f"Missing credit hours for master's degree for tag: {tag} in PDF {pdf_file}")
            info_complete = False

        missing_grad_status = self.tag_graduation_status[tag][0]
        missing_paper_req = self.tag_paper_requirements[tag][0]
        if missing_grad_status and missing_paper_req:
            self.logger.warning(f"Missing paper requirements or graduation status for tag: {tag} in PDF {pdf_file}")
            info_complete = False
        

        return info_complete

    def split_pdf_by_tag(self, pdf_path):
        doc = fitz.open(pdf_path)
        total_pages = doc.page_count
        self.total_pages_all_pdfs += total_pages

        try:
            for page in doc:
                paper_requirements = None
                missing_paper_requirements = None
                graduation_status = None
                missing_graduation_status = None

                pdf_filename = self.save_page_as_pdf(page, pdf_path)
                tag = self.extract_tag(page)
                degree = self.extract_degree(page)
                hours = self.extract_hours(page)
                self.tag_signature[tag[0]] = self.check_signature(page)
                missing_paper_requirements, paper_requirements = self.log_paper_requirements(page)
                missing_graduation_status, graduation_status = self.log_graduation_status(page)

                if tag:
                    self.tag_to_pdf[tag[0]] = pdf_filename
                    # if not missing_paper_requirements and paper_requirements:
                    self.tag_paper_requirements[tag[0]] = (missing_paper_requirements, paper_requirements)
                    # if not missing_graduation_status and graduation_status:
                    self.tag_graduation_status[tag[0]] = (missing_graduation_status, graduation_status)
                    if degree:
                        self.tag_to_degree[tag[0]] = degree[0]
                    if hours:
                        self.tag_to_hours[tag[0]] = hours[0]
                else:
                    self.logger.warning(f"No tag found on page {page.number + 1} of {pdf_path}")

            return total_pages

        except Exception as e:
            self.logger.error(f"Failed to process PDF {pdf_path}: {e}")
        finally:
            doc.close()

    def find_target_directory(self, tag):
        for root, dirs, _ in os.walk(self.target_root):
            for name in dirs:
                if tag in name:
                    return os.path.join(root, name)
        return None

    def save_page_as_pdf(self,page, pdf_path):
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        pdf_filename = f"{base_name}-page-{page.number}-pcform.pdf"
        pdf_full_path = os.path.join(self.target_root, pdf_filename)
        mat = fitz.Matrix(2, 2)
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.save(pdf_full_path, "PDF", resolution=200.0)
        return pdf_full_path

    @staticmethod
    def open_pdf(pdf_path):
        os.startfile(pdf_path)
        return True

    def handle_folder_move(self, tag, target_dir):
        if self.tag_graduation_status[tag][1] == "postpone":
            DirectoryManager.move_folder(target_dir, self.config["postpone_path"], self.logger, tag)
        elif self.tag_graduation_status != "postpone":
            if self.tag_paper_requirements[tag][1] == "No Paper Required":
                DirectoryManager.move_folder(target_dir, self.config["No Paper Required"], self.logger, tag)
            else:
                DirectoryManager.move_folder(target_dir, self.config["Paper Required"], self.logger, tag)



    @staticmethod
    def extract_text_from_coordinates(page, clip_rect, regex):
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=clip_rect)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        # Save the image for debugging
        # debug_img_path = os.path.join(self.output_dir, f"debug_{requirement}_page_{page.number + 1}.png")
        # img.save(debug_img_path)        
        text = pytesseract.image_to_string(img)
        if not text:
            # Preprocess the image
            img_cv = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2GRAY)
            _, img_cv = cv2.threshold(img_cv, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
            
            # Perform OCR
            text = pytesseract.image_to_string(img_cv, config='--psm 6')
        search=re.search(regex, text)
        if search:
            return search[0]
        else:
            return ""

    @staticmethod
    def extract_tag(page):
        return PDFProcessor.extract_text_from_coordinates(page, fitz.Rect(473, 134, 528, 146), r'\d+')
    
    @staticmethod
    def extract_name(page):
        return PDFProcessor.extract_text_from_coordinates(page, fitz.Rect(80.31113011036848, 133.4940929581844,196.7109435722059, 144.68638271413022), r'[a-zA-Z]+(?:\s+[a-zA-Z]+)*')

    @staticmethod
    def extract_program( page):
        return PDFProcessor.extract_text_from_coordinates(page, fitz.Rect(130.67643401212513, 145.80561168972486,259.387766205503, 159.23635939685994), r'[a-zA-Z]+(?:\s+[a-zA-Z]+)*')
    
    @staticmethod
    def extract_degree(self, page):
        return PDFProcessor.extract_text_from_coordinates(page, fitz.Rect(89.34336475707028, 162.19180565627266, 173.9263959390862, 173.6783901377811), r'[a-zA-Z]+(?:\s+[a-zA-Z]+)*')
   
    @staticmethod
    def extract_cehcklist_tag(page):
        return PDFProcessor.extract_text_from_coordinates(page, fitz.Rect(441.2741166817183, 96.63094381729934,554.8994715223395, 108.73856359539832)
, r'\d+')
    
    @staticmethod
    def extract_checklist_name(page):
        return PDFProcessor.extract_text_from_coordinates(page, fitz.Rect(67.800614295742, 96.63094381729934,317.40385279809027, 108.73856359539832), r'[a-zA-Z]+(?:\s+[a-zA-Z]+)*')

    @staticmethod
    def extract_cehcklist_program( page):
        return PDFProcessor.extract_text_from_coordinates(page, fitz.Rect(79.0235663255059, 114.65437490521418,370.5992044136393, 131.77116337597505)
, r'[a-zA-Z]+(?:\s+[a-zA-Z]+)*')
    
    @staticmethod
    def extract_cehcklist_degree( page):
        return PDFProcessor.extract_text_from_coordinates(page, fitz.Rect(435.68598447644183, 117.120761903313,553.9681161547934, 128.2970263138659)
, r'[a-zA-Z]+(?:\s+[a-zA-Z]+)*')




    def extract_hours(self, page):
        return PDFProcessor.extract_text_from_coordinates(page, fitz.Rect(525.57, 226.72, 592.02261, 236.90), r'\d+')

    def check_signature(self, page):
        return PDFProcessor.extract_text_from_coordinates(page, fitz.Rect(379.68, 672.63, 562.93, 684.78), r'[a-zA-Z]+(?:\s+[a-zA-Z]+)*') is not None


    def log_paper_requirements(self, page):
        missing_paper_requirements = True
        paper_requirements = None
        coordinates = {
            "No Paper Required": fitz.Rect(55.67, 290.35, 72.76, 308.50),
            "Research Paper": fitz.Rect(188.10, 290.35, 205.19, 308.50),
            "Thesis": fitz.Rect(301.31, 290.35, 318.40, 308.50),
            "Capstone Report": fitz.Rect(368.60, 290.35, 385.69, 308.50),
            "Dissertation": fitz.Rect(487.15, 290.35, 504.24, 308.50)
        }
        for requirement, coord in coordinates.items():
            if self.check_checkbox_selection(page, coord, 0.1824812030075188):
                # self.logger.info(f"'{requirement}' is selected on page {page.number + 1}")
                paper_requirements = requirement
                missing_paper_requirements = False
        return missing_paper_requirements, paper_requirements

    def log_graduation_status(self, page):
        missing_graduation_status = True
        graduation_status = None
        coordinates = {
            "selection_1": fitz.Rect(44.61, 349.8837622348279 - 17.075, 44.61 + 44.4, 349.8837622348279),
            "postpone": fitz.Rect(44.61, 377.68551106888515 - 17.075, 44.61 + 44.4, 377.68551106888515),
            "selection_3": fitz.Rect(44.61, 405.7061713110846 - 17.075, 44.61 + 44.4, 405.7061713110846),
            "selection_4": fitz.Rect(44.61, 435.91594563470585 - 17.075, 44.61 + 44.4, 435.91594563470585)
        }
        for status, coord in coordinates.items():
            if self.check_checkbox_selection(page, coord, 0.03):
                # self.logger.info(f"'{status}' is selected on page {page.number + 1}")
                graduation_status = status
                missing_graduation_status = False
        return missing_graduation_status, graduation_status

    def check_checkbox_selection(self, page, rect, threshold):
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=rect)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img_cv = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)
        white_pixels = cv2.countNonZero(binary)
        total_pixels = binary.size
        density = white_pixels / total_pixels
        return density > threshold

    def rename_logic(self):
        pdf_files = self.collect_pdf_files()
        for path, files in pdf_files.items():
            if 1 in files and 2 in files:
                self.rename_pcform(files[1], files[2])

    def collect_pdf_files(self):
        pdf_files = {}
        for p in self.config["paths"]:
            directory = os.fsencode(p)
            for path, subdirs, files in os.walk(directory):
                path = path.decode("utf-8")
                for filename in files:
                    filename = filename.decode("utf-8")
                    if filename.endswith("pdf"):
                        if path not in pdf_files:
                            pdf_files[path] = {}
                        if "checklist" in filename.lower():
                            pdf_files[path][1] = os.path.join(path, filename)
                        elif "pcform" in filename.lower() or "page" in filename.lower():
                            pdf_files[path][2] = os.path.join(path, filename)
        return pdf_files

    def rename_pcform(self, checklist_filename, pcform_filename):
        try:
            checklist_basename = os.path.basename(checklist_filename)
            name_part_after_checklist = checklist_basename.split(" - Checklist(244) - ")[1]
            new_pcform_filename = f"2 - PCForm(244) - {name_part_after_checklist}.pdf"
            new_pcform_path = os.path.join(os.path.dirname(pcform_filename), new_pcform_filename)
            if ".pdf.pdf" in new_pcform_path:
                new_pcform_path = new_pcform_path.replace(".pdf.pdf", ".pdf")
            if pcform_filename != new_pcform_path:
                os.rename(pcform_filename, new_pcform_path)
                self.logger.info(f'Renamed "{pcform_filename}" to "{new_pcform_path}"')
        except Exception as e:
            self.logger.error(f"Failed to rename {pcform_filename}: {e}")
