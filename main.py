# zip included
import fitz  # PyMuPDF
from PIL import Image, ImageEnhance, ImageFilter
from pyzbar.pyzbar import decode
import io
import os
import csv
import logging
import re
import numpy as np
import cv2
import tkinter as tk
from tkinter import filedialog, messagebox
import zipfile


# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def extract_images_from_pdf(pdf_path):
    """Extracts images from a given PDF and returns a list of PIL images with their page numbers."""
    pdf_document = fitz.open(pdf_path)
    images = []
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        image_list = page.get_images(full=True)
        for img_index, img in enumerate(image_list):
            xref = img[0]
            try:
                base_image = pdf_document.extract_image(xref)
                image_bytes = base_image["image"]
                image = Image.open(io.BytesIO(image_bytes))
                images.append((image, page_num + 1))  # Store image with page number
            except IOError as e:
                logging.warning(
                    f"Skipping invalid image on page {page_num + 1}, image {img_index + 1} due to error: {e}")
            except Exception as e:
                logging.error(f"Unsupported image format on page {page_num + 1}, image {img_index + 1}: {e}")
    return images


def preprocess_image(image):
    """Applies various preprocessing techniques to enhance the image for QR code detection."""
    # Convert PIL image to OpenCV format
    image_cv = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)

    # Resize image
    image_cv = cv2.resize(image_cv, (0, 0), fx=1.5, fy=1.5)

    # Convert to grayscale
    gray = cv2.cvtColor(image_cv, cv2.COLOR_BGR2GRAY)

    # Apply Gaussian Blur to reduce noise
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)

    # Apply adaptive thresholding
    adaptive_thresh = cv2.adaptiveThreshold(blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                            cv2.THRESH_BINARY, 11, 2)

    # Convert back to PIL image
    preprocessed_image = Image.fromarray(adaptive_thresh)

    return preprocessed_image


def read_qr_codes_from_image(image):
    """Decodes QR codes from a given PIL image and returns a list of QR data."""
    qr_data_list = []
    try:
        # Try reading the QR codes from the original image
        decoded_objects = decode(image)
        for obj in decoded_objects:
            qr_data_list.append(obj.data.decode('utf-8'))

        # If no QR codes found, try with enhanced image
        if not qr_data_list:
            enhanced_image = enhance_image(image)
            decoded_objects = decode(enhanced_image)
            for obj in decoded_objects:
                qr_data_list.append(obj.data.decode('utf-8'))

        # If still no QR codes found, try with preprocessed image
        if not qr_data_list:
            preprocessed_image = preprocess_image(image)
            decoded_objects = decode(preprocessed_image)
            for obj in decoded_objects:
                qr_data_list.append(obj.data.decode('utf-8'))
    except Exception as e:
        logging.error(f"Error decoding QR code from image: {e}")
    return qr_data_list


def enhance_image(image):
    """Enhances the image quality to improve QR code detection."""
    # Convert to grayscale
    gray_image = image.convert('L')

    # Increase contrast
    enhancer = ImageEnhance.Contrast(gray_image)
    enhanced_image = enhancer.enhance(2)

    # Sharpen the image
    sharpener = ImageEnhance.Sharpness(enhanced_image)
    sharp_image = sharpener.enhance(2)

    # Binarize the image
    binarized_image = sharp_image.point(lambda x: 0 if x < 128 else 255, '1')

    return binarized_image


def sanitize_filename(filename):
    """Sanitizes a string to be used as a valid filename."""
    return re.sub(r'[\\/:"*?<>|]+', '_', filename)


def process_pdf(pdf_path, image_folder):
    """Processes the PDF and generates a report of QR codes, saving any found QR code images."""
    pdf_filename = os.path.basename(pdf_path)
    pdf_number = pdf_filename.split('.')[0]  # Extract PDF number from filename
    pdf_save_dir = os.path.join(image_folder, pdf_number)
    os.makedirs(pdf_save_dir, exist_ok=True)

    images = extract_images_from_pdf(pdf_path)
    qr_code_data_list = []
    qr_image_paths = []

    for image, page_num in images:
        qr_codes = read_qr_codes_from_image(image)
        if qr_codes:
            for qr_index, qr_code_data in enumerate(qr_codes):
                # Sanitize QR code data to create a valid filename
                qr_code_id = sanitize_filename(qr_code_data)

                # Save QR code image with PDF number, page number, and QR code ID in filename
                image_filename = f"{pdf_number}page{page_num}qr{qr_index + 1}id{qr_code_id}.png"
                image_path = os.path.join(pdf_save_dir, image_filename)
                try:
                    image.save(image_path)
                    qr_image_paths.append(image_path)

                    # Append QR code data to list
                    qr_code_data_list.append({
                        "PDF Name": pdf_filename,
                        "Page Number": page_num,
                        "QR Code Data": qr_code_data,
                        "QR Image Path": image_path
                    })
                except Exception as e:
                    logging.error(f"Error saving image {image_path}: {e}")

    # If no QR codes found, add a record indicating "NO QR code found"
    if not qr_code_data_list:
        qr_code_data_list.append({
            "PDF Name": pdf_filename,
            "Page Number": "",
            "QR Code Data": "NO QR code found",
            "QR Image Path": ""
        })

    # Prepare report
    report = {
        "PDF Name": pdf_filename,
        "QR Codes Found": len(qr_code_data_list) if qr_code_data_list[0]['QR Code Data'] != "NO QR code found" else 0,
        "QR Code Details": qr_code_data_list
    }
    return report


def save_report_to_csv(report, csv_path):
    """Saves the report to a CSV file."""
    with open(csv_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=["PDF Name", "Page Number", "QR Code Data", "QR Image Path"])
        writer.writeheader()
        for qr_code_info in report["QR Code Details"]:
            writer.writerow(qr_code_info)


def process_pdfs_in_folder(folder_path, image_folder, report_file):
    """Processes all PDFs in a given folder and generates a consolidated report."""
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    all_qr_code_data_list = []

    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        logging.info(f"Processing {pdf_path}")
        try:
            report = process_pdf(pdf_path, image_folder)
            all_qr_code_data_list.extend(report["QR Code Details"])
            logging.info(f"Found {report['QR Codes Found']} QR codes in {pdf_file}")
        except Exception as e:
            logging.error(f"Failed to process {pdf_path}: {e}")
            # If an error occurs, log it in the report
            all_qr_code_data_list.append({
                "PDF Name": pdf_file,
                "Page Number": "",
                "QR Code Data": "Error processing PDF",
                "QR Image Path": ""
            })

    # Generate consolidated report
    with open(report_file, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=["PDF Name", "Page Number", "QR Code Data", "QR Image Path"])
        writer.writeheader()
        for qr_code_info in all_qr_code_data_list:
            writer.writerow(qr_code_info)

    logging.info(f"Report generated at {report_file}")


def extract_pdfs_from_zip(zip_path, extract_to_folder):
    """Extracts all PDFs from a given ZIP file to the specified folder."""
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to_folder)
    logging.info(f"Extracted all PDFs from {zip_path} to {extract_to_folder}")


# Tkinter GUI implementation
class QRCodeExtractorApp:

    def __init__(self, root):
        self.root = root
        self.root.title("PDFs QR Code Extractor")
        self.root.geometry("600x400")
        self.root.configure(bg='#F0F8FF')

        # Frame for selecting file or folder
        self.selection_frame = tk.Frame(root, bg='#F0F8FF')
        self.selection_frame.pack(pady=10)

        self.file_radio = tk.Radiobutton(self.selection_frame, text="Single PDF File", value=1,
                                         command=self.select_file, bg='#F0F8FF', fg='#FF6347')
        self.file_radio.grid(row=0, column=0, padx=10)
        self.folder_radio = tk.Radiobutton(self.selection_frame, text="Folder of PDFs", value=2,
                                           command=self.select_folder, bg='#F0F8FF', fg='#FF6347')
        self.folder_radio.grid(row=0, column=1, padx=10)
        self.zip_radio = tk.Radiobutton(self.selection_frame, text="ZIP File of PDFs", value=3, command=self.select_zip,
                                        bg='#F0F8FF', fg='#FF6347')
        self.zip_radio.grid(row=0, column=2, padx=10)

        # Label and entry for selected file or folder path
        self.path_label = tk.Label(root, text="Selected Path:", bg='#F0F8FF', fg='#4682B4')
        self.path_label.pack()
        self.path_entry = tk.Entry(root, width=50)
        self.path_entry.pack(pady=5)

        # Browse button
        self.browse_button = tk.Button(root, text="Browse", command=self.browse, bg='#87CEFA', fg='#4B0082')
        self.browse_button.pack(pady=5)

        # Save directory
        self.save_dir_label = tk.Label(root, text="Save Directory:", bg='#F0F8FF', fg='#4682B4')
        self.save_dir_label.pack()
        self.save_dir_entry = tk.Entry(root, width=50)
        self.save_dir_entry.pack(pady=5)
        self.browse_save_dir_button = tk.Button(root, text="Browse", command=self.browse_save_dir, bg='#87CEFA',
                                                fg='#4B0082')
        self.browse_save_dir_button.pack(pady=5)

        # Report file name
        self.report_label = tk.Label(root, text="Report and Image Folder Name:", bg='#F0F8FF', fg='#4682B4')
        self.report_label.pack()
        self.report_entry = tk.Entry(root, width=50)
        self.report_entry.pack(pady=5)

        # Process button
        self.process_button = tk.Button(root, text="Process", command=self.process, bg='#32CD32', fg='#FFFFFF')
        self.process_button.pack(pady=20)

        self.file_or_folder = None

    def select_file(self):
        self.file_or_folder = "file"

    def select_folder(self):
        self.file_or_folder = "folder"

    def select_zip(self):
        self.file_or_folder = "zip"

    def browse(self):
        if self.file_or_folder == "file":
            file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if file_path:
                self.path_entry.delete(0, tk.END)
                self.path_entry.insert(0, file_path)
        elif self.file_or_folder == "folder":
            folder_path = filedialog.askdirectory()
            if folder_path:
                self.path_entry.delete(0, tk.END)
                self.path_entry.insert(0, folder_path)
        elif self.file_or_folder == "zip":
            zip_path = filedialog.askopenfilename(filetypes=[("ZIP files", "*.zip")])
            if zip_path:
                self.path_entry.delete(0, tk.END)
                self.path_entry.insert(0, zip_path)
        else:
            messagebox.showerror("Error",
                                 "Please select either 'Single PDF File', 'Folder of PDFs', or 'ZIP File of PDFs'")

    def browse_save_dir(self):
        save_dir = filedialog.askdirectory()
        if save_dir:
            self.save_dir_entry.delete(0, tk.END)
            self.save_dir_entry.insert(0, save_dir)

    def process(self):
        path = self.path_entry.get()
        save_dir = self.save_dir_entry.get()
        report_file_name = self.report_entry.get()

        if not path or not save_dir or not report_file_name:
            messagebox.showerror("Error", "Please provide all the required paths and report file name.")
            return

        sanitized_name = sanitize_filename(report_file_name)
        report_file = os.path.join(save_dir, f"{sanitized_name}.csv")
        image_folder = os.path.join(save_dir, sanitized_name)
        os.makedirs(image_folder, exist_ok=True)

        if self.file_or_folder == "file":
            logging.info(f"Processing single PDF file: {path}")
            try:
                report = process_pdf(path, os.path.dirname(path))  # Save images within the same directory as the PDF
                save_report_to_csv(report, report_file)
                messagebox.showinfo("Success", f"QR code extraction completed.\nReport saved at {report_file}")
            except Exception as e:
                logging.error(f"Error processing file: {e}")
                messagebox.showerror("Error", f"Failed to process the file.\nError: {e}")
        elif self.file_or_folder == "folder":
            logging.info(f"Processing folder of PDFs: {path}")
            try:
                process_pdfs_in_folder(path, image_folder, report_file)
                messagebox.showinfo("Success", f"QR code extraction completed.\nReport saved at {report_file}")
            except Exception as e:
                logging.error(f"Error processing folder: {e}")
                messagebox.showerror("Error", f"Failed to process the folder.\nError: {e}")
        elif self.file_or_folder == "zip":
            logging.info(f"Processing ZIP file: {path}")
            extract_to_folder = os.path.join(save_dir, "extracted_pdfs")
            os.makedirs(extract_to_folder, exist_ok=True)
            try:
                extract_pdfs_from_zip(path, extract_to_folder)
                process_pdfs_in_folder(extract_to_folder, image_folder, report_file)
                messagebox.showinfo("Success", f"QR code extraction completed.\nReport saved at {report_file}")
            except Exception as e:
                logging.error(f"Error processing ZIP file: {e}")
                messagebox.showerror("Error", f"Failed to process the ZIP file.\nError: {e}")
        else:
            messagebox.showerror("Error",
                                 "Please select either 'Single PDF File', 'Folder of PDFs', or 'ZIP File of PDFs'")


if __name__ == "__main__":
    root = tk.Tk()
    app = QRCodeExtractorApp(root)
    root.mainloop()