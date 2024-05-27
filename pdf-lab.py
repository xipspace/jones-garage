# xipspace

import os
import re
import sys
from datetime import datetime

import cv2
import numpy as np
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image
import fitz

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower()
            for text in re.split(r'(\d+)', s)]

def get_files_with_extension(folder, extension):
    return [f for f in os.listdir(folder) if f.lower().endswith(extension)]

def find_valid_pdfs():
    pdf_files = get_files_with_extension('.', '.pdf')
    num_pdfs = len(pdf_files)

    if num_pdfs == 0:
        print("No valid PDF files found in the current directory.")
    else:
        print(f"Found {num_pdfs} valid PDF files in the current directory.")

    return pdf_files

def transform_filename(filename):
    # Convert to lowercase
    filename = filename.lower()
    # Replace multiple spaces with a single space
    filename = re.sub(r'\s+', ' ', filename)
    # Replace spaces with underscores
    filename = filename.replace(' ', '_')
    return filename

def generate_default_output_folder(input_pdf):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    input_filename = os.path.splitext(os.path.basename(input_pdf))[0]
    input_filename = transform_filename(input_filename)
    return os.path.join(os.path.dirname(input_pdf), f"{timestamp}_{input_filename}")

def split_pdf(input_pdf, output_folder):
    pdf_reader = PdfReader(input_pdf)
    input_filename = os.path.splitext(os.path.basename(input_pdf))[0]
    input_filename = transform_filename(input_filename)
    os.makedirs(output_folder, exist_ok=True)

    num_pages = len(pdf_reader.pages)
    print(f"Processing '{input_filename}' with {num_pages} pages.")

    processed_folders = []

    for page_num, page in enumerate(pdf_reader.pages, start=1):
        output_filename = f'{input_filename}_page_{page_num}.pdf'
        output_path = os.path.join(output_folder, output_filename)
        write_page_to_pdf(page, pdf_reader.metadata, output_path)

    processed_folders.append(output_folder)

    return processed_folders

def write_page_to_pdf(page, metadata, output_path):
    pdf_writer = PdfWriter()
    pdf_writer.add_page(page)
    pdf_writer.add_metadata(metadata)

    with open(output_path, 'wb') as output_pdf:
        pdf_writer.write(output_pdf)

def list_folder_info(folders):
    for folder in folders:
        pdf_files = get_files_with_extension(folder, '.pdf')
        image_files = get_files_with_extension(folder, '.jpg')

        num_pdfs = len(pdf_files)
        num_images = len(image_files)

        print(f"Folder: {folder}, PDF Count: {num_pdfs}, Image Count: {num_images}")


def extract_pages(pdf_files):
    processed_folders = []
    for pdf_file in pdf_files:
        processed_folders.extend(split_pdf(pdf_file, generate_default_output_folder(pdf_file)))
    print("Extraction complete.")
    return processed_folders

def convert_to_image(folders):
    for folder in folders:
        pdf_files = get_files_with_extension(folder, '.pdf')
        for pdf_file in pdf_files:
            pdf_path = os.path.join(folder, pdf_file)
            doc = fitz.open(pdf_path)
            for page_num in range(doc.page_count):
                process_page(folder, pdf_file, page_num, doc)

def process_page(folder, pdf_file, page_num, doc):
    page = doc[page_num]
    pixmap = page.get_pixmap(matrix=fitz.Matrix(200 / 72, 200 / 72))
    img = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
    img_path = os.path.join(folder, f"{os.path.splitext(pdf_file)[0]}.jpg")
    img.save(img_path, 'JPEG')

def treat_images(folders):
    for folder in folders:
        pdf_files = get_files_with_extension(folder, '.pdf')
        jpg_files = get_files_with_extension(folder, '.jpg')
        treated_files = [f for f in jpg_files if f.startswith('treated_')]

        for pdf_file in pdf_files:
            pdf_path = os.path.join(folder, pdf_file)

            with fitz.open(pdf_path) as doc:
                for page_num in range(doc.page_count):
                    process_image(folder, pdf_file, page_num)

        for file in pdf_files + jpg_files:
            file_path = os.path.join(folder, file)
            if file not in treated_files:
                os.remove(file_path)

def process_image(folder, pdf_file, page_num, quality=35, max_width=1500):
    pdf_path = os.path.join(folder, pdf_file)
    doc = fitz.open(pdf_path)
    page = doc[page_num]
    pixmap = page.get_pixmap(matrix=fitz.Matrix(200 / 72, 200 / 72))
    img = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
    img_path = os.path.join(folder, f"{os.path.splitext(pdf_file)[0]}_page_{page_num + 1}.jpg")

    try:
        Image.MAX_IMAGE_PIXELS = None
        image = np.array(img)

        if image is not None:
            if image.shape[1] > max_width:
                ratio = max_width / image.shape[1]
                new_height = int(image.shape[0] * ratio)

                if new_height > 0:
                    raw_image = Image.fromarray(image)
                    new_size = (max_width, new_height)
                    raw_resampled_image = raw_image.resize(new_size, resample=Image.LANCZOS)
                    image = np.array(raw_resampled_image)

            filtered_image = filter_image(image)

            output_filename = "treated_" + os.path.splitext(os.path.basename(pdf_file))[0] + ".jpg"
            output_path = os.path.join(folder, output_filename)

            img = Image.fromarray(filtered_image)
            img.save(output_path, 'JPEG', quality=quality)

        else:
            print(f"Error! Skipping file: {img_path}")
    except Exception as e:
        print(f"Error processing file {img_path}: {e}")

def filter_image(image):
    try:
        gray_image = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
        denoised_image = cv2.GaussianBlur(gray_image, (5, 5), 0)
        bilateral_filtered_image = cv2.bilateralFilter(denoised_image, d=5, sigmaColor=10, sigmaSpace=10)
        sharpened_image = sharpen_image(bilateral_filtered_image)
        restore_image = cv2.cvtColor(sharpened_image, cv2.COLOR_GRAY2RGB)

        return restore_image
    except Exception as e:
        print(f"Error filtering image: {e}")
        return None

def sharpen_image(image):
    kernel = np.array([[-1, -1, -1],
                       [-1,  9, -1],
                       [-1, -1, -1]])
    sharpened_image = cv2.filter2D(image, -1, kernel)
    return sharpened_image

def convert_image_to_pdf(processed_folders):
    try:
        for folder_path in processed_folders:
            jpg_files = get_files_with_extension(folder_path, '.jpg')

            if not jpg_files:
                print(f"No JPG files found in the folder: {folder_path}")
                continue

            jpg_files.sort(key=natural_sort_key)

            for jpg_file in jpg_files:
                jpg_file_path = os.path.join(folder_path, jpg_file)
                image = Image.open(jpg_file_path)

                pdf_file_path = jpg_file_path.replace(".jpg", ".pdf")
                image.save(pdf_file_path, "PDF")

            for jpg_file in jpg_files:
                jpg_file_path = os.path.join(folder_path, jpg_file)
                os.remove(jpg_file_path)

    except Exception as e:
        print(f"Error converting images to PDF in folder {folder_path}: {e}")

def merge_pdfs(processed_folders):
    try:
        for folder_path in processed_folders:
            pdf_files = get_files_with_extension(folder_path, '.pdf')

            if not pdf_files:
                print(f"No PDF files found in the folder: {folder_path}")
                continue

            if len(pdf_files) == 1:
                print(f"Only one PDF file found in the folder: {folder_path}. Renaming the file.")
                pdf_file = pdf_files[0]
                old_path = os.path.join(folder_path, pdf_file)
                new_path = os.path.join(folder_path, f"{os.path.basename(folder_path)}.pdf")
                os.rename(old_path, new_path)
                continue

            pdf_files.sort(key=natural_sort_key)

            pdf_document = fitz.open()
            folder_name = os.path.basename(folder_path)
            output_pdf_file = os.path.join(folder_path, f"{folder_name}_merged.pdf")

            for pdf_file in pdf_files:
                pdf_file_path = os.path.join(folder_path, pdf_file)
                pdf_document.insert_pdf(fitz.open(pdf_file_path), from_page=0)
                os.remove(pdf_file_path)

            pdf_document.save(output_pdf_file)
            pdf_document.close()

    except Exception as e:
        print(f"Error merging PDFs in folder {folder_path}: {e}")
        if os.path.exists(output_pdf_file):
            os.remove(output_pdf_file)
            print(f"Removed incomplete merged file: '{output_pdf_file}'")

def scan_folders(processed_folders):
    try:
        root_folder = os.getcwd()
        subfolders = [f.path for f in os.scandir(root_folder) if f.is_dir()]

        valid_folders_found = False

        for folder_path in subfolders:
            pdf_files = get_files_with_extension(folder_path, '.pdf')
            jpg_files = get_files_with_extension(folder_path, '.jpg')

            if pdf_files or jpg_files:
                processed_folders.append(folder_path)
                valid_folders_found = True

        if not valid_folders_found:
            print("No valid files were found in any subfolder.")

    except Exception as e:
        print(f"Error scanning folders: {e}")

def list_files(folder_path):
    files = os.listdir(folder_path)
    files.sort(key=natural_sort_key)
    for file in files:
        print(file)

def main():
    try:
        pdf_files = find_valid_pdfs()
        processed_folders = []

        while True:
            try:
                user_choice = int(input("Choose an option:\n1. Extract Pages\n2. Convert to Image\n3. Convert and Treat Image\n4. Convert Images to PDF\n5. Merge PDFs\n8. Scan Folders\n9. List Folder Info\n0. Exit\n"))

                if user_choice == 1:
                    processed_folders = extract_pages(pdf_files)
                elif user_choice == 2:
                    if not processed_folders:
                        print("No folders processed yet.")
                    else:
                        convert_to_image(processed_folders)
                        print("Conversion to image complete.")
                elif user_choice == 3:
                    if not processed_folders:
                        print("No folders processed yet.")
                    else:
                        treat_images(processed_folders)
                        print("Image treatment complete.")
                elif user_choice == 4:
                    if not processed_folders:
                        print("No folders processed yet.")
                    else:
                        convert_image_to_pdf(processed_folders)
                        print("Conversion to PDF complete.")
                elif user_choice == 5:
                    if not processed_folders:
                        print("No folders processed yet.")
                    else:
                        merge_pdfs(processed_folders)
                        print("Merging PDFs complete.")
                elif user_choice == 8:
                    if not processed_folders:
                        scan_folders(processed_folders)
                        print("Scanning complete.")
                    else:
                        print("Active folders identified. Run option 9 for more information.")
                elif user_choice == 9:
                    if not processed_folders:
                        print("No folders processed yet.")
                    else:
                        list_folder_info(processed_folders)
                elif user_choice == 0:
                    print("Exiting the script.")
                    sys.exit()
                else:
                    print("Invalid option. Please try again.")
            except ValueError:
                print("Invalid input. Please enter a valid number.")

    except Exception as e:
        print(f"An error occurred: {e}")

    input("Press any key to quit...")

if __name__ == '__main__':
    main()

