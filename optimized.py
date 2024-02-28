import cv2
import pytesseract
from pdf2image import convert_from_path
import numpy as np
import re
import os
import subprocess
import platform
import time
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor

def clear_screen():
    if platform.system() == "Windows":
        os.system('cls')
    else:
        os.system('clear')

# Display welcome message and copyright notice
clear_screen()
print("ECSSA PDF CONVERTER")
print("Copyright © 2024 ECSSA. All rights reserved.")
print("This software is provided 'as-is', without any express or implied warranty.")
print("In no event will the authors be held liable for any damages arising from the use of this software.\n")
time.sleep(2)  # Pause for 2 seconds for the user to see the message

# Check if Tesseract is installed
print("Checking if Tesseract is installed...")
try:
    # If you don't have tesseract executable in your PATH, include the following:
    pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
    tesseract_version = subprocess.check_output([pytesseract.pytesseract.tesseract_cmd, '--version'], stderr=subprocess.STDOUT)
    print("Tesseract is installed.")
except FileNotFoundError:
    print("Tesseract is not installed or it's not in the system's PATH.")
    exit(1)

time.sleep(2)  # Pause for 2 seconds for the user to see the message

# Check if Poppler is installed
print("Checking if Poppler is installed...")
poppler_path = r'C:\\Release-23.11.0-0\\poppler-23.11.0\\Library\\bin'
if os.path.isdir(poppler_path):
    print("Poppler is installed.")
else:
    print("Poppler is not installed or the provided path is incorrect.")
    print("Please install Poppler and place it in the C: drive.")
    exit(1)

time.sleep(2)  # Pause for 2 seconds for the user to see the message

# Ask user for the text file path
text_file_path = input("Enter the path where you want to save the text file: ").strip('"')

# Function to process a page/image
def process_page(img_page_number):
    img, page_number = img_page_number
    # Convert image to grayscale
    img_gray = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)

    # Apply adaptive thresholding
    img_bin = cv2.adaptiveThreshold(img_gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)

    # Use Tesseract to extract text
    custom_config = r'--oem 3 --psm 6'
    text = pytesseract.image_to_string(img_bin, config=custom_config, lang='eng+spa')

    # Store the text in the dictionary
    ocr_output[page_number] = text
    # Update the progress bar
    pbar.update(1)

def load_new_pdf():
    global pdf_path, images, ocr_output

    # Ask user for the new PDF file path
    pdf_path = input("Enter the path to the new PDF file: ").strip('"')

    # Convert PDF to images
    images = convert_from_path(pdf_path, poppler_path=poppler_path)

    # Create a dictionary to store the OCR output
    ocr_output = {}

    # Process images with tqdm progress bar and multithreading
    with tqdm(total=len(images), desc="Processing pages") as pbar:
        # Use ThreadPoolExecutor to limit the number of threads
        with ThreadPoolExecutor(max_workers=8) as executor:
            # Map the process_page function to the pool of threads
            executor.map(process_page, [(img, i+1) for i, img in enumerate(images)])
    print("PDF loaded and processed successfully!")
    input("Press Enter to continue...")

# Initial processing of the first PDF
pdf_path = input("Enter the path to the PDF file: ").strip('"')
images = convert_from_path(pdf_path, poppler_path=poppler_path)
ocr_output = {}

# Process images with tqdm progress bar and multithreading
with tqdm(total=len(images), desc="Processing pages") as pbar:
    # Use ThreadPoolExecutor to limit the number of threads
    with ThreadPoolExecutor(max_workers=8) as executor:
        # Map the process_page function to the pool of threads
        executor.map(process_page, [(img, i+1) for i, img in enumerate(images)])

def convert_pdf_to_text():
    # Ask user for the page numbers to convert
    pages_to_convert = input("Enter the page numbers to convert, separated by commas: ")
    pages_to_convert = list(map(int, pages_to_convert.split(',')))

    # Write the OCR output to a .txt file
    filename = input("Enter a filename for the output (without extension): ")
    with open(f'{text_file_path}/{filename}.txt', 'w', encoding='utf-8') as f:
        for page in pages_to_convert:
            text = ocr_output.get(page)
            if text:
                f.write(f'Page {page}:\n{text}\n\n')
    print(f"Pages {' '.join(map(str, pages_to_convert))} converted to text, and saved as {filename}.txt")
    input("Press Enter to continue...")

def search_word_in_pdf():
    clear_screen()
    print("Search Options:")
    print("1. Search for dates and hours")
    print("2. Search for specific words")
    search_option = input("Enter your choice (1 or 2): ")

    if search_option == '1':
        return search_dates_and_hours()
    elif search_option == '2':
        return search_specific_words()
    else:
        return "Invalid choice. Please enter 1 or 2."

def search_dates_and_hours():
    # Define the patterns for date and time
    date_pattern = r'\b\d{1,2}\s*de\s*[a-zA-Z]+\s*de\s*\d{4}\b'
    time_pattern = r'\b([01]?[0-9]|2[0-3]):[0-5][0-9]\b'

    # Create a file to store the search results
    with open(f'{text_file_path}/dates.txt', 'w', encoding='utf-8') as f:
        for page, text in ocr_output.items():
            # Split the text into sentences
            sentences = text.split('.')
            for sentence in sentences:
                # Search for date and time in the sentence
                date_match = re.search(date_pattern, sentence)
                time_match = re.search(time_pattern, sentence)
                
                # If a date or time is found, write the sentence to the file
                if date_match or time_match:
                    f.write(f"Found on page {page}:\n{sentence}\n")
    
    return "Dates and hours searched. Results saved to 'dates.txt'."

def search_specific_words():
    # Ask user for the words to search
    search_words = input("Enter the words to search, separated by commas: ")
    search_words = [word.strip() for word in search_words.split(',')]

    # Create a file to store the search results
    with open(f'{text_file_path}/Search.txt', 'w', encoding='utf-8') as f:
        for page, text in ocr_output.items():
            # Split the text into sentences
            sentences = text.split('.')
            for sentence in sentences:
                # Search for each word in the sentence
                for word in search_words:
                    if word in sentence:
                        f.write(f"'{word}' found on page {page}:\n{sentence}\n")

    return "Specific words searched. Results saved to 'Search.txt'."

def main():
    while True:
        clear_screen()
        print("1. Convert PDF to text")
        print("2. Search word in PDF")
        print("3. Load new PDF")
        print("4. Exit")
        choice = input("Enter your choice: ")

        if choice == '1':
            convert_pdf_to_text()
        elif choice == '2':
            message = search_word_in_pdf()
            print(message)
            input("Press Enter to continue...")
        elif choice == '3':
            load_new_pdf()
        elif choice == '4':
            break
        else:
            print("Invalid choice. Please enter 1, 2, 3, or 4.")

if __name__ == "__main__":
    main()
