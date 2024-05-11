"""
Modul ini menyediakan fungsi untuk mengonversi dokumen .docx menjadi PDF yang terenkripsi.
"""
import os
import re
import subprocess
from PyPDF2 import PdfReader, PdfWriter
import secrets
import string

def generate_high_security_password(length=24):
    """Generate a high-security password containing a mix of upper-case, lower-case, digits, and special characters."""
    # Define possible character sets
    alphabet_upper = string.ascii_uppercase
    alphabet_lower = string.ascii_lowercase
    digits = string.digits
    special_chars = "!@#$%^&*()_+-=[]{}|;:',.<>?/`~"

    # Combine all characters into one string
    all_chars = alphabet_upper + alphabet_lower + digits + special_chars
    
    # Ensure the password has at least one character from each set
    password = [
        secrets.choice(alphabet_upper),
        secrets.choice(alphabet_lower),
        secrets.choice(digits),
        secrets.choice(special_chars)
    ]
    
    # Fill the rest of the password length with a random choice of all available characters
    password += [secrets.choice(all_chars) for _ in range(length - 4)]
    
    # Shuffle the list to remove any predictable patterns and convert to string
    secrets.SystemRandom().shuffle(password)
    return ''.join(password)

def load_mappings(file_path):
    """Load renaming mappings from a file."""
    phrase_map = {}
    with open(file_path, 'r') as file:
        for line in file:
            parts = line.strip().split(',')
            if len(parts) == 2:
                phrase, abbreviation = parts
                # Use case-insensitive regex pattern
                phrase_map[r'\b' + re.escape(phrase.strip().lower()) + r'\b'] = abbreviation.strip()
    return phrase_map

def rename_file(file, phrase_map):
    """Rename the file based on loaded mappings, checking both full phrases and key terms."""
    file_lower = file.lower()  # Normalize file name to lower case for case insensitive comparison

    # Check full phrase mappings
    for pattern, abbreviation in phrase_map.items():
        if re.search(pattern, file_lower):
            return abbreviation + '.pdf'

    # If no full match, check for embedded key terms from the same mapping
    key_list = sorted(phrase_map.values(), key=len, reverse=True)
    for key in key_list:
        if re.search(r'\b' + re.escape(key.lower()) + r'\b', file_lower):
            return key + '.pdf'

    return "ALL Soal.pdf"  # Default return if no specific pattern is matc

def run_pandoc(command):
    """
    Attempt to run a pandoc conversion command.
    """
    try:
        result = subprocess.run(command, check=True, capture_output=True)
        print("Successfully converted:", result.stdout.decode())
        return True
    except subprocess.CalledProcessError as e:
        print("Error converting:", e.stderr.decode())
        return False

def convert_with_libreoffice(doc_path, pdf_path):
    """
    Convert a document using LibreOffice when Pandoc fails.
    Ensures that the output PDF is named correctly and attempts to maintain a standard format with a white background and black text, although explicit formatting control may require additional LibreOffice configuration or scripting.
    """
    libreoffice_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    output_dir = os.path.dirname(pdf_path)
    temp_pdf_path = os.path.join(output_dir, os.path.splitext(os.path.basename(doc_path))[0] + ".pdf")

    command = [
        libreoffice_path, '--headless', '--convert-to', 'pdf:writer_pdf_Export',
        '--outdir', output_dir, doc_path
    ]

    try:
        subprocess.run(command, check=True)
        # After conversion, rename the default output file if it exists
        if os.path.exists(temp_pdf_path):
            os.rename(temp_pdf_path, pdf_path)
            print(f"Successfully converted with LibreOffice and renamed to: {pdf_path}")
        else:
            print(f"Conversion succeeded but expected file {temp_pdf_path} not found.")
            return False
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error converting with LibreOffice: {e}")
        return False
    
def decrypt_pdf(reader, password):
    """
    Attempt to decrypt a PDF file if it is encrypted.
    """
    if reader.is_encrypted:
        if reader.decrypt(password) != 1:
            print(f"Failed to decrypt the PDF with provided password: {password}")
            return False
    return True

def process_pdf_file(pdf_path, password):
    """
    Process the PDF file: decrypt, add pages, and encrypt.
    """
    try:
        reader = PdfReader(pdf_path)
        if not decrypt_pdf(reader, password):
            return False

        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        writer.encrypt(password)

        with open(pdf_path, "wb") as f:
            writer.write(f)
        print(f"Processed and encrypted '{pdf_path}' with password: {password}")
        return True
    except Exception as e:
        print(f"Failed to process PDF {pdf_path}: {e}")
        return False

def convert_docs_to_pdf(source_folder, output_folder, password_log_path, mapping_file):
    """
    Convert all docx files in the source folder to encrypted PDFs, with logging.
    """
    phrase_map = load_mappings(mapping_file)
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    with open(password_log_path, "w") as password_file:
        for root, dirs, files in os.walk(source_folder):
            for file in files:
                if file.endswith(".docx") and not file.startswith("~$"):
                    doc_path = os.path.join(root, file)
                    new_filename = rename_file(file, phrase_map)
                    pdf_path = os.path.join(output_folder, new_filename)
                    password = generate_high_security_password(24)
                    command = [
                        "pandoc", doc_path, "--pdf-engine=xelatex",
                        "-V", "mainfont=DejaVu Serif",
                        "-V", "colorlinks=true",
                        "-V", "urlcolor=black",
                        "-V", "linkcolor=black",
                        "--citeproc", "--from=docx", "-o", pdf_path
                    ]

                    if not run_pandoc(command):
                        print("Pandoc conversion failed, trying with LibreOffice...")
                        if not convert_with_libreoffice(doc_path, pdf_path):
                            print("LibreOffice conversion also failed.")
                            continue

                    if not process_pdf_file(pdf_path, password):
                        continue

                    password_file.write(f'{new_filename} = "{password}"\n')
                    print(f"Converted and encrypted '{file}' to '{pdf_path}' with password: {password}")


SOURCE = "/Users/adityaramadhan/python/pdfco/source"
OUTPUT = "/Users/adityaramadhan/python/pdfco/output"
PASSWORD_LOG_PATH = "/Users/adityaramadhan/python/pdfco/passwords.txt"
MAPPING_FILE = "/Users/adityaramadhan/python/pdfco/mappings.txt"
convert_docs_to_pdf(SOURCE, OUTPUT, PASSWORD_LOG_PATH, MAPPING_FILE)

