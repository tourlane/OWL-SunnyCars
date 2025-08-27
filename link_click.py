import os
import re
import requests
import fitz  # PyMuPDF
import openpyxl
from datetime import datetime

LOG_FILE = 'download_log.txt'
duplicated_files = []

def log(message):
    timestamp = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    full_message = f"{timestamp} {message}"
    print(full_message)
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(full_message + '\n')





# Folders
ATTACHMENTS_DIR = 'attachments'
PDF_DOWNLOAD_DIR = 'pdfs'
os.makedirs(PDF_DOWNLOAD_DIR, exist_ok=True)

# Extract links from .txt and .csv files
def extract_links_from_text(file_path):
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        text = f.read()
    return re.findall(r'https?://\S+', text)

# Extract links from .xlsx files
def extract_links_from_excel(file_path):
    links = []
    wb = openpyxl.load_workbook(file_path, data_only=True)
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(values_only=True):
            for cell in row:
                if isinstance(cell, str) and 'http' in cell:
                    matches = re.findall(r'https?://\S+', cell)
                    links.extend(matches)
    return links

# Extract links from .pdf files
def extract_links_from_pdf(file_path):
    links = []
    doc = fitz.open(file_path)
    for page in doc:
        for link in page.get_links():
            if 'uri' in link:
                links.append(link['uri'])
    return links

downloaded_files = set()  # To track duplicates
failed_links = []

def download_from_link(url, output_folder, source_file):
    try:
        response = requests.get(url, timeout=15)
        response.raise_for_status()

        if 'application/pdf' not in response.headers.get('Content-Type', ''):
            print(f"[SKIPPED] Not a PDF: {url}")
            return

        # Get filename
        content_disp = response.headers.get('content-disposition', '')
        if 'filename=' in content_disp:
            filename = content_disp.split('filename=')[1].strip('"')
        else:
            filename = url.split('/')[-1].split('?')[0] or 'file.pdf'

        filepath = os.path.join(output_folder, filename)

        # Check for duplicates
        if filename in downloaded_files:
            print(f"[DUPLICATE] Already downloaded: {filename} (from {source_file})")
            duplicated_files.append((filename, source_file))
            return
        downloaded_files.add(filename)

        with open(filepath, 'wb') as f:
            f.write(response.content)
        log(f"[DOWNLOADED] {filename} (from {source_file})")
    except Exception as e:
        print(f"[FAILED] {url} (from {source_file}) — {e}")
        failed_links.append((url, source_file, str(e)))


def process_attachments():
    total_links = 0
    for fname in os.listdir(ATTACHMENTS_DIR):
        if fname.startswith("~$"):  # Skip temp/lock files
            continue
        path = os.path.join(ATTACHMENTS_DIR, fname)
        print(f"\n>>> Processing: {fname}")

        if fname.endswith(('.txt', '.csv')):
            links = extract_links_from_text(path)
        elif fname.endswith('.pdf'):
            links = extract_links_from_pdf(path)
        elif fname.endswith('.xlsx'):
            links = extract_links_from_excel(path)
        else:
            print(f"[SKIPPED] Unsupported file: {fname}")
            continue

        print(f"Found {len(links)} link(s) in {fname}")
        total_links += len(links)

        for link in links:
            download_from_link(link, PDF_DOWNLOAD_DIR, fname)

    print("\n=== SUMMARY ===")
    print(f"Total files in '{ATTACHMENTS_DIR}': {len(os.listdir(ATTACHMENTS_DIR))}")
    print(f"Total links found: {total_links}")
    print(f"Total unique PDFs downloaded: {len(downloaded_files)}")
    print(f"Failed downloads: {len(failed_links)}")

    if failed_links:
        print("\n--- Failed URLs ---")
        for url, source, err in failed_links:
            print(f"{url} (from {source}) => {err}")

    if duplicated_files:
        print("\n--- Duplicated PDFs ---")
        for filename, source in duplicated_files:
            print(f"{filename} (from {source})")


if __name__ == '__main__':
    process_attachments()
