import pdfplumber
from PyPDF2 import PdfReader, PdfWriter
import re
import os

def extract_text_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        return [(i, page.extract_text()) for i, page in enumerate(pdf.pages)]

def detect_fields(text_pages):
    pattern = re.compile(r'([A-Z \./]+)\s*:\s*([^\n]+)')
    fields = set()
    for _, text in text_pages:
        matches = pattern.findall(text)
        for key, _ in matches:
            fields.add(key.strip().upper())
    return sorted(fields)

def get_user_input(detected_fields):
    print("\nDetected searchable fields:\n")
    for idx, field in enumerate(detected_fields):
        print(f"{idx+1}. {field}")
    
    field_idx = int(input("\nEnter the number of the field you want to search by: ")) - 1
    value = input(f"Enter value to search in field '{detected_fields[field_idx]}': ").strip()
    return detected_fields[field_idx], value

def find_matching_pages(text_pages, search_field, search_value):
    matched_pages = []
    pattern = re.compile(fr"{re.escape(search_field)}\s*:\s*.*{re.escape(search_value)}", re.IGNORECASE)
    for page_num, text in text_pages:
        if pattern.search(text):
            matched_pages.append(page_num)
    return matched_pages

def create_output_pdf(source_pdf_path, pages, field, value):
    reader = PdfReader(source_pdf_path)
    writer = PdfWriter()
    for page_num in pages:
        writer.add_page(reader.pages[page_num])
    
    safe_value = value.replace("/", "-").replace(" ", "_")
    filename = f"Extracted_{field.replace(' ', '_')}_{safe_value}.pdf"
    with open(filename, "wb") as f:
        writer.write(f)
    print(f"\n✅ Extracted PDF saved as: {filename}")

def main():
    print("📄 PDF Certificate Extractor Tool")
    pdf_path = input("Enter path to PDF file (or drag & drop): ").strip().strip('"')

    if not os.path.exists(pdf_path):
        print("❌ File not found!")
        return

    text_pages = extract_text_from_pdf(pdf_path)
    detected_fields = detect_fields(text_pages)
    
    if not detected_fields:
        print("❌ No searchable fields found.")
        return

    search_field, search_value = get_user_input(detected_fields)
    matched_pages = find_matching_pages(text_pages, search_field, search_value)

    if not matched_pages:
        print("❌ No matches found for the given input.")
        return

    create_output_pdf(pdf_path, matched_pages, search_field, search_value)

if __name__ == "__main__":
    main()