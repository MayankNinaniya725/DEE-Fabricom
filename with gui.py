import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber
from pdf2image import convert_from_path
import pytesseract
from PIL import Image

# Fields to extract
FIELDS = ["HEAT NO", "PLATE NO", "TEST CERTIFICATE NO"]

# Optional: set Tesseract path if needed
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def extract_text_with_ocr(pdf_path, page_num):
    images = convert_from_path(pdf_path, first_page=page_num+1, last_page=page_num+1, dpi=300)
    if images:
        return pytesseract.image_to_string(images[0])
    return ""

def extract_all_pages(pdf_path):
    data = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            try:
                text = page.extract_text()
            except:
                text = None
            if not text or len(text.strip()) < 20:
                text = extract_text_with_ocr(pdf_path, i)
            data.append((i, text))
    return data

def find_matching_pages(text_pages, value):
    matched = []
    for page_num, text in text_pages:
        lines = text.split('\n')
        for line in lines:
            if value.lower() in line.lower():
                matched.append(page_num)
                break
    return matched

def save_single_page_pdf(source_path, page_num, fields):
    reader = PdfReader(source_path)
    writer = PdfWriter()
    writer.add_page(reader.pages[page_num])

    safe_name = "_".join(fields.get(key, "NA").replace("/", "-") for key in [
        "FLANGE NO", "HEAT NO", "PLATE NO", "PRODUCT NO", "PART NO", "TEST CERTIFICATE NO"
    ])
    default_name = f"{safe_name}.pdf"

    output_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        initialfile=default_name,
        filetypes=[("PDF files", "*.pdf")],
        title="Save Page PDF As"
    )

    if not output_path:
        return None
    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path


# GUI App
class PDFExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Certificate Extractor")
        self.pdf_path = ""

        tk.Label(root, text="PDF File:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.file_entry = tk.Entry(root, width=50)
        self.file_entry.grid(row=0, column=1)
        tk.Button(root, text="Browse", command=self.browse_file).grid(row=0, column=2)

        tk.Label(root, text="Search Field:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
        self.field_combo = Combobox(root, values=FIELDS, state="readonly")
        self.field_combo.grid(row=1, column=1)
        self.field_combo.current(0)

        tk.Label(root, text="Search Value:").grid(row=2, column=0, padx=10, pady=10, sticky="e")
        self.value_entry = tk.Entry(root, width=50)
        self.value_entry.grid(row=2, column=1)

        tk.Button(root, text="Extract & Save", command=self.extract_and_save).grid(row=3, column=1, pady=20)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.pdf_path = file_path
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)

    def extract_and_save(self):
        if not self.pdf_path:
            messagebox.showerror("Error", "Please select a PDF file.")
            return

        field = self.field_combo.get()
        value = self.value_entry.get().strip()
        if not value:
            messagebox.showerror("Error", "Please enter a value to search.")
            return

        text_pages = extract_all_pages(self.pdf_path)
        matched_pages = find_matching_pages(text_pages, value)

        if not matched_pages:
            messagebox.showinfo("No Match", "No matching pages found.")
            return

        output_path = save_matched_pdf(self.pdf_path, matched_pages, value)
        if output_path:
            messagebox.showinfo("Success", f"✅ Extracted PDF saved at:\n{output_path}")

# Run the GUI
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFExtractorGUI(root)
    root.mainloop()