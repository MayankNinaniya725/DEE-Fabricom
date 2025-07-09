import os
import re
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber
import pandas as pd
from pdf2image import convert_from_path
import pytesseract

# Appearance
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

FIELDS = ["FLANGE NO", "HEAT NO", "PLATE NO", "PRODUCT NO", "PART NO", "TEST CERTIFICATE NO"]

def extract_text_with_ocr(pdf_path, page_num):
    images = convert_from_path(pdf_path, first_page=page_num + 1, last_page=page_num + 1, dpi=300)
    if images:
        return pytesseract.image_to_string(images[0])
    return ""

def extract_fields_from_text(text):
    def search(pattern):
        match = re.search(pattern, text, re.IGNORECASE)
        return match.group(1).strip() if match else "NA"

    return {
        "FLANGE NO": search(r"FLANGE\s*NO[:\-]?\s*([A-Z0-9\-\/]+)"),
        "HEAT NO": search(r"HEAT\s*NO[:\-]?\s*([A-Z0-9\-\/]+)"),
        "PLATE NO": search(r"PLATE\s*NO[:\-]?\s*([A-Z0-9\-\/]+)"),
        "PRODUCT NO": search(r"PRODUCT\s*NO[:\-]?\s*([A-Z0-9\-\/]+)"),
        "PART NO": search(r"PART\s*NO[:\-]?\s*([A-Z0-9\-\/]+)"),
        "TEST CERTIFICATE NO": search(r"TEST\s+CERTIFICATE\s+NO[:\-]?\s*([A-Z0-9\-\/]+)")
    }

def extract_pdf_by_field(pdf_path, field, value, output_base="output", progress_callback=None):
    results = []
    writer = PdfWriter()
    folder_name = f"{field.replace(' ', '')}_{value.replace('/', '-')}"
    output_dir = os.path.join(output_base, folder_name)
    os.makedirs(output_dir, exist_ok=True)

    reader = PdfReader(pdf_path)
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            try:
                text = page.extract_text()
            except:
                text = None
            if not text or len(text.strip()) < 20:
                text = extract_text_with_ocr(pdf_path, i)

            if value.lower() in text.lower():
                fields = extract_fields_from_text(text)
                results.append({"Page": i + 1, **fields})
                writer.add_page(reader.pages[i])

            if progress_callback:
                progress_callback(i + 1, total)

    if results:
        combined_filename = f"{field.lower().replace(' ', '')}-{value}.pdf"
        combined_path = os.path.join(output_dir, combined_filename)
        with open(combined_path, "wb") as f:
            writer.write(f)

        df = pd.DataFrame(results)
        df.to_excel(os.path.join(output_dir, "summary.xlsx"), index=False)

    return output_dir, len(results)

class PDFExtractorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("PDF Field-Based Extractor")
        self.geometry("650x420")
        self.resizable(False, False)

        self.pdf_path = ""
        self.output_folder = ""

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=3)
        self.grid_columnconfigure(5, weight=1)

        # PDF file
        ctk.CTkLabel(self, text="PDF File:").grid(row=0, column=0, padx=15, pady=10, sticky="e")
        self.file_entry = ctk.CTkEntry(self, width=400)
        self.file_entry.grid(row=0, column=1, padx=5, pady=5)
        ctk.CTkButton(self, text="Browse", command=self.browse_pdf).grid(row=0, column=2, padx=10)

        # Field
        ctk.CTkLabel(self, text="Field:").grid(row=1, column=0, padx=15, pady=10, sticky="e")
        self.field_combo = ctk.CTkComboBox(self, values=FIELDS, width=400)
        self.field_combo.grid(row=1, column=1, padx=5, pady=5, columnspan=2, sticky="w")
        self.field_combo.set(FIELDS[0])

        # Search Value
        ctk.CTkLabel(self, text="Search Value:").grid(row=2, column=0, padx=15, pady=10, sticky="e")
        self.value_entry = ctk.CTkEntry(self, width=400)
        self.value_entry.grid(row=2, column=1, padx=5, pady=5, columnspan=2, sticky="w")

        # Output Folder
        ctk.CTkLabel(self, text="Output Folder:").grid(row=3, column=0, padx=15, pady=10, sticky="e")
        self.output_entry = ctk.CTkEntry(self, width=400)
        self.output_entry.grid(row=3, column=1, padx=5, pady=5)
        ctk.CTkButton(self, text="Browse", command=self.browse_output_folder).grid(row=3, column=2, padx=10)

        # Spinner / Loading indicator
        self.spinner = ctk.CTkLabel(self, text="", text_color="gray")
        self.spinner.grid(row=4, column=0, columnspan=3, pady=10)

        # Extract button (centered)
        self.extract_button = ctk.CTkButton(self, text="Extract & Save", command=self.run_extraction)
        self.extract_button.grid(row=5, column=0, columnspan=3, pady=20)

    def browse_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if path:
            self.pdf_path = path
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, path)

    def browse_output_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.output_folder = path
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, path)

    def update_progress(self, current, total):
        pass

    def run_extraction(self):
        if not self.pdf_path:
            messagebox.showerror("Error", "Please select a PDF file.")
            return
        if not self.output_folder:
            messagebox.showerror("Error", "Please select an output folder.")
            return

        field = self.field_combo.get()
        value = self.value_entry.get().strip()
        if not value:
            messagebox.showerror("Error", "Please enter a value to search.")
            return

        self.spinner.configure(text="🔄 Extracting... Please wait.")
        self.update_idletasks()

        # Start in a new thread
        threading.Thread(target=self._threaded_extraction, args=(field, value), daemon=True).start()

    def _threaded_extraction(self, field, value):
        out_dir, count = extract_pdf_by_field(
            self.pdf_path,
            field,
            value,
            output_base=self.output_folder,
            progress_callback=self.update_progress
        )
        self.after(100, self._on_extraction_complete, out_dir, count)

    def _on_extraction_complete(self, out_dir, count):
        self.spinner.configure(text="")  # Stop loading
        if count == 0:
            messagebox.showinfo("Result", "No matching pages found.")
        else:
            messagebox.showinfo("Success", f"Extracted {count} pages.\nSaved in:\n{out_dir}")

if __name__ == "__main__":
    app = PDFExtractorApp()
    app.mainloop()
# This code is a GUI application for extracting specific fields from PDF files.
# It allows users to select a PDF, specify a field and value to search for, and
# saves the extracted pages and a summary in an output folder.
