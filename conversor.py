import os
import pandas as pd
from docx import Document
from PyPDF2 import PdfReader
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import logging

# Configure logging
logging.basicConfig(filename="conversion.log", level=logging.INFO, format="%(asctime)s - %(message)s")

# Conversion functions
def convert_pdf_to_docx(pdf_file, docx_file):
    try:
        reader = PdfReader(pdf_file)
        doc = Document()
        for page in reader.pages:
            doc.add_paragraph(page.extract_text())
        doc.save(docx_file)
    except Exception as e:
        logging.error(f"Error converting PDF to DOCX: {str(e)}")
        raise

def convert_docx_to_pdf(docx_file, pdf_file):
    try:
        os.system(f'libreoffice --headless --convert-to pdf "{docx_file}" --outdir "{os.path.dirname(pdf_file)}"')
    except Exception as e:
        logging.error(f"Error converting DOCX to PDF: {str(e)}")
        raise

def convert_docx_to_txt(docx_file, txt_file):
    try:
        doc = Document(docx_file)
        with open(txt_file, 'w', encoding='utf-8') as f:
            for para in doc.paragraphs:
                f.write(para.text + '\n')
    except Exception as e:
        logging.error(f"Error converting DOCX to TXT: {str(e)}")
        raise

def convert_txt_to_docx(txt_file, docx_file):
    try:
        doc = Document()
        with open(txt_file, 'r', encoding='utf-8') as f:
            for line in f:
                doc.add_paragraph(line.strip())
        doc.save(docx_file)
    except Exception as e:
        logging.error(f"Error converting TXT to DOCX: {str(e)}")
        raise

def convert_xls_to_csv(xls_file, csv_file):
    try:
        df = pd.read_excel(xls_file)
        df.to_csv(csv_file, index=False)
    except Exception as e:
        logging.error(f"Error converting XLS to CSV: {str(e)}")
        raise

def convert_csv_to_xls(csv_file, xls_file):
    try:
        df = pd.read_csv(csv_file)
        df.to_excel(xls_file, index=False)
    except Exception as e:
        logging.error(f"Error converting CSV to XLS: {str(e)}")
        raise

# Dictionary for conversion functions
conversion_functions = {
    "PDF para DOCX": (convert_pdf_to_docx, ".docx"),
    "DOCX para PDF": (convert_docx_to_pdf, ".pdf"),
    "DOCX para TXT": (convert_docx_to_txt, ".txt"),
    "TXT para DOCX": (convert_txt_to_docx, ".docx"),
    "XLS para CSV": (convert_xls_to_csv, ".csv"),
    "CSV para XLS": (convert_csv_to_xls, ".xls"),
}

# Perform conversion with threading
def perform_conversion():
    input_file = input_file_path.get()
    conversion_type = conversion_type_var.get()

    if not input_file:
        messagebox.showerror("Erro", "Por favor, selecione o arquivo de entrada.")
        return

    if not os.path.exists(input_file):
        messagebox.showerror("Erro", "O arquivo de entrada não existe.")
        return

    try:
        conversion_function, output_extension = conversion_functions.get(conversion_type, (None, None))
        if conversion_function:
            # Generate output file name
            base_name, _ = os.path.splitext(input_file)
            output_file = base_name + output_extension

            # Perform conversion
            conversion_function(input_file, output_file)
            logging.info(f"Converted {input_file} to {output_file} ({conversion_type})")
            messagebox.showinfo("Sucesso", f"Arquivo convertido com sucesso! Salvo como: {output_file}")
        else:
            messagebox.showerror("Erro", "Tipo de conversão inválido.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

def perform_conversion_threaded():
    threading.Thread(target=perform_conversion).start()

# File selection functions
def select_input_file():
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo de entrada",
        filetypes=[("Todos os arquivos", "*.*"), ("PDF Files", "*.pdf"), ("DOCX Files", "*.docx"), ("TXT Files", "*.txt"), ("XLS Files", "*.xls"), ("CSV Files", "*.csv")]
    )
    input_file_path.set(file_path)

# GUI Configuration
root = tk.Tk()
root.title("Conversor de Arquivos")

input_file_path = tk.StringVar()
conversion_type_var = tk.StringVar()

# Layout
tk.Label(root, text="Arquivo de Entrada:").grid(row=0, column=0, padx=10, pady=10)
tk.Entry(root, textvariable=input_file_path, width=50).grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Selecionar", command=select_input_file).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Tipo de Conversão:").grid(row=1, column=0, padx=10, pady=10)
conversion_dropdown = ttk.Combobox(root, textvariable=conversion_type_var, state="readonly")
conversion_dropdown['values'] = list(conversion_functions.keys())
conversion_dropdown.grid(row=1, column=1, padx=10, pady=10)

tk.Button(root, text="Converter", command=perform_conversion_threaded).grid(row=2, column=1, pady=20)

# Progress bar
progress = ttk.Progressbar(root, orient="horizontal", length=300, mode="indeterminate")
progress.grid(row=3, column=1, pady=10)

root.mainloop()