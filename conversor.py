import os
import pandas as pd
from docx import Document
from PyPDF2 import PdfReader
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def convert_pdf_to_docx(pdf_file, docx_file):
    reader = PdfReader(pdf_file)
    doc = Document()
    for page in reader.pages:
        doc.add_paragraph(page.extract_text())
    doc.save(docx_file)

def convert_docx_to_pdf(docx_file, pdf_file):
    os.system(f'libreoffice --headless --convert-to pdf "{docx_file}" --outdir "{os.path.dirname(pdf_file)}"')

def convert_docx_to_txt(docx_file, txt_file):
    doc = Document(docx_file)
    with open(txt_file, 'w', encoding='utf-8') as f:
        for para in doc.paragraphs:
            f.write(para.text + '\n')

def convert_txt_to_docx(txt_file, docx_file):
    doc = Document()
    with open(txt_file, 'r', encoding='utf-8') as f:
        for line in f:
            doc.add_paragraph(line.strip())
    doc.save(docx_file)

def convert_xls_to_csv(xls_file, csv_file):
    df = pd.read_excel(xls_file)
    df.to_csv(csv_file, index=False)

def convert_csv_to_xls(csv_file, xls_file):
    df = pd.read_csv(csv_file)
    df.to_excel(xls_file, index=False)

def perform_conversion():
    input_file = input_file_path.get()
    output_file = output_file_path.get()
    conversion_type = conversion_type_var.get()

    if not input_file or not output_file:
        messagebox.showerror("Erro", "Por favor, selecione os arquivos de entrada e saída.")
        return

    try:
        if conversion_type == "PDF para DOCX":
            convert_pdf_to_docx(input_file, output_file)
            messagebox.showinfo("Sucesso", "Arquivo convertido de PDF para DOCX com sucesso!")
        elif conversion_type == "DOCX para PDF":
            convert_docx_to_pdf(input_file, output_file)
            messagebox.showinfo("Sucesso", "Arquivo convertido de DOCX para PDF com sucesso!")
        elif conversion_type == "DOCX para TXT":
            convert_docx_to_txt(input_file, output_file)
            messagebox.showinfo("Sucesso", "Arquivo convertido de DOCX para TXT com sucesso!")
        elif conversion_type == "TXT para DOCX":
            convert_txt_to_docx(input_file, output_file)
            messagebox.showinfo("Sucesso", "Arquivo convertido de TXT para DOCX com sucesso!")
        elif conversion_type == "XLS para CSV":
            convert_xls_to_csv(input_file, output_file)
            messagebox.showinfo("Sucesso", "Arquivo convertido de XLS para CSV com sucesso!")
        elif conversion_type == "CSV para XLS":
            convert_csv_to_xls(input_file, output_file)
            messagebox.showinfo("Sucesso", "Arquivo convertido de CSV para XLS com sucesso!")
        else:
            messagebox.showerror("Erro", "Tipo de conversão inválido.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

def select_input_file():
    file_path = filedialog.askopenfilename(title="Selecione o arquivo de entrada")
    input_file_path.set(file_path)

def select_output_file():
    file_path = filedialog.asksaveasfilename(title="Salvar como", defaultextension=".*")
    output_file_path.set(file_path)

# Configuração da interface gráfica
root = tk.Tk()
root.title("Conversor de Arquivos")

input_file_path = tk.StringVar()
output_file_path = tk.StringVar()
conversion_type_var = tk.StringVar()

# Layout
tk.Label(root, text="Arquivo de Entrada:").grid(row=0, column=0, padx=10, pady=10)
tk.Entry(root, textvariable=input_file_path, width=50).grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Selecionar", command=select_input_file).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Arquivo de Saída:").grid(row=1, column=0, padx=10, pady=10)
tk.Entry(root, textvariable=output_file_path, width=50).grid(row=1, column=1, padx=10, pady