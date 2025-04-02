import os
import pandas as pd
from docx import Document
from PyPDF2 import PdfReader
from tkinter import Tk, filedialog, messagebox

def convert_pdf_to_docx(pdf_file, docx_file):
    reader = PdfReader(pdf_file)
    doc = Document()
    for page in reader.pages:
        doc.add_paragraph(page.extract_text())
    doc.save(docx_file)

def convert_docx_to_pdf(docx_file, pdf_file):
    # Para converter DOCX para PDF, você pode usar um software externo como o LibreOffice
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

def main():
    Tk().withdraw()  # Oculta a janela principal do Tkinter
    while True:
        print("Conversor de Arquivos")
        print("1. PDF para DOCX")
        print("2. DOCX para PDF")
        print("3. DOCX para TXT")
        print("4. TXT para DOCX")
        print("5. XLS para CSV")
        print("6. CSV para XLS")
        print("0. Sair")
        
        choice = input("Escolha uma opção: ")
        
        if choice == '0':
            break
        
        file_path = filedialog.askopenfilename(title="Selecione o arquivo")
        
        if not file_path:
            messagebox.showerror("Erro", "Nenhum arquivo selecionado.")
            continue
        
        if choice == '1':
            output_file = filedialog.asksaveasfilename(defaultextension=".docx", title="Salvar como DOCX")
            convert_pdf_to_docx(file_path, output_file)
            messagebox.showinfo("Sucesso", "Arquivo convertido com sucesso!")
        
        elif choice == '2':
            output_file = filedialog.asksaveasfilename(defaultextension=".pdf", title="Salvar como PDF")
            convert_docx_to_pdf(file_path, output_file)
            messagebox.showinfo("Sucesso", "Arquivo convertido com sucesso!")
        
        elif choice == '3':
            output_file = filedialog.asksaveasfilename(defaultextension=".txt", title="Salvar como TXT")
            convert_docx_to_txt(file_path, output_file)
            messagebox.showinfo("Sucesso", "Arquivo convertido com sucesso!")
        
        elif choice == '4':
            output_file = filedialog.asksaveasfilename(defaultextension=".docx", title="Salvar como DOCX")
            convert_txt_to_docx(file_path, output_file)
            messagebox.showinfo("Sucesso", "Arquivo convertido com sucesso!")
        
        elif choice == '5':
            output_file = filedialog.asksaveasfilename(defaultextension=".csv", title="Salvar como CSV")
            convert_xls_to_csv(file_path, output_file)
            messagebox.showinfo("Sucesso", "Arquivo convertido com sucesso!")
        
        elif choice == '6':
            output_file = filedialog.asksaveasfilename(defaultextension=".xls", title="Salvar como XLS")
            convert_csv_to_xls(file_path, output_file)
            messagebox.showinfo("Sucesso", "Arquivo convertido com sucesso!")
        
        else:
            messagebox.showerror("Erro", "Opção inválida.")

if __name__ == "__main__":
    main()