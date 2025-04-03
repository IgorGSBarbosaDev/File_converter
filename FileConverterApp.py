import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from FileConverter import FileConverter


class FileConverterApp:
    """Aplicativo GUI para conversão de arquivos."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Conversor de Arquivos")
        self.root.configure(bg="#f0f0f0")

        self.input_file_path = tk.StringVar()
        self.conversion_type_var = tk.StringVar()

        self.setup_ui()

    def setup_ui(self):
        """Configuração da interface gráfica."""
        # Configuração do estilo
        style = ttk.Style()
        style.configure("TLabel", background="#f0f0f0", font=("Arial", 10))
        style.configure("TButton", font=("Arial", 10))
        style.configure("TCombobox", font=("Arial", 10))

        header = tk.Label(self.root, text="Conversor de Arquivos", font=("Arial", 16, "bold"), bg="#f0f0f0")
        header.grid(row=0, column=0, columnspan=3, pady=10)

        tk.Label(self.root, text="Arquivo de Entrada:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
        tk.Entry(self.root, textvariable=self.input_file_path, width=50).grid(row=1, column=1, padx=10, pady=10)
        tk.Button(self.root, text="Selecionar", command=self.select_input_file).grid(row=1, column=2, padx=10, pady=10)

        tk.Label(self.root, text="Tipo de Conversão:").grid(row=2, column=0, padx=10, pady=10, sticky="e")
        conversion_dropdown = ttk.Combobox(self.root, textvariable=self.conversion_type_var, state="readonly")
        conversion_dropdown['values'] = list(FileConverter.conversion_functions.keys())
        conversion_dropdown.grid(row=2, column=1, padx=10, pady=10)

        convert_button = tk.Button(self.root, text="Converter", command=self.perform_conversion_threaded, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
        convert_button.grid(row=3, column=1, pady=20)

        progress = ttk.Progressbar(self.root, orient="horizontal", length=300, mode="indeterminate")
        progress.grid(row=4, column=1, pady=10)

    def select_input_file(self):
        """Abre uma caixa de diálogo para selecionar o arquivo de entrada."""
        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo de entrada",
            filetypes=[("Todos os arquivos", "*.*"), ("PDF Files", "*.pdf"), ("DOCX Files", "*.docx"), ("TXT Files", "*.txt"), ("XLS Files", "*.xls"), ("CSV Files", "*.csv")]
        )
        self.input_file_path.set(file_path)

    def perform_conversion(self):
        """Garante que o arquivo de entrada e o tipo de conversão sejam válidos e chama a função de conversão apropriada."""
        input_file = self.input_file_path.get()
        conversion_type = self.conversion_type_var.get()

        if not input_file:
            messagebox.showerror("Erro", "Por favor, selecione o arquivo de entrada.")
            return

        if not os.path.exists(input_file):
            messagebox.showerror("Erro", "O arquivo de entrada não existe.")
            return

        try:
            conversion_function, output_extension = FileConverter.conversion_functions.get(conversion_type, (None, None))
            if conversion_function:
                base_name, _ = os.path.splitext(input_file)
                output_file = base_name + output_extension
                conversion_function(input_file, output_file)
                messagebox.showinfo("Sucesso", f"Arquivo convertido com sucesso! Salvo como: {output_file}")
            else:
                messagebox.showerror("Erro", "Tipo de conversão inválido.")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

    def perform_conversion_threaded(self):
        """Executa a conversão em uma thread separada."""
        threading.Thread(target=self.perform_conversion).start()

    def run(self):
        """Executa o loop principal do aplicativo."""
        self.root.mainloop()