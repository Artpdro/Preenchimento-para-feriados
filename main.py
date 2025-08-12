
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

# Importar a função de preenchimento e o mapeamento
from pdf_filler import fill_pdf_document
from pdf_mapping import pdf_fields

co0 = "#2C3E50"  # Azul escuro
co1 = "#FFFFFF"  # Branco
co2 = "#3498DB"  # Azul moderno
co3 = "#27AE60"  # Verde sucesso
co4 = "#E74C3C"  # Vermelho erro
co5 = "#ECF0F1"  # Cinza claro
co6 = "#34495E"  # Cinza escuro
co7 = "#F39C12"  # Laranja
co8 = "#9B59B6"  # Roxo
co9 = "#1ABC9C"  # Verde agua

class PDFillerApp:
    def __init__(self, master):
        self.master = master
        master.title("Preenchedor de PDF")
        master.geometry("1000x400") # Definindo o tamanho da janela
        master.resizable(False, False) # Impedir redimensionamento

        # Configurar estilo ttk
        style = ttk.Style()
        style.theme_use("clam") # Ou "alt", "default", "classic"
        style.configure("TLabel", font=("Arial", 10))
        style.configure("TEntry", font=("Arial", 10))
        style.configure("TButton", font=("Arial", 10, "bold"), padding=5)
        style.configure("TLabelframe.Label", font=("Arial", 12, "bold"))

        # pdf_fields agora é importado de pdf_mapping.py
        self.pdf_fields = pdf_fields

        # Variáveis para armazenar os dados
        self.cnpj_var = tk.StringVar()
        self.nome_fantasia_var = tk.StringVar()
        self.endereco_var = tk.StringVar()
        self.data_feriado_var = tk.StringVar()
        self.output_dir_path = tk.StringVar()

        # Widgets da interface
        self.create_widgets()

    def create_widgets(self):
        # Frame para os campos de entrada
        input_frame = ttk.LabelFrame(self.master, text="Dados para Preenchimento")
        input_frame.pack(padx=20, pady=20, fill="x", expand=True)

        # Configurar grid para centralizar e expandir
        input_frame.columnconfigure(1, weight=1)

        # CNPJ
        ttk.Label(input_frame, text="CNPJ:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(input_frame, textvariable=self.cnpj_var, width=60).grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        # Nome da Empresa
        ttk.Label(input_frame, text="Nome da Empresa:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(input_frame, textvariable=self.nome_fantasia_var, width=60).grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        # Endereço
        ttk.Label(input_frame, text="Endereço:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(input_frame, textvariable=self.endereco_var, width=60).grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # Dia do Feriado
        ttk.Label(input_frame, text="Dia do Feriado:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(input_frame, textvariable=self.data_feriado_var, width=60).grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        # Frame para seleção de arquivos
        file_frame = ttk.LabelFrame(self.master, text="Seleção de Arquivos")
        file_frame.pack(padx=20, pady=10, fill="x", expand=True)
        file_frame.columnconfigure(1, weight=1)

        # Diretório de Saída
        ttk.Label(file_frame, text="Diretório de Saída:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(file_frame, textvariable=self.output_dir_path, width=50, state="readonly").grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        ttk.Button(file_frame, text="Selecionar Pasta", command=self.select_output_dir).grid(row=0, column=2, padx=10, pady=5)

        # Botão de Preenchimento
        fill_button = ttk.Button(self.master, text="Preencher PDF", command=self.fill_pdf)
        fill_button.pack(pady=20)

    def select_output_dir(self):
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.output_dir_path.set(dir_path)

    def fill_pdf(self):
        input_pdf = "SOLICITAÇÃO PARA AUTORIZAÇÃO DE ABERTURA NO FERIADO  - SINDNORTE.pdf" # Caminho fixo
        output_dir = self.output_dir_path.get()

        if not input_pdf or not output_dir:
            messagebox.showerror("Erro", "Por favor, selecione o diretório de saída e certifique-se que o PDF de entrada está disponível.")
            return

        data_to_fill = {
            'cnpj': self.cnpj_var.get(),
            'nome_fantasia': self.nome_fantasia_var.get(),
            'endereco': self.endereco_var.get(),
            'data': self.data_feriado_var.get(),
        }

        # Chamar a função de preenchimento do módulo pdf_filler
        filled_pdf_path = fill_pdf_document(input_pdf, output_dir, data_to_fill)

        if filled_pdf_path:
            messagebox.showinfo("Sucesso", f"PDF preenchido e salvo em: {filled_pdf_path}")
        else:
            messagebox.showerror("Erro", "Ocorreu um erro ao preencher o PDF. Verifique o console para mais detalhes.")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFillerApp(root)
    root.mainloop()


