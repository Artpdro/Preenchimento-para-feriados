import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

# Importar a função de preenchimento e o mapeamento
from pdf_filler import fill_pdf_document
from pdf_mapping import pdf_fields

class PDFillerApp:
    def __init__(self, master):
        self.master = master
        master.title("Preencher solicitação de abertura")
        master.geometry("1000x400") # Definindo o tamanho da janela
        master.resizable(False, False) # Impedir redimensionamento
        master.configure(bg="#F0F0F0") # Fundo cinza claro para a janela principal

        # Cores baseadas na imagem de referência do Sindnorte
        self.colors = {
            'primary_blue': '#4A90E2',  # Azul principal para cabeçalhos e botões
            'light_blue': '#5BA3F5',    # Azul claro para hover/ativo
            'dark_blue': '#2C5AA0',     # Azul escuro para texto em azul
            'white': '#FFFFFF',         # Branco para fundo de campos e texto
            'light_gray': '#F5F5F5',    # Cinza muito claro para frames
            'medium_gray': '#E0E0E0',   # Cinza médio para bordas ou divisores
            'dark_gray': '#666666',     # Cinza escuro para texto secundário
            'text_dark': '#333333'      # Preto quase total para texto principal
        }

        # Configurar estilo ttk
        style = ttk.Style()
        style.theme_use("clam") # Tema base

        # Estilo para Labels
        style.configure("TLabel", 
                        font=("Arial", 10), 
                        background=self.colors['light_gray'], 
                        foreground=self.colors['text_dark'])
        
        # Estilo para Entry (campos de texto)
        style.configure("TEntry", 
                        font=("Arial", 10), 
                        fieldbackground=self.colors['white'], 
                        foreground=self.colors['text_dark'],
                        bordercolor=self.colors['medium_gray'],
                        borderwidth=1,
                        relief="solid")

        # Estilo para Buttons
        style.configure("TButton", 
                        font=("Arial", 10, "bold"), 
                        padding=5, 
                        background=self.colors['primary_blue'], 
                        foreground=self.colors['white'],
                        relief="flat")
        style.map("TButton", 
                  background=[('active', self.colors['light_blue'])],
                  foreground=[('active', self.colors['white'])])

        # Estilo para LabelFrames (molduras)
        style.configure("TLabelframe.Label", 
                        font=("Arial", 12, "bold"), 
                        background=self.colors['light_gray'], 
                        foreground=self.colors['dark_blue'])
        style.configure("TLabelframe", 
                        background=self.colors['light_gray'],
                        bordercolor=self.colors['medium_gray'],
                        borderwidth=1,
                        relief="solid")

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
        # Cabeçalho (simples, apenas para cor)
        header_frame = tk.Frame(self.master, bg=self.colors['primary_blue'], height=40)
        header_frame.pack(fill='x', padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        header_label = tk.Label(header_frame, text="Preencher abertura de feriados", 
                                font=("Arial", 14, "bold"), 
                                fg=self.colors['white'], 
                                bg=self.colors['primary_blue'])
        header_label.pack(pady=8)

        # Frame principal para o conteúdo
        main_content_frame = tk.Frame(self.master, bg=self.colors["light_gray"])
        main_content_frame.pack(padx=10, pady=10, fill="both", expand=True)

        # Frame para os campos de entrada
        input_frame = ttk.LabelFrame(main_content_frame, text="Dados para Preenchimento")
        input_frame.pack(padx=10, pady=10, fill="x", expand=True)

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
        file_frame = ttk.LabelFrame(main_content_frame, text="Seleção de Arquivos")
        file_frame.pack(padx=10, pady=10, fill="x", expand=True)
        file_frame.columnconfigure(1, weight=1)

        # Diretório de Saída
        ttk.Label(file_frame, text="Local para salvar:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(file_frame, textvariable=self.output_dir_path, width=50, state="readonly").grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        ttk.Button(file_frame, text="Selecionar Pasta", command=self.select_output_dir).grid(row=0, column=2, padx=10, pady=5)

        # Botão de Preenchimento
        fill_button = ttk.Button(main_content_frame, text="Preencher PDF", command=self.fill_pdf)
        fill_button.pack(pady=10)

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


