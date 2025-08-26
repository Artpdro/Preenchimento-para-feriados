import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sqlite3
import os
from datetime import datetime

# Importar a função de preenchimento e o mapeamento
from pdf_filler import fill_pdf_document
from pdf_mapping import pdf_fields
from cnpj_formatter import format_cnpj, validate_cnpj_format, clean_cnpj

class PDFillerApp:
    def __init__(self, master):
        self.master = master
        master.title("Preencher solicitação de abertura")
        master.geometry("1200x900")
        master.resizable(True, True)
        master.configure(bg="#F0F0F0")

        # Cores baseadas na imagem de referência do Sindnorte
        self.colors = {
            'primary_blue': '#4A90E2',
            'light_blue': '#5BA3F5',
            'dark_blue': '#2C5AA0',
            'white': '#FFFFFF',
            'light_gray': '#F5F5F5',
            'medium_gray': '#E0E0E0',
            'dark_gray': '#666666',
            'text_dark': '#333333'
        }

        # Configurar estilo ttk
        style = ttk.Style()
        style.theme_use("clam")

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

        # Estilo especial para o botão de pesquisa
        style.configure("Search.TButton", 
                        font=("Arial", 10, "bold"), 
                        padding=5, 
                        background='#28a745', 
                        foreground=self.colors['white'],
                        relief="flat")
        style.map("Search.TButton", 
                  background=[('active', '#218838')],
                  foreground=[('active', self.colors['white'])])

        # Estilo para botão de preenchimento em lote
        style.configure("Batch.TButton", 
                        font=("Arial", 10, "bold"), 
                        padding=5, 
                        background='#FF6B35', 
                        foreground=self.colors['white'],
                        relief="flat")
        style.map("Batch.TButton", 
                  background=[('active', '#E55A2B')],
                  foreground=[('active', self.colors['white'])])

        # Estilos para botões CRUD
        style.configure("Add.TButton", 
                        font=("Arial", 10, "bold"), 
                        padding=5, 
                        background='#28a745', 
                        foreground=self.colors['white'],
                        relief="flat")
        style.map("Add.TButton", 
                  background=[('active', '#218838')],
                  foreground=[('active', self.colors['white'])])

        style.configure("Update.TButton", 
                        font=("Arial", 10, "bold"), 
                        padding=5, 
                        background='#ffc107', 
                        foreground=self.colors['text_dark'],
                        relief="flat")
        style.map("Update.TButton", 
                  background=[('active', '#e0a800')],
                  foreground=[('active', self.colors['text_dark'])])

        style.configure("Delete.TButton", 
                        font=("Arial", 10, "bold"), 
                        padding=5, 
                        background='#dc3545', 
                        foreground=self.colors['white'],
                        relief="flat")
        style.map("Delete.TButton", 
                  background=[('active', '#c82333')],
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

        # Estilo para Treeview (tabela)
        style.configure("Treeview", 
                        font=("Arial", 9),
                        background=self.colors['white'],
                        foreground=self.colors['text_dark'],
                        fieldbackground=self.colors['white'])
        style.configure("Treeview.Heading", 
                        font=("Arial", 10, "bold"),
                        background=self.colors['primary_blue'],
                        foreground=self.colors['white'])

        # pdf_fields agora é importado de pdf_mapping.py
        self.pdf_fields = pdf_fields

        # Variáveis para armazenar os dados
        self.cnpj_var = tk.StringVar()
        self.razao_social_var = tk.StringVar()
        self.telefone_var = tk.StringVar()
        self.responsavel_var = tk.StringVar()
        self.data_feriado_var = tk.StringVar()
        self.filter_cnpj_var = tk.StringVar()
        self.filter_razao_social_var = tk.StringVar()
        self.output_dir_path = tk.StringVar()
        self.nome_fantasia_var = tk.StringVar()
        self.endereco_var = tk.StringVar()
        
        # Variáveis para pesquisa
        self.search_term_var = tk.StringVar()
        
        # Variáveis para preenchimento em lote
        self.batch_data_feriado_var = tk.StringVar()
        self.batch_output_dir_path = tk.StringVar()
        self.batch_filter_cnpj_var = tk.StringVar()
        self.batch_filter_razao_social_var = tk.StringVar()

        # Variável para controlar se estamos editando uma empresa existente
        self.editing_company_id = None

        # Criar notebook (abas)
        self.notebook = ttk.Notebook(master)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Criar as abas
        self.create_individual_tab()
        self.create_batch_tab()
        
        # Carregar dados das empresas
        self.load_companies()

    def create_individual_tab(self):
        """Cria a aba de preenchimento individual"""
        # Frame para a aba individual
        individual_frame = ttk.Frame(self.notebook)
        self.notebook.add(individual_frame, text="Preenchimento Individual")

        # Cabeçalho
        header_frame = tk.Frame(individual_frame, bg=self.colors['primary_blue'], height=40)
        header_frame.pack(fill='x', padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        header_label = tk.Label(header_frame, text="Preencher abertura de feriados - Individual", 
                                font=("Arial", 14, "bold"), 
                                fg=self.colors['white'], 
                                bg=self.colors['primary_blue'])
        header_label.pack(pady=8)

        # Frame principal para o conteúdo
        main_content_frame = tk.Frame(individual_frame, bg=self.colors["light_gray"])
        main_content_frame.pack(padx=10, pady=10, fill="both", expand=True)

        # Frame para pesquisa de empresas
        search_frame = ttk.LabelFrame(main_content_frame, text="Pesquisar Empresa")
        search_frame.pack(padx=10, pady=10, fill="x")
        search_frame.columnconfigure(1, weight=1)

        # Campo de pesquisa
        ttk.Label(search_frame, text="CNPJ ou Razão Social:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_term_var, width=50)
        search_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        
        # Botão de pesquisa
        search_button = ttk.Button(search_frame, text="Pesquisar e Preencher", 
                                 command=self.search_and_fill_company, style="Search.TButton")
        search_button.grid(row=0, column=2, padx=10, pady=5)

        # Bind Enter key para pesquisa
        search_entry.bind('<Return>', lambda event: self.search_and_fill_company())

        # Frame para os campos de entrada
        input_frame = ttk.LabelFrame(main_content_frame, text="Dados para Preenchimento")
        input_frame.pack(padx=10, pady=10, fill="x")
        input_frame.columnconfigure(1, weight=1)

        # CNPJ
        ttk.Label(input_frame, text="CNPJ:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.cnpj_entry = ttk.Entry(input_frame, textvariable=self.cnpj_var, width=60)
        self.cnpj_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        self.cnpj_entry.bind('<KeyRelease>', self.format_cnpj_on_type)
        self.cnpj_entry.bind('<FocusOut>', self.validate_cnpj_on_focus_out)

        # Nome da Empresa
        ttk.Label(input_frame, text="Nome da empresa (Razão social):").grid(row=1, column=0, sticky="w", padx=10, pady=5)        
        ttk.Entry(input_frame, textvariable=self.razao_social_var, width=60).grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        # Nome Fantasia
        ttk.Label(input_frame, text="Nome Fantasia:").grid(row=2, column=0, sticky="w", padx=10, pady=5)        
        ttk.Entry(input_frame, textvariable=self.nome_fantasia_var, width=60).grid(row=2, column=1, padx=10, pady=5, sticky="ew")



        # Telefone
        ttk.Label(input_frame, text="Telefone:").grid(row=4, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(input_frame, textvariable=self.telefone_var, width=60).grid(row=4, column=1, padx=10, pady=5, sticky="ew")



        # Endereço
        ttk.Label(input_frame, text="Endereço:").grid(row=6, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(input_frame, textvariable=self.endereco_var, width=60).grid(row=6, column=1, padx=10, pady=5, sticky="ew")



        # Responsável
        ttk.Label(input_frame, text="Responsável:").grid(row=8, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(input_frame, textvariable=self.responsavel_var, width=60).grid(row=8, column=1, padx=10, pady=5, sticky="ew")

        # Data do Feriado
        ttk.Label(input_frame, text="Data do Feriado:").grid(row=9, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(input_frame, textvariable=self.data_feriado_var, width=60).grid(row=9, column=1, padx=10, pady=5, sticky="ew")

        # Frame para botões CRUD
        crud_frame = ttk.LabelFrame(main_content_frame, text="Gerenciar Empresas")
        crud_frame.pack(padx=10, pady=10, fill="x")

        # Botões CRUD
        buttons_frame = tk.Frame(crud_frame, bg=self.colors['light_gray'])
        buttons_frame.pack(pady=10)

        ttk.Button(buttons_frame, text="Adicionar Empresa", 
                  command=self.add_company, style="Add.TButton").pack(side="left", padx=5)
        
        ttk.Button(buttons_frame, text="Atualizar Empresa", 
                  command=self.update_company, style="Update.TButton").pack(side="left", padx=5)
        
        ttk.Button(buttons_frame, text="Deletar Empresa", 
                  command=self.delete_company, style="Delete.TButton").pack(side="left", padx=5)
        

        # Frame para seleção de arquivos
        file_frame = ttk.LabelFrame(main_content_frame, text="Seleção de Arquivos")
        file_frame.pack(padx=10, pady=10, fill="x")
        file_frame.columnconfigure(1, weight=1)

        # Diretório de Saída
        ttk.Label(file_frame, text="Local para salvar:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(file_frame, textvariable=self.output_dir_path, width=50, state="readonly").grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        ttk.Button(file_frame, text="Selecionar Pasta", command=self.select_output_dir).grid(row=0, column=2, padx=10, pady=5)

        # Botão de Preenchimento
        fill_button = ttk.Button(main_content_frame, text="Preencher PDF", command=self.fill_pdf)
        fill_button.pack(pady=10)

        # Frame para a tabela de empresas
        table_frame = ttk.LabelFrame(main_content_frame, text="Empresas Cadastradas")
        table_frame.pack(padx=10, pady=10, fill="both", expand=True)

        # Criar a tabela (Treeview)
        columns = ("CNPJ", "Razão Social", "Nome Fantasia", "Telefone", "Endereço", "Responsável")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=15)
        
        # Configurar cabeçalhos
        for col in columns:
            self.tree.heading(col, text=col)
            if col == "CNPJ":
                self.tree.column(col, width=120, minwidth=100)
            elif col == "Razão Social":
                self.tree.column(col, width=150, minwidth=120)
            elif col == "Nome Fantasia":
                self.tree.column(col, width=120, minwidth=100)

        # Scrollbars para a tabela
        v_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        # Posicionar tabela e scrollbars
        self.tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")

        # Configurar expansão do grid
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # Bind para seleção de linha
        self.tree.bind("<<TreeviewSelect>>", self.on_company_select)

        # Frame para filtros
        filter_frame = ttk.Frame(table_frame)
        filter_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        
        ttk.Label(filter_frame, text="Filtrar por CNPJ:").pack(side="left", padx=5)
        self.filter_cnpj_var = tk.StringVar()
        self.filter_cnpj_var.trace("w", self.filter_companies)
        filter_cnpj_entry = ttk.Entry(filter_frame, textvariable=self.filter_cnpj_var, width=20)
        filter_cnpj_entry.pack(side="left", padx=5)

        ttk.Label(filter_frame, text="Filtrar por Nome/Fantasia:").pack(side="left", padx=5)
        self.filter_razao_social_var = tk.StringVar()
        self.filter_razao_social_var.trace("w", self.filter_companies)
        filter_razao_social_entry = ttk.Entry(filter_frame, textvariable=self.filter_razao_social_var, width=30)
        filter_razao_social_entry.pack(side="left", padx=5)

    def create_batch_tab(self):
        """Cria a aba de preenchimento em lote"""
        # Frame para a aba de lote
        batch_frame = ttk.Frame(self.notebook)
        self.notebook.add(batch_frame, text="Preenchimento em Lote")

        # Cabeçalho
        header_frame = tk.Frame(batch_frame, bg=self.colors['primary_blue'], height=40)
        header_frame.pack(fill='x', padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        header_label = tk.Label(header_frame, text="Preencher abertura de feriados - Em Lote", 
                                font=("Arial", 14, "bold"), 
                                fg=self.colors['white'], 
                                bg=self.colors['primary_blue'])
        header_label.pack(pady=8)

        # Frame principal para o conteúdo
        main_content_frame = tk.Frame(batch_frame, bg=self.colors["light_gray"])
        main_content_frame.pack(padx=10, pady=10, fill="both", expand=True)

        # Frame para configuração do lote
        batch_config_frame = ttk.LabelFrame(main_content_frame, text="Configuração do Preenchimento em Lote")
        batch_config_frame.pack(padx=10, pady=10, fill="x")
        batch_config_frame.columnconfigure(1, weight=1)

        # Data do Feriado para lote
        ttk.Label(batch_config_frame, text="Data do Feriado:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        batch_date_entry = ttk.Entry(batch_config_frame, textvariable=self.batch_data_feriado_var, width=60)
        batch_date_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        
        # Adicionar dica de formato
        ttk.Label(batch_config_frame, text="(Formato: DD/MM/AAAA)", 
                 font=("Arial", 8), foreground=self.colors['dark_gray']).grid(row=1, column=1, sticky="w", padx=10)

        # Frame para seleção de diretório de saída
        batch_file_frame = ttk.LabelFrame(main_content_frame, text="Local para Salvar")
        batch_file_frame.pack(padx=10, pady=10, fill="x")
        batch_file_frame.columnconfigure(1, weight=1)

        # Diretório de Saída para lote
        ttk.Label(batch_file_frame, text="Pasta de destino:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(batch_file_frame, textvariable=self.batch_output_dir_path, width=50, state="readonly").grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        ttk.Button(batch_file_frame, text="Selecionar Pasta", command=self.select_batch_output_dir).grid(row=0, column=2, padx=10, pady=5)

        # Frame para a tabela de empresas no modo lote
        batch_table_frame = ttk.LabelFrame(main_content_frame, text="Empresas Cadastradas (Selecione para Preencher)")
        batch_table_frame.pack(padx=10, pady=10, fill="both", expand=True)

        # Criar a tabela (Treeview) para seleção em lote
        columns = ("CNPJ", "Razão Social", "Nome Fantasia", "Telefone", "Endereço", "Responsável")
        self.batch_tree = ttk.Treeview(batch_table_frame, columns=columns, show="headings", height=10, selectmode="extended")
        
        # Configurar cabeçalhos
        for col in columns:
            self.batch_tree.heading(col, text=col)
            if col == "CNPJ":
                self.batch_tree.column(col, width=100, minwidth=80)
            elif col == "Razão Social":
                self.batch_tree.column(col, width=120, minwidth=100)
            elif col == "Nome Fantasia":
                self.batch_tree.column(col, width=100, minwidth=80)

        # Scrollbars para a tabela de lote
        batch_v_scrollbar = ttk.Scrollbar(batch_table_frame, orient="vertical", command=self.batch_tree.yview)
        batch_h_scrollbar = ttk.Scrollbar(batch_table_frame, orient="horizontal", command=self.batch_tree.xview)
        self.batch_tree.configure(yscrollcommand=batch_v_scrollbar.set, xscrollcommand=batch_h_scrollbar.set)

        # Posicionar tabela e scrollbars
        self.batch_tree.grid(row=0, column=0, sticky="nsew")
        batch_v_scrollbar.grid(row=0, column=1, sticky="ns")
        batch_h_scrollbar.grid(row=1, column=0, sticky="ew")

        # Configurar expansão do grid
        batch_table_frame.grid_rowconfigure(0, weight=1)
        batch_table_frame.grid_columnconfigure(0, weight=1)

        # Bind para seleção de linha na tabela de lote
        self.batch_tree.bind("<<TreeviewSelect>>", self.on_company_select)


        # Scrollbars para a tabela de lote
        batch_v_scrollbar = ttk.Scrollbar(batch_table_frame, orient="vertical", command=self.batch_tree.yview)
        batch_h_scrollbar = ttk.Scrollbar(batch_table_frame, orient="horizontal", command=self.batch_tree.xview)
        self.batch_tree.configure(yscrollcommand=batch_v_scrollbar.set, xscrollcommand=batch_h_scrollbar.set)

        # Posicionar tabela e scrollbars
        self.batch_tree.grid(row=0, column=0, sticky="nsew")
        batch_v_scrollbar.grid(row=0, column=1, sticky="ns")
        batch_h_scrollbar.grid(row=1, column=0, sticky="ew")

        # Configurar expansão do grid
        batch_table_frame.grid_rowconfigure(0, weight=1)
        batch_table_frame.grid_columnconfigure(0, weight=1)

        # Frame para filtros da tabela de lote
        batch_filter_frame = ttk.Frame(batch_table_frame)
        batch_filter_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        
        ttk.Label(batch_filter_frame, text="Filtrar por CNPJ:").pack(side="left", padx=5)
        self.batch_filter_cnpj_var = tk.StringVar()
        self.batch_filter_cnpj_var.trace("w", self.filter_batch_companies)
        batch_filter_cnpj_entry = ttk.Entry(batch_filter_frame, textvariable=self.batch_filter_cnpj_var, width=20)
        batch_filter_cnpj_entry.pack(side="left", padx=5)

        ttk.Label(batch_filter_frame, text="Filtrar por Nome/Fantasia:").pack(side="left", padx=5)
        self.batch_filter_razao_social_var = tk.StringVar()
        self.batch_filter_razao_social_var.trace("w", self.filter_batch_companies)
        batch_filter_razao_social_entry = ttk.Entry(batch_filter_frame, textvariable=self.batch_filter_razao_social_var, width=30)
        batch_filter_razao_social_entry.pack(side="left", padx=5)

        # Botões para seleção em lote
        batch_selection_frame = tk.Frame(main_content_frame, bg=self.colors["light_gray"])
        batch_selection_frame.pack(pady=10)

        ttk.Button(batch_selection_frame, text="Selecionar Todas", command=self.select_all_companies).pack(side="left", padx=5)
        ttk.Button(batch_selection_frame, text="Desmarcar Todas", command=self.deselect_all_companies).pack(side="left", padx=5)

        # Botão de Preenchimento em Lote
        batch_fill_button = ttk.Button(main_content_frame, text="Preencher PDFs em Lote", 
                                      command=self.fill_batch_pdfs, style="Batch.TButton")
        batch_fill_button.pack(pady=10)

    # Métodos CRUD para empresas
    def add_company(self):
        """Adiciona uma nova empresa ao banco de dados"""
        # Validar campos obrigatórios
        cnpj = self.cnpj_var.get().strip()
        razao_social = self.razao_social_var.get().strip()
        
        if not cnpj or not razao_social:
            messagebox.showerror("Erro", "CNPJ e Razão Social são obrigatórios!")
            return
        
        # Validar formato do CNPJ
        if not validate_cnpj_format(cnpj):
            messagebox.showerror("Erro", "CNPJ deve ter formato válido (XX.XXX.XXX/XXXX-XX)!")
            return
        
        try:
            conn = sqlite3.connect('empresas.db')
            cursor = conn.cursor()
            
            # Verificar se o CNPJ já existe
            cnpj_clean = clean_cnpj(cnpj)
            cursor.execute("SELECT cnpj FROM empresas WHERE REPLACE(REPLACE(REPLACE(cnpj, '.', ''), '/', ''), '-', '') = ?", (cnpj_clean,))
            if cursor.fetchone():
                messagebox.showerror("Erro", "CNPJ já cadastrado!")
                conn.close()
                return
            
            # Inserir nova empresa (adaptado para a estrutura real da tabela)
            cursor.execute("""
                INSERT INTO empresas (cnpj, razao_social, nome_fantasia, email, telefone, celular, endereco, cidade)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                cnpj,
                razao_social,
                self.nome_fantasia_var.get().strip(),
                self.email_var.get().strip(),
                self.telefone_var.get().strip()
            ))
            
            conn.commit()
            conn.close()
            
            messagebox.showinfo("Sucesso", "Empresa adicionada com sucesso!")
            self.load_companies()
            self.clear_fields()
            
        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao adicionar empresa: {e}")

    def update_company(self):
        """Atualiza uma empresa existente no banco de dados"""
        # Verificar se uma empresa está selecionada
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Aviso", "Selecione uma empresa na tabela para atualizar!")
            return
        
        # Validar campos obrigatórios
        cnpj = self.cnpj_var.get().strip()
        razao_social = self.razao_social_var.get().strip()
        
        if not cnpj or not razao_social:
            messagebox.showerror("Erro", "CNPJ e Razão Social são obrigatórios!")
            return
        
        # Validar formato do CNPJ
        if not validate_cnpj_format(cnpj):
            messagebox.showerror("Erro", "CNPJ deve ter formato válido (XX.XXX.XXX/XXXX-XX)!")
            return
        
        try:
            conn = sqlite3.connect('empresas.db')
            cursor = conn.cursor()
            
            # Obter CNPJ original da empresa selecionada
            item = self.tree.item(selection[0])
            original_cnpj = item['values'][0]
            original_cnpj_clean = clean_cnpj(original_cnpj)
            
            # Se o CNPJ foi alterado, verificar se o novo CNPJ já existe
            cnpj_clean = clean_cnpj(cnpj)
            if cnpj_clean != original_cnpj_clean:
                cursor.execute("SELECT cnpj FROM empresas WHERE REPLACE(REPLACE(REPLACE(cnpj, '.', ''), '/', ''), '-', '') = ?", (cnpj_clean,))
                if cursor.fetchone():
                    messagebox.showerror("Erro", "CNPJ já cadastrado para outra empresa!")
                    conn.close()
                    return
            
            # Atualizar empresa (adaptado para a estrutura real da tabela)
            cursor.execute("""
                UPDATE empresas 
                SET cnpj = ?, razao_social = ?, nome_fantasia = ?, telefone = ?, endereco = ?, responsavel = ?
                WHERE REPLACE(REPLACE(REPLACE(cnpj, ".", ""), "/", ""), "-", "") = ?
            """, (
                cnpj,
                razao_social,
                self.nome_fantasia_var.get().strip(),
                self.telefone_var.get().strip(),
                self.endereco_var.get().strip(),
                self.responsavel_var.get().strip(),
                original_cnpj_clean
            ))
            
            conn.commit()
            conn.close()
            
            messagebox.showinfo("Sucesso", "Empresa atualizada com sucesso!")
            self.load_companies()
            self.clear_fields()
            
        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao atualizar empresa: {e}")

    def delete_company(self):
        """Deleta uma empresa do banco de dados"""
        # Verificar se uma empresa está selecionada
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Aviso", "Selecione uma empresa na tabela para deletar!")
            return
        
        # Obter dados da empresa selecionada
        item = self.tree.item(selection[0])
        cnpj = item['values'][0]
        razao_social = item['values'][1]
        
        # Confirmar exclusão
        if not messagebox.askyesno("Confirmar Exclusão", 
                                   f"Tem certeza que deseja deletar a empresa:\n{razao_social}\nCNPJ: {cnpj}?"):
            return
        
        try:
            conn = sqlite3.connect('empresas.db')
            cursor = conn.cursor()
            
            # Deletar empresa
            cnpj_clean = clean_cnpj(cnpj)
            cursor.execute("DELETE FROM empresas WHERE REPLACE(REPLACE(REPLACE(cnpj, '.', ''), '/', ''), '-', '') = ?", (cnpj_clean,))
            
            conn.commit()
            conn.close()
            
            messagebox.showinfo("Sucesso", "Empresa deletada com sucesso!")
            self.load_companies()
            self.clear_fields()
            
        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao deletar empresa: {e}")

    def clear_fields(self):
        """Limpa todos os campos de entrada"""
        self.cnpj_var.set("")
        self.razao_social_var.set("")
        self.nome_fantasia_var.set("")
        self.telefone_var.set("")
        self.endereco_var.set("")
        self.responsavel_var.set("")
        self.data_feriado_var.set("")
        self.editing_company_id = None

    def select_batch_output_dir(self):
        """Seleciona o diretório de saída para lote"""
        directory = filedialog.askdirectory(title="Selecionar pasta para salvar PDFs em lote")
        if directory:
            self.batch_output_dir_path.set(directory)

    def select_all_companies(self):
        """Seleciona todas as empresas na tabela de lote"""
        for item in self.batch_tree.get_children():
            self.batch_tree.selection_add(item)

    def deselect_all_companies(self):
        """Desmarca todas as empresas na tabela de lote"""
        self.batch_tree.selection_remove(self.batch_tree.selection())

    def fill_batch_pdfs(self):
        """Preenche PDFs em lote para as empresas selecionadas"""
        # Validar seleção
        selected_items = self.batch_tree.selection()
        if not selected_items:
            messagebox.showwarning("Aviso", "Selecione pelo menos uma empresa para preencher!")
            return

        # Validar data do feriado
        data_feriado = self.batch_data_feriado_var.get().strip()
        if not data_feriado:
            messagebox.showerror("Erro", "Por favor, informe a data do feriado!")
            return

        # Validar diretório de saída
        output_dir = self.batch_output_dir_path.get()
        if not output_dir:
            messagebox.showerror("Erro", "Por favor, selecione a pasta de destino!")
            return

        # Verificar se o template existe
        template_path = "formulario.pdf"
        if not os.path.exists(template_path):
            messagebox.showerror("Erro", f"Arquivo template não encontrado: {template_path}")
            return

        try:
            # Coletar dados das empresas selecionadas
            empresas_selecionadas = []
            for item in selected_items:
                values = self.batch_tree.item(item)['values']
                empresa_data = {
                    'cnpj': values[0],
                    'razao_social': values[1],
                    'nome_fantasia': values[2],
                    'telefone': values[3],
                    'endereco': values[4],
                    'responsavel': values[5]
                }
                empresas_selecionadas.append(empresa_data)

            total_empresas = len(empresas_selecionadas)
            sucessos = 0
            erros = []

            # Criar janela de progresso
            progress_window = tk.Toplevel(self.master)
            progress_window.title("Processando...")
            progress_window.geometry("400x150")
            progress_window.resizable(False, False)
            progress_window.transient(self.master)
            progress_window.grab_set()

            # Centralizar janela
            progress_window.update_idletasks()
            x = (progress_window.winfo_screenwidth() // 2) - (400 // 2)
            y = (progress_window.winfo_screenheight() // 2) - (150 // 2)
            progress_window.geometry(f"400x150+{x}+{y}")

            # Widgets da janela de progresso
            progress_label = tk.Label(progress_window, text=f"Processando 0 de {total_empresas} empresas...", 
                                    font=("Arial", 12))
            progress_label.pack(pady=10)

            progress_bar = ttk.Progressbar(progress_window, length=300, mode='determinate')
            progress_bar.pack(pady=10)
            progress_bar['maximum'] = total_empresas

            status_label = tk.Label(progress_window, text="Iniciando...", font=("Arial", 10))
            status_label.pack(pady=5)

            # Processar cada empresa selecionada
            for i, empresa_data in enumerate(empresas_selecionadas):
                cnpj_formatted = empresa_data['cnpj']
                razao_social = empresa_data['razao_social']
                nome_fantasia = empresa_data['nome_fantasia']
                email = empresa_data['email']
                telefone = empresa_data['telefone']
                celular = empresa_data['celular']
                endereco = empresa_data['endereco']
                municipio = empresa_data['municipio']
                
                # Atualizar interface
                progress_label.config(text=f"Processando empresa {i+1} de {total_empresas}")
                status_label.config(text=f"Empresa: {razao_social[:40]}...")
                progress_bar['value'] = i + 1
                progress_window.update()

                try:
                    # Preparar dados para preenchimento
                    data = {
                        'cnpj': cnpj_formatted,
                        'razao_social': razao_social or '',
                        'nome_fantasia': nome_fantasia or '',
                        'telefone': telefone or '',
                        'endereco': endereco or '',
                        'responsavel': responsavel or '',
                        'data_feriado': data_feriado
                    }

                    # Nome do arquivo de saída
                    nome_arquivo_limpo = razao_social.replace('/', '_').replace('\\', '_')
                    nome_arquivo = f"{nome_arquivo_limpo[:50]}.pdf"
                    output_path = os.path.join(output_dir, nome_arquivo)

                    # Preencher PDF
                    if os.path.exists(template_path):
                        fill_pdf_document(template_path, output_dir, data)
                        sucessos += 1
                    else:
                        erros.append(f"{razao_social}: Arquivo template não encontrado")

                except Exception as e:
                    erros.append(f"{razao_social}: {str(e)}")

            # Fechar janela de progresso
            progress_window.destroy()

            # Mostrar relatório
            relatorio = f"Processamento concluído!\n\n"
            relatorio += f"Total de empresas selecionadas: {total_empresas}\n"
            relatorio += f"PDFs criados com sucesso: {sucessos}\n"
            relatorio += f"Erros: {len(erros)}\n\n"

            if erros:
                relatorio += "Empresas com erro:\n"
                for erro in erros[:10]:  # Mostrar apenas os primeiros 10 erros
                    relatorio += f"• {erro}\n"
                if len(erros) > 10:
                    relatorio += f"... e mais {len(erros) - 10} erros.\n"

            messagebox.showinfo("Relatório", relatorio)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro durante o processamento: {e}")

    def search_and_fill_company(self):
        """Pesquisa empresa e preenche campos automaticamente"""
        search_term = self.search_term_var.get().strip()
        
        if not search_term:
            messagebox.showwarning("Aviso", "Por favor, digite um CNPJ ou razão social para pesquisar.")
            return
        
        try:
            conn = sqlite3.connect('empresas.db')
            cursor = conn.cursor()
            
            # Remover caracteres especiais do CNPJ se for numérico
            search_term_clean = ''.join(filter(str.isdigit, search_term))
            
            # Pesquisar por CNPJ (exato) ou razão social/nome fantasia (parcial)
            cursor.execute("""
                SELECT cnpj, razao_social, nome_fantasia, telefone, endereco, responsavel
                FROM empresas 
                WHERE REPLACE(REPLACE(REPLACE(cnpj, ".", ""), "/", ""), "-", "") LIKE ? 
                   OR UPPER(razao_social) LIKE UPPER(?) 
                   OR UPPER(nome_fantasia) LIKE UPPER(?)
                ORDER BY 
                    CASE 
                        WHEN REPLACE(REPLACE(REPLACE(cnpj, ".", ""), "/", ""), "-", "") = ? THEN 1
                        WHEN UPPER(razao_social) = UPPER(?) THEN 2
                        WHEN UPPER(nome_fantasia) = UPPER(?) THEN 3
                        ELSE 4
                    END
                LIMIT 1
            """, (f'%{search_term_clean}%', f'%{search_term}%', f'%{search_term}%', 
                  search_term_clean, search_term, search_term))
            
            result = cursor.fetchone()
            
            if result:
                # Preencher os campos com os dados encontrados
                cnpj = str(result[0]).replace('.0', '') if result[0] else ''
                cnpj_formatted = format_cnpj(cnpj) if cnpj else ''
                
                razao_social = result[1] or ''
                nome_fantasia = result[2] or ''
                telefone = str(result[3]).replace(".0", "") if result[3] else ""
                endereco = result[4] or ""
                responsavel = result[5] or ""
                
                # Preencher os campos da interface
                self.cnpj_var.set(cnpj_formatted)
                self.razao_social_var.set(razao_social)
                self.nome_fantasia_var.set(nome_fantasia)
                self.telefone_var.set(telefone)
                self.endereco_var.set(endereco)
                self.responsavel_var.set(responsavel)                # Limpar campo de pesquisa
                self.search_term_var.set("")
                
                # Mostrar mensagem de sucesso
                messagebox.showinfo("Sucesso", f"Empresa encontrada: {razao_social}")
                
                # Destacar a empresa na tabela se estiver visível
                self.highlight_company_in_table(cnpj_formatted)
                
            else:
                messagebox.showwarning("Não encontrado", 
                                     f"Nenhuma empresa encontrada com o termo: {search_term}")
            
            conn.close()
            
        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao pesquisar empresa: {e}")

    def highlight_company_in_table(self, cnpj):
        """Destaca a empresa na tabela se ela estiver visível"""
        try:
            for item in self.tree.get_children():
                values = self.tree.item(item)['values']
                if values and str(values[0]) == cnpj:
                    self.tree.selection_set(item)
                    self.tree.focus(item)
                    self.tree.see(item)
                    break
        except Exception as e:
            print(f"Erro ao destacar empresa na tabela: {e}")

    def load_companies(self):
        """Carrega as empresas do banco de dados na tabela individual e na tabela de lote"""
        try:
            conn = sqlite3.connect('empresas.db')
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT cnpj, razao_social, nome_fantasia, telefone, endereco, responsavel
                FROM empresas 
                ORDER BY razao_social
            """)
            
            companies = cursor.fetchall()
            
            # Limpar tabela individual existente
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Limpar tabela de lote existente
            for item in self.batch_tree.get_children():
                self.batch_tree.delete(item)
            
            # Adicionar empresas às tabelas
            for company in companies:
                cnpj = str(company[0]).replace('.0', '') if company[0] else ''
                cnpj_formatted = format_cnpj(cnpj) if cnpj else ''
                
                values = (
                    cnpj_formatted,
                    company[1] or '',  # razao_social
                    company[2] or '',  # nome_fantasia
                    company[3] or '',  # email
                    str(company[4]).replace('.0', '') if company[4] else '',  # telefone
                    str(company[5]).replace('.0', '') if company[5] else '',  # celular
                    company[6] if len(company) > 6 and company[6] is not None else 
                    company[7] if len(company) > 7 and company[7] is not None else ''   # cidade
                )
                self.tree.insert('', 'end', values=values)
                self.batch_tree.insert('', 'end', values=values)
            
            conn.close()
            
        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao carregar empresas: {e}")

    def filter_companies(self, *args):
        """Filtra as empresas na tabela individual"""
        cnpj_filter = self.filter_cnpj_var.get().lower()
        name_filter = self.filter_razao_social_var.get().lower()
        
        # Se não há filtros, mostrar todas as empresas
        if not cnpj_filter and not name_filter:
            self.load_companies()
            return
        
        try:
            conn = sqlite3.connect('empresas.db')
            cursor = conn.cursor()
            
            # Construir query com filtros
            query = """
                SELECT cnpj, razao_social, nome_fantasia, telefone, endereco, responsavel
                FROM empresas 
                WHERE 1=1
            """
            params = []
            
            if cnpj_filter:
                query += " AND LOWER(cnpj) LIKE ?"
                params.append(f'%{cnpj_filter}%')
            
            if name_filter:
                query += " AND (LOWER(razao_social) LIKE ? OR LOWER(nome_fantasia) LIKE ?)"
                params.extend([f'%{name_filter}%', f'%{name_filter}%'])
            
            query += " ORDER BY razao_social"
            
            cursor.execute(query, params)
            companies = cursor.fetchall()
            
            # Limpar tabela existente
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Adicionar empresas filtradas à tabela
            for company in companies:
                cnpj = str(company[0]).replace('.0', '') if company[0] else ''
                cnpj_formatted = format_cnpj(cnpj) if cnpj else ''
                
                values = (
                    cnpj_formatted,
                    company[1] or '',  # razao_social
                    company[2] or '',  # nome_fantasia
                    company[3] or '',  # email
                    str(company[4]).replace('.0', '') if company[4] else '',  # telefone
                    str(company[5]).replace('.0', '') if company[5] else '',  # celular
                    company[6] if len(company) > 6 and company[6] is not None else 
                    company[7] if len(company) > 7 and company[7] is not None else ''   # cidade
                )
                self.tree.insert('', 'end', values=values)
            
            conn.close()
            
        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao filtrar empresas: {e}")

    def filter_batch_companies(self, *args):
        """Filtra as empresas na tabela de lote"""
        cnpj_filter = self.batch_filter_cnpj_var.get().lower()
        name_filter = self.batch_filter_razao_social_var.get().lower()
        
        # Se não há filtros, mostrar todas as empresas
        if not cnpj_filter and not name_filter:
            self.load_companies()
            return
        
        try:
            conn = sqlite3.connect('empresas.db')
            cursor = conn.cursor()
            
            # Construir query com filtros
            query = """
                SELECT cnpj, razao_social, nome_fantasia, telefone, endereco, responsavel
                FROM empresas 
                WHERE 1=1
            """
            params = []
            
            if cnpj_filter:
                query += " AND LOWER(cnpj) LIKE ?"
                params.append(f'%{cnpj_filter}%')
            
            if name_filter:
                query += " AND (LOWER(razao_social) LIKE ? OR LOWER(nome_fantasia) LIKE ?)"
                params.extend([f'%{name_filter}%', f'%{name_filter}%'])
            
            query += " ORDER BY razao_social"
            
            cursor.execute(query, params)
            companies = cursor.fetchall()
            
            # Limpar tabela existente
            for item in self.batch_tree.get_children():
                self.batch_tree.delete(item)
            
            # Adicionar empresas filtradas à tabela
            for company in companies:
                cnpj = str(company[0]).replace('.0', '') if company[0] else ''
                cnpj_formatted = format_cnpj(cnpj) if cnpj else ''
                
                values = (
                    cnpj_formatted,
                    company[1] or '',  # razao_social
                    company[2] or '',  # nome_fantasia
                    company[3] or '',  # email
                    str(company[4]).replace('.0', '') if company[4] else '',  # telefone
                    str(company[5]).replace('.0', '') if company[5] else '',  # celular
                    company[6] if len(company) > 6 and company[6] is not None else 
                    company[7] if len(company) > 7 and company[7] is not None else ''   # cidade
                )
                self.batch_tree.insert('', 'end', values=values)
            
            conn.close()
            
        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao filtrar empresas: {e}")

    def on_company_select(self, event):
        """Preenche os campos quando uma empresa é selecionada na tabela individual"""
        selection = self.tree.selection()
        if selection:
            item = self.tree.item(selection[0])
            values = item['values']
            
            if values:
                self.cnpj_var.set(values[0])
                self.razao_social_var.set(values[1])
                self.nome_fantasia_var.set(values[2])
                self.email_var.set(values[3])
                self.telefone_var.set(values[4])
                self.celular_var.set(values[5])
                self.endereco_var.set(values[6])
                self.municipio_var.set(values[7])

    def select_output_dir(self):
        """Seleciona o diretório de saída"""
        directory = filedialog.askdirectory(title="Selecionar pasta para salvar PDF")
        if directory:
            self.output_dir_path.set(directory)

    def fill_pdf(self):
        """Preenche o PDF com os dados fornecidos"""
        # Coletar dados dos campos
        data = {
            'cnpj': self.cnpj_var.get(),
            'razao_social': self.razao_social_var.get(),
            'nome_fantasia': self.nome_fantasia_var.get(),
            'email': self.email_var.get(),
            'telefone': self.telefone_var.get(),
            'celular': self.celular_var.get(),
            'responsavel': self.responsavel_var.get(),
            'endereco': self.endereco_var.get(),
            'municipio': self.municipio_var.get(),
            'data_feriado': self.data_feriado_var.get()
        }

        # Validar dados obrigatórios
        if not data['cnpj'] or not data['razao_social'] or not data['data_feriado']:
            messagebox.showerror("Erro", "Por favor, preencha pelo menos CNPJ, Razão Social e Data do Feriado.")
            return

        # Validar diretório de saída
        output_dir = self.output_dir_path.get()
        if not output_dir:
            messagebox.showerror("Erro", "Por favor, selecione o local para salvar o PDF.")
            return

        # Definir caminhos
        template_path = "formulario.pdf"
        output_filename_limpo = data['razao_social'].replace('/', '_')
        output_filename = f"{output_filename_limpo[:50]}.pdf"
        output_path = os.path.join(output_dir, output_filename)

        try:
            # Verificar se o template existe
            if not os.path.exists(template_path):
                messagebox.showerror("Erro", f"Arquivo template não encontrado: {template_path}")
                return

            # Preencher o PDF
            fill_pdf_document(template_path, output_dir, data)
            
            messagebox.showinfo("Sucesso", f"PDF preenchido com sucesso!\nSalvo em: {output_dir}")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao preencher PDF: {str(e)}")

    def format_cnpj_on_type(self, event):
        """Formata o CNPJ enquanto o usuário digita"""
        try:
            current_text = self.cnpj_var.get()
            cursor_position = self.cnpj_entry.index(tk.INSERT)
            
            # Remover caracteres não numéricos
            numbers_only = ''.join(filter(str.isdigit, current_text))
            
            # Aplicar formatação se tiver pelo menos alguns dígitos
            if len(numbers_only) > 2:
                formatted = format_cnpj(numbers_only)
                if formatted != current_text:
                    self.cnpj_var.set(formatted)
                    # Tentar manter a posição do cursor
                    try:
                        self.cnpj_entry.icursor(min(cursor_position, len(formatted)))
                    except:
                        pass
        except Exception as e:
            print(f"Erro na formatação do CNPJ: {e}")

    def validate_cnpj_on_focus_out(self, event):
        """Valida o CNPJ quando o campo perde o foco"""
        cnpj = self.cnpj_var.get()
        if cnpj and not validate_cnpj_format(cnpj):
            messagebox.showwarning("CNPJ Inválido", "O CNPJ deve ter o formato: XX.XXX.XXX/XXXX-XX")
            self.cnpj_entry.focus_set()

def main():
    root = tk.Tk()
    app = PDFillerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()

