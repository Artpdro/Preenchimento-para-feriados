import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sqlite3
import os
import threading
from datetime import datetime

# Importar a função de preenchimento e o mapeamento
from pdf_filler import fill_pdf_document
from pdf_mapping import pdf_fields
from cnpj_formatter import format_cnpj, validate_cnpj_format, clean_cnpj

class PDFillerApp:
    def __init__(self, master):
        self.master = master
        master.title("Preencher solicitação de abertura - Individual e Lote")
        master.geometry("1400x1000") # Aumentando ainda mais para acomodar as abas
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
        self.setup_styles()

        # pdf_fields agora é importado de pdf_mapping.py
        self.pdf_fields = pdf_fields

        # Variáveis para armazenar os dados (aba individual)
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
        self.municipio_var = tk.StringVar()
        self.search_term_var = tk.StringVar()

        # Variáveis para a aba de lote
        self.batch_data_feriado_var = tk.StringVar()

        self.batch_output_dir_path = tk.StringVar()
        self.batch_filter_cnpj_var = tk.StringVar()
        self.batch_filter_razao_social_var = tk.StringVar()

        # Lista para armazenar empresas selecionadas para lote
        self.selected_companies = []

        # Widgets da interface
        self.create_widgets()
        
        # Carregar dados das empresas
        self.load_companies()

    def setup_styles(self):
        """Configurar estilos ttk"""
        style = ttk.Style()
        style.theme_use("clam")

        # Estilo para Labels
        style.configure("TLabel", 
                        font=("Arial", 10), 
                        background=self.colors['light_gray'], 
                        foreground=self.colors['text_dark'])
        
        # Estilo para Entry
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

        # Estilo para botão de lote
        style.configure("Batch.TButton", 
                        font=("Arial", 10, "bold"), 
                        padding=5, 
                        background='#dc3545', 
                        foreground=self.colors['white'],
                        relief="flat")
        style.map("Batch.TButton", 
                  background=[('active', '#c82333')],
                  foreground=[('active', self.colors['white'])])

        # Estilo para LabelFrames
        style.configure("TLabelframe.Label", 
                        font=("Arial", 12, "bold"), 
                        background=self.colors['light_gray'], 
                        foreground=self.colors['dark_blue'])
        style.configure("TLabelframe", 
                        background=self.colors['light_gray'],
                        bordercolor=self.colors['medium_gray'],
                        borderwidth=1,
                        relief="solid")

        # Estilo para Treeview
        style.configure("Treeview", 
                        font=("Arial", 9),
                        background=self.colors['white'],
                        foreground=self.colors['text_dark'],
                        fieldbackground=self.colors['white'])
        style.configure("Treeview.Heading", 
                        font=("Arial", 10, "bold"),
                        background=self.colors['primary_blue'],
                        foreground=self.colors['white'])

    def create_widgets(self):
        # Cabeçalho
        header_frame = tk.Frame(self.master, bg=self.colors['primary_blue'], height=40)
        header_frame.pack(fill='x', padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        header_label = tk.Label(header_frame, text="Preencher abertura de feriados - Individual e Lote", 
                                font=("Arial", 14, "bold"), 
                                fg=self.colors['white'], 
                                bg=self.colors['primary_blue'])
        header_label.pack(pady=8)

        # Frame principal para o conteúdo
        main_content_frame = tk.Frame(self.master, bg=self.colors["light_gray"])
        main_content_frame.pack(padx=10, pady=10, fill="both", expand=True)

        # Criar sistema de abas
        self.notebook = ttk.Notebook(main_content_frame)
        self.notebook.pack(fill="both", expand=True)

        # Criar abas
        self.create_individual_tab()
        self.create_batch_tab()

    def create_individual_tab(self):
        """Criar aba para preenchimento individual"""
        individual_frame = ttk.Frame(self.notebook)
        self.notebook.add(individual_frame, text="Preenchimento Individual")

        # Frame para pesquisa de empresas
        search_frame = ttk.LabelFrame(individual_frame, text="Pesquisar Empresa")
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
        input_frame = ttk.LabelFrame(individual_frame, text="Dados para Preenchimento")
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

        # Telefone
        ttk.Label(input_frame, text="Telefone:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(input_frame, textvariable=self.telefone_var, width=60).grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # Responsável
        ttk.Label(input_frame, text="Responsável:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(input_frame, textvariable=self.responsavel_var, width=60).grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        # Data do Feriado
        ttk.Label(input_frame, text="Data do Feriado:").grid(row=4, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(input_frame, textvariable=self.data_feriado_var, width=60).grid(row=4, column=1, padx=10, pady=5, sticky="ew")

        # Frame para seleção de arquivos
        file_frame = ttk.LabelFrame(individual_frame, text="Seleção de Arquivos")
        file_frame.pack(padx=10, pady=10, fill="x")
        file_frame.columnconfigure(1, weight=1)

        # Diretório de Saída
        ttk.Label(file_frame, text="Local para salvar:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(file_frame, textvariable=self.output_dir_path, width=50, state="readonly").grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        ttk.Button(file_frame, text="Selecionar Pasta", command=self.select_output_dir).grid(row=0, column=2, padx=10, pady=5)

        # Botão de Preenchimento
        fill_button = ttk.Button(individual_frame, text="Preencher PDF", command=self.fill_pdf)
        fill_button.pack(pady=10)

        # Frame para a tabela de empresas
        table_frame = ttk.LabelFrame(individual_frame, text="Empresas Cadastradas")
        table_frame.pack(padx=10, pady=10, fill="both", expand=True)

        # Criar a tabela (Treeview)
        columns = ("CNPJ", "Nome da Empresa", "Nome Fantasia", "Telefone", "Endereço", "Município")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=10)
        
        # Configurar cabeçalhos
        for col in columns:
            self.tree.heading(col, text=col)
            if col == "CNPJ":
                self.tree.column(col, width=120, minwidth=100)
            elif col == "Nome da Empresa":
                self.tree.column(col, width=150, minwidth=120)
            elif col == "Nome Fantasia":
                self.tree.column(col, width=150, minwidth=120)
            elif col == "Telefone":
                self.tree.column(col, width=100, minwidth=80)
            elif col == "Endereço":
                self.tree.column(col, width=250, minwidth=200)
            elif col == "Município":
                self.tree.column(col, width=150, minwidth=100)
        
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
        self.filter_cnpj_var.trace("w", self.filter_companies)
        filter_cnpj_entry = ttk.Entry(filter_frame, textvariable=self.filter_cnpj_var, width=20)
        filter_cnpj_entry.pack(side="left", padx=5)

        ttk.Label(filter_frame, text="Filtrar por Nome/Fantasia:").pack(side="left", padx=5)
        self.filter_razao_social_var.trace("w", self.filter_companies)
        filter_razao_social_entry = ttk.Entry(filter_frame, textvariable=self.filter_razao_social_var, width=30)
        filter_razao_social_entry.pack(side="left", padx=5)

    def create_batch_tab(self):
        """Criar aba para preenchimento em lote"""
        batch_frame = ttk.Frame(self.notebook)
        self.notebook.add(batch_frame, text="Preenchimento em Lote")

        # Frame para dados comuns do lote
        batch_input_frame = ttk.LabelFrame(batch_frame, text="Dados Comuns para Lote")
        batch_input_frame.pack(padx=10, pady=10, fill="x")
        batch_input_frame.columnconfigure(1, weight=1)

        # Responsável (comum para todas)


        # Data do Feriado (comum para todas)
        ttk.Label(batch_input_frame, text="Data do Feriado:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(batch_input_frame, textvariable=self.batch_data_feriado_var, width=60).grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        # Frame para seleção de pasta de saída
        batch_file_frame = ttk.LabelFrame(batch_frame, text="Local para Salvar PDFs")
        batch_file_frame.pack(padx=10, pady=10, fill="x")
        batch_file_frame.columnconfigure(1, weight=1)

        # Diretório de Saída
        ttk.Label(batch_file_frame, text="Local para salvar:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(batch_file_frame, textvariable=self.batch_output_dir_path, width=50, state="readonly").grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        ttk.Button(batch_file_frame, text="Selecionar Pasta", command=self.select_batch_output_dir).grid(row=0, column=2, padx=10, pady=5)

        # Frame para controles de seleção
        selection_frame = ttk.LabelFrame(batch_frame, text="Controles de Seleção")
        selection_frame.pack(padx=10, pady=10, fill="x")

        # Botões de seleção
        buttons_frame = ttk.Frame(selection_frame)
        buttons_frame.pack(pady=5)

        ttk.Button(buttons_frame, text="Selecionar Todas", command=self.select_all_companies).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Desmarcar Todas", command=self.deselect_all_companies).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Inverter Seleção", command=self.invert_selection).pack(side="left", padx=5)

        # Label para mostrar quantas empresas estão selecionadas
        self.selection_count_label = ttk.Label(selection_frame, text="Nenhuma empresa selecionada")
        self.selection_count_label.pack(pady=5)

        # Botão de processamento em lote
        batch_process_button = ttk.Button(batch_frame, text="Processar Lote", 
                                        command=self.process_batch, style="Batch.TButton")
        batch_process_button.pack(pady=10)

        # Frame para a tabela de empresas (com checkboxes)
        batch_table_frame = ttk.LabelFrame(batch_frame, text="Selecionar Empresas para Lote")
        batch_table_frame.pack(padx=10, pady=10, fill="both", expand=True)

        # Criar a tabela para lote (com checkbox)
        batch_columns = ("Selecionado", "CNPJ", "Nome da Empresa", "Nome Fantasia", "Telefone", "Endereço", "Município")
        self.batch_tree = ttk.Treeview(batch_table_frame, columns=batch_columns, show="headings", height=12)
        
        # Configurar cabeçalhos
        for col in batch_columns:
            self.batch_tree.heading(col, text=col)
            if col == "Selecionado":
                self.batch_tree.column(col, width=80, minwidth=80)
            elif col == "CNPJ":
                self.batch_tree.column(col, width=120, minwidth=100)
            elif col == "Nome da Empresa":
                self.batch_tree.column(col, width=150, minwidth=120)
            elif col == "Nome Fantasia":
                self.batch_tree.column(col, width=150, minwidth=120)
            elif col == "Telefone":
                self.batch_tree.column(col, width=100, minwidth=80)
            elif col == "Endereço":
                self.batch_tree.column(col, width=250, minwidth=200)
            elif col == "Município":
                self.batch_tree.column(col, width=150, minwidth=100)
        
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

        # Bind para clique na tabela (toggle selection)
        self.batch_tree.bind("<Button-1>", self.on_batch_tree_click)

        # Frame para filtros da aba de lote
        batch_filter_frame = ttk.Frame(batch_table_frame)
        batch_filter_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        
        ttk.Label(batch_filter_frame, text="Filtrar por CNPJ:").pack(side="left", padx=5)
        self.batch_filter_cnpj_var.trace("w", self.filter_batch_companies)
        batch_filter_cnpj_entry = ttk.Entry(batch_filter_frame, textvariable=self.batch_filter_cnpj_var, width=20)
        batch_filter_cnpj_entry.pack(side="left", padx=5)

        ttk.Label(batch_filter_frame, text="Filtrar por Nome/Fantasia:").pack(side="left", padx=5)
        self.batch_filter_razao_social_var.trace("w", self.filter_batch_companies)
        batch_filter_razao_social_entry = ttk.Entry(batch_filter_frame, textvariable=self.batch_filter_razao_social_var, width=30)
        batch_filter_razao_social_entry.pack(side="left", padx=5)

    # Métodos da aba individual (mantidos do código original)
    def search_and_fill_company(self):
        """Pesquisar empresa e preencher campos automaticamente"""
        search_term = self.search_term_var.get().strip()
        
        if not search_term:
            messagebox.showwarning("Aviso", "Por favor, digite um CNPJ ou razão social para pesquisar.")
            return
        
        try:
            conn = sqlite3.connect('empresas.db')
            cursor = conn.cursor()
            
            search_term_clean = ''.join(filter(str.isdigit, search_term))
            
            cursor.execute("""
                SELECT cnpj, razao_social, nome_fantasia, telefone, responsavel, endereco, cidade
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
                cnpj = str(result[0]).replace('.0', '') if result[0] else ''
                cnpj_formatted = format_cnpj(cnpj) if cnpj else ''
                
                razao_social = result[1] or ''
                nome_fantasia = result[2] or ''
                telefone = result[3] or ''
                responsavel = result[4] or ''
                endereco = result[5] or ''
                municipio = result[6] or ''
                
                self.cnpj_var.set(cnpj_formatted)
                self.razao_social_var.set(razao_social)
                self.nome_fantasia_var.set(nome_fantasia)
                self.telefone_var.set(telefone)
                self.responsavel_var.set(responsavel)
                self.endereco_var.set(endereco)
                self.municipio_var.set(municipio)
                self.search_term_var.set("")
                
                messagebox.showinfo("Sucesso", f"Empresa encontrada: {razao_social}")
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
        """Carrega as empresas do banco de dados nas tabelas"""
        try:
            conn = sqlite3.connect('empresas.db')
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT cnpj, razao_social, nome_fantasia, telefone, endereco, cidade
                FROM empresas
                ORDER BY razao_social
            """)
            
            companies = cursor.fetchall()
            
            # Limpar tabelas existentes
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Carregar dados na tabela individual
            for company in companies:
                cnpj = str(company[0]).replace('.0', '') if company[0] else ''
                cnpj_formatted = format_cnpj(cnpj) if cnpj else ''
                
                razao_social = company[1] or ''
                nome_fantasia = company[2] or ''
                telefone = company[3] or ''
                endereco = company[4] or ''
                municipio = company[5] or ''
                
                self.tree.insert("", "end", values=(cnpj_formatted, razao_social, nome_fantasia, telefone, endereco, municipio))
            
            # Carregar dados na tabela de lote também
            self.load_batch_companies()
            
            conn.close()
            
        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao carregar empresas: {e}")
        except Exception as e:
            print(f"Erro inesperado ao carregar empresas: {e}")

    def load_batch_companies(self):
        """Carrega as empresas na tabela de lote"""
        try:
            conn = sqlite3.connect('empresas.db')
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT cnpj, razao_social, nome_fantasia, telefone, endereco, cidade
                FROM empresas
                ORDER BY razao_social
            """)
            
            companies = cursor.fetchall()
            
            # Limpar tabela de lote
            for item in self.batch_tree.get_children():
                self.batch_tree.delete(item)
            
            # Carregar dados na tabela de lote
            for company in companies:
                cnpj = str(company[0]).replace('.0', '') if company[0] else ''
                cnpj_formatted = format_cnpj(cnpj) if cnpj else ''
                
                razao_social = company[1] or ''
                nome_fantasia = company[2] or ''
                telefone = company[3] or ''
                endereco = company[4] or ''
                municipio = company[5] or ''
                
                # Adicionar com checkbox desmarcado
                self.batch_tree.insert("", "end", values=("☐", cnpj_formatted, razao_social, nome_fantasia, telefone, endereco, municipio))
            
            conn.close()
            self.update_selection_count()
            
        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao carregar empresas para lote: {e}")

    # Métodos para a aba de lote
    def select_batch_output_dir(self):
        """Selecionar diretório de saída para lote"""
        directory = filedialog.askdirectory()
        if directory:
            self.batch_output_dir_path.set(directory)

    def on_batch_tree_click(self, event):
        """Manipular clique na tabela de lote para toggle de seleção"""
        item = self.batch_tree.identify('item', event.x, event.y)
        column = self.batch_tree.identify('column', event.x, event.y)
        
        if item and column == '#1':  # Primeira coluna (checkbox)
            values = list(self.batch_tree.item(item)['values'])
            if values[0] == "☐":
                values[0] = "☑"
            else:
                values[0] = "☐"
            
            self.batch_tree.item(item, values=values)
            self.update_selection_count()

    def select_all_companies(self):
        """Selecionar todas as empresas visíveis"""
        for item in self.batch_tree.get_children():
            values = list(self.batch_tree.item(item)['values'])
            values[0] = "☑"
            self.batch_tree.item(item, values=values)
        self.update_selection_count()

    def deselect_all_companies(self):
        """Desmarcar todas as empresas"""
        for item in self.batch_tree.get_children():
            values = list(self.batch_tree.item(item)['values'])
            values[0] = "☐"
            self.batch_tree.item(item, values=values)
        self.update_selection_count()

    def invert_selection(self):
        """Inverter seleção de todas as empresas visíveis"""
        for item in self.batch_tree.get_children():
            values = list(self.batch_tree.item(item)['values'])
            values[0] = "☑" if values[0] == "☐" else "☐"
            self.batch_tree.item(item, values=values)
        self.update_selection_count()

    def update_selection_count(self):
        """Atualizar contador de empresas selecionadas"""
        selected_count = 0
        total_count = 0
        
        for item in self.batch_tree.get_children():
            total_count += 1
            values = self.batch_tree.item(item)['values']
            if values[0] == "☑":
                selected_count += 1
        
        self.selection_count_label.config(text=f"{selected_count} de {total_count} empresas selecionadas")

    def get_selected_companies(self):
        """Obter lista de empresas selecionadas"""
        selected = []
        for item in self.batch_tree.get_children():
            values = self.batch_tree.item(item)['values']
            if values[0] == "☑":
                # Buscar dados completos da empresa no banco
                cnpj = values[1]  # CNPJ formatado
                try:
                    conn = sqlite3.connect('empresas.db')
                    cursor = conn.cursor()
                    
                    cnpj_clean = clean_cnpj(cnpj)
                    cursor.execute("""
                        SELECT cnpj, razao_social, nome_fantasia, telefone, endereco, cidade
                        FROM empresas 
                        WHERE REPLACE(REPLACE(REPLACE(cnpj, ".", ""), "/", ""), "-", "") = ?
                    """, (cnpj_clean,))
                    
                    result = cursor.fetchone()
                    if result:
                        company_data = {
                            'cnpj': format_cnpj(str(result[0]).replace('.0', '')) if result[0] else '',
                            'razao_social': result[1] or '',
                            'nome_fantasia': result[2] or '',
                            'telefone': result[3] or '',

                            'endereco': result[4] or '',
                            'municipio': result[5] or ''
                        }
                        selected.append(company_data)
                    
                    conn.close()
                    
                except sqlite3.Error as e:
                    print(f"Erro ao buscar dados da empresa {cnpj}: {e}")
        
        return selected

    def process_batch(self):
        """Processar lote de empresas"""
        # Validar dados comuns

        data_feriado = self.batch_data_feriado_var.get().strip()
        

        
        output_dir = self.batch_output_dir_path.get().strip()
        if not output_dir:
            messagebox.showwarning("Aviso", "Por favor, selecione o local para salvar os PDFs.")
            return
        
        # Obter empresas selecionadas
        selected_companies = self.get_selected_companies()
        
        if not selected_companies:
            messagebox.showwarning("Aviso", "Por favor, selecione pelo menos uma empresa.")
            return
        
        # Confirmar processamento
        result = messagebox.askyesno("Confirmar Processamento", 
                                   f"Processar {len(selected_companies)} empresas?\n\n"                                   f"Data do Feriado: {data_feriado}\n" +                                   f"Local de Saída: {output_dir}")
        
        if not result:
            return
        
        # Processar em thread separada para não travar a interface
        self.process_batch_thread(selected_companies, data_feriado, output_dir)

    def process_batch_thread(self, companies, data_feriado, output_dir):
        """Processar lote em thread separada"""
        def process():
            success_count = 0
            error_count = 0
            errors = []
            
            for i, company in enumerate(companies):
                try:
                    # Preparar dados para preenchimento
                    data = {
                        'cnpj': company['cnpj'],
                        'razao_social': company['razao_social'],
                        'nome_fantasia': company['nome_fantasia'],


                        'endereco': company['endereco'],
                        'municipio': company['municipio'],
                        'data_feriado': data_feriado  # Usar data comum
                    }
                    
                    # Gerar nome do arquivo
                    safe_name = "".join(c for c in company['razao_social'] if c.isalnum() or c in (' ', '-', '_')).rstrip()
                    safe_name = safe_name.replace(' ', '_')
                    output_filename = f"{safe_name}_{company['cnpj'].replace('.', '').replace('/', '').replace('-', '')}.pdf"
                    output_path = os.path.join(output_dir, output_filename)
                    
                    # Preencher PDF
                    fill_pdf_document(data, output_path, self.pdf_fields)
                    success_count += 1
                    
                    # Atualizar progresso na interface principal (thread-safe)
                    self.master.after(0, lambda i=i+1, total=len(companies): 
                                    self.update_batch_progress(i, total))
                    
                except Exception as e:
                    error_count += 1
                    errors.append(f"{company['razao_social']}: {str(e)}")
                    print(f"Erro ao processar {company['razao_social']}: {e}")
            
            # Mostrar resultado final
            self.master.after(0, lambda: self.show_batch_result(success_count, error_count, errors))
        
        # Executar em thread separada
        thread = threading.Thread(target=process)
        thread.daemon = True
        thread.start()

    def update_batch_progress(self, current, total):
        """Atualizar progresso do processamento em lote"""
        self.selection_count_label.config(text=f"Processando... {current}/{total}")

    def show_batch_result(self, success_count, error_count, errors):
        """Mostrar resultado do processamento em lote"""
        self.update_selection_count()  # Restaurar contador normal
        
        message = f"Processamento concluído!\n\n"
        message += f"Sucessos: {success_count}\n"
        message += f"Erros: {error_count}\n"
        
        if errors:
            message += f"\nErros encontrados:\n"
            for error in errors[:5]:  # Mostrar apenas os primeiros 5 erros
                message += f"• {error}\n"
            if len(errors) > 5:
                message += f"... e mais {len(errors) - 5} erros."
        
        if error_count > 0:
            messagebox.showwarning("Processamento Concluído com Erros", message)
        else:
            messagebox.showinfo("Processamento Concluído", message)

    def filter_batch_companies(self, *args):
        """Filtrar empresas na tabela de lote"""
        cnpj_filter = self.batch_filter_cnpj_var.get().lower()
        name_filter = self.batch_filter_razao_social_var.get().lower()
        

        
        if not cnpj_filter and not name_filter:
            return
        
        # Filtrar itens
        for item in self.batch_tree.get_children():
            values = self.batch_tree.item(item)['values']
            cnpj = str(values[1]).lower()
            razao_social = str(values[2]).lower()
            nome_fantasia = str(values[3]).lower()
            
            show_item = True
            
            if cnpj_filter and cnpj_filter not in cnpj:
                show_item = False
            
            if name_filter and name_filter not in razao_social and name_filter not in nome_fantasia:
                show_item = False
            
            if not show_item:
                self.batch_tree.detach(item)
            else:
                self.batch_tree.reattach(item, '', 'end')

    # Métodos originais mantidos
    def on_company_select(self, event):
        """Manipular seleção de empresa na tabela individual"""
        try:
            selection = self.tree.selection()
            if selection:
                item = selection[0]
                values = self.tree.item(item)['values']
                
                if values:
                    cnpj = values[0]
                    
                    # Buscar dados completos da empresa
                    conn = sqlite3.connect('empresas.db')
                    cursor = conn.cursor()
                    
                    cnpj_clean = clean_cnpj(cnpj)
                    cursor.execute("""
                        SELECT cnpj, razao_social, nome_fantasia, telefone, endereco, cidade
                        FROM empresas 
                        WHERE REPLACE(REPLACE(REPLACE(cnpj, ".", ""), "/", ""), "-", "") = ?
                    """, (cnpj_clean,))
                    
                    result = cursor.fetchone()
                    
                    if result:
                        cnpj_formatted = format_cnpj(str(result[0]).replace('.0', '')) if result[0] else ''
                        
                        self.cnpj_var.set(cnpj_formatted)
                        self.razao_social_var.set(result[1] or '')
                        self.nome_fantasia_var.set(result[2] or '')
                        self.telefone_var.set(result[3] or '')
                        self.responsavel_var.set(result[4] or '')
                        self.endereco_var.set(result[5] or '')
                        self.municipio_var.set(result[6] or '')
                    
                    conn.close()
                    
        except Exception as e:
            print(f"Erro ao selecionar empresa: {e}")

    def filter_companies(self, *args):
        """Filtrar empresas na tabela individual"""
        cnpj_filter = self.filter_cnpj_var.get().lower()
        name_filter = self.filter_razao_social_var.get().lower()
        
        # Recarregar todas as empresas
        self.load_companies()
        
        if not cnpj_filter and not name_filter:
            return
        
        # Filtrar itens
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            cnpj = str(values[0]).lower()
            razao_social = str(values[1]).lower()
            nome_fantasia = str(values[2]).lower()
            
            show_item = True
            
            if cnpj_filter and cnpj_filter not in cnpj:
                show_item = False
            
            if name_filter and name_filter not in razao_social and name_filter not in nome_fantasia:
                show_item = False
            
            if not show_item:
                self.tree.delete(item)

    def select_output_dir(self):
        """Selecionar diretório de saída"""
        directory = filedialog.askdirectory()
        if directory:
            self.output_dir_path.set(directory)

    def fill_pdf(self):
        """Preencher PDF individual"""
        # Validar dados
        cnpj = self.cnpj_var.get().strip()
        razao_social = self.razao_social_var.get().strip()
        telefone = self.telefone_var.get().strip()
        data_feriado = self.data_feriado_var.get().strip()
        output_dir = self.output_dir_path.get().strip()
        
        if not all([cnpj, razao_social, responsavel, data_feriado, output_dir]):
            messagebox.showwarning("Aviso", "Por favor, preencha todos os campos obrigatórios.")
            return
        
        try:
            # Preparar dados
            data = {
                'cnpj': cnpj,
                'razao_social': razao_social,
                'nome_fantasia': self.nome_fantasia_var.get().strip(),
                'telefone': telefone,
                'responsavel': responsavel,
                'endereco': self.endereco_var.get().strip(),
                'municipio': self.municipio_var.get().strip(),
                'data_feriado': data_feriado
            }
            
            # Gerar nome do arquivo
            safe_name = "".join(c for c in razao_social if c.isalnum() or c in (' ', '-', '_')).rstrip()
            safe_name = safe_name.replace(' ', '_')
            output_filename = f"{safe_name}_{cnpj.replace('.', '').replace('/', '').replace('-', '')}.pdf"
            output_path = os.path.join(output_dir, output_filename)
            
            # Preencher PDF
            fill_pdf_document(data, output_path, self.pdf_fields)
            
            messagebox.showinfo("Sucesso", f"PDF preenchido com sucesso!\nSalvo em: {output_path}")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao preencher PDF: {e}")

    def format_cnpj_on_type(self, event):
        """Formatar CNPJ enquanto digita"""
        try:
            current_text = self.cnpj_var.get()
            numbers_only = ''.join(filter(str.isdigit, current_text))
            
            if len(numbers_only) <= 14:
                formatted = format_cnpj(numbers_only)
                if formatted != current_text:
                    cursor_pos = self.cnpj_entry.index(tk.INSERT)
                    self.cnpj_var.set(formatted)
                    # Tentar manter a posição do cursor
                    try:
                        self.cnpj_entry.icursor(min(cursor_pos, len(formatted)))
                    except:
                        pass
        except:
            pass

    def validate_cnpj_on_focus_out(self, event):
        """Validar CNPJ quando o campo perde o foco"""
        cnpj_text = self.cnpj_var.get().strip()
        
        if cnpj_text and not validate_cnpj_format(cnpj_text):
            numbers_only = clean_cnpj(cnpj_text)
            if len(numbers_only) == 14:
                formatted_cnpj = format_cnpj(numbers_only)
                self.cnpj_var.set(formatted_cnpj)
            elif len(numbers_only) > 0:
                messagebox.showwarning("CNPJ Inválido", 
                                     f"CNPJ deve ter 14 dígitos. Você digitou {len(numbers_only)} dígitos.")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFillerApp(root)
    root.mainloop()



        self.update_selection_count()
