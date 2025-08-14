import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sqlite3
import os

# Importar a função de preenchimento e o mapeamento
from pdf_filler import fill_pdf_document
from pdf_mapping import pdf_fields

class PDFillerApp:
    def __init__(self, master):
        self.master = master
        master.title("Preencher solicitação de abertura")
        master.geometry("1200x900") # Aumentando o tamanho da janela para acomodar a nova funcionalidade
        master.resizable(True, True) # Permitir redimensionamento
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
        self.endereco_var = tk.StringVar()
        self.data_feriado_var = tk.StringVar()
        self.filter_cnpj_var = tk.StringVar()
        self.filter_razao_social_var = tk.StringVar()
        self.output_dir_path = tk.StringVar()
        
        # Variáveis para pesquisa
        self.search_term_var = tk.StringVar()

        # Widgets da interface
        self.create_widgets()
        
        # Carregar dados das empresas
        self.load_companies()

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

        # NOVO: Frame para pesquisa de empresas
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

        # Configurar grid para centralizar e expandir
        input_frame.columnconfigure(1, weight=1)

        # CNPJ
        ttk.Label(input_frame, text="CNPJ:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(input_frame, textvariable=self.cnpj_var, width=60).grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        # Nome da Empresa
        ttk.Label(input_frame, text="Nome da Empresa:").grid(row=1, column=0, sticky="w", padx=10, pady=5)        
        ttk.Entry(input_frame, textvariable=self.razao_social_var, width=60).grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        # Endereço
        ttk.Label(input_frame, text="Endereço:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(input_frame, textvariable=self.endereco_var, width=60).grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # Dia do Feriado
        ttk.Label(input_frame, text="Dia do Feriado:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(input_frame, textvariable=self.data_feriado_var, width=60).grid(row=3, column=1, padx=10, pady=5, sticky="ew")

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
        columns = ("CNPJ", "Nome da Empresa", "Nome Fantasia", "Endereço", "Município", "Situação")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=15)
        
        # Configurar cabeçalhos
        for col in columns:
            self.tree.heading(col, text=col)
            if col == "CNPJ":
                self.tree.column(col, width=120, minwidth=100)
            elif col == "Nome da Empresa":
                self.tree.column(col, width=150, minwidth=120)
            elif col == "Nome Fantasia":
                self.tree.column(col, width=150, minwidth=120)
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
        self.filter_cnpj_var = tk.StringVar()
        self.filter_cnpj_var.trace("w", self.filter_companies)
        filter_cnpj_entry = ttk.Entry(filter_frame, textvariable=self.filter_cnpj_var, width=20)
        filter_cnpj_entry.pack(side="left", padx=5)

        ttk.Label(filter_frame, text="Filtrar por Nome/Fantasia:").pack(side="left", padx=5)
        self.filter_razao_social_var = tk.StringVar()
        self.filter_razao_social_var.trace("w", self.filter_companies)
        filter_razao_social_entry = ttk.Entry(filter_frame, textvariable=self.filter_razao_social_var, width=30)
        filter_razao_social_entry.pack(side="left", padx=5)

    def search_and_fill_company(self):
        """Nova função para pesquisar empresa e preencher campos automaticamente"""
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
                SELECT cnpj, razao_social, nome_fantasia, endereco, complemento, 
                       bairro, municipio, cep, uf, telefone, email
                FROM empresas 
                WHERE REPLACE(REPLACE(REPLACE(cnpj, '.', ''), '/', ''), '-', '') LIKE ? 
                   OR UPPER(razao_social) LIKE UPPER(?) 
                   OR UPPER(nome_fantasia) LIKE UPPER(?)
                ORDER BY 
                    CASE 
                        WHEN REPLACE(REPLACE(REPLACE(cnpj, '.', ''), '/', ''), '-', '') = ? THEN 1
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
                razao_social = result[1] or ''
                nome_fantasia = result[2] or ''
                endereco = result[3] or ''
                complemento = result[4] or ''
                bairro = result[5] or ''
                municipio = result[6] or ''
                cep = result[7] or ''
                uf = result[8] or ''
                telefone = result[9] or ''
                email = result[10] or ''
                
                # Montar endereço completo
                endereco_completo = []
                if endereco:
                    endereco_completo.append(endereco)
                if complemento:
                    endereco_completo.append(complemento)
                if bairro:
                    endereco_completo.append(bairro)
                if municipio:
                    endereco_completo.append(municipio)
                if uf:
                    endereco_completo.append(uf)
                if cep:
                    endereco_completo.append(f"CEP: {cep}")
                
                # Preencher os campos da interface
                self.cnpj_var.set(cnpj)
                self.razao_social_var.set(razao_social)
                self.endereco_var.set(", ".join(endereco_completo))
                
                # Limpar campo de pesquisa
                self.search_term_var.set("")
                
                # Mostrar mensagem de sucesso
                messagebox.showinfo("Sucesso", f"Empresa encontrada: {razao_social}")
                
                # Destacar a empresa na tabela se estiver visível
                self.highlight_company_in_table(cnpj)
                
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
        """Carrega as empresas do banco de dados na tabela"""
        try:
            conn = sqlite3.connect('empresas.db')
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT cnpj, razao_social, nome_fantasia, endereco, municipio, situacao 
                FROM empresas 
                ORDER BY razao_social
            """)
            
            self.companies_data = cursor.fetchall()
            self.display_companies(self.companies_data)
            
            conn.close()
            
        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao carregar empresas: {e}")

    def display_companies(self, companies):
        """Exibe as empresas na tabela"""
        # Limpar tabela
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Adicionar empresas
        for company in companies:
            # Formatar CNPJ se for numérico
            cnpj = str(company[0])
            if cnpj.replace('.', '').isdigit():
                cnpj = cnpj.replace('.0', '')  # Remove .0 se existir
            
            self.tree.insert("", "end", values=(
                cnpj,
                company[1] or "",  # razao_social
                company[2] or "",  # nome_fantasia
                company[3] or "",  # endereco
                company[4] or "",  # municipio
                company[5] or ""   # situacao
            ))

    def filter_companies(self, *args):
        """Filtra as empresas baseado no texto digitado"""
        filter_cnpj_text = self.filter_cnpj_var.get().lower()
        filter_razao_social_text = self.filter_razao_social_var.get().lower()
        
        if not filter_cnpj_text and not filter_razao_social_text:
            self.display_companies(self.companies_data)
            return
        
        filtered_companies = []
        for company in self.companies_data:
            cnpj = str(company[0]).lower()
            razao_social = (company[1] or "").lower()
            nome_fantasia = (company[2] or "").lower()
            
            cnpj_match = filter_cnpj_text in cnpj
            name_match = filter_razao_social_text in razao_social or filter_razao_social_text in nome_fantasia
            
            if (not filter_cnpj_text or cnpj_match) and (not filter_razao_social_text or name_match):
                filtered_companies.append(company)
        
        self.display_companies(filtered_companies)

    def on_company_select(self, event):
        """Preenche os campos quando uma empresa é selecionada"""
        selection = self.tree.selection()
        if selection:
            item = self.tree.item(selection[0])
            values = item['values']
            
            # Preencher campos
            self.cnpj_var.set(values[0])
            self.razao_social_var.set(values[1])  # Nome da Empresa (Razão Social)            
            # Preencher endereço diretamente da tabela se disponível
            if values[3]:  # Se há endereço na tabela
                self.endereco_var.set(values[3])
            else:
                # Buscar endereço completo no banco se não estiver na tabela
                try:
                    conn = sqlite3.connect('empresas.db')
                    cursor = conn.cursor()
                    
                    cursor.execute("""
                        SELECT endereco, complemento, bairro, municipio, cep
                        FROM empresas 
                        WHERE cnpj = ?
                    """, (values[0],))
                    
                    result = cursor.fetchone()
                    if result:
                        endereco_completo = []
                        if result[0]:  # endereco
                            endereco_completo.append(result[0])
                        if result[1]:  # complemento
                            endereco_completo.append(result[1])
                        if result[2]:  # bairro
                            endereco_completo.append(result[2])
                        if result[3]:  # municipio
                            endereco_completo.append(result[3])
                        if result[4]:  # cep
                            endereco_completo.append(f"CEP: {result[4]}")
                        
                        self.endereco_var.set(", ".join(endereco_completo))
                    
                    conn.close()
                    
                except sqlite3.Error as e:
                    print(f"Erro ao buscar endereço: {e}")

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

        # Verificar se há uma empresa selecionada na tabela para obter o nome fantasia
        nome_fantasia = ""
        selection = self.tree.selection()
        if selection:
            item = self.tree.item(selection[0])
            values = item['values']
            nome_fantasia = values[2] if len(values) > 2 else ""

        data_to_fill = {
            'cnpj': self.cnpj_var.get(),
            'razao_social': self.razao_social_var.get(),
            'endereco': self.endereco_var.get(),
            'data': self.data_feriado_var.get(),
            'nome_fantasia': nome_fantasia,
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

