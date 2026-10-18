"""
POC - Sistema de Mapeamento e Gestão de Estoque Farmacêutico
Inspirado na Identidade Visual Prati-Donaduzzi (Roxo, Preto e Branco)
Tecnologia: Python Nativo (Tkinter + SQLite3) - 100% Portável, Zero Instalação
"""

import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox
import re
import os

# --- PALETA DE CORES PRATI-DONADUZZI ---
COLOR_BG_DARK = "#121214"       # Fundo principal escuro
COLOR_CARD_DARK = "#1E1E24"     # Fundo de cartões e painéis
COLOR_PRIMARY = "#5B2C82"       # Roxo institucional principal
COLOR_PRIMARY_HOVER = "#7338A3" # Roxo hover
COLOR_ACCENT = "#9B51E0"        # Roxo vibrante para destaques
COLOR_TEXT_MAIN = "#FFFFFF"     # Texto principal
COLOR_TEXT_MUTED = "#A0A0B0"    # Texto secundário
COLOR_BORDER = "#2E2E38"        # Bordas
COLOR_SLOT_EMPTY = "#272730"    # Posição vazia
COLOR_SLOT_OCCUPIED = "#5B2C82" # Posição ocupada
COLOR_ALERT = "#E06A3B"         # Alerta estoque baixo

DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "estoque_medicamentos.db")


class DatabaseManager:
    """Gerenciador de Banco de Dados Local SQLite (Zero Instalação)."""
    def __init__(self, db_path=DB_FILE):
        self.db_path = db_path
        self.init_database()

    def get_connection(self):
        return sqlite3.connect(self.db_path)

    def init_database(self):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS itens_estoque (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    rua TEXT NOT NULL,
                    bloco INTEGER,
                    posicao INTEGER,
                    codigo_localizacao TEXT NOT NULL,
                    codigo_produto TEXT NOT NULL,
                    nome_medicamento TEXT NOT NULL,
                    lote TEXT NOT NULL,
                    quantidade INTEGER NOT NULL CHECK(quantidade >= 0),
                    criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_busca ON itens_estoque (codigo_produto, lote, nome_medicamento)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_loc ON itens_estoque (codigo_localizacao)")
            conn.commit()
            
            # Insere dados de demonstração caso a base esteja vazia
            cursor.execute("SELECT COUNT(*) FROM itens_estoque")
            if cursor.fetchone()[0] == 0:
                self._seed_demo_data(cursor)
                conn.commit()

    def _seed_demo_data(self, cursor):
        demo_items = [
            ("Rua C", 28, 9, "C2809", "MED-00102", "Dipirona 500mg/mL Solução", "L24089A", 120),
            ("Rua C", 28, 10, "C2810", "MED-00103", "Paracetamol 750mg Comprimidos", "L24090B", 45),
            ("Rua C", 1, 1, "C0101", "MED-00104", "Amoxicilina 500mg Cápsulas", "L24101A", 80),
            ("Rua C", 2, 3, "C0203", "MED-00105", "Ibuprofeno 400mg Drágeas", "L24105C", 15),
            ("Rua A", 2, 1, "A0201", "MED-00201", "Omeprazol 20mg Cápsulas", "L24011D", 200),
            ("Rua B", 4, 5, "B0405", "MED-00301", "Losartana Potássica 50mg", "L24022E", 60),
            ("Frações", 0, 1, "FRAC-01", "MED-00401", "Clonazepam 2,5mg/mL Gotas", "L24033F", 8),
            ("Frações", 0, 2, "FRAC-02", "MED-00402", "Sinvastatina 20mg", "L24034G", 4)
        ]
        cursor.executemany("""
            INSERT INTO itens_estoque (rua, bloco, posicao, codigo_localizacao, codigo_produto, nome_medicamento, lote, quantidade)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, demo_items)


def parse_localizacao(texto):
    """
    Decodifica códigos como 'C2809' -> Rua C, Bloco 28 (Direito), Posição 09.
    Ou aceita identificação de 'FRAC-01', 'FRAC' para Frações.
    """
    texto = texto.strip().upper()
    if texto.startswith("FRAC") or "FRA" in texto:
        return {"rua": "Frações", "bloco": 0, "posicao": 1, "codigo": texto if texto.startswith("FRAC") else "FRAC-01", "valido": True}
    
    match = re.match(r"^([A-D])(\d{1,2})(\d{2})$", texto)
    if match:
        rua_letra = match.group(1)
        bloco = int(match.group(2))
        posicao = int(match.group(3))
        
        # Validação de lado par/ímpar e limites de posições (1 a 14)
        if 1 <= posicao <= 14:
            rua_nome = f"Rua {rua_letra}"
            return {
                "rua": rua_nome,
                "bloco": bloco,
                "posicao": posicao,
                "codigo": f"{rua_letra}{bloco:02d}{posicao:02d}",
                "lado": "Direito (Par)" if bloco % 2 == 0 else "Esquerdo (Ímpar)",
                "valido": True
            }
    return {"valido": False}


class EstoqueApp(tk.Tk):
    """Interface Gráfica do POC no padrão Prati-Donaduzzi."""
    def __init__(self):
        super().__init__()
        self.title("Prati WMS - Sistema de Mapeamento de Estoque (POC)")
        self.geometry("1180x720")
        self.minsize(980, 600)
        self.configure(bg=COLOR_BG_DARK)

        self.db = DatabaseManager()
        self.rua_selecionada = tk.StringVar(value="Rua C")
        self.selected_item_id = None

        self._configurar_estilos()
        self._construir_interface()
        self._carregar_tabela()
        self._desenhar_mapa_2d()

    def _configurar_estilos(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        
        style.configure("Treeview", 
                        background=COLOR_CARD_DARK, 
                        foreground=COLOR_TEXT_MAIN, 
                        fieldbackground=COLOR_CARD_DARK, 
                        rowheight=28,
                        bordercolor=COLOR_BORDER,
                        borderwidth=0)
        style.configure("Treeview.Heading", 
                        background=COLOR_PRIMARY, 
                        foreground=COLOR_TEXT_MAIN, 
                        relief="flat",
                        font=("Segoe UI", 10, "bold"))
        style.map("Treeview", 
                  background=[("selected", COLOR_ACCENT)],
                  foreground=[("selected", COLOR_TEXT_MAIN)])

    def _construir_interface(self):
        # 1. Header Superior Institucional
        header = tk.Frame(self, bg=COLOR_PRIMARY, height=54)
        header.pack(fill=tk.X, side=tk.TOP)

        lbl_logo = tk.Label(header, text="◆ PRATI LOGÍSTICA", font=("Segoe UI", 14, "bold"), fg=COLOR_TEXT_MAIN, bg=COLOR_PRIMARY)
        lbl_logo.pack(side=tk.LEFT, padx=16, pady=10)

        lbl_sub = tk.Label(header, text="Mapeamento 2D de Estoque & Controle de Lotes", font=("Segoe UI", 10), fg="#E2D0F0", bg=COLOR_PRIMARY)
        lbl_sub.pack(side=tk.LEFT, padx=6)

        # 2. Container Principal Dividido (Esquerda: Mapa 2D, Direita: Gestão e Tabela)
        main_container = tk.Frame(self, bg=COLOR_BG_DARK)
        main_container.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        # --- PAINEL ESQUERDO: MAPA 2D ---
        painel_mapa = tk.Frame(main_container, bg=COLOR_CARD_DARK, bd=1, relief=tk.SOLID)
        painel_mapa.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))

        # Seletor de Ruas
        toolbar_mapa = tk.Frame(painel_mapa, bg=COLOR_CARD_DARK)
        toolbar_mapa.pack(fill=tk.X, padx=10, pady=8)

        tk.Label(toolbar_mapa, text="Visualizar:", font=("Segoe UI", 10, "bold"), fg=COLOR_TEXT_MAIN, bg=COLOR_CARD_DARK).pack(side=tk.LEFT, padx=(0, 6))

        for rua in ["Rua A", "Rua B", "Rua C", "Rua D", "Frações"]:
            btn = tk.Radiobutton(toolbar_mapa, text=rua, value=rua, variable=self.rua_selecionada,
                                 indicatoron=0, selectcolor=COLOR_PRIMARY, fg=COLOR_TEXT_MAIN, bg=COLOR_BORDER,
                                 activebackground=COLOR_PRIMARY_HOVER, activeforeground=COLOR_TEXT_MAIN,
                                 padx=10, pady=3, font=("Segoe UI", 9, "bold"),
                                 command=self._on_rua_changed)
            btn.pack(side=tk.LEFT, padx=3)

        # Canvas do Mapa 2D
        self.canvas_mapa = tk.Canvas(painel_mapa, bg="#16161A", highlightthickness=0)
        self.canvas_mapa.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        self.canvas_mapa.bind("<Configure>", lambda e: self._desenhar_mapa_2d())

        # Legenda no rodapé do mapa
        legenda_frame = tk.Frame(painel_mapa, bg=COLOR_CARD_DARK)
        legenda_frame.pack(fill=tk.X, padx=10, pady=(0, 8))
        self._criar_legenda_item(legenda_frame, COLOR_SLOT_EMPTY, "Livre")
        self._criar_legenda_item(legenda_frame, COLOR_SLOT_OCCUPIED, "Ocupado")
        self._criar_legenda_item(legenda_frame, COLOR_ALERT, "Estoque Crítico (<20 un)")

        # --- PAINEL DIREITO: FORMULÁRIO + BUSCA + LISTAGEM ---
        painel_dados = tk.Frame(main_container, bg=COLOR_CARD_DARK, width=480, bd=1, relief=tk.SOLID)
        painel_dados.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(8, 0))

        # Barra de Pesquisa Rápida Inteligente
        frame_busca = tk.Frame(painel_dados, bg=COLOR_CARD_DARK)
        frame_busca.pack(fill=tk.X, padx=12, pady=10)

        tk.Label(frame_busca, text="🔍 Busca Rápida:", font=("Segoe UI", 10, "bold"), fg=COLOR_TEXT_MAIN, bg=COLOR_CARD_DARK).pack(anchor="w")
        
        self.entry_busca = tk.Entry(frame_busca, bg="#2A2A36", fg=COLOR_TEXT_MAIN, insertbackground=COLOR_TEXT_MAIN,
                                    font=("Segoe UI", 10), bd=1, relief=tk.FLAT)
        self.entry_busca.pack(fill=tk.X, pady=4, ipady=4)
        self.entry_busca.bind("<KeyRelease>", lambda e: self._carregar_tabela())

        # Formulário de Cadastro / Edição Rápida
        frame_form = tk.LabelFrame(painel_dados, text=" Cadastrar / Editar Medicamento ", 
                                   font=("Segoe UI", 9, "bold"), fg=COLOR_ACCENT, bg=COLOR_CARD_DARK, bd=1)
        frame_form.pack(fill=tk.X, padx=12, pady=6)

        campos = [
            ("Código da Localização (ex: C2809 ou FRAC):", "loc"),
            ("Código do Medicamento:", "codigo"),
            ("Nome do Medicamento:", "nome"),
            ("Lote:", "lote"),
            ("Quantidade:", "qtd")
        ]

        self.form_entries = {}
        for idx, (label_txt, key) in enumerate(campos):
            f_row = tk.Frame(frame_form, bg=COLOR_CARD_DARK)
            f_row.pack(fill=tk.X, padx=8, pady=2)
            tk.Label(f_row, text=label_txt, font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg=COLOR_CARD_DARK).pack(anchor="w")
            ent = tk.Entry(f_row, bg="#2A2A36", fg=COLOR_TEXT_MAIN, insertbackground=COLOR_TEXT_MAIN, font=("Segoe UI", 9), bd=1, relief=tk.FLAT)
            ent.pack(fill=tk.X, ipady=2)
            self.form_entries[key] = ent

        # Botões de Ação
        frame_botoes = tk.Frame(frame_form, bg=COLOR_CARD_DARK)
        frame_botoes.pack(fill=tk.X, padx=8, pady=8)

        btn_salvar = tk.Button(frame_botoes, text="Salvar Item", bg=COLOR_PRIMARY, fg=COLOR_TEXT_MAIN,
                               activebackground=COLOR_PRIMARY_HOVER, activeforeground=COLOR_TEXT_MAIN,
                               font=("Segoe UI", 9, "bold"), relief=tk.FLAT, padx=12, pady=4, command=self._salvar_item)
        btn_salvar.pack(side=tk.LEFT, padx=(0, 4))

        btn_limpar = tk.Button(frame_botoes, text="Limpar", bg=COLOR_BORDER, fg=COLOR_TEXT_MAIN,
                               font=("Segoe UI", 9), relief=tk.FLAT, padx=10, pady=4, command=self._limpar_formulario)
        btn_limpar.pack(side=tk.LEFT, padx=4)

        btn_remover = tk.Button(frame_botoes, text="Remover Selecionado", bg="#8B0000", fg=COLOR_TEXT_MAIN,
                                font=("Segoe UI", 9), relief=tk.FLAT, padx=10, pady=4, command=self._remover_item)
        btn_remover.pack(side=tk.RIGHT)

        # Tabela de Itens (Treeview)
        tk.Label(painel_dados, text="Itens Cadastrados no Estoque:", font=("Segoe UI", 9, "bold"), fg=COLOR_TEXT_MAIN, bg=COLOR_CARD_DARK).pack(anchor="w", padx=12, pady=(8, 2))

        tabela_frame = tk.Frame(painel_dados, bg=COLOR_CARD_DARK)
        tabela_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 10))

        colunas = ("id", "localizacao", "codigo", "nome", "lote", "qtd")
        self.tree = ttk.Treeview(tabela_frame, columns=colunas, show="headings", selectmode="browse")
        self.tree.heading("id", text="#")
        self.tree.heading("localizacao", text="Local")
        self.tree.heading("codigo", text="Código")
        self.tree.heading("nome", text="Nome")
        self.tree.heading("lote", text="Lote")
        self.tree.heading("qtd", text="Qtd")

        self.tree.column("id", width=30, anchor="center")
        self.tree.column("localizacao", width=65, anchor="center")
        self.tree.column("codigo", width=80, anchor="w")
        self.tree.column("nome", width=140, anchor="w")
        self.tree.column("lote", width=70, anchor="center")
        self.tree.column("qtd", width=45, anchor="e")

        scroll_y = ttk.Scrollbar(tabela_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<<TreeviewSelect>>", self._on_tabela_select)

    def _criar_legenda_item(self, parent, cor, texto):
        f = tk.Frame(parent, bg=COLOR_CARD_DARK)
        f.pack(side=tk.LEFT, padx=6)
        c = tk.Canvas(f, width=12, height=12, bg=COLOR_CARD_DARK, highlightthickness=0)
        c.pack(side=tk.LEFT, padx=(0, 4))
        c.create_rectangle(0, 0, 12, 12, fill=cor, outline="")
        tk.Label(f, text=texto, font=("Segoe UI", 8), fg=COLOR_TEXT_MUTED, bg=COLOR_CARD_DARK).pack(side=tk.LEFT)

    def _on_rua_changed(self):
        self._desenhar_mapa_2d()

    def _desenhar_mapa_2d(self):
        """Renderiza o Mapa 2D de corredor e blocos na rua selecionada."""
        self.canvas_mapa.delete("all")
        w = self.canvas_mapa.winfo_width()
        h = self.canvas_mapa.winfo_height()
        if w < 100 or h < 100:
            return

        rua = self.rua_selecionada.get()

        # Obter itens ocupados nesta rua do banco
        with self.db.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT codigo_localizacao, quantidade, nome_medicamento, lote FROM itens_estoque WHERE rua = ?", (rua,))
            ocupados = {row[0]: {"qtd": row[1], "nome": row[2], "lote": row[3]} for row in cursor.fetchall()}

        if rua == "Frações":
            self._desenhar_palete_fracoes(w, h, ocupados)
        else:
            self._desenhar_corredor_rua(w, h, rua, ocupados)

    def _desenhar_palete_fracoes(self, w, h, ocupados):
        """Visualização específica para o Palete de Frações."""
        self.canvas_mapa.create_text(w/2, 28, text="ÁREA DE FRAÇÕES (PALETE AVULSO)", font=("Segoe UI", 12, "bold"), fill=COLOR_TEXT_MAIN)
        self.canvas_mapa.create_text(w/2, 48, text="Medicamentos em pequenas quantidades fracionadas", font=("Segoe UI", 9), fill=COLOR_TEXT_MUTED)

        # Desenhar grade de 12 posições fracionadas
        cols = 4
        rows = 3
        slot_w = 100
        slot_h = 70
        start_x = (w - (cols * slot_w + (cols - 1) * 15)) / 2
        start_y = 90

        for r in range(rows):
            for c in range(cols):
                num = r * cols + c + 1
                cod = f"FRAC-{num:02d}"
                x1 = start_x + c * (slot_w + 15)
                y1 = start_y + r * (slot_h + 15)
                x2 = x1 + slot_w
                y2 = y1 + slot_h

                info = ocupados.get(cod)
                cor = COLOR_SLOT_EMPTY
                if info:
                    cor = COLOR_ALERT if info["qtd"] < 20 else COLOR_SLOT_OCCUPIED

                self.canvas_mapa.create_rectangle(x1, y1, x2, y2, fill=cor, outline=COLOR_BORDER, width=2)
                self.canvas_mapa.create_text((x1+x2)/2, y1 + 16, text=cod, font=("Segoe UI", 9, "bold"), fill=COLOR_TEXT_MAIN)
                
                if info:
                    self.canvas_mapa.create_text((x1+x2)/2, y1 + 36, text=f"{info['qtd']} un", font=("Segoe UI", 9, "bold"), fill="#FFD700")
                    nome_curto = info['nome'][:12] + ".." if len(info['nome']) > 12 else info['nome']
                    self.canvas_mapa.create_text((x1+x2)/2, y1 + 52, text=nome_curto, font=("Segoe UI", 7), fill=COLOR_TEXT_MAIN)
                else:
                    self.canvas_mapa.create_text((x1+x2)/2, y1 + 42, text="Livre", font=("Segoe UI", 8), fill=COLOR_TEXT_MUTED)

    def _desenhar_corredor_rua(self, w, h, rua, ocupados):
        """Desenha corredor com Lado Esquerdo (Ímpares) e Lado Direito (Pares)."""
        letra_rua = rua.replace("Rua ", "")
        self.canvas_mapa.create_text(w/2, 22, text=f"MAPA 2D: {rua.upper()}", font=("Segoe UI", 12, "bold"), fill=COLOR_TEXT_MAIN)

        # Corredor central
        corredor_w = 50
        cx = w / 2
        self.canvas_mapa.create_rectangle(cx - corredor_w/2, 45, cx + corredor_w/2, h - 20, fill="#1F1F24", outline="")
        self.canvas_mapa.create_text(cx, h/2, text="CORREDOR\nCENTRAL", font=("Segoe UI", 8, "bold"), fill="#444452", justify=tk.CENTER)

        # Lado Esquerdo (Blocos Ímpares: 01, 03)
        self.canvas_mapa.create_text(cx - 140, 40, text="LADO ESQUERDO (ÍMPAR)", font=("Segoe UI", 9, "bold"), fill=COLOR_TEXT_MUTED)
        self._desenhar_bloco_vertical(cx - 240, 60, letra_rua, 1, ocupados)
        self._desenhar_bloco_vertical(cx - 120, 60, letra_rua, 3, ocupados)

        # Lado Direito (Blocos Pares: 02, 28)
        self.canvas_mapa.create_text(cx + 140, 40, text="LADO DIREITO (PAR)", font=("Segoe UI", 9, "bold"), fill=COLOR_TEXT_MUTED)
        self._desenhar_bloco_vertical(cx + 60, 60, letra_rua, 2, ocupados)
        self._desenhar_bloco_vertical(cx + 180, 60, letra_rua, 28, ocupados)

    def _desenhar_bloco_vertical(self, start_x, start_y, letra_rua, bloco_num, ocupados):
        """Renderiza um bloco com níveis de 1 a 14 dispostos em pares."""
        box_w = 34
        box_h = 24
        
        self.canvas_mapa.create_text(start_x + box_w, start_y, text=f"Bloco {bloco_num:02d}", font=("Segoe UI", 8, "bold"), fill=COLOR_ACCENT)

        # São 7 andares com 2 posições por andar (Total 14 posições)
        # Nível 7 (13, 14) no topo -> Nível 1 (01, 02) na base
        for nivel in range(7):
            pos_esq = 13 - (nivel * 2)
            pos_dir = 14 - (nivel * 2)
            y = start_y + 14 + (nivel * (box_h + 3))

            for idx, pos in enumerate([pos_esq, pos_dir]):
                x = start_x + (idx * (box_w + 3))
                cod_loc = f"{letra_rua}{bloco_num:02d}{pos:02d}"
                info = ocupados.get(cod_loc)

                cor = COLOR_SLOT_EMPTY
                if info:
                    cor = COLOR_ALERT if info["qtd"] < 20 else COLOR_SLOT_OCCUPIED

                rect = self.canvas_mapa.create_rectangle(x, y, x + box_w, y + box_h, fill=cor, outline=COLOR_BORDER)
                self.canvas_mapa.create_text(x + box_w/2, y + box_h/2, text=f"{pos:02d}", font=("Segoe UI", 7, "bold"), fill=COLOR_TEXT_MAIN)

    def _carregar_tabela(self):
        """Filtro inteligente de busca por Código, Lote, Nome ou Quantidade."""
        termo = self.entry_busca.get().strip()
        for row in self.tree.get_children():
            self.tree.delete(row)

        with self.db.get_connection() as conn:
            cursor = conn.cursor()
            if termo:
                query = """
                    SELECT id, codigo_localizacao, codigo_produto, nome_medicamento, lote, quantidade 
                    FROM itens_estoque 
                    WHERE codigo_produto LIKE ? OR lote LIKE ? OR nome_medicamento LIKE ? OR CAST(quantidade AS TEXT) = ?
                    ORDER BY id DESC
                """
                param = f"%{termo}%"
                cursor.execute(query, (param, param, param, termo))
            else:
                cursor.execute("SELECT id, codigo_localizacao, codigo_produto, nome_medicamento, lote, quantidade FROM itens_estoque ORDER BY id DESC")

            for item in cursor.fetchall():
                self.tree.insert("", tk.END, values=item)

    def _on_tabela_select(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        vals = self.tree.item(selected[0], "values")
        self.selected_item_id = vals[0]
        self.form_entries["loc"].delete(0, tk.END)
        self.form_entries["loc"].insert(0, vals[1])
        self.form_entries["codigo"].delete(0, tk.END)
        self.form_entries["codigo"].insert(0, vals[2])
        self.form_entries["nome"].delete(0, tk.END)
        self.form_entries["nome"].insert(0, vals[3])
        self.form_entries["lote"].delete(0, tk.END)
        self.form_entries["lote"].insert(0, vals[4])
        self.form_entries["qtd"].delete(0, tk.END)
        self.form_entries["qtd"].insert(0, vals[5])

    def _salvar_item(self):
        loc = self.form_entries["loc"].get().strip()
        cod = self.form_entries["codigo"].get().strip()
        nome = self.form_entries["nome"].get().strip()
        lote = self.form_entries["lote"].get().strip()
        qtd_str = self.form_entries["qtd"].get().strip()

        if not (loc and cod and nome and lote and qtd_str):
            messagebox.showwarning("Atenção", "Preencha todos os campos para salvar.")
            return

        try:
            qtd = int(qtd_str)
            if qtd < 0:
                raise ValueError()
        except ValueError:
            messagebox.showerror("Erro", "Quantidade deve ser um número inteiro positivo.")
            return

        parsed = parse_localizacao(loc)
        if not parsed["valido"]:
            messagebox.showerror("Erro de Localização", 
                                 "Formato inválido!\nExemplos válidos:\n- C2809 (Rua C, Bloco 28, Posição 09)\n- FRAC-01 (Área de Frações)")
            return

        with self.db.get_connection() as conn:
            cursor = conn.cursor()
            if self.selected_item_id:
                cursor.execute("""
                    UPDATE itens_estoque 
                    SET rua = ?, bloco = ?, posicao = ?, codigo_localizacao = ?, codigo_produto = ?, nome_medicamento = ?, lote = ?, quantidade = ?
                    WHERE id = ?
                """, (parsed["rua"], parsed.get("bloco", 0), parsed.get("posicao", 1), parsed["codigo"], cod, nome, lote, qtd, self.selected_item_id))
            else:
                cursor.execute("""
                    INSERT INTO itens_estoque (rua, bloco, posicao, codigo_localizacao, codigo_produto, nome_medicamento, lote, quantidade)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (parsed["rua"], parsed.get("bloco", 0), parsed.get("posicao", 1), parsed["codigo"], cod, nome, lote, qtd))
            conn.commit()

        messagebox.showinfo("Sucesso", "Medicamento registrado com sucesso no estoque!")
        self._limpar_formulario()
        self._carregar_tabela()
        self._desenhar_mapa_2d()

    def _remover_item(self):
        if not self.selected_item_id:
            messagebox.showwarning("Atenção", "Selecione um item na tabela para remover.")
            return

        if messagebox.askyesno("Confirmar Baixa", "Deseja realmente dar baixa e remover este item do estoque?"):
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM itens_estoque WHERE id = ?", (self.selected_item_id,))
                conn.commit()

            messagebox.showinfo("Removido", "Item removido do estoque.")
            self._limpar_formulario()
            self._carregar_tabela()
            self._desenhar_mapa_2d()

    def _limpar_formulario(self):
        self.selected_item_id = None
        for ent in self.form_entries.values():
            ent.delete(0, tk.END)
        if self.tree.selection():
            self.tree.selection_remove(self.tree.selection()[0])


if __name__ == "__main__":
    app = EstoqueApp()
    app.mainloop()
