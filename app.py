"""Prati WMS - gestão local de estoque farmacêutico em Tkinter e SQLite."""

import contextlib
import csv
import os
import re
import sqlite3
import sys
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

PRIMARY = "#3F1E69"
PRIMARY_HOVER = "#573087"
BG = "#100C18"
HEADER = "#171020"
SURFACE = "#1C1527"
SURFACE_ALT = "#251B32"
BORDER = "#392A4B"
TEXT = "#F7F2FC"
MUTED = "#AD9CBE"
SUCCESS = "#42C58A"
DANGER = "#FF6B7A"
WARNING = "#F1B95B"

BASE_DIR = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "estoque_medicamentos.db")


class DatabaseManager:
    """Gerencia o banco local e mantém o histórico das operações."""

    def __init__(self, db_path=DB_PATH):
        self.db_path = db_path
        self.init_database()

    @contextlib.contextmanager
    def get_connection(self):
        connection = sqlite3.connect(self.db_path)
        connection.row_factory = sqlite3.Row
        try:
            yield connection
            connection.commit()
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()

    def init_database(self):
        with self.get_connection() as connection:
            cursor = connection.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS itens_estoque (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    rua TEXT NOT NULL,
                    bloco INTEGER NOT NULL,
                    posicao INTEGER NOT NULL CHECK(posicao BETWEEN 1 AND 14),
                    codigo_localizacao TEXT NOT NULL UNIQUE,
                    codigo_produto TEXT NOT NULL,
                    nome_medicamento TEXT NOT NULL,
                    lote TEXT NOT NULL,
                    quantidade INTEGER NOT NULL CHECK(quantidade >= 0),
                    status TEXT NOT NULL DEFAULT 'Normal' CHECK(status IN ('Normal', 'Bloqueado')),
                    criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    atualizado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS historico_movimentacoes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    data_hora TEXT NOT NULL,
                    tipo TEXT NOT NULL,
                    codigo_localizacao TEXT NOT NULL,
                    codigo_produto TEXT NOT NULL,
                    nome_medicamento TEXT NOT NULL,
                    lote TEXT NOT NULL,
                    quantidade INTEGER NOT NULL,
                    motivo TEXT
                )
            """)
            columns = {row[1] for row in cursor.execute("PRAGMA table_info(itens_estoque)")}
            if "status" not in columns:
                cursor.execute("ALTER TABLE itens_estoque ADD COLUMN status TEXT NOT NULL DEFAULT 'Normal'")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_itens_busca ON itens_estoque (codigo_produto, lote, nome_medicamento)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_itens_loc ON itens_estoque (codigo_localizacao)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_hist_data ON historico_movimentacoes (data_hora)")
            if cursor.execute("SELECT COUNT(*) FROM itens_estoque").fetchone()[0] == 0:
                self._seed_inicial(cursor)

    def _seed_inicial(self, cursor):
        dados = [
            ("Rua C", 28, 9, "C2809", "PRATI-101", "Dipirona 500mg/mL Gotas", "L24089A", 140),
            ("Rua C", 28, 10, "C2810", "PRATI-102", "Paracetamol 750mg Comp.", "L24090B", 85),
            ("Rua C", 1, 1, "C0101", "PRATI-103", "Amoxicilina 500mg Cáps.", "L24101A", 60),
            ("Rua C", 2, 3, "C0203", "PRATI-104", "Ibuprofeno 400mg Drágeas", "L24105C", 12),
            ("Rua A", 2, 1, "A0201", "PRATI-201", "Omeprazol 20mg Cáps.", "L24011D", 220),
            ("Rua A", 4, 5, "A0405", "PRATI-202", "Losartana Potássica 50mg", "L24022E", 18),
            ("Rua B", 1, 2, "B0102", "PRATI-301", "Metformina 850mg", "L24044F", 95),
            ("Frações", 0, 1, "FRAC-01", "PRATI-401", "Clonazepam 2,5mg/mL Gotas", "L24055G", 6),
            ("Frações", 0, 2, "FRAC-02", "PRATI-402", "Sinvastatina 20mg Comp.", "L24056H", 9),
        ]
        for item in dados:
            cursor.execute("""
                INSERT INTO itens_estoque
                    (rua, bloco, posicao, codigo_localizacao, codigo_produto,
                     nome_medicamento, lote, quantidade, status)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'Normal')
            """, item)
            self._registrar_historico(cursor, "ENTRADA", item, "Carga inicial do sistema")

    @staticmethod
    def _registrar_historico(cursor, tipo, item, motivo):
        cursor.execute("""
            INSERT INTO historico_movimentacoes
                (data_hora, tipo, codigo_localizacao, codigo_produto,
                 nome_medicamento, lote, quantidade, motivo)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (datetime.now().strftime("%d/%m/%Y %H:%M:%S"), tipo, item[3], item[4], item[5], item[6], item[7], motivo))

    @staticmethod
    def _validar_campos(codigo, nome, lote, quantidade, status):
        codigo, nome, lote = codigo.strip().upper(), nome.strip(), lote.strip().upper()
        if not codigo or not nome or not lote:
            raise ValueError("Preencha código, medicamento e lote.")
        try:
            quantidade = int(quantidade)
        except (TypeError, ValueError) as error:
            raise ValueError("A quantidade deve ser um número inteiro.") from error
        if quantidade < 0:
            raise ValueError("A quantidade não pode ser negativa.")
        if status not in ("Normal", "Bloqueado"):
            raise ValueError("Selecione um status válido.")
        return codigo, nome, lote, quantidade, status

    def metricas(self):
        with self.get_connection() as connection:
            total = connection.execute("SELECT COUNT(*), COALESCE(SUM(quantidade), 0) FROM itens_estoque").fetchone()
            bloqueados = connection.execute("SELECT COUNT(*) FROM itens_estoque WHERE status = 'Bloqueado'").fetchone()[0]
        return total[0], total[1], bloqueados

    def listar_itens(self, busca="", status="Todos"):
        query = "SELECT id, codigo_localizacao, codigo_produto, nome_medicamento, lote, quantidade, status FROM itens_estoque WHERE 1=1"
        params = []
        if busca.strip():
            termo = f"%{busca.strip()}%"
            query += " AND (codigo_localizacao LIKE ? OR codigo_produto LIKE ? OR nome_medicamento LIKE ? OR lote LIKE ? OR CAST(quantidade AS TEXT) LIKE ?)"
            params.extend([termo] * 5)
        if status in ("Normal", "Bloqueado"):
            query += " AND status = ?"
            params.append(status)
        query += " ORDER BY nome_medicamento COLLATE NOCASE, lote"
        with self.get_connection() as connection:
            return connection.execute(query, params).fetchall()

    def obter_item(self, item_id):
        with self.get_connection() as connection:
            return connection.execute("SELECT * FROM itens_estoque WHERE id = ?", (item_id,)).fetchone()

    def cadastrar_item(self, localizacao, codigo, nome, lote, quantidade, status):
        endereco = parse_localizacao(localizacao)
        if not endereco["valido"]:
            raise ValueError(endereco["msg"])
        dados = self._validar_campos(codigo, nome, lote, quantidade, status)
        item = (endereco["rua"], endereco["bloco"], endereco["posicao"], endereco["codigo"], *dados[:4])
        with self.get_connection() as connection:
            cursor = connection.cursor()
            cursor.execute("""
                INSERT INTO itens_estoque
                    (rua, bloco, posicao, codigo_localizacao, codigo_produto,
                     nome_medicamento, lote, quantidade, status)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (*item, dados[4]))
            self._registrar_historico(cursor, "ENTRADA", item, "Cadastro de novo item")

    def atualizar_item(self, item_id, localizacao, codigo, nome, lote, quantidade, status, motivo):
        endereco = parse_localizacao(localizacao)
        if not endereco["valido"]:
            raise ValueError(endereco["msg"])
        dados = self._validar_campos(codigo, nome, lote, quantidade, status)
        item = (endereco["rua"], endereco["bloco"], endereco["posicao"], endereco["codigo"], *dados[:4])
        with self.get_connection() as connection:
            cursor = connection.cursor()
            cursor.execute("""
                UPDATE itens_estoque
                SET rua=?, bloco=?, posicao=?, codigo_localizacao=?, codigo_produto=?,
                    nome_medicamento=?, lote=?, quantidade=?, status=?, atualizado_em=CURRENT_TIMESTAMP
                WHERE id=?
            """, (*item, dados[4], item_id))
            if cursor.rowcount == 0:
                raise ValueError("Item não encontrado.")
            self._registrar_historico(cursor, "ALTERAÇÃO", item, motivo.strip() or "Dados atualizados pela aba Consultar")

    def remover_item(self, item_id):
        with self.get_connection() as connection:
            cursor = connection.cursor()
            row = cursor.execute("SELECT * FROM itens_estoque WHERE id = ?", (item_id,)).fetchone()
            if row is None:
                raise ValueError("Item não encontrado.")
            item = (row["rua"], row["bloco"], row["posicao"], row["codigo_localizacao"], row["codigo_produto"], row["nome_medicamento"], row["lote"], row["quantidade"])
            cursor.execute("DELETE FROM itens_estoque WHERE id = ?", (item_id,))
            self._registrar_historico(cursor, "BAIXA / REMOÇÃO", item, "Item removido pela aba Consultar")

    def listar_historico(self, busca="", tipo="Todos", limite=None):
        query = "SELECT id, data_hora, tipo, codigo_localizacao, codigo_produto, nome_medicamento, lote, quantidade, COALESCE(motivo, '') AS motivo FROM historico_movimentacoes WHERE 1=1"
        params = []
        if busca.strip():
            termo = f"%{busca.strip()}%"
            query += " AND (codigo_localizacao LIKE ? OR codigo_produto LIKE ? OR nome_medicamento LIKE ? OR lote LIKE ? OR motivo LIKE ?)"
            params.extend([termo] * 5)
        if tipo != "Todos":
            query += " AND tipo = ?"
            params.append(tipo)
        query += " ORDER BY id DESC"
        if limite:
            query += " LIMIT ?"
            params.append(limite)
        with self.get_connection() as connection:
            return connection.execute(query, params).fetchall()


def parse_localizacao(texto):
    """Valida C2809 (ruas A-D) e FRAC-01 (frações 01-24)."""
    texto = texto.strip().upper()
    if texto.startswith("FRAC") or "FRA" in texto:
        match = re.search(r"(\d{1,2})", texto)
        numero = max(1, min(24, int(match.group(1)) if match else 1))
        return {"valido": True, "rua": "Frações", "bloco": 0, "posicao": numero, "codigo": f"FRAC-{numero:02d}", "descricao": f"Palete de Frações - Posição {numero:02d}"}
    match = re.match(r"^([A-D])(\d{1,2})(\d{2})$", texto)
    if match:
        letra, bloco, posicao = match.group(1), int(match.group(2)), int(match.group(3))
        if 1 <= posicao <= 14:
            lado = "Direito (Par)" if bloco % 2 == 0 else "Esquerdo (Ímpar)"
            nivel = (posicao + 1) // 2
            return {"valido": True, "rua": f"Rua {letra}", "bloco": bloco, "posicao": posicao, "codigo": f"{letra}{bloco:02d}{posicao:02d}", "lado": lado, "nivel": nivel, "descricao": f"Rua {letra} | Bloco {bloco:02d} - {lado} | Nível {nivel} | Posição {posicao:02d}"}
        return {"valido": False, "msg": f"Posição {posicao} inválida! Todos os blocos possuem exatamente 14 posições (01 a 14)."}
    return {"valido": False, "msg": "Formato inválido! Use C2809 ou FRAC-01."}


def criar_botao(parent, texto, comando, variante="primary"):
    fundo, frente = {
        "primary": (PRIMARY, TEXT),
        "secondary": (SURFACE_ALT, TEXT),
        "danger": (DANGER, "#21060A"),
    }[variante]
    return tk.Button(
        parent, text=texto, command=comando, bg=fundo, fg=frente,
        activebackground=PRIMARY_HOVER if variante == "primary" else BORDER,
        activeforeground=TEXT, relief="flat", bd=0, cursor="hand2",
        font=("Segoe UI Semibold", 10), padx=18, pady=10, highlightthickness=0,
    )


class PratiWMSApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Prati WMS | Gestão de estoque")
        self.geometry("1280x780")
        self.minsize(1050, 660)
        self.configure(bg=BG)
        self.db = DatabaseManager()
        self.paginas = {}
        self.botoes_nav = {}
        self._configurar_estilos()
        self._construir_interface()
        self.mostrar_pagina("inicio")

    def _configurar_estilos(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Treeview", background=SURFACE, fieldbackground=SURFACE, foreground=TEXT, rowheight=35, borderwidth=0, font=("Segoe UI", 9))
        style.map("Treeview", background=[("selected", PRIMARY)])
        style.configure("Treeview.Heading", background=SURFACE_ALT, foreground=MUTED, borderwidth=0, relief="flat", padding=(8, 10), font=("Segoe UI Semibold", 9))
        style.map("Treeview.Heading", background=[("active", SURFACE_ALT)])
        style.configure("TCombobox", fieldbackground=SURFACE_ALT, background=SURFACE_ALT, foreground=TEXT, arrowcolor=TEXT, bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER, padding=7)
        style.map("TCombobox", fieldbackground=[("readonly", SURFACE_ALT)], foreground=[("readonly", TEXT)])
        style.configure("Vertical.TScrollbar", background=SURFACE_ALT, troughcolor=BG, bordercolor=BG, arrowcolor=MUTED)

    def _construir_interface(self):
        header = tk.Frame(self, bg=HEADER, height=76, highlightbackground=BORDER, highlightthickness=1)
        header.pack(fill="x")
        header.pack_propagate(False)
        marca = tk.Frame(header, bg=HEADER)
        marca.pack(side="left", padx=(28, 50), fill="y")
        tk.Label(marca, text="PRATI", bg=HEADER, fg=TEXT, font=("Segoe UI Black", 18)).pack(side="left")
        tk.Label(marca, text="WMS", bg=PRIMARY, fg=TEXT, font=("Segoe UI Bold", 9), padx=7, pady=3).pack(side="left", padx=(8, 0))
        navegacao = tk.Frame(header, bg=HEADER)
        navegacao.pack(side="left", fill="y")
        for chave, texto in (("inicio", "Início"), ("consultar", "Consultar"), ("cadastrar", "Cadastrar"), ("historico", "Histórico")):
            botao = tk.Button(
                navegacao, text=texto, command=lambda pagina=chave: self.mostrar_pagina(pagina),
                bg=HEADER, fg=MUTED, activebackground=HEADER, activeforeground=TEXT,
                relief="flat", bd=0, cursor="hand2", font=("Segoe UI Semibold", 10),
                padx=22, pady=25,
            )
            botao.pack(side="left", fill="y")
            self.botoes_nav[chave] = botao
        tk.Label(header, text="SQLite local  •  Dados neste dispositivo", bg=HEADER, fg=MUTED, font=("Segoe UI", 9)).pack(side="right", padx=28)

        self.container = tk.Frame(self, bg=BG)
        self.container.pack(fill="both", expand=True)
        for chave, classe in (("inicio", InicioView), ("consultar", ConsultarView), ("cadastrar", CadastrarView), ("historico", HistoricoView)):
            pagina = classe(self.container, self)
            pagina.place(relx=0, rely=0, relwidth=1, relheight=1)
            self.paginas[chave] = pagina

    def mostrar_pagina(self, chave):
        for nome, botao in self.botoes_nav.items():
            ativo = nome == chave
            botao.configure(fg=TEXT if ativo else MUTED, bg=PRIMARY if ativo else HEADER, activebackground=PRIMARY if ativo else HEADER)
        pagina = self.paginas[chave]
        pagina.tkraise()
        pagina.on_show()

    def atualizar_todas(self):
        for pagina in self.paginas.values():
            pagina.on_show()


class BaseView(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=BG)
        self.app = app

    def cabecalho(self, titulo, descricao):
        bloco = tk.Frame(self, bg=BG)
        bloco.pack(fill="x", padx=34, pady=(28, 20))
        tk.Label(bloco, text=titulo, bg=BG, fg=TEXT, font=("Segoe UI Semibold", 22)).pack(anchor="w")
        tk.Label(bloco, text=descricao, bg=BG, fg=MUTED, font=("Segoe UI", 10)).pack(anchor="w", pady=(5, 0))

    def on_show(self):
        pass


class InicioView(BaseView):
    def __init__(self, parent, app):
        super().__init__(parent, app)
        self.metric_labels = {}
        self.cabecalho("Visão geral", "Acompanhe o estoque e as últimas movimentações.")
        metricas = tk.Frame(self, bg=BG)
        metricas.pack(fill="x", padx=34)
        for indice, (chave, titulo, cor) in enumerate((
            ("itens", "Itens cadastrados", TEXT),
            ("unidades", "Unidades em estoque", SUCCESS),
            ("bloqueados", "Itens bloqueados", WARNING),
        )):
            card = tk.Frame(metricas, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
            card.grid(row=0, column=indice, sticky="nsew", padx=(0, 12 if indice < 2 else 0))
            metricas.grid_columnconfigure(indice, weight=1)
            tk.Label(card, text=titulo, bg=SURFACE, fg=MUTED, font=("Segoe UI", 10)).pack(anchor="w", padx=20, pady=(18, 6))
            label = tk.Label(card, text="—", bg=SURFACE, fg=cor, font=("Segoe UI Semibold", 28))
            label.pack(anchor="w", padx=20, pady=(0, 18))
            self.metric_labels[chave] = label

        painel = tk.Frame(self, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        painel.pack(fill="both", expand=True, padx=34, pady=22)
        topo = tk.Frame(painel, bg=SURFACE)
        topo.pack(fill="x", padx=20, pady=(18, 12))
        tk.Label(topo, text="Movimentações recentes", bg=SURFACE, fg=TEXT, font=("Segoe UI Semibold", 13)).pack(side="left")
        criar_botao(topo, "Ver histórico", lambda: self.app.mostrar_pagina("historico"), "secondary").pack(side="right")
        colunas = ("data", "tipo", "local", "produto", "medicamento", "qtd")
        self.tabela = ttk.Treeview(painel, columns=colunas, show="headings", height=9)
        titulos = ("DATA / HORA", "MOVIMENTO", "LOCAL", "CÓDIGO", "MEDICAMENTO", "QTD.")
        larguras = (145, 125, 90, 110, 390, 70)
        for coluna, titulo, largura in zip(colunas, titulos, larguras):
            self.tabela.heading(coluna, text=titulo)
            self.tabela.column(coluna, width=largura, anchor="center" if coluna != "medicamento" else "w")
        self.tabela.pack(fill="both", expand=True, padx=20, pady=(0, 20))

    def on_show(self):
        itens, unidades, bloqueados = self.app.db.metricas()
        self.metric_labels["itens"].configure(text=f"{itens:,}".replace(",", "."))
        self.metric_labels["unidades"].configure(text=f"{unidades:,}".replace(",", "."))
        self.metric_labels["bloqueados"].configure(text=str(bloqueados))
        self.tabela.delete(*self.tabela.get_children())
        for row in self.app.db.listar_historico(limite=8):
            self.tabela.insert("", "end", values=(row["data_hora"], row["tipo"], row["codigo_localizacao"], row["codigo_produto"], row["nome_medicamento"], row["quantidade"]))


class ConsultarView(BaseView):
    def __init__(self, parent, app):
        super().__init__(parent, app)
        self.busca = tk.StringVar()
        self.filtro_status = tk.StringVar(value="Todos")
        self.contagem = tk.StringVar(value="0 itens")
        self.cabecalho("Consultar estoque", "Busque em tempo real e edite ou remova diretamente pela linha.")
        painel = tk.Frame(self, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        painel.pack(fill="both", expand=True, padx=34, pady=(0, 28))
        filtros = tk.Frame(painel, bg=SURFACE)
        filtros.pack(fill="x", padx=20, pady=18)
        tk.Entry(
            filtros, textvariable=self.busca, bg=SURFACE_ALT, fg=TEXT,
            insertbackground=TEXT, relief="flat", font=("Segoe UI", 10), width=48,
            highlightbackground=BORDER, highlightthickness=1,
        ).pack(side="left", ipady=9, padx=(0, 12))
        combo = ttk.Combobox(filtros, textvariable=self.filtro_status, values=("Todos", "Normal", "Bloqueado"), state="readonly", width=15)
        combo.pack(side="left")
        tk.Label(filtros, textvariable=self.contagem, bg=SURFACE, fg=MUTED, font=("Segoe UI", 9)).pack(side="right")
        criar_botao(filtros, "+ Novo item", lambda: self.app.mostrar_pagina("cadastrar")).pack(side="right", padx=(0, 12))
        self.busca.trace_add("write", lambda *_: self.carregar())
        combo.bind("<<ComboboxSelected>>", lambda _event: self.carregar())

        tabela_frame = tk.Frame(painel, bg=SURFACE)
        tabela_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        colunas = ("local", "codigo", "medicamento", "lote", "qtd", "status", "acoes")
        self.tabela = ttk.Treeview(tabela_frame, columns=colunas, show="headings")
        config = (
            ("local", "LOCAL", 95, "center"), ("codigo", "CÓDIGO", 115, "center"),
            ("medicamento", "MEDICAMENTO", 360, "w"), ("lote", "LOTE", 110, "center"),
            ("qtd", "QTD.", 70, "center"), ("status", "STATUS", 100, "center"),
            ("acoes", "AÇÕES", 160, "center"),
        )
        for coluna, titulo, largura, ancora in config:
            self.tabela.heading(coluna, text=titulo)
            self.tabela.column(coluna, width=largura, anchor=ancora)
        barra = ttk.Scrollbar(tabela_frame, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=barra.set)
        self.tabela.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        self.tabela.tag_configure("bloqueado", foreground=WARNING)
        self.tabela.bind("<ButtonRelease-1>", self._acao_linha)
        self.tabela.bind("<Double-1>", lambda _event: self._editar_selecionado())

    def on_show(self):
        self.carregar()

    def carregar(self):
        if not hasattr(self, "tabela"):
            return
        self.tabela.delete(*self.tabela.get_children())
        rows = self.app.db.listar_itens(self.busca.get(), self.filtro_status.get())
        for row in rows:
            tag = "bloqueado" if row["status"] == "Bloqueado" else ""
            self.tabela.insert("", "end", iid=str(row["id"]), values=(
                row["codigo_localizacao"], row["codigo_produto"], row["nome_medicamento"],
                row["lote"], row["quantidade"], row["status"], "Editar     |     Remover",
            ), tags=(tag,))
        self.contagem.set(f"{len(rows)} item" + ("" if len(rows) == 1 else "s"))

    def _acao_linha(self, event):
        if self.tabela.identify_region(event.x, event.y) != "cell":
            return
        item_id = self.tabela.identify_row(event.y)
        if not item_id or self.tabela.identify_column(event.x) != "#7":
            return
        x, _y, largura, _altura = self.tabela.bbox(item_id, "acoes")
        self._editar(item_id) if event.x < x + largura / 2 else self._remover(item_id)

    def _editar_selecionado(self):
        selecionado = self.tabela.selection()
        if selecionado:
            self._editar(selecionado[0])

    def _editar(self, item_id):
        item = self.app.db.obter_item(int(item_id))
        if item:
            ItemDialog(self, self.app, item)

    def _remover(self, item_id):
        valores = self.tabela.item(item_id, "values")
        if not messagebox.askyesno(
            "Remover item", f"Remover {valores[2]} do local {valores[0]}?\n\nA operação ficará registrada no histórico.", parent=self,
        ):
            return
        try:
            self.app.db.remover_item(int(item_id))
            self.app.atualizar_todas()
        except (ValueError, sqlite3.Error) as error:
            messagebox.showerror("Não foi possível remover", str(error), parent=self)


class CadastrarView(BaseView):
    def __init__(self, parent, app):
        super().__init__(parent, app)
        self.campos = {}
        self.status = tk.StringVar(value="Normal")
        self.previa = tk.StringVar(value="Informe um endereço, como C2809 ou FRAC-01.")
        self.cabecalho("Cadastrar item", "Inclua um novo medicamento em uma posição livre do estoque.")
        card = tk.Frame(self, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill="both", expand=True, padx=34, pady=(0, 28))
        form = tk.Frame(card, bg=SURFACE)
        form.pack(fill="x", padx=34, pady=28)
        campos = (
            ("localizacao", "Localização"), ("codigo", "Código do produto"),
            ("nome", "Medicamento"), ("lote", "Lote"), ("quantidade", "Quantidade"),
        )
        for indice, (chave, rotulo) in enumerate(campos):
            linha, coluna = divmod(indice, 2)
            bloco = tk.Frame(form, bg=SURFACE)
            bloco.grid(row=linha, column=coluna, sticky="ew", padx=(0, 22 if coluna == 0 else 0), pady=(0, 18))
            form.grid_columnconfigure(coluna, weight=1)
            tk.Label(bloco, text=rotulo, bg=SURFACE, fg=MUTED, font=("Segoe UI Semibold", 9)).pack(anchor="w", pady=(0, 7))
            entrada = tk.Entry(bloco, bg=SURFACE_ALT, fg=TEXT, insertbackground=TEXT, relief="flat", font=("Segoe UI", 11), highlightbackground=BORDER, highlightthickness=1)
            entrada.pack(fill="x", ipady=10)
            self.campos[chave] = entrada
            if chave == "localizacao":
                entrada.bind("<KeyRelease>", self._atualizar_previa)

        bloco_status = tk.Frame(form, bg=SURFACE)
        bloco_status.grid(row=2, column=1, sticky="ew", pady=(0, 18))
        tk.Label(bloco_status, text="Status", bg=SURFACE, fg=MUTED, font=("Segoe UI Semibold", 9)).pack(anchor="w", pady=(0, 7))
        ttk.Combobox(bloco_status, textvariable=self.status, values=("Normal", "Bloqueado"), state="readonly").pack(fill="x")
        previa = tk.Frame(card, bg=SURFACE_ALT, highlightbackground=BORDER, highlightthickness=1)
        previa.pack(fill="x", padx=34, pady=(0, 24))
        tk.Label(previa, text="VALIDAÇÃO DO ENDEREÇO", bg=SURFACE_ALT, fg=MUTED, font=("Segoe UI Semibold", 8)).pack(anchor="w", padx=16, pady=(12, 4))
        self.previa_label = tk.Label(previa, textvariable=self.previa, bg=SURFACE_ALT, fg=MUTED, font=("Segoe UI", 10))
        self.previa_label.pack(anchor="w", padx=16, pady=(0, 12))
        acoes = tk.Frame(card, bg=SURFACE)
        acoes.pack(fill="x", padx=34, pady=(0, 28))
        criar_botao(acoes, "Limpar", self.limpar, "secondary").pack(side="right")
        criar_botao(acoes, "Cadastrar item", self.salvar).pack(side="right", padx=(0, 10))

    def _atualizar_previa(self, _event=None):
        resultado = parse_localizacao(self.campos["localizacao"].get())
        self.previa.set(resultado.get("descricao", resultado.get("msg")))
        self.previa_label.configure(fg=SUCCESS if resultado["valido"] else DANGER)

    def salvar(self):
        try:
            self.app.db.cadastrar_item(
                self.campos["localizacao"].get(), self.campos["codigo"].get(),
                self.campos["nome"].get(), self.campos["lote"].get(),
                self.campos["quantidade"].get(), self.status.get(),
            )
        except sqlite3.IntegrityError as error:
            mensagem = "A localização informada já está ocupada." if "UNIQUE" in str(error) else str(error)
            messagebox.showerror("Cadastro não realizado", mensagem, parent=self)
            return
        except (ValueError, sqlite3.Error) as error:
            messagebox.showerror("Cadastro não realizado", str(error), parent=self)
            return
        messagebox.showinfo("Item cadastrado", "O item foi incluído com sucesso.", parent=self)
        self.limpar()
        self.app.mostrar_pagina("consultar")

    def limpar(self):
        for entrada in self.campos.values():
            entrada.delete(0, "end")
        self.status.set("Normal")
        self.previa.set("Informe um endereço, como C2809 ou FRAC-01.")
        self.previa_label.configure(fg=MUTED)
        self.campos["localizacao"].focus_set()


class HistoricoView(BaseView):
    def __init__(self, parent, app):
        super().__init__(parent, app)
        self.busca = tk.StringVar()
        self.tipo = tk.StringVar(value="Todos")
        self.cabecalho("Histórico", "Rastreabilidade completa de entradas, alterações e remoções.")
        painel = tk.Frame(self, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        painel.pack(fill="both", expand=True, padx=34, pady=(0, 28))
        filtros = tk.Frame(painel, bg=SURFACE)
        filtros.pack(fill="x", padx=20, pady=18)
        tk.Entry(
            filtros, textvariable=self.busca, bg=SURFACE_ALT, fg=TEXT,
            insertbackground=TEXT, relief="flat", font=("Segoe UI", 10), width=42,
            highlightbackground=BORDER, highlightthickness=1,
        ).pack(side="left", ipady=9, padx=(0, 12))
        combo = ttk.Combobox(
            filtros, textvariable=self.tipo,
            values=("Todos", "ENTRADA", "ALTERAÇÃO", "BAIXA / REMOÇÃO"),
            state="readonly", width=20,
        )
        combo.pack(side="left")
        criar_botao(filtros, "Exportar CSV", self.exportar, "secondary").pack(side="right")
        self.busca.trace_add("write", lambda *_: self.carregar())
        combo.bind("<<ComboboxSelected>>", lambda _event: self.carregar())

        tabela_frame = tk.Frame(painel, bg=SURFACE)
        tabela_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        colunas = ("data", "tipo", "local", "codigo", "medicamento", "lote", "qtd", "motivo")
        self.tabela = ttk.Treeview(tabela_frame, columns=colunas, show="headings")
        config = (
            ("data", "DATA / HORA", 140), ("tipo", "MOVIMENTO", 125),
            ("local", "LOCAL", 80), ("codigo", "CÓDIGO", 105),
            ("medicamento", "MEDICAMENTO", 260), ("lote", "LOTE", 95),
            ("qtd", "QTD.", 60), ("motivo", "DETALHES", 260),
        )
        for coluna, titulo, largura in config:
            self.tabela.heading(coluna, text=titulo)
            self.tabela.column(coluna, width=largura, anchor="w" if coluna in ("medicamento", "motivo") else "center")
        barra = ttk.Scrollbar(tabela_frame, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=barra.set)
        self.tabela.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")

    def on_show(self):
        self.carregar()

    def carregar(self):
        if not hasattr(self, "tabela"):
            return
        self.tabela.delete(*self.tabela.get_children())
        for row in self.app.db.listar_historico(self.busca.get(), self.tipo.get()):
            self.tabela.insert("", "end", values=(row["data_hora"], row["tipo"], row["codigo_localizacao"], row["codigo_produto"], row["nome_medicamento"], row["lote"], row["quantidade"], row["motivo"]))

    def exportar(self):
        caminho = filedialog.asksaveasfilename(
            parent=self, title="Exportar histórico", defaultextension=".csv",
            filetypes=(("Arquivo CSV", "*.csv"),),
            initialfile=f"historico_prati_{datetime.now():%Y%m%d}.csv",
        )
        if not caminho:
            return
        rows = self.app.db.listar_historico(self.busca.get(), self.tipo.get())
        with open(caminho, "w", newline="", encoding="utf-8-sig") as arquivo:
            writer = csv.writer(arquivo, delimiter=";")
            writer.writerow(("Data/Hora", "Movimento", "Localização", "Código", "Medicamento", "Lote", "Quantidade", "Detalhes"))
            for row in rows:
                writer.writerow((row["data_hora"], row["tipo"], row["codigo_localizacao"], row["codigo_produto"], row["nome_medicamento"], row["lote"], row["quantidade"], row["motivo"]))
        messagebox.showinfo("Exportação concluída", "O histórico foi exportado com sucesso.", parent=self)


class ItemDialog(tk.Toplevel):
    def __init__(self, parent, app, item):
        super().__init__(parent)
        self.app = app
        self.item = item
        self.title("Editar item")
        self.geometry("570x570")
        self.resizable(False, False)
        self.configure(bg=BG)
        self.transient(parent.winfo_toplevel())
        self.grab_set()
        self.campos = {}
        self.status = tk.StringVar(value=item["status"])
        self._construir()

    def _construir(self):
        tk.Label(self, text="Editar item", bg=BG, fg=TEXT, font=("Segoe UI Semibold", 20)).pack(anchor="w", padx=28, pady=(24, 4))
        tk.Label(self, text="Atualize os dados e registre o motivo da alteração.", bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w", padx=28, pady=(0, 18))
        form = tk.Frame(self, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        form.pack(fill="both", expand=True, padx=28, pady=(0, 20))
        valores = {
            "localizacao": self.item["codigo_localizacao"], "codigo": self.item["codigo_produto"],
            "nome": self.item["nome_medicamento"], "lote": self.item["lote"],
            "quantidade": str(self.item["quantidade"]), "motivo": "",
        }
        rotulos = {
            "localizacao": "Localização", "codigo": "Código do produto", "nome": "Medicamento",
            "lote": "Lote", "quantidade": "Quantidade", "motivo": "Motivo / detalhes da alteração",
        }
        for indice, chave in enumerate(valores):
            linha, coluna = divmod(indice, 2)
            bloco = tk.Frame(form, bg=SURFACE)
            bloco.grid(row=linha, column=coluna, sticky="ew", padx=(18, 9 if coluna == 0 else 18), pady=(16, 0))
            form.grid_columnconfigure(coluna, weight=1)
            tk.Label(bloco, text=rotulos[chave], bg=SURFACE, fg=MUTED, font=("Segoe UI Semibold", 8)).pack(anchor="w", pady=(0, 5))
            entrada = tk.Entry(bloco, bg=SURFACE_ALT, fg=TEXT, insertbackground=TEXT, relief="flat", font=("Segoe UI", 10), highlightbackground=BORDER, highlightthickness=1)
            entrada.pack(fill="x", ipady=8)
            entrada.insert(0, valores[chave])
            self.campos[chave] = entrada
        status_bloco = tk.Frame(form, bg=SURFACE)
        status_bloco.grid(row=3, column=0, sticky="ew", padx=(18, 9), pady=(16, 18))
        tk.Label(status_bloco, text="Status", bg=SURFACE, fg=MUTED, font=("Segoe UI Semibold", 8)).pack(anchor="w", pady=(0, 5))
        ttk.Combobox(status_bloco, textvariable=self.status, values=("Normal", "Bloqueado"), state="readonly").pack(fill="x")
        acoes = tk.Frame(self, bg=BG)
        acoes.pack(fill="x", padx=28, pady=(0, 24))
        criar_botao(acoes, "Cancelar", self.destroy, "secondary").pack(side="right")
        criar_botao(acoes, "Salvar alterações", self.salvar).pack(side="right", padx=(0, 10))

    def salvar(self):
        try:
            self.app.db.atualizar_item(
                self.item["id"], self.campos["localizacao"].get(), self.campos["codigo"].get(),
                self.campos["nome"].get(), self.campos["lote"].get(),
                self.campos["quantidade"].get(), self.status.get(), self.campos["motivo"].get(),
            )
        except sqlite3.IntegrityError as error:
            mensagem = "A localização informada já está ocupada." if "UNIQUE" in str(error) else str(error)
            messagebox.showerror("Alteração não realizada", mensagem, parent=self)
            return
        except (ValueError, sqlite3.Error) as error:
            messagebox.showerror("Alteração não realizada", str(error), parent=self)
            return
        self.destroy()
        self.app.atualizar_todas()


if __name__ == "__main__":
    app = PratiWMSApp()
    app.mainloop()
