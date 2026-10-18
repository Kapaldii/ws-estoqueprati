"""Prati WMS - gestão local de estoque farmacêutico em Tkinter e SQLite."""

import contextlib
import csv
import os
import re
import sqlite3
import sys
import tkinter as tk
import unicodedata
import zipfile
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
from xml.etree import ElementTree

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

RUAS = ("Rua A", "Rua B", "Rua C", "Rua D", "Frações")
BLOCOS_POR_RUA = range(1, 35)

BASE_DIR = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "estoque_medicamentos.db")


def normalizar_texto(valor):
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    return "".join(caractere for caractere in texto if not unicodedata.combining(caractere)).strip().lower()


def eh_fracao(valor):
    return normalizar_texto(valor).startswith("frac")


def formatar_localizacao(codigo):
    return "Fração" if eh_fracao(codigo) else codigo


def _indice_coluna_excel(referencia):
    letras = re.match(r"[A-Z]+", referencia or "A").group(0)
    indice = 0
    for letra in letras:
        indice = indice * 26 + ord(letra) - 64
    return indice - 1


def ler_planilha_excel(caminho):
    """Lê a primeira planilha de um arquivo XLSX usando apenas a biblioteca padrão."""
    namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with zipfile.ZipFile(caminho) as arquivo:
        compartilhados = []
        if "xl/sharedStrings.xml" in arquivo.namelist():
            raiz = ElementTree.fromstring(arquivo.read("xl/sharedStrings.xml"))
            compartilhados = ["".join(no.text or "" for no in item.iter(f"{namespace}t")) for item in raiz]
        planilhas = sorted(nome for nome in arquivo.namelist() if nome.startswith("xl/worksheets/sheet") and nome.endswith(".xml"))
        if not planilhas:
            raise ValueError("O arquivo Excel não contém planilhas.")
        raiz = ElementTree.fromstring(arquivo.read(planilhas[0]))
        linhas = []
        for linha in raiz.iter(f"{namespace}row"):
            valores = []
            for celula in linha.findall(f"{namespace}c"):
                indice = _indice_coluna_excel(celula.get("r", "A1"))
                while len(valores) <= indice:
                    valores.append("")
                tipo = celula.get("t")
                if tipo == "inlineStr":
                    valor = "".join(no.text or "" for no in celula.iter(f"{namespace}t"))
                else:
                    no_valor = celula.find(f"{namespace}v")
                    valor = no_valor.text if no_valor is not None else ""
                    if tipo == "s" and valor:
                        valor = compartilhados[int(valor)]
                    elif valor and re.fullmatch(r"-?\d+\.0", valor):
                        valor = valor[:-2]
                valores[indice] = valor
            if any(str(valor).strip() for valor in valores):
                linhas.append(valores)
    if not linhas:
        raise ValueError("A planilha está vazia.")
    return linhas


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

    def listar_localizacoes_livres(self, rua):
        if rua == "Frações":
            return ["Fração"]
        if rua not in RUAS:
            return []
        letra = rua[-1]
        with self.get_connection() as connection:
            ocupadas = {
                row[0] for row in connection.execute(
                    "SELECT codigo_localizacao FROM itens_estoque WHERE rua = ?", (rua,)
                )
            }
        return [
            codigo
            for bloco in BLOCOS_POR_RUA
            for posicao in range(1, 15)
            if (codigo := f"{letra}{bloco:02d}{posicao:02d}") not in ocupadas
        ]

    @staticmethod
    def _proximo_codigo_fracao(cursor):
        usados = set()
        for row in cursor.execute("SELECT codigo_localizacao FROM itens_estoque WHERE rua = 'Frações'"):
            match = re.search(r"(\d+)$", row[0])
            if match:
                usados.add(int(match.group(1)))
        numero = 1
        while numero in usados:
            numero += 1
        return f"FRAC-{numero:02d}", numero

    def cadastrar_item(self, localizacao, codigo, nome, lote, quantidade, status):
        dados = self._validar_campos(codigo, nome, lote, quantidade, status)
        fracao = eh_fracao(localizacao)
        endereco = None if fracao else parse_localizacao(localizacao)
        if endereco is not None and not endereco["valido"]:
            raise ValueError(endereco["msg"])
        with self.get_connection() as connection:
            cursor = connection.cursor()
            if fracao:
                codigo_fracao, posicao = self._proximo_codigo_fracao(cursor)
                endereco = {"rua": "Frações", "bloco": 0, "posicao": posicao, "codigo": codigo_fracao}
            item = (endereco["rua"], endereco["bloco"], endereco["posicao"], endereco["codigo"], *dados[:4])
            cursor.execute("""
                INSERT INTO itens_estoque
                    (rua, bloco, posicao, codigo_localizacao, codigo_produto,
                     nome_medicamento, lote, quantidade, status)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (*item, dados[4]))
            self._registrar_historico(cursor, "ENTRADA", item, "Cadastro de novo item")

    def importar_excel(self, caminho):
        linhas = ler_planilha_excel(caminho)
        cabecalhos = {
            re.sub(r"[^a-z0-9]+", "_", normalizar_texto(valor)).strip("_"): indice
            for indice, valor in enumerate(linhas[0])
        }
        aliases = {
            "rua": ("rua",),
            "localizacao": ("localizacao", "codigo_localizacao", "codigo_da_localizacao", "endereco", "posicao"),
            "codigo": ("codigo", "codigo_produto", "codigo_do_produto", "cod_produto"),
            "nome": ("medicamento", "nome", "nome_medicamento", "nome_do_medicamento", "produto"),
            "lote": ("lote",),
            "quantidade": ("quantidade", "qtd", "estoque"),
            "status": ("status", "situacao"),
        }

        def indice_de(campo):
            return next((cabecalhos[nome] for nome in aliases[campo] if nome in cabecalhos), None)

        indices = {campo: indice_de(campo) for campo in aliases}
        faltantes = [campo for campo in ("codigo", "nome", "lote", "quantidade") if indices[campo] is None]
        if faltantes or (indices["localizacao"] is None and indices["rua"] is None):
            nomes = ", ".join(faltantes + (["localização ou rua"] if indices["localizacao"] is None and indices["rua"] is None else []))
            raise ValueError(f"Colunas obrigatórias ausentes: {nomes}.")

        importados, erros = 0, []
        for numero_linha, linha in enumerate(linhas[1:], start=2):
            def valor(campo, padrao=""):
                indice = indices[campo]
                return str(linha[indice]).strip() if indice is not None and indice < len(linha) else padrao

            if not any(str(celula).strip() for celula in linha):
                continue
            rua, localizacao = valor("rua"), valor("localizacao")
            if eh_fracao(rua) or eh_fracao(localizacao):
                localizacao = "Fração"
            elif localizacao.isdigit() and normalizar_texto(rua) in ("rua a", "rua b", "rua c", "rua d"):
                localizacao = f"{rua.strip()[-1].upper()}{localizacao.zfill(4)}"
            status = "Bloqueado" if normalizar_texto(valor("status", "Normal")) == "bloqueado" else "Normal"
            try:
                self.cadastrar_item(localizacao, valor("codigo"), valor("nome"), valor("lote"), valor("quantidade"), status)
                importados += 1
            except (ValueError, sqlite3.Error) as error:
                erros.append(f"Linha {numero_linha}: {error}")
        return importados, erros

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
        destaque = tk.Frame(self, bg=PRIMARY)
        destaque.pack(fill="x", padx=34, pady=(0, 18))
        chamada = tk.Frame(destaque, bg=PRIMARY)
        chamada.pack(side="left", fill="both", expand=True, padx=24, pady=18)
        tk.Label(chamada, text="O que você deseja fazer?", bg=PRIMARY, fg=TEXT, font=("Segoe UI Semibold", 16)).pack(anchor="w")
        tk.Label(chamada, text="Cadastre um item ou encontre rapidamente o que precisa.", bg=PRIMARY, fg="#DCCDEB", font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 0))
        atalhos = tk.Frame(destaque, bg=PRIMARY)
        atalhos.pack(side="right", padx=20, pady=16)
        criar_botao(atalhos, "+ Novo cadastro", lambda: self.app.mostrar_pagina("cadastrar"), "secondary").pack(side="left", padx=5)
        criar_botao(atalhos, "Importar Excel", self._abrir_importacao, "secondary").pack(side="left", padx=5)
        criar_botao(atalhos, "Consultar estoque", lambda: self.app.mostrar_pagina("consultar"), "secondary").pack(side="left", padx=5)
        metricas = tk.Frame(self, bg=BG)
        metricas.pack(fill="x", padx=34)
        for indice, (chave, titulo, cor) in enumerate((
            ("itens", "Itens cadastrados", TEXT),
            ("unidades", "Unidades em estoque", SUCCESS),
            ("bloqueados", "Itens bloqueados", WARNING),
        )):
            card = tk.Frame(metricas, bg=SURFACE)
            card.grid(row=0, column=indice, sticky="nsew", padx=(0, 12 if indice < 2 else 0))
            metricas.grid_columnconfigure(indice, weight=1)
            tk.Frame(card, bg=cor, height=3).pack(fill="x")
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
            self.tabela.insert("", "end", values=(row["data_hora"], row["tipo"], formatar_localizacao(row["codigo_localizacao"]), row["codigo_produto"], row["nome_medicamento"], row["quantidade"]))

    def _abrir_importacao(self):
        self.app.mostrar_pagina("cadastrar")
        self.app.paginas["cadastrar"].importar_excel()


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
                formatar_localizacao(row["codigo_localizacao"]), row["codigo_produto"], row["nome_medicamento"],
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
        self.rua = tk.StringVar(value="Rua A")
        self.localizacao = tk.StringVar()
        self.status = tk.StringVar(value="Normal")
        self.previa = tk.StringVar(value="Selecione uma rua e uma localização disponível.")
        self.cabecalho("Cadastrar item", "Inclua um novo medicamento em uma posição livre do estoque.")
        card = tk.Frame(self, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill="both", expand=True, padx=34, pady=(0, 28))
        form = tk.Frame(card, bg=SURFACE)
        form.pack(fill="x", padx=34, pady=28)

        bloco_rua = tk.Frame(form, bg=SURFACE)
        bloco_rua.grid(row=0, column=0, sticky="ew", padx=(0, 22), pady=(0, 18))
        form.grid_columnconfigure(0, weight=1)
        tk.Label(bloco_rua, text="Rua", bg=SURFACE, fg=MUTED, font=("Segoe UI Semibold", 9)).pack(anchor="w", pady=(0, 7))
        self.rua_combo = ttk.Combobox(bloco_rua, textvariable=self.rua, values=RUAS, state="readonly")
        self.rua_combo.pack(fill="x")
        self.rua_combo.bind("<<ComboboxSelected>>", self._atualizar_localizacoes)

        bloco_local = tk.Frame(form, bg=SURFACE)
        bloco_local.grid(row=0, column=1, sticky="ew", pady=(0, 18))
        form.grid_columnconfigure(1, weight=1)
        tk.Label(bloco_local, text="Localização livre", bg=SURFACE, fg=MUTED, font=("Segoe UI Semibold", 9)).pack(anchor="w", pady=(0, 7))
        self.localizacao_combo = ttk.Combobox(bloco_local, textvariable=self.localizacao, state="readonly")
        self.localizacao_combo.pack(fill="x")
        self.localizacao_combo.bind("<<ComboboxSelected>>", self._atualizar_previa)

        campos = (
            ("codigo", "Código do produto"), ("nome", "Medicamento"),
            ("lote", "Lote"), ("quantidade", "Quantidade"),
        )
        for indice, (chave, rotulo) in enumerate(campos):
            linha, coluna = divmod(indice, 2)
            linha += 1
            bloco = tk.Frame(form, bg=SURFACE)
            bloco.grid(row=linha, column=coluna, sticky="ew", padx=(0, 22 if coluna == 0 else 0), pady=(0, 18))
            tk.Label(bloco, text=rotulo, bg=SURFACE, fg=MUTED, font=("Segoe UI Semibold", 9)).pack(anchor="w", pady=(0, 7))
            entrada = tk.Entry(bloco, bg=SURFACE_ALT, fg=TEXT, insertbackground=TEXT, relief="flat", font=("Segoe UI", 11), highlightbackground=BORDER, highlightthickness=1)
            entrada.pack(fill="x", ipady=10)
            self.campos[chave] = entrada

        bloco_status = tk.Frame(form, bg=SURFACE)
        bloco_status.grid(row=3, column=0, sticky="ew", padx=(0, 22), pady=(0, 18))
        tk.Label(bloco_status, text="Status", bg=SURFACE, fg=MUTED, font=("Segoe UI Semibold", 9)).pack(anchor="w", pady=(0, 7))
        ttk.Combobox(bloco_status, textvariable=self.status, values=("Normal", "Bloqueado"), state="readonly").pack(fill="x")
        previa = tk.Frame(card, bg=SURFACE_ALT, highlightbackground=BORDER, highlightthickness=1)
        previa.pack(fill="x", padx=34, pady=(0, 24))
        tk.Label(previa, text="VALIDAÇÃO DO ENDEREÇO", bg=SURFACE_ALT, fg=MUTED, font=("Segoe UI Semibold", 8)).pack(anchor="w", padx=16, pady=(12, 4))
        self.previa_label = tk.Label(previa, textvariable=self.previa, bg=SURFACE_ALT, fg=MUTED, font=("Segoe UI", 10))
        self.previa_label.pack(anchor="w", padx=16, pady=(0, 12))
        acoes = tk.Frame(card, bg=SURFACE)
        acoes.pack(fill="x", padx=34, pady=(0, 28))
        criar_botao(acoes, "Importar Excel", self.importar_excel, "secondary").pack(side="left")
        criar_botao(acoes, "Limpar", self.limpar, "secondary").pack(side="right")
        criar_botao(acoes, "Cadastrar item", self.salvar).pack(side="right", padx=(0, 10))
        self._atualizar_localizacoes()

    def _atualizar_previa(self, _event=None):
        localizacao = self.localizacao.get()
        if not localizacao:
            self.previa.set("Não há localizações livres nesta rua.")
            self.previa_label.configure(fg=WARNING)
            return
        if localizacao == "Fração":
            self.previa.set("Palete de Frações • vários itens podem compartilhar este local")
        else:
            self.previa.set(parse_localizacao(localizacao)["descricao"])
        self.previa_label.configure(fg=SUCCESS)

    def _atualizar_localizacoes(self, _event=None):
        livres = self.app.db.listar_localizacoes_livres(self.rua.get())
        self.localizacao_combo.configure(values=livres)
        self.localizacao.set(livres[0] if livres else "")
        self._atualizar_previa()

    def on_show(self):
        if hasattr(self, "localizacao_combo"):
            self._atualizar_localizacoes()

    def salvar(self):
        try:
            self.app.db.cadastrar_item(
                self.localizacao.get(), self.campos["codigo"].get(),
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
        self.rua.set("Rua A")
        self.status.set("Normal")
        self._atualizar_localizacoes()
        self.campos["codigo"].focus_set()

    def importar_excel(self):
        caminho = filedialog.askopenfilename(
            parent=self, title="Importar itens do Excel",
            filetypes=(("Planilha Excel", "*.xlsx"),),
        )
        if not caminho:
            return
        try:
            importados, erros = self.app.db.importar_excel(caminho)
        except (ValueError, OSError, zipfile.BadZipFile) as error:
            messagebox.showerror("Importação não realizada", str(error), parent=self)
            return
        self.app.atualizar_todas()
        if erros:
            detalhes = "\n".join(erros[:8])
            restante = f"\n... e mais {len(erros) - 8} erro(s)." if len(erros) > 8 else ""
            messagebox.showwarning("Importação concluída com ressalvas", f"{importados} item(ns) importado(s).\n{len(erros)} linha(s) ignorada(s).\n\n{detalhes}{restante}", parent=self)
        else:
            messagebox.showinfo("Importação concluída", f"{importados} item(ns) importado(s) com sucesso.", parent=self)
        if importados:
            self.app.mostrar_pagina("consultar")


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
            self.tabela.insert("", "end", values=(row["data_hora"], row["tipo"], formatar_localizacao(row["codigo_localizacao"]), row["codigo_produto"], row["nome_medicamento"], row["lote"], row["quantidade"], row["motivo"]))

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
