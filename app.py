"""Prati WMS - gestão local de estoque farmacêutico em Tkinter e SQLite."""

import contextlib
import csv
import base64
import io
import json
import os
import re
import sqlite3
import sys
import tkinter as tk
import unicodedata
import zipfile
from datetime import datetime, timedelta
from tkinter import filedialog, messagebox, ttk
from xml.etree import ElementTree

import win32con
import win32print
import win32ui
from PIL import Image, ImageDraw, ImageFont, ImageGrab, ImageTk, ImageWin

THEMES = {
    "claro": {
        "PRIMARY": "#3F1E69", "PRIMARY_HOVER": "#573087",
        "BG": "#F7F5FA", "HEADER": "#FFFFFF", "SURFACE": "#FFFFFF",
        "SURFACE_ALT": "#F0EBF5", "BORDER": "#DED5E8",
        "TEXT": "#21182B", "MUTED": "#71647D", "SUCCESS": "#16875D",
        "DANGER": "#D64555", "DANGER_HOVER": "#B93646", "WARNING": "#A66A12", "ON_PRIMARY": "#FFFFFF",
        "PRIMARY_MUTED": "#DFD0EC", "HANDLE": "#7C3AED",
        "BRAND": "#3F1E69", "BRAND_SUB": "#74558F",
    },
    "escuro": {
        "PRIMARY": "#3F1E69", "PRIMARY_HOVER": "#573087",
        "BG": "#100C18", "HEADER": "#171020", "SURFACE": "#1C1527",
        "SURFACE_ALT": "#251B32", "BORDER": "#392A4B",
        "TEXT": "#F7F2FC", "MUTED": "#AD9CBE", "SUCCESS": "#42C58A",
        "DANGER": "#FF6B7A", "DANGER_HOVER": "#E64F61", "WARNING": "#F1B95B", "ON_PRIMARY": "#FFFFFF",
        "PRIMARY_MUTED": "#DCCDEB", "HANDLE": "#A78BFA",
        "BRAND": "#F7F2FC", "BRAND_SUB": "#AD9CBE",
    },
}


def definir_tema(nome):
    paleta = THEMES.get(nome, THEMES["claro"])
    globals().update(paleta)


definir_tema("claro")

RUAS = ("Rua A", "Rua B", "Rua C", "Rua D", "Frações")
BLOCOS_POR_RUA = range(1, 35)

LABEL_WIDTH_IN = 6
LABEL_HEIGHT_IN = 4
LABEL_DPI = 203
LABEL_WIDTH_PX = LABEL_WIDTH_IN * LABEL_DPI
LABEL_HEIGHT_PX = LABEL_HEIGHT_IN * LABEL_DPI
ARIAL_BLACK_PATH = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "ariblk.ttf")

BASE_DIR = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "estoque_medicamentos.db")
CONFIG_PATH = os.path.join(BASE_DIR, "prati_wms_config.json")
HISTORY_RETENTION_DAYS = 2
BACKUP_RETENTION = 7


def carregar_tema():
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as arquivo:
            nome = json.load(arquivo).get("tema")
            return nome if nome in THEMES else "claro"
    except (OSError, ValueError, TypeError, json.JSONDecodeError):
        return "claro"


def salvar_tema(nome):
    temporario = f"{CONFIG_PATH}.tmp"
    with open(temporario, "w", encoding="utf-8") as arquivo:
        json.dump({"tema": nome}, arquivo, ensure_ascii=False, indent=2)
    os.replace(temporario, CONFIG_PATH)


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


def fonte_arial_black(tamanho_pontos):
    tamanho_pixels = max(8, round(float(tamanho_pontos) * LABEL_DPI / 72))
    if not os.path.exists(ARIAL_BLACK_PATH):
        raise RuntimeError("A fonte Arial Black não foi encontrada no Windows.")
    return ImageFont.truetype(ARIAL_BLACK_PATH, tamanho_pixels)


def dimensoes_texto_etiqueta(elemento):
    fonte = fonte_arial_black(elemento["tamanho"])
    desenho = ImageDraw.Draw(Image.new("L", (1, 1)))
    espacamento = max(2, fonte.size // 5)
    caixa = desenho.multiline_textbbox(
        (0, 0), elemento["texto"], font=fonte, spacing=espacamento, align="center",
    )
    largura_natural = max(1, caixa[2] - caixa[0])
    altura_natural = max(1, caixa[3] - caixa[1])
    largura = max(1, round(largura_natural * float(elemento.get("escala_x", 1.0))))
    altura = max(1, round(altura_natural * float(elemento.get("escala_y", 1.0))))
    return largura, altura, caixa, fonte, espacamento


def imagem_texto_etiqueta(elemento):
    largura, altura, caixa, fonte, espacamento = dimensoes_texto_etiqueta(elemento)
    largura_natural = max(1, caixa[2] - caixa[0])
    altura_natural = max(1, caixa[3] - caixa[1])
    imagem = Image.new("RGBA", (largura_natural, altura_natural), (255, 255, 255, 0))
    desenho = ImageDraw.Draw(imagem)
    desenho.multiline_text(
        (-caixa[0], -caixa[1]), elemento["texto"], fill="black", font=fonte,
        spacing=espacamento, align="center",
    )
    if imagem.size != (largura, altura):
        imagem = imagem.resize((largura, altura), Image.Resampling.LANCZOS)
    return imagem


def renderizar_etiqueta(elementos):
    """Renderiza uma etiqueta Zebra SAMPLE horizontal 6x4 em 203 dpi."""
    etiqueta = Image.new("RGB", (LABEL_WIDTH_PX, LABEL_HEIGHT_PX), "white")
    for elemento in elementos:
        if elemento["tipo"] == "texto":
            imagem = imagem_texto_etiqueta(elemento)
            x = int(elemento["x"] - imagem.width / 2)
            y = int(elemento["y"] - imagem.height / 2)
            etiqueta.paste(imagem, (x, y), imagem)
        elif elemento["tipo"] == "imagem":
            imagem = elemento["imagem"].convert("RGBA")
            largura = max(1, int(elemento["largura"]))
            altura = max(1, int(elemento["altura"]))
            imagem = imagem.resize((largura, altura), Image.Resampling.LANCZOS)
            x = int(elemento["x"] - largura / 2)
            y = int(elemento["y"] - altura / 2)
            etiqueta.paste(imagem, (x, y), imagem)
    return etiqueta


def salvar_modelo_etiqueta(caminho, elementos):
    dados = {"formato": "Prati WMS Zebra SAMPLE horizontal", "largura": LABEL_WIDTH_PX, "altura": LABEL_HEIGHT_PX, "elementos": []}
    for elemento in elementos:
        item = {chave: valor for chave, valor in elemento.items() if chave != "imagem"}
        if elemento["tipo"] == "imagem":
            buffer = io.BytesIO()
            elemento["imagem"].save(buffer, format="PNG")
            item["imagem_png"] = base64.b64encode(buffer.getvalue()).decode("ascii")
        dados["elementos"].append(item)
    with open(caminho, "w", encoding="utf-8") as arquivo:
        json.dump(dados, arquivo, ensure_ascii=False, indent=2)


def abrir_modelo_etiqueta(caminho):
    with open(caminho, "r", encoding="utf-8") as arquivo:
        dados = json.load(arquivo)
    if dados.get("largura") != LABEL_WIDTH_PX or dados.get("altura") != LABEL_HEIGHT_PX:
        raise ValueError("O modelo não usa o formato Zebra SAMPLE horizontal 6x4 em 203 dpi.")
    elementos = []
    for item in dados.get("elementos", []):
        tipo = item.get("tipo")
        if tipo not in ("texto", "imagem"):
            raise ValueError("O modelo contém um elemento não suportado.")
        elemento = {chave: valor for chave, valor in item.items() if chave != "imagem_png"}
        if tipo == "imagem":
            conteudo = base64.b64decode(item["imagem_png"])
            elemento["imagem"] = Image.open(io.BytesIO(conteudo)).convert("RGBA")
        else:
            elemento.setdefault("escala_x", 1.0)
            elemento.setdefault("escala_y", 1.0)
        elementos.append(elemento)
    return elementos


def listar_impressoras(modulo=win32print):
    flags = modulo.PRINTER_ENUM_LOCAL | modulo.PRINTER_ENUM_CONNECTIONS
    return sorted({item[2] for item in modulo.EnumPrinters(flags)}, key=str.casefold)


def imprimir_imagem_windows(imagem, impressora, criar_dc=win32ui.CreateDC, criar_dib=ImageWin.Dib):
    """Envia a etiqueta como bitmap para qualquer impressora instalada no Windows."""
    if not impressora:
        raise ValueError("Selecione uma impressora.")
    dc = criar_dc()
    documento_iniciado = False
    try:
        dc.CreatePrinterDC(impressora)
        largura_area = dc.GetDeviceCaps(win32con.HORZRES)
        altura_area = dc.GetDeviceCaps(win32con.VERTRES)
        dpi_x = dc.GetDeviceCaps(win32con.LOGPIXELSX)
        dpi_y = dc.GetDeviceCaps(win32con.LOGPIXELSY)
        largura_alvo = round(LABEL_WIDTH_IN * dpi_x)
        altura_alvo = round(LABEL_HEIGHT_IN * dpi_y)
        escala = min(1.0, largura_area / largura_alvo, altura_area / altura_alvo)
        largura_alvo = max(1, round(largura_alvo * escala))
        altura_alvo = max(1, round(altura_alvo * escala))
        x = max(0, (largura_area - largura_alvo) // 2)
        y = max(0, (altura_area - altura_alvo) // 2)
        dc.StartDoc("Prati WMS - Etiqueta Zebra SAMPLE")
        documento_iniciado = True
        dc.StartPage()
        criar_dib(imagem.convert("RGB")).draw(dc.GetHandleOutput(), (x, y, x + largura_alvo, y + altura_alvo))
        dc.EndPage()
        dc.EndDoc()
        documento_iniciado = False
    except Exception:
        if documento_iniciado:
            dc.AbortDoc()
        raise
    finally:
        dc.DeleteDC()


class DatabaseManager:
    """Gerencia o banco local e mantém o histórico das operações."""

    def __init__(self, db_path=DB_PATH, manutencao_automatica=True):
        self.db_path = os.path.abspath(db_path)
        self.base_dir = os.path.dirname(self.db_path)
        self.backup_dir = os.path.join(self.base_dir, "Backups")
        self.avisos_inicializacao = []
        try:
            os.makedirs(self.base_dir, exist_ok=True)
        except OSError as error:
            raise OSError("Sem permissão para acessar a pasta de dados do Prati WMS.") from error
        self.init_database()
        self.verificar_integridade()
        if manutencao_automatica:
            try:
                self.limpar_historico_antigo()
            except (OSError, sqlite3.Error) as error:
                self.avisos_inicializacao.append(f"Não foi possível limpar o histórico antigo: {error}")
            try:
                self.criar_backup_automatico()
            except (OSError, sqlite3.Error) as error:
                self.avisos_inicializacao.append(f"Não foi possível criar o backup diário: {error}")

    @contextlib.contextmanager
    def get_connection(self):
        connection = sqlite3.connect(self.db_path, timeout=10)
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

    def verificar_integridade(self):
        connection = sqlite3.connect(self.db_path, timeout=10)
        try:
            resultado = connection.execute("PRAGMA quick_check").fetchone()[0]
        finally:
            connection.close()
        if resultado != "ok":
            raise sqlite3.DatabaseError(
                "O banco de dados apresentou um problema de integridade. "
                "Feche o programa e restaure um arquivo da pasta Backups."
            )

    def limpar_historico_antigo(self, agora=None):
        limite = (agora or datetime.now()) - timedelta(days=HISTORY_RETENTION_DAYS)
        ids_antigos = []
        with self.get_connection() as connection:
            for row in connection.execute("SELECT id, data_hora FROM historico_movimentacoes"):
                try:
                    data_registro = datetime.strptime(row["data_hora"], "%d/%m/%Y %H:%M:%S")
                except (TypeError, ValueError):
                    continue
                if data_registro < limite:
                    ids_antigos.append((row["id"],))
            connection.executemany("DELETE FROM historico_movimentacoes WHERE id = ?", ids_antigos)
        return len(ids_antigos)

    def criar_backup(self, caminho):
        caminho = os.path.abspath(caminho)
        if caminho == self.db_path:
            raise ValueError("Escolha um arquivo diferente do banco em uso.")
        os.makedirs(os.path.dirname(caminho), exist_ok=True)
        temporario = f"{caminho}.tmp"
        if os.path.exists(temporario):
            os.remove(temporario)
        origem = sqlite3.connect(self.db_path, timeout=10)
        destino = sqlite3.connect(temporario)
        try:
            origem.backup(destino)
            destino.commit()
            resultado = destino.execute("PRAGMA integrity_check").fetchone()[0]
            if resultado != "ok":
                raise sqlite3.DatabaseError("O backup criado não passou na verificação de integridade.")
        finally:
            destino.close()
            origem.close()
        os.replace(temporario, caminho)
        return caminho

    def criar_backup_automatico(self, agora=None):
        agora = agora or datetime.now()
        os.makedirs(self.backup_dir, exist_ok=True)
        caminho = os.path.join(self.backup_dir, f"Prati_WMS_auto_{agora:%Y-%m-%d}.db")
        if not os.path.exists(caminho):
            self.criar_backup(caminho)
        self._limpar_backups_automaticos()
        return caminho

    def _limpar_backups_automaticos(self):
        arquivos = sorted(
            (
                os.path.join(self.backup_dir, nome)
                for nome in os.listdir(self.backup_dir)
                if nome.startswith("Prati_WMS_auto_") and nome.endswith(".db")
            ),
            key=os.path.basename,
            reverse=True,
        )
        for caminho in arquivos[BACKUP_RETENTION:]:
            os.remove(caminho)

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
        "primary": (PRIMARY, ON_PRIMARY),
        "secondary": (SURFACE_ALT, TEXT),
        "danger": (DANGER, ON_PRIMARY),
    }[variante]
    botao = tk.Button(
        parent, text=texto, command=comando, bg=fundo, fg=frente,
        activebackground=PRIMARY_HOVER if variante == "primary" else BORDER,
        activeforeground=ON_PRIMARY if variante in ("primary", "danger") else TEXT,
        relief="flat", bd=0, cursor="hand2", font=("Segoe UI Semibold", 10),
        padx=18, pady=10, highlightthickness=1, highlightbackground=fundo,
    )
    botao.variante_prati = variante

    def ao_entrar(_event):
        cor = {"primary": PRIMARY_HOVER, "secondary": BORDER, "danger": DANGER_HOVER}[variante]
        botao.configure(bg=cor, highlightbackground=cor)

    def ao_sair(_event):
        cor = {"primary": PRIMARY, "secondary": SURFACE_ALT, "danger": DANGER}[variante]
        botao.configure(bg=cor, highlightbackground=cor)

    botao.bind("<Enter>", ao_entrar)
    botao.bind("<Leave>", ao_sair)
    return botao


def criar_card_acao(parent, icone, titulo, descricao, comando, destaque=False):
    fundo = PRIMARY if destaque else SURFACE
    frente = ON_PRIMARY if destaque else TEXT
    detalhe = PRIMARY_MUTED if destaque else MUTED
    card = tk.Frame(
        parent, bg=fundo, cursor="hand2", highlightbackground=PRIMARY if destaque else BORDER,
        highlightthickness=1, padx=16, pady=14,
    )
    topo = tk.Frame(card, bg=fundo, cursor="hand2")
    topo.pack(fill="x")
    tk.Label(
        topo, text=icone, bg=fundo, fg=ON_PRIMARY if destaque else PRIMARY,
        font=("Segoe UI Semibold", 16), cursor="hand2",
    ).pack(side="left", padx=(0, 10))
    textos = tk.Frame(topo, bg=fundo, cursor="hand2")
    textos.pack(side="left", fill="x", expand=True)
    tk.Label(textos, text=titulo, bg=fundo, fg=frente, font=("Segoe UI Semibold", 10), cursor="hand2").pack(anchor="w")
    tk.Label(textos, text=descricao, bg=fundo, fg=detalhe, font=("Segoe UI", 8), cursor="hand2").pack(anchor="w", pady=(2, 0))

    def widgets():
        return (card, topo, textos, *topo.winfo_children(), *textos.winfo_children())

    def configurar_cor(cor):
        for widget in widgets():
            widget.configure(bg=cor)

    def entrar(_event):
        configurar_cor(PRIMARY_HOVER if destaque else SURFACE_ALT)
        card.configure(highlightbackground=PRIMARY)

    def sair(_event):
        configurar_cor(PRIMARY if destaque else SURFACE)
        card.configure(highlightbackground=PRIMARY if destaque else BORDER)

    for widget in widgets():
        widget.bind("<Button-1>", lambda _event: comando())
        widget.bind("<Enter>", entrar)
        widget.bind("<Leave>", sair)
    return card


class PratiWMSApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.tema = carregar_tema()
        definir_tema(self.tema)
        self.title("Prati WMS | Gestão de estoque")
        self.geometry("1280x780")
        self.minsize(1050, 660)
        self.configure(bg=BG)
        try:
            self.db = DatabaseManager()
        except (OSError, sqlite3.Error) as error:
            messagebox.showerror(
                "Banco de dados indisponível",
                f"O Prati WMS não conseguiu abrir o banco de dados.\n\n{error}\n\n"
                "Mova o programa para uma pasta gravável ou restaure um backup.",
                parent=self,
            )
            self.destroy()
            raise SystemExit(1) from error
        self.paginas = {}
        self.botoes_nav = {}
        self.indicadores_nav = {}
        self.pagina_atual = "inicio"
        self._configurar_estilos()
        self._construir_interface()
        self.mostrar_pagina("inicio")
        if self.db.avisos_inicializacao:
            self.after(300, self._mostrar_avisos_inicializacao)

    def _configurar_estilos(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Treeview", background=SURFACE, fieldbackground=SURFACE, foreground=TEXT, rowheight=35, borderwidth=0, font=("Segoe UI", 9))
        style.map("Treeview", background=[("selected", PRIMARY)], foreground=[("selected", ON_PRIMARY)])
        style.configure("Treeview.Heading", background=SURFACE_ALT, foreground=MUTED, borderwidth=0, relief="flat", padding=(8, 10), font=("Segoe UI Semibold", 9))
        style.map("Treeview.Heading", background=[("active", SURFACE_ALT)])
        style.configure("TCombobox", fieldbackground=SURFACE_ALT, background=SURFACE_ALT, foreground=TEXT, arrowcolor=TEXT, bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER, padding=7)
        style.map("TCombobox", fieldbackground=[("readonly", SURFACE_ALT)], foreground=[("readonly", TEXT)])
        style.configure("Vertical.TScrollbar", background=SURFACE_ALT, troughcolor=BG, bordercolor=BG, arrowcolor=MUTED)

    def _construir_interface(self):
        header = tk.Frame(self, bg=HEADER, height=82)
        header.pack(fill="x")
        header.pack_propagate(False)

        marca = tk.Frame(header, bg=HEADER)
        marca.pack(side="left", padx=(26, 34), fill="y")
        logo = tk.Canvas(marca, width=42, height=42, bg=HEADER, highlightthickness=0)
        logo.pack(side="left", pady=20)
        logo.create_polygon(21, 2, 39, 12, 21, 22, 3, 12, fill=PRIMARY_HOVER, outline="")
        logo.create_polygon(3, 14, 20, 24, 20, 41, 3, 31, fill=PRIMARY, outline="")
        logo.create_polygon(22, 24, 39, 14, 39, 31, 22, 41, fill=HANDLE, outline="")
        nome_marca = tk.Frame(marca, bg=HEADER)
        nome_marca.pack(side="left", fill="y", padx=(9, 0))
        tk.Label(nome_marca, text="PRATI", bg=HEADER, fg=BRAND, font=("Segoe UI Black", 18)).pack(side="left", pady=22)
        tk.Label(nome_marca, text="WMS", bg=HEADER, fg=BRAND_SUB, font=("Segoe UI Semibold", 18)).pack(side="left", padx=(6, 0), pady=22)

        navegacao = tk.Frame(header, bg=HEADER)
        navegacao.pack(side="left", fill="y")
        itens_navegacao = (
            ("inicio", "⌂  Início"),
            ("consultar", "⌕  Consultar"),
            ("cadastrar", "+  Cadastrar"),
            ("historico", "◷  Histórico"),
            ("etiquetas", "◇  Etiquetas"),
        )
        for chave, texto in itens_navegacao:
            item = tk.Frame(navegacao, bg=HEADER)
            item.pack(side="left", fill="y")
            indicador = tk.Frame(item, bg=HEADER, height=3)
            indicador.pack(side="bottom", fill="x", padx=14)
            botao = tk.Button(
                item, text=texto, command=lambda pagina=chave: self.mostrar_pagina(pagina),
                bg=HEADER, fg=MUTED, activebackground=HEADER, activeforeground=TEXT,
                relief="flat", bd=0, cursor="hand2", font=("Segoe UI Semibold", 10),
                padx=14, pady=23,
            )
            botao.pack(side="top", fill="both", expand=True)
            botao.bind("<Enter>", lambda _event, pagina=chave: self._hover_navegacao(pagina, True))
            botao.bind("<Leave>", lambda _event, pagina=chave: self._hover_navegacao(pagina, False))
            self.botoes_nav[chave] = botao
            self.indicadores_nav[chave] = indicador

        self.botao_tema = criar_botao(header, self._texto_botao_tema(), self.alternar_tema, "secondary")
        self.botao_tema.configure(padx=13, pady=8)
        self.botao_tema.pack(side="right", padx=(8, 20))
        status = tk.Frame(header, bg=HEADER)
        status.pack(side="right", padx=(8, 2))
        tk.Label(status, text="●", bg=HEADER, fg=SUCCESS, font=("Segoe UI", 8)).pack(side="left")
        tk.Label(status, text="Dados locais", bg=HEADER, fg=MUTED, font=("Segoe UI", 9)).pack(side="left", padx=(5, 0))
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")

        self.container = tk.Frame(self, bg=BG)
        self.container.pack(fill="both", expand=True)
        for chave, classe in (("inicio", InicioView), ("consultar", ConsultarView), ("cadastrar", CadastrarView), ("historico", HistoricoView), ("etiquetas", EtiquetasView)):
            pagina = classe(self.container, self)
            pagina.place(relx=0, rely=0, relwidth=1, relheight=1)
            self.paginas[chave] = pagina

    def mostrar_pagina(self, chave):
        self.pagina_atual = chave
        for nome, botao in self.botoes_nav.items():
            ativo = nome == chave
            botao.configure(fg=PRIMARY if ativo else MUTED, bg=HEADER, activebackground=SURFACE_ALT)
            self.indicadores_nav[nome].configure(bg=PRIMARY if ativo else HEADER)
        pagina = self.paginas[chave]
        pagina.tkraise()
        pagina.on_show()

    def _hover_navegacao(self, chave, ativo):
        if chave == self.pagina_atual:
            return
        self.botoes_nav[chave].configure(
            bg=SURFACE_ALT if ativo else HEADER,
            fg=TEXT if ativo else MUTED,
        )

    def _texto_botao_tema(self):
        return "🌙 Escuro" if self.tema == "claro" else "☀ Claro"

    def alternar_tema(self):
        pagina_atual = self.pagina_atual
        editor = self.paginas.get("etiquetas")
        elementos = editor.elementos if editor else []
        selecionado = editor.selecionado if editor else None
        self.tema = "escuro" if self.tema == "claro" else "claro"
        definir_tema(self.tema)
        try:
            salvar_tema(self.tema)
        except OSError as error:
            messagebox.showwarning("Preferência não salva", str(error), parent=self)
        for filho in self.winfo_children():
            filho.destroy()
        self.configure(bg=BG)
        self.paginas = {}
        self.botoes_nav = {}
        self.indicadores_nav = {}
        self._configurar_estilos()
        self._construir_interface()
        novo_editor = self.paginas["etiquetas"]
        novo_editor.elementos = elementos
        novo_editor.selecionado = selecionado
        novo_editor._atualizar_previa()
        self.mostrar_pagina(pagina_atual)

    def atualizar_todas(self):
        for pagina in self.paginas.values():
            pagina.on_show()

    def _mostrar_avisos_inicializacao(self):
        messagebox.showwarning(
            "Manutenção do banco",
            "\n\n".join(self.db.avisos_inicializacao),
            parent=self,
        )


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
        self.action_cards = []
        self.metric_cards = []
        self.cabecalho("Visão geral", "Acompanhe o estoque e as últimas movimentações.")
        destaque = tk.Frame(self, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        destaque.pack(fill="x", padx=34, pady=(0, 18))
        tk.Frame(destaque, bg=PRIMARY, width=5).pack(side="left", fill="y")
        conteudo_destaque = tk.Frame(destaque, bg=SURFACE)
        conteudo_destaque.pack(side="left", fill="both", expand=True, padx=20, pady=15)
        chamada = tk.Frame(conteudo_destaque, bg=SURFACE)
        chamada.pack(fill="x", pady=(0, 12))
        tk.Label(chamada, text="AÇÕES RÁPIDAS", bg=SURFACE, fg=PRIMARY, font=("Segoe UI Semibold", 8)).pack(anchor="w")
        tk.Label(chamada, text="O que você deseja fazer?", bg=SURFACE, fg=TEXT, font=("Segoe UI Semibold", 15)).pack(anchor="w", pady=(3, 0))
        atalhos = tk.Frame(conteudo_destaque, bg=SURFACE)
        atalhos.pack(fill="x")
        acoes = (
            ("+", "Novo cadastro", "Adicionar item", lambda: self.app.mostrar_pagina("cadastrar"), True),
            ("⇧", "Importar Excel", "Cadastrar em lote", self._abrir_importacao, False),
            ("⌕", "Consultar", "Buscar no estoque", lambda: self.app.mostrar_pagina("consultar"), False),
            ("⤓", "Fazer backup", "Salvar uma cópia", self._fazer_backup, False),
        )
        for indice, (icone, titulo, descricao, comando, destaque_acao) in enumerate(acoes):
            card_acao = criar_card_acao(atalhos, icone, titulo, descricao, comando, destaque_acao)
            card_acao.grid(row=0, column=indice, sticky="nsew", padx=(0, 10 if indice < 3 else 0))
            atalhos.grid_columnconfigure(indice, weight=1, uniform="acoes")
            self.action_cards.append(card_acao)

        metricas = tk.Frame(self, bg=BG)
        metricas.pack(fill="x", padx=34)
        for indice, (chave, titulo, descricao, icone, cor) in enumerate((
            ("itens", "Itens cadastrados", "posições em uso", "▦", PRIMARY),
            ("unidades", "Unidades em estoque", "soma das quantidades", "Σ", SUCCESS),
            ("bloqueados", "Itens bloqueados", "requerem atenção", "!", WARNING),
        )):
            card = tk.Frame(metricas, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
            card.grid(row=0, column=indice, sticky="nsew", padx=(0, 12 if indice < 2 else 0))
            metricas.grid_columnconfigure(indice, weight=1)
            cabecalho_card = tk.Frame(card, bg=SURFACE)
            cabecalho_card.pack(fill="x", padx=18, pady=(15, 3))
            tk.Label(cabecalho_card, text=icone, bg=SURFACE_ALT, fg=cor, width=3, pady=5, font=("Segoe UI Semibold", 11)).pack(side="left")
            tk.Label(cabecalho_card, text=titulo, bg=SURFACE, fg=MUTED, font=("Segoe UI Semibold", 9)).pack(side="left", padx=(10, 0))
            label = tk.Label(card, text="—", bg=SURFACE, fg=TEXT, font=("Segoe UI Semibold", 25))
            label.pack(anchor="w", padx=18, pady=(2, 0))
            tk.Label(card, text=descricao, bg=SURFACE, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w", padx=19, pady=(0, 14))
            self.metric_labels[chave] = label
            self.metric_cards.append(card)

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

    def _fazer_backup(self):
        caminho = filedialog.asksaveasfilename(
            parent=self,
            title="Salvar backup do Prati WMS",
            initialdir=self.app.db.backup_dir,
            initialfile=f"Prati_WMS_backup_{datetime.now():%Y-%m-%d_%H%M}.db",
            defaultextension=".db",
            filetypes=(("Banco SQLite", "*.db"),),
        )
        if not caminho:
            return
        try:
            self.app.db.criar_backup(caminho)
        except (OSError, ValueError, sqlite3.Error) as error:
            messagebox.showerror("Backup não realizado", str(error), parent=self)
            return
        messagebox.showinfo(
            "Backup concluído",
            f"O backup foi salvo com segurança em:\n{caminho}",
            parent=self,
        )


class ConsultarView(BaseView):
    def __init__(self, parent, app):
        super().__init__(parent, app)
        self.busca = tk.StringVar()
        self.filtro_status = tk.StringVar(value="Todos")
        self.contagem = tk.StringVar(value="0 itens")
        self.acao_hover = None
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
        self.tabela.bind("<Motion>", self._hover_acoes)
        self.tabela.bind("<Leave>", lambda _event: self._limpar_hover_acoes())

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

    def _hover_acoes(self, event):
        item_id = self.tabela.identify_row(event.y)
        if self.tabela.identify_region(event.x, event.y) != "cell" or self.tabela.identify_column(event.x) != "#7" or not item_id:
            self._limpar_hover_acoes()
            return
        x, _y, largura, _altura = self.tabela.bbox(item_id, "acoes")
        lado = "editar" if event.x < x + largura / 2 else "remover"
        if self.acao_hover == (item_id, lado):
            return
        self._limpar_hover_acoes()
        valores = list(self.tabela.item(item_id, "values"))
        valores[6] = "▶ Editar     |     Remover" if lado == "editar" else "Editar     |     ▶ Remover"
        self.tabela.item(item_id, values=valores)
        self.tabela.configure(cursor="hand2")
        self.acao_hover = (item_id, lado)

    def _limpar_hover_acoes(self):
        if self.acao_hover:
            item_id, _lado = self.acao_hover
            if self.tabela.exists(item_id):
                valores = list(self.tabela.item(item_id, "values"))
                valores[6] = "Editar     |     Remover"
                self.tabela.item(item_id, values=valores)
        self.tabela.configure(cursor="")
        self.acao_hover = None

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
        self.cabecalho("Histórico", "Entradas, alterações e remoções dos últimos 2 dias.")
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


class TextoEtiquetaDialog(tk.Toplevel):
    def __init__(self, parent, titulo, texto="", tamanho=28):
        super().__init__(parent)
        self.resultado = None
        self.title(titulo)
        self.geometry("620x430")
        self.minsize(520, 360)
        self.configure(bg=BG)
        self.transient(parent.winfo_toplevel())
        self.grab_set()

        tk.Label(self, text=titulo, bg=BG, fg=TEXT, font=("Segoe UI Semibold", 20)).pack(
            anchor="w", padx=28, pady=(24, 4),
        )
        tk.Label(
            self, text="Digite o conteúdo com clareza. Use Enter para criar novas linhas.",
            bg=BG, fg=MUTED, font=("Segoe UI", 9),
        ).pack(anchor="w", padx=28, pady=(0, 16))

        card = tk.Frame(self, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill="both", expand=True, padx=28, pady=(0, 18))
        tk.Label(card, text="TEXTO DA ETIQUETA", bg=SURFACE, fg=MUTED, font=("Segoe UI Semibold", 9)).pack(
            anchor="w", padx=18, pady=(16, 7),
        )
        self.entrada = tk.Text(
            card, height=8, wrap="word", undo=True, bg=SURFACE_ALT, fg=TEXT,
            insertbackground=TEXT, selectbackground=PRIMARY, selectforeground=ON_PRIMARY,
            relief="flat", bd=0, font=("Segoe UI", 13), padx=12, pady=10,
            highlightbackground=BORDER, highlightcolor=PRIMARY, highlightthickness=1,
        )
        self.entrada.pack(fill="both", expand=True, padx=18)
        self.entrada.insert("1.0", texto)

        tamanho_frame = tk.Frame(card, bg=SURFACE)
        tamanho_frame.pack(fill="x", padx=18, pady=14)
        tk.Label(tamanho_frame, text="Tamanho da fonte", bg=SURFACE, fg=TEXT, font=("Segoe UI Semibold", 9)).pack(side="left")
        self.tamanho = tk.IntVar(value=int(tamanho))
        tk.Spinbox(
            tamanho_frame, from_=8, to=160, textvariable=self.tamanho, width=6,
            bg=SURFACE_ALT, fg=TEXT, insertbackground=TEXT, buttonbackground=SURFACE_ALT,
            relief="flat", font=("Segoe UI", 10), highlightbackground=BORDER,
            highlightcolor=PRIMARY, highlightthickness=1,
        ).pack(side="left", padx=(12, 6), ipady=5)
        tk.Label(tamanho_frame, text="pt  •  Fonte fixa: Arial Black", bg=SURFACE, fg=MUTED, font=("Segoe UI", 9)).pack(side="left")

        acoes = tk.Frame(self, bg=BG)
        acoes.pack(fill="x", padx=28, pady=(0, 24))
        criar_botao(acoes, "Cancelar", self.destroy, "secondary").pack(side="right")
        criar_botao(acoes, "Aplicar texto", self._confirmar).pack(side="right", padx=(0, 10))
        self.bind("<Escape>", lambda _event: self.destroy())
        self.bind("<Control-Return>", lambda _event: self._confirmar())
        self.after_idle(self._preparar_foco)

    def _preparar_foco(self):
        self.entrada.focus_set()
        self.entrada.mark_set("insert", "end-1c")

    def _confirmar(self):
        texto = self.entrada.get("1.0", "end-1c").strip()
        if not texto:
            messagebox.showwarning("Texto vazio", "Digite o texto que será colocado na etiqueta.", parent=self)
            return
        try:
            tamanho = int(self.tamanho.get())
        except (TypeError, ValueError):
            tamanho = 28
        self.resultado = {"texto": texto, "tamanho": max(8, min(160, tamanho))}
        self.destroy()


class EtiquetasView(BaseView):
    LARGURA_PREVIA = 558
    ALTURA_PREVIA = 372

    def __init__(self, parent, app):
        super().__init__(parent, app)
        self.elementos = []
        self.selecionado = None
        self.modo_arraste = None
        self.ultimo_ponto = None
        self.dados_redimensionamento = None
        self.fantasma_texto = None
        self.imagem_previa = None
        self.impressora = tk.StringVar()
        self.status_impressora = tk.StringVar(value="Buscando impressoras...")
        self.cabecalho("Etiquetas", "Monte etiquetas Zebra SAMPLE horizontais com texto ou imagem.")
        self._construir()
        self.nova_etiqueta(confirmar=False)

    def _construir(self):
        barra = tk.Frame(self, bg=BG)
        barra.pack(fill="x", padx=34, pady=(0, 14))
        criar_botao(barra, "Nova etiqueta", self.nova_etiqueta).pack(side="left")
        criar_botao(barra, "Abrir modelo", self.abrir_modelo, "secondary").pack(side="left", padx=8)
        criar_botao(barra, "Salvar modelo", self.salvar_modelo, "secondary").pack(side="left")
        tk.Label(barra, text="Zebra SAMPLE  •  6 × 4 pol.  •  203 dpi  •  Horizontal", bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(side="right")

        area = tk.Frame(self, bg=BG)
        area.pack(fill="both", expand=True, padx=34, pady=(0, 24))

        painel_conteudo = tk.Frame(area, bg=SURFACE, width=220, highlightbackground=BORDER, highlightthickness=1)
        painel_conteudo.pack(side="left", fill="y", padx=(0, 16))
        painel_conteudo.pack_propagate(False)
        tk.Label(painel_conteudo, text="ELEMENTOS", bg=SURFACE, fg=MUTED, font=("Segoe UI Semibold", 9)).pack(anchor="w", padx=18, pady=(18, 14))
        self.ferramenta_texto = tk.Label(
            painel_conteudo, text="T   Adicionar texto", bg=SURFACE_ALT, fg=TEXT,
            relief="flat", cursor="hand2", font=("Segoe UI Semibold", 10),
            padx=14, pady=14, highlightbackground=BORDER, highlightthickness=1,
        )
        self.ferramenta_texto.pack(fill="x", padx=18)
        self.ferramenta_texto.bind("<ButtonPress-1>", self._iniciar_ferramenta_texto)
        self.ferramenta_texto.bind("<B1-Motion>", self._mover_ferramenta_texto)
        self.ferramenta_texto.bind("<ButtonRelease-1>", self._soltar_ferramenta_texto)
        self.ferramenta_texto.bind("<Enter>", lambda _event: self.ferramenta_texto.configure(bg=BORDER))
        self.ferramenta_texto.bind("<Leave>", lambda _event: self.ferramenta_texto.configure(bg=SURFACE_ALT))
        tk.Label(
            painel_conteudo,
            text="Arraste a ferramenta para a etiqueta.\n\nDuplo clique: editar texto\nAlças laterais: largura\nAlças superior/inferior: altura\nCtrl+V: colar imagem\nDelete: excluir selecionado",
            justify="left", bg=SURFACE, fg=MUTED, font=("Segoe UI", 8),
        ).pack(anchor="w", padx=18, pady=18)
        tk.Label(painel_conteudo, text="Fonte fixa\nArial Black", justify="left", bg=SURFACE, fg=TEXT, font=("Arial Black", 11)).pack(side="bottom", anchor="w", padx=18, pady=18)

        centro = tk.Frame(area, bg=BG)
        centro.pack(side="left", fill="both", expand=True)
        tk.Label(centro, text="PRÉ-VISUALIZAÇÃO", bg=BG, fg=MUTED, font=("Segoe UI Semibold", 9)).pack(pady=(0, 8))
        self.canvas = tk.Canvas(
            centro, width=self.LARGURA_PREVIA, height=self.ALTURA_PREVIA,
            bg="white", bd=0, highlightbackground=BORDER, highlightthickness=2,
            cursor="arrow",
        )
        self.canvas.pack()
        self.canvas.bind("<Button-1>", self._iniciar_arraste)
        self.canvas.bind("<B1-Motion>", self._arrastar)
        self.canvas.bind("<ButtonRelease-1>", self._finalizar_arraste)
        self.canvas.bind("<Double-1>", self._editar_texto)
        self.canvas.bind("<Motion>", self._atualizar_cursor)
        self.canvas.bind("<Delete>", lambda _event: self.excluir_selecionado())
        self.canvas.bind("<Control-v>", lambda _event: self.colar_imagem())

        painel_impressao = tk.Frame(area, bg=SURFACE, width=260, highlightbackground=BORDER, highlightthickness=1)
        painel_impressao.pack(side="right", fill="y", padx=(16, 0))
        painel_impressao.pack_propagate(False)
        tk.Label(painel_impressao, text="IMPRESSÃO", bg=SURFACE, fg=MUTED, font=("Segoe UI Semibold", 9)).pack(anchor="w", padx=18, pady=(18, 14))
        tk.Label(painel_impressao, text="Impressora do Windows", bg=SURFACE, fg=TEXT, font=("Segoe UI Semibold", 9)).pack(anchor="w", padx=18)
        self.combo_impressora = ttk.Combobox(painel_impressao, textvariable=self.impressora, state="readonly")
        self.combo_impressora.pack(fill="x", padx=18, pady=(7, 10))
        criar_botao(painel_impressao, "Atualizar impressoras", self.atualizar_impressoras, "secondary").pack(fill="x", padx=18)
        tk.Label(painel_impressao, textvariable=self.status_impressora, wraplength=220, justify="left", bg=SURFACE, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w", padx=18, pady=14)
        informacao = tk.Frame(painel_impressao, bg=SURFACE_ALT)
        informacao.pack(fill="x", padx=18, pady=(6, 18))
        tk.Label(informacao, text="FORMATO DA MÍDIA", bg=SURFACE_ALT, fg=MUTED, font=("Segoe UI Semibold", 8)).pack(anchor="w", padx=12, pady=(12, 4))
        tk.Label(informacao, text="152,4 × 101,6 mm\n1218 × 812 pontos\nOrientação: horizontal", justify="left", bg=SURFACE_ALT, fg=TEXT, font=("Segoe UI", 9)).pack(anchor="w", padx=12, pady=(0, 12))
        criar_botao(painel_impressao, "Imprimir etiqueta", self.imprimir).pack(side="bottom", fill="x", padx=18, pady=18)

    def on_show(self):
        self.atualizar_impressoras()
        self._atualizar_previa()

    def nova_etiqueta(self, confirmar=True):
        if confirmar and self.elementos and not messagebox.askyesno("Nova etiqueta", "Descartar o conteúdo atual e criar uma nova etiqueta?", parent=self):
            return
        self.elementos = []
        self.selecionado = None
        self._atualizar_previa()

    def abrir_modelo(self):
        caminho = filedialog.askopenfilename(parent=self, title="Abrir modelo de etiqueta", filetypes=(("Modelo Prati", "*.prati-etiqueta"), ("Arquivo JSON", "*.json")))
        if not caminho:
            return
        try:
            self.elementos = abrir_modelo_etiqueta(caminho)
            self.selecionado = None
            self._atualizar_previa()
        except (OSError, ValueError, KeyError, json.JSONDecodeError, base64.binascii.Error) as error:
            messagebox.showerror("Modelo inválido", str(error), parent=self)

    def salvar_modelo(self):
        caminho = filedialog.asksaveasfilename(parent=self, title="Salvar modelo de etiqueta", defaultextension=".prati-etiqueta", filetypes=(("Modelo Prati", "*.prati-etiqueta"),))
        if not caminho:
            return
        try:
            salvar_modelo_etiqueta(caminho, self.elementos)
            messagebox.showinfo("Modelo salvo", "O modelo foi salvo com sucesso.", parent=self)
        except OSError as error:
            messagebox.showerror("Não foi possível salvar", str(error), parent=self)

    def _iniciar_ferramenta_texto(self, event):
        self._destruir_fantasma_texto()
        self.fantasma_texto = tk.Toplevel(self)
        self.fantasma_texto.overrideredirect(True)
        self.fantasma_texto.attributes("-topmost", True)
        tk.Label(self.fantasma_texto, text="T  Texto", bg=PRIMARY, fg=ON_PRIMARY, font=("Segoe UI Semibold", 10), padx=12, pady=7).pack()
        self._mover_ferramenta_texto(event)

    def _mover_ferramenta_texto(self, event):
        if self.fantasma_texto is not None:
            self.fantasma_texto.geometry(f"+{event.x_root + 12}+{event.y_root + 12}")

    def _soltar_ferramenta_texto(self, event):
        self._destruir_fantasma_texto()
        x_canvas = event.x_root - self.canvas.winfo_rootx()
        y_canvas = event.y_root - self.canvas.winfo_rooty()
        if 0 <= x_canvas <= self.LARGURA_PREVIA and 0 <= y_canvas <= self.ALTURA_PREVIA:
            x = x_canvas * LABEL_WIDTH_PX / self.LARGURA_PREVIA
            y = y_canvas * LABEL_HEIGHT_PX / self.ALTURA_PREVIA
            self._adicionar_texto(x, y)

    def _destruir_fantasma_texto(self):
        if self.fantasma_texto is not None:
            self.fantasma_texto.destroy()
            self.fantasma_texto = None

    def _adicionar_texto(self, x, y):
        dialogo = TextoEtiquetaDialog(self, "Adicionar texto")
        self.wait_window(dialogo)
        if dialogo.resultado is None:
            return
        self.elementos.append({
            "tipo": "texto", "x": x, "y": y,
            "texto": dialogo.resultado["texto"], "tamanho": dialogo.resultado["tamanho"],
            "escala_x": 1.0, "escala_y": 1.0,
        })
        self.selecionado = len(self.elementos) - 1
        self._atualizar_previa()

    def _editar_texto(self, event):
        x, y = self._ponto_modelo(event)
        indice = self._elemento_no_ponto(x, y)
        if indice is None or self.elementos[indice]["tipo"] != "texto":
            return
        elemento = self.elementos[indice]
        dialogo = TextoEtiquetaDialog(self, "Editar texto", elemento["texto"], elemento["tamanho"])
        self.wait_window(dialogo)
        if dialogo.resultado is not None:
            elemento["texto"] = dialogo.resultado["texto"]
            elemento["tamanho"] = dialogo.resultado["tamanho"]
            self.selecionado = indice
            self._atualizar_previa()

    def colar_imagem(self):
        try:
            conteudo = ImageGrab.grabclipboard()
            if isinstance(conteudo, list) and conteudo:
                conteudo = Image.open(conteudo[0])
            if not isinstance(conteudo, Image.Image):
                raise ValueError("A área de transferência não contém uma imagem.")
            imagem = conteudo.convert("RGBA")
            escala = min(600 / imagem.width, 600 / imagem.height, 1.0)
            largura = max(1, round(imagem.width * escala))
            altura = max(1, round(imagem.height * escala))
            self.elementos.append({"tipo": "imagem", "x": LABEL_WIDTH_PX // 2, "y": LABEL_HEIGHT_PX // 2, "largura": largura, "altura": altura, "imagem": imagem})
            self.selecionado = len(self.elementos) - 1
            self._atualizar_previa()
        except (OSError, ValueError) as error:
            messagebox.showwarning("Imagem não encontrada", str(error), parent=self)

    def excluir_selecionado(self):
        if self.selecionado is None:
            return
        del self.elementos[self.selecionado]
        self.selecionado = None
        self._atualizar_previa()

    def atualizar_impressoras(self):
        try:
            impressoras = listar_impressoras()
            self.combo_impressora.configure(values=impressoras)
            if impressoras:
                try:
                    padrao = win32print.GetDefaultPrinter()
                except Exception:
                    padrao = impressoras[0]
                self.impressora.set(padrao if padrao in impressoras else impressoras[0])
                self.status_impressora.set(f"{len(impressoras)} impressora(s) disponível(is). Usa o driver instalado no Windows.")
            else:
                self.impressora.set("")
                self.status_impressora.set("Nenhuma impressora instalada. O editor e os modelos continuam disponíveis.")
        except Exception as error:
            self.impressora.set("")
            self.status_impressora.set(f"Não foi possível consultar as impressoras: {error}")

    def imprimir(self):
        if not self.elementos:
            messagebox.showwarning("Etiqueta vazia", "Adicione texto ou imagem antes de imprimir.", parent=self)
            return
        if not self.impressora.get():
            messagebox.showwarning("Sem impressora", "Instale ou conecte uma impressora no Windows e atualize a lista.", parent=self)
            return
        try:
            imprimir_imagem_windows(renderizar_etiqueta(self.elementos), self.impressora.get())
            messagebox.showinfo("Etiqueta enviada", f"A etiqueta foi enviada para {self.impressora.get()}.", parent=self)
        except Exception as error:
            messagebox.showerror("Falha na impressão", str(error), parent=self)

    def _caixa_elemento(self, elemento):
        if elemento["tipo"] == "imagem":
            meia_largura = elemento["largura"] / 2
            meia_altura = elemento["altura"] / 2
            return (elemento["x"] - meia_largura, elemento["y"] - meia_altura, elemento["x"] + meia_largura, elemento["y"] + meia_altura)
        largura, altura = dimensoes_texto_etiqueta(elemento)[:2]
        return (elemento["x"] - largura / 2, elemento["y"] - altura / 2, elemento["x"] + largura / 2, elemento["y"] + altura / 2)

    def _atualizar_previa(self):
        if not hasattr(self, "canvas"):
            return
        imagem = renderizar_etiqueta(self.elementos).resize((self.LARGURA_PREVIA, self.ALTURA_PREVIA), Image.Resampling.LANCZOS)
        self.imagem_previa = ImageTk.PhotoImage(imagem)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, image=self.imagem_previa, anchor="nw")
        if self.selecionado is not None and self.selecionado < len(self.elementos):
            x1, y1, x2, y2 = self._caixa_elemento(self.elementos[self.selecionado])
            escala_x = self.LARGURA_PREVIA / LABEL_WIDTH_PX
            escala_y = self.ALTURA_PREVIA / LABEL_HEIGHT_PX
            caixa = (x1 * escala_x, y1 * escala_y, x2 * escala_x, y2 * escala_y)
            self.canvas.create_rectangle(*caixa, outline=HANDLE, width=2, dash=(4, 2))
            for _nome, x, y in self._alcas_selecao(caixa):
                self.canvas.create_rectangle(x - 5, y - 5, x + 5, y + 5, fill=HANDLE, outline="white", width=1)

    def _ponto_modelo(self, event):
        return (event.x * LABEL_WIDTH_PX / self.LARGURA_PREVIA, event.y * LABEL_HEIGHT_PX / self.ALTURA_PREVIA)

    @staticmethod
    def _alcas_selecao(caixa):
        x1, y1, x2, y2 = caixa
        meio_x = (x1 + x2) / 2
        meio_y = (y1 + y2) / 2
        return (
            ("nw", x1, y1), ("n", meio_x, y1), ("ne", x2, y1),
            ("e", x2, meio_y), ("se", x2, y2), ("s", meio_x, y2),
            ("sw", x1, y2), ("w", x1, meio_y),
        )

    def _caixa_selecao_canvas(self):
        if self.selecionado is None or self.selecionado >= len(self.elementos):
            return None
        x1, y1, x2, y2 = self._caixa_elemento(self.elementos[self.selecionado])
        return (
            x1 * self.LARGURA_PREVIA / LABEL_WIDTH_PX,
            y1 * self.ALTURA_PREVIA / LABEL_HEIGHT_PX,
            x2 * self.LARGURA_PREVIA / LABEL_WIDTH_PX,
            y2 * self.ALTURA_PREVIA / LABEL_HEIGHT_PX,
        )

    def _alca_no_ponto(self, x, y):
        caixa = self._caixa_selecao_canvas()
        if caixa is None:
            return None
        for nome, alca_x, alca_y in self._alcas_selecao(caixa):
            if abs(x - alca_x) <= 9 and abs(y - alca_y) <= 9:
                return nome
        return None

    def _elemento_no_ponto(self, x, y):
        for indice in range(len(self.elementos) - 1, -1, -1):
            x1, y1, x2, y2 = self._caixa_elemento(self.elementos[indice])
            if x1 <= x <= x2 and y1 <= y <= y2:
                return indice
        return None

    def _atualizar_cursor(self, event):
        alca = self._alca_no_ponto(event.x, event.y)
        if alca in ("e", "w"):
            self.canvas.configure(cursor="sb_h_double_arrow")
            return
        if alca in ("n", "s"):
            self.canvas.configure(cursor="sb_v_double_arrow")
            return
        if alca is not None:
            self.canvas.configure(cursor="sizing")
            return
        x, y = self._ponto_modelo(event)
        self.canvas.configure(cursor="fleur" if self._elemento_no_ponto(x, y) is not None else "arrow")

    def _iniciar_arraste(self, event):
        self.canvas.focus_set()
        x, y = self._ponto_modelo(event)
        alca = self._alca_no_ponto(event.x, event.y)
        if alca is not None and self.selecionado is not None:
            elemento = self.elementos[self.selecionado]
            x1, y1, x2, y2 = self._caixa_elemento(elemento)
            self.dados_redimensionamento = {
                "alca": alca,
                "meia_largura": max(1, (x2 - x1) / 2),
                "meia_altura": max(1, (y2 - y1) / 2),
                "tamanho": elemento.get("tamanho"),
                "largura": elemento.get("largura"),
                "altura": elemento.get("altura"),
                "escala_x": float(elemento.get("escala_x", 1.0)),
                "escala_y": float(elemento.get("escala_y", 1.0)),
            }
            self.modo_arraste = "redimensionar"
        else:
            self.selecionado = self._elemento_no_ponto(x, y)
            self.modo_arraste = "mover" if self.selecionado is not None else None
        self.ultimo_ponto = (x, y)
        self._atualizar_previa()

    def _arrastar(self, event):
        if self.modo_arraste is None or self.selecionado is None:
            return
        x, y = self._ponto_modelo(event)
        elemento = self.elementos[self.selecionado]
        if self.modo_arraste == "redimensionar":
            dados = self.dados_redimensionamento
            alca = dados["alca"]
            altera_largura = alca in ("nw", "ne", "e", "se", "sw", "w")
            altera_altura = alca in ("nw", "n", "ne", "se", "s", "sw")
            fator_x = max(0.05, abs(x - elemento["x"]) / dados["meia_largura"])
            fator_y = max(0.05, abs(y - elemento["y"]) / dados["meia_altura"])
            if elemento["tipo"] == "texto":
                if altera_largura:
                    elemento["escala_x"] = max(0.2, min(8.0, dados["escala_x"] * fator_x))
                if altera_altura:
                    elemento["escala_y"] = max(0.2, min(8.0, dados["escala_y"] * fator_y))
            else:
                if altera_largura:
                    elemento["largura"] = max(20, min(LABEL_WIDTH_PX, round(dados["largura"] * fator_x)))
                if altera_altura:
                    elemento["altura"] = max(20, min(LABEL_HEIGHT_PX, round(dados["altura"] * fator_y)))
        else:
            anterior_x, anterior_y = self.ultimo_ponto
            elemento["x"] = max(0, min(LABEL_WIDTH_PX, elemento["x"] + x - anterior_x))
            elemento["y"] = max(0, min(LABEL_HEIGHT_PX, elemento["y"] + y - anterior_y))
        self.ultimo_ponto = (x, y)
        self._atualizar_previa()

    def _finalizar_arraste(self, _event):
        self.modo_arraste = None
        self.ultimo_ponto = None
        self.dados_redimensionamento = None


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
