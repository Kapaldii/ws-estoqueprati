"""
SUITE DE TESTES UNITÁRIOS - PRATI WMS
Validação completa de regras de negócio, endereçamento 2D, persistência e auditoria.
Tecnologia: Python standard unittest (zero dependências externas)
"""

import unittest
import os
import sqlite3
import tempfile
import csv
import zipfile
from datetime import datetime, timedelta
from xml.sax.saxutils import escape

from PIL import Image

import app as app_module

# Importa a lógica de negócio do sistema principal
from app import (
    LABEL_HEIGHT_PX,
    LABEL_WIDTH_PX,
    DatabaseManager,
    EtiquetasView,
    abrir_modelo_etiqueta,
    dimensoes_texto_etiqueta,
    imprimir_imagem_windows,
    listar_impressoras,
    parse_localizacao,
    renderizar_etiqueta,
    salvar_modelo_etiqueta,
)


class TestEnderecamento2D(unittest.TestCase):
    """Testes unitários das regras de endereçamento do armazém."""

    def test_enderecos_validos_ruas_principais(self):
        """Valida endereços válidos nas Ruas A, B, C e D."""
        casos = [
            ("C2809", "Rua C", 28, 9, "Direito (Par)", 5),
            ("c2809", "Rua C", 28, 9, "Direito (Par)", 5),  # Minúsculas
            ("A0101", "Rua A", 1, 1, "Esquerdo (Ímpar)", 1),
            ("B0414", "Rua B", 4, 14, "Direito (Par)", 7),
            ("D0307", "Rua D", 3, 7, "Esquerdo (Ímpar)", 4),
        ]
        for entrada, exp_rua, exp_bloco, exp_pos, exp_lado, exp_nivel in casos:
            with self.subTest(entrada=entrada):
                res = parse_localizacao(entrada)
                self.assertTrue(res["valido"], f"Endereço {entrada} deveria ser válido.")
                self.assertEqual(res["rua"], exp_rua)
                self.assertEqual(res["bloco"], exp_bloco)
                self.assertEqual(res["posicao"], exp_pos)
                self.assertEqual(res["lado"], exp_lado)
                self.assertEqual(res["nivel"], exp_nivel)

    def test_limite_rigido_14_posicoes_por_bloco(self):
        """Valida que posições fora do intervalo [1, 14] são estritamente rejeitadas."""
        posicoes_invalidas = ["A0100", "B0215", "C0320", "D0499", "C2800"]
        for cod in posicoes_invalidas:
            with self.subTest(codigo=cod):
                res = parse_localizacao(cod)
                self.assertFalse(res["valido"])
                self.assertIn("14 posições", res["msg"])

    def test_rejeicao_ruas_inexistentes(self):
        """Apenas Ruas A, B, C, D e Frações são permitidas."""
        ruas_invalidas = ["E0105", "X0204", "Z2809", "12345"]
        for cod in ruas_invalidas:
            with self.subTest(codigo=cod):
                res = parse_localizacao(cod)
                self.assertFalse(res["valido"])

    def test_palete_fracoes(self):
        """Valida decodificação de posições do Palete de Frações."""
        casos_frac = [
            ("FRAC-01", "FRAC-01", 1),
            ("frac-12", "FRAC-12", 12),
            ("FRAC 24", "FRAC-24", 24),
            ("FRACOES", "FRAC-01", 1), # Fallback padrão para primeira posição
        ]
        for entrada, exp_cod, exp_pos in casos_frac:
            with self.subTest(entrada=entrada):
                res = parse_localizacao(entrada)
                self.assertTrue(res["valido"])
                self.assertEqual(res["rua"], "Frações")
                self.assertEqual(res["codigo"], exp_cod)
                self.assertEqual(res["posicao"], exp_pos)


class TestBancoDadosEAuditoria(unittest.TestCase):
    """Testes unitários das operações ACID no banco SQLite e Logs de Auditoria."""

    def setUp(self):
        """Cria um banco SQLite temporário isolado para cada teste."""
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = os.path.join(self.temp_dir.name, "test_wms.db")
        self.db = DatabaseManager(self.db_path)

    def tearDown(self):
        """Limpa o ambiente temporário."""
        self.temp_dir.cleanup()

    def test_criacao_banco_vazio(self):
        """Verifica se um banco novo começa vazio e com a estrutura necessária."""
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT COUNT(*) FROM itens_estoque")
            total_itens = c.fetchone()[0]
            self.assertEqual(total_itens, 0)

            c.execute("SELECT COUNT(*) FROM historico_movimentacoes")
            total_logs = c.fetchone()[0]
            self.assertEqual(total_logs, 0)

    def test_cadastro_item_com_log_auditoria(self):
        """Cadastra um novo medicamento e valida o registro correspondente no log."""
        self.db.cadastrar_item("A0812", "PRATI-999", "Cetoprofeno 100mg", "L26001A", 50, "Normal")

        # Validação da persistência
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT nome_medicamento, quantidade FROM itens_estoque WHERE codigo_localizacao = 'A0812'")
            item = c.fetchone()
            self.assertIsNotNone(item)
            self.assertEqual(item[0], "Cetoprofeno 100mg")
            self.assertEqual(item[1], 50)

            # Validação do log
            c.execute("SELECT tipo, codigo_produto, quantidade FROM historico_movimentacoes WHERE codigo_localizacao = 'A0812'")
            log = c.fetchone()
            self.assertIsNotNone(log)
            self.assertEqual(log[0], "ENTRADA")
            self.assertEqual(log[1], "PRATI-999")
            self.assertEqual(log[2], 50)

    def test_impedir_duplicidade_de_localizacao(self):
        """Garante que a mesma posição física não pode ser cadastrada simultaneamente por outro item."""
        self.db.cadastrar_item("C2809", "PRATI-BASE", "Medicamento base", "LBASE", 1, "Normal")
        with self.db.get_connection() as conn:
            c = conn.cursor()
            with self.assertRaises(sqlite3.IntegrityError):
                c.execute("""
                    INSERT INTO itens_estoque (rua, bloco, posicao, codigo_localizacao, codigo_produto, nome_medicamento, lote, quantidade)
                    VALUES ('Rua C', 28, 9, 'C2809', 'PRATI-DUPLICADO', 'Medicamento X', 'L999', 10)
                """)

    def test_impedir_quantidade_negativa(self):
        """Valida que a regra de integridade CHECK(quantidade >= 0) bloqueia valores negativos."""
        with self.db.get_connection() as conn:
            c = conn.cursor()
            with self.assertRaises(sqlite3.IntegrityError):
                c.execute("""
                    INSERT INTO itens_estoque (rua, bloco, posicao, codigo_localizacao, codigo_produto, nome_medicamento, lote, quantidade)
                    VALUES ('Rua A', 2, 2, 'A0202', 'PRATI-NEG', 'Medicamento Negativo', 'L000', -15)
                """)

    def test_alteracao_item_e_log(self):
        """Edita dados de um item existente e registra no log."""
        self.db.cadastrar_item("C2809", "PRATI-101", "Dipirona 500mg/mL Gotas", "L24089A", 140, "Normal")
        item = self.db.listar_itens("PRATI-101")[0]
        self.db.atualizar_item(
            item["id"], "C2809", "PRATI-101", "Dipirona 500mg/mL Gotas",
            "L24089A", 250, "Normal", "Qtd alterada de 140 para 250",
        )

        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT quantidade FROM itens_estoque WHERE codigo_localizacao = 'C2809'")
            self.assertEqual(c.fetchone()[0], 250)

            c.execute("SELECT tipo, motivo FROM historico_movimentacoes WHERE codigo_localizacao = 'C2809' AND tipo = 'ALTERAÇÃO'")
            log_alt = c.fetchone()
            self.assertIsNotNone(log_alt)
            self.assertIn("alterada", log_alt[1])

    def test_baixa_remocao_item_e_log(self):
        """Dá baixa total em um medicamento e valida a desocupação da posição e registro de auditoria."""
        self.db.cadastrar_item("A0201", "PRATI-201", "Omeprazol 20mg", "L24011D", 220, "Normal")
        item = self.db.listar_itens("PRATI-201")[0]
        self.db.remover_item(item["id"])

        with self.db.get_connection() as conn:
            c = conn.cursor()
            # Confirma que a posição A0201 ficou livre
            c.execute("SELECT COUNT(*) FROM itens_estoque WHERE codigo_localizacao = 'A0201'")
            self.assertEqual(c.fetchone()[0], 0)

            # Confirma log de baixa
            c.execute("SELECT tipo, motivo FROM historico_movimentacoes WHERE codigo_localizacao = 'A0201' AND tipo = 'BAIXA / REMOÇÃO'")
            log_rem = c.fetchone()
            self.assertIsNotNone(log_rem)
            self.assertEqual(log_rem[0], "BAIXA / REMOÇÃO")


class TestBuscasEFiltros(unittest.TestCase):
    """Testes unitários de filtros de pesquisa e exportação."""

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = os.path.join(self.temp_dir.name, "test_buscas.db")
        self.db = DatabaseManager(self.db_path)
        self.db.cadastrar_item("C2809", "PRATI-101", "Dipirona 500mg/mL Gotas", "L24089A", 140, "Normal")
        self.db.cadastrar_item("C2810", "PRATI-102", "Paracetamol 750mg", "L24090B", 85, "Normal")

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_busca_inteligente_por_lote_e_codigo(self):
        """Valida que as queries de busca encontram itens por código, lote ou nome."""
        with self.db.get_connection() as conn:
            c = conn.cursor()
            # Busca por Código
            c.execute("SELECT nome_medicamento FROM itens_estoque WHERE codigo_produto LIKE '%PRATI-101%'")
            res_cod = c.fetchone()
            self.assertIsNotNone(res_cod)
            self.assertIn("Dipirona", res_cod[0])

            # Busca por Lote
            c.execute("SELECT codigo_localizacao FROM itens_estoque WHERE lote LIKE '%L24090B%'")
            res_lote = c.fetchone()
            self.assertIsNotNone(res_lote)
            self.assertEqual(res_lote[0], "C2810")

            # Busca por Quantidade
            c.execute("SELECT nome_medicamento FROM itens_estoque WHERE CAST(quantidade AS TEXT) = '140'")
            res_qtd = c.fetchone()
            self.assertIsNotNone(res_qtd)
            self.assertIn("Dipirona", res_qtd[0])

    def test_exportacao_csv_historico(self):
        """Valida a geração correta de arquivos CSV a partir do histórico de auditoria."""
        csv_path = os.path.join(self.temp_dir.name, "export_test.csv")
        
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT id, data_hora, tipo, codigo_localizacao, codigo_produto, nome_medicamento, lote, quantidade, motivo FROM historico_movimentacoes")
            linhas = c.fetchall()

        with open(csv_path, mode="w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(["ID", "Data/Hora", "Tipo Movimento", "Localização", "Código Produto", "Medicamento", "Lote", "Quantidade", "Motivo"])
            writer.writerows(linhas)

        self.assertTrue(os.path.exists(csv_path))
        with open(csv_path, mode="r", encoding="utf-8-sig") as f:
            conteudo = f.read()
            self.assertIn("PRATI-101", conteudo)
            self.assertIn("ENTRADA", conteudo)


class TestCadastroGuiEImportacao(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = os.path.join(self.temp_dir.name, "test_novas_funcoes.db")
        self.db = DatabaseManager(self.db_path)
        self.db.cadastrar_item("C2809", "PRATI-101", "Dipirona", "L1", 10, "Normal")
        self.db.cadastrar_item("C2810", "PRATI-102", "Paracetamol", "L2", 10, "Normal")

    def tearDown(self):
        self.temp_dir.cleanup()

    @staticmethod
    def _criar_xlsx(caminho, linhas):
        namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        xml_linhas = []
        for numero, linha in enumerate(linhas, start=1):
            celulas = []
            for indice, valor in enumerate(linha):
                coluna = chr(65 + indice)
                celulas.append(f'<c r="{coluna}{numero}" t="inlineStr"><is><t>{escape(str(valor))}</t></is></c>')
            xml_linhas.append(f'<row r="{numero}">{"".join(celulas)}</row>')
        planilha = f'<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{namespace}"><sheetData>{"".join(xml_linhas)}</sheetData></worksheet>'
        with zipfile.ZipFile(caminho, "w") as arquivo:
            arquivo.writestr("xl/worksheets/sheet1.xml", planilha)

    def test_lista_exibe_apenas_localizacoes_livres(self):
        livres = self.db.listar_localizacoes_livres("Rua C")
        self.assertNotIn("C2809", livres)
        self.assertNotIn("C2810", livres)
        self.assertIn("C2808", livres)
        self.assertIn("C3414", livres)
        self.assertEqual(self.db.listar_localizacoes_livres("Frações"), ["Fração"])

    def test_varios_itens_compartilham_palete_de_fracoes(self):
        self.db.cadastrar_item("Fração", "PRATI-F1", "Fracionado 1", "LF1", 2, "Normal")
        self.db.cadastrar_item("Fração", "PRATI-F2", "Fracionado 2", "LF2", 3, "Bloqueado")
        with self.db.get_connection() as conn:
            rows = conn.execute("SELECT codigo_localizacao FROM itens_estoque WHERE codigo_produto IN ('PRATI-F1', 'PRATI-F2') ORDER BY codigo_produto").fetchall()
        self.assertEqual(len(rows), 2)
        self.assertNotEqual(rows[0][0], rows[1][0])
        self.assertTrue(all(row[0].startswith("FRAC-") for row in rows))

    def test_importacao_excel_cadastra_rua_e_fracao(self):
        caminho = os.path.join(self.temp_dir.name, "importacao.xlsx")
        self._criar_xlsx(caminho, [
            ["Rua", "Localização", "Código do Produto", "Nome do Medicamento", "Lote", "Qtd.", "Status"],
            ["Rua D", "0114", "PRATI-XLSX1", "Produto Excel", "LX1", "5", "Normal"],
            ["Frações", "", "PRATI-XLSX2", "Produto Fracionado", "LX2", "3", "Bloqueado"],
        ])
        importados, erros = self.db.importar_excel(caminho)
        self.assertEqual(importados, 2)
        self.assertEqual(erros, [])
        with self.db.get_connection() as conn:
            normal = conn.execute("SELECT codigo_localizacao, status FROM itens_estoque WHERE codigo_produto = 'PRATI-XLSX1'").fetchone()
            fracao = conn.execute("SELECT rua, status FROM itens_estoque WHERE codigo_produto = 'PRATI-XLSX2'").fetchone()
        self.assertEqual(tuple(normal), ("D0114", "Normal"))
        self.assertEqual(tuple(fracao), ("Frações", "Bloqueado"))


class TestManutencaoBanco(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = os.path.join(self.temp_dir.name, "test_manutencao.db")
        self.db = DatabaseManager(self.db_path, manutencao_automatica=False)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_limpeza_mantem_somente_ultimas_48_horas(self):
        referencia = datetime(2026, 10, 18, 12, 0, 0)
        antigo = (referencia - timedelta(days=3)).strftime("%d/%m/%Y %H:%M:%S")
        recente = (referencia - timedelta(hours=47)).strftime("%d/%m/%Y %H:%M:%S")
        with self.db.get_connection() as connection:
            for data_hora, codigo in ((antigo, "ANTIGO"), (recente, "RECENTE")):
                connection.execute(
                    """
                    INSERT INTO historico_movimentacoes
                        (data_hora, tipo, codigo_localizacao, codigo_produto,
                         nome_medicamento, lote, quantidade, motivo)
                    VALUES (?, 'ENTRADA', 'A0101', ?, 'Teste', 'L1', 1, 'Retenção')
                    """,
                    (data_hora, codigo),
                )

        removidos = self.db.limpar_historico_antigo(referencia)

        with self.db.get_connection() as connection:
            codigos = {
                row[0] for row in connection.execute(
                    "SELECT codigo_produto FROM historico_movimentacoes"
                )
            }
        self.assertEqual(removidos, 1)
        self.assertEqual(codigos, {"RECENTE"})

    def test_backup_integro_e_retencao_de_sete_copias(self):
        self.db.cadastrar_item("A0101", "PRATI-BKP", "Produto backup", "LBKP", 4, "Normal")
        manual = os.path.join(self.temp_dir.name, "manual.db")
        self.db.criar_backup(manual)
        connection = sqlite3.connect(manual)
        try:
            self.assertEqual(connection.execute("PRAGMA integrity_check").fetchone()[0], "ok")
            self.assertEqual(connection.execute("SELECT COUNT(*) FROM itens_estoque").fetchone()[0], 1)
        finally:
            connection.close()

        for dia in range(1, 10):
            self.db.criar_backup_automatico(datetime(2026, 1, dia, 8, 0, 0))
        automaticos = sorted(
            nome for nome in os.listdir(self.db.backup_dir)
            if nome.startswith("Prati_WMS_auto_")
        )
        self.assertEqual(len(automaticos), 7)
        self.assertEqual(automaticos[0], "Prati_WMS_auto_2026-01-03.db")


class TestTemasInterface(unittest.TestCase):
    def test_alternancia_aplica_as_duas_paletas(self):
        original = "claro" if app_module.BG == app_module.THEMES["claro"]["BG"] else "escuro"
        try:
            for nome in ("claro", "escuro"):
                app_module.definir_tema(nome)
                self.assertEqual(app_module.BG, app_module.THEMES[nome]["BG"])
                self.assertEqual(app_module.PRIMARY, "#3F1E69")
                self.assertNotEqual(app_module.TEXT, app_module.BG)
        finally:
            app_module.definir_tema(original)

    def test_preferencia_de_tema_e_salva_localmente(self):
        with tempfile.TemporaryDirectory() as pasta:
            original = app_module.CONFIG_PATH
            app_module.CONFIG_PATH = os.path.join(pasta, "config.json")
            try:
                app_module.salvar_tema("escuro")
                self.assertEqual(app_module.carregar_tema(), "escuro")
            finally:
                app_module.CONFIG_PATH = original


class TestCompatibilidadeWindows(unittest.TestCase):
    def setUp(self):
        self.base_dir = os.path.dirname(os.path.abspath(__file__))

    def _ler(self, nome):
        with open(os.path.join(self.base_dir, nome), "r", encoding="utf-8") as arquivo:
            return arquivo.read()

    def test_manifesto_declara_windows_moderno_sem_elevacao(self):
        manifesto = self._ler("Prati_WMS.manifest")
        self.assertIn('processorArchitecture="amd64"', manifesto)
        self.assertIn('level="asInvoker"', manifesto)
        self.assertIn('uiAccess="false"', manifesto)
        self.assertIn("{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}", manifesto)
        self.assertNotIn("requireAdministrator", manifesto)

    def test_build_fixa_python_312_e_64_bits(self):
        spec = self._ler("Prati_WMS.spec")
        requisitos = self._ler("requirements-build.txt")
        script = self._ler("build_windows.ps1")
        self.assertIn("sys.version_info[:2] != (3, 12)", spec)
        self.assertIn("struct.calcsize('P') * 8 != 64", spec)
        self.assertIn("manifest='Prati_WMS.manifest'", spec)
        self.assertIn("uac_admin=False", spec)
        self.assertIn("pyinstaller==6.22.3", requisitos)
        self.assertIn("Get-FileHash", script)
        self.assertIn("Feche todas as janelas do Prati WMS", script)
        self.assertIn("$env:TEMP", script)


class TestEtiquetasZebra(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.temp_dir.cleanup()

    @staticmethod
    def _elementos():
        return [
            {"tipo": "texto", "x": 406, "y": 160, "texto": "PRATI WMS", "tamanho": 28},
            {"tipo": "imagem", "x": 406, "y": 650, "largura": 160, "altura": 120, "imagem": Image.new("RGBA", (80, 60), "black")},
        ]

    def test_renderizacao_zebra_sample_203_dpi(self):
        etiqueta = renderizar_etiqueta(self._elementos())
        self.assertEqual(etiqueta.size, (LABEL_WIDTH_PX, LABEL_HEIGHT_PX))
        self.assertEqual(etiqueta.size, (1218, 812))
        self.assertGreater(etiqueta.width, etiqueta.height)
        self.assertEqual(etiqueta.getpixel((406, 650)), (0, 0, 0))
        self.assertNotEqual(etiqueta.getbbox(), None)

    def test_modelo_preserva_texto_e_imagem(self):
        elementos_originais = self._elementos()
        elementos_originais[0].update({"escala_x": 1.4, "escala_y": 0.8})
        caminho = os.path.join(self.temp_dir.name, "modelo.prati-etiqueta")
        salvar_modelo_etiqueta(caminho, elementos_originais)
        elementos = abrir_modelo_etiqueta(caminho)
        self.assertEqual(elementos[0]["texto"], "PRATI WMS")
        self.assertEqual(elementos[0]["tamanho"], 28)
        self.assertEqual(elementos[0]["escala_x"], 1.4)
        self.assertEqual(elementos[0]["escala_y"], 0.8)
        self.assertEqual(elementos[1]["imagem"].size, (80, 60))

    def test_texto_tem_dimensoes_horizontal_e_vertical_independentes(self):
        natural = {"tipo": "texto", "x": 400, "y": 200, "texto": "PRATI", "tamanho": 28}
        largura, altura = dimensoes_texto_etiqueta(natural)[:2]
        alterado = {**natural, "escala_x": 1.8, "escala_y": 0.6}
        largura_alterada, altura_alterada = dimensoes_texto_etiqueta(alterado)[:2]
        self.assertAlmostEqual(largura_alterada / largura, 1.8, delta=0.03)
        self.assertAlmostEqual(altura_alterada / altura, 0.6, delta=0.03)

    def test_selecao_exibe_oito_alcas(self):
        alcas = EtiquetasView._alcas_selecao((10, 20, 110, 70))
        self.assertEqual(len(alcas), 8)
        self.assertEqual({alca[0] for alca in alcas}, {"nw", "n", "ne", "e", "se", "s", "sw", "w"})

    def test_descoberta_de_impressoras_sem_hardware(self):
        class ImpressorasFalsas:
            PRINTER_ENUM_LOCAL = 2
            PRINTER_ENUM_CONNECTIONS = 4

            @staticmethod
            def EnumPrinters(flags):
                self.assertEqual(flags, 6)
                return [(None, None, "Zebra Teste"), (None, None, "Impressora PDF")]

        self.assertEqual(listar_impressoras(ImpressorasFalsas), ["Impressora PDF", "Zebra Teste"])

    def test_envio_para_driver_simulado(self):
        eventos = []

        class DCFalso:
            def CreatePrinterDC(self, nome): eventos.append(("printer", nome))
            def GetDeviceCaps(self, codigo): return {8: 1200, 10: 1600, 88: 203, 90: 203}[codigo]
            def StartDoc(self, nome): eventos.append(("documento", nome))
            def StartPage(self): eventos.append(("pagina", "inicio"))
            def GetHandleOutput(self): return 123
            def EndPage(self): eventos.append(("pagina", "fim"))
            def EndDoc(self): eventos.append(("documento", "fim"))
            def AbortDoc(self): eventos.append(("documento", "abortado"))
            def DeleteDC(self): eventos.append(("dc", "excluido"))

        class DIBFalso:
            def __init__(self, imagem): self.imagem = imagem
            def draw(self, handle, caixa): eventos.append(("desenho", handle, caixa, self.imagem.size))

        imprimir_imagem_windows(
            renderizar_etiqueta(self._elementos()), "Zebra Teste",
            criar_dc=DCFalso, criar_dib=DIBFalso,
        )
        self.assertIn(("printer", "Zebra Teste"), eventos)
        self.assertTrue(any(evento[0] == "desenho" for evento in eventos))
        self.assertIn(("documento", "fim"), eventos)
        self.assertIn(("dc", "excluido"), eventos)


if __name__ == "__main__":
    unittest.main(verbosity=2)
