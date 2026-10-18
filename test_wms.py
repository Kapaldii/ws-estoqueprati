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
from datetime import datetime
from xml.sax.saxutils import escape

# Importa a lógica de negócio do sistema principal
from app import parse_localizacao, DatabaseManager


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

    def test_criacao_e_seed_inicial(self):
        """Verifica se tabelas, índices e carga inicial de demonstração são gerados."""
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT COUNT(*) FROM itens_estoque")
            total_itens = c.fetchone()[0]
            self.assertGreater(total_itens, 0, "O banco deve conter dados iniciais.")

            c.execute("SELECT COUNT(*) FROM historico_movimentacoes")
            total_logs = c.fetchone()[0]
            self.assertEqual(total_itens, total_logs, "Cada item inicial deve ter um registro de ENTRADA no log.")

    def test_cadastro_item_com_log_auditoria(self):
        """Cadastra um novo medicamento e valida o registro correspondente no log."""
        agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("""
                INSERT INTO itens_estoque (rua, bloco, posicao, codigo_localizacao, codigo_produto, nome_medicamento, lote, quantidade)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, ("Rua A", 8, 12, "A0812", "PRATI-999", "Cetoprofeno 100mg", "L26001A", 50))

            c.execute("""
                INSERT INTO historico_movimentacoes (data_hora, tipo, codigo_localizacao, codigo_produto, nome_medicamento, lote, quantidade, motivo)
                VALUES (?, 'ENTRADA', ?, ?, ?, ?, ?, 'Cadastro de novo lote')
            """, (agora, "A0812", "PRATI-999", "Cetoprofeno 100mg", "L26001A", 50))

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
        agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("UPDATE itens_estoque SET quantidade = 250 WHERE codigo_localizacao = 'C2809'")
            c.execute("""
                INSERT INTO historico_movimentacoes (data_hora, tipo, codigo_localizacao, codigo_produto, nome_medicamento, lote, quantidade, motivo)
                VALUES (?, 'ALTERAÇÃO', 'C2809', 'PRATI-101', 'Dipirona 500mg/mL Gotas', 'L24089A', 250, 'Qtd alterada de 140 para 250')
            """, (agora,))

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
        agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        with self.db.get_connection() as conn:
            c = conn.cursor()
            # Pega dados antes de remover
            c.execute("SELECT codigo_produto, nome_medicamento, lote, quantidade FROM itens_estoque WHERE codigo_localizacao = 'A0201'")
            item = c.fetchone()

            # Remove da tabela principal
            c.execute("DELETE FROM itens_estoque WHERE codigo_localizacao = 'A0201'")

            # Registra baixa no log
            c.execute("""
                INSERT INTO historico_movimentacoes (data_hora, tipo, codigo_localizacao, codigo_produto, nome_medicamento, lote, quantidade, motivo)
                VALUES (?, 'BAIXA / REMOÇÃO', 'A0201', ?, ?, ?, ?, 'Saída por expedição')
            """, (agora, item[0], item[1], item[2], item[3]))

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


if __name__ == "__main__":
    unittest.main(verbosity=2)
