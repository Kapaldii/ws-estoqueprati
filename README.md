# Prati WMS

Aplicação desktop local para gestão de estoque farmacêutico, desenvolvida em Python com Tkinter e SQLite.

## Executar

- Dê duplo clique em `Prati_WMS.exe`; não é necessário instalar nem executar como administrador.
- Para executar pelo código-fonte, use `python app.py`.
- O banco `estoque_medicamentos.db` fica na mesma pasta da aplicação.

## Interface

- **Início:** indicadores essenciais e movimentações recentes.
- **Consultar:** busca instantânea, filtro por status e ações de editar/remover em cada linha.
- **Cadastrar:** inclusão de itens em posições livres, sem campos de movimentação.
- **Histórico:** auditoria de entradas, alterações e remoções, com exportação CSV.

Os únicos status disponíveis são `Normal` e `Bloqueado`.

### Importar do Excel

Na aba **Cadastrar**, use **Importar Excel** e selecione um arquivo `.xlsx`. A primeira linha deve conter os cabeçalhos:

| Rua | Localização | Código | Medicamento | Lote | Quantidade | Status |
| --- | --- | --- | --- | --- | --- | --- |
| Rua A | A0102 | PRATI-500 | Exemplo 500mg | L26001 | 20 | Normal |
| Frações |  | PRATI-501 | Exemplo fracionado | L26002 | 5 | Bloqueado |

`Status` é opcional e assume `Normal`. Para itens fracionados, informe apenas `Frações` na coluna Rua; vários produtos podem compartilhar esse palete.

## Endereçamento

- Ruas A a D: letra, bloco e posição, como `C2809`.
- Frações: um único palete compartilhado, apresentado como `Fração` na interface.
- Cada rua possui blocos de `01` a `34`, com 14 posições por bloco.

## Gerar o executável

```powershell
pyinstaller --onefile --noconsole --name "Prati_WMS" app.py
```

O executável portátil é gerado na pasta `dist`.

## Testes

```powershell
python -m unittest -v test_wms.py
```
