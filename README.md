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

## Endereçamento

- Ruas A a D: letra, bloco e posição, como `C2809`.
- Frações: posições de `FRAC-01` a `FRAC-24`.
- Cada bloco possui 14 posições.

## Gerar o executável

```powershell
pyinstaller --onefile --noconsole --name "Prati_WMS" app.py
```

O executável portátil é gerado na pasta `dist`.

## Testes

```powershell
python -m unittest -v test_wms.py
```
