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
- **Histórico:** movimentações dos últimos 2 dias, com exportação CSV.
- **Etiquetas:** editor Zebra SAMPLE horizontal 6×4 para texto em Arial Black ou imagem colada, com impressão pelo driver do Windows.

Os únicos status disponíveis são `Normal` e `Bloqueado`.

## Dados e backups

- Bancos novos começam vazios, sem medicamentos demonstrativos.
- Ao abrir o programa, históricos com mais de 2 dias são removidos automaticamente.
- Um backup automático é criado por dia na pasta `Backups`, mantendo as 7 cópias mais recentes.
- O botão **Fazer backup**, na tela Início, permite salvar uma cópia adicional onde desejar.
- A integridade do banco é verificada durante a abertura. Se houver problema ou falta de permissão, o programa mostra uma mensagem clara.

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
pyinstaller --clean --onefile --noconsole --name "Prati_WMS" app.py
```

O executável portátil é gerado na pasta `dist`.

Para levar a outro computador, copie `Prati_WMS.exe` e `estoque_medicamentos.db` para a mesma pasta. Não coloque o programa em `Program Files`; use uma pasta dentro de `Documentos` para que o banco possa ser atualizado sem permissão de administrador.

## Etiquetas e impressão

- Formato fixo horizontal: `152,4 × 101,6 mm` (6×4 polegadas), renderizado em `203 dpi` (`1218 × 812` pontos).
- Arraste **Adicionar texto** da barra lateral até a etiqueta; o conteúdo é solicitado após soltar.
- O texto usa sempre Arial Black. Arraste as alças roxas para alterar o tamanho diretamente com o cursor e use duplo clique para editar o conteúdo.
- Imagens são coladas exclusivamente com `Ctrl+V`, sem botão dedicado.
- Os elementos podem ser selecionados, movidos com o cursor e excluídos com `Delete`.
- Modelos são abertos e salvos no formato `.prati-etiqueta`.
- A lista apresenta qualquer impressora instalada ou compartilhada no Windows. A impressão usa o driver selecionado, inclusive em impressoras não Zebra.

Sem uma impressora física, a renderização, descoberta de dispositivos e envio ao driver são validados por testes automatizados com uma impressora simulada. Antes do uso em produção, configure no driver o papel 6×4 em orientação horizontal e faça uma impressão de calibração no equipamento real.

## Testes

```powershell
python -m unittest -v test_wms.py
```
