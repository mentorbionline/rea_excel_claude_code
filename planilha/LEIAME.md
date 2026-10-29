# planilha_code

Planilha simples parecida com o Excel, escrita do zero em Python (tkinter + openpyxl).

## Como rodar

1. Instale o Python 3 e o openpyxl: `pip install openpyxl`
2. Na pasta `planilha`, rode: `python planilha.py`
   (ou `python planilha.py arquivo.xlsx` para já abrir um arquivo)
3. Testes: `python test_planilha.py`
4. Arquivo para teste manual (dados fictícios): `exemplo_compatibilidade.xlsx`.
5. Programa pronto (Windows): `dist\planilha_code.exe`, com atalho `planilha_code` na Área de Trabalho.

## Atalhos

| Tecla | Ação |
|---|---|
| Setas, Enter, Tab, Shift+Tab, Page Up/Down, Home | Mover a seleção |
| Shift+setas, Shift+clique, arrastar | Selecionar intervalo |
| Clique no cabeçalho / no canto | Selecionar coluna, linha / tudo (Ctrl+A) |
| Digitar, F2, duplo clique | Editar célula (Esc cancela, Enter confirma) |
| Delete / Backspace | Limpar seleção / limpar e editar |
| Ctrl+C, Ctrl+X, Ctrl+V | Copiar, recortar, colar |
| Ctrl+Z, Ctrl+Y | Desfazer, refazer |
| Ctrl+B, Ctrl+I | Negrito, itálico |
| Ctrl+N, Ctrl+O, Ctrl+S | Novo, abrir, salvar |
| Duplo clique na aba | Renomear planilha |
| Arrastar a borda no cabeçalho da coluna/linha | Mudar largura/altura |
| Duplo clique na borda do cabeçalho da coluna | Ajustar largura ao conteúdo |
| Menu Exibir > Congelar painéis / Descongelar | Fixa as linhas acima e as colunas à esquerda da célula atual |

## O que faz

- Grade de 100 linhas × 26 colunas (cresce ao navegar além), barra de fórmulas editável.
- Várias abas: adicionar (+), renomear, excluir (menu Planilha), alternar.
- Fórmulas: `SOMA/SUM`, `MÉDIA/MEDIA/AVERAGE`, `MÍNIMO/MINIMO/MIN`, `MÁXIMO/MAXIMO/MAX`,
  intervalos (`A1:B5`), separador `;` ou `,`, contas `+ - * /` com parênteses e referências (`$A$1` também).
  Erros: `#ERRO!` (fórmula inválida), `#DIV/0!`, `#CIRC!` (referência circular).
- Números: use vírgula como decimal (`1,5` ou `1.234,56`). Dentro de fórmulas, o decimal é ponto (`=A1*1.5`).
- Formatos: negrito, itálico, Geral, Número (`1.234,56`), Moeda (`R$ 1.234,56`), Porcentagem (`12,50%`),
  Data (`26/01/1991`), com botões **− decimais / + decimais** (0 a 4 casas).
- Datas: digite `26/01/1991` (ou `26/01/1991 08:30`). Aparecem como `dd/mm/aaaa` e são gravadas como data
  de verdade no Excel. Nas fórmulas contam como número (nº de série do Excel): `MÍNIMO/MÁXIMO` de datas
  mostram a data; `=A1+30` mostra o número de série (escolha o formato Data para ver como data).
- Cor do texto e cor de fundo (botões com "Escolher cor..." e "Sem cor"), alinhamento Esquerda/Centro/Direita
  e Quebrar texto (as linhas quebradas são cortadas na altura da célula). Tudo entra no desfazer.
- Larguras de coluna, alturas de linha e painéis congelados: lidos do arquivo, editáveis e salvos.
- Copiar/colar intervalos, com ajuste de referências relativas (recortar não ajusta). Copia texto
  separado por tabulação para a área de transferência e aceita colar do Excel.
- Desfazer/refazer edições, colar, recortar, limpar e formatação.
- Abrir/salvar `.xlsx`. Ao salvar, as funções vão em inglês (`=SUM(A1:A3)`) para o Excel/LibreOffice entenderem.
- **Salvar preserva o arquivo original**: ao salvar (ou "Salvar como") uma planilha aberta de um `.xlsx`,
  o programa reabre o arquivo original e grava por cima dele só o que mudou aqui (valores, fórmulas,
  formatos, larguras, painéis, abas novas/renomeadas/excluídas). O resto continua como estava: bordas,
  células mescladas, filtros, formatação condicional, validação de dados, tabelas, comentários,
  hiperlinks, cores de tema, formatos de número desconhecidos, proteção, nomes definidos.
  Se o original sumiu ou não abre, grava do zero e avisa.
- Ao abrir, datas gravadas como TEXTO no formato `aaaa-mm-dd` (ou `aaaa-mm-dd hh:mm[:ss]`, como
  `1989-06-10 00:00:00`) viram data automaticamente: aparecem como `dd/mm/aaaa` (com `hh:mm` se a hora
  não for meia-noite) e, ao salvar, vão para o arquivo como data de verdade (`dd/mm/yyyy`), mesmo sem
  edição; fonte, preenchimento, bordas e alinhamento ficam como estavam. Datas inválidas
  (`2023-02-30`) e texto `dd/mm/aaaa` continuam texto.
- Ao abrir, mostra um relatório: ajustes feitos ao abrir (datas em texto convertidas), o que fica
  preservado mas não aparece aqui, e o que será perdido ao salvar.

## Limitações conhecidas

- Só as 4 funções acima. Outras funções (PROCV, SE, etc.) e referências a outras abas (`Plan2!A1`)
  ficam como texto na célula; no arquivo continuam como fórmula.
- Não aparecem aqui (mas são preservados): bordas, mesclagens, filtros, formatação condicional,
  validação, comentários, cores de tema, preenchimentos com padrão, sublinhado/tachado, alinhamento
  justificado/recuo/texto girado. Tipo e tamanho de fonte não são lidos.
- Em células mescladas, só a primeira célula da mesclagem é gravada; editar as outras não tem efeito no arquivo.
- Perdidos ao salvar (limite do openpyxl): gráficos, imagens, formas/caixas de texto, segmentações,
  macros, formatação parcial dentro do texto de uma célula, e extensões avisadas no relatório.
- Horas sem data e booleanos aparecem como texto (ficam intactos no arquivo se não forem editados).
  Datas só no formato `dd/mm/aaaa` ao digitar (ano com 4 dígitos).
- Renomear uma aba não corrige fórmulas de outras abas que citam o nome antigo.
- Excluir/renomear abas, mudar largura/altura e congelar painéis não entram no desfazer.
- Arredondamento de ,5 na exibição segue o Python (1234,5 com 0 casas mostra 1.234), não o Excel.
- Sem texto entre aspas em fórmulas, sem operador `^`, sem comparação (`>`, `=`).
- Recalcula a aba inteira a cada mudança (rápido para planilhas pequenas/médias).
