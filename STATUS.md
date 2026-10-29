# STATUS — rea_excel_claude_code

## Decisões
- 29/10/2026 — Pacote de compatibilidade .xlsx: salvar por cima do original (preserva o que o app não mostra), datas, formatos de número com casas decimais, larguras/alturas, painéis congelados, cores de texto/fundo, alinhamento e quebra de texto. Pedido pelo Lula.
- 29/10/2026 — Projeto vai para o GitHub público mentorbionline/rea_excel_claude_code (repositório criado pelo Lula). Fora do repositório: pasta rea\ (projeto de terceiros) e o .exe (vai como Release, anexado pelo Lula pelo site).
- 29/10/2026 — Nome do aplicativo trocado para planilha_code (título da janela, .exe e atalho). Pedido pelo Lula; a pasta do projeto continua planilha\.
- 29/10/2026 — Não descompilar nem copiar código do Excel com o REA. Motivo: a licença da Microsoft proíbe engenharia reversa do Excel, e o próprio REA (DISCLAIMER.md) não apoia "copiar software proprietário ou reproduzir seu código". Decidido por: Claude; o Lula aceitou o caminho alternativo (planilha própria) em 29/10/2026.
- Caminho sugerido: criar uma planilha própria parecida com o Excel (células, fórmulas, abrir/salvar .xlsx), usando documentação pública (formato OOXML/ECMA-376 e lista pública de funções).

- 29/10/2026 — Planilha própria desktop em Python (tkinter + openpyxl), só stdlib + openpyxl, 2 arquivos (modelo + UI em planilha.py). Pedido pelo Lula via Claude (agente principal).

## Feito (conferido)
- 29/10/2026 — Pacote de compatibilidade pronto: 15 testes passam; planilha_code.exe refeito; planilha\exemplo_compatibilidade.xlsx (dados fictícios) aberto no .exe com cabeçalho colorido, datas dd/mm/aaaa, R$ e relatório só com Bordas, Células mescladas e Filtros (preservados).
- 29/10/2026 — Ícone verde aparece na janela e na barra de tarefas (conferido por captura de tela). App renomeado: dist\planilha_code.exe + atalho C:\Users\USER\Desktop\planilha_code.lnk; aberto por ele, título 'Sem título - planilha_code'. Planilha.exe e Planilha.lnk antigos removidos. 11 testes passam.
- 29/10/2026 — Executável criado (depois renomeado para planilha_code.exe): planilha\dist\Planilha.exe (32,8 MB, PyInstaller --onefile, ícone planilha\planilha.ico). Atalho C:\Users\USER\Desktop\Planilha.lnk criado; aberto por ele, a janela 'Sem título - Planilha' subiu.
- 29/10/2026 — planilha\planilha.py, planilha\test_planilha.py e planilha\LEIAME.md criados. 11 testes passam (fórmulas, recálculo, formatos pt-BR, desfazer/refazer, copiar/colar, ida e volta .xlsx, relatório de compatibilidade, abertura da janela). Janela conferida por captura de tela.
- Repositório morluto/rea clonado em H:\dev\work\rea_excel_claude_code\rea (6.787 arquivos).
- REA = "Reverse Engineer Anything": servidor MCP + CLI (Node.js 22+) para analisar binários, apps Electron/JS, .NET e sites. Análise nativa exige Hopper, Ghidra, IDA ou Kuna.
- Excel localizado: C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE — versão 16.0.19127.20860 (Microsoft 365), 70 MB, binário nativo.

## Falta
1. Claude: enviar o código ao GitHub (mentorbionline/rea_excel_claude_code).
2. Lula: anexar planilha\dist\planilha_code.exe numa Release no site do GitHub.
3. Lula: testar com o arquivo real de funcionários, salvar e abrir no Excel para conferir.
4. Ainda não feito: mostrar bordas, células mescladas e filtros (hoje só são preservados); desfazer de largura/congelar/abas.

## Fatos
- PyInstaller 6.22.3 instalado (pip --user) em 29/10/2026. Para refazer o .exe depois de mudar o código: python -m PyInstaller --noconfirm --onefile --windowed --name planilha_code --icon planilha.ico --add-data "planilha.ico;." --distpath dist planilha.py (dentro de planilha\).
- O atalho aponta para planilha\dist\planilha_code.exe: se mover a pasta do projeto, o atalho quebra.
- O Excel é nativo (não é Electron/.NET); o REA só mostraria pseudocódigo/assembly, não o código-fonte.
- gh (GitHub CLI) não está instalado neste PC; git funciona.
- Créditos do Firecrawl esgotados em 29/10/2026.
- A planilha só entende SOMA, MÉDIA, MÍNIMO, MÁXIMO; outras fórmulas do .xlsx ficam como texto e voltam como fórmula ao salvar.
