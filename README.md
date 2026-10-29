# planilha_code

Planilha desktop parecida com o Excel, escrita do zero em Python (tkinter + openpyxl), sem nenhum código do Microsoft Excel.

## Como rodar

```
pip install openpyxl
python planilha/planilha.py
```

## Testes

```
python planilha/test_planilha.py
```

## Como gerar o .exe

```
pip install pyinstaller
cd planilha
python -m PyInstaller --noconfirm --onefile --windowed --name planilha_code --icon planilha.ico --add-data "planilha.ico;." planilha.py
```

O executável pronto está na página de [Releases](https://github.com/mentorbionline/rea_excel_claude_code/releases).

Recursos e limitações: veja [planilha/LEIAME.md](planilha/LEIAME.md).
