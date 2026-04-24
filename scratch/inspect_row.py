import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')
wb = openpyxl.load_workbook("D:/git/dev/dbi_patcher/data/dictionary.xlsx", data_only=True)
ws = wb["Translations"]
for r in range(270, 280):
    print(f"Row {r}: {ws.cell(r, 1).value}")
