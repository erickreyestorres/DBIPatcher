import openpyxl
from pathlib import Path

path = Path(r'data\dictionary.xlsx')
wb = openpyxl.load_workbook(path)
ws = wb['Translations']
val = ws.cell(35, 1).value

with open('debug_xlsx.txt', 'w', encoding='utf-8') as f:
    f.write(f"Row 35 Value: {repr(val)}\n")
    if val:
        f.write(f"Hex: {val.encode('utf-8').hex()}\n")
print("Done, check debug_xlsx.txt")
