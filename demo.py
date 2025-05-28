import os
import sys
import pathlib
import xlwings as xw
import re
import datetime

file_path = sys.argv[1]

# Ensure the file exists
if os.path.isfile(file_path) is False:
    print(f"O arquivo {file_path} não existe.")
    sys.exit(1)
    
app = None  # declare before try

try:
    app = xw.App(visible=False)
    src_wb = xw.Book(file_path)
    dest_wb = xw.Book()
    
    for sheet in src_wb.sheets:
        sheet.copy(after=dest_wb.sheets[-1])

    # Remove external references
    # for sheet in dest_wb.sheets:
    #     for cell in sheet.used_range:
    #         if isinstance(cell.formula, str) and '[arsal.xlsx]' in cell.formula:
    #             cell.formula = cell.formula.replace('[arsal.xlsx]', '')

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    copy_path = os.path.join(os.path.dirname(file_path), f"{base_name}_COPY_{int(datetime.datetime.timestamp(datetime.datetime.now()))}.xlsx")

except Exception as e:
    print(f"Erro durante execução: {e}")

finally:
    dest_wb.save(copy_path)

    src_wb.close()
    dest_wb.close()

    app.kill()