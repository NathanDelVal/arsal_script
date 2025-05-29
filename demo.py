import os
import sys
import xlwings as xw
import datetime
import tempfile
import shutil
import win32com.client as win32

file_path = sys.argv[1]

# Ensure the file exists
if not os.path.isfile(file_path):
    print(f"O arquivo {file_path} não existe.")
    sys.exit(1)
    
# def add_column():

def clean_temp_folder(path):
    for entry in os.listdir(path):
        entry_path = os.path.join(path, entry)
        try:
            if os.path.isfile(entry_path) or os.path.islink(entry_path):
                os.unlink(entry_path)
            elif os.path.isdir(entry_path):
                shutil.rmtree(entry_path)
        except Exception:
            pass  # skip files in use or protected

try:
    with xw.App(visible=False) as app:
        src_wb = xw.Book(file_path)
        dest_wb = xw.Book()

        file_name = os.path.basename(file_path)

        for sheet in src_wb.sheets:
            sheet.copy(after=dest_wb.sheets[-1])
        dest_wb.sheets.__delitem__(0)  # remove initial blank sheet

        # 🔧 Optional: External reference cleanup (commented)
        # for sheet in dest_wb.sheets:
        #     formulas = [list(row) for row in sheet.used_range.formula]
        #     for r in range(len(formulas)):
        #         for c in range(len(formulas[r])):
        #             if '[arsal.xlsx]' in formulas[r][c]:
        #                 formulas[r][c] = formulas[r][c].replace('[arsal.xlsx]', '')
        #     sheet.used_range.formula = tuple(tuple(row) for row in formulas)

        copy_path = os.path.join(
            os.path.dirname(file_path),
            f"{os.path.splitext(file_name)[0]}_COPY_{int(datetime.datetime.timestamp(datetime.datetime.now()))}{os.path.splitext(file_name)[1]}"
        )

        # print(DumpCOMObject())

        # dest_wb.save(copy_path)

        src_wb.close()
        dest_wb.close()
        print("Processamento concluído com sucesso.")

except Exception as e:
    print(f"Erro durante execução: {e}")

finally:
    pass
    # print("Limpando arquivos temporários...")
    # clean_temp_folder(tempfile.gettempdir())
    # clean_temp_folder("C:\\Windows\\Temp")