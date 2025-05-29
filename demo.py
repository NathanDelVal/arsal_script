import os
import sys
import xlwings as xw
import datetime
import tempfile
import shutil

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
        
        # sheet = dest_wb.sheets['08-RST_ANL_VRF']

        # # Column to copy (e.g., column B = 2)
        # source_col = 2

        # # Get the number of rows in use
        # n_rows = sheet.used_range.last_cell.row

        # # Get the last used column and calculate the new column index
        # last_col = sheet.used_range.last_cell.column
        # target_col = last_col + 1

        # # Copy values from source to new column
        # source_range = sheet.range((1, source_col), (n_rows, source_col))
        # target_range = sheet.range((1, target_col), (n_rows, target_col))

        # # Copy values and formulas
        # target_range.value = source_range.value
        
        dest_wb.save(copy_path)

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