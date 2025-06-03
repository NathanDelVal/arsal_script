import os
import sys
import xlwings as xw
import datetime
import tempfile
import shutil
import win32com.client as win32
import formulas

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

def index_to_column_letter(index):
    letters = ''
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters

try:
    with xw.App(visible=False) as app:
        src_wb = xw.Book(file_path)
        dest_wb = xw.Book()

        file_name = os.path.basename(file_path)

        for sheet in src_wb.sheets:
            sheet.copy(after=dest_wb.sheets[-1])
        dest_wb.sheets.__delitem__(0)  # remove initial blank sheet
        
        try:
            #variáveis de referência das tabelas
            target_sht = dest_wb.sheets[formulas.panilha_alvo]        
            lista = dest_wb.sheets[formulas.lista_de_referencia]
            rst = dest_wb.sheets[formulas.planilha_de_referencia]
        except KeyError as e:
            print(f"ERRO: A planilha {e} não foi encontrada no arquivo.")
            src_wb.close()
            dest_wb.close()
            sys.exit(1)
            
        # print([v for v in target_sht[0, :].value if v is not None])  #find first row with values
        # next_col_n = index_to_column_letter(last_col_i + 1) #next column letter
        
        lista_de_referencia = dest_wb.sheets[formulas.lista_de_referencia]
        planilha_de_referencia = dest_wb.sheets[formulas.planilha_de_referencia]
        planilha_alvo = dest_wb.sheets[formulas.panilha_alvo]             
        
        headers = planilha_alvo.range((formulas.linha_headers, 1), (formulas.linha_headers, planilha_alvo.used_range.last_cell.column)).value
        
        cidade2 = planilha_alvo.range((1, headers.index("Conta") + 1), (planilha_alvo.used_range.last_cell.row, headers.index("Conta") + 1)).value
        
        cidade2 = formulas.procx([c for c in cidade2[3:] if c is not None], 
                                 lista_de_referencia.range("Q3:Q80").value, 
                                 lista_de_referencia.range("R3:R80").value)
        
        copy_path = os.path.join(
            os.path.dirname(file_path),
            f"{os.path.splitext(file_name)[0]}_COPY_{int(datetime.datetime.timestamp(datetime.datetime.now()))}{os.path.splitext(file_name)[1]}"
        )
    
        # dest_wb.save(copy_path)

        src_wb.close()
        dest_wb.close()
        print("Processamento concluído com sucesso.")

except Exception as e:
    print(f"ERRO DURANTE EXECUÇÃO: {e}")

finally:
    pass
    # print("Limpando arquivos temporários...")
    # clean_temp_folder(tempfile.gettempdir())
    # clean_temp_folder("C:\\Windows\\Temp")