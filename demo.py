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
            lista_de_referencia = dest_wb.sheets[formulas.lista_de_referencia]
            planilha_de_referencia = dest_wb.sheets[formulas.planilha_de_referencia]
            planilha_alvo = dest_wb.sheets[formulas.planilha_alvo]
        except KeyError as e:
            print(f"ERRO AO ACESSAR O ARQUIVO: {e}")
            src_wb.close()
            dest_wb.close()
            sys.exit(1)
        
        last_row = planilha_alvo.used_range.last_cell.row
        last_col = planilha_alvo.used_range.last_cell.column
         
        headers = planilha_alvo.range((formulas.linha_headers, 1), (formulas.linha_headers, last_col)).value
        
        temp_c = planilha_alvo.range((1, headers.index("Conta") + 1), (last_row, headers.index("Conta") + 1)).value
        temp_c = formulas.procx([c for c in temp_c[formulas.linha_headers:] if c is not None], 
                                 lista_de_referencia.range(f"Q{formulas.linha_headers}:Q80").value, 
                                 lista_de_referencia.range(f"R{formulas.linha_headers}:R80").value)
        planilha_alvo[formulas.linha_headers - 1:, last_col].options(transpose=True).value = ["Cidade2"] + temp_c
        last_col += 1

        temp_c = planilha_alvo.range(
            (1, headers.index("Valor") + 1), (last_row, headers.index("Valor") + 1)
            ).value 
        temp_c = [0 if c == "N.D" else c for c in temp_c[formulas.linha_headers:]]
        planilha_alvo[formulas.linha_headers - 1:, last_col].options(transpose=True).value = ["Resultado independente"] + temp_c
        last_col += 1

        temp_c = planilha_alvo.range((1, headers.index("Análise") + 1), (last_row, headers.index("Análise") + 1)).value
        temp_c = formulas.procx([c for c in temp_c[formulas.linha_headers:] if c is not None], 
                                 lista_de_referencia.range(f"T{formulas.linha_headers}:T34").value, 
                                 lista_de_referencia.range(f"V{formulas.linha_headers}:V34").value)
        planilha_alvo[formulas.linha_headers - 1:, last_col].options(transpose=True).value = ["Possui Frequência na Portaria?"] + temp_c
        last_col += 1
        
        temp_c = planilha_alvo.range((1, headers.index("Análise") + 1), (last_row, headers.index("Análise") + 1)).value
        temp_c = formulas.procx([c for c in temp_c[formulas.linha_headers:] if c is not None], 
                                 lista_de_referencia.range(f"T{formulas.linha_headers}:T34").value, 
                                 lista_de_referencia.range(f"W{formulas.linha_headers}:W34").value)
        planilha_alvo[formulas.linha_headers - 1:, last_col].options(transpose=True).value = ["Possui VMP na Portaria?"] + temp_c
        last_col += 1
        
        temp_c = planilha_alvo.range((1, headers.index("Análise") + 1), (last_row, headers.index("Análise") + 1)).value
        temp_c = formulas.procx([c for c in temp_c[formulas.linha_headers:] if c is not None], 
                                 lista_de_referencia.range(f"T{formulas.linha_headers}:T34").value, 
                                 lista_de_referencia.range(f"U{formulas.linha_headers}:U34").value)
        planilha_alvo[formulas.linha_headers - 1:, last_col].options(transpose=True).value = ["Tipo Análise"] + temp_c
        last_col += 1

        planilha_alvo[formulas.linha_headers - 1:, last_col].options(transpose=True).value = ["Responsável Análise"] 
        last_col += 1
        
        planilha_alvo[formulas.linha_headers - 1:, last_col].options(transpose=True).value = ["ID_Conc"] 
        last_col += 1
        
        planilha_alvo[formulas.linha_headers - 1:, last_col].options(transpose=True).value = ["Justificativa  Concessionária"] 
        last_col += 1
        
        planilha_alvo[formulas.linha_headers - 1:, last_col].options(transpose=True).value = ["ID_VI"] 
        last_col += 1
        
        planilha_alvo[formulas.linha_headers - 1:, last_col].options(transpose=True).value = ["Justificativa VI"] 
        last_col += 1
        
        planilha_alvo[formulas.linha_headers - 1:, last_col].options(transpose=True).value = ["Acatar"] 
        last_col += 1
        
        planilha_alvo[formulas.linha_headers - 1:, last_col].options(transpose=True).value = ["Etapa do Processo"] 
        last_col += 1
        
        planilha_alvo[formulas.linha_headers - 1:, last_col].options(transpose=True).value = ["Observações"] 
        last_col += 1
        
        # temp_c = temp_c = planilha_alvo.range((1, headers.index("Formula para calculo3") + 1), (last_row, headers.index("Formula para calculo3") + 1)).value
        # temp_c = ["Sim" if c == 1 else "Não" if c == 0 else "" for c in temp_c[formulas.linha_headers:]]
        # planilha_alvo[formulas.linha_headers - 1:, last_col].options(transpose=True).value = ["Justificativa aceita?"] + temp_c
        # last_col += 1        
        
        copy_path = os.path.join(
            os.path.dirname(file_path),
            f"{os.path.splitext(file_name)[0]}_COPY_{int(datetime.datetime.timestamp(datetime.datetime.now()))}{os.path.splitext(file_name)[1]}"
        )
    
        dest_wb.save(copy_path)

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