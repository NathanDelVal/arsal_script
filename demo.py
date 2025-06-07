import os
import sys
import xlwings as xw
import datetime
import tempfile
import shutil
import params
import win32com.client as win32
import formulas
import traceback

file_path = sys.argv[1]

### FUNCTIONS ###
def procx(value, lookup_array, return_array, if_not_found=None):
    value = list(value) if not isinstance(value, list) else value
    v = 0
    try:
        for v in range(len(value)):
            index = lookup_array.index(value[v])
            value[v] = return_array[index]
    except Exception as e:
        # print(f"{e} \n {lookup_array.index(value[v])}")
        value[v] = ""
    return value

def adjust_cols_width(sheet):
    cols = sheet.used_range.columns
    for i, val in enumerate(cols):
        letter = index_to_column_letter(i + 1)
        largest_str = [len(str(v)) for v in val.value]
        largest_str = 14 if max(largest_str) < 14 else max(largest_str)
        sheet[f"{letter}:{letter}"].column_width = largest_str
    
def index_to_column_letter(index):
    letters = ''
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters
#########################################################################
planilha_alvo = "08-RST_ANL_VRF"
lista_de_referencia = "Listas"
planilha_de_referencia = "07-RST_ANL_JTF"
linha_headers = 3
##########################################################################
# Ensure the file exists
if not os.path.isfile(file_path):
    print(f"O arquivo {file_path} não existe.")
    sys.exit(1)

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
            lista_de_referencia = dest_wb.sheets[lista_de_referencia]
            planilha_de_referencia = dest_wb.sheets[planilha_de_referencia]
            planilha_alvo = dest_wb.sheets[planilha_alvo]
        except KeyError as e:
            print(f"ERRO AO ACESSAR O ARQUIVO: {e}")
            src_wb.close()
            dest_wb.close()
            sys.exit(1)
        
        last_row = planilha_alvo.used_range.last_cell.row
        last_col = planilha_alvo.used_range.last_cell.column
         
        headers = planilha_alvo.range((linha_headers, 1), (linha_headers, last_col)).value
        
        conta = planilha_alvo.range((linha_headers + 1, headers.index("Conta") + 1), (last_row, headers.index("Conta") + 1)).value
        
        cidade2 = procx([c for c in conta if c is not None], 
                                 lista_de_referencia.range(f"Q:Q").value, 
                                 lista_de_referencia.range(f"R:R").value)
        
        analise = planilha_alvo.range((linha_headers + 1, headers.index("Análise") + 1), (last_row, headers.index("Análise") + 1)).value
        
        parecer_analise = planilha_alvo.range((linha_headers + 1, headers.index("Parecer Análise") + 1), (last_row, headers.index("Parecer Análise") + 1)).value
        
        planilha_alvo[linha_headers - 1:, last_col].options(transpose=True).value = ["Cidade2"] + cidade2
        planilha_alvo[linha_headers - 1, last_col].color = (0, 0, 139)
        last_col = planilha_alvo.used_range.last_cell.column

        temp_c = planilha_alvo.range(
            (1, headers.index("Valor") + 1), (last_row, headers.index("Valor") + 1)
            ).value 
        temp_c = [0 if c == "N.D" else c for c in temp_c[linha_headers:]]
        planilha_alvo[linha_headers - 1:, last_col].options(transpose=True).value = ["Resultado independente"] + temp_c
        planilha_alvo[linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1

        temp_c = procx([c for c in analise if c is not None], 
                                 lista_de_referencia.range(f"T:T").value, 
                                 lista_de_referencia.range(f"V:V").value)
        planilha_alvo[linha_headers - 1:, last_col].options(transpose=True).value = ["Possui Frequência na Portaria?"]  + temp_c
        planilha_alvo[linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        temp_c = procx([c for c in analise if c is not None], 
                                 lista_de_referencia.range(f"T:T").value, 
                                 lista_de_referencia.range(f"W:W").value)
        planilha_alvo[linha_headers - 1:, last_col].options(transpose=True).value = ["Possui VMP na Portaria?"] + temp_c
        planilha_alvo[linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        temp_c = procx([c for c in analise if c is not None], 
                                 lista_de_referencia.range(f"T:T").value, 
                                 lista_de_referencia.range(f"U:U").value)
        planilha_alvo[linha_headers - 1:, last_col].options(transpose=True).value = ["Tipo Análise"] + temp_c
        planilha_alvo[linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1

        planilha_alvo[linha_headers - 1:, last_col].options(transpose=True).value = ["Responsável Análise"] 
        planilha_alvo[linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        planilha_alvo[linha_headers - 1:, last_col].options(transpose=True).value = ["ID_Conc"] 
        planilha_alvo[linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        planilha_alvo[linha_headers - 1:, last_col].options(transpose=True).value = ["Justificativa  Concessionária"] 
        planilha_alvo[linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        planilha_alvo[linha_headers - 1:, last_col].options(transpose=True).value = ["ID_VI"] 
        planilha_alvo[linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        planilha_alvo[linha_headers - 1:, last_col].options(transpose=True).value = ["Justificativa VI"] 
        planilha_alvo[linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        planilha_alvo[linha_headers - 1:, last_col].options(transpose=True).value = ["Acatar"] 
        planilha_alvo[linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        etapa_processo = planilha_alvo.range((1, headers.index("Ponto Coleta") + 1), (last_row, headers.index("Ponto Coleta") + 1)).value
        etapa_processo = procx([c for c in etapa_processo[linha_headers:] if c is not None], 
                                 lista_de_referencia.range(f"Z:Z").value, 
                                 lista_de_referencia.range(f"AH:AH").value)
        planilha_alvo[linha_headers - 1:, last_col].options(transpose=True).value = ["Etapa do Processo"] + etapa_processo
        planilha_alvo[linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        planilha_alvo[linha_headers - 1:, last_col].options(transpose=True).value = ["Observações"] 
        planilha_alvo[linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        formula1 = [f"{v1}{v2}{v3}{v4}" for v1,v2,v3,v4 in zip(cidade2,conta,analise,parecer_analise)]
        planilha_alvo[linha_headers - 1:, last_col].options(transpose=True).value = ["Fórmula para Cálculo"] + formula1
        planilha_alvo[linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        formula2 = [f"{v1}{v2}{v3}{v4}{v5}" for v1,v2,v3,v4,v5 in zip(cidade2,conta,analise,parecer_analise, etapa_processo)] 
        planilha_alvo[linha_headers - 1:, last_col].options(transpose=True).value = ["Fórmula para Cálculo 2"] + formula2
        planilha_alvo[linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        planilha_alvo[linha_headers - 1:, last_col].options(transpose=True).value = ["Fórmula para Cálculo 3"] 
        planilha_alvo[linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        adjust_cols_width(planilha_alvo)

        copy_path = os.path.join(
            os.path.dirname(file_path),
            f"{os.path.splitext(file_name)[0]}_COPY_{int(datetime.datetime.timestamp(datetime.datetime.now()))}{os.path.splitext(file_name)[1]}"
        )
    
        dest_wb.save(copy_path)

        src_wb.close()
        dest_wb.close()
        print("Processamento concluído com sucesso.")

except Exception as e:
    print(traceback.format_exc())
    print(f"ERRO DURANTE EXECUÇÃO: {e}")

finally:
    pass
    # print("Limpando arquivos temporários...")
    # clean_temp_folder(tempfile.gettempdir())
    # clean_temp_folder("C:\\Windows\\Temp")