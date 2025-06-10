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
            lista_de_referencia = dest_wb.sheets[params.lista_de_referencia]
            planilha_de_referencia = dest_wb.sheets[params.planilha_de_referencia]
            planilha_alvo = dest_wb.sheets[params.planilha_alvo]
        except KeyError as e:
            print(f"ERRO AO ACESSAR PLANILHA: {e}")
            src_wb.close()
            dest_wb.close()
            sys.exit(1)
        
        last_row = planilha_alvo.used_range.last_cell.row
        last_col = planilha_alvo.used_range.last_cell.column
         
        headers = planilha_alvo.range((params.linha_headers, 1), (params.linha_headers, last_col)).value
        
        conta = planilha_alvo.range((params.linha_headers + 1, headers.index("Conta") + 1), (last_row, headers.index("Conta") + 1)).value
        
        cidade2 = formulas.procx([c for c in conta], 
                                 [c for c in lista_de_referencia.range(f"Q:Q").value if c is not None], 
                                 [c for c in lista_de_referencia.range(f"R:R").value if c is not None])
        
        analise = planilha_alvo.range((params.linha_headers + 1, headers.index("Análise") + 1), (last_row, headers.index("Análise") + 1)).value
        
        parecer_analise = planilha_alvo.range((params.linha_headers + 1, headers.index("Parecer Análise") + 1), (last_row, headers.index("Parecer Análise") + 1)).value
        
        etapa_processo = planilha_alvo.range((1, headers.index("Ponto Coleta") + 1), (last_row, headers.index("Ponto Coleta") + 1)).value
        etapa_processo = formulas.procx([c for c in etapa_processo[params.linha_headers:]], 
                                 [c for c in lista_de_referencia.range(f"Z:Z").value if c is not None], 
                                 [c for c in lista_de_referencia.range(f"AH:AH").value if c is not None])

        formula1 = [f"{v1}{v2}{v3}{v4}" for v1,v2,v3,v4 in zip(cidade2,conta,analise,parecer_analise)]

        formula2 = [f"{v1}{v2}{v3}{v4}{v5}" for v1,v2,v3,v4,v5 in zip(cidade2,conta,analise,parecer_analise, etapa_processo)] 
        
        id_vi = formulas.procx([c for c in formula2], 
                [c for c in planilha_de_referencia.range(f"Y:Y").value if c is not None], 
                [c for c in planilha_de_referencia.range(f"Z:Z").value if c is not None])

        observacoes = formulas.procx([c for c in id_vi], 
                [c for c in planilha_de_referencia.range(f"F:F").value if c is not None], 
                [c for c in planilha_de_referencia.range(f"U:U").value if c is not None])

        planilha_alvo[params.linha_headers - 1:, last_col].options(transpose=True).value = ["Cidade2"] + cidade2
        planilha_alvo[params.linha_headers - 1, last_col].color = (0, 0, 139)
        last_col = planilha_alvo.used_range.last_cell.column

        temp_c = planilha_alvo.range(
            (1, headers.index("Valor") + 1), (last_row, headers.index("Valor") + 1)
            ).value 
        temp_c = [0 if c == "N.D" else c for c in temp_c[params.linha_headers:]]
        planilha_alvo[params.linha_headers - 1:, last_col].options(transpose=True).value = ["Resultado independente"] + temp_c
        planilha_alvo[params.linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1

        temp_c = formulas.procx([c for c in analise], 
                                 [c for c in lista_de_referencia.range(f"T:T").value if c is not None], 
                                 [c for c in lista_de_referencia.range(f"V:V").value if c is not None])
        planilha_alvo[params.linha_headers - 1:, last_col].options(transpose=True).value = ["Possui Frequência na Portaria?"]  + temp_c
        planilha_alvo[params.linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        temp_c = formulas.procx([c for c in analise], 
                                 [c for c in lista_de_referencia.range(f"T:T").value if c is not None], 
                                 [c for c in lista_de_referencia.range(f"W:W").value if c is not None])
        planilha_alvo[params.linha_headers - 1:, last_col].options(transpose=True).value = ["Possui VMP na Portaria?"] + temp_c
        planilha_alvo[params.linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        temp_c = formulas.procx([c for c in analise], 
                                 [c for c in lista_de_referencia.range(f"T:T").value if c is not None], 
                                 [c for c in lista_de_referencia.range(f"U:U").value if c is not None])
        planilha_alvo[params.linha_headers - 1:, last_col].options(transpose=True).value = ["Tipo Análise"] + temp_c
        planilha_alvo[params.linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1

        planilha_alvo[params.linha_headers - 1:, last_col].options(transpose=True).value = ["Responsável Análise"] 
        planilha_alvo[params.linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        planilha_alvo[params.linha_headers - 1:, last_col].options(transpose=True).value = ["ID_Conc"] 
        planilha_alvo[params.linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        planilha_alvo[params.linha_headers - 1:, last_col].options(transpose=True).value = ["Justificativa  Concessionária"] 
        planilha_alvo[params.linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        planilha_alvo[params.linha_headers - 1:, last_col].options(transpose=True).value = ["ID_VI"] 
        planilha_alvo[params.linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        planilha_alvo[params.linha_headers - 1:, last_col].options(transpose=True).value = ["Justificativa VI"] 
        planilha_alvo[params.linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        planilha_alvo[params.linha_headers - 1:, last_col].options(transpose=True).value = ["Acatar"] 
        planilha_alvo[params.linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        planilha_alvo[params.linha_headers - 1:, last_col].options(transpose=True).value = ["Etapa do Processo"] + etapa_processo
        planilha_alvo[params.linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        planilha_alvo[params.linha_headers - 1:, last_col].options(transpose=True).value = ["Observações"] 
        planilha_alvo[params.linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        planilha_alvo[params.linha_headers - 1:, last_col].options(transpose=True).value = ["Fórmula para Cálculo"] + formula1
        planilha_alvo[params.linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        planilha_alvo[params.linha_headers - 1:, last_col].options(transpose=True).value = ["Fórmula para Cálculo 2"] + formula2
        planilha_alvo[params.linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        planilha_alvo[params.linha_headers - 1:, last_col].options(transpose=True).value = ["Fórmula para Cálculo 3"] 
        planilha_alvo[params.linha_headers - 1, last_col].color = (0, 0, 139)
        last_col += 1
        
        formulas.adjust_cols_width(planilha_alvo)

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