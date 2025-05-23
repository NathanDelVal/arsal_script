import openpyxl as openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook
import openpyxl.workbook.properties
import pandas as pd
import argparse
import sys
import xlsxwriter
# import numpy as np
# import matplotlib.pyplot as plt

openpyxl.workbook.properties.CalcProperties.fullCalcOnLoad = True
# openpyxl.workbook.properties.CalcProperties.calcOnSave = True

excel_expressions = [
    '=_xlfn.XLOOKUP([@Conta],Listas!$Q$3:$Q$80,Listas!$R$3:$R$80,"",0,1)',
    '=SE([@Valor]="N.D","0",[@Valor])',
    '=_xlfn.XLOOKUP([@Análise],Listas!$T$3:$T$34,Listas!$V$3:$V$34,"",0,1)',
    '=_xlfn.XLOOKUP([@Análise],Listas!$T$3:$T$34,Listas!$W$3:$W$34,"",0,1)',
    '=SE([@[Formula para calculo3]]=2,"",SE([@[Formula para calculo3]]=1,"Sim",SE([@[Formula para calculo3]]=0,"Não","")))',
    '=SE(E(CONT.VALORES([@Observações])>0,[@[Formula para calculo4]]=1),"Não",SE([@[Posssui Frequência na Portaria?]]="Sim","Sim","Não"))',
    '=SE([@[Aceito no IQA?]]="Não","NA",SE(E([@[Posssui Frequência na Portaria?]]="Sim",[@[Possui VMP na Portaria?]]="Não",[@[Aceito no IQA?]]="Sim"),"Conforme",SE(E([@[Posssui Frequência na Portaria?]]="Sim",[@[Possui VMP na Portaria?]]="Sim",[@[Aceito no IQA?]]="Sim",[@[Parecer Análise]]="Conforme"),"Conforme",SE(E([@[Aceito no IQA?]]="Sim",[@[Justificativa aceita?]]="Sim"),"Conforme",SE(E([@[Posssui Frequência na Portaria?]]="Sim",[@[Possui VMP na Portaria?]]="Sim",[@[Aceito no IQA?]]="Sim",[@[Parecer Análise]]="Não Conforme",OU([@[Justificativa aceita?]]="",[@[Justificativa aceita?]]="Não")),"Não Conforme","Continuar Formula")))))',
    '=_xlfn.XLOOKUP([@Análise],Listas!$T$3:$T$34,Listas!$U$3:$U$34,"",0,1)',
    '',
    '',
    '',
    '=_xlfn.XLOOKUP([@[Formula para calculo2]],\'07-RST_ANL_JTF\'!$U$5:$U$1008,\'07-RST_ANL_JTF\'!$V$5:$V$1008,"",0,1)',
    '',
    '',
    '=_xlfn.XLOOKUP([@[Ponto Coleta]],Listas!$Z$3:$Z$12897,Listas!$AH$3:$AH$12897,"",0,1)',
    '',
    '=[@Cidade2]&[@Conta]&[@Análise]&[@[Parecer Análise]]',
    '=[@Cidade2]&[@Conta]&[@Análise]&[@[Parecer Análise]]&[@[Etapa do Processo]]',
    '=_xlfn.XLOOKUP(@[[ID_VI]],\'07-RST_ANL_JTF\'!$V$5:$V$1561,\'07-RST_ANL_JTF\'!$W$5:$W$1561,"Sem valor",0,1)',
    ''
]

#  ['Cidade2' 'Resultado independente' 'Posssui Frequência na Portaria?'
#  'Possui VMP na Portaria?' 'Justificativa aceita?' 'Aceito no IQA?'
#  'Análise de conformidade VI' 'Tipo Análise' 'Responsável Análise'
#  'ID_Conc' 'Justificativa  Concessionária' 'ID_VI' 'Justificativa VI'
#  'Acatar' 'Etapa do Processo' 'Observações' 'Formula para calculo'
#  'Formula para calculo2' 'Formula para calculo3' 'Formula para calculo4']

sheets = pd.read_excel("arsal.xlsx", sheet_name=None, engine="openpyxl")    

sheet = sheets["08-RST_ANL_VRF"]

heads_ref = (sheet.iloc[1, :].values[15:])

rows_num = len(sheet.loc[:])

parser = argparse.ArgumentParser(description="Authomate Excel file processing")

parser.add_argument("column_heads", nargs="+", help="column heads to create")

args = parser.parse_args()

# print(args.column_heads)

dict_items = {key: val for key, val in zip(heads_ref, excel_expressions)}

df = pd.DataFrame(dict_items, index=[0])

# Get the first row (index 0)
row_to_repeat = df.loc[[0]]  # note the double brackets to keep it as a DataFrame

# Repeat it X times
df = pd.concat([row_to_repeat] * rows_num, ignore_index=True)

with pd.ExcelWriter('$testeee.xlsx', engine='xlsxwriter') as writer:
    workbook  = writer.book
    worksheet = workbook.add_worksheet("Sheet1")
    writer.sheets["Sheet1"] = worksheet

    # Write headers
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value)

    # Write formula strings as actual Excel formulas
    for row_num, row in enumerate(df.itertuples(index=False), start=0):
        for col_num, cell in enumerate(row):
            if isinstance(cell, str) and cell.strip().startswith('='):
                worksheet.write_formula(row_num, col_num, cell)
            else:
                worksheet.write(row_num, col_num, cell)

# Open the file you've written
# wb = load_workbook("$testeee.xlsx")
# wb.save("$testeee.xlsx")