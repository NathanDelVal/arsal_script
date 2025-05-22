import openpyxl as openpyxl
from openpyxl import load_workbook
import openpyxl.workbook
import openpyxl.workbook.properties
import pandas as pd

openpyxl.workbook.properties.CalcProperties.fullCalcOnLoad = True
# import numpy as np
# import matplotlib.pyplot as plt

sheets = pd.read_excel("arsal.xlsx", sheet_name=None, engine="openpyxl")    

sheet = sheets["08-RST_ANL_VRF"]

heads_ref = (sheet.iloc[1, :].values[15:])

for value in heads_ref:
    matches = (sheet == value)
    for row_idx, col_idx in zip(*matches.values.nonzero()):
        row_label = sheet.index[row_idx]
        col_label = sheet.columns[col_idx]
        print(f"Value '{value}' found at row '{row_label}', column '{col_label}'")

# heads_ref = ["08-RST_ANL_VRF"]

# for s in sheets:

head_ref = ''

# print(sheet[head_ref])

# Open the file you've written
wb = load_workbook("$testeee.xlsx")
wb.save("$testeee.xlsx")