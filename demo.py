import pandas as pd
# import numpy as np
# import matplotlib.pyplot as plt

# sheets = pd.read_excel("arsal.xlsx", sheet_name=None, engine="openpyxl")    

# head_ref = ''

# for df in sheets["08-RST_ANL_VRF"].head().items():
#     if "Cidade2" in df[1].values:
#         head_ref = df[0]

# print(sheets["08-RST_ANL_VRF"][head_ref])

# Create a DataFrame with values and a formula
df = pd.DataFrame({
    "A": [10, 20, 30],
    "B": [1, 2, 3],
    "C": ["=A2+B2", "=A3+B3", "=A4+B4"]  # Excel-style formulas
})

# Save to Excel — keep index=False to avoid shifting references
df.to_excel("formulas.xlsx", index=False)
