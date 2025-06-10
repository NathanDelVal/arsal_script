import xlwings as xw
import pandas as pd
import os
import shutil
import traceback

def procx(value, lookup_array, return_array, if_not_found=None):
    value = list(value) if not isinstance(value, list) else value
    v = 0
    try:
        for v in range(len(value)):
            index = lookup_array.index(value[v])
            value[v] = return_array[index]
    except Exception as e:
        # print(f"{len(lookup_array)} // {len(return_array)}")
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

def find_headers_index(sht):
    for x in range(len(sht.used_range.columns)):
        for y in range(len(sht.used_range.columns[x].value)):
            if sht.used_range.columns[x].value[y] is not None:
                return y

def col_letter_to_index(col_letter: str) -> int:
    col_letter = col_letter.upper()
    index = 0
    for char in col_letter:
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index

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