import xlwings as xw
import pandas as pd

### EDITEM OS VALORES AQUI ###
panilha_alvo = "08-RST_ANL_VRF"
lista_de_referencia = 'Listas'
planilha_de_referencia = '07-RST_ANL_JTF'
linha_headers = 3
################

def procx(value, lookup_array, return_array, if_not_found=None):
    value = list(value) if not isinstance(value, list) else value
    
    for v in range(len(value)):
        index = lookup_array.index(value[v])
        value[v] = return_array[index]
        
    return value
        
    
# def formula1(conta):
#     return procx(conta, lista_de_referencia.range(f"Q{linha_headers + 1}:Q80").value, lista_de_referencia.range(f"R{linha_headers + 1}:R80").value)

# def formula2(valor):
#     return "0" if valor == "N.D" else valor

# def formulaX(analise):
#     return procx(analise, lista_de_referencia.range(f"T{linha_headers + 1}:T{linha_headers + 1}4").value, lista_de_referencia.range(f"V{linha_headers + 1}:V{linha_headers + 1}4").value) 

# def formula4(analise):
#     return procx(analise, lista_de_referencia.range(f"T{linha_headers + 1}:T{linha_headers + 1}4").value, lista_de_referencia.range(f"W{linha_headers + 1}:W{linha_headers + 1}4").value)

def formula5(formula):
    if formula == 2:
        return ""
    elif formula == 1:
        return "Sim"
    elif formula == 0:
        return "Não"
    else:
        return ""

def formula6(observacoes, formula4, freq_portaria):
    if observacoes and formula4 == 1:
        return "Não"
    elif freq_portaria == "Sim":
        return "Sim"
    else:
        return "Não"

def formula7(iqa, freq_portaria, vmp_portaria, u4, k4, justificativa, parecer):
    if iqa == "Não":
        return "NA"
    elif freq_portaria == "Sim" and vmp_portaria == "Não" and u4 == "Sim":
        return "Conforme"
    elif freq_portaria == "Sim" and vmp_portaria == "Sim" and u4 == "Sim" and k4 == "Conforme":
        return "Conforme"
    elif iqa == "Sim" and justificativa == "Sim":
        return "Conforme"
    elif (freq_portaria == "Sim" and vmp_portaria == "Sim" and u4 == "Sim"
          and parecer == "Não Conforme" and (justificativa in ["", "Não"])):
        return "Não Conforme"
    else:
        return "Continuar Formula"

# def formula8(analise):
#     return procx(analise, lista_de_referencia.range(f"T{linha_headers + 1}:T{linha_headers + 1}4").value, lista_de_referencia.range(f"U{linha_headers + 1}:U{linha_headers + 1}4").value)

# def formula9(valor):
#     return procx(valor, planilha_de_referencia.range(f"U5:U1008").value, planilha_de_referencia.range("V5:V1008").value)

# def formula10(e4):
#     return procx(e4, lista_de_referencia.range(f"Z{linha_headers + 1}:Z12897").value, lista_de_referencia.range(f"AH{linha_headers + 1}:AH12897").value)

def formula11(cidade2, conta, analise, parecer):
    return f"{cidade2}{conta}{analise}{parecer}"

def formula12(cidade2, conta, analise, parecer, etapa):
    return f"{cidade2}{conta}{analise}{parecer}{etapa}"

def formula1X(whatever):
    return whatever

# def formula14(aa4):
#     return procx(aa4, planilha_de_referencia.range("V5:V1561").value, planilha_de_referencia.range("W5:W1561").value, if_not_found="Sem valor")