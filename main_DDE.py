import pprint

try:
    input_file
except NameError:
    print("Erro na obtenção do input file. Contactar administrador do sistema.")
    import sys
    sys.path.append(r'G:\aaa\3. Python')
    input_file = r"G:\aaa\1. Ficheiros originais\BE_BP_06junho_2021.xlsx"
    data = {}

# Importação dos packages
import tabulizer
import pandas as pd
import numpy as np
from openpyxl import Workbook, load_workbook
from tabulizer import *


carregar_excel = tabulizer.excel.base.ExcelLoader.load(input_file)

#fazer um dicionário

Folha = carregar_excel.get_worksheet_by_index(0)
Folha2 = carregar_excel.get_worksheet_by_index(1)

##########################################
#First sheet
##########################################

RowMemberSet = {"left":"b","right":"c","top":10}
ColMemberSet = {"top":6,"bottom":8,"left":"D"}

#Chamar o scanner
scanner = tabulizer.excel.scan.TableScanner()

# Argumentos = (Folha,RowMemberset,ColMemberset) outros  row_max_skip, col_max_skip,row_member_strip_text)
scanner.scan_table(Folha,RowMemberSet,ColMemberSet, row_max_skip=2, col_max_skip=1,row_member_strip_text=True,col_member_strip_text=True)

#Passar linhas True
df = scanner.get_table(skip_nulls=True)

#adicionar coluna
df.insert(3, "Divisa", "Euro")

#Altera o nome das colunas
df.columns.values[4] = 'Ano'
df.columns.values[5] = 'Mês'

# Para linhas com periodo < 100, converter para float e adicionar 1900
df.loc[df['Ano'].astype(float) < 100, 'Ano'] = df['Ano'].astype(float) + 1900

# Na coluna 'Ano', converter para texto e ficar apenas com o ano (deixar cair as casas decimais)
df.loc[:, 'Ano'] = df['Ano'].astype(str).str[:4]

#ajustar mês
df.loc[df['Mês'] == 'Jan', 'Mês'] = '01'
df.loc[df['Mês'] == 'jan', 'Mês'] = '01'
df.loc[df['Mês'] == 'Fev', 'Mês'] = '02'
df.loc[df['Mês'] == 'fev', 'Mês'] = '02'
df.loc[df['Mês'] == 'Mar', 'Mês'] = '03'
df.loc[df['Mês'] == 'mar', 'Mês'] = '03'
df.loc[df['Mês'] == 'Abr', 'Mês'] = '04'
df.loc[df['Mês'] == 'abr', 'Mês'] = '04'
df.loc[df['Mês'] == 'Mai', 'Mês'] = '05'
df.loc[df['Mês'] == 'mai', 'Mês'] = '05'
df.loc[df['Mês'] == 'Jun', 'Mês'] = '06'
df.loc[df['Mês'] == 'jun', 'Mês'] = '06'
df.loc[df['Mês'] == 'Jul', 'Mês'] = '07'
df.loc[df['Mês'] == 'jul', 'Mês'] = '07'
df.loc[df['Mês'] == 'Ago', 'Mês'] = '08'
df.loc[df['Mês'] == 'ago', 'Mês'] = '08'
df.loc[df['Mês'] == 'Set', 'Mês'] = '09'
df.loc[df['Mês'] == 'set', 'Mês'] = '09'
df.loc[df['Mês'] == 'Out', 'Mês'] = '10'
df.loc[df['Mês'] == 'out', 'Mês'] = '10'
df.loc[df['Mês'] == 'Nov', 'Mês'] = '11'
df.loc[df['Mês'] == 'nov', 'Mês'] = '11'
df.loc[df['Mês'] == 'Dez', 'Mês'] = '12'
df.loc[df['Mês'] == 'dez', 'Mês'] = '12'

df.insert(5, "MêsAux", df.Mês)
df.loc[:, 'Mês'] = df['Mês'].astype(str).str[:2]
lista_meses = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
df.loc[~df['Mês'].isin(lista_meses), 'Mês'] = df['MêsAux'].str[5:7]

#criar coluna Periodo
df.insert(3,'Periodo',df.Ano+'-'+df.Mês)

#Correção da caraterização dos instrumentos
df.loc[df['Membro de colunas 1'] == 'Total 15=10+11+13+14', 'Membro de colunas 1'] = 'Total'
df.loc[df['Membro de colunas 2'] == 'Total 10=1+5', 'Membro de colunas 2'] = 'Total'
df.loc[df['Membro de colunas 1'] == 'Títulos', 'Membro de colunas 1'] = 'Títulos de dívida'
df.loc[df['Membro de colunas 1'] == 'Certificados de Aforro 11', 'Membro de colunas 3'] = 'Certificados de Aforro'
df.loc[df['Membro de colunas 3'] == 'Certificados de Aforro', 'Membro de colunas 1'] = 'Depósitos'
df.loc[df['Membro de colunas 1'] == 'Certificados do Tesouro 13', 'Membro de colunas 3'] = 'Certificados do Tesouro'
df.loc[df['Membro de colunas 3'] == 'Certificados do Tesouro', 'Membro de colunas 1'] = 'Depósitos'
df.loc[df['Membro de colunas 1'] == 'Outros empréstimos 14', 'Membro de colunas 1'] = 'Empréstimos'
df.loc[df['Membro de colunas 1'] == 'dos quais:', 'Membro de colunas 3'] = 'Certificados Aforro d.q.: capitalização acumulada'
df.loc[df['Membro de colunas 3'] == 'Certificados de Aforro d.q.: capitalização acumulada', 'Membro de colunas 1'] = 'Depósitos'
df.loc[df['Membro de colunas 2'] == 'Curto prazo 1', 'Membro de colunas 2'] = 'Curto prazo'
df.loc[df['Membro de colunas 2'] == 'Médio e longo prazos 5', 'Membro de colunas 2'] = 'Médio e longo prazos'
df.loc[df['Membro de colunas 2'] == 'Capitalização acumulada 12', 'Membro de colunas 1'] = 'Depósitos'
df.loc[df['Membro de colunas 2'] == 'Capitalização acumulada 12', 'Membro de colunas 2'] = 'Total'
df.loc[df['Membro de colunas 3'] == 'Certificados de Aforro d.q.: capitalização acumulada', 'Membro de colunas 2'] = 'Total'
df.loc[df['Membro de colunas 3'] == 'BT 2', 'Membro de colunas 2'] = 'Curto prazo'
df.loc[df['Membro de colunas 3'] == 'BT 2', 'Membro de colunas 3'] = 'Bilhetes do Tesouro'
df.loc[df['Membro de colunas 3'] == 'CEDIC 3', 'Membro de colunas 2'] = 'Curto prazo'
df.loc[df['Membro de colunas 3'] == 'CEDIC 3', 'Membro de colunas 3'] = 'CEDIC'
df.loc[df['Membro de colunas 3'] == 'ECP 4', 'Membro de colunas 2'] = 'Curto prazo'
df.loc[df['Membro de colunas 3'] == 'ECP 4', 'Membro de colunas 3'] = 'Papel comercial'
df.loc[df['Membro de colunas 3'] == 'OT e outros títulos de taxa fixa 6', 'Membro de colunas 2'] = 'Médio e longo prazos'
df.loc[df['Membro de colunas 3'] == 'OT e outros títulos de taxa fixa 6', 'Membro de colunas 3'] = 'OT e outros títulos de taxa fixa'
df.loc[df['Membro de colunas 3'] == 'Títulos de taxa variável 7', 'Membro de colunas 2'] = 'Médio e longo prazos'
df.loc[df['Membro de colunas 3'] == 'Títulos de taxa variável 7', 'Membro de colunas 3'] = 'Títulos de taxa variável'
df.loc[df['Membro de colunas 3'] == 'Com vencimento no prazo de 1 ano 9', 'Membro de colunas 2'] = 'Médio e longo prazos'
df.loc[df['Membro de colunas 3'] == 'Com vencimento no prazo de 1 ano 9', 'Membro de colunas 3'] = 'Títulos de dívida com vencimento no prazo de 1 ano'
df.loc[df['Membro de colunas 3'] == 'CEDIM 8', 'Membro de colunas 2'] = 'Médio e longo prazos'
df.loc[df['Membro de colunas 3'] == 'CEDIM 8', 'Membro de colunas 3'] = 'CEDIM'
df.loc[df['Membro de colunas 2'].isna(), 'Membro de colunas 2'] = 'Total'
df.loc[df['Membro de colunas 3'].isna(), 'Membro de colunas 3'] = 'Total'

#Apaga colunas que não interessam
df.drop(df.columns[[0,1,2,5,6,7]],axis=1,inplace=True)

#Altera o nome das colunas
df.columns.values[2] = 'Instrumento'
df.columns.values[3] = 'Prazo_contratual'
df.columns.values[4] = 'Instrumento_detalhe'
df.columns.values[5] = 'Valor_MEUR'

#Ordenar colunas
df=df.reindex(columns=['Periodo','Divisa','Instrumento','Instrumento_detalhe','Prazo_contratual','Valor_MEUR'])

##########################################
#Second sheet
##########################################

RowMemberSet = {"left":"b","right":"c","top":10}
ColMemberSet = {"top":6,"bottom":8,"left":"D"}

#Chamar o scanner
scanner = tabulizer.excel.scan.TableScanner()

# Argumentos = (Folha,RowMemberset,ColMemberset) outros  row_max_skip, col_max_skip,row_member_strip_text)
scanner.scan_table(Folha2,RowMemberSet,ColMemberSet, row_max_skip=2, col_max_skip=2,row_member_strip_text=True,col_member_strip_text=True)

#Passar linhas True
df2 = scanner.get_table(skip_nulls=True)

#adicionar colunas
df2.insert(3, "Divisa", "Não euro")
df2.insert(4,"Periodo",df.Periodo)

#Correção da caraterização dos instrumentos
df2.loc[df2['Coluna (texto)'] == 'I', 'Divisa'] = 'Todas as divisas'
df2.loc[df2['Coluna (texto)'] == 'I', 'Membro de colunas 1'] = 'Total'
df2.loc[df2['Coluna (texto)'] == 'I', 'Membro de colunas 2'] = 'Total'
df2.loc[df2['Coluna (texto)'] == 'H', 'Membro de colunas 1'] = 'Total'
df2.loc[df2['Coluna (texto)'] == 'H', 'Membro de colunas 3'] = 'Total'
df2.loc[df2['Membro de colunas 1'] == 'Outros empréstimos 19', 'Membro de colunas 1'] = 'Empréstimos'
df2.loc[df2['Membro de colunas 1'] == 'Títulos', 'Membro de colunas 1'] = 'Títulos de dívida'
df2.loc[df2['Membro de colunas 3'] == 'Com vencimento no prazo de 1 ano 18', 'Membro de colunas 3'] = 'Títulos de dívida com vencimento no prazo de 1 ano'
df2.loc[df2['Membro de colunas 3'] == 'Títulos de dívida com vencimento no prazo de 1 ano', 'Membro de colunas 2'] = 'Médio e longo prazos'
df2.loc[df2['Membro de colunas 2'] == 'Curto prazo 16', 'Membro de colunas 2'] = 'Curto prazo'
df2.loc[df2['Membro de colunas 2'] == 'Médio e longo prazos 17', 'Membro de colunas 2'] = 'Médio e longo prazos'
df2.loc[df2['Membro de colunas 3'] == 'Total 20=16+17+19', 'Membro de colunas 3'] = 'Total'
df2.loc[df2['Coluna (texto)'] == 'F', 'Membro de colunas 2'] = 'Total'
df2.loc[df2['Membro de colunas 2'].isna(), 'Membro de colunas 2'] = 'Total'
df2.loc[df2['Membro de colunas 3'].isna(), 'Membro de colunas 3'] = 'Total'

#Apaga colunas que não interessam
df2.drop(df2.columns[[0,1,2,5,6]],axis=1,inplace=True)

#Altera o nome das colunas
df2.columns.values[2] = 'Instrumento'
df2.columns.values[3] = 'Prazo_contratual'
df2.columns.values[4] = 'Instrumento_detalhe'
df2.columns.values[5] = 'Valor_MEUR'

#Ordenar colunas
df2=df2.reindex(columns=['Periodo','Divisa','Instrumento','Instrumento_detalhe','Prazo_contratual','Valor_MEUR'])

#CONCATENAR DUAS SHEETS
final = pd.concat([df, df2], ignore_index=True, sort=False)

#adicionar colunas
final.insert(5, "Tipo_Metrica", "Valor nominal")
final.insert(6,"Tipo_Valor", "Stock")

#arredondar valores
final=final.round(6)

final.to_excel(r"G:\aaa\TabelaFinal.xlsx")

data['DDE'] = final

"""
errors = SourceVersionValidator.validate(
    source="source004",
    source_version="V001",
    data=data
)
pprint.pprint(errors)

"""
