

import os
from openpyxl import Workbook

arquivo_excel = Workbook()
planilha1 = arquivo_excel.active

print("*****************************************")
print('DE ACORDO!\n')

regiao = input("DIGITE (1) PARA R1 OU (2) PARA R2: ")
UF = input("DIGITE A UF: ")
if(regiao == '1'):
   LOC = input("DIGITE A LOCALIDADE: ")
COD = input("DIGITE O CODLOG: ")

server = 'LGT008612\SQLEXPRESS'
banco = 'FIBRA'

#comando = (f'"set nocount on SELECT top 10 * FROM FIBRA.dbo.END_TOTAIS WHERE UF = {UF} AND LOCALIDADE_ABREV = {LOC} AND COD_LOGRADOURO = {COD}"')
if(regiao == '1'):
   comando = (f'"set nocount on SELECT top 10 * FROM FIBRA.dbo.END_TOTAIS WHERE UF = \'{UF}\' AND LOCALIDADE_ABREV = \'{LOC}\' AND COD_LOGRADOURO = \'{COD}\'"')
else:
   comando = (f'"set nocount on SELECT top 10 * FROM FIBRA.dbo.END_TOTAIS WHERE UF = \'{UF}\' AND COD_LOGRADOURO = \'{COD}\'"')
   
var = 'sqlcmd -S '+server+' -d '+banco+' -Q ' + comando + ' -o SURVEYS.csv -W -s,'

print("*****************************************")
print("Executando comando SQL ...")
os.system(var)

print("Resultado Gerado em CSV ...")

print("Contando as linhas do arquivo ...")
with open("SURVEYS.csv") as x:
   contador = sum(1 for line in x)

print("Lendo arquivo CSV ...")
f = open("SURVEYS.csv", "r")

print("Adicionando as linhas do arquivo CSV para um arquivo Excel-xlsx ...")
for i in range(contador):
   linha = f.readline().split(",")
   if i != 1:
      planilha1.append(linha)

print("Salvando arquivo  ...")
arquivo_excel.save("SURVEYS.xlsx")

f.close()
os.remove("SURVEYS.csv")

print("Programa concluido  ...")
input("Tecle para encerrar")


