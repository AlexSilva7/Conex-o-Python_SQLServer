

import webbrowser
import datetime
from pathlib import Path
import time
import os
import pyodbc 
import pyautogui

server = '.\SQLEXPRESS' 
database = 'FIBRA' 
cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';Trusted_Connection=yes;')
cursor = cnxn.cursor()

os.startfile('C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Cisco\\Cisco AnyConnect Secure Mobility Client\\Cisco AnyConnect Secure Mobility Client.lnk')

login = ''
senha = ''

time.sleep(10)

pyautogui.press('enter')
time.sleep(10)
pyautogui.press('tab')
time.sleep(5)
pyautogui.press('enter')	
time.sleep(10)
pyautogui.typewrite(login, interval=0.2)
time.sleep(2)
pyautogui.press('tab')
time.sleep(2)
pyautogui.typewrite(senha, interval=0.2)
time.sleep(2)
pyautogui.press('enter')	
time.sleep(10)

currentDateTime = datetime.datetime.now()

date = currentDateTime.date()

ano = date.year
mes = date.month
dia = date.day

if mes < 10:
	mes = (f"0{mes}")

if dia < 10:
	dia = (f"0{dia}")

ano = (f"{ano}")
mes = (f"{mes}")
dia = (f"{dia}")

url = "http://netwin.intranet/RelatoriosOffline/surveys/Enderecos_Totais_" + ano + mes + dia + ".csv"
webbrowser.open(url, 2)

fileName = r"C:\Users\alex.silva\Downloads\Enderecos_Totais_" + ano + mes + dia + ".csv"
fileObj = Path(fileName)


while(fileObj.is_file() == False):
	print("Aguardando conclusão do download")
	time.sleep(60)
	fileName = r"C:\Users\alex.silva\Downloads\Enderecos_Totais_" + ano + mes + dia + ".csv"
	fileObj = Path(fileName)



print("Dowload finalizado")
time.sleep(2)

print("Transportando CSV para a pasta do projeto")
os.rename("C:\\Users\\alex.silva\\Downloads\\Enderecos_Totais_" + ano + mes + dia + ".csv", "C:\\Users\\alex.silva\\Desktop\\END_TOTAIS\\END_TOTAIS.csv")

print("Verificando arquivo")
with open("END_TOTAIS.csv") as x:
	contador = sum(1 for line in x)
x.close()

print("Abrindo o arquivo")
f = open("END_TOTAIS.csv","r")

#Cabecalho
print("Lendo cabecalho")
f.readline()

vetor = []

#vetor.append(f.readline().split("|"))
print("Lendo arquivo")
for i in range(contador):
	vetor.append(f.readline())	
f.close()

print("Crianco copia em txt")
arquivo = open("END_TOTAIS.txt", "a")
arquivo.writelines(vetor)
arquivo.close()

print("Apagando tabela anterior")
cursor.execute("""
	TRUNCATE TABLE FIBRA.DBO.END_TOTAIS
""")
cnxn.commit()
cnxn.close()

print("Carregando para o banco de dados")
os.system("sqlcmd -S ""LGT008612\\SQLEXPRESS"" -d ""FIBRA"" -i C:\\Users\\alex.silva\\Desktop\\END_TOTAIS\\BULK_INSERT_END_TOTAIS.sql -o C:\\Users\\alex.silva\\Desktop\\END_TOTAIS\\RESULT.txt")
	
print("Carga concluida!")

print("Removendo txt")
os.remove("END_TOTAIS.txt")
	
print("Salvando backup do Csv na pasta downloads")
os.rename("C:\\Users\\alex.silva\\Desktop\\END_TOTAIS\\END_TOTAIS.csv", "C:\\Users\\alex.silva\\Downloads\\Enderecos_Totais_" + ano + mes + dia + ".csv")

print("Processo finalizado!")
fim = input("  TECLE ENTER PARA ENCERRAR!")




