import ssl
import datetime
# import mysql.connector

from datetime import date
from datetime import datetime

import xlwt
from tkinter import *
from tkinter import messagebox
from cryptography import x509
# @Cnum#
from entities.Certificado import Certificado

maxDays = 10

# Way to save it into a SQL table (import mysql.connector)
# ########################################################################################################################
# # -----------------------------------------Operações e conexão do Banco de Dados---------------------------------------#
# ########################################################################################################################
#
# db = mysql.connector.connect(
#     host='localhost',
#     user='root',
#     password='',
#     database='info_certificados'
# )
# cursor = db.cursor()
#
# cursor.execute('drop table tb_certificados')
# sqlCreateTable = "CREATE TABLE tb_certificados (id INT unsigned auto_increment primary key not null," \
#                  "emitido_por varchar(255)," \
#                  "valido_a_partir varchar(255)," \
#                  "valido_ate varchar(255)," \
#                  "tempo_restante varchar(255)," \
#                  "created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP," \
#                  "updated_at DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP)"
# cursor.execute(sqlCreateTable)
#
# cursor.execute('drop table tb_certificadosWarning')
# sqlCreateTable = "CREATE TABLE tb_certificadosWarning (id INT unsigned auto_increment primary key not null," \
#                  "emitido_por varchar(255)," \
#                  "valido_a_partir varchar(255)," \
#                  "valido_ate varchar(255)," \
#                  "tempo_restante varchar(255)," \
#                  "created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP," \
#                  "updated_at DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP)"
# cursor.execute(sqlCreateTable)
#
#
# def create(emitidoPor, validoApartirDe, validoAte, tempoRestante, table):
#     sqlInsert = f"INSERT INTO {table} (emitido_por, valido_a_partir, valido_ate, tempo_restante) VALUES (%s, %s, %s, %s)"
#     data = (rf"{emitidoPor}", rf"{validoApartirDe}", rf"{validoAte}", rf"{tempoRestante}")
#     cursor.execute(sqlInsert, data)
#     db.commit()
#
#
# def readAll(table):
#     cursor = db.cursor()
#     sqlSelect = f"SELECT * FROM {table}"
#     cursor.execute(sqlSelect)
#     return cursor.fetchall()
#
#
# def read_emitidoPor(emitidoPor, table):
#     sqlSelect = f"SELECT * FROM {table} where emitido_por = '{emitidoPor}'"
#     cursor.execute(sqlSelect)
#     return cursor.fetchall()
#
#
# def delete(emitidoPor, table):
#     sqlDelete = f"DELETE FROM {table} WHERE emitido_por = '{emitidoPor}'"
#     cursor.execute(sqlDelete)
#     db.commit()
#
#     if len(read_emitidoPor(emitidoPor, 'tb_certificados')) == 0:
#         print('Deletado com sucesso.')
#     else:
#         print('Houve um problema.')
#
#
# def update_emitidoPor(actualName, newName, table):
#     sqlUpdate = f"UPDATE {table} SET emitido_por = '{newName}' WHERE emitido_por = '{actualName}'"
#     cursor.execute(sqlUpdate)
#     db.commit()

########################################################################################################################
# -------------------------------------------------------Funções-------------------------------------------------------#
########################################################################################################################


def removeCharacters(item, characters):
    aux = item
    characters = characters

    for i in range(0, len(characters)):
        aux = aux.replace(characters[i], "")
    return aux


def transformToDateTime(data):
    return datetime.strptime(data, '%Y-%m-%d').date()


def printObjectList(objectList):
    for i in range(len(objectList)):
        print(objectList[i].__str__())


certificadosList = []
certificadosWarning = []
today_str = date.today().strftime("%d/%m/%Y")
today_dateTime = datetime.strptime(today_str, '%d/%m/%Y').date()

########################################################################################################################
# ---------------------------------------Captura dados dos certificados instalados-------------------------------------#
########################################################################################################################

# Lista com todos os certificados
for store in ["MY"]:
    for cert, encoding, trust in ssl.enum_certificates(store):
        certificate = x509.load_der_x509_certificate(cert, backend=None)

        certificateSubject = (str(certificate.subject).split(","))
        for i in range(len(certificateSubject)):
            if certificateSubject[i][:3] == "CN=":
                emitidoPor = removeCharacters(certificateSubject[i][3:], ')>')
                break

        certificateNotValidBefore_str = (str(certificate.not_valid_before)).split(' ')
        certificateNotValidBefore_date = transformToDateTime(certificateNotValidBefore_str[0])

        certificateNotValidAfter_str = (str(certificate.not_valid_after)).split(' ')
        certificateNotValidAfter_date = transformToDateTime(certificateNotValidAfter_str[0])

        tempoRestante = (certificateNotValidAfter_date - today_dateTime).days

        certificadosList.append(
            Certificado(emitidoPor, certificateNotValidBefore_date, certificateNotValidAfter_date, tempoRestante))

# Lista com os certficados com o prazo próximo
for i in range(len(certificadosList)):
    if int(certificadosList[i].getTempoRestante()) <= maxDays:
        certificadosWarning.append(certificadosList[i])

# # Atualiza as informações no banco de dados
# for i in range(len(certificadosList)):
#     create(certificadosList[i].getEmitidoPor(), certificadosList[i].getValidoApartirDe(),
#            certificadosList[i].getValidoAte(), certificadosList[i].getTempoRestante(), 'tb_certificados')
#
# for i in range(len(certificadosWarning)):
#     create(certificadosWarning[i].getEmitidoPor(), certificadosWarning[i].getValidoApartirDe(),
#            certificadosWarning[i].getValidoAte(), certificadosWarning[i].getTempoRestante(), 'tb_certificadosWarning')


########################################################################################################################
# -------------------------------------------------Salva dados em um .xls----------------------------------------------#
########################################################################################################################

# Arquivo com todos os certificados
workbook = xlwt.Workbook()
resultSheet = workbook.add_sheet('ResultSheet')

# Escreve o header
header = ['Emitido por', 'Valido a partir de', 'Válido ate', 'Tempo restante (Dias)']
for h in range(len(header)):
    resultSheet.write(0, h, header[h])

for i in range(len(certificadosList)):
    resultSheet.write(i + 1, 0, certificadosList[i].getEmitidoPor())
    resultSheet.write(i + 1, 1, certificadosList[i].getValidoApartirDe().strftime("%d/%m/%Y"))
    resultSheet.write(i + 1, 2, certificadosList[i].getValidoAte().strftime("%d/%m/%Y"))
    resultSheet.write(i + 1, 3, certificadosList[i].getTempoRestante())

workbook.save(rf'resumo.xls')

# Arquivo com os certificados próximo do vencimento
workbook = xlwt.Workbook()
resultSheet = workbook.add_sheet('ResultSheet')

# Escreve o header
header = ['Emitido por', 'Valido a partir de', 'Válido ate', 'Tempo restante (Dias)']
for h in range(len(header)):
    resultSheet.write(0, h, header[h])

for i in range(len(certificadosWarning)):
    resultSheet.write(i + 1, 0, certificadosWarning[i].getEmitidoPor())
    resultSheet.write(i + 1, 1, certificadosWarning[i].getValidoApartirDe().strftime("%d/%m/%Y"))
    resultSheet.write(i + 1, 2, certificadosWarning[i].getValidoAte().strftime("%d/%m/%Y"))
    resultSheet.write(i + 1, 3, certificadosWarning[i].getTempoRestante())

workbook.save(rf'resumo_proxVencimento.xls')
########################################################################################################################
# ----------------------------------------------------Janela de aviso--------------------------------------------------#
########################################################################################################################

alertMessage = ''
# Caixa de alerta
if len(certificadosWarning) != 0:
    alertMessage += '#' * 45 + '\n'
    alertMessage += 'ATENÇÃO! CERTIFICADOS COM O VENCIMENTO PRÓXIMO:\n'
    alertMessage += '#' * 45 + '\n'
    alertMessage += '\n' + '-' * 79
    for i in range(len(certificadosWarning)):
        alertMessage += '\n' + f'Certificado: {certificadosWarning[i].getEmitidoPor()}' + \
                        '\n' + f'Valido até: {certificadosWarning[i].getValidoAte().strftime("%d/%m/%Y")}' + \
                        '\n' + f'Tempo restante (Dias): {certificadosWarning[i].getTempoRestante()}' + \
                        '\n' + '-' * 79
if len(certificadosWarning) == 0:
    alertMessage = 'Não há nenhum certificado com a data de vencimento próxima.'

app = Tk()
app.title("VCDI")
app.geometry("500x500")
messagebox.showinfo(title="VCDI", message=alertMessage)
scrollBar = Scrollbar(app)
scrollBar.pack(side=RIGHT, fill=Y)
text = Text(app, font="Arial 12")
text.pack(expand=YES, fill=BOTH)

# Janela principal
for i in range(len(certificadosList)):
    text.insert(0.0, '\n')
    text.insert(0.0, f'Tempo restante (Dias): {certificadosList[i].getTempoRestante()}')
    text.insert(0.0, '\n')
    text.insert(0.0, f'Válido até: {certificadosList[i].getValidoAte().strftime("%d/%m/%Y")}')
    text.insert(0.0, '\n')
    text.insert(0.0, f'Válido a partir de: {certificadosList[i].getValidoApartirDe().strftime("%d/%m/%Y")}')
    text.insert(0.0, '\n')
    text.insert(0.0, f'Emitido por: {certificadosList[i].getEmitidoPor()}')
    text.insert(0.0, '\n')
    text.insert(0.0, '-'*10)

text.config(yscrollcommand=scrollBar.set)
scrollBar.config(command=text.yview())

app.mainloop()