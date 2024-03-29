import ssl
import datetime

from datetime import date
from datetime import datetime

import xlwt
from tkinter import *
from tkinter import messagebox
from cryptography import x509
from entities.Certificado import Certificado

maxDays = 10

########################################################################################################################
# -------------------------------------------------------Functions-----------------------------------------------------#
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

########################################################################################################################
# -------------------------------------------------Salva dados em um .xls----------------------------------------------#
########################################################################################################################
# Arquivo com todos os certificados

if len(certificadosList) > 0:
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

    workbook.save(rf'resum.xls')

    # Arquivo com os certificados próximo do vencimento
    if len(certificadosWarning) > 0:
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

        workbook.save(rf'near_from_due_date.xls')
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