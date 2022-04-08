class Certificado:
    def __init__(self, emitidoPor, validoApartirDe, validoAte, tempoRestante):
        self.emitidoPor = emitidoPor
        self.validoApartirDe = validoApartirDe
        self.validoAte = validoAte
        self.tempoRestante = tempoRestante

    def getEmitidoPor(self):
        return self.emitidoPor

    def getValidade(self):
        return self.validadeInteira

    def getValidoApartirDe(self):
        return self.validoApartirDe

    def getValidoAte(self):
        return self.validoAte

    def getTempoRestante(self):
        return self.tempoRestante

    def setNomeCertificado(self, emitidoPor):
        self.emitidoPor = emitidoPor

    def setValidade(self, validadeInteira):
        self.validadeInteira = validadeInteira

    def setValidoApartirDe(self, validoApartirDe):
        self.validoApartirDe = validoApartirDe

    def setValidoAte(self, validoAte):
        self.validoAte = validoAte

    def setTempoRestante(self, tempoRestante):
        self.tempoRestante = tempoRestante


    def __str__(self):
        return 'Emitido por: %s  |  validoApartirDe: %s  |  validoAte: %s  |  tempoRestante: %s' % (
        self.emitidoPor, self.validoApartirDe, self.validoAte, self.tempoRestante)
