from pyzbar.pyzbar import decode
from pdf2image import convert_from_path 
import os
import numpy as np

def linha_digitavel(linha):
    # https://www.bb.com.br/docs/pub/emp/mpe/dwn/PadraoCodigoBarras.pdf
    def modulo10(num):
        soma = 0
        peso = 2
        for c in reversed(num):
            parcial = int(c) * peso
            if parcial > 9:
                s = str(parcial)
                parcial = int(s[0]) + int(s[1])
            soma += parcial
            if peso == 2:
                peso = 1
            else:
                peso = 2

        resto10 = soma % 10
        if resto10 == 0:
            modulo10 = 0
        else:
            modulo10 = 10 - resto10
        return modulo10

    def monta_campo(campo):
        campo_dv = "%s%s" % (campo, modulo10(campo))
        return "%s.%s" % (campo_dv[0:5], campo_dv[5:])
    
    return ' '.join([monta_campo(linha[0:4] + linha[19:24]),
                        monta_campo(linha[24:34]),
                        monta_campo(linha[34:44]),
                        linha[4],
                        linha[5:19]])

def leitor_cod(pdf_path):
    img = convert_from_path(pdf_path, 300)[0]
    detec_cod = decode(img)

    if not detec_cod:
        return False
    else:
        for cod in detec_cod:
            if cod.data != "" and cod.type == "I25":
                return cod.data.decode('utf-8')
