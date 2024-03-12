from docx import Document
from datetime import datetime


try:
    documento = Document("./source/contrato-original.docx")

    nome = "Fulano de Tal"
    item1 = "Carro"
    item2 = "Celular"
    item3 = "Notebook"

    dia = str(datetime.now().day)
    mes = str(datetime.now().month)
    ano = str(datetime.now().year)

    referencias = {
        "XXXX": nome,
        "YYYY": item1,
        "ZZZZ": item2,
        "WWWW": item3,
        "DD": dia,
        "MM": mes,
        "AAAA": ano
    }

    for paragrafo in documento.paragraphs:
        # print(paragrafo.text)
        for codigo in referencias:
            valor = referencias[codigo]
            paragrafo.text = paragrafo.text.replace(codigo, valor)

    documento.save("./output/contrato-final.docx")
except Exception as e:
    print(e)

