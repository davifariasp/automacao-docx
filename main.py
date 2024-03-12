from docx import Document
from datetime import datetime
import pandas as pd

tabela = pd.read_excel("./source/database.xlsx")

for linha in tabela.index:
    try:
        documento = Document("./source/contrato-original.docx")

        nome = tabela.loc[linha, "nome"]
        item1 = tabela.loc[linha, "item1"]
        item2 = tabela.loc[linha, "item2"]
        item3 = tabela.loc[linha, "item3"]
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

        nome_documento = f"Contrato - {nome}"
        documento.save(f"./output/{nome_documento}.docx")

    except Exception as e:
        print(e)

