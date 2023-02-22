from docx import Document

modelo = Document('clientes/ADONIAS/RECIBOS/MODELO.docx')

contador_titulos = 0
contador_recibos = 0
temp = 0
for paragrafo in modelo.paragraphs:
    linha = paragrafo.text
        
    if contador_titulos > 1:
        contador_recibos = contador_recibos + 1
        contador_titulos = 0
        temp = temp + 1
        
    if linha == 'RECIBO DE ALUGUEL':
        contador_titulos =  contador_titulos + 1
        
        
print(temp)