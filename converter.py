import PyPDF2
import pytesseract
import pandas as pd
from pdf2image import convert_from_path
from PIL import Image

# CONVIERTE LAS PAGINAS DEL PDF EN IMAGENES
paginas_imagenes = convert_from_path('Ignacio Guidali.pdf')


with open ('Ignacio Guidali.pdf', 'rb') as file:
    # lee el archivo y se obtiene el numero de paginas
    reader = PyPDF2.PdfReader(file)
    num_paginas = len(reader.pages)

    # string vacio para luego almacenar el texto importado del pdf
    texto_pdf = ''

    # SE RECORRE CADA PAGINA/IMAGEN, SE CONVIERTE ESA IMAGEN EN GRIS Y SE EXTRAE EL TEXTO CON PYTESSERACT
    for pagina, imagen in enumerate(paginas_imagenes):
        imagen_gris = imagen.convert('L')
        texto = pytesseract.image_to_string(imagen_gris)
        # el texto extraido se concatena en la variable texto_pdf
        texto_pdf += texto
# COPIA EL CONTENIDO DE TEXTO QUE SE GUARDO EN LA VARIABLE TEXTO_PDF Y LO PEGA EN EL ARCHIVO TXT
with open ('texto.txt', 'w', encoding='utf-8') as file:
    file.write(texto_pdf)



# DE ESTA FORMA EN VEZ DE GUARDARLO EN UN ARCHIVO DE TEXTO, CREAMOS UN ARCHIVO .XLSX CON EL CONTENIDO 

    # Reemplaza los caracteres no imprimibles en el archivo excel
texto_pdf = texto_pdf.replace('\n', ' ').replace('\r', '')
texto_pdf = ''.join(filter(lambda x: x.isprintable(), texto_pdf))

    # Separa el texto extraido en un listado de palabras utilizando el metodo split
palabras = texto_pdf.split()

    # Luego crea una lista de 20 palabras para imprimir en cada celda, con esto resolvemos el error que puede ocacionar copiar mucho texto en una sola celda de excel
max_palabras_por_celda = 20
contenido_celdas = [palabras[i:i+max_palabras_por_celda] for i in range(0, len(palabras), max_palabras_por_celda)]

    # Convertimos ese listado en string
contenido_celdas = [' '.join(celda) for celda in contenido_celdas]

    # Crear un DataFrame de pandas con el texto extraído
df = pd.DataFrame({'Texto': contenido_celdas})

    # Guardar el DataFrame en un archivo de Excel
df.to_excel('Ignacio Guidali.xlsx', index=False, engine='openpyxl')            

print("ready")
