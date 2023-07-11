import os
from tkinter.filedialog import askdirectory
import pandas as pd
import re
import time
from tkinter.messagebox import showinfo
import pdfplumber

# Preguntar por el directorio
directorio = askdirectory(title="Seleccionar carpetas con PDF de Facturas")

Start = time.time()

# Listar todos los archivos del directorio y sus subdirectorios en una lista sin backslash
lista_archivos = []
for root, dirs, files in os.walk(directorio):
    for file in files:
        if file.endswith(".pdf"):
            lista_archivos.append(os.path.join(root, file).replace("\\", "/"))

# Exportar lista a un archivo txt 
with open("lista_archivos.txt", "w") as output:
    output.write(str(lista_archivos))

# Crear un dataframe vacío
df = pd.DataFrame(columns=["Archivo", "CUIT del Emisor" , "COD" , "Punto de Venta", "Número de Factura", "Fecha", "Desde" , "Hasta"])

# Extraer el texto de los archivos PDF solamente de la primera página
for i in lista_archivos:
    with pdfplumber.open(i) as pdf:
        primera_pagina = pdf.pages[0]
        texto = primera_pagina.extract_text()
        #print(texto)
        #print("--------------------------------------------------")
        
        Archivo = i.split("/")[-1].replace(".pdf", "")

        # Extraer el COD de de factura
        Cod = re.search(r"COD. (\d+)", texto)
        Cod = Cod.group(1)
        #print(Cod)

        # Extraer el CUIT del emisor
        CUIT = re.search(r"CUIT: (\d+)", texto)
        CUIT = CUIT.group(1)
        #print(CUIT)

        # Extraer el punto de venta
        punto_venta = re.search(r"Punto de Venta: (\d+)", texto)
        punto_venta = punto_venta.group(1)
        #print(punto_venta)

        # Extraer el número de factura
        numero_factura = re.search(r"Comp. Nro: (\d+)", texto)
        numero_factura = numero_factura.group(1)
        #print(numero_factura)

        # Extraer la fecha
        fecha = re.search(r"Fecha de Emisión: (\d+/\d+/\d+)", texto)
        fecha = fecha.group(1)
        #print(fecha)

        # Extraer el rango de facturas, si no existe se deja vacío
        Desde = re.search(r"Desde:(\d+/\d+/\d+)", texto)
        if Desde == None:
            Desde = ""
        else:
            Desde = Desde.group(1)
        #print(Desde)
        Hasta = re.search(r"Hasta:(\d+/\d+/\d+)", texto)
        if Hasta == None:
            Hasta = ""
        else:
            Hasta = Hasta.group(1)
        #print(Hasta)

        # Agregar una linea nueva con los datos extraidos
        df = pd.concat([df, pd.DataFrame([[Archivo, Cod, CUIT, punto_venta, numero_factura, fecha, Desde, Hasta]], columns=["Archivo", "COD" , "CUIT del emisor" , "Punto de Venta", "Número de Factura", "Fecha", "Desde" , "Hasta"])], ignore_index=True)


# Exportar el dataframe a un archivo csv
df.to_csv("Facturas.csv", index=False)

End = time.time()

Tiempo_Total = End - Start

showinfo("Extracción de datos", f"Se extrajeron los datos de {len(lista_archivos)} archivos en {Tiempo_Total} segundos")