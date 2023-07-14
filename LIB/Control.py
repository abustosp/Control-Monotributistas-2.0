import pandas as pd
from tkinter.messagebox import showinfo
import numpy as np
import os
import time
import re
import pdfplumber
import openpyxl

def Extraer_PDF_info(PDFpath: str):
    '''
    Extrae los Datos de todos de las facturas PDF de una carpeta que se descargaron del servicio Comprobantes en Línea de AFIP

    Parameters
    ----------
    path : str
        Path de la carpeta donde se encuentran los PDF de las facturas
    '''

    directorio = PDFpath

    Start = time.time()

    # Listar todos los archivos del directorio y sus subdirectorios en una lista sin backslash
    lista_archivos = os.listdir(directorio)

    # Filtrar la lista para que solo queden los archivos PDF
    lista_archivos = [i for i in lista_archivos if i.endswith(".pdf")]

    # Agregar el directorio a cada archivo de la lista
    lista_archivos = [directorio + "/" + i for i in lista_archivos]

    # Cambiar los backslash por slash
    lista_archivos = [i.replace("\\", "/") for i in lista_archivos] 

    # Eliminar de la lista los PDF que pesen menos de 10 KB
    lista_archivos = [i for i in lista_archivos if os.path.getsize(i) > 10000]

    # Crear un dataframe vacío
    df = pd.DataFrame(columns=["Archivo PDF", "CUIT del Emisor" , "COD" , "Punto de Venta", "Número de Factura", "Fecha", "Desde" , "Hasta"])

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
            Desde = re.search(r"Desde: (\d+/\d+/\d+)", texto)
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
            df = pd.concat([df, pd.DataFrame([[Archivo, Cod, CUIT, punto_venta, numero_factura, fecha, Desde, Hasta]], columns=["Archivo PDF", "COD" , "CUIT del emisor" , "Punto de Venta", "Número de Factura", "Fecha", "Desde" , "Hasta"])], ignore_index=True)

    # Transformar las columnas 'COD' , 'CUIT del emisor' , 'Punto de Venta' y 'Número de Factura' a int
    df["COD"] = df["COD"].astype(int)
    df["CUIT del emisor"] = df["CUIT del emisor"].astype(np.int64)
    df["Punto de Venta"] = df["Punto de Venta"].astype(int)
    df["Número de Factura"] = df["Número de Factura"].astype(int)

    # Crear una columna 'AUX' con el 'COD' , 'CUIT del emisor' , el 'Punto de Venta' y 'Número de Factura'
    df["AUX"] = df["COD"].astype(str) + "-" + df["Punto de Venta"].astype(str) + "-" + df["Número de Factura"].astype(str)

    # Exportar el dataframe a un archivo csv
    df.to_excel("Datos de Facturas PDF.xlsx", index=False)

    End = time.time()

    Tiempo_Total = End - Start

    showinfo("Extracción de datos", f"Se extrajeron los datos de {len(lista_archivos)} archivos en {Tiempo_Total} segundos")

    return df

    

def Control(MCpath: str , PDFPath: str):
    '''
    Controla los datos de los archivos de 'Mis Comprobantes' con las escalas de categorías de AFIP

    Parameters
    ----------
    MCpath : str
        Path de la carpeta donde se encuentran los archivos de 'Mis Comprobantes'
    PDFPath : str
        Path de la carpeta donde se encuentran los PDF de las facturas
    '''
    # Mostrar mensaje de inicio
    showinfo("Control de datos", "Se iniciará el control de los datos de los archivos de 'Mis Comprobantes' con las escalas de categorías de AFIP y los PDF de las facturas")

    # Leer Excel con las tablas de las escalas
    Categorias = pd.read_excel('Categorias.xlsx')

    # Leer la celda 'A2' de la hoja 'Rango de Fechas' y guardarla en la variable 'fecha_inicial' en formato datetime
    fecha_inicial = pd.read_excel('Categorias.xlsx', sheet_name='Rango de Fechas', header=None, skiprows=1, usecols=[0]).iloc[0,0]
    fecha_inicial = pd.to_datetime(fecha_inicial , format='%d/%m/%Y')
    # leer la celda 'B2' en fomato fecha
    fecha_final = pd.read_excel('Categorias.xlsx', sheet_name='Rango de Fechas', header=None, skiprows=1, usecols=[1]).iloc[0,0]
    fecha_final = pd.to_datetime(fecha_final , format='%d/%m/%Y')

    # Preguntar por el Excel con los Archivos de 'Mis Comprobantes'
    Archivos = os.listdir(MCpath)

    # Agregar el directorio a cada archivo de la lista
    Archivos = [MCpath + "/" + i for i in Archivos]

    # Filtrar la lista para que solo queden los archivos Excel
    Archivos = [i for i in Archivos if i.endswith(".xlsx")]

    # Crear un DataFrame vacio para guardar los datos consolidados
    Consolidado = pd.DataFrame()

    # Leer cada uno de los archivos de 'Mis Comprobantes' y concat en el DataFrame Consolidado
    # Consolidar archivos y renombrar columnas
    # consolidadar columnas
    for f in Archivos:
        #Si el existe el archivo, leerlo
        if os.path.isfile(f):  
            data = pd.read_excel(f, header = None, skiprows=2 , )
            # si el datsaframe esta vacio, no hacer nada
            if len(data) > 0:

                # Crear la columna 'Archivo' con el ultimo elemento de 'f' separado por "/"
                data['Archivo'] = f.split("/")[-1]
                #data['Archivo'] = f.str.split("/")[-1]
                data['CUIT Cliente'] = data["Archivo"].str.split("-").str[3].str.strip().astype(np.int64)
                data['Fin CUIT'] = data["Archivo"].str.split("-").str[0].str.strip().astype(np.int64)
                data['Cliente'] = data["Archivo"].str.split("-").str[-1].str.strip().replace('.xlsx','', regex=True)
                Consolidado = pd.concat([Consolidado , data])
            
    # Renombrar columnas
    Consolidado.columns = [ 'Fecha' , 'Tipo' , 'Punto de Venta' , 'Número Desde' , 'Número Hasta' , 'Cód. Autorización' , 'Tipo Doc. Receptor' , 'Nro. Doc. Receptor/Emisor' , 'Denominación Receptor/Emisor' , 'Tipo Cambio' , 'Moneda' , 'Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total' , 'Archivo' , 'CUIT Cliente' , 'Fin CUIT' , 'Cliente']

    #Eliminar las columas 'Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA'
    Consolidado.drop(['Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA'], axis=1, inplace=True)

    #Cambiar de signo si es una Nota de Crédito
    Consolidado.loc[Consolidado["Tipo"].str.contains("Nota de Crédito"), ['Imp. Total']] *= -1

    #Crear columna de 'MC' con los valores 'archivo' que van desde el caracter 5 al 8 en la Consolidado
    Consolidado['MC'] = Consolidado['Archivo'].str.split("-").str[1].str.strip()


    Info_Facturas_PDF = Extraer_PDF_info(PDFpath=PDFPath)


    Consolidado['Tipo'] = Consolidado['Tipo'].str.split(" ").str[0].str.strip().astype(int)
    
    Consolidado['AUX'] = Consolidado['Tipo'].astype(str) + "-" + Consolidado['Punto de Venta'].astype(str) + "-" + Consolidado['Número Desde'].astype(str)

    # Merge con la tabla Info_Facturas_PDF 
    Consolidado = pd.merge(Consolidado , 
                           Info_Facturas_PDF[['AUX' , 'Desde' , 'Hasta' , 'Archivo PDF']] , 
                           how='left' , 
                           left_on='AUX' , 
                           right_on='AUX')

    # Crear la columna 'Cruzado' con valores 'Si' o 'No' dependiendo si se cruzó o no la información
    Consolidado['Cruzado'] = np.nan
    Consolidado.loc[Consolidado['Archivo PDF'].notnull() , 'Cruzado'] = 'Si'
    Consolidado.loc[Consolidado['Archivo PDF'].isnull() , 'Cruzado'] = 'No'

    # Si las columnas 'Desde' y 'Hasta' son NaN entonces Eliminar todas filas donde la columna 'Fecha' no se encuentre entre el rango de fechas iniciales y finales
    ##########Consolidado = Consolidado[(Consolidado['Fecha'] >= fecha_inicial) & (Consolidado['Fecha'] <= fecha_final) & (Consolidado['Desde'].isnull()) & (Consolidado['Hasta'].isnull())]

    # Si las columnas 'Desde' y 'Hasta' son vacíos entonces toman el valor de 'Fecha'
    Consolidado['Desde'] = Consolidado['Desde'].fillna(Consolidado['Fecha'])
    Consolidado['Hasta'] = Consolidado['Hasta'].fillna(Consolidado['Fecha'])
    Consolidado.loc[Consolidado['Desde'] == "", 'Desde'] = Consolidado['Fecha']
    Consolidado.loc[Consolidado['Hasta'] == "", 'Hasta'] = Consolidado['Fecha']

    # Transformar las columnas 'FECHA EMISION' y 'HASTA' en formato datetime
    Consolidado['Desde'] = pd.to_datetime(Consolidado['Desde'], format='%d/%m/%Y')
    Consolidado['Hasta'] = pd.to_datetime(Consolidado['Hasta'], format='%d/%m/%Y')
    Consolidado['Fecha'] = pd.to_datetime(Consolidado['Fecha'], format='%d/%m/%Y')

    # Crear columna auxiliar con la diferencia entre fecha_inicial y 'FECHA EMISION' en dias
    Consolidado['Diferencia_inicial'] = Consolidado['Desde'] - fecha_inicial
    Consolidado['Diferencia_inicial'] = Consolidado['Diferencia_inicial'].dt.days

    # Crear columna auxiliar con la diferencia entre fecha_final y 'HASTA'
    Consolidado['Diferencia_final'] =  fecha_final - Consolidado['Hasta']
    Consolidado['Diferencia_final'] = Consolidado['Diferencia_final'].dt.days

    # Calcular los dias de diferencia entre 'FECHA EMISION' y 'HASTA'
    Consolidado['Dias de facturación'] = Consolidado['Hasta'] - Consolidado['Desde'] 
    Consolidado['Dias de facturación'] = Consolidado['Dias de facturación'].dt.days +1

    # Crear una columa 'Días Efectivos' con el valor de la columna 'Dias de facturación'
    Consolidado['Días Efectivos'] = Consolidado['Dias de facturación']

    # si la columna 'Diferencia_inicial' es negativa, entonces al valor de la columna 'Días Efectivos' se le resta el valor de la columna 'Diferencia_inicial'
    Consolidado.loc[Consolidado['Diferencia_inicial'] < 0, 'Días Efectivos'] = Consolidado['Días Efectivos'] + Consolidado['Diferencia_inicial']

    # si la columna 'Diferencia_final' es negativa, entonces al valor de la columna 'Días Efectivos' se le resta el valor de la columna 'Diferencia_final'
    Consolidado.loc[Consolidado['Diferencia_final'] < 0, 'Días Efectivos'] = Consolidado['Días Efectivos'] + Consolidado['Diferencia_final']

    # Crear una columna 'Importe por día' con el valor de la columna 'Imp. Total' dividido entre el valor de la columna 'Dias de facturación'
    Consolidado['Importe por día'] = Consolidado['Imp. Total'] / Consolidado['Dias de facturación']

    # Multiplicar el valor de la columna 'Importe por día' por el valor de la columna 'Días Efectivos' y guardar el resultado en la columna 'Importe Prorrateado'
    Consolidado['Importe Prorrateado'] = Consolidado['Importe por día'] * Consolidado['Días Efectivos']

    # Volver a mostrar las columnas 'Desde', 'Hasta' y 'Fecha' en formato fecha
    Consolidado['Desde'] = Consolidado['Desde'].dt.strftime('%d/%m/%Y')
    Consolidado['Hasta'] = Consolidado['Hasta'].dt.strftime('%d/%m/%Y')
    Consolidado['Fecha'] = Consolidado['Fecha'].dt.strftime('%d/%m/%Y')

    # Contar las cantidades de 'No' en la columna 'Cruzado' y guardar el resultado en la variable 'No_Cruzado'
    No_Cruzado = Consolidado['Cruzado'].value_counts()['No'] 


    #Crear Tabla dinámica con los totales de las columnas  'Importe Prorrateado' por 'Archivo'
    TablaDinamica = pd.pivot_table(Consolidado, values=['Importe Prorrateado' , 'Tipo'], index=['Cliente' , 'MC'], aggfunc={'Importe Prorrateado': np.sum , 'Tipo': 'count'})


    # Renombrar la columna 'Tipo' por 'Cantidad de Comprobantes' de la TablaDinamica1 , TablaDinamica2 y TablaDinamica3
    TablaDinamica.rename(columns={'Tipo': 'Cantidad de Comprobantes'}, inplace=True)

    # Buscar el valor de 'Importe Prorrateado' en la escala de categorias donde el valor esta en 'Ingresos brutos'
    TablaDinamica['Ingresos brutos máximos por la categoría'] = TablaDinamica['Importe Prorrateado'].apply(lambda x: Categorias.loc[Categorias['Ingresos brutos'] >= x, 'Ingresos brutos'].iloc[0])

    #Buscar la 'Categoría' en la escala de categorias donde el valor esta en 'Ingresos brutos máximos por la categoría'
    TablaDinamica['Categoría'] = TablaDinamica['Importe Prorrateado'].apply(lambda x: Categorias.loc[Categorias['Ingresos brutos'] >= x, 'Categoria'].iloc[0])

    # Exportar el Consolidado y la Tabla Dinámica a un archivo de Excel
    Archivo_final = pd.ExcelWriter('Reporte Recategorizaciones de Monotributistas.xlsx', engine='openpyxl')
    TablaDinamica.to_excel(Archivo_final, sheet_name='Tabla Dinámica', index=True)
    Consolidado.to_excel(Archivo_final, sheet_name='Consolidado', index=False)
    Archivo_final.close()

    #Guardar el archivo
    #Archivo_final.save()

    #Mostrar mensaje de finalización
    showinfo(title="Finalizado", message=f"El archivo se ha generado correctamente.\n \nCantidad de Facturas no cruzados: {No_Cruzado}")

if __name__ == "__main__":
    Control()