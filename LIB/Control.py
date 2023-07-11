import pandas as pd
from tkinter.messagebox import showinfo
import numpy as np
import os
from LIB.Extractor_Facturas import Extraer_PDF_info

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

    #Multiplicar Total tipo de cambio
    Consolidado['Imp. Total'] *= Consolidado['Tipo Cambio']   

    #Cambiar de signo si es una Nota de Crédito
    Consolidado.loc[Consolidado["Tipo"].str.contains("Nota de Crédito"), ['Imp. Total']] *= -1

    #Crear columna de 'MC' con los valores 'archivo' que van desde el caracter 5 al 8 en la Consolidado
    Consolidado['MC'] = Consolidado['Archivo'].str.split("-").str[1].str.strip()


    Info_Facturas_PDF = Extraer_PDF_info(PDFpath=PDFPath)


    # Transformar la columna 'Fecha' en formato datetime
    Consolidado['Fecha'] = pd.to_datetime(Consolidado['Fecha'] , format='%d/%m/%Y')


    Consolidado['Tipo'] = Consolidado['Tipo'].str.split(" ").str[0].str.strip().astype(int)

    
    Consolidado['AUX'] = Consolidado['Tipo'].astype(str) + "-" + Consolidado['Punto de Venta'].astype(str) + "-" + Consolidado['Número Desde'].astype(str)

    # Merge con la tabla Info_Facturas_PDF 
    Consolidado = pd.merge(Consolidado , 
                           Info_Facturas_PDF[['AUX' , 'Desde' , 'Hasta']] , 
                           how='left' , 
                           left_on='AUX' , 
                           right_on='AUX')

    # Si las columnas 'Desde' y 'Hasta' son NaN entonces Eliminar todas filas donde la columna 'Fecha' no se encuentre entre el rango de fechas iniciales y finales
    ##########Consolidado = Consolidado[(Consolidado['Fecha'] >= fecha_inicial) & (Consolidado['Fecha'] <= fecha_final) & (Consolidado['Desde'].isnull()) & (Consolidado['Hasta'].isnull())]





    # Transformar nuevamente la columna 'Fecha' en formato fecha de excel
    Consolidado['Fecha'] = Consolidado['Fecha'].dt.strftime('%d/%m/%Y')

    #Crear Tabla dinámica con los totales de las columnas  'Imp. Total' por 'Archivo'
    TablaDinamica = pd.pivot_table(Consolidado, values=['Imp. Total' , 'Tipo'], index=['Cliente' , 'MC'], aggfunc={'Imp. Total': np.sum , 'Tipo': 'count'})


    # Renombrar la columna 'Tipo' por 'Cantidad de Comprobantes' de la TablaDinamica1 , TablaDinamica2 y TablaDinamica3
    TablaDinamica.rename(columns={'Tipo': 'Cantidad de Comprobantes'}, inplace=True)

    # Buscar el valor de 'Imp. Total' en la escala de categorias donde el valor esta en 'Ingresos brutos'
    TablaDinamica['Ingresos brutos máximos por la categoría'] = TablaDinamica['Imp. Total'].apply(lambda x: Categorias.loc[Categorias['Ingresos brutos'] >= x, 'Ingresos brutos'].iloc[0])

    #Buscar la 'Categoría' en la escala de categorias donde el valor esta en 'Ingresos brutos máximos por la categoría'
    TablaDinamica['Categoría'] = TablaDinamica['Imp. Total'].apply(lambda x: Categorias.loc[Categorias['Ingresos brutos'] >= x, 'Categoria'].iloc[0])

    # Exportar el Consolidado y la Tabla Dinámica a un archivo de Excel
    Archivo_final = pd.ExcelWriter('Reporte Recategorizaciones de monotirbutistas.xlsx', engine='openpyxl')
    TablaDinamica.to_excel(Archivo_final, sheet_name='Tabla Dinámica', index=True)
    Consolidado.to_excel(Archivo_final, sheet_name='Consolidado', index=False)
    Archivo_final.close()

    #Guardar el archivo
    #Archivo_final.save()

    #Mostrar mensaje de finalización
    showinfo(title="Finalizado", message="El archivo se ha generado correctamente")

if __name__ == "__main__":
    Control()