import pandas as pd
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showinfo
import numpy as np
import os

def Control():

    # Leer Excel con las tablas de las escalas
    Categorias = pd.read_excel('Categorias.xlsx')

    # Leer la celda 'A2' de la hoja 'Rango de Fechas' y guardarla en la variable 'fecha_inicial' en formato datetime
    fecha_inicial = pd.read_excel('Categorias.xlsx', sheet_name='Rango de Fechas', header=None, skiprows=1, usecols=[0]).iloc[0,0]
    fecha_inicial = pd.to_datetime(fecha_inicial , format='%d/%m/%Y')
    # leer la celda 'B2' en fomato fecha
    fecha_final = pd.read_excel('Categorias.xlsx', sheet_name='Rango de Fechas', header=None, skiprows=1, usecols=[1]).iloc[0,0]
    fecha_final = pd.to_datetime(fecha_final , format='%d/%m/%Y')

    # Preguntar por el Excel con los Archivos de 'Mis Comprobantes'
    Archivos = askopenfilename(title="Seleccione el Excel con las ubicaciones Archivos de 'Mis Comprobantes'")

    # Leer el Excel con los Archivos de 'Mis Comprobantes'
    Archivos = pd.read_excel(Archivos) 

    # transformar la primer columna en una lista
    Archivos = Archivos.iloc[:,0].tolist()

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

                #Leer la primer celde de cada archivo y guardarla en la columna 'Primera celda'
                data['Primera celda'] = pd.read_excel(f, header=None).iloc[0,0]

                #data['Archivo'] = f.str.split("/")[-1]
                data['CUIT Cliente'] = data["Primera celda"].str.split(" ").str[-1].str.strip().astype(np.int64)
                
                #Crear columna 'Fin CUIT' con el ultimo caracter de 'CUIT Cliente'
                data['Fin CUIT'] = data['CUIT Cliente'].astype(str).str[-1].astype(np.int64)
                #data['Cliente'] = data["Archivo"].str.split("-").str[-1].str.strip().replace('.xlsx','', regex=True)
                Consolidado = pd.concat([Consolidado , data])
            
    # Renombrar columnas
    Consolidado.columns = [ 'Fecha' , 'Tipo' , 'Punto de Venta' , 'Número Desde' , 'Número Hasta' , 'Cód. Autorización' , 'Tipo Doc. Receptor' , 'Nro. Doc. Receptor/Emisor' , 'Denominación Receptor/Emisor' , 'Tipo Cambio' , 'Moneda' , 'Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total' , 'Archivo' , 'Primera celda' , 'CUIT Cliente' , 'Fin CUIT']

    #Eliminar las columas 'Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA'
    Consolidado.drop(['Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA'], axis=1, inplace=True)

    #Multiplicar Total tipo de cambio
    Consolidado['Imp. Total'] *= Consolidado['Tipo Cambio']   

    #Cambiar de signo si es una Nota de Crédito
    Consolidado.loc[Consolidado["Tipo"].str.contains("Nota de Crédito"), ['Imp. Total']] *= -1

    #Crear columna de 'MC' con los valores 'Primera celda' que corresponden al tercer elemento del nombre del archivo
    Consolidado['MC'] = Consolidado['Primera celda'].str.split(" ").str[2].str.strip()

    # Transformar la columna 'Fecha' en formato datetime
    Consolidado['Fecha'] = pd.to_datetime(Consolidado['Fecha'] , format='%d/%m/%Y')

    # Eliminar todas las filas que no esten en el rango de fechas
    Consolidado = Consolidado[(Consolidado['Fecha'] >= fecha_inicial) & (Consolidado['Fecha'] <= fecha_final)]

    # Transformar nuevamente la columna 'Fecha' en formato fecha de excel
    Consolidado['Fecha'] = Consolidado['Fecha'].dt.strftime('%d/%m/%Y')

    #Crear Tabla dinámica con los totales de las columnas  'Imp. Total' por 'Archivo'
    TablaDinamica = pd.pivot_table(Consolidado, values=['Imp. Total' , 'Tipo'], index=['CUIT Cliente' , 'MC'], aggfunc={'Imp. Total': np.sum , 'Tipo': 'count'})


    # Renombrar la columna 'Tipo' por 'Cantidad de Comprobantes' de la TablaDinamica1 , TablaDinamica2 y TablaDinamica3
    TablaDinamica.rename(columns={'Tipo': 'Cantidad de Comprobantes'}, inplace=True)

    # Buscar el valor de 'Imp. Total' en la escala de categorias donde el valor esta en 'Ingresos brutos'
    TablaDinamica['Ingresos brutos máximos por la categoría'] = TablaDinamica['Imp. Total'].apply(lambda x: Categorias.loc[Categorias['Ingresos brutos'] >= x, 'Ingresos brutos'].iloc[0])

    #Buscar la 'Categoría' en la escala de categorias donde el valor esta en 'Ingresos brutos máximos por la categoría'
    TablaDinamica['Categoría'] = TablaDinamica['Imp. Total'].apply(lambda x: Categorias.loc[Categorias['Ingresos brutos'] >= x, 'Categoria'].iloc[0])

    # Exportar
    Archivo_final = pd.ExcelWriter('Control de Monotributistas.xlsx', engine='openpyxl')
    Consolidado.to_excel(Archivo_final, sheet_name="Consolidado" , index=False)

    #Exportar Tabla Dinámica a la hoja 'TD'
    TablaDinamica.to_excel(Archivo_final, sheet_name="TD" , index=True , merge_cells=False)

    #Guardar el archivo
    Archivo_final.save()

    #Mostrar mensaje de finalización
    showinfo(title="Finalizado", message="El archivo se ha generado correctamente")

if __name__ == "__main__":
    Control()