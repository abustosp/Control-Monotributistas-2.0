import os
from tkinter.filedialog import askdirectory
import time
from tkinter.messagebox import showinfo

def Listador():

    StartTime = time.time()

    # Listar todos los archivos de una carpeta y todas sus subcarpetas
    Ruta = askdirectory(title="Seleccionar carpeta")
    Archivos = []
    for root, dirs, files in os.walk(Ruta):
        for file in files:
            if file.endswith(".xlsx"):
                Archivos.append(os.path.join(root, file))

    # Reemplazar todos los '\' por '/' en el listado de archivos
    Archivos = [w.replace('\\', '/') for w in Archivos]
            
    #Exportar Archivos a un TXT
    with open('Archivos.txt', 'w') as f:
        for item in Archivos:
            f.write("%s\n" % item)

    EndTime = time.time()

    Ejecucion = EndTime - StartTime

    #Mostrarlo en 2 decimales
    Ejecucion = round(Ejecucion, 2)

    print("Tiempo de ejecución: " + str(EndTime - StartTime) + " segundos")

    # Crear una ventana de mensaje con "El archivo se ha consolidado correctamente en Segundo"
    showinfo("Consolidador" , "El archivo con la lista se ha creado correctamente en " + str(Ejecucion) + " segundos")

if __name__ == '__main__':
    Listador()
