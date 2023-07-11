#!/usr/bin/python3
import tkinter as tk
import tkinter.ttk as ttk
from LIB.Control import Control

#!/usr/bin/python3
import tkinter as tk
import tkinter.ttk as ttk

def Comando():
    Control(MCpath="BOTS/MC Descargados" , PDFPath="BOTS/Facturas Descargadas")

class ModeloPygubuApp:
    def __init__(self, master=None):
        # build ui
        Toplevel_1 = tk.Tk() if master is None else tk.Toplevel(master)
        Toplevel_1.configure(
            background="#2e2e2e",
            cursor="arrow",
            height=430,
            width=400)
        Toplevel_1.iconbitmap("LIB/ABP-blanco-en-fondo-negro.ico")
        Toplevel_1.minsize(400, 430)
        Toplevel_1.overrideredirect("False")
        Toplevel_1.title("Control de Monotributistas 2.0")
        Label_3 = ttk.Label(Toplevel_1)
        self.img_ABPblancoenfondonegro111 = tk.PhotoImage(
            file="LIB/ABP blanco en sin fondo .png")
        Label_3.configure(
            background="#2e2e2e",
            image=self.img_ABPblancoenfondonegro111)
        Label_3.pack(side="top")
        Label_1 = ttk.Label(Toplevel_1)
        Label_1.configure(
            background="#2e2e2e",
            font="TkDefaultFont",
            foreground="#ffffff",
            justify="center",
            relief="flat",
            takefocus=False,
            text='Control de Monotributistas 2.0 en Base a las Escalas de Facturación.\n\nEl programa utiliza como base el Excel de Categorías (que se debe actualizar manualmente), los Archivos de Mis comprobantes y las Facturas Descargados por los Bots que se encuentran en la carpeta BOTS',
            wraplength=400)
        Label_1.pack(expand="true", side="top")
        Label_2 = ttk.Label(Toplevel_1)
        Label_2.configure(
            background="#2e2e2e",
            foreground="#ffffff",
            justify="center",
            text='por Agustín Bustos Piasentini\nhttps://www.Agustin-Bustos-Piasentini.com.ar/')
        Label_2.pack(expand="true", side="top")
        self.Mensual_XLS = ttk.Button(Toplevel_1)
        self.Mensual_XLS.configure(
            text='Procesar' , command=Comando)
        self.Mensual_XLS.pack(expand="true", pady=4, side="top")
        # self.Button_1 = ttk.Button(Toplevel_1)
        # self.Button_1.configure(
        #     text='Listador de Archivos(Genera Archivos.txt)' , command=Listador)
        # self.Button_1.pack(expand="true", pady=4, side="top")

        # Main widget
        self.mainwindow = Toplevel_1

    def run(self):
        self.mainwindow.mainloop()


if __name__ == "__main__":
    app = ModeloPygubuApp()
    app.run()
