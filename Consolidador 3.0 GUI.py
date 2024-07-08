import os
import customtkinter as ctk
from LIB.Consolidador import Consolidador_Excel, Consolidador_Carpetas
from LIB.Listador import Listador
from LIB.Transformar_ZIP_MC import Transformar_ZIP_MC
from tkinter.filedialog import askdirectory
from LIB.filtrar_archivos import filtrar_archivos
from LIB.arreglo_para_sos import arreglo_para_sos
from LIB.control_diff_comprobantes import control_diff_comprobantes
from LIB.control_contenido_mc import control_contenido

def Donaciones():
    # Funcion para redirigir a "https://cafecito.app/abustos"
    os.system("start https://cafecito.app/abustos")
    
def transformar_archivos():
    """
    Función para transformar los archivos ZIP de Mis Comprobantes a Excel.
    
    """
    directorio = askdirectory(title="Seleccionar la carpeta que posee los archivos ZIP de Mis Comprobantes")
    Transformar_ZIP_MC(directorio)
    
def abrir_Excel_consolidador():
    """
    Función para abrir el excel "Lista Consolidador.xlsx"
    
    """
    os.system("start Lista-Consolidador.xlsx")

###### TKinter #############################################

# Crea la ventana principal
ventana = ctk.CTk()
ventana.title("Consolidador 3.0")
ventana.geometry("400x560")

# Crea una etiqueta para mostrar la lista de archivos
etiqueta = ctk.CTkLabel(ventana, text="")
etiqueta.pack()


#Crea un botón para ejecutar el listador
boton1 = ctk.CTkButton(ventana, text="Listador de Excels de carpetas y subcarpetas en Archivos.txt", command=Listador)
boton1.pack(pady=10)

# Abre el Excel "Lista Consolidador.xlsx"
boton7 = ctk.CTkButton(ventana, text="Abrir Excel Lista-Consolidador.xlsx", command=abrir_Excel_consolidador)
boton7.pack(pady=10)

# Crea un botón para transformar los archivos ZIP de Mis Comprobantes a Excel
boton6 = ctk.CTkButton(ventana, text="Transformar ZIP de Mis Comprobantes a Excel", command=transformar_archivos)
boton6.pack(pady=10)

# Crear boton de control de contenido
boton11 = ctk.CTkButton(ventana, text="Control de Contenido de Mis Comprobantes", command=control_contenido)
boton11.pack(pady=10)

# Crear boton de filtrar archivos
boton8 = ctk.CTkButton(ventana, text="Filtrar Archivos", command=filtrar_archivos)
boton8.pack(pady=10)

# Crear boton de arreglo para SOS
boton9 = ctk.CTkButton(ventana, text="Arreglo para SOS-Contador", command=arreglo_para_sos)
boton9.pack(pady=10)

#Crea un botón para seleccionar la carpeta desde una ventana del explorador de archivos
boton2 = ctk.CTkButton(ventana, text="Consolidar Excels de Mis Comprobantes en base a un Excel", command=Consolidador_Excel)
boton2.pack(pady=10)

#Crea un botón para seleccionar la carpeta desde una ventana del explorador de archivos
boton3 = ctk.CTkButton(ventana, text="Consolidar Excels de Mis Comprobantes de una Carpeta", command=Consolidador_Carpetas)
boton3.pack(pady=10)

# Crear un botón para controlar la diferencia de comprobantes
boton10 = ctk.CTkButton(ventana, text="Controlar Diferencia de Mis Comprobantes", command=control_diff_comprobantes)
boton10.pack(pady=10)

# Crear un boton de donación
boton4 = ctk.CTkButton(ventana, text="Donaciones", command=Donaciones)
boton4.pack(pady=10)

#Crear un botón para salir de la aplicación al lado del botón anterior
boton5 = ctk.CTkButton(ventana, text="Salir", command=ventana.quit)
boton5.pack(pady=10)



# Inicia el bucle principal de la ventana
ventana.mainloop()