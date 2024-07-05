import os
import customtkinter as ctk
from LIB.Consolidador import Consolidador_Excel, Consolidador_Carpetas
from LIB.Listador import Listador
from LIB.Transformar_ZIP_MC import Transformar_ZIP_MC
from tkinter.filedialog import askdirectory

def Donaciones():
    # Funcion para redirigir a "https://cafecito.app/abustos"
    os.system("start https://cafecito.app/abustos")
    
def transformar_archivos():
    """
    Función para transformar los archivos ZIP de Mis Comprobantes a Excel.
    
    """
    directorio = askdirectory(title="Seleccionar la carpeta que posee los archivos ZIP de Mis Comprobantes")
    Transformar_ZIP_MC(directorio)

###### TKinter #############################################

# Crea la ventana principal
ventana = ctk.CTk()
ventana.title("Consolidador 3.0")
ventana.geometry("400x320")

# Crea una etiqueta para mostrar la lista de archivos
etiqueta = ctk.CTkLabel(ventana, text="")
etiqueta.pack()


#Crea un botón para ejecutar el listador
boton1 = ctk.CTkButton(ventana, text="Listador de Excels de carpetas y subcarpetas", command=Listador)
boton1.pack(pady=10)

#Crea un botón para seleccionar la carpeta desde una ventana del explorador de archivos
boton2 = ctk.CTkButton(ventana, text="Seleccionar Excel", command=Consolidador_Excel)
boton2.pack(pady=10)

#Crea un botón para seleccionar la carpeta desde una ventana del explorador de archivos
boton3 = ctk.CTkButton(ventana, text="Seleccionar Carpeta", command=Consolidador_Carpetas)
boton3.pack(pady=10)

# Crea un botón para transformar los archivos ZIP de Mis Comprobantes a Excel
boton6 = ctk.CTkButton(ventana, text="Transformar ZIP de Mis Comprobantes a Excel", command=transformar_archivos)
boton6.pack(pady=10)

# Crear un boton de donación
boton4 = ctk.CTkButton(ventana, text="Donaciones", command=Donaciones)
boton4.pack(pady=10)

#Crear un botón para salir de la aplicación al lado del botón anterior
boton5 = ctk.CTkButton(ventana, text="Salir", command=ventana.quit)
boton5.pack(pady=10)



# Inicia el bucle principal de la ventana
ventana.mainloop()