import pandas as pd
import numpy as np
import os
import customtkinter as ctk
from tkinter import filedialog
from LIB.Consolidador import Consolidador
from LIB.Listador import Listador



###### TKinter #############################################

# Crea la ventana principal
ventana = ctk.CTk()
ventana.title("Consolidador 3.0")
ventana.geometry("400x200")

# Crea una etiqueta para mostrar la lista de archivos
etiqueta = ctk.CTkLabel(ventana, text="")
etiqueta.pack()


#Crea un botón para ejecutar el listador
boton1 = ctk.CTkButton(ventana, text="Listador de Excels de carpetas y subcarpetas", command=Listador)
boton1.pack(pady=10)

#Crea un botón para seleccionar la carpeta desde una ventana del explorador de archivos
boton2 = ctk.CTkButton(ventana, text="Seleccionar Excel", command=Consolidador)
boton2.pack(pady=10)

#Crear un botón para salir de la aplicación al lado del botón anterior
boton3 = ctk.CTkButton(ventana, text="Salir", command=ventana.quit)
boton3.pack(pady=10)



# Inicia el bucle principal de la ventana
ventana.mainloop()