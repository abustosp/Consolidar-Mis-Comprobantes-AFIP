import pandas as pd
import numpy as np
import os
import tkinter as tk
from tkinter import filedialog


###### TKinter #############################################

# Crea la ventana principal
ventana = tk.Tk()
ventana.title("Consolidador 2.0")
ventana.geometry("400x200")

# Crea una etiqueta para mostrar la lista de archivos
etiqueta = tk.Label(ventana, text="")
etiqueta.pack()

#crear una funcion para seleccionar todos los archivos .xlsx de una carpeta
def seleccionar_carpeta():

    # Abre una ventana para seleccionar la carpeta
    ruta = filedialog.askdirectory(parent=ventana)

    # Si se seleccionó una carpeta
    if ruta != "":

        # Obtiene la lista de archivos .xlsx de la carpeta
        archivos = os.listdir(ruta)
        archivos = [x for x in archivos if x.endswith(".xlsx")]

        # Si hay archivos
        if len(archivos) > 0:

            # Muestra la lista de archivos en la etiqueta
            etiqueta["text"] = "Procesando " + str(len(archivos)) + " archivos .xlsx"

            # Crear tabla a donde consolidar
            TablaBase = pd.DataFrame()

            # Consolidar archivos y renombrar columnas
            # consolidadar columnas
            for f in archivos:
                data = pd.read_excel(ruta + "/" + f, header = None, skiprows=2 , )
                data['Archivo'] = f
                data['CUIT Cliente'] = data["Archivo"].str[19:30].astype(np.int64)
                data['Fin CUIT'] = data["CUIT Cliente"].astype(str).str[-1].astype(int)
                #TablaBase = TablaBase.append(data)
                TablaBase = pd.concat([TablaBase , data])
                
            # Renombrar columnas
            TablaBase.columns = [ 'Fecha' , 'Tipo' , 'Punto de Venta' , 'Número Desde' , 'Número Hasta' , 'Cód. Autorización' , 'Tipo Doc. Receptor' , 'Nro. Doc. Receptor/Emisor' , 'Denominación Receptor/Emisor' , 'Tipo Cambio' , 'Moneda' , 'Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total' , 'Archivo' , 'CUIT Cliente' , 'Fin CUIT']

            #Multiplicar por tipo de cambio
            TablaBase['Imp. Neto Gravado'] *= TablaBase['Tipo Cambio']
            TablaBase['Imp. Neto No Gravado'] *= TablaBase['Tipo Cambio']
            TablaBase['Imp. Op. Exentas'] *= TablaBase['Tipo Cambio']
            TablaBase['IVA'] *= TablaBase['Tipo Cambio']
            TablaBase['Imp. Total'] *= TablaBase['Tipo Cambio']   

            #Cambiar de signo si es una Nota de Crédito
            TablaBase.loc[TablaBase["Tipo"].str.contains("Nota de Crédito"), ['Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total']] *= -1

            #Crear Tabla dinámica con los totales de las columnas  'Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total' por 'Archivo'
            TablaDinamica = pd.pivot_table(TablaBase, values=['Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total'], index=['Archivo'], aggfunc=np.sum)

            # Exportar
            Archivo_final = pd.ExcelWriter('Consolidado.xlsx', engine='openpyxl')
            TablaBase.to_excel(Archivo_final, sheet_name="Consolidado" , index=False)

            #Exportar Tabla Dinámica a la hoja 'TD 'Consolidado.xlsx'
            TablaDinamica.to_excel(Archivo_final, sheet_name="TD" , index=True)
            Archivo_final.save()
        else:
            
            # Muestra un mensaje de error en la etiqueta
            etiqueta["text"] = "No hay archivos .xlsx en la carpeta seleccionada"

#Crea un botón para seleccionar la carpeta desde una ventana del explorador de archivos

boton = tk.Button(ventana, text="Seleccionar carpeta", command=seleccionar_carpeta)
boton.pack()

# Inicia el bucle principal de la ventana
ventana.mainloop()