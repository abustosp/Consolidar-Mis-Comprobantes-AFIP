import pandas as pd
import numpy as np
from tkinter import filedialog
import time
import os

def Consolidador():

    StartTime = time.time()

    # #crear una funcion para seleccionar todos los archivos .xlsx de una carpeta
    # def seleccionar_carpeta():

    # Abre una ventana para seleccionar la carpeta
    ruta = filedialog.askopenfile(title="Seleccionar el Excel que posee los Directorios con los archivos a consolidar" , filetypes = (("Excel files","*.xlsx"),("all files","*.*")) )

    archivos = pd.read_excel(ruta.name)

    # Hacer una lista con el la primer columna del excel
    archivos = archivos.iloc[:,0].tolist() 

    TablaBase = pd.DataFrame()

    # Consolidar archivos y renombrar columnas
    # consolidadar columnas
    for f in archivos:
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

    #Crear columna de 'MC' con los valores 'archivo' que van desde el caracter 5 al 8 en la TablaBase
    TablaBase['MC'] = TablaBase['Archivo'].str.split("-").str[1].str.strip()

    #Crear Tabla dinámica con los totales de las columnas  'Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total' por 'Archivo'
    TablaDinamica = pd.pivot_table(TablaBase, values=['Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total' , 'Tipo'], index=['Archivo'], aggfunc={'Imp. Neto Gravado': np.sum , 'Imp. Neto No Gravado': np.sum , 'Imp. Op. Exentas': np.sum , 'IVA': np.sum , 'Imp. Total': np.sum , 'Tipo': 'count'})

    #Crear Tabla dinámica con los totales de las columnas  'Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total' por 'CUIT Cliente'
    TablaDinamica2 = pd.pivot_table(TablaBase, values=['Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total'], index=['CUIT Cliente' , 'MC' , 'Archivo' , 'Tipo'], aggfunc={'Imp. Neto Gravado': np.sum , 'Imp. Neto No Gravado': np.sum , 'Imp. Op. Exentas': np.sum , 'IVA': np.sum , 'Imp. Total': np.sum , 'Tipo': 'count'})

    #Crear Tabla dinámica con los totales de las columnas  'Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total' por 'CUIT Cliente' , 'MC' , 'Archivo' y 'Tipo'
    #TablaDinamica3 = pd.pivot_table(TablaBase, values=['Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total'], index=['CUIT Cliente' , 'MC' , 'Archivo' , 'Tipo'], aggfunc=np.sum)
    #Crear ua tabla dinámica como la anterior pero agregándo la cantidad de registros que conforman el tipo
    #TablaDinamica3 = pd.pivot_table(TablaBase, values=['Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total'], index=['CUIT Cliente' , 'MC' , 'Archivo' , 'Tipo'], aggfunc={'Imp. Neto Gravado' : np.sum , 'Imp. Neto No Gravado' : np.sum , 'Imp. Op. Exentas' : np.sum , 'IVA' : np.sum , 'Imp. Total' : np.sum , 'Tipo' : 'count'})

    # Renombrar la columna 'Tipo' por 'Cantidad de Comprobantes' de la TablaDinamica1 , TablaDinamica2 y TablaDinamica3
    TablaDinamica.rename(columns={'Tipo': 'Cantidad de Comprobantes'}, inplace=True)
    TablaDinamica2.rename(columns={'Tipo': 'Cantidad de Comprobantes'}, inplace=True)
    #TablaDinamica3.rename(columns={'Tipo': 'Cantidad de Comprobantes'}, inplace=True)

    # Exportar
    Archivo_final = pd.ExcelWriter('Consolidado.xlsx', engine='openpyxl')
    TablaBase.to_excel(Archivo_final, sheet_name="Consolidado" , index=False)

    #Exportar Tabla Dinámica a la hoja 'TD' de 'Consolidado.xlsx'
    TablaDinamica.to_excel(Archivo_final, sheet_name="TD" , index=True , merge_cells=False)

    #Exportar Tabla Dinámica a la hoja 'TD2' de 'Consolidado.xlsx'
    TablaDinamica2.to_excel(Archivo_final, sheet_name="TD Cruce" , index=True , merge_cells=False)

    #Exportar Tabla Dinámica a la hoja 'TD por CBTE' de 'Consolidado.xlsx'
    #TablaDinamica3.to_excel(Archivo_final, sheet_name="TD por CBTE" , index=True , merge_cells=False)

    #Guardar el archivo
    Archivo_final.save()

    EndTime = time.time()

    print("Tiempo de ejecución: " + str(EndTime - StartTime) + " segundos")

if __name__ == "__main__":
    Consolidador()
