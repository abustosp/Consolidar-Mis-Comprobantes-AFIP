import pandas as pd
import numpy as np
from zipfile import ZipFile
import os
import openpyxl
from tkinter.filedialog import askdirectory

def Transformar_ZIP_MC(Directorio):
    '''
    Esta función recibe un directorio y transforma los archivos .zip de Mis Comprobantes en archivos .xlsx con el formato correcto para ser importados a la base de datos.

    ### Parámetros:
    - Directorio: Directorio donde se encuentran los archivos .zip de Mis Comprobantes.
    '''

    Archivos = os.listdir(Directorio)

    Zips = [Archivo for Archivo in Archivos if Archivo.endswith(".zip")]

    for Zip in Zips:

        # Obtener el nombre del archivo sin la extensión
        Nombre = Zip.split(".zip")[0]

        # Obtener el CUIT del archivo
        Cuit = Nombre.split("-")[3].strip()

        # Obtener el Tipo de Archivo (Emitidas o Recibidas)
        Tipo = Nombre.split("-")[1].strip()

        # Si el tipo es MCE, cambiarlo a "Mis Comprobantes Emitidos". si es MCR, cambiarlo a "Mis Comprobantes Recibidos"
        if Tipo == "MCE":
            Tipo = "Mis Comprobantes Emitidos"
        elif Tipo == "MCR":
            Tipo = "Mis Comprobantes Recibidos"

        # Obtener el nombre del archivo dentro del zip
        with ZipFile(Directorio + "/" + Zip, 'r') as zip:
            # Listar los archivos dentro del zip
            Archivos = zip.namelist()
            # Obtener el nombre del primer archivo
            Archivo = Archivos[0]
            zip.extract(Archivo, Directorio)

        df = pd.read_csv(Directorio + "\\" + Zip, sep=";", encoding="UTF-8" , decimal=",")

        # Transformar la "Fecha de Emisión" a datetime
        df["Fecha de Emisión"] = pd.to_datetime(df["Fecha de Emisión"], format="%Y-%m-%d")
        # Mostrar como dd/mm/aaaa
        df["Fecha de Emisión"] = df["Fecha de Emisión"].dt.strftime("%d/%m/%Y")

        Diccionario_Tipo = {
            '1':'1 - Factura A',
            '11':'11 - Factura C',
            '13':'13 - Nota de Crédito C',
            '15':'15 - Recibo C',
            '2':'2 - Nota de Débito A',
            '201':'201 - Factura de Crédito electrónica MiPyMEs (FCE) A',
            '203':'203 - Nota de Crédito electrónica MiPyMEs (FCE) A',
            '211':'211 - Factura de Crédito electrónica MiPyMEs (FCE) C',
            '3':'3 - Nota de Crédito A',
            '6':'6 - Factura B',
            '8':'8 - Nota de Crédito B',
    }
        
        # Cambiar el tipo de comprobante por el nombre
        df["Tipo de Comprobante"] = df["Tipo de Comprobante"].astype(str)
        df["Tipo de Comprobante"] = df["Tipo de Comprobante"].map(Diccionario_Tipo)

        # # Ordenar por "Denominación Vendedor" en orden ascendente
        # df.sort_values(by="Denominación Vendedor", ascending=True, inplace=True)

        # Eliminar el archivo
        os.remove(Directorio + "\\" + Archivo)
        # Eliminar el directorio
        #os.rmdir(Directorio)

        df.to_excel(f"{Directorio}/{Nombre}.xlsx", index=False , sheet_name="Sheet1")

        Header = f"{Tipo} - CUIT {Cuit}"

        # mover los datos una fila hacia abajo y en la celda A1 poner el Header
        wb = openpyxl.load_workbook(f"{Directorio}/{Nombre}.xlsx")
        ws = wb.active
        ws.insert_rows(1)
        ws["A1"] = Header
        wb.save(f"{Directorio}/{Nombre}.xlsx")
        wb.close()

if __name__ == "__main__":
    Directorio = askdirectory(title="Seleccionar el directorio donde se encuentran los archivos de MC")
    Transformar_ZIP_MC(Directorio)

