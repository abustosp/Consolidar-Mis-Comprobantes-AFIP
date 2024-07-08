import os
import pandas as pd
import numpy as np
from tkinter.filedialog import askdirectory
import openpyxl
from tkinter.messagebox import showinfo
import LIB.formatos as fmt

def arreglo_para_sos():

    # seleccionar una carpeta
    ruta = askdirectory(title="Seleccionar la carpeta que posee los archivos de Mis Comprobantes a procesar para SOS-Contador")
    
    if not ruta:
        return

    # Crear una lista con los nombres de los archivos en la carpeta
    archivos = os.listdir(ruta)

    # Filtrar los que no son '.xlsx'
    archivos = [archivo for archivo in archivos if '.xlsx' in archivo]

    # Quitar los archivos que empiezan con '~$'
    archivos = [archivo for archivo in archivos if '~$' not in archivo]

    # reemplazar '.xlsx' por ''
    archivos = [archivo.replace('.xlsx', '') for archivo in archivos]

    # Crear un dataframe con los nombres de los archivos
    Tabla_Archivos = pd.DataFrame(archivos, columns=['Archivo'])

    # Agregar la 'ruta' + '/' al 'Nompre'
    Tabla_Archivos['Archivo con ruta'] = ruta + '/' + Tabla_Archivos['Archivo'] + '.xlsx'

    # Crear una columna que se llame primera celda
    Tabla_Archivos['Primera celda'] = np.nan

    # Leer cada archivo y guardar el contenido de la primera celda en una columna llamada 'Primera celda'
    # for i in range(len(Tabla_Archivos)):
    #     Tabla_Archivos['Primera celda'][i] = pd.read_excel(Tabla_Archivos['Archivo con ruta'][i], header=None).iloc[0,0]
    #     del i

    def obtener_primera_celda(archivo):
        return pd.read_excel(archivo, header=None).iloc[0,0]

    Tabla_Archivos['Primera celda'] = Tabla_Archivos['Archivo con ruta'].apply(obtener_primera_celda)

    #crear la Carpeta 'Procesado' en la 'ruta' si no existe
    if not os.path.exists(ruta + '/Procesado'):
        os.makedirs(ruta + '/Procesado')

    #Leer cada Excel en un Dataframe_Temporal y Exportarlo a la carpeta 'Procesado'
    for i in range(len(Tabla_Archivos)):
        df = pd.read_excel(Tabla_Archivos['Archivo con ruta'][i], skiprows=1)
        # si la columna 'Archivo' contiene ' MCR ' entonces filtrar los datos donde el 'Tipo' contenga ' B'
        if ' MCR ' in Tabla_Archivos['Archivo'][i]:
            df = df[df['Tipo'].str.contains(' B') == False]
        #Exportar el dataframe solo si contiene datos
        if len(df) > 0:
            df.to_excel(ruta + '/Procesado/' + Tabla_Archivos['Archivo'][i] + " - Procesado.xlsx", index=False)


        if len(df) > 0:
            #Al archivo recien creado, agregar la primera celda en la primera fila
            wb = openpyxl.load_workbook(ruta + '/Procesado/' + Tabla_Archivos['Archivo'][i] + " - Procesado.xlsx")
            # Seleccionar hoja activa
            ws = wb.active

            # Agregar Formatos
            #fmt.Agregar_filtros(ws)
            fmt.Aplicar_formato_encabezado(ws)
            fmt.Aplicar_formato_moneda(ws, ColumnaInicial=12 , ColumnaFinal=16)
            fmt.Autoajustar_columnas(ws)

            ws.insert_rows(1)
            ws['A1'] = Tabla_Archivos['Primera celda'][i]
            wb.save(ruta + '/Procesado/' + Tabla_Archivos['Archivo'][i] + " - Procesado.xlsx")
        
    del i, df, wb, ws

    showinfo('Proceso terminado', 'Se han procesado los archivos')
