import os
import pandas as pd
from tkinter.filedialog import askdirectory, askopenfilename
import shutil

def filtrar_archivos(
    directorio:str=None , 
    periodo:str=None , 
    archivo_con_datos:str=None):
    """Filtrar archivos en base a un archivo de Excel para luego copiarlos en un directorio objetivo.

    ---
    ### Args:
        - directorio (str): directorio donde se encuentran los archivos a copiar 
        - periodo (str): periodo de los archivos a copiar
        - archivo_con_datos (str): archivo de Excel con los datos de los archivos a copiar
    """
    if directorio == None:
        # Seleccionar el directorio donde se encuentran los archivos
        directorio = askdirectory(title="Seleccionar el directorio donde se encuentran los archivos a filtrar")
        
    if periodo == None:
        periodo = "-".join([directorio.split("/")[3][:4], directorio.split("/")[3][4:6]])
        
        
    if archivo_con_datos == None:
        # Archivo de Excel con los datos de los libros que deben ser procesados e importados automáticamente
        #archivo_con_datos = r"F:\UiPath\Libros Ventas y Compras\Libro Ventas y Compras SOS\Importar MC a SOS.xlsx"
        archivo_con_datos = askopenfilename(title="Seleccionar el Excel con los datos de los archivos a filtrar")
        
    if not directorio or not periodo or not archivo_con_datos or directorio == None or periodo == None or archivo_con_datos == None:
        return
        
    
    df = pd.read_excel(archivo_con_datos , sheet_name="Listado")

    # Descartar las filas que contengan valores en la columna 'Importar'
    df = df[df['Importar'].isnull()]
    # resetear el index
    df.reset_index(drop=True, inplace=True)

    # crear el directorio objetivo
    directorio_objetivo = f"{directorio}/Importación Masiva {periodo}"
    os.makedirs(directorio_objetivo , exist_ok=True)

    # Recorrer el dataframe y copiar los archivos en el directorio objetivo
    for i in range(len(df)):
        # Si existe el archivo de la columa MCR/MCE en el directorio copiarlo en el directorio objetivo
        if os.path.exists(f"{directorio}/{df['MCR'][i]}.xlsx"):
            shutil.copy(f"{directorio}/{df['MCR'][i]}.xlsx", f"{directorio_objetivo}/{df['MCR'][i]}.xlsx")
        if os.path.exists(f"{directorio}/{df['MCE'][i]}.xlsx"):
            shutil.copy(f"{directorio}/{df['MCE'][i]}.xlsx", f"{directorio_objetivo}/{df['MCE'][i]}.xlsx")


if __name__ == '__main__':
    filtrar_archivos()