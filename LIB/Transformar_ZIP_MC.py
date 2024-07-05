import pandas as pd
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
            '1': '1 - FACTURA A',
            '2': '2 - NOTA DE DÉBITO A',
            '3': '3 - NOTA DE CREDITO A',
            '4': '4 - RECIBOS A',
            '5': '5 - NOTA DE VENTA AL CONTADO A',
            '6': '6 - FACTURA B',
            '7': '7 - NOTA DE DÉBITO B',
            '8': '8 - NOTA DE CREDITO B',
            '9': '9 - RECIBOS B',
            '10': '10 - NOTA DE VENTA AL CONTADO B',
            '11': '11 - FACTURA C',
            '12': '12 - NOTA DE DÉBITO C',
            '13': '13 - NOTA DE CREDITO C',
            '15': '15 - RECIBOS C',
            '16': '16 - NOTA DE VENTA AL CONTADO C',
            '17': '17 - LIQUIDACION DE SERVICIOS PUBLICOS CLASE A',
            '18': '18 - LIQUIDACION DE SERVICIOS PUBLICOS CLASE B',
            '19': '19 - FACTURA DE EXPORTACION',
            '20': '20 - NOTA DE DÉBITO POR OPERACIONES CON EL EXTERIOR',
            '21': '21 - NOTA DE CREDITO POR OPERACIONES CON EL EXTERIOR',
            '22': '22 - FACTURA - PERMISO EXPORTACION SIMPLIFICADO - DTO. 855/97',
            '23': '23 - COMPROBANTES “A” DE COMPRA PRIMARIA PARA EL SECTOR PESQUERO MARITIMO',
            '24': '24 - COMPROBANTES “A” DE CONSIGNACION PRIMARIA PARA EL SECTOR PESQUERO MARITIMO',
            '25': '25 - COMPROBANTES “B” DE COMPRA PRIMARIA PARA EL SECTOR PESQUERO MARITIMO',
            '26': '26 - COMPROBANTES “B” DE CONSIGNACION PRIMARIA PARA EL SECTOR PESQUERO MARITIMO',
            '27': '27 - LIQUIDACION UNICA COMERCIAL IMPOSITIVA CLASE A',
            '28': '28 - LIQUIDACION UNICA COMERCIAL IMPOSITIVA CLASE B',
            '29': '29 - LIQUIDACION UNICA COMERCIAL IMPOSITIVA CLASE C',
            '30': '30 - COMPROBANTES DE COMPRA DE BIENES USADOS',
            '31': '31 - MANDATO - CONSIGNACION',
            '32': '32 - COMPROBANTES PARA RECICLAR MATERIALES',
            '33': '33 - LIQUIDACION PRIMARIA DE GRANOS',
            '34': '34 - COMPROBANTES A DEL APARTADO A INCISO F) R.G. N° 1415',
            '35': '35 - COMPROBANTES B DEL ANEXO I, APARTADO A, INC. F), R.G. N° 1415',
            '36': '36 - COMPROBANTES C DEL Anexo I, Apartado A, INC. F), R.G. N° 1415',
            '37': '37 - NOTA DE DÉBITO O DOCUMENTO EQUIVALENTE QUE CUMPLAN CON LA R.G. N° 1415',
            '38': '38 - NOTA DE CRÉDITO O DOCUMENTO EQUIVALENTE QUE CUMPLAN CON LA R.G. N° 1415',
            '39': '39 - OTROS COMPROBANTES A QUE CUMPLEN CON LA R G 1415',
            '40': '40 - OTROS COMPROBANTES B QUE CUMPLAN CON LA R.G. N° 1415',
            '41': '41 - OTROS COMPROBANTES C QUE CUMPLAN CON LA R.G. N° 1415',
            '43': '43 - NOTA DE CRÉDITO LIQUIDACIÓN UNICA COMERCIAL IMPOSITIVA CLASE B',
            '44': '44 - NOTA DE CRÉDITO LIQUIDACIÓN UNICA COMERCIAL IMPOSITIVA CLASE C',
            '45': '45 - NOTA DE DÉBITO LIQUIDACIÓN UNICA COMERCIAL IMPOSITIVA CLASE A',
            '46': '46 - NOTA DE DÉBITO LIQUIDACIÓN UNICA COMERCIAL IMPOSITIVA CLASE B',
            '47': '47 - NOTA DE DÉBITO LIQUIDACIÓN UNICA COMERCIAL IMPOSITIVA CLASE C',
            '48': '48 - NOTA DE CRÉDITO LIQUIDACIÓN UNICA COMERCIAL IMPOSITIVA CLASE A',
            '49': '49 - COMPROBANTES DE COMPRA DE BIENES NO REGISTRABLES A CONSUMIDORES FINALES',
            '50': '50 - RECIBO FACTURA A RÉGIMEN DE FACTURA DE CRÉDITO',
            '51': '51 - FACTURA M',
            '52': '52 - NOTA DE DÉBITO M',
            '53': '53 - NOTA DE CRÉDITO M',
            '54': '54 - RECIBOS M',
            '55': '55 - NOTA DE VENTA AL CONTADO M',
            '56': '56 - COMPROBANTES M DEL ANEXO I APARTADO A INC F) R.G. N° 1415',
            '57': '57 - OTROS COMPROBANTES M QUE CUMPLAN CON LA R.G. N° 1415',
            '58': '58 - CUENTAS DE VENTA Y LIQUIDO PRODUCTO M',
            '59': '59 - LIQUIDACIONES M',
            '60': '60 - CUENTAS DE VENTA Y LIQUIDO PRODUCTO A',
            '61': '61 - CUENTAS DE VENTA Y LIQUIDO PRODUCTO B',
            '63': '63 - LIQUIDACIONES A',
            '64': '64 - LIQUIDACIONES B',
            '66': '66 - DESPACHO DE IMPORTACIÓN',
            '68': '68 - LIQUIDACIÓN C',
            '70': '70 - RECIBOS FACTURA DE CRÉDITO',
            '80': '80 - INFORME DIARIO DE CIERRE (ZETA) - CONTROLADORES FISCALES',
            '81': '81 - TIQUE FACTURA A',
            '82': '82 - TIQUE FACTURA B',
            '83': '83 - TIQUE',
            '88': '88 - REMITO ELECTRÓNICO',
            '89': '89 - RESUMEN DE DATOS',
            '90': '90 - OTROS COMPROBANTES - DOCUMENTOS EXCEPTUADOS - NOTA DE CRÉDITO',
            '91': '91 - REMITOS R',
            '99': '99 - OTROS COMPROBANTES QUE NO CUMPLEN O ESTÁN EXCEPTUADOS DE LA R.G. 1415 Y SUS MODIF',
            '110': '110 - TIQUE NOTA DE CRÉDITO',
            '111': '111 - TIQUE FACTURA C',
            '112': '112 - TIQUE NOTA DE CRÉDITO A',
            '113': '113 - TIQUE NOTA DE CRÉDITO B',
            '114': '114 - TIQUE NOTA DE CRÉDITO C',
            '115': '115 - TIQUE NOTA DE DÉBITO A',
            '116': '116 - TIQUE NOTA DE DÉBITO B',
            '117': '117 - TIQUE NOTA DE DÉBITO C',
            '118': '118 - TIQUE FACTURA M',
            '119': '119 - TIQUE NOTA DE CRÉDITO M',
            '120': '120 - TIQUE NOTA DE DÉBITO M',
            '201': '201 - FACTURA DE CRÉDITO ELECTRÓNICA MiPyMEs (FCE) A',
            '202': '202 - NOTA DE DÉBITO ELECTRÓNICA MiPyMEs (FCE) A',
            '203': '203 - NOTA DE CRÉDITO ELECTRÓNICA MiPyMEs (FCE) A',
            '206': '206 - FACTURA DE CRÉDITO ELECTRÓNICA MiPyMEs (FCE) B',
            '207': '207 - NOTA DE DÉBITO ELECTRÓNICA MiPyMEs (FCE) B',
            '208': '208 - NOTA DE CRÉDITO ELECTRÓNICA MiPyMEs (FCE) B',
            '211': '211 - FACTURA DE CRÉDITO ELECTRÓNICA MiPyMEs (FCE) C',
            '212': '212 - NOTA DE DÉBITO ELECTRÓNICA MiPyMEs (FCE) C',
            '213': '213 - NOTA DE CRÉDITO ELECTRÓNICA MiPyMEs (FCE) C',
            '331': '331 - LIQUIDACIÓN SECUNDARIA DE GRANOS',
            '332': '332 - CERTIFICACIÓN ELECTRÓNICA (GRANOS)',
            '995': '995 - REMITO ELECTRÓNICO CÁRNICO'
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

