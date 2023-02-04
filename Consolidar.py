import pandas as pd
import numpy as np
import os

DirectorioObjetivo = "Consolidar"
path = os.getcwd()
files = os.listdir(DirectorioObjetivo)


############ Crear tabla a donde consolidar ##########

TablaBase = pd.DataFrame()


############# loop del consolidado ###################

for f in files:
    data = pd.read_excel((DirectorioObjetivo + "/" + f), header = None, skiprows=2 , )
    data['Archivo'] = f
    data['CUIT Cliente'] = data["Archivo"].str[19:30].astype(np.int64)
    data['Fin CUIT'] = data["CUIT Cliente"].astype(str).str[-1].astype(int)
    #TablaBase = TablaBase.append(data)
    TablaBase = pd.concat([TablaBase , data])


############## Renombrar columnas ####################

TablaBase.columns = [ 'Fecha' , 'Tipo' , 'Punto de Venta' , 'Número Desde' , 'Número Hasta' , 'Cód. Autorización' , 'Tipo Doc. Receptor' , 'Nro. Doc. Receptor/Emisor' , 'Denominación Receptor/Emisor' , 'Tipo Cambio' , 'Moneda' , 'Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total' , 'Archivo' , 'CUIT Cliente' , 'Fin CUIT']


####### Exportar #####################################

TablaBase.to_excel("Consolidado.xlsx" , index=False)

print('Terminado')