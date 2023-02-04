import pandas as pd
import numpy as np
import os

print('''
XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
+XXXXXXXXXXXXXXXXXXXXXXX/    /XXXXXXXXXXXXXXXXXXXXXXXXXXXXX
XXXXXXXXXXXXXXXXXXXXXXX/    /XX/  \XXXXXXXXXXXXXXXXXXXXXXXX
XXXXXXXXXXXXXXXXXXXXXX/    /XX/    \XXXXXXXXXXXXXXXXXXXXXXX
XXXXXXXXXXXXXXXXXXXXX/    /XX/      \XXXXXXXXXXXXXXXXXXXXXX
XXXXXXXXXXXXXXXXXXXX/    /XX/        \XXXXXXXXXXXXXXXXXXXXX
XXXXXXXXXXXXXXXXXXX/    /XX/    /\    \XXXXXXXXXXXXXXXXXXXX
XXXXXXXXXXXXXXXXXX/    /XX/    /XX\    \XXXXXXXXXXXXXXXXXXX
XXXXXXXXXXXXXXXXX/    /XX/    /XXXX\    \XXXXXXXXXXXXXXXXXX
XXXXXXXXXXXXXXXX/    /XX/    /XXXXXX\    \XXXXXXXXXXXXXXXXX
XXXXXXXXXXXXXXX/    /XX/    /XXXXXXXX\    \XXXXXXXXXXXXXXXX
XXXXXXXXXXXXXX/    /XX/    /XXXXXXXXXX\    \XXXXXXXXXXXXXXX
XXXXXXXXXXXXX/    /XX/    /XXXX\    \XX\    \XXXXXXXXXXXXXX
XXXXXXXXXXXX/    /XX/    /XXXXXX\    \XX\    \XXXXXXXXXXXXX
XXXXXXXXXXX/    /XX/    /XXXXXXXX\    \XX\    \XXXXXXXXXXXX
XXXXXXXXXX/    /XX/    /XXXXXXXXXX\    \XX\    \XXXXXXXXXXX
XXXXXXXXX/    /XXXXXXXXXXXXXXXXXXXX\    \XX\    \XXXXXXXXXX
XXXXXXXX/    /XXXXXXXXXXXXXXXXXXXXXX\    \XX\    \XXXXXXXXX
XXXXXXX/                          \XX\    \XX\    \XXXXXXXX
XXXXXX/                            \XX\    \XX\    \XXXXXXX
XXXXX/    /XXXXXXXXXXXXXXXXXXXXXXXXXXXX\    \XX\    \XXXXXX
XXXX/    /XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX\    \XX\    \XXXXX
XXX/    /XX/                                  \XX\    \XXXX
XX/    /XX/                                    \XX\    \XXX
X/    /XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX\    \XX\    \XX
/    /XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX\    \XX\    \X
XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
X=X                                                     X=X
X=X                     Versión 1.0                     X=X
X=X                                                     X=X
X=X Por Agustín Bustos Piasentini                       X=X
X=X bustos-agustin@hotmail.com                          X=X
X=X agustin.bustos.p@gmail.com                          X=X
X=X                                                     X=X
XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


''')

DirectorioObjetivo = "Consolidar"
path = os.getcwd()
files = os.listdir(DirectorioObjetivo)

############ Crear tabla a donde consolidar ##########

TablaBase = pd.DataFrame()


############# loop del consolidado ###################

for f in files:
    data = pd.read_excel((DirectorioObjetivo + "/" + f), header = None, skiprows=2 , )
    data['Archivo'] = f
    #data['CUIT Cliente'] = data["Archivo"].str[19:30].astype(np.int64)
    #data['Fin CUIT'] = data["CUIT Cliente"].astype(str).str[-1].astype(int)
    #TablaBase = TablaBase.append(data)
    TablaBase = pd.concat([TablaBase , data])

del f, data, files, path, DirectorioObjetivo

############## Renombrar columnas ####################

#Renombrar columnas
TablaBase.columns = [ 'Fecha' , 'Tipo' , 'Punto de Venta' , 'Número Desde' , 'Número Hasta' , 'Cód. Autorización' , 'Tipo Doc. Receptor' , 'Nro. Doc. Receptor/Emisor' , 'Denominación Receptor/Emisor' , 'Tipo Cambio' , 'Moneda' , 'Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total' , 'Archivo']

#rellenar con 0 las columnas de 'Imp. Neto Gravado', 'Imp. Neto No Gravado', 'Imp. Op. Exentas', 'IVA', 'Imp. Total'
TablaBase[['Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total']] = TablaBase[['Imp. Neto Gravado', 'Imp. Neto No Gravado', 'Imp. Op. Exentas', 'IVA', 'Imp. Total']].fillna(0)

#Multiplicar 'Imp. Neto Gravado', 'Imp. Neto No Gravado', 'Imp. Op. Exentas', 'IVA' y 'Imp. Total' por 'Tipo Cambio' cuando la moneda es distinta a '$'
#TablaBase.loc[TablaBase['Moneda'] != '$', ['Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total']] *= TablaBase.loc[TablaBase['Moneda'] != '$', ['Tipo Cambio']]
TablaBase['Imp. Neto Gravado'] *= TablaBase['Tipo Cambio']
TablaBase['Imp. Neto No Gravado'] *= TablaBase['Tipo Cambio']
TablaBase['Imp. Op. Exentas'] *= TablaBase['Tipo Cambio']
TablaBase['IVA'] *= TablaBase['Tipo Cambio']
TablaBase['Imp. Total'] *= TablaBase['Tipo Cambio']

#Cambiar de signo si es una Nota de Crédito
TablaBase.loc[TablaBase["Tipo"].str.contains("Nota de Crédito"), ['Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total']] *= -1


#Crear Tabla dinámica con los totales de las columnas  'Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total' por 'Archivo'
TablaDinamica = pd.pivot_table(TablaBase, values=['Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total'], index=['Archivo'], aggfunc=np.sum)

####### Exportar #####################################
Archivo_final = pd.ExcelWriter('Consolidado.xlsx', engine='openpyxl')
TablaBase.to_excel(Archivo_final, sheet_name="Consolidado" , index=False)
#Exportar Tabla Dinámica a la hoja 'TD 'Consolidado.xlsx'
TablaDinamica.to_excel(Archivo_final, sheet_name="TD" , index=True)
Archivo_final.save()

print('Terminado')