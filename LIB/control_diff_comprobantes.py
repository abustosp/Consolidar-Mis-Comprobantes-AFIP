import pandas as pd
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showinfo
import openpyxl
import LIB.formatos as fmt

def control_diff_comprobantes():
    """
    Función que compara dos archivos excel y muestra las diferencias entre ellos. Generando un archivo excel con las diferencias detalladas, resumidas en un tabla dinámica y los archivos originales.
    """
        
    # Seleccionar los archivos
    archivo1 = askopenfilename(title='Seleccione el primer archivo')
    archivo2 = askopenfilename(title='Seleccione el segundo archivo')
    
    if not archivo1 or not archivo2:
        return

    # Leer los archivos excel
    df1 = pd.read_excel(archivo1, sheet_name='Consolidado')
    df2 = pd.read_excel(archivo2, sheet_name='Consolidado')

    # Editar tipo de comprobantes
    df1['Tipo'] = df1['Tipo'].str.split(' ').str[0].astype(int)
    df2['Tipo'] = df2['Tipo'].str.split(' ').str[0].astype(int)
    
    # Editar moneda
    df1['Moneda'] = df1['Moneda'].str.replace('PES', '$')
    df2['Moneda'] = df2['Moneda'].str.replace('PES', '$')

    # Transformar todos los strings a mayúsculas
    #df1 = df1.map(lambda x: x.upper() if type(x) == str else x)
    #df2 = df2.map(lambda x: x.upper() if type(x) == str else x)
    #df1 = pd.DataFrame(df1)
    #df2 = pd.DataFrame(df2)

    # Crear una tabla dinámica donde las filas sean 'Fin CUIT' 'CUIT Cliente' y 'Archivo', los valores sean sumas de 'Imp. Neto Gravado' 'Imp. Neto No Gravado' 'Imp. Op. Exentas' 'IVA' 'Imp. Total' y se cuente la cantidad de 'Archivo'
    td1 = pd.pivot_table(df1, index=['Fin CUIT', 'CUIT Cliente', 'Archivo'], values=['Imp. Neto Gravado', 'Imp. Neto No Gravado', 'Imp. Op. Exentas', 'IVA', 'Imp. Total', 'MC'], aggfunc={'Imp. Neto Gravado':'sum', 'Imp. Neto No Gravado':'sum', 'Imp. Op. Exentas':'sum', 'IVA':'sum', 'Imp. Total':'sum', 'MC':'count'}, dropna=True)
    td2 = pd.pivot_table(df2, index=['Fin CUIT', 'CUIT Cliente', 'Archivo'], values=['Imp. Neto Gravado', 'Imp. Neto No Gravado', 'Imp. Op. Exentas', 'IVA', 'Imp. Total', 'MC'], aggfunc={'Imp. Neto Gravado':'sum', 'Imp. Neto No Gravado':'sum', 'Imp. Op. Exentas':'sum', 'IVA':'sum', 'Imp. Total':'sum', 'MC':'count'}, dropna=True)

    # Crear una tabla con las difererencias entre df1 y df2
    df_diff = pd.concat([df1, df2])
    df_diff = df_diff.drop_duplicates(keep=False, subset=['CUIT Cliente' , 'Tipo' , 'Punto de Venta', 'Número Desde'])

    # Crear una tabla con las diferencias entre td1 y td2
    td_diff = pd.pivot_table(df_diff, index=['Fin CUIT', 'CUIT Cliente', 'Archivo'], values=['Imp. Neto Gravado', 'Imp. Neto No Gravado', 'Imp. Op. Exentas', 'IVA', 'Imp. Total' , 'MC'], aggfunc={'Imp. Neto Gravado': 'sum', 'Imp. Neto No Gravado': 'sum', 'Imp. Op. Exentas': 'sum', 'IVA': 'sum', 'Imp. Total': 'sum', 'MC':'count'}, dropna=True)

    # resetear los índices de las tablas dinámicas
    td1.reset_index(inplace=True)
    td2.reset_index(inplace=True)
    td_diff.reset_index(inplace=True)

    # Exportar a un Excel
    with pd.ExcelWriter('Resumen MC.xlsx') as writer:
        df1.to_excel(writer, sheet_name='MC1' , index=False)
        df2.to_excel(writer, sheet_name='MC2' , index=False)
        td1.to_excel(writer, sheet_name='TD1', index=False)
        td2.to_excel(writer, sheet_name='TD2', index=False)
        df_diff.to_excel(writer, sheet_name='Diff Detalle', index=False)
        td_diff.to_excel(writer, sheet_name='Diff TD', index=False)
        
    # Aplicar formato al archivo
    hojas = ['MC1', 'MC2', 'TD1', 'TD2', 'Diff Detalle', 'Diff TD']
    hojasMC = ['MC1', 'MC2' , 'Diff Detalle']
    hojasTD = ['TD1', 'TD2', 'Diff TD']
    wb = openpyxl.load_workbook('Resumen MC.xlsx')
        
    # Aplicar formato de moneda    
    for hoja in hojasMC:
        fmt.Aplicar_formato_moneda(wb[hoja], 12, 17)
        
    for hoja in hojasTD:
        fmt.Aplicar_formato_moneda(wb[hoja], 4, 8)
    
    # Aplicar formatos
    for hoja in hojas:
        fmt.Aplicar_formato_encabezado(wb[hoja])
        fmt.Autoajustar_columnas(wb[hoja])
        fmt.Agregar_filtros(wb[hoja])
        
    wb.save('Resumen MC.xlsx')
        
    # Mostrar un mensaje de que el proceso ha finalizado
    showinfo('Proceso finalizado', 'El proceso ha finalizado, revise el archivo Resumen MC.xlsx')
    
if __name__ == '__main__':
    control_diff_comprobantes()