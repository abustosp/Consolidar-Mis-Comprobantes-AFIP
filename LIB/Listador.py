import os
from tkinter.filedialog import askdirectory

def Listador():

    # Listar todos los archivos de una carpeta y todas sus subcarpetas
    Ruta = askdirectory(title="Seleccionar carpeta")
    Archivos = []
    for root, dirs, files in os.walk(Ruta):
        for file in files:
            if file.endswith(".xlsx"):
                Archivos.append(os.path.join(root, file))
            
    #Exportar Archivos a un TXT
    with open('Archivos.txt', 'w') as f:
        for item in Archivos:
            f.write("%s\n" % item)

if __name__ == '__main__':
    Listador()
