import pandas as pd
import glob
import os
"""este codigo unifica los dos archivos bajados de sap.
las columnas "E y "F" las trae vacias.
Columna AM entregado/retirado los elimina.
Columna AN elimina los datos distintos del codigo iata
"""

def borrar():
    while True:
        pregunta = input("queres borrar los archivos MHTML?: S/N ").lower()
        if pregunta == "s":
            encontrar = glob.glob("*MHTML")
            if not encontrar:
                print("no hay archivos MHTML")
            else:
                for archivo in encontrar:
                    os.remove(archivo)
            break
        if pregunta == "n":
            print("no se borraron archivos")
            break
        else:
            print("Ingresa 's' o 'n'")

def ejecutarExcelFinalizado(iata):
    os.startfile(f"archivoUnificado{iata}.xlsx")


while True:
    iata = input("Ingresa el codigo IATA: ").upper()
    if len(iata) == 3:
        break
    else:
        print("error intente nuevametne")
encontrar = glob.glob("*xlsx")
lista = []
for archivo in encontrar:
    leer = pd.read_excel(archivo)
    lista.append(leer)
df = pd.concat(lista, ignore_index = True)

"""manipulacion de datos"""
df = df[df["Motivo Descripción"] != "Retirado"]
df = df[df["Motivo Descripción"] != "Entregado"]
df = df[df["Destino"] == iata]
df["Distrito Destino"] = ""
df["Provincia"] = ""
print(df)
#df.to_excel(f"archivoUnificado{iata}.xlsx", index = False)
#borrar()
#ejecutarExcelFinalizado(iata)