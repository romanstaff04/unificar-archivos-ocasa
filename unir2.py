import pandas as pd
import glob
import os
"""este codigo unifica los dos archivos bajados de sap.
las columnas "E y "F" las trae vacias.
Columna AM entregado/retirado los elimina.
Columna AN elimina los datos distintos del codigo iata
filtra por peso mayor o igual a 200kg y 503 columna "s"
filtra por volumen mayor o igual a 1.5 y 503 columna "s"
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

def manipularDatos(df):
    df.loc[df["Peso del objeto"] >= 200, "Ruta Virtual"] = 503
    df.loc[df["Volumen"] >= 1.5, "Ruta Virtual"] = 503
    df = df[df["Motivo Descripción"] != "Retirado"]
    df = df[df["Motivo Descripción"] != "Entregado"]
    df = df[df["Destino"] == iata]
    df["Distrito Destino"] = ""
    df["Provincia"] = ""
    return df

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

df = manipularDatos(df)

df.to_excel(f"archivoUnificado{iata}.xlsx", index = False)
borrar()
ejecutarExcelFinalizado(iata)