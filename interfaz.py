import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import glob
import os
import sys

# Función para obtener la ruta correcta al archivo, tanto si se ejecuta como .py o como .exe
def obtener_ruta_recurso(nombre_archivo):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, nombre_archivo)
    return os.path.abspath(nombre_archivo)

def borrarMHTML():
    encontrar = glob.glob("*MHTML")
    if encontrar:
        for archivo in encontrar:
            os.remove(archivo)
        messagebox.showinfo("Operación completada", "Se borraron los archivos MHTML correctamente.")


def obtener_archivos():
    return [archivo for archivo in glob.glob("*.xlsx") if archivo != "Canalizador con Provincia y Sucursal UM - ruteo JUNIO2025.xlsx"]

def cargar_datos(iata):
    archivos = obtener_archivos()
    if not archivos:
        return None
    lista = [pd.read_excel(archivo) for archivo in archivos]
    df = pd.concat(lista, ignore_index=True)
    return df

def vaciarGeo(df):
    condicion = (df["Calidad – GEO"] != "ROOFTOP") & (df["Calidad – GEO"] != "APPROXIMATE")
    df.loc[condicion, ["Latitud", "Longitud"]] = ""
    return df

#esta funcion no la estoy usando
def manipularDatosGeneral(df, iata):
    #vaciarGeo(df)
    df.loc[df["Peso del objeto"] >= 200, "Ruta Virtual"] = 503
    df.loc[df["Volumen"] >= 0.7, "Ruta Virtual"] = 503

    df["Distrito Destino"] = ""
    df["Provincia"] = ""

    df = df[
            (df["Motivo Descripción"] != "Retirado") & 
            (df["Motivo Descripción"] != "Entregado") & 
            (df["Destino"] == iata)
        ].copy()
    return df

def manipularDatos(df, iata):
    #vaciarGeo(df)
    #eliminar duplicados en todas las sucursaless
    duplicados = df.duplicated(subset = "Nro. identificación pieza según cliente", keep = "first")
    df.loc[duplicados, "Nro. identificación pieza según cliente"] = df.loc[duplicados, "Equipo"]

    if iata == "TOR":
        df["Distrito Destino"] = ""
        df["Provincia"] = ""
    if iata == "FMA":
        df.loc[df["Nombre Solicitante"] == "TRANSFARMACO S.A.", "Ruta Virtual"] = 502
        df.loc[df["Nombre Solicitante"] == "Fresenius Medical Care Argentina SA", "Ruta Virtual"] = 502
        df.loc[df["Nombre Solicitante"] == "SPP Servicio Puerta a Puerta S.A.", "Ruta Virtual"] = 502
        df.loc[df["Nombre Solicitante"] == "OCASA DISTRIBUCION POSTAL", "Ruta Virtual"] = 1001
        df.loc[df["Nombre Solicitante"] == "ORG COURIER ARG", "Ruta Virtual"] = 700

        df.loc[(df["Ruta Virtual"].isna()) & (df["CP Destino"] != 3600), "Ruta Virtual"] = 504
        df.loc[(df["Ruta Virtual"].isna()) & (df["Peso del objeto"] >= 200), "Ruta Virtual"] = 503
        df.loc[(df["Ruta Virtual"].isna()) & (df["Volumen"] >= 0.7), "Ruta Virtual"] = 503

        df["Distrito Destino"] = ""
        df["Provincia"] = ""

        df = df[
            (df["Motivo Descripción"] != "Retirado") & 
            (df["Motivo Descripción"] != "Entregado") & 
            (df["Destino"] == iata)
        ].copy()

    if iata == "IRJ":
        df.loc[(df["Ruta Virtual"].isna()) & (df["Peso del objeto"] >= 200), "Ruta Virtual"] = 503
        df.loc[(df["Ruta Virtual"].isna()) & (df["Volumen"] >= 0.7), "Ruta Virtual"] = 503
        df.loc[df["Nombre Solicitante"] == "ORG COURIER ARG", "Ruta Virtual"] = 700
        df["Distrito Destino"] = ""
        df["Provincia"] = ""

    if iata == "CRD":
        df.loc[df["Nombre Solicitante"] == "OCASA DISTRIBUCION POSTAL", "Ruta Virtual"] = 502
        df = df[
            (df["Motivo Descripción"] != "Retirado") & 
            (df["Motivo Descripción"] != "Entregado") & 
            (df["Destino"] == iata)
        ].copy()
        df.loc[df["Nombre Solicitante"] == "ORG COURIER ARG", "Ruta Virtual"] = 700
        df.loc[df["Peso del objeto"] >= 200, "Ruta Virtual"] = 503
        df.loc[df["Volumen"] >= 0.7, "Ruta Virtual"] = 503
        df["Distrito Destino"] = ""
        df["Provincia"] = ""

    if iata == "LUQ":
        df = df[
            (df["Motivo Descripción"] != "Retirado") & 
            (df["Motivo Descripción"] != "Entregado") & 
            (df["Destino"] == iata)
        ].copy()
        df.loc[df["Nombre Solicitante"] == "ORG COURIER ARG", "Ruta Virtual"] = 700
        df.loc[df["Peso del objeto"] >= 200, "Ruta Virtual"] = 503
        df.loc[df["Volumen"] >= 0.7, "Ruta Virtual"] = 503
        df["Distrito Destino"] = ""
        df["Provincia"] = ""
    return df

def canalizadorLocalidad(df):
    ruta = obtener_ruta_recurso("Canalizador con Provincia y Sucursal UM - ruteo JUNIO2025.xlsx")
    canalizador = pd.read_excel(ruta)
    canalizador_reducido = canalizador[["CP Destino", "Distrito Destino"]]
    df = df.drop(columns=["Distrito Destino"], errors="ignore")
    merge = pd.merge(df, canalizador_reducido, on="CP Destino", how="left")

    columna_referencia = "Altura"
    if columna_referencia in merge.columns:
        indice = merge.columns.get_loc(columna_referencia) + 1
        columna_valores = merge.pop("Distrito Destino")
        merge.insert(indice, "Distrito Destino", columna_valores)
    else:
        print(f"No se encontró '{columna_referencia}', 'Distrito Destino' se dejó al final.")

    return merge

def canalizadorProvincia(df):
    ruta = obtener_ruta_recurso("Canalizador con Provincia y Sucursal UM - ruteo JUNIO2025.xlsx")
    canalizador = pd.read_excel(ruta)
    canalizador_reducido = canalizador[["CP Destino", "Provincia", "ZONIFICACION"]]

    df = df.drop(columns=["Provincia"], errors="ignore")
    merge = pd.merge(df, canalizador_reducido, on="CP Destino", how="left")

    columna_referencia = "Población"
    if columna_referencia in merge.columns:
        indice = merge.columns.get_loc(columna_referencia) + 1
        columna_valores = merge.pop("Provincia")
        merge.insert(indice, "Provincia", columna_valores)
    else:
        print(f"No se encontró '{columna_referencia}', 'Provincia' se dejó al final.")

    return merge

def procesar(iata):
    df = cargar_datos(iata)
    if df is None:
        messagebox.showerror("Error", "No se encontraron archivos para procesar.")
        return

    df = manipularDatos(df, iata)
    df = canalizadorLocalidad(df)
    df = canalizadorProvincia(df)
    borrarMHTML()

    nombre_salida = f"subirUnigis{iata}.xlsx"
    df.to_excel(nombre_salida, index=False)
    os.startfile(nombre_salida)

def crear_interfaz():
    ventana = tk.Tk()
    ventana.title("Procesador de Ruteo por IATA")
    ventana.geometry("400x200")

    etiqueta = ttk.Label(ventana, text="Seleccione un código IATA:", font=("Arial", 12))
    etiqueta.pack(pady=10)

    opciones_iata = ["FMA", "IRJ", "CRD", "LUQ", "TUC", "RES", "TOR", "REL", "PSS", "CTC", "ROS", "RGL", "FCO", "EQS"]
    seleccion = tk.StringVar()
    combobox = ttk.Combobox(ventana, textvariable=seleccion, values=opciones_iata, state="readonly")
    combobox.pack(pady=5)

    def ejecutar():
        iata = seleccion.get()
        if not iata:
            messagebox.showwarning("Advertencia", "Debe seleccionar un código IATA.")
        else:
            procesar(iata)
            ventana.destroy()

    boton = ttk.Button(ventana, text="Procesar", command=ejecutar)
    boton.pack(pady=20)

    ventana.mainloop()

if __name__ == "__main__":
    crear_interfaz()