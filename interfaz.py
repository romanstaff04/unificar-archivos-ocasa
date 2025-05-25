import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import glob
import os

def borrarMHTML():
    encontrar = glob.glob("*MHTML")
    for archivo in encontrar:
        os.remove(archivo)

def obtener_archivos():
    return [archivo for archivo in glob.glob("*.xlsx") if archivo != "CANALIZADOR MADRE.xlsx"]

def cargar_datos():
    archivos = obtener_archivos()
    if not archivos:
        return None
    lista = [pd.read_excel(archivo) for archivo in archivos]
    df = pd.concat(lista, ignore_index=True)
    return df


def manipularDatos(df, iata):
    if iata == "FMA":
        #agregar valores
        df.loc[df["Nombre Solicitante"] == "TRANSFARMACO S.A.", "Ruta Virtual"] = 502
        df.loc[df["Nombre Solicitante"] == "Fresenius Medical Care Argentina SA", "Ruta Virtual"] = 502
        df.loc[df["Nombre Solicitante"] == "SPP Servicio Puerta a Puerta S.A.", "Ruta Virtual"] = 502
        df.loc[df["Nombre Solicitante"] == "OCASA DISTRIBUCION POSTAL", "Ruta Virtual"] = 1001
        df.loc[df["Nombre Solicitante"] == "ORG COURIER ARG", "Ruta Virtual"] = 700

        df.loc[(df["Ruta Virtual"].isna()) & (df["CP Destino"] != 3600), "Ruta Virtual"] = 504
        
        # agregar valores SOLO si "Ruta Virtual" está vacía
        df.loc[(df["Ruta Virtual"].isna()) & (df["Peso del objeto"] >= 200), "Ruta Virtual"] = 503
        df.loc[(df["Ruta Virtual"].isna()) & (df["Volumen"] >= 0.7), "Ruta Virtual"] = 503

        #limpiar columnas
        df["Distrito Destino"] = ""
        df["Provincia"] = ""

        # filtrar columnas
        df = df [
            (df["Motivo Descripción"] != "Retirado") & 
            (df["Motivo Descripción"] != "Entregado") &
            (df["Destino"] == iata)
        ].copy()
    else:
        pass

    if iata == "IRJ":
        # agregar valores SOLO si "Ruta Virtual" está vacía
        df.loc[(df["Ruta Virtual"].isna()) & (df["Peso del objeto"] >= 200), "Ruta Virtual"] = 503
        df.loc[(df["Ruta Virtual"].isna()) & (df["Volumen"] >= 0.7), "Ruta Virtual"] = 503
        #agregar valores
        df.loc[df["Nombre Solicitante"] == "ORG COURIER ARG", "Ruta Virtual"] = 700
        #limpiar columnas
        df["Distrito Destino"] = ""
        df["Provincia"] = ""

    if iata == "CRD":
        df.loc[df["Nombre Solicitante"] == "OCASA DISTRIBUCION POSTAL", "Ruta Virtual"] = 502
        df = df [
            (df["Motivo Descripción"] != "Retirado") & 
            (df["Motivo Descripción"] != "Entregado") &
            (df["Destino"] == iata)
        ].copy()

        # agregar valores
        df.loc[df["Nombre Solicitante"] == "ORG COURIER ARG", "Ruta Virtual"] = 700
        df.loc[df["Peso del objeto"] >= 200, "Ruta Virtual"] = 503
        df.loc[df["Volumen"] >= 0.7, "Ruta Virtual"] = 503

        #limpiar columnas
        df["Distrito Destino"] = ""
        df["Provincia"] = ""
    else:
        pass
    if iata == "LUQ":
        df = df [
            (df["Motivo Descripción"] != "Retirado") & 
            (df["Motivo Descripción"] != "Entregado") &
            (df["Destino"] == iata)
        ].copy()

        # agregar valores
        df.loc[df["Nombre Solicitante"] == "ORG COURIER ARG", "Ruta Virtual"] = 700
        df.loc[df["Peso del objeto"] >= 200, "Ruta Virtual"] = 503
        df.loc[df["Volumen"] >= 0.7, "Ruta Virtual"] = 503

        #limpiar columnas
        df["Distrito Destino"] = ""
        df["Provincia"] = ""
    else:
        pass
    return df

def canalizadorLocalidad(df):
    # --- MERGE con CANALIZADOR para que traiga la localidad ---
        canalizadorLocalidad = pd.read_excel("CANALIZADOR MADRE.xlsx")
        canalizador_reducidoLocalidad = canalizadorLocalidad[["CP Destino", "Distrito Destino"]]

        #elimina la columna distrito destino original para luego reemplazarla por el merge-
        df = df.drop(columns=["Distrito Destino"], errors="ignore")
        merge = pd.merge(df, canalizador_reducidoLocalidad, on="CP Destino", how="left")

        # Insertar "Distrito Destino" después de "Altura"
        columna_referencia = "Altura"
        if columna_referencia in merge.columns:
            indice_destino = merge.columns.get_loc(columna_referencia) + 1
            columna_valores = merge.pop("Distrito Destino")
            merge.insert(indice_destino, "Distrito Destino", columna_valores)
        else:
            print(f"Advertencia: no se encontró la columna '{columna_referencia}' para ubicar 'Distrito Destino'. Se dejó al final.")
        return merge

def canalizadorProvincia(df):
    # --- MERGE con CANALIZADOR para que traiga la provincia ---
        canalizadorProvincia = pd.read_excel("CANALIZADOR MADRE.xlsx")
        canalizador_reducidoProvincia = canalizadorProvincia[["CP Destino", "Provincia"]]

        #elimina la columna distrito destino original para luego reemplazarla por el merge-
        df = df.drop(columns=["Provincia"], errors="ignore")
        merge = pd.merge(df, canalizador_reducidoProvincia, on="CP Destino", how="left")

        # Insertar "Provincia" después de "Poblacion"
        columna_referencia = "Población"
        if columna_referencia in merge.columns:
            indice_destino = merge.columns.get_loc(columna_referencia) + 1
            columna_valores = merge.pop("Provincia")
            merge.insert(indice_destino, "Provincia", columna_valores)
        else:
            print(f"Advertencia: no se encontró la columna '{columna_referencia}' para ubicar 'Provincia'. Se dejó al final.")
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

    # Guardar y abrir archivo
    nombre_salida = f"archivoUnificado{iata}.xlsx"
    df.to_excel(nombre_salida, index=False)
    os.startfile(nombre_salida)

# --- Interfaz Gráfica ---
def crear_interfaz():
    ventana = tk.Tk()
    ventana.title("Procesador de Ruteo por IATA")
    ventana.geometry("400x200")

    etiqueta = ttk.Label(ventana, text="Seleccione un código IATA:", font=("Arial", 12))
    etiqueta.pack(pady=10)

    opciones_iata = ["FMA", "IRJ", "CRD", "LUQ", "TUC", "RES"]
    seleccion = tk.StringVar()
    combobox = ttk.Combobox(ventana, textvariable=seleccion, values=opciones_iata, state="readonly")
    combobox.pack(pady=5)

    def ejecutar():
        iata = seleccion.get()
        if not iata:
            messagebox.showwarning("Advertencia", "Debe seleccionar un código IATA.")
        else:
            procesar(iata)

    boton = ttk.Button(ventana, text="Procesar", command=ejecutar)
    boton.pack(pady=20)

    ventana.mainloop()

if __name__ == "__main__":
    crear_interfaz()
