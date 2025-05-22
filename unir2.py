import pandas as pd
import glob
import os

def orgCourier(df, iata, lista1):
    if iata in lista1:
            df.loc[df["Nombre Solicitante"] == "ORG COURIER ARG", "Ruta Virtual"] = 700
    else:
        print("error")
    return df

def borrarMHTML():
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

def manipularDatos(df, iata):
    # filtrar columnas
    df = df [
        (df["Motivo Descripción"] != "Retirado") & 
        (df["Motivo Descripción"] != "Entregado") &
        (df["Destino"] == iata)
    ].copy()

    # agregar valores
    df.loc[df["Peso del objeto"] >= 200, "Ruta Virtual"] = 503
    df.loc[df["Volumen"] >= 0.7, "Ruta Virtual"] = 503

    #limpiar columnas
    df["Distrito Destino"] = ""
    df["Provincia"] = ""

    #duplicados = df.duplicated(subset= "Nro. identificación pieza según cliente", keep = False)
    return df

def canalizadorLocalidad(df, iata):
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

def canalizadorProvincia(df, iata):
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

def main():
    while True:
        iata = input("Ingresa el codigo IATA: ").upper()
        if len(iata) == 3:
            break
        else:
            print("Error. Intente nuevamente.")

    lista1 = ["CRD", "LUQ"]
    if iata in lista1:
        encontrar = glob.glob("*xlsx")
        lista = []

        for archivo in encontrar:
            if archivo == "CANALIZADOR MADRE.xlsx":
                continue  # saltea el canalizador
            leer = pd.read_excel(archivo)
            lista.append(leer)

        df = pd.concat(lista, ignore_index=True)
        df = manipularDatos(df, iata)
        df = orgCourier(df, iata, lista1)
        df = canalizadorLocalidad(df, iata)
        df = canalizadorProvincia(df,iata)

        # Guardar y abrir
        nombre_salida = f"archivoUnificado{iata}.xlsx"
        df.to_excel(nombre_salida, index=False)
        os.startfile(nombre_salida)

if __name__ == "__main__": 
    main()
