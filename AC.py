import pandas as pd
import streamlit as st
import zipfile
import os
from datetime import datetime

# Función para transformar el 'Tipo de Documento' basado en las reglas
def transformar_tipo(tipo, rut):
    if tipo == "FÑ":
        return "33"
    elif tipo == "FO":
        return "34"
    elif tipo == "ZV":
        if rut in ["60503000-9", "76516999-2", "9297612-2"]:
            return "34"
        else:
            return "33"
    else:
        return tipo

# Función principal para procesar el archivo
def procesar_archivo(df):
    
    # Filtrar filas donde 'Referencia' contiene "-"
    df["Referencia"] = df['Referencia'].astype(str).str.replace('.', '', regex=True).str.split('-').str[0]
    df = df[~df["Referencia"].str.contains("-", na=False)]
    
    # Renombrar columnas según los requerimientos
    columnas_nuevas = {
        "Acreedor": "Rut emisor",
        "Clase de documento": "Tipo de Documento",
        "Referencia": "Folio",
        "Importe en moneda local": "Monto a pagar",
        "Vencimiento neto": "Fecha a pagar"
    }
    df = df.rename(columns=columnas_nuevas)

    # Aplicar transformación al 'Tipo de Documento'
    df["Tipo de Documento"] = df.apply(
        lambda row: transformar_tipo(row["Tipo de Documento"], row["Rut emisor"]), axis=1
    )

    # Limpiar y transformar el 'Monto a pagar'
    df["Monto a pagar"] = (
        df["Monto a pagar"]
        .astype(str)  # Convertir a texto para manipular
        .str.replace(".", "", regex=False)  # Eliminar puntos
        .astype(float)  # Convertir a número
        .abs()  # Hacer positivo
        .astype(int)  # Convertir a entero
    )

    # Formatear la fecha a 'dd-mm-aaaa'
    df["Fecha a pagar"] = pd.to_datetime(df["Fecha a pagar"], errors="coerce").dt.strftime("%d-%m-%Y")

    # Agrupar por 'Sociedad' y crear un archivo por cada grupo
    archivos_por_sociedad = {}
    for sociedad, grupo in df.groupby("Sociedad"):
        archivos_por_sociedad[sociedad] = grupo[["Rut emisor", "Tipo de Documento", "Folio", "Monto a pagar", "Fecha a pagar"]]

    return archivos_por_sociedad


# Configuración de la app Streamlit
st.title("Procesador de archivos de confirmación")
st.write("Sube un archivo Excel con los datos a procesar:")

# Subida de archivo
archivo_subido = st.file_uploader("Subir archivo", type=["xlsx", "xls"])

if archivo_subido is not None:
    try:
        # Leer archivo Excel
        df = pd.read_excel(archivo_subido)

        # Procesar archivo
        dfs_por_sociedad = procesar_archivo(df)

        # Crear un archivo ZIP para guardar los archivos procesados
        zip_nombre = "archivos_confirmacion.zip"
        with zipfile.ZipFile(zip_nombre, "w") as zipf:
            for sociedad, df_sociedad in dfs_por_sociedad.items():
                # Crear un nombre de archivo único para cada sociedad
                timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
                nombre_archivo = f"Data_{sociedad}_{timestamp}.xlsx"

                # Guardar el DataFrame en un archivo Excel
                df_sociedad.to_excel(nombre_archivo, index=False)

                # Añadir el archivo al ZIP
                zipf.write(nombre_archivo)

                # Eliminar el archivo temporal para no acumular basura
                os.remove(nombre_archivo)

        # Ofrecer el archivo ZIP para descargar
        with open(zip_nombre, "rb") as file:
            st.download_button(
                label="Descargar todos los archivos en un ZIP",
                data=file,
                file_name=zip_nombre,
                mime="application/zip"
            )

        # Eliminar el archivo ZIP temporal después de la descarga
        os.remove(zip_nombre)

    except Exception as e:
        st.error(f"Hubo un error al procesar el archivo: {e}")
