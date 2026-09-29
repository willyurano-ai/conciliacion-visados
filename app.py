import streamlit as st
import pandas as pd

st.set_page_config(page_title="Conciliación de Visados y Sellos", layout="wide")

st.title("📊 Sistema de Conciliación: Planilla Diaria vs. Totales")
st.write("Sube ambas planillas para verificar los valores por mes y rango de días.")

# Sección de carga de archivos
st.sidebar.header("1. Carga de Archivos")
file_diaria = st.sidebar.file_uploader("Subir Planilla Diaria (Ingreso Visados)", type=["xlsx", "xls"])
file_totales = st.sidebar.file_uploader("Subir Planilla de Totales", type=["xlsx", "xls"])

if file_diaria is not None and file_totales is not None:
    try:
        # Leer archivos Excel
        df_diaria = pd.read_excel(file_diaria)
        df_totales = pd.read_excel(file_totales)

        st.success("¡Archivos cargados con éxito!")

        # Vista previa de datos opcional
        with st.expander("Ver vista previa de los datos brutos"):
            st.write("Planilla Diaria:", df_diaria.head(3))
            st.write("Planilla Totales:", df_totales.head(3))

        st.info("Configurá los filtros de fecha y ejecuta la conciliación.")

        # Aquí incorporaremos los filtros de fecha y la lógica exacta de cruce 
        # basada en las columnas indicadas en tu documento.

    except Exception as e:
        st.error(f"Ocurrió un error al procesar los archivos: {e}")
else:
    st.warning("Por favor, sube ambos archivos en la barra lateral para comenzar.")
