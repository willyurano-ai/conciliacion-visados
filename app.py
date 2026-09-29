import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Conciliación de Visados y Sellos", layout="wide")

st.title("📊 Sistema de Conciliación: Planilla Diaria vs. Totales")
st.write("Sube ambas planillas para verificar los valores por mes y rango de días.")

# Sección de carga de archivos en la barra lateral
st.sidebar.header("1. Carga de Archivos")
file_diaria = st.sidebar.file_uploader("Subir Planilla Diaria (Ingreso Visados)", type=["xlsx", "xls"])
file_totales = st.sidebar.file_uploader("Subir Planilla de Totales", type=["xlsx", "xls"])

if file_diaria is not None and file_totales is not None:
    try:
        # Lectura de los archivos excel
        xls_d = pd.ExcelFile(file_diaria)
        sheet_d = st.sidebar.selectbox("Hoja Planilla Diaria", xls_d.sheet_names)
        df_diaria_raw = pd.read_excel(file_diaria, sheet_name=sheet_d)

        xls_t = pd.ExcelFile(file_totales)
        sheet_t = st.sidebar.selectbox("Hoja Planilla Totales", xls_t.sheet_names)
        df_totales_raw = pd.read_excel(file_totales, sheet_name=sheet_t)

        st.success("¡Archivos cargados con éxito!")

        with st.expander("🔍 Ver vista previa de datos originales"):
            col1, col2 = st.columns(2)
            with col1:
                st.write("**Planilla Diaria (Primeras filas):**")
                st.dataframe(df_diaria_raw.head(5))
            with col2:
                st.write("**Planilla de Totales (Primeras filas):**")
                st.dataframe(df_totales_raw.head(5))

        st.sidebar.divider()
        st.sidebar.header("2. Configuración de Columnas y Filtros")

        columns_d = list(df_diaria_raw.columns)
        columns_t = list(df_totales_raw.columns)

        # Selección automática o manual de columnas clave según tus indicaciones
        col_fecha_d = st.sidebar.selectbox("Columna Fecha (Diaria)", columns_d, index=min(1, len(columns_d)-1))
        col_visado_f = st.sidebar.selectbox("Columna Importe Visado / Oblea (Diaria)", columns_d, index=min(5, len(columns_d)-1))
        col_sellos_d = st.sidebar.selectbox("Columna Sellos (Diaria)", columns_d, index=min(9, len(columns_d)-1))
        col_tot_visados_l = st.sidebar.selectbox("Columna Total Visados (Diaria)", columns_d, index=min(11, len(columns_d)-1))
        col_tot_sellos_p = st.sidebar.selectbox("Columna Total Sellos (Diaria)", columns_d, index=min(15, len(columns_d)-1))

        st.sidebar.divider()
        col_concepto_t = st.sidebar.selectbox("Columna Concepto (Totales)", columns_t, index=min(2, len(columns_t)-1))
        col_monto_t = st.sidebar.selectbox("Columna Monto (Totales)", columns_t, index=min(3, len(columns_t)-1))
        col_fecha_t = st.sidebar.selectbox("Columna Fecha/Hora (Totales)", columns_t, index=min(4, len(columns_t)-1))

        # Procesamiento de Planilla Diaria
        df_d = df_diaria_raw.copy()
        df_d['Fecha_dt'] = pd.to_datetime(df_d[col_fecha_d], errors='coerce')
        
        # Función para limpiar importes y detectar la palabra OBLEA (importe cero)
        def limpiar_visado(val):
            if pd.isna(val):
                return 0.0
            val_str = str(val).strip().upper()
            if "OBLEA" in val_str:
                return 0.0
            try:
                return float(str(val).replace('$', '').replace('.', '').replace(',', '.'))
            except:
                return 0.0

        df_d['Visado_Limpio'] = df_d[col_visado_f].apply(limpiar_visado)
        df_d['Sellos_Limpio'] = pd.to_numeric(df_d[col_sellos_d].astype(str).str.replace('$', '').str.replace('.', '').str.replace(',', '.'), errors='coerce').fillna(0)
        df_d['Tot_Visados_Limpio'] = pd.to_numeric(df_d[col_tot_visados_l].astype(str).str.replace('$', '').str.replace('.', '').str.replace(',', '.'), errors='coerce').fillna(0)
        df_d['Tot_Sellos_Limpio'] = pd.to_numeric(df_d[col_tot_sellos_p].astype(str).str.replace('$', '').str.replace('.', '').str.replace(',', '.'), errors='coerce').fillna(0)

        # Procesamiento de Planilla Totales
        df_t = df_totales_raw.copy()
        df_t['Fecha_dt'] = pd.to_datetime(df_t[col_fecha_t], errors='coerce')
        df_t['Monto_Limpio'] = pd.to_numeric(df_t[col_monto_t].astype(str).str.replace('$', '').str.replace('.', '').str.replace(',', '.'), errors='coerce').fillna(0)
        df_t['Concepto_Limpio'] = df_t[col_concepto_t].astype(str).str.strip().str.upper()

        # Selector de Mes y Rango de Días
        valid_dates = df_d['Fecha_dt'].dropna()
        if not valid_dates.empty:
            min_date = valid_dates.min().date()
            max_date = valid_dates.max().date()
            
            st.sidebar.header("3. Filtro de Período")
            rango_fechas = st.sidebar.date_input("Seleccionar Rango de Fechas", [min_date, max_date], min_value=min_date, max_value=max_date)
            
            if len(rango_fechas) == 2:
                start_date, end_date = rango_fechas
                
                # Filtrar dataframes por el período seleccionado
                df_d_filtered = df_d[(df_d['Fecha_dt'].dt.date >= start_date) & (df_d['Fecha_dt'].dt.date <= end_date)]
                df_t_filtered = df_t[(df_t['Fecha_dt'].dt.date >= start_date) & (df_t['Fecha_dt'].dt.date <= end_date)]

                st.subheader(f"📅 Resultados de Conciliación: {start_date} al {end_date}")

                # Totales Planilla Diaria
                tot_visado_diaria = df_d_filtered['Visado_Limpio'].sum()
                tot_sellos_diaria = df_d_filtered['Sellos_Limpio'].sum()

                # Totales Planilla Totales por concepto
                mask_tasa = df_t_filtered['Concepto_Limpio'].str.contains('TASA DE VISADO|VISADO', na=False)
                mask_sellos = df_t_filtered['Concepto_Limpio'].str.contains('SELLOS|RECAUDACION', na=False)

                tot_visado_totales = df_t_filtered.loc[mask_tasa, 'Monto_Limpio'].sum()
                tot_sellos_totales = df_t_filtered.loc[mask_sellos, 'Monto_Limpio'].sum()

                # Mostrar métricas y conciliación en columnas
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("### 🏛️ Tasa de Visado")
                    st.metric(label="Planilla Diaria", value=f"${tot_visado_diaria:,.2f}")
                    st.metric(label="Planilla Totales", value=f"${tot_visado_totales:,.2f}")
                    diff_visado = round(tot_visado_diaria - tot_visado_totales, 2)
                    if abs(diff_visado) < 0.01:
                        st.success("✅ **ESTÁ OK (Sin diferencias en Visados)**")
                    else:
                        st.error(f"❌ **DIFERENCIA DETECTADA:** ${diff_visado:,.2f}")

                with col2:
                    st.markdown("### 🏷️ Recaudación Sellos")
                    st.metric(label="Planilla Diaria", value=f"${tot_sellos_diaria:,.2f}")
                    st.metric(label="Planilla Totales", value=f"${tot_sellos_totales:,.2f}")
                    diff_sellos = round(tot_sellos_diaria - tot_sellos_totales, 2)
                    if abs(diff_sellos) < 0.01:
                        st.success("✅ **ESTÁ OK (Sin diferencias en Sellos)**")
                    else:
                        st.error(f"❌ **DIFERENCIA DETECTADA:** ${diff_sellos:,.2f}")

                # Análisis diario detallado
                with st.expander("📊 Ver detalle diario y desglose"):
                    st.write("**Detalle Diaria Agrupada por Día:**")
                    diaria_diaria = df_d_filtered.groupby(df_d_filtered['Fecha_dt'].dt.date)[['Visado_Limpio', 'Sellos_Limpio']].sum().reset_index()
                    st.dataframe(diaria_diaria)
                    
                    st.write("**Registros en Planilla Totales:**")
                    st.dataframe(df_t_filtered[['Fecha_dt', col_concepto_t, col_monto_t]])

            else:
                st.info("Por favor, selecciona un rango de fechas completo en la barra lateral.")
        else:
            st.warning("No se pudieron detectar fechas válidas en la Planilla Diaria.")

    except Exception as e:
        st.error(f"Ocurrió un error al procesar los archivos: {e}")
else:
    st.info("Por favor, sube ambos archivos en la barra lateral para comenzar.")
