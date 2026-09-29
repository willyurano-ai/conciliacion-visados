import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Conciliación y Auditoría de Visados y Sellos", layout="wide")

st.title("📊 Sistema de Conciliación y Auditoría Inteligente")
st.write("Planilla Diaria vs. Planilla de Totales con Detección Automática de Discrepancias.")

# Barra lateral para carga de archivos
st.sidebar.header("1. Carga de Archivos")
file_diaria = st.sidebar.file_uploader("Subir Planilla Diaria (Ingreso Visados)", type=["xlsx", "xls"])
file_totales = st.sidebar.file_uploader("Subir Planilla de Totales", type=["xlsx", "xls"])

skip_rows_d = st.sidebar.number_input("Filas a omitir al inicio (Planilla Diaria)", min_value=0, max_value=5, value=2, help="Normalmente 2 para saltar títulos y ubicar las columnas correctamente.")

if file_diaria is not None and file_totales is not None:
    try:
        # Lectura de Planilla Diaria
        xls_d = pd.ExcelFile(file_diaria)
        sheet_d = st.sidebar.selectbox("Hoja Planilla Diaria", xls_d.sheet_names)
        df_diaria_raw = pd.read_excel(file_diaria, sheet_name=sheet_d, skiprows=skip_rows_d)

        # Lectura de Planilla Totales
        xls_t = pd.ExcelFile(file_totales)
        sheet_t = st.sidebar.selectbox("Hoja Planilla Totales", xls_t.sheet_names)
        df_totales_raw = pd.read_excel(file_totales, sheet_name=sheet_t)

        st.success("¡Archivos cargados con éxito!")

        columns_d = list(df_diaria_raw.columns)
        columns_t = list(df_totales_raw.columns)

        st.sidebar.divider()
        st.sidebar.header("2. Selección de Columnas Clave")
        
        default_fecha_d = columns_d[1] if len(columns_d) > 1 else columns_d[0]
        default_visado_d = columns_d[4] if len(columns_d) > 4 else columns_d[0]
        default_sellos_d = columns_d[8] if len(columns_d) > 8 else columns_d[0]

        col_fecha_d = st.sidebar.selectbox("Columna Fecha (Diaria - Col B)", columns_d, index=columns_d.index(default_fecha_d) if default_fecha_d in columns_d else 0)
        col_visado_f = st.sidebar.selectbox("Columna Importe Visado (Diaria - Col F)", columns_d, index=columns_d.index(default_visado_d) if default_visado_d in columns_d else 0)
        col_sellos_d = st.sidebar.selectbox("Columna Sellos (Diaria - Col J)", columns_d, index=columns_d.index(default_sellos_d) if default_sellos_d in columns_d else 0)

        st.sidebar.divider()
        default_concepto_t = columns_t[2] if len(columns_t) > 2 else columns_t[0]
        default_monto_t = columns_t[3] if len(columns_t) > 3 else columns_t[0]
        default_fecha_t = columns_t[4] if len(columns_t) > 4 else columns_t[0]

        col_concepto_t = st.sidebar.selectbox("Columna Concepto (Totales - Col C)", columns_t, index=columns_t.index(default_concepto_t) if default_concepto_t in columns_t else 0)
        col_monto_t = st.sidebar.selectbox("Columna Monto (Totales - Col D)", columns_t, index=columns_t.index(default_monto_t) if default_monto_t in columns_t else 0)
        col_fecha_t = st.sidebar.selectbox("Columna Fecha/Hora (Totales - Col E)", columns_t, index=columns_t.index(default_fecha_t) if default_fecha_t in columns_t else 0)

        # Procesamiento Planilla Diaria
        df_d = df_diaria_raw.copy()
        df_d['Fecha_dt'] = pd.to_datetime(df_d[col_fecha_d], errors='coerce', dayfirst=True)
        
        def limpiar_monto(val):
            if pd.isna(val):
                return 0.0
            if isinstance(val, (int, float)):
                return float(val)
            val_str = str(val).strip().upper()
            if val_str in ["OBLEA", "OBLEAS", "ARBA", "S/D", "N/A", ""]:
                return 0.0
            try:
                val_str = val_str.replace('$', '').strip()
                if '.' in val_str and ',' in val_str:
                    val_str = val_str.replace('.', '').replace(',', '.')
                elif ',' in val_str and '.' not in val_str:
                    val_str = val_str.replace(',', '.')
                return float(val_str)
            except:
                return 0.0

        df_d['Visado_Limpio'] = df_d[col_visado_f].apply(limpiar_monto)
        df_d['Sellos_Limpio'] = df_d[col_sellos_d].apply(limpiar_monto)

        # Procesamiento Planilla Totales
        df_t = df_totales_raw.copy()
        if str(df_t.iloc[0][col_concepto_t]).strip().upper() in ["CUENTA CONTABLE", "CONCEPTO", "NAN"]:
            df_t = df_t.iloc[1:].copy()

        df_t['Fecha_dt'] = pd.to_datetime(df_t[col_fecha_t], errors='coerce', dayfirst=True)
        df_t['Monto_Limpio'] = df_t[col_monto_t].apply(limpiar_monto)
        df_t['Concepto_Limpio'] = df_t[col_concepto_t].astype(str).str.strip().str.upper()

        # Detección automática de la fecha presente en los archivos
        valid_dates_d = df_d['Fecha_dt'].dropna()
        if not valid_dates_d.empty:
            detected_date = valid_dates_d.min().date()
        else:
            detected_date = pd.Timestamp.today().date()

        st.sidebar.divider()
        st.sidebar.header("3. Período de Análisis")
        # Selector de fecha optimizado con la fecha detectada por defecto
        selected_date = st.sidebar.date_input("Fecha de Auditoría", value=detected_date)

        # Filtrado estricto usando .dt.date para evitar conflictos de tipo con datetime64[s]
        df_d_filtered = df_d[df_d['Fecha_dt'].dt.date == selected_date].copy()
        df_t_filtered = df_t[df_t['Fecha_dt'].dt.date == selected_date].copy()

        st.subheader(f"📅 Conciliación y Auditoría del día: {selected_date}")

        # Totales Planilla Diaria
        tot_visado_diaria = df_d_filtered['Visado_Limpio'].sum()
        tot_sellos_diaria = df_d_filtered['Sellos_Limpio'].sum()

        # Totales Planilla Totales
        mask_tasa = df_t_filtered['Concepto_Limpio'].str.contains('INGRESO POR TASA DE VISADO', na=False)
        mask_sellos = df_t_filtered['Concepto_Limpio'].str.contains('RECAUDACION SELLOS', na=False)

        tot_visado_totales = df_t_filtered.loc[mask_tasa, 'Monto_Limpio'].sum()
        tot_sellos_totales = df_t_filtered.loc[mask_sellos, 'Monto_Limpio'].sum()

        # Métricas lado a lado
        col1, col2 = st.columns(2)

        with col1:
            st.markdown("### 🏛️ Tasa de Visado")
            st.metric(label="Planilla Diaria (Suma Renglones)", value=f"${tot_visado_diaria:,.2f}")
            st.metric(label="Planilla Totales (Suma Concepto)", value=f"${tot_visado_totales:,.2f}")
            diff_visado = round(tot_visado_diaria - tot_visado_totales, 2)
            if abs(diff_visado) < 0.01:
                st.success("✅ **ESTÁ OK (Sin diferencias en Visados)**")
            else:
                st.error(f"❌ **DIFERENCIA:** ${diff_visado:,.2f}")

        with col2:
            st.markdown("### 🏷️ Recaudación Sellos")
            st.metric(label="Planilla Diaria (Suma Renglones)", value=f"${tot_sellos_diaria:,.2f}")
            st.metric(label="Planilla Totales (Suma Concepto)", value=f"${tot_sellos_totales:,.2f}")
            diff_sellos = round(tot_sellos_diaria - tot_sellos_totales, 2)
            if abs(diff_sellos) < 0.01:
                st.success("✅ **ESTÁ OK (Sin diferencias en Sellos)**")
            else:
                st.error(f"❌ **DIFERENCIA:** ${diff_sellos:,.2f}")

        st.divider()
        st.subheader("🕵️‍♂️ Auditor Inteligente (Ordenado por Importe de Mayor a Menor)")

        # Auditoría de Visados ordenada estrictamente por importe descendente
        st.markdown("#### Detalle: Tasa de Visado")
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**Planilla Diaria (Visados)**")
            df_d_visado_sorted = df_d_filtered[[col_fecha_d, col_visado_f, 'Visado_Limpio']].sort_values(by='Visado_Limpio', ascending=False)
            st.dataframe(df_d_visado_sorted, use_container_width=True)
        with col_b:
            st.markdown("**Planilla Totales (Tasa de Visado)**")
            df_t_visado_sorted = df_t_filtered.loc[mask_tasa, [col_fecha_t, col_concepto_t, col_monto_t, 'Monto_Limpio']].sort_values(by='Monto_Limpio', ascending=False)
            st.dataframe(df_t_visado_sorted, use_container_width=True)

        st.divider()

        # Auditoría de Sellos ordenada estrictamente por importe descendente
        st.markdown("#### Detalle: Recaudación de Sellos")
        col_c, col_d = st.columns(2)
        with col_c:
            st.markdown("**Planilla Diaria (Sellos)**")
            df_d_sellos_sorted = df_d_filtered[[col_fecha_d, col_sellos_d, 'Sellos_Limpio']].sort_values(by='Sellos_Limpio', ascending=False)
            st.dataframe(df_d_sellos_sorted, use_container_width=True)
        with col_d:
            st.markdown("**Planilla Totales (Recaudación Sellos)**")
            df_t_sellos_sorted = df_t_filtered.loc[mask_sellos, [col_fecha_t, col_concepto_t, col_monto_t, 'Monto_Limpio']].sort_values(by='Monto_Limpio', ascending=False)
            st.dataframe(df_t_sellos_sorted, use_container_width=True)

    except Exception as e:
        st.error(f"Ocurrió un error al procesar: {e}")
else:
    st.info("Por favor, sube ambos archivos en la barra lateral para comenzar la conciliación.")
