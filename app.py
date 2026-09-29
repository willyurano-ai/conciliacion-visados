import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Conciliación y Auditoría de Visados y Sellos", layout="wide")

st.title("📊 Sistema de Conciliación y Auditoría Inteligente")
st.write("Planilla Diaria vs. Planilla de Totales con Detección Automática de Discrepancias.")

# Barra lateral para carga
st.sidebar.header("1. Carga de Archivos")
file_diaria = st.sidebar.file_uploader("Subir Planilla Diaria (Ingreso Visados)", type=["xlsx", "xls"])
file_totales = st.sidebar.file_uploader("Subir Planilla de Totales", type=["xlsx", "xls"])

skip_rows_d = st.sidebar.number_input("Filas a omitir al inicio (Planilla Diaria)", min_value=0, max_value=5, value=2, help="Normalmente 2 para saltar títulos y ubicar las columnas correctamente.")

if file_diaria is not None and file_totales is not None:
    try:
        xls_d = pd.ExcelFile(file_diaria)
        sheet_d = st.sidebar.selectbox("Hoja Planilla Diaria", xls_d.sheet_names)
        df_diaria_raw = pd.read_excel(file_diaria, sheet_name=sheet_d, skiprows=skip_rows_d)

        xls_t = pd.ExcelFile(file_totales)
        sheet_t = st.sidebar.selectbox("Hoja Planilla Totales", xls_t.sheet_names)
        df_totales_raw = pd.read_excel(file_totales, sheet_name=sheet_t)

        st.success("¡Archivos cargados con éxito!")

        columns_d = list(df_diaria_raw.columns)
        columns_t = list(df_totales_raw.columns)

        st.sidebar.divider()
        st.sidebar.header("2. Selección de Columnas Clave")
        
        default_fecha_d = columns_d[1] if len(columns_d) > 1 else columns_d[0]
        default_visado_d = columns_d[5] if len(columns_d) > 5 else columns_d[0]
        default_sellos_d = columns_d[9] if len(columns_d) > 9 else columns_d[0]

        col_fecha_d = st.sidebar.selectbox("Columna Fecha (Diaria - Col B)", columns_d, index=columns_d.index(default_fecha_d) if default_fecha_d in columns_d else 0)
        col_visado_f = st.sidebar.selectbox("Columna Importe Visado / Oblea (Diaria - Col F)", columns_d, index=columns_d.index(default_visado_d) if default_visado_d in columns_d else 0)
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
        # MODIFICACIÓN CRUCIAL: dayfirst=True para interpretar correctamente el formato español (DD/MM/YYYY)
        df_d['Fecha_dt'] = pd.to_datetime(df_d[col_fecha_d], errors='coerce', dayfirst=True)
        
        def limpiar_monto_visado(val):
            if pd.isna(val):
                return 0.0
            val_str = str(val).strip().upper()
            if val_str in ["OBLEA", "OBLEAS", "ARBA", "S/D", "N/A", ""]:
                return 0.0
            try:
                val_clean = str(val).replace('$', '').replace('.', '').replace(',', '.').strip()
                return float(val_clean)
            except:
                return 0.0

        def limpiar_monto_general(val):
            if pd.isna(val):
                return 0.0
            val_str = str(val).strip().upper()
            if val_str in ["OBLEA", "OBLEAS", "ARBA", "S/D", "N/A", ""]:
                return 0.0
            try:
                val_clean = str(val).replace('$', '').replace('.', '').replace(',', '.').strip()
                return float(val_clean)
            except:
                return 0.0

        df_d['Visado_Limpio'] = df_d[col_visado_f].apply(limpiar_monto_visado)
        df_d['Sellos_Limpio'] = df_d[col_sellos_d].apply(limpiar_monto_general)

        # Procesamiento Planilla Totales
        df_t = df_totales_raw.copy()
        if str(df_t.iloc[0][col_concepto_t]).strip().upper() in ["CUENTA CONTABLE", "CONCEPTO"]:
            df_t = df_t.iloc[1:].copy()

        # MODIFICACIÓN CRUCIAL: dayfirst=True también para la planilla de totales
        df_t['Fecha_dt'] = pd.to_datetime(df_t[col_fecha_t], errors='coerce', dayfirst=True)
        df_t['Monto_Limpio'] = df_t[col_monto_t].apply(limpiar_monto_general)
        df_t['Concepto_Limpio'] = df_t[col_concepto_t].astype(str).str.strip().str.upper()

        # Rango de fechas
        valid_dates_t = df_t['Fecha_dt'].dropna()
        valid_dates_d = df_d['Fecha_dt'].dropna()

        if not valid_dates_t.empty or not valid_dates_d.empty:
            min_date = min(valid_dates_t.min() if not valid_dates_t.empty else valid_dates_d.min(), valid_dates_d.min() if not valid_dates_d.empty else valid_dates_t.min()).date()
            max_date = max(valid_dates_t.max() if not valid_dates_t.empty else valid_dates_d.max(), valid_dates_d.max() if not valid_dates_d.empty else valid_dates_t.max()).date()

            st.sidebar.divider()
            st.sidebar.header("3. Período de Análisis")
            rango_fechas = st.sidebar.date_input("Seleccionar Rango de Fechas", [min_date, max_date], min_value=min_date, max_value=max_date)

            if len(rango_fechas) == 2:
                start_date, end_date = rango_fechas

                # Filtrado estricto por rango de fechas
                df_d_filtered = df_d[(df_d['Fecha_dt'].dt.date >= start_date) & (df_d['Fecha_dt'].dt.date <= end_date)]
                df_t_filtered = df_t[(df_t['Fecha_dt'].dt.date >= start_date) & (df_t['Fecha_dt'].dt.date <= end_date)]

                st.subheader(f"📅 Conciliación y Auditoría del período: {start_date} al {end_date}")

                # Totales Planilla Diaria (Suma estricta renglón por renglón)
                tot_visado_diaria = df_d_filtered['Visado_Limpio'].sum()
                tot_sellos_diaria = df_d_filtered['Sellos_Limpio'].sum()

                # Totales Planilla Totales filtrando por concepto específico
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
                st.subheader("🕵️‍♂️ Auditor Inteligente: Detalle para el Análisis Humano")

                # Auditoría de Visados
                if abs(diff_visado) >= 0.01:
                    st.warning(f"⚠️ **Desvío detectado en Tasa de Visado por ${diff_visado:,.2f}**")
                    col_a, col_b = st.columns(2)
                    with col_a:
                        st.markdown("**Detalle Planilla Diaria (Visados)**")
                        st.dataframe(df_d_filtered[[col_fecha_d, col_visado_f, 'Visado_Limpio']])
                    with col_b:
                        st.markdown("**Detalle Planilla Totales (Tasa de Visado)**")
                        st.dataframe(df_t_filtered.loc[mask_tasa, [col_fecha_t, col_concepto_t, col_monto_t]])
                else:
                    st.success("🔍 Los visados coinciden exactamente en el total del período.")

                # Auditoría de Sellos
                if abs(diff_sellos) >= 0.01:
                    st.warning(f"⚠️ **Desvío detectado en Recaudación de Sellos por ${diff_sellos:,.2f}**")
                    col_c, col_d = st.columns(2)
                    with col_c:
                        st.markdown("**Detalle Planilla Diaria (Sellos)**")
                        st.dataframe(df_d_filtered[[col_fecha_d, col_sellos_d, 'Sellos_Limpio']])
                    with col_d:
                        st.markdown("**Detalle Planilla Totales (Recaudación Sellos)**")
                        st.dataframe(df_t_filtered.loc[mask_sellos, [col_fecha_t, col_concepto_t, col_monto_t]])
                else:
                    st.success("🔍 Los sellos coinciden exactamente en el total del período.")

            else:
                st.warning("Seleccioná un rango de fechas válido.")
        else:
            st.warning("No se encontraron fechas válidas en los archivos.")

    except Exception as e:
        st.error(f"Ocurrió un error al procesar: {e}")
else:
    st.info("Por favor, sube ambos archivos en la barra lateral.")
