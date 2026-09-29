import streamlit as st
import pandas as pd
import numpy as np
from collections import Counter

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

        # Función de búsqueda inteligente con exclusión de palabras trampa (como Número, Cantidad, Operador)
        def encontrar_columna_inteligente(columns, exact_keywords, avoid_keywords, default_idx):
            for idx, col in enumerate(columns):
                c_up = str(col).upper()
                if any(k in c_up for k in exact_keywords) and not any(a in c_up for a in avoid_keywords):
                    return idx
            for idx, col in enumerate(columns):
                c_up = str(col).upper()
                if any(k in c_up for k in exact_keywords):
                    return idx
            return default_idx if default_idx < len(columns) else 0

        st.sidebar.divider()
        st.sidebar.header("2. Selección de Columnas Clave")
        
        idx_fecha_d = encontrar_columna_inteligente(columns_d, ['FECHA', 'DATE'], [], 1)
        # Búsqueda específica para evitar que tome "Número Visado" u otros campos erróneos
        idx_visado_d = encontrar_columna_inteligente(columns_d, ['IMPORTE VISADOS', 'VISADOS', 'VISADO'], ['NUMERO', 'NRO', 'CANTIDAD', 'OPERADOR'], 4)
        idx_sellos_d = encontrar_columna_inteligente(columns_d, ['SELLO', 'SELLOS'], [], 8)

        col_fecha_d = st.sidebar.selectbox("Columna Fecha (Diaria)", columns_d, index=idx_fecha_d)
        col_visado_f = st.sidebar.selectbox("Columna Importe Visado (Diaria)", columns_d, index=idx_visado_d)
        col_sellos_d = st.sidebar.selectbox("Columna Sellos (Diaria)", columns_d, index=idx_sellos_d)

        st.sidebar.divider()
        idx_concepto_t = encontrar_columna_inteligente(columns_t, ['CONCEPTO', 'CUENTA', 'DESCRIPCION'], [], 2)
        idx_monto_t = encontrar_columna_inteligente(columns_t, ['MONTO', 'IMPORTE', 'VALOR'], [], 3)
        idx_fecha_t = encontrar_columna_inteligente(columns_t, ['FECHA', 'HORA', 'DATE'], [], 4)

        col_concepto_t = st.sidebar.selectbox("Columna Concepto (Totales)", columns_t, index=idx_concepto_t)
        col_monto_t = st.sidebar.selectbox("Columna Monto (Totales)", columns_t, index=idx_monto_t)
        col_fecha_t = st.sidebar.selectbox("Columna Fecha/Hora (Totales)", columns_t, index=idx_fecha_t)

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

        # Detección automática de fecha
        valid_dates_d = df_d['Fecha_dt'].dropna()
        detected_date = valid_dates_d.min().date() if not valid_dates_d.empty else pd.Timestamp.today().date()

        st.sidebar.divider()
        st.sidebar.header("3. Período de Análisis")
        selected_date = st.sidebar.date_input("Fecha de Auditoría", value=detected_date)

        # Filtrado estricto
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
            st.metric(label="Planilla Diaria", value=f"${tot_visado_diaria:,.2f}")
            st.metric(label="Planilla Totales", value=f"${tot_visado_totales:,.2f}")
            diff_visado = round(tot_visado_diaria - tot_visado_totales, 2)
            if abs(diff_visado) < 0.01:
                st.success("✅ **ESTÁ OK (Sin diferencias en Visados)**")
            else:
                st.error(f"❌ **DIFERENCIA TOTAL:** ${diff_visado:,.2f}")

        with col2:
            st.markdown("### 🏷️ Recaudación Sellos")
            st.metric(label="Planilla Diaria", value=f"${tot_sellos_diaria:,.2f}")
            st.metric(label="Planilla Totales", value=f"${tot_sellos_totales:,.2f}")
            diff_sellos = round(tot_sellos_diaria - tot_sellos_totales, 2)
            if abs(diff_sellos) < 0.01:
                st.success("✅ **ESTÁ OK (Sin diferencias en Sellos)**")
            else:
                st.error(f"❌ **DIFERENCIA TOTAL:** ${diff_sellos:,.2f}")

        st.divider()
        st.subheader("🕵️‍♂️ Auditor Inteligente (Vista Unificada y Detección de Errores)")

        # Función auxiliar para diagnóstico limpio por conjuntos de importes
        def generar_diagnostico_limpio(lista_diaria, lista_totales, nombre_concepto):
            d_vals = [round(float(x), 2) for x in lista_diaria if pd.notna(x) and float(x) != 0.0]
            t_vals = [round(float(x), 2) for x in lista_totales if pd.notna(x) and float(x) != 0.0]
            
            c_d = Counter(d_vals)
            c_t = Counter(t_vals)
            
            solo_d = list((c_d - c_t).elements())
            solo_t = list((c_t - c_d).elements())
            
            if not solo_d and not solo_t:
                st.success(f"✨ **{nombre_concepto}: Sin errores. Todos los registros coinciden perfectamente entre ambas planillas.**")
            else:
                st.error(f"🚨 **Desvíos encontrados en {nombre_concepto}:**")
                if solo_d:
                    st.markdown(f"- **Importes presentes en Planilla Diaria que NO están en Totales:**")
                    for m in sorted(solo_d, reverse=True):
                        st.markdown(f"  - 📌 `${m:,.2f}`")
                if solo_t:
                    st.markdown(f"- **Importes presentes en Planilla Totales que NO están en Diaria:**")
                    for m in sorted(solo_t, reverse=True):
                        st.markdown(f"  - 📌 `${m:,.2f}`")

        # --- VISADOS: Tabla Unificada y Diagnóstico Limpio ---
        st.markdown("#### 🏛️ Detalle Comparativo: Tasa de Visado")
        df_d_v = df_d_filtered[[col_fecha_d, col_visado_f, 'Visado_Limpio']].sort_values(by='Visado_Limpio', ascending=False).reset_index(drop=True)
        df_t_v = df_t_filtered.loc[mask_tasa, [col_fecha_t, col_concepto_t, col_monto_t, 'Monto_Limpio']].sort_values(by='Monto_Limpio', ascending=False).reset_index(drop=True)
        
        df_merged_v = pd.concat([df_d_v.add_prefix('Diaria_'), df_t_v.add_prefix('Totales_')], axis=1)
        st.dataframe(df_merged_v, use_container_width=True)

        generar_diagnostico_limpio(df_d_v['Visado_Limpio'], df_t_v['Monto_Limpio'], "Tasa de Visado")

        st.divider()

        # --- SELLOS: Tabla Unificada y Diagnóstico Limpio ---
        st.markdown("#### 🏷️ Detalle Comparativo: Recaudación de Sellos")
        df_d_s = df_d_filtered[[col_fecha_d, col_sellos_d, 'Sellos_Limpio']].sort_values(by='Sellos_Limpio', ascending=False).reset_index(drop=True)
        df_t_s = df_t_filtered.loc[mask_sellos, [col_fecha_t, col_concepto_t, col_monto_t, 'Monto_Limpio']].sort_values(by='Monto_Limpio', ascending=False).reset_index(drop=True)
        
        df_merged_s = pd.concat([df_d_s.add_prefix('Diaria_'), df_t_s.add_prefix('Totales_')], axis=1)
        st.dataframe(df_merged_s, use_container_width=True)

        generar_diagnostico_limpio(df_d_s['Sellos_Limpio'], df_t_s['Monto_Limpio'], "Recaudación de Sellos")

    except Exception as e:
        st.error(f"Ocurrió un error al procesar: {e}")
else:
    st.info("Por favor, sube ambos archivos en la barra lateral para comenzar la conciliación.")
