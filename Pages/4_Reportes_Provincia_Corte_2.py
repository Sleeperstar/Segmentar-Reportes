# pages/4_Reportes_Provincia_Corte_2.py
import streamlit as st
import pandas as pd
import io
import zipfile
import re
from datetime import datetime

# --- Funciones de ayuda (reutilizadas y adaptadas) ---
def normalizar_nombre(nombre):
    """Convierte un nombre a un formato estándar: mayúsculas, sin puntos/comas y con espacios simples."""
    if not isinstance(nombre, str): return ""
    nombre_limpio = nombre.upper().replace('.', '').replace(',', '').replace('-', '')
    return re.sub(r'\s+', ' ', nombre_limpio).strip()

def get_agencia_base(nombre_completo, lista_departamentos):
    """
    Separa el nombre base de la agencia del departamento de forma robusta.
    Ej: 'MI AGENCIA PIURA' -> 'MI AGENCIA'
    """
    if not isinstance(nombre_completo, str):
        return ""
    
    # La lista de departamentos ya viene ordenada del más largo al más corto.
    for depto in lista_departamentos:
        # Creamos un patrón para buscar el departamento al final del string,
        # precedido de al menos un espacio. Es insensible a mayúsculas/minúsculas.
        pattern = r'\s+' + re.escape(depto) + '$'
        
        # Intentamos sustituir el patrón encontrado por una cadena vacía.
        cleaned_name, num_subs = re.subn(pattern, '', nombre_completo, flags=re.IGNORECASE)
        
        # Si se hizo una sustitución, encontramos el departamento.
        if num_subs > 0:
            return cleaned_name.strip()
            
    # Si no se encontró ningún departamento como sufijo, devolvemos el nombre original.
    return nombre_completo.strip()

def procesar_provincia_corte_2(archivo_excel_cargado):
    log_output = []
    log_output.append("--- INICIO DEL PROCESO: PROVINCIA CORTE 2 ---")
    
    # Mapa de alias para agencias con múltiples nombres de asesor
    mapeo_asesor_alias = {
        'EXPORTEL SAC': ['EXPORTEL SAC', 'EXPORTEL PROVINCIA']
    }
    log_output.append(f"Usando mapa de alias para: {', '.join(mapeo_asesor_alias.keys())}")

    # --- 1. Validación de Cabeceras ---
    try:
        df_headers_reporte = pd.read_excel(archivo_excel_cargado, sheet_name='Reporte CORTE 2', header=None, nrows=2)
        fila2_headers = [str(h).strip().upper() for h in df_headers_reporte.iloc[1].values]
        if 'AGENCIA' not in fila2_headers or 'RUC' not in fila2_headers:
            log_output.append("ALERTA: Cabeceras 'AGENCIA' o 'RUC' no encontradas en 'Reporte CORTE 2'.")
            return None, log_output

        df_headers_base = pd.read_excel(archivo_excel_cargado, sheet_name='BASE', header=None, nrows=1)
        base_headers = [str(h).strip().upper() for h in df_headers_base.iloc[0].values]
        if 'ASESOR' not in base_headers or 'DEPARTAMENTO' not in base_headers:
            log_output.append("ALERTA: Cabeceras 'ASESOR' o 'DEPARTAMENTO' no encontradas en la hoja 'BASE'.")
            return None, log_output
        log_output.append("Validación de cabeceras exitosa.")

    except Exception as e:
        log_output.append(f"ERROR al validar cabeceras: {e}")
        return None, log_output

    # --- 2. Lectura y Preparación de Datos ---
    try:
        log_output.append("Leyendo datos completos...")
        df_reporte_total = pd.read_excel(archivo_excel_cargado, sheet_name='Reporte CORTE 2', header=[0, 1])
        df_base_total = pd.read_excel(archivo_excel_cargado, sheet_name='BASE')
        df_base_total.columns = df_base_total.columns.str.strip().str.upper()

        lista_departamentos = df_base_total['DEPARTAMENTO'].dropna().unique().tolist()
        lista_departamentos.sort(key=len, reverse=True)
        log_output.append(f"Detectados {len(lista_departamentos)} departamentos para limpieza de nombres.")

        col_agencia_reporte = next((col for col in df_reporte_total.columns if 'AGENCIA' in col[1]), None)
        if not col_agencia_reporte:
            log_output.append("ERROR: No se encontró la columna 'AGENCIA' en 'Reporte CORTE 2'.")
            return None, log_output
        
        df_reporte_total['AGENCIA_BASE'] = df_reporte_total[col_agencia_reporte].apply(lambda x: get_agencia_base(x, lista_departamentos))
        df_reporte_total['AGENCIA_BASE_NORMALIZADA'] = df_reporte_total['AGENCIA_BASE'].apply(normalizar_nombre)
        df_base_total['ASESOR_NORMALIZADO'] = df_base_total['ASESOR'].apply(normalizar_nombre)

    except Exception as e:
        log_output.append(f"ERROR al leer o preparar datos: {e}")
        return None, log_output

    # --- 3. Proceso de Segmentación ---
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        agencias_a_procesar = df_reporte_total['AGENCIA_BASE_NORMALIZADA'].dropna().unique().tolist()
        log_output.append(f"Se encontraron {len(agencias_a_procesar)} agencias únicas para procesar.")

        for agencia_norm in agencias_a_procesar:
            reporte_agencia = df_reporte_total[df_reporte_total['AGENCIA_BASE_NORMALIZADA'] == agencia_norm].copy()
            
            # --- Lógica de cruce con mapa de alias ---
            if agencia_norm in mapeo_asesor_alias:
                nombres_a_buscar = mapeo_asesor_alias[agencia_norm]
                base_agencia = df_base_total[df_base_total['ASESOR_NORMALIZADO'].isin(nombres_a_buscar)].copy()
            else:
                base_agencia = df_base_total[df_base_total['ASESOR_NORMALIZADO'] == agencia_norm].copy()

            if reporte_agencia.empty: continue
            
            # --- INICIO: Bloque de validación ---
            try:
                # Encontrar la columna 'ALTAS' en el MultiIndex del reporte
                col_altas = next((col for col in reporte_agencia.columns if 'ALTAS' in col[1]), None)
                if col_altas:
                    altas_reporte = pd.to_numeric(reporte_agencia[col_altas], errors='coerce').fillna(0).sum() # type: ignore
                    registros_base = len(base_agencia)
                    if int(altas_reporte) == registros_base: # type: ignore
                        log_output.append(f"ÉXITO    | {agencia_norm:<40} | ALTAS: {int(altas_reporte):<5} | Registros BASE: {registros_base:<5} | OK") # type: ignore
                    else:
                        log_output.append(f"DESCUADRE | {agencia_norm:<40} | ALTAS: {int(altas_reporte):<5} | Registros BASE: {registros_base:<5} | REVISAR") # type: ignore
                else:
                    log_output.append(f"INFO     | {agencia_norm:<40} | No se pudo encontrar la columna ALTAS para validar.")
            except Exception as e:
                log_output.append(f"Error validando la agencia '{agencia_norm}': {e}")
            # --- FIN: Bloque de validación ---
            
            # --- INICIO: Corrección de formato de cabeceras y columnas ---
            
            # 1. Obtener el nombre original ANTES de eliminar la columna auxiliar.
            #    Se accede a la columna por su nombre de tupla en el MultiIndex.
            nombre_original_agencia = reporte_agencia[('AGENCIA_BASE', '')].iloc[0] # type: ignore

            # 2. Eliminar las columnas auxiliares.
            reporte_agencia = reporte_agencia.drop(columns=[('AGENCIA_BASE', ''), ('AGENCIA_BASE_NORMALIZADA', '')], errors='ignore')

            # 3. Aplanar las cabeceras de dos niveles en una sola, de forma limpia.
            new_cols = []
            for col in reporte_agencia.columns:
                level1 = str(col[0]).strip()
                level2 = str(col[1]).strip().replace('\n', ' ')
                # Si la cabecera superior es 'Unnamed' o un duplicado, usar solo la inferior.
                if 'unnamed' in level1.lower() or level1 == level2:
                    new_cols.append(level2)
                else:
                    new_cols.append(f"{level1} - {level2}")
            reporte_agencia.columns = new_cols
            
            reporte_agencia_final = reporte_agencia # Renombrar para claridad

            # --- FIN: Corrección de formato ---

            output_buffer = io.BytesIO()
            with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer: # type: ignore
                reporte_agencia_final.to_excel(writer, sheet_name='Reporte CORTE 2', index=False)
                base_agencia.drop(columns=['ASESOR_NORMALIZADO'], errors='ignore').to_excel(writer, sheet_name='BASE', index=False)
                
                workbook, worksheet = writer.book, writer.sheets['Reporte CORTE 2']
                percent_format = workbook.add_format({'num_format': '0.00%'})
                header_penalidad = workbook.add_format({'bold': True, 'font_color': 'white', 'fg_color': '#0070C0', 'border': 1})
                header_clawback = workbook.add_format({'bold': True, 'font_color': 'white', 'fg_color': '#002060', 'border': 1})
                default_header = workbook.add_format({'bold': True, 'fg_color': '#FFC000', 'border': 1})

                header = reporte_agencia_final.columns.tolist()
                for i, h_text in enumerate(header):
                    if h_text.startswith('PENALIDAD 1 -'): worksheet.write(0, i, h_text, header_penalidad)
                    elif h_text.startswith('CLAWBACK 1 -'): worksheet.write(0, i, h_text, header_clawback)
                    else: worksheet.write(0, i, h_text, default_header)

                for col_name in ['Cumplimiento Altas %', 'CLAWBACK 1 - Cumplimiento Corte 2 %']:
                    try:
                        worksheet.set_column(header.index(col_name), header.index(col_name), 18, percent_format)
                    except ValueError: pass
            
            nombre_archivo_limpio = "".join(c for c in nombre_original_agencia if c.isalnum() or c in (' ', '_')).rstrip()
            zf.writestr(f"Reporte Provincia Corte 2 {nombre_archivo_limpio}.xlsx", output_buffer.getvalue())

    log_output.append("--- FIN DEL PROCESO ---")
    zip_buffer.seek(0)
    return zip_buffer, log_output

# --- Interfaz de Usuario ---
st.title("Segmentador de Reportes - Provincia Corte 2")
st.markdown("Sube el archivo consolidado de **Provincia CORTE 2** para generar los reportes individuales.")
st.warning("El archivo debe contener las hojas 'Reporte CORTE 2' y 'BASE'.")

uploaded_file = st.file_uploader("Sube tu archivo Excel de Provincia CORTE 2", type=["xlsx"], key="provincia_corte_2_uploader")

if uploaded_file:
    st.success(f"Archivo '{uploaded_file.name}' cargado.")
    if st.button("Procesar y Generar Reportes", type="primary"):
        with st.spinner("Procesando archivo de Provincia Corte 2..."):
            zip_file, log_data = procesar_provincia_corte_2(uploaded_file)
        
        if zip_file:
            st.success("¡Proceso completado!")
            st.subheader("Log de Validación")
            st.text_area("Resultado:", "\n".join(log_data), height=300)
            st.subheader("Descargar Resultados")
            st.download_button(
                label="Descargar todos los reportes (.zip)",
                data=zip_file,
                file_name=f"Reportes_Provincia_Corte_2_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip")
        else:
            st.error("Ocurrió un error al procesar el archivo.")
            st.text_area("Log de Errores:", "\n".join(log_data), height=300) 