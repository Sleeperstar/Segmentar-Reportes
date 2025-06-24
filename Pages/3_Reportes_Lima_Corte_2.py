# pages/3_Reportes_Lima_Corte_2.py
import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime

def procesar_reporte_corte_2(archivo_excel_cargado):
    """
    Procesa un archivo Excel con la estructura de "Corte 2", que contiene
    cabeceras de múltiples niveles, y lo segmenta por agencia.
    """
    log_output = []
    log_output.append("--- INICIO DEL PROCESO DE SEGMENTACIÓN (CORTE 2) ---")

    # --- 1. Validación de Cabeceras ---
    try:
        # Validación para 'Reporte CORTE 2' con cabeceras en dos filas
        df_headers_reporte = pd.read_excel(archivo_excel_cargado, sheet_name='Reporte CORTE 2', header=None, nrows=2)
        fila1_headers = [str(h).strip().upper() for h in df_headers_reporte.iloc[0].values]
        fila2_headers = [str(h).strip().upper() for h in df_headers_reporte.iloc[1].values]
        
        cabeceras_fila1_esperadas = ['PENALIDAD 1', 'CLAWBACK 1']
        cabeceras_fila2_esperadas = ['RUC', 'AGENCIA', 'ALTAS', 'TOTAL A PAGAR CORTE 2']

        if not all(h in fila1_headers for h in cabeceras_fila1_esperadas) or not all(h in fila2_headers for h in cabeceras_fila2_esperadas):
            log_output.append("ALERTA DE ARCHIVO: No se encontraron las cabeceras esperadas en las dos primeras filas de la hoja 'Reporte CORTE 2'.")
            log_output.append("Asegúrese de que 'PENALIDAD 1', 'CLAWBACK 1' (fila 1) y 'RUC', 'AGENCIA', etc. (fila 2) estén presentes.")
            return None, log_output

        # Validación para 'BASE' (cabecera simple)
        df_headers_base = pd.read_excel(archivo_excel_cargado, sheet_name='BASE', header=None, nrows=1)
        base_headers = [str(h).strip().upper() for h in df_headers_base.iloc[0].values]
        if 'ASESOR' not in base_headers or 'COD_PEDIDO' not in base_headers:
            log_output.append("ALERTA DE ARCHIVO: Las cabeceras 'ASESOR' y 'COD_PEDIDO' no se encontraron en la hoja 'BASE'.")
            return None, log_output
        
        log_output.append("Validación de cabeceras exitosa.")

    except Exception as e:
        log_output.append(f"ERROR al validar cabeceras: {e}. Asegúrese de que las hojas 'Reporte CORTE 2' y 'BASE' existan.")
        return None, log_output

    # --- 2. Lectura de Datos Completos ---
    try:
        log_output.append("Leyendo datos completos del archivo...")
        # Leer el reporte con las dos primeras filas como cabecera
        df_reporte_total = pd.read_excel(archivo_excel_cargado, sheet_name='Reporte CORTE 2', header=[0, 1])
        df_base_total = pd.read_excel(archivo_excel_cargado, sheet_name='BASE')

        # Estandarizar cabeceras de la hoja BASE
        df_base_total.columns = df_base_total.columns.str.strip().str.upper()
        log_output.append("Datos cargados y cabeceras de la BASE estandarizadas.")

    except Exception as e:
        log_output.append(f"ERROR: No se pudo leer el archivo Excel. Error: {e}")
        return None, log_output

    # --- 3. Proceso de Segmentación ---
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        # La columna 'AGENCIA' está en el segundo nivel de la cabecera. 
        # Pandas crea tuplas para MultiIndex. Necesitamos encontrar la tupla correcta.
        columna_agencia = next((col for col in df_reporte_total.columns if 'AGENCIA' in col), None)
        if not columna_agencia:
             log_output.append(f"ERROR: No se pudo encontrar la columna 'AGENCIA' en la hoja 'Reporte CORTE 2'.")
             return None, log_output

        agencias_a_procesar = df_reporte_total[columna_agencia].dropna().unique().tolist()
        log_output.append(f"Se encontraron {len(agencias_a_procesar)} agencias únicas para procesar.")

        columna_altas = next((col for col in df_reporte_total.columns if 'ALTAS' in col), None)

        for agencia in agencias_a_procesar:
            reporte_agencia = df_reporte_total[df_reporte_total[columna_agencia] == agencia].copy()
            if reporte_agencia.empty:
                continue

            # La lógica de filtrado en la BASE sigue siendo por 'ASESOR'
            base_agencia = df_base_total[df_base_total['ASESOR'] == agencia]

            # Validación de consistencia
            try:
                if columna_altas:
                    altas_reporte = int(reporte_agencia.iloc[0][columna_altas])
                    registros_base = len(base_agencia)
                    if altas_reporte == registros_base:
                        log_output.append(f"ÉXITO    | {agencia:<40} | ALTAS: {altas_reporte:<5} | Registros BASE: {registros_base:<5} | OK")
                    else:
                        log_output.append(f"DESCUADRE | {agencia:<40} | ALTAS: {altas_reporte:<5} | Registros BASE: {registros_base:<5} | REVISAR")
                else:
                    log_output.append(f"INFO     | {agencia:<40} | No se pudo validar conteo de ALTAS.")
            except Exception as e:
                log_output.append(f"Error validando la agencia '{agencia}': {e}")

            # Aplanar el MultiIndex de las columnas para resolver el NotImplementedError.
            # Esto convierte la cabecera de dos filas en una sola, más limpia.
            new_cols = []
            for col in reporte_agencia.columns:
                level1 = str(col[0]).strip()
                level2 = str(col[1]).strip().replace('\n', ' ')
                # Si la cabecera superior es 'Unnamed' o es igual a la inferior, usar solo la inferior.
                if 'unnamed' in level1.lower() or level1 == level2:
                    new_cols.append(level2)
                else:
                    new_cols.append(f"{level1} - {level2}")
            reporte_agencia.columns = new_cols

            # Crear el archivo Excel para la agencia
            output_buffer = io.BytesIO()
            with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer: # type: ignore
                reporte_agencia.to_excel(writer, sheet_name='Reporte CORTE 2', index=False)
                base_agencia.to_excel(writer, sheet_name='BASE', index=False)
                
                # --- INICIO: Aplicar formato estético al reporte ---
                workbook = writer.book
                worksheet_reporte = writer.sheets['Reporte CORTE 2']

                # Definir formatos de celda y cabecera con los nuevos colores
                percent_format = workbook.add_format({'num_format': '0.00%'})
                header_penalidad_format = workbook.add_format({'bold': True, 'font_color': 'white', 'fg_color': '#0070C0', 'border': 1})
                header_clawback_format = workbook.add_format({'bold': True, 'font_color': 'white', 'fg_color': '#002060', 'border': 1})
                default_header_format = workbook.add_format({'bold': True, 'fg_color': '#FFC000', 'border': 1})

                # Aplicar formato a las cabeceras (reescribiéndolas con estilo)
                header = reporte_agencia.columns.tolist() # type: ignore
                for col_idx, header_text in enumerate(header):
                    if header_text.startswith('PENALIDAD 1 -'):
                        worksheet_reporte.write(0, col_idx, header_text, header_penalidad_format)
                    elif header_text.startswith('CLAWBACK 1 -'):
                        worksheet_reporte.write(0, col_idx, header_text, header_clawback_format)
                    else:
                        worksheet_reporte.write(0, col_idx, header_text, default_header_format)

                # Aplicar formato de porcentaje a columnas específicas
                cols_to_format_percent = ['Cumplimiento Altas %', 'CLAWBACK 1 - Cumplimiento Corte 2 %']
                for col_name in cols_to_format_percent:
                    try:
                        col_idx = header.index(col_name)
                        # Parámetros: primera_col, ultima_col, ancho, formato
                        worksheet_reporte.set_column(col_idx, col_idx, 18, percent_format)
                    except ValueError:
                        # La columna no existe en este dataframe, se ignora para evitar errores.
                        pass
                # --- FIN: Aplicar formato estético ---
            
            nombre_archivo_limpio = "".join(c for c in agencia if c.isalnum() or c in (' ', '_')).rstrip()
            zf.writestr(f"Reporte Corte 2 {nombre_archivo_limpio}.xlsx", output_buffer.getvalue())

    log_output.append("--- FIN DEL PROCESO ---")
    zip_buffer.seek(0)
    return zip_buffer, log_output


# --- Interfaz de Usuario para la página de Reportes Lima Corte 2 ---
st.title("Segmentador de Reportes - Lima Corte 2")
st.markdown("Sube el archivo consolidado de **Lima CORTE 2** para generar los reportes individuales por agencia.")
st.warning("Asegúrate de que el archivo tenga las hojas 'Reporte CORTE 2' y 'BASE', y que las cabeceras del reporte estén en las dos primeras filas.")

uploaded_file = st.file_uploader("Sube tu archivo Excel de CORTE 2", type=["xlsx"], key="lima_corte_2_uploader")

if uploaded_file is not None:
    st.success(f"Archivo '{uploaded_file.name}' cargado exitosamente.")
    if st.button("Procesar y Generar Reportes de Corte 2", type="primary"):
        with st.spinner("Procesando... La lectura de cabeceras complejas puede tardar un poco."):
            zip_file, log_data = procesar_reporte_corte_2(uploaded_file)
        
        if zip_file:
            st.success("¡Proceso completado!")
            st.subheader("Log de Validación del Proceso")
            st.text_area("Resultado de la validación:", "\n".join(log_data), height=300)
            st.subheader("Descargar Resultados")
            st.download_button(
                label="Descargar todos los reportes de Corte 2 (.zip)",
                data=zip_file,
                file_name=f"Reportes_Lima_Corte_2_Segmentados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip"
            )
        else:
            st.error("Ocurrió un error al validar o procesar el archivo. Por favor, revisa los detalles a continuación.")
            st.subheader("Log de Errores")
            st.text_area("Detalles del error:", "\n".join(log_data), height=300) 