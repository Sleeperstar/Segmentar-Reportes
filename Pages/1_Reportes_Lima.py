# pages/1_Reportes_Lima.py
import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime

# --- Las funciones de lógica no cambian ---
def validar_cabeceras(archivo_excel, nombre_hoja, cabeceras_esperadas):
    try:
        df_primera_fila = pd.read_excel(archivo_excel, sheet_name=nombre_hoja, header=None, nrows=1)
        cabeceras_reales = [str(col).strip().upper() for col in df_primera_fila.iloc[0].values]
        for cabecera in cabeceras_esperadas:
            if cabecera.upper() not in cabeceras_reales: return False
        return True
    except Exception: return False

def procesar_archivos_excel(archivo_excel_cargado):
    log_output = []
    log_output.append("--- INICIO DEL PROCESO DE SEGMENTACIÓN Y VALIDACIÓN ---")
    cabeceras_esenciales_reporte = ['AGENCIA', 'RUC', 'ALTAS', 'TOTAL A PAGAR']
    if not validar_cabeceras(archivo_excel_cargado, 'Reporte CORTE 1', cabeceras_esenciales_reporte):
        log_output.append("ALERTA DE ARCHIVO: Las cabeceras esperadas (como 'AGENCIA', 'RUC', etc.) no se encontraron en la primera fila de la hoja 'Reporte CORTE 1'.")
        log_output.append("Por favor, asegúrese de que los encabezados de su reporte estén en la Fila 1 del archivo Excel y vuelva a intentarlo.")
        return None, log_output
    cabeceras_esenciales_base = ['COD_PEDIDO', 'DNI_CLIENTE', 'ASESOR']
    if not validar_cabeceras(archivo_excel_cargado, 'BASE', cabeceras_esenciales_base):
        log_output.append("ALERTA DE ARCHIVO: Las cabeceras esperadas (como 'COD_PEDIDO', 'ASESOR', etc.) no se encontraron en la primera fila de la hoja 'BASE'.")
        log_output.append("Por favor, asegúrese de que los encabezados de su base estén en la Fila 1 del archivo Excel y vuelva a intentarlo.")
        return None, log_output
    log_output.append("Validación de cabeceras exitosa. Los encabezados se encontraron en la primera fila.")
    try:
        log_output.append("Leyendo datos completos del archivo...")
        df_reporte_total = pd.read_excel(archivo_excel_cargado, sheet_name='Reporte CORTE 1')
        df_base_total = pd.read_excel(archivo_excel_cargado, sheet_name='BASE')
        df_reporte_total.columns = df_reporte_total.columns.str.strip().str.upper()
        df_base_total.columns = df_base_total.columns.str.strip().str.upper()
        log_output.append("Nombres de columnas estandarizados (sin espacios y en mayúsculas).")
    except Exception as e:
        log_output.append(f"ERROR: No se pudo leer el archivo Excel. Error: {e}")
        return None, log_output
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        agencias_a_procesar = df_reporte_total['AGENCIA'].dropna().unique().tolist()
        log_output.append(f"Se encontraron {len(agencias_a_procesar)} agencias únicas para procesar.")
        mapeo_agencias_alias = {"EXPORTEL S.A.C.": ["EXPORTEL S.A.C.", "EXPORTEL PROVINCIA"]}
        try:
            columnas_base_deseadas = df_base_total.columns.tolist()
            indice_final = columnas_base_deseadas.index('RECIBO1_PAGADO')
            columnas_a_mantener_en_base = columnas_base_deseadas[:indice_final + 1]
        except ValueError:
            log_output.append("ERROR: La columna 'RECIBO1_PAGADO' no se encontró en la hoja 'BASE'.")
            return None, log_output
        for agencia in agencias_a_procesar:
            reporte_agencia = df_reporte_total[df_reporte_total['AGENCIA'] == agencia].copy()
            if reporte_agencia.empty: continue
            if agencia in mapeo_agencias_alias:
                nombres_a_buscar = mapeo_agencias_alias[agencia]
                base_agencia = df_base_total[df_base_total['ASESOR'].isin(nombres_a_buscar)]
            else:
                base_agencia = df_base_total[df_base_total['ASESOR'] == agencia]
            base_agencia_final = base_agencia[columnas_a_mantener_en_base]
            try:
                altas_reporte = int(reporte_agencia.iloc[0]['ALTAS'])
                registros_base = len(base_agencia_final)
                if altas_reporte == registros_base: log_output.append(f"ÉXITO    | {agencia:<40} | ALTAS: {altas_reporte:<5} | Registros BASE: {registros_base:<5} | OK")
                else: log_output.append(f"DESCUADRE | {agencia:<40} | ALTAS: {altas_reporte:<5} | Registros BASE: {registros_base:<5} | REVISAR")
            except Exception as e: log_output.append(f"Error validando la agencia '{agencia}': {e}")
            output_buffer = io.BytesIO()
            with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer: # type: ignore
                reporte_agencia.to_excel(writer, sheet_name='Reporte Agencia', index=False) # type: ignore
                base_agencia_final.to_excel(writer, sheet_name='BASE', index=False) # type: ignore
                workbook, worksheet = writer.book, writer.sheets['Reporte Agencia']
                percent_format, number_format = workbook.add_format({'num_format': '0.00%'}), workbook.add_format({'num_format': '#,##0.00'})
                header = list(reporte_agencia.columns)
                try:
                    worksheet.set_column(header.index('CUMPLIMIENTO ALTAS %'), header.index('CUMPLIMIENTO ALTAS %'), 18, percent_format)
                    worksheet.set_column(header.index('TOTAL A PAGAR'), header.index('TOTAL A PAGAR'), 18, number_format)
                except ValueError: pass
            nombre_archivo_limpio = "".join(c for c in agencia if c.isalnum() or c in (' ', '_')).rstrip()
            zf.writestr(f"Reporte {nombre_archivo_limpio}.xlsx", output_buffer.getvalue())
    log_output.append("--- FIN DEL PROCESO ---")
    zip_buffer.seek(0)
    return zip_buffer, log_output


# --- Interfaz de Usuario para la página de Reportes Lima ---
st.title("Segmentador de Reportes - Lima")
st.markdown("Sube el archivo consolidado de Lima para generar los reportes individuales por agencia.")

uploaded_file = st.file_uploader("Sube tu archivo Excel de reportes de Lima", type=["xlsx"], key="lima_uploader")
if uploaded_file is not None:
    st.success(f"Archivo '{uploaded_file.name}' cargado exitosamente.")
    if st.button("Procesar y Generar Reportes", type="primary"):
        with st.spinner("Procesando... Esto puede tardar unos minutos para archivos grandes."):
            zip_file, log_data = procesar_archivos_excel(uploaded_file)
        if zip_file:
            st.success("¡Proceso completado!")
            st.subheader("Log de Validación del Proceso")
            st.text_area("Resultado de la validación:", "\n".join(log_data), height=300)
            st.subheader("Descargar Resultados")
            st.download_button(label="Descargar todos los reportes (.zip)", data=zip_file,
                              file_name=f"Reportes_Lima_Segmentados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                              mime="application/zip")
        else:
            st.error("Ocurrió un error al validar el archivo. Por favor, revisa los detalles a continuación.")
            st.subheader("Log de Errores")
            st.text_area("Detalles del error:", "\n".join(log_data), height=300)