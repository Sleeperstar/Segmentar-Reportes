import pandas as pd
import os
import logging

def setup_logging(log_path):
    # ... (Esta función no cambia)
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s',
                        handlers=[logging.FileHandler(log_path, mode='w'), logging.StreamHandler()])

def segmentar_y_validar_reportes():
    # --- 1. CONFIGURACIÓN INICIAL ---
    ruta_archivo_maestro = r"C:\Users\jhamp\WIN LOCAL-AUTOMATIZACIO\Segmentar Reportes\Reportes AGENCIA LIMA Corte 1 MAYO 2025 V2.xlsx"
    directorio_salida = r"C:\Users\jhamp\WIN LOCAL-AUTOMATIZACIO\Segmentar Reportes\Reportes Segmentados"
    ruta_log = os.path.join(directorio_salida, "validation_log.txt")

    os.makedirs(directorio_salida, exist_ok=True)
    setup_logging(ruta_log)
    
    agencias_a_procesar = [
        'AKLLA DISTRIBUIDORES SAC', 'ALIV TELECOM S.A.C.', 'C & C SALES EIRL',
        'CORPORACION DE TODO PERU S.A.C', 'CORPORACION VISUAL CONNECTIONS S.A.C.',
        'DAS SOLUCIONES S.A.C.', 'DATANTENNA S.A.C.', 'EXPORTEL S.A.C.',
        'FUTURA CONNECTION', 'S L SELECTRA', 'TELECENTER S.A.C.',
        'TELENET CALL SAC', 'A & C HYBRIDCOM E.I.R.L.', 'NOVO NETWORKS SAC',
        'TELCO CONTACT S.A.C', 'COMUNICACIONES CLV CENTER S.A.', 'LEAD',
        'UPSALES S.A.C.', 'JOMA', 'CMBRANDO VENTAS S.A.C.', 'BELFECOM E.I.R.L.',
        'GRP SOLUTIONS S.A.C'
    ]

    # =================================================================================
    # === NUEVO: MAPA DE ALIAS DE AGENCIAS ===
    # Aquí defines el nombre principal y una lista de todos los nombres que le corresponden en la BASE.
    # =================================================================================
    mapeo_agencias_alias = {
        "EXPORTEL S.A.C.": ["EXPORTEL S.A.C.", "EXPORTEL PROVINCIA"]
        # Si en el futuro tienes otro caso, solo añádelo aquí. Ejemplo:
        # "OTRA AGENCIA S.A.C.": ["OTRA AGENCIA S.A.C.", "OTRA AGENCIA LIMA", "OTRA AGENCIA SUR"]
    }

    # --- 2. LECTURA DE DATOS ---
    logging.info("--- INICIO DEL PROCESO DE SEGMENTACIÓN Y VALIDACIÓN ---")
    # ... (El bloque de lectura de datos no cambia) ...
    try:
        logging.info("Leyendo archivo maestro...")
        df_reporte_total = pd.read_excel(ruta_archivo_maestro, sheet_name='Reporte CORTE 1')
        df_base_total = pd.read_excel(ruta_archivo_maestro, sheet_name='BASE')
        logging.info("Lectura completada exitosamente.")
    except Exception as e:
        logging.error(f"No se pudo leer el archivo Excel. Error: {e}")
        return
        
    # --- 3. DEFINICIÓN DE COLUMNAS PARA LA BASE ---
    # ... (Este bloque no cambia) ...
    try:
        columnas_base_deseadas = df_base_total.columns.tolist()
        indice_final = columnas_base_deseadas.index('RECIBO1_PAGADO')
        columnas_a_mantener_en_base = columnas_base_deseadas[:indice_final + 1]
    except ValueError:
        logging.error("La columna 'RECIBO1_PAGADO' no se encontró en la hoja 'BASE'.")
        return

    # --- 4. PROCESAMIENTO, VALIDACIÓN Y ESCRITURA ---
    for agencia in agencias_a_procesar:
        # Filtrar el reporte principal (esto no cambia)
        reporte_agencia = df_reporte_total[df_reporte_total['AGENCIA'] == agencia].copy()

        # =================================================================================
        # === LÓGICA DE FILTRADO MEJORADA USANDO EL MAPA DE ALIAS ===
        # =================================================================================
        # Revisa si la agencia actual tiene alias definidos en el mapa
        if agencia in mapeo_agencias_alias:
            nombres_a_buscar = mapeo_agencias_alias[agencia]
            base_agencia = df_base_total[df_base_total['ASESOR'].isin(nombres_a_buscar)]
        else:
            # Si no está en el mapa, usa el comportamiento normal (búsqueda exacta)
            base_agencia = df_base_total[df_base_total['ASESOR'] == agencia]
        
        base_agencia_final = base_agencia[columnas_a_mantener_en_base]

        # --- VALIDACIÓN ---
        # ... (El bloque de validación y log no cambia) ...
        if reporte_agencia.empty:
            logging.warning(f"Agencia '{agencia}' no encontrada en el reporte principal. Saltando...")
            continue
        try:
            altas_reporte = int(reporte_agencia.iloc[0]['ALTAS'])
            registros_base = len(base_agencia_final)
            if altas_reporte == registros_base:
                logging.info(f"ÉXITO    | {agencia:<40} | ALTAS: {altas_reporte:<5} | Registros BASE: {registros_base:<5} | OK")
            else:
                logging.warning(f"DESCUADRE | {agencia:<40} | ALTAS: {altas_reporte:<5} | Registros BASE: {registros_base:<5} | REVISAR")
        except Exception as e:
            logging.error(f"Error validando la agencia '{agencia}': {e}")


        # --- ESCRITURA Y FORMATO DEL EXCEL ---
        # ... (El resto del script para escribir y formatear no cambia) ...
        nombre_archivo_limpio = "".join(c for c in agencia if c.isalnum() or c in (' ', '_')).rstrip()
        ruta_salida_agencia = os.path.join(directorio_salida, f"Reporte {nombre_archivo_limpio}.xlsx")
        with pd.ExcelWriter(ruta_salida_agencia, engine='xlsxwriter') as writer:
            reporte_agencia.to_excel(writer, sheet_name='Reporte Agencia', index=False)
            base_agencia_final.to_excel(writer, sheet_name='BASE', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Reporte Agencia']
            percent_format = workbook.add_format({'num_format': '0.00%'})
            number_format = workbook.add_format({'num_format': '#,##0.00'})
            header = list(reporte_agencia.columns)
            try:
                cumplimiento_idx = header.index('Cumplimiento Altas %')
                worksheet.set_column(cumplimiento_idx, cumplimiento_idx, 18, percent_format)
                total_pagar_idx = header.index('TOTAL A PAGAR')
                worksheet.set_column(total_pagar_idx, total_pagar_idx, 18, number_format)
            except ValueError as e:
                logging.error(f"No se pudo encontrar una columna para formatear en '{agencia}': {e}")
    
    logging.info("--- FIN DEL PROCESO ---")

# --- INICIAR EL SCRIPT ---
if __name__ == "__main__":
    segmentar_y_validar_reportes()