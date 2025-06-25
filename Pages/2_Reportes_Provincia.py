# pages/2_Reportes_Provincia.py
import streamlit as st
import pandas as pd
import io
import zipfile
import re # Necesitamos importar la librería de expresiones regulares
from datetime import datetime

# --- Las funciones de lógica (validar_cabeceras, procesar_reportes_provincia) no necesitan cambios ---
# Las dejamos tal como estaban en la versión anterior.
def validar_cabeceras_provincia(archivo_excel, nombre_hoja, cabeceras_esperadas):
    try:
        df_primera_fila = pd.read_excel(archivo_excel, sheet_name=nombre_hoja, header=None, nrows=1)
        cabeceras_reales = [str(col).strip().upper() for col in df_primera_fila.iloc[0].values]
        for cabecera in cabeceras_esperadas:
            if cabecera.upper() not in cabeceras_reales: return False
        return True
    except Exception: return False


def normalizar_nombre(nombre):
    if not isinstance(nombre, str): return ""
    nombre_limpio = nombre.upper().replace('.', '').replace(',', '').replace('-', '')
    return re.sub(r'\s+', ' ', nombre_limpio).strip()

def get_agencia_base(nombre_completo, lista_departamentos):
    if not isinstance(nombre_completo, str): return ""
    nombre_completo_norm = normalizar_nombre(nombre_completo)
    for depto in lista_departamentos:
        if nombre_completo_norm.endswith(normalizar_nombre(depto)):
            nombre_base = nombre_completo[:-len(depto)].strip()
            return nombre_base
    return nombre_completo.strip()

def procesar_reportes_provincia(archivo_excel_cargado, zona_seleccionada):
    log_output = []
    log_output.append(f"--- INICIO DEL PROCESO PARA ZONA: {zona_seleccionada} ---")

    # ... (Validación de cabeceras no cambia) ...
    cabeceras_reporte = ['AGENCIA', 'RUC', 'ALTAS']
    if not validar_cabeceras_provincia(archivo_excel_cargado, 'Reporte CORTE 1', cabeceras_reporte):
        log_output.append("ALERTA: Cabeceras esperadas no encontradas en la hoja 'Reporte CORTE 1'.")
        return None, log_output
    cabeceras_base = ['COD_PEDIDO', 'ASESOR', 'ZONA', 'DEPARTAMENTO']
    if not validar_cabeceras_provincia(archivo_excel_cargado, 'BASE', cabeceras_base):
        log_output.append("ALERTA: Cabeceras esperadas no encontradas en la hoja 'BASE'.")
        return None, log_output
    log_output.append("Validación de cabeceras exitosa.")

    # ==============================================================================
    # === NUEVO: MAPA DE ALIAS PARA ASESORES ===
    # Se define aquí qué nombres de asesor en la BASE corresponden a una misma agencia.
    # Los nombres deben estar NORMALIZADOS (mayúsculas, sin puntos, etc.)
    # ==============================================================================
    mapeo_asesor_alias = {
        'EXPORTEL SAC': ['EXPORTEL SAC', 'EXPORTEL PROVINCIA']
        # Si tienes otros casos, los puedes añadir aquí. Ejemplo:
        # 'OTRA AGENCIA': ['OTRA AGENCIA', 'OTRA AGENCIA SOPORTE']
    }

    try:
        # ... (La lógica de lectura y filtrado inicial no cambia) ...
        log_output.append("Leyendo datos completos del archivo...")
        df_reporte_total = pd.read_excel(archivo_excel_cargado, sheet_name='Reporte CORTE 1', dtype=str)
        df_base_total = pd.read_excel(archivo_excel_cargado, sheet_name='BASE', dtype=str)
        df_reporte_total.columns = df_reporte_total.columns.str.strip().str.upper()
        df_base_total.columns = df_base_total.columns.str.strip().str.upper()
        base_filtrada_por_zona = df_base_total[df_base_total['ZONA'].str.strip().str.upper() == zona_seleccionada.upper()]
        if base_filtrada_por_zona.empty:
            log_output.append(f"ALERTA: No se encontraron registros en la hoja 'BASE' para la zona '{zona_seleccionada}'.")
            return None, log_output
        
        # Corrección: Aseguramos que trabajamos con una Serie de Pandas
        lista_departamentos = pd.Series(base_filtrada_por_zona['DEPARTAMENTO']).dropna().unique().tolist()
        lista_departamentos.sort(key=len, reverse=True)
        df_reporte_total['AGENCIA_BASE'] = df_reporte_total['AGENCIA'].apply(lambda x: get_agencia_base(x, lista_departamentos))
        
        # Aplicamos la misma corrección para futuras operaciones
        asesores_normalizados = pd.Series(base_filtrada_por_zona['ASESOR']).apply(normalizar_nombre)
        base_filtrada_por_zona = base_filtrada_por_zona.assign(ASESOR_NORMALIZADO=asesores_normalizados)

        agencias_de_la_zona = base_filtrada_por_zona['ASESOR_NORMALIZADO'].dropna().unique().tolist()
        
        # Continuamos con la lógica, asegurando el tipo correcto donde sea necesario
        df_reporte_total['AGENCIA_BASE_NORMALIZADA'] = df_reporte_total['AGENCIA_BASE'].apply(normalizar_nombre)
        reporte_filtrado_por_zona = df_reporte_total[df_reporte_total['AGENCIA_BASE_NORMALIZADA'].isin(agencias_de_la_zona)].copy()
        if reporte_filtrado_por_zona.empty:
            log_output.append(f"ALERTA: No se encontraron datos en la hoja 'Reporte CORTE 1' para las agencias de la zona '{zona_seleccionada}'.")
            return None, log_output
    except Exception as e:
        log_output.append(f"ERROR: No se pudo leer o filtrar el archivo Excel. Error: {e}")
        return None, log_output

    reporte_filtrado_por_zona['ALTAS'] = pd.to_numeric(reporte_filtrado_por_zona['ALTAS'])
    
    try:
        columnas_base_original = list(base_filtrada_por_zona.columns)
        indice_final = columnas_base_original.index('RECIBO1_PAGADO')
        columnas_a_mantener_en_base = columnas_base_original[:indice_final + 1]
        if 'ZONA' not in columnas_a_mantener_en_base: columnas_a_mantener_en_base.append('ZONA')
        if 'ASESOR_NORMALIZADO' not in columnas_a_mantener_en_base: columnas_a_mantener_en_base.append('ASESOR_NORMALIZADO')
    except ValueError as e:
        log_output.append(f"ERROR: No se encontró una columna esencial como 'RECIBO1_PAGADO'. Error: {e}")
        return None, log_output
        
    agencias_base_a_procesar = pd.Series(reporte_filtrado_por_zona['AGENCIA_BASE_NORMALIZADA']).dropna().unique().tolist()
    log_output.append(f"Se van a generar reportes para {len(agencias_base_a_procesar)} agencias base (normalizadas).")
    
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for agencia_base_norm in agencias_base_a_procesar:
            reporte_agencia = reporte_filtrado_por_zona[reporte_filtrado_por_zona['AGENCIA_BASE_NORMALIZADA'] == agencia_base_norm].copy()
            
            # ==============================================================================
            # === MEJORA CLAVE: Usamos el mapa de alias para buscar en la BASE ===
            # ==============================================================================
            if agencia_base_norm in mapeo_asesor_alias:
                nombres_a_buscar = mapeo_asesor_alias[agencia_base_norm]
                base_agencia = base_filtrada_por_zona[base_filtrada_por_zona['ASESOR_NORMALIZADO'].isin(nombres_a_buscar)].copy()
            else:
                # Si no está en el mapa, se usa la lógica normal
                base_agencia = base_filtrada_por_zona[base_filtrada_por_zona['ASESOR_NORMALIZADO'] == agencia_base_norm].copy()
            
            base_agencia_sin_asesor = pd.DataFrame(base_agencia).drop(columns=['ASESOR_NORMALIZADO'], errors='ignore')
            base_agencia_final = base_agencia_sin_asesor[columnas_a_mantener_en_base[:-1]]
            
            try:
                altas_reporte = reporte_agencia['ALTAS'].sum()
                registros_base = len(base_agencia_final)
                if altas_reporte == registros_base:
                    log_output.append(f"ÉXITO    | {agencia_base_norm:<40} | ALTAS: {altas_reporte:<5} | Registros BASE: {registros_base:<5} | OK")
                else:
                    log_output.append(f"DESCUADRE | {agencia_base_norm:<40} | ALTAS: {altas_reporte:<5} | Registros BASE: {registros_base:<5} | REVISAR")
            except Exception as e:
                log_output.append(f"Error validando la agencia '{agencia_base_norm}': {e}")
                
            nombre_original_agencia = pd.Series(reporte_agencia['AGENCIA_BASE']).iloc[0]
            output_buffer = io.BytesIO()
            with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer: # type: ignore
                # Corrección final: guardar el resultado de drop en una variable intermedia
                reporte_agencia_final = pd.DataFrame(reporte_agencia).drop(columns=['AGENCIA_BASE', 'AGENCIA_BASE_NORMALIZADA'], errors='ignore')
                reporte_agencia_final.to_excel(writer, sheet_name='Reporte Agencia', index=False)
                base_agencia_final.to_excel(writer, sheet_name='BASE', index=False)
            zf.writestr(f"Reporte {nombre_original_agencia.strip()}.xlsx", output_buffer.getvalue())
            
    log_output.append("--- FIN DEL PROCESO ---")
    zip_buffer.seek(0)
    return zip_buffer, log_output



# --- Interfaz de Usuario para la página de Reportes Provincia ---
st.title("Segmentador de Reportes - Provincia")
st.markdown("Sube el archivo consolidado de Provincia para generar los reportes por zona.")

# --- LÓGICA DE UI MEJORADA ---
# 1. El usuario sube el archivo PRIMERO.
uploaded_file = st.file_uploader("1. Sube tu archivo Excel de reportes de Provincia", type=["xlsx"], key="provincia_uploader")

# 2. Si el archivo se sube, LEEMOS las zonas y MOSTRAMOS el menú desplegable.
if uploaded_file is not None:
    try:
        # Hacemos una lectura rápida solo de la columna ZONA para obtener las opciones.
        df_zonas = pd.read_excel(uploaded_file, sheet_name='BASE', usecols=['ZONA'])
        # Obtenemos los valores únicos, eliminamos nulos y los convertimos a una lista.
        lista_zonas_dinamica = df_zonas['ZONA'].dropna().unique().tolist()

        if not lista_zonas_dinamica:
            st.warning("No se encontraron zonas en la columna 'ZONA' de la hoja 'BASE' del archivo subido.")
        else:
            st.info(f"Zonas detectadas en el archivo: {', '.join(lista_zonas_dinamica)}")
            zona_seleccionada = st.selectbox(
                "2. Selecciona la Zona a procesar",
                options=lista_zonas_dinamica,
                index=None,
                placeholder="Elige una de las zonas detectadas"
            )

            # 3. Si el usuario selecciona una zona, MOSTRAMOS el botón para procesar.
            if zona_seleccionada:
                if st.button("Procesar y Generar Reportes de Provincia", type="primary"):
                    with st.spinner(f"Procesando {zona_seleccionada}..."):
                        # Pasamos el archivo cargado, que ya está en memoria.
                        zip_file, log_data = procesar_reportes_provincia(uploaded_file, zona_seleccionada)
                    if zip_file:
                        st.success("¡Proceso completado!")
                        st.subheader("Log de Validación del Proceso")
                        st.text_area("Resultado:", "\n".join(log_data), height=300)
                        st.subheader("Descargar Resultados")
                        st.download_button(
                            label=f"Descargar reportes de {zona_seleccionada} (.zip)",
                            data=zip_file,
                            file_name=f"Reportes_{zona_seleccionada.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                            mime="application/zip"
                        )
                    else:
                        st.error("Ocurrió un error. Revisa los detalles a continuación.")
                        st.text_area("Log de Errores:", "\n".join(log_data), height=300)

    except Exception as e:
        st.error(f"No se pudo procesar el archivo. ¿Estás seguro de que tiene una hoja 'BASE' con una columna 'ZONA'? Error: {e}")