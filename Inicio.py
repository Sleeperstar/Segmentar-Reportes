# streamlit_app: Inicio
# app.py (Página Principal)
import streamlit as st

# Configuración general de la página. Esto se aplicará a todas las páginas.
st.set_page_config(
    page_title="Automatización de Reportes",
    page_icon="📄", # Puedes usar un emoji como ícono
    layout="wide"
)

# --- BARRA LATERAL (SIDEBAR) ---
# El logo se mostrará aquí y será visible en todas las páginas.
with st.sidebar:
    st.image("logo.png", width=200)
    st.title("Menú de Navegación")
    st.info("Selecciona el tipo de reporte que deseas procesar en el menú de arriba.")

# --- CONTENIDO PRINCIPAL DE LA PÁGINA DE BIENVENIDA ---
st.title("Bienvenido a la Herramienta de Automatización del cortee de los Reportes")

st.markdown("---")

st.header("Instrucciones de Uso")
st.markdown("""
1.  **Navega a la sección deseada** usando el menú de la izquierda (por ejemplo, `Reportes Lima`).
2.  **Sube el archivo Excel** consolidado cuando se te solicite.
3.  Haz clic en el botón **"Procesar y Generar Reportes"**.
4.  **Espera** a que la herramienta valide y segmente los datos.
5.  **Descarga el archivo .zip** con todos los reportes individuales.
""")

st.markdown("---")
st.success("¡Comienza seleccionando una opción en la barra lateral izquierda!")