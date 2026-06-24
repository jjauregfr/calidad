import streamlit as st
import os

# 1. Configuración global (Debe ser lo primero)
st.set_page_config(
    page_title="Centro de Control de Automatizaciones | Calidad",
    page_icon="⚡",
    layout="wide"
)

# 2. Rutas absolutas a los archivos
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Rutas anteriores
ruta_excel = os.path.join(BASE_DIR, "pages", "unificar_excel.py")

# NUEVAS RUTAS (Asegúrate de que los archivos existan en tu carpeta 'views')
ruta_auditor = os.path.join(BASE_DIR, "pages", "auditor_sql.py")
ruta_mail_planvital = os.path.join(BASE_DIR, "pages", "PlanvitalCorreo.py")
ruta_mail_capital = os.path.join(BASE_DIR, "pages", "CapitalCorreo.py")
ruta_gestor_capital = os.path.join(BASE_DIR, "pages", "gestor_doctos_capital.py")

# ==========================================
# FUNCION DEL HUB PRINCIPAL (INTERFAZ)
# ==========================================
def mostrar_hub():
    """Esta función contiene toda la interfaz visual de tu Home / Hub"""
    st.markdown("""
        <style>
        .hub-title { text-align: center; color: #1F497D; font-family: 'Segoe UI', sans-serif; }
        .tool-card {
            background-color: #ffffff;
            padding: 24px;
            border-radius: 12px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
            border-top: 4px solid #4F81BD;
            margin-bottom: 20px;
            transition: transform 0.2s;
            height: 180px; /* Altura fija para que todas las tarjetas se vean alineadas */
        }
        .tool-card:hover {
            transform: translateY(-5px);
        }
        </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='hub-title'>⚡Asistente Proceso Despliegue </h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: gray;'>Plataforma de asistencia al proceso de Gestión de Despliegue</p>", unsafe_allow_html=True)
    st.markdown("---")

    # -----------------------------------------------------------------
    # FILA 1: Herramientas Originales
    # -----------------------------------------------------------------
    col1, col2, col3 = st.columns(3)

    with col1:
       st.markdown("""
            <div class="tool-card" style="border-top-color: #34495E;">
                <h4>🗂️ Gestor de Archivos Capital</h4>
                <p style='color: #555; font-size: 14px;'>Organizador automatizado, renombrador y validador de estructuras de archivos para Capital.</p>
            </div>
        """, unsafe_allow_html=True)
       if st.button("Abrir Gestor Capital 📂", key="btn_gestor_capital", use_container_width=True):
            st.switch_page(pagina_gestor_capital)

    with col2:
        st.markdown("""
            <div class="tool-card" style="border-top-color: #2ECC71;">
                <h4>📁 Unificador Excel Genérico</h4>
                <p style='color: #555; font-size: 14px;'>Une múltiples libros de cálculo planos en un único DataFrame consolidado de manera masiva.</p>
            </div>
        """, unsafe_allow_html=True)
        if st.button("Abrir Unificador Excel 📂", key="btn_unificar", use_container_width=True):
            st.switch_page(pagina_excel)

    with col3:
        st.markdown("""
            <div class="tool-card" style="border-top-color: #9B59B6;">
                <h4>🔍 Calidad - Auditor SQL</h4>
                <p style='color: #555; font-size: 14px;'>Auditoría automatizada de bases de datos y cruces de información mediante queries SQL estructuradas.</p>
            </div>
        """, unsafe_allow_html=True)
        if st.button("Abrir Auditor SQL 🛠️", key="btn_auditor", use_container_width=True):
            st.switch_page(pagina_auditor)

    st.markdown("<br>", unsafe_allow_html=True) # Espacio entre filas

    # -----------------------------------------------------------------
    # FILA 2: Herramientas de Correos y Gestión
    # -----------------------------------------------------------------
    col4, col5, col6 = st.columns(3)

    with col4:
        st.markdown("""
            <div class="tool-card" style="border-top-color: #E67E22;">
                <h4>✉️ Reporte Semanal Planvital</h4>
                <p style='color: #555; font-size: 14px;'>Automatización, procesamiento de correos (Catalogación en TEST) y generación de archivo semanal para AFP Planvital.</p>
            </div>
        """, unsafe_allow_html=True)
        if st.button("Abrir Correos Planvital 📨", key="btn_mail_planvital", use_container_width=True):
            st.switch_page(pagina_mail_planvital)

    with col5:
        st.markdown("""
            <div class="tool-card" style="border-top-color: #E74C3C;">
                <h4>✉️ Reporte Semanal Capital</h4>
                <p style='color: #555; font-size: 14px;'>Automatización, procesamiento de correos (SCRIPT EN PROD) y generación de archivo semanal AFP Capital.</p>
            </div>
        """, unsafe_allow_html=True)
        if st.button("Abrir Correos Capital 📧", key="btn_mail_capital", use_container_width=True):
            st.switch_page(pagina_mail_capital)

   
    st.markdown("<br>", unsafe_allow_html=True) # Espacio entre filas

    # -----------------------------------------------------------------
    # FILA 3: Herramientas de Correos y Gestión
    # -----------------------------------------------------------------
    col7, col8, col9 = st.columns(3)

            
    st.markdown("---")
    st.caption("🔒 Acceso restringido para el equipo de Arquitectura | Generado en jjf")


# ==========================================
# CONFIGURACIÓN DE NAVEGACIÓN
# ==========================================
pagina_principal = st.Page(mostrar_hub, title="Asistente Proceso Despliegue", icon="⚡", default=True)
pagina_excel = st.Page(ruta_excel, title="Unificador Excel Genérico", icon="📁")

# NUEVAS PÁGINAS EN EL MENÚ IZQUIERDO
pagina_auditor = st.Page(ruta_auditor, title="Calidad - Auditor SQL", icon="🔍")
pagina_mail_planvital = st.Page(ruta_mail_planvital, title="Calidad - Reporte Semanal Planvital", icon="✉️")
pagina_mail_capital = st.Page(ruta_mail_capital, title="Calidad - Reporte Semanal Capital", icon="📨")
pagina_gestor_capital = st.Page(ruta_gestor_capital, title="Calidad - Gestor Archivos Capital", icon="🗂️")

# Inicializamos la navegación con TODAS las páginas
pg = st.navigation([
    pagina_principal, 
    pagina_excel, 
    pagina_auditor, 
    pagina_mail_planvital, 
    pagina_mail_capital, 
    pagina_gestor_capital,
   
])

pg.run()