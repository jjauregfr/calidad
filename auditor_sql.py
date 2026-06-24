import streamlit as st
import os
import re
import pandas as pd
from io import BytesIO
st.set_page_config(page_title="Auditor SQL Avanzado - CALIDAD", layout="wide", page_icon="🛡️")


def auditar_contenido_sql(contenido):
    """
    Analiza el script línea por línea capturando errores exactos.
    Clasifica de forma específica los INSERT, UPDATE y DELETE sin esquema.
    """
    errores_detallados = []
    version_hallada = "No detectada"
    
    # Determinar clasificaciones principales del archivo
    es_un_package = bool(re.search(r'\bPACKAGE\b', contenido, re.IGNORECASE))
    
    # Evalúa si tiene la sentencia de creación de un package (especificación o cuerpo)
    tiene_creacion_pkg = bool(re.search(r'\bCREATE\s+(?:OR\s+REPLACE\s+)?PACKAGE\b', contenido, re.IGNORECASE))
    es_sql_ejecucion = not tiene_creacion_pkg  # Si no crea un PKG, es un SQL de ejecución

    # Asignación de valores por defecto basados en el tipo de script
    tiene_check_version = "❌" if not es_sql_ejecucion else "No Aplica"
    termina_con_barra = "❌" if not es_sql_ejecucion else "No Aplica"
    tiene_pkg_esquema = "❌" if es_un_package else "No Aplica"
    tiene_body_esquema = "❌" if es_un_package else "No Aplica"
    
    # Contadores específicos por archivo
    num_updates_sin_esq = 0
    num_inserts_sin_esq = 0
    num_deletes_sin_esq = 0
    
    tablas_con_esquema = set()
    tablas_sin_esquema = set()
    
    esquemas_validos = r'(?:NEWAFP|SNDFCOFRE|NEWBYP)'
    palabras_ignorar = {'DUAL', 'SELECT', 'WHERE', 'AND', 'OR', 'SET', 'VALUES', 'BEGIN', 'EXCEPTION', 'INTO', 'ON'}
    
    # 1. Buscar Versión (Solo aplica si NO es un SQL de ejecución)
    if not es_sql_ejecucion:
        match_version = re.search(r"(?:VERSION|VERSIÓN)\s*[:\s]\s*([\d\.]+)", contenido, re.IGNORECASE)
        if match_version:
            version_hallada = match_version.group(1)
            tiene_check_version = "✅"
        else:
            errores_detallados.append("Línea N/A: No se encontró la cabecera de versión ('VERSION : X.X.X').")
    else:
        version_hallada = "No Aplica"

    # 2. Checks de Estructura de Packages
    if es_un_package:
        # Validar especificación de Package (CREATE PACKAGE...)
        tiene_especificacion = bool(re.search(r'\bCREATE\s+(?:OR\s+REPLACE\s+)?PACKAGE\s+(?!BODY\b)', contenido, re.IGNORECASE))
        if tiene_especificacion:
            patron_estricto_pkg = rf'\bCREATE\s+(?:OR\s+REPLACE\s+)?PACKAGE\s+{esquemas_validos}\.'
            tiene_pkg_esquema = "✅" if re.search(patron_estricto_pkg, contenido, re.IGNORECASE) else "❌"
            if tiene_pkg_esquema == "❌":
                errores_detallados.append("Línea CREATE: La especificación del PACKAGE no contiene un esquema válido.")
        else:
            tiene_pkg_esquema = "No Aplica"

        # Validar cuerpo de Package (CREATE PACKAGE BODY...)
        tiene_cuerpo = bool(re.search(r'\bCREATE\s+(?:OR\s+REPLACE\s+)?PACKAGE\s+BODY\b', contenido, re.IGNORECASE))
        if tiene_cuerpo:
            patron_estricto_body = rf'\bCREATE\s+(?:OR\s+REPLACE\s+)?PACKAGE\s+BODY\s+{esquemas_validos}\.'
            tiene_body_esquema = "✅" if re.search(patron_estricto_body, contenido, re.IGNORECASE) else "❌"
            if tiene_body_esquema == "❌":
                errores_detallados.append("Línea CREATE BODY: El cuerpo del PACKAGE BODY no contiene un esquema válido.")
        else:
            tiene_body_esquema = "No Aplica"

    # 3. Expresiones Regulares para capturas generales de Tablas
    patron_con_esquema = rf'\b(?:FROM|JOIN|UPDATE|DELETE\s+FROM|INSERT\s+INTO|ALTER\s+TABLE|CREATE\s+TABLE|CREATE\s+INDEX|ON)\s+({esquemas_validos})\.([a-zA-Z_]\w*)'
    patron_vars_con = rf'\b\w+\s+({esquemas_validos})\.([a-zA-Z_]\w*)(?:\.\w+)?\s*%(?:ROWTYPE|TYPE)'
    patron_vars_sin = rf'\b\w+\s+(?!{esquemas_validos}\.)([a-zA-Z_]\w*)(?:\.\w+)?\s*%(?:ROWTYPE|TYPE)'

    # 4. Expresiones Regulares ATÓMICAS para DML sin esquema (Modificaciones)
    patron_update_sin = rf'\bUPDATE\s+(?!{esquemas_validos}\.)([a-zA-Z_]\w*)'
    patron_insert_sin = rf'\bINSERT\s+INTO\s+(?!{esquemas_validos}\.)([a-zA-Z_]\w*)'
    patron_delete_sin = rf'\bDELETE\s+FROM\s+(?!{esquemas_validos}\.)([a-zA-Z_]\w*)'
    patron_residual_sin = rf'\b(?:FROM|JOIN|ALTER\s+TABLE|CREATE\s+TABLE|CREATE\s+INDEX|ON)\s+(?!{esquemas_validos}\.)([a-zA-Z_]\w*)'

    # 5. Análisis línea por línea
    lineas = contenido.splitlines()
    for num_linea, linea in enumerate(lineas, start=1):
        linea_limpia = linea.strip()
        if not linea_limpia or linea_limpia.startswith('--'):  
            continue
            
        # --- Recolección CON Esquema ---
        for match in re.finditer(patron_con_esquema, linea, re.IGNORECASE):
            esq, esc_objeto = match.group(1), match.group(2)
            if esc_objeto.upper() not in palabras_ignorar and len(esc_objeto) > 2:
                tablas_con_esquema.add(f"{esq.upper()}.{esc_objeto.upper()}")
        for match in re.finditer(patron_vars_con, linea, re.IGNORECASE):
            esq, esc_objeto = match.group(1), match.group(2)
            if esc_objeto.upper() not in palabras_ignorar and len(esc_objeto) > 2:
                tablas_con_esquema.add(f"{esq.upper()}.{esc_objeto.upper()}")

        # --- Análisis Específico de Objetos SIN Esquema ---
        # SE OMITE POR COMPLETO SI EL ARCHIVO ES UN PACKAGE
        if not es_un_package:
            # UPDATE sin esquema
            for match in re.finditer(patron_update_sin, linea, re.IGNORECASE):
                objeto = match.group(1)
                if objeto.upper() not in palabras_ignorar and len(objeto) > 2:
                    tablas_sin_esquema.add(objeto.upper())
                    num_updates_sin_esq += 1
                    errores_detallados.append(f"Línea {num_linea}: [UPDATE] Tabla '{objeto}' sin esquema. Código: `{linea_limpia}`")

            # INSERT sin esquema
            for match in re.finditer(patron_insert_sin, linea, re.IGNORECASE):
                objeto = match.group(1)
                if objeto.upper() not in palabras_ignorar and len(objeto) > 2:
                    tablas_sin_esquema.add(objeto.upper())
                    num_inserts_sin_esq += 1
                    errores_detallados.append(f"Línea {num_linea}: [INSERT] Tabla '{objeto}' sin esquema. Código: `{linea_limpia}`")

            # DELETE sin esquema
            for match in re.finditer(patron_delete_sin, linea, re.IGNORECASE):
                objeto = match.group(1)
                if objeto.upper() not in palabras_ignorar and len(objeto) > 2:
                    tablas_sin_esquema.add(objeto.upper())
                    num_deletes_sin_esq += 1
                    errores_detallados.append(f"Línea {num_linea}: [DELETE] Tabla '{objeto}' sin esquema. Código: `{linea_limpia}`")

            # Otros sin esquema (FROM, JOIN, etc.)
            for match in re.finditer(patron_residual_sin, linea, re.IGNORECASE):
                objeto = match.group(1)
                if objeto.upper() not in palabras_ignorar and len(objeto) > 2:
                    tablas_sin_esquema.add(objeto.upper())
                    errores_detallados.append(f"Línea {num_linea}: Objeto/Tabla '{objeto}' sin esquema. Código: `{linea_limpia}`")
                    
            for match in re.finditer(patron_vars_sin, linea, re.IGNORECASE):
                objeto = match.group(1)
                if objeto.upper() not in palabras_ignorar and len(objeto) > 2:
                    tablas_sin_esquema.add(objeto.upper())
                    errores_detallados.append(f"Línea {num_linea}: Variable apunta a '{objeto}' sin esquema. Código: `{linea_limpia}`")

    # 6. Validar cierre del archivo (Solo aplica si NO es un SQL de ejecución)
    if not es_sql_ejecucion:
        if contenido.strip().endswith('/'):
            termina_con_barra = "✅"
        else:
            errores_detallados.append("Línea FINAL: El script no finaliza con '/'")

    num_con = len(tablas_con_esquema)
    num_sin = len(tablas_sin_esquema)

    txt_con_esquema = f"({num_con}) " + (", ".join(tablas_con_esquema) if num_con > 0 else "Ninguna")
    txt_sin_esquema = f"({num_sin}) " + (", ".join(tablas_sin_esquema) if num_sin > 0 else "✅ Todo OK")

    return (errores_detallados, version_hallada, tiene_check_version, tiene_pkg_esquema, tiene_body_esquema, 
            termina_con_barra, txt_con_esquema, txt_sin_esquema, num_con, num_sin,
            num_updates_sin_esq, num_inserts_sin_esq, num_deletes_sin_esq, es_un_package)


# --- INTERFAZ DE STREAMLIT ---
st.title("🛡️ Auditor SQL Avanzado - CALIDAD")
st.markdown("Tablero de Calidad DML.")

ruta = st.text_input("Ruta del directorio local a auditar:", placeholder="Ejemplo: C:/Proyectos/Scripts_SQL")

if st.button("🚀 Analizar Directorio"):
    if not ruta:
        st.warning("Por favor, ingresa una ruta válida.")
    elif os.path.exists(ruta):
        reportes = []
        archivos = [f for f in os.listdir(ruta) if f.lower().endswith(('.sql', '.pck', '.pkh', '.pkb'))]

        if len(archivos) == 0:
            st.info("No se encontraron archivos válidos en la ruta especificada.")
        else:
            progreso = st.progress(0)
            
            # Inicializadores globales de contadores
            g_total_packages = 0
            
            for i, archivo in enumerate(archivos):
                ruta_completa = os.path.join(ruta, archivo)
                try:
                    with open(ruta_completa, 'rb') as f:
                        contenido = f.read().decode('utf-8', errors='replace')
                        
                    (errs, ver, c_ver, c_pkg, c_body, c_barra, t_con, t_sin, num_con, num_sin,
                     n_upd, n_ins, n_del, es_un_pkg) = auditar_contenido_sql(contenido)
                    
                    if es_un_pkg:
                        g_total_packages += 1
                        val_upd = "No Aplica"
                        val_ins = "No Aplica"
                        val_del = "No Aplica"
                        t_sin_reporte = "No Aplica"
                    else:
                        val_upd = "✅" if n_upd == 0 else n_upd
                        val_ins = "✅" if n_ins == 0 else n_ins
                        val_del = "✅" if n_del == 0 else n_del
                        t_sin_reporte = t_sin
                        
                    reportes.append({
                        'Archivo': archivo,
                        'Estado': "❌ Errores" if errs else "✅ OK",
                        'Update sin esquema': val_upd,
                        'Insert sin esquema': val_ins,
                        'Delete sin esquema': val_del,
                        'Check Versión': c_ver,
                        'Check Package (Espec.)': c_pkg,
                        'Check Body Esquema': c_body,
                        'Check Termina "/"': c_barra,
                        'Versión Detectada': ver,
                        'Suma Tablas Con': num_con,
                        'Suma Tablas Sin': 0 if es_un_pkg else num_sin,
                        'Tablas Detectadas (Con)': t_con,
                        'Tablas Sin Esquema': t_sin_reporte,
                        'Detalle Línea a Línea': "\n\n".join(errs) if errs else "Ninguno"
                    })
                except Exception as e:
                    reportes.append({
                        'Archivo': archivo, 'Estado': "🚨 Fallo Crítico",
                        'Update sin esquema': "🚨", 'Insert sin esquema': "🚨", 'Delete sin esquema': "🚨",
                        'Check Versión': "🚨", 'Check Package (Espec.)': "🚨", 'Check Body Esquema': "🚨",
                        'Check Termina "/"': "🚨", 'Versión Detectada': "Error", 'Suma Tablas Con': 0, 'Suma Tablas Sin': 0,
                        'Tablas Detectadas (Con)': "Error", 'Tablas Sin Esquema': "Error", 'Detalle Línea a Línea': str(e)
                    })
                progreso.progress((i + 1) / len(archivos))
            
            df = pd.DataFrame(reportes)
            
            # --- CONTEOS DE MÉTRICAS NUEVAS ---
            total_archivos = len(df)
            total_packages = g_total_packages
            total_scripts = total_archivos - total_packages
            total_con_errores = len(df[df['Estado'].str.contains("❌|🚨")])
            total_ok = len(df[df['Estado'] == "✅ OK"])
            
            # --- PANEL DE MÉTRICAS CENTRALES ---
            st.subheader("Métricas de Control SQL")
            m1, m2, m3, m4, m5 = st.columns(5)
            with m1:
                st.metric("Total Archivos", total_archivos)
            with m2:
                st.metric("Total Package", total_packages)
            with m3:
                st.metric("Total Script", total_scripts)
            with m4:
                st.metric("Total Archivos con Errores", total_con_errores, delta="Por corregir" if total_con_errores > 0 else "Ninguno", delta_color="inverse" if total_con_errores > 0 else "normal")
            with m5:
                st.metric("Total Archivos OK", total_ok)
            
            # --- TABLA DE RESULTADOS ---
            st.subheader("Reporte de Calidad de Código")
            
            columnas_ordenadas = [
                'Archivo', 'Estado', 'Update sin esquema', 'Insert sin esquema', 'Delete sin esquema',
                'Check Versión', 'Check Package (Espec.)', 'Check Body Esquema', 'Check Termina "/"', 
                'Versión Detectada', 'Suma Tablas Con', 'Suma Tablas Sin', 'Tablas Sin Esquema', 'Detalle Línea a Línea'
            ]
            
            st.dataframe(df[columnas_ordenadas], use_container_width=True, hide_index=True)

            # Generar Excel en memoria
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df[columnas_ordenadas].to_excel(writer, index=False, sheet_name='Auditoria_Calidad')
            
            st.download_button(
                label="📥 Descargar Reporte Completo (Excel)",
                data=output.getvalue(),
                file_name="reporte_calidad_operaciones.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("La ruta especificada no existe.")