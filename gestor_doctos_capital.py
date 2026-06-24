# -*- coding: utf-8 -*-
import streamlit as st
import subprocess
import sys
import os
import pandas as pd
import re
from pathlib import Path
from datetime import datetime


# Configuración de página
st.set_page_config(page_title="Gestor de Calidad Sonda", page_icon="✅")

st.title("Automatización de Calidad AFP CAPITAL")
st.markdown("---")

# --- CONFIGURACIÓN ---
RF_FOLDER = r"C:\Users\jjauregfr\OneDrive - Sonda S.A\Calidad\Gestion_Calidad\Generacion_Archivos_Capital\RF"
SCRIPTS = {
    "FDA": r"C:\Users\jjauregfr\OneDrive - Sonda S.A\Calidad\Gestion_Calidad\pages\PY\FDA1.py",
    "Catalogacion": r"C:\Users\jjauregfr\OneDrive - Sonda S.A\Calidad\Gestion_Calidad\pages\PY\Catalogacion.py",
    "Manual": r"C:\Users\jjauregfr\OneDrive - Sonda S.A\Calidad\Gestion_Calidad\pages\PY\Manual.py",
    "FDB": r"C:\Users\jjauregfr\OneDrive - Sonda S.A\Calidad\Gestion_Calidad\pages\PY\FDB.py"
}

COLUMNAS_VALIDAS = [
    "Requerimiento", "Nombre Item", "Tipo Item", "Módulo", 
    "Motivo Implementación", "Fecha Catalogación", "Versión Item", 
    "Ruta Repositorio", "Nombre Rama", "Observaciones"
]



def estandarizar_nombre_columna(df):
    # 1. Borrar filas que estén completamente vacías (líneas en blanco)
    df = df.dropna(how='all')
    # Eliminar espacios fantasmas en los nombres de las columnas antes de mapear
    df.columns = df.columns.str.strip()
    posibles_nombres = ['Cod. Requerimiento', 'Cod. Req', 'Requerimiento', 'Código Requerimiento']
    for nombre in posibles_nombres:
        if nombre in df.columns:
            df = df.rename(columns={nombre: 'Requerimiento'})
            break
    return df

def registrar_en_historico(rf_id, total_elementos, fecha, hora_inicio, hora_termino):
    archivo_log = r"C:\Users\jjauregfr\OneDrive - Sonda S.A\Calidad\RF-Procesados\Rf_procesados.xlsx"
    # Crear directorio si no existe
    os.makedirs(os.path.dirname(archivo_log), exist_ok=True)
    
    nuevo_dato = {
        "Requerimiento": [rf_id],
        "Total Elementos": [total_elementos],
        "Fecha": [fecha],
        "Hora Inicio": [hora_inicio],
        "Hora Termino": [hora_termino]
    }
    df_nuevo = pd.DataFrame(nuevo_dato)
    
    if os.path.exists(archivo_log):
        df_existente = pd.read_excel(archivo_log)
        df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
    else:
        df_final = df_nuevo
        
    df_final.to_excel(archivo_log, index=False)
    
# --- FUNCIONES ---
def get_latest_rf():
    files = [os.path.join(RF_FOLDER, f) for f in os.listdir(RF_FOLDER) 
             if f.upper().startswith("RF") and f.lower().endswith(".xlsx")]
    return max(files, key=os.path.getmtime) if files else None

def validar_encabezados(file_path):
    try:
        # Leemos las primeras filas para validar la estructura interna
        df = pd.read_excel(file_path, nrows=5)
        # Aplicamos la estandarización para homologar "Cod. Requerimiento" a "Requerimiento"
        df = estandarizar_nombre_columna(df)
        
        columnas_actuales = df.columns.tolist()
        faltantes = [col for col in COLUMNAS_VALIDAS if col not in columnas_actuales]
        
        if faltantes:
            return False, f"Columnas faltantes: {', '.join(faltantes)}"
        return True, None
    except Exception as e:
        return False, f"No se pudo leer el archivo: {str(e)}"

#==================================
# --- INTERFAZ ---
#======================================
if st.button("Detectar último archivo y Procesar"):
    rf_file = get_latest_rf()
    hora_inicio = datetime.now().strftime("%H:%M:%S")
    fecha_ejec = datetime.now().strftime("%Y-%m-%d")
    
    if not rf_file:
        st.error("No se encontraron archivos RF en la carpeta.")
    else:
        valido, error = validar_encabezados(rf_file)
        if not valido:
            st.error(f"Error de encabezado: {error}")
        else:
        # 1. Leer el archivo, estandarizar columnas y contar elementos reales
            df_rf = pd.read_excel(rf_file)
            df_rf = estandarizar_nombre_columna(df_rf)
            total_elementos = len(df_rf)            
            
          
            
            rf_name = Path(rf_file).stem
            match = re.search(r'RF-\d+', rf_name)
            rf_id = match.group(0) if match else "RF_Desconocido"
            target_folder = os.path.join(os.path.dirname(rf_file), rf_id)
            
            if not os.path.exists(target_folder):
                os.makedirs(target_folder)
            
            st.success(f"Archivo validado y estandarizado correctamente: **{os.path.basename(rf_file)}**")
            st.info(f"Carpeta de trabajo: {target_folder}")
            
            # Ejecutar Scripts
            for nombre, path in SCRIPTS.items():
                with st.status(f"Ejecutando {nombre}...", expanded=True) as status:
                    try:
                        # Usamos encoding 'cp1252' (Windows estándar) y errors='replace' para no fallar
                        result = subprocess.run(
                            [sys.executable, path, rf_file], 
                            capture_output=True, 
                            encoding='cp1252', 
                            errors='replace'
                        )
                        
                        if result.returncode == 0:
                            st.write(f"{nombre} finalizado.")
                            status.update(label=f"{nombre} completado", state="complete")
                            
                            if nombre=="FDA":
                                    st.success("🎉 **FDA generado correctamente**")        
                                    # Buscamos la línea que tiene el OUTPUT en la consola procesada
                                    st.info(f"📂 **Carpeta de trabajo:** `{RF_FOLDER}`")
                            if nombre=="Catalogacion":
                                    st.success("🎉 **Catalogación generado correctamente**")        
                                    # Buscamos la línea que tiene el OUTPUT en la consola procesada
                                    st.info(f"📂 **Carpeta de trabajo:** `{RF_FOLDER}`")
                            if nombre=="Manual":
                                    st.success("🎉 **READM generado correctamente**")        
                                    # Buscamos la línea que tiene el OUTPUT en la consola procesada
                                    st.info(f"📂 **Carpeta de trabajo:** `{RF_FOLDER}`")
                            if nombre=="FDB":
                                    st.success("🎉 **FBD generado correctamente**")        
                                    # Buscamos la línea que tiene el OUTPUT en la consola procesada
                                    st.info(f"📂 **Carpeta de trabajo:** `{RF_FOLDER}`")       
                                            
                        else:
                            if "no aplica" in result.stderr.lower() or "no aplica" in result.stdout.lower():
                                st.warning(f"{nombre}: No aplica. Se omite.")
                                status.update(label=f"{nombre} omitido", state="complete")
                            else:
                                st.error(f"Error en {nombre}")
                                st.code(result.stderr)
                                status.update(label=f"Error en {nombre}", state="error")
                                break 
                    except Exception as e:
                        st.error(f"Excepción: {str(e)}")
                        break
            else:
                hora_termino = datetime.now().strftime("%H:%M:%S")
                registrar_en_historico(rf_id, total_elementos, fecha_ejec, hora_inicio, hora_termino)
                st.balloons()         
                st.success("¡Proceso finalizado y registrado exitosamente!")

st.sidebar.info("Cualquier duda, Arquitectura TI")