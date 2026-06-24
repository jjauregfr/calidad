import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Configuración de la página web
st.set_page_config(page_title="Unificador de Excel", page_icon="📊", layout="centered")

st.title("📊 Unificador de Archivos Excel")
st.write("Coloca tus archivos `.xlsx` en la carpeta o ejecuta la unificación directa.")

# 1. Configura la ruta de la carpeta
carpeta_origen = st.text_input("Ruta de la carpeta contenedora:", value=".")

if st.button("🚀 Iniciar Unificación", type="primary"):
    if os.path.exists(carpeta_origen):
        # 2. Buscar archivos
        archivos = [f for f in os.listdir(carpeta_origen) if f.endswith('.xlsx') and not f.startswith('~$') and not f.startswith('Resultado_Unificado')]
        
        if len(archivos) == 0:
            st.warning("⚠️ No se encontraron archivos Excel (.xlsx) en la carpeta especificada.")
        else:
            st.write(f"**Archivos detectados ({len(archivos)}):**")
            lista_dataframes = []
            
            # Progreso en Streamlit
            progreso = st.progress(0)
            
            for idx, archivo in enumerate(archivos):
                ruta_completa = os.path.join(carpeta_origen, archivo)
                try:
                    df = pd.read_excel(ruta_completa)
                    df['Origen_Archivo'] = archivo  # Columna de trazabilidad
                    lista_dataframes.append(df)
                    st.success(f"✔ Cargado: {archivo} ({len(df)} filas)")
                except Exception as e:
                    st.error(f"❌ Error en {archivo}: {e}")
                
                # Actualizar barra de progreso
                progreso.progress((idx + 1) / len(archivos))
            
            # 3. Concatenar y guardar
            if lista_dataframes:
                df_unificado = pd.concat(lista_dataframes, ignore_index=True)
                
                fecha_hoy = datetime.now().strftime('%Y%m%d_%H%M%S')
                nombre_salida = f"Resultado_Unificado_{fecha_hoy}.xlsx"
                archivo_salida = os.path.join(carpeta_origen, nombre_salida)
                
                df_unificado.to_excel(archivo_salida, index=False)
                
                st.balloons() # Animación de éxito 🎉
                st.success(f"✨ **¡Proceso completado con éxito!** ✨")
                st.info(f"**Guardado en:** {archivo_salida}")
                st.metric(label="Total Filas Consolidadas", value=len(df_unificado))
    else:
        st.error("❌ La ruta de la carpeta no existe de forma local.")