# -*- coding: utf-8 -*-

# Limpieza absoluta de procesos colgados al arrancar el script
import os
import sys
import psutil

def limpiar_solo_outlook_congelado():
    # NO tocamos Python para que Streamlit no se caiga. 
    # Solo fulminamos Outlook si quedó bloqueado en segundo plano.
    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'].lower() == 'outlook.exe':
                proc.kill()
        except:
            pass

# Ejecuta la limpieza segura antes de que Streamlit intente conectar el nuevo COM
limpiar_solo_outlook_congelado()

import win32com.client
import pandas as pd
from datetime import datetime, timedelta
import re
from win11toast import toast
import winsound
import streamlit as st
import pythoncom
import time

# --- Configuración Inicial de la Página ---
st.title("📦 Automatización de Correos Planvital")
st.subheader("Seguimiento del proceso en tiempo real")
st.markdown("---")

# --- Configuración de Entorno ---
base_path = r"C:\Calidad\CorreosPlanvital"
nombre_carpeta_analista = "A-PlanVitalAnalista"

if not os.path.exists(base_path):
    os.makedirs(base_path)
# 2. Mostrarlo en la interfaz de Streamlit
st.info(f"📂 **Los archivos se guardarán en:** `{base_path}`")

def estandarizar_nombre_columna(df):
    posibles_nombres = ['Cod. Requerimiento', 'Cod. Req', 'Requerimiento', 'Código Requerimiento']
    for nombre in posibles_nombres:
        if nombre in df.columns:
            df = df.rename(columns={nombre: 'Requerimiento'})
            return df
    return df

def obtener_o_crear_carpeta(parent_folder, folder_name):
    try:
        return parent_folder.Folders[folder_name]
    except:
        try:
            return parent_folder.Folders.Add(folder_name)
        except Exception as e:
            raise Exception(f"No pude acceder ni crear la carpeta '{folder_name}': {str(e)}")

def limpiar_asunto(asunto):
    if not asunto: return ""
    asunto_limpio = re.sub(r'\s+', ' ', asunto)
    return asunto_limpio.strip().upper()
    
# =========================================================
# CONTENEDOR DINÁMICO EN MOVIMIENTO (PASO A PASO)
# =========================================================
with st.status("Ejecutando Automatización PlanVital...", expanded=True) as status:
    # Declaramos variables COM globales para asegurar su liberación en el bloque 'finally'
    outlook = None
    inbox = None
    folder_analista = None
    carpeta_destino = None
    
    try:
        # 1. Inicializar subproceso COM de Windows para Streamlit
        pythoncom.CoInitialize()
        
        st.write("🔗 Conectando de forma segura con la API de Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)

        # Cargar mapeo de carpetas de la cuenta predeterminada
        folder_analista = obtener_o_crear_carpeta(inbox, nombre_carpeta_analista)
        carpeta_destino = obtener_o_crear_carpeta(folder_analista, "PLANVITAL_Procesados")

        # --- NOTIFICACIÓN NATIVA DE WINDOWS ---
        toast('Automatización PlanVital', 'Proceso Iniciado. Recopilando Correos CATALOGACION TEST.')
        winsound.Beep(1000, 500)
        
        # --- PASO 1: Mover correos del Inbox a la carpeta de trabajo ---
        st.write("🔍 Analizando la Bandeja de Entrada (`Inbox`)...")
        mensajes_inbox = inbox.Items
        correos_clasificados = 0
        
        # Recorrido inverso para evitar desfase de índices en la API COM al desplazar elementos
        for i in range(len(mensajes_inbox), 0, -1):
            msg = mensajes_inbox[i-1]
            if hasattr(msg, 'Subject'):
                asunto_limpio = limpiar_asunto(msg.Subject)
                if "[PLANVITAL]" in asunto_limpio and "CATALOGA" in asunto_limpio and "TEST" in asunto_limpio:
                    msg.Move(folder_analista)
                    correos_clasificados += 1
                    
        if correos_clasificados > 0:
            st.write(f"📁 Se movieron `{correos_clasificados}` correos nuevos a la bandeja de análisis.")
        time.sleep(0.5)

        # --- PASO 2: Filtrar y contar los mensajes elegibles ---
        messages = folder_analista.Items
        messages.Sort("[ReceivedTime]", True)
        lista_mensajes = [msg for msg in messages]
        
        mensajes_a_procesar = []
        for msg in lista_mensajes:
            if hasattr(msg, 'Subject'):
                asunto_actual_limpio = limpiar_asunto(msg.Subject)
                if "[PLANVITAL]" in asunto_actual_limpio and "CATALOGA" in asunto_actual_limpio and "TEST" in asunto_actual_limpio:
                    mensajes_a_procesar.append(msg)

        total_correos = len(mensajes_a_procesar)
        
        if total_correos == 0:
            st.write("⚠️ No se encontraron correos pendientes por procesar.")
            barra_progreso = st.progress(1.0)
        else:
            st.write(f"📂 Detectados `{total_correos}` correos válidos para consolidación.")
            barra_progreso = st.progress(0.0)

        # --- PASO 3: Extracción de datos y Consolidación de DataFrames ---
        ids_correos_a_mover = []

        for indice, msg in enumerate(mensajes_a_procesar):
            asunto_actual_limpio = limpiar_asunto(msg.Subject)
            fecha_llegada_str = msg.ReceivedTime.strftime('%Y/%m/%d')
            
            st.write(f"📨 **[{indice + 1}/{total_correos}] Procesando:** {msg.Subject[:40]}...")
            
            # Cálculo del rango semanal basado en la fecha de envío
            fecha_correo = msg.SentOn.replace(tzinfo=None)
            start_week = fecha_correo - timedelta(days=fecha_correo.weekday())
            end_week = start_week + timedelta(days=6)
            rango_semana = f"{start_week.strftime('%d%m%Y')} {end_week.strftime('%d%m%Y')}"
            
            semana_path = os.path.join(base_path, f"PLV TEST {rango_semana}")
            os.makedirs(semana_path, exist_ok=True)
            nombre_archivo_final = os.path.join(semana_path, f"PLV TEST {rango_semana}.xlsx")
            
            # ruta final:
            st.info(f"📂 Guardando datos en: `{nombre_archivo_final}`")
            
            # Respaldar correo electrónico original en el directorio semanal
            safe_subject = re.sub(r'[\\/*?:"<>|]', "", msg.Subject)
            msg_path = os.path.join(semana_path, f"{safe_subject}.msg")
            msg.SaveAs(msg_path)
            toast('Automatización PlanVital', 'Procesando Archivos Adjuntos CATALOGACION TEST.')
                
            # Procesamiento iterativo de libros Excel adjuntos
            for attachment in msg.Attachments:
                if attachment.FileName.startswith("RF") and attachment.FileName.endswith(".xlsx"):
                    temp_path = os.path.join(semana_path, attachment.FileName)
                    attachment.SaveAsFile(temp_path)
                    
                    # CORRECCIÓN: Forzamos el motor openpyxl para evitar corrupción de codificación
                    df_temp = pd.read_excel(temp_path, engine='openpyxl')
                    df_temp = df_temp.dropna(how='all') 
                    df_temp.columns = df_temp.columns.str.strip()
                    df_temp = estandarizar_nombre_columna(df_temp)
                    df_temp = df_temp.dropna(subset=['Requerimiento'])
                        
                    mensaje_toast = f"Analista: {msg.SenderName}\nAsunto: {msg.Subject[:30]}..." 
                    toast('Procesando Archivos', mensaje_toast)
                    
                    # Lógica matemática de incremento secuencial de ítems
                    ultimo_n = 0
                    ultimo_rf = None
                    
                    if os.path.exists(nombre_archivo_final):
                        df_base = pd.read_excel(nombre_archivo_final, engine='openpyxl')
                        if not df_base.empty:
                            ultimo_n = df_base['N°'].iloc[-1]
                            ultimo_rf = df_base['RF'].iloc[-1]

                    cambios = df_temp['Requerimiento'].ne(df_temp['Requerimiento'].shift())
                    if ultimo_rf is not None and df_temp['Requerimiento'].iloc[0] == ultimo_rf:
                        cambios.iloc[0] = False
                    df_temp['N°'] = ultimo_n + cambios.cumsum()

                    # Estructuración final de la matriz de datos comprimida
                    df_nuevo = pd.DataFrame({
                        'N°': df_temp['N°'],
                        'ANALISTA': msg.SenderName,
                        'RF': df_temp['Requerimiento'],
                        'REFERENCIA': asunto_actual_limpio.split("TEST")[-1].strip(),
                        'PIEZA': df_temp['Nombre Item'],
                        'TIPO': df_temp['Tipo Item'],
                        'FECHA': fecha_llegada_str,
                        'VERSION': df_temp['Versión Item']
                    })
                    
                    if os.path.exists(nombre_archivo_final):
                        df_base = pd.read_excel(nombre_archivo_final, engine='openpyxl')
                        df_final = pd.concat([df_base, df_nuevo], ignore_index=True)
                    else:
                        df_final = df_nuevo
                    
                    df_final.to_excel(nombre_archivo_final, index=False)
                    os.remove(temp_path)

            # Guardamos el hash único de identificación del mensaje para moverlo al final de forma segura
            ids_correos_a_mover.append(msg.EntryID)
            
            # Actualización de la barra progresiva en la UI
            porcentaje = (indice + 1) / total_correos
            barra_progreso.progress(porcentaje)

        # --- PASO 4: Reubicación de Correos por EntryID Desacoplado ---
        if ids_correos_a_mover:
            st.write(f"📥 Trasladando {len(ids_correos_a_mover)} elementos procesados a la subcarpeta de destino...")
            for entry_id in ids_correos_a_mover:
                # Recuperación aislada del puntero del correo para ejecutar el cambio de repositorio limpio
                correo_individual = outlook.GetItemFromID(entry_id)
                correo_individual.Move(carpeta_destino)
            st.write("✅ La bandeja interna de Outlook fue reestructurada correctamente.")

        # --- NOTIFICACIÓN FINAL DE ÉXITO ---
        toast('Automatización PlanVital', 'Proceso finalizado con éxito. Los archivos fueron actualizados.')
        winsound.Beep(1000, 500)
        
        status.update(label="🎉 ¡Proceso completado!", state="complete", expanded=False)
        st.success("✅ La consolidación masiva se ejecutó de forma correcta y los históricos se encuentran resguardados.")

    except Exception as e:
        status.update(label="❌ El proceso falló", state="error", expanded=True)
        st.error(f"Error crítico en tiempo de ejecución: {str(e)}")
        toast('Error en PlanVital', f'El proceso falló: {str(e)}')
        
        # CORRECCIÓN: Guardar el log forzando UTF-8
        with open("error_log.txt", "a", encoding="utf-8") as f:
            f.write(f"{datetime.now()}: {str(e)}\n")

    finally:
        # --- DESBLOQUEO CRÍTICO DE PUNTEROS Y CONEXIÓN DE WINDOWS ---
        st.write("⚙️ Liberando hilos del núcleo del sistema operativo...")
        inbox = None
        folder_analista = None
        carpeta_destino = None
        outlook = None
        pythoncom.CoUninitialize()