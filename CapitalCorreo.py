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
import locale
import time
import streamlit as st
import pythoncom
import signal

# --- Configuración Inicial del Dashboard de Streamlit ---
st.title("📬 Automatización de Correos Capital")
st.subheader("📊 Seguimiento del proceso en tiempo real")
st.markdown("---")

# Ajusta esta ruta a tu entorno real de OneDrive / Disco Duro
base_path = r"C:\Calidad\CorreosCapital"
fecha_ejecucion = datetime.now().strftime('%Y/%m/%d %H:%M:%S')  

if not os.path.exists(base_path):
    os.makedirs(base_path)

# 2. Mostrarlo en la interfaz de Streamlit
st.info(f"📂 **Los archivos se guardarán en:** `{base_path}`")    

try:
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8') 
except:
    try:
        locale.setlocale(locale.LC_TIME, 'spanish') 
    except:
        st.warning("No se pudo establecer el idioma español, se usará el predeterminado.")

# ===========================================
# Funciones Base del Script Original
# ===========================================
def estandarizar_nombre_columna(df):
    posibles_nombres = ['Cod. Requerimiento', 'Cod. Req', 'Requerimiento', 'Código Requerimiento']
    for nombre in posibles_nombres:
        if nombre in df.columns:
            df = df.rename(columns={nombre: 'Requerimiento'})
            return df
    return df

def limpiar_asunto(asunto):
    if not asunto:
        return ""
    asunto_limpio = " ".join(asunto.split())
    return asunto_limpio

def mover_seguro(msg, carpeta_destino, intentos=3):
    for i in range(intentos):
        try:
            msg.Move(carpeta_destino)
            return True
        except Exception as e:
            time.sleep(1.5)
    return False

# =========================================================
# CONTENEDOR EN VIVO DE STREAMLIT (PROCESO PASO A PASO)
# =========================================================
with st.status("Iniciando componentes de automatización Capital...", expanded=True) as status:
    # Definición de punteros para limpieza forzada en bloque finally
    outlook = None
    bandeja_entrada = None
    carpeta_raiz = None
    carpeta_destino = None

    try:
        # Inicializar el subproceso COM de Windows para evitar caídas de Streamlit
        pythoncom.CoInitialize()

        st.write("🔗 Conectando con la API de escritorio de Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        bandeja_entrada = outlook.GetDefaultFolder(6)  # Bandeja de Entrada (Inbox)
        raiz_cuenta = bandeja_entrada.Parent           # Raíz de la cuenta (tu_correo@sonda.com)
        
        # --- DETECCIÓN AUTOMÁTICA DE UBICACIÓN ---
        nombre_carpeta = "A-Capital_Script_Produccion"
        
        # Intento 1: Buscar en la raíz de la cuenta
        try:
            carpeta_raiz = raiz_cuenta.Folders[nombre_carpeta]
            st.write(f"📁 Carpeta `{nombre_carpeta}` detectada en la raíz principal de la cuenta.")
        except:
            # Intento 2: Si falla, buscar dentro de la Bandeja de Entrada (Inbox)
            try:
                carpeta_raiz = bandeja_entrada.Folders[nombre_carpeta]
                st.write(f"📁 Carpeta `{nombre_carpeta}` detectada dentro de la Bandeja de Entrada.")
            except Exception as e:
                raise Exception(f"No se encontró la carpeta '{nombre_carpeta}' en ninguna ubicación de Outlook. Verifica que no existan espacios invisibles al final del nombre.")

        # Una vez ubicada la raíz, asignamos de forma segura el repositorio de destino
        try:
            carpeta_destino = carpeta_raiz.Folders["Capital_Procesados"]
        except Exception as e:
            raise Exception(f"Se localizó '{nombre_carpeta}', pero falta crear la subcarpeta 'Capital_Procesados' en su interior desde Outlook.")

        # Notificación nativa de Windows al iniciar
        toast('Automatización CAPITAL', 'Proceso Iniciado. Recopilando Correos CATALOGACION PROD.')
        winsound.Beep(1000, 500)

        # --- 1. IDENTIFICAR MENSAJES (FILTRADO INICIAL) ---
        st.write("🔍 Escaneando ítems elegibles en la carpeta de origen...")
        items = carpeta_raiz.Items
        ids_a_procesar = []

        for msg in items:
            try:
                if hasattr(msg, 'Subject'):
                    asunto = limpiar_asunto(msg.Subject)
                    if "CAPITAL" in asunto and "EJECUTAR" in asunto and "PROD" in asunto and "SCRIPT" in asunto:
                        ids_a_procesar.append(msg.EntryID)
            except Exception as e:
                continue

        # Quitar duplicados por descalce de sincronización en servidores Exchange
        ids_unicos = list(set(ids_a_procesar))
        total_correos = len(ids_unicos)
        
        if total_correos == 0:
            st.write("⚠️ No se encontraron correos pendientes de procesamiento.")
            barra_progreso = st.progress(1.0)
        else:
            st.write(f"📂 Se detectaron `{total_correos}` correos listos para su procesamiento.")
            barra_progreso = st.progress(0.0)

        # --- 2. PROCESAR CADA CORREO (BUCLE PRINCIPAL CON INTERFAZ VIVA) ---
        for indice, entry_id in enumerate(ids_unicos, 1):
            msg = None  
            try:
                msg = outlook.GetItemFromID(entry_id)
                
                # ESCUDO CRÍTICO: Comprobación de existencia en la carpeta origen
                if not msg or msg.Parent.Name != nombre_carpeta:
                    if msg: del msg
                    continue

                remitente = getattr(msg, 'SenderEmailAddress', '').strip().lower()
                asunto = limpiar_asunto(msg.Subject)
                
                st.write(f"📨 **[{indice}/{total_correos}] Procesando:** {asunto[:50]}...")
                toast('Automatización CAPITAL', f"Procesando correo {indice}/{total_correos}...")

                # --- CONFIGURACIÓN DE RUTAS TEMPORALES ---
                fecha_correo = msg.SentOn.replace(tzinfo=None)
                start_week = fecha_correo - timedelta(days=fecha_correo.weekday())
                mes_carpeta = fecha_correo.strftime('%B').upper() 
                
                end_week = start_week + timedelta(days=6)
                rango_semana = f"{start_week.strftime('%d%m%Y')} {end_week.strftime('%d%m%Y')}"

                carpeta_mes = os.path.join(base_path, mes_carpeta)
                semana_path = os.path.join(carpeta_mes, f"SCRIPT_EN_PRODUCCION_{rango_semana}")
                os.makedirs(semana_path, exist_ok=True)
                
                path_sonda = os.path.join(semana_path, "SONDA")
                path_explotacion = os.path.join(semana_path, "EXPLOTACION")
                os.makedirs(path_sonda, exist_ok=True)
                os.makedirs(path_explotacion, exist_ok=True)

                nombre_archivo_final = os.path.join(semana_path, f"[CAPITAL]-EJECUTAR_SCRIPT_EN_PRODUCCION_{rango_semana}.xlsx")

                # --- CAMINO A: LÓGICA PARA CARPETA EXPLOTACION ---
                if "explotacion@sura.cl" in remitente:
                    safe_subject = re.sub(r'[\\/*?:"<>|]', "", msg.Subject)
                    msg_path = os.path.join(path_explotacion, f"{safe_subject}.msg")
                    
                    if not os.path.exists(msg_path):
                        msg.SaveAs(msg_path)

                # --- CAMINO B: LÓGICA PARA CARPETA SONDA / OTROS ---
                elif "calidad.afp@sonda.com" not in remitente:
                    safe_subject = re.sub(r'[\\/*?:"<>|]', "", msg.Subject)
                    msg_path = os.path.join(path_sonda, f"{safe_subject}.msg")
                    msg.SaveAs(msg_path)
                                
                    match_servidor = re.search(r"BD\s*[:\s]*(\w+)", msg.Body, re.IGNORECASE)
                    nombre_servidor = match_servidor.group(1) if match_servidor else "DESCONOCIDO"

                    for adj in msg.Attachments:
                        if adj.FileName.startswith("IM") and adj.FileName.endswith(".xlsx"):
                            temp_path = os.path.join(path_sonda, adj.FileName)
                            adj.SaveAsFile(temp_path)
                            
                            # Corrección: Uso explícito de openpyxl para mantener codificaciones íntegras
                            df = pd.read_excel(temp_path, engine='openpyxl')
                            df = df.dropna(how='all')
                            df.columns = df.columns.str.strip()
                            df = estandarizar_nombre_columna(df)
                            df = df.dropna(subset=['Requerimiento'])
                            
                            ultimo_n = 0
                            ultimo_rf = None
                            if os.path.exists(nombre_archivo_final):
                                df_base = pd.read_excel(nombre_archivo_final, engine='openpyxl')
                                if not df_base.empty:
                                    ultimo_n = df_base['N°'].iloc[-1]
                                    ultimo_rf = df_base['NUMERO TICKET'].iloc[-1]

                            cambios = df['Requerimiento'].ne(df['Requerimiento'].shift())
                            if ultimo_rf is not None and df['Requerimiento'].iloc[0] == ultimo_rf:
                                cambios.iloc[0] = False
                            df['N°'] = ultimo_n + cambios.cumsum()

                            match_motivo = re.search(r"IM-\d+\s*(.*)", asunto)
                            motivo = match_motivo.group(1) if match_motivo else "SIN MOTIVO"

                            df_nuevo = pd.DataFrame({
                                'N°': df['N°'],
                                'SERVIDOR': nombre_servidor,
                                'SCRIPT': df['Nombre Item'].iloc[0] if 'Nombre Item' in df.columns else "N/A",
                                'MOTIVO': motivo,
                                'FECHA EJECUCIÓN': df['Fecha Catalogación'].iloc[0] if 'Fecha Catalogación' in df.columns else "N/A",
                                'NUMERO TICKET': df['Requerimiento'],
                                'AUTOMATIZADO FEC': fecha_ejecucion
                            })
                                        
                            if os.path.exists(nombre_archivo_final):
                                df_base = pd.read_excel(nombre_archivo_final, engine='openpyxl')
                                df_final = pd.concat([df_base, df_nuevo], ignore_index=True)
                            else:
                                df_final = df_nuevo
                            
                            df_final.to_excel(nombre_archivo_final, index=False)
                            os.remove(temp_path)

                # --- FINALIZACIÓN DE ITERACIÓN: MOVER Y DESVINCULAR ELEMENTO ---
                msg.Save()
                if mover_seguro(msg, carpeta_destino, intentos=3):
                    pass
                else:
                    msg.UnRead = False
                    msg.Save()

                del msg
                time.sleep(1.0) # Sincronización del búfer MAPI de Exchange

                # ACTUALIZACIÓN EN TIEMPO REAL DE LA BARRA DE PROGRESO
                porcentaje_actual = indice / total_correos
                barra_progreso.progress(porcentaje_actual)

            except Exception as e:
                if msg: del msg
                continue

        # Mensaje de éxito al salir del bucle
        toast('Automatización CAPITAL', 'Proceso Finalizado. CATALOGACION PROD.')
        winsound.Beep(1000, 500)
        
        status.update(label="🎉 ¡Consolidación exitosa!", state="complete", expanded=False)
        st.success("✨ Todo el set de correos de Capital ha sido procesado y archivado.")

    except Exception as e:
        status.update(label="❌ Ocurrió un error en la ejecución", state="error", expanded=True)
        st.error(f"Falla crítica: {str(e)}")
        toast('Error en CAPITAL', f'El proceso falló: {str(e)}')
        with open("error_log.txt", "a", encoding="utf-8") as f:
            f.write(f"{datetime.now()}: {str(e)}\n")

    finally:
        # --- DESBLOQUEO SISTÉMICO DE RECURSOS COM ---
        st.write("⚙️ Limpiando hilos de ejecución de Windows...")
        carpeta_destino = None
        carpeta_raiz = None
        bandeja_entrada = None
        outlook = None
        pythoncom.CoUninitialize()