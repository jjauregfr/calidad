# -*- coding: utf-8 -*-
import win32com.client
import pandas as pd
import os
from datetime import datetime, timedelta
import re
from win11toast import toast
import winsound
import locale
import time
import streamlit as st
import pythoncom

# --- Configuración Inicial del Dashboard de Streamlit ---
st.title("馃摠 Automatizaci贸n de Correos Capital")
st.subheader("Seguimiento del proceso en tiempo real")
st.markdown("---")

# Ajusta esta ruta a tu entorno real de OneDrive / Disco Duro
base_path = r"C:\Users\jjauregfr\OneDrive - Sonda S.A\Calidad\CorreosCapital"
fecha_ejecucion = datetime.now().strftime('%Y/%m/%d %H:%M:%S')  

try:
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8') 
except:
    try:
        locale.setlocale(locale.LC_TIME, 'spanish') 
    except:
        st.warning("No se pudo establecer el idioma espa帽ol, se usar谩 el predeterminado.")

# ===========================================
# Funciones Base del Script Original
# ===========================================
def estandarizar_nombre_columna(df):
    posibles_nombres = ['Cod. Requerimiento', 'Cod. Req', 'Requerimiento','C贸digo Requerimiento']
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
# CONTENEDOR EN VIVO DE STREAMLIT (PROCESO PASO A PASEO)
# =========================================================
with st.status("Iniciando componentes de automatizaci贸n Capital...", expanded=True) as status:
    # Definici贸n de punteros para limpieza forzada en bloque finally
    outlook = None
    bandeja_entrada = None
    carpeta_raiz = None
    carpeta_destino = None

    try:
        # Inicializar el subproceso COM de Windows para evitar ca铆das de Streamlit
        pythoncom.CoInitialize()

        st.write("馃攲 Conectando con la API de escritorio de Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        bandeja_entrada = outlook.GetDefaultFolder(6)  # Bandeja de Entrada (Inbox)
        raiz_cuenta = bandeja_entrada.Parent           # Ra铆z de la cuenta (tu_correo@sonda.com)
        
        # --- DETECCI脫N AUTOM脕TICA DE UBICACI脫N ---
        nombre_carpeta = "A-Capital_Script_Produccion"
        
        # Intento 1: Buscar en la ra铆z de la cuenta
        try:
            carpeta_raiz = raiz_cuenta.Folders[nombre_carpeta]
            st.write(f"馃搨 Carpeta `{nombre_carpeta}` detectada en la ra铆z principal de la cuenta.")
        except:
            # Intento 2: Si falla, buscar dentro de la Bandeja de Entrada (Inbox)
            try:
                carpeta_raiz = bandeja_entrada.Folders[nombre_carpeta]
                st.write(f"馃搨 Carpeta `{nombre_carpeta}` detectada dentro de la Bandeja de Entrada.")
            except Exception as e:
                raise Exception(f"No se encontr贸 la carpeta '{nombre_carpeta}' en ninguna ubicaci贸n de Outlook. Verifica que no existan espacios invisibles al final del nombre.")

        # Una vez ubicada la ra铆z, asignamos de forma segura el repositorio de destino
        try:
            carpeta_destino = carpeta_raiz.Folders["Capital_Procesados"]
        except Exception as e:
            raise Exception(f"Se localiz贸 '{nombre_carpeta}', pero falta crear la subcarpeta 'Capital_Procesados' en su interior desde Outlook.")

        # Notificaci贸n nativa de Windows al iniciar
        toast('Automatizaci贸n CAPITAL', 'Proceso Iniciado. Recopilando Correos CATALOGACION PROD.')
        winsound.Beep(1000, 500)

        # --- 1. IDENTIFICAR MENSAJES (FILTRADO INICIAL) ---
        st.write("馃摜 Escaneando 铆tems elegibles en la carpeta de origen...")
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

        # Quitar duplicados por descalce de sincronizaci贸n en servidores Exchange
        ids_unicos = list(set(ids_a_procesar))
        total_correos = len(ids_unicos)
        
        if total_correos == 0:
            st.write("鈿狅笍 No se encontraron correos pendientes de procesamiento.")
            barra_progreso = st.progress(1.0)
        else:
            st.write(f"馃搳 Se detectaron `{total_correos}` correos listos para su procesamiento.")
            barra_progreso = st.progress(0.0)

        # --- 2. PROCESAR CADA CORREO (BUCLE PRINCIPAL CON INTERFAZ VIVA) ---
        for indice, entry_id in enumerate(ids_unicos, 1):
            msg = None  
            try:
                msg = outlook.GetItemFromID(entry_id)
                
                # ESCUDO CR脥TICO: Comprobaci贸n de existencia en la carpeta origen
                if not msg or msg.Parent.Name != nombre_carpeta:
                    if msg: del msg
                    continue

                remitente = getattr(msg, 'SenderEmailAddress', '').strip().lower()
                asunto = limpiar_asunto(msg.Subject)
                
                st.write(f"馃摟 **[{indice}/{total_correos}] Procesando:** {asunto[:50]}...")
                toast('Automatizaci贸n CAPITAL', f"Procesando correo {indice}/{total_correos}...")

                # --- CONSTRUCCI脫N DE RUTAS TEMPORALES ---
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

                # --- CAMINO A: L脫GICA PARA CARPETA EXPLOTACION ---
                if "explotacion@sura.cl" in remitente:
                    safe_subject = re.sub(r'[\\/*?:"<>|]', "", msg.Subject)
                    msg_path = os.path.join(path_explotacion, f"{safe_subject}.msg")
                    
                    if not os.path.exists(msg_path):
                        msg.SaveAs(msg_path)

                # --- CAMINO B: L脫GICA PARA CARPETA SONDA / OTROS ---
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
                            
                            df = pd.read_excel(temp_path)
                            df = df.dropna(how='all')
                            df.columns = df.columns.str.strip()
                            df = estandarizar_nombre_columna(df)
                            df = df.dropna(subset=['Requerimiento'])
                            
                            ultimo_n = 0
                            ultimo_rf = None
                            if os.path.exists(nombre_archivo_final):
                                df_base = pd.read_excel(nombre_archivo_final)
                                if not df_base.empty:
                                    ultimo_n = df_base['N掳'].iloc[-1]
                                    ultimo_rf = df_base['NUMERO TICKET'].iloc[-1]

                            cambios = df['Requerimiento'].ne(df['Requerimiento'].shift())
                            if ultimo_rf is not None and df['Requerimiento'].iloc[0] == ultimo_rf:
                                cambios.iloc[0] = False
                            df['N掳'] = ultimo_n + cambios.cumsum()

                            match_motivo = re.search(r"IM-\d+\s*(.*)", asunto)
                            motivo = match_motivo.group(1) if match_motivo else "SIN MOTIVO"

                            df_nuevo = pd.DataFrame({
                                'N掳': df['N掳'],
                                'SERVIDOR': nombre_servidor,
                                'SCRIPT': df['Nombre Item'].iloc[0] if 'Nombre Item' in df.columns else "N/A",
                                'MOTIVO': motivo,
                                'FECHA EJECUCI脫N': df['Fecha Catalogaci贸n'].iloc[0] if 'Fecha Catalogaci贸n' in df.columns else "N/A",
                                'NUMERO TICKET': df['Requerimiento'],
                                'AUTOMATIZADO FEC': fecha_ejecucion
                            })
                                        
                            if os.path.exists(nombre_archivo_final):
                                df_base = pd.read_excel(nombre_archivo_final)
                                df_final = pd.concat([df_base, df_nuevo], ignore_index=True)
                            else:
                                df_final = df_nuevo
                            
                            df_final.to_excel(nombre_archivo_final, index=False)
                            os.remove(temp_path)

                # --- FINALIZACI脫N DE ITERACI脫N: MOVER Y DESVINCULAR ELEMENTO ---
                msg.Save()
                if mover_seguro(msg, carpeta_destino, intentos=3):
                    pass
                else:
                    msg.UnRead = False
                    msg.Save()

                del msg
                time.sleep(1.0) # Sincronizaci贸n del b煤fer MAPI de Exchange

                # 馃搳 ACTUALIZACI脫N EN TIEMPO REAL DE LA BARRA DE PROGRESO
                porcentaje_actual = indice / total_correos
                barra_progreso.progress(porcentaje_actual)

            except Exception as e:
                if msg: del msg
                continue

        # Mensaje de 茅xito al salir del bucle
        toast('Automatizaci贸n CAPITAL', 'Proceso Finalizado. CATALOGACION PROD.')
        winsound.Beep(1000, 500)
        
        status.update(label="馃帀 隆Consolidaci贸n exitosa!", state="complete", expanded=False)
        st.success("鉁?Todo el set de correos de Capital ha sido procesado y archivado.")

    except Exception as e:
        status.update(label="鉂?Ocurri贸 un error en la ejecuci贸n", state="error", expanded=True)
        st.error(f"Falla cr铆tica: {str(e)}")
        toast('Error en CAPITAL', f'El proceso fall贸: {str(e)}')
        with open("error_log.txt", "a") as f:
            f.write(f"{datetime.now()}: {str(e)}\n")

    finally:
        # --- DESBLOQUEO SIST脡MICO DE RECURSOS COM ---
        st.write("鈾伙笍 Limpiando hilos de ejecuci贸n de Windows...")
        carpeta_destino = None
        carpeta_raiz = None
        bandeja_entrada = None
        outlook = None
        pythoncom.CoUninitialize()

# --- Bot贸n inferior para regresar al panel principal ---
if st.button("猬咃笍 Volver al Hub Principal", use_container_width=True):
    st.switch_page("panel.py")