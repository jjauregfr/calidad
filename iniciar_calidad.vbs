Set WshShell = CreateObject("WScript.Shell")
ruta = "C:\Users\jjauregfr\OneDrive - Sonda S.A\Calidad\Gestion_Calidad"

' 1. Cierra Outlook si quedó colgado (Evita el congelamiento en el puerto 8507)
' Eliminamos 'python.exe' de aquí para que no rompa la conexión de Streamlit al recargar
WshShell.Run "taskkill /f /im outlook.exe", 0, True

' Cambiar al directorio de trabajo
WshShell.CurrentDirectory = ruta

' 2. Ejecución oculta forzando UTF-8 (chcp 65001) para que NUNCA se rompan los emojis
WshShell.Run "cmd /c chcp 65001 && python -m streamlit run panel_calidad.py --server.headless true --server.port 8507", 0

' Esperar 4 segundos a que el servidor de Streamlit levante
WScript.Sleep 4000

' Abrir el navegador directamente en el puerto 8507
WshShell.Run "http://localhost:8507"