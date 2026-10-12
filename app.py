#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Servidor y Backend de la Aplicación de Reembolsos.
Permite seleccionar carpetas de forma interactiva, previsualizar los archivos de cada trabajador,
monitorear en tiempo real el proceso y descargar/abrir los archivos Excel generados.

El código original de procesar_reembolsos.py se mantiene 100% intacto y es importado.
"""

import os
import sys
import json
import time
import urllib.parse
import threading
import subprocess
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
from socketserver import ThreadingMixIn

import procesar_reembolsos as core

# Soporte para ejecución estándar y empaquetado con PyInstaller (.exe)
if getattr(sys, 'frozen', False):
    BASE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATES_DIR = os.path.join(BASE_DIR, 'templates')
if not os.path.exists(TEMPLATES_DIR):
    exe_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
    alt_templates = os.path.join(exe_dir, 'templates')
    if os.path.exists(alt_templates):
        TEMPLATES_DIR = alt_templates

# Estado global del proceso
app_state = {
    "status": "idle",       # "idle", "running", "completed", "error"
    "current_folder": "",
    "progress_msg": "Listo",
    "logs": [],
    "results": None,
    "error": None,
    "start_time": None,
    "duration": 0
}
state_lock = threading.Lock()


class StreamCapture:
    """Captura la salida estándar de Python y la envía al log en vivo de la aplicación."""
    def __init__(self, callback):
        self.callback = callback
        self.stdout = sys.stdout

    def write(self, text):
        self.stdout.write(text)
        if text:
            self.callback(text)

    def flush(self):
        self.stdout.flush()


def append_log(text):
    with state_lock:
        app_state["logs"].append({
            "time": time.strftime("%H:%M:%S"),
            "text": text
        })
        # Actualizar mensaje de progreso según la línea procesada
        clean = text.strip()
        if clean.startswith("Trabajador:"):
            worker = clean.split("Trabajador:")[-1].strip()
            app_state["progress_msg"] = f"Procesando trabajador: {worker}"
        elif "PROCESO COMPLETADO" in clean:
            app_state["progress_msg"] = "¡Proceso completado con éxito!"


def parse_execution_results(folder_path, raw_logs):
    """
    Analiza la carpeta procesada y los logs para construir un resumen estructurado
    de los trabajadores, comprobantes y montos resultantes.
    """
    trabajadores = []
    total_monto_global = 0.0
    total_comprobantes_global = 0

    subdirs = [d for d in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, d))]
    carpetas_trabajadores = [os.path.join(folder_path, d) for d in sorted(subdirs) if d.lower() != 'zextra']
    if not carpetas_trabajadores:
        carpetas_trabajadores = [folder_path]

    # Parsear logs para extraer detalles de items por trabajador
    trabajador_actual = None
    info_por_trabajador = {}

    for line_obj in raw_logs:
        line = line_obj["text"].strip()
        if line.startswith("Trabajador:"):
            trabajador_actual = line.split("Trabajador:")[-1].strip()
            if trabajador_actual not in info_por_trabajador:
                info_por_trabajador[trabajador_actual] = {
                    "excel": "",
                    "total": 0.0,
                    "items": []
                }
        elif trabajador_actual and "-> Archivo Excel:" in line:
            info_por_trabajador[trabajador_actual]["excel"] = line.split("-> Archivo Excel:")[-1].strip()
        elif trabajador_actual and "-> Total reembolsos:" in line:
            m_str = line.split("-> Total reembolsos:")[-1].replace("$", "").replace(",", "").strip()
            try:
                info_por_trabajador[trabajador_actual]["total"] = float(m_str)
            except ValueError:
                pass
        elif trabajador_actual and line.startswith("•"):
            # Ej: • [XML]      2026-08-16 | $  745.33 | SERVICIOS GASOLINEROS DE MEXIC | CLI 63308140
            partes = [p.strip() for p in line[1:].split("|")]
            origen_fecha = partes[0] if len(partes) > 0 else ""
            monto_str = partes[1] if len(partes) > 1 else "$0.00"
            proveedor = partes[2] if len(partes) > 2 else ""
            folio = partes[3] if len(partes) > 3 else ""

            origen = ""
            fecha = ""
            if "[" in origen_fecha and "]" in origen_fecha:
                origen = origen_fecha[origen_fecha.find("[")+1:origen_fecha.find("]")]
                fecha = origen_fecha[origen_fecha.find("]")+1:].strip()
            else:
                origen = "OTRO"
                fecha = origen_fecha

            monto_val = 0.0
            try:
                monto_val = float(monto_str.replace("$", "").replace(",", "").strip())
            except ValueError:
                pass

            info_por_trabajador[trabajador_actual]["items"].append({
                "origen": origen,
                "fecha": fecha,
                "monto": monto_val,
                "proveedor": proveedor,
                "folio": folio
            })

    # Consolidar información con las carpetas reales
    for c_path in carpetas_trabajadores:
        nombre = os.path.basename(c_path)
        archivos = os.listdir(c_path) if os.path.exists(c_path) else []
        excels = [f for f in archivos if f.lower().endswith('.xlsx') and not f.startswith('~$')]
        excel_name = excels[0] if excels else f"Reembolsos {nombre} WXX 2026XXXX.xlsx"
        excel_path = os.path.join(c_path, excel_name)

        datos = info_por_trabajador.get(nombre, {"excel": excel_name, "total": 0.0, "items": []})
        total_monto_global += datos["total"]
        total_comprobantes_global += len(datos["items"])

        trabajadores.append({
            "nombre": nombre,
            "carpeta_path": c_path,
            "excel_nombre": datos.get("excel") or excel_name,
            "excel_path": excel_path,
            "existe_excel": os.path.exists(excel_path),
            "total_monto": datos["total"],
            "total_items": len(datos["items"]),
            "items": datos["items"]
        })

    return {
        "carpeta": folder_path,
        "total_trabajadores": len(trabajadores),
        "total_comprobantes": total_comprobantes_global,
        "total_monto": total_monto_global,
        "trabajadores": trabajadores
    }


def ejecutar_proceso_reembolsos(folder_path):
    """Ejecuta el procesamiento en segundo plano capturando la salida."""
    global app_state
    start_time = time.time()
    with state_lock:
        app_state["status"] = "running"
        app_state["current_folder"] = folder_path
        app_state["logs"] = []
        app_state["results"] = None
        app_state["error"] = None
        app_state["progress_msg"] = "Iniciando procesamiento de comprobantes..."
        app_state["start_time"] = start_time

    capture = StreamCapture(append_log)
    original_stdout = sys.stdout

    try:
        sys.stdout = capture
        core.procesar_reembolsos(folder_path)
    except Exception as e:
        with state_lock:
            app_state["status"] = "error"
            app_state["error"] = str(e)
            app_state["progress_msg"] = f"Error: {e}"
        append_log(f"ERROR DURANTE EL PROCESO: {e}\n")
    finally:
        sys.stdout = original_stdout

    # Si no hubo error fatal, parsear resultados
    with state_lock:
        if app_state["status"] != "error":
            app_state["status"] = "completed"
            app_state["progress_msg"] = "¡Proceso finalizado!"
            app_state["duration"] = round(time.time() - start_time, 1)
            app_state["results"] = parse_execution_results(folder_path, app_state["logs"])


def escanear_carpeta(folder_path):
    """
    Inspecciona la carpeta seleccionada y devuelve un resumen detallado
    de los trabajadores y archivos encontrados antes de procesar.
    """
    if not os.path.exists(folder_path) or not os.path.isdir(folder_path):
        return {"ok": False, "error": f"La carpeta '{folder_path}' no existe o no es válida."}

    subdirs = [d for d in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, d))]
    
    # Identificar si es una carpeta contenedora de trabajadores o un trabajador directo
    es_semanal = len(subdirs) > 0 and any(d.lower() != 'zextra' for d in subdirs)
    
    if es_semanal:
        carpetas_trabajadores = [os.path.join(folder_path, d) for d in sorted(subdirs) if d.lower() != 'zextra']
    else:
        carpetas_trabajadores = [folder_path]

    trabajadores = []
    total_xml = 0
    total_pdf = 0
    total_img = 0

    for c_path in carpetas_trabajadores:
        nombre = os.path.basename(c_path)
        archivos = os.listdir(c_path)
        
        xmls = [f for f in archivos if f.lower().endswith('.xml')]
        pdfs = [f for f in archivos if f.lower().endswith('.pdf')]
        imgs = [f for f in archivos if f.lower().endswith(('.jpg', '.jpeg', '.png', '.webp'))]
        excels = [f for f in archivos if f.lower().endswith('.xlsx') and not f.startswith('~$')]

        total_xml += len(xmls)
        total_pdf += len(pdfs)
        total_img += len(imgs)

        trabajadores.append({
            "nombre": nombre,
            "ruta": c_path,
            "xmls_count": len(xmls),
            "pdfs_count": len(pdfs),
            "imgs_count": len(imgs),
            "excel_detectado": excels[0] if excels else None
        })

    # Verificar plantilla base
    plantilla = core.encontrar_plantilla_base(folder_path)
    if not plantilla:
        # Intentar buscar en el directorio base del proyecto
        plantilla = core.encontrar_plantilla_base(BASE_DIR)

    return {
        "ok": True,
        "ruta": folder_path,
        "nombre_carpeta": os.path.basename(folder_path),
        "es_semanal": es_semanal,
        "total_trabajadores": len(trabajadores),
        "total_xml": total_xml,
        "total_pdf": total_pdf,
        "total_img": total_img,
        "total_archivos": total_xml + total_pdf + total_img,
        "plantilla_encontrada": os.path.basename(plantilla) if plantilla else None,
        "plantilla_ruta": plantilla,
        "trabajadores": trabajadores
    }


def listar_carpetas_proyecto():
    """Detecta carpetas de semanas o trabajo disponibles en el proyecto."""
    carpetas = []
    for item in sorted(os.listdir(BASE_DIR)):
        full_path = os.path.join(BASE_DIR, item)
        if os.path.isdir(full_path) and not item.startswith('.') and item != '__pycache__' and item != 'templates' and item != 'static':
            # Verificar si contiene subcarpetas o archivos de comprobantes
            subdirs = [d for d in os.listdir(full_path) if os.path.isdir(os.path.join(full_path, d))]
            mtime = os.path.getmtime(full_path)
            carpetas.append({
                "nombre": item,
                "ruta": full_path,
                "subcarpetas": len(subdirs),
                "fecha_modificacion": time.strftime("%d/%m/%Y %H:%M", time.localtime(mtime))
            })
    return carpetas


def seleccionar_carpeta_dialogo(initial_dir=None):
    """
    Abre un diálogo nativo de selección de carpetas compatible con Windows, Linux y macOS.
    """
    if not initial_dir or not os.path.exists(initial_dir):
        initial_dir = BASE_DIR

    # 1. En Windows: intentar Tkinter o PowerShell
    if sys.platform == 'win32':
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            folder = filedialog.askdirectory(initialdir=initial_dir, title="Seleccionar Carpeta de Reembolsos")
            root.destroy()
            if folder:
                return os.path.abspath(folder)
        except Exception:
            pass

    # 2. En macOS: intentar AppleScript
    elif sys.platform == 'darwin':
        try:
            script = f'POSIX path of (choose folder with prompt "Seleccionar Carpeta de Reembolsos" default location POSIX file "{initial_dir}")'
            res = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=120)
            selected = res.stdout.strip()
            if res.returncode == 0 and selected and os.path.isdir(selected):
                return selected
        except Exception:
            pass

    # 3. En Linux: intentar Zenity o KDialog
    else:
        try:
            cmd = [
                'zenity',
                '--file-selection',
                '--directory',
                '--title=Seleccionar Carpeta de Reembolsos',
                f'--filename={initial_dir}/'
            ]
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
            selected = res.stdout.strip()
            if res.returncode == 0 and selected and os.path.isdir(selected):
                return selected
        except Exception:
            pass

        try:
            cmd = ['kdialog', '--getexistingdirectory', initial_dir, '--title', 'Seleccionar Carpeta de Reembolsos']
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
            selected = res.stdout.strip()
            if res.returncode == 0 and selected and os.path.isdir(selected):
                return selected
        except Exception:
            pass

    # Fallback general con Tkinter
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        folder = filedialog.askdirectory(initialdir=initial_dir, title="Seleccionar Carpeta de Reembolsos")
        root.destroy()
        if folder and os.path.isdir(folder):
            return os.path.abspath(folder)
    except Exception:
        pass

    return None


def abrir_en_sistema(ruta):
    """Abre un archivo o carpeta en el gestor/aplicación predeterminada del sistema."""
    if not ruta or not os.path.exists(ruta):
        return False, "La ruta no existe"
    try:
        abs_path = os.path.abspath(ruta)
        if sys.platform == 'win32':
            os.startfile(abs_path)
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', abs_path])
        else:
            subprocess.Popen(['xdg-open', abs_path])
        return True, None
    except Exception as e:
        return False, str(e)


class ThreadedHTTPServer(ThreadingMixIn, HTTPServer):
    """Servidor HTTP multi-hilo para responder rápidamente a peticiones concurrentes."""
    daemon_threads = True


class ReembolsosHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        # Silenciar logs ruidosos del servidor HTTP en consola
        pass

    def send_json_response(self, data, status_code=200):
        self.send_response(status_code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_HEAD(self):
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query)

        # Página Principal
        if path == '/' or path == '/index.html':
            html_path = os.path.join(TEMPLATES_DIR, 'index.html')
            if not os.path.exists(html_path):
                self.send_error(404, "Plantilla index.html no encontrada")
                return
            with open(html_path, 'rb') as f:
                content = f.read()
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Cache-Control', 'no-cache')
            self.end_headers()
            self.wfile.write(content)
            return

        # API: Carpetas del espacio de trabajo
        if path == '/api/workspace-folders':
            carpetas = listar_carpetas_proyecto()
            self.send_json_response({"ok": True, "folders": carpetas})
            return

        # API: Estado actual del procesamiento
        if path == '/api/status':
            with state_lock:
                data = {
                    "status": app_state["status"],
                    "current_folder": app_state["current_folder"],
                    "progress_msg": app_state["progress_msg"],
                    "logs": app_state["logs"],
                    "results": app_state["results"],
                    "error": app_state["error"],
                    "duration": app_state["duration"]
                }
            self.send_json_response(data)
            return

        # API: Descarga de archivo Excel generado
        if path == '/api/download':
            file_param = query.get('file', [''])[0]
            if not file_param or not os.path.exists(file_param):
                self.send_error(404, "Archivo no encontrado")
                return
            filename = os.path.basename(file_param)
            with open(file_param, 'rb') as f:
                excel_data = f.read()
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
            self.send_header('Content-Length', str(len(excel_data)))
            self.end_headers()
            self.wfile.write(excel_data)
            return

        self.send_error(404, "Ruta no encontrada")

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        # Leer cuerpo JSON
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length).decode('utf-8') if content_length > 0 else '{}'
        try:
            req_data = json.loads(body) if body else {}
        except Exception:
            req_data = {}

        # API: Diálogo nativo de selección de carpetas
        if path == '/api/browse-folder':
            initial_dir = req_data.get('initial_dir') or BASE_DIR
            selected = seleccionar_carpeta_dialogo(initial_dir)
            if selected and os.path.isdir(selected):
                self.send_json_response({"ok": True, "path": selected})
            else:
                self.send_json_response({"ok": False, "cancelled": True})
            return

        # API: Escanear carpeta seleccionada
        if path == '/api/scan':
            target_path = req_data.get('folder_path', '').strip()
            if not target_path:
                self.send_json_response({"ok": False, "error": "Debes especificar una ruta de carpeta."})
                return
            scan_data = escanear_carpeta(target_path)
            self.send_json_response(scan_data)
            return

        # API: Iniciar procesamiento de reembolsos
        if path == '/api/process':
            with state_lock:
                if app_state["status"] == "running":
                    self.send_json_response({"ok": False, "error": "Ya hay un proceso en ejecución."})
                    return

            target_path = req_data.get('folder_path', '').strip()
            if not target_path or not os.path.isdir(target_path):
                self.send_json_response({"ok": False, "error": "La carpeta especificada no existe."})
                return

            # Lanzar el proceso en un hilo separado
            worker_thread = threading.Thread(
                target=ejecutar_proceso_reembolsos,
                args=(target_path,),
                daemon=True
            )
            worker_thread.start()

            self.send_json_response({"ok": True, "message": "Proceso iniciado correctamente."})
            return

        # API: Abrir carpeta en el gestor de archivos del sistema
        if path == '/api/open-folder':
            folder_to_open = req_data.get('path', '').strip()
            ok, err = abrir_en_sistema(folder_to_open)
            if ok:
                self.send_json_response({"ok": True})
            else:
                self.send_json_response({"ok": False, "error": err or "Carpeta no encontrada"})
            return

        # API: Abrir archivo con la aplicación predeterminada del sistema
        if path == '/api/open-file':
            file_to_open = req_data.get('path', '').strip()
            ok, err = abrir_en_sistema(file_to_open)
            if ok:
                self.send_json_response({"ok": True})
            else:
                self.send_json_response({"ok": False, "error": err or "Archivo no encontrado"})
            return

        self.send_error(404, "Ruta no encontrada")


def iniciar_servidor(puerto=5050, abrir_navegador=True):
    """Inicia el servidor web local y abre el navegador."""
    server_address = ('127.0.0.1', puerto)
    
    # Si el puerto está ocupado, probar puertos subsiguientes
    max_intentos = 10
    httpd = None
    for p in range(puerto, puerto + max_intentos):
        try:
            server_address = ('127.0.0.1', p)
            httpd = ThreadedHTTPServer(server_address, ReembolsosHandler)
            puerto = p
            break
        except OSError:
            continue

    if not httpd:
        print(f"Error: No se pudo enlazar el servidor en los puertos {puerto} a {puerto+max_intentos}.")
        sys.exit(1)

    url = f"http://127.0.0.1:{puerto}"
    print("=" * 70)
    print(f"  APLICACIÓN DE REEMBOLSOS INICIADA")
    print(f"  URL Local: {url}")
    print("=" * 70)

    if abrir_navegador:
        def open_browser():
            time.sleep(1.0)
            webbrowser.open(url)
        threading.Thread(target=open_browser, daemon=True).start()

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nServidor detenido.")
        httpd.server_close()


if __name__ == '__main__':
    puerto = int(sys.argv[1]) if len(sys.argv) > 1 else 5050
    # No abrir navegador si se pasa la bandera --no-browser
    abrir_nav = '--no-browser' not in sys.argv
    iniciar_servidor(puerto=puerto, abrir_navegador=abrir_nav)
