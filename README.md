# Gestor de Reembolsos

Aplicación que automatiza el procesamiento semanal de comprobantes
fiscales: extrae datos de **XML CFDI** (SAT), **PDFs** y **tickets
fotográficos** (vía OCR), deduplica, valida contra plantillas Excel
y genera una hoja de cálculo por persona.

Incluye una interfaz web local para seleccionar carpetas, monitorear
el progreso y descargar los Excel resultantes.

---

## 🚀 Uso rápido (usuario final)

1. Doble clic en `Gestor_Reembolsos.exe` (Windows) o `Gestor_Reembolsos`
   (Linux/macOS, dentro de `dist/`).
2. Se abre automáticamente en el navegador predeterminado en
   `http://127.0.0.1:5050`.
3. Selecciona la carpeta de la semana con comprobantes y pulsa
   **Procesar**.
4. Al finalizar, descarga o abre los Excel generados.

Para cerrar la aplicación: cierra la ventana de la consola del servidor
(o `Ctrl+C` en ella) y la pestaña del navegador.

---

## 🧑‍💻 Uso desde código fuente (desarrolladores)

### Windows
Doble clic en **`iniciar_app.bat`** (configura el entorno la primera vez
y abre el navegador).

### Manual (cualquier OS)
```bash
python -m pip install -r requirements.txt
python app.py                # → http://127.0.0.1:5050
```

Argumentos opcionales:
```bash
python app.py 8080           # puerto personalizado
python app.py 5050 --no-browser   # no abrir navegador
```

---

## 🔨 Compilar el ejecutable

El proyecto se compila con **PyInstaller** (por defecto, recomendado
por su compatibilidad con rapidocr/pdfplumber/onnxruntime).

### PyInstaller (recomendado — funciona bien con OCR/PDF)

PyInstaller incluye hooks para empaquetar correctamente modelos ONNX,
binarios nativos y archivos de configuración de las librerías usadas.

**Requisitos:** Python 3.10+ (no requiere gcc ni MSVC).

**Build:**
```bash
bash build.sh
```

Salida en `dist/`:
- Windows → `Gestor_Reembolsos.exe`
- Linux/macOS → `Gestor_Reembolsos`

### Nuitka (alternativa experimental)

Nuitka traduce Python → C → binario nativo (más difícil de decompilar),
pero su soporte para `rapidocr_onnxruntime` y `onnxruntime` es **frágil**
y suele requerir flags adicionales (`--enable-plugin=data-files`,
`--include-data-dir=<ruta de modelos>`).

```bash
pip install -r requirements.txt nuitka ordered-set
python -m nuitka \
    --standalone --onefile \
    --output-filename=Gestor_Reembolsos.exe \
    --include-data-dir=templates=templates \
    --include-package=rapidocr_onnxruntime \
    --include-package=pdfplumber --include-package=onnxruntime \
    --enable-plugin=anti-bloat \
    app.py
```

### ⚠️ Sobre el `.exe` vs binario sin extensión
El compilador genera el binario nativamente para el SO donde corres
el build:
- **Corres en Windows** → produces `dist/Gestor_Reembolsos.exe`
- **Corres en Linux/macOS** → produces `dist/Gestor_Reembolsos`

Si necesitas un `.exe` para Windows pero solo tienes Linux, ejecuta el
build en una **VM Windows**, en **WSL**, o con **Wine** (engorroso).
No hay cross-compilation oficial.

---

## 📦 Distribución

El entregable al usuario final es **un único archivo** dentro de `dist/`:

- `Gestor_Reembolsos.exe` (Windows)
- `Gestor_Reembolsos` (Linux/macOS)

No requiere Python ni librerías adicionales en la máquina destino.

---

## 📁 Estructura del proyecto

```
app_reembolsos/
├── README.md                            Este archivo
├── app.py                               Servidor web + GUI (FastAPI-style)
├── procesar_reembolsos.py               Lógica de extracción XML/PDF/OCR/Excel
├── requirements.txt                     Dependencias de runtime
├── build.sh                             Compilador (PyInstaller por defecto)
├── iniciar_app.bat                      Lanzador Windows desde código fuente
├── iniciar_app.sh                       Lanzador Linux/macOS desde código fuente
├── gestor_reembolsos_pyinstaller.spec   Configuración de PyInstaller
└── templates/
    └── index.html                       Interfaz web (frontend)
```

---

## ⚙️ ¿Cómo funciona?

1. **Entrada:** carpetas semanales con subcarpetas por trabajador, cada
   una conteniendo XMLs CFDI, PDFs de facturas, PDFs de tickets e
   imágenes de tickets.
2. **Extracción:**
   - **XML CFDI** → parseo directo con `xml.etree.ElementTree`
   - **PDF CFDI** → `pdfplumber` extrae UUID, total, fecha y emisor
   - **Tickets (imagen/PDF)** → OCR con **RapidOCR** (ONNX Runtime)
3. **Deduplicación:** si el mismo gasto aparece como XML, PDF y ticket,
   se conserva una sola entrada (XML gana como fuente primaria).
4. **Plantilla Excel:** se inyectan los registros en la tabla `Table1`
   de una plantilla `.xlsx` existente, preservando estilos y fórmulas.
5. **Salida:** un Excel actualizado por trabajador en su carpeta.

---

## 📋 Dependencias clave

- `pdfplumber` — lectura de PDFs CFDI
- `rapidocr-onnxruntime` — OCR de tickets fotográficos
- `onnxruntime` — backend de inferencia ONNX
- `numpy` — manipulación de imágenes para OCR

Todas las dependencias se instalan con `pip install -r requirements.txt`.