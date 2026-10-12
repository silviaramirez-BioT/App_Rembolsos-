#!/usr/bin/env bash
set -e

# =============================================================
# Compilador del Gestor de Reembolsos (PyInstaller, multiplataforma)
# =============================================================
# Genera un ejecutable standalone a partir de app.py usando PyInstaller
# (mejor compatibilidad con rapidocr_onnxruntime / pdfplumber / onnxruntime
#  que Nuitka, gracias a sus hooks y collect_all).
#
# Compatible con Linux, macOS y Windows (Git Bash / WSL / nativo).
#
# Salida:
#   Windows: dist\Gestor_Reembolsos.exe
#   Linux/macOS: dist/Gestor_Reembolsos
# =============================================================

DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$DIR"

echo "=================================================================="
echo "  COMPILADOR — Gestor de Reembolsos (PyInstaller)"
echo "=================================================================="

# 1) Detectar Python disponible
if command -v python3 &> /dev/null; then
    PYTHON_CMD=python3
elif command -v python &> /dev/null; then
    PYTHON_CMD=python
elif command -v py &> /dev/null; then
    PYTHON_CMD=py
else
    echo "[ERROR] No se encontró Python en el PATH."
    echo "Instala Python 3.10+ desde https://www.python.org/downloads/"
    exit 1
fi

echo "[INFO] Usando Python: $($PYTHON_CMD --version)"

# 2) Limpiar builds previos
echo "[0/4] Limpiando builds anteriores..."
rm -rf build/ dist/ *.spec.bak 2>/dev/null || true

# 3) Crear/activar entorno virtual
if [ ! -d ".venv" ]; then
    echo "[1/4] Creando entorno virtual (.venv)..."
    $PYTHON_CMD -m venv .venv
fi

if [ -f ".venv/bin/activate" ]; then
    # shellcheck disable=SC1091
    source ".venv/bin/activate"
    PYTHON_CMD=python
elif [ -f ".venv/Scripts/activate" ]; then
    # shellcheck disable=SC1091
    source ".venv/Scripts/activate"
    PYTHON_CMD=python
fi

# 4) Instalar dependencias + PyInstaller
echo "[2/4] Instalando dependencias y PyInstaller..."
python -m pip install --upgrade pip --quiet
python -m pip install -r requirements.txt --quiet
python -m pip install pyinstaller --quiet

# 5) Compilar con PyInstaller usando el .spec
echo "[3/4] Compilando Gestor_Reembolsos (puede tardar 1-3 min)..."
echo "      Incluye OCR (RapidOCR + ONNX), lectura de PDFs y servidor web."
echo ""
python -m PyInstaller gestor_reembolsos_pyinstaller.spec --noconfirm --clean

# 6) Resultado final
if [[ "$OSTYPE" == "msys" || "$OSTYPE" == "cygwin" || "$OSTYPE" == "win32" ]]; then
    OUTPUT_PATH="dist/Gestor_Reembolsos.exe"
else
    OUTPUT_PATH="dist/Gestor_Reembolsos"
fi

echo ""
echo "=================================================================="
echo "  ¡COMPILACIÓN EXITOSA!"
echo "  Ejecutable generado: $DIR/$OUTPUT_PATH"
echo "=================================================================="
echo ""
echo "  Puedes enviar este único archivo a cualquier usuario."
echo "  Al hacer doble clic, se abrirá en su navegador predeterminado"
echo "  en http://127.0.0.1:5050"
echo ""
echo "  NOTA: El .exe/.bin se genera nativamente para el SO donde"
echo "  corriste este script. Para un .exe de Windows desde Linux,"
echo "  necesitas ejecutar build.sh en Windows (Git Bash / WSL / VM)."
echo "=================================================================="