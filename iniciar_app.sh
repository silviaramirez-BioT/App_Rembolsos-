#!/bin/bash

# ==============================================================================
# Lanzador de la Aplicación de Reembolsos
# ==============================================================================

# Ir al directorio del script
cd "$(dirname "$0")"

echo "=================================================================="
echo "    Iniciando Aplicación de Reembolsos con Selector de Carpetas"
echo "=================================================================="

# Detectar comando de Python disponible (usando uv o el entorno virtual .venv)
if [ -f "$HOME/.local/bin/uv" ]; then
    PYTHON_CMD="$HOME/.local/bin/uv run python"
elif command -v uv &> /dev/null; then
    PYTHON_CMD="uv run python"
elif [ -f ".venv/bin/python3" ]; then
    PYTHON_CMD=".venv/bin/python3"
else
    PYTHON_CMD="python3"
fi

echo "Iniciando servidor local y abriendo el navegador..."
$PYTHON_CMD app.py 5050
