#!/bin/bash
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
PYTHON="/Users/gianpier/INTERBANK_GENERADOR/.venv_app/bin/python"

if [ ! -f "$PYTHON" ]; then
    PYTHON="python3"
fi

echo "================================================================="
echo "🏦 INTERBANK - CONSOLIDADOR FINANCIERO VP RETAIL 2026"
echo "================================================================="
echo "Iniciando aplicación de escritorio CustomTkinter..."
"$PYTHON" "$DIR/app_desktop.py"
