#!/usr/bin/env sh
set -eu
cd "$(dirname "$0")"
if [ ! -x .venv/bin/python ]; then
  echo "Ambiente non trovato. Eseguire prima ./install.sh"
  exit 1
fi
exec .venv/bin/python main.py
