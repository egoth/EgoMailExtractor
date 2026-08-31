#!/usr/bin/env sh
set -eu
cd "$(dirname "$0")"
if [ ! -x .venv/bin/python ]; then
  echo "Ambiente non trovato. Eseguire prima ./install.sh"
  exit 1
fi
. .venv/bin/activate
python -m pip install pyinstaller
ICON="$(pwd)/assets/egomail_icon.ico"
if [ ! -f "$ICON" ]; then
  echo "Icona applicazione non trovata: $ICON"
  exit 1
fi
rm -rf build dist/EgoMailExtractor EgoMailExtractor.spec
SEP=":"
pyinstaller --noconfirm --clean --windowed --name EgoMailExtractor --add-data "builtin_profiles${SEP}builtin_profiles" --add-data "help${SEP}help" --add-data "assets${SEP}assets" --add-data "LICENSE.txt${SEP}." --icon "$ICON" main.py
echo "Applicazione creata nella directory dist/."
