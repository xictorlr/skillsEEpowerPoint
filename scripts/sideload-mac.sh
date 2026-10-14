#!/usr/bin/env bash
# Copia un manifest a la carpeta de add-ins sideloaded de PowerPoint para Mac.
set -euo pipefail
MANIFEST="${1:-manifest.xml}"
WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
mkdir -p "$WEF"
cp "$MANIFEST" "$WEF/slide-tools-$(basename "$MANIFEST")"
echo "Copiado a: $WEF"
echo "Reinicia PowerPoint y abre Insertar > Complementos > Mis complementos > Complementos para desarrolladores."
