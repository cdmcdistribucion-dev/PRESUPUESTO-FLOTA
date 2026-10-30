#!/bin/sh
cd "$(dirname "$0")" || exit 1
node actualizar_datos.js || { echo "Hubo un error. No se publicó nada."; exit 1; }
git add -A && git commit -m "Actualización de datos $(date '+%d/%m/%Y %H:%M')" && git push
echo "Listo. En 1-2 minutos el tablero publicado muestra los datos nuevos."
