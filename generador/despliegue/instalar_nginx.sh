#!/usr/bin/env bash
# Agrega /formularios/ al sitio HTTPS de nginx (cloud-https). Correr en tecsito con:
#   sudo bash ~/web/formularios_oec/generador/despliegue/instalar_nginx.sh
# Hace respaldo, valida con nginx -t y, si falla, deja todo como estaba.
set -euo pipefail
SITIO=/etc/nginx/sites-available/cloud-https
FRAGMENTO="$(dirname "$(readlink -f "$0")")/nginx-formularios.conf"
if grep -q "location /formularios/" "$SITIO"; then
  echo "Ya está configurado /formularios/ en $SITIO. No se cambió nada."; exit 0
fi
RESPALDO="$SITIO.bak-antes-formularios-$(date +%Y%m%d-%H%M%S)"
cp -p "$SITIO" "$RESPALDO"
# Insertar el fragmento antes de la última llave de cierre del bloque server
python3 - "$SITIO" "$FRAGMENTO" <<'PY'
import sys
sitio, fragmento = sys.argv[1], open(sys.argv[2]).read()
texto = open(sitio).read()
fin = texto.rstrip().rfind("}")
open(sitio, "w").write(texto[:fin].rstrip() + "\n\n" + fragmento.rstrip() + "\n}\n")
PY
if nginx -t; then
  systemctl reload nginx
  echo "Listo: https://cloud.capteec.com/formularios/  (respaldo: $RESPALDO)"
else
  cp -p "$RESPALDO" "$SITIO"
  echo "nginx -t falló: se restauró el respaldo. No se recargó nginx."; exit 1
fi
