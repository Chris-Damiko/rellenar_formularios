#!/usr/bin/env bash
# Sube el portal a tecsito y lo (re)inicia con pm2. Correr desde la PC:
#   bash generador/despliegue/desplegar_tecsito.sh
# La primera vez también copia las bases Excel, ajustes.json y estado.json.
# Las siguientes veces NO pisa los datos del servidor: bases, salida/, respaldos/,
# asignaciones.json, ajustes.json, estado.json, usuarios y clave de sesión.
set -euo pipefail
cd "$(dirname "$0")/.."
SSH_KEY=~/.ssh/id_ed25519_tecsito
DEST=capteec@100.95.125.45
REMOTO=/home/capteec/web/formularios_oec
SSH="ssh -i $SSH_KEY -o BatchMode=yes"

echo "== código y plantillas"
$SSH $DEST "mkdir -p $REMOTO/generador $REMOTO/datos ~/.local/bin"
rsync -az --delete -e "$SSH" \
  --exclude '.venv/' --exclude '__pycache__/' --exclude 'salida/' --exclude 'respaldos/' \
  --exclude 'informes/' --exclude '*.log' --exclude 'asignaciones.json' --exclude 'ajustes.json' \
  --exclude 'estado.json' --exclude 'portal/usuarios.json' --exclude 'portal/secret.key' --exclude 'portal/subidas/' --exclude 'portal/portal.log*' \
  --exclude 'generar_todo.log' --exclude 'verificar_todo.log' \
  ./ "$DEST:$REMOTO/generador/"
echo "== ajustes y estado (solo si no existen en el servidor)"
rsync -az --ignore-existing -e "$SSH" ajustes.json estado.json "$DEST:$REMOTO/generador/"

echo "== bases maestras (solo si no existen en el servidor)"
.venv/bin/python - <<'PY' > /tmp/bases_formularios.txt
import os, config
for p in config.PERFILES:
    print(os.path.relpath(config.ruta_base(p), config.RAIZ_OEC))
PY
rsync -az --ignore-existing --files-from=/tmp/bases_formularios.txt -e "$SSH" \
  "$(.venv/bin/python -c 'import config; print(config.RAIZ_OEC)')/" "$DEST:$REMOTO/datos/"

echo "== uv y entorno de Python"
if ! $SSH $DEST "test -x ~/.local/bin/uv"; then
  scp -i $SSH_KEY "$(command -v uv)" "$DEST:~/.local/bin/uv"
fi
$SSH $DEST "cd $REMOTO/generador && ( test -x .venv/bin/python || ~/.local/bin/uv venv -q --python /usr/bin/python3 .venv ) \
  && ~/.local/bin/uv pip install -q --python .venv/bin/python -r despliegue/requirements.txt"

echo "== pm2"
$SSH $DEST "cd $REMOTO/generador && (pm2 describe formularios-oec >/dev/null 2>&1 \
  && pm2 restart formularios-oec --update-env >/dev/null || pm2 start despliegue/ecosystem.config.cjs >/dev/null) \
  && pm2 save >/dev/null && sleep 3 && pm2 describe formularios-oec | grep -E 'status|restarts' \
  && curl -s -o /dev/null -w 'portal local: HTTP %{http_code}\n' http://127.0.0.1:5010/entrar"
