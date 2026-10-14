#!/usr/bin/env bash
# Abre el portal de registro de candidatos.
#   ./iniciar_portal.sh          solo en esta PC  -> http://127.0.0.1:5000
#   ./iniciar_portal.sh --red    también desde otras PCs de la red local
cd "$(dirname "$0")"
exec .venv/bin/python portal/app.py "$@"
