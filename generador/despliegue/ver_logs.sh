#!/usr/bin/env bash
# Muestra los registros del portal en tecsito. Desde la PC:
#   bash generador/despliegue/ver_logs.sh            # últimas 80 líneas de actividad + errores de pm2
#   bash generador/despliegue/ver_logs.sh errores    # solo errores con su traza
#   bash generador/despliegue/ver_logs.sh vivo       # seguir en vivo (Ctrl+C para salir)
SSH="ssh -i $HOME/.ssh/id_ed25519_tecsito -o BatchMode=yes capteec@100.95.125.45"
LOG=/home/capteec/web/formularios_oec/generador/portal/portal.log
case "${1:-}" in
  errores) $SSH "grep -n -A25 -E ' (ERROR|WARNING) ' $LOG | tail -200" ;;
  vivo)    $SSH -t "tail -f $LOG" ;;
  *)       $SSH "echo '== pm2'; pm2 describe formularios-oec | grep -E 'status|restarts|uptime'; \
                 echo '== actividad (portal.log)'; tail -n 80 $LOG; \
                 echo '== errores de pm2'; tail -n 40 ~/.pm2/logs/formularios-oec-error.log" ;;
esac
