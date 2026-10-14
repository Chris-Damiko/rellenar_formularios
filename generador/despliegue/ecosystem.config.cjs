// pm2: portal de formularios OEC en tecsito (127.0.0.1:5010, publicado por nginx en /formularios/)
module.exports = {
  apps: [{
    name: "formularios-oec",
    cwd: "/home/capteec/web/formularios_oec/generador",
    script: "/home/capteec/web/formularios_oec/generador/.venv/bin/python",
    args: "portal/app.py --puerto 5010 --servidor waitress",
    interpreter: "none",
    env: {
      PORTAL_LOGIN: "1",
      PORTAL_HTTPS: "1",
      RAIZ_OEC: "/home/capteec/web/formularios_oec/datos",
      PYTHONUNBUFFERED: "1",
    },
    max_memory_restart: "600M",
    restart_delay: 3000,
  }],
};
