"""
Portal web para registrar candidatos nuevos en la base Excel de cada perfil y
generar sus formularios al instante.

USO:
    cd generador
    .venv/bin/python portal/app.py            # http://127.0.0.1:5000 (solo esta PC)
    .venv/bin/python portal/app.py --red      # también desde otras PCs de la red local

En el servidor (detrás de nginx, con inicio de sesión):
    PORTAL_LOGIN=1 PORTAL_HTTPS=1 .venv/bin/python portal/app.py --puerto 5010 --servidor waitress
    Usuarios: .venv/bin/python portal/usuarios.py agregar <usuario>

Qué hace al registrar:
    1. Agrega la fila al final de la base maestra del perfil (con respaldo previo).
    2. Guarda el examinador, las UC y el lugar elegidos (asignaciones.json).
    3. Genera los formularios en salida/<PERFIL>/<CODIGO>/FORMULARIOS/.
"""

import argparse
import datetime
import io
import json
import logging
import logging.handlers
import traceback
import os
import re
import secrets
import shutil
import sys
import time
import zipfile

from flask import Flask, abort, flash, redirect, render_template, request, send_file, session, url_for
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.security import check_password_hash

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import ajustes  # noqa: E402
import base  # noqa: E402
import base_escritura  # noqa: E402
import config  # noqa: E402
import generar  # noqa: E402

CARPETA_PORTAL = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_USUARIOS = os.path.join(CARPETA_PORTAL, "usuarios.json")
ARCHIVO_SECRETO = os.path.join(CARPETA_PORTAL, "secret.key")


def clave_secreta():
    """Clave de las sesiones: se crea una vez al azar y se guarda en portal/secret.key."""
    if not os.path.exists(ARCHIVO_SECRETO):
        with open(os.open(ARCHIVO_SECRETO, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600), "w") as f:
            f.write(secrets.token_hex(32))
    with open(ARCHIVO_SECRETO) as f:
        return f.read().strip()


app = Flask(__name__)
app.secret_key = clave_secreta()
# Detrás de nginx: respeta https, el host y el prefijo /formularios (X-Forwarded-Prefix)
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1)
app.config.update(
    MAX_CONTENT_LENGTH=25 * 1024 * 1024,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("PORTAL_HTTPS") == "1",
    PERMANENT_SESSION_LIFETIME=datetime.timedelta(hours=10),
)
# PORTAL_LOGIN=1 obliga a iniciar sesión (en el servidor). En la PC local, sin la
# variable, el portal abre directo como antes.
LOGIN_OBLIGATORIO = os.environ.get("PORTAL_LOGIN") == "1"
_fallos = {}

# Registro de actividad y errores: portal/portal.log (rota a los 5 MB, guarda 5) y también en pm2
registro = logging.getLogger("portal")
registro.setLevel(logging.INFO)
_archivo_log = logging.handlers.RotatingFileHandler(os.path.join(CARPETA_PORTAL, "portal.log"),
                                                    maxBytes=5 * 1024 * 1024, backupCount=5, encoding="utf-8")
_archivo_log.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
registro.addHandler(_archivo_log)
registro.addHandler(logging.StreamHandler(sys.stdout))


def _quien():
    return f"{session.get('usuario') or '-'}@{request.remote_addr or '?'}"


@app.after_request
def anotar_peticion(respuesta):
    if request.endpoint != "static":
        registro.info(f"{_quien()} {request.method} {request.full_path.rstrip('?')} -> {respuesta.status_code}")
    return respuesta


@app.errorhandler(Exception)
def anotar_error(error):
    from werkzeug.exceptions import HTTPException
    if isinstance(error, HTTPException):
        return error
    registro.error(f"{_quien()} {request.method} {request.full_path} ERROR {error!r}\n{traceback.format_exc()}")
    return render_template("error.html", error=error), 500


def leer_usuarios():
    try:
        with open(ARCHIVO_USUARIOS, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}


@app.before_request
def exigir_sesion():
    if not LOGIN_OBLIGATORIO or request.endpoint in ("entrar", "static"):
        return None
    if not session.get("usuario") or session["usuario"] not in leer_usuarios():
        return redirect(url_for("entrar", siguiente=request.full_path if request.method == "GET" else None))
    return None


@app.route("/entrar", methods=["GET", "POST"])
def entrar():
    if request.method == "POST":
        ip = request.remote_addr or "?"
        recientes = [t for t in _fallos.get(ip, []) if time.time() - t < 900]
        if len(recientes) >= 8:
            flash("Demasiados intentos. Espere 15 minutos.", "error")
            return render_template("entrar.html"), 429
        usuario = (request.form.get("usuario") or "").strip().lower()
        datos = leer_usuarios().get(usuario)
        if datos and check_password_hash(datos["hash"], request.form.get("clave") or ""):
            session.clear()
            session.permanent = True
            session["usuario"] = usuario
            _fallos.pop(ip, None)
            registro.info(f"LOGIN ok {usuario}@{ip}")
            siguiente = request.args.get("siguiente") or ""
            if not siguiente.startswith("/") or siguiente.startswith("//"):
                siguiente = None
            # "siguiente" es la ruta sin el prefijo de nginx (/formularios): se le antepone
            return redirect(request.script_root + siguiente if siguiente else url_for("inicio"))
        _fallos[ip] = recientes + [time.time()]
        registro.warning(f"LOGIN fallido usuario={usuario!r} ip={ip}")
        time.sleep(1)
        flash("Usuario o contraseña incorrectos.", "error")
    return render_template("entrar.html")


@app.route("/salir")
def salir():
    session.clear()
    return redirect(url_for("entrar") if LOGIN_OBLIGATORIO else url_for("inicio"))


NOMBRES_PERFIL = {
    "PRC": "Prevención de riesgos: construcción", "PRE": "Prevención de riesgos: energía eléctrica",
    "GS": "Gestión de soldadura", "IE": "Instalaciones eléctricas", "MA": "Mecatrónica automotriz",
    "OR": "Operaciones de líneas y redes energizadas", "AL": "Auxiliar de liniero",
    "AO": "Asistente en ofimática", "FF": "Formación de formadores", "MQ": "Maquillaje",
}
PROVINCIAS = ["AZUAY", "BOLIVAR", "CAÑAR", "CARCHI", "CHIMBORAZO", "COTOPAXI", "EL ORO", "ESMERALDAS",
              "GALAPAGOS", "GUAYAS", "IMBABURA", "LOJA", "LOS RIOS", "MANABI", "MORONA SANTIAGO", "NAPO",
              "ORELLANA", "PASTAZA", "PICHINCHA", "SANTA ELENA", "SANTO DOMINGO DE LOS TSACHILAS",
              "SUCUMBIOS", "TUNGURAHUA", "ZAMORA CHINCHIPE"]
ESTADOS_CIVILES = ["SOLTERO", "CASADO", "DIVORCIADO", "VIUDO", "UNION DE HECHO"]
NIVELES = [("secundaria", "Secundaria / bachillerato"), ("superior", "Tercer nivel"), ("cuarto", "Cuarto nivel")]

# Lectura de bases con caché según la fecha de modificación del archivo
_cache = {}


def leer(perfil):
    ruta = config.ruta_base(perfil)
    mtime = os.path.getmtime(ruta)
    if _cache.get(perfil, (None,))[0] != mtime:
        candidatos, col = base.leer_base(ruta, config.PERFILES[perfil]["hoja"])
        _cache[perfil] = (mtime, candidatos, col)
    return _cache[perfil][1], _cache[perfil][2]


def perfil_o_404(perfil):
    perfil = perfil.upper()
    if perfil not in config.PERFILES:
        abort(404)
    return perfil


def numero(codigo):
    return int(codigo.split("-")[1])


def estado_formularios(perfil, codigo):
    if os.path.isdir(os.path.join(generar.CARPETA_SALIDA, perfil, codigo, "FORMULARIOS")):
        return "generados"
    if generar.tiene_formularios(perfil, codigo):
        return "en el OEC"
    return "pendiente"


@app.template_filter("fecha")
def filtro_fecha(valor):
    if isinstance(valor, (datetime.date, datetime.datetime)):
        return valor.strftime("%d/%m/%Y")
    return valor or ""


@app.context_processor
def comunes():
    return {"PERFILES": config.PERFILES, "NOMBRES_PERFIL": NOMBRES_PERFIL,
            "usuario_actual": session.get("usuario"), "login": LOGIN_OBLIGATORIO}


# ---------------------------------------------------------------- inicio

@app.route("/")
def inicio():
    aj = ajustes.leer()
    estado = generar.leer_estado()
    tarjetas = []
    for p in config.PERFILES:
        try:
            candidatos, _ = leer(p)
            ultimo_base = max(candidatos, key=numero) if candidatos else "—"
            total = len(candidatos)
        except Exception as e:  # base dañada o ausente: se muestra el error en la tarjeta
            ultimo_base, total = f"error: {e}", 0
        exa = ajustes.examinador(aj, aj["perfiles"][p]["examinador"])
        tarjetas.append({"perfil": p, "nombre": NOMBRES_PERFIL[p], "total": total, "ultimo_base": ultimo_base,
                         "ultimo_generado": estado.get(p, "—"), "examinador": exa["nombre"] if exa else "—",
                         "ucs": ", ".join(aj["perfiles"][p]["ucs"])})
    return render_template("inicio.html", tarjetas=tarjetas)


# ---------------------------------------------------------------- registrar

def _fecha(texto):
    return datetime.datetime.strptime(texto, "%Y-%m-%d") if texto else None


def _hora(texto):
    return datetime.datetime.strptime(texto, "%H:%M").time() if texto else None


def candidato_desde_formulario(perfil, form, col):
    """Arma el candidato (mismo formato que base.leer_base) con lo que llegó del formulario."""
    t = lambda k: (form.get(k) or "").strip()
    may = lambda k: re.sub(r"\s+", " ", t(k)).upper()
    aj = ajustes.leer()
    lugar = ajustes.lugar(aj, perfil, t("lugar"))
    educacion = {}
    for nivel in ["primaria"] + [n for n, _ in NIVELES]:
        if t(f"edu_{nivel}_institucion"):
            educacion[nivel] = {k: may(f"edu_{nivel}_{k}") for k in ("institucion", "pais", "ciudad", "titulo")}
    solo_primaria = form.get("solo_primaria") == "on"
    if solo_primaria and "primaria" not in educacion:
        educacion["primaria"] = dict(generar.ESTUDIO_LECTO)
    if solo_primaria:
        marca = "PRI"
    else:
        marca = "IV" if "cuarto" in educacion else "III" if "superior" in educacion else "II" if "secundaria" in educacion else ""
    experiencias = []
    for i in range(1, len(col.experiencias) + 1):
        if t(f"exp{i}_empresa"):
            experiencias.append({"empresa": may(f"exp{i}_empresa"), "direccion": may(f"exp{i}_direccion"),
                                 "telefono": t(f"exp{i}_telefono"), "desde": _fecha(t(f"exp{i}_desde")),
                                 "hasta": _fecha(t(f"exp{i}_hasta")), "actividad": t(f"exp{i}_actividad")})
    capacitaciones = []
    for i in range(1, len(col.capacitaciones) + 1):
        if t(f"cap{i}_curso"):
            horas = t(f"cap{i}_horas")
            capacitaciones.append({"curso": may(f"cap{i}_curso"), "institucion": may(f"cap{i}_institucion"),
                                   "fechas": t(f"cap{i}_fechas"), "horas": int(horas) if horas.isdigit() else horas})
    edad = t("edad")
    costo = t("costo")
    return {
        "codigo": may("codigo"), "marca": marca, "nombre": may("nombre"), "cedula": t("cedula"),
        "cod_exam": may("cod_exam"), "f_naci": _fecha(t("f_naci")), "edad": int(edad) if edad.isdigit() else edad,
        "estado": may("estado"), "provincia": may("provincia"), "canton": may("canton"), "parroquia": may("parroquia"),
        "direccion": may("direccion"), "correo": t("correo").lower(), "telefono": t("telefono"), "celular": t("celular"),
        "experiencias": experiencias, "capacitaciones": capacitaciones, "educacion": educacion,
        "costo": float(costo) if re.fullmatch(r"\d+(\.\d+)?", costo) else costo,
        "fecha_reg": _fecha(t("fecha_reg")), "fecha_acep": _fecha(t("fecha_acep")), "fecha_exam": _fecha(t("fecha_exam")),
        "hora": _hora(t("hora")),
        "lugar": lugar["lugar"] if lugar else "", "tipo_lugar": lugar["tipo"] if lugar else "",
        "ciudad_examen": lugar["ciudad"] if lugar else "", "direccion_examen": lugar["direccion"] if lugar else "",
        "amarilla": form.get("ausente") == "on",
        "permitir_repetido": form.get("permitir_repetido") == "on",
    }


def errores_de(cand, perfil, ucs, aj):
    errores = []
    for campo, nombre in [("codigo", "código"), ("nombre", "nombres y apellidos"), ("cedula", "cédula"),
                          ("fecha_reg", "fecha de registro"), ("fecha_acep", "fecha de aceptación"),
                          ("fecha_exam", "fecha de examen")]:
        if not cand[campo]:
            errores.append(f"Falta {nombre}.")
    if cand["codigo"] and not re.fullmatch(rf"{perfil}-\d+", cand["codigo"]):
        errores.append(f"El código debe tener la forma {perfil}-0000.")
    if not ucs:
        errores.append("Marque al menos una unidad de competencia.")
    if len(ucs) > 5:
        errores.append("El C02 tiene espacio para 5 unidades de competencia como máximo.")
    return errores


@app.route("/perfil/<perfil>/nuevo", methods=["GET", "POST"])
def nuevo(perfil):
    perfil = perfil_o_404(perfil)
    aj = ajustes.leer()
    candidatos, col = leer(perfil)
    p = aj["perfiles"][perfil]
    examinadores = ajustes.examinadores_de(aj, perfil)
    if request.method == "POST":
        form = request.form
        cand = candidato_desde_formulario(perfil, form, col)
        ucs = [u for u in form.getlist("ucs") if u]
        errores = errores_de(cand, perfil, ucs, aj)
        if cand["cedula"] and not generar.cedula_valida(cand["cedula"]) and form.get("cedula_extranjera") != "on":
            errores.append("La cédula no pasa el dígito verificador. Si es pasaporte o extranjera, marque la casilla.")
        if not errores:
            try:
                codigo, respaldo = base_escritura.agregar(perfil, cand)
                generar.guardar_asignacion(codigo, {"examinador": form.get("examinador") or p["examinador"],
                                                    "ucs": ucs, "lugar": form.get("lugar") or p["lugar"]})
                _cache.pop(perfil, None)
                candidatos, _ = leer(perfil)
                _, avisos = generar.generar_uno(perfil, codigo, candidatos)
                registro.info(f"REGISTRO {_quien()} {codigo} {cand['nombre']} respaldo={os.path.basename(respaldo)} avisos={avisos}")
                flash(f"{codigo} registrado en la base y formularios generados.", "ok")
                for a in avisos:
                    flash(a, "aviso")
                return redirect(url_for("candidato", perfil=perfil, codigo=codigo))
            except base_escritura.BaseOcupada as e:
                registro.warning(f"REGISTRO {_quien()} {perfil} base ocupada: {e}")
                errores.append(str(e))
            except Exception as e:
                registro.error(f"REGISTRO {_quien()} {perfil} {cand.get('codigo')} FALLÓ {e!r}\n{traceback.format_exc()}")
                errores.append(f"No se pudo guardar: {e}")
        for e in errores:
            flash(e, "error")
        valores = dict(form)
        valores["ucs"] = ucs
    else:
        hoy = datetime.date.today().isoformat()
        valores = {"codigo": base_escritura.siguiente_codigo(perfil, candidatos), "fecha_reg": hoy, "fecha_acep": hoy,
                   "fecha_exam": hoy, "hora": "16:00", "telefono": "07-9999999", "costo": p.get("costo", 10),
                   "examinador": p["examinador"], "lugar": p["lugar"], "ucs": p["ucs"],
                   **{f"edu_{n}_pais": "ECUADOR" for n, _ in NIVELES}}
    return render_template("nuevo.html", perfil=perfil, p=p, v=valores, examinadores=examinadores,
                           n_exp=len(col.experiencias), n_cap=len(col.capacitaciones),
                           niveles=NIVELES, provincias=PROVINCIAS, estados=ESTADOS_CIVILES,
                           ruta_base=config.ruta_base(perfil))


# ---------------------------------------------------------------- candidatos

@app.route("/perfil/<perfil>")
def candidatos(perfil):
    perfil = perfil_o_404(perfil)
    todos, _ = leer(perfil)
    asignaciones = generar.leer_asignaciones()
    aj = ajustes.leer()
    q = (request.args.get("q") or "").strip().upper()
    lista = sorted(todos.values(), key=lambda c: numero(c["codigo"]), reverse=True)
    if q:
        lista = [c for c in lista if q in c["codigo"] or q in c["nombre"].upper() or q in c["cedula"]]
    lista = lista[:200]
    filas = []
    for c in lista:
        a = asignaciones.get(c["codigo"], {})
        exa = ajustes.examinador(aj, a.get("examinador") or aj["perfiles"][perfil]["examinador"])
        filas.append({**c, "estado_form": estado_formularios(perfil, c["codigo"]),
                      "examinador": exa["nombre"] if exa else "—"})
    return render_template("candidatos.html", perfil=perfil, filas=filas, q=q, total=len(todos))


@app.route("/perfil/<perfil>/<codigo>", methods=["GET", "POST"])
def candidato(perfil, codigo):
    perfil = perfil_o_404(perfil)
    todos, _ = leer(perfil)
    if codigo not in todos:
        abort(404)
    aj = ajustes.leer()
    p = aj["perfiles"][perfil]
    if request.method == "POST":
        ucs = [u for u in request.form.getlist("ucs") if u]
        if not ucs:
            flash("Marque al menos una unidad de competencia.", "error")
        else:
            generar.guardar_asignacion(codigo, {"examinador": request.form.get("examinador"), "ucs": ucs,
                                                "lugar": request.form.get("lugar")})
            try:
                _, avisos = generar.generar_uno(perfil, codigo, todos)
                registro.info(f"GENERAR {_quien()} {codigo} examinador={request.form.get('examinador')} ucs={ucs}")
                flash("Formularios generados.", "ok")
                for a in avisos:
                    flash(a, "aviso")
            except Exception as e:
                registro.error(f"GENERAR {_quien()} {codigo} FALLÓ {e!r}\n{traceback.format_exc()}")
                flash(f"No se pudieron generar: {e}", "error")
        return redirect(url_for("candidato", perfil=perfil, codigo=codigo))
    c = todos[codigo]
    a = generar.leer_asignaciones().get(codigo, {})
    exa, ucs, lugar = generar.eleccion(c, perfil, aj, generar.leer_asignaciones())
    carpeta = os.path.join(generar.CARPETA_SALIDA, perfil, codigo, "FORMULARIOS")
    archivos = sorted(os.listdir(carpeta)) if os.path.isdir(carpeta) else []
    return render_template("candidato.html", perfil=perfil, c=c, p=p, archivos=archivos,
                           examinadores=ajustes.examinadores_de(aj, perfil), exa=exa, ucs=ucs, lugar=lugar,
                           asignado=bool(a), estado_form=estado_formularios(perfil, codigo),
                           c03=generar.necesita_c03(c), educacion=generar.educacion_legible(c))


@app.route("/descargar/<perfil>/<codigo>.zip")
def descargar_zip(perfil, codigo):
    perfil = perfil_o_404(perfil)
    carpeta = os.path.join(generar.CARPETA_SALIDA, perfil, codigo, "FORMULARIOS")
    if not os.path.isdir(carpeta):
        abort(404)
    memoria = io.BytesIO()
    with zipfile.ZipFile(memoria, "w", zipfile.ZIP_DEFLATED) as z:
        for a in sorted(os.listdir(carpeta)):
            z.write(os.path.join(carpeta, a), f"{codigo}/FORMULARIOS/{a}")
    memoria.seek(0)
    return send_file(memoria, as_attachment=True, download_name=f"{codigo}_FORMULARIOS.zip", mimetype="application/zip")


@app.route("/descargar/<perfil>/<codigo>/<path:archivo>")
def descargar_archivo(perfil, codigo, archivo):
    perfil = perfil_o_404(perfil)
    carpeta = os.path.join(generar.CARPETA_SALIDA, perfil, codigo, "FORMULARIOS")
    ruta = os.path.realpath(os.path.join(carpeta, archivo))
    if not ruta.startswith(os.path.realpath(carpeta) + os.sep) or not os.path.isfile(ruta):
        abort(404)
    return send_file(ruta, as_attachment=True)


# ---------------------------------------------------------------- subir Excel actualizado

CARPETA_SUBIDAS = os.path.join(CARPETA_PORTAL, "subidas")
CARPETA_LOTES = os.path.join(generar.CARPETA_SALIDA, "_lotes")


@app.route("/perfil/<perfil>/subir", methods=["GET", "POST"])
def subir(perfil):
    """Paso 1: se sube el Excel y se muestra qué pasaría. Paso 2 (confirmar): se reemplaza y se genera."""
    perfil = perfil_o_404(perfil)
    actuales, _ = leer(perfil)
    estado = generar.leer_estado()
    ultimo = estado.get(perfil) or generar.ultimo_inicial(perfil, actuales)
    if request.method == "GET":
        return render_template("subir.html", perfil=perfil, ultimo=ultimo, vista=None,
                               nombre_base=os.path.basename(config.ruta_base(perfil)))
    archivo = request.files.get("archivo")
    if not archivo or not archivo.filename.lower().endswith((".xlsx", ".xlsm")):
        flash("Suba el archivo Excel (.xlsx) de la base del perfil.", "error")
        return redirect(url_for("subir", perfil=perfil))
    os.makedirs(CARPETA_SUBIDAS, exist_ok=True)
    ficha = secrets.token_hex(8)
    ruta = os.path.join(CARPETA_SUBIDAS, f"{perfil}_{ficha}.xlsx")
    archivo.save(ruta)
    registro.info(f"SUBIDA {_quien()} {perfil} archivo={archivo.filename!r} ficha={ficha} tamaño={os.path.getsize(ruta)}")
    try:
        nuevos, _ = base.leer_base(ruta, config.PERFILES[perfil]["hoja"])
    except SystemExit as e:
        registro.warning(f"SUBIDA {_quien()} {perfil} rechazada: {e}")
        os.remove(ruta)
        flash(f"El archivo no tiene el formato de la base: {e}", "error")
        return redirect(url_for("subir", perfil=perfil))
    except Exception as e:
        registro.error(f"SUBIDA {_quien()} {perfil} no se pudo leer {e!r}\n{traceback.format_exc()}")
        os.remove(ruta)
        flash(f"No se pudo leer el archivo: {e}", "error")
        return redirect(url_for("subir", perfil=perfil))
    ajenos = [c for c in nuevos if not c.startswith(perfil + "-")]
    if not nuevos or len(ajenos) > len(nuevos) / 2:
        os.remove(ruta)
        flash(f"Ese Excel no parece ser la base de {perfil} (códigos: {', '.join(list(nuevos)[:3])}…).", "error")
        return redirect(url_for("subir", perfil=perfil))
    desde = numero(ultimo) + 1 if ultimo else 0
    por_generar = sorted((c for c in nuevos if c.startswith(perfil + "-") and numero(c) >= desde), key=numero)
    incompletos = [c for c in por_generar if not nuevos[c]["nombre"] or not nuevos[c]["cedula"]]
    desaparecen = sorted(set(actuales) - set(nuevos), key=numero)
    vista = {"ficha": ficha, "nombre": archivo.filename, "total": len(nuevos), "total_actual": len(actuales),
             "por_generar": [nuevos[c] for c in por_generar], "incompletos": incompletos,
             "desaparecen": desaparecen, "ultimo_nuevo": max(nuevos, key=numero)}
    return render_template("subir.html", perfil=perfil, ultimo=ultimo, vista=vista,
                           nombre_base=os.path.basename(config.ruta_base(perfil)))


@app.route("/perfil/<perfil>/subir/<ficha>/confirmar", methods=["POST"])
def confirmar_subida(perfil, ficha):
    perfil = perfil_o_404(perfil)
    if not re.fullmatch(r"[0-9a-f]{16}", ficha):
        abort(404)
    ruta = os.path.join(CARPETA_SUBIDAS, f"{perfil}_{ficha}.xlsx")
    if not os.path.exists(ruta):
        flash("La subida ya no existe. Vuelva a subir el archivo.", "error")
        return redirect(url_for("subir", perfil=perfil))
    destino = config.ruta_base(perfil)
    if base_escritura.archivo_abierto(destino):
        flash("La base está abierta en Excel o LibreOffice. Ciérrela e intente otra vez.", "error")
        return redirect(url_for("subir", perfil=perfil))
    respaldo = base_escritura.respaldar(perfil, destino)
    registro.info(f"REEMPLAZO {_quien()} {perfil} base reemplazada, respaldo={os.path.basename(respaldo)}")
    temporal = destino + ".subiendo"
    shutil.copyfile(ruta, temporal)
    os.replace(temporal, destino)
    os.remove(ruta)
    _cache.pop(perfil, None)
    candidatos, _ = leer(perfil)
    estado = generar.leer_estado()
    ultimo = estado.get(perfil) or generar.ultimo_inicial(perfil, candidatos)
    desde = numero(ultimo) + 1 if ultimo else 0
    codigos = [c for c in candidatos if c.startswith(perfil + "-") and numero(c) >= desde]
    try:
        generados, revisar = generar.generar_lote(perfil, codigos, candidatos)
    except Exception as e:
        registro.error(f"LOTE {_quien()} {perfil} FALLÓ {e!r}\n{traceback.format_exc()}")
        flash(f"La base se reemplazó (respaldo: {os.path.basename(respaldo)}), pero falló la generación: {e}", "error")
        return redirect(url_for("candidatos", perfil=perfil))
    os.makedirs(CARPETA_LOTES, exist_ok=True)
    lote = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    with open(os.path.join(CARPETA_LOTES, f"{perfil}_{lote}.json"), "w", encoding="utf-8") as f:
        json.dump({"generados": generados, "revisar": revisar, "respaldo": os.path.basename(respaldo)}, f, ensure_ascii=False)
    registro.info(f"LOTE {_quien()} {perfil} lote={lote} generados={len(generados)} "
                  f"({generados[0] if generados else '-'} a {generados[-1] if generados else '-'}) revisar={len(revisar)}")
    flash(f"Base de {perfil} reemplazada (respaldo guardado). Formularios generados: {len(generados)}.", "ok")
    return redirect(url_for("ver_lote", perfil=perfil, lote=lote))


def _leer_lote(perfil, lote):
    if not re.fullmatch(r"\d{8}-\d{6}", lote):
        abort(404)
    ruta = os.path.join(CARPETA_LOTES, f"{perfil}_{lote}.json")
    if not os.path.exists(ruta):
        abort(404)
    with open(ruta, encoding="utf-8") as f:
        return json.load(f)


@app.route("/perfil/<perfil>/lote/<lote>")
def ver_lote(perfil, lote):
    perfil = perfil_o_404(perfil)
    datos = _leer_lote(perfil, lote)
    candidatos, _ = leer(perfil)
    filas = [candidatos[c] for c in datos["generados"] if c in candidatos]
    return render_template("lote.html", perfil=perfil, lote=lote, filas=filas, revisar=datos["revisar"],
                           respaldo=datos["respaldo"])


@app.route("/descargar/<perfil>/lote/<lote>.zip")
def descargar_lote(perfil, lote):
    perfil = perfil_o_404(perfil)
    datos = _leer_lote(perfil, lote)
    memoria = io.BytesIO()
    with zipfile.ZipFile(memoria, "w", zipfile.ZIP_DEFLATED) as z:
        for codigo in datos["generados"]:
            carpeta = os.path.join(generar.CARPETA_SALIDA, perfil, codigo, "FORMULARIOS")
            if os.path.isdir(carpeta):
                for a in sorted(os.listdir(carpeta)):
                    z.write(os.path.join(carpeta, a), f"{codigo}/FORMULARIOS/{a}")
        if datos["revisar"]:
            z.writestr("REVISAR.txt", "\n".join(f"{c}: {'; '.join(a)}" for c, a in datos["revisar"]) + "\n")
    memoria.seek(0)
    return send_file(memoria, as_attachment=True, download_name=f"{perfil}_{lote}_FORMULARIOS.zip",
                     mimetype="application/zip")


# ---------------------------------------------------------------- configuración

@app.route("/configuracion")
def configuracion():
    aj = ajustes.leer()
    return render_template("configuracion.html", aj=aj, perfiles=list(config.PERFILES))


@app.route("/configuracion/examinador", methods=["POST"])
def guardar_examinador():
    aj = ajustes.leer()
    f = request.form
    id_ = (f.get("id") or "").strip().upper()
    if not id_:
        usados = [int(m.group(1)) for e in aj["examinadores"] for m in [re.fullmatch(r"EXA(\d+)", e["id"])] if m]
        id_ = f"EXA{max(usados, default=0) + 1}"
    datos = {"id": id_, "nombre": re.sub(r"\s+", " ", f.get("nombre", "")).strip().upper(),
             "cedula": f.get("cedula", "").strip(), "celular": f.get("celular", "").strip(),
             "correo": f.get("correo", "").strip(), "trato": f.get("trato") or "el Sr.",
             "fecha_ingreso": f.get("fecha_ingreso") or "2023-03-06", "perfiles": f.getlist("perfiles")}
    if not datos["nombre"] or not datos["cedula"]:
        flash("El examinador necesita nombre y cédula.", "error")
        return redirect(url_for("configuracion"))
    existente = ajustes.examinador(aj, id_)
    if existente:
        existente.update(datos)
    else:
        aj["examinadores"].append(datos)
    ajustes.guardar(aj)
    registro.info(f"AJUSTES {_quien()} examinador {id_} guardado {datos}")
    flash(f"Examinador {id_} guardado.", "ok")
    return redirect(url_for("configuracion") + "#examinadores")


@app.route("/configuracion/examinador/<id_>/borrar", methods=["POST"])
def borrar_examinador(id_):
    aj = ajustes.leer()
    en_uso = [p for p, d in aj["perfiles"].items() if d["examinador"] == id_]
    if en_uso:
        flash(f"{id_} es el examinador predeterminado de {', '.join(en_uso)}. Cambie eso primero.", "error")
    else:
        aj["examinadores"] = [e for e in aj["examinadores"] if e["id"] != id_]
        ajustes.guardar(aj)
        flash(f"Examinador {id_} eliminado.", "ok")
    return redirect(url_for("configuracion") + "#examinadores")


@app.route("/configuracion/perfil/<perfil>", methods=["POST"])
def guardar_perfil(perfil):
    perfil = perfil_o_404(perfil)
    aj = ajustes.leer()
    p = aj["perfiles"][perfil]
    f = request.form
    ucs = [u.strip().upper() for u in re.split(r"[,\s]+", f.get("ucs", "")) if u.strip()]
    malos = [u for u in ucs if not re.fullmatch(r"UC[1-5]", u)]
    if not ucs or malos:
        flash(f"{perfil}: las unidades deben ser UC1 a UC5 separadas por coma (el C02 tiene 5 columnas).", "error")
        return redirect(url_for("configuracion") + f"#perfil-{perfil}")
    p["ucs"] = sorted(set(ucs), key=lambda u: int(u[2:]))
    exa = f.get("examinador")
    if exa and ajustes.examinador(aj, exa):
        p["examinador"] = exa
        e = ajustes.examinador(aj, exa)
        if perfil not in e["perfiles"]:
            e["perfiles"].append(perfil)
    costo = f.get("costo", "").strip()
    p["costo"] = float(costo) if re.fullmatch(r"\d+(\.\d+)?", costo) else p.get("costo", 10)
    # Lugares: filas lugar_<i>_campo
    lugares = []
    indices = sorted({int(m.group(1)) for k in f for m in [re.fullmatch(r"lugar_(\d+)_direccion", k)] if m})
    for i in indices:
        campos = {k: (f.get(f"lugar_{i}_{k}") or "").strip()
                  for k in ("id", "lugar", "tipo", "ciudad", "direccion", "sector", "telefono", "texto")}
        if f.get(f"lugar_{i}_borrar") == "on" or not campos["direccion"]:
            continue
        campos["id"] = campos["id"] or f"L{max([int(l['id'][1:]) for l in lugares + p['lugares'] if l.get('id', '')[1:].isdigit()] + [0]) + 1}"
        campos["texto"] = campos["texto"] or campos["direccion"]
        lugares.append(campos)
    if not lugares:
        flash(f"{perfil}: debe quedar al menos un lugar de examinación.", "error")
        return redirect(url_for("configuracion") + f"#perfil-{perfil}")
    p["lugares"] = lugares
    elegido = f.get("lugar_predeterminado")
    p["lugar"] = elegido if any(l["id"] == elegido for l in lugares) else lugares[0]["id"]
    ajustes.guardar(aj)
    registro.info(f"AJUSTES {_quien()} perfil {perfil} ucs={p['ucs']} examinador={p['examinador']} lugares={len(p['lugares'])}")
    flash(f"Perfil {perfil} guardado.", "ok")
    return redirect(url_for("configuracion") + f"#perfil-{perfil}")


def main():
    ap = argparse.ArgumentParser(description="Portal de registro de candidatos CAPTEEC OEC")
    ap.add_argument("--red", action="store_true", help="aceptar conexiones de otras PCs de la red local")
    ap.add_argument("--puerto", type=int, default=5000)
    ap.add_argument("--servidor", choices=["flask", "waitress"], default="flask",
                    help="waitress para el servidor (más estable que el de desarrollo de Flask)")
    args = ap.parse_args()
    host = "0.0.0.0" if args.red else "127.0.0.1"
    print(f"Portal en http://{'<IP-de-esta-PC>' if args.red else '127.0.0.1'}:{args.puerto}"
          f" · inicio de sesión {'obligatorio' if LOGIN_OBLIGATORIO else 'desactivado'}", flush=True)
    if LOGIN_OBLIGATORIO and not leer_usuarios():
        print("⚠️  No hay usuarios. Cree uno con: python portal/usuarios.py agregar <usuario>", flush=True)
    if args.servidor == "waitress":
        from waitress import serve
        # Los encabezados X-Forwarded-* de nginx los procesa ProxyFix (arriba)
        serve(app, host=host, port=args.puerto, threads=8)
    else:
        app.run(host=host, port=args.puerto, debug=False, threaded=True)


if __name__ == "__main__":
    main()
