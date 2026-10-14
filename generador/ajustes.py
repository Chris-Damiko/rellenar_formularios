"""
Ajustes editables (examinadores, unidades de competencia y lugares de cada
perfil). Se guardan en ajustes.json y se cambian desde el portal web
(Configuración) o a mano.

    examinadores: lista de {id, nombre, cedula, celular, correo, trato,
                  fecha_ingreso (AAAA-MM-DD), perfiles: [PRC, ...]}
    perfiles:     {PRC: {esquema, ucs: [UC1, UC2], examinador: id por defecto,
                  lugares: [{id, lugar, tipo, ciudad, direccion, sector,
                  telefono, texto}], lugar: id por defecto, costo}}

La primera vez se llena solo con config.py y con lo que plantillas.py lee de
los formularios reales de cada muestra.
"""

import copy
import json
import os
import threading

import config

ARCHIVO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ajustes.json")
_candado = threading.Lock()

# Códigos de examinador del F01 (1. Doc Registros Calificacion/.../F001/F01.xls)
CODIGOS_EXA = {"ADRIAN": "EXA2", "BERNARDA": "EXA3", "ANA": "EXA4", "HECTOR": "EXA5",
               "MARLO": "EXA6", "BYRON": "EXA7", "MILTON": "EXA8"}
TRATO = {"BERNARDA": "la Sra.", "ANA": "la Sra."}


def _iniciales():
    examinadores = []
    for clave, e in config.EXAMINADORES.items():
        examinadores.append({
            "id": CODIGOS_EXA[clave], "nombre": e["nombre"], "cedula": e["cedula"], "celular": e["celular"],
            "correo": "", "trato": TRATO.get(clave, "el Sr."), "fecha_ingreso": "2023-03-06",
            "perfiles": [p for p, d in config.PERFILES.items() if d["examinador"] == clave],
        })
    perfiles = {p: {"examinador": CODIGOS_EXA[d["examinador"]], "esquema": "", "ucs": [], "lugares": [],
                    "lugar": "", "costo": 10}
                for p, d in config.PERFILES.items()}
    return {"examinadores": examinadores, "perfiles": perfiles}


def leer():
    with _candado:
        if not os.path.exists(ARCHIVO):
            datos = _iniciales()
            _escribir(datos)
            return datos
        with open(ARCHIVO, encoding="utf-8") as f:
            return json.load(f)


def _escribir(datos):
    tmp = ARCHIVO + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(datos, f, indent=2, ensure_ascii=False)
    os.replace(tmp, ARCHIVO)


def guardar(datos):
    with _candado:
        _escribir(datos)


def examinador(datos, id_):
    return next((e for e in datos["examinadores"] if e["id"] == id_), None)


def examinadores_de(datos, perfil):
    return [e for e in datos["examinadores"] if perfil in e["perfiles"]]


def lugar(datos, perfil, id_=None):
    p = datos["perfiles"][perfil]
    id_ = id_ or p.get("lugar")
    return next((l for l in p["lugares"] if l["id"] == id_), p["lugares"][0] if p["lugares"] else None)


def completar_perfil(perfil, esquema, ucs, lugar_muestra):
    """Lo llama plantillas.py con lo leído de la muestra, solo si el perfil aún no lo tiene."""
    datos = leer()
    p = datos["perfiles"].setdefault(perfil, copy.deepcopy(_iniciales()["perfiles"][perfil]))
    p["esquema"] = p.get("esquema") or esquema
    p["ucs"] = p.get("ucs") or ucs
    if not p.get("lugares"):
        lugar_muestra["id"] = "L1"
        p["lugares"] = [lugar_muestra]
        p["lugar"] = "L1"
    guardar(datos)
    return p
