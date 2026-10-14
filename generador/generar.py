"""
Genera los formularios de cada candidato a partir de la base maestra del
perfil y de las plantillas armadas con plantillas.py.

Sigue desde donde se quedó: estado.json guarda el último código generado de
cada perfil. Cuando llega una base actualizada, se reemplaza el archivo del
perfil (ver la ruta en config.py o en la salida del script) y se vuelve a
correr: solo se generan los códigos posteriores al último.

Salida: salida/<PERFIL>/<CODIGO>/FORMULARIOS/  (misma estructura que en el OEC)
y salida/<PERFIL>/REVISAR.txt con lo que hay que mirar a mano en esa corrida.

USO:
    python generar.py PRC                      # sigue desde el último (p. ej. PRC-2101 si quedó en PRC-2100)
    python generar.py                          # todos los perfiles, cada uno desde su último
    python generar.py PRC --estado             # solo muestra dónde quedó cada perfil
    python generar.py PRC --desde 2101         # empieza en PRC-2101 y sigue hasta el final (y guarda el estado)
    python generar.py PRC --desde 2101 --hasta 2150
    python generar.py PRC --codigos PRC-2105,PRC-2110   # rehacer sueltos (no cambia el estado)
    python generar.py PRC --pendientes         # todos los que no tienen formularios en el OEC
"""

import argparse
import datetime
import glob
import json
import os
import re
import shutil
import sys
from concurrent.futures import ProcessPoolExecutor

import ajustes
import base
import config
import docx_texto as dt
import formatos
import xlsx_texto
from plantillas import CARPETA_PLANTILLAS

CARPETA_SALIDA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "salida")
CAPTEEC = "CAPTEEC CENTRO DE CAPACITACIÓN TÉCNICA - EMPRESARIAL & TUTORIAS S.A.S."
# Fila "Lecto-escritura (C03)" del C02 para quien rinde el C03
ESTUDIO_LECTO = {"institucion": CAPTEEC, "pais": "ECUADOR", "ciudad": "LA TRONCAL", "titulo": "LECTO-ESCRITURA"}
MARCA = re.compile(r"\{\{([A-Za-z0-9_]+)\}\}")


def cedula_valida(cedula):
    """Cédula ecuatoriana: 10 dígitos, provincia válida y dígito verificador (módulo 10)."""
    if not re.fullmatch(r"\d{10}", cedula):
        return False
    provincia = int(cedula[:2])
    if not (1 <= provincia <= 24 or provincia == 30) or int(cedula[2]) >= 6:
        return False
    suma = 0
    for i, digito in enumerate(cedula[:9]):
        v = int(digito) * (2 if i % 2 == 0 else 1)
        suma += v - 9 if v > 9 else v
    return (10 - suma % 10) % 10 == int(cedula[9])


def necesita_c03(cand):
    """Solo primaria (marca PRI en la columna A) -> evaluación de lecto-escritura C03."""
    if cand["marca"].startswith("PRI"):
        return True
    edu = cand["educacion"]
    return "primaria" in edu and not any(n in edu for n in ("secundaria", "superior", "cuarto"))


ARCHIVO_ASIGNACIONES = os.path.join(os.path.dirname(os.path.abspath(__file__)), "asignaciones.json")


def leer_asignaciones():
    """Examinador, UC y lugar elegidos en el portal para cada candidato (código -> datos)."""
    try:
        with open(ARCHIVO_ASIGNACIONES, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}


def guardar_asignacion(codigo, datos):
    todas = leer_asignaciones()
    todas[codigo] = datos
    tmp = ARCHIVO_ASIGNACIONES + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(todas, f, indent=2, ensure_ascii=False)
    os.replace(tmp, ARCHIVO_ASIGNACIONES)


def eleccion(cand, perfil, aj, asignaciones):
    """Examinador, unidades de competencia y lugar del candidato: lo elegido en el
    portal o, si no hay, lo predeterminado del perfil (el lugar se busca primero
    por la dirección que trae la base)."""
    p = aj["perfiles"][perfil]
    a = asignaciones.get(cand["codigo"], {})
    exa = ajustes.examinador(aj, a.get("examinador") or p["examinador"]) or ajustes.examinador(aj, p["examinador"])
    ucs = a.get("ucs") or p["ucs"]
    lugar = None
    if a.get("lugar"):
        lugar = ajustes.lugar(aj, perfil, a["lugar"])
    elif cand["direccion_examen"]:
        d = base.normalizar(cand["direccion_examen"])
        lugar = next((l for l in p["lugares"] if d in (base.normalizar(l["direccion"]), base.normalizar(l["texto"]))), None)
    return exa, ucs, lugar or ajustes.lugar(aj, perfil)


def datos_planos(cand, perfil, aj=None, asignaciones=None):
    """Valores base de cada campo. Las fechas quedan como datetime (para Excel)."""
    aj = aj or ajustes.leer()
    asignaciones = leer_asignaciones() if asignaciones is None else asignaciones
    examinador, ucs, lugar = eleccion(cand, perfil, aj, asignaciones)
    ingreso = datetime.datetime.strptime(examinador.get("fecha_ingreso") or "2023-03-06", "%Y-%m-%d")
    edu = dict(cand["educacion"])
    # C02, fila "Lecto-escritura (C03) cuando aplique": el estudio en CAPTEEC de
    # quien rinde el C03. En la base va en el bloque de primaria; no es la primaria,
    # así que esa fila solo se llena si la base trae una escuela de verdad.
    lecto = {}
    if "CAPTEEC" in base.normalizar(edu.get("primaria", {}).get("institucion")):
        lecto = edu.pop("primaria")
    if necesita_c03(cand):
        edu["lecto"] = {**ESTUDIO_LECTO, **{k: v for k, v in lecto.items() if v}}
    edad = int(cand["edad"]) if cand["edad"].isdigit() else cand["edad"]
    d = {
        "CODIGO": cand["codigo"], "NOMBRE": cand["nombre"], "CEDULA": cand["cedula"],
        "CORREO": cand["correo"], "EDAD": edad, "PROVINCIA": cand["provincia"], "CANTON": cand["canton"],
        "PARROQUIA": cand["parroquia"], "DIRECCION": cand["direccion"], "TELEFONO": cand["telefono"],
        "CELULAR": cand["celular"], "FECHA_REG": cand["fecha_reg"], "FECHA_ACEP": cand["fecha_acep"],
        "FECHA_EXAM": cand["fecha_exam"], "HORA": cand["hora"],
        "EXAMINADOR_ID": examinador["id"], "EXAMINADOR_NOMBRE": examinador["nombre"],
        "EXAMINADOR_CEDULA": examinador["cedula"], "EXAMINADOR_CELULAR": examinador["celular"],
        "EXAMINADOR_TRATO": examinador.get("trato") or "el Sr.",
        "EXA_DIA": str(ingreso.day), "EXA_MES": formatos.MESES[ingreso.month - 1].lower(), "EXA_ANIO": str(ingreso.year),
        "UC": ", ".join(ucs),
        **{f"UC{i}_X": ("X" if f"UC{i}" in ucs else "") for i in range(1, 6)},
        "LUGAR_TEXTO": lugar["texto"], "LUGAR_OEC": lugar["lugar"], "LUGAR_DIRECCION": lugar["direccion"],
        "LUGAR_SECTOR": lugar["sector"], "LUGAR_TELEFONO": lugar["telefono"],
    }
    for nivel in ("lecto", "primaria", "secundaria", "superior", "cuarto"):
        e = edu.get(nivel, {})
        for k in ("institucion", "pais", "ciudad", "titulo"):
            d[f"EDU_{nivel.upper()}_{k.upper()}"] = e.get(k, "")
    for i in (1, 2):
        c = cand["capacitaciones"][i - 1] if len(cand["capacitaciones"]) >= i else {}
        for k in ("curso", "institucion", "fechas", "horas"):
            v = c.get(k, "")
            d[f"CAP{i}_{k.upper()}"] = int(v) if k == "horas" and v.isdigit() else v
        e = cand["experiencias"][i - 1] if len(cand["experiencias"]) >= i else {}
        d[f"EXP{i}_DESDE"] = e.get("desde", "")
        d[f"EXP{i}_HASTA"] = e.get("hasta", "")
        d[f"EXP{i}_EMPRESA"] = e.get("empresa", "")
        d[f"EXP{i}_DIRECCION"] = e.get("direccion", "")
        d[f"EXP{i}_TELEFONO"] = e.get("telefono", "")
        d[f"EXP{i}_FUNCIONES"] = e.get("actividad", "")
    return d


def resolver(clave, datos):
    """'FECHA_EXAM__LARGA_DEL' -> '25 de Junio del 2026'; 'NOMBRE__TITULO' -> 'Lopez ...'."""
    nombre, _, formato = clave.partition("__")
    if nombre not in datos:
        raise KeyError(clave)
    v = datos[nombre]
    if nombre.startswith("FECHA_") and formato:
        return formatos.escribir_fecha(v, formato)
    if nombre == "HORA":
        return formatos.escribir_hora(v, formato or "H")
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        return f"{v.day:02d}/{v.month:02d}/{v.year}"
    v = str(v)
    return {"MAY": v.upper(), "MIN": v.lower(), "TITULO": v.title()}.get(formato, v)


def rellenar_excel(origen, destino, datos):
    def tipado(clave):
        if clave not in datos:
            return resolver(clave, datos)  # con formato (__) o inexistente -> KeyError
        return datos[clave]
    return xlsx_texto.rellenar(origen, destino, lambda c: resolver(c, datos), tipado)


class Valores(dict):
    """Diccionario que resuelve las claves con formato al pedirlas."""

    def __init__(self, datos):
        super().__init__()
        self.datos = datos

    def __contains__(self, clave):
        try:
            resolver(clave, self.datos)
            return True
        except KeyError:
            return False

    def __getitem__(self, clave):
        return resolver(clave, self.datos)


def generar_candidato(trabajo):
    datos, carpeta, plantillas = trabajo
    if os.path.isdir(carpeta):
        shutil.rmtree(carpeta)
    os.makedirs(carpeta)
    faltan = set()
    for ruta in plantillas:
        nombre = os.path.basename(ruta)
        if nombre.startswith("T04"):
            nombre = f"T04_10 {datos['EXAMINADOR_ID']}.docx"
        destino = os.path.join(carpeta, nombre)
        if ruta.endswith(".docx"):
            faltan |= dt.rellenar(ruta, destino, Valores(datos))
        else:
            faltan |= rellenar_excel(ruta, destino, datos)
    return carpeta, faltan


def advertencias(cand, perfil, muestra):
    avisos = []
    if cand["amarilla"]:
        avisos.append("fila amarilla: se generó C13; llenar a mano motivo, fecha del registro y nueva fecha")
    for campo in ("fecha_reg", "fecha_acep", "fecha_exam"):
        if cand[campo] is None:
            bruto = cand[campo + "_bruta"]
            avisos.append(f"{campo} vacía o inválida ({bruto!r})" if bruto not in (None, "") else f"{campo} vacía")
    if cand["hora"] is None:
        avisos.append("hora vacía")
    if not cedula_valida(cand["cedula"]):
        avisos.append(f"cédula no pasa el dígito verificador ({cand['cedula']})")
    for campo, nombre in (("correo", "correo"), ("celular", "celular"), ("direccion", "dirección de domicilio")):
        if not cand[campo]:
            avisos.append(f"sin {nombre}")
    if not cand["experiencias"]:
        avisos.append("sin experiencia laboral en la base")
    if not cand["capacitaciones"]:
        avisos.append("sin capacitación en la base")
    if not cand["educacion"] and not cand["marca"].startswith("PRI"):
        avisos.append("sin educación")
    lugares = ajustes.leer()["perfiles"][perfil]["lugares"]
    conocidos = {base.normalizar(l[k]) for l in lugares for k in ("direccion", "texto")}
    lugar_c = base.normalizar(cand["direccion_examen"])
    if lugar_c and lugar_c not in conocidos:
        avisos.append(f"lugar de examen de la base no está en los lugares del perfil: '{cand['direccion_examen']}' "
                      f"(se usó el lugar predeterminado; agréguelo en Configuración si es nuevo)")
    return avisos


ARCHIVO_ESTADO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "estado.json")


def leer_estado():
    try:
        with open(ARCHIVO_ESTADO, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}


def escribir_estado(estado):
    with open(ARCHIVO_ESTADO, "w", encoding="utf-8") as f:
        json.dump(dict(sorted(estado.items())), f, indent=2, ensure_ascii=False)


def ultimo_inicial(perfil, candidatos):
    """Si el perfil aún no está en estado.json: el mayor entre lo último generado
    en salida/ y el último candidato que ya tiene formularios en el OEC."""
    numero = lambda c: int(c.split("-")[1])
    salida = os.path.join(CARPETA_SALIDA, perfil)
    hechos = [c for c in (os.listdir(salida) if os.path.isdir(salida) else []) if base.PATRON_CODIGO.fullmatch(c)]
    hechos += [c for c in candidatos if tiene_formularios(perfil, c)]
    return max(hechos, key=numero) if hechos else None


ARCHIVO_INDICE_OEC = os.path.join(os.path.dirname(os.path.abspath(__file__)), "indice_oec.json")
_indice_oec = None


def indice_oec():
    """Códigos que ya tienen formularios en el OEC. Se usa donde la carpeta del OEC
    no está (el servidor): crear con  python generar.py --indice-oec"""
    global _indice_oec
    if _indice_oec is None:
        try:
            with open(ARCHIVO_INDICE_OEC, encoding="utf-8") as f:
                _indice_oec = {p: set(c) for p, c in json.load(f).items()}
        except FileNotFoundError:
            _indice_oec = {}
    return _indice_oec


_con_carpetas = {}


def oec_completo(perfil):
    """True si aquí están las carpetas de los candidatos (la copia completa del OEC)."""
    if perfil not in _con_carpetas:
        carpeta = os.path.join(config.CARPETA_V03, config.PERFILES[perfil]["carpeta"])
        _con_carpetas[perfil] = os.path.isdir(carpeta) and any(
            base.PATRON_CODIGO.fullmatch(c) and os.path.isdir(os.path.join(carpeta, c)) for c in os.listdir(carpeta))
    return _con_carpetas[perfil]


def tiene_formularios(perfil, codigo):
    if not oec_completo(perfil):
        return codigo in indice_oec().get(perfil, set())
    f = config.carpeta_formularios(perfil, codigo)
    return bool(f and os.listdir(f))


def crear_indice_oec():
    indice = {}
    for p, d in config.PERFILES.items():
        carpeta = os.path.join(config.CARPETA_V03, d["carpeta"])
        indice[p] = sorted(c for c in os.listdir(carpeta) if base.PATRON_CODIGO.fullmatch(c) and tiene_formularios(p, c))
    with open(ARCHIVO_INDICE_OEC, "w", encoding="utf-8") as f:
        json.dump(indice, f, indent=1)
    print(f"Índice guardado en {ARCHIVO_INDICE_OEC}: " + ", ".join(f"{p} {len(c)}" for p, c in indice.items()))


def plantillas_de(perfil):
    """(comunes, c03, c13): rutas de las plantillas del perfil."""
    carpeta_plantillas = os.path.join(CARPETA_PLANTILLAS, perfil)
    todas = sorted(glob.glob(os.path.join(carpeta_plantillas, "*.docx")) +
                   glob.glob(os.path.join(carpeta_plantillas, "*.xlsx")))
    c03 = [p for p in todas if os.path.basename(p).startswith("C03")]
    c13 = [p for p in todas if os.path.basename(p).startswith("C13")]
    return [p for p in todas if p not in c03 and p not in c13], c03, c13


def generar_uno(perfil, codigo, candidatos=None):
    """Genera los formularios de un candidato (lo usa el portal web).
    Devuelve (carpeta, avisos)."""
    if candidatos is None:
        candidatos, _ = base.leer_base(config.ruta_base(perfil), config.PERFILES[perfil]["hoja"])
    cand = candidatos[codigo]
    comunes, c03, c13 = plantillas_de(perfil)
    if not comunes:
        raise RuntimeError(f"No hay plantillas de {perfil}. Corre: python plantillas.py {perfil}")
    plantillas = comunes + (c03 if necesita_c03(cand) else []) + (c13 if cand["amarilla"] else [])
    muestra = candidatos.get(config.PERFILES[perfil]["muestra"], cand)
    avisos = advertencias(cand, perfil, muestra)
    carpeta = os.path.join(CARPETA_SALIDA, perfil, codigo, "FORMULARIOS")
    _, faltan = generar_candidato((datos_planos(cand, perfil), carpeta, plantillas))
    if faltan:
        avisos.append(f"marcas sin dato: {', '.join(sorted(faltan))}")
    # Si es el más nuevo del perfil, el estado avanza hasta él
    estado = leer_estado()
    ultimo = estado.get(perfil)
    if not ultimo or int(codigo.split("-")[1]) > int(ultimo.split("-")[1]):
        estado[perfil] = codigo
        escribir_estado(estado)
    return carpeta, avisos


def generar_lote(perfil, codigos, candidatos, guardar_estado=True):
    """Genera varios candidatos en paralelo (lo usa el portal al subir un Excel).
    Devuelve (generados, revisar) con revisar = [(código, [avisos])].
    Si guardar_estado, el estado avanza como en la consola: hasta antes de la
    primera fila sin nombre o cédula."""
    comunes, c03, c13 = plantillas_de(perfil)
    aj, asignaciones = ajustes.leer(), leer_asignaciones()
    muestra = candidatos.get(config.PERFILES[perfil]["muestra"])
    numero = lambda c: int(c.split("-")[1])
    codigos = sorted(codigos, key=numero)
    trabajos, revisar, generados, primero_incompleto = [], [], [], None
    for codigo in codigos:
        cand = candidatos[codigo]
        if not cand["nombre"] or not cand["cedula"]:
            revisar.append((codigo, ["SIN NOMBRE O CÉDULA: no se generó (se intentará la próxima vez)"]))
            primero_incompleto = primero_incompleto or codigo
            continue
        plantillas = comunes + (c03 if necesita_c03(cand) else []) + (c13 if cand["amarilla"] else [])
        avisos = advertencias(cand, perfil, muestra or cand)
        if avisos:
            revisar.append((codigo, avisos))
        carpeta = os.path.join(CARPETA_SALIDA, perfil, codigo, "FORMULARIOS")
        trabajos.append((datos_planos(cand, perfil, aj, asignaciones), carpeta, plantillas))
    with ProcessPoolExecutor(max_workers=4) as procesos:
        for carpeta, faltan in procesos.map(generar_candidato, trabajos):
            codigo = os.path.basename(os.path.dirname(carpeta))
            generados.append(codigo)
            if faltan:
                revisar.append((codigo, [f"marcas sin dato: {', '.join(sorted(faltan))}"]))
    if guardar_estado and generados:
        estado = leer_estado()
        anteriores = [c for c in codigos if not primero_incompleto or numero(c) < numero(primero_incompleto)]
        nuevo = anteriores[-1] if anteriores else None
        ultimo = estado.get(perfil)
        if nuevo and (not ultimo or numero(nuevo) > numero(ultimo)):
            estado[perfil] = nuevo
            escribir_estado(estado)
    return generados, sorted(revisar)


def generar_perfil(perfil, args):
    datos_perfil = config.PERFILES[perfil]
    carpeta_plantillas = os.path.join(CARPETA_PLANTILLAS, perfil)
    if not os.path.isdir(carpeta_plantillas):
        print(f"❌ No hay plantillas de {perfil}. Corre primero: python plantillas.py {perfil}")
        return
    todas = sorted(glob.glob(os.path.join(carpeta_plantillas, "*.docx")) +
                   glob.glob(os.path.join(carpeta_plantillas, "*.xlsx")))
    c03 = [p for p in todas if os.path.basename(p).startswith("C03")]
    c13 = [p for p in todas if os.path.basename(p).startswith("C13")]
    comunes = [p for p in todas if p not in c03 and p not in c13]

    candidatos, _ = base.leer_base(config.ruta_base(perfil), datos_perfil["hoja"])
    muestra = candidatos[datos_perfil["muestra"]]
    aj, asignaciones = ajustes.leer(), leer_asignaciones()

    def numero(c):
        return int(c.split("-")[1])

    estado = leer_estado()
    ultimo = estado.get(perfil) or ultimo_inicial(perfil, candidatos)
    if perfil not in estado and ultimo:
        estado[perfil] = ultimo
        escribir_estado(estado)
    guardar = False
    if args.codigos:
        elegidos = [c.strip() for c in args.codigos.split(",")]
        no_estan = [c for c in elegidos if c not in candidatos]
        if no_estan:
            print(f"⚠️  No están en la base: {', '.join(no_estan)}")
        elegidos = [c for c in elegidos if c in candidatos]
    elif args.pendientes:
        elegidos = [c for c in candidatos if not tiene_formularios(perfil, c)]
    else:
        desde = args.desde if args.desde is not None else (numero(ultimo) + 1 if ultimo else 0)
        elegidos = [c for c in candidatos
                    if numero(c) >= desde and (args.hasta is None or numero(c) <= args.hasta)]
        guardar = True
    elegidos.sort(key=numero)

    print(f"\n===== {perfil} =====")
    print(f"Base: {config.ruta_base(perfil)}")
    print(f"Último generado: {ultimo or '(ninguno)'}")
    if not elegidos:
        print("Nada nuevo en la base.")
        return
    print(f"Se generan {len(elegidos)}: {elegidos[0]} a {elegidos[-1]}")
    salida = os.path.join(CARPETA_SALIDA, perfil)
    os.makedirs(salida, exist_ok=True)

    trabajos, revisar = [], []
    primero_incompleto = None
    for codigo in elegidos:
        cand = candidatos[codigo]
        if not cand["nombre"] or not cand["cedula"]:
            revisar.append((codigo, ["SIN NOMBRE O CÉDULA: no se generó (se intentará otra vez la próxima corrida)"]))
            primero_incompleto = primero_incompleto or codigo
            continue
        # C13 (ausencia / cambio de fecha) solo para las filas amarillas de la base
        plantillas = comunes + (c03 if necesita_c03(cand) else []) + (c13 if cand["amarilla"] else [])
        avisos = advertencias(cand, perfil, muestra)
        if necesita_c03(cand) and not c03:
            avisos.append("necesita C03 pero el perfil no tiene plantilla C03")
        if avisos:
            revisar.append((codigo, avisos))
        carpeta = os.path.join(salida, codigo, "FORMULARIOS")
        trabajos.append((datos_planos(cand, perfil, aj, asignaciones), carpeta, plantillas))

    with ProcessPoolExecutor() as procesos:
        for carpeta, faltan in procesos.map(generar_candidato, trabajos):
            codigo = os.path.basename(os.path.dirname(carpeta))
            if faltan:
                revisar.append((codigo, [f"marcas sin dato: {', '.join(sorted(faltan))}"]))
            print(f"✅ {codigo}")

    with open(os.path.join(salida, "REVISAR.txt"), "w", encoding="utf-8") as f:
        f.write(f"Perfil {perfil}. Generados {len(trabajos)} de {len(elegidos)} candidatos "
                f"({elegidos[0]} a {elegidos[-1]}).\n\n")
        for codigo, avisos in sorted(revisar):
            f.write(f"{codigo}: {'; '.join(avisos)}\n")
    print(f"Listo: {len(trabajos)} candidatos en {salida}/")
    if revisar:
        print(f"⚠️  {len(revisar)} con datos a revisar: {salida}/REVISAR.txt")
    if guardar and trabajos:
        # Si hay una fila sin nombre/cédula, el estado se queda justo antes de ella
        # para que la próxima corrida la vuelva a intentar.
        if primero_incompleto:
            anteriores = [c for c in elegidos if numero(c) < numero(primero_incompleto)]
            nuevo = anteriores[-1] if anteriores else ultimo
        else:
            nuevo = elegidos[-1]
        if nuevo and (not ultimo or numero(nuevo) > numero(ultimo) or args.desde is not None):
            estado[perfil] = nuevo
            escribir_estado(estado)
            print(f"📌 Estado guardado: {perfil} quedó en {nuevo}. La próxima corrida empieza en el siguiente.")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("perfiles", nargs="*", help=f"perfiles ({', '.join(config.PERFILES)}); vacío = todos")
    ap.add_argument("--desde", type=int)
    ap.add_argument("--hasta", type=int)
    ap.add_argument("--codigos")
    ap.add_argument("--pendientes", action="store_true", help="todos los que no tienen formularios en el OEC")
    ap.add_argument("--estado", action="store_true", help="solo mostrar dónde quedó cada perfil")
    ap.add_argument("--indice-oec", action="store_true",
                    help="guardar en indice_oec.json qué candidatos ya tienen formularios (para el servidor)")
    args = ap.parse_args()
    if args.indice_oec:
        crear_indice_oec()
        return
    perfiles = [p.upper() for p in args.perfiles] or list(config.PERFILES)
    if args.estado:
        estado = leer_estado()
        for p in perfiles:
            print(f"{p}: quedó en {estado.get(p, '(sin estado aún: se fija en la primera corrida)')}")
        return
    for p in perfiles:
        if p not in config.PERFILES:
            print(f"❌ Perfil desconocido: {p}")
            continue
        generar_perfil(p, args)


if __name__ == "__main__":
    sys.exit(main())
