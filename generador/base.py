"""
Lectura de las bases maestras de cada perfil
(0. Certificaciones Registradas/Version V03/<perfil>/<archivo>.xlsx).

Cada perfil tiene las columnas en letras distintas (unos tienen el teléfono de
la empresa, IE tiene dos bloques de capacitación, etc.), así que aquí no se usan
letras fijas: se recorre la fila de encabezados (fila 2) en orden y se
reconoce cada bloque por el nombre de sus encabezados.
"""

import datetime
import os
import re
import unicodedata
import warnings

from openpyxl import load_workbook

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

FILA_ENCABEZADOS = 2
FILA_INICIO_DATOS = 3
AMARILLO = "FFFFFF00"
PATRON_CODIGO = re.compile(r"[A-Z]+-\d+")


def normalizar(texto):
    """'Dirección ' -> 'DIRECCION'. Sin tildes, mayúsculas y espacios simples."""
    if texto is None:
        return ""
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", texto).strip().upper()


def texto(valor):
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return str(valor).strip()


def cedula_texto(valor):
    """Las cédulas guardadas como número pierden el 0 inicial."""
    if isinstance(valor, (int, float)):
        return str(int(valor)).zfill(10)
    return texto(valor)


def fecha(valor):
    """Devuelve datetime si el valor es una fecha (o texto tipo '24/07/2026'), si no None."""
    if isinstance(valor, datetime.datetime):
        return valor
    if isinstance(valor, datetime.date):
        return datetime.datetime(valor.year, valor.month, valor.day)
    if isinstance(valor, str):
        m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/?(\d{4})", valor.strip())
        if m:
            try:
                return datetime.datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except ValueError:
                return None
    return None


def hora(valor):
    if isinstance(valor, datetime.datetime):
        return valor.time()
    if isinstance(valor, datetime.time):
        return valor
    if isinstance(valor, str):
        m = re.fullmatch(r"(\d{1,2})[:hH](\d{2})(?::\d{2})?", valor.strip())
        if m:
            return datetime.time(int(m.group(1)), int(m.group(2)))
    return None


class Columnas:
    """Ubica cada dato en la fila de encabezados, bloque por bloque."""

    def __init__(self, encabezados):
        self.enc = [normalizar(e) for e in encabezados]
        self.mapa = {}
        self._analizar()

    def _buscar(self, nombres, desde=0, hasta=None):
        hasta = len(self.enc) if hasta is None else hasta
        nombres = [nombres] if isinstance(nombres, str) else nombres
        for i in range(desde, hasta):
            if self.enc[i] in nombres:
                return i
        return None

    def _todos(self, nombres, desde=0, hasta=None):
        hasta = len(self.enc) if hasta is None else hasta
        nombres = [nombres] if isinstance(nombres, str) else nombres
        return [i for i in range(desde, hasta) if self.enc[i] in nombres]

    def _analizar(self):
        b = self._buscar
        m = self.mapa
        m["codigo"] = b("CODIGO")
        if m["codigo"] is None:
            raise ValueError("No se encontró la columna CODIGO en la fila de encabezados")

        empresas = self._todos("EMPRESA")
        cursos = self._todos("CURSO")
        costo = b("COSTO")
        if costo is None:
            raise ValueError("No se encontró la columna COSTO")
        fin_personal = empresas[0] if empresas else (cursos[0] if cursos else costo)

        # Datos personales (antes de la primera EMPRESA)
        for clave, nombre in [("nombre", "APELLIDOS Y NOMBRES"), ("cedula", "CEDULA"), ("cod_exam", "COD. EXAM"),
                              ("f_naci", "F. NACI."), ("edad", "EDAD"), ("estado", "ESTADO"),
                              ("provincia", "PROVINCIA"), ("canton", "CANTON"), ("parroquia", "PARROQUIA"),
                              ("direccion", "DIRECCION"), ("correo", "CORREO"), ("telefono", "TELEFONO"),
                              ("celular", "CELULAR")]:
            m[clave] = b(nombre, 0, fin_personal)

        # Experiencias laborales: cada bloque empieza en EMPRESA y llega hasta
        # la siguiente EMPRESA o el primer CURSO posterior.
        limites = sorted(set(empresas + cursos + [costo]))
        self.experiencias = []
        for e in empresas:
            fin = min(x for x in limites if x > e)
            self.experiencias.append({
                "empresa": e,
                "direccion": b("DIRECCION", e, fin),
                "telefono": b(["TELEFONO", "CELULAR"], e, fin),
                "desde": b("F. INICIO", e, fin),
                "hasta": b("F. FIN", e, fin),
                "actividad": b("ACTIVIDAD", e, fin),
                "mes": b("MES", e, fin),
            })

        # Capacitaciones (IE tiene dos bloques CURSO)
        self.capacitaciones = []
        for c in cursos:
            fin = min([x for x in limites if x > c] + [costo])
            self.capacitaciones.append({
                "curso": c,
                "institucion": b("INSTITUCION", c, fin),
                "fechas": b("FECHAS.CURSO", c, fin),
                "horas": b("HORAS.CURSO", c, fin),
            })

        # Educación: bloques INSITUCION/PAIS/CIUDAD/TITULO entre la capacitación y COSTO.
        # 4 bloques = primaria (estudio CAPTEEC), secundaria, superior, cuarto nivel.
        # 3 bloques = secundaria, superior, cuarto nivel.
        inicio_edu = (cursos[-1] if cursos else fin_personal)
        bloques = self._todos(["INSITUCION", "INSTITUCION"], inicio_edu + 1, costo)
        bloques = [i for i in bloques if self.enc[i + 1:i + 2] == ["PAIS"]]
        niveles = ["primaria", "secundaria", "superior", "cuarto"][-len(bloques):] if bloques else []
        if len(bloques) > 4:
            raise ValueError(f"Se encontraron {len(bloques)} bloques de educación, se esperaban 3 o 4")
        self.educacion = {n: {"institucion": i, "pais": i + 1, "ciudad": i + 2, "titulo": i + 3}
                          for n, i in zip(niveles, bloques)}
        # EDAD y ESTADO repetidos entre la capacitación y la educación
        fin_cap = bloques[0] if bloques else costo
        m["edad2"] = b("EDAD", inicio_edu + 1, fin_cap)
        m["estado2"] = b("ESTADO", inicio_edu + 1, fin_cap)

        # Fechas, lugar y hora del examen (después de COSTO)
        m["costo"] = costo
        m["fecha_reg"] = b("FECHA REG", costo)
        m["fecha_acep"] = b("FECHA ACEP", costo)
        m["fecha_exam"] = b(["FECHA EXAMINACION", "FECHA EXAM"], costo)
        m["fecha_cert"] = b(["FECHA VISUALIZACION CERTIFICADO", "FECHA CERTIFICADO"], costo)
        m["hora"] = b("HORA", costo)
        # Las 4 columnas entre la fecha de certificado y la hora:
        # LUGAR (dentro/fuera), TIPO/INFRAESTRUCTURA, CIUDAD, DIRECCION/LUGAR
        ini = (m["fecha_cert"] or m["fecha_exam"] or costo) + 1
        fin = m["hora"] if m["hora"] is not None else ini + 4
        lugar = list(range(ini, fin))
        for clave, i in zip(["lugar", "tipo_lugar", "ciudad_examen", "direccion_examen"], lugar):
            m[clave] = i
        m["examen"] = b("EXAMEN", costo)
        m["cedula2"] = b("CEDULA", m["hora"]) if m["hora"] is not None else None


def _es_amarilla(celda):
    relleno = celda.fill
    return bool(relleno and relleno.fill_type) and relleno.fgColor.type == "rgb" and relleno.fgColor.rgb == AMARILLO


def buscar_hoja(wb, esperada):
    """Si la base nueva trae la hoja con otro nombre, usa la que tenga los
    encabezados CODIGO y APELLIDOS Y NOMBRES (prefiere la de la versión 3)."""
    candidatas = []
    for ws in wb.worksheets:
        enc = {normalizar(c.value) for c in ws[FILA_ENCABEZADOS]}
        if {"CODIGO", "APELLIDOS Y NOMBRES", "COSTO"} <= enc:
            candidatas.append(ws)
    if not candidatas:
        raise SystemExit(f"❌ La base no tiene la hoja '{esperada}' ni otra con los encabezados CODIGO / APELLIDOS Y NOMBRES")
    elegida = next((ws for ws in candidatas if "3" in ws.title), candidatas[0])
    print(f"⚠️  No está la hoja '{esperada}'; se usa '{elegida.title}'")
    return elegida


def leer_base(ruta, hoja):
    """Devuelve {codigo: candidato} con los datos ya limpios de la base maestra."""
    if not os.path.exists(ruta):
        raise SystemExit(f"❌ No existe la base: {ruta}")
    wb = load_workbook(ruta, data_only=True)
    ws = wb[hoja] if hoja in wb.sheetnames else buscar_hoja(wb, hoja)
    encabezados = [c.value for c in ws[FILA_ENCABEZADOS]]
    col = Columnas(encabezados)
    candidatos = {}
    for fila in ws.iter_rows(min_row=FILA_INICIO_DATOS):
        valores = [c.value for c in fila]
        v = lambda i: valores[i] if i is not None and i < len(valores) else None
        codigo = texto(v(col.mapa["codigo"]))
        if not PATRON_CODIGO.fullmatch(codigo):
            continue
        m = col.mapa
        cand = {
            "codigo": codigo,
            "marca": normalizar(valores[0]) if col.mapa["codigo"] > 0 else "",
            "amarilla": _es_amarilla(fila[m["codigo"]]),
            "fila": fila[0].row,
            "nombre": re.sub(r"\s+", " ", texto(v(m["nombre"]))),
            "cedula": cedula_texto(v(m["cedula"])),
            "cod_exam": texto(v(m["cod_exam"])),
            "f_naci": fecha(v(m["f_naci"])),
            "edad": texto(v(m["edad"])),
            "provincia": texto(v(m["provincia"])),
            "canton": texto(v(m["canton"])),
            "parroquia": texto(v(m["parroquia"])),
            "direccion": texto(v(m["direccion"])),
            "correo": texto(v(m["correo"])),
            "telefono": texto(v(m["telefono"])),
            "celular": texto(v(m["celular"])),
            "fecha_reg": fecha(v(m["fecha_reg"])),
            "fecha_acep": fecha(v(m["fecha_acep"])),
            "fecha_exam": fecha(v(m["fecha_exam"])),
            "fecha_reg_bruta": v(m["fecha_reg"]),
            "fecha_acep_bruta": v(m["fecha_acep"]),
            "fecha_exam_bruta": v(m["fecha_exam"]),
            "hora": hora(v(m["hora"])),
            "lugar": texto(v(m["lugar"])),
            "tipo_lugar": texto(v(m["tipo_lugar"])),
            "ciudad_examen": texto(v(m["ciudad_examen"])),
            "direccion_examen": texto(v(m["direccion_examen"])),
            "experiencias": [],
            "capacitaciones": [],
            "educacion": {},
        }
        for e in col.experiencias:
            exp = {k: v(i) for k, i in e.items()}
            exp = {"empresa": texto(exp["empresa"]), "direccion": texto(exp["direccion"]),
                   "telefono": texto(exp["telefono"]), "desde": fecha(exp["desde"]) or texto(exp["desde"]),
                   "hasta": fecha(exp["hasta"]) or texto(exp["hasta"]), "actividad": texto(exp["actividad"])}
            if exp["empresa"]:
                cand["experiencias"].append(exp)
        for c in col.capacitaciones:
            cap = {k: texto(v(i)) for k, i in c.items()}
            if cap["curso"]:
                cand["capacitaciones"].append(cap)
        for nivel, bloque in col.educacion.items():
            edu = {k: texto(v(i)) for k, i in bloque.items()}
            if edu["institucion"]:
                cand["educacion"][nivel] = edu
        candidatos[codigo] = cand
    return candidatos, col
