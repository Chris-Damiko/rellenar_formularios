"""
Formatos de fecha y hora que aparecen en los formularios, y cómo reconocerlos.

Una marca de fecha en una plantilla se ve así: {{FECHA_EXAM__LARGA_DEL}}.
La parte después de "__" dice cómo escribir la fecha:

    CORTA            25/06/2026
    CORTA_SIN_CEROS  25/6/2026
    LARGA            25 de Junio de 2026
    LARGA_DEL        25 de Junio del 2026
    LARGA_MAY_DEL    25 DE JUNIO DEL 2026
    LARGA_MIN        25 de junio de 2026
    ..._0            el día con cero (02 de julio de 2026)

Una marca de hora: {{HORA__H}} -> 9h00, {{HORA__H_0}} -> 09h00, {{HORA__PUNTOS}} -> 09:00
"""

import re

MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO",
         "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]


# Si la base no tiene el dato, queda una línea para llenarla a mano
EN_BLANCO_FECHA = "____ de ____________ de ______"
EN_BLANCO_HORA = "____h____"


def escribir_fecha(fecha, formato):
    if fecha is None:
        return "____/____/______" if formato.startswith("CORTA") else EN_BLANCO_FECHA
    partes = formato.split("_")
    cero = "0" in partes
    dia = f"{fecha.day:02d}" if cero else str(fecha.day)
    if partes[0] == "CORTA":
        if "SIN" in partes:
            return f"{fecha.day}/{fecha.month}/{fecha.year}"
        return f"{fecha.day:02d}/{fecha.month:02d}/{fecha.year}"
    mes = MESES[fecha.month - 1]
    mayusculas = "MAY" in partes
    if not mayusculas:
        mes = mes.lower() if "MIN" in partes else mes.capitalize()
    de = "DE" if mayusculas else "de"
    union = "del" if "DEL" in partes else "de"
    union = union.upper() if mayusculas else union
    return f"{dia} {de} {mes} {union} {fecha.year}"


def escribir_hora(hora, formato):
    if hora is None:
        return EN_BLANCO_HORA
    if formato == "PUNTOS":
        return f"{hora.hour:02d}:{hora.minute:02d}"
    if formato == "H_0":
        return f"{hora.hour:02d}h{hora.minute:02d}"
    return f"{hora.hour}h{hora.minute:02d}"


def patron_fecha(fecha):
    """Regex que encuentra `fecha` escrita de cualquiera de las formas conocidas."""
    mes = MESES[fecha.month - 1]
    d, m, a = fecha.day, fecha.month, fecha.year
    larga = rf"(?P<dia>0?{d})\s+(?P<de>de|DE|De)\s+(?P<mes>{mes}|{mes.capitalize()}|{mes.lower()})\s+(?P<union>de|del|DE|DEL|De|Del)\s+{a}"
    corta = rf"(?P<cd>0?{d})/(?P<cm>0?{m})/{a}"
    return re.compile(rf"(?<!\d)(?:{larga}|{corta})(?!\d)")


def formato_de(m):
    """Nombre del formato a partir de un match de patron_fecha."""
    if m.group("cd") is not None:
        sin_ceros = (len(m.group("cd")) == 1 and m.group("cd") != "0") or len(m.group("cm")) == 1
        return "CORTA_SIN_CEROS" if sin_ceros else "CORTA"
    partes = ["LARGA"]
    mes = m.group("mes")
    if mes.isupper():
        partes.append("MAY")
    elif mes.islower():
        partes.append("MIN")
    if m.group("union").lower() == "del":
        partes.append("DEL")
    if m.group("dia").startswith("0"):
        partes.append("0")
    return "_".join(partes)


def patron_hora(hora):
    return re.compile(rf"(?<!\d)0?{hora.hour}(?P<sep>h|H|:){hora.minute:02d}(?!\d)")


def formato_hora_de(m):
    if m.group("sep") == ":":
        return "PUNTOS"
    return "H_0" if m.group(0).startswith("0") else "H"
