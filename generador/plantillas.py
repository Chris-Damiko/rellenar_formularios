"""
Arma las plantillas de un perfil a partir de un candidato ya terminado (la
"muestra" de config.py).

Se copian los formularios reales de la muestra, se pasan a .docx/.xlsx y se
cambian los datos de ese candidato por marcas {{CAMPO}}. Todo lo demás
(examinador, esquema, unidades de competencia, lugar, firmas y fotos) queda
igual que en la muestra.

    Word:  se buscan en el texto los datos de la muestra (nombre, cédula,
           código, fechas, hora...) tal como están en la base maestra.
    Excel: las celdas tienen siempre la misma posición, así que se usa un
           mapa fijo (MAPA_EXCEL) y se comprueban las etiquetas antes de tocar
           nada.

Al final se escribe INDUCCION.txt con cada cambio hecho, para revisarlo.

USO:
    python plantillas.py PRE            # un perfil
    python plantillas.py                # todos
"""

import os
import re
import shutil
import subprocess
import sys
import tempfile

from docx import Document
from openpyxl import load_workbook

import ajustes
import base
import config
import docx_texto as dt
import formatos

CARPETA_PLANTILLAS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "plantillas")

# Nombre final de cada formulario según cómo empieza el archivo de la muestra
NOMBRES = [
    (r"^C0?1[_\-]", "C01_0001"), (r"^C0?2[_\-]", "C02-0001"), (r"^C0?3", "C03"),
    (r"^C0?4[_\-]", "C04-0001"), (r"^C0?5[_\-]", "C05-0001"), (r"^C0?6[_\-]", "C06-0001"),
    (r"^C0?7[_\-]", "C07-0001"), (r"^C0?8[_\-]", "C08-0001"), (r"^C0?9[_\-]", "C09-0001"),
    (r"^C10[_\-]", "C10-0001"), (r"^C11[_\-]", "C11-0001"), (r"^C12[_\-]", "C12-0001"),
    (r"^Reporte.*Lugar", "Reporte Fotografico Lugar C007"),
    (r"^Reporte.*Proceso", "Reporte Fotografico Proceso C008"),
    (r"^T04", None),  # se conserva el nombre (T04_10 EXA5, etc.)
]

# Campos de texto que se buscan en cada Word. Los de domicilio solo en C09,
# porque valores cortos como "LA TRONCAL" o "CENTRO" también aparecen como
# texto fijo ("CAPTEEC CENTRO DE CAPACITACION..."). Aun en C09 se buscan solo
# en líneas cortas sin el nombre del OEC (ver DOMICILIO).
CAMPOS_WORD = {
    "*": ["CODIGO", "CEDULA", "NOMBRE"],
    "C09-0001": ["CODIGO", "CEDULA", "NOMBRE", "CORREO", "CELULAR", "DIRECCION", "CANTON", "PROVINCIA"],
    "C03": ["CEDULA", "NOMBRE"],
}

# Palabras que deciden a qué fecha corresponde un texto cuando la muestra tiene
# la misma fecha de registro, aceptación y examen. Gana la más cercana.
CLAVES_FECHA = [("FECHA_ACEP", ["acept"]), ("FECHA_REG", ["recib", "recepci", "solicitud"])]
# El C04 es la carta de aceptación: la fecha de la carta ("La Troncal, ...") es la de aceptación
FECHA_ENCABEZADO = {"C04-0001": "FECHA_ACEP"}
VENTANA_CONTEXTO = 60
DOMICILIO = {"CORREO", "CELULAR", "DIRECCION", "CANTON", "PROVINCIA"}

# Excel: hoja (None = hoja activa) -> etiquetas que deben estar + celdas a marcar
MAPA_EXCEL = {
    # C02: las filas cambian un poco entre perfiles (OR tiene dos filas más
    # antes de la capacitación), así que se ubican por la etiqueta de la columna A.
    "C02-0001": {"hoja": "MATRIZ", "etiquetas": {"A8": "Formulario"}, "celdas": "C02"},
    "C07-0001": {"hoja": None, "etiquetas": {"A10": "NOMBRE DEL LUGAR", "A11": "FECHA DE VERIFICACI"},
                 "celdas": {"H10": "{{LUGAR_TEXTO}}", "H11": "{{FECHA_EXAM}}"}},
    "C08-0001": {
        "hoja": "HOJA DE TRABAJO",
        "etiquetas": {"A8": "LUGAR DE LA EXAMINACI", "A9": "FECHA", "C18": "EXAMINADOR", "C19": "CANDIDATO"},
        "celdas": {"C8": "{{LUGAR_TEXTO}}", "C9": "FECHA_TEXTO", "B18": "{{EXAMINADOR_NOMBRE}}", "E18": "{{EXAMINADOR_CEDULA}}",
                   "G18": "{{EXAMINADOR_CELULAR}}", "B19": "{{NOMBRE}}", "E19": "{{CEDULA}}", "G19": "{{CELULAR}}",
                   "LIMPIAR_FILA": [19, 20]},
    },
    "C11-0001": {
        "hoja": "OBSERVACIÓN AL EXAMINADOR",
        "etiquetas": {"A8": "Apellidos y nombres del examinador", "O9": "Fecha", "A10": "Lugar de la examinaci"},
        "celdas": {"H8": "{{EXAMINADOR_NOMBRE}}", "B9": "{{EXAMINADOR_CEDULA}}", "P9": "{{FECHA_EXAM}}",
                   "A10": "Lugar de la examinación: {{LUGAR_TEXTO}}"},
    },
    "C12-0001": {
        "hoja": "ENCUESTA DE SATISFACCIÓN",
        "etiquetas": {"A85": "LUGAR", "T85": "FECHA", "A87": "NOMBRE DEL EXAMINADO"},
        "celdas": {"G85": "{{LUGAR_TEXTO}}", "X85": "{{FECHA_EXAM}}", "E87": "{{NOMBRE}}"},
    },
}


def fila_con(ws, etiqueta, columna="A"):
    """Primera fila cuya celda de `columna` empieza con `etiqueta` (sin tildes ni mayúsculas)."""
    buscado = base.normalizar(etiqueta)
    for fila in range(1, ws.max_row + 1):
        if base.normalizar(ws[f"{columna}{fila}"].value).startswith(buscado):
            return fila
    raise ValueError(f"C02: no se encontró la fila '{etiqueta}' en la columna {columna}")


def celdas_c02(ws):
    """Celda -> marca del C02, ubicando cada sección por su etiqueta."""
    f = lambda etiqueta: fila_con(ws, etiqueta)
    celdas = {
        "B8": "CCLT-{{CODIGO}}", "I8": "{{FECHA_REG}}",
        f"E{f('Nombres y Apellidos')}": "{{NOMBRE}}", f"E{f('Cédula')}": "{{CEDULA}}",
        f"E{f('Correo')}": "{{CORREO}}", f"E{f('Edad')}": "{{EDAD}}", f"E{f('Provincia')}": "{{PROVINCIA}}",
        f"E{f('Ciudad')}": "{{CANTON}}", f"E{f('Dirección:')}": "{{DIRECCION__MAY}}",
        f"E{f('Teléfono:')}": "{{TELEFONO}}", f"I{f('Teléfono:')}": "{{CELULAR}}",
    }
    if not base.normalizar(ws[f"H{f('Teléfono:')}"].value).startswith("CELULAR"):
        raise ValueError("C02: no está la etiqueta 'Celular:' junto al teléfono")
    for etiqueta, nivel in [("Primaria", "PRIMARIA"), ("Secundaria", "SECUNDARIA"),
                            ("Tercer nivel", "SUPERIOR"), ("Cuarto nivel", "CUARTO")]:
        fila = f(etiqueta)
        for col, k in [("B", "INSTITUCION"), ("H", "PAIS"), ("I", "CIUDAD"), ("J", "TITULO")]:
            celdas[f"{col}{fila}"] = f"{{{{EDU_{nivel}_{k}}}}}"
    curso = f("Nombre del curso")
    for i in range(1, 6):
        for col, k in [("A", "CURSO"), ("F", "INSTITUCION"), ("I", "FECHAS"), ("K", "HORAS")]:
            celdas[f"{col}{curso + i}"] = f"{{{{CAP{i}_{k}}}}}" if i <= 2 else None
    desde = f("Desde")
    for i in (1, 2):
        for col, k in [("A", "DESDE"), ("B", "HASTA"), ("C", "EMPRESA"), ("G", "DIRECCION"),
                       ("I", "TELEFONO"), ("J", "FUNCIONES")]:
            celdas[f"{col}{desde + i}"] = f"{{{{EXP{i}_{k}}}}}"
    # Restos de texto desbordado de otros candidatos (columna O)
    celdas[f"O{desde + 2}"] = None
    celdas[f"O{desde + 3}"] = None
    # Unidades de competencia: X debajo de UC1..UC5
    fila_uc = fila_con(ws, "UC1", "I")
    for i, col in enumerate("IJKLM", start=1):
        if base.normalizar(ws[f"{col}{fila_uc}"].value) != f"UC{i}":
            raise ValueError(f"C02: se esperaba UC{i} en {col}{fila_uc}")
        celdas[f"{col}{fila_uc + 1}"] = f"{{{{UC{i}_X}}}}"
    # Lugar donde desea ser examinado (sección 1.2)
    lugar = fila_con(ws, "Instalaciones del OEC")
    for i, (etiqueta, marca) in enumerate([("Instalaciones del OEC", "LUGAR_OEC"), ("Dirección", "LUGAR_DIRECCION"),
                                           ("Sector", "LUGAR_SECTOR"), ("Teléfono", "LUGAR_TELEFONO")]):
        if not base.normalizar(ws[f"A{lugar + i}"].value).startswith(base.normalizar(etiqueta)):
            raise ValueError(f"C02: se esperaba '{etiqueta}' en A{lugar + i}")
        celdas[f"B{lugar + i}"] = "{{" + marca + "}}"
    return celdas


def lugar_c02(ws):
    """Lugar de examinación tal como está en el C02 de la muestra (para ajustes.json)."""
    fila = fila_con(ws, "Instalaciones del OEC")
    v = lambda c: str(ws[c].value or "").strip()
    return {"lugar": v(f"B{fila}"), "direccion": v(f"B{fila + 1}"), "sector": v(f"B{fila + 2}"),
            "telefono": v(f"B{fila + 3}")}


def nombre_final(archivo):
    raiz = os.path.splitext(archivo)[0]
    for patron, nombre in NOMBRES:
        if re.search(patron, archivo, re.I):
            return nombre or raiz.replace("_", " ").replace("T04 10", "T04_10")
    return None


def convertir(origen, destino_dir, extension):
    """Convierte .doc/.xls con LibreOffice. Devuelve la ruta del archivo nuevo."""
    with tempfile.TemporaryDirectory() as tmp:
        copia = os.path.join(tmp, os.path.basename(origen))
        shutil.copy(origen, copia)
        subprocess.run(["soffice", "--headless", "--convert-to", extension, "--outdir", tmp, copia],
                       check=True, capture_output=True)
        convertido = os.path.splitext(copia)[0] + "." + extension
        destino = os.path.join(destino_dir, os.path.basename(convertido))
        shutil.move(convertido, destino)
        return destino


def valores_muestra(cand):
    v = {
        "CODIGO": cand["codigo"], "CEDULA": cand["cedula"], "NOMBRE": cand["nombre"],
        "CORREO": cand["correo"], "CELULAR": cand["celular"], "DIRECCION": cand["direccion"],
        "CANTON": cand["canton"], "PROVINCIA": cand["provincia"],
    }
    return {k: x for k, x in v.items() if x and len(x) >= 4}


VOCALES = {"A": "AÁÀ", "E": "EÉÈ", "I": "IÍÌ", "O": "OÓÒ", "U": "UÚÙÜ", "N": "NÑ"}


def _sin_tildes(palabra):
    """Regex de una palabra que acepta la versión con o sin tilde (ANDRES / ANDRÉS)."""
    salida = ""
    for c in palabra:
        base_c = base.normalizar(c)
        if base_c in VOCALES:
            letras = VOCALES[base_c]
            salida += "[" + letras + letras.lower() + "]"
        else:
            salida += re.escape(c)
    return salida


def patron_texto(valor):
    return re.compile(r"(?<![\w@.])" + r"\s+".join(_sin_tildes(p) for p in valor.split()) + r"(?![\w@])", re.I)


def sufijo_caso(encontrado, valor):
    if encontrado == valor or encontrado.upper() == valor.upper() and encontrado.isupper() == valor.isupper():
        return ""
    if encontrado.isupper():
        return "__MAY"
    if encontrado.islower():
        return "__MIN"
    return "__TITULO"


def campo_fecha(posibles, contexto):
    """Elige entre FECHA_REG / FECHA_ACEP / FECHA_EXAM según las palabras previas."""
    if len(posibles) == 1:
        return posibles[0]
    ctx = base.normalizar(contexto).lower()
    mejor, pos_mejor = None, -1
    for campo, palabras in CLAVES_FECHA:
        if campo not in posibles:
            continue
        for palabra in palabras:
            pos = ctx.rfind(palabra)
            if pos > pos_mejor:
                mejor, pos_mejor = campo, pos
    if mejor:
        return mejor
    return "FECHA_EXAM" if "FECHA_EXAM" in posibles else posibles[0]


# Datos del perfil que pasan a ser campos (se eligen en el portal por candidato)
FIJOS = ["EXAMINADOR_NOMBRE", "EXAMINADOR_CEDULA", "LUGAR_TEXTO", "UC"]
# T04: compromiso del examinador. Es igual para todos menos nombre, cédula,
# trato y fecha de ingreso (algunas copias reales tienen el nombre de otro).
CAMBIOS_T04 = [
    (r"El día \d{1,2} del mes de \w+ de \d{4}", "El día {{EXA_DIA}} del mes de {{EXA_MES}} de {{EXA_ANIO}}"),
    (r"\b(?:el|la)\s+Sra?\.\s+(?:\{\{EXAMINADOR_NOMBRE\}\}|[A-ZÁÉÍÓÚÑ ]+?)(?=\s+comienza)",
     "{{EXAMINADOR_TRATO}} {{EXAMINADOR_NOMBRE}}"),
    (r"^Nombre:\s*(?!\{\{)[A-ZÁÉÍÓÚÑ ]+$", "Nombre: {{EXAMINADOR_NOMBRE}}"),
    (r"^CC:\s*\d{10}", "CC: {{EXAMINADOR_CEDULA}}"),
]


def inducir_word(ruta, nombre, cand, registro, fijos):
    documento = Document(ruta)
    campos = CAMPOS_WORD.get(nombre, CAMPOS_WORD["*"])
    muestra = valores_muestra(cand)
    textos = [(c, patron_texto(muestra[c])) for c in campos if c in muestra]
    if nombre != "C03":
        textos = [(c, patron_texto(fijos[c]) if c != "UC" else re.compile(r"\bUC1(?:\s*,\s*UC\d)*\b"))
                  for c in FIJOS if fijos.get(c)] + textos
        muestra = {**muestra, **fijos}

    fechas = {}
    for campo in ("FECHA_REG", "FECHA_ACEP", "FECHA_EXAM"):
        f = cand[campo.lower()]
        if f:
            fechas.setdefault(f, []).append(campo)

    previo = ""
    for p in dt.parrafos(documento):
        original = dt.texto_parrafo(p)
        if not original.strip():
            continue
        linea_libre = len(original) < 150 and "CAPTEEC" not in original.upper()
        for campo, patron in textos:
            if campo in DOMICILIO and not linea_libre:
                continue

            def cambio(m, campo=campo):
                caso = "" if campo in FIJOS else sufijo_caso(m.group(0), muestra[campo])
                marca = "{{" + campo + caso + "}}"
                registro.append(f"{nombre}: {m.group(0)!r} -> {marca}")
                return marca
            dt.reemplazar_en_parrafo(p, patron, cambio)
        for f, posibles in fechas.items():
            texto_actual = dt.texto_parrafo(p)

            def cambio_fecha(m, posibles=posibles, texto_actual=texto_actual):
                campo = campo_fecha(posibles, (previo + " " + texto_actual[:m.start()])[-VENTANA_CONTEXTO:])
                if nombre in FECHA_ENCABEZADO and FECHA_ENCABEZADO[nombre] in posibles \
                        and re.search(r"La Troncal,\s*$", texto_actual[:m.start()]):
                    campo = FECHA_ENCABEZADO[nombre]
                marca = "{{" + campo + "__" + formatos.formato_de(m) + "}}"
                registro.append(f"{nombre}: {m.group(0)!r} -> {marca}   [antes: ...{texto_actual[max(0, m.start() - 40):m.start()]!r}]")
                return marca
            dt.reemplazar_en_parrafo(p, formatos.patron_fecha(f), cambio_fecha)
        if cand["hora"]:
            def cambio_hora(m):
                marca = "{{HORA__" + formatos.formato_hora_de(m) + "}}"
                registro.append(f"{nombre}: {m.group(0)!r} -> {marca}")
                return marca
            dt.reemplazar_en_parrafo(p, formatos.patron_hora(cand["hora"]), cambio_hora)

        if nombre.startswith("T04"):
            for patron, nuevo in CAMBIOS_T04:
                if dt.reemplazar_en_parrafo(p, re.compile(patron), lambda m, nuevo=nuevo: nuevo):
                    registro.append(f"{nombre}: {patron} -> {nuevo}")

        # Tratamiento neutro: "Estimado Sr." -> "Estimado(a) Sr.(a)"
        dt.reemplazar_en_parrafo(p, re.compile(r"\bEstimad[oa]\s+Sr\.?(?:a\.?)?(?!\s*\(a\))"),
                                 lambda m: "Estimado(a) Sr.(a)")
        dt.reemplazar_en_parrafo(p, re.compile(r"\bSr(?:a)?\.(?!\s*\(a\))(?=\s*\{\{NOMBRE)"), lambda m: "Sr.(a)")
        previo = (previo + " " + original)[-VENTANA_CONTEXTO:]
    documento.save(ruta)


def inducir_excel(ruta, nombre, cand, registro):
    mapa = MAPA_EXCEL.get(nombre)
    if not mapa:
        return
    wb = load_workbook(ruta)
    ws = wb[mapa["hoja"]] if mapa["hoja"] else wb.active
    for celda, esperado in mapa["etiquetas"].items():
        real = base.normalizar(ws[celda].value)
        if not real.startswith(base.normalizar(esperado)):
            raise ValueError(f"{nombre}: la celda {celda} debería empezar con '{esperado}' y tiene '{ws[celda].value}'")
    celdas = celdas_c02(ws) if mapa["celdas"] == "C02" else mapa["celdas"]
    for celda, marca in celdas.items():
        if celda == "LIMPIAR_FILA":
            continue
        anterior = ws[celda].value
        # Fechas escritas como texto en la celda ("04 de Junio de 2026"): se copia su formato
        fecha_campo = re.fullmatch(r"\{\{(FECHA_[A-Z]+)\}\}", marca) if isinstance(marca, str) else None
        if marca == "FECHA_TEXTO" or (fecha_campo and isinstance(anterior, str)):
            campo = fecha_campo.group(1) if fecha_campo else "FECHA_EXAM"
            f = cand[campo.lower()]
            m = formatos.patron_fecha(f).search(str(anterior or "")) if f else None
            marca = "{{" + campo + "__" + (formatos.formato_de(m) if m else "LARGA") + "}}"
        ws[celda].value = marca
        registro.append(f"{nombre}: {ws.title}!{celda} {str(anterior)[:50]!r} -> {marca}")
    # Restos de otros candidatos en las filas del candidato (C08)
    for fila in celdas.get("LIMPIAR_FILA", []):
        for c in ws[fila]:
            if c.coordinate in celdas or c.coordinate in mapa["etiquetas"] or c.column == 1:
                continue
            if c.value is not None and re.fullmatch(r"[\d\s.]+", str(c.value)):
                registro.append(f"{nombre}: {ws.title}!{c.coordinate} {c.value!r} -> (vacío)")
                c.value = None
    wb.save(ruta)


PUNTOS = r"[….]+[….\s]*"


def datos_fijos(perfil_dir):
    """Esquema, unidades de competencia y lugar del perfil, leídos del C04 ya armado.
    En el C04 la etiqueta y el valor están en celdas distintas: el valor es la
    línea siguiente a la etiqueta."""
    lineas = [l.strip() for l in dt.texto_documento(os.path.join(perfil_dir, "C04-0001.docx")).split("\n") if l.strip()]
    esquema = next((m.group(1) for l in lineas
                    for m in [re.search(r"le notifica .*esquema de certificación:?\s+(.+?)\.?\s+ha sido", l)] if m), None)

    def despues_de(etiqueta):
        for i, l in enumerate(lineas[:-1]):
            if l.startswith(etiqueta):
                resto = l[len(etiqueta):].strip()
                return resto or lineas[i + 1]
        return None
    uc, lugar = despues_de("Unidades de competencia"), despues_de("Lugar de examinación")
    if not (esquema and uc and lugar) or "{{" in lugar or not uc.startswith("UC"):
        raise ValueError(f"C13: no se pudo leer esquema / UC / lugar del C04 ({esquema!r}, {uc!r}, {lugar!r})")
    return esquema.strip(), uc.rstrip(", "), lugar


def armar_c13(perfil, destino, registro):
    """C13 (ausencia y cambio de fecha) con nombre, solicitud, esquema, UC, fecha,
    hora, lugar y examinador. El motivo, la fecha del registro y la nueva fecha
    quedan en blanco para llenar a mano, como en el original."""
    esquema = ajustes.leer()["perfiles"][perfil]["esquema"]
    uc, lugar, examinador = "{{UC}}", "{{LUGAR_TEXTO}}", "{{EXAMINADOR_NOMBRE}}"
    final = os.path.join(destino, "C13.docx")
    os.replace(convertir(config.PLANTILLA_C13, destino, "docx"), final)
    cambios = [
        (rf"Candidato/a:{PUNTOS},", "Candidato/a: {{NOMBRE}},"),
        (rf"Solicitud Nº:\s*{PUNTOS},", "Solicitud Nº: CCLT-{{CODIGO}},"),
        (rf"certificación:{PUNTOS}(?=en las)", f"certificación: {esquema} "),
        (rf"\(UC\):\s*{PUNTOS}(?=para)", f"(UC): {uc} "),
        (rf"el día:\s*{PUNTOS},\s*de\s*{PUNTOS}de 20{PUNTOS},", "el día: {{FECHA_EXAM__LARGA}},"),
        (rf"a las\s*{PUNTOS}h{PUNTOS},", "a las {{HORA__H}},"),
        (rf"en: \(lugar/dirección{PUNTOS}(?=con el)", f"en: {lugar}, "),
        (rf"Examinador/a:\s*{PUNTOS}$", f"Examinador/a: {examinador}."),
        (rf"Yo\s*{PUNTOS}(?=candidato)", "Yo {{NOMBRE}} "),
        (rf"siguiente esquema de certificación:\s*{PUNTOS}$", f"siguiente esquema de certificación: {esquema}."),
        (r"^Nombre: _+$", "Nombre: {{NOMBRE}}"),
    ]
    documento = Document(final)
    hechos = set()
    for p in dt.parrafos(documento):
        for i, (patron, nuevo) in enumerate(cambios):
            if dt.reemplazar_en_parrafo(p, re.compile(patron), lambda m, nuevo=nuevo: nuevo):
                hechos.add(i)
                registro.append(f"C13: {patron} -> {nuevo}")
    faltan = [cambios[i][0] for i in range(len(cambios)) if i not in hechos]
    if faltan:
        raise ValueError(f"C13: no se encontraron estos espacios en la plantilla: {faltan}")
    documento.save(final)


def buscar_c03(perfil, muestra):
    """C03 de la muestra o, si no tiene, el más reciente del perfil."""
    carpeta_perfil = os.path.join(config.CARPETA_V03, config.PERFILES[perfil]["carpeta"])
    codigos = sorted((c for c in os.listdir(carpeta_perfil) if base.PATRON_CODIGO.fullmatch(c)), reverse=True)
    for codigo in [muestra] + codigos:
        f = config.carpeta_formularios(perfil, codigo)
        if not f:
            continue
        for a in os.listdir(f):
            if re.match(r"^C0?3", a, re.I) and a.lower().endswith((".doc", ".docx")):
                return codigo, os.path.join(f, a)
    return None, None


def armar(perfil):
    datos = config.PERFILES[perfil]
    muestra = datos["muestra"]
    candidatos, _ = base.leer_base(config.ruta_base(perfil), datos["hoja"])
    cand = candidatos[muestra]
    origen = config.carpeta_formularios(perfil, muestra)
    if not origen:
        raise SystemExit(f"❌ {muestra} no tiene carpeta de formularios")

    destino = os.path.join(CARPETA_PLANTILLAS, perfil)
    if os.path.isdir(destino):
        shutil.rmtree(destino)
    os.makedirs(destino)

    archivos = [(muestra, os.path.join(origen, a)) for a in sorted(os.listdir(origen))
                if not a.startswith(("~", ".")) and not re.match(r"^C0?3", a, re.I)]
    codigo_c03, ruta_c03 = buscar_c03(perfil, muestra)
    if ruta_c03:
        archivos.append((codigo_c03, ruta_c03))

    registro = [f"Perfil {perfil}. Muestra {muestra}: {cand['nombre']} ({cand['cedula']})",
                f"C03 tomado de {codigo_c03}" if ruta_c03 else "Sin C03 en este perfil", ""]
    convertidos = []
    for codigo, ruta in archivos:
        nombre = nombre_final(os.path.basename(ruta))
        if not nombre:
            registro.append(f"(se ignora {os.path.basename(ruta)})")
            continue
        ext = os.path.splitext(ruta)[1].lower()
        es_word = ext in (".doc", ".docx")
        nuevo_ext = "docx" if es_word else "xlsx"
        if ext in (".docx", ".xlsx"):
            tmp = shutil.copy(ruta, os.path.join(destino, os.path.basename(ruta)))
        else:
            tmp = convertir(ruta, destino, nuevo_ext)
        final = os.path.join(destino, f"{nombre}.{nuevo_ext}")
        os.replace(tmp, final)
        convertidos.append((codigo, nombre, final, es_word))

    # Datos del perfil tal como están en los formularios de la muestra
    esquema, uc, lugar_texto = datos_fijos(destino)
    lugar = lugar_c02(load_workbook(os.path.join(destino, "C02-0001.xlsx"))["MATRIZ"])
    lugar.update({"texto": lugar_texto, "tipo": cand["tipo_lugar"] or "TALLER",
                  "ciudad": cand["ciudad_examen"] or "LA TRONCAL"})
    ajustes.completar_perfil(perfil, esquema, [u.strip() for u in uc.split(",")], lugar)
    exa = config.EXAMINADORES[datos["examinador"]]
    fijos = {"EXAMINADOR_NOMBRE": exa["nombre"], "EXAMINADOR_CEDULA": exa["cedula"],
             "LUGAR_TEXTO": lugar_texto, "UC": uc}
    registro.append(f"Datos fijos de la muestra: {fijos}")

    for codigo, nombre, final, es_word in convertidos:
        c = candidatos.get(codigo, cand)
        if es_word:
            inducir_word(final, nombre, c, registro, fijos)
        else:
            inducir_excel(final, nombre, c, registro)

    armar_c13(perfil, destino, registro)

    with open(os.path.join(destino, "INDUCCION.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(registro) + "\n")
    print(f"✅ {perfil}: {len(os.listdir(destino)) - 1} plantillas en {destino}")


if __name__ == "__main__":
    for p in sys.argv[1:] or list(config.PERFILES):
        armar(p)
