"""
Agrega un candidato nuevo al final de la base maestra de un perfil.

Se edita el XML del .xlsx directamente (como en xlsx_texto.py) para no tocar
nada más del libro: comentarios, vínculos externos, formatos y las otras
hojas quedan exactamente igual. Cada celda nueva toma el estilo (formato de
fecha, bordes, letra) de la misma columna en la fila del candidato anterior.

Antes de escribir:
  - se comprueba que el archivo no esté abierto en Excel / LibreOffice,
  - se guarda una copia en respaldos/<PERFIL>/.
Después se vuelve a leer la base y se comprueba que la fila quedó bien.
"""

import datetime
import os
import re
import shutil
import threading
import zipfile
from xml.sax.saxutils import escape

from openpyxl.utils import get_column_letter

import base
import config

CARPETA_RESPALDOS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "respaldos")
AMARILLO = "FFFFFF00"
EPOCA = datetime.datetime(1899, 12, 30)
_candado = threading.Lock()


class BaseOcupada(Exception):
    pass


def archivo_abierto(ruta):
    carpeta, nombre = os.path.split(ruta)
    return any(os.path.exists(os.path.join(carpeta, n)) for n in (f"~${nombre}", f".~lock.{nombre}#"))


def respaldar(perfil, ruta):
    destino = os.path.join(CARPETA_RESPALDOS, perfil)
    os.makedirs(destino, exist_ok=True)
    marca = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    copia = os.path.join(destino, f"{marca}_{os.path.basename(ruta)}")
    shutil.copy2(ruta, copia)
    return copia


def ruta_hoja(zin, nombre_hoja):
    """xl/worksheets/sheetN.xml de la hoja con ese nombre."""
    libro = zin.read("xl/workbook.xml").decode("utf-8")
    rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    for m in re.finditer(r"<sheet\b[^>]*/>", libro):
        etiqueta = m.group(0)
        nombre = re.search(r'name="([^"]*)"', etiqueta).group(1)
        if base.normalizar(nombre.replace("&amp;", "&")) == base.normalizar(nombre_hoja):
            rid = re.search(r'r:id="([^"]+)"', etiqueta).group(1)
            rel = next(r for r in re.findall(r"<Relationship\b[^>]*/>", rels) if f'Id="{rid}"' in r)
            destino = re.search(r'Target="([^"]+)"', rel).group(1)
            return destino.lstrip("/") if destino.startswith("/xl/") else "xl/" + destino
    raise ValueError(f"No se encontró la hoja '{nombre_hoja}' en el libro")


def valor_xml(ref, estilo, valor):
    s = f' s="{estilo}"' if estilo else ""
    if valor is None or valor == "":
        return f'<c r="{ref}"{s}/>' if estilo else ""
    if isinstance(valor, datetime.datetime):
        d = valor - EPOCA
        return f'<c r="{ref}"{s}><v>{d.days + d.seconds / 86400}</v></c>'
    if isinstance(valor, datetime.date):
        return valor_xml(ref, estilo, datetime.datetime(valor.year, valor.month, valor.day))
    if isinstance(valor, datetime.time):
        return f'<c r="{ref}"{s}><v>{(valor.hour * 3600 + valor.minute * 60) / 86400}</v></c>'
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return f'<c r="{ref}"{s}><v>{valor}</v></c>'
    return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{escape(str(valor))}</t></is></c>'


class Estilos:
    """Clona estilos de styles.xml cuando hace falta: formato de fecha/hora para
    celdas que reciben una fecha, y relleno amarillo para las filas de C13."""

    FECHA, HORA = 14, 20  # formatos integrados de Excel: fecha corta y h:mm

    def __init__(self, xml):
        self.xml = xml
        xfs = re.search(r"<cellXfs\b[^>]*>(.*?)</cellXfs>", xml, re.S)
        self.xfs = re.findall(r"<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>", xfs.group(1), re.S)
        self.nuevos = []
        self.cache = {}
        formatos = dict(re.findall(r'<numFmt\b[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"', xml))
        self.formato_fecha = {int(k) for k, v in formatos.items() if re.search(r"[dmyhs]", re.sub(r'"[^"]*"|\[[^]]*\]', "", v), re.I)}
        self.formato_fecha |= set(range(14, 23)) | {45, 46, 47}
        self.relleno_amarillo = None
        self.modificado = False

    def _xf(self, indice):
        todos = self.xfs + self.nuevos
        return todos[indice] if indice < len(todos) else todos[0]

    def _atributo(self, xf, nombre, valor):
        if f'{nombre}="' in xf:
            return re.sub(rf'{nombre}="[^"]*"', f'{nombre}="{valor}"', xf)
        return xf.replace("<xf ", f'<xf {nombre}="{valor}" ', 1)

    def _clonar(self, indice, cambios):
        clave = (indice, tuple(sorted(cambios.items())))
        if clave not in self.cache:
            xf = self._xf(indice)
            for k, v in cambios.items():
                xf = self._atributo(xf, k, v)
            self.cache[clave] = len(self.xfs) + len(self.nuevos)
            self.nuevos.append(xf)
            self.modificado = True
        return self.cache[clave]

    def para(self, indice, valor, amarilla):
        indice = int(indice or 0)
        num = int((re.search(r'numFmtId="(\d+)"', self._xf(indice)) or [0, 0])[1])
        if isinstance(valor, (datetime.date, datetime.time)) and num not in self.formato_fecha:
            formato = self.HORA if isinstance(valor, datetime.time) else self.FECHA
            indice = self._clonar(indice, {"numFmtId": formato, "applyNumberFormat": 1})
        if amarilla:
            indice = self._clonar(indice, {"fillId": self._amarillo(), "applyFill": 1})
        return indice

    def _amarillo(self):
        if self.relleno_amarillo is None:
            fills = re.search(r"<fills\b[^>]*>(.*?)</fills>", self.xml, re.S)
            lista = re.findall(r"<fill>.*?</fill>|<fill/>", fills.group(1), re.S)
            existente = next((i for i, f in enumerate(lista) if f'rgb="{AMARILLO}"' in f and "solid" in f), None)
            if existente is None:
                nuevo = (f'<fill><patternFill patternType="solid"><fgColor rgb="{AMARILLO}"/>'
                         f'<bgColor indexed="64"/></patternFill></fill>')
                apertura = re.sub(r'count="\d+"', f'count="{len(lista) + 1}"', fills.group(0).split(">", 1)[0] + ">", 1)
                self.xml = self.xml.replace(fills.group(0), apertura + fills.group(1) + nuevo + "</fills>")
                existente = len(lista)
            self.relleno_amarillo = existente
        return self.relleno_amarillo

    def resultado(self):
        if not self.nuevos:
            return self.xml
        xfs = re.search(r"<cellXfs\b[^>]*>(.*?)</cellXfs>", self.xml, re.S)
        total = len(self.xfs) + len(self.nuevos)
        apertura = re.sub(r'count="\d+"', f'count="{total}"', xfs.group(0).split(">", 1)[0] + ">", 1)
        return self.xml.replace(xfs.group(0), apertura + xfs.group(1) + "".join(self.nuevos) + "</cellXfs>")


def escribir_fila(ruta, hoja, numero_fila, fila_modelo, valores, amarilla=False):
    """valores: {índice de columna (0 = A): valor}. fila_modelo: fila de la que se copian los estilos."""
    with zipfile.ZipFile(ruta) as zin:
        hoja_xml = ruta_hoja(zin, hoja)
        xml = zin.read(hoja_xml).decode("utf-8")
        estilos = zin.read("xl/styles.xml").decode("utf-8")
        entradas = [(i, zin.read(i.filename)) for i in zin.infolist()]

    def fila_xml(n):
        return re.search(rf'<row\b[^>]*\br="{n}"[^>]*?(?:/>|>.*?</row>)', xml, re.S)

    modelo = fila_xml(fila_modelo)
    # Estilo de cada columna: el de la fila más reciente que tenga un valor en
    # esa columna (si el candidato anterior no tenía experiencia, su celda de
    # fecha no tiene formato de fecha).
    estilo_modelo = {}
    for n in range(fila_modelo, max(base.FILA_INICIO_DATOS, fila_modelo - 200) - 1, -1):
        fila = fila_xml(n)
        if not fila:
            continue
        for c in re.finditer(r'<c\b([^>]*)>(.*?)</c>', fila.group(0), re.S):
            ref = re.search(r'\br="([A-Z]+)\d+"', c.group(1)).group(1)
            estilo = re.search(r'\bs="(\d+)"', c.group(1))
            if ref not in estilo_modelo and estilo and re.search(r"<v>|<is>", c.group(2)):
                estilo_modelo[ref] = estilo.group(1)
    if modelo:
        for ref, estilo in re.findall(r'<c\b[^>]*\br="([A-Z]+)\d+"[^>]*?\bs="(\d+)"', modelo.group(0)):
            estilo_modelo.setdefault(ref, estilo)
    actual = fila_xml(numero_fila)
    existentes = {}
    if actual:
        for c in re.finditer(r'<c\b[^>]*\br="([A-Z]+)\d+"[^>]*?(?:/>|>.*?</c>)', actual.group(0), re.S):
            existentes[c.group(1)] = c.group(0)

    columnas = {get_column_letter(i + 1): v for i, v in valores.items()}
    gestor = Estilos(estilos)
    celdas = dict(existentes)
    for letra, valor in columnas.items():
        estilo = estilo_modelo.get(letra)
        if estilo is None and letra in existentes:
            m = re.search(r's="(\d+)"', existentes[letra])
            estilo = m.group(1) if m else None
        estilo = gestor.para(estilo, valor, amarilla)
        celdas[letra] = valor_xml(f"{letra}{numero_fila}", str(estilo), valor)
    estilos = gestor.resultado()
    orden = sorted(celdas, key=lambda l: (len(l), l))
    contenido = "".join(celdas[l] for l in orden if celdas[l])

    if actual:
        apertura = re.match(r"<row\b[^>]*?(?=/?>)", actual.group(0)).group(0)
        xml = xml[:actual.start()] + f"{apertura}>{contenido}</row>" + xml[actual.end():]
    else:
        atributos = ""
        if modelo:
            atributos = re.sub(r'\s+(r|spans)="[^"]*"', "", re.match(r"<row\b([^>]*?)(?=/?>)", modelo.group(0)).group(1))
        nueva = f'<row r="{numero_fila}"{atributos}>{contenido}</row>'
        # insertar en orden dentro de sheetData
        anteriores = [m for m in re.finditer(r'<row\b[^>]*\br="(\d+)"[^>]*?(?:/>|>.*?</row>)', xml, re.S)
                      if int(m.group(1)) < numero_fila]
        pos = anteriores[-1].end() if anteriores else xml.index("<sheetData>") + len("<sheetData>")
        xml = xml[:pos] + nueva + xml[pos:]
    dim = re.search(r'<dimension ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', xml)
    if dim and int(dim.group(4)) < numero_fila:
        xml = xml.replace(dim.group(0), f'<dimension ref="{dim.group(1)}{dim.group(2)}:{dim.group(3)}{numero_fila}"')

    tmp = ruta + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, datos in entradas:
            if info.filename == hoja_xml:
                datos = xml.encode("utf-8")
            elif info.filename == "xl/styles.xml" and gestor.modificado:
                datos = estilos.encode("utf-8")
            zout.writestr(info, datos)
    os.replace(tmp, ruta)


def meses_entre(desde, hasta):
    if isinstance(desde, datetime.datetime) and isinstance(hasta, datetime.datetime):
        return max(0, (hasta.year - desde.year) * 12 + hasta.month - desde.month)
    return None


def valores_fila(col, cand):
    """{índice de columna: valor} a partir de un candidato con el mismo formato que base.leer_base()."""
    m = col.mapa
    v = {}

    def poner(indice, valor):
        if indice is not None and valor not in (None, ""):
            v[indice] = valor
    if m["codigo"] > 0:
        poner(0, cand.get("marca"))
    for clave in ("codigo", "nombre", "cedula", "cod_exam", "f_naci", "edad", "estado", "provincia", "canton",
                  "parroquia", "direccion", "correo", "telefono", "celular", "costo", "fecha_reg", "fecha_acep",
                  "fecha_exam", "lugar", "tipo_lugar", "ciudad_examen", "direccion_examen", "hora"):
        poner(m.get(clave), cand.get(clave))
    poner(m.get("edad2"), cand.get("edad"))
    poner(m.get("estado2"), cand.get("estado"))
    poner(m.get("cedula2"), cand.get("cedula"))
    poner(m.get("examen"), cand.get("cod_exam"))
    for bloque, exp in zip(col.experiencias, cand.get("experiencias", [])):
        for k in ("empresa", "direccion", "telefono", "desde", "hasta", "actividad"):
            poner(bloque[k], exp.get(k))
        poner(bloque["mes"], meses_entre(exp.get("desde"), exp.get("hasta")))
    for bloque, cap in zip(col.capacitaciones, cand.get("capacitaciones", [])):
        for k in ("curso", "institucion", "fechas", "horas"):
            poner(bloque[k], cap.get(k))
    for nivel, bloque in col.educacion.items():
        edu = cand.get("educacion", {}).get(nivel)
        if edu:
            for k in ("institucion", "pais", "ciudad", "titulo"):
                poner(bloque[k], edu.get(k))
    return v


def siguiente_codigo(perfil, candidatos):
    numeros = [int(c.split("-")[1]) for c in candidatos]
    ancho = max((len(c.split("-")[1]) for c in candidatos), default=4)
    return f"{perfil}-{(max(numeros) + 1 if numeros else 1):0{ancho}d}"


def agregar(perfil, cand):
    """Escribe el candidato en la base del perfil. Devuelve (código, ruta del respaldo)."""
    with _candado:
        ruta = config.ruta_base(perfil)
        hoja = config.PERFILES[perfil]["hoja"]
        if archivo_abierto(ruta):
            raise BaseOcupada(f"La base está abierta en Excel o LibreOffice. Ciérrela e intente otra vez: {ruta}")
        candidatos, col = base.leer_base(ruta, hoja)
        if cand["codigo"] in candidatos:
            raise ValueError(f"El código {cand['codigo']} ya existe en la base")
        repetido = next((c for c, d in candidatos.items() if d["cedula"] == cand["cedula"]), None)
        if repetido and not cand.get("permitir_repetido"):
            raise ValueError(f"La cédula {cand['cedula']} ya está registrada como {repetido}")
        ultima = max((d["fila"] for d in candidatos.values()), default=base.FILA_INICIO_DATOS - 1)
        respaldo = respaldar(perfil, ruta)
        escribir_fila(ruta, hoja, ultima + 1, ultima, valores_fila(col, cand), cand.get("amarilla", False))

        # Comprobar leyendo otra vez
        releidos, _ = base.leer_base(ruta, hoja)
        nuevo = releidos.get(cand["codigo"])
        errores = []
        if not nuevo:
            errores.append("la fila no aparece al volver a leer la base")
        else:
            for k in ("nombre", "cedula", "correo", "celular"):
                if (cand.get(k) or "") != nuevo[k]:
                    errores.append(f"{k}: se escribió {cand.get(k)!r} y se leyó {nuevo[k]!r}")
            for k in ("fecha_reg", "fecha_acep", "fecha_exam"):
                if cand.get(k) and nuevo[k] and cand[k].date() != nuevo[k].date():
                    errores.append(f"{k} distinta al volver a leer")
            if cand.get("amarilla") and not nuevo["amarilla"]:
                errores.append("la fila no quedó en amarillo")
        if errores:
            shutil.copy2(respaldo, ruta)
            raise RuntimeError("No se guardó (se restauró el respaldo): " + "; ".join(errores))
        return cand["codigo"], respaldo
