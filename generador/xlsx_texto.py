"""
Relleno de plantillas .xlsx editando el XML directamente.

openpyxl tarda casi un minuto en abrir y guardar el C07 (21 hojas) y además
pierde formas y otros objetos al guardar. Aquí se abre el .xlsx como zip, se
buscan las celdas cuyo texto (sharedStrings o texto en línea) tiene marcas {{CAMPO}} y se
reescribe solo esa celda. El resto del archivo se copia byte a byte.

Una celda que es solo una marca y cuyo valor es fecha o número queda como
fecha/número de Excel (conserva el formato de la celda); en los demás casos
queda como texto.
"""

import datetime
import re
import zipfile
from html import unescape
from xml.sax.saxutils import escape

from lxml import etree

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
MARCA = re.compile(r"\{\{([A-Za-z0-9_]+)\}\}")
CELDA = re.compile(r'<c\b([^>]*?)(/>|>(.*?)</c>)', re.S)
EPOCA = datetime.datetime(1899, 12, 30)


def textos_compartidos(xml):
    raiz = etree.fromstring(xml)
    return ["".join(t.text or "" for t in si.iter(NS + "t")) for si in raiz.iter(NS + "si")]


def serial(fecha):
    delta = fecha - EPOCA
    return delta.days + delta.seconds / 86400


def celda_nueva(atributos, valor):
    """XML de la celda con el valor nuevo, manteniendo r= y s= (estilo)."""
    atributos = re.sub(r'\s+t="[^"]*"', "", atributos)
    if valor is None or valor == "":
        return f"<c{atributos}/>"
    if isinstance(valor, datetime.datetime):
        return f"<c{atributos}><v>{serial(valor)}</v></c>"
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return f"<c{atributos}><v>{valor}</v></c>"
    texto = escape(str(valor))
    return f'<c{atributos} t="inlineStr"><is><t xml:space="preserve">{texto}</t></is></c>'


def rellenar(origen, destino, valor_de, valor_tipado):
    """
    valor_de(clave) -> str      (para marcas dentro de un texto)
    valor_tipado(clave) -> obj  (para celdas que son solo una marca)
    Ambas lanzan KeyError si la clave no existe. Devuelve las claves sin dato.
    """
    faltan = set()
    with zipfile.ZipFile(origen) as zin:
        nombres = zin.namelist()
        compartidos = textos_compartidos(zin.read("xl/sharedStrings.xml")) if "xl/sharedStrings.xml" in nombres else []
        con_marca = {i for i, t in enumerate(compartidos) if "{{" in t}

        def reemplazar_celda(m):
            atributos, cuerpo = m.group(1), m.group(3) or ""
            if 't="s"' in atributos:
                v = re.search(r"<v>(\d+)</v>", cuerpo)
                if not v or int(v.group(1)) not in con_marca:
                    return m.group(0)
                texto = compartidos[int(v.group(1))]
            elif 't="inlineStr"' in atributos and "{{" in cuerpo:
                texto = unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", cuerpo, re.S)))
            else:
                return m.group(0)
            solo = MARCA.fullmatch(texto.strip())
            if solo:
                try:
                    return celda_nueva(atributos, valor_tipado(solo.group(1)))
                except KeyError:
                    faltan.add(solo.group(1))
                    return celda_nueva(atributos, None)

            def cambio(mm):
                try:
                    return valor_de(mm.group(1))
                except KeyError:
                    faltan.add(mm.group(1))
                    return ""
            nuevo = MARCA.sub(cambio, texto)
            return celda_nueva(atributos, nuevo if nuevo.strip() else None)

        with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                datos = zin.read(info.filename)
                if info.filename.startswith("xl/worksheets/sheet") and info.filename.endswith(".xml") \
                        and (con_marca or b"{{" in datos):
                    datos = CELDA.sub(reemplazar_celda, datos.decode("utf-8")).encode("utf-8")
                elif info.filename == "xl/workbook.xml":
                    # Que Excel recalcule las fórmulas (p. ej. =P9 en el C11) al abrir.
                    # Las plantillas guardadas con openpyxl ya traen <calcPr fullCalcOnLoad="1"/>.
                    xml = datos.decode("utf-8")
                    if "<calcPr" in xml and "fullCalcOnLoad" not in xml:
                        xml = xml.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
                    datos = xml.encode("utf-8")
                zout.writestr(info, datos)
    return faltan
