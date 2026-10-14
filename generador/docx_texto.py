"""
Búsqueda y reemplazo de texto en .docx sin perder el formato.

Word parte el texto de un párrafo en varios "runs" (w:t) según el formato, así
que un nombre puede quedar repartido en tres pedazos. Aquí se trabaja con el
texto completo del párrafo y se reescriben solo los w:t que toca cada cambio,
conservando el formato del primero.

Se recorren todas las partes del documento: cuerpo, tablas, encabezados, pies
de página y cuadros de texto.
"""

import re

from docx import Document

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def partes_xml(documento):
    """Elementos raíz de document.xml, header*.xml, footer*.xml, footnotes, etc."""
    vistos = set()
    for parte in documento.part.package.iter_parts():
        if not parte.content_type.startswith("application/vnd.openxmlformats-officedocument.wordprocessingml"):
            continue
        elemento = getattr(parte, "element", None)
        if elemento is None or id(elemento) in vistos:
            continue
        vistos.add(id(elemento))
        yield elemento


def parrafos(documento):
    for raiz in partes_xml(documento):
        for p in raiz.iter(W + "p"):
            yield p


def textos_de(p):
    """Los w:t del párrafo que no están dentro de un párrafo anidado (cuadros de texto)."""
    salida = []
    for t in p.iter(W + "t"):
        padre = t.getparent()
        while padre is not None and padre.tag != W + "p":
            padre = padre.getparent()
        if padre is p:
            salida.append(t)
    return salida


def texto_parrafo(p):
    return "".join(t.text or "" for t in textos_de(p))


def reemplazar_rango(ts, inicio, fin, nuevo):
    """Reemplaza los caracteres [inicio, fin) del texto unido de `ts` por `nuevo`."""
    pos = 0
    primero = None
    for t in ts:
        txt = t.text or ""
        a, b = pos, pos + len(txt)
        pos = b
        if b <= inicio or a >= fin:
            continue
        ini_local = max(inicio, a) - a
        fin_local = min(fin, b) - a
        if primero is None:
            primero = t
            t.text = txt[:ini_local] + nuevo + txt[fin_local:]
        else:
            t.text = txt[:ini_local] + txt[fin_local:]
        t.set(XML_SPACE, "preserve")


def reemplazar_en_parrafo(p, patron, funcion):
    """Aplica `funcion(match) -> str` a cada coincidencia de `patron` en el párrafo.

    Se procesan de derecha a izquierda para no mover las posiciones pendientes.
    Devuelve cuántos cambios hizo.
    """
    ts = textos_de(p)
    if not ts:
        return 0
    completo = "".join(t.text or "" for t in ts)
    coincidencias = list(patron.finditer(completo))
    for m in reversed(coincidencias):
        reemplazar_rango(ts, m.start(), m.end(), funcion(m))
    return len(coincidencias)


MARCA = re.compile(r"\{\{([A-Z0-9_]+)\}\}")


def rellenar(ruta_entrada, ruta_salida, valores):
    """Cambia cada {{CAMPO}} por valores[CAMPO]. Devuelve los campos sin valor."""
    documento = Document(ruta_entrada)
    faltan = set()

    def valor(m):
        clave = m.group(1)
        if clave not in valores:
            faltan.add(clave)
            return ""
        return str(valores[clave])

    for p in parrafos(documento):
        if "{{" in texto_parrafo(p):
            reemplazar_en_parrafo(p, MARCA, valor)
    documento.save(ruta_salida)
    return faltan


def texto_documento(ruta):
    """Todo el texto del documento, un párrafo por línea (para verificar)."""
    return "\n".join(texto_parrafo(p) for p in parrafos(Document(ruta)))
