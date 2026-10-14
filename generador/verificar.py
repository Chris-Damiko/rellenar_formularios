"""
Revisa los formularios generados de un perfil.

1. Ida y vuelta: se generan los formularios de la propia muestra y se comparan,
   línea por línea (Word) y celda por celda (Excel), con los formularios reales
   de la muestra en el OEC. Lo que salga distinto es un dato que la plantilla
   no reemplaza bien, o un dato que en la base está distinto al formulario real.
2. Fugas: en los formularios de cada candidato generado no debe aparecer
   ningún dato de la muestra (nombre, cédula, código, correo, celular).
3. Presencia: el nombre y la cédula del candidato tienen que estar en sus
   formularios, y no puede quedar ninguna marca {{ }}.

USO:
    python verificar.py PRE
"""

import difflib
import glob
import os
import re
import sys
import tempfile
import zipfile
from html import unescape

from openpyxl import load_workbook

import base
import config
import docx_texto as dt
import generar
import plantillas


def texto_excel(ruta):
    """Celdas con valor como 'Hoja!A1=valor' (para comparar con la muestra real)."""
    wb = load_workbook(ruta, data_only=False)
    lineas = []
    for ws in wb.worksheets:
        for c in ws._cells.values():
            if c.value not in (None, ""):
                lineas.append(f"{ws.title}!{c.coordinate}={c.value}")
    return sorted(lineas)


def texto_excel_rapido(ruta):
    """Todo el texto del .xlsx leyendo el XML (para buscar fugas en cientos de archivos)."""
    with zipfile.ZipFile(ruta) as z:
        partes = [z.read(n).decode("utf-8", "ignore") for n in z.namelist()
                  if n == "xl/sharedStrings.xml" or n.startswith("xl/worksheets/sheet")]
    textos = re.findall(r"<t[^>]*>(.*?)</t>|<v>(.*?)</v>", "".join(partes), re.S)
    return [unescape(a or b) for a, b in textos]


def texto_de(ruta, rapido=False):
    if ruta.endswith(".docx"):
        return [l.strip() for l in dt.texto_documento(ruta).split("\n") if l.strip()]
    return texto_excel_rapido(ruta) if rapido else texto_excel(ruta)


def ida_y_vuelta(perfil, candidatos):
    muestra = config.PERFILES[perfil]["muestra"]
    cand = candidatos[muestra]
    origen = config.carpeta_formularios(perfil, muestra)
    plantillas_dir = os.path.join(plantillas.CARPETA_PLANTILLAS, perfil)
    diferencias = 0
    with tempfile.TemporaryDirectory() as tmp:
        lista = [p for p in sorted(glob.glob(os.path.join(plantillas_dir, "*.docx")) +
                                   glob.glob(os.path.join(plantillas_dir, "*.xlsx")))
                 if not os.path.basename(p).startswith(("C03", "C13"))]
        carpeta, faltan = generar.generar_candidato((generar.datos_planos(cand, perfil), os.path.join(tmp, "gen"), lista))
        for a in sorted(os.listdir(origen)):
            nombre = plantillas.nombre_final(a)
            if not nombre or nombre == "C03":
                continue
            ext = "docx" if a.lower().endswith((".doc", ".docx")) else "xlsx"
            real = os.path.join(origen, a)
            if not a.lower().endswith(ext):
                real = plantillas.convertir(real, tmp, ext)
            generado = os.path.join(carpeta, f"{nombre}.{ext}")
            if nombre.startswith("T04"):  # el T04 lleva el código del examinador elegido
                generado = glob.glob(os.path.join(carpeta, "T04*"))[0]
            r, g = texto_de(real), texto_de(generado)
            diff = [l for l in difflib.unified_diff(r, g, "real", "generado", n=0, lineterm="")
                    if l[:1] in "+-" and not l.startswith(("+++", "---"))]
            if diff:
                diferencias += 1
                print(f"\n  ~ {nombre}: {len(diff)} líneas distintas (- real, + generado)")
                for l in diff[:30]:
                    print("     " + l[:160])
    return diferencias


def fugas_y_presencia(perfil, candidatos):
    muestra = candidatos[config.PERFILES[perfil]["muestra"]]
    salida = os.path.join(generar.CARPETA_SALIDA, perfil)
    datos_muestra = {k: muestra[k] for k in ("nombre", "cedula", "codigo", "correo", "celular") if len(muestra[k]) >= 5}
    problemas = 0
    marcas_plantilla = {}
    carpetas = sorted(glob.glob(os.path.join(salida, "*", "FORMULARIOS")))
    for carpeta in carpetas:
        codigo = os.path.basename(os.path.dirname(carpeta))
        if codigo == muestra["codigo"] or codigo not in candidatos:
            continue
        cand = candidatos[codigo]
        for ruta in sorted(glob.glob(os.path.join(carpeta, "*"))):
            nombre = os.path.basename(ruta)
            texto = "\n".join(texto_de(ruta, rapido=True))
            for campo, valor in datos_muestra.items():
                if valor == cand[campo]:
                    continue
                if re.search(re.escape(valor), texto, re.I):
                    print(f"  ✗ {codigo}/{nombre}: aparece {campo} de la muestra ({valor})")
                    problemas += 1
            if "{{" in texto:
                print(f"  ✗ {codigo}/{nombre}: quedó una marca {{{{ }}}}")
                problemas += 1
            if nombre not in marcas_plantilla:
                marcas_plantilla[nombre] = "\n".join(texto_de(os.path.join(plantillas.CARPETA_PLANTILLAS, perfil, nombre), rapido=True))
            plantilla = marcas_plantilla[nombre]
            for campo in ("nombre", "cedula"):
                if "{{" + campo.upper() in plantilla:
                    if cand[campo] and cand[campo] not in texto:
                        print(f"  ✗ {codigo}/{nombre}: no aparece el {campo} del candidato")
                        problemas += 1
    return len(carpetas), problemas


def main():
    for perfil in sys.argv[1:] or list(config.PERFILES):
        candidatos, _ = base.leer_base(config.ruta_base(perfil), config.PERFILES[perfil]["hoja"])
        print(f"\n===== {perfil} =====")
        print(f"1) Ida y vuelta con la muestra {config.PERFILES[perfil]['muestra']}:")
        d = ida_y_vuelta(perfil, candidatos)
        print(f"   {'sin diferencias' if not d else f'{d} formularios con diferencias (ver arriba)'}")
        print("2-3) Fugas y presencia en lo generado:")
        n, p = fugas_y_presencia(perfil, candidatos)
        print(f"   {n} candidatos revisados, {p} problemas")


if __name__ == "__main__":
    main()
