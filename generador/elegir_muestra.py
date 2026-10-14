"""
Busca, para cada perfil, un buen candidato "muestra" para armar las plantillas:
el más reciente que tenga todos los formularios, cuyo nombre y cédula
aparezcan en ellos (algunas carpetas tienen formularios copiados de otro
candidato) y cuyas fechas y hora estén en la base maestra.

No cambia nada: solo imprime la sugerencia para config.py.

USO:
    python elegir_muestra.py            # todos los perfiles
    python elegir_muestra.py IE MA
"""

import os
import re
import subprocess
import sys
import tempfile

import base
import config

# Formularios que deben tener el nombre del candidato
CON_NOMBRE = ["C01", "C04", "C05", "C08", "C10", "C12"]
REVISAR_MAX = 10


def texto(ruta, tmp):
    """Texto plano de un .doc/.docx/.xls/.xlsx usando LibreOffice."""
    ext = os.path.splitext(ruta)[1].lower()
    destino = "txt:Text" if ext in (".doc", ".docx") else "csv:Text - txt - csv (StarCalc):44,34,76"
    subprocess.run(["soffice", "--headless", "--convert-to", destino, "--outdir", tmp, ruta],
                   capture_output=True, check=False)
    salida = os.path.join(tmp, os.path.splitext(os.path.basename(ruta))[0] + (".txt" if "txt:" in destino else ".csv"))
    try:
        with open(salida, encoding="utf-8", errors="ignore") as f:
            return f.read()
    except FileNotFoundError:
        return ""


def revisar(perfil):
    datos = config.PERFILES[perfil]
    candidatos, _ = base.leer_base(config.ruta_base(perfil), datos["hoja"])
    codigos = sorted(candidatos, key=lambda c: int(c.split("-")[1]), reverse=True)
    revisados = 0
    for codigo in codigos:
        cand = candidatos[codigo]
        carpeta = config.carpeta_formularios(perfil, codigo)
        if not carpeta or len(os.listdir(carpeta)) < 13:
            continue
        if not (cand["fecha_exam"] and cand["fecha_reg"] and cand["hora"]):
            continue
        revisados += 1
        if revisados > REVISAR_MAX:
            break
        malos = []
        with tempfile.TemporaryDirectory() as tmp:
            for prefijo in CON_NOMBRE:
                archivo = next((a for a in os.listdir(carpeta) if re.match(prefijo + r"[_\-]", a, re.I)), None)
                if not archivo:
                    malos.append(f"{prefijo} no está")
                    continue
                t = base.normalizar(texto(os.path.join(carpeta, archivo), tmp))
                if base.normalizar(cand["nombre"]) not in t:
                    malos.append(f"{prefijo} con otro nombre")
        estado = "✅ sirve" if not malos else "✗ " + ", ".join(malos)
        print(f"  {codigo}: {estado}")
        if not malos:
            print(f"  → sugerencia para config.py: \"muestra\": \"{codigo}\"   (actual: {datos['muestra']})")
            return codigo
    print("  → ninguna muestra limpia entre las revisadas")
    return None


if __name__ == "__main__":
    for p in sys.argv[1:] or list(config.PERFILES):
        print(f"\n== {p}")
        revisar(p)
