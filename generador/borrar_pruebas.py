"""
Quita candidatos de prueba de un perfil restaurando el respaldo de la base
tomado justo antes de registrarlos. Correr dentro de generador/:

    .venv/bin/python borrar_pruebas.py PRC PRC-3067 PRC-3068 PRC-3069            # solo muestra
    .venv/bin/python borrar_pruebas.py PRC PRC-3067 PRC-3068 PRC-3069 --aplicar  # lo hace
"""
import json
import os
import shutil
import sys

import base
import base_escritura
import config
import generar

args = [a for a in sys.argv[1:] if a != "--aplicar"]
aplicar = "--aplicar" in sys.argv
perfil, pruebas = args[0], set(args[1:])
ruta, hoja = config.ruta_base(perfil), config.PERFILES[perfil]["hoja"]
numero = lambda c: int(c.split("-")[1])

actual, _ = base.leer_base(ruta, hoja)
no_estan = pruebas - set(actual)
if no_estan:
    sys.exit(f"❌ No están en la base: {', '.join(sorted(no_estan))}")
for c in sorted(pruebas):
    print(f"   prueba: {c} {actual[c]['nombre']} ({actual[c]['cedula']})")
esperado = set(actual) - pruebas

# Respaldo más reciente que tenga exactamente los demás candidatos
carpeta = os.path.join(base_escritura.CARPETA_RESPALDOS, perfil)
elegido = None
for f in sorted(os.listdir(carpeta), reverse=True):
    try:
        cands, _ = base.leer_base(os.path.join(carpeta, f), hoja)
    except BaseException:
        continue
    if set(cands) == esperado:
        elegido = f
        break
if not elegido:
    sys.exit("❌ Ningún respaldo tiene exactamente los demás candidatos. No se cambió nada.")
print(f"✅ Respaldo a restaurar: {elegido}  ({len(esperado)} candidatos, último {max(esperado, key=numero)})")

estado = generar.leer_estado()
nuevo_estado = max((c for c in esperado if numero(c) < min(map(numero, pruebas))), key=numero)
print(f"   estado.json {perfil}: {estado.get(perfil)} -> {nuevo_estado if estado.get(perfil) in pruebas else '(sin cambio)'}")
salidas = [os.path.join(generar.CARPETA_SALIDA, perfil, c) for c in sorted(pruebas)]
print(f"   se borran: {', '.join(s for s in salidas if os.path.isdir(s)) or '(no hay carpetas)'}")

if not aplicar:
    print("\nNo se cambió nada. Para hacerlo, repita con --aplicar")
    sys.exit()
if base_escritura.archivo_abierto(ruta):
    sys.exit("❌ La base está abierta en Excel/LibreOffice. Ciérrela y repita.")

copia = base_escritura.respaldar(perfil, ruta)  # por si acaso: la base con las pruebas
shutil.copy2(os.path.join(carpeta, elegido), ruta)
if set(base.leer_base(ruta, hoja)[0]) != esperado:
    shutil.copy2(copia, ruta)
    sys.exit("❌ Al volver a leer no cuadró; se dejó la base como estaba.")
if estado.get(perfil) in pruebas:
    estado[perfil] = nuevo_estado
    generar.escribir_estado(estado)
asignaciones = generar.leer_asignaciones()
if pruebas & set(asignaciones):
    for c in pruebas:
        asignaciones.pop(c, None)
    tmp = generar.ARCHIVO_ASIGNACIONES + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(asignaciones, f, indent=2, ensure_ascii=False)
    os.replace(tmp, generar.ARCHIVO_ASIGNACIONES)
for s in salidas:
    if os.path.isdir(s):
        shutil.rmtree(s)
print(f"\n✅ Listo. Base restaurada ({len(esperado)} candidatos). La base con las pruebas quedó en {copia}")
