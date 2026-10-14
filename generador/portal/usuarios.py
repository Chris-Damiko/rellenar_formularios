"""
Usuarios del portal (solo se usan si el portal corre con PORTAL_LOGIN=1).
Las contraseñas se guardan con hash en portal/usuarios.json.

USO:
    python portal/usuarios.py agregar <usuario>     # pide la contraseña (también sirve para cambiarla)
    python portal/usuarios.py quitar <usuario>
    python portal/usuarios.py lista
"""

import getpass
import json
import os
import sys

from werkzeug.security import generate_password_hash

ARCHIVO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "usuarios.json")


def leer():
    try:
        with open(ARCHIVO, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}


def guardar(datos):
    tmp = ARCHIVO + ".tmp"
    with open(os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600), "w", encoding="utf-8") as f:
        json.dump(datos, f, indent=2, ensure_ascii=False)
    os.replace(tmp, ARCHIVO)


def main():
    if len(sys.argv) < 2 or sys.argv[1] not in ("agregar", "quitar", "lista"):
        print(__doc__)
        return 1
    datos = leer()
    if sys.argv[1] == "lista":
        print("\n".join(sorted(datos)) or "(sin usuarios)")
        return 0
    if len(sys.argv) < 3:
        print(__doc__)
        return 1
    usuario = sys.argv[2].strip().lower()
    if sys.argv[1] == "quitar":
        datos.pop(usuario, None)
        guardar(datos)
        print(f"Usuario {usuario} eliminado.")
        return 0
    clave = getpass.getpass(f"Contraseña para {usuario}: ")
    if len(clave) < 10:
        print("La contraseña debe tener al menos 10 caracteres.")
        return 1
    if getpass.getpass("Repita la contraseña: ") != clave:
        print("No coinciden.")
        return 1
    datos[usuario] = {"hash": generate_password_hash(clave)}
    guardar(datos)
    print(f"Usuario {usuario} guardado.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
