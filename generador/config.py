"""
Datos fijos de cada perfil.

`muestra` es un candidato ya terminado (formularios llenos a mano y revisados)
del que se sacan las plantillas del perfil: todo lo que no depende del
candidato (examinador, esquema, unidades de competencia, lugar, firmas,
fotos) queda tal como está en esa muestra.

Los examinadores salen del F01 (1. Doc Registros Calificacion/2. V03 REC 2023/
Version 1/1. Documentos/F001/F01.xls) y coinciden con los C06 reales.
"""

import os

RAIZ_OEC = os.environ.get(
    "RAIZ_OEC",
    os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "CATEEC-OEC"),
)
CARPETA_V03 = os.path.join(RAIZ_OEC, "0. Certificaciones Registradas", "Version V03")
# C13 con el formato CCLT (el que se presentó en la calificación V03). En el
# OEC se imprime en blanco y se llena a mano; aquí se prellena lo conocido.
PLANTILLA_C13 = os.path.join(
    RAIZ_OEC, "1. Doc Registros Calificacion", "2. V03 REC 2023", "Version 1", "1. Documentos", "C13", "C13.doc",
)

EXAMINADORES = {
    "BERNARDA": {"nombre": "BERNARDA PRISCILA AGUIRRE ARIAS", "cedula": "0104111760", "celular": "0960809613"},
    "HECTOR": {"nombre": "HECTOR ANDRES ROMERO AUCANCELA", "cedula": "0302713722", "celular": "0987136236"},
    "ADRIAN": {"nombre": "ADRIAN FABRICIO PESÁNTEZ PESÁNTEZ", "cedula": "0106682172", "celular": "0990852866"},
    "BYRON": {"nombre": "BYRON MARCELO CABRERA RODRIGUEZ", "cedula": "0706410487", "celular": "0969346359"},
    "MILTON": {"nombre": "MILTON FERNANDO VELECELA MEDINA", "cedula": "0302438825", "celular": "0991413460"},
    "MARLO": {"nombre": "MARLO RENE CORONEL CALDERON", "cedula": "0923369565", "celular": "0969145126"},
    "ANA": {"nombre": "ANA BELEN AVILA CRESPO", "cedula": "0302866165", "celular": "0993836068"},
}

PERFILES = {
    "PRC": {"carpeta": "Prevencion en Construccion", "base": "LISTADO PRC.xlsx", "hoja": "VERSION 3",
            "muestra": "PRC-2106", "examinador": "BERNARDA"},
    "PRE": {"carpeta": "Prevencion en Energia", "base": "CERTIFICACIONES PRE.xlsx", "hoja": "Versio03",
            "muestra": "PRE-0670", "examinador": "HECTOR"},
    "GS": {"carpeta": "Gestion Soldadura", "base": "Certificaciones GS.xlsx", "hoja": "Versio03",
           "muestra": "GS-0155", "examinador": "ADRIAN"},
    "IE": {"carpeta": "Instalaciones Electricas", "base": "Certificaciones IE.xlsx", "hoja": "Version3",
           "muestra": "IE-0148", "examinador": "BYRON"},
    "MA": {"carpeta": "Mecatronica Automotriz", "base": "Certificaciones MA.xlsx", "hoja": "Version 3",
           "muestra": "MA-0081", "examinador": "ADRIAN"},
    "OR": {"carpeta": "Operaciones de Lineas y Redes Energizadas", "base": "CERTIFICACIONES OR.xlsx", "hoja": "Hoja1",
           "muestra": "OR-0121", "examinador": "HECTOR"},
    "AL": {"carpeta": "Auxiliar de Liniero", "base": "Certificaciones AL.xlsx", "hoja": "VERSION3",
           "muestra": "AL-0032", "examinador": "MILTON"},
    "AO": {"carpeta": "Asistente en Ofimatica", "base": "Certificaciones AO.xlsx", "hoja": "VERSION3",
           "muestra": "AO-0032", "examinador": "MARLO"},
    "FF": {"carpeta": "Formacion de Formadores", "base": "Copia de Certificaciones FF.xlsx", "hoja": "VERSION3",
           "muestra": "FF-0045", "examinador": "MARLO"},
    "MQ": {"carpeta": "Maquillaje", "base": "LISTADO MQ.xlsx", "hoja": "VERSION 3",
           "muestra": "MQ-0002", "examinador": "ANA"},
}


def ruta_base(perfil):
    p = PERFILES[perfil]
    return os.path.join(CARPETA_V03, p["carpeta"], p["base"])


def carpeta_candidato(perfil, codigo):
    return os.path.join(CARPETA_V03, PERFILES[perfil]["carpeta"], codigo)


def carpeta_formularios(perfil, codigo):
    """Carpeta FORMULARIOS/Formulario/... del candidato, o None si no existe."""
    c = carpeta_candidato(perfil, codigo)
    if not os.path.isdir(c):
        return None
    for s in sorted(os.listdir(c)):
        if s.lower().startswith(("form", "fomr")) and os.path.isdir(os.path.join(c, s)):
            return os.path.join(c, s)
    return None
