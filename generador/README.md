# Generador de formularios CCLT (CAPTEEC OEC)

Llena los formularios de certificación de cada candidato a partir de la
**base maestra** del perfil, que está en
`CATEEC-OEC/0. Certificaciones Registradas/Version V03/<perfil>/`.

Perfiles: PRC, PRE, GS, IE, MA, OR, AL, AO, FF, MQ.

## Instalación

```bash
cd generador
uv venv .venv
uv pip install --python .venv/bin/python openpyxl xlrd python-docx pillow lxml flask
```

LibreOffice (`soffice`) tiene que estar instalado. Se usa para pasar los .doc
y .xls viejos a .docx y .xlsx cuando se arman las plantillas.

## Portal web (registrar candidatos y generar al instante)

```bash
cd generador
./iniciar_portal.sh          # abre http://127.0.0.1:5000 en esta PC
./iniciar_portal.sh --red    # también desde otras PCs de la red (http://<IP-de-esta-PC>:5000)
```

- **Perfiles:** tarjetas con cuántos candidatos hay en la base y cuál fue el último generado.
- **Registrar candidato:** formulario con todos los datos de la base: proceso, datos personales, domicilio,
  educación, capacitación y experiencia. Se elige el examinador, el lugar y las unidades de competencia.
  Al guardar:
  1. se agrega la fila al final de la base Excel del perfil;
  2. se guarda antes una copia en `respaldos/<PERFIL>/`;
  3. se vuelve a leer la fila para comprobarla;
  4. se generan los formularios y se pueden descargar en .zip.
  - "Solo primaria" agrega el C03.
  - "Ausente" pinta la fila de amarillo y agrega el C13.
  - La base tiene que estar cerrada en Excel/LibreOffice; si está abierta, el portal avisa y no escribe nada.
- **Candidatos:** lista y búsqueda. En cada candidato se puede cambiar el examinador, las UC o el lugar
  y volver a generar. Lo elegido se guarda en `asignaciones.json`.
- **Configuración** (`ajustes.json`):
  - examinadores: nombre, cédula, celular, trato (el Sr./la Sra.), fecha de ingreso y perfiles que examina;
  - por perfil: unidades de competencia (UC1 a UC5), examinador predeterminado, costo y lugares de examinación.

El examinador, las UC y el lugar ya no están fijos en las plantillas:

| Dato | Formularios donde aparece |
|---|---|
| Examinador | C06, C08, C11, C13 y T04 (el archivo se llama `T04_10 <código del examinador>`) |
| UC | C04, C13 y las X del C02 |
| Lugar | C02, C04, C07, C08, C11, C12 y C13 |

## Servidor de prueba (tecsito, https://cloud.capteec.com/formularios/)

- **Ubicación:** `capteec@tecsito:/home/capteec/web/formularios_oec/`.
  - `generador/`: código, plantillas y entorno `.venv`.
  - `datos/`: copia de las 10 bases maestras, con la misma estructura de carpetas del OEC.
- **Proceso:** pm2 `formularios-oec`, en 127.0.0.1:5010, con inicio de sesión obligatorio.
- **nginx:** `/formularios/` dentro del sitio `cloud-https`. No abre ningún puerto nuevo.
- **Índice:** el servidor no tiene las carpetas de los candidatos, así que `indice_oec.json` le indica quién ya
  tiene formularios en el OEC. Se crea con `python generar.py --indice-oec`.

```bash
bash despliegue/desplegar_tecsito.sh      # desde la PC: subir cambios y reiniciar (no pisa los datos del servidor)
# en tecsito, una sola vez, con sudo:
sudo bash ~/web/formularios_oec/generador/despliegue/instalar_nginx.sh
# usuarios del portal (en tecsito):
cd ~/web/formularios_oec/generador && .venv/bin/python portal/usuarios.py agregar <usuario>
pm2 logs formularios-oec                  # ver errores
```

## Uso diario: seguir donde se quedó

1. Llega la base actualizada de un perfil (p. ej. `LISTADO PRC.xlsx`).
2. Reemplaza el archivo en su lugar (el script muestra la ruta en cada corrida):

   | Perfil | Archivo |
   |---|---|
   | PRC | `Version V03/Prevencion en Construccion/LISTADO PRC.xlsx` |
   | PRE | `Version V03/Prevencion en Energia/CERTIFICACIONES PRE.xlsx` |
   | GS | `Version V03/Gestion Soldadura/Certificaciones GS.xlsx` |
   | IE | `Version V03/Instalaciones Electricas/Certificaciones IE.xlsx` |
   | MA | `Version V03/Mecatronica Automotriz/Certificaciones MA.xlsx` |
   | OR | `Version V03/Operaciones de Lineas y Redes Energizadas/CERTIFICACIONES OR.xlsx` |
   | AL | `Version V03/Auxiliar de Liniero/Certificaciones AL.xlsx` |
   | AO | `Version V03/Asistente en Ofimatica/Certificaciones AO.xlsx` |
   | FF | `Version V03/Formacion de Formadores/Copia de Certificaciones FF.xlsx` |
   | MQ | `Version V03/Maquillaje/LISTADO MQ.xlsx` |

   (todas dentro de `CATEEC-OEC/0. Certificaciones Registradas/`). Si la hoja
   viene con otro nombre, el script busca sola la que tiene los encabezados.
3. Corre:

   ```bash
   cd generador
   .venv/bin/python generar.py PRC        # si quedó en PRC-2100, genera desde PRC-2101 hasta el final
   ```

El último código generado de cada perfil se guarda en `estado.json`.

- Si una fila nueva todavía no tiene nombre o cédula, el estado se queda justo
  antes de esa fila. Así la próxima corrida la vuelve a intentar.
- Si el estado está vacío, el punto de partida es el código más alto que ya
  esté generado o que ya tenga formularios en el OEC.

```bash
.venv/bin/python generar.py --estado                 # dónde quedó cada perfil
.venv/bin/python generar.py                          # todos los perfiles, cada uno desde su último
.venv/bin/python generar.py PRC --desde 2101         # forzar el punto de partida (y guardar el estado)
.venv/bin/python generar.py PRC --desde 2101 --hasta 2150
.venv/bin/python generar.py PRC --codigos PRC-2105,PRC-2110   # rehacer sueltos (no cambia el estado)
.venv/bin/python generar.py PRC --pendientes         # todos los que no tienen formularios en el OEC
.venv/bin/python verificar.py PRC                    # revisión
```

Salida: `salida/<PERFIL>/<CODIGO>/FORMULARIOS/`. Tiene la misma estructura
que la carpeta del candidato en el OEC, para copiarla tal cual.
`salida/<PERFIL>/REVISAR.txt` lista lo que hay que revisar a mano en la última corrida.
El generador no escribe nada dentro de `CATEEC-OEC`.

Las plantillas ya están armadas en `plantillas/`. Solo hay que rehacerlas
(`.venv/bin/python plantillas.py PRC`) si cambian los formatos o el examinador.

## Cómo funciona

- **base.py** lee la base maestra. Cada perfil tiene las columnas en letras
  distintas, así que cada dato se ubica por su encabezado (fila 2), bloque
  por bloque: datos personales, experiencia 1 y 2, capacitación,
  educación, fechas, lugar y hora.
- **plantillas.py** arma las plantillas de cada perfil a partir de un
  candidato ya terminado (la `muestra` de `config.py`).
  - Se copian sus formularios reales y se cambian los datos de ese
    candidato por marcas `{{CAMPO}}`.
  - Lo demás queda igual que en la muestra: examinador, esquema, unidades
    de competencia, lugar, firmas y fotos.
  - Cada cambio se anota en `plantillas/<PERFIL>/INDUCCION.txt`, para revisarlo.
- **generar.py** reemplaza las marcas con los datos de cada candidato.
  - Word: reemplaza el texto párrafo por párrafo, sin perder el formato.
  - Excel: edita el XML de las celdas directamente, sin abrir el libro con openpyxl.
- **verificar.py** hace tres revisiones:
  1. Genera los formularios de la muestra y los compara con los reales.
  2. Revisa que ningún dato de la muestra aparezca en los demás candidatos.
  3. Revisa que no quede ninguna marca `{{ }}`.
- **elegir_muestra.py** busca, por perfil, una muestra cuyos formularios tengan
  realmente su nombre. Algunas carpetas del OEC tienen formularios copiados de
  otro candidato, por ejemplo IE-0150, IE-0149, OR-0122 y OR-0123.

### Marcas

| Marca | Ejemplo |
|---|---|
| `{{NOMBRE}}`, `{{CEDULA}}`, `{{CODIGO}}` | `CCLT-{{CODIGO}}` → `CCLT-PRE-0700` |
| `{{FECHA_REG}}`, `{{FECHA_ACEP}}`, `{{FECHA_EXAM}}` | en Excel queda como fecha |
| `{{FECHA_EXAM__LARGA_DEL}}` | `9 de Julio del 2026` (ver `formatos.py`) |
| `{{HORA__H}}` / `{{HORA__H_0}}` / `{{HORA__PUNTOS}}` | `9h00` / `09h00` / `09:00` |
| `{{DIRECCION__MAY}}` | en mayúsculas |
| `{{EDU_SECUNDARIA_INSTITUCION}}`, `{{CAP1_CURSO}}`, `{{EXP1_DESDE}}`… | C02 |
| `{{EXAMINADOR_NOMBRE}}`, `{{EXAMINADOR_CEDULA}}` | C08, C11 (sale de `config.py`) |

### Reglas

- **C03** (lecto-escritura): se genera si la columna A de la base dice `PRI`.
  Si el candidato no tiene primaria registrada, en el C02 se pone
  "CAPTEEC … / ECUADOR / LA TRONCAL", como en los formularios reales.
- **Tratamiento:** "Estimado Sr." se cambia por "Estimado(a) Sr.(a)", porque la
  base no dice el sexo del candidato.
- **Cédula del examinador en el C11:** se toma de `config.py`. La muestra PRE
  tenía una cédula equivocada (1314001072).
- **Lugar del examen:** es el de la muestra del perfil. Si el lugar de un
  candidato en la base es otro, aparece en `REVISAR.txt`.
- **C13** (ausencia y cambio de fecha): se genera solo para las filas pintadas de
  amarillo en la base. PRC-2093, que está en amarillo, es el único con C13 en el OEC.
  - Se prellenan nombre, solicitud, esquema, UC, fecha, hora, lugar y examinador.
  - El motivo (NO ASISTIÓ / SOLICITA CAMBIO), la fecha del registro, la nueva
    fecha y las firmas se llenan a mano, igual que en el formato real.
  - La plantilla es `2. V03 REC 2023/Version 1/1. Documentos/C13/C13.doc`.
- **C04, columna CUMPLE:** se marcan las 4 X siempre, como en todos los C04 reales,
  aunque falte el dato en la base (los documentos están en la carpeta DOC).
- **Fechas en el C02:** salen como m/d/aaaa porque el .xls original usa `m/d/yy`.
