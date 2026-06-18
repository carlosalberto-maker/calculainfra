# calculainfra

Script en Python (`calcular_infraestructura.py`) para calcular automáticamente la infraestructura tecnológica (equipos de cómputo, impresoras, access points y switches) requerida en una unidad médica EDS UM, a partir de una plantilla de Excel que llenas con los datos del hospital.

El script lee el Excel, aplica los **criterios de estimación oficiales** (los que están como imágenes en la pestaña `Criterios_estimación`) y genera un reporte con tres columnas: lo que el hospital tiene **actualmente**, lo que **requiere**, y lo que le **falta**.

---

## Cómo ejecutarlo paso a paso

> **Requisitos previos** (ya instalados en este equipo): Python 3.13 y las dependencias `pandas` + `openpyxl`. Si vas a usar otro equipo, ve al final del README (sección [Instalación en un equipo nuevo](#instalación-en-un-equipo-nuevo)).

### Paso 1 — Copia la plantilla y dale el nombre del hospital

Tienes el archivo base `Información implementación EDS UM.xlsx`. Para **cada hospital** haz una copia y nómbrala con un nombre claro, por ejemplo:

- `Hospital_Gomez_Palacio.xlsx`
- `Hospital_Torreon_Centro.xlsx`
- `Hospital_Saltillo.xlsx`

### Paso 2 — Llena el Excel del hospital

Llena **únicamente la columna "Cantidad (número de)"** y las columnas de "actual" en las tres pestañas relevantes:

#### Pestaña `Servicios`
Por cada fila (Consultorios de triage, Camas de observación, etc.), llena:
- **Cantidad (número de)** → cuántas unidades hay de esa área (consultorios, camas, salas, módulos…). Si no aplica, déjala vacía o pon `0`.
- **Número de equipos de cómputo (actual)** → cuántas computadoras ya hay instaladas en esa área.
- **Número de impresoras** → cuántas impresoras ya hay.
- (Opcional) Piso, especialidad, horario, propietario, observaciones.

> **Importante:** las filas que empiezan con "Áreas de…" son agrupadores y **no se cuentan** (las camas se cuentan en la fila "Camas de observación", no en "Áreas de observación").

#### Pestaña `Personal (cuerpo de gobierno)`
Por cada perfil/turno:
- **Cantidad de personal** → cuántas personas hay en ese turno (`0` si no hay).
- **# equipos cómputo por perfil** → cuántas computadoras tiene ese perfil actualmente.
- **# impresoras por perfil** → cuántas impresoras.

#### Pestaña `Red`
La hoja `Red` tiene **dos bloques** de columnas con los mismos encabezados:

| Columna | Bloque | Significado |
|---|---|---|
| B – Acces Point | **Izquierdo (ACTUAL)** | Cuántos APs hay instalados hoy en ese servicio |
| C – Switch (24 Puertos) | **Izquierdo (ACTUAL)** | Cuántos switches de 24 puertos hay hoy |
| D – Observaciones | **Izquierdo (ACTUAL)** | Tus notas del levantamiento |
| E – Acces Point | **Derecho (FALTANTE)** | Lo que falta de APs *(lo calcula el script)* |
| F – Switch (24 Puertos) | **Derecho (FALTANTE)** | Lo que falta de switches *(lo calcula el script)* |
| G – Observaciones | **Derecho (FALTANTE)** | Comentario del análisis *(lo escribe el script)* |

**Tú sólo debes llenar el bloque izquierdo (B, C, D)** con los valores actuales del hospital. El bloque derecho (E, F, G) lo deja en blanco — el script lo llenará por ti si usas el flag `--actualizar-excel` (ver paso 4).

### Paso 3 — Abre una terminal en la carpeta del script

En el Explorador de Windows, navega a `c:\Users\ECONAS12\Documents\Dev\calculainfra`, haz clic derecho dentro de la carpeta y elige **"Abrir en Terminal"** (o "Abrir PowerShell aquí").

### Paso 4 — Ejecuta el script

> **Por defecto el script NO modifica tu Excel**: sólo lo lee e imprime el reporte en la terminal.

Pásale el nombre del Excel del hospital con `--archivo`:

```powershell
python calcular_infraestructura.py --archivo "Hospital_Gomez_Palacio.xlsx"
```

Si el archivo se llama exactamente `Información implementación EDS UM.xlsx` (el nombre por defecto), puedes omitir el parámetro:

```powershell
python calcular_infraestructura.py
```

**(Opcional, no recomendado salvo que lo quieras explícitamente)** Con el flag `--actualizar-excel` el script escribirá los faltantes de AP y Switch en el bloque derecho de la hoja `Red` del Excel:

```powershell
python calcular_infraestructura.py --archivo "Hospital_Gomez_Palacio.xlsx" --actualizar-excel
```

### Paso 5 — Lee el reporte

El script imprime en la terminal dos tablas:

**1. Tabla global** — el total para todo el hospital:

```
======================================================================
REPORTE DE ESTIMACIÓN DE INFRAESTRUCTURA (EDS UM) — GLOBAL
======================================================================
Equipo             |   Actual |  Requerido |  Faltante
----------------------------------------------------------------------
Computadoras       |        6 |         13 |         7
Impresoras         |        2 |          3 |         1
Access Point       |        3 |          2 |         0
Switch             |        1 |          2 |         1
======================================================================
```

**2. Desglose por servicio** — lo mismo pero separado por servicio, que es lo que pide la segunda mitad de la hoja `Red`:

```
DESGLOSE POR SERVICIO
(A=Actual  R=Requerido  F=Faltante)
-----------------------------------------------------------------------------------------------------
Servicio                            |  Computadoras  |  Impresoras  |  Access Point  |     Switch
                                    |    A    R    F |   A   R   F |    A    R    F |    A    R    F
-----------------------------------------------------------------------------------------------------
Urgencias                           |    5   11    6 |   1   2   1 |    2    1    0 |    1    1    0
Hospitalización (incluye terapi...) |    0    0    0 |   0   0   0 |    1    0    0 |    0    0    0
Áreas de estancia corta (Hemodiá... |    0    0    0 |   0   0   0 |    0    0    0 |    0    0    0
Quirófano                           |    0    0    0 |   0   0   0 |    0    0    0 |    0    0    0
Tococirugía                         |    0    0    0 |   0   0   0 |    0    0    0 |    0    0    0
Consulta externa                    |    0    0    0 |   0   0   0 |    0    0    0 |    0    0    0
Cuerpo de gobierno                  |    1    2    1 |   1   1   0 |    0    1    1 |    0    1    1
-----------------------------------------------------------------------------------------------------
```

La columna **F (Faltante)** de cada bloque es exactamente lo que tienes que requisitar.

### Paso 6 — Guarda el reporte (opcional)

Si quieres guardar el reporte en un archivo de texto, redirige la salida:

```powershell
python calcular_infraestructura.py --archivo "Hospital_Gomez_Palacio.xlsx" > "Reporte_Gomez_Palacio.txt"
```

---

## Criterios de estimación implementados

Los criterios están codificados en el script (sección `calcular_computadoras_area`) porque en la plantilla están como imágenes. Se aplican exactamente como en `Criterios_estimación`:

| Servicio | Área | Criterio |
|---|---|---|
| Urgencias / Admisión continua | Módulo de admisión | 1 x módulo |
| Urgencias | Módulo de TRIAGE | 1 x módulo |
| Urgencias | Consultorios de primer contacto | 1 x consultorio |
| Urgencias | Camas reanimación / corta estancia / observación / tococirugía / choque | 1 x 5 camas |
| Urgencias | Central de enfermería | 1 x central |
| Urgencias | Jefatura de servicio / enfermería | 1 |
| Hospitalización | Camas hospitalización | 1 x 5 camas |
| Hospitalización | Cunero patológico | 1 x 5 cunas |
| Hospitalización | Camas UCIA, UCIP, UCIN | 1 x 5 camas |
| Hospitalización | Camas cirugía ambulatoria | 1 x 5 camas |
| Hospitalización | Recuperación post parto / post quirúrgico | 1 x 5 camas |
| Hospitalización | Central de enfermería | 1 x central |
| Hospitalización | Jefatura de servicio / enfermería | 1 |
| Quirófano | Quirófanos urgencias / centrales | 1 x 2 salas |
| Quirófano | Central de enfermería | 1 x central |
| Quirófano | Jefatura de servicio / enfermería | 1 |
| Consulta externa | Consultorios | 1 x consultorio |
| Consulta externa | Auxiliar administrativo | 1 x 2 salas |
| SAI Farmacia | Farmacia hospitalaria | 1 x área |
| SAI Farmacia | Farmacia de consulta externa | 1 x ventanilla |
| Cuerpo de gobierno | Dirección | 1 |
| Cuerpo de gobierno | Subdirección Médica | 1 |
| Cuerpo de gobierno | Asistente de dirección / coordinador de turno | 1 |
| Cuerpo de gobierno | Jefatura de enfermería | 1 |
| Impresoras (todos) | — | 1 x cada 10 equipos de cómputo |

### Criterios de red (estimación complementaria)

La plantilla **no incluye criterios oficiales** para Access Points ni Switches. El script usa los siguientes valores por defecto, configurables en la parte superior de `calcular_infraestructura.py`:

- **Access Points:** 1 por cada `COMPUTADORAS_POR_AP = 25` equipos de cómputo requeridos.
- **Switches (24 puertos):** 1 por cada `PUERTOS_POR_SWITCH - PUERTOS_DE_RESERVA = 22` nodos (computadoras + impresoras + APs), reservando 2 puertos por switch para uplink.

Si tienes un criterio oficial distinto, modifica esas tres constantes al inicio del script.

---

## Instalación en un equipo nuevo

Si vas a usar el script en otra computadora:

```powershell
# 1. Instala Python 3 (en Windows con winget):
winget install --id Python.Python.3.13 --scope user

# 2. (Abre una nueva terminal para que reconozca python.) Luego instala las librerías:
pip install -r requirements.txt

# 3. Ejecuta como en el paso 4 de arriba.
```

Las dependencias necesarias están en `requirements.txt`: `pandas` y `openpyxl`.

---

## ¿Qué resuelve el script?

1. **Evita cálculos manuales propensos a error:** lee las tablas del Excel (manejando filas en blanco y celdas combinadas) e interpreta cantidades.
2. **Aplica los criterios oficiales:** traduce reglas como "1 equipo por cada 5 camas de observación" o "1 impresora por cada 10 computadoras" en estimaciones precisas de hardware.
3. **Cálculo de red:** con base en el total de computadoras requeridas, estima preliminarmente la cantidad de APs y switches de 24 puertos.
4. **Reporte diferencial:** muestra `Actual | Requerido | Faltante` por cada rubro (Computadoras, Impresoras, APs, Switches), listo para levantar requisiciones.
5. **Desglose por área:** lista cuántos equipos generó cada fila del Excel para que puedas auditar el resultado.
