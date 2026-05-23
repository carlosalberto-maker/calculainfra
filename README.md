# calculainfra

Este repositorio contiene un script en Python (`calcular_infraestructura.py`) diseñado para calcular de forma automática la infraestructura tecnológica (equipos de cómputo, impresoras, access points y switches) requerida para las unidades médicas EDS UM. 

El cálculo se basa estrictamente en reglas y criterios definidos (número de consultorios, camas, salas quirúrgicas, áreas administrativas, etc.) leyendo una plantilla de Excel que el personal médico y de TI llena. Si la plantilla está vacía o sin cantidades, el script reporta 0 equipos faltantes.

## Para Agentes de IA y Desarrolladores

Si eres un agente de IA o un desarrollador clonando este repositorio en un equipo nuevo, necesitas instalar las dependencias de Python para que el script pueda leer el archivo de Excel correctamente. 

En Python, el equivalente a `package.json` es el archivo `requirements.txt`.

### Instalación de dependencias

Ejecuta el siguiente comando en la terminal para instalar las librerías necesarias (`pandas` y `openpyxl`):

```bash
pip install -r requirements.txt
```

*(Si el comando `pip` no funciona, intenta usar `pip3` o `python3 -m pip`).*

### Ejecución del script

Una vez que tengas el archivo de Excel (`Información implementación EDS UM.xlsx`) lleno con los datos correspondientes en las pestañas "Servicios", "Personal" y "Red", ejecuta el script de la siguiente manera:

```bash
python3 calcular_infraestructura.py
```

Si el archivo de Excel tiene otro nombre, puedes pasárselo como parámetro:

```bash
python3 calcular_infraestructura.py --archivo "Nombre_del_archivo.xlsx"
```

### ¿Qué resuelve el script?
1. **Evita cálculos manuales y propensos a error:** Lee las tablas del Excel (incluso manejando filas en blanco o celdas combinadas) e interpreta cantidades.
2. **Aplica criterios exactos:** Traduce criterios cualitativos ("1 equipo por cada 5 camas de observación", "1 impresora por cada 10 computadoras", etc.) en estimaciones precisas de hardware.
3. **Cálculo de Red:** Basado en la cantidad total de equipos (computadoras) demandados en cada área, asume de forma preliminar la cantidad de Access Points y el número de Switches de 24 puertos necesarios para cubrir esos nodos.
4. **Reporte diferencial:** Genera una tabla en terminal mostrando `Actual | Requerido | Faltante` por cada rubro (Computadoras, Impresoras, APs y Switches), facilitando directamente el levantamiento de requisiciones.
