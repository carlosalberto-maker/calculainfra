import pandas as pd
import math
import argparse
import sys

def parse_excel(file_path):
    print(f"Cargando archivo: {file_path}")
    try:
        xls = pd.ExcelFile(file_path)
    except Exception as e:
        print(f"Error al cargar el archivo: {e}")
        sys.exit(1)
        
    # Leer Criterios de estimación
    # Asumimos que el usuario creará columnas como: 'Criterio', 'Valor' 
    # o 'Tipo de Área/Personal', 'Equipo Requerido', 'Cantidad'
    print("--- Leyendo Criterios de Estimación ---")
    try:
        criterios_df = pd.read_excel(xls, 'Criterios_estimación')
        criterios_df.columns = criterios_df.columns.str.strip()
        # Limpiar filas vacías
        criterios_df = criterios_df.dropna(how='all')
        print(f"Se encontraron {len(criterios_df)} criterios definidos.")
        if criterios_df.empty:
            print("ADVERTENCIA: La pestaña 'Criterios_estimación' está vacía.")
            print("Asegúrate de llenarla con columnas como 'Criterio' y 'Valor'.")
    except Exception as e:
        print(f"Error al leer Criterios_estimación: {e}")
        criterios_df = pd.DataFrame()

    # Leer Servicios
    print("--- Leyendo Servicios ---")
    try:
        servicios_df = pd.read_excel(xls, 'Servicios', header=4)
        servicios_df.columns = servicios_df.columns.str.strip()
        # Rellenar celdas combinadas de 'Servicio' hacia abajo
        if 'Servicio' in servicios_df.columns:
            servicios_df['Servicio'] = servicios_df['Servicio'].ffill()
        servicios_df = servicios_df.dropna(subset=['Servicio', 'Área'], how='all')
    except Exception as e:
        print(f"Error al leer Servicios: {e}")
        servicios_df = pd.DataFrame()

    # Leer Personal
    print("--- Leyendo Personal ---")
    try:
        personal_df = pd.read_excel(xls, 'Personal (cuerpo de gobierno)', header=4)
        personal_df.columns = personal_df.columns.str.strip()
        # Rellenar celdas combinadas de 'Perfil' hacia abajo
        col_perfil = [col for col in personal_df.columns if 'Perfil' in col][0]
        personal_df[col_perfil] = personal_df[col_perfil].ffill()
        personal_df = personal_df.dropna(subset=[col_perfil], how='all')
    except Exception as e:
        print(f"Error al leer Personal: {e}")
        personal_df = pd.DataFrame()

    # Leer Red
    print("--- Leyendo Red ---")
    try:
        red_df = pd.read_excel(xls, 'Red', header=4)
        red_df.columns = red_df.columns.str.strip()
        if 'Servicio' in red_df.columns:
            red_df['Servicio'] = red_df['Servicio'].ffill()
        red_df = red_df.dropna(subset=['Servicio'], how='all')
    except Exception as e:
        print(f"Error al leer Red: {e}")
        red_df = pd.DataFrame()
        
    return criterios_df, servicios_df, personal_df, red_df

def calcular_computadoras_area(servicio, area, cantidad):
    servicio = str(servicio).strip().lower()
    area = str(area).strip().lower()
    
    # Si la cantidad no está definida o es 0, requerimos 0 unidades
    try:
        cantidad = float(cantidad)
        if math.isnan(cantidad) or cantidad <= 0:
            return 0
    except:
        return 0

    # Urgencias o admisión continua
    if "urgencia" in servicio or "admisión" in servicio:
        if "admisión" in area: return math.ceil(cantidad / 1)
        if "triage" in area: return math.ceil(cantidad / 1)
        if "consultorio" in area: return math.ceil(cantidad / 1)
        if "cama" in area or "observación" in area or "reanimación" in area: return math.ceil(cantidad / 5)
        if "central de enfermería" in area: return 1
        if "jefatura" in area: return 1

    # Hospitalización
    if "hospitalización" in servicio:
        if "cama" in area or "cunero" in area or "recuperación" in area: return math.ceil(cantidad / 5)
        if "central de enfermería" in area: return 1
        if "jefatura" in area: return 1
        
    # Quirófano
    if "quirófano" in servicio or "quirófanos" in area:
        if "sala" in area or "quirófano" in area: return math.ceil(cantidad / 2)
        if "central de enfermería" in area: return 1
        if "jefatura" in area: return 1
        
    # Consulta externa
    if "consulta externa" in servicio:
        if "consultorio" in area: return math.ceil(cantidad / 1)
        if "auxiliar" in area or "administrativo" in area: return math.ceil(cantidad / 2) # Asume 1x2 salas

    # Farmacia
    if "farmacia" in servicio:
        if "hospitalaria" in area: return 1
        if "ventanilla" in area or "consulta externa" in area: return math.ceil(cantidad / 1)

    # Reglas genéricas de fallback si el nombre no hace match exacto pero tiene palabras clave
    if "cama" in area or "cuna" in area: return math.ceil(cantidad / 5)
    if "consultorio" in area: return math.ceil(cantidad / 1)
    if "jefatura" in area or "central" in area: return 1
    
    # Por default, si está listado requiere 1
    return 1

def calcular_faltantes(criterios, servicios, personal, red):
    # Aquí definimos diccionarios para ir sumando lo actual y lo requerido
    requerido = {'Computadoras': 0, 'Impresoras': 0, 'Access Point': 0, 'Switch': 0}
    actual = {'Computadoras': 0, 'Impresoras': 0, 'Access Point': 0, 'Switch': 0}

    # 1. Analizar Servicios (Equipos de cómputo actuales y requeridos)
    if not servicios.empty:
        col_comp_actual = 'Número de equipos de cómputo (actual)'
        col_imp_actual = 'Número de impresoras'
        
        if col_comp_actual in servicios.columns:
            actual['Computadoras'] += pd.to_numeric(servicios[col_comp_actual], errors='coerce').fillna(0).sum()
        if col_imp_actual in servicios.columns:
            actual['Impresoras'] += pd.to_numeric(servicios[col_imp_actual], errors='coerce').fillna(0).sum()
            
        # Iterar sobre cada fila para aplicar las reglas de la imagen
        for idx, row in servicios.iterrows():
            serv = row.get('Servicio', '')
            area = row.get('Área', '')
            cant = row.get('Cantidad (número de)', 1)
            
            comp_req = calcular_computadoras_area(serv, area, cant)
            requerido['Computadoras'] += comp_req

    # 2. Analizar Personal (Cuerpo de gobierno)
    if not personal.empty:
        col_pers_cant = [c for c in personal.columns if 'Cantidad de personal' in c]
        col_comp_pers = '# equipos computo por perfil (cantidad indicada en la columna C)'
        col_imp_pers = '# impresoras por perfil (cantidad indicada en la columna C)'
        col_perfil = [col for col in personal.columns if 'Perfil' in col][0]
        
        if col_comp_pers in personal.columns:
            actual['Computadoras'] += pd.to_numeric(personal[col_comp_pers], errors='coerce').fillna(0).sum()
        if col_imp_pers in personal.columns:
            actual['Impresoras'] += pd.to_numeric(personal[col_imp_pers], errors='coerce').fillna(0).sum()
            
        # Para cuerpo de gobierno, agrupar por perfil y ver si tienen personal asignado
        if col_pers_cant:
            personal[col_pers_cant[0]] = pd.to_numeric(personal[col_pers_cant[0]], errors='coerce').fillna(0)
            agrupado = personal.groupby(col_perfil)[col_pers_cant[0]].sum()
            
            for perfil, cant_personal in agrupado.items():
                if cant_personal > 0:
                    perfil_str = str(perfil).lower()
                    if any(x in perfil_str for x in ['dirección', 'subdirección', 'asistente', 'coordinador', 'jefatura']):
                        # 1 equipo para cada área del cuerpo de gobierno especificada en la imagen si hay personal
                        requerido['Computadoras'] += 1

    # Calcular impresoras requeridas: Regla de 1 x cada 10 equipos de cómputo
    requerido['Impresoras'] = math.ceil(requerido['Computadoras'] / 10.0)

    # 3. Analizar Red
    if not red.empty:
        col_ap_actual = red.columns[1] if len(red.columns) > 1 else None
        col_sw_actual = red.columns[2] if len(red.columns) > 2 else None
        
        if col_ap_actual:
            actual['Access Point'] += pd.to_numeric(red[col_ap_actual], errors='coerce').fillna(0).sum()
        if col_sw_actual:
            actual['Switch'] += pd.to_numeric(red[col_sw_actual], errors='coerce').fillna(0).sum()
        
        # Como las imágenes no especifican criterios de Red, usaremos valores por defecto:
        # Pero solo requeriremos APs para aquellos servicios que tengan al menos algún equipo requerido o personal.
        # En una estimación inicial real, calcular APs y Switches requiere mapeo de planos.
        # Proponemos 1 AP por cada 2 áreas que realmente tengan actividad (con al menos 1 computadora requerida).
        areas_activas = requerido['Computadoras']
        requerido['Access Point'] += math.ceil(areas_activas * 0.5)
        
        # Switch de 24 puertos para cubrir nodos (PCs + APs)
        nodos_totales = requerido['Computadoras'] + requerido['Access Point']
        requerido['Switch'] += math.ceil(nodos_totales / 24.0)

    # Calcular faltantes
    faltantes = {}
    for item in requerido:
        faltantes[item] = requerido[item] - actual[item]
        if faltantes[item] < 0:
            faltantes[item] = 0

    return actual, requerido, faltantes

def generar_reporte(actual, requerido, faltantes):
    print("\n" + "="*50)
    print("REPORTE DE ESTIMACIÓN DE INFRAESTRUCTURA (EDS UM)")
    print("="*50)
    print(f"{'Equipo':<20} | {'Actual':<10} | {'Requerido':<10} | {'Faltante':<10}")
    print("-" * 55)
    for item in faltantes:
        print(f"{item:<20} | {int(actual[item]):<10} | {int(requerido[item]):<10} | {int(faltantes[item]):<10}")
    print("="*50)
    print("\nNota: El cálculo se basó en los datos de las pestañas 'Servicios', 'Personal' y 'Red'.")
    print("Se utilizaron los criterios de la pestaña 'Criterios_estimación'.")

def main():
    parser = argparse.ArgumentParser(description='Calculadora de infraestructura EDS UM basada en criterios de Excel.')
    parser.add_argument('--archivo', type=str, default='Información implementación EDS UM.xlsx',
                        help='Ruta al archivo Excel')
    args = parser.parse_args()

    criterios_df, servicios_df, personal_df, red_df = parse_excel(args.archivo)
    
    actual, requerido, faltantes = calcular_faltantes(criterios_df, servicios_df, personal_df, red_df)
    
    generar_reporte(actual, requerido, faltantes)

if __name__ == "__main__":
    main()
