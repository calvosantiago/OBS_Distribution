#!/usr/bin/env python
# coding: utf-8

# ##  PASO 1: LIBRERÍAS Y CONFIGURACIÓN INICIAL

# In[1]:

from __future__ import annotations
import pandas as pd
import numpy as np
import os
from datetime import date, datetime, timedelta
import shutil
import warnings
from pathlib import Path
import country_converter as coco
import pulp
from collections import defaultdict
import itertools
import math
import re
from typing import Dict, Iterable, List, Sequence, Tuple, Optional, Set
warnings.simplefilter(action='ignore', category=UserWarning)

from os.path import expanduser
home = expanduser("~")
print(home)

#Activate virtual environment: .venv\Scripts\activate


# ## BUSCAR ARCHIVOS DE OPs ABIERTA NO ASIGNADAS Y QB_CN MÁS RECIENTE

# In[2]:


# Buscar archivo más reciente de OP No Asignadas 
def absoluteFilePaths(directory):
    l_aux = []
    for dirpath, _, filenames in os.walk(directory):
        for f in filenames:
            x = os.path.abspath(os.path.join(dirpath, f))
            l_aux.append(x)
    return l_aux

pre_ruta = home + "/Downloads"
archivos = absoluteFilePaths(pre_ruta)

maximo, f_maxima = None, None

for x in archivos:
    nombre = os.path.basename(x)
    if nombre.startswith("Oportunidades abiertas No Asignadas JE Totales") and "~" not in nombre and nombre.endswith(".xlsx"):
        try:
            # Separar por espacios y buscar el bloque de fecha + hora
            partes = nombre.replace(".xlsx", "").split(" ")
            for i in range(len(partes)):
                bloque = " ".join(partes[i:i+2])  # ejemplo "18-08-2025 11-36-30"
                try:
                    datetime_object = datetime.strptime(bloque, "%d-%m-%Y %H-%M-%S")
                    if maximo is None or datetime_object > maximo:
                        maximo = datetime_object
                        f_maxima = x
                except ValueError:
                    continue
        except Exception as e:
            print("Error procesando", nombre, "->", e)
            continue

if f_maxima is None:
    raise FileNotFoundError("No se encontró ningún archivo de Oportunidades abiertas No Asignadas JE Totales.")

print("📂 Último archivo detectado:", f_maxima)
print("⏰ Fecha/Hora detectada:", maximo)

# Abrir y cargar
xls = pd.ExcelFile(f_maxima)
df_cupones = xls.parse(0)
df_cupones.rename(columns={'País (Contacto) (Contacto)': 'País'}, inplace=True)
xls.close()


# In[3]:


def absoluteFilePaths(directory):
    l_aux = []
    for dirpath, _, filenames in os.walk(directory):
        for f in filenames:
            x = os.path.abspath(os.path.join(dirpath, f))
            l_aux.append(x)
    return l_aux

# Buscar archivo más reciente de histórico qb_CN_V3 en Descargas
pre_ruta = home + "/Downloads"
archivos = absoluteFilePaths(pre_ruta)

maximo_hist, f_maxima_hist = None, None

for x in archivos:
    nombre = os.path.basename(x)
    if nombre.startswith("qb_CN_V3_OBS") and "~" not in nombre and nombre.endswith(".xlsx"):
        try:
            partes = nombre.replace(".xlsx", "").split(" ")
            for i in range(len(partes)):
                bloque = " ".join(partes[i:i+2])  # ej: "18-08-2025 11-36-30"
                try:
                    # Caso con fecha + hora
                    datetime_object = datetime.strptime(bloque, "%d-%m-%Y %H-%M-%S")
                    if maximo_hist is None or datetime_object > maximo_hist:
                        maximo_hist = datetime_object
                        f_maxima_hist = x
                except ValueError:
                    # Caso con solo fecha (archivos antiguos)
                    try:
                        datetime_object = datetime.strptime(partes[i], "%d-%m-%Y")
                        if maximo_hist is None or datetime_object > maximo_hist:
                            maximo_hist = datetime_object
                            f_maxima_hist = x
                    except ValueError:
                        continue
        except Exception as e:
            print("Error procesando", nombre, "->", e)
            continue

if f_maxima_hist is None:
    raise FileNotFoundError("No se encontró ningún archivo qb_CN_V3 válido en Descargas.")

print("📂 Último archivo histórico detectado:", f_maxima_hist)
print("⏰ Fecha/Hora detectada:", maximo_hist)

# Leer archivo histórico
xls_hist = pd.ExcelFile(f_maxima_hist)
df_hist = xls_hist.parse(0)
xls_hist.close()


# In[4]:


dic_ingles = {
    'Opportunity Id': 'ID de la Oportunidad',
    'Topic': 'Tema',
    'Advised Program of interest from webform': 'Programa de Interes',
    'Country (Originating Lead) (Lead)': 'País',
    'Pillar (Source Campaign) (Campaign)': 'Pillar (Campaña de origen) (Campaña)',
    'Owner': 'Propietario',
    'Program Version of Interest (Originating Lead) (Lead)': 'Programa de Interes',
    'Status Reason': 'Razón para el estado'
}

if 'Opportunity Id' in df_cupones.columns:
    df_cupones.rename(columns=dic_ingles, inplace=True)

# Conservar las columnas clave
df_cupones = df_cupones[['ID de la Oportunidad', 'Programa de Interes', 'País', 'Pillar (Campaña de origen) (Campaña)'] + 
                        [col for col in df_cupones.columns if col not in ['ID de la Oportunidad', 'Programa de Interes',  'País', 'Pillar (Campaña de origen) (Campaña)']]]


# ## ÁREAS DE LOS PROGRAMAS 

# In[5]:


# Leer archivo Areas_Paises.xlsx
ruta_area_paises = home + r"/Grupo Planeta/BI POWER - General/Distribución Atenea/Codigos/Distribución OBS/Areas_Paises.xlsx"
xls_ap = pd.ExcelFile(ruta_area_paises) 
df_areas = xls_ap.parse("Areas")
df_areas.columns = [col.strip() for col in df_areas.columns]

df_cupones['Programa de Interes'] = df_cupones['Programa de Interes'].astype(str).str.strip().str.upper()
df_areas['PROGRAMA'] = df_areas['PROGRAMA'].astype(str).str.strip().str.upper()

# Unir directamente la columna 'Área' desde el Excel
df_cupones = df_cupones.merge(
    df_areas[['PROGRAMA', 'Área', 'TIPO', 'IDIOMA']].rename(columns={'PROGRAMA': 'Programa de Interes'}),
    on='Programa de Interes',
    how='left'
)

# Renombrar columna para que se llame 'AREA' en tu lógica posterior
df_cupones = df_cupones.rename(columns={'Área': 'AREA'})

df_hist['Programa de Interes'] = df_hist['Programa de Interes'].astype(str).str.strip().str.upper()

df_hist = df_hist.drop(columns=['TIPO', 'IDIOMA', 'AREA'], errors='ignore')

df_hist = df_hist.merge(
    df_areas[['PROGRAMA', 'Área', 'TIPO', 'IDIOMA']].rename(columns={'PROGRAMA': 'Programa de Interes'}),
    on='Programa de Interes',
    how='left'
)

df_hist = df_hist.rename(columns={'Área': 'AREA'})



# ## AGRUPACIÓN PAÍSES

# In[6]:


# Cargar hoja de países
df_paises = xls_ap.parse("Paises")
df_paises.columns = [col.strip() for col in df_paises.columns]
# Correcciones manuales antes de usar country_converter
def corregir_pais(pais):
    if not isinstance(pais, str) or pais.strip() == "":
        return ""

    pais_mayus = pais.upper().strip()

    correcciones_pais = {
        "PERÚ": "PERU",
        "PANAMÁ": "PANAMA",
        "MÉXICO": "MEXICO",
        "PHILIPINNES": "PHILIPPINES",
        "BOSTWANA": "BOTSWANA",               
        "GUADALOUPE": "GUADELOUPE",           
        "NETHERLANDS ANTILLES": "CURACAO",    
        "NETHERLAND ANTILLES": "CURACAO",
        "ANTILLAS HOLANDESAS": "CURACAO",
        "UNITED STATES": "USA",
        "UNITED STATES OF AMERICA": "USA",
        "U.S.A.": "USA",
        "EEUU": "USA",
        "": "",
        np.nan: ""
    }

    return correcciones_pais.get(pais_mayus, pais_mayus)


# Aplicar corrección antes de pasar a coco
df_cupones['País Corregido'] = df_cupones['País'].apply(corregir_pais)
df_hist['País Corregido']= df_hist['País (Contacto) (Contacto)'].apply(corregir_pais)
# Paso 1: Normalizar país usando country_converter
df_cupones[ 'País'] = df_cupones[ 'País'].astype(str).str.upper()
df_cupones['País Normalizado'] = coco.convert(names=df_cupones['País Corregido'], to='name_short', not_found=None)
df_hist['País (Contacto) (Contacto)']=df_hist['País (Contacto) (Contacto)'].astype(str).str.upper()
df_hist['País Normalizado']=coco.convert(names=df_hist['País Corregido'],to='name_short',not_found=None)

# Paso 2: Diccionario de agrupación OBS (casos específicos)
dic_agrupacion_obs = {
    "ESPAÑA": "ES", "SPAIN": "ES",
    "CHILE": "CL", "COLOMBIA": "CO", "COSTA RICA": "CR", "ECUADOR": "EC",
    "MÉXICO": "MX", "MEXICO": "MX", "PERÚ": "PE", "PERU": "PE",
    "EL SALVADOR": "VIP", "PANAMÁ": "VIP", "PANAMA": "VIP",
    "PUERTO RICO": "VIP", "URUGUAY": "VIP", "HONDURAS": "VIP"
}

# Paso 3: Diccionario de continentes (desde hoja PAIS A NORMALIZAR → AGRUP ENG)
dic_continente = dict(zip(df_paises['PAIS A NORMALIZAR'].str.upper(), df_paises['AGRUP ENG']))

# Paso 4: Funciones de clasificación
def clasificar_obs(pais_norm):
    pais_mayus = pais_norm.upper() if isinstance(pais_norm, str) else ""
    return dic_agrupacion_obs.get(pais_mayus, "RI")

def clasificar_continente(pais_norm):
    pais_mayus = pais_norm.upper() if isinstance(pais_norm, str) else ""
    return dic_continente.get(pais_mayus, "Desconocido")

# Paso 5: Aplicar clasificación al DataFrame
df_cupones['Agrupación OBS'] = df_cupones['País Normalizado'].apply(clasificar_obs)
df_cupones['Continente'] = df_cupones['País Normalizado'].apply(clasificar_continente)
df_hist['Agrupación OBS'] = df_hist['País Normalizado'].apply(clasificar_obs)
df_hist['Continente'] = df_hist['País Normalizado'].apply(clasificar_continente)


# ## NORMALIZAMOS PILARES

# In[7]:


#  normalización de pilares
ruta_pilares = home + r"\Grupo Planeta\BI POWER - General\PBI\OBS v2\0_TTAA\OBS_PULL_PUSH.xlsx"

df_pilares_norm = pd.read_excel(ruta_pilares, sheet_name='PULL-PUSH')

# Crear diccionario de mapeo
mapa_pilares = dict(zip(df_pilares_norm['PILAR'], df_pilares_norm['PILAR PARA DISTRIBUCION']))

# Normalizar pilares en df_cupones
df_cupones['PILAR_NORM'] = df_cupones['Pillar (Campaña de origen) (Campaña)'].map(mapa_pilares)

# Normalizar pilares en df_hist
df_hist['PILAR_NORM'] = df_hist['Pillar (Campaña de origen) (Campaña)'].map(mapa_pilares)



# ## SUDOKU--- AÑADIMOS LAS HORAS DEL ARCHIVO SUDOKU

# In[8]:


#Obtener las horas con el archivo SUDOKU
def absoluteFilePaths(directory):
    l_aux = []
    for dirpath, _, filenames in os.walk(directory):
        for f in filenames:
            x = os.path.abspath(os.path.join(dirpath, f))
            l_aux.append(x)
    return l_aux
# Buscar archivo con "SUDOKU" en la carpeta Descargas
archivos = absoluteFilePaths(home + "/Grupo Planeta/BI POWER - General/Distribución Atenea/Codigos/Distribución OBS")
ruta_sudoku = ""
for f in archivos:
    if "SUDOKU" in f.upper() and f.endswith(".xlsx") and "~" not in f:
        ruta_sudoku = f
        break

if ruta_sudoku == "":
    raise FileNotFoundError("Archivo con 'SUDOKU' no encontrado en Directorio.")
# Leer solo ese rango
df_sudoku_raw = pd.read_excel(
    ruta_sudoku,
    sheet_name="Estatus diario",
    usecols="P:U",
    skiprows=9,  # 0-based index →
    nrows=6,
    header=1     
)


ruta = rf"{home}\Grupo Planeta\BI POWER - General\Distribución Atenea\Codigos\Distribución OBS\SUDOKU.xlsx"

print(f"⚠️ **ADVERTENCIA:**  \n**Recuerda actualizar las horas en:  \n({ruta})")


# CALCULAMOS PESOS POR AREA A PARTIR DEL SUDOKU

# In[9]:


# Función para calcular semana comercial OBS
def obtener_semana_comercial(fecha_actual: datetime) -> str:
    """Devuelve la semana comercial OBS en formato AÑO-MES-Sn (inicio martes)"""
    año = fecha_actual.year
    mes = fecha_actual.month

    primer_dia_mes = datetime(año, mes, 1)
    primer_martes = primer_dia_mes + timedelta(days=(1 - primer_dia_mes.weekday() + 7) % 7)

    if fecha_actual < primer_martes:
        if mes == 1:
            año -= 1
            mes = 12
        else:
            mes -= 1
        return obtener_semana_comercial(datetime(año, mes, 1))

    dias_diferencia = (fecha_actual - primer_martes).days
    numero_semana = dias_diferencia // 7 + 1

    return f"{año}-{mes:02d}-S{numero_semana}"

# -------------------------------------------------------------------------
# Función auxiliar para obtener el primer martes del mes comercial
def obtener_inicio_mes_comercial(fecha_actual: datetime) -> datetime:
    año = fecha_actual.year
    mes = fecha_actual.month

    primer_dia_mes = datetime(año, mes, 1)
    primer_martes = primer_dia_mes + timedelta(days=(1 - primer_dia_mes.weekday() + 7) % 7)

    if fecha_actual < primer_martes:
        if mes == 1:
            año -= 1
            mes = 12
        else:
            mes -= 1
        primer_dia_mes = datetime(año, mes, 1)
        primer_martes = primer_dia_mes + timedelta(days=(1 - primer_dia_mes.weekday() + 7) % 7)

    return primer_martes

# -------------------------------------------------------------------------
#  función para calcular pesos desde la última fila del SUDOKU
def get_pesos_mensuales(df_sudoku_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Extrae la última fila (totales mensuales) de df_sudoku_raw y calcula los pesos base.
    """
    # Tomar la última fila como totales acumulados del mes
    totales = df_sudoku_raw.iloc[-1]

    df_resultado = pd.DataFrame({
        'EQUIPO': ['Equipo_' + col.strip() for col in totales.index],
        'HORAS': totales.values
    })

    total_horas = df_resultado['HORAS'].sum()
    df_resultado['PESO_BASE'] = df_resultado['HORAS'] / total_horas

    # --- Validación visual ---
    print("\n✅ [VALIDACIÓN] Totales acumulados desde última fila del SUDOKU:")
    print(df_resultado)
    print(f"\n🎯 Total horas acumuladas: {total_horas:.2f}")

    return df_resultado

# -------------------------------------------------------------------------
# Ejecutar función con la última fila como referencia
df_pesos_actuales = get_pesos_mensuales(df_sudoku_raw)

# -------------------------------------------------------------------------
# Función para calcular pesos por área
def get_pesos_por_area(df_pesos_actuales: pd.DataFrame) -> dict:
    equipos_area = {
        'A': ['Equipo_A1', 'Equipo_A2'],
        'B': ['Equipo_B1', 'Equipo_B2'],
        'C': ['Equipo_C1', 'Equipo_C2'],
        'T': df_pesos_actuales['EQUIPO'].tolist(),  # Todos los equipos
        'E': {  # Pesos fijos para programas en inglés
            'Equipo_A1': 0.125,
            'Equipo_B1': 0.25,
            'Equipo_C1': 0.125,
            'Equipo_A2': 0.125,
            'Equipo_B2': 0.25,
            'Equipo_C2': 0.125
        }
    }

    pesos_por_area = {}

    for area, equipos in equipos_area.items():
        if area == 'E':
            df_e = pd.DataFrame({
                'EQUIPO': list(equipos.keys()),
                'PESO_BASE': list(equipos.values())
            })
            pesos_por_area['E'] = df_e
        else:
            df_area = df_pesos_actuales[df_pesos_actuales['EQUIPO'].isin(equipos)].copy()
            total_area = df_area['HORAS'].sum()
            df_area['PESO_BASE'] = df_area['HORAS'] / total_area
            pesos_por_area[area] = df_area[['EQUIPO', 'PESO_BASE']].reset_index(drop=True)

    return pesos_por_area

# -------------------------------------------------------------------------
# Ejecutar función actualizada
pesos_areas = get_pesos_por_area(df_pesos_actuales)

# DF de horas para cálculo de cadencia
df_horas_eq = df_pesos_actuales[['EQUIPO', 'HORAS']].copy()

# Mostrar ejemplo de los pesos de área E
print("\n🎯 Pesos fijos para el área E (programas en inglés):")
print(pesos_areas['E'])

# -------------------------------------------------------------------------
# Para convertirlo en DataFrame consolidado
df_pesos_areas = pd.concat(
    [df.assign(AREA=area) for area, df in pesos_areas.items()],
    ignore_index=True
)
df_pesos_areas = df_pesos_areas[['AREA', 'EQUIPO', 'PESO_BASE']]
print("\n✅ DataFrame consolidado de pesos por área:")
print(df_pesos_areas)


# ## PROCESAMOS EL QB_CN y OP NO ASIGNADAS
# ### QB_CN MAS RECIENTE

# In[10]:


df_hist = df_hist.drop(columns=['TIPO', 'IDIOMA', 'AREA'], errors='ignore')

df_hist = df_hist.merge(
    df_areas[['PROGRAMA', 'Área', 'TIPO', 'IDIOMA']].rename(columns={'PROGRAMA': 'Programa de Interes'}),
    on='Programa de Interes',
    how='left'
)

df_hist = df_hist.rename(columns={'Área': 'AREA'})

# Filtros base
filtro_base = (
    df_hist['Equipo Asignado'] != 'Equipo_Referidos'
) & (
    ~df_hist['PILAR_NORM'].isin(['REF/RECUP', 'OTROS'])
)

# Equipos por área principal
equipos_area = {
    'A': ['Equipo_A1', 'Equipo_A2'],
    'B': ['Equipo_B1', 'Equipo_B2'],
    'C': ['Equipo_C1', 'Equipo_C2'],
}

col_equipo = 'Equipo de Ventas (Usuario propietario) (Usuario)'

tabla = []

for area, equipos in equipos_area.items():
    for equipo in equipos:
        # Área propia: MST + ESP + equipo + su área + filtros extra
        total_area = df_hist[
            (df_hist['TIPO'] == 'MST') &
            (df_hist['IDIOMA'] == 'ESP') &
            (df_hist[col_equipo] == equipo) &
            filtro_base
        ].shape[0] if equipo.startswith(f"Equipo_{area}") else 0

        # Transversal (MBA): MBA + ESP + equipo + filtros extra
        total_T = df_hist[
            (df_hist['TIPO'] == 'MBA') &
            (df_hist['IDIOMA'] == 'ESP') &
            (df_hist[col_equipo] == equipo) &
            filtro_base
        ].shape[0]

        # Inglés: cualquier tipo + ENG + equipo + filtros extra
        total_E = df_hist[
            (df_hist['IDIOMA'] == 'ENG') &
            (df_hist[col_equipo] == equipo) &
            filtro_base
        ].shape[0]

        tabla.append({
            'EQUIPO': equipo,
            f'cupones_{area}': total_area,
            'cupones_T': total_T,
            'cupones_E': total_E
        })

# Convertir a DataFrame
df_resumen_cupones = pd.DataFrame(tabla).sort_values('EQUIPO').reset_index(drop=True)

print(df_resumen_cupones)


# Ruta de destino (carpeta Descargas)
#ruta_salida = str(Path.home() / "Downloads" / "df_hist_exportado.xlsx")

# 1. Primeros pasos idénticos a los tuyos
df_hist_qbCN = df_hist.drop(columns=['TIPO', 'IDIOMA', 'AREA'], errors='ignore')

df_hist_qbCN = df_hist_qbCN.merge(
    df_areas[['PROGRAMA', 'Área', 'TIPO', 'IDIOMA']]
            .rename(columns={'PROGRAMA': 'Programa de Interes'}),
    on='Programa de Interes',
    how='left'
)

df_hist_qbCN = df_hist_qbCN.rename(columns={'Área': 'AREA'})

# 2. Definimos el filtro base
filtro_base = (
    (df_hist_qbCN['Equipo Asignado'] != 'Equipo_Referidos') &
    (~df_hist_qbCN['PILAR_NORM'].isin(['REF/RECUP', 'OTROS']))
)

# 3. Listado de todos los equipos que nos importan
equipos = sum(equipos_area.values(), [])  # ['Equipo_A1','Equipo_A2',..., 'Equipo_C3']

# 4. Filtros de validez de cupón histórico
#    - MST o MBA + ESP
filtro_mst_mba_esp = (
    df_hist_qbCN['TIPO'].isin(['MST', 'MBA']) &
    (df_hist_qbCN['IDIOMA'] == 'ESP')
)
#    - O bien cualquier tipo + ENG
filtro_eng = df_hist_qbCN['IDIOMA'] == 'ENG'

# 5. Filtro de equipo asignado
filtro_equipo = df_hist_qbCN[col_equipo].isin(equipos)

# 6. Aplicamos todos los filtros juntos
df_hist_qbCN = df_hist_qbCN[
    filtro_base &
    filtro_equipo &
    (filtro_mst_mba_esp | filtro_eng)
].copy()





# ###  OP NO ASIGNADAS

# In[11]:


# 1) Copiar y limpiar
df_cupones_open = df_cupones.copy()
for col in ['TIPO', 'IDIOMA', 'AREA']:
    df_cupones_open[col] = (
        df_cupones_open[col]
        .astype(str).str.strip().str.upper()
    )
df_cupones_open["INDEX_ORIGINAL"] = df_cupones_open.reset_index().index

# 1.5) NORMALIZAR y marcar duplicados por Email/Teléfono
def _norm_email(x):
    s = str(x).strip().lower()
    return s if s and s not in {"nan", "none"} else pd.NA

def _norm_phone(x):
    # deja SOLO dígitos; si hay >=9, nos quedamos con los últimos 9 (típico ES).
    s = re.sub(r"\D+", "", str(x))
    if len(s) >= 9:
        return s[-9:]
    return pd.NA

email_col = 'Email (Contacto) (Contacto)'
phone_col = 'Teléfono (Cliente potencial) (Contacto)'

# columnas normalizadas
df_cupones_open['_EMAIL_N'] = (
    df_cupones_open[email_col].map(_norm_email)
    if email_col in df_cupones_open.columns else pd.NA
)
df_cupones_open['_PHONE_N'] = (
    df_cupones_open[phone_col].map(_norm_phone)
    if phone_col in df_cupones_open.columns else pd.NA
)

# duplicado = 2ª+ aparición del mismo email o del mismo teléfono (ignorando nulos)
dup_email = df_cupones_open['_EMAIL_N'].duplicated(keep='first') & df_cupones_open['_EMAIL_N'].notna()
dup_phone = df_cupones_open['_PHONE_N'].duplicated(keep='first') & df_cupones_open['_PHONE_N'].notna()
mask_dup_any = dup_email | dup_phone

# 🎯 Reglas para duplicados:
# - Mantener/poner Propietario = 'Equipo_Z' (quedan como "especiales")
# - Fijar EQUIPO_FINAL = 'Joaquim Barnola Fontrodona'
df_cupones_open.loc[mask_dup_any, 'Propietario'] = 'Equipo_Z'
if 'EQUIPO_FINAL' in df_cupones_open.columns:
    df_cupones_open.loc[mask_dup_any, 'EQUIPO_FINAL'] = 'Joaquim Barnola Fontrodona'

# --- Equipos por área principal (definir antes de usarlos) ---
equipos_area = {
    'A': ['Equipo_A1', 'Equipo_A2'],
    'B': ['Equipo_B1', 'Equipo_B2'],
    'C': ['Equipo_C1', 'Equipo_C2'],
}

# 2) Equipos a mayúsculas
equipos = sum(equipos_area.values(), [])     # ['Equipo_A1', ...]
equipos_upper = [e.upper() for e in equipos] # ['EQUIPO_A1', ...]

# === Especiales claros: Referidos, Z manuales y Duplicados->Joaquim ===

# 1) Conjunto de especiales
special_set = {'EQUIPO_REFERIDOS', 'EQUIPO_Z'}
prop_upper = df_cupones_open['Propietario'].astype(str).str.strip().str.upper()
mask_special = prop_upper.isin(special_set)

# 2) Subset especiales (conserva índices del original)
df_special = df_cupones_open.loc[mask_special].copy()

# 3) Columnas helper
prop_upper_sp = df_special['Propietario'].astype(str).str.strip().str.upper()

# 4) Clasificaciones
mask_ref = prop_upper_sp.eq('EQUIPO_REFERIDOS')
mask_z   = prop_upper_sp.eq('EQUIPO_Z')

# máscara de duplicados del dataset completo, alineada por índice
mask_dup_in_special = mask_dup_any.reindex(df_special.index, fill_value=False)

# 5) Asegurar columnas y asignar con protección
if 'EQUIPO_FINAL' not in df_special.columns:
    df_special['EQUIPO_FINAL'] = pd.NA
if 'SPECIAL_KIND' not in df_special.columns:
    df_special['SPECIAL_KIND'] = pd.NA

# 5a) Z duplicados -> Joaquim
sel_dup_z = (mask_z & mask_dup_in_special)
if sel_dup_z.any():
    df_special.loc[sel_dup_z, 'EQUIPO_FINAL'] = 'Joaquim Barnola Fontrodona'
    df_special.loc[sel_dup_z, 'SPECIAL_KIND'] = 'DUPLICADO->JOAQUIM'

# 5b) Referidos -> Equipo_Referidos (solo si faltaba EQUIPO_FINAL)
sel_ref_sin_final = mask_ref & (
    df_special['EQUIPO_FINAL'].isna() |
    (df_special['EQUIPO_FINAL'].astype(str).str.strip() == '')
)
if sel_ref_sin_final.any():
    df_special.loc[sel_ref_sin_final, 'EQUIPO_FINAL'] = 'Equipo_Referidos'
    df_special.loc[mask_ref, 'SPECIAL_KIND'] = 'REFERIDOS'

# 5c) Z manuales (no duplicados), si falta EQUIPO_FINAL poner "Equipo_Z"
sel_z_no_dup = mask_z & (~mask_dup_in_special)
sel_z_no_dup_sin_final = sel_z_no_dup & (
    df_special['EQUIPO_FINAL'].isna() |
    (df_special['EQUIPO_FINAL'].astype(str).str.strip() == '')
)
if sel_z_no_dup_sin_final.any():
    df_special.loc[sel_z_no_dup_sin_final, 'EQUIPO_FINAL'] = 'Equipo_Z'
    df_special.loc[sel_z_no_dup, 'SPECIAL_KIND'] = 'Z_MANUAL'

# 6) (Opcional) Validación rápida
print("\n=== Resumen especiales (SPECIAL_KIND) ===")
print(df_special['SPECIAL_KIND'].value_counts(dropna=False))

# 7) Volcar cambios a df_cupones_open por índice
cols_push = ['EQUIPO_FINAL', 'SPECIAL_KIND']
if not df_special.empty:
    df_cupones_open.loc[df_special.index, cols_push] = df_special[cols_push]

# ============================================================

# 3) Máscaras de validez (para el resto)
mask_equipo  = prop_upper.isin(equipos_upper)
mask_mst_esp = df_cupones_open['TIPO'].eq('MST') & df_cupones_open['IDIOMA'].eq('ESP')
mask_mba_esp = df_cupones_open['TIPO'].eq('MBA') & df_cupones_open['IDIOMA'].eq('ESP')
mask_eng     = df_cupones_open['IDIOMA'].eq('ENG')

# válidos del flujo “normal” (excluye especiales)
mask_valid_core = mask_equipo & (mask_mst_esp | mask_mba_esp | mask_eng)

# Para el reporte de “perdidos”, NO cuentes los especiales como perdidos:
mask_valid = mask_valid_core | mask_special

# filas que SE PIERDEN (de verdad)
df_perdidos = df_cupones_open.loc[~mask_valid, ['Propietario','TIPO','IDIOMA','AREA']]
print("=== Cupones que se perderían al filtrar (sin contar especiales) ===")
print(df_perdidos)
print(f"Número de cupones perdidos: {len(df_perdidos)}")

# 4) Subset final que sí entra en optimización (sin los especiales)
df_cupones_open = df_cupones_open[mask_valid_core].reset_index(drop=True)

# ------------------- Resumen de cupones -------------------

# Filtros base
filtro_base = (
    df_cupones['Equipo Asignado'] != 'Equipo_Referidos'
) & (
    ~df_cupones['PILAR_NORM'].isin(['REF/RECUP', 'OTROS'])
)

col_equipo = 'Propietario'
tabla = []

for area, equipos in equipos_area.items():
    for equipo in equipos:
        # Área propia: MST + ESP + equipo + filtros extra
        total_area = df_cupones[
            (df_cupones['TIPO'] == 'MST') &
            (df_cupones['IDIOMA'] == 'ESP') &
            (df_cupones[col_equipo] == equipo) &
            filtro_base
        ].shape[0] if equipo.startswith(f"Equipo_{area}") else 0

        # Transversal (MBA): MBA + ESP + equipo + filtros extra
        total_T = df_cupones[
            (df_cupones['TIPO'] == 'MBA') &
            (df_cupones['IDIOMA'] == 'ESP') &
            (df_cupones[col_equipo] == equipo) &
            filtro_base
        ].shape[0]

        # Inglés: cualquier tipo + ENG + equipo + filtros extra
        total_E = df_cupones[
            (df_cupones['IDIOMA'] == 'ENG') &
            (df_cupones[col_equipo] == equipo) &
            filtro_base
        ].shape[0]

        tabla.append({
            'EQUIPO': equipo,
            f'cupones_{area}': total_area,
            'cupones_T': total_T,
            'cupones_E': total_E
        })

# Convertir a DataFrame
df_resumen_cupones_open = pd.DataFrame(tabla).sort_values('EQUIPO').reset_index(drop=True)
print(df_resumen_cupones_open)


# ## REAPERTURAS

# In[12]:


# Clasificación inicial REAP vs FRESH 
df_cupones_open['TIPO_REPARTO'] = df_cupones_open['Tipo de Re-Apertura'].apply(
    lambda x: 'REAP' if pd.notna(x) and str(x).strip() != '' else 'FRESH'
)

# Cargar archivo de estructura comercial
ruta_estructura = home + r"\Grupo Planeta\BI POWER - General\PBI\OBS v2\0_TTAA\OBS_ESTRUCTURA_COMERCIAL.xlsx"
xls_estructura = pd.ExcelFile(ruta_estructura)

# Cargar hoja ESTRUCTURA
df_estructura = xls_estructura.parse("ESTRUCTURA")
df_estructura.columns = [col.strip() for col in df_estructura.columns]

# Cargar hoja de normalización de nombres
df_norm = xls_estructura.parse("NORMALIZACION NOMBRES")
df_norm.columns = [col.strip() for col in df_norm.columns]

# Crear diccionario y normalizar nombres
dic_normalizacion = dict(zip(df_norm["NOMBRE A NORMALIZAR"], df_norm["NOMBRE CORTO"]))
col_origen = 'Propietario (Oportunidad de Origen) (Oportunidad)'
df_cupones_open['NOMBRE_NORMALIZADO'] = df_cupones_open[col_origen].map(dic_normalizacion)

# Semana comercial OBS (usando función ya definida)
semana_actual = obtener_semana_comercial(datetime.today())    

# Filtrar asesores AC activos de la semana actual
df_ac = df_estructura[
    (df_estructura["ACTIVO"] == "ACTIVO") &
    (df_estructura["ROL"] == "AC") &
    (df_estructura["AÑO-MES-SEM"] == semana_actual)
][["NOMBRE CORTO", "EQUIPO HISTORICO AC"]].copy()

# Incluir JE aunque estén inactivos
df_je = df_estructura[
    (df_estructura["ROL"] == "JE") &
    (df_estructura["AÑO-MES-SEM"] == semana_actual)
][["NOMBRE CORTO", "EQUIPO HISTORICO AC"]].copy()

# Unir asesores AC y jefes JE
df_activos = pd.concat([df_ac, df_je], ignore_index=True)


df_activos["EQUIPO_REAP"] = df_activos["EQUIPO HISTORICO AC"].str.strip().apply(lambda x: f"Equipo_{x}")
df_activos = df_activos.rename(columns={"NOMBRE CORTO": "NOMBRE_NORMALIZADO"})

# Reaperturas (REAP)
df_reap = df_cupones_open[df_cupones_open['TIPO_REPARTO'] == 'REAP'].copy()
df_reap = df_reap.merge(df_activos, on="NOMBRE_NORMALIZADO", how="left")

# Dividir REAP válidas y no válidas
df_reap_validas = df_reap[df_reap['EQUIPO_REAP'].notna()].copy()
df_reap_invalidas = df_reap[df_reap['EQUIPO_REAP'].isna()].copy()

# Asignar equipo final en válidas
df_reap_validas['EQUIPO_FINAL'] = df_reap_validas['EQUIPO_REAP']

# Tratar reaperturas inválidas como FRESH
df_reap_invalidas = df_reap_invalidas.copy()
df_reap_invalidas['TIPO_REPARTO'] = 'FRESH'
df_reap_invalidas['PAIS_NORM'] = df_reap_invalidas['Agrupación OBS']
df_reap_invalidas['PROGRAMA_NORM'] = df_reap_invalidas['Programa de Interes']

# Verificar columnas requeridas
columnas_requeridas = ['AREA', 'PILAR_NORM', 'Agrupación OBS', 'Programa de Interes']
for col in columnas_requeridas:
    if col not in df_reap_invalidas.columns:
        raise ValueError(f"Falta la columna requerida: {col}")

# --- LIMPIEZA DE DUPLICADOS PARA EVITAR _x / _y ---
def limpiar_columnas_duplicadas(df):
    for col in df.columns:
        if col.endswith('_x') and col[:-2] in df.columns:
            df.drop(columns=[col], inplace=True)
        elif col.endswith('_y') and col[:-2] in df.columns:
            df.drop(columns=[col], inplace=True)
        elif col.endswith('_x'):
            df.rename(columns={col: col[:-2]}, inplace=True)
        elif col.endswith('_y'):
            df.rename(columns={col: col[:-2]}, inplace=True)
    return df

df_cupones_open = limpiar_columnas_duplicadas(df_cupones_open)
df_reap_invalidas = limpiar_columnas_duplicadas(df_reap_invalidas)

# --- HORA DE CORTE (bloque separado) ---
# Ajusta tu hora de corte. Si no quieres usarla hoy, pon cutoff_dt = None

# Lee exactamente la celda L4 de la hoja TTAA de OBS-DISTRI ATENEA
ruta_distri_og = home + r"\Grupo Planeta\BI POWER - General\Distribución Atenea\Codigos\Distribución OBS\SUDOKU.xlsx"
xls_distri_og = pd.ExcelFile(ruta_distri_og)
_cutoff_raw = pd.read_excel(
    xls_distri_og,
    sheet_name='Estatus diario',
    header=None,
    usecols='O',   # columna O
    skiprows=21,     # salta las 4 primeras filas -> fila 5
    nrows=1         # solo 1 fila
).iat[0, 0]

cutoff_dt = pd.to_datetime(_cutoff_raw, errors='coerce', dayfirst=True)
if pd.isna(cutoff_dt):
    raise ValueError("Cutoff inválido en DISTRI!L4 (esperaba fecha/hora).")

print(f"Hora de corte: {cutoff_dt}")

df_corte = df_cupones_open.iloc[0:0].copy()  # df vacío con mismas columnas, por si no hay corte


if cutoff_dt is not None:
    # Parse robusto (dd/mm/yyyy hh:mm soportado)
    fechas_open = pd.to_datetime(df_cupones_open['Fecha de creación'], errors='coerce', dayfirst=True)

    # Solo FRESH anteriores al corte
    mask_cutoff = (
        df_cupones_open['TIPO_REPARTO'].eq('FRESH')
        & fechas_open.notna()
        & (fechas_open < cutoff_dt)
    )

    # Guardar aparte para histórico (mantener equipo tal cual)
    df_corte = df_cupones_open.loc[mask_cutoff].copy()

    # Asegurar EQUIPO_FINAL: si faltara, usar Propietario
    if 'EQUIPO_FINAL' not in df_corte.columns:
        df_corte['EQUIPO_FINAL'] = df_corte['Propietario']
    else:
        df_corte['EQUIPO_FINAL'] = df_corte['EQUIPO_FINAL'].fillna(df_corte['Propietario'])

    # (Opcional) marca de trazabilidad
    df_corte['REAP_REASON'] = 'CUTOFF'
    df_corte['TIPO_REPARTO'] = 'REAP'

    # Sacarlos del circuito de FRESH (no se redistribuyen)
    df_cupones_open.loc[mask_cutoff, 'TIPO_REPARTO'] = 'REAP'
# --- Corte también para REAP inválidas (que ahora son FRESH) ---
if (cutoff_dt is not None) and (not df_reap_invalidas.empty) and ('Fecha de creación' in df_reap_invalidas.columns):
    fechas_inv = pd.to_datetime(df_reap_invalidas['Fecha de creación'], errors='coerce', dayfirst=True)
    mask_cutoff_inv = fechas_inv.notna() & (fechas_inv < cutoff_dt)

    if mask_cutoff_inv.any():
        # Estos NO deben redistribuirse: se van al histórico conservando equipo
        df_corte_inv = df_reap_invalidas.loc[mask_cutoff_inv].copy()

        # Asegurar EQUIPO_FINAL
        if 'EQUIPO_FINAL' not in df_corte_inv.columns:
            df_corte_inv['EQUIPO_FINAL'] = df_corte_inv['Propietario']
        else:
            df_corte_inv['EQUIPO_FINAL'] = df_corte_inv['EQUIPO_FINAL'].fillna(df_corte_inv['Propietario'])

        df_corte_inv['REAP_REASON'] = 'CUTOFF'
        df_corte_inv['TIPO_REPARTO'] = 'REAP'

        # Acumular en df_corte para anexarlo luego a df_hist_total
        df_corte = pd.concat([df_corte, df_corte_inv], ignore_index=True)

        # Quitar de df_reap_invalidas para que NO entren a df_fresh
        df_reap_invalidas = df_reap_invalidas.loc[~mask_cutoff_inv].copy()



# Unir con FRESH originales
df_fresh = pd.concat([
    df_cupones_open[df_cupones_open['TIPO_REPARTO'] == 'FRESH'],
    df_reap_invalidas
], ignore_index=True)

df_fresh = df_fresh[df_fresh['AREA'].isin(['A', 'B', 'C', 'T', 'E'])].copy()

# Crear clave compuesta
df_fresh['CLAVE'] = df_fresh['AREA'] + '__' + df_fresh['PILAR_NORM'].fillna('') + '__' + df_fresh['Agrupación OBS'].fillna('') + '__' + df_fresh['Programa de Interes'].fillna('')

# Asignar índice único para cada cupón
df_fresh = df_fresh.reset_index(drop=True)
df_fresh['ID'] = df_fresh.index.astype(str)

# --- Preparar histórico acumulado ---
df_hist_qbCN['CLAVE'] = df_hist_qbCN['AREA'] + '__' + df_hist_qbCN['PILAR_NORM'].fillna('') + '__' + df_hist_qbCN['Agrupación OBS'].fillna('') + '__' + df_hist_qbCN['Programa de Interes'].fillna('')
df_hist_qbCN['EQUIPO_FINAL'] = df_hist_qbCN['Equipo de Ventas (Usuario propietario) (Usuario)']

df_reap_validas['CLAVE'] = df_reap_validas['AREA'] + '__' + df_reap_validas['PILAR_NORM'].fillna('') + '__' + df_reap_validas['Agrupación OBS'].fillna('') + '__' + df_reap_validas['Programa de Interes'].fillna('')
if not df_corte.empty:
    df_corte['CLAVE'] = df_corte['AREA'] + '__' + df_corte['PILAR_NORM'].fillna('') + '__' + df_corte['Agrupación OBS'].fillna('') + '__' + df_corte['Programa de Interes'].fillna('')

df_hist_total = pd.concat([df_hist_qbCN, df_reap_validas, df_corte], ignore_index=True)
df_hist_total = df_hist_total[df_hist_total['EQUIPO_FINAL'].notna()]
print("Reaperturas  válidas:", df_reap_validas.shape[0])
print("Reaperturas  inválidas:", df_reap_invalidas.shape[0])
if cutoff_dt is not None:
    print("Corte horario (promovidos a histórico):", df_corte.shape[0])




# ## CALCULO CADENCIAS PREVIAS

# In[13]:


# ——— 1. Resumen histórico de cupones por equipo ———
# Asumiendo que ya tienes df_hist_total listo y el diccionario equipos_area:
col_equipo = 'EQUIPO_FINAL'
# 1.1) Lista única de equipos que aparezcan en df_hist_qbCN
equipos_unicos = (
    df_hist_total
    [col_equipo]
    .astype(str)
    .str.strip()
    .unique()
)

tabla = []
for equipo in equipos_unicos:
    # Area A: MST + ESP + cupón de ese equipo
    total_A = df_hist_total[
        (df_hist_total['TIPO']=='MST') &
        (df_hist_total['IDIOMA']=='ESP') &
        (df_hist_total[col_equipo]==equipo) 

    ].shape[0]
    # Transversal T: MBA + ESP
    total_T = df_hist_total[
        (df_hist_total['TIPO']=='MBA') &
        (df_hist_total['IDIOMA']=='ESP') &
        (df_hist_total[col_equipo]==equipo) 

    ].shape[0]
    # Inglés E: cualquier tipo + ENG
    total_E = df_hist_total[
        (df_hist_total['IDIOMA']=='ENG') &
        (df_hist_total[col_equipo]==equipo) 

    ].shape[0]

    tabla.append({
        'EQUIPO': equipo,
        'cupones_A': total_A,
        'cupones_T': total_T,
        'cupones_E': total_E
    })

df_resumen_hist = pd.DataFrame(tabla).fillna(0)
df_resumen_hist = df_resumen_hist.sort_values('EQUIPO').reset_index(drop=True)

# ——— 2. Resumen inicial df_cupones ———
# Asegurarnos de tener esta lista con todos los equipos de ventas relevantes
equipos = [
    'Equipo_A1', 'Equipo_A2',
    'Equipo_B1', 'Equipo_B2',
    'Equipo_C1', 'Equipo_C2'
]

# Inicializar tabla resumen
resumen = []

for equipo in equipos:
    registro = {'EQUIPO': equipo}

    # Área A, B, C → tipo MST + idioma ESP (independiente del área del cupón)
    if equipo.startswith('Equipo_A'):
        total_A = df_fresh[
            (df_fresh['TIPO'] == 'MST') &
            (df_fresh['IDIOMA'] == 'ESP') &
            (df_fresh['Propietario'] == equipo)
        ].shape[0]
        registro['cupones_A'] = total_A
    if equipo.startswith('Equipo_B'):
        total_B = df_fresh[
            (df_fresh['TIPO'] == 'MST') &
            (df_fresh['IDIOMA'] == 'ESP') &
            (df_fresh['Propietario'] == equipo)
        ].shape[0]
        registro['cupones_B'] = total_B
    if equipo.startswith('Equipo_C'):
        total_C = df_fresh[
            (df_fresh['TIPO'] == 'MST') &
            (df_fresh['IDIOMA'] == 'ESP') &
            (df_fresh['Propietario'] == equipo)
        ].shape[0]
        registro['cupones_C'] = total_C

    # Área T → solo si el cupón es del área T, tipo MBA, idioma ESP
    total_T = df_fresh[
        (df_fresh['AREA'] == 'T') &
        (df_fresh['TIPO'] == 'MBA') &
        (df_fresh['IDIOMA'] == 'ESP') &
        (df_fresh['Propietario'] == equipo)
    ].shape[0]
    registro['cupones_T'] = total_T

    # Área E → solo si el cupón es del área E, tipo MST o MBA, idioma ENG
    total_E = df_fresh[
        (df_fresh['AREA'] == 'E') &
        (df_fresh['TIPO'].isin(['MST', 'MBA'])) &
        (df_fresh['IDIOMA'] == 'ENG') &
        (df_fresh['Propietario'] == equipo)
    ].shape[0]
    registro['cupones_E'] = total_E

    resumen.append(registro)
print(resumen)
# Crear DataFrame resumen
df_resumen_final = pd.DataFrame(resumen).fillna(0)
df_resumen_final = df_resumen_final.sort_values('EQUIPO').reset_index(drop=True)
# Reordenar columnas si lo deseas
columnas_orden = ['EQUIPO', 'cupones_A', 'cupones_B', 'cupones_C', 'cupones_T', 'cupones_E']
df_resumen_final = df_resumen_final.reindex(columns=columnas_orden, fill_value=0)

df_resumen_init = df_resumen_final.copy()
# Suma total inicial por equipo:
df_resumen_init['CUPONES_INIT'] = df_resumen_init[['cupones_A','cupones_B','cupones_C']].sum(axis=1)
df_resumen_init['CUPONES_T'] = df_resumen_init[['cupones_T']].sum(axis=1)

# ——— 3. Juntar histórico + inicial ———
df_counts = (
    pd.merge(df_resumen_hist[['EQUIPO','cupones_A','cupones_T']],
             df_resumen_init[['EQUIPO','CUPONES_INIT','cupones_T']],
             on='EQUIPO', how='outer')
      .fillna(0)
)
# Total preliminar
df_counts['CUPONES_PRELIM_A'] = df_counts['CUPONES_INIT'] + df_counts['cupones_A']
df_counts['CUPONES_PRELIM_T']= df_counts['cupones_T_x']+df_counts['cupones_T_y']

# ——— 4. Cadencia preliminar ———
# Asumiendo que tu df_horas_eq tiene columnas ['EQUIPO','HORAS']
df_cad = pd.merge(df_horas_eq[['EQUIPO','HORAS']],
                  df_counts[['EQUIPO','CUPONES_PRELIM_A','CUPONES_PRELIM_T']],
                  on='EQUIPO', how='left').fillna(0)

df_cad['CAD_PRELIM_A'] = df_cad.apply(
    lambda r: r['CUPONES_PRELIM_A']/(r['HORAS']/6) if r['HORAS']>0 else 0,
    axis=1
)
df_cad['CAD_PRELIM_T'] = df_cad.apply(
    lambda r: r['CUPONES_PRELIM_T']/(r['HORAS']/6) if r['HORAS']>0 else 0,
    axis=1
)
print(df_cad)
# Convertir a diccionario: Cadencia preliminar A y T
cad_prelim_A_dict = df_cad.set_index('EQUIPO')['CAD_PRELIM_A'].to_dict()
cad_prelim_T_dict = df_cad.set_index('EQUIPO')['CAD_PRELIM_T'].to_dict()


#------------------------------------------------------------------------------------------

# 1. Cadencia teórica para A (filas 0 y 1)
num_A = df_cad.iloc[0:2]['CUPONES_PRELIM_A'].sum()
den_A = df_cad.iloc[0:2]['HORAS'].sum() / 6
cad_teo_A = num_A / den_A if den_A > 0 else 0

# 2. Cadencia teórica para B (filas 2 y 3)
num_B = df_cad.iloc[2:4]['CUPONES_PRELIM_A'].sum()
den_B = df_cad.iloc[2:4]['HORAS'].sum() / 6
cad_teo_B = num_B / den_B if den_B > 0 else 0

# 3. Cadencia teórica para C (filas 4 y 6)
#    Ajusta el segundo índice si tienes tres equipos C (por ejemplo df_cad.iloc[4:7])
num_C = df_cad.iloc[4:6]['CUPONES_PRELIM_A'].sum()
den_C = df_cad.iloc[4:6]['HORAS'].sum() / 6
cad_teo_C = num_C / den_C if den_C > 0 else 0

# 4. Cadencia teórica para T (toda la columna)
num_T = df_cad['CUPONES_PRELIM_T'].sum()
den_T = df_cad['HORAS'].sum() / 6
cad_teo_T = num_T / den_T if den_T > 0 else 0

# Diccionario de cadencias teóricas por área
cad_teo_map = {
    'A': cad_teo_A,
    'B': cad_teo_B,
    'C': cad_teo_C,
    'T': cad_teo_T,
    'E': cad_teo_T  # O ajusta según tu lógica para E
}

# 5. Mostrar resultados
print(f"Cadencia teórica A: {cad_teo_A:.2f} cupones por bloque de 6h")
print(f"Cadencia teórica B: {cad_teo_B:.2f} cupones por bloque de 6h")
print(f"Cadencia teórica C: {cad_teo_C:.2f} cupones por bloque de 6h")
print(f"Cadencia teórica T: {cad_teo_T:.2f} cupones por bloque de 6h")


# ## PARTE I --- MODELO OPTIMIZACIÓN CADENCIA/PRESENCIALIDAD
# 

# Qué es MILP y qué son los *slacks* (resumen para este proyecto)
# 
#  ¿Qué es un MILP?
# **MILP (Mixed-Integer Linear Programming)** es un modelo de optimización:
# - **Objetivo lineal**: minimizar (o maximizar) una suma ponderada de variables.
# - **Restricciones lineales**: igualdades o desigualdades lineales.
# - **Variables mixtas**: algunas son **enteras/binarias** (p. ej. `x_{cupón, equipo} ∈ {0,1}`), otras **continuas**.
# 
# En nuestro caso:
# - Variables binarias `x_{c,t}` indican si el **cupón c** va al **equipo t**.
# - El objetivo penaliza desviaciones respecto a **metas por pilar** y (blandamente) desviaciones de bandas en **Web** y **Buscadores**.
# - Usamos el solver **CBC** a través de `PuLP`.
# 
# ---
# 
#  ¿Qué son los *slacks*?
# Los **slacks** son variables de “holgura” que miden **cuánto te saltas** una restricción cuando la tratas **blanda** (no dura).
# - Si una condición debe estar dentro de un margen (**banda blanda**), el slack recoge el exceso fuera de ese margen.
# - Ese exceso se **penaliza en el objetivo**, así el solver “prefiere” cumplir la banda, pero, si es imposible, la viola **pagando multa** antes de volver el problema infactible.
# 
# Aquí usamos slacks en:
# - **Web** y **Buscadores**: bandas **blandas** ±`pilar_band_*` del **estimado acumulado** por equipo (`s_web_pos/neg`, `s_bsc_pos/neg`).
# - **Desvío por pilar** (en general): `diff_{t,p}` mide distancia del acumulado real al **estimado por pilar**.
# 
# **Nota:** La **cadencia global exacta** se fuerza sin slack (igualdad exacta de FRESH por equipo).
# 
# ---
# 
#  Cómo se aplica en este modelo
# - **Restricciones duras** (sin slack):
#   - Cada cupón se asigna a **un único** equipo.
#   - **FRESH exacto** por equipo (cadencia global exacta → el acumulado por equipo se alinea).
# - **Restricciones blandas** (con slack + penalización):
#   - Ajuste por **pilar** en el acumulado: `diff_{t,p}`.
#   - **Bandas** para **Web** y **Buscadores** por equipo: `± pilar_band_*` con slacks `s_*`.
# 
# ---
# 
#  Parámetros clave para afinar
# - **`pilar_band_web` / `pilar_band_busc`** (p. ej. 0.05): ancho de la **banda blanda** (±5%) sobre el **estimado acumulado** por equipo en Web/Buscadores.
#   - Más pequeño ⇒ más precisión (puede tardar más o necesitar slacks > 0).
# - **Pesos del objetivo**:
#   - `BIG_WEB` / `BIG_BUSC`: **multa muy alta** si Web/Busc salen de su banda.
#   - `PENAL_DIFF['Web'|'Buscadores']`: penaliza **cualquier** desviación al estimado por pilar.
#   - `PENAL_DIFF['Redes Sociales'|'P.Verticales']`: **muy bajos** para usarlos de “pulmón”.
# - **`time_limit`** del solver: tiempo máximo de búsqueda (segundos).
# 
# ---
# 
#  Guía rápida de ajuste (si algo no cuadra)
# 1. **Web/Busc con poco ajuste**:
#    - Baja `pilar_band_*` (ej. 0.05 → 0.03).
#    - Sube `BIG_WEB` / `BIG_BUSC` (ej. 5000 → 8000, 3000 → 6000).
# 2. **Solver lento o “TimeLimit” con gaps grandes**:
#    - Sube `time_limit` (ej. 90 → 120/180).
#    - O afloja `pilar_band_*` (ej. 0.03 → 0.05).
# 3. **RS / P.Verticales desbalanceados**:
#    - Sube **ligeramente** `PENAL_DIFF['Redes Sociales']` o `['P.Verticales']` (p. ej. 0.1 → 0.2).
#    - (Opcional) añadir banda blanda superior para RS si quieres limitar extremos.
# 
# ---
# 
#  Diagnóstico exprés (qué mirar en cada run)
# - **Cadencia global**: suma de FRESH por equipo = F del área (siempre exacta).
# - **Web/Busc**:
#   - Chequear slacks `s_web_*` / `s_bsc_*`: **0** ⇒ dentro de banda; **>0** ⇒ violación (pagando multa).
#   - Si hay muchos slacks > 0, subir `time_limit` o aflojar `pilar_band_*` un punto.
# - **Tiempos**: si suele llegar a `TimeLimit` y la solución te vale ⇒ OK; si no, ajustar parámetros. Actualmente está en 120s el TimeLimit
# 
# ---
# 

# In[14]:


# --- Constantes y utilidades robustas ---
PILLARS = ['Web', 'Buscadores', 'P.Verticales', 'Redes Sociales']

def _force_pillars_columns(m, index=None):
    """
    Garantiza que 'm' (equipo x pilar) tenga SIEMPRE las 4 columnas canónicas.
    - Suma columnas duplicadas si las hubiera.
    - Reindexa columnas a PILLARS con fill_value=0.
    - Si se pasa 'index', reindexa también filas a ese índice.
    """
    if m is None or getattr(m, "empty", True):
        base = pd.DataFrame(0, columns=PILLARS, index=(index if index is not None else []))
        return base
    # Sumar duplicadas tras normalizaciones
    cols = m.columns
    if hasattr(cols, "duplicated") and cols.duplicated().any():
        m = m.groupby(level=0, axis=1).sum()
    if index is not None:
        m = m.reindex(index=index, fill_value=0)
    return m.reindex(columns=PILLARS, fill_value=0)


# ================== NUEVO: helpers de redondeo robusto ==================
def _sanitize_nonneg_series(x):
    """Devuelve Series no negativa y sin NaNs, preservando índice."""
    if isinstance(x, np.ndarray):
        x = pd.Series(x)
    x = pd.Series(x, index=getattr(x, "index", None))
    x = x.fillna(0.0).astype(float).clip(lower=0.0)
    return x

def ajustar_redondeo_sum_exacta(target_float: pd.Series, total: int) -> pd.Series:
    """
    Redondea un vector no negativo (Series) a enteros no negativos cuya suma es EXACTAMENTE 'total'.
    Estrategia:
      1) Sanitiza y, si la suma>0, reescala a suma 'total'.
      2) Floor y asignación por Largest Remainder (Hamilton).
      3) Corrección fina por posibles ±1 de coma flotante.
    Retorna una Series de int con el mismo índice.
    """
    s = _sanitize_nonneg_series(target_float)
    T = int(total)

    # Casos borde
    if T <= 0 or len(s) == 0 or s.sum() == 0.0:
        return pd.Series(0, index=s.index, dtype=int)

    # Reescalar para que la suma sea exactamente T (antes del redondeo)
    factor = T / s.sum()
    s_scaled = s * factor

    floors = np.floor(s_scaled).astype(int)
    residual = T - int(floors.sum())

    # Partes fraccionales
    frac = s_scaled - floors
    order_asc = np.argsort(frac.values)          # menor → mayor
    order_desc = order_asc[::-1]                 # mayor → menor

    alloc = floors.copy()

    if residual > 0:
        # Añadir +1 a los de fracción más grande
        take = order_desc[:residual]
        alloc[take] += 1
    elif residual < 0:
        # Quitar 1 empezando por los de fracción más pequeña (sin ir a negativo)
        need = -residual
        for idx in order_asc:
            if need == 0:
                break
            if alloc[idx] > 0:
                alloc[idx] -= 1
                need -= 1
        # Si aún queda por quitar (raro), retirar de los mayores fraccionales con >0
        i = 0
        while need > 0 and i < len(order_desc):
            j = order_desc[i]
            if alloc[j] > 0:
                alloc[j] -= 1
                need -= 1
            i += 1

    result = pd.Series(alloc, index=s.index, dtype=int)

    # Corrección ultra-fina por si queda ±1
    diff = T - int(result.sum())
    if diff != 0:
        if diff > 0:
            for idx in order_desc:
                if diff == 0:
                    break
                result.iloc[idx] += 1
                diff -= 1
        else:
            need = -diff
            for idx in order_asc:
                if need == 0:
                    break
                if result.iloc[idx] > 0:
                    result.iloc[idx] -= 1
                    need -= 1

    # Garantía final de no-negatividad y suma exacta
    result = result.clip(lower=0)
    # Si por alguna rareza flotante todavía no cierra, ajustar en el primer índice disponible
    delta = T - int(result.sum())
    if delta != 0 and len(result) > 0:
        if delta > 0:
            result.iloc[0] += delta
        else:
            # quitar sin ir a negativo
            for i in range(len(result)):
                q = min(result.iloc[i], -delta)
                result.iloc[i] -= q
                delta += q
                if delta == 0:
                    break
    return result.astype(int)
# =======================================================================


def distribuir_area_X(
    df_fresh, df_hist_total, df_horas_eq, df_pesos_areas,
    cad_prelim_dict, cad_teo, df_reap_validas,
    sigla_area, equipos_X,
    pilar_band_web=0.05,     # ±5% alrededor del estimado ACUMULADO de Web por equipo (banda blanda)
    pilar_band_busc=0.05,    # ±5% alrededor del estimado ACUMULADO de Buscadores por equipo (banda blanda)
    time_limit=90
):
    """
    Cadencia GLOBAL exacta por equipo (igualando #FRESH del día) y Web/Buscadores muy ajustados
    al estimado ACUMULADO (redes/p.verticales flexibles).
    El ESTIMADO por pilar se calcula como: (total pilar hist+fresh) * (PESO_BASE normalizado del equipo).
    """

    pillars = PILLARS[:]  
    # ---------- Helper: estimado puro por PESO_BASE ----------
    def make_estimado_pilar_base(df_hist_area, df_fresh_area, df_pesos_areas, equipos_X, pillars, sigla_area, share_override=None):
        """
        Estimado por pilar = Total ACUMULADO del pilar (hist+fresh) * % del equipo en el área.
        % por equipo viene de PESO_BASE (normalizado) o de share_override (Series que suma 1).
        """
        equipos_X = list(map(str, equipos_X))

        # Totales ACUMULADOS por pilar
        tot_hist_p = (df_hist_area.groupby('PILAR_NORM').size()
                      .reindex(pillars, fill_value=0).astype(int))
        tot_fresh_p = df_fresh_area['PILAR_NORM'].value_counts().reindex(pillars, fill_value=0).astype(int)
        T = (tot_hist_p + tot_fresh_p)  # Series por pilar

        # % por equipo
        if share_override is not None:
            s = pd.Series(share_override, index=equipos_X, dtype=float).fillna(0.0)
        else:
            pesos = (df_pesos_areas[df_pesos_areas['AREA'] == sigla_area]
                     .set_index('EQUIPO')['PESO_BASE']
                     .reindex(equipos_X).astype(float).fillna(0.0))
            total_pesos = float(pesos.sum())
            s = pesos / (total_pesos if total_pesos > 0 else 1.0)

        # fallback uniforme si todo es 0
        if not np.isfinite(s.to_numpy()).all() or s.sum() <= 0:
            s = pd.Series(1.0 / len(equipos_X), index=equipos_X)

        # Estimado float por pilar = T[p] * s
        est_float = pd.DataFrame(index=equipos_X, columns=pillars, dtype=float)
        for p in pillars:
            est_float[p] = float(T[p]) * s

        # Redondeo por "largest remainder" por columna para cuadrar T[p]
        est_floor = np.floor(est_float).astype(int)
        est = est_floor.copy()

        for p in pillars:
            need = int(T[p] - est[p].sum())  # cuánto falta para llegar al total exacto
            if need > 0:
                frac = (est_float[p] - est_floor[p]).reset_index()
                frac.columns = ['EQUIPO_FINAL', 'frac']
                frac = frac.sort_values(['frac', 'EQUIPO_FINAL'], ascending=[False, True])
                for eq in frac['EQUIPO_FINAL'].iloc[:need]:
                    est.at[eq, p] += 1
            elif need < 0:
                frac = (est_float[p] - est_floor[p]).reset_index()
                frac.columns = ['EQUIPO_FINAL', 'frac']
                frac = frac.sort_values(['frac', 'EQUIPO_FINAL'], ascending=[True, True])
                k = 0
                for eq in frac['EQUIPO_FINAL']:
                    if est.at[eq, p] > 0:
                        est.at[eq, p] -= 1
                        k += 1
                        if k == -need:
                            break

        return est.reindex(index=equipos_X, columns=pillars, fill_value=0).astype(int)

    # ---------- Ajuste fino por acumulado (swaps) ----------
    def ajuste_fino_cadencia_acum(df_fresh_area, estimado_pilar, df_hist_area):
        """Swaps 1↔1 priorizando Web/Buscadores; compensación con RS o P.Verticales. Evalúa ACUMULADO (hist+fresh)."""
        df = df_fresh_area.copy()
        pilares_clave = ['Web', 'Buscadores']
        pilares_compensa = ['Redes Sociales', 'P.Verticales']
        equipos = list(map(str, equipos_X))  # incluir equipos aunque no tengan FRESH

        def matriz_acum(dframe):
            m = (dframe.groupby(['EQUIPO_FINAL','PILAR_NORM']).size()
                          .unstack(fill_value=0))
            return _force_pillars_columns(m, index=equipos)

        acum = pd.concat([
            df_hist_area[['EQUIPO_FINAL','PILAR_NORM']],
            df[['EQUIPO_FINAL','PILAR_NORM']]
        ], ignore_index=True)
        count_acum = matriz_acum(acum)
        est_safe = estimado_pilar.reindex(index=equipos, columns=pillars, fill_value=0)
        delta = (count_acum - est_safe).fillna(0).astype('Int64')

        movimientos = 0
        for pilar in pilares_clave:
            while True:
                delta_p = delta[pilar]
                exceso_eq = delta_p.idxmax()
                falta_eq  = delta_p.idxmin()
                if int(delta_p.get(exceso_eq,0)) <= 0 or int(delta_p.get(falta_eq,0)) >= 0:
                    break

                cupones_exceso = df[(df['EQUIPO_FINAL']==exceso_eq) & (df['PILAR_NORM']==pilar)]
                if cupones_exceso.empty:
                    break

                realizado = False
                for p_comp in pilares_compensa:
                    cupones_comp = df[(df['EQUIPO_FINAL']==falta_eq) & (df['PILAR_NORM']==p_comp)]
                    if cupones_comp.empty:
                        continue
                    for i in cupones_exceso.index:
                        if realizado: break
                        for j in cupones_comp.index:
                            if i == j: continue
                            ei, ej = df.at[i,'EQUIPO_FINAL'], df.at[j,'EQUIPO_FINAL']
                            df.at[i,'EQUIPO_FINAL'], df.at[j,'EQUIPO_FINAL'] = ej, ei

                            acum2 = pd.concat([
                                df_hist_area[['EQUIPO_FINAL','PILAR_NORM']],
                                df[['EQUIPO_FINAL','PILAR_NORM']]
                            ], ignore_index=True)
                            new_count = matriz_acum(acum2)
                            new_delta = (new_count - est_safe).fillna(0).astype('Int64')

                            if int(new_delta.abs().to_numpy().sum()) < int(delta.abs().to_numpy().sum()):
                                delta = new_delta
                                movimientos += 1
                                realizado = True
                                break
                            else:
                                df.at[i,'EQUIPO_FINAL'], df.at[j,'EQUIPO_FINAL'] = ei, ej
                    if realizado:
                        break
                if not realizado:
                    break

        print(f"\n🔁 Ajuste fino (ACUM) completado. Movimientos: {movimientos}")
        return df

    # ---------- Filtros base ----------
    df_fresh_area = df_fresh[(df_fresh['AREA'] == sigla_area) & (df_fresh['PILAR_NORM'].isin(pillars))].copy()
    df_fresh_area = df_fresh_area.reset_index(drop=True)
    coupons = df_fresh_area.index.tolist()

    # Alinear histórico al mix TIPO/IDIOMA del día
    filt_idiomas = df_fresh_area['IDIOMA'].dropna().unique().tolist()
    filt_tipos   = df_fresh_area['TIPO'].dropna().unique().tolist()
    df_hist_area = df_hist_total[
        (df_hist_total['EQUIPO_FINAL'].isin(equipos_X)) &
        (df_hist_total['PILAR_NORM'].isin(pillars)) &
        (~df_hist_total['AREA'].isin(['T','E'])) &
        (df_hist_total['IDIOMA'].isin(filt_idiomas)) &
        (df_hist_total['TIPO'].isin(filt_tipos))
    ].copy()

    # Horas y bloques
    horas_area = (
        df_horas_eq[df_horas_eq['EQUIPO'].isin(equipos_X)]
        .set_index('EQUIPO')['HORAS'].reindex(equipos_X)
    )
    bloques = (horas_area / 6).fillna(0).astype(float)

    # Hist por equipo
    cupones_hist = (
        df_hist_area.groupby('EQUIPO_FINAL').size()
        .reindex(equipos_X, fill_value=0).astype(int)
    )

    # ---------- Estimado por pilar (PESO_BASE) ----------
    estimado_pilar = make_estimado_pilar_base(
        df_hist_area=df_hist_area,
        df_fresh_area=df_fresh_area,
        df_pesos_areas=df_pesos_areas,
        equipos_X=equipos_X,
        pillars=pillars,
        sigla_area=sigla_area,
        share_override=None  # o pasa un dict/Series {'Equipo_A1':0.4,...} si quieres forzar %
    ).astype(int)

    # ---------- Cadencia GLOBAL EXACTA vía #FRESH por equipo ----------
    F = int(len(df_fresh_area))                        # fresh del día (fijo)
    cad_raw = (bloques * float(cad_teo)).fillna(0.0)   # objetivo en total ACUMULADO
    sum_obj = float(cad_raw.sum())
    total_acum_real = int(cupones_hist.sum() + F)      # ACUMULADO real tras asignar el día

    # Recentrar a total real
    scale = (total_acum_real / sum_obj) if sum_obj > 0 else 1.0
    cad_target = cad_raw * scale                       # objetivo total por equipo (float)

    # Convertir a objetivo de FRESH por equipo (ENTERO, suma F) usando helper ROBUSTO
    fresh_target_float = (cad_target - cupones_hist).clip(lower=0.0)
    fresh_target_int = ajustar_redondeo_sum_exacta(fresh_target_float, total=F)

    # Seguridad
    assert int(fresh_target_int.sum()) == F
    assert (fresh_target_int >= 0).all()

    # ---------- MILP ----------
    team_vars = {(c, t): pulp.LpVariable(f"x_{c}_{t}", cat="Binary")
                 for c in coupons for t in equipos_X}

    # Slacks por pilar (desviación absoluta al estimado ACUMULADO)
    diff_pilar = {(t, p): pulp.LpVariable(f"diff_{t}_{p}", lowBound=0)
                  for t in equipos_X for p in pillars}

    # Slacks para bandas BLANDAS en Web/Buscadores
    s_web_pos  = {t: pulp.LpVariable(f"s_web_pos_{t}",  lowBound=0) for t in equipos_X}
    s_web_neg  = {t: pulp.LpVariable(f"s_web_neg_{t}",  lowBound=0) for t in equipos_X}
    s_bsc_pos  = {t: pulp.LpVariable(f"s_bsc_pos_{t}",  lowBound=0) for t in equipos_X}
    s_bsc_neg  = {t: pulp.LpVariable(f"s_bsc_neg_{t}",  lowBound=0) for t in equipos_X}

    # Pesos objetivo
    PENAL_DIFF = {'Web': 50, 'Buscadores': 30, 'P.Verticales': 0.5, 'Redes Sociales': 0.01}
    BIG_WEB  = 5000.0
    BIG_BUSC = 3000.0

    prob = pulp.LpProblem(f"Distribucion_Area_{sigla_area}", pulp.LpMinimize)
    prob += (
        pulp.lpSum(PENAL_DIFF[p] * diff_pilar[(t, p)] for t in equipos_X for p in pillars) +
        BIG_WEB  * pulp.lpSum(s_web_pos[t] + s_web_neg[t]   for t in equipos_X) +   
        BIG_BUSC * pulp.lpSum(s_bsc_pos[t] + s_bsc_neg[t]   for t in equipos_X)
    )

    # 1) Cada cupón se asigna a 1 equipo
    for c in coupons:
        prob += pulp.lpSum(team_vars[(c, t)] for t in equipos_X) == 1

    # 2) FRESH exacto por equipo (cadencia global exacta)
    for t in equipos_X:
        prob += pulp.lpSum(team_vars[(c, t)] for c in coupons) == int(fresh_target_int.loc[t])

    # 3) Desvío por pilar vs estimado ACUMULADO + bandas blandas Web/Busc
    hist_map = df_hist_area.groupby(['EQUIPO_FINAL', 'PILAR_NORM']).size().to_dict()

    for t in equipos_X:
        for p in pillars:
            h = int(hist_map.get((t, p), 0))
            assign_tp = pulp.lpSum(team_vars[(c, t)] for c in coupons if df_fresh_area.at[c, 'PILAR_NORM'] == p)
            total_tp  = h + assign_tp
            est_tp    = int(estimado_pilar.loc[t, p])

            prob +=  total_tp - est_tp <= diff_pilar[(t, p)]
            prob +=  est_tp - total_tp <= diff_pilar[(t, p)]

        # Bandas BLANDAS Web
        hW = int(hist_map.get((t, 'Web'), 0))
        aW = pulp.lpSum(team_vars[(c, t)] for c in coupons if df_fresh_area.at[c, 'PILAR_NORM'] == 'Web')
        totW = hW + aW
        estW = int(estimado_pilar.loc[t, 'Web'])
        epsW = max(1, int(round(estW * float(pilar_band_web))))
        prob += totW - estW <=  epsW + s_web_pos[t]
        prob += estW - totW <=  epsW + s_web_neg[t]

        # Bandas BLANDAS Buscadores
        hB = int(hist_map.get((t, 'Buscadores'), 0))
        aB = pulp.lpSum(team_vars[(c, t)] for c in coupons if df_fresh_area.at[c, 'PILAR_NORM'] == 'Buscadores')
        totB = hB + aB
        estB = int(estimado_pilar.loc[t, 'Buscadores'])
        epsB = max(1, int(round(estB * float(pilar_band_busc))))
        prob += totB - estB <=  epsB + s_bsc_pos[t]
        prob += estB - totB <=  epsB + s_bsc_neg[t]

    print(f"🧩 Resolviendo modelo para área {sigla_area} (cadencia exacta + Web/Busc ajustados)...")
    status = prob.solve(pulp.PULP_CBC_CMD(msg=False, timeLimit=time_limit))
    print("📌 Estado del solver:", pulp.LpStatus[prob.status])

    # ---------- Recuperar asignaciones ----------
    asignaciones = []
    for c in coupons:
        asignado = None
        for t in equipos_X:
            val = pulp.value(team_vars[(c, t)])
            if val is not None and round(val) == 1:
                asignado = t
                break
        if asignado is None:
            # Fallback teórico (no debería pasar con igualdad exacta de FRESH por equipo)
            deficits = (fresh_target_int - 0).reindex(equipos_X).fillna(0)
            asignado = deficits.idxmax()
        asignaciones.append(asignado)
    df_fresh_area['EQUIPO_FINAL'] = asignaciones

    # ---------- Ajuste fino (ACUM) ----------
    df_fresh_area = ajuste_fino_cadencia_acum(df_fresh_area, estimado_pilar[pillars], df_hist_area)

    # ---------- Métricas y salida ----------
    df_real_fresh = (
        df_fresh_area.groupby(['EQUIPO_FINAL', 'PILAR_NORM']).size()
        .unstack(fill_value=0)
    )
    df_real_fresh = _force_pillars_columns(df_real_fresh, index=equipos_X)

    df_acum_real = pd.concat([df_hist_area, df_fresh_area], ignore_index=True)
    df_real_acum = (
        df_acum_real.groupby(['EQUIPO_FINAL', 'PILAR_NORM']).size()
        .unstack(fill_value=0)
    )
    df_real_acum = _force_pillars_columns(df_real_acum, index=equipos_X)

    # Cadencias de reporte
    bloques_rep = (df_horas_eq.set_index('EQUIPO').reindex(equipos_X)['HORAS'] / 6).replace(0, pd.NA)
    cad_pilar = df_real_fresh.div(bloques_rep, axis=0).astype(float).round(2)

    total_final = (
        df_fresh_area['EQUIPO_FINAL'].value_counts()
        .add(cupones_hist, fill_value=0)
        .reindex(equipos_X, fill_value=0)
    )
    cad_final = (total_final / bloques_rep).astype(float).round(4)

    df_cadencia = pd.DataFrame({
        "Equipo": equipos_X,
        "Cadencia Original": pd.Series({eq: cad_prelim_dict.get(eq, 0) for eq in equipos_X}).reindex(equipos_X).fillna(0).values,
        "Cadencia Final": cad_final.values,
        "Cadencia Teórica": [cad_teo] * len(equipos_X)
    })

    df_est = estimado_pilar.copy(); df_est.columns = [p + "_Estimado" for p in df_est.columns]
    df_real_a = df_real_acum.copy();  df_real_a.columns  = [p + "_Real_ACUM"  for p in df_real_a.columns]
    df_real_f = df_real_fresh.copy(); df_real_f.columns = [p + "_Real_FRESH" for p in df_real_f.columns]
    df_comparativa = pd.concat([df_est, df_real_a, df_real_f], axis=1)

    df_final_area = pd.concat([
        df_fresh_area,
        df_reap_validas[
            (df_reap_validas['EQUIPO_FINAL'].isin(equipos_X)) &
            (~df_reap_validas['AREA'].isin(['T', 'E']))
        ]
    ], ignore_index=True)

    print("\n🔎 Comparativa (Estimado / Real_ACUM / Real_FRESH) tras ajuste fino:")
    print(df_comparativa)
    print("\n📊 Cadencia por Pilar y Equipo (FRESH):")
    print(cad_pilar)
    print("\n📏 Verificación cadencias Área", sigla_area)
    print(df_cadencia)

    return df_final_area, df_cadencia, df_comparativa, cad_pilar


# ### AREA A

# In[15]:


equipos_A = ["Equipo_A1", "Equipo_A2"]

df_final_A, df_cadencia, df_comparativa, cad_pilar = distribuir_area_X(
    df_fresh,
    df_hist_total,
    df_horas_eq,
    df_pesos_areas,
    cad_prelim_A_dict,
    cad_teo_A,
    df_reap_validas,
    "A",
    equipos_A
)


# ### AREA B

# In[16]:


equipos_B = ["Equipo_B1", "Equipo_B2"]

df_final_B, df_cadencia, df_comparativa, cad_pilar = distribuir_area_X(
    df_fresh,
    df_hist_total,
    df_horas_eq,
    df_pesos_areas,
    cad_prelim_A_dict,
    cad_teo_B,
    df_reap_validas,                         
    "B",
    equipos_B                     
)
#Exportamos a Excel (prueba)
#df_final_B.to_excel("Distribucion_Final_Area_B.xlsx", index=False)


# ### AREA C

# In[17]:


equipos_C = ["Equipo_C1", "Equipo_C2"]

df_final_C, df_cadencia, df_comparativa, cad_pilar = distribuir_area_X(
    df_fresh,
    df_hist_total,
    df_horas_eq,
    df_pesos_areas,
    cad_prelim_A_dict,
    cad_teo_C,
    df_reap_validas,
    "C",
    equipos_C
)
#Exportamos a Excel (prueba)
#df_final_C.to_excel("Distribucion_Final_Area_C.xlsx", index=False)


# ### AREA T

# In[18]:


def distribuir_area_T(
    df_fresh, df_hist_total, df_horas_eq, df_pesos_areas,
    cad_prelim_T_dict, cad_teo_T, df_reap_validas,
    pilar_band_web=0.05,   # ±5% sobre el estimado ACUMULADO de Web por equipo (banda BLANDA)
    pilar_band_busc=0.05,  # ±5% sobre el estimado ACUMULADO de Buscadores por equipo (banda BLANDA)
    time_limit=90
):
    """
    Área T con:
    - Cadencia GLOBAL exacta por equipo (fresh por equipo fijado).
    - Split IG/XP variable: se calcula con la suma de PESO_BASE de los equipos IG vs XP en T.
      Se fuerza que el ACUMULADO (hist + reaps + fresh) cumpla esas cuotas (vía cuotas de fresh por directora).
    - Web/Buscadores muy ajustados al estimado ACUMULADO por equipo (bandas blandas con multa alta).
    - RS y P.Verticales como pulmón.
    El ESTIMADO por pilar = (hist+fresh) * (% equipo por PESO_BASE normalizado en T).
    Histórico T: TIPO=MBA, IDIOMA=ESP.
    """

    # --- Utils ---
    def _ajustar_redondeo_sum_exacta(target_float: pd.Series, total: int) -> pd.Series:
        """
        Redondeo robusto tipo 'largest remainder' garantizando que la suma == total,
        tolerando floats, NaNs e imposibilidades locales con ajuste final.
        """
        s = pd.Series(target_float).copy()
        s = s.astype(float).replace([np.inf, -np.inf], np.nan).fillna(0.0).clip(lower=0.0)
        s = s.reindex(sorted(s.index))  # orden estable determinista

        total = int(total)
        if total <= 0 or len(s) == 0:
            return pd.Series(0, index=s.index, dtype=int)

        floors = np.floor(s).astype(int)
        res = floors.copy()

        need = int(total - int(res.sum()))
        if need == 0:
            return res.clip(lower=0)

        frac = (s - floors).to_numpy()
        idxs = np.arange(len(s))

        if need > 0:
            order = idxs[np.argsort(-frac)]
            k = 0; L = max(1, len(order))
            while need > 0:
                i = order[k % L]
                res.iloc[i] += 1
                need -= 1
                k += 1
        else:
            take = -need
            order = idxs[np.argsort(frac)]  # menores fracciones primero
            k = 0; L = max(1, len(order)); safety = 0
            while take > 0 and safety < 5 * len(res):
                i = order[k % L]
                if res.iloc[i] > 0:
                    res.iloc[i] -= 1
                    take -= 1
                k += 1; safety += 1
            if take > 0:
                order2 = idxs[np.argsort(-res.to_numpy())]
                j = 0; L2 = max(1, len(order2)); safety2 = 0
                while take > 0 and safety2 < 5 * len(res):
                    i = order2[j % L2]
                    if res.iloc[i] > 0:
                        res.iloc[i] -= 1
                        take -= 1
                    j += 1; safety2 += 1

        final_sum = int(res.sum())
        if final_sum != total:
            diff = total - final_sum
            if diff > 0:
                for k in range(diff):
                    i = k % len(res)
                    res.iloc[i] += 1
            elif diff < 0:
                need_take = -diff
                k = 0; safety = 0
                while need_take > 0 and safety < 10 * len(res):
                    i = k % len(res)
                    if res.iloc[i] > 0:
                        res.iloc[i] -= 1
                        need_take -= 1
                    k += 1; safety += 1

        assert int(res.sum()) == total, "Ajuste robusto no logró igualar la suma al total."
        return res.clip(lower=0).astype(int)

    # Ajusta per-equipo los enteros para cumplir una cuota de grupo exacta (versión simple; ya no la usamos)
    def _forzar_cuota_grupo(fresh_int: pd.Series,
                            fresh_float: pd.Series,
                            grupo_add: list,  # a quién sumar (+1 cada vez)
                            grupo_sub: list,  # a quién restar (-1 cada vez)
                            diff: int) -> pd.Series:
        """Transfiere 'diff' unidades desde grupo_sub hacia grupo_add respetando fracciones."""
        if diff == 0:
            return fresh_int
        # Orden para sumar: mayor decimal pendiente en grupo_add
        orden_add = pd.Index(grupo_add)[np.argsort(-(fresh_float.loc[grupo_add] - fresh_int.loc[grupo_add]))]
        # Orden para quitar: menor decimal (o más exceso) en grupo_sub
        orden_sub = pd.Index(grupo_sub)[np.argsort((fresh_float.loc[grupo_sub] - fresh_int.loc[grupo_sub]))]

        res = fresh_int.copy()
        i_add = i_sub = 0
        steps = abs(int(diff))
        for _ in range(steps):
            while i_add < len(orden_add) and pd.isna(orden_add[i_add]):
                i_add += 1
            while i_sub < len(orden_sub) and (pd.isna(orden_sub[i_sub]) or res.at[orden_sub[i_sub]] <= 0):
                i_sub += 1
            if i_add >= len(orden_add) or i_sub >= len(orden_sub):
                break
            a = orden_add[i_add]; b = orden_sub[i_sub]
            if diff > 0:
                res.at[a] += 1; res.at[b] -= 1
            else:
                res.at[a] -= 1; res.at[b] += 1
            i_add = (i_add + 1) % len(orden_add) if len(orden_add) else i_add
            i_sub = (i_sub + 1) % len(orden_sub) if len(orden_sub) else i_sub
        return res

    # Fallback de asignación proporcional minimizando desviación a estimados por pilar
    def _fallback_asignacion(df_fresh_T, fresh_target_int, estimado_pilar, df_hist_T, equipos_T, pillars):
        df = df_fresh_T.copy()
        # acumulado actual (histórico)
        hist_map = df_hist_T.groupby(['EQUIPO_FINAL', 'PILAR_NORM']).size().unstack(fill_value=0)
        hist_map = hist_map.reindex(index=equipos_T, columns=pillars, fill_value=0)

        # contadores de asignación fresh en curso
        assigned = pd.DataFrame(0, index=equipos_T, columns=pillars, dtype=int)
        remaining = fresh_target_int.reindex(equipos_T).astype(int).fillna(0).to_dict()

        # Para cada cupón, elige el equipo con más reducción de error cuadrático hacia el estimado
        for c in df.index:
            p = df.at[c, 'PILAR_NORM']
            best_t, best_gain = None, None
            for t in equipos_T:
                if remaining.get(t, 0) <= 0:
                    continue
                before = hist_map.at[t, p] + assigned.at[t, p]
                after  = before + 1
                est    = int(estimado_pilar.loc[t, p]) if (t in estimado_pilar.index and p in estimado_pilar.columns) else 0
                # coste marginal (cuadrático)
                cost_before = (before - est)**2
                cost_after  = (after  - est)**2
                gain = cost_before - cost_after  # queremos maximizar la "ganancia"
                # desempate por mayor restante para no bloquear
                tie = remaining.get(t, 0)
                key = (gain, tie)
                if (best_gain is None) or (key > best_gain):
                    best_gain = key
                    best_t = t
            if best_t is None:
                # si todos 0, asigna al equipo con menor acumulado total
                sums = (hist_map + assigned).sum(axis=1)
                best_t = sums.idxmin()
            df.at[c, 'EQUIPO_FINAL'] = best_t
            assigned.at[best_t, p] += 1
            remaining[best_t] = remaining.get(best_t, 0) - 1

        return df

    # ========== Parámetros y filtros ==========

    pillars = PILLARS[:]  # ['Web','Buscadores','P.Verticales','Redes Sociales']
    equipos_T = ['Equipo_A1', 'Equipo_A2', 'Equipo_B1', 'Equipo_B2', 'Equipo_C1', 'Equipo_C2']

    # Map directoras
    IG_teams = {'Equipo_A1', 'Equipo_B1', 'Equipo_C1'}
    XP_teams = {'Equipo_A2', 'Equipo_B2', 'Equipo_C2'}

    # FRESH / HIST / REAPS (filtros T)
    df_fresh_T = df_fresh[(df_fresh['AREA'] == 'T') & (df_fresh['PILAR_NORM'].isin(pillars))].copy().reset_index(drop=True)
    coupons_T = df_fresh_T.index.tolist()

    df_hist_T = df_hist_total[
        (df_hist_total['TIPO'] == 'MBA') &
        (df_hist_total['IDIOMA'] == 'ESP') &
        (df_hist_total['EQUIPO_FINAL'].isin(equipos_T)) &
        (df_hist_total['PILAR_NORM'].isin(pillars))
    ].copy()

    df_reap_T = df_reap_validas[
        (df_reap_validas['AREA'] == 'T') &
        (df_reap_validas['TIPO'] == 'MBA') &
        (df_reap_validas['IDIOMA'] == 'ESP') &
        (df_reap_validas['EQUIPO_FINAL'].isin(equipos_T))
    ].copy()

    # Horas / bloques y hist por equipo
    horas_T = (
        df_horas_eq[df_horas_eq['EQUIPO'].isin(equipos_T)]
        .set_index('EQUIPO')['HORAS']
        .reindex(equipos_T)
    )
    bloques = (horas_T / 6).fillna(0).astype(float)

    cupones_hist = (
        df_hist_T.groupby('EQUIPO_FINAL').size()
        .reindex(equipos_T, fill_value=0).astype(int)
    )
    cupones_reap = (
        df_reap_T.groupby('EQUIPO_FINAL').size()
        .reindex(equipos_T, fill_value=0).astype(int)
    )

    # Estimado por pilar (PESO_BASE)
    def make_estimado_pilar_base(df_hist_area, df_fresh_area, df_pesos_areas, equipos, pillars, sigla_area='T', share_override=None):
        equipos = list(map(str, equipos))
        tot_hist_p = (df_hist_area.groupby('PILAR_NORM').size().reindex(pillars, fill_value=0).astype(int))
        tot_fresh_p = df_fresh_area['PILAR_NORM'].value_counts().reindex(pillars, fill_value=0).astype(int)
        T = (tot_hist_p + tot_fresh_p)
        if share_override is not None:
            s = pd.Series(share_override, index=equipos, dtype=float).fillna(0.0)
        else:
            pesos = (
                df_pesos_areas[df_pesos_areas['AREA'] == sigla_area]
                .set_index('EQUIPO')['PESO_BASE']
                .reindex(equipos).astype(float).fillna(0.0)
            )
            total_pesos = float(pesos.sum())
            s = pesos / (total_pesos if total_pesos > 0 else 1.0)
        if not np.isfinite(s.to_numpy()).all() or s.sum() <= 0:
            s = pd.Series(1.0 / len(equipos), index=equipos)
        est_float = pd.DataFrame(index=equipos, columns=pillars, dtype=float)
        for p in pillars:
            est_float[p] = float(T[p]) * s
        est_floor = np.floor(est_float).astype(int)
        est = est_floor.copy()
        for p in pillars:
            need = int(T[p] - est[p].sum())
            if need > 0:
                frac = (est_float[p] - est_floor[p]).reset_index()
                frac.columns = ['EQUIPO_FINAL', 'frac']
                frac = frac.sort_values(['frac','EQUIPO_FINAL'], ascending=[False, True])
                for eq in frac['EQUIPO_FINAL'].iloc[:need]:
                    est.at[eq, p] += 1
            elif need < 0:
                frac = (est_float[p] - est_floor[p]).reset_index()
                frac.columns = ['EQUIPO_FINAL', 'frac']
                frac = frac.sort_values(['frac','EQUIPO_FINAL'], ascending=[True, True])
                k = 0
                for eq in frac['EQUIPO_FINAL']:
                    if est.at[eq, p] > 0:
                        est.at[eq, p] -= 1
                        k += 1
                        if k == -need: break
        return est.reindex(index=equipos, columns=pillars, fill_value=0).astype(int)

    estimado_pilar = make_estimado_pilar_base(
        df_hist_area=df_hist_T, df_fresh_area=df_fresh_T,
        df_pesos_areas=df_pesos_areas, equipos=equipos_T, pillars=pillars, sigla_area='T'
    ).astype(int)

    # Cuotas IG/XP VARIABLES por PESO_BASE
    df_pesos_T = (
        df_pesos_areas[df_pesos_areas['AREA'] == 'T'][['EQUIPO','PESO_BASE']]
        .dropna().set_index('EQUIPO').reindex(equipos_T).fillna(0.0)
    )
    total_pesos_T = float(df_pesos_T['PESO_BASE'].sum())
    if total_pesos_T <= 0:
        share_IG, share_XP = 0.5, 0.5
    else:
        share_IG = float(df_pesos_T.loc[list(IG_teams & set(equipos_T)), 'PESO_BASE'].sum()) / total_pesos_T
        share_XP = 1.0 - share_IG

    n_hist = int(len(df_hist_T))
    n_reap = int(len(df_reap_T))
    n_fresh = int(len(df_fresh_T))
    total_T = n_hist + n_reap + n_fresh

    # Guardia: si no hay fresh, devolvemos coherente
    if n_fresh == 0:
        df_final_T = pd.concat([df_fresh_T.assign(EQUIPO_FINAL=np.nan).iloc[0:0], df_reap_T], ignore_index=True)
        df_cadencia     = pd.DataFrame(columns=["Equipo","Cadencia Original","Cadencia Final","Cadencia Teórica"])
        df_comparativa  = pd.DataFrame(index=equipos_T)
        cad_pilar       = pd.DataFrame(index=equipos_T, columns=pillars).fillna(0.0)
        return df_final_T, df_cadencia, df_comparativa, cad_pilar

    target_IG_total = int(round(share_IG * total_T))
    target_XP_total = total_T - target_IG_total

    fixed_IG = int((cupones_hist + cupones_reap).reindex(sorted(IG_teams & set(equipos_T)), fill_value=0).sum())
    fixed_XP = int((cupones_hist + cupones_reap).reindex(sorted(XP_teams & set(equipos_T)), fill_value=0).sum())

    quota_IG_fresh = max(0, target_IG_total - fixed_IG)
    quota_XP_fresh = max(0, target_XP_total - fixed_XP)

    if (quota_IG_fresh + quota_XP_fresh) != n_fresh:
        diff_quota = n_fresh - (quota_IG_fresh + quota_XP_fresh)
        gap_IG = target_IG_total - fixed_IG
        gap_XP = target_XP_total - fixed_XP
        if diff_quota > 0:
            if gap_IG >= gap_XP:
                quota_IG_fresh += diff_quota
            else:
                quota_XP_fresh += diff_quota
        else:
            remove = -diff_quota
            if gap_IG <= gap_XP:
                r = min(quota_IG_fresh, remove); quota_IG_fresh -= r; remove -= r
                quota_XP_fresh = max(0, quota_XP_fresh - remove)
            else:
                r = min(quota_XP_fresh, remove); quota_XP_fresh -= r; remove -= r
                quota_IG_fresh = max(0, quota_IG_fresh - remove)

    # Cadencia GLOBAL por equipo y targets FRESH
    cad_raw = (bloques * float(cad_teo_T)).fillna(0.0)
    sum_obj = float(cad_raw.sum())
    total_acum_real = int(cupones_hist.sum() + n_fresh)  # ACUMULADO tras asignar el día
    scale = (total_acum_real / sum_obj) if sum_obj > 0 else 1.0
    cad_target = cad_raw * scale

    # Fresh objetivo por equipo (float) y enteros base
    fresh_target_float = (cad_target - cupones_hist).astype(float)
    fresh_target_float = pd.Series(fresh_target_float, index=equipos_T)
    fresh_target_float = fresh_target_float.replace([np.inf, -np.inf], np.nan).fillna(0.0).clip(lower=0.0)
    fresh_target_int = _ajustar_redondeo_sum_exacta(fresh_target_float, total=n_fresh)

    # ---------- NUEVO BLOQUE ROBUSTO: Forzar CUOTAS IG/XP EXACTAS ----------
    def _pick_recipient_pos(fresh_int, fresh_float, group):
        gap = (fresh_float.reindex(group).fillna(0) - fresh_int.reindex(group).fillna(0))
        if (gap > 0).any():
            return gap.idxmax()
        return fresh_int.reindex(group).fillna(0).idxmin()

    def _pick_donor_neg(fresh_int, fresh_float, group):
        excess = (fresh_int.reindex(group).fillna(0) - fresh_float.reindex(group).fillna(0))
        candidates = fresh_int.reindex(group).fillna(0)
        candidates = candidates[candidates > 0]
        if candidates.empty:
            return None
        excess = excess.loc[candidates.index]
        if (excess > 0).any():
            return excess.idxmax()
        return candidates.idxmax()

    def _enforce_quota_exact(fresh_int: pd.Series,
                             fresh_float: pd.Series,
                             group_pos: list, group_neg: list,
                             target_pos: int) -> pd.Series:
        """
        Ajusta fresh_int por movimientos 1↔1 entre grupos para que
        sum(group_pos) == target_pos (y por tanto sum(group_neg) se ajusta).
        Conserva suma total y evita negativos.
        """
        fresh_int = fresh_int.copy().astype(int)

        for g in (group_pos, group_neg):
            for k in g:
                if k not in fresh_int.index:
                    fresh_int.loc[k] = 0
                    fresh_float.loc[k] = 0.0
        fresh_int = fresh_int.reindex(group_pos + group_neg).fillna(0).astype(int)
        fresh_float = fresh_float.reindex(fresh_int.index).fillna(0.0).astype(float)

        safety = 0
        max_steps = 10 * (len(group_pos) + len(group_neg) + 1)

        while True:
            sum_pos = int(fresh_int.reindex(group_pos).sum())
            need = int(target_pos - sum_pos)
            if need == 0:
                break
            if safety > max_steps:
                if need > 0:
                    for _ in range(need):
                        recip = _pick_recipient_pos(fresh_int, fresh_float, group_pos)
                        donor = _pick_donor_neg(fresh_int, fresh_float, group_neg)
                        if donor is None: break
                        fresh_int.at[recip] += 1
                        fresh_int.at[donor] -= 1
                else:
                    for _ in range(-need):
                        donor_pos = _pick_donor_neg(fresh_int, fresh_float, group_pos)
                        recip_neg = _pick_recipient_pos(fresh_int, fresh_float, group_neg)
                        if donor_pos is None: break
                        fresh_int.at[donor_pos] -= 1
                        fresh_int.at[recip_neg] += 1
                break

            if need > 0:
                recip = _pick_recipient_pos(fresh_int, fresh_float, group_pos)
                donor = _pick_donor_neg(fresh_int, fresh_float, group_neg)
                if donor is None:
                    break
                fresh_int.at[recip] += 1
                fresh_int.at[donor] -= 1
            else:
                donor_pos = _pick_donor_neg(fresh_int, fresh_float, group_pos)
                recip_neg = _pick_recipient_pos(fresh_int, fresh_float, group_neg)
                if donor_pos is None:
                    break
                fresh_int.at[donor_pos] -= 1
                fresh_int.at[recip_neg] += 1

            safety += 1

        if int(fresh_int.sum()) != int(fresh_target_int.sum()):
            raise RuntimeError("Conservación de suma violada al ajustar cuotas IG/XP.")
        if (fresh_int < 0).any():
            raise RuntimeError("Aparecieron enteros negativos al ajustar cuotas IG/XP.")
        return fresh_int

    IG_equipos = sorted(list(IG_teams & set(equipos_T)))
    XP_equipos = sorted(list(XP_teams & set(equipos_T)))

    sumIG = int(fresh_target_int.reindex(IG_equipos, fill_value=0).sum())
    needIG = int(quota_IG_fresh - sumIG)  # >0: subir IG; <0: bajar IG
    if needIG != 0:
        fresh_target_int = _enforce_quota_exact(
            fresh_target_int, fresh_target_float,
            group_pos=IG_equipos, group_neg=XP_equipos,
            target_pos=int(quota_IG_fresh)
        )

    sumIG2 = int(fresh_target_int.reindex(IG_equipos, fill_value=0).sum())
    sumXP2 = int(fresh_target_int.reindex(XP_equipos, fill_value=0).sum())
    assert int(fresh_target_int.sum()) == n_fresh, \
        f"Fresh por equipo no suma n_fresh: {int(fresh_target_int.sum())} != {n_fresh}"
    assert sumIG2 == int(quota_IG_fresh), \
        f"Cuota IG no cumplida (post-ajuste): IG={sumIG2}, objetivo={int(quota_IG_fresh)} (XP={sumXP2}, objetivo={int(quota_XP_fresh)})"
    assert sumXP2 == int(quota_XP_fresh), \
        f"Cuota XP no cumplida (post-ajuste): XP={sumXP2}, objetivo={int(quota_XP_fresh)} (IG={sumIG2}, objetivo={int(quota_IG_fresh)})"
    # ---------- FIN NUEVO BLOQUE ROBUSTO ----------

    # ========== MILP (pilares ajustados, cadencia fija por equipo) ==========
    team_vars = {(c, t): pulp.LpVariable(f"x_{c}_{t}", cat="Binary")
                 for c in coupons_T for t in equipos_T}

    diff_pilar = {(t, p): pulp.LpVariable(f"diff_{t}_{p}", lowBound=0)
                  for t in equipos_T for p in pillars}

    s_web_pos  = {t: pulp.LpVariable(f"s_web_pos_{t}",  lowBound=0) for t in equipos_T}
    s_web_neg  = {t: pulp.LpVariable(f"s_web_neg_{t}",  lowBound=0) for t in equipos_T}
    s_bsc_pos  = {t: pulp.LpVariable(f"s_bsc_pos_{t}",  lowBound=0) for t in equipos_T}
    s_bsc_neg  = {t: pulp.LpVariable(f"s_bsc_neg_{t}",  lowBound=0) for t in equipos_T}

    PENAL_DIFF = {'Web': 50, 'Buscadores': 30, 'P.Verticales': 0.5, 'Redes Sociales': 0.01}
    BIG_WEB, BIG_BUSC = 5000.0, 3000.0

    prob = pulp.LpProblem("Distribucion_Area_T", pulp.LpMinimize)
    prob += (
        pulp.lpSum(PENAL_DIFF[p] * diff_pilar[(t, p)] for t in equipos_T for p in pillars) +
        BIG_WEB  * pulp.lpSum(s_web_pos[t] + s_web_neg[t] for t in equipos_T) +
        BIG_BUSC * pulp.lpSum(s_bsc_pos[t] + s_bsc_neg[t] for t in equipos_T)
    )

    # 1) Cada cupón a un único equipo
    for c in coupons_T:
        prob += pulp.lpSum(team_vars[(c, t)] for t in equipos_T) == 1

    # 2) FRESH EXACTO por equipo (ya respeta cuotas IG/XP)
    for t in equipos_T:
        prob += pulp.lpSum(team_vars[(c, t)] for c in coupons_T) == int(fresh_target_int.loc[t])

    # 3) Ajuste por pilar vs estimado ACUMULADO + bandas blandas
    hist_map = df_hist_T.groupby(['EQUIPO_FINAL', 'PILAR_NORM']).size().to_dict()

    for t in equipos_T:
        for p in pillars:
            h = int(hist_map.get((t, p), 0))
            assign_tp = pulp.lpSum(team_vars[(c, t)] for c in coupons_T if df_fresh_T.at[c, 'PILAR_NORM'] == p)
            total_tp  = h + assign_tp
            est_tp    = int(estimado_pilar.loc[t, p])
            prob +=  total_tp - est_tp <= diff_pilar[(t, p)]
            prob +=  est_tp - total_tp <= diff_pilar[(t, p)]

        # Bandas blandas Web
        hW = int(hist_map.get((t, 'Web'), 0))
        aW = pulp.lpSum(team_vars[(c, t)] for c in coupons_T if df_fresh_T.at[c, 'PILAR_NORM'] == 'Web')
        totW = hW + aW
        estW = int(estimado_pilar.loc[t, 'Web'])
        epsW = max(1, int(round(estW * float(pilar_band_web))))
        prob += totW - estW <=  epsW + s_web_pos[t]
        prob += estW - totW <=  epsW + s_web_neg[t]

        # Bandas blandas Buscadores
        hB = int(hist_map.get((t, 'Buscadores'), 0))
        aB = pulp.lpSum(team_vars[(c, t)] for c in coupons_T if df_fresh_T.at[c, 'PILAR_NORM'] == 'Buscadores')
        totB = hB + aB
        estB = int(estimado_pilar.loc[t, 'Buscadores'])
        epsB = max(1, int(round(estB * float(pilar_band_busc))))
        prob += totB - estB <=  epsB + s_bsc_pos[t]
        prob += estB - totB <=  epsB + s_bsc_neg[t]

    print("🧩 Resolviendo modelo para área T (cadencia exacta + cuotas IG/XP + Web/Busc ajustados)...")
    try:
        status = prob.solve(pulp.PULP_CBC_CMD(msg=False, timeLimit=time_limit))
        lp_status = pulp.LpStatus[prob.status]
    except Exception as e:
        lp_status = f"Exception: {e}"

    print("📌 Estado:", lp_status)

    infeasible = (lp_status not in ("Optimal",))

    # ---------- Recuperar asignaciones (si MILP OK) ----------
    if not infeasible:
        asignaciones = []
        for c in coupons_T:
            asignado = None
            for t in equipos_T:
                val = pulp.value(team_vars[(c, t)])
                if val is not None and round(val) == 1:
                    asignado = t
                    break
            if asignado is None:
                deficits = fresh_target_int.reindex(equipos_T).fillna(0)
                asignado = deficits.idxmax()
            asignaciones.append(asignado)
        df_fresh_T['EQUIPO_FINAL'] = asignaciones
    else:
        print("⚠️ MILP no resolvió (Infeasible/No Optimal). Activando Fallback proporcional por pilares...")
        df_fresh_T = _fallback_asignacion(df_fresh_T, fresh_target_int, estimado_pilar, df_hist_T, equipos_T, pillars)

    # ---------- Ajuste fino (ACUM) ----------
    def _force_pillars_columns_local(m, index=None):
        return _force_pillars_columns(m, index=index)  # usa tu helper global

    def ajuste_fino_cadencia_acum(df_fresh_area, estimado_pilar, df_hist_area):
        df = df_fresh_area.copy()
        pilares_clave = ['Web', 'Buscadores']
        pilares_compensa = ['Redes Sociales', 'P.Verticales']
        equipos = list(map(str, equipos_T))
        def matriz_acum(dframe):
            m = (dframe.groupby(['EQUIPO_FINAL','PILAR_NORM']).size().unstack(fill_value=0))
            return _force_pillars_columns_local(m, index=equipos)
        acum = pd.concat([df_hist_area[['EQUIPO_FINAL','PILAR_NORM']],
                          df[['EQUIPO_FINAL','PILAR_NORM']]], ignore_index=True)
        count_acum = matriz_acum(acum)
        est_safe = estimado_pilar.reindex(index=equipos, columns=pillars, fill_value=0)
        delta = (count_acum - est_safe).fillna(0).astype('Int64')
        movimientos = 0
        for pilar in pilares_clave:
            while True:
                delta_p = delta[pilar]; exceso_eq = delta_p.idxmax(); falta_eq = delta_p.idxmin()
                if int(delta_p.get(exceso_eq,0)) <= 0 or int(delta_p.get(falta_eq,0)) >= 0: break
                cupones_exceso = df[(df['EQUIPO_FINAL']==exceso_eq) & (df['PILAR_NORM']==pilar)]
                if cupones_exceso.empty: break
                realizado = False
                for p_comp in pilares_compensa:
                    cupones_comp = df[(df['EQUIPO_FINAL']==falta_eq) & (df['PILAR_NORM']==p_comp)]
                    if cupones_comp.empty: continue
                    for i in cupones_exceso.index:
                        if realizado: break
                        for j in cupones_comp.index:
                            if i == j: continue
                            ei, ej = df.at[i,'EQUIPO_FINAL'], df.at[j,'EQUIPO_FINAL']
                            df.at[i,'EQUIPO_FINAL'], df.at[j,'EQUIPO_FINAL'] = ej, ei
                            acum2 = pd.concat([df_hist_area[['EQUIPO_FINAL','PILAR_NORM']],
                                               df[['EQUIPO_FINAL','PILAR_NORM']]], ignore_index=True)
                            new_count = matriz_acum(acum2)
                            new_delta = (new_count - est_safe).fillna(0).astype('Int64')
                            if int(new_delta.abs().to_numpy().sum()) < int(delta.abs().to_numpy().sum()):
                                delta = new_delta; movimientos += 1; realizado = True; break
                            else:
                                df.at[i,'EQUIPO_FINAL'], df.at[j,'EQUIPO_FINAL'] = ei, ej
                    if realizado: break
                if not realizado: break
        print(f"\n🔁 Ajuste fino (ACUM) completado. Movimientos: {movimientos}")
        return df

    df_fresh_T = ajuste_fino_cadencia_acum(df_fresh_T, estimado_pilar[pillars], df_hist_T)

    # ---------- Métricas y salida ----------
    df_real_fresh = (df_fresh_T.groupby(['EQUIPO_FINAL','PILAR_NORM']).size().unstack(fill_value=0))
    df_real_fresh = _force_pillars_columns(df_real_fresh, index=equipos_T)

    df_acum_real = pd.concat([df_hist_T, df_fresh_T], ignore_index=True)
    df_real_acum = (df_acum_real.groupby(['EQUIPO_FINAL','PILAR_NORM']).size().unstack(fill_value=0))
    df_real_acum = _force_pillars_columns(df_real_acum, index=equipos_T)

    bloques_rep = (horas_T / 6).replace(0, pd.NA)
    cad_pilar = df_real_fresh.div(bloques_rep, axis=0).astype(float).round(2)

    total_final = (
        df_fresh_T['EQUIPO_FINAL'].value_counts()
        .add(cupones_hist, fill_value=0)
        .reindex(equipos_T, fill_value=0)
    )
    cad_final = (total_final / bloques_rep).astype(float).round(4)

    df_cadencia = pd.DataFrame({
        "Equipo": equipos_T,
        "Cadencia Original": pd.Series({eq: cad_prelim_T_dict.get(eq, 0) for eq in equipos_T}).reindex(equipos_T).fillna(0).values,
        "Cadencia Final": cad_final.values,
        "Cadencia Teórica": [cad_teo_T] * len(equipos_T)
    })

    df_est = estimado_pilar.copy(); df_est.columns = [p + "_Estimado" for p in df_est.columns]
    df_real_a = df_real_acum.copy();  df_real_a.columns  = [p + "_Real_ACUM"  for p in df_real_a.columns]
    df_real_f = df_real_fresh.copy(); df_real_f.columns = [p + "_Real_FRESH" for p in df_real_f.columns]
    df_comparativa = pd.concat([df_est, df_real_a, df_real_f], axis=1)

    df_final_T = pd.concat([df_fresh_T, df_reap_T], ignore_index=True)

    # --- Resumen IG/XP final para ver cumplimiento ---
    final_counts_equipo = (
        df_final_T['EQUIPO_FINAL'].value_counts()
        .add(0, fill_value=0).reindex(equipos_T, fill_value=0).astype(int)
    )
    final_IG = int(final_counts_equipo.reindex(IG_equipos, fill_value=0).sum())
    final_XP = int(final_counts_equipo.reindex(XP_equipos, fill_value=0).sum())
    share_IG_final = final_IG / (len(df_final_T) + len(df_hist_T)*0) if total_T > 0 else 0  # solo display

    print("\n🔎 Comparativa (Estimado / Real_ACUM / Real_FRESH) tras ajuste fino:")
    print(df_comparativa)
    print("\n📊 Cadencia por Pilar y Equipo (FRESH):")
    print(cad_pilar)
    print("\n📏 Verificación cadencias Área T")
    print(df_cadencia)
    print(f"\n✅ IG fresh cuota: {fresh_target_int.reindex(IG_equipos, fill_value=0).sum()} / {quota_IG_fresh} | XP: {fresh_target_int.reindex(XP_equipos, fill_value=0).sum()} / {quota_XP_fresh}")
    print(f"🎯 Shares objetivo (PESOS T): IG={share_IG:.2%} XP={share_XP:.2%}")
    if infeasible:
        print("⚠️ Resultado con Fallback proporcional (no MILP). Revisa penalizaciones/bandas/cuotas si deseas una solución óptima MILP.")

    return df_final_T, df_cadencia, df_comparativa, cad_pilar


# In[19]:


df_final_T, df_cadencia, df_comparativa, cad_pilar = distribuir_area_T(
    df_fresh, df_hist_total, df_horas_eq, df_pesos_areas,
    cad_prelim_T_dict, cad_teo_T, df_reap_validas,
    pilar_band_web=0.05, pilar_band_busc=0.05, time_limit=90
)


# ### AREA E

# In[20]:


def distribuir_area_E(df_fresh, df_hist_total, df_pesos_areas, df_reap_validas):
    """
    Área E:
    1) Sin cadencia. Objetivo 1: IG=50% vs XP=40% del total (hist + reaps + fresh).
    2) Objetivo 2: dentro de cada directora, aproximar a los PESO_BASE por equipo (de df_pesos_areas).
    3) Reaperturas válidas se pegan al final (cuentan como hoy), pero NO se reasignan.
    """

    # --- Config ---
    pillars = ['Buscadores', 'Redes Sociales', 'Web', 'P.Verticales']

    # Map de equipos por directora
    IG_teams = {'Equipo_A1', 'Equipo_B1', 'Equipo_C1'}
    XP_teams = {'Equipo_A2', 'Equipo_B2', 'Equipo_C2'}

    # --- Equipos E y pesos base ---
    df_pesos_E = df_pesos_areas[df_pesos_areas['AREA'] == 'E'][['EQUIPO', 'PESO_BASE']].dropna()
    equipos_E = df_pesos_E['EQUIPO'].unique().tolist()
    if not equipos_E:
        raise ValueError("No hay equipos para área E en df_pesos_areas.")

    # Filtra conjuntos IG/XP efectivos (solo equipos presentes en E)
    IG_equipos = sorted([e for e in equipos_E if e in IG_teams])
    XP_equipos = sorted([e for e in equipos_E if e in XP_teams])

    if len(IG_equipos) == 0 or len(XP_equipos) == 0:
        print("ADVERTENCIA: Falta al menos un grupo de directora en los equipos de E. "
              "Verifica que E incluya A1,B1,C1 (IG) y A2,B2,C2 (XP).")

    # Normaliza y reindexa pesos a los equipos efectivos
    df_pesos_E = df_pesos_E.set_index('EQUIPO').reindex(equipos_E).fillna(0.0)
    # Asegura que sumen 1.0 (si tus PESO_BASE ya vienen normalizados, esto no cambia nada)
    suma_pesos = df_pesos_E['PESO_BASE'].sum()
    if suma_pesos <= 0:
        raise ValueError("PESO_BASE para E no es válido (suma <= 0).")
    df_pesos_E['PESO_NORM'] = df_pesos_E['PESO_BASE'] / suma_pesos

    # --- Filtrado FRESH E ---
    df_fresh_E = (
        df_fresh[
            (df_fresh['AREA'] == 'E') &
            (df_fresh['TIPO'].isin(['MST', 'MBA'])) &
            (df_fresh['IDIOMA'] == 'ENG') &
            (df_fresh['PILAR_NORM'].isin(pillars))
        ]
        .copy()
        .reset_index(drop=True)
    )
    coupons_E = df_fresh_E.index.tolist()

    # --- Histórico válido E (solo equipos E) ---
    df_hist_E = (
        df_hist_total[
            (df_hist_total['TIPO'].isin(['MST', 'MBA'])) &
            (df_hist_total['IDIOMA'] == 'ENG') &
            (df_hist_total['EQUIPO_FINAL'].isin(equipos_E)) &
            (df_hist_total['PILAR_NORM'].isin(pillars))
        ]
        .copy()
    )

    # --- Reaperturas válidas E (se respetan) ---
    df_reap_validas_E = (
        df_reap_validas[
            (df_reap_validas['AREA'] == 'E') &
            (df_reap_validas['TIPO'].isin(['MST', 'MBA'])) &
            (df_reap_validas['IDIOMA'] == 'ENG') &
            (df_reap_validas['EQUIPO_FINAL'].isin(equipos_E))
        ]
        .copy()
    )

    # --- Totales base ---
    n_hist  = len(df_hist_E)
    n_reaps = len(df_reap_validas_E)
    n_fresh = len(df_fresh_E)
    total_E = n_hist + n_reaps + n_fresh

    # Targets por directora (enteros)
    target_IG_total = int(round(0.50 * total_E))
    target_XP_total = total_E - target_IG_total

    # Conteos fijos por directora (hist + reaps)
    hist_counts = df_hist_E['EQUIPO_FINAL'].value_counts().reindex(equipos_E, fill_value=0)
    reap_counts = df_reap_validas_E['EQUIPO_FINAL'].value_counts().reindex(equipos_E, fill_value=0)

    fixed_IG = int(hist_counts.reindex(IG_equipos, fill_value=0).sum() + reap_counts.reindex(IG_equipos, fill_value=0).sum())
    fixed_XP = int(hist_counts.reindex(XP_equipos, fill_value=0).sum() + reap_counts.reindex(XP_equipos, fill_value=0).sum())

    # Cuotas de fresh por directora (intentos)
    quota_IG_fresh = target_IG_total - fixed_IG
    quota_XP_fresh = target_XP_total - fixed_XP

    # Ajuste de factibilidad: que no sean negativos y que sumen n_fresh
    quota_IG_fresh = max(0, quota_IG_fresh)
    quota_XP_fresh = max(0, quota_XP_fresh)

    # Si la suma no cuadra, ajusta proporcionalmente o por resto
    suma_quota = quota_IG_fresh + quota_XP_fresh
    if suma_quota != n_fresh:
        # Distribuye el desfase al grupo con mayor gap relativo
        diff = n_fresh - suma_quota
        # Métrica simple de prioridad: quien esté más lejos del target total tras fijos
        gap_IG = target_IG_total - fixed_IG
        gap_XP = target_XP_total - fixed_XP
        if diff > 0:
            # tenemos fresh "libres": añadir al que tenga mayor gap
            if gap_IG >= gap_XP:
                quota_IG_fresh += diff
            else:
                quota_XP_fresh += diff
        else:
            # nos sobran cuotas: restar al que tenga menor gap
            remove = -diff
            if gap_IG <= gap_XP:
                quitar_IG = min(quota_IG_fresh, remove)
                quota_IG_fresh -= quitar_IG
                remove -= quitar_IG
                quota_XP_fresh = max(0, quota_XP_fresh - remove)
            else:
                quitar_XP = min(quota_XP_fresh, remove)
                quota_XP_fresh -= quitar_XP
                remove -= quitar_XP
                quota_IG_fresh = max(0, quota_IG_fresh - remove)

    # --- Targets por equipo (en función de PESO_NORM sobre total_E) ---
    team_targets = (df_pesos_E['PESO_NORM'] * total_E).round().astype(int)
    # Ajuste para que sume exactamente total_E
    dif_total = total_E - int(team_targets.sum())
    if dif_total != 0:
        # corrige el residuo en el equipo con mayor decimal original
        decimals = (df_pesos_E['PESO_NORM'] * total_E) - (df_pesos_E['PESO_NORM'] * total_E).astype(int)
        if dif_total > 0:
            eq_fix = decimals.sort_values(ascending=False).index.tolist()
        else:
            eq_fix = decimals.sort_values(ascending=True).index.tolist()
        for eq in eq_fix:
            if dif_total == 0: break
            team_targets.loc[eq] += 1 if dif_total > 0 else -1
            dif_total += -1 if dif_total > 0 else 1

    # Conteo acumulado actual (hist + reaps) por equipo
    current_total = (hist_counts + reap_counts).reindex(equipos_E, fill_value=0)

    # --- Asignación GREEDY de fresh ---
    # 1) Particiona la lista de fresh en dos bolsas: IG y XP según cuotas calculadas
    fresh_indices = list(df_fresh_E.index)
    IG_need = int(quota_IG_fresh)
    XP_need = int(quota_XP_fresh)
    # Por simplicidad, usa el orden actual. Si quieres aleatorizar: np.random.shuffle(fresh_indices)
    IG_bucket = fresh_indices[:IG_need]
    XP_bucket = fresh_indices[IG_need:IG_need+XP_need]

    # 2) Dentro de cada bolsa, asigna al equipo con mayor déficit relativo respecto a su team_target
    asignaciones = {}

    def asignar_en_bolsa(indices, equipos_grupo):
        nonlocal asignaciones, current_total
        if not equipos_grupo:
            # Si no hay equipos en el grupo (caso patológico), reparte proporcionalmente entre todos
            equipos_grupo = equipos_E[:]
        for idx in indices:
            # Déficit = target - (actual + ya-asignados)
            deficits = {eq: int(team_targets.get(eq, 0) - current_total.get(eq, 0)) for eq in equipos_grupo}
            # Si todos los déficits <= 0, asigna al de menor exceso (o cualquiera)
            # Ordena por mayor déficit (desc), y desempata por menor carga actual
            orden = sorted(equipos_grupo, key=lambda e: (deficits[e], -current_total[e]), reverse=True)
            elegido = orden[0]
            asignaciones[idx] = elegido
            current_total[elegido] += 1

    asignar_en_bolsa(IG_bucket, IG_equipos)
    asignar_en_bolsa(XP_bucket, XP_equipos)

    # Si quedaran cupones (por alguna razón), reparte entre todos priorizando déficit
    restantes = [i for i in fresh_indices if i not in asignaciones]
    if restantes:
        for idx in restantes:
            deficits_all = {eq: int(team_targets.get(eq, 0) - current_total.get(eq, 0)) for eq in equipos_E}
            orden_all = sorted(equipos_E, key=lambda e: (deficits_all[e], -current_total[e]), reverse=True)
            elegido = orden_all[0]
            asignaciones[idx] = elegido
            current_total[elegido] += 1

    df_fresh_E['EQUIPO_FINAL'] = df_fresh_E.index.map(asignaciones)

    # --- Salida final (fresh asignado + reaps válidas) ---
    df_final_E = pd.concat([df_fresh_E, df_reap_validas_E], ignore_index=True)

    # --- Comparativas ---
    # Totales finales por equipo = hist + reaps + fresh_asignado
    final_counts_equipo = (
        df_final_E['EQUIPO_FINAL'].value_counts().add(hist_counts, fill_value=0).astype(int)
    ).reindex(equipos_E, fill_value=0)

    df_objetivo_equipo = team_targets.rename("Estimado_Equipo")
    df_asignado_equipo = final_counts_equipo.rename("Asignado_Equipo")
    df_comparativa_equipo = pd.concat([df_objetivo_equipo, df_asignado_equipo], axis=1)
    df_comparativa_equipo['Delta'] = df_comparativa_equipo['Asignado_Equipo'] - df_comparativa_equipo['Estimado_Equipo']

    # Resumen por directora IG/XP (final)
    final_IG = int(final_counts_equipo.reindex(IG_equipos, fill_value=0).sum())
    final_XP = int(final_counts_equipo.reindex(XP_equipos, fill_value=0).sum())
    share_IG = final_IG / total_E if total_E > 0 else 0
    share_XP = final_XP / total_E if total_E > 0 else 0

    df_resumen_directora = pd.DataFrame({
        'Total_Final': [final_IG, final_XP],
        'Share_Final': [share_IG, share_XP],
        'Share_Target': [0.50, 0.50],
        'Delta_puntos': [share_IG - 0.50, share_XP - 0.50]
    }, index=['IG', 'XP'])

    # Por pilar (solo fresh hoy, para inspección)
    pillars_cols = pillars[:]  # asegura orden
    df_real_pilar_hoy = (
        df_fresh_E
        .groupby(['EQUIPO_FINAL', 'PILAR_NORM'])
        .size()
        .unstack(fill_value=0)
        .reindex(index=equipos_E, columns=pillars_cols, fill_value=0)
    )

    # Avisos de factibilidad
    msgs = []
    if (final_IG + final_XP) != total_E:
        msgs.append("ADVERTENCIA: Totales IG+XP no igualan Total E (inspeccionar filtros).")
    if (final_IG != target_IG_total) or (final_XP != target_XP_total):
        msgs.append("Nota: No se pudo igualar exactamente 50/50 por fijos históricos/reaperturas. "
                    "Se compensó al máximo con los fresh.")

    if msgs:
        for m in msgs:
            print(m)

    # Prints útiles
    print("\n=== Comparativa por EQUIPO (Estimado vs Asignado) ===")
    print(df_comparativa_equipo.sort_index())
    print("\n=== Resumen por DIRECTORA (IG/XP) ===")
    print(df_resumen_directora)

    return df_final_E, df_comparativa_equipo, df_real_pilar_hoy, df_resumen_directora


# ===== EJECUCIÓN =====
df_final_E, df_comparativa_E, df_real_E, df_resumen_directora_E = distribuir_area_E(
    df_fresh=df_fresh,
    df_hist_total=df_hist_total,
    df_pesos_areas=df_pesos_areas,
    df_reap_validas=df_reap_validas
)

# # Export (opcional)
# df_final_E.to_excel("Distribucion_Final_Area_E.xlsx", index=False)
# df_comparativa_E.to_excel("Comparativa_Equipo_E.xlsx")
# df_resumen_directora_E.to_excel("Comparativa_Directora_E.xlsx")
# df_real_E.to_excel("Fresh_por_Pilar_E.xlsx")


# ### JUNTAMOS TODOS LOS DF EN UNO SOLO; ORGANIZAMOS POR INDEX_ORIGINAL Y EXPORTAMOS A EXCEL

# In[21]:


df_final_total = pd.concat([
    df_final_A,
    df_final_B,
    df_final_C,
    df_final_T,
    df_final_E,
    df_corte,
    df_special,
], ignore_index=True)
#Ordenamos por el índice original para mantener el orden de los cupones
df_final_total = (
    df_final_total
      .sort_values(['ID de la Oportunidad', 'INDEX_ORIGINAL'], kind='stable')
      .drop_duplicates(subset=['ID de la Oportunidad'], keep='first')
      .sort_values('INDEX_ORIGINAL', kind='stable')
      .reset_index(drop=True)
)

# df_final_total.to_excel("Distribucion_Final_I.xlsx", index=False)  -->Por si queremos exportar la primera parte individual


# ## PARTE II --- OPTIMIZACIÓN PAIS/PROGRAMA

# ### FUNCIÓN 

# Distribución de cupones – Segunda etapa (V24)
# 
# > **Objetivo**: Optimizar **País** (Agrupación OBS) y **Programa** dentro de cada área sin alterar **ningún conteo** producido por la primera etapa (cadencia y repartos por equipo/pilar/DV). Esta etapa opera **exclusivamente sobre FRESH** por **bloques** (Área → Grupo de equipos → Pilar → [DV en T]), y garantiza invariantes **por construcción** mediante **slots exactos**.
# 
# ---
# 
# ##1) Resumen ejecutivo
# - **Qué hace**: Reordena **qué cupón específico** lleva cada equipo dentro de su mismo pilar (y DV en T), para acercar el acumulado **(HIST+FRESH)** por **País** y **Programa** al reparto esperado según **PESO_BASE** del área.
# - **Qué no hace**: No cambia **cuántos cupones** tiene cada equipo. Ni por área, ni por pilar, ni por DV (en T). Por ello, **cadencia** y **HIST+FRESH por equipo×pilar** quedan idénticos.
# - **Cómo lo garantiza**: Fija **slots** por bloque (= #FRESH actuales por equipo×pilar×[DV]). La reasignación solo **permuta** cupones entre equipos que **comparten ese bloque** y con el **mismo presupuesto de slots**.
# - **Prioridad**: Optimiza con prioridad **País > Programa** (lexicográfico). Si no hay mejora posible sin romper invariantes, **no cambia nada** y lo **registra**.
# - **REAP**: Para métricas, deduplica REAP contra HIST (cuentan como HIST para objetivos). En el export final **reinyecta** REAP sin duplicar claves.
# 
# ---
# 
# ##2) Entradas y salidas
# **Entradas mínimas de DataFrames**
# - `df_final_total`: filas **FRESH** (y opcionalmente **REAP**). Columnas clave: `INDEX_ORIGINAL`, `TIPO_REPARTO`, `EQUIPO_FINAL`, `PILAR_NORM`, `Agrupación OBS`, `Programa de Interes`, `AREA`, y en T a veces `DV`.
# - `df_hist_total`: histórico **(incluye reaperturas no atendidas)** con mismas columnas análogas necesarias.
# - `df_pesos_areas`: columnas `AREA`, `EQUIPO`, `PESO_BASE` (para objetivos esperados).
# - `df_horas_eq`: columnas `EQUIPO`, `HORAS` (para validar **cadencia**).
# 
# **Salida pública** (firma estable):
# ```python
# res_2a, df_final_ajustado = run_segunda_etapa_v19(...)
# ```
# - `res_2a`: df final **optimizado** de la **segunda etapa** (CLEAN, sin REAP reinyectadas).
# - `df_final_ajustado`: `res_2a` + **REAP originales reinyectadas** sin duplicar.
# 
# ---
# 
# ##3) Invariantes duros (se cumplen al 100%)
# 1. **#FRESH por equipo** (global, por área) **invariante**.
# 2. **FRESH×PILAR por equipo** (global, por área y por grupo) **invariante**.
# 3. En **área T**: **FRESH×PILAR×DV por equipo** **invariante**.
# 4. **(HIST + FRESH) por (equipo×pilar)** **invariante**.
# 5. **Cadencia por equipo** = (HIST + FRESH) / (HORAS/6) **invariante** (tolerancia ≤ 1e-9).
# 
# ---
# 
# ##4) Concepto de “slots”
# - En un bloque (p. ej. **Área A → Grupo [Equipo_A1, Equipo_A2] → Pilar Web**), si **A1** tiene 7 FRESH y **A2** tiene 5 FRESH, entonces los **slots** son: A1=7, A2=5.
# - La segunda etapa **no puede** cambiar esos números; solo decide **qué** 7 cupones concretos van a A1 y **qué** 5 a A2, buscando mejorar **País/Programa**.
# - En **T** con `DV`, el slot se define a nivel **(equipo×pilar×DV)** y también queda **bloqueado**.
# 
# ---
# 
# ##5) Flujo de la segunda etapa (vista general)
# 1. **Dedupe REAP** para métricas (`dedup_reap_hist_vs_final`).
# 2. **Foto “antes”**: conteos globales de #FRESH, HIST+FRESH×pilar y cadencia.
# 3. **Core**: por área → grupo → pilar (→ DV en T). Reasignación en slots con optimización País > Programa.
# 4. **Validaciones** grupo → área → global. Rollback si falla algún invariante.
# 5. **Foto “después”** con asserts.
# 6. **Reinyección REAP** en export final.
# 
# ---
# 
# ##6) Funciones principales
# - `run_segunda_etapa_v19`: orquesta toda la etapa.
# - `dedup_reap_hist_vs_final`: elimina duplicados REAP vs HIST.
# - `_segunda_etapa_core`: núcleo transaccional por área.
# - `_balanced_reassign_pillar`: reasignación en slots con optimización lexicográfica País > Programa.
# - `_groups_for_area`: agrupa equipos según reglas de base/directora.
# - `_cadencia_por_equipo`: recalcula cadencia y valida que no cambie.
# - `construir_final_con_reap`: reinyecta REAP originales.
# 
# ---
# 
# ##7) Algoritmo de reasignación (detalle)
# **Dentro de cada bloque** (área→grupo→pilar→[DV]):
# 1. Fijar **slots** (ej. A1=7, A2=5).
# 2. Calcular **objetivos esperados** por País y Programa con (HIST+FRESH) y PESO_BASE.
# 3. Para cada cupón:
#    - Evaluar coste lexicográfico `(cost_country, cost_program)` en cada equipo con slot.
#    - Asignar al equipo con menor coste.
#    - Actualizar contadores de País/Programa del equipo.
# 4. Verificar que **slots finales = slots iniciales**.
# 
# ---
# 
# ##8) Ejemplo numérico (mini-bloque)
# **Área A → Grupo [A1,A2] → Pilar Web**
# - Slots iniciales: A1=2, A2=1.
# - Cupones (País, Programa):
#   - Cupón 1: (ES, MBA)
#   - Cupón 2: (FR, MSC)
#   - Cupón 3: (ES, MSC)
# - Objetivo esperado (HIST+FRESH):
#   - A1: ES=2, FR=0
#   - A2: ES=0, FR=1
# 
# **Asignación paso a paso:**
# 1. Cupón 1 (ES,MBA): A1 necesita ES → coste=0,0 → asignado A1.
# 2. Cupón 2 (FR,MSC): A2 necesita FR → coste=0,0 → asignado A2.
# 3. Cupón 3 (ES,MSC): A1 aún tiene slot y necesita ES → asignado A1.
# 
# **Resultado**: A1 recibe cupones 1 y 3, A2 recibe cupón 2. Slots intactos (2+1).
# 
# ---
# 
# ##9) Diagrama de flujo simplificado
# ```
#       START
#         │
#    Dedupe REAP
#         │
#    Foto "antes"
#         │
#    ┌──────────────┐
#    │ Por cada área│
#    └──────┬───────┘
#           │
#    Agrupar equipos
#           │
#      ┌─────────────┐
#      │ Por pilar   │
#      └──────┬──────┘
#             │
#    Fijar slots exactos
#             │
#    Reasignar (País>Programa)
#             │
#    Validar grupo
#             │
#    Validar área
#             │
#    Commit/rollback
#             │
#    Foto "después"
#             │
#    Assert invariantes
#             │
#    Reinyectar REAP
#         │
#        END
# ```
# 
# ---
# 
# 10) Checklist de verificación rápida
# - [ ] `PILAR_NORM` en {Web, Buscadores, P.Verticales, Redes Sociales}.
# - [ ] DV en T bloqueado (`lock_cols_by_area={'T':['DV']}`).
# - [ ] `HORAS>0` para todos los equipos.
# - [ ] `INDEX_ORIGINAL` único en FRESH.
# - [ ] REAP deduplicadas correctamente.
# 
# ---
# 
# ##11) Glosario
# - **Slot**: #FRESH fijo de un equipo en un bloque (equipo×pilar×DV).
# - **Bloque**: Área→Grupo→Pilar→[DV].
# - **Lexicográfico**: prioriza País sobre Programa.
# - **Commit/Rollback**: aplicar/descartar cambios según validaciones.
# 
# ---
# 
# ##12) Conclusión
# La segunda etapa **no altera la cadencia ni los conteos**, solo refina la **calidad del mix País/Programa**. Usa **slots exactos**, validaciones estrictas y un **algoritmo greedy lexicográfico**, garantizando seguridad y mejoras locales sin riesgo de romper la lógica central de distribución.
# 
# 
# 
# ---
# 
#  Anexo A · Ejemplo numérico paso a paso (bloque mini)
# **Bloque**: Área **T** → Grupo **[Equipo_A1, Equipo_A2]** → Pilar **Web** → **DV=IG**
# 
# **Datos del bloque**
# - **Slots (FRESH actuales en el bloque)**: A1=3, A2=2  
# - **Cupones FRESH** (5 filas):
#   1. C1: País=ES, Programa=MBA
#   2. C2: País=ES, Programa=UX
#   3. C3: País=MX, Programa=Data
#   4. C4: País=MX, Programa=MBA
#   5. C5: País=CL, Programa=UX
# - **Histórico en el grupo (mismo pilar y DV)**:
#   - A1: ES=1, MX=0, CL=1 ; Programas: MBA=1, UX=0, Data=1
#   - A2: ES=0, MX=1, CL=0 ; Programas: MBA=0, UX=1, Data=0
# - **PESOS_BASE del grupo** (normalizados): A1=0.6, A2=0.4
# 
# **Objetivo esperado (HIST+FRESH) por País**  
# Total HIST+FRESH por país en el bloque (sumando los 5 FRESH y el HIST):
# - ES: 2 (HIST 1 + FRESH 1 ya contabilizado al total) + 1 adicional = 2 (para el ejemplo simplificamos al total de países en el bloque = {ES:2, MX:2, CL:1})
# - MX: 2
# - CL: 1
# 
# Reparto esperado por equipo (multiplicar por pesos y redondeo Hamilton):
# - Para **ES=2** → A1≈1.2→1; A2≈0.8→1  
# - Para **MX=2** → A1≈1.2→1; A2≈0.8→1  
# - Para **CL=1** → A1≈0.6→1; A2≈0.4→0  
# **Esperado País** → A1: {ES:1, MX:1, CL:1} ; A2: {ES:1, MX:1, CL:0}
# 
# **Objetivo esperado (HIST+FRESH) por Programa** (totales en bloque: MBA=2, UX=2, Data=1):
# - MBA=2 → A1≈1.2→1; A2≈0.8→1
# - UX =2 → A1≈1.2→1; A2≈0.8→1
# - Data=1 → A1≈0.6→1; A2≈0.4→0
# **Esperado Programa** → A1: {MBA:1, UX:1, Data:1} ; A2: {MBA:1, UX:1, Data:0}
# 
# **Actuales (HIST)**
# - País A1: ES=1, MX=0, CL=1  → faltan MX(+1), sobran nada  
# - País A2: ES=0, MX=1, CL=0  → falta ES(+1), sobra nada
# - Prog A1: MBA=1, UX=0, Data=1 → falta UX(+1)  
# - Prog A2: MBA=0, UX=1, Data=0 → falta MBA(+1)
# 
# **Asignación greedy lexicográfica (País > Programa)**
# - **Slots**: A1×3, A2×2
# - Paso 1: Cupón **C1(ES,MBA)**
#   - A1: País ES → A1 necesita ES? (esperado ES=1, ya tiene ES=1 HIST) → está **justo** ⇒ coste país=1. Prog MBA → A1 necesita MBA? (ya tiene 1 HIST) ⇒ **justo** ⇒ coste prog=1 → (1,1)
#   - A2: País ES → **falta** ⇒ coste=0. Prog MBA → **falta** ⇒ coste=0 → **(0,0)** → asignar a **A2**.  
#   **Update**: A2 gana ES y MBA (HIST+FRESH). Slots A2 restantes=1.
# - Paso 2: Cupón **C2(ES,UX)**
#   - A1: País ES → está **justo** (1 esperado, 1 actual) ⇒ 1; Prog UX → **falta** ⇒ 0 → (1,0)
#   - A2: País ES → ahora está **justo** ⇒ 1; Prog UX → **justo** (ya tenía 1 HIST) ⇒ 1 → (1,1)
#   → Gana **A1** por (1,0) < (1,1).  
#   **Update**: A1 gana ES y UX. Slots A1 restantes=2.
# - Paso 3: Cupón **C3(MX,Data)**
#   - A1: País MX → **falta** ⇒ 0; Prog Data → **justo** (ya tenía 1 HIST, espera 1) ⇒ 1 → (0,1)
#   - A2: País MX → **justo** ⇒ 1; Prog Data → **falta?** espera 0, tiene 0 ⇒ **justo** ⇒ 1 → (1,1)
#   → Gana **A1** (0,1) < (1,1).  
#   **Update**: A1 gana MX y Data. Slots A1 restantes=1.
# - Paso 4: Cupón **C4(MX,MBA)**
#   - A1: País MX → ahora **justo** ⇒ 1; Prog MBA → **justo** ⇒ 1 → (1,1)
#   - A2: País MX → **justo** ⇒ 1; Prog MBA → **falta** (sigue faltando 1) ⇒ 0 → (1,0)
#   → **A2** por (1,0).  
#   **Update**: A2 gana MX y MBA. Slots A2 restantes=0.
# - Paso 5: Cupón **C5(CL,UX)**
#   - A2 no tiene slots.  
#   - A1: País CL → **justo** (espera 1, tiene 1 HIST) ⇒ 1; Prog UX → ya **justo** (con C2) ⇒ 1 → (1,1)  
#   Único posible: **A1**.  
#   **Update**: A1 gana CL y UX. Slots A1 restantes=0.
# 
# **Resultado final en el bloque**
# - **Slots respetados**: A1=3, A2=2.
# - **País (HIST+FRESH)**: A1 {ES:2, MX:1, CL:2} vs esperado {1,1,1} ⇒ aún justos en MX y cerca en ES/CL (ejemplo educativo).  
#   A2 {ES:1, MX:2, CL:0} vs esperado {1,1,0} ⇒ cumple ES/CL y se aproxima en MX.
# - **Programa (HIST+FRESH)**: A1 {MBA:2, UX:2, Data:2} vs esperado {1,1,1}; A2 {MBA:1, UX:1, Data:0} vs {1,1,0}.  
#   *Nota*: Los números exactos dependen del HIST y del total por valor; el ejemplo muestra la **lógica de decisión**, no el cierre perfecto.
# - **Invariantes**: `#FRESH por equipo` (A1=3, A2=2), `FRESH×PILAR`, y `FRESH×PILAR×DV` **inmutables**.
# 
# > Si ningún movimiento mejora (País/Programa) sin romper slots, el bloque se deja **tal cual** y queda registrado con `[WARN]`.
# 
# ---
# 
#  Anexo B · Diagrama de flujo (ASCII)
# ```
# ┌───────────────────────────────────────────────┐
# │ INICIO 2ª ETAPA                              │
# └───────────────┬──────────────────────────────┘
#                 │
#                 ▼
#      DEDUPE REAP para métricas
#                 │
#                 ▼
#    Foto ANTES (global): #FRESH, H+F eq×pilar,
#              cadencia
#                 │
#                 ▼
#       Por cada ÁREA en {A,B,C,T,E}
#                 │
#                 ▼
#         Agrupar equipos (reglas área)
#                 │
#                 ▼
#       Por cada GRUPO → por cada PILAR
#                 │              │
#                 │              └─(en T) segmentar por DV
#                 ▼
#        Calcular SLOTS por equipo
#                 │
#                 ▼
#   Calcular ESPERADO (País/Programa)
#                 │
#                 ▼
#   Reasignación GREEDY (lexicográfico):
#    - evalúa coste (País>Programa)
#    - asigna cupón al mejor equipo
#    - actualiza contadores H+F
#                 │
#                 ▼
#   Verificar slots del segmento (= invariante)
#                 │
#                 ▼
#   Validar grupo: FRESH eq, FRESH×PILAR,
#       (H+F) eq×pilar, (en T) F×P×DV
#                 │
#        ┌────────┴─────────┐
#        │                  │
#        ▼                  ▼
#    Grupo OK           Grupo NO-OK
#    (acumular)          (descartar)
#        │
#        ▼
#  Validar ÁREA (commit/rollback)
#        │
#        ▼
# Aplicar cambios aprobados globalmente
#        │
#        ▼
#  Foto DESPUÉS (global) + asserts
#   #FRESH, H+F eq×pilar, FRESH×PILAR,
#            cadencia
#        │
#        ▼
#   Reinyectar REAP (sin duplicar)
#        │
#        ▼
#                  FIN
# ```
# 
# ---
# 
# Anexo C · Consejos prácticos
# - Asegúrate de que `PILAR_NORM` ∈ {Web, Buscadores, P.Verticales, Redes Sociales}.
# - En **T**, si existe `DV`, deja el lock activo (`{'T':['DV']}`) para respetar **FRESH×PILAR×DV**.
# - Si un área/grupo sale con `[WARN]`, es habitual cuando el bloque ya está cercano al objetivo o no hay cupones con los valores necesarios para cubrir déficits.
# 
# 

# In[22]:


# ============================================
#  SEGUNDA ETAPA – V24 (slots exactos + checks)

PILLARS = ['Web', 'Buscadores', 'P.Verticales', 'Redes Sociales']
IG_teams = {'Equipo_A1', 'Equipo_B1', 'Equipo_C1'}
XP_teams = {'Equipo_A2', 'Equipo_B2', 'Equipo_C2'}

# -----------------------
# Utilidades básicas
# -----------------------
def _ensure_columns(df: pd.DataFrame, cols: Sequence[str]):
    faltan = [c for c in cols if c not in df.columns]
    if faltan:
        raise KeyError(f"Faltan columnas requeridas: {faltan}")

def _pick_key_column(df: pd.DataFrame, prefer: Optional[str]=None) -> str:
    candidates = [
        prefer,
        'ID Oportunidad', 'Id Oportunidad (Oportunidad)', 'ID_OPORTUNIDAD',
        'ID_CUPON', 'INDEX_ORIGINAL', 'ID', 'Id'
    ]
    for c in candidates:
        if c and c in df.columns:
            return c
    if 'INDEX_ORIGINAL' in df.columns:
        return 'INDEX_ORIGINAL'
    raise KeyError("No encuentro una columna ID única (pasa 'prefer' con el nombre exacto de tu ID).")

def _fill_na(series: pd.Series) -> pd.Series:
    return series.fillna('<NA>').astype(str)

# ------------------------------------------
# Esperado vs actual por variable (país/programa)
# ------------------------------------------
def _build_expected_by_pillar_value(
    df_hist_p: pd.DataFrame,
    df_fresh_p: pd.DataFrame,
    equipos: Sequence[str],
    pesos_norm: pd.Series,
    variable: str
) -> pd.DataFrame:
    equipos = [str(e) for e in equipos]
    hist_v  = _fill_na(df_hist_p[variable]) if variable in df_hist_p.columns else pd.Series([], dtype=object)
    fresh_v = _fill_na(df_fresh_p[variable]) if variable in df_fresh_p.columns else pd.Series([], dtype=object)

    tot_val = (hist_v.value_counts() + fresh_v.value_counts()).fillna(0).astype(int)
    valores = sorted(tot_val.index.tolist())
    if not valores:
        return pd.DataFrame(0, index=equipos, columns=[], dtype=int)

    exp = pd.DataFrame(0.0, index=equipos, columns=valores)
    for v in valores:
        exp[v] = (pesos_norm * float(tot_val[v]))
    flo = np.floor(exp).astype(int)
    resid = (exp.sum(axis=0) - flo.sum(axis=0)).astype(int)
    dec = (exp - flo)

    for v in valores:
        r = int(resid.get(v, 0))
        if r > 0:
            order = dec[v].sort_values(ascending=False).index.tolist()
            for k in order[:r]:
                flo.at[k, v] += 1
    return flo.astype(int)

def _actual_accum_by_pillar_value(
    df_hist_p: pd.DataFrame,
    df_fresh_p: pd.DataFrame,
    equipos: Sequence[str],
    variable: str
) -> pd.DataFrame:
    equipos = [str(e) for e in equipos]
    dfh = df_hist_p.copy()
    dff = df_fresh_p.copy()
    if variable in dfh.columns: dfh[variable] = _fill_na(dfh[variable])
    if variable in dff.columns: dff[variable] = _fill_na(dff[variable])

    hist = (dfh.groupby(['EQUIPO_FINAL', variable]).size()
            .unstack(fill_value=0).reindex(index=equipos, fill_value=0))
    fres = (dff.groupby(['EQUIPO_FINAL', variable]).size()
            .unstack(fill_value=0).reindex(index=equipos, fill_value=0))
    actual = (hist.add(fres, fill_value=0)).astype(int)
    cols = sorted(set(actual.columns))
    return actual.reindex(columns=cols, fill_value=0).astype(int)

# -----------------------
# Agrupación de equipos
# -----------------------
def _groups_for_area(
    area: str,
    equipos: Sequence[str],
    ig_set: Set[str] = IG_teams,
    xp_set: Set[str] = XP_teams,
) -> List[List[str]]:
    """
    Reglas de agrupación por área:
    - 'T': agrupa por base completa (A/B/C/...) → cada base es un grupo (p.ej. [C1,C2,C3]).
    - 'E': separa IG / XP / otros.
    - 'C': **nuevo** → agrupa por base completa (p.ej. [C1,C2,C3] en un único grupo).
    - 'A' y 'B': empareja 1 con 2 si existe y deja sueltos los demás (comportamiento previo).
    """
    equipos = sorted(str(e) for e in equipos if pd.notna(e))

    if area == 'E':
        IG = sorted([e for e in equipos if e in ig_set])
        XP = sorted([e for e in equipos if e in xp_set])
        out = []
        if IG: out.append(IG)
        if XP: out.append(XP)
        otros = sorted([e for e in equipos if e not in ig_set and e not in xp_set])
        out += [[e] for e in otros]
        return out

    # Construir mapa base -> lista de equipos (p.ej. base 'C' -> [Equipo_C1, Equipo_C2, Equipo_C3])
    base_map: Dict[str, List[str]] = {}
    for e in equipos:
        suf = e.split('_', 1)[-1] if '_' in e else e
        base = suf[0] if suf else ''
        base_map.setdefault(base, []).append(e)

    groups: List[List[str]] = []
    for base, lst in sorted(base_map.items()):
        lst = sorted(lst)
        if area in ('T', 'C'):
            # NUEVO: en 'C' hacemos lo mismo que en 'T': grupo completo por base (ej. [C1, C2, C3])
            groups.append(lst)
            continue
        # Áreas 'A' y 'B': comportamiento anterior (empareja 1 con 2; resto sueltos)
        has1 = [x for x in lst if x.endswith('1')]
        has2 = [x for x in lst if x.endswith('2')]
        used = set()
        if has1 and has2:
            groups.append([has1[0], has2[0]]); used.update([has1[0], has2[0]])
        for x in lst:
            if x not in used:
                groups.append([x])
    return groups


# -----------------------
# Validación / métricas
# -----------------------
def _counts_fresh_by_team(df_fresh: pd.DataFrame) -> pd.Series:
    return df_fresh['EQUIPO_FINAL'].value_counts().sort_index()

def _counts_fresh_by_team_pilar(df_fresh: pd.DataFrame) -> pd.DataFrame:
    return (df_fresh.groupby(['EQUIPO_FINAL','PILAR_NORM']).size()
            .unstack(fill_value=0).reindex(columns=PILLARS, fill_value=0))

def _counts_histplusfresh_by_team_pilar(df_hist: pd.DataFrame, df_fresh: pd.DataFrame) -> pd.DataFrame:
    both = pd.concat([df_hist[['EQUIPO_FINAL','PILAR_NORM']], df_fresh[['EQUIPO_FINAL','PILAR_NORM']]], ignore_index=True)
    return (both.groupby(['EQUIPO_FINAL','PILAR_NORM']).size()
            .unstack(fill_value=0).reindex(columns=PILLARS, fill_value=0))

def _counts_fresh_by_team_pilar_lock(df_fresh: pd.DataFrame, lock_cols: Sequence[str]) -> pd.DataFrame:
    if not lock_cols:
        return _counts_fresh_by_team_pilar(df_fresh)
    group_cols = ['EQUIPO_FINAL','PILAR_NORM'] + list(lock_cols)
    c = (df_fresh.groupby(group_cols).size().rename('CNT').reset_index())
    c[lock_cols] = c[lock_cols].astype(str).apply(lambda s: s.fillna('<NA>'))
    c['LOCK_KEY'] = c[lock_cols].astype(str).agg('|'.join, axis=1)
    pivot = c.pivot_table(index=['EQUIPO_FINAL','PILAR_NORM','LOCK_KEY'], values='CNT', aggfunc='sum').sort_index()
    return pivot

def _assert_equal_msg(left, right, msg: str):
    if isinstance(left, pd.DataFrame):
        left  = left.sort_index().sort_index(axis=1)
        right = right.reindex(index=left.index, columns=left.columns, fill_value=0)
        equal = left.equals(right)
    elif isinstance(left, pd.Series):
        left  = left.sort_index()
        right = right.reindex(left.index, fill_value=0)
        equal = left.equals(right)
    else:
        equal = (left == right)
    assert equal, msg

def _cadencia_por_equipo(df_hist_total_clean: pd.DataFrame,
                         df_fresh_like: pd.DataFrame,
                         df_horas_eq: pd.DataFrame) -> pd.Series:
    _ensure_columns(df_horas_eq, ['EQUIPO','HORAS'])
    hist = df_hist_total_clean[['EQUIPO_FINAL']].assign(TIPO='HIST')
    fresh = df_fresh_like[['EQUIPO_FINAL']].assign(TIPO='FRESH')
    tot = pd.concat([hist, fresh], ignore_index=True)
    por_equipo = tot['EQUIPO_FINAL'].value_counts().sort_index()
    h = (df_horas_eq.set_index('EQUIPO')['HORAS']
         .reindex(por_equipo.index)
         .astype(float))
    if h.isna().any() or (h <= 0).any():
        malos = h[h.isna() | (h<=0)].index.tolist()
        raise ValueError(f"Equipos sin horas válidas para cadencia: {malos}")
    cad = (por_equipo / (h / 6.0)).rename('CADENCIA')
    return cad

# ------------------------------------------
# Núcleo: reasignación por slots (weighted País+Programa)
# ------------------------------------------
def _balanced_reassign_pillar(
    df_hist_p: pd.DataFrame,
    df_fresh_p: pd.DataFrame,
    grupo: Sequence[str],
    pesos_norm_g: pd.Series,
    var_pais: str,
    var_prog: str,
    w_country: float,
    w_program: float,
    lock_cols: Optional[Sequence[str]] = None
) -> pd.DataFrame:
    """
    Reasigna SOLO dentro del pilar (df_fresh_p) con slots exactos:
      - slots por equipo = #FRESH actuales del equipo en este pilar (y lock si aplica).
      - si lock_cols != None: slots por (equipo × lock_tuple) (ej. DV en T).
    Objetivo: minimizar la **suma ponderada de la variación de error absoluto**
    frente al esperado (HIST+FRESH) en País y Programa:
        coste = w_country * Δerror_país + w_program * Δerror_programa
    donde Δerror_v = |(act+1) - exp| - |act - exp| para el valor v del cupón.
    """
    if df_fresh_p.empty:
        return df_fresh_p

    lock_cols = list(lock_cols or [])
    dff = df_fresh_p.copy()
    dfh = df_hist_p.copy()

    # Normalizar valores a '<NA>'
    for col in [var_pais, var_prog] + lock_cols:
        if col in dff.columns: dff[col] = _fill_na(dff[col])
        if col in dfh.columns: dfh[col] = _fill_na(dfh[col])

    # Esperados y actuales (HIST+FRESH) por país y programa
    exp_c = _build_expected_by_pillar_value(dfh, dff, grupo, pesos_norm_g, var_pais)
    act_c = _actual_accum_by_pillar_value  (dfh, dff, grupo, var_pais)
    all_c = sorted(set(exp_c.columns) | set(act_c.columns))
    exp_c = exp_c.reindex(index=grupo, columns=all_c, fill_value=0).astype(int)
    act_c = act_c.reindex(index=grupo, columns=all_c, fill_value=0).astype(int)

    exp_p = _build_expected_by_pillar_value(dfh, dff, grupo, pesos_norm_g, var_prog)
    act_p = _actual_accum_by_pillar_value  (dfh, dff, grupo, var_prog)
    all_p = sorted(set(exp_p.columns) | set(act_p.columns))
    exp_p = exp_p.reindex(index=grupo, columns=all_p, fill_value=0).astype(int)
    act_p = act_p.reindex(index=grupo, columns=all_p, fill_value=0).astype(int)

    # Segmentos por lock
    if lock_cols:
        def _lock_key_row(r: pd.Series) -> Tuple:
            return tuple(r.get(c, '<NA>') for c in lock_cols)
        segments: List[Tuple[Tuple, List[int]]] = []
        for key, sub in dff.groupby(dff.apply(_lock_key_row, axis=1)):
            segments.append((key, sub.index.tolist()))
    else:
        segments = [((), dff.index.tolist())]

    def _segment_team_counts(idxs: List[int]) -> pd.Series:
        return dff.loc[idxs, 'EQUIPO_FINAL'].value_counts().reindex(grupo, fill_value=0).astype(int)

    # Coste ponderado: reducción/aumento de error absoluto (menor es mejor)
    # Δerror = |(act+1)-exp| - |act-exp|  → -1 mejora, +1 empeora, 0 indiferente
    def marginal_penalty(team: str, pais_val: str, prog_val: str) -> Tuple[float, float, float]:
        # País
        a_c = int(act_c.at[team, pais_val]) if pais_val in act_c.columns else 0
        e_c = int(exp_c.at[team, pais_val]) if pais_val in exp_c.columns else 0
        delta_c = abs((a_c + 1) - e_c) - abs(a_c - e_c)

        # Programa
        a_p = int(act_p.at[team, prog_val]) if prog_val in act_p.columns else 0
        e_p = int(exp_p.at[team, prog_val]) if prog_val in exp_p.columns else 0
        delta_p = abs((a_p + 1) - e_p) - abs(a_p - e_p)

        total = (w_country * float(delta_c)) + (w_program * float(delta_p))
        return total, float(delta_c), float(delta_p)  # total para ordenar, y (delta_c, delta_p) para desempate

    EPS = 1e-12  # tolerancia para empates numéricos

    # Asignación por segmentos con slots exactos
    for seg_key, idxs in segments:
        if not idxs:
            continue

        seg_counts_before = _segment_team_counts(idxs)
        slots = {t: int(seg_counts_before[t]) for t in grupo if int(seg_counts_before[t]) > 0}
        if not slots:
            continue

        remaining = slots.copy()
        remaining_coupons = list(idxs)  # orden estable
        assigned_team: Dict[int, str] = {}

        # Greedy ponderado: para cada cupón, elige equipo con slot que minimice el coste total
        # Desempate: menor (delta_c, delta_p) para preservar una prioridad suave a País
        while remaining_coupons:
            best_pair = None
            best_score = None  # (total_cost, delta_c, delta_p)

            for i in remaining_coupons:
                ri = dff.loc[i]
                pais_i = ri[var_pais]
                prog_i = ri[var_prog]
                for t, cap in list(remaining.items()):
                    if cap <= 0:
                        continue
                    total, dc, dp = marginal_penalty(t, pais_i, prog_i)
                    score = (total, dc, dp)
                    if best_score is None:
                        best_score = score
                        best_pair = (i, t, pais_i, prog_i)
                    else:
                        # Primero por coste total; si casi empata, desempata por (dc, dp)
                        if (total < best_score[0] - EPS) or \
                           (abs(total - best_score[0]) <= EPS and (dc < best_score[1] - EPS or
                                                                   (abs(dc - best_score[1]) <= EPS and dp < best_score[2] - EPS))):
                            best_score = score
                            best_pair = (i, t, pais_i, prog_i)

            # Asignar y actualizar “actuales” (HIST+FRESH) del equipo elegido
            i, t, pais_i, prog_i = best_pair
            assigned_team[i] = t
            remaining[t] -= 1
            if remaining[t] == 0:
                del remaining[t]
            remaining_coupons.remove(i)

            if pais_i in act_c.columns:
                act_c.at[t, pais_i] += 1
            if prog_i in act_p.columns:
                act_p.at[t, prog_i] += 1

        # Aplicar reasignación del segmento
        for i, t in assigned_team.items():
            dff.at[i, 'EQUIPO_FINAL'] = t

        # Verificación dura del segmento (slots invariantes)
        seg_counts_after = dff.loc[idxs, 'EQUIPO_FINAL'].value_counts().reindex(grupo, fill_value=0).astype(int)
        _assert_equal_msg(seg_counts_after, seg_counts_before,
                          f"Slots del segmento {seg_key} cambiaron dentro del pilar (lock={lock_cols}).")

    return dff


# ------------------------------------------
# Núcleo 2ª etapa (transaccional por ÁREA + slots)
# ------------------------------------------
def _segunda_etapa_core(
    df_final_total_clean: pd.DataFrame,
    df_hist_total_clean: pd.DataFrame,
    df_pesos_areas: pd.DataFrame,
    variables: Tuple[str,str]=('Agrupación OBS','Programa de Interes'),
    pillars: Sequence[str]=PILLARS,
    areas: Sequence[str]=('A','B','C','T','E'),
    w_country: float=1.0,
    w_program: float=0.9,
    lock_cols_by_area: Optional[Dict[str, List[str]]]=None
) -> pd.DataFrame:

    var_pais, var_prog = variables
    lock_cols_by_area = lock_cols_by_area or {}

    _ensure_columns(df_final_total_clean, ['INDEX_ORIGINAL','TIPO_REPARTO','EQUIPO_FINAL','PILAR_NORM', var_pais, var_prog, 'AREA'])
    _ensure_columns(df_hist_total_clean,  ['EQUIPO_FINAL','PILAR_NORM', var_pais, var_prog])
    _ensure_columns(df_pesos_areas,       ['AREA','EQUIPO','PESO_BASE'])

    res = df_final_total_clean.copy()
    res['EQUIPO_FINAL'] = res['EQUIPO_FINAL'].astype(str)
    df_hist_total_clean = df_hist_total_clean.copy()
    df_hist_total_clean['EQUIPO_FINAL'] = df_hist_total_clean['EQUIPO_FINAL'].astype(str)

    # Baselines globales (blindajes)
    mask_fresh = res['TIPO_REPARTO'].eq('FRESH')
    fresh_baseline_global = _counts_fresh_by_team(res.loc[mask_fresh, ['EQUIPO_FINAL']])
    tp_baseline_global    = _counts_histplusfresh_by_team_pilar(
        df_hist_total_clean, res.loc[mask_fresh, ['EQUIPO_FINAL','PILAR_NORM']]
    )
    fresh_pilar_before_global = _counts_fresh_by_team_pilar(res.loc[mask_fresh, ['EQUIPO_FINAL','PILAR_NORM']])

    updates_all: List[pd.DataFrame] = []
    processed_idx: Set[int] = set()

    for area in areas:
        equipos_area = (df_pesos_areas[df_pesos_areas['AREA']==area]['EQUIPO']
                        .dropna().astype(str).unique().tolist())
        if not equipos_area:
            continue

        fresh_area = res[mask_fresh & res['EQUIPO_FINAL'].isin(equipos_area) & res['AREA'].eq(area)].copy()
        if fresh_area.empty:
            continue
        hist_area  = df_hist_total_clean[df_hist_total_clean['EQUIPO_FINAL'].isin(equipos_area)].copy()

        # Snapshots área (antes)
        fp_area_before    = _counts_fresh_by_team_pilar(fresh_area)
        fresh_area_before = _counts_fresh_by_team(fresh_area[['EQUIPO_FINAL']])

        # DV lock en T (si columna existe y no la pasaron)
        area_lock_cols = list(lock_cols_by_area.get(area, []))
        if area == 'T' and 'DV' in fresh_area.columns and 'DV' not in area_lock_cols:
            area_lock_cols = ['DV']

        # Baseline FRESH×PILAR×LOCK por área (si aplica)
        if area_lock_cols:
            fp_lock_area_before = _counts_fresh_by_team_pilar_lock(fresh_area, area_lock_cols)

        # Pesos de área
        pesos_area = (df_pesos_areas[df_pesos_areas['AREA']==area]
                      .set_index('EQUIPO')['PESO_BASE'].astype(float))
        pesos_area = pesos_area.reindex(equipos_area).fillna(0.0)
        if pesos_area.sum() <= 0:
            pesos_area = pd.Series(1.0/len(equipos_area), index=equipos_area, dtype=float)
        else:
            pesos_area = pesos_area / float(pesos_area.sum())

        grupos = _groups_for_area(area, equipos_area)

        area_updates_parts: List[pd.DataFrame] = []
        processed_idx_area: Set[int] = set()

        for grupo in grupos:
            if len(grupo) <= 1:
                continue

            fresh_g = fresh_area[fresh_area['EQUIPO_FINAL'].isin(grupo)].copy()
            if fresh_g.empty:
                continue
            hist_g  = hist_area[hist_area['EQUIPO_FINAL'].isin(grupo)].copy()

            # Baselines duros del grupo
            baseline_pilar_g = _counts_histplusfresh_by_team_pilar(hist_g, fresh_g).reindex(index=grupo, fill_value=0)
            fresh_before     = _counts_fresh_by_team(fresh_g[['EQUIPO_FINAL']]).reindex(grupo, fill_value=0)
            fp_before        = _counts_fresh_by_team_pilar(fresh_g).reindex(index=grupo, fill_value=0)

            if area_lock_cols:
                fp_lock_before = _counts_fresh_by_team_pilar_lock(fresh_g, area_lock_cols)

            # Pesos normalizados dentro del grupo
            pesos_norm_g = pesos_area.reindex(grupo).fillna(0.0)
            s = float(pesos_norm_g.sum())
            if s <= 0:
                pesos_norm_g = pd.Series(1.0/len(grupo), index=grupo, dtype=float)
            else:
                pesos_norm_g = (pesos_norm_g / s).astype(float)

            parts_grp: List[pd.DataFrame] = []
            for p in PILLARS:
                df_fresh_p = fresh_g[fresh_g['PILAR_NORM']==p].copy()
                if df_fresh_p.empty:
                    continue
                df_hist_p  = hist_g[hist_g['PILAR_NORM']==p].copy()

                df_opt = _balanced_reassign_pillar(
                    df_hist_p=df_hist_p,
                    df_fresh_p=df_fresh_p,
                    grupo=grupo,
                    pesos_norm_g=pesos_norm_g,
                    var_pais=var_pais, var_prog=var_prog,
                    w_country=w_country, w_program=w_program,
                    lock_cols=area_lock_cols
                )
                parts_grp.append(df_opt[['INDEX_ORIGINAL','EQUIPO_FINAL','PILAR_NORM']])

            if not parts_grp:
                continue

            upd_g = (pd.concat(parts_grp, ignore_index=True)
                     .drop_duplicates(subset=['INDEX_ORIGINAL'], keep='last'))
            upd_g = upd_g[~upd_g['INDEX_ORIGINAL'].isin(processed_idx | processed_idx_area)]

            # Aplicar provisional al grupo y validar invariantes duros
            m_g = dict(zip(upd_g['INDEX_ORIGINAL'], upd_g['EQUIPO_FINAL']))
            fresh_g_post = fresh_g.copy()
            fresh_g_post.loc[:, 'EQUIPO_FINAL'] = fresh_g_post['INDEX_ORIGINAL'].map(m_g).fillna(fresh_g_post['EQUIPO_FINAL'])

            fresh_after  = _counts_fresh_by_team(fresh_g_post[['EQUIPO_FINAL']]).reindex(grupo, fill_value=0)
            fp_after     = _counts_fresh_by_team_pilar(fresh_g_post).reindex(index=grupo, fill_value=0)
            post_pilar_g = _counts_histplusfresh_by_team_pilar(hist_g, fresh_g_post).reindex(index=grupo, fill_value=0)

            inv_pilar_ok       = post_pilar_g.equals(baseline_pilar_g)
            inv_fresh_ok       = fresh_after.equals(fresh_before)
            inv_fresh_pilar_ok = fp_after.equals(fp_before)

            inv_lock_ok = True
            if area_lock_cols:
                fp_lock_after = _counts_fresh_by_team_pilar_lock(fresh_g_post, area_lock_cols)
                inv_lock_ok = fp_lock_after.equals(fp_lock_before)

            if inv_pilar_ok and inv_fresh_ok and inv_fresh_pilar_ok and inv_lock_ok:
                area_updates_parts.append(upd_g[['INDEX_ORIGINAL','EQUIPO_FINAL']])
                processed_idx_area.update(upd_g['INDEX_ORIGINAL'].tolist())
            else:
                print(f"[WARN Área {area} Grupo {grupo}] descartado "
                      f"(H+F eq×pilar={inv_pilar_ok}, FRESH eq={inv_fresh_ok}, "
                      f"FRESH eq×pilar={inv_fresh_pilar_ok}, LOCK={inv_lock_ok}).")

        # Commit por área y validación dura por área
        if area_updates_parts:
            upd_area = (pd.concat(area_updates_parts, ignore_index=True)
                        .drop_duplicates(subset=['INDEX_ORIGINAL'], keep='last'))
            m_area = dict(zip(upd_area['INDEX_ORIGINAL'], upd_area['EQUIPO_FINAL']))

            fresh_area_post = fresh_area.copy()
            fresh_area_post.loc[:, 'EQUIPO_FINAL'] = fresh_area_post['INDEX_ORIGINAL'].map(m_area).fillna(fresh_area_post['EQUIPO_FINAL'])

            fp_area_after     = _counts_fresh_by_team_pilar(fresh_area_post)
            fresh_area_after  = _counts_fresh_by_team(fresh_area_post[['EQUIPO_FINAL']])

            ok_area = fp_area_after.equals(fp_area_before) and fresh_area_after.equals(fresh_area_before)

            ok_area_lock = True
            if area_lock_cols:
                fp_lock_area_after = _counts_fresh_by_team_pilar_lock(fresh_area_post, area_lock_cols)
                ok_area_lock = fp_lock_area_after.equals(fp_lock_area_before)

            if ok_area and ok_area_lock:
                updates_all.append(upd_area[['INDEX_ORIGINAL','EQUIPO_FINAL']])
                processed_idx.update(processed_idx_area)
            else:
                print(f"[ROLLBACK Área {area}] cambios descartados "
                      f"(FRESH×PILAR / FRESH por equipo / LOCK no invariante).")

    # Aplicar updates aprobados (global)
    if updates_all:
        updates = (pd.concat(updates_all, ignore_index=True)
                   .drop_duplicates(subset=['INDEX_ORIGINAL'], keep='last'))
        m = dict(zip(updates['INDEX_ORIGINAL'], updates['EQUIPO_FINAL']))
        res.loc[mask_fresh, 'EQUIPO_FINAL'] = res.loc[mask_fresh, 'INDEX_ORIGINAL'].map(m)\
                                                .fillna(res.loc[mask_fresh, 'EQUIPO_FINAL'])

    res = res.sort_values('INDEX_ORIGINAL', kind='stable').reset_index(drop=True)

    # Blindajes globales (post)
    fresh_after_global = _counts_fresh_by_team(res.loc[mask_fresh, ['EQUIPO_FINAL']])\
                         .reindex(fresh_baseline_global.index, fill_value=0)
    _assert_equal_msg(fresh_after_global, fresh_baseline_global, "2ª etapa: cambió el #FRESH por equipo.")

    tp_after_global = _counts_histplusfresh_by_team_pilar(
        df_hist_total_clean, res.loc[mask_fresh, ['EQUIPO_FINAL','PILAR_NORM']]
    ).reindex(index=tp_baseline_global.index, columns=tp_baseline_global.columns, fill_value=0)
    _assert_equal_msg(tp_after_global, tp_baseline_global, "2ª etapa: cambió el (equipo×pilar) HIST+FRESH.")

    fresh_pilar_after_global = _counts_fresh_by_team_pilar(res.loc[mask_fresh, ['EQUIPO_FINAL','PILAR_NORM']])\
                               .reindex(index=fresh_pilar_before_global.index, columns=fresh_pilar_before_global.columns, fill_value=0)
    _assert_equal_msg(fresh_pilar_after_global, fresh_pilar_before_global, "2ª etapa: cambió el FRESH×PILAR global.")

    return res

# ------------------------------
# Dedupe REAP para métricas
# ------------------------------
def dedup_reap_hist_vs_final(
    df_hist_total: pd.DataFrame,
    df_final_total: pd.DataFrame,
    prefer_key: Optional[str]=None,
    verbose: bool=True
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    dfh = df_hist_total.copy()
    dff = df_final_total.copy()
    for (name, df, cols) in [
        ("df_hist_total", dfh, ['EQUIPO_FINAL']),
        ("df_final_total", dff, ['EQUIPO_FINAL','TIPO_REPARTO'])
    ]:
        _ensure_columns(df, cols)

    if 'REAP' not in set(dff['TIPO_REPARTO']):
        if verbose: print("[DEDUP] No hay REAP en df_final_total → nada que deduplicar.")
        return dfh, dff

    key_h = _pick_key_column(dfh, prefer=prefer_key)
    key_f = _pick_key_column(dff, prefer=prefer_key)

    reap_final = dff.loc[dff['TIPO_REPARTO'].eq('REAP') & dff[key_f].notna(), [key_f]].drop_duplicates()
    if reap_final.empty:
        if verbose: print("[DEDUP] df_final_total no tiene REAP con clave no nula.")
        return dfh, dff

    keys_hist = set(dfh.loc[dfh[key_h].notna(), key_h].astype(str).unique())
    keys_final_reap = set(reap_final[key_f].astype(str).unique())
    dup_keys = keys_final_reap & keys_hist

    if not dup_keys:
        if verbose: print("[DEDUP] No hay REAP duplicadas entre HIST y FINAL.")
        return dfh, dff

    mask_drop = dff['TIPO_REPARTO'].eq('REAP') & dff[key_f].astype(str).isin(dup_keys)
    n_drop = int(mask_drop.sum())
    dff = dff.loc[~mask_drop].copy()
    if verbose:
        print(f"[DEDUP] REAP duplicadas detectadas por clave: {len(dup_keys)} únicas.")
        print(f"[DEDUP] Filas REAP eliminadas de df_final_total: {n_drop} (se cuentan como HIST).")
    return dfh, dff

# -------------------------------------------------------
# Reconstruir export final con REAP (reinyección final)
# -------------------------------------------------------
def construir_final_con_reap(
    df_final_total_original: pd.DataFrame,
    df_result_2a_etapa: pd.DataFrame,
    prefer_key: Optional[str]=None
) -> pd.DataFrame:
    key_res = _pick_key_column(df_result_2a_etapa, prefer=prefer_key)
    key_org = _pick_key_column(df_final_total_original, prefer=prefer_key)

    mask_reap_org = df_final_total_original['TIPO_REPARTO'].eq('REAP')
    df_reap_org = df_final_total_original.loc[mask_reap_org].copy()

    base = df_result_2a_etapa.copy()

    if df_reap_org.empty or key_org not in df_reap_org.columns:
        df_final_ajustado = base
    else:
        keys_base = set(base.loc[base[key_res].notna(), key_res].astype(str).unique())
        reap_to_add = df_reap_org.loc[
            df_reap_org[key_org].notna() &
            ~df_reap_org[key_org].astype(str).isin(keys_base)
        ].copy()
        df_final_ajustado = pd.concat([base, reap_to_add], ignore_index=True)
        if key_res in df_final_ajustado.columns:
            df_final_ajustado = (df_final_ajustado
                                 .sort_values(by=[key_res])
                                 .drop_duplicates(subset=[key_res], keep='first')
                                 .reset_index(drop=True))

    if 'INDEX_ORIGINAL' in df_final_ajustado.columns:
        df_final_ajustado = df_final_ajustado.sort_values('INDEX_ORIGINAL', kind='stable').reset_index(drop=True)

    print(f"[FINAL] REAP presentes en df_final_ajustado: {int((df_final_ajustado['TIPO_REPARTO']=='REAP').sum())}")
    return df_final_ajustado

# -----------------------
# Runner público (v19)
# -----------------------
def run_segunda_etapa_v19(
    df_final_total: pd.DataFrame,
    df_hist_total: pd.DataFrame,
    df_pesos_areas: pd.DataFrame,
    df_horas_eq: pd.DataFrame,
    variables: Tuple[str,str]=('Agrupación OBS','Programa de Interes'),
    pillars: Sequence[str]=PILLARS,
    areas: Sequence[str]=('A','B','C','T','E'),
    max_iter_vals: int=60,                 # compat (no usado)
    max_pairs_sample: int=40,              # compat (no usado)
    pillar_weights: Dict[str,float]=None,  # compat (no usado)
    prefer_key: Optional[str]=None,
    verbose: bool=True,
    country_soft_abs: int=1,               # compat (no usado)
    country_soft_rel: float=0.05,          # compat (no usado)
    w_country: float=1.0,
    w_program: float=0.9,
    lock_cols_by_area: Optional[Dict[str, List[str]]]=None  # p.ej. {'T': ['DV']}
) -> Tuple[pd.DataFrame, pd.DataFrame]:

    _ensure_columns(df_final_total, ['INDEX_ORIGINAL','TIPO_REPARTO','EQUIPO_FINAL','PILAR_NORM','AREA'])
    mask_fresh_orig = df_final_total['TIPO_REPARTO'].eq('FRESH')
    assert df_final_total.loc[mask_fresh_orig, 'INDEX_ORIGINAL'].is_unique, "INDEX_ORIGINAL duplicado en FRESH."

    # 1) DEDUPE (REAP cuentan como HIST para métricas)
    df_hist_total_clean, df_final_total_clean = dedup_reap_hist_vs_final(
        df_hist_total=df_hist_total,
        df_final_total=df_final_total,
        prefer_key=prefer_key,
        verbose=verbose
    )

    # Fotos antes (CLEAN)
    mask_fresh_clean = df_final_total_clean['TIPO_REPARTO'].eq('FRESH')
    cnt_fresh_before = _counts_fresh_by_team(
        df_final_total_clean.loc[mask_fresh_clean, ['EQUIPO_FINAL']]
    )
    tp_before = _counts_histplusfresh_by_team_pilar(
        df_hist_total_clean,
        df_final_total_clean.loc[mask_fresh_clean, ['EQUIPO_FINAL','PILAR_NORM']]
    )
    cad_antes = _cadencia_por_equipo(
        df_hist_total_clean,
        df_final_total_clean.loc[mask_fresh_clean, ['EQUIPO_FINAL']],
        df_horas_eq
    )

    # 2) Core (slots exactos, locks por área, transacciones por grupo)
    res_2a = _segunda_etapa_core(
        df_final_total_clean=df_final_total_clean,
        df_hist_total_clean=df_hist_total_clean,
        df_pesos_areas=df_pesos_areas,
        variables=variables,
        pillars=pillars,
        areas=areas,
        w_country=w_country,
        w_program=w_program,
        lock_cols_by_area=lock_cols_by_area
    )

    # 3) Postchecks globales duros
    mask_fresh_after = res_2a['TIPO_REPARTO'].eq('FRESH')
    cnt_fresh_after = _counts_fresh_by_team(res_2a.loc[mask_fresh_after, ['EQUIPO_FINAL']])\
                      .reindex(cnt_fresh_before.index, fill_value=0)
    tp_after = _counts_histplusfresh_by_team_pilar(
        df_hist_total_clean,
        res_2a.loc[mask_fresh_after, ['EQUIPO_FINAL','PILAR_NORM']]
    ).reindex(index=tp_before.index, columns=tp_before.columns, fill_value=0)

    cad_desp = _cadencia_por_equipo(
        df_hist_total_clean,
        res_2a.loc[mask_fresh_after, ['EQUIPO_FINAL']],
        df_horas_eq
    )

    if verbose:
        print("ΔFRESH:", (cnt_fresh_after - cnt_fresh_before).to_dict())
        print("Δ(Equipo×Pilar):", (tp_after - tp_before).stack()[lambda s: s!=0].to_dict() if not tp_after.equals(tp_before) else {})
        print("ΔCadencia:", (cad_desp - cad_antes).round(12).to_dict())

    _assert_equal_msg(cnt_fresh_after, cnt_fresh_before, "Cambió el #FRESH por equipo.")
    _assert_equal_msg(tp_after, tp_before, "Cambió el (equipo×pilar) HIST+FRESH.")
    diff_cad = (cad_desp - cad_antes).abs()
    assert (diff_cad < 1e-9).all(), f"Cadencia cambió en: {diff_cad[diff_cad>=1e-9].to_dict()}"

    # 4) Export final: reinyectar REAP originales sin duplicar
    df_final_ajustado = construir_final_con_reap(
        df_final_total_original=df_final_total,
        df_result_2a_etapa=res_2a,
        prefer_key=prefer_key
    )

    return res_2a, df_final_ajustado


# ### CALL FUNCIÓN

# In[23]:


# res_2a: distribución ajustada (solo FRESH dentro del clean)
# final_export: con REAP reinyectadas (sin duplicar)
res_2a, df_final_ajustado = run_segunda_etapa_v19(
    df_final_total=df_final_total,
    df_hist_total=df_hist_total,
    df_pesos_areas=df_pesos_areas,
    df_horas_eq=df_horas_eq,
    # puedes ajustar la “suavidad” aquí:
    country_soft_abs=1,      # permite +1 de error en país
    country_soft_rel=0.05,   # o +5% (toma el mayor)
    w_country=1.0,
    w_program=0.9
)


df_final_ajustado.to_excel("Distribucion_Final.xlsx", index=False)


# ## FINAL 

# In[24]:


print("✅ Distribución realizada correctamente ✅")
input("\nPresiona ENTER para cerrar la aplicación...")

