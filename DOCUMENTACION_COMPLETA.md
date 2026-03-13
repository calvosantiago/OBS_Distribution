# Documentacion Completa - Distribucion OBS

## 1) Objetivo del proyecto
Este proyecto distribuye cupones abiertos entre equipos comerciales OBS respetando reglas de negocio, pesos operativos (horas), cadencia, pilar, y restricciones de reaperturas.

La salida principal es un Excel final con todos los cupones del dia (fresh + reap + especiales) y el `EQUIPO_FINAL` asignado.

Reglas de negocio 
 - Áreas A, B, C: distribución proporcional según horas trabajadas.
 - Área T: programas MBA en español.
 - Área E: programas MBA y MST en inglés (sin cadencia, solo pesos teóricos).
 - Reaperturas: se asignan primero al equipo original si cumple criterios.
 - Priorización en ajustes: País → Programa → Pilares.
 - Pilares protegidos: Web y Buscadores mantienen proporción acumulada.


## 2) Arquitectura general
Entrada de datos -> Preproceso y enriquecimiento -> Etapa 1 (asignacion por area) -> Etapa 2 (rebalanceo interno sin romper invariantes) -> Export a Excel.

### Componentes del codigo
- `run.py`: punto de entrada de ejecucion.
- `distribucion_obs/config.py`: configuracion de rutas y parametros.
- `distribucion_obs/loaders.py`: lectura de archivos fuente Excel.
- `distribucion_obs/preprocess.py`: limpieza, enriquecimiento, separacion fresh/reap/historico y calculos preliminares.
- `distribucion_obs/stages.py`: orquestacion de etapa 1 y etapa 2.
- `distribucion_obs/pipeline.py`: pipeline end-to-end y controles globales.
- `distribucion_obs/extracted_functions.py`: motor principal de distribucion (A/B/C/T/E) y segunda etapa.

## 3) Orquestacion exacta del pipeline
Flujo principal en `run.py` -> `run_and_export()` -> `run_pipeline()`.

`run_pipeline()` ejecuta (en este orden):
1. `load_base_inputs()`, `load_pilares_map()`, `load_sudoku_raw()`
2. Control de PMAX (solo conteo para trazabilidad)
3. `enrich_with_area_country_pillar()`
4. `build_weights()`
5. `build_hist_qbcn()`
6. `preprocess_open_coupons()`
7. `split_reap_fresh_hist()`
8. `compute_cadencia_preliminar()`
9. `run_first_stage()`
10. `run_second_stage()`
11. Control de filas entrada/salida
12. Export a `Distribucion_Final.xlsx`

## 4) Archivos de entrada y salida

### Entradas principales
1. Archivo de cupones abiertos (se toma el mas reciente en Downloads):
- Prefijo esperado: `Oportunidades abiertas No Asignadas JE Totales...`

2. Archivo historico qb_CN (se toma el mas reciente en Downloads):
- Prefijo esperado: `qb_CN_V3_OBS...`

3. `Areas_Paises.xlsx` (en raiz del proyecto):
- Hoja `Areas` (mapeo programa -> area/tipo/idioma)
- Hoja `Paises` (normalizacion continente)

4. `SUDOKU.xlsx` (en raiz del proyecto):
- Hoja `Estatus diario`:
  - columnas `P:U` para horas por equipo
  - celda `O22` para hora de corte

5. `OBS_ESTRUCTURA_COMERCIAL.xlsx` (ruta configurada en `PipelineConfig`):
- Hoja `ESTRUCTURA`
- Hoja `NORMALIZACION NOMBRES`

6. `OBS_PULL_PUSH.xlsx` (ruta configurada en `PipelineConfig`):
- Hoja `PULL-PUSH` (normalizacion de pilar)

### Salida
- `Distribucion_Final.xlsx` en la raiz del proyecto.

## 5) Librerias usadas

### Runtime real del programa (las importantes)
- `pandas`: transformaciones tabulares y lectura/escritura Excel.
- `numpy`: redondeos robustos y operaciones numericas.
- `pulp`: modelos MILP (asignacion optimizada en areas A/B/C/T).
- `country_converter`: normalizacion de nombres de pais.
- `openpyxl` (via `pandas.read_excel`/`to_excel`).
- Stdlib: `datetime`, `pathlib`, `dataclasses`, `typing`, `re`.

Nota: `requirements.txt` contiene muchas librerias de otros usos del entorno. Para este pipeline, las criticas son las anteriores.

## 6) Configuracion
Archivo: `distribucion_obs/config.py`

Parametros clave:
- `downloads_dir`: carpeta donde se busca el ultimo archivo de abiertos e historico.
- `areas_paises_path`, `sudoku_path`, `estructura_path`, `pilares_path`.
- `output_path`.
- `pilar_band_web`, `pilar_band_busc`: bandas blandas para desviacion de Web/Buscadores.
- `time_limit`: limite del solver CBC (segundos).
- `stage2_w_country`, `stage2_w_program`: pesos de etapa 2.
- `enforce_row_count_control`: control estricto de filas.

## 7) Reglas de negocio implementadas (preproceso)

### 7.1 PMAX
- Historico: PMAX se excluye de calculos.
- Hoy: PMAX no se redistribuye, pero se conserva en salida final como passthrough.

### 7.2 Duplicados por contacto
En abiertos:
- Se normaliza email y telefono.
- Si detecta duplicado (email o telefono), se marca como `Equipo_Z`.
- Si corresponde, `EQUIPO_FINAL` se fija en `Joaquim Barnola Fontrodona`.

### 7.3 Especiales
- `Equipo_Referidos` y `Equipo_Z` se separan como `df_special`.
- No entran al motor de distribucion normal.

### 7.4 Fresh vs Reap
- `TIPO_REPARTO = REAP` si hay `Tipo de Re-Apertura`, si no `FRESH`.
- REAP valida se asigna via estructura comercial (normalizacion de propietario origen).
- Fresh anterior a hora de corte (SUDOKU O22) se pasa a `REAP` (`df_corte`).
- Historico acumulado de calculo = `hist_qbcn + reap_validas + corte`.

### 7.5 Filtros exactos para que un cupon cuente o no cuente

#### A) Filtros para que un cupon entre a distribucion de hoy
- Se excluye de distribucion si subpilar contiene `pmax` (se preserva en salida final como passthrough).
- Si es duplicado por email/telefono normalizado, se marca como `Equipo_Z` (sale del flujo normal).
- Debe pertenecer a equipos comerciales del modelo (`Equipo_A1` a `Equipo_C2`).
- Debe cumplir alguno de estos criterios tipo/idioma:
  - `TIPO = MST` y `IDIOMA = ESP`
  - `TIPO = MBA` y `IDIOMA = ESP`
  - `IDIOMA = ENG`

#### B) Filtros del historico base de calculo (`df_hist_qbcn`)
- Excluye `Equipo Asignado = Equipo_Referidos`.
- Excluye `PILAR_NORM in ['REF/RECUP', 'OTROS']`.
- Solo equipos comerciales (`Equipo_A1` a `Equipo_C2`).
- Mismo criterio tipo/idioma que en distribucion:
  - MST/ESP, MBA/ESP o ENG.

#### C) Filtros para separar `FRESH` y `REAP`
- `REAP` si `Tipo de Re-Apertura` no esta vacio.
- `REAP valida` si se puede mapear a equipo por estructura comercial semanal.
- `FRESH` creado antes del corte (`SUDOKU!O22`) pasa a `REAP` por regla `CUTOFF`.
- `REAP` invalida vuelve a `FRESH` (salvo si cae en corte).

#### D) Filtros de conteo usados por el modelo en Areas A/B/C
- Fresh del area:
  - `AREA` de esa area (`A` o `B` o `C`)
  - `PILAR_NORM in ['Web', 'Buscadores', 'P.Verticales', 'Redes Sociales']`
- Historico para el modelo del area:
  - `EQUIPO_FINAL` en el par del area (ejemplo A1/A2 para area A)
  - `AREA not in ['T', 'E']`
  - `PILAR_NORM` en los 4 pilares del modelo
  - `TIPO` e `IDIOMA` contenidos en el set presente en el fresh de esa area

#### E) Filtro de referencia para cruces de control A1/A2 \"sin T ni E\"
Para reproducir conteos como 359/339:
- `EQUIPO_FINAL in ['Equipo_A1', 'Equipo_A2']`
- `AREA not in ['T', 'E']`
- `PILAR_NORM in ['Web', 'Buscadores', 'P.Verticales', 'Redes Sociales']`
- Base de conteo:
  - historico filtrado + fresh final asignado del dia.

## 8) Como se calculan los pesos y cadencias

### 8.1 Pesos base por equipo
`get_pesos_mensuales()`:
- Toma la ultima fila de SUDOKU (`Estatus diario`, columnas P:U).
- Calcula `PESO_BASE = HORAS_equipo / HORAS_totales`.

`get_pesos_por_area()`:
- Areas A/B/C/T: normaliza por horas dentro del area.
- Area E: usa pesos fijos configurables (no depende de horas).

### 8.2 Cadencia preliminar
`compute_cadencia_preliminar()` calcula cadencia por equipo:
- Cadencia A: usando MST+ESP.
- Cadencia T: usando MBA+ESP.
- Formula: `cadencia = cupones_acumulados / (HORAS/6)`.

## 9) Etapa 1 - Asignacion por area

### 9.1 Area A/B/C (`distribuir_area_X`)
Metodo principal: MILP + ajuste fino por swaps.

Objetivos y restricciones:
1. Cada cupon fresh se asigna a un solo equipo.
2. Se fija exactamente cuantos fresh recibe cada equipo (`fresh_target_int`) para cuadrar cadencia global.
3. Se minimiza desviacion contra estimado acumulado por pilar (`hist + fresh`) con pesos:
- Web: muy alto
- Buscadores: alto
- P.Verticales: bajo
- Redes Sociales: casi libre (buffer)
4. Bandas blandas para Web/Buscadores con penalizacion alta.

Estimado por pilar:
- Para cada pilar `p`: `estimado(equipo,p) = total_acumulado_pilar(p) * share_equipo_area`.
- `share_equipo_area` proviene de `PESO_BASE` del area.
- Se redondea con Largest Remainder para que la suma por pilar sea exacta.

Ajuste fino posterior:
- Swaps 1 a 1 entre equipos para mejorar score acumulado sin romper conteos.

Subpilares (solo A/B/C):
- Se intenta ajustar subpilares:
  - Web -> `SEO`
  - Buscadores -> `GOOGLE BT`
- Prioridad: primero se respeta pilar global.
- Los swaps de subpilar son solo dentro del mismo pilar (target vs no-target).
- Invariante duro: si altera conteos por equipo+pilar, se revierte.
- Se imprime subreporte en terminal para SEO/Google BT.

Importante: en area T NO se ajustan subpilares.

### 9.2 Area T (`distribuir_area_T`)
Metodo: MILP con doble nivel de control (equipo + directora IG/XP).

Reglas principales:
1. Cadencia global exacta por equipo (fresh fijado por equipo).
2. Cuotas IG/XP variables segun peso real en area T:
- Share IG = suma pesos de A1/B1/C1 en T
- Share XP = 1 - Share IG
3. Cuota fresh IG/XP se fuerza exactamente (bloque robusto de ajuste entero).
4. Ademas del ajuste team-level, hay objetivo DV-level (IG/XP agregado) por pilar para:
- Web
- Buscadores
- Redes Sociales
5. Web/Buscadores tienen penalizacion mas alta; RS actua como compensador.

Si MILP no es optimo:
- Activa fallback proporcional pillar-aware.

### 9.3 Area E (`distribuir_area_E`)
Metodo: greedy pillar-aware con cuotas de directora y pesos fijos.

Reglas:
1. No usa cadencia.
2. Objetivo 1: split IG/XP del acumulado (`hist + fresh`) segun `E_SHARE_TARGET_IG`.
3. Objetivo 2: dentro de cada directora, acercarse a pesos por equipo de area E.
4. Reaperturas validas se respetan (no se reasignan).

#### Donde actualizar porcentajes de area E
En `distribucion_obs/extracted_functions.py` (parte superior):
- `E_TEAM_FIXED_WEIGHTS`: pesos fijos por equipo en area E.
- `E_SHARE_TARGET_IG`: objetivo de share IG (XP = 1 - IG).

## 10) Etapa 2 - Rebalanceo interno (sin romper invariantes)
Funcion principal: `run_segunda_etapa_v19()`.

Objetivo:
- Mejorar distribucion por `Agrupacion OBS` (pais) y `Programa de Interes` dentro de cada pilar.
- Sin cambiar:
  - cantidad fresh por equipo
  - cantidad hist+fresh por equipo x pilar
  - cadencia por equipo

### 10.1 Dedupe inicial de reaperturas
`dedup_reap_hist_vs_final()`:
- Si un REAP de `df_final_total` ya existe en historico (misma clave), se elimina del final clean para no duplicar conteo.

### 10.2 Core de optimizacion
`_segunda_etapa_core()`:
- Opera por area y por grupos de equipos.
- Agrupacion de equipos (`_groups_for_area`):
  - A/B: pares 1-2.
  - C: grupo completo por base (p.ej. C1,C2,C3 si existen).
  - T: grupo completo por base.
  - E: IG vs XP.
- En cada pilar aplica `_balanced_reassign_pillar()`:
  - slots exactos por equipo (no cambia volumen).
  - opcional `lock_cols` por area.
  - Para T, si existe columna `DV`, se bloquea por `DV` automaticamente.

Funcion de costo por cupon asignado:
- `coste = w_country * DeltaErrorPais + w_program * DeltaErrorPrograma`
- `DeltaError = |(actual+1)-esperado| - |actual-esperado|`
- Luego aplica 2-opt por swaps para mejora local adicional.

### 10.3 Blindajes finales
Postchecks duros:
- Fresh por equipo igual antes/despues.
- Hist+Fresh por equipo x pilar igual antes/despues.
- Cadencia igual antes/despues.

Si hay violacion, lanza error/assert.

### 10.4 Reinyeccion de REAP
`construir_final_con_reap()`:
- Reinyecta REAP originales no duplicados para el archivo exportable final.

## 11) Estructura de salida del pipeline
`PipelineResult` retorna:
- `input_rows`
- `df_pesos_actuales`
- `df_pesos_areas`
- `df_final_stage1`
- `df_final_stage2_clean`
- `df_final_export`
- `df_hist_total`
- `df_fresh`

`run_and_export()` guarda `df_final_export` en `output_path`.

## 12) Ejecucion desde cero

### 12.1 Requisitos
- Python 3.11+ recomendado.
- Archivos Excel en rutas configuradas.
- CBC solver disponible via `pulp` (normalmente viene con `PULP_CBC_CMD`).

### 12.2 Instalacion rapida
```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

### 12.3 Ejecutar
```bash
python run.py
```

Salida esperada en terminal:
- resumen de PMAX
- pesos por area
- comparativas por area
- controles historico antes/despues
- `OK - control filas entrada/salida: X -> X`

### 12.4 Ejecucion con rutas custom
Puede crearse `PipelineConfig` manual en un script corto:
```python
from pathlib import Path
from distribucion_obs import PipelineConfig, run_and_export

workspace = Path(r"C:\ruta\proyecto")
cfg = PipelineConfig.default(workspace)
cfg = PipelineConfig(
    **{**cfg.__dict__, "downloads_dir": Path(r"C:\ruta\mis_archivos")}
)
run_and_export(cfg)
```

## 13) Troubleshooting

1. `No se encontro archivo con prefijo ...`
- Verificar que el archivo mas reciente este en `downloads_dir` con prefijo correcto.

2. `No se pudo leer hora de corte ... O22`
- Revisar `SUDOKU.xlsx` hoja `Estatus diario`, celda O22.

3. `Control de filas fallo: entrada != salida`
- Revisar deduplicacion por ID y columna `INDEX_ORIGINAL`.

4. MILP area T infeasible
- El codigo ya hace fallback.
- Si se quiere evitar fallback, revisar bandas, cuotas IG/XP y pesos.

5. Descuadres en area E
- Revisar `E_TEAM_FIXED_WEIGHTS` y `E_SHARE_TARGET_IG`.
- Confirmar que equipos E incluyan IG y XP.

## 14) Parametros de negocio que mas se tocan
- `E_TEAM_FIXED_WEIGHTS` (area E).
- `E_SHARE_TARGET_IG` (area E).
- `pilar_band_web`, `pilar_band_busc` (config).
- `stage2_w_country`, `stage2_w_program` (config).
- `time_limit` de solver (config).

## 15) Notas sobre archivos legacy
- `Distribucion_OBS_legacy.ipynb`: version historica/no modular.
- `Manual_Codigo_Distribucion_OBS.docx`: documentacion previa.
- `build/` y `dist/`: artefactos de ejecutable.

El flujo vigente y mantenible para desarrollo actual es el modular (`run.py` + paquete `distribucion_obs`).
