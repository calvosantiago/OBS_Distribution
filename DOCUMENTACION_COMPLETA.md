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
- `extraccion_atenea.py`: conexion con Atenea/Dataverse, extraccion de `qb_cn` y oportunidades abiertas no asignadas.
- `probar_ownerid.py`: prueba/asignacion automatica de `ownerid` en Atenea desde `Distribucion_Final.xlsx`.
- `distribucion_obs/config.py`: configuracion de rutas y parametros.
- `distribucion_obs/loaders.py`: lectura de archivos fuente Excel o DataFrames extraidos desde Atenea.
- `distribucion_obs/preprocess.py`: limpieza, enriquecimiento, separacion fresh/reap/historico y calculos preliminares.
- `distribucion_obs/stages.py`: orquestacion de etapa 1 y etapa 2.
- `distribucion_obs/pipeline.py`: pipeline end-to-end y controles globales.
- `distribucion_obs/extracted_functions.py`: motor principal de distribucion (A/B/C/T/E) y segunda etapa.
- `launcher_atenea.py`: punto de entrada del ejecutable de Atenea.
- `build_atenea.bat`: recompila el ejecutable `Distribucion_OBS_Atenea.exe`.

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

### Entradas principales - flujo Excel legacy
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

### Entradas principales - flujo Atenea
El flujo Atenea sustituye los dos Excels de Downloads por DataFrames extraidos directamente desde Dataverse:

1. `qb_cn`:
- Sale de `extraccion_atenea.py`.
- Equivale al historico `qb_CN_V3_OBS...` que esperaba el codigo original.

2. `op_no_asig`:
- Sale de `extraccion_atenea.py`.
- Equivale al Excel de oportunidades abiertas no asignadas que esperaba el codigo original.

3. `Areas_Paises.xlsx`, `SUDOKU.xlsx`, `OBS_ESTRUCTURA_COMERCIAL.xlsx` y `OBS_PULL_PUSH.xlsx`:
- Se siguen usando igual que en el flujo Excel.
- `Areas_Paises.xlsx` sigue normalizando programas y paises.

4. `token_cache.bin`:
- Cache local de autenticacion MSAL para Atenea/Dataverse.
- No contiene el codigo del proyecto; contiene tokens de usuario.
- Si otro usuario ejecuta el programa, lo correcto es que tenga su propia cache/token de autenticacion, no compartir una cache personal.

### Salida
- `Distribucion_Final.xlsx` en la raiz del proyecto.
- En flujo Excel legacy: una hoja con el resultado final.
- En flujo Atenea: tres hojas:
  - `Distribucion_Final`
  - `qb_cn`
  - `op_no_asig`
- Las hojas del flujo Atenea salen ordenadas por `Fecha de creación`.
- `INDEX_ORIGINAL` se conserva en `Distribucion_Final` para trazabilidad: puede cambiar el orden visual del Excel, pero no se pierde el indice original del pipeline.

### Export auxiliar de leads
Cuando se ejecuta el flujo Atenea, tambien se guarda una copia de la extraccion en:

```text
leads\leads_YYYY-MM-DD_YYYY-MM-DD.xlsx
```

Ese archivo contiene las hojas `qb_cn` y `op_no_asig` con el formato y orden de columnas esperado por los CSV/Excels historicos.

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
- `input_source`: `excel` o `atenea`.
- `atenea_fecha_inicio`, `atenea_fecha_fin`: periodo de extraccion desde Atenea.
- `atenea_cache_file`: ruta de `token_cache.bin`.
- `atenea_export_excel`: si `True`, exporta copia de leads en carpeta `leads`.
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
- `atenea_qbcn_export`
- `atenea_op_no_asig_export`

`run_and_export()` guarda `df_final_export` en `output_path`.

Si el origen es Atenea y existen `atenea_qbcn_export` / `atenea_op_no_asig_export`, `run_and_export()` escribe `Distribucion_Final.xlsx` con 3 hojas:
- `Distribucion_Final`
- `qb_cn`
- `op_no_asig`

La hoja `Distribucion_Final` se ordena por `Fecha de creación` solo para el Excel. La columna `INDEX_ORIGINAL` se mantiene para poder reconstruir el orden interno del pipeline y validar que no se perdieron filas.

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
#### Flujo Excel legacy
```bash
python run.py
```

Salida esperada en terminal:
- resumen de PMAX
- pesos por area
- comparativas por area
- controles historico antes/despues
- `OK - control filas entrada/salida: X -> X`

#### Flujo Atenea
```bash
python run.py --source atenea --fecha-inicio 2026-06-01 --fecha-fin 2026-06-05
```

Notas:
- `--fecha-inicio` y `--fecha-fin` se informan en formato `YYYY-MM-DD`.
- El script pide la hora de corte por pantalla.
- Se extraen datos de Atenea/Dataverse usando MSAL.
- Se genera `Distribucion_Final.xlsx` con 3 hojas.
- Se genera copia de leads en `leads\leads_YYYY-MM-DD_YYYY-MM-DD.xlsx`.

Orden recomendado de prueba:
1. Ejecutar distribucion Atenea.
2. Revisar `Distribucion_Final.xlsx`.
3. Ejecutar dry-run de asignacion automatica.
4. Solo si el dry-run es correcto, aplicar ownerid.

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

## 13) Asignacion automatica de ownerid en Atenea

Archivo: `probar_ownerid.py`

Este script usa la API de Dataverse, que es lo que hay detras de Atenea, para actualizar el campo `ownerid` de oportunidades.

La distribucion calcula el destino en la columna `EQUIPO_FINAL`. Ese valor se resuelve contra Atenea como:
1. `team.name`
2. si no existe como equipo, `systemuser.fullname`

Si un nombre coincide con varios equipos o usuarios, el script falla. No elige a ciegas.

### 13.1 Dry-run de una oportunidad
No cambia Atenea:

```bash
python probar_ownerid.py --id 2021-002714891
```

### 13.2 Aplicar una oportunidad
Actualiza Atenea solo para esa oportunidad:

```bash
python probar_ownerid.py --id 2021-002714891 --apply
```

### 13.3 Probar destino manual
Usa `--owner` solo con una oportunidad individual:

```bash
python probar_ownerid.py --id 2021-002714891 --owner Equipo_A1
python probar_ownerid.py --id 2021-002714891 --owner Equipo_A1 --apply
```

### 13.4 Dry-run masivo
No cambia Atenea:

```bash
python probar_ownerid.py --all
```

Para probar pocas filas:

```bash
python probar_ownerid.py --all --limit 10
```

Si Dataverse responde con limite `429`, el script reintenta automaticamente. Parametros utiles:

```bash
python probar_ownerid.py --all --sleep 1 --max-retries 8 --retry-wait 45
```

- `--sleep`: pausa entre oportunidades.
- `--max-retries`: numero de reintentos por peticion HTTP.
- `--retry-wait`: espera base si Dataverse no informa `Retry-After`.

### 13.5 Aplicar masivo
Actualiza Atenea:

```bash
python probar_ownerid.py --all --apply
```

Para aplicar con mas margen frente a limites de Dataverse:

```bash
python probar_ownerid.py --all --apply --sleep 1 --max-retries 8 --retry-wait 45
```

### 13.6 Controles de seguridad del masivo
El masivo:
- Lee `Distribucion_Final.xlsx`.
- Busca cada oportunidad por `opportunityid` o `ID de la Oportunidad`.
- Resuelve `EQUIPO_FINAL` como equipo o usuario en Atenea.
- Si el owner actual ya coincide con `EQUIPO_FINAL`, lo salta (`SKIPPED_ALREADY_OWNER`).
- Si el owner actual en Atenea no coincide con el owner original guardado en el Excel (`_ownerid_value`, `owneridname` o `Propietario`), lo marca como `SKIPPED_OWNER_CHANGED`.
- En modo `--apply`, si hay errores o owners cambiados, no aplica ningun cambio masivo.
- Genera log Excel en `crm_ownerid_logs`.
- El log incluye `http_retries`, con los reintentos usados en cada fila.

Esto evita aplicar cambios sobre oportunidades que alguien haya reasignado manualmente despues de generar `Distribucion_Final.xlsx`.

### 13.7 Estados del log
- `DRY_RUN_CHANGE`: se cambiaria si se ejecutara con `--apply`.
- `READY_TO_APPLY`: fila validada antes de aplicar.
- `APPLIED`: cambio aplicado correctamente.
- `SKIPPED_ALREADY_OWNER`: ya estaba en el owner correcto.
- `SKIPPED_OWNER_CHANGED`: el owner actual cambio respecto al Excel.
- `ERROR`: error de validacion.
- `ERROR_APPLY`: fallo durante el PATCH real.

### 13.8 Riesgos principales
- Nombres de `EQUIPO_FINAL` que no existan exactamente en Atenea.
- Usuarios/equipos duplicados por nombre.
- Cambios manuales entre la generacion del Excel y la aplicacion.
- Rendimiento: 600-700 oportunidades implican varias llamadas API por oportunidad; el dry-run puede tardar.
- Autenticacion: cada usuario debe autenticarse con sus permisos. Compartir `token_cache.bin` no es recomendable.

## 14) Troubleshooting

1. `No se encontro archivo con prefijo ...`
- Verificar que el archivo mas reciente este en `downloads_dir` con prefijo correcto.

2. `No se pudo leer hora de corte ... O22`
- Solo aplica al ejecutar via `run.py`. En el ejecutable `.exe`, la hora de corte se introduce por teclado al inicio (ver seccion 17.6).

3. `Control de filas fallo: entrada != salida`
- Revisar deduplicacion por ID y columna `INDEX_ORIGINAL`.

4. MILP area T infeasible
- El codigo ya hace fallback.
- Si se quiere evitar fallback, revisar bandas, cuotas IG/XP y pesos.

5. Descuadres en area E
- Revisar `E_TEAM_FIXED_WEIGHTS` y `E_SHARE_TARGET_IG`.
- Confirmar que equipos E incluyan IG y XP.

6. Atenea pide login o falla autenticacion
- Revisar `token_cache.bin`.
- Si es otro usuario, debe autenticarse con su propia cuenta y permisos.

7. `No se encontro owner destino como team ni systemuser`
- Revisar que `EQUIPO_FINAL` exista exactamente en Atenea.
- Ejemplo: `Equipo_A1` no es lo mismo que `Equipo A1`.

8. `SKIPPED_OWNER_CHANGED`
- La oportunidad ya no esta en el mismo owner que cuando se genero el Excel.
- Revisar manualmente antes de forzar.

9. `Distribucion_Final.xlsx` no tiene 3 hojas
- Solo tendra 3 hojas cuando se ejecute con `--source atenea`.
- El flujo Excel legacy mantiene una sola hoja.

## 15) Parametros de negocio que mas se tocan
- `E_TEAM_FIXED_WEIGHTS` (area E).
- `E_SHARE_TARGET_IG` (area E).
- `pilar_band_web`, `pilar_band_busc` (config).
- `stage2_w_country`, `stage2_w_program` (config).
- `time_limit` de solver (config).

## 16) Notas sobre archivos legacy
- `Distribucion_OBS_legacy.ipynb`: version historica/no modular.
- `Manual_Codigo_Distribucion_OBS.docx`: documentacion previa.
- `build/` y `dist/`: artefactos de PyInstaller (ver seccion 17).
- `diagnostico_atenea.py`: herramienta de diagnostico usada durante desarrollo. Eliminada porque el flujo Atenea ya quedo integrado y validado.

El flujo vigente y mantenible para desarrollo actual es el modular (`run.py` + paquete `distribucion_obs`).

## 17) Ejecutable distribuible (.exe)

### 17.1 Descripcion general
Ademas del flujo de desarrollo (`run.py`), el proyecto genera un ejecutable `Distribucion_OBS.exe` autocontenido mediante PyInstaller. Cualquier usuario con acceso a la carpeta OneDrive puede ejecutarlo sin tener Python instalado.

Tambien existe un ejecutable especifico para el flujo Atenea:
- `Distribucion_OBS_Atenea.exe`
- Se compila con `build_atenea.bat`.
- Usa `launcher_atenea.py`.
- Pide hora de corte, fecha inicio y fecha fin.

### 17.2 Archivos involucrados

| Archivo | Rol |
|---|---|
| `launcher.py` | Punto de entrada del ejecutable (no es `run.py`) |
| `build.bat` | Script para compilar/recompilar el ejecutable con doble click |
| `dist/Distribucion_OBS.exe` | Ejecutable generado por PyInstaller |
| `Distribucion_OBS.exe` | Copia del exe en la raiz del proyecto (esta es la version que se usa) |
| `launcher_atenea.py` | Punto de entrada del ejecutable Atenea |
| `build_atenea.bat` | Script para compilar/recompilar `Distribucion_OBS_Atenea.exe` |
| `dist/Distribucion_OBS_Atenea.exe` | Ejecutable Atenea generado por PyInstaller |

### 17.3 Como recompilar el ejecutable legacy
Cada vez que se modifique el codigo fuente, ejecutar `build.bat` con doble click. El proceso:
1. Llama a PyInstaller con Python 3.14
2. Genera `dist\Distribucion_OBS.exe`
3. El exe debe copiarse manualmente a la raiz del proyecto para que encuentre `Areas_Paises.xlsx` y `SUDOKU.xlsx`

La carpeta `build/` que genera PyInstaller es temporal y puede borrarse sin problema.

### 17.4 Como recompilar el ejecutable Atenea
Ejecutar:

```bat
build_atenea.bat
```

Genera:

```text
dist\Distribucion_OBS_Atenea.exe
```

Este ejecutable sigue el flujo Atenea, no el flujo Excel legacy.

### 17.5 Resolucion de rutas en modo frozen
En PyInstaller `--onefile`, `__file__` apunta a una carpeta temporal de extraccion (`_MEIxxx`), no a la ubicacion real del exe. `launcher.py` resuelve esto con:

```python
if getattr(sys, "frozen", False):
    workspace = Path(sys.executable).resolve().parent  # carpeta del .exe
else:
    workspace = Path(__file__).resolve().parent         # modo desarrollo
```

Por eso el exe debe estar en la misma carpeta que `Areas_Paises.xlsx` y `SUDOKU.xlsx` (la raiz del proyecto en OneDrive).

### 17.6 Hora de corte interactiva
Al ejecutar el `.exe`, el programa pide la hora de corte por terminal antes de iniciar el pipeline:

```
==================================================
  HORA DE CORTE
  Oportunidades creadas ANTES de esta fecha/hora
  se trataran como REAP (no redistribuidas).
==================================================
  Introduce la hora de corte (DD/MM/AAAA HH:MM): 12/03/2026 11:20
  -> Corte establecido: 12/03/2026 11:20
```

Formato requerido: `DD/MM/AAAA HH:MM`. Si el formato es incorrecto, vuelve a pedirlo.

Esto sustituye la lectura de la celda `O22` de `SUDOKU.xlsx`. La funcion `split_reap_fresh_hist()` acepta `cutoff_dt` como parametro opcional: si se pasa, lo usa directamente; si es `None`, lee del SUDOKU (comportamiento legacy de `run.py`).

En `Distribucion_OBS_Atenea.exe`, ademas de hora de corte, se piden:
- Hora de corte (`YYYY-MM-DD HH:MM`)
- Fecha inicio qb_cn/op_no_asig (`YYYY-MM-DD`)
- Fecha fin qb_cn/op_no_asig (`YYYY-MM-DD`)

Ejemplo:

```text
Hora de corte: 2026-06-05 10:00
Fecha inicio: 2026-06-01
Fecha fin:    2026-06-05
```

### 17.7 Paquetes con datos o binarios incluidos explicitamente
PyInstaller no detecta automaticamente todos los archivos de datos de terceros. El `build.bat` incluye:

| Flag | Paquete | Motivo |
|---|---|---|
| `--collect-data country_converter` | country_converter | Archivo `country_data.tsv` |
| `--collect-all pulp` | pulp | Binario `cbc.exe` del solver CBC |
| `--collect-data openpyxl` | openpyxl | Templates de archivos Excel |

### 17.8 Distribucion a otros usuarios
Los usuarios solo necesitan:
1. `Distribucion_OBS.exe` en la carpeta raiz del proyecto (OneDrive compartida)
2. `Areas_Paises.xlsx` y `SUDOKU.xlsx` en la misma carpeta (ya estan en OneDrive)
3. Los archivos `OBS_ESTRUCTURA_COMERCIAL.xlsx` y `OBS_PULL_PUSH.xlsx` sincronizados en su OneDrive
4. Los archivos de cupones e historico en su carpeta `Downloads` con los prefijos correctos

Para el flujo Atenea:
1. Usar `Distribucion_OBS_Atenea.exe`.
2. Tener permisos en Atenea/Dataverse.
3. Autenticarse con su usuario cuando el programa lo pida.
4. No compartir `token_cache.bin` entre usuarios.

