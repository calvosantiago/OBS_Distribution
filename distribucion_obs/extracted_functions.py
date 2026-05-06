"""Funciones extraídas de Distribucion_OBS.py sin ejecución top-level."""

from __future__ import annotations
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import warnings
import pulp
from collections import defaultdict
import re
from typing import Dict, List, Sequence, Tuple, Optional, Set

PILLARS = ['Web', 'Buscadores', 'P.Verticales', 'Redes Sociales']
IG_teams = {'Equipo_A1', 'Equipo_B1', 'Equipo_C1'}
XP_teams = {'Equipo_A2', 'Equipo_B2', 'Equipo_C2'}

# Área E: configuración editable de negocio.
E_TEAM_FIXED_WEIGHTS = {
    'Equipo_A1': 0.00,
    'Equipo_B1': 0.00,
    'Equipo_C1': 0.00,
    'Equipo_A2': 0.00,
    'Equipo_B2': 1.00,
    'Equipo_C2': 0.00,
}
E_SHARE_TARGET_IG = 0.00

def obtener_semana_comercial(fecha_actual: datetime, calendario_path=None) -> str:
    """Devuelve la semana comercial OBS en formato AÑO-MES-Sn.
    Consulta OBS_CALENDARIO_COMERCIAL.xlsx (hoja CAL_COM, columna 'SEM NAT').
    Si no se puede leer el fichero, usa la fórmula de fallback."""
    from pathlib import Path

    if calendario_path is not None:
        try:
            df_cal = pd.read_excel(calendario_path, sheet_name="CAL_COM", usecols=["FECHA", "ams"])
            df_cal["FECHA"] = pd.to_datetime(df_cal["FECHA"]).dt.date
            fecha_date = fecha_actual.date() if hasattr(fecha_actual, "date") else fecha_actual
            fila = df_cal[df_cal["FECHA"] == fecha_date]
            if not fila.empty:
                # "ams" tiene formato "2026-10S1" → añadir guión antes de S → "2026-10-S1"
                raw = str(fila["ams"].iloc[0]).strip()
                semana = raw[:-2] + "-" + raw[-2:] if raw[-2] == "S" else raw
                print(f"[OK] Semana comercial actual: {semana}")
                return semana
            else:
                print(
                    f"[ADVERTENCIA] La fecha {fecha_date} no se encontro en el calendario comercial "
                    f"({calendario_path.name}). Se usara calculo por formula."
                )
        except Exception as e:
            print(
                f"[ADVERTENCIA] No se pudo leer el calendario comercial "
                f"({calendario_path.name}): {e}. Se usara calculo por formula."
            )

    # --- Fallback por fórmula ---
    año = fecha_actual.year
    mes = fecha_actual.month
    primer_dia_mes = datetime(año, mes, 1)
    anchor = primer_dia_mes - timedelta(days=2)
    primer_martes = anchor + timedelta(days=(1 - anchor.weekday() + 7) % 7)
    if fecha_actual < primer_martes:
        ultimo_dia_mes_ant = primer_dia_mes - timedelta(days=1)
        año, mes = ultimo_dia_mes_ant.year, ultimo_dia_mes_ant.month
        primer_dia_mes = datetime(año, mes, 1)
        anchor = primer_dia_mes - timedelta(days=2)
        primer_martes = anchor + timedelta(days=(1 - anchor.weekday() + 7) % 7)
    numero_semana = (fecha_actual - primer_martes).days // 7 + 1
    return f"{año}-{mes:02d}-S{numero_semana}"
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
def get_pesos_por_area(df_pesos_actuales: pd.DataFrame) -> dict:
    equipos_area = {
        'A': ['Equipo_A1', 'Equipo_A2'],
        'B': ['Equipo_B1', 'Equipo_B2'],
        'C': ['Equipo_C1', 'Equipo_C2'],
        'T': df_pesos_actuales['EQUIPO'].tolist(),  # Todos los equipos
        'E': E_TEAM_FIXED_WEIGHTS.copy(),  # Pesos fijos para programas en inglés
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
def _norm_email(x):
    s = str(x).strip().lower()
    return s if s and s not in {"nan", "none"} else pd.NA
def _norm_phone(x):
    # deja SOLO dígitos; si hay >=9, nos quedamos con los últimos 9 (típico ES).
    s = re.sub(r"\D+", "", str(x))
    if len(s) >= 9:
        return s[-9:]
    return pd.NA
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

    # Trabajar en numpy evita ambigüedad entre indexado por etiqueta/posición de pandas
    s_scaled_np = s_scaled.to_numpy(dtype=float)
    floors = np.floor(s_scaled_np).astype(int)
    residual = T - int(floors.sum())

    # Partes fraccionales
    frac = s_scaled_np - floors
    order_asc = np.argsort(frac)          # menor → mayor
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

        # Mismos pesos que el MILP: RS es válvula (coste ~0), Web/Busc son prioritarios
        PENAL_AJUSTE = {'Web': 50, 'Buscadores': 30, 'P.Verticales': 0.5, 'Redes Sociales': 0.01}

        def score_delta(d):
            """Puntuación ponderada del delta (mismo criterio que el MILP)."""
            return sum(PENAL_AJUSTE.get(p, 1) * float(d[p].abs().sum()) for p in d.columns)

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
        score = score_delta(delta)

        def _pick_subpillar_col(dframe: pd.DataFrame) -> str | None:
            candidates = [
                "SubPillar (Campaña de origen) (Campaña)",
                "SubPillar Name (Campaña de origen) (Campaña)",
            ]
            for c in candidates:
                if c in dframe.columns:
                    return c
            return None

        _sub_col_fresh = _pick_subpillar_col(df)
        _sn_fresh = (
            df[_sub_col_fresh].astype(str).str.upper().str.strip()
            if _sub_col_fresh is not None
            else pd.Series("", index=df.index)
        )

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

                # Para Web: mover primero el sub-pilar (SEO/no-SEO) que más necesita el receptor.
                # Así el swap global y el ajuste de sub-pilar trabajan en la misma dirección.
                if pilar == 'Web' and _sub_col_fresh is not None:
                    _recv_web_idx = df.index[(df['EQUIPO_FINAL'] == falta_eq) & (df['PILAR_NORM'] == 'Web')]
                    _donor_seo_n = _sn_fresh.loc[cupones_exceso.index].str.contains('SEO', na=False).sum()
                    _recv_seo_n  = _sn_fresh.loc[_recv_web_idx].str.contains('SEO', na=False).sum() if len(_recv_web_idx) > 0 else 0
                    _donor_web_n = max(1, len(cupones_exceso))
                    _recv_web_n  = max(1, len(_recv_web_idx))
                    _is_seo = _sn_fresh.loc[cupones_exceso.index].str.contains('SEO', na=False)
                    if (_donor_seo_n / _donor_web_n) > (_recv_seo_n / _recv_web_n):
                        # Donante tiene más SEO proporcionalmente → mover SEO primero al receptor
                        seo_idx    = _is_seo[_is_seo].index.tolist()
                        no_seo_idx = _is_seo[~_is_seo].index.tolist()
                        cupones_exceso = cupones_exceso.loc[seo_idx + no_seo_idx]
                    else:
                        # Receptor ya tiene suficiente SEO → mover no-SEO primero
                        seo_idx    = _is_seo[_is_seo].index.tolist()
                        no_seo_idx = _is_seo[~_is_seo].index.tolist()
                        cupones_exceso = cupones_exceso.loc[no_seo_idx + seo_idx]

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

                            new_score = score_delta(new_delta)

                            if new_score < score:
                                delta = new_delta
                                score = new_score
                                movimientos += 1
                                realizado = True
                                break
                            else:
                                df.at[i,'EQUIPO_FINAL'], df.at[j,'EQUIPO_FINAL'] = ei, ej
                    if realizado:
                        break
                if not realizado:
                    break

        # ---------- Ajuste fino de subpilares (sin tocar global por pilar) ----------
        # Regla: solo swaps dentro del mismo pilar (target-subpilar <-> no-target-subpilar).
        # Así se conservan exactamente los conteos por equipo y por pilar.
        # (_pick_subpillar_col definida antes del loop global)
        sub_col_fresh = _pick_subpillar_col(df)
        sub_col_hist = _pick_subpillar_col(df_hist_area)

        movimientos_sub = 0
        # Subpilares solo aplican en A/B/C y nunca deben alterar los conteos por pilar/equipo.
        if (sigla_area in {"A", "B", "C"}) and (sub_col_fresh is not None):
            df_before_sub = df.copy(deep=True)
            fresh_sub_norm = df[sub_col_fresh].astype(str).str.upper().str.strip()
            hist_sub_norm = (
                df_hist_area[sub_col_hist].astype(str).str.upper().str.strip()
                if sub_col_hist is not None
                else pd.Series("", index=df_hist_area.index)
            )

            pesos = (
                df_pesos_areas[df_pesos_areas['AREA'] == sigla_area]
                .set_index('EQUIPO')['PESO_BASE']
                .reindex(equipos)
                .astype(float)
                .fillna(0.0)
            )
            total_pesos = float(pesos.sum())
            shares = (
                pesos / total_pesos
                if total_pesos > 0
                else pd.Series(1.0 / len(equipos), index=equipos)
            )

            # target_subpilar contiene texto buscado (case-insensitive) dentro del subpilar.
            sub_targets = [
                ("Web", "SEO"),
                ("Buscadores", "GOOGLE BT"),
            ]

            for pilar_obj, token_sub in sub_targets:
                mask_pilar_fresh = df["PILAR_NORM"].eq(pilar_obj)
                mask_target_fresh = mask_pilar_fresh & fresh_sub_norm.str.contains(token_sub, case=False, regex=False, na=False)
                mask_other_fresh = mask_pilar_fresh & (~mask_target_fresh)

                # Objetivo acumulado del subpilar por equipo (hist + fresh), repartido por pesos de área.
                mask_pilar_hist = df_hist_area["PILAR_NORM"].eq(pilar_obj)
                hist_target_counts = df_hist_area[mask_pilar_hist].copy()
                if not hist_target_counts.empty:
                    hist_target_counts = hist_target_counts[
                        hist_sub_norm.loc[hist_target_counts.index].str.contains(token_sub, case=False, regex=False, na=False)
                    ]
                hist_target_counts = (
                    hist_target_counts.groupby("EQUIPO_FINAL").size()
                    .reindex(equipos, fill_value=0)
                    .astype(int)
                )
                fresh_target_counts = (
                    df[mask_target_fresh].groupby("EQUIPO_FINAL").size()
                    .reindex(equipos, fill_value=0)
                    .astype(int)
                )
                actual_target_counts = (hist_target_counts + fresh_target_counts).astype(int)
                total_target = int(actual_target_counts.sum())
                if total_target <= 0:
                    continue

                target_float = shares * float(total_target)
                target_int = ajustar_redondeo_sum_exacta(target_float, total=total_target).reindex(equipos, fill_value=0).astype(int)
                delta_sub = (actual_target_counts - target_int).astype(int)

                # === DIAGNÓSTICO TEMPORAL ===
                print(f"\n📊 DIAG subpilar [{pilar_obj}/{token_sub}] Área {sigla_area}")
                print(f"  Total {token_sub} (hist+fresh): {total_target}")
                print(f"  Shares:\n{shares.to_string()}")
                print(f"  Target por equipo:\n{target_int.to_string()}")
                print(f"  Actual (hist+fresh):\n{actual_target_counts.to_string()}")
                print(f"  Delta (actual-target):\n{delta_sub.to_string()}")
                for _eq in equipos:
                    _n_seo   = int(len(df.index[(df['EQUIPO_FINAL'] == _eq) & mask_target_fresh]))
                    _n_other = int(len(df.index[(df['EQUIPO_FINAL'] == _eq) & mask_other_fresh]))
                    print(f"  {_eq}: fresh_{token_sub}={_n_seo}  fresh_no_{token_sub}={_n_other}")
                # === FIN DIAGNÓSTICO ===

                # Mueve 1 unidad de subpilar target de donante->receptor y compensa con otro cupón del mismo pilar.
                while True:
                    donors = [e for e in equipos if int(delta_sub.get(e, 0)) > 0]
                    receivers = [e for e in equipos if int(delta_sub.get(e, 0)) < 0]
                    if not donors or not receivers:
                        break

                    donors = sorted(donors, key=lambda e: int(delta_sub[e]), reverse=True)
                    receivers = sorted(receivers, key=lambda e: int(delta_sub[e]))
                    moved = False

                    for d in donors:
                        idx_d = df.index[(df["EQUIPO_FINAL"] == d) & mask_target_fresh]
                        if len(idx_d) == 0:
                            continue
                        for r in receivers:
                            idx_r = df.index[(df["EQUIPO_FINAL"] == r) & mask_other_fresh]
                            if len(idx_r) == 0:
                                continue
                            i = idx_d[0]
                            j = idx_r[0]
                            # Swap dentro del mismo pilar: mantiene globales por equipo/pilar.
                            df.at[i, "EQUIPO_FINAL"] = r
                            df.at[j, "EQUIPO_FINAL"] = d
                            delta_sub.at[d] = int(delta_sub.at[d]) - 1
                            delta_sub.at[r] = int(delta_sub.at[r]) + 1
                            movimientos_sub += 1
                            moved = True
                            break
                        if moved:
                            break

                    if not moved:
                        break

            # Invariante duro: los sub-swaps no pueden tocar totales por equipo+pilar.
            before_counts = (
                df_before_sub.groupby(["EQUIPO_FINAL", "PILAR_NORM"]).size()
                .reindex(
                    pd.MultiIndex.from_product([equipos, pillars], names=["EQUIPO_FINAL", "PILAR_NORM"]),
                    fill_value=0,
                )
                .astype(int)
            )
            after_counts = (
                df.groupby(["EQUIPO_FINAL", "PILAR_NORM"]).size()
                .reindex(
                    pd.MultiIndex.from_product([equipos, pillars], names=["EQUIPO_FINAL", "PILAR_NORM"]),
                    fill_value=0,
                )
                .astype(int)
            )
            if not before_counts.equals(after_counts):
                df = df_before_sub
                movimientos_sub = 0
                print("⚠️ Ajuste subpilares revertido: alteraba conteos por pilar/equipo.")

        if movimientos_sub > 0:
            print(f"🔧 Ajuste subpilares (SEO/Google BT) completado. Movimientos: {movimientos_sub}")
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
    # Los REAP de esta área que aparecen en el output final pueden tener PILAR_NORM / IDIOMA / TIPO
    # que no coinciden con los filtros de df_hist_area (p.ej. pilar sin mapear, tipo diferente).
    # Si esos REAP no se cuentan en cupones_hist, el equipo que los recibe obtiene demasiados FRESH
    # porque el fresh_target no descuenta sus reaperturas → ratio final desviado del objetivo.
    # Solución: añadir a cupones_hist los REAP del output que NO pasaron el filtro de df_hist_area.
    _df_reap_area_out = df_reap_validas[
        (df_reap_validas['EQUIPO_FINAL'].isin(equipos_X)) &
        (~df_reap_validas['AREA'].isin(['T', 'E']))
    ].copy()
    _reap_in_hist_mask = (
        _df_reap_area_out['PILAR_NORM'].isin(pillars) &
        _df_reap_area_out['IDIOMA'].isin(filt_idiomas) &
        _df_reap_area_out['TIPO'].isin(filt_tipos)
    )
    _cupones_reap_extra = (
        _df_reap_area_out[~_reap_in_hist_mask]
        .groupby('EQUIPO_FINAL').size()
        .reindex(equipos_X, fill_value=0).astype(int)
    )
    cupones_hist = cupones_hist + _cupones_reap_extra

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
    # Prioridad explícita:
    # 1) WEB (más importante), 2) Global por equipo (ya va exacto por restricción),
    # 3) Buscadores, 4) P.Verticales, 5) Redes Sociales (válvula de escape).
    PENAL_DIFF = {'Web': 120, 'Buscadores': 45, 'P.Verticales': 15, 'Redes Sociales': 0.01}
    BIG_WEB  = 9000.0
    BIG_BUSC = 4500.0

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

    # ---------- Subreporte SEO / Google BT (terminal) ----------
    def _pick_subpillar_col_report(dframe: pd.DataFrame) -> str | None:
        candidates = [
            "SubPillar (Campaña de origen) (Campaña)",
            "SubPillar Name (Campaña de origen) (Campaña)",
        ]
        for c in candidates:
            if c in dframe.columns:
                return c
        return None

    sub_col_hist = _pick_subpillar_col_report(df_hist_area)
    sub_col_fresh = _pick_subpillar_col_report(df_fresh_area)
    if (sub_col_hist is not None) and (sub_col_fresh is not None):
        equipos = list(map(str, equipos_X))
        pesos_sub = (
            df_pesos_areas[df_pesos_areas["AREA"] == sigla_area]
            .set_index("EQUIPO")["PESO_BASE"]
            .reindex(equipos)
            .astype(float)
            .fillna(0.0)
        )
        sum_p = float(pesos_sub.sum())
        shares_sub = (
            pesos_sub / sum_p
            if sum_p > 0
            else pd.Series(1.0 / len(equipos), index=equipos)
        )

        sub_targets_report = [
            ("WEB_SEO", "Web", "SEO"),
            ("BUSC_GBT", "Buscadores", "GOOGLE BT"),
        ]
        rep = pd.DataFrame(index=equipos)

        hist_sub_norm = df_hist_area[sub_col_hist].astype(str).str.upper().str.strip()
        fresh_sub_norm = df_fresh_area[sub_col_fresh].astype(str).str.upper().str.strip()

        for label, pilar_obj, token_sub in sub_targets_report:
            hist_mask = (
                df_hist_area["PILAR_NORM"].eq(pilar_obj)
                & hist_sub_norm.str.contains(token_sub, case=False, regex=False, na=False)
            )
            fresh_mask = (
                df_fresh_area["PILAR_NORM"].eq(pilar_obj)
                & fresh_sub_norm.str.contains(token_sub, case=False, regex=False, na=False)
            )

            hist_counts = (
                df_hist_area.loc[hist_mask].groupby("EQUIPO_FINAL").size()
                .reindex(equipos, fill_value=0)
                .astype(int)
            )
            fresh_counts = (
                df_fresh_area.loc[fresh_mask].groupby("EQUIPO_FINAL").size()
                .reindex(equipos, fill_value=0)
                .astype(int)
            )
            acum_counts = (hist_counts + fresh_counts).astype(int)
            total_sub = int(acum_counts.sum())
            obj_counts = ajustar_redondeo_sum_exacta(shares_sub * float(total_sub), total=total_sub).reindex(equipos, fill_value=0).astype(int)

            rep[f"{label}_HIST"] = hist_counts
            rep[f"{label}_FRESH"] = fresh_counts
            rep[f"{label}_ACUM"] = acum_counts
            rep[f"{label}_OBJ"] = obj_counts
            rep[f"{label}_DELTA"] = (acum_counts - obj_counts).astype(int)

        total_row = pd.DataFrame(rep.sum(axis=0)).T
        total_row.index = ["TOTAL_AREA"]
        rep_print = pd.concat([rep, total_row], axis=0)
        print(f"\n📎 Subreporte Subpilares Área {sigla_area} (SEO / Google BT)")
        print(rep_print.to_string())
    else:
        print(f"\n📎 Subreporte Subpilares Área {sigla_area}: sin columna de subpilar en datos.")

    return df_final_area, df_cadencia, df_comparativa, cad_pilar
def distribuir_area_T(
    df_fresh, df_hist_total, df_horas_eq, df_pesos_areas,
    cad_prelim_T_dict, cad_teo_T, df_reap_validas,
    pilar_band_web=0.08,   # ±8% sobre el estimado ACUMULADO de Web por equipo (banda BLANDA, ampliada para dar flexibilidad a restricciones DV)
    pilar_band_busc=0.08,  # ±8% sobre el estimado ACUMULADO de Buscadores por equipo (banda BLANDA, ampliada para dar flexibilidad a restricciones DV)
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

    # Fallback de asignación proporcional minimizando desviación a estimados por pilar (con DV-level)
    def _fallback_asignacion(df_fresh_T, fresh_target_int, estimado_pilar, df_hist_T, equipos_T, pillars,
                             dv_groups_fb=None, est_dv_fb=None):
        df = df_fresh_T.copy()
        # acumulado actual (histórico)
        hist_map_fb = df_hist_T.groupby(['EQUIPO_FINAL', 'PILAR_NORM']).size().unstack(fill_value=0)
        hist_map_fb = hist_map_fb.reindex(index=equipos_T, columns=pillars, fill_value=0)

        # contadores de asignación fresh en curso
        assigned = pd.DataFrame(0, index=equipos_T, columns=pillars, dtype=int)
        remaining = fresh_target_int.reindex(equipos_T).astype(int).fillna(0).to_dict()

        # Pesos DV-level para fallback (dominan sobre team-level)
        DV_WEIGHT_FB = {'Web': 5.0, 'Buscadores': 3.0, 'Redes Sociales': 0.5}

        # Para cada cupón, elige el equipo con más reducción de error cuadrático hacia el estimado
        for c in df.index:
            p = df.at[c, 'PILAR_NORM']
            best_t, best_gain = None, None
            for t in equipos_T:
                if remaining.get(t, 0) <= 0:
                    continue
                before = hist_map_fb.at[t, p] + assigned.at[t, p]
                after  = before + 1
                est    = int(estimado_pilar.loc[t, p]) if (t in estimado_pilar.index and p in estimado_pilar.columns) else 0
                # coste marginal team-level (cuadrático)
                cost_before = (before - est)**2
                cost_after  = (after  - est)**2
                gain = cost_before - cost_after  # queremos maximizar la "ganancia"

                # coste marginal DV-level (NUEVO)
                dv_gain = 0.0
                if dv_groups_fb and est_dv_fb and p in DV_WEIGHT_FB:
                    dv = 'IG' if t in IG_teams else 'XP'
                    dv_tms = dv_groups_fb.get(dv, [])
                    dv_before = sum(hist_map_fb.at[tt, p] + assigned.at[tt, p] for tt in dv_tms)
                    dv_after  = dv_before + 1
                    dv_est    = est_dv_fb.get((dv, p), 0)
                    dv_gain   = DV_WEIGHT_FB[p] * ((dv_before - dv_est)**2 - (dv_after - dv_est)**2)

                total_gain = gain + dv_gain
                # desempate por mayor restante para no bloquear
                tie = remaining.get(t, 0)
                key = (total_gain, tie)
                if (best_gain is None) or (key > best_gain):
                    best_gain = key
                    best_t = t
            if best_t is None:
                # si todos 0, asigna al equipo con menor acumulado total
                sums = (hist_map_fb + assigned).sum(axis=1)
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

    # ---------- Estimaciones a nivel DV (IG/XP agregado) ----------
    dv_groups = {
        'IG': sorted(list(IG_teams & set(equipos_T))),
        'XP': sorted(list(XP_teams & set(equipos_T)))
    }
    dv_pillars = ['Web', 'Buscadores', 'Redes Sociales']  # P.Verticales queda como buffer
    est_dv = {}
    for dv, dv_tms in dv_groups.items():
        for p in dv_pillars:
            est_dv[(dv, p)] = int(estimado_pilar.loc[dv_tms, p].sum())

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

    # Ajuste DV (IG/XP) por pilar en T:
    # - mantiene el total IG agregado original de dv_pillars,
    # - prioriza redondeo fraccional en Web/Buscadores,
    # - usa Redes Sociales como compensador.
    tot_acum_p = (
        df_hist_T["PILAR_NORM"].value_counts()
        .add(df_fresh_T["PILAR_NORM"].value_counts(), fill_value=0)
        .reindex(dv_pillars, fill_value=0)
        .astype(int)
    )
    ig_total_locked = int(sum(int(est_dv[("IG", p)]) for p in dv_pillars))
    ig_float_dv = {p: float(tot_acum_p[p]) * float(share_IG) for p in dv_pillars}

    ig_web = int(np.floor(ig_float_dv["Web"] + 0.5))
    ig_busc = int(np.floor(ig_float_dv["Buscadores"] + 0.5))
    ig_web = int(np.clip(ig_web, 0, int(tot_acum_p["Web"])))
    ig_busc = int(np.clip(ig_busc, 0, int(tot_acum_p["Buscadores"])))
    ig_rs = int(ig_total_locked - (ig_web + ig_busc))

    rs_cap = int(tot_acum_p["Redes Sociales"])
    if ig_rs < 0:
        excess = -ig_rs
        # Web es intocable: solo recorta Buscadores para compensar.
        take = min(excess, ig_busc)
        ig_busc -= take
        excess -= take
        ig_rs = 0
    elif ig_rs > rs_cap:
        # Si RS no tiene capacidad, reparte excedente solo en Buscadores (Web intocable).
        need = ig_rs - rs_cap
        room = int(tot_acum_p["Buscadores"]) - ig_busc
        add = min(need, room)
        ig_busc += add
        need -= add
        ig_rs = rs_cap

    ig_target_dv = {
        "Web": int(ig_web),
        "Buscadores": int(ig_busc),
        "Redes Sociales": int(ig_rs),
    }
    for p in dv_pillars:
        est_dv[("IG", p)] = int(ig_target_dv[p])
        est_dv[("XP", p)] = int(tot_acum_p[p]) - int(ig_target_dv[p])

    n_hist = int(len(df_hist_T))
    n_fresh = int(len(df_fresh_T))
    # df_hist_T viene de df_hist_total, que ya incluye REAP/CORTE válidos.
    # Evita doble conteo en objetivos IG/XP.
    total_T = n_hist + n_fresh

    # Guardia: si no hay fresh, devolvemos coherente
    if n_fresh == 0:
        df_final_T = pd.concat([df_fresh_T.assign(EQUIPO_FINAL=np.nan).iloc[0:0], df_reap_T], ignore_index=True)
        df_cadencia     = pd.DataFrame(columns=["Equipo","Cadencia Original","Cadencia Final","Cadencia Teórica"])
        df_comparativa  = pd.DataFrame(index=equipos_T)
        cad_pilar       = pd.DataFrame(index=equipos_T, columns=pillars).fillna(0.0)
        return df_final_T, df_cadencia, df_comparativa, cad_pilar

    target_IG_total = int(round(share_IG * total_T))
    target_XP_total = total_T - target_IG_total

    fixed_IG = int(cupones_hist.reindex(sorted(IG_teams & set(equipos_T)), fill_value=0).sum())
    fixed_XP = int(cupones_hist.reindex(sorted(XP_teams & set(equipos_T)), fill_value=0).sum())

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

    # --- Variables DV-level (IG/XP agregado) ---
    diff_dv   = {(dv, p): pulp.LpVariable(f"diff_dv_{dv}_{p}", lowBound=0)
                 for dv in ['IG', 'XP'] for p in dv_pillars}
    s_dv_pos  = {(dv, p): pulp.LpVariable(f"s_dv_pos_{dv}_{p}", lowBound=0)
                 for dv in ['IG', 'XP'] for p in dv_pillars}
    s_dv_neg  = {(dv, p): pulp.LpVariable(f"s_dv_neg_{dv}_{p}", lowBound=0)
                 for dv in ['IG', 'XP'] for p in dv_pillars}

    # --- Penalizaciones ---
    # Team-level (existente, secundario)
    PENAL_DIFF = {'Web': 120, 'Buscadores': 45, 'P.Verticales': 0.5, 'Redes Sociales': 0.01}
    BIG_WEB, BIG_BUSC = 9000.0, 4500.0

    # DV-level (NUEVO, prioridad máxima - domina sobre team-level)
    PENAL_DV = {'Web': 500, 'Buscadores': 300, 'Redes Sociales': 50}
    BIG_DV   = {'Web': 50000.0, 'Buscadores': 30000.0, 'Redes Sociales': 5000.0}
    EPS_DV   = {'Web': 0, 'Buscadores': 1, 'Redes Sociales': 3}

    # Pre-computar cupones por pilar (optimización de rendimiento)
    coupons_by_pilar = {p: [c for c in coupons_T if df_fresh_T.at[c, 'PILAR_NORM'] == p] for p in pillars}

    prob = pulp.LpProblem("Distribucion_Area_T", pulp.LpMinimize)
    prob += (
        # DV-level (PRIORIDAD MÁXIMA)
        pulp.lpSum(PENAL_DV[p] * diff_dv[(dv, p)]
                   for dv in ['IG', 'XP'] for p in dv_pillars) +
        pulp.lpSum(BIG_DV[p] * (s_dv_pos[(dv, p)] + s_dv_neg[(dv, p)])
                   for dv in ['IG', 'XP'] for p in dv_pillars) +
        # Team-level (PRIORIDAD SECUNDARIA)
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

    # 3) Ajuste por pilar vs estimado ACUMULADO + bandas blandas (TEAM-LEVEL)
    hist_map = df_hist_T.groupby(['EQUIPO_FINAL', 'PILAR_NORM']).size().to_dict()

    for t in equipos_T:
        for p in pillars:
            h = int(hist_map.get((t, p), 0))
            assign_tp = pulp.lpSum(team_vars[(c, t)] for c in coupons_by_pilar.get(p, []))
            total_tp  = h + assign_tp
            est_tp    = int(estimado_pilar.loc[t, p])
            prob +=  total_tp - est_tp <= diff_pilar[(t, p)]
            prob +=  est_tp - total_tp <= diff_pilar[(t, p)]

        # Bandas blandas Web (team-level)
        hW = int(hist_map.get((t, 'Web'), 0))
        aW = pulp.lpSum(team_vars[(c, t)] for c in coupons_by_pilar.get('Web', []))
        totW = hW + aW
        estW = int(estimado_pilar.loc[t, 'Web'])
        epsW = max(1, int(round(estW * float(pilar_band_web))))
        prob += totW - estW <=  epsW + s_web_pos[t]
        prob += estW - totW <=  epsW + s_web_neg[t]

        # Bandas blandas Buscadores (team-level)
        hB = int(hist_map.get((t, 'Buscadores'), 0))
        aB = pulp.lpSum(team_vars[(c, t)] for c in coupons_by_pilar.get('Buscadores', []))
        totB = hB + aB
        estB = int(estimado_pilar.loc[t, 'Buscadores'])
        epsB = max(1, int(round(estB * float(pilar_band_busc))))
        prob += totB - estB <=  epsB + s_bsc_pos[t]
        prob += estB - totB <=  epsB + s_bsc_neg[t]

    # 4) Restricciones DV-level (IG/XP agregado) - PRIORIDAD MÁXIMA
    for dv in ['IG', 'XP']:
        dv_tms = dv_groups[dv]
        for p in dv_pillars:
            h_dv = sum(int(hist_map.get((t, p), 0)) for t in dv_tms)
            assign_dv_p = pulp.lpSum(
                team_vars[(c, t)]
                for t in dv_tms
                for c in coupons_by_pilar.get(p, [])
            )
            total_dv_p = h_dv + assign_dv_p
            est_dv_p = est_dv[(dv, p)]

            # Capturar desviación absoluta
            prob += total_dv_p - est_dv_p <= diff_dv[(dv, p)]
            prob += est_dv_p - total_dv_p <= diff_dv[(dv, p)]

            # Banda blanda DV (ajustada: ±1 para Web/Busc, ±3 para RS)
            eps = EPS_DV[p]
            prob += total_dv_p - est_dv_p <= eps + s_dv_pos[(dv, p)]
            prob += est_dv_p - total_dv_p <= eps + s_dv_neg[(dv, p)]

    print("🧩 Resolviendo modelo para área T (cadencia exacta + cuotas IG/XP + DV-level + Web/Busc ajustados)...")
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
        df_fresh_T = _fallback_asignacion(df_fresh_T, fresh_target_int, estimado_pilar, df_hist_T, equipos_T, pillars,
                                           dv_groups_fb=dv_groups, est_dv_fb=est_dv)

    # ---------- Ajuste fino (ACUM) ----------
    def _force_pillars_columns_local(m, index=None):
        return _force_pillars_columns(m, index=index)  # usa tu helper global

    def ajuste_fino_cadencia_acum(df_fresh_area, estimado_pilar, df_hist_area):
        df = df_fresh_area.copy()
        pilares_clave = ['Web', 'Buscadores']
        pilares_compensa = ['Redes Sociales', 'P.Verticales']
        equipos = list(map(str, equipos_T))

        # Team-level pesos (secundarios)
        PENAL_AJUSTE = {'Web': 50, 'Buscadores': 30, 'P.Verticales': 0.5, 'Redes Sociales': 0.01}

        # DV-level pesos (PRIORIDAD - dominan sobre team-level)
        PENAL_DV_AJUSTE = {'Web': 500, 'Buscadores': 300, 'Redes Sociales': 50}

        ig_eq = sorted(list(IG_teams & set(equipos)))
        xp_eq = sorted(list(XP_teams & set(equipos)))

        def score_delta(d):
            # Team-level score (existente)
            team_score = sum(PENAL_AJUSTE.get(p, 1) * float(d[p].abs().sum()) for p in d.columns)

            # DV-level score (NUEVO - domina)
            dv_score = 0.0
            for p in ['Web', 'Buscadores', 'Redes Sociales']:
                if p in d.columns:
                    dv_ig = abs(float(d.loc[ig_eq, p].sum())) if ig_eq else 0.0
                    dv_xp = abs(float(d.loc[xp_eq, p].sum())) if xp_eq else 0.0
                    dv_score += PENAL_DV_AJUSTE.get(p, 1) * (dv_ig + dv_xp)

            return dv_score + team_score

        def matriz_acum(dframe):
            m = (dframe.groupby(['EQUIPO_FINAL','PILAR_NORM']).size().unstack(fill_value=0))
            return _force_pillars_columns_local(m, index=equipos)
        acum = pd.concat([df_hist_area[['EQUIPO_FINAL','PILAR_NORM']],
                          df[['EQUIPO_FINAL','PILAR_NORM']]], ignore_index=True)
        count_acum = matriz_acum(acum)
        est_safe = estimado_pilar.reindex(index=equipos, columns=pillars, fill_value=0)
        delta = (count_acum - est_safe).fillna(0).astype('Int64')
        score = score_delta(delta)
        movimientos = 0
        ig_set = set(ig_eq)
        xp_set = set(xp_eq)

        def _same_dv(eq1, eq2):
            return (eq1 in ig_set and eq2 in ig_set) or (eq1 in xp_set and eq2 in xp_set)

        for pilar in pilares_clave:
            while True:
                delta_p = delta[pilar]; exceso_eq = delta_p.idxmax(); falta_eq = delta_p.idxmin()
                if int(delta_p.get(exceso_eq,0)) <= 0 or int(delta_p.get(falta_eq,0)) >= 0: break
                # Web: solo permitir swaps dentro del mismo DV (IG↔IG o XP↔XP)
                if pilar == 'Web' and not _same_dv(exceso_eq, falta_eq):
                    break
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
                            new_score = score_delta(new_delta)
                            if new_score < score:
                                delta = new_delta; score = new_score; movimientos += 1; realizado = True; break
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

    # --- Comparativa DV-level (IG/XP agregado) ---
    print()
    print("📊 Comparativa DV-level (IG/XP agregado) por pilar:")
    for dv in ['IG', 'XP']:
        dv_tms = dv_groups[dv]
        for p in dv_pillars:
            real_val = int(df_real_acum.loc[dv_tms, p].sum()) if p in df_real_acum.columns else 0
            est_val = est_dv.get((dv, p), 0)
            diff_val = real_val - est_val
            marker = "✅" if abs(diff_val) <= EPS_DV.get(p, 1) else "⚠️"
            print(f"  {marker} {dv} {p}: real={real_val} est={est_val} diff={diff_val:+d}")

    return df_final_T, df_cadencia, df_comparativa, cad_pilar
def distribuir_area_E(df_fresh, df_hist_total, df_pesos_areas, df_reap_validas):
    """
    Área E:
    1) Sin cadencia. Objetivo 1: split IG/XP configurable (hist + fresh).
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
    n_fresh = len(df_fresh_E)
    # df_hist_E viene de df_hist_total, que ya incluye REAP/CORTE válidos.
    # Evita doble conteo en objetivos IG/XP.
    total_E = n_hist + n_fresh

    # Targets por directora (enteros)
    share_target_IG = float(E_SHARE_TARGET_IG)
    share_target_XP = 1.0 - share_target_IG
    target_IG_total = int(round(share_target_IG * total_E))
    target_XP_total = total_E - target_IG_total

    # Conteos fijos por directora (hist, que ya incluye reaps válidas)
    hist_counts = df_hist_E['EQUIPO_FINAL'].value_counts().reindex(equipos_E, fill_value=0)

    fixed_IG = int(hist_counts.reindex(IG_equipos, fill_value=0).sum())
    fixed_XP = int(hist_counts.reindex(XP_equipos, fill_value=0).sum())

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
    team_targets = ajustar_redondeo_sum_exacta(
        df_pesos_E['PESO_NORM'] * float(total_E), total=total_E
    )

    # Conteo acumulado actual por equipo (hist ya contiene reaps válidas)
    current_total = hist_counts.reindex(equipos_E, fill_value=0)

    # --- Estimado por pilar (para orientar asignación pillar-aware) ---
    def _pillar_counts_E(df_src, equipos, pils):
        if df_src is None or len(df_src) == 0:
            return pd.DataFrame(0, index=equipos, columns=pils)
        return (df_src.groupby(['EQUIPO_FINAL', 'PILAR_NORM']).size()
                .unstack(fill_value=0)
                .reindex(index=equipos, columns=pils, fill_value=0))

    pillar_total_hist  = df_hist_E['PILAR_NORM'].value_counts().reindex(pillars, fill_value=0)
    pillar_total_reap  = (df_reap_validas_E['PILAR_NORM'].value_counts().reindex(pillars, fill_value=0)
                          if len(df_reap_validas_E) > 0 else pd.Series(0, index=pillars))
    pillar_total_fresh = df_fresh_E['PILAR_NORM'].value_counts().reindex(pillars, fill_value=0)
    pillar_total_acum  = pillar_total_hist + pillar_total_reap + pillar_total_fresh

    estimado_pilar_E = pd.DataFrame(index=equipos_E, columns=pillars, dtype=int)
    for p in pillars:
        estimado_pilar_E[p] = ajustar_redondeo_sum_exacta(
            df_pesos_E['PESO_NORM'] * float(pillar_total_acum[p]),
            total=int(pillar_total_acum[p])
        ).values

    # Objetivo DV por pilar en E:
    # fija Web/Buscadores por fracción IG y compensa en Redes Sociales,
    # manteniendo el total IG agregado de estos pilares.
    dv_pillars_E = ['Web', 'Buscadores', 'Redes Sociales']
    est_dv_E = {p: int(estimado_pilar_E.loc[IG_equipos, p].sum()) for p in dv_pillars_E}
    ig_total_locked_E = int(sum(est_dv_E.values()))
    ig_float_dv_E = {p: float(pillar_total_acum[p]) * share_target_IG for p in dv_pillars_E}

    ig_web_E = int(np.floor(ig_float_dv_E['Web'] + 0.5))
    ig_busc_E = int(np.floor(ig_float_dv_E['Buscadores'] + 0.5))
    ig_web_E = int(np.clip(ig_web_E, 0, int(pillar_total_acum['Web'])))
    ig_busc_E = int(np.clip(ig_busc_E, 0, int(pillar_total_acum['Buscadores'])))
    ig_rs_E = int(ig_total_locked_E - (ig_web_E + ig_busc_E))

    rs_cap_E = int(pillar_total_acum['Redes Sociales'])
    if ig_rs_E < 0:
        excess = -ig_rs_E
        # Web es intocable: solo recorta Buscadores para compensar.
        take = min(excess, ig_busc_E)
        ig_busc_E -= take
        excess -= take
        ig_rs_E = 0
    elif ig_rs_E > rs_cap_E:
        # Si RS no tiene capacidad, reparte excedente solo en Buscadores (Web intocable).
        need = ig_rs_E - rs_cap_E
        room = int(pillar_total_acum['Buscadores']) - ig_busc_E
        add = min(need, room)
        ig_busc_E += add
        need -= add
        ig_rs_E = rs_cap_E

    ig_target_dv_E = {
        'Web': int(ig_web_E),
        'Buscadores': int(ig_busc_E),
        'Redes Sociales': int(ig_rs_E),
    }

    # Conteo actual por equipo y pilar (hist ya incluye reaps válidas)
    current_pilar = _pillar_counts_E(df_hist_E, equipos_E, pillars).astype(int)

    # Orden de pilar por prioridad descendente (Web > Busc > PV > RS)
    PILAR_PESO_E  = {'Web': 50, 'Buscadores': 30, 'P.Verticales': 0.5, 'Redes Sociales': 0.01}
    PILAR_ORDEN_E = sorted(pillars, key=lambda p: -PILAR_PESO_E.get(p, 1))

    # --- Asignación GREEDY pillar-aware con cuota DV ---
    # Procesa pilares de mayor a menor prioridad; dentro de cada pilar reparte
    # la cuota IG/XP proporcionalmente y asigna al equipo con mayor déficit de pilar.
    asignaciones = {}
    IG_remaining  = int(quota_IG_fresh)
    XP_remaining  = int(quota_XP_fresh)
    pilar_counts_fresh = {p: int((df_fresh_E['PILAR_NORM'] == p).sum()) for p in PILAR_ORDEN_E}

    def _mejor_equipo_E(grupo, pilar):
        """Equipo del grupo con mayor déficit de pilar; desempate por déficit total."""
        return max(
            grupo,
            key=lambda e: (
                int(estimado_pilar_E.at[e, pilar]) - int(current_pilar.at[e, pilar]),
                int(team_targets.get(e, 0))         - int(current_total.get(e, 0)),
            )
        )

    for idx_p, pilar in enumerate(PILAR_ORDEN_E):
        coupon_idx_p = df_fresh_E[df_fresh_E['PILAR_NORM'] == pilar].index.tolist()
        if not coupon_idx_p:
            continue
        n_p       = len(coupon_idx_p)
        total_rem = IG_remaining + XP_remaining

        # Reparte los n_p cupones:
        # - En Web/Buscadores/RS usa objetivo DV por pilar (con compensación RS),
        # - En P.Verticales usa reparto proporcional de cuotas restantes.
        if total_rem <= 0:
            n_IG_p, n_XP_p = 0, 0
        else:
            future_coupons = int(sum(pilar_counts_fresh[p] for p in PILAR_ORDEN_E[idx_p + 1:]))
            min_ig_here = max(0, IG_remaining - future_coupons)
            max_ig_here = min(n_p, IG_remaining)

            if pilar in ig_target_dv_E:
                current_ig_p = int(current_pilar.reindex(index=IG_equipos, columns=[pilar], fill_value=0)[pilar].sum())
                desired_ig_here = int(ig_target_dv_E[pilar] - current_ig_p)
                n_IG_p = int(np.clip(desired_ig_here, min_ig_here, max_ig_here))
            else:
                n_IG_p = int(round(n_p * IG_remaining / total_rem))
                n_IG_p = max(min_ig_here, min(max_ig_here, n_IG_p))

            n_XP_p = n_p - n_IG_p

        for idx in coupon_idx_p[:n_IG_p]:
            elegido = (_mejor_equipo_E(IG_equipos, pilar) if IG_equipos
                       else max(equipos_E, key=lambda e: int(team_targets.get(e, 0)) - int(current_total.get(e, 0))))
            asignaciones[idx]                = elegido
            current_pilar.at[elegido, pilar] += 1
            current_total[elegido]           += 1
            IG_remaining                     -= 1

        for idx in coupon_idx_p[n_IG_p:n_IG_p + n_XP_p]:
            elegido = (_mejor_equipo_E(XP_equipos, pilar) if XP_equipos
                       else max(equipos_E, key=lambda e: int(team_targets.get(e, 0)) - int(current_total.get(e, 0))))
            asignaciones[idx]                = elegido
            current_pilar.at[elegido, pilar] += 1
            current_total[elegido]           += 1
            XP_remaining                     -= 1

    # Cupones sin asignar (residuo de redondeo)
    restantes = [i for i in df_fresh_E.index if i not in asignaciones]
    if restantes:
        for idx in restantes:
            pilar_r = df_fresh_E.at[idx, 'PILAR_NORM']
            if IG_remaining > 0 and IG_equipos:
                elegido      = _mejor_equipo_E(IG_equipos, pilar_r)
                IG_remaining -= 1
            elif XP_remaining > 0 and XP_equipos:
                elegido       = _mejor_equipo_E(XP_equipos, pilar_r)
                XP_remaining -= 1
            else:
                elegido = max(equipos_E, key=lambda e: (
                    int(estimado_pilar_E.at[e, pilar_r]) - int(current_pilar.at[e, pilar_r]),
                    int(team_targets.get(e, 0))          - int(current_total.get(e, 0)),
                ))
            asignaciones[idx]                = elegido
            current_pilar.at[elegido, pilar_r] += 1
            current_total[elegido]             += 1

    df_fresh_E['EQUIPO_FINAL'] = df_fresh_E.index.map(asignaciones)

    # --- Salida final (fresh asignado + reaps válidas) ---
    df_final_E = pd.concat([df_fresh_E, df_reap_validas_E], ignore_index=True)

    # --- Comparativas ---
    # Totales finales por equipo = hist + reaps + fresh_asignado
    final_counts_equipo = (
        df_fresh_E['EQUIPO_FINAL'].value_counts().add(hist_counts, fill_value=0).astype(int)
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
        'Share_Target': [share_target_IG, share_target_XP],
        'Delta_puntos': [share_IG - share_target_IG, share_XP - share_target_XP]
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
        msgs.append(
                    f"Nota: No se pudo igualar exactamente IG={share_target_IG:.0%} / XP={share_target_XP:.0%} "
                    "por fijos históricos/reaperturas. "
                    "Se compensó al máximo con los fresh.")

    if msgs:
        for m in msgs:
            print(m)

    # Prints útiles
    print("\n=== Comparativa por EQUIPO (Estimado vs Asignado) ===")
    print(df_comparativa_equipo.sort_index())
    print("\n=== Resumen por DIRECTORA (IG/XP) ===")
    print(df_resumen_directora)
    print("\n=== Estimado vs Real FRESH por Pilar (Area E) ===")
    df_pilar_cmp_E = pd.concat(
        [estimado_pilar_E.add_suffix("_Est"), df_real_pilar_hoy.add_suffix("_Real")],
        axis=1
    ).sort_index()
    print(df_pilar_cmp_E)

    return df_final_E, df_comparativa_equipo, df_real_pilar_hoy, df_resumen_directora
def _ensure_columns(df: pd.DataFrame, cols: Sequence[str]):
    faltan = [c for c in cols if c not in df.columns]
    if faltan:
        raise KeyError(f"Faltan columnas requeridas: {faltan}")
def _pick_key_column(df: pd.DataFrame, prefer: Optional[str]=None) -> str:
    candidates = [
        prefer,
        'ID de la Oportunidad',
        'ID Oportunidad', 'Id Oportunidad (Oportunidad)', 'ID_OPORTUNIDAD',
        'ID_CUPON', 'INDEX_ORIGINAL', 'ID', 'Id'
    ]
    for c in candidates:
        if c and c in df.columns:
            return c
    if 'INDEX_ORIGINAL' in df.columns:
        return 'INDEX_ORIGINAL'
    raise KeyError("No encuentro una columna ID única (pasa 'prefer' con el nombre exacto de tu ID).")
def _key_values(series: pd.Series) -> pd.Series:
    values = series.dropna()
    if pd.api.types.is_numeric_dtype(values):
        values = values.map(lambda x: str(int(x)) if float(x).is_integer() else str(x))
    else:
        values = values.astype(str).str.strip()
    return values
def _fill_na(series: pd.Series) -> pd.Series:
    return series.fillna('<NA>').astype(str)
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
def _lp_reassign_segment(
    dff_seg: pd.DataFrame,
    grupo: List[str],
    slots: Dict[str, int],
    exp_c: pd.DataFrame,
    act_c_base: pd.DataFrame,
    exp_p: pd.DataFrame,
    act_p_base: pd.DataFrame,
    var_pais: str,
    var_prog: str,
    w_country: float,
    w_program: float,
) -> Dict[int, str]:
    """
    Resuelve un ILP de transporte para asignar cupones a equipos minimizando
    la desviación absoluta ponderada de país y programa vs esperado.
    Retorna {coupon_index: equipo_asignado}.
    """
    teams_with_slots = [t for t in grupo if slots.get(t, 0) > 0]
    if not teams_with_slots or dff_seg.empty:
        return {}

    # Agrupar cupones por tipo (país, programa)
    pais_vals = _fill_na(dff_seg[var_pais])
    prog_vals = _fill_na(dff_seg[var_prog])
    type_keys = list(zip(pais_vals, prog_vals))
    type_to_indices: Dict[Tuple[str,str], List[int]] = {}
    for idx, tk in zip(dff_seg.index, type_keys):
        type_to_indices.setdefault(tk, []).append(idx)
    types = sorted(type_to_indices.keys())
    supply = {tk: len(idxs) for tk, idxs in type_to_indices.items()}

    # Columnas de país y programa presentes
    countries = sorted(set(tk[0] for tk in types))
    programs  = sorted(set(tk[1] for tk in types))

    # --- Formulación ILP ---
    prob = pulp.LpProblem("seg_transport", pulp.LpMinimize)

    # Variables: n[t][(c,p)] = cuántos cupones de tipo (c,p) van al equipo t
    n = {}
    for t in teams_with_slots:
        for tk in types:
            n[(t, tk)] = pulp.LpVariable(f"n_{t}_{tk[0]}_{tk[1]}", lowBound=0, cat="Integer")

    # Restricciones de capacidad: cada equipo recibe exactamente sus slots
    for t in teams_with_slots:
        prob += pulp.lpSum(n[(t, tk)] for tk in types) == slots[t], f"cap_{t}"

    # Restricciones de conservación: cada tipo se asigna completamente
    for tk in types:
        prob += pulp.lpSum(n[(t, tk)] for t in teams_with_slots) == supply[tk], f"sup_{tk}"

    # Variables slack para desviación absoluta
    dc_pos = {}; dc_neg = {}
    for t in teams_with_slots:
        for c in countries:
            dc_pos[(t,c)] = pulp.LpVariable(f"dcp_{t}_{c}", lowBound=0)
            dc_neg[(t,c)] = pulp.LpVariable(f"dcn_{t}_{c}", lowBound=0)
    dp_pos = {}; dp_neg = {}
    for t in teams_with_slots:
        for p in programs:
            dp_pos[(t,p)] = pulp.LpVariable(f"dpp_{t}_{p}", lowBound=0)
            dp_neg[(t,p)] = pulp.LpVariable(f"dpn_{t}_{p}", lowBound=0)

    # Restricciones de desviación: |actual - expected| <= d_pos + d_neg
    for t in teams_with_slots:
        for c in countries:
            # actual_country[t][c] = hist_c[t][c] + sum de n[t][(c,p)] para todo p
            hist_tc = int(act_c_base.at[t, c]) if (t in act_c_base.index and c in act_c_base.columns) else 0
            fresh_tc = pulp.lpSum(n[(t, tk)] for tk in types if tk[0] == c)
            actual_tc = hist_tc + fresh_tc
            exp_tc = int(exp_c.at[t, c]) if (t in exp_c.index and c in exp_c.columns) else 0
            prob += actual_tc - exp_tc <= dc_pos[(t,c)]
            prob += exp_tc - actual_tc <= dc_neg[(t,c)]

        for p in programs:
            hist_tp = int(act_p_base.at[t, p]) if (t in act_p_base.index and p in act_p_base.columns) else 0
            fresh_tp = pulp.lpSum(n[(t, tk)] for tk in types if tk[1] == p)
            actual_tp = hist_tp + fresh_tp
            exp_tp = int(exp_p.at[t, p]) if (t in exp_p.index and p in exp_p.columns) else 0
            prob += actual_tp - exp_tp <= dp_pos[(t,p)]
            prob += exp_tp - actual_tp <= dp_neg[(t,p)]

    # Objetivo: minimizar desviación ponderada
    prob += (
        w_country * pulp.lpSum(dc_pos[(t,c)] + dc_neg[(t,c)] for t in teams_with_slots for c in countries) +
        w_program * pulp.lpSum(dp_pos[(t,p)] + dp_neg[(t,p)] for t in teams_with_slots for p in programs)
    )

    prob.solve(pulp.PULP_CBC_CMD(msg=False, timeLimit=30))

    if pulp.LpStatus[prob.status] != "Optimal":
        return {}  # fallback: caller usará asignación original

    # Extraer solución y asignar cupones específicos
    assigned: Dict[int, str] = {}
    for tk in types:
        coupon_idxs = list(type_to_indices[tk])
        pos = 0
        for t in teams_with_slots:
            count = int(round(pulp.value(n[(t, tk)]) or 0))
            for idx in coupon_idxs[pos:pos + count]:
                assigned[idx] = t
            pos += count
        # Cupones residuales (por redondeo float) → asignar al equipo con más slots restantes
        for idx in coupon_idxs[pos:]:
            remaining_slots = {t: slots[t] - sum(1 for v in assigned.values() if v == t) for t in teams_with_slots}
            best_t = max(remaining_slots, key=remaining_slots.get)
            assigned[idx] = best_t

    return assigned
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

    # Esperados: basados en (HIST+FRESH) para saber el volumen total a distribuir.
    # Actuales: solo HIST; el greedy (y el 2-opt) los incrementarán al asignar cada fresh.
    _empty = dff.iloc[0:0]
    exp_c = _build_expected_by_pillar_value(dfh, dff, grupo, pesos_norm_g, var_pais)
    act_c = _actual_accum_by_pillar_value  (dfh, _empty, grupo, var_pais)
    all_c = sorted(set(exp_c.columns) | set(act_c.columns))
    exp_c = exp_c.reindex(index=grupo, columns=all_c, fill_value=0).astype(int)
    act_c = act_c.reindex(index=grupo, columns=all_c, fill_value=0).astype(int)

    exp_p = _build_expected_by_pillar_value(dfh, dff, grupo, pesos_norm_g, var_prog)
    act_p = _actual_accum_by_pillar_value  (dfh, _empty, grupo, var_prog)
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

    # Asignación por segmentos con slots exactos (ILP de transporte)
    for seg_key, idxs in segments:
        if not idxs:
            continue

        seg_counts_before = _segment_team_counts(idxs)
        slots = {t: int(seg_counts_before[t]) for t in grupo if int(seg_counts_before[t]) > 0}
        if not slots:
            continue

        dff_seg = dff.loc[idxs]
        assigned = _lp_reassign_segment(
            dff_seg=dff_seg,
            grupo=grupo,
            slots=slots,
            exp_c=exp_c, act_c_base=act_c.copy(),
            exp_p=exp_p, act_p_base=act_p.copy(),
            var_pais=var_pais, var_prog=var_prog,
            w_country=w_country, w_program=w_program,
        )

        if assigned:
            for i, t in assigned.items():
                dff.at[i, 'EQUIPO_FINAL'] = t
            for i, t in assigned.items():
                pais_i = _fill_na(pd.Series([dff.at[i, var_pais]])).iloc[0]
                prog_i = _fill_na(pd.Series([dff.at[i, var_prog]])).iloc[0]
                if pais_i in act_c.columns and t in act_c.index:
                    act_c.at[t, pais_i] += 1
                if prog_i in act_p.columns and t in act_p.index:
                    act_p.at[t, prog_i] += 1

        # Verificación dura del segmento (slots invariantes)
        seg_counts_after = dff.loc[idxs, 'EQUIPO_FINAL'].value_counts().reindex(grupo, fill_value=0).astype(int)
        _assert_equal_msg(seg_counts_after, seg_counts_before,
                          f"Slots del segmento {seg_key} cambiaron dentro del pilar (lock={lock_cols}).")

    return dff


def _balanced_reassign_group_cross_pilar(
    df_hist_g: pd.DataFrame,
    df_fresh_g: pd.DataFrame,
    grupo: Sequence[str],
    pesos_norm_g: pd.Series,
    var_pais: str,
    var_prog: str,
    w_country: float,
    w_program: float,
    pillars: Sequence[str] = PILLARS,
) -> pd.DataFrame:
    """
    Reasigna TODOS los fresh de un grupo (todos los pilares a la vez).
    Invariantes:
      - Total FRESH por equipo (suma de todos los pilares) = no cambia.
      - Total FRESH por pilar a nivel de grupo = no cambia (por supply).
    Se libera:
      - FRESH por equipo×pilar: PUEDE cambiar (compensación cruzada entre pilares).
    Objetivo: minimizar desvío de país y programa vs esperado, pilar a pilar.
    """
    dff = df_fresh_g.copy()
    dfh = df_hist_g.copy()

    if dff.empty:
        return dff

    # Normalizar NA
    for col in [var_pais, var_prog]:
        if col in dff.columns:
            dff[col] = _fill_na(dff[col])
        if col in dfh.columns:
            dfh[col] = _fill_na(dfh[col])

    _empty = dff.iloc[0:0]

    # --- Esperados y actuales POR PILAR ---
    exp_c_pil: Dict[str, pd.DataFrame] = {}
    act_c_pil: Dict[str, pd.DataFrame] = {}
    exp_p_pil: Dict[str, pd.DataFrame] = {}
    act_p_pil: Dict[str, pd.DataFrame] = {}

    for p in pillars:
        dfh_p = dfh[dfh['PILAR_NORM'] == p]
        dff_p = dff[dff['PILAR_NORM'] == p]
        if dff_p.empty and dfh_p.empty:
            continue

        ec = _build_expected_by_pillar_value(dfh_p, dff_p, grupo, pesos_norm_g, var_pais)
        ac = _actual_accum_by_pillar_value(dfh_p, _empty, grupo, var_pais)
        all_c = sorted(set(ec.columns) | set(ac.columns))
        exp_c_pil[p] = ec.reindex(index=grupo, columns=all_c, fill_value=0).astype(int)
        act_c_pil[p] = ac.reindex(index=grupo, columns=all_c, fill_value=0).astype(int)

        ep = _build_expected_by_pillar_value(dfh_p, dff_p, grupo, pesos_norm_g, var_prog)
        ap = _actual_accum_by_pillar_value(dfh_p, _empty, grupo, var_prog)
        all_pr = sorted(set(ep.columns) | set(ap.columns))
        exp_p_pil[p] = ep.reindex(index=grupo, columns=all_pr, fill_value=0).astype(int)
        act_p_pil[p] = ap.reindex(index=grupo, columns=all_pr, fill_value=0).astype(int)

    if not exp_c_pil:
        return dff

    # --- Tipo clave: (país, programa, pilar) ---
    pais_vals = _fill_na(dff[var_pais])
    prog_vals = _fill_na(dff[var_prog])
    pil_vals = dff['PILAR_NORM']
    type_keys_raw = list(zip(pais_vals, prog_vals, pil_vals))

    type_to_indices: Dict[Tuple[str, str, str], List[int]] = {}
    for idx, tk in zip(dff.index, type_keys_raw):
        type_to_indices.setdefault(tk, []).append(idx)
    types = sorted(type_to_indices.keys())
    supply = {tk: len(idxs) for tk, idxs in type_to_indices.items()}

    # Capacidad: total fresh por equipo (invariante de la 1ª etapa)
    total_per_team = dff['EQUIPO_FINAL'].value_counts().reindex(grupo, fill_value=0)
    teams = [t for t in grupo if int(total_per_team[t]) > 0]

    if not teams or not types:
        return dff

    # --- ILP ---
    prob = pulp.LpProblem("cross_pilar", pulp.LpMinimize)

    n: Dict[Tuple[str, Tuple], pulp.LpVariable] = {}
    for t in teams:
        for tk in types:
            n[(t, tk)] = pulp.LpVariable(
                f"n_{t}_{tk[0][:8]}_{tk[1][:8]}_{tk[2][:4]}", lowBound=0, cat="Integer"
            )

    # Supply: cada (país, programa, pilar) se asigna completamente
    for tk in types:
        prob += pulp.lpSum(n[(t, tk)] for t in teams) == supply[tk], f"sup_{hash(tk)}"

    # Capacidad: total por equipo = invariante
    for t in teams:
        prob += pulp.lpSum(n[(t, tk)] for tk in types) == int(total_per_team[t]), f"cap_{t}"

    # --- Desviaciones por PILAR×EQUIPO×PAÍS y PILAR×EQUIPO×PROGRAMA ---
    dc_pos: Dict = {}
    dc_neg: Dict = {}
    dp_pos: Dict = {}
    dp_neg: Dict = {}

    for p in exp_c_pil:
        types_p = [tk for tk in types if tk[2] == p]
        countries_p = sorted(exp_c_pil[p].columns)
        programs_p = sorted(exp_p_pil[p].columns)

        for t in teams:
            for c in countries_p:
                key = (t, p, c)
                dc_pos[key] = pulp.LpVariable(f"dcp_{t}_{p[:4]}_{c[:8]}", lowBound=0)
                dc_neg[key] = pulp.LpVariable(f"dcn_{t}_{p[:4]}_{c[:8]}", lowBound=0)
                hist_tc = int(act_c_pil[p].at[t, c]) if t in act_c_pil[p].index else 0
                fresh_tc = pulp.lpSum(n[(t, tk)] for tk in types_p if tk[0] == c)
                exp_tc = int(exp_c_pil[p].at[t, c]) if t in exp_c_pil[p].index else 0
                prob += (hist_tc + fresh_tc) - exp_tc <= dc_pos[key]
                prob += exp_tc - (hist_tc + fresh_tc) <= dc_neg[key]

            for pr in programs_p:
                key = (t, p, pr)
                dp_pos[key] = pulp.LpVariable(f"dpp_{t}_{p[:4]}_{pr[:8]}", lowBound=0)
                dp_neg[key] = pulp.LpVariable(f"dpn_{t}_{p[:4]}_{pr[:8]}", lowBound=0)
                hist_tp = int(act_p_pil[p].at[t, pr]) if t in act_p_pil[p].index else 0
                fresh_tp = pulp.lpSum(n[(t, tk)] for tk in types_p if tk[1] == pr)
                exp_tp = int(exp_p_pil[p].at[t, pr]) if t in exp_p_pil[p].index else 0
                prob += (hist_tp + fresh_tp) - exp_tp <= dp_pos[key]
                prob += exp_tp - (hist_tp + fresh_tp) <= dp_neg[key]

    # Objetivo
    prob += (
        w_country * pulp.lpSum(dc_pos[k] + dc_neg[k] for k in dc_pos)
        + w_program * pulp.lpSum(dp_pos[k] + dp_neg[k] for k in dp_pos)
    )

    prob.solve(pulp.PULP_CBC_CMD(msg=False, timeLimit=60))

    if pulp.LpStatus[prob.status] != "Optimal":
        print(f"[WARN] Cross-pilar ILP sin solucion optima para grupo {grupo}. Se mantiene asignacion original.")
        return dff

    # --- Mapear solución a cupones individuales ---
    assigned: Dict[int, str] = {}
    for tk in types:
        coupon_idxs = list(type_to_indices[tk])
        pos = 0
        for t in teams:
            count = int(round(pulp.value(n[(t, tk)]) or 0))
            for idx in coupon_idxs[pos:pos + count]:
                assigned[idx] = t
            pos += count
        # Residuales por redondeo
        for idx in coupon_idxs[pos:]:
            remaining = {t: int(total_per_team[t]) - sum(1 for v in assigned.values() if v == t)
                         for t in teams}
            best_t = max(remaining, key=remaining.get)
            assigned[idx] = best_t

    for i, t in assigned.items():
        dff.at[i, 'EQUIPO_FINAL'] = t

    return dff


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
                print(f"[WARN Area {area} Grupo {grupo}] descartado "
                      f"(H+F eq*pilar={inv_pilar_ok}, FRESH eq={inv_fresh_ok}, "
                      f"FRESH eq*pilar={inv_fresh_pilar_ok}, LOCK={inv_lock_ok}).")

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
                print(f"[ROLLBACK Area {area}] cambios descartados "
                      f"(FRESH*PILAR / FRESH por equipo / LOCK no invariante).")

    # Aplicar updates aprobados (global)
    if updates_all:
        updates = (pd.concat(updates_all, ignore_index=True)
                   .drop_duplicates(subset=['INDEX_ORIGINAL'], keep='last'))
        m = dict(zip(updates['INDEX_ORIGINAL'], updates['EQUIPO_FINAL']))
        res.loc[mask_fresh, 'EQUIPO_FINAL'] = res.loc[mask_fresh, 'INDEX_ORIGINAL'].map(m)\
                                                .fillna(res.loc[mask_fresh, 'EQUIPO_FINAL'])

    res = res.sort_values('INDEX_ORIGINAL', kind='stable').reset_index(drop=True)
    mask_fresh = res['TIPO_REPARTO'].eq('FRESH')

    # Blindajes globales (post)
    fresh_after_global = _counts_fresh_by_team(res.loc[mask_fresh, ['EQUIPO_FINAL']])\
                         .reindex(fresh_baseline_global.index, fill_value=0)
    _assert_equal_msg(fresh_after_global, fresh_baseline_global, "2a etapa: cambio el #FRESH por equipo.")

    tp_after_global = _counts_histplusfresh_by_team_pilar(
        df_hist_total_clean, res.loc[mask_fresh, ['EQUIPO_FINAL','PILAR_NORM']]
    ).reindex(index=tp_baseline_global.index, columns=tp_baseline_global.columns, fill_value=0)
    _assert_equal_msg(tp_after_global, tp_baseline_global, "2a etapa: cambio el (equipo*pilar) HIST+FRESH.")

    fresh_pilar_after_global = _counts_fresh_by_team_pilar(res.loc[mask_fresh, ['EQUIPO_FINAL','PILAR_NORM']])\
                               .reindex(index=fresh_pilar_before_global.index, columns=fresh_pilar_before_global.columns, fill_value=0)
    _assert_equal_msg(fresh_pilar_after_global, fresh_pilar_before_global, "2a etapa: cambio el FRESH*PILAR global.")

    return res
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

    keys_hist = set(_key_values(dfh[key_h]).unique())
    keys_final_reap = set(_key_values(reap_final[key_f]).unique())
    dup_keys = keys_final_reap & keys_hist

    if not dup_keys:
        if verbose: print("[DEDUP] No hay REAP duplicadas entre HIST y FINAL.")
        return dfh, dff

    key_f_norm = _key_values(dff[key_f])
    mask_drop = dff['TIPO_REPARTO'].eq('REAP') & key_f_norm.reindex(dff.index).isin(dup_keys).fillna(False)
    n_drop = int(mask_drop.sum())
    dff = dff.loc[~mask_drop].copy()
    if verbose:
        print(f"[DEDUP] REAP duplicadas detectadas por clave: {len(dup_keys)} únicas.")
        print(f"[DEDUP] Filas REAP eliminadas de df_final_total: {n_drop} (se cuentan como HIST).")
    return dfh, dff
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
        keys_base = set(_key_values(base[key_res]).unique())
        key_org_norm = _key_values(df_reap_org[key_org])
        reap_to_add = df_reap_org.loc[
            df_reap_org[key_org].notna() &
            ~key_org_norm.reindex(df_reap_org.index).isin(keys_base).fillna(False)
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
