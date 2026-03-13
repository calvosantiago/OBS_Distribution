from __future__ import annotations

import pandas as pd

from .config import PipelineConfig
from .extracted_functions import (
    distribuir_area_E,
    distribuir_area_T,
    distribuir_area_X,
    run_segunda_etapa_v19,
)


def run_first_stage(
    df_fresh: pd.DataFrame,
    df_hist_total: pd.DataFrame,
    df_horas_eq: pd.DataFrame,
    df_pesos_areas: pd.DataFrame,
    cad_prelim_abc_dict: dict[str, float],   # cadencias MST/ESP para todos los equipos A,B,C
    cad_prelim_t_dict: dict[str, float],
    cad_teo: dict[str, float],
    df_reap_validas: pd.DataFrame,
    df_corte: pd.DataFrame,
    df_special: pd.DataFrame,
    cfg: PipelineConfig,
) -> pd.DataFrame:
    equipos_a = ["Equipo_A1", "Equipo_A2"]
    equipos_b = ["Equipo_B1", "Equipo_B2"]
    equipos_c = ["Equipo_C1", "Equipo_C2"]

    df_final_a, _, _, _ = distribuir_area_X(
        df_fresh,
        df_hist_total,
        df_horas_eq,
        df_pesos_areas,
        cad_prelim_abc_dict,
        cad_teo["A"],
        df_reap_validas,
        "A",
        equipos_a,
        pilar_band_web=cfg.pilar_band_web,
        pilar_band_busc=cfg.pilar_band_busc,
        time_limit=cfg.time_limit,
    )
    df_final_b, _, _, _ = distribuir_area_X(
        df_fresh,
        df_hist_total,
        df_horas_eq,
        df_pesos_areas,
        cad_prelim_abc_dict,
        cad_teo["B"],
        df_reap_validas,
        "B",
        equipos_b,
        pilar_band_web=cfg.pilar_band_web,
        pilar_band_busc=cfg.pilar_band_busc,
        time_limit=cfg.time_limit,
    )
    df_final_c, _, _, _ = distribuir_area_X(
        df_fresh,
        df_hist_total,
        df_horas_eq,
        df_pesos_areas,
        cad_prelim_abc_dict,
        cad_teo["C"],
        df_reap_validas,
        "C",
        equipos_c,
        pilar_band_web=cfg.pilar_band_web,
        pilar_band_busc=cfg.pilar_band_busc,
        time_limit=cfg.time_limit,
    )
    df_final_t, _, _, _ = distribuir_area_T(
        df_fresh,
        df_hist_total,
        df_horas_eq,
        df_pesos_areas,
        cad_prelim_t_dict,
        cad_teo["T"],
        df_reap_validas,
        pilar_band_web=cfg.pilar_band_web,
        pilar_band_busc=cfg.pilar_band_busc,
        time_limit=cfg.time_limit,
    )
    df_final_e, _, _, _ = distribuir_area_E(
        df_fresh=df_fresh,
        df_hist_total=df_hist_total,
        df_pesos_areas=df_pesos_areas,
        df_reap_validas=df_reap_validas,
    )

    df_final_total = pd.concat(
        [df_final_a, df_final_b, df_final_c, df_final_t, df_final_e, df_corte, df_special],
        ignore_index=True,
    )
    df_final_total = (
        df_final_total.sort_values(["ID de la Oportunidad", "INDEX_ORIGINAL"], kind="stable")
        .drop_duplicates(subset=["ID de la Oportunidad"], keep="first")
        .sort_values("INDEX_ORIGINAL", kind="stable")
        .reset_index(drop=True)
    )
    return df_final_total


def run_second_stage(
    df_final_total: pd.DataFrame,
    df_hist_total: pd.DataFrame,
    df_pesos_areas: pd.DataFrame,
    df_horas_eq: pd.DataFrame,
    cfg: PipelineConfig,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    return run_segunda_etapa_v19(
        df_final_total=df_final_total,
        df_hist_total=df_hist_total,
        df_pesos_areas=df_pesos_areas,
        df_horas_eq=df_horas_eq,
        w_country=cfg.stage2_w_country,
        w_program=cfg.stage2_w_program,
    )

