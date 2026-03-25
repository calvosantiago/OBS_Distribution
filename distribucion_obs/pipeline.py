from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime

import pandas as pd

from .config import PipelineConfig
from .loaders import load_base_inputs, load_pilares_map, load_sudoku_raw
from .preprocess import (
    build_hist_qbcn,
    build_weights,
    compute_cadencia_preliminar,
    enrich_with_area_country_pillar,
    preprocess_open_coupons,
    split_reap_fresh_hist,
)
from .stages import run_first_stage, run_second_stage


@dataclass
class PipelineResult:
    input_rows: int
    df_pesos_actuales: pd.DataFrame
    df_pesos_areas: pd.DataFrame
    df_final_stage1: pd.DataFrame
    df_final_stage2_clean: pd.DataFrame
    df_final_export: pd.DataFrame
    df_hist_total: pd.DataFrame
    df_fresh: pd.DataFrame


def run_pipeline(cfg: PipelineConfig, cutoff_dt: datetime | None = None) -> PipelineResult:
    df_cupones, df_hist, df_areas, df_paises = load_base_inputs(cfg)
    input_rows = len(df_cupones)
    df_pilares_norm = load_pilares_map(cfg)
    df_sudoku_raw = load_sudoku_raw(cfg)

    def _count_pmax(df: pd.DataFrame) -> int:
        cols = [
            "SubPillar (Campaña de origen) (Campaña)",
            "SubPillar Name (Campaña de origen) (Campaña)",
        ]
        if df.empty:
            return 0
        mask = pd.Series(False, index=df.index)
        for col in cols:
            if col in df.columns:
                mask = mask | df[col].astype(str).str.contains("pmax", case=False, na=False)
        return int(mask.sum())

    pmax_hoy = _count_pmax(df_cupones)
    pmax_hist = _count_pmax(df_hist)
    print(f"\n=== CONTROL PMAX ===")
    print(f"PMAX hoy (passthrough, fuera de cálculo): {pmax_hoy}")
    print(f"PMAX histórico (excluidos de cálculo): {pmax_hist}")

    df_cupones, df_hist = enrich_with_area_country_pillar(
        df_cupones=df_cupones,
        df_hist=df_hist,
        df_areas=df_areas,
        df_paises=df_paises,
        df_pilares_norm=df_pilares_norm,
    )
    df_pesos_actuales, df_horas_eq, df_pesos_areas = build_weights(df_sudoku_raw)
    print("\n=== PESO BASE TRANSVERSAL (HORAS) ===")
    print(
        df_pesos_actuales[["EQUIPO", "HORAS", "PESO_BASE"]]
        .assign(PCT=lambda d: (d["PESO_BASE"] * 100).round(4))
        .sort_values("EQUIPO")
        .to_string(index=False)
    )
    print("\n=== PESOS POR ÁREA (usados en distribución) ===")
    for area in ["A", "B", "C", "T", "E"]:
        d = df_pesos_areas[df_pesos_areas["AREA"] == area][["EQUIPO", "PESO_BASE"]].copy()
        if d.empty:
            continue
        d["PCT"] = (d["PESO_BASE"] * 100).round(4)
        print(f"\nAREA {area}")
        print(d.sort_values("EQUIPO").to_string(index=False))

    recup_hoy = int(df_cupones["PILAR_NORM"].isin(["REF/RECUP", "OTROS"]).sum()) if "PILAR_NORM" in df_cupones.columns else 0
    print(f"\n=== CONTROL REF/RECUP + OTROS ===")
    print(f"REF/RECUP y OTROS hoy (passthrough, fuera de cálculo): {recup_hoy}")

    df_hist_qbcn = build_hist_qbcn(df_hist, df_areas)
    df_cupones_open, df_special = preprocess_open_coupons(df_cupones)
    df_fresh, df_reap_validas, df_corte, df_hist_total = split_reap_fresh_hist(
        df_cupones_open=df_cupones_open,
        df_hist_qbcn=df_hist_qbcn,
        cfg=cfg,
        cutoff_dt=cutoff_dt,
    )

    def _area_team_matrix(df: pd.DataFrame, areas: list[str], teams: list[str]) -> pd.DataFrame:
        if df.empty:
            m = pd.DataFrame(0, index=teams, columns=areas, dtype=int)
            m["TOTAL"] = 0
            return m
        x = df.copy()
        x["AREA"] = x["AREA"].astype(str).str.strip()
        x["EQUIPO_FINAL"] = x["EQUIPO_FINAL"].astype(str).str.strip()
        m = (
            x[x["AREA"].isin(areas) & x["EQUIPO_FINAL"].isin(teams)]
            .groupby(["EQUIPO_FINAL", "AREA"])
            .size()
            .unstack(fill_value=0)
            .reindex(index=teams, columns=areas, fill_value=0)
            .astype(int)
        )
        m["TOTAL"] = m.sum(axis=1)
        return m

    areas_show = ["A", "B", "C", "E", "T"]
    teams_show = (
        df_pesos_areas[df_pesos_areas["AREA"].isin(areas_show)]["EQUIPO"]
        .astype(str)
        .dropna()
        .unique()
        .tolist()
    )
    teams_show = sorted(teams_show)
    hist_before = _area_team_matrix(df_hist_total, areas_show, teams_show)
    print("\n=== HISTÓRICO ANTES DE DISTRIBUCIÓN (AREA x EQUIPO) ===")
    print(hist_before.to_string())
    print("\nTotales por área (antes):")
    print(hist_before[areas_show].sum(axis=0).to_string())

    cad_prelim_a_dict, cad_prelim_t_dict, cad_teo = compute_cadencia_preliminar(
        df_hist_total=df_hist_total,
        df_fresh=df_fresh,
        df_horas_eq=df_horas_eq,
    )

    df_final_stage1 = run_first_stage(
        df_fresh=df_fresh,
        df_hist_total=df_hist_total,
        df_horas_eq=df_horas_eq,
        df_pesos_areas=df_pesos_areas,
        cad_prelim_abc_dict=cad_prelim_a_dict,
        cad_prelim_t_dict=cad_prelim_t_dict,
        cad_teo=cad_teo,
        df_reap_validas=df_reap_validas,
        df_corte=df_corte,
        df_special=df_special,
        cfg=cfg,
    )
    df_final_stage2_clean, df_final_export = run_second_stage(
        df_final_total=df_final_stage1,
        df_hist_total=df_hist_total,
        df_pesos_areas=df_pesos_areas,
        df_horas_eq=df_horas_eq,
        cfg=cfg,
    )

    # "Después" como acumulado histórico + fresh final asignado (no altera histórico base).
    fresh_final = df_final_export.copy()
    if "TIPO_REPARTO" in fresh_final.columns:
        fresh_final = fresh_final[fresh_final["TIPO_REPARTO"].astype(str).str.strip().eq("FRESH")]
    fresh_final_counts = fresh_final[["AREA", "EQUIPO_FINAL"]].copy()
    hist_after = hist_before.copy()
    if not fresh_final_counts.empty:
        add = _area_team_matrix(fresh_final_counts, areas_show, teams_show)
        hist_after[areas_show] = hist_before[areas_show].add(add[areas_show], fill_value=0).astype(int)
        hist_after["TOTAL"] = hist_after[areas_show].sum(axis=1)

    print("\n=== HISTÓRICO DESPUÉS (HIST + FRESH FINAL) ===")
    print(hist_after.to_string())
    print("\nTotales por área (después):")
    print(hist_after[areas_show].sum(axis=0).to_string())
    if cfg.enforce_row_count_control:
        output_rows = len(df_final_export)
        if output_rows != input_rows:
            detail = ""
            if "INDEX_ORIGINAL" in df_final_export.columns:
                in_idx = set(range(input_rows))
                out_idx = set(df_final_export["INDEX_ORIGINAL"].dropna().astype(int).tolist())
                missing = sorted(in_idx - out_idx)
                extra = sorted(out_idx - in_idx)
                detail = f" missing_INDEX_ORIGINAL={missing[:20]} extra_INDEX_ORIGINAL={extra[:20]}"
            raise ValueError(
                f"Control de filas falló: entrada={input_rows}, salida={output_rows}.{detail}"
            )
    return PipelineResult(
        input_rows=input_rows,
        df_pesos_actuales=df_pesos_actuales,
        df_pesos_areas=df_pesos_areas,
        df_final_stage1=df_final_stage1,
        df_final_stage2_clean=df_final_stage2_clean,
        df_final_export=df_final_export,
        df_hist_total=df_hist_total,
        df_fresh=df_fresh,
    )


def run_and_export(cfg: PipelineConfig, cutoff_dt: datetime | None = None) -> PipelineResult:
    result = run_pipeline(cfg, cutoff_dt=cutoff_dt)
    result.df_final_export.to_excel(cfg.output_path, index=False)
    return result
