"""
Punto de entrada para el ejecutable Atenea empaquetado con PyInstaller.
Pide hora de corte y periodo de extraccion antes de ejecutar el pipeline.
"""
from __future__ import annotations

import sys
import traceback
from dataclasses import replace
from datetime import datetime
from pathlib import Path


def _get_workspace() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _ask_cutoff() -> datetime:
    fmt = "%d/%m/%Y %H:%M"
    print("\n" + "=" * 50)
    print("  HORA DE CORTE")
    print("  Oportunidades creadas ANTES de esta fecha/hora")
    print("  se trataran como REAP (no redistribuidas).")
    print("=" * 50)
    while True:
        raw = input("  Introduce la hora de corte (DD/MM/AAAA HH:MM): ").strip()
        try:
            cutoff = datetime.strptime(raw, fmt)
            print(f"  -> Corte establecido: {cutoff.strftime(fmt)}\n")
            return cutoff
        except ValueError:
            print("  Formato incorrecto. Ejemplo valido: 05/08/2025 10:00\n")


def _ask_date(label: str) -> str:
    fmt = "%Y-%m-%d"
    while True:
        raw = input(f"  {label} (YYYY-MM-DD): ").strip()
        try:
            datetime.strptime(raw, fmt)
            return raw
        except ValueError:
            print("  Formato incorrecto. Ejemplo valido: 2026-05-09\n")


def _ask_atenea_period() -> tuple[str, str]:
    print("\n" + "=" * 50)
    print("  PERIODO QBCN ATENEA")
    print("  Fecha inicio y fecha fin se interpretan en hora Madrid.")
    print("=" * 50)
    while True:
        fecha_inicio = _ask_date("Fecha inicio")
        fecha_fin = _ask_date("Fecha fin   ")
        if fecha_inicio <= fecha_fin:
            print(f"  -> Periodo establecido: {fecha_inicio} -> {fecha_fin}\n")
            return fecha_inicio, fecha_fin
        print("  La fecha inicio no puede ser posterior a la fecha fin.\n")


def main() -> None:
    from distribucion_obs import PipelineConfig, run_and_export

    workspace = _get_workspace()
    print(f"Carpeta de trabajo: {workspace}")

    cutoff_dt = _ask_cutoff()
    fecha_inicio, fecha_fin = _ask_atenea_period()

    cfg = PipelineConfig.default(workspace_dir=workspace)
    cfg = replace(
        cfg,
        input_source="atenea",
        atenea_fecha_inicio=fecha_inicio,
        atenea_fecha_fin=fecha_fin,
        atenea_export_excel=True,
    )
    result = run_and_export(cfg, cutoff_dt=cutoff_dt)
    print(
        f"\nOK - control filas entrada/salida: "
        f"{result.input_rows} -> {len(result.df_final_export)}"
    )
    print(f"Archivo generado: {cfg.output_path}")
    print(f"Extraccion generada: leads_{fecha_inicio}_{fecha_fin}.xlsx")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        print("\n[ERROR] El proceso termino con errores:")
        traceback.print_exc()
    finally:
        input("\nPresiona ENTER para cerrar...")
