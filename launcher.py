"""
Punto de entrada para el ejecutable empaquetado con PyInstaller.
Resuelve la ruta del workspace correctamente tanto en modo frozen (.exe) como en desarrollo.
"""
from __future__ import annotations

import sys
import traceback
from datetime import datetime
from pathlib import Path


def _get_workspace() -> Path:
    """
    En modo frozen (PyInstaller --onefile), __file__ apunta a la carpeta temporal
    de extraccion (_MEIxxx). Se usa sys.executable para obtener la carpeta del .exe.
    En desarrollo normal, usa la carpeta del script.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _ask_cutoff() -> datetime:
    """Pide la hora de corte al usuario por terminal y la valida."""
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
            print(f"  Formato incorrecto. Ejemplo valido: 05/08/2025 10:00\n")


def main() -> None:
    from distribucion_obs import PipelineConfig, run_and_export

    workspace = _get_workspace()
    print(f"Carpeta de trabajo: {workspace}")

    cutoff_dt = _ask_cutoff()

    cfg = PipelineConfig.default(workspace_dir=workspace)
    result = run_and_export(cfg, cutoff_dt=cutoff_dt)
    print(
        f"\nOK - control filas entrada/salida: "
        f"{result.input_rows} -> {len(result.df_final_export)}"
    )
    print(f"Archivo generado: {cfg.output_path}")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        print("\n[ERROR] El proceso termino con errores:")
        traceback.print_exc()
    finally:
        input("\nPresiona ENTER para cerrar...")
