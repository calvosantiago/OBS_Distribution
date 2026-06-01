from __future__ import annotations

import argparse
from dataclasses import replace
from pathlib import Path
import sys

from distribucion_obs import PipelineConfig, run_and_export


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Ejecuta la distribución OBS.")
    parser.add_argument(
        "--source",
        choices=["excel", "atenea"],
        default="excel",
        help="Origen de datos para cupones e historico.",
    )
    parser.add_argument("--fecha-inicio", help="Inicio del periodo Atenea en formato YYYY-MM-DD.")
    parser.add_argument("--fecha-fin", help="Fin del periodo Atenea en formato YYYY-MM-DD.")
    return parser.parse_args()


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    args = _parse_args()
    if args.source == "atenea" and (not args.fecha_inicio or not args.fecha_fin):
        print("Para --source atenea debes indicar --fecha-inicio y --fecha-fin.")
        sys.exit(2)

    workspace = Path(__file__).resolve().parent
    cfg = PipelineConfig.default(workspace_dir=workspace)
    cfg = replace(
        cfg,
        input_source=args.source,
        atenea_fecha_inicio=args.fecha_inicio,
        atenea_fecha_fin=args.fecha_fin,
        atenea_export_excel=args.source == "atenea",
    )
    result = run_and_export(cfg)
    print(f"OK - control filas entrada/salida: {result.input_rows} -> {len(result.df_final_export)}")
    print(f"Archivo: {cfg.output_path}")


if __name__ == "__main__":
    main()
