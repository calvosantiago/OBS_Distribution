from __future__ import annotations

from pathlib import Path
import sys

from distribucion_obs import PipelineConfig, run_and_export


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    workspace = Path(__file__).resolve().parent
    cfg = PipelineConfig.default(workspace_dir=workspace)
    result = run_and_export(cfg)
    print(f"OK - control filas entrada/salida: {result.input_rows} -> {len(result.df_final_export)}")
    print(f"Archivo: {cfg.output_path}")


if __name__ == "__main__":
    main()
