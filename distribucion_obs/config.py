from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class PipelineConfig:
    workspace_dir: Path
    downloads_dir: Path
    areas_paises_path: Path
    sudoku_path: Path
    estructura_path: Path
    pilares_path: Path
    output_path: Path
    pilar_band_web: float = 0.05
    pilar_band_busc: float = 0.05
    time_limit: int = 90
    stage2_w_country: float = 1.0
    stage2_w_program: float = 0.9
    enforce_row_count_control: bool = True

    @staticmethod
    def default(workspace_dir: Path) -> "PipelineConfig":
        home = Path.home()
        return PipelineConfig(
            workspace_dir=workspace_dir,
            downloads_dir=home / "Downloads",
            areas_paises_path=workspace_dir / "Areas_Paises.xlsx",
            sudoku_path=workspace_dir / "SUDOKU.xlsx",
            estructura_path=home
            / "Grupo Planeta"
            / "BI POWER - General"
            / "PBI"
            / "OBS v2"
            / "0_TTAA"
            / "OBS_ESTRUCTURA_COMERCIAL.xlsx",
            pilares_path=home
            / "Grupo Planeta"
            / "BI POWER - General"
            / "PBI"
            / "OBS v2"
            / "0_TTAA"
            / "OBS_PULL_PUSH.xlsx",
            output_path=workspace_dir / "Distribucion_Final.xlsx",
        )
