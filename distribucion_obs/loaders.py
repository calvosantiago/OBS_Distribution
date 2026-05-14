from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd

from .config import PipelineConfig


_DIC_INGLES_CUPONES = {
    "Opportunity Id": "ID de la Oportunidad",
    "Topic": "Tema",
    "Advised Program of interest from webform": "Programa de Interes",
    "Country (Originating Lead) (Lead)": "País",
    "Country (Contact) (Contact)": "País (Contacto) (Contacto)",
    "Pillar (Source Campaign) (Campaign)": "Pillar (Campaña de origen) (Campaña)",
    "SubPillar (Source Campaign) (Campaign)": "SubPillar (Campaña de origen) (Campaña)",
    "Owner": "Propietario",
    "Owner (Originating Opportunity) (Opportunity)": "Propietario (Oportunidad de Origen) (Oportunidad)",
    "Reopening type": "Tipo de Re-Apertura",
    "Created On": "Fecha de creación",
    "Email (Contact) (Contact)": "Email (Contacto) (Contacto)",
    "Address 1: Phone (Potential Customer) (Contact)": "Teléfono (Cliente potencial) (Contacto)",
    "Status Reason": "Razón para el estado",
}

_DIC_INGLES_HIST = {
    "Pillar (Source Campaign) (Campaign)": "Pillar (Campaña de origen) (Campaña)",
    "SubPillar (Source Campaign) (Campaign)": "SubPillar (Campaña de origen) (Campaña)",
    "SubPillar Name (Source Campaign) (Campaign)": "SubPillar Name (Campaña de origen) (Campaña)",
    "Assigned team": "Equipo Asignado",
    "Sales Team (Owning User) (User)": "Equipo de Ventas (Usuario propietario) (Usuario)",
    "Country (Contact) (Contact)": "País (Contacto) (Contacto)",
    "Advised Program of interest from webform": "Programa de Interes",
}


def _parse_dt_from_filename(stem: str) -> Optional[datetime]:
    parts = stem.split(" ")
    for i in range(len(parts)):
        block = " ".join(parts[i : i + 2])
        try:
            return datetime.strptime(block, "%d-%m-%Y %H-%M-%S")
        except ValueError:
            pass
        try:
            return datetime.strptime(parts[i], "%d-%m-%Y")
        except ValueError:
            pass
    return None


def _find_latest_file(downloads_dir: Path, prefix: str) -> Path:
    best_dt = None
    best_file: Optional[Path] = None
    for path in downloads_dir.rglob("*.xlsx"):
        name = path.name
        if "~" in name or not name.startswith(prefix):
            continue
        dt = _parse_dt_from_filename(path.stem)
        if dt is None:
            continue
        if best_dt is None or dt > best_dt:
            best_dt = dt
            best_file = path
    if best_file is None:
        raise FileNotFoundError(f"No se encontró archivo con prefijo: {prefix}")
    return best_file


def load_base_inputs(cfg: PipelineConfig) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    cupones_path = _find_latest_file(cfg.downloads_dir, "Oportunidades abiertas No Asignadas JE Totales")
    hist_path = _find_latest_file(cfg.downloads_dir, "qb_CN_V3_OBS")

    print("\n=== ARCHIVOS DE DESCARGAS EN USO ===")
    print(f"CUPONES: {cupones_path.name}")
    print(f"  Ruta: {cupones_path}")
    print(f"HISTORICO: {hist_path.name}")
    print(f"  Ruta: {hist_path}")

    df_cupones = pd.read_excel(cupones_path)
    df_hist = pd.read_excel(hist_path)
    df_areas = pd.read_excel(cfg.areas_paises_path, sheet_name="Areas")
    df_paises = pd.read_excel(cfg.areas_paises_path, sheet_name="Paises")

    if "Opportunity Id" in df_cupones.columns:
        df_cupones = df_cupones.rename(columns=_DIC_INGLES_CUPONES)
    if "País (Contacto) (Contacto)" in df_cupones.columns and "País" not in df_cupones.columns:
        df_cupones = df_cupones.rename(columns={"País (Contacto) (Contacto)": "País"})

    if "Assigned team" in df_hist.columns:
        df_hist = df_hist.rename(columns=_DIC_INGLES_HIST)

    df_areas.columns = [c.strip() for c in df_areas.columns]
    df_paises.columns = [c.strip() for c in df_paises.columns]

    base_cols = [
        "ID de la Oportunidad",
        "Programa de Interes",
        "País",
        "Pillar (Campaña de origen) (Campaña)",
    ]
    ordered = base_cols + [c for c in df_cupones.columns if c not in base_cols]
    df_cupones = df_cupones[ordered]

    return df_cupones, df_hist, df_areas, df_paises


def load_pilares_map(cfg: PipelineConfig) -> pd.DataFrame:
    return pd.read_excel(cfg.pilares_path, sheet_name="PULL-PUSH")


def load_sudoku_raw(cfg: PipelineConfig) -> pd.DataFrame:
    return pd.read_excel(
        cfg.sudoku_path,
        sheet_name="Estatus diario",
        usecols="P:V",
        skiprows=9,
        nrows=6,
        header=1,
    )

