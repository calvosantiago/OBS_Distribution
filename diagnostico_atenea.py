from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from distribucion_obs.loaders import normalize_atenea_inputs
from extraccion_atenea import extraer_datos_atenea


REQUIRED_CUPONES = [
    "ID de la Oportunidad",
    "Programa de Interes",
    "País",
    "Pillar (Campaña de origen) (Campaña)",
    "SubPillar (Campaña de origen) (Campaña)",
    "Propietario",
    "Tipo de Re-Apertura",
    "Propietario (Oportunidad de Origen) (Oportunidad)",
    "Equipo Asignado",
    "Email (Contacto) (Contacto)",
    "Teléfono (Cliente potencial) (Contacto)",
]

REQUIRED_HIST = [
    "ID de la Oportunidad",
    "Programa de Interes",
    "País (Contacto) (Contacto)",
    "Pillar (Campaña de origen) (Campaña)",
    "SubPillar (Campaña de origen) (Campaña)",
    "Equipo Asignado",
    "Equipo de Ventas (Usuario propietario) (Usuario)",
    "Propietario",
    "Fecha de creación",
    "Tipo de Re-Apertura",
    "Propietario (Oportunidad de Origen) (Oportunidad)",
]

KEY_CUPONES = [
    "ID de la Oportunidad",
    "Programa de Interes",
    "País",
    "Pillar (Campaña de origen) (Campaña)",
    "Propietario",
    "Equipo Asignado",
]

KEY_HIST = [
    "ID de la Oportunidad",
    "Programa de Interes",
    "País (Contacto) (Contacto)",
    "Pillar (Campaña de origen) (Campaña)",
    "Equipo Asignado",
    "Equipo de Ventas (Usuario propietario) (Usuario)",
]

DEBUG_CUPONES_PROGRAM = [
    "ID de la Oportunidad",
    "Programa de Interes",
    "program_interest2.mcs_marketingname",
    "program_interest2.sis_qualificationidname",
    "mcs_programidname",
    "_mcs_programid_value",
    "program_main2.mcs_marketingname",
    "lead_origin2.mcs_programversionidname",
    "mcs_programmeversioncampusidname",
    "pfu_marketingnamepvc",
    "Propietario",
    "Equipo Asignado",
]

DEBUG_HIST_PROGRAM = [
    "ID de la Oportunidad",
    "Programa de Interes",
    "mcs_programidname",
    "_mcs_programid_value",
    "program_interest.mcs_marketingname",
    "pfu_marketingnamepvc",
    "mcs_programmeversioncampusidname",
    "program_main.mcs_marketingname",
    "Propietario",
    "Equipo Asignado",
]


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Diagnostica columnas Atenea vs contrato del pipeline.")
    parser.add_argument("--fecha-inicio", help="Inicio del periodo Atenea en formato YYYY-MM-DD.")
    parser.add_argument("--fecha-fin", help="Fin del periodo Atenea en formato YYYY-MM-DD.")
    parser.add_argument(
        "--excel",
        type=Path,
        help="Excel leads_*.xlsx ya exportado por extraccion_atenea.py; evita llamar a Dataverse.",
    )
    return parser.parse_args()


def _load_from_excel(path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    xls = pd.ExcelFile(path)
    hist_sheet = next((s for s in xls.sheet_names if s.startswith("qb_cn_")), None)
    cupones_sheet = next((s for s in xls.sheet_names if s.startswith("op_no_asig_")), None)
    if hist_sheet is None or cupones_sheet is None:
        raise ValueError(
            f"No encuentro hojas qb_cn_* y op_no_asig_* en {path}. Hojas: {xls.sheet_names}"
        )
    df_hist = pd.read_excel(path, sheet_name=hist_sheet)
    df_cupones = pd.read_excel(path, sheet_name=cupones_sheet)
    return df_hist, df_cupones


def _missing(df: pd.DataFrame, required: list[str]) -> list[str]:
    return [col for col in required if col not in df.columns]


def _empty_counts(df: pd.DataFrame, columns: list[str]) -> pd.Series:
    counts = {}
    for col in columns:
        if col not in df.columns:
            counts[col] = "NO_EXISTE"
            continue
        s = df[col]
        counts[col] = int(s.isna().sum() + s.astype(str).str.strip().eq("").sum())
    return pd.Series(counts)


def _print_block(title: str, value: object) -> None:
    print("\n" + "=" * 80)
    print(title)
    print("=" * 80)
    print(value)


def _blank_mask(df: pd.DataFrame, column: str) -> pd.Series:
    if column not in df.columns:
        return pd.Series(True, index=df.index)
    return df[column].isna() | df[column].astype(str).str.strip().eq("")


def _sample_missing_program(df: pd.DataFrame, debug_cols: list[str]) -> str:
    mask = _blank_mask(df, "Programa de Interes")
    sample = df.loc[mask, [c for c in debug_cols if c in df.columns]]
    if sample.empty:
        return "sin ejemplos"
    return sample.to_string(index=False)


def main() -> None:
    args = _parse_args()
    if args.excel:
        df_hist_raw, df_cupones_raw = _load_from_excel(args.excel)
        source = f"Excel local: {args.excel}"
    else:
        if not args.fecha_inicio or not args.fecha_fin:
            raise SystemExit("Indica --excel o bien --fecha-inicio y --fecha-fin.")
        df_hist_raw, df_cupones_raw = extraer_datos_atenea(
            args.fecha_inicio,
            args.fecha_fin,
            export_excel=False,
        )
        source = f"Atenea: {args.fecha_inicio} -> {args.fecha_fin}"

    df_cupones, df_hist = normalize_atenea_inputs(df_cupones_raw, df_hist_raw)

    _print_block("ORIGEN", source)
    _print_block("FORMAS CRUDAS", f"qb_cn={df_hist_raw.shape} | op_no_asig={df_cupones_raw.shape}")
    _print_block("FORMAS NORMALIZADAS", f"hist={df_hist.shape} | cupones={df_cupones.shape}")

    _print_block("FALTAN EN CUPONES/op_no_asig", _missing(df_cupones, REQUIRED_CUPONES) or "ninguna")
    _print_block("FALTAN EN HIST/qb_cn", _missing(df_hist, REQUIRED_HIST) or "ninguna")

    _print_block("NULOS CLAVE CUPONES/op_no_asig", _empty_counts(df_cupones, KEY_CUPONES).to_string())
    _print_block("NULOS CLAVE HIST/qb_cn", _empty_counts(df_hist, KEY_HIST).to_string())
    _print_block(
        "EJEMPLOS SIN PROGRAMA CUPONES/op_no_asig",
        _sample_missing_program(df_cupones, DEBUG_CUPONES_PROGRAM),
    )
    _print_block(
        "EJEMPLOS SIN PROGRAMA HIST/qb_cn",
        _sample_missing_program(df_hist, DEBUG_HIST_PROGRAM),
    )

    _print_block("COLUMNAS CRUDAS CUPONES/op_no_asig", "\n".join(df_cupones_raw.columns))
    _print_block("COLUMNAS CRUDAS HIST/qb_cn", "\n".join(df_hist_raw.columns))


if __name__ == "__main__":
    main()
