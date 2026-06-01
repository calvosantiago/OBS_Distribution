from __future__ import annotations

from datetime import datetime
from zoneinfo import ZoneInfo

import country_converter as coco
import pandas as pd

from .config import PipelineConfig
from .extracted_functions import (
    AREA_A_TEAMS,
    AREA_B_TEAMS,
    AREA_C_TEAMS,
    CORE_TEAMS,
    _norm_email,
    _norm_phone,
    get_pesos_mensuales,
    get_pesos_por_area,
    limpiar_columnas_duplicadas,
    obtener_semana_comercial,
)


def _mask_pmax(df: pd.DataFrame) -> pd.Series:
    """Máscara de filas cuyo subpilar contenga 'pmax' (case-insensitive, coincidencia parcial)."""
    subpilar_cols = [
        "SubPillar (Campaña de origen) (Campaña)",
        "SubPillar Name (Campaña de origen) (Campaña)",
    ]
    if df.empty:
        return pd.Series(False, index=df.index)
    mask = pd.Series(False, index=df.index)
    for col in subpilar_cols:
        if col in df.columns:
            mask = mask | df[col].astype(str).str.contains("pmax", case=False, na=False)
    return mask


def _mask_google_search_argentina(df: pd.DataFrame) -> pd.Series:
    """Filas Google Search Argentina que se gestionan fuera de la distribución."""
    if df.empty:
        return pd.Series(False, index=df.index)

    subpilar_cols = [
        "SubPillar (Campaña de origen) (Campaña)",
        "SubPillar Name (Campaña de origen) (Campaña)",
    ]
    subpilar_mask = pd.Series(False, index=df.index)
    for col in subpilar_cols:
        if col in df.columns:
            subpilar_mask = subpilar_mask | df[col].astype(str).str.contains(
                "google search", case=False, na=False
            )

    pilar_mask = (
        df["PILAR_NORM"].astype(str).str.strip().str.upper().eq("BUSCADORES")
        if "PILAR_NORM" in df.columns
        else pd.Series(False, index=df.index)
    )

    pais_cols = [
        "País Normalizado",
        "País Corregido",
        "País",
        "País (Contacto) (Contacto)",
        "Country (Contact) (Contact)",
        "Country (Originating Lead) (Lead)",
    ]
    pais_mask = pd.Series(False, index=df.index)
    for col in pais_cols:
        if col in df.columns:
            pais_mask = pais_mask | df[col].astype(str).str.upper().str.contains(
                "ARGENT", na=False
            )

    return pilar_mask & subpilar_mask & pais_mask


def _map_equipo_e1_esp_to_b2(df: pd.DataFrame, equipo_cols: list[str]) -> pd.DataFrame:
    if df.empty or "IDIOMA" not in df.columns:
        return df
    mask_esp = df["IDIOMA"].astype(str).str.strip().str.upper().eq("ESP")
    for col in equipo_cols:
        if col in df.columns:
            mask_e1 = df[col].astype(str).str.strip().str.upper().eq("EQUIPO_E1")
            df.loc[mask_esp & mask_e1, col] = "Equipo_B2"
    return df


def _warn_missing_program_normalization(df: pd.DataFrame, source_name: str) -> None:
    if df.empty or "Programa de Interes" not in df.columns:
        return
    required_cols = {"AREA", "TIPO", "IDIOMA"}
    if not required_cols.issubset(df.columns):
        return

    missing_norm = df["AREA"].isna() | df["TIPO"].isna() | df["IDIOMA"].isna()
    if not missing_norm.any():
        return

    cols = ["Programa de Interes"]
    if "ID de la Oportunidad" in df.columns:
        cols.append("ID de la Oportunidad")
    sample = df.loc[missing_norm, cols].copy()
    sample["INDEX_ORIGINAL"] = sample.index
    summary = (
        sample.groupby("Programa de Interes", dropna=False)
        .agg(
            FILAS=("Programa de Interes", "size"),
            INDEX_ORIGINAL=("INDEX_ORIGINAL", lambda s: ", ".join(map(str, list(s)[:5]))),
        )
        .reset_index()
        .sort_values(["FILAS", "Programa de Interes"], ascending=[False, True])
    )
    if "ID de la Oportunidad" in sample.columns:
        ids = sample.groupby("Programa de Interes", dropna=False)["ID de la Oportunidad"].agg(
            lambda s: ", ".join(map(str, list(s)[:3]))
        )
        summary["IDS_EJEMPLO"] = summary["Programa de Interes"].map(ids)

    print(f"\n[ADVERTENCIA] Programas sin normalizar en Areas_Paises.xlsx ({source_name}):")
    print("Añádelos en la hoja Areas con Área, TIPO e IDIOMA para que no se descarten.")
    print(summary.to_string(index=False))


def _build_program_lookup(df_areas: pd.DataFrame) -> pd.DataFrame:
    alias_cols = ["PROGRAMA", "PROG", "PROG NORM", "NOMBRE CORTO"]
    frames = []
    for alias_col in alias_cols:
        if alias_col not in df_areas.columns:
            continue
        tmp = df_areas[[alias_col, "Área", "TIPO", "IDIOMA"]].copy()
        tmp = tmp.rename(columns={alias_col: "Programa de Interes"})
        tmp["Programa de Interes"] = tmp["Programa de Interes"].astype(str).str.strip().str.upper()
        tmp = tmp[tmp["Programa de Interes"].ne("") & tmp["Programa de Interes"].ne("NAN")]
        frames.append(tmp)
    if not frames:
        return pd.DataFrame(columns=["Programa de Interes", "Área", "TIPO", "IDIOMA"])
    return (
        pd.concat(frames, ignore_index=True)
        .drop_duplicates(subset=["Programa de Interes"], keep="first")
        [["Programa de Interes", "Área", "TIPO", "IDIOMA"]]
    )


def enrich_with_area_country_pillar(
    df_cupones: pd.DataFrame,
    df_hist: pd.DataFrame,
    df_areas: pd.DataFrame,
    df_paises: pd.DataFrame,
    df_pilares_norm: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    # Regla de negocio: PMAX no participa en distribución/cálculos.
    # En histórico sí se excluye de cálculos; en cupones de hoy se preserva para export final.
    df_cupones = df_cupones.copy()
    df_hist = df_hist.loc[~_mask_pmax(df_hist)].copy()

    df_cupones["Programa de Interes"] = df_cupones["Programa de Interes"].astype(str).str.strip().str.upper()
    df_hist["Programa de Interes"] = df_hist["Programa de Interes"].astype(str).str.strip().str.upper()
    merge_cols = _build_program_lookup(df_areas)
    df_cupones = df_cupones.merge(merge_cols, on="Programa de Interes", how="left").rename(columns={"Área": "AREA"})
    _warn_missing_program_normalization(df_cupones, "cupones abiertos")

    df_hist = df_hist.drop(columns=["TIPO", "IDIOMA", "AREA"], errors="ignore")
    df_hist = df_hist.merge(merge_cols, on="Programa de Interes", how="left").rename(columns={"Área": "AREA"})
    _warn_missing_program_normalization(df_hist, "histórico")

    def corregir_pais(pais: object) -> str:
        if not isinstance(pais, str) or pais.strip() == "":
            return ""
        corrections = {
            "PERÚ": "PERU",
            "PANAMÁ": "PANAMA",
            "MÉXICO": "MEXICO",
            "PHILIPINNES": "PHILIPPINES",
            "BOSTWANA": "BOTSWANA",
            "GUADALOUPE": "GUADELOUPE",
            "NETHERLANDS ANTILLES": "CURACAO",
            "NETHERLAND ANTILLES": "CURACAO",
            "ANTILLAS HOLANDESAS": "CURACAO",
            "UNITED STATES": "USA",
            "UNITED STATES OF AMERICA": "USA",
            "U.S.A.": "USA",
            "EEUU": "USA",
        }
        return corrections.get(pais.upper().strip(), pais.upper().strip())

    df_cupones["País Corregido"] = df_cupones["País"].apply(corregir_pais)
    df_hist["País Corregido"] = df_hist["País (Contacto) (Contacto)"].apply(corregir_pais)
    df_cupones["País Normalizado"] = coco.convert(names=df_cupones["País Corregido"], to="name_short", not_found=None)
    df_hist["País Normalizado"] = coco.convert(names=df_hist["País Corregido"], to="name_short", not_found=None)

    agrup_obs = {
        "ESPAÑA": "ES",
        "SPAIN": "ES",
        "CHILE": "CL",
        "COLOMBIA": "CO",
        "COSTA RICA": "CR",
        "ECUADOR": "EC",
        "MÉXICO": "MX",
        "MEXICO": "MX",
        "PERÚ": "PE",
        "PERU": "PE",
        "EL SALVADOR": "VIP",
        "PANAMÁ": "VIP",
        "PANAMA": "VIP",
        "PUERTO RICO": "VIP",
        "URUGUAY": "VIP",
        "HONDURAS": "VIP",
    }
    cont_map = dict(zip(df_paises["PAIS A NORMALIZAR"].astype(str).str.upper(), df_paises["AGRUP ENG"]))

    def clas_obs(x: object) -> str:
        y = x.upper() if isinstance(x, str) else ""
        return agrup_obs.get(y, "RI")

    def clas_cont(x: object) -> str:
        y = x.upper() if isinstance(x, str) else ""
        return cont_map.get(y, "Desconocido")

    df_cupones["Agrupación OBS"] = df_cupones["País Normalizado"].apply(clas_obs)
    df_cupones["Continente"] = df_cupones["País Normalizado"].apply(clas_cont)
    df_hist["Agrupación OBS"] = df_hist["País Normalizado"].apply(clas_obs)
    df_hist["Continente"] = df_hist["País Normalizado"].apply(clas_cont)

    mapa_pilares = dict(zip(df_pilares_norm["PILAR"], df_pilares_norm["PILAR PARA DISTRIBUCION"]))
    df_cupones["PILAR_NORM"] = df_cupones["Pillar (Campaña de origen) (Campaña)"].map(mapa_pilares)
    df_hist["PILAR_NORM"] = df_hist["Pillar (Campaña de origen) (Campaña)"].map(mapa_pilares)

    # Google Search Argentina se gestiona manualmente, igual que PMAX:
    # no participa en histórico/cadencia ni en distribución.
    mask_gs_arg_hist = _mask_google_search_argentina(df_hist)
    if mask_gs_arg_hist.any():
        print(
            "\n=== CONTROL GOOGLE SEARCH ARGENTINA ===\n"
            f"Histórico excluido de cálculo: {int(mask_gs_arg_hist.sum())}"
        )
        df_hist = df_hist.loc[~mask_gs_arg_hist].copy()
    return df_cupones, df_hist


def build_weights(df_sudoku_raw: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    df_pesos_actuales = get_pesos_mensuales(df_sudoku_raw)
    pesos_areas = get_pesos_por_area(df_pesos_actuales)
    df_horas_eq = df_pesos_actuales[["EQUIPO", "HORAS"]].copy()
    df_pesos_areas = pd.concat([df.assign(AREA=area) for area, df in pesos_areas.items()], ignore_index=True)
    df_pesos_areas = df_pesos_areas[["AREA", "EQUIPO", "PESO_BASE"]]
    return df_pesos_actuales, df_horas_eq, df_pesos_areas


def build_hist_qbcn(df_hist: pd.DataFrame, df_areas: pd.DataFrame) -> pd.DataFrame:
    df_hist_qbcn = df_hist.drop(columns=["TIPO", "IDIOMA", "AREA"], errors="ignore")
    df_hist_qbcn = df_hist_qbcn.merge(
        _build_program_lookup(df_areas),
        on="Programa de Interes",
        how="left",
    ).rename(columns={"Área": "AREA"})

    df_hist_qbcn = _map_equipo_e1_esp_to_b2(
        df_hist_qbcn,
        ["Equipo de Ventas (Usuario propietario) (Usuario)", "Equipo Asignado"],
    )

    equipos_area = {"A": AREA_A_TEAMS, "B": AREA_B_TEAMS, "C": AREA_C_TEAMS}
    equipos = sum(equipos_area.values(), []) + ["Equipo_E1"]

    filtro_base = (df_hist_qbcn["Equipo Asignado"] != "Equipo_Referidos") & (~df_hist_qbcn["PILAR_NORM"].isin(["REF/RECUP", "OTROS"]))
    filtro_mst_mba_esp = df_hist_qbcn["TIPO"].isin(["MST", "MBA"]) & (df_hist_qbcn["IDIOMA"] == "ESP")
    filtro_eng = df_hist_qbcn["IDIOMA"] == "ENG"
    filtro_equipo = df_hist_qbcn["Equipo de Ventas (Usuario propietario) (Usuario)"].isin(equipos)

    return df_hist_qbcn[filtro_base & filtro_equipo & (filtro_mst_mba_esp | filtro_eng)].copy()


def preprocess_open_coupons(df_cupones: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = df_cupones.copy()
    for col in ["TIPO", "IDIOMA", "AREA"]:
        df[col] = df[col].astype(str).str.strip().str.upper()
    df["INDEX_ORIGINAL"] = df.reset_index().index
    mask_e1_esp = (
        df["IDIOMA"].astype(str).str.strip().str.upper().eq("ESP")
        & df["Propietario"].astype(str).str.strip().str.upper().eq("EQUIPO_E1")
    )
    if mask_e1_esp.any():
        df.loc[mask_e1_esp, "EQUIPO_FINAL_EXPORT"] = df.loc[mask_e1_esp, "Propietario"]
    df = _map_equipo_e1_esp_to_b2(df, ["Propietario", "EQUIPO_FINAL"])

    # PMAX de hoy: no entran en distribución, pero se preservan en la salida final.
    mask_pmax_today = _mask_pmax(df)
    df_pmax = df.loc[mask_pmax_today].copy()
    if not df_pmax.empty:
        if "EQUIPO_FINAL" not in df_pmax.columns:
            df_pmax["EQUIPO_FINAL"] = df_pmax["Propietario"]
        else:
            df_pmax["EQUIPO_FINAL"] = df_pmax["EQUIPO_FINAL"].fillna(df_pmax["Propietario"])
        df_pmax["SPECIAL_KIND"] = "PMAX_PASSTHROUGH"
    df = df.loc[~mask_pmax_today].copy()

    # Google Search Argentina de hoy: misma regla que PMAX.
    mask_gs_arg_today = _mask_google_search_argentina(df)
    df_gs_arg = df.loc[mask_gs_arg_today].copy()
    if not df_gs_arg.empty:
        if "EQUIPO_FINAL" not in df_gs_arg.columns:
            df_gs_arg["EQUIPO_FINAL"] = df_gs_arg["Propietario"]
        else:
            df_gs_arg["EQUIPO_FINAL"] = df_gs_arg["EQUIPO_FINAL"].fillna(df_gs_arg["Propietario"])
        df_gs_arg["SPECIAL_KIND"] = "GOOGLE_SEARCH_ARG_PASSTHROUGH"
    df = df.loc[~mask_gs_arg_today].copy()

    # REF/RECUP y OTROS de hoy: igual que PMAX, no entran en distribución ni en el
    # cálculo de cadencia. Se preservan en la salida final con su propietario original.
    _PILARES_EXCLUIDOS = {"REF/RECUP", "OTROS"}
    if "PILAR_NORM" in df.columns:
        mask_recup_today = df["PILAR_NORM"].isin(_PILARES_EXCLUIDOS)
    else:
        mask_recup_today = pd.Series(False, index=df.index)
    df_recup = df.loc[mask_recup_today].copy()
    if not df_recup.empty:
        if "EQUIPO_FINAL" not in df_recup.columns:
            df_recup["EQUIPO_FINAL"] = df_recup["Propietario"]
        else:
            df_recup["EQUIPO_FINAL"] = df_recup["EQUIPO_FINAL"].fillna(df_recup["Propietario"])
        df_recup["SPECIAL_KIND"] = "RECUP_PASSTHROUGH"
    df = df.loc[~mask_recup_today].copy()

    email_col = "Email (Contacto) (Contacto)"
    phone_col = "Teléfono (Cliente potencial) (Contacto)"
    df["_EMAIL_N"] = df[email_col].map(_norm_email) if email_col in df.columns else pd.NA
    df["_PHONE_N"] = df[phone_col].map(_norm_phone) if phone_col in df.columns else pd.NA

    dup_email = df["_EMAIL_N"].duplicated(keep="first") & df["_EMAIL_N"].notna()
    dup_phone = df["_PHONE_N"].duplicated(keep="first") & df["_PHONE_N"].notna()
    mask_dup_any = dup_email | dup_phone
    df.loc[mask_dup_any, "Propietario"] = "Equipo_Z"
    if "EQUIPO_FINAL" in df.columns:
        df.loc[mask_dup_any, "EQUIPO_FINAL"] = "Joaquim Barnola Fontrodona"

    special_set = {"EQUIPO_REFERIDOS", "EQUIPO_Z"}
    prop_upper = df["Propietario"].astype(str).str.strip().str.upper()
    mask_special = prop_upper.isin(special_set)
    df_special = df.loc[mask_special].copy()

    if "EQUIPO_FINAL" not in df_special.columns:
        df_special["EQUIPO_FINAL"] = pd.NA
    if "SPECIAL_KIND" not in df_special.columns:
        df_special["SPECIAL_KIND"] = pd.NA

    prop_upper_sp = df_special["Propietario"].astype(str).str.strip().str.upper()
    mask_ref = prop_upper_sp.eq("EQUIPO_REFERIDOS")
    mask_z = prop_upper_sp.eq("EQUIPO_Z")
    mask_dup_in_special = mask_dup_any.reindex(df_special.index, fill_value=False)

    sel_dup_z = mask_z & mask_dup_in_special
    df_special.loc[sel_dup_z, "EQUIPO_FINAL"] = "Joaquim Barnola Fontrodona"
    df_special.loc[sel_dup_z, "SPECIAL_KIND"] = "DUPLICADO->JOAQUIM"

    sel_ref_sin_final = mask_ref & (df_special["EQUIPO_FINAL"].isna() | (df_special["EQUIPO_FINAL"].astype(str).str.strip() == ""))
    df_special.loc[sel_ref_sin_final, "EQUIPO_FINAL"] = "Equipo_Referidos"
    df_special.loc[mask_ref, "SPECIAL_KIND"] = "REFERIDOS"

    sel_z_no_dup = mask_z & (~mask_dup_in_special)
    sel_z_no_dup_sin_final = sel_z_no_dup & (df_special["EQUIPO_FINAL"].isna() | (df_special["EQUIPO_FINAL"].astype(str).str.strip() == ""))
    df_special.loc[sel_z_no_dup_sin_final, "EQUIPO_FINAL"] = "Equipo_Z"
    df_special.loc[sel_z_no_dup, "SPECIAL_KIND"] = "Z_MANUAL"

    if not df_special.empty:
        df.loc[df_special.index, ["EQUIPO_FINAL", "SPECIAL_KIND"]] = df_special[["EQUIPO_FINAL", "SPECIAL_KIND"]]

    equipos_area = {"A": AREA_A_TEAMS, "B": AREA_B_TEAMS, "C": AREA_C_TEAMS, "E": ["Equipo_E1"]}
    equipos_upper = [e.upper() for e in sum(equipos_area.values(), [])]
    mask_equipo = prop_upper.isin(equipos_upper)
    mask_mst_esp = df["TIPO"].eq("MST") & df["IDIOMA"].eq("ESP")
    mask_mba_esp = df["TIPO"].eq("MBA") & df["IDIOMA"].eq("ESP")
    mask_eng = df["IDIOMA"].eq("ENG")
    mask_valid_core = mask_equipo & (mask_mst_esp | mask_mba_esp | mask_eng)
    df = df[mask_valid_core].reset_index(drop=True)
    if not df_pmax.empty:
        # Se agregan a especiales para preservarlos en el resultado final, sin redistribución.
        df_special = pd.concat([df_special, df_pmax], ignore_index=False)
    if not df_gs_arg.empty:
        # Se agregan a especiales para preservarlos en el resultado final, sin redistribución.
        df_special = pd.concat([df_special, df_gs_arg], ignore_index=False)
    if not df_recup.empty:
        # REF/RECUP y OTROS: passthrough idéntico al PMAX.
        df_special = pd.concat([df_special, df_recup], ignore_index=False)
    return df, df_special


def split_reap_fresh_hist(
    df_cupones_open: pd.DataFrame,
    df_hist_qbcn: pd.DataFrame,
    cfg: PipelineConfig,
    cutoff_dt: datetime | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    df_cupones_open = df_cupones_open.copy()
    df_cupones_open["TIPO_REPARTO"] = df_cupones_open["Tipo de Re-Apertura"].apply(
        lambda x: "REAP" if pd.notna(x) and str(x).strip() != "" else "FRESH"
    )

    df_estructura = pd.read_excel(cfg.estructura_path, sheet_name="ESTRUCTURA")
    df_norm = pd.read_excel(cfg.estructura_path, sheet_name="NORMALIZACION NOMBRES")
    df_estructura.columns = [c.strip() for c in df_estructura.columns]
    df_norm.columns = [c.strip() for c in df_norm.columns]

    dic_normalizacion = dict(zip(df_norm["NOMBRE A NORMALIZAR"], df_norm["NOMBRE CORTO"]))
    col_origen = "Propietario (Oportunidad de Origen) (Oportunidad)"
    df_cupones_open["NOMBRE_NORMALIZADO"] = df_cupones_open[col_origen].map(dic_normalizacion)

    semana_actual = obtener_semana_comercial(datetime.today(), calendario_path=cfg.calendario_path)
    df_ac = df_estructura[
        (df_estructura["ACTIVO"] == "ACTIVO")
        & (df_estructura["ROL"] == "AC")
        & (df_estructura["AÑO-MES-SEM"] == semana_actual)
    ][["NOMBRE CORTO", "EQUIPO HISTORICO AC"]].copy()
    df_je = df_estructura[(df_estructura["ROL"] == "JE") & (df_estructura["AÑO-MES-SEM"] == semana_actual)][
        ["NOMBRE CORTO", "EQUIPO HISTORICO AC"]
    ].copy()
    df_activos = pd.concat([df_ac, df_je], ignore_index=True)
    df_activos["EQUIPO_REAP"] = df_activos["EQUIPO HISTORICO AC"].str.strip().apply(lambda x: f"Equipo_{x}")
    df_activos = df_activos.rename(columns={"NOMBRE CORTO": "NOMBRE_NORMALIZADO"})

    df_reap = df_cupones_open[df_cupones_open["TIPO_REPARTO"] == "REAP"].copy().merge(df_activos, on="NOMBRE_NORMALIZADO", how="left")
    df_reap_validas = df_reap[df_reap["EQUIPO_REAP"].notna()].copy()
    df_reap_invalidas = df_reap[df_reap["EQUIPO_REAP"].isna()].copy()
    df_reap_validas["EQUIPO_FINAL"] = df_reap_validas["EQUIPO_REAP"]
    df_reap_validas["EQUIPO_FINAL_EXPORT"] = df_reap_validas["EQUIPO_REAP"]
    df_reap_validas = _map_equipo_e1_esp_to_b2(df_reap_validas, ["EQUIPO_REAP", "EQUIPO_FINAL"])
    df_reap_invalidas["TIPO_REPARTO"] = "FRESH"

    df_cupones_open = limpiar_columnas_duplicadas(df_cupones_open)
    df_reap_invalidas = limpiar_columnas_duplicadas(df_reap_invalidas)

    if cutoff_dt is None:
        cutoff_raw = pd.read_excel(cfg.sudoku_path, sheet_name="Estatus diario", header=None, usecols="O", skiprows=21, nrows=1).iat[0, 0]
        cutoff_dt = pd.to_datetime(cutoff_raw, errors="coerce", dayfirst=True)
        if pd.isna(cutoff_dt):
            raise ValueError("No se pudo leer hora de corte desde SUDOKU (Estatus diario!O22).")

    def _parse_created_dates(values: pd.Series | None) -> pd.Series:
        if values is None:
            return pd.Series(pd.NaT, index=df_cupones_open.index)
        if cfg.input_source == "atenea":
            return (
                pd.to_datetime(values, errors="coerce", utc=True)
                .dt.tz_convert(ZoneInfo("Europe/Madrid"))
                .dt.tz_localize(None)
            )
        return pd.to_datetime(values, errors="coerce", dayfirst=True)

    fechas_open = _parse_created_dates(df_cupones_open.get("Fecha de creación"))
    mask_cutoff = df_cupones_open["TIPO_REPARTO"].eq("FRESH") & fechas_open.notna() & (fechas_open < cutoff_dt)
    df_corte = df_cupones_open.loc[mask_cutoff].copy()
    if not df_corte.empty:
        if "EQUIPO_FINAL" not in df_corte.columns:
            df_corte["EQUIPO_FINAL"] = df_corte["Propietario"]
        else:
            df_corte["EQUIPO_FINAL"] = df_corte["EQUIPO_FINAL"].fillna(df_corte["Propietario"])
        df_corte["REAP_REASON"] = "CUTOFF"
        df_corte["TIPO_REPARTO"] = "REAP"
        df_cupones_open.loc[mask_cutoff, "TIPO_REPARTO"] = "REAP"

    if (not df_reap_invalidas.empty) and ("Fecha de creación" in df_reap_invalidas.columns):
        if cfg.input_source == "atenea":
            fechas_inv = (
                pd.to_datetime(df_reap_invalidas["Fecha de creación"], errors="coerce", utc=True)
                .dt.tz_convert(ZoneInfo("Europe/Madrid"))
                .dt.tz_localize(None)
            )
        else:
            fechas_inv = pd.to_datetime(df_reap_invalidas["Fecha de creación"], errors="coerce", dayfirst=True)
        mask_cutoff_inv = fechas_inv.notna() & (fechas_inv < cutoff_dt)
        if mask_cutoff_inv.any():
            df_corte_inv = df_reap_invalidas.loc[mask_cutoff_inv].copy()
            if "EQUIPO_FINAL" not in df_corte_inv.columns:
                df_corte_inv["EQUIPO_FINAL"] = df_corte_inv["Propietario"]
            else:
                df_corte_inv["EQUIPO_FINAL"] = df_corte_inv["EQUIPO_FINAL"].fillna(df_corte_inv["Propietario"])
            df_corte_inv["REAP_REASON"] = "CUTOFF"
            df_corte_inv["TIPO_REPARTO"] = "REAP"
            df_corte = pd.concat([df_corte, df_corte_inv], ignore_index=True)
            df_reap_invalidas = df_reap_invalidas.loc[~mask_cutoff_inv].copy()

    df_fresh = pd.concat([df_cupones_open[df_cupones_open["TIPO_REPARTO"] == "FRESH"], df_reap_invalidas], ignore_index=True)
    df_fresh = df_fresh[df_fresh["AREA"].isin(["A", "B", "C", "T", "E"])].copy()
    df_fresh = df_fresh.reset_index(drop=True)

    df_hist_qbcn = df_hist_qbcn.copy()
    df_hist_qbcn["EQUIPO_FINAL"] = df_hist_qbcn["Equipo de Ventas (Usuario propietario) (Usuario)"]

    df_hist_total = pd.concat([df_hist_qbcn, df_reap_validas, df_corte], ignore_index=True)
    df_hist_total = df_hist_total[df_hist_total["EQUIPO_FINAL"].notna()].copy()
    return df_fresh, df_reap_validas, df_corte, df_hist_total


def compute_cadencia_preliminar(
    df_hist_total: pd.DataFrame,
    df_fresh: pd.DataFrame,
    df_horas_eq: pd.DataFrame,
) -> tuple[dict[str, float], dict[str, float], dict[str, float]]:
    equipos = CORE_TEAMS
    df_hist_base = df_hist_total[df_hist_total["EQUIPO_FINAL"].isin(equipos)].copy()
    hist_a = (
        df_hist_base[(df_hist_base["TIPO"] == "MST") & (df_hist_base["IDIOMA"] == "ESP")]
        .groupby("EQUIPO_FINAL")
        .size()
        .reindex(equipos, fill_value=0)
        .rename("cupones_A")
    )
    hist_t = (
        df_hist_base[(df_hist_base["TIPO"] == "MBA") & (df_hist_base["IDIOMA"] == "ESP")]
        .groupby("EQUIPO_FINAL")
        .size()
        .reindex(equipos, fill_value=0)
        .rename("cupones_T")
    )
    fresh_a = (
        df_fresh[(df_fresh["TIPO"] == "MST") & (df_fresh["IDIOMA"] == "ESP")]
        .groupby("Propietario")
        .size()
        .reindex(equipos, fill_value=0)
        .rename("fresh_A")
    )
    fresh_t = (
        df_fresh[(df_fresh["TIPO"] == "MBA") & (df_fresh["IDIOMA"] == "ESP")]
        .groupby("Propietario")
        .size()
        .reindex(equipos, fill_value=0)
        .rename("fresh_T")
    )
    df_counts = pd.concat([hist_a, hist_t, fresh_a, fresh_t], axis=1).fillna(0)
    df_counts["CUPONES_PRELIM_A"] = df_counts["cupones_A"] + df_counts["fresh_A"]
    df_counts["CUPONES_PRELIM_T"] = df_counts["cupones_T"] + df_counts["fresh_T"]

    df_cad = (
        df_horas_eq[df_horas_eq["EQUIPO"].isin(equipos)][["EQUIPO", "HORAS"]]
        .merge(df_counts[["CUPONES_PRELIM_A", "CUPONES_PRELIM_T"]], left_on="EQUIPO", right_index=True, how="left")
        .fillna(0)
    )
    df_cad["CAD_PRELIM_A"] = df_cad.apply(lambda r: r["CUPONES_PRELIM_A"] / (r["HORAS"] / 6) if r["HORAS"] > 0 else 0, axis=1)
    df_cad["CAD_PRELIM_T"] = df_cad.apply(lambda r: r["CUPONES_PRELIM_T"] / (r["HORAS"] / 6) if r["HORAS"] > 0 else 0, axis=1)

    cad_prelim_a = df_cad.set_index("EQUIPO")["CAD_PRELIM_A"].to_dict()
    cad_prelim_t = df_cad.set_index("EQUIPO")["CAD_PRELIM_T"].to_dict()

    # Helper: cadencia teórica por área filtrando por nombre de equipo explícito,
    # sin depender del orden de filas de df_cad (que viene del orden de columnas del SUDOKU).
    def _cad_teo_area(eq_list: list, col: str) -> float:
        rows = df_cad[df_cad["EQUIPO"].isin(eq_list)]
        h = float(rows["HORAS"].sum())
        return float(rows[col].sum()) / (h / 6) if h > 0 else 0.0

    cad_teo = {
        "A": _cad_teo_area(AREA_A_TEAMS, "CUPONES_PRELIM_A"),
        "B": _cad_teo_area(AREA_B_TEAMS, "CUPONES_PRELIM_A"),
        "C": _cad_teo_area(AREA_C_TEAMS, "CUPONES_PRELIM_A"),
        "T": _cad_teo_area(equipos,                    "CUPONES_PRELIM_T"),
    }
    cad_teo["E"] = cad_teo["T"]
    return cad_prelim_a, cad_prelim_t, cad_teo
