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

_ATENEA_CUPONES_COLUMNS = {
    "ID de la Oportunidad": ["mcs_opportunityautonumber", "opportunityid"],
    "Fecha de creación": ["createdon"],
    "Programa de Interes": [
        "mcs_programidname",
        "program_interest2.mcs_marketingname",
        "pfu_marketingnamepvc",
        "program_main2.mcs_marketingname",
        "lead_origin2.mcs_programversionidname",
        "mcs_programmeversioncampusidname",
    ],
    "País": [
        "contact_parent2.sis_address1countryidname",
        "lead_origin2.mcs_address1countryidname",
    ],
    "País (Contacto) (Contacto)": ["contact_parent2.sis_address1countryidname"],
    "Country (Cliente potencial original) (Lead)": ["lead_origin2.mcs_address1countryidname"],
    "Pillar (Campaña de origen) (Campaña)": [
        "campaign_origin2.mcs_pillaridname",
        "campaign_origin2.mcs_pillarname",
    ],
    "Pillar Name (Campaña de origen) (Campaña)": ["campaign_origin2.mcs_pillarname"],
    "SubPillar (Campaña de origen) (Campaña)": ["campaign_origin2.mcs_subpillaridname"],
    "Pull/Push": ["mcs_pullpushname", "campaign_origin2.mcs_pullpushname"],
    "Propietario": ["owneridname"],
    "Tipo de Re-Apertura": ["mcs_reopeningtypename"],
    "Asesor Sugerido": ["mcs_suggestedadvisoridname"],
    "Tipo": ["mcs_typename"],
    "Re-Abierta": ["mcs_reopenedopportunityname"],
    "Propietario (Oportunidad de Origen) (Oportunidad)": ["opp_origin2.owneridname"],
    "Equipo Asignado": ["mcs_assignedteamidname"],
    "Calificación (Programa de Interes) (Programa)": ["program_interest2.sis_qualificationidname"],
    "Campaña de origen (Cliente potencial original) (Lead)": ["lead_origin2.campaignidname"],
    "Región": ["mcs_regionname"],
    "Edad (Cliente potencial original) (Lead)": ["lead_origin2.mcs_leadage"],
    " Nombre Completo (Contacto) (Contacto)": ["contact_parent2.fullname"],
    "Dirección : Estado (Contacto) (Contacto)": ["contact_parent2.sis_address1_stateorprovince"],
    "Email (Contacto) (Contacto)": ["contact_parent2.emailaddress1"],
    "Program Version of Interest (Cliente potencial original) (Lead)": ["lead_origin2.mcs_programversionidname"],
    "Oferta de la Versión del Programa": ["mcs_programmeversioncampusidname"],
    "Nombre Marketing (PVC)": ["pfu_marketingnamepvc"],
    "Teléfono (Cliente potencial) (Contacto)": ["contact_customer2.address1_telephone1"],
    "key_crmnet (Cliente potencial) (Contacto)": ["contact_customer2.sis_key_crmnet"],
    "key_Migration (Cliente potencial) (Contacto)": ["contact_customer2.sis_key_migration"],
    "Lead Duplicated": ["mcs_leadduplicatedname"],
}

_ATENEA_HIST_COLUMNS = {
    "ID de la Oportunidad": ["mcs_opportunityautonumber", "opportunityid"],
    "Tema": ["name"],
    "Nombre completo (Contacto) (Contacto)": ["contact_parent.sis_fullname"],
    "Programa de Interes": ["mcs_programidname", "program_interest.mcs_marketingname", "pfu_marketingnamepvc"],
    "Calificación (Programa de Interes) (Programa)": ["program_interest.sis_qualificationidname"],
    "Teléfono (Contacto) (Contacto)": ["contact_parent.address1_telephone1"],
    "Teléfono móvil (Contacto) (Contacto)": ["contact_parent.mobilephone"],
    "Edad (Cliente potencial original) (Lead)": ["lead_origin.mcs_leadage"],
    "País (Contacto) (Contacto)": ["contact_parent.sis_address1countryidname"],
    "Country (Cliente potencial original) (Lead)": ["lead_origin.mcs_address1countryidname"],
    "Estado": ["statecodename"],
    "Fase de canalización": ["stepname"],
    "Academic Period": ["mcs_academicperiodidname"],
    "Pull/Push": ["mcs_pullpushname", "campaign_origin.mcs_pullpushname"],
    "Pillar (Campaña de origen) (Campaña)": [
        "campaign_origin.mcs_pillaridname",
        "campaign_origin.mcs_pillarname",
    ],
    "Pillar Name (Campaña de origen) (Campaña)": ["campaign_origin.mcs_pillarname"],
    "SubPillar (Campaña de origen) (Campaña)": [
        "campaign_origin.mcs_subpillaridname",
        "campaign_origin.mcs_subpillarname",
    ],
    "Campaña de origen (Cliente potencial original) (Lead)": ["lead_origin.campaignidname"],
    "Fecha de creación": ["createdon"],
    "Fecha de Asignación": ["mcs_dateofassignment"],
    "Fecha Realización de Entrevista": ["mcs_interviewdonedate"],
    "Días desde la última comunicación": ["mcs_dayssincelastcommunication"],
    "Razón para el estado": ["statuscodename"],
    "Razón de Cierre": ["mcs_closurereasonidname"],
    "Sub-Razón de Cierre": ["mcs_closuresubreasonidname"],
    "Propietario": ["owneridname"],
    "Equipo de Ventas (Usuario propietario) (Usuario)": ["systemuser_owner.mcs_salesteamidname"],
    "Level of Study (Cliente potencial original) (Lead)": ["lead_origin.mcs_levelofstudyidname"],
    "Tipo": ["mcs_typename"],
    "Unidad de negocio propietaria": ["owningbusinessunitname"],
    "Programa Principal": ["program_main.mcs_marketingname"],
    "Usuario Propietario de la Oportunidad (Oportunidad de Origen) (Oportunidad)": [
        "opp_origin.mcs_owninguseridname"
    ],
    "Program Version of Interest (Cliente potencial original) (Lead)": ["lead_origin.mcs_programversionidname"],
    "Equipo Asignado": ["mcs_assignedteamidname"],
    "Usuario Propietario de la Oportunidad": ["mcs_owninguseridname"],
    "Fecha de localización": ["mcs_firstcontactdate"],
    "Creado en (Admisión) (Admisión)": ["admission.createdon"],
    "Ingresos reales": ["actualvalue"],
    "Porcentaje de Descuento (Cotización Ganada) (Oferta)": ["quote_won.mcs_discountpercentage"],
    "Importe detallado total (Cotización Ganada) (Oferta)": ["quote_won.totallineitemamount"],
    "Importe total (Cotización Ganada) (Oferta)": ["quote_won.totalamount"],
    "Monto de Inscripción (Cotización Ganada) (Oferta)": ["quote_won.mcs_inscriptionammount"],
    "Installments": ["mcs_installments"],
    "Cuotas (Cotización Ganada) (Oferta)": ["quote_won.mcs_installments"],
    "Número de Cuotas por Defecto (Payment term) (Plan de pago)": [
        "paymentterm.mcs_defaultnumberofinstallments"
    ],
    "Número de Cuotas por Defecto (Oferta de la Versión del Programa) (Versión del programa Campus)": [
        "pvc.mcs_defaultnumberofinstallments"
    ],
    "Sexo legal (Cliente potencial) (Contacto)": ["contact_customer.gendercodename"],
    "Fecha de cierre real": ["actualclosedate"],
    "Email (Contacto) (Contacto)": ["contact_parent.emailaddress1"],
    "Fecha de la Validación de la Documentación": ["mcs_documentationvalidateddate"],
    "Fecha de Entrevista Agendada": ["mcs_interviewscheduleddate"],
    "Localizado": ["mcs_unlocatedname"],
    "Oferta de la Versión del Programa": ["mcs_programmeversioncampusidname"],
    "Number of emails (Deprecated)": ["mcs_numberofemails"],
    "Number of phone calls (Deprecated)": ["mcs_numberofphonecalls"],
    "Insight Score": ["pfu_insightscore"],
    "Fecha del último contacto": ["mcs_lastcontactdate"],
    "Fecha de modificación": ["modifiedon"],
    "Tipo de Re-Apertura": ["mcs_reopeningtypename"],
    "Re-Abierta": ["mcs_reopenedopportunityname"],
    "Propietario (Oportunidad de Origen) (Oportunidad)": ["opp_origin.owneridname"],
    "key_crmnet (Cliente potencial) (Contacto)": ["contact_customer.sis_key_crmnet"],
    "Lead Duplicated": ["mcs_leadduplicatedname"],
    "Date de fixation d'entretien": ["mcs_interviewarrangementdate"],
    "Intentos de Localización": ["mcs_ilocalizedattempts"],
    "Fecha de No Localización": ["mcs_unreachabledate"],
    "Primer intento de llamada": ["mcs_firstcontactattemptname"],
    "Fecha del primer intento de llamada": ["mcs_firstcontactattemptdate"],
    "Fecha de primera llamada": ["mcs_firstcontactdate"],
    "Código de campaña (Campaña de origen) (Campaña)": ["campaign_origin.codename"],
    "Modality Code (Campaña de origen) (Campaña)": ["campaign_origin.mcs_modalitycode"],
    "Marketing Name (Programa de Interes) (Programa)": ["program_interest.mcs_marketingname"],
    "Marketing Name (Programa Principal) (Programa)": ["program_main.mcs_marketingname"],
    "Marketing Name (PV) (Program Version) (Versiones del Programa)": ["programversion.mcs_marketingnamepv"],
    "Nombre Marketing (PVC)": ["pfu_marketingnamepvc"],
    "Fecha de último intento de contacto": ["mcs_lastcontactattemptdateuserlocal"],
    "Campaña de origen": ["campaign_origin.codename"],
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


def _copy_first_available(df: pd.DataFrame, target: str, sources: list[str]) -> None:
    values = None
    for source in sources:
        if source not in df.columns:
            continue
        candidate = df[source]
        if values is None:
            values = candidate.copy()
        else:
            values = values.where(values.notna() & values.astype(str).str.strip().ne(""), candidate)
    if values is not None:
        df[target] = values


def normalize_atenea_inputs(
    df_cupones: pd.DataFrame,
    df_hist: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Adapta nombres técnicos de Dataverse al contrato histórico del pipeline."""
    df_cupones = df_cupones.copy()
    df_hist = df_hist.copy()

    for target, sources in _ATENEA_CUPONES_COLUMNS.items():
        _copy_first_available(df_cupones, target, sources)
    for target, sources in _ATENEA_HIST_COLUMNS.items():
        _copy_first_available(df_hist, target, sources)

    base_cols = [
        "ID de la Oportunidad",
        "Programa de Interes",
        "País",
        "Pillar (Campaña de origen) (Campaña)",
    ]
    ordered = [c for c in base_cols if c in df_cupones.columns]
    ordered += [c for c in df_cupones.columns if c not in ordered]
    df_cupones = df_cupones[ordered]

    return df_cupones, df_hist


def _sort_atenea_raw_by_createdon(df: pd.DataFrame) -> pd.DataFrame:
    if "createdon" not in df.columns:
        return df.copy()
    sort_df = df.copy()
    sort_key = pd.to_datetime(sort_df["createdon"], errors="coerce", utc=True)
    sort_df = sort_df.assign(_SORT_CREATEDON=sort_key)
    sort_df = sort_df.sort_values("_SORT_CREATEDON", kind="stable", na_position="last")
    return sort_df.drop(columns=["_SORT_CREATEDON"]).reset_index(drop=True)


def load_base_inputs(cfg: PipelineConfig) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if cfg.input_source == "atenea":
        return load_atenea_inputs(cfg)
    if cfg.input_source != "excel":
        raise ValueError(f"input_source no soportado: {cfg.input_source}")

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


def load_atenea_inputs(cfg: PipelineConfig) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if not cfg.atenea_fecha_inicio or not cfg.atenea_fecha_fin:
        raise ValueError("Para input_source='atenea' informa atenea_fecha_inicio y atenea_fecha_fin.")

    from extraccion_atenea import extraer_datos_atenea, format_op_no_asig_export, format_qbcn_export

    print("\n=== DATOS DE ATENEA EN USO ===")
    print(f"Periodo: {cfg.atenea_fecha_inicio} -> {cfg.atenea_fecha_fin}")

    df_hist, df_cupones = extraer_datos_atenea(
        cfg.atenea_fecha_inicio,
        cfg.atenea_fecha_fin,
        export_excel=cfg.atenea_export_excel,
        cache_file=str(cfg.atenea_cache_file or cfg.workspace_dir / "token_cache.bin"),
    )
    df_hist = _sort_atenea_raw_by_createdon(df_hist)
    df_cupones = _sort_atenea_raw_by_createdon(df_cupones)
    df_hist_export = format_qbcn_export(df_hist)
    df_cupones_export = format_op_no_asig_export(df_cupones)
    df_cupones, df_hist = normalize_atenea_inputs(df_cupones, df_hist)
    df_hist.attrs["atenea_export_sheet"] = df_hist_export
    df_cupones.attrs["atenea_export_sheet"] = df_cupones_export

    df_areas = pd.read_excel(cfg.areas_paises_path, sheet_name="Areas")
    df_paises = pd.read_excel(cfg.areas_paises_path, sheet_name="Paises")
    df_areas.columns = [c.strip() for c in df_areas.columns]
    df_paises.columns = [c.strip() for c in df_paises.columns]

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

