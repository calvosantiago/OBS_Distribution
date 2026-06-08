from __future__ import annotations

import argparse
import os
import time
from datetime import datetime
from pathlib import Path
from typing import Any

import msal
import pandas as pd
import requests

from extraccion_atenea import AUTHORITY, CACHE_FILE, CLIENT_ID, ORG_URL, SCOPES


REQUEST_SETTINGS = {
    "max_retries": 5,
    "retry_wait": 30.0,
}
REQUEST_RETRY_COUNT = 0


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Prueba/asignacion masiva de ownerid en Atenea/Dataverse desde Distribucion_Final.xlsx."
    )
    parser.add_argument(
        "--id",
        help="ID de la Oportunidad visible (2021-...) u opportunityid GUID.",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="Procesa todas las oportunidades del Excel. Sin --apply solo hace dry-run masivo.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        help="Limita el numero de filas a procesar en modo --all, util para pruebas.",
    )
    parser.add_argument(
        "--excel",
        default="Distribucion_Final.xlsx",
        help="Excel de salida de distribucion. Por defecto: Distribucion_Final.xlsx",
    )
    parser.add_argument(
        "--owner",
        help="Owner destino manual. Solo se permite con --id. Si no se informa, usa EQUIPO_FINAL del Excel.",
    )
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Aplica el cambio real en Atenea. Sin este flag solo muestra el dry-run.",
    )
    parser.add_argument(
        "--force-owner-mismatch",
        action="store_true",
        help="Permite aplicar aunque el owner actual en Atenea no coincida con el owner original del Excel.",
    )
    parser.add_argument(
        "--cache-file",
        default=CACHE_FILE,
        help="Ruta de token_cache.bin. Por defecto usa el del workspace.",
    )
    parser.add_argument(
        "--sleep",
        type=float,
        default=0.35,
        help="Pausa en segundos entre oportunidades en modo --all. Por defecto: 0.35.",
    )
    parser.add_argument(
        "--max-retries",
        type=int,
        default=5,
        help="Reintentos por peticion HTTP si Dataverse responde 429/5xx. Por defecto: 5.",
    )
    parser.add_argument(
        "--retry-wait",
        type=float,
        default=30.0,
        help="Espera base en segundos para 429 si Dataverse no informa Retry-After. Por defecto: 30.",
    )
    args = parser.parse_args()
    if bool(args.id) == bool(args.all):
        parser.error("Indica una sola opcion: --id para una oportunidad o --all para masivo.")
    if args.all and args.owner:
        parser.error("--owner no se permite con --all; en masivo se usa EQUIPO_FINAL de cada fila.")
    if args.limit is not None and args.limit <= 0:
        parser.error("--limit debe ser mayor que cero.")
    if args.sleep < 0:
        parser.error("--sleep no puede ser negativo.")
    if args.max_retries < 0:
        parser.error("--max-retries no puede ser negativo.")
    if args.retry_wait <= 0:
        parser.error("--retry-wait debe ser mayor que cero.")
    return args


def _clean_value(value: Any) -> str:
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def _row_ref(row: pd.Series) -> str:
    for col in ["ID de la Oportunidad", "opportunityid", "mcs_opportunityautonumber"]:
        value = _clean_value(row.get(col, ""))
        if value:
            return value
    return ""


def _escape_odata(value: str) -> str:
    return value.replace("'", "''")


def _get_token(cache_file: str) -> str:
    cache = msal.SerializableTokenCache()
    if os.path.exists(cache_file):
        cache.deserialize(Path(cache_file).read_text(encoding="utf-8"))

    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache,
    )
    accounts = app.get_accounts()
    token_result = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None
    if not token_result:
        print("Abriendo navegador para autenticacion...")
        token_result = app.acquire_token_interactive(scopes=SCOPES)

    if cache.has_state_changed:
        Path(cache_file).write_text(cache.serialize(), encoding="utf-8")

    if "access_token" not in token_result:
        raise RuntimeError(token_result.get("error_description", "Error de autenticacion desconocido"))
    return token_result["access_token"]


def _headers(token: str) -> dict[str, str]:
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "Prefer": 'odata.include-annotations="*"',
    }


def _retry_after_seconds(response: requests.Response, attempt: int) -> float:
    retry_after = response.headers.get("Retry-After")
    if retry_after:
        try:
            return max(float(retry_after), 1.0)
        except ValueError:
            pass
    return max(REQUEST_SETTINGS["retry_wait"] * attempt, 1.0)


def _request_with_retry(method: str, url: str, **kwargs: Any) -> requests.Response:
    global REQUEST_RETRY_COUNT
    max_retries = int(REQUEST_SETTINGS["max_retries"])
    for attempt in range(1, max_retries + 2):
        response = requests.request(method, url, **kwargs)
        retryable = response.status_code == 429 or 500 <= response.status_code <= 599
        if not retryable or attempt > max_retries:
            return response

        REQUEST_RETRY_COUNT += 1
        wait_seconds = _retry_after_seconds(response, attempt)
        print(
            f"[RETRY] {method} {response.status_code}. "
            f"Intento {attempt}/{max_retries}. Esperando {wait_seconds:.0f}s..."
        )
        time.sleep(wait_seconds)
    return response


def _get_json(token: str, path: str, params: dict[str, str] | None = None) -> dict[str, Any]:
    response = _request_with_retry(
        "GET",
        f"{ORG_URL}/api/data/v9.2/{path}",
        headers=_headers(token),
        params=params,
    )
    if response.status_code != 200:
        raise RuntimeError(f"GET {path} fallo {response.status_code}: {response.text}")
    return response.json()


def _load_target_row(excel_path: Path, opportunity_ref: str) -> pd.Series:
    df = pd.read_excel(excel_path)
    ref = opportunity_ref.strip()
    masks = []
    for col in ["ID de la Oportunidad", "opportunityid", "mcs_opportunityautonumber"]:
        if col in df.columns:
            masks.append(df[col].astype(str).str.strip().eq(ref))
    if not masks:
        raise ValueError("El Excel no contiene columnas para identificar oportunidad.")

    mask = masks[0]
    for extra in masks[1:]:
        mask = mask | extra
    matches = df.loc[mask].copy()
    if matches.empty:
        raise ValueError(f"No se encontro la oportunidad {ref} en {excel_path}.")
    if len(matches) > 1:
        raise ValueError(f"La oportunidad {ref} aparece {len(matches)} veces en {excel_path}.")
    return matches.iloc[0]


def _load_all_rows(excel_path: Path, limit: int | None = None) -> pd.DataFrame:
    df = pd.read_excel(excel_path)
    if "EQUIPO_FINAL" not in df.columns:
        raise ValueError("El Excel no contiene EQUIPO_FINAL.")
    id_cols = [
        col
        for col in ["ID de la Oportunidad", "opportunityid", "mcs_opportunityautonumber"]
        if col in df.columns
    ]
    if not id_cols:
        raise ValueError("El Excel no contiene columnas para identificar oportunidad.")

    refs = df.apply(_row_ref, axis=1)
    missing = refs.eq("")
    if missing.any():
        raise ValueError(f"Hay {int(missing.sum())} filas sin ID de oportunidad.")

    dupes = refs[refs.duplicated(keep=False)]
    if not dupes.empty:
        sample = sorted(dupes.unique().tolist())[:10]
        raise ValueError(f"Hay oportunidades duplicadas en el Excel: {sample}")

    if limit is not None:
        df = df.head(limit).copy()
    return df


def _find_owner_record(token: str, owner_name: str) -> dict[str, str]:
    owner = owner_name.strip()
    if not owner:
        raise ValueError("Owner destino vacio.")

    team_data = _get_json(
        token,
        "teams",
        {
            "$select": "teamid,name",
            "$filter": f"name eq '{_escape_odata(owner)}'",
        },
    ).get("value", [])
    if len(team_data) == 1:
        return {
            "type": "team",
            "entity_set": "teams",
            "id": team_data[0]["teamid"],
            "name": team_data[0]["name"],
        }
    if len(team_data) > 1:
        raise ValueError(f"Owner {owner!r} coincide con varios teams.")

    user_data = _get_json(
        token,
        "systemusers",
        {
            "$select": "systemuserid,fullname,domainname",
            "$filter": f"fullname eq '{_escape_odata(owner)}'",
        },
    ).get("value", [])
    if len(user_data) == 1:
        return {
            "type": "systemuser",
            "entity_set": "systemusers",
            "id": user_data[0]["systemuserid"],
            "name": user_data[0]["fullname"],
        }
    if len(user_data) > 1:
        raise ValueError(f"Owner {owner!r} coincide con varios usuarios.")

    raise ValueError(f"No se encontro owner destino como team ni systemuser: {owner!r}")


def _resolve_owner_cached(token: str, owner_name: str, cache: dict[str, dict[str, str]]) -> dict[str, str]:
    key = owner_name.strip().casefold()
    if key not in cache:
        cache[key] = _find_owner_record(token, owner_name)
    return cache[key]


def _find_opportunity(token: str, row: pd.Series) -> dict[str, str]:
    opportunityid = _clean_value(row.get("opportunityid", ""))
    if opportunityid:
        data = _get_json(
            token,
            f"opportunities({opportunityid})",
            {"$select": "opportunityid,mcs_opportunityautonumber,_ownerid_value"},
        )
        return {
            "id": data["opportunityid"],
            "number": str(data.get("mcs_opportunityautonumber", "")),
            "owner_id": str(data.get("_ownerid_value", "")),
            "owner_name": str(
                data.get(
                    "_ownerid_value@OData.Community.Display.V1.FormattedValue",
                    "",
                )
            ),
        }

    number = _clean_value(row.get("ID de la Oportunidad", row.get("mcs_opportunityautonumber", "")))
    data = _get_json(
        token,
        "opportunities",
        {
            "$select": "opportunityid,mcs_opportunityautonumber,_ownerid_value",
            "$filter": f"mcs_opportunityautonumber eq '{_escape_odata(number)}'",
        },
    ).get("value", [])
    if len(data) != 1:
        raise ValueError(f"No se encontro una unica oportunidad en Atenea para {number}: {len(data)}")
    item = data[0]
    return {
        "id": item["opportunityid"],
        "number": str(item.get("mcs_opportunityautonumber", "")),
        "owner_id": str(item.get("_ownerid_value", "")),
        "owner_name": str(
            item.get("_ownerid_value@OData.Community.Display.V1.FormattedValue", "")
        ),
    }


def _patch_owner(token: str, opportunityid: str, owner: dict[str, str]) -> None:
    body = {"ownerid@odata.bind": f"/{owner['entity_set']}({owner['id']})"}
    response = _request_with_retry(
        "PATCH",
        f"{ORG_URL}/api/data/v9.2/opportunities({opportunityid})",
        headers={**_headers(token), "If-Match": "*"},
        json=body,
    )
    if response.status_code not in (204, 1223):
        raise RuntimeError(f"PATCH ownerid fallo {response.status_code}: {response.text}")


def _expected_owner_from_excel(row: pd.Series) -> dict[str, str]:
    owner_id = _clean_value(row.get("_ownerid_value", "")).lower()
    owner_name = _clean_value(row.get("owneridname", row.get("Propietario", "")))
    return {"id": owner_id, "name": owner_name}


def _owner_matches_expected(current: dict[str, str], expected: dict[str, str]) -> bool:
    if expected["id"]:
        return current["owner_id"].strip().lower() == expected["id"]
    if expected["name"]:
        return current["owner_name"].strip().casefold() == expected["name"].casefold()
    return True


def _owner_matches_destination(current: dict[str, str], destination: dict[str, str]) -> bool:
    current_id = current["owner_id"].strip().lower()
    destination_id = destination["id"].strip().lower()
    if current_id and destination_id:
        return current_id == destination_id
    return current["owner_name"].strip().casefold() == destination["name"].strip().casefold()


def _base_log_row(
    excel_path: Path,
    row: pd.Series,
    apply: bool,
    force_owner_mismatch: bool,
) -> dict[str, Any]:
    return {
        "status": "PENDING",
        "error": "",
        "excel": str(excel_path),
        "row_ref": _row_ref(row),
        "opportunity_number": "",
        "opportunityid": "",
        "owner_actual": "",
        "owner_actual_id": "",
        "owner_esperado_excel": "",
        "owner_esperado_excel_id": "",
        "owner_actual_coincide_excel": "",
        "owner_destino": "",
        "owner_destino_tipo": "",
        "owner_destino_id": "",
        "apply": apply,
        "force_owner_mismatch": force_owner_mismatch,
        "http_retries": 0,
    }


def _prepare_owner_change(
    token: str,
    row: pd.Series,
    excel_path: Path,
    owner_cache: dict[str, dict[str, str]],
    apply: bool,
    force_owner_mismatch: bool,
    owner_override: str | None = None,
) -> tuple[dict[str, Any], dict[str, str] | None, dict[str, str] | None]:
    log_row = _base_log_row(excel_path, row, apply, force_owner_mismatch)
    retries_before = REQUEST_RETRY_COUNT
    try:
        owner_name = owner_override or _clean_value(row.get("EQUIPO_FINAL", ""))
        if not owner_name:
            raise ValueError("EQUIPO_FINAL vacio en el Excel.")

        opportunity = _find_opportunity(token, row)
        owner = _resolve_owner_cached(token, owner_name, owner_cache)
        expected_owner = _expected_owner_from_excel(row)
        owner_matches_expected = _owner_matches_expected(opportunity, expected_owner)

        log_row.update(
            {
                "opportunity_number": opportunity["number"],
                "opportunityid": opportunity["id"],
                "owner_actual": opportunity["owner_name"],
                "owner_actual_id": opportunity["owner_id"],
                "owner_esperado_excel": expected_owner["name"],
                "owner_esperado_excel_id": expected_owner["id"],
                "owner_actual_coincide_excel": owner_matches_expected,
                "owner_destino": owner["name"],
                "owner_destino_tipo": owner["type"],
                "owner_destino_id": owner["id"],
            }
        )

        if not owner_matches_expected and not force_owner_mismatch:
            log_row["status"] = "SKIPPED_OWNER_CHANGED"
            log_row["error"] = (
                "Owner actual en Atenea no coincide con el owner original del Excel. "
                "No se aplica el cambio."
            )
            return log_row, None, None

        if _owner_matches_destination(opportunity, owner):
            log_row["status"] = "SKIPPED_ALREADY_OWNER"
            return log_row, None, None

        log_row["status"] = "READY_TO_APPLY" if apply else "DRY_RUN_CHANGE"
        return log_row, opportunity, owner
    except Exception as exc:
        log_row["status"] = "ERROR"
        log_row["error"] = str(exc)
        return log_row, None, None
    finally:
        log_row["http_retries"] = REQUEST_RETRY_COUNT - retries_before


def _write_log_rows(log_rows: list[dict[str, Any]]) -> Path:
    out_dir = Path("crm_ownerid_logs")
    out_dir.mkdir(exist_ok=True)
    path = out_dir / f"crm_ownerid_log_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    pd.DataFrame(log_rows).to_excel(path, index=False)
    return path


def _print_summary(log_rows: list[dict[str, Any]]) -> None:
    counts = pd.Series([row["status"] for row in log_rows]).value_counts().sort_index()
    print("\n=== RESUMEN OWNERID ATENEA ===")
    for status, count in counts.items():
        print(f"{status}: {count}")


def _run_all(args: argparse.Namespace, token: str, excel_path: Path) -> None:
    df = _load_all_rows(excel_path, limit=args.limit)
    owner_cache: dict[str, dict[str, str]] = {}
    prepared: list[tuple[dict[str, Any], dict[str, str] | None, dict[str, str] | None]] = []

    for _, row in df.iterrows():
        prepared.append(
            _prepare_owner_change(
                token=token,
                row=row,
                excel_path=excel_path,
                owner_cache=owner_cache,
                apply=args.apply,
                force_owner_mismatch=args.force_owner_mismatch,
            )
        )
        if args.sleep:
            time.sleep(args.sleep)

    log_rows = [item[0] for item in prepared]
    blocking = [row for row in log_rows if row["status"] in {"ERROR", "SKIPPED_OWNER_CHANGED"}]
    if args.apply and blocking:
        log_path = _write_log_rows(log_rows)
        _print_summary(log_rows)
        print("\n[STOP] No se aplico ningun cambio porque hay filas con error o owner cambiado.")
        print(f"Log: {log_path}")
        return

    if args.apply:
        for log_row, opportunity, owner in prepared:
            if log_row["status"] != "READY_TO_APPLY":
                continue
            try:
                retries_before = REQUEST_RETRY_COUNT
                _patch_owner(token, opportunity["id"], owner)
                log_row["status"] = "APPLIED"
                log_row["http_retries"] = int(log_row.get("http_retries", 0)) + (
                    REQUEST_RETRY_COUNT - retries_before
                )
                if args.sleep:
                    time.sleep(args.sleep)
            except Exception as exc:
                log_row["status"] = "ERROR_APPLY"
                log_row["error"] = str(exc)
                break

    log_path = _write_log_rows(log_rows)
    _print_summary(log_rows)
    if args.apply:
        print(f"\nCambios aplicados: {sum(row['status'] == 'APPLIED' for row in log_rows)}")
    else:
        print(f"\nCambios que se aplicarian: {sum(row['status'] == 'DRY_RUN_CHANGE' for row in log_rows)}")
    print(f"Log: {log_path}")


def _run_one(args: argparse.Namespace, token: str, excel_path: Path) -> None:
    owner_cache: dict[str, dict[str, str]] = {}
    row = _load_target_row(excel_path, args.id)
    log_row, opportunity, owner = _prepare_owner_change(
        token=token,
        row=row,
        excel_path=excel_path,
        owner_cache=owner_cache,
        apply=args.apply,
        force_owner_mismatch=args.force_owner_mismatch,
        owner_override=args.owner,
    )

    print("\n=== OWNERID ATENEA ===")
    print(f"Excel usado: {excel_path}")
    print(f"Oportunidad: {log_row['opportunity_number']} ({log_row['opportunityid']})")
    print(f"Owner actual: {log_row['owner_actual']} ({log_row['owner_actual_id']})")
    print(
        "Owner esperado Excel: "
        f"{log_row['owner_esperado_excel'] or '(sin nombre)'} "
        f"({log_row['owner_esperado_excel_id'] or 'sin id'})"
    )
    print(f"Owner actual coincide con Excel: {log_row['owner_actual_coincide_excel']}")
    print(
        "Owner destino: "
        f"{log_row['owner_destino']} [{log_row['owner_destino_tipo']}] "
        f"({log_row['owner_destino_id']})"
    )
    print(f"Status: {log_row['status']}")
    print(f"Apply: {args.apply}")

    if args.apply and log_row["status"] == "READY_TO_APPLY":
        _patch_owner(token, opportunity["id"], owner)
        log_row["status"] = "APPLIED"
        print("\n[OK] ownerid actualizado en Atenea.")
    elif args.apply and log_row["status"] != "SKIPPED_ALREADY_OWNER":
        print(f"\n[SKIP] {log_row['error'] or 'No hay cambio aplicable.'}")
    else:
        print("\nSin cambios.")

    log_path = _write_log_rows([log_row])
    print(f"Log: {log_path}")


def main() -> None:
    args = _parse_args()
    REQUEST_SETTINGS["max_retries"] = args.max_retries
    REQUEST_SETTINGS["retry_wait"] = args.retry_wait
    excel_path = Path(args.excel)
    token = _get_token(args.cache_file)
    if args.all:
        _run_all(args, token, excel_path)
    else:
        _run_one(args, token, excel_path)


if __name__ == "__main__":
    main()
