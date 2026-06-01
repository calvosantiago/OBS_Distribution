# =============================================================================
# EXTRACCIÓN DE LEADS DESDE ATENEA (Dataverse)
# =============================================================================
# Pega este bloque AL PRINCIPIO de tu script existente.
# Al final de este bloque tendrás un DataFrame "df" listo para usar.
#
# INSTALACIÓN (solo la primera vez, en tu terminal):
#   pip install msal requests pandas
# =============================================================================

import msal
import requests
import pandas as pd
import urllib.parse
import os
import copy
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone
from html import escape as html_escape

# zoneinfo viene incluido en Python 3.9+
# En Windows puede necesitar: pip install tzdata
try:
    from zoneinfo import ZoneInfo
except ImportError:
    try:
        from backports.zoneinfo import ZoneInfo
    except ImportError:
        raise ImportError(
            "Instala el paquete de zonas horarias: pip install tzdata"
        )

MADRID_TZ = ZoneInfo("Europe/Madrid")

# -----------------------------------------------------------------------------
# CONFIGURACIÓN — Solo tocar esta sección si algo cambia
# -----------------------------------------------------------------------------
ORG_URL    = "https://atenea.crm4.dynamics.com"
CLIENT_ID  = "51f81489-12ee-4a9e-aaae-a2591f45987d"  # App Microsoft Dynamics 365
AUTHORITY  = "https://login.microsoftonline.com/organizations"
SCOPES     = [f"{ORG_URL}/.default"]
CACHE_FILE = "token_cache.bin"  # Guarda el token para no pedir login cada día

# -----------------------------------------------------------------------------
# PASO 1 — Pedir fechas del periodo
# -----------------------------------------------------------------------------
print("=" * 50)
print("  Extracción de Leads — Atenea")
print("=" * 50)
print("Introduce las fechas del periodo (en hora Madrid).")
print("  Fecha inicio → primer día que quieres incluir  (ej: 2026-05-05)")
print("  Fecha fin    → último día que quieres incluir  (ej: 2026-06-01)\n")

fecha_inicio_input = input("Fecha inicio (YYYY-MM-DD): ").strip()
fecha_fin_input    = input("Fecha fin    (YYYY-MM-DD): ").strip()

# Conversión automática a UTC respetando horario de verano/invierno de Madrid
# Inicio : medianoche Madrid del primer día elegido
# Fin    : medianoche Madrid del día SIGUIENTE al último día elegido
#          (para incluir el último día completo con operador "lt")
dt_inicio_parsed = datetime.strptime(fecha_inicio_input, "%Y-%m-%d")
dt_fin_parsed    = datetime.strptime(fecha_fin_input,    "%Y-%m-%d") + timedelta(days=1)

dt_inicio_madrid = datetime(dt_inicio_parsed.year, dt_inicio_parsed.month,
                            dt_inicio_parsed.day, 0, 0, 0, tzinfo=MADRID_TZ)
dt_fin_madrid    = datetime(dt_fin_parsed.year, dt_fin_parsed.month,
                            dt_fin_parsed.day, 0, 0, 0, tzinfo=MADRID_TZ)

fecha_inicio = dt_inicio_madrid.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
fecha_fin    = dt_fin_madrid.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

offset_i = int(dt_inicio_madrid.utcoffset().total_seconds() / 3600)
offset_f = int(dt_fin_madrid.utcoffset().total_seconds() / 3600)

print(f"\n  Periodo Madrid : {fecha_inicio_input}  →  {fecha_fin_input} (inclusive)")
print(f"  Horario activo : UTC+{offset_i} (inicio)   UTC+{offset_f} (fin)")
print(f"  Rango UTC      : {fecha_inicio}  →  {fecha_fin}\n")

# -----------------------------------------------------------------------------
# PASO 2 — Autenticación con caché de token
# La primera vez abrirá el navegador para que inicies sesión.
# Las siguientes veces usará el token cacheado automáticamente.
# -----------------------------------------------------------------------------
cache = msal.SerializableTokenCache()
if os.path.exists(CACHE_FILE):
    cache.deserialize(open(CACHE_FILE, "r").read())

app = msal.PublicClientApplication(
    client_id=CLIENT_ID,
    authority=AUTHORITY,
    token_cache=cache
)

token_result = None
accounts = app.get_accounts()

if accounts:
    # Intentar renovar en silencio desde caché
    token_result = app.acquire_token_silent(SCOPES, account=accounts[0])

if not token_result:
    print("Abriendo navegador para autenticación (solo necesario la primera vez)...")
    token_result = app.acquire_token_interactive(scopes=SCOPES)

# Guardar caché actualizado
if cache.has_state_changed:
    open(CACHE_FILE, "w").write(cache.serialize())

if "access_token" not in token_result:
    raise Exception(
        f"Error de autenticación: {token_result.get('error_description', 'desconocido')}\n"
        "Si el error persiste, pide a IT un Client ID propio (App Registration en Azure AD)."
    )

token = token_result["access_token"]
print("✓ Autenticación correcta\n")

# -----------------------------------------------------------------------------
# PASO 3 — FetchXML con fechas dinámicas
# -----------------------------------------------------------------------------
fetchxml = f"""<fetch>
  <entity name="opportunity">
    <attribute name="name" />
    <attribute name="mcs_opportunityautonumber" />
    <!-- ID primario — necesario para deduplicación y detección de bucles -->
    <attribute name="opportunityid" />

    <!-- Fechas y números — campos reales, sin cambios -->
    <attribute name="createdon" />
    <attribute name="modifiedon" />
    <attribute name="mcs_dateofassignment" />
    <attribute name="mcs_interviewdonedate" />
    <attribute name="mcs_interviewscheduleddate" />
    <attribute name="mcs_interviewarrangementdate" />
    <attribute name="mcs_documentationvalidateddate" />
    <attribute name="mcs_unreachabledate" />
    <attribute name="mcs_firstcontactdate" />
    <attribute name="mcs_firstcontactattemptdate" />
    <attribute name="mcs_lastcontactdate" />
    <attribute name="mcs_lastcontactdate2" />
    <attribute name="mcs_lastcontactattemptdateuserlocal" />
    <attribute name="mcs_winclosingdate" />
    <attribute name="mcs_committeedate" />
    <attribute name="mcs_committeeresolutiondate" />
    <attribute name="mcs_quoteduedate" />
    <attribute name="mcs_wonquotedate" />
    <attribute name="actualclosedate" />
    <attribute name="mcs_dayssincelastcommunication" />
    <attribute name="mcs_ilocalizedattempts" />
    <attribute name="mcs_numberofemails" />
    <attribute name="mcs_numberofphonecalls" />
    <attribute name="actualvalue" />
    <attribute name="mcs_installments" />
    <attribute name="pfu_insightscore" />
    <attribute name="pfu_marketingnamepvc" />
    <attribute name="pfu_iasentprovider" />
    <attribute name="stepname" />
    <attribute name="mcs_leadduplicatedname" />

    <!-- Option sets — usar nombre base (sin "name"); la etiqueta llega como FormattedValue -->
    <attribute name="statuscode" />
    <attribute name="statecode" />
    <attribute name="mcs_pullpush" />
    <attribute name="mcs_type" />
    <attribute name="mcs_unlocated" />
    <attribute name="mcs_reopeningtype" />
    <attribute name="mcs_reopenedopportunity" />
    <attribute name="mcs_firstcontactattempt" />
    <attribute name="mcs_automaticassignment" />
    <attribute name="mcs_prevision" />
    <attribute name="mcs_paymentmethod" />

    <!-- Lookups — usar nombre base; GUID en _campo_value, etiqueta como FormattedValue -->
    <attribute name="ownerid" />
    <attribute name="owningbusinessunit" />
    <attribute name="mcs_academicperiodid" />
    <attribute name="mcs_programmeversioncampusid" />
    <attribute name="mcs_assignedteamid" />
    <attribute name="mcs_owninguserid" />
    <attribute name="mcs_closurereasonid" />
    <attribute name="mcs_closuresubreasonid" />
    <attribute name="mcs_teamassigmentruleid" />

    <filter type="and">
      <filter type="or">
        <condition attribute="ownerid" operator="not-in">
          <value>8BBE1932-3890-ED11-AAD1-6045BD8C9CB4</value>
          <value>FBBA6AA3-15B1-ED11-83FF-6045BD8C99FA</value>
          <value>D96F212C-3890-ED11-AAD1-6045BD8C9CB4</value>
          <value>E610189F-A3FB-ED11-8849-6045BD8C9A29</value>
          <value>1A843DC6-A0EA-EE11-A203-000D3A29B4D4</value>
          <value>E3965A26-73F5-EE11-A1FD-000D3A38858A</value>
          <value>821A5D98-87AA-EC11-983F-002248852F2C</value>
          <value>7F72212C-3890-ED11-AAD1-6045BD8C9CB4</value>
          <value>D0A18F3F-94C9-ED11-B597-6045BD8C966B</value>
        </condition>
        <condition attribute="ownerid" operator="null" />
      </filter>
      <condition attribute="mcs_opportunityautonumber" operator="begins-with" value="2021" />
      <condition attribute="owningbusinessunit" operator="eq" value="C4623CF3-A698-EC11-B400-000D3ABF3052" />
      <condition attribute="createdon" operator="ge" value="{fecha_inicio}" />
      <condition attribute="createdon" operator="lt" value="{fecha_fin}" />
    </filter>

    <link-entity name="systemuser" from="systemuserid" to="ownerid" link-type="outer" alias="systemuser_owner">
      <attribute name="mcs_salesteamid" />
      <attribute name="mcs_salescateory" />
    </link-entity>
    <link-entity name="mshied_program" from="mshied_programid" to="mcs_programid" link-type="outer" alias="program_interest">
      <attribute name="sis_qualificationid" />
      <attribute name="mcs_marketingname" />
    </link-entity>
    <link-entity name="contact" from="contactid" to="parentcontactid" link-type="outer" alias="contact_parent">
      <attribute name="sis_fullname" />
      <attribute name="address1_telephone1" />
      <attribute name="mobilephone" />
      <attribute name="sis_address1countryid" />
      <attribute name="sis_address1_stateorprovince" />
      <attribute name="address1_city" />
      <attribute name="emailaddress1" />
    </link-entity>
    <link-entity name="lead" from="leadid" to="originatingleadid" link-type="outer" alias="lead_origin">
      <attribute name="mcs_leadage" />
      <attribute name="mcs_address1countryid" />
      <attribute name="campaignid" />
      <attribute name="mcs_levelofstudyid" />
      <attribute name="mcs_programversionid" />
      <attribute name="mcs_observations" />
      <attribute name="mcs_observaciones" />
    </link-entity>
    <link-entity name="opportunity" from="opportunityid" to="mcs_origintingopportunityid" link-type="outer" alias="opp_origin">
      <attribute name="mcs_owninguserid" />
      <attribute name="ownerid" />
      <attribute name="mcs_automaticassignment" />
    </link-entity>
    <link-entity name="sis_admission" from="sis_admissionid" to="mcs_admissionid" link-type="outer" alias="admission">
      <attribute name="createdon" />
    </link-entity>
    <link-entity name="quote" from="quoteid" to="mcs_wonquoteid" link-type="outer" alias="quote_won">
      <attribute name="mcs_discountpercentage" />
      <attribute name="totallineitemamount" />
      <attribute name="totalamount" />
      <attribute name="mcs_inscriptionammount" />
      <attribute name="mcs_installments" />
    </link-entity>
    <link-entity name="mcs_paymentterm" from="mcs_paymenttermid" to="mcs_paymenttermid" link-type="outer" alias="paymentterm">
      <attribute name="mcs_defaultnumberofinstallments" />
    </link-entity>
    <link-entity name="sis_programmeversioncampus" from="sis_programmeversioncampusid" to="mcs_programmeversioncampusid" link-type="outer" alias="pvc">
      <attribute name="mcs_defaultnumberofinstallments" />
    </link-entity>
    <link-entity name="contact" from="contactid" to="customerid" link-type="outer" alias="contact_customer">
      <attribute name="gendercode" />
      <attribute name="sis_key_crmnet" />
    </link-entity>
    <link-entity name="campaign" from="campaignid" to="campaignid" link-type="outer" alias="campaign_origin">
      <attribute name="mcs_pillarid" />
      <attribute name="mcs_subpillarid" />
      <attribute name="mcs_pillarname" />
      <attribute name="mcs_subpillarname" />
      <attribute name="mcs_pullpush" />
      <attribute name="codename" />
      <attribute name="mcs_modalitycode" />
    </link-entity>
    <link-entity name="mshied_program" from="mshied_programid" to="mcs_mainprogramid" link-type="outer" alias="program_main">
      <attribute name="mcs_marketingname" />
    </link-entity>
    <link-entity name="mshied_programversion" from="mshied_programversionid" to="sis_programversionid" link-type="outer" alias="programversion">
      <attribute name="mcs_marketingnamepv" />
    </link-entity>
    <link-entity name="systemuser" from="systemuserid" to="mcs_assignedowner" link-type="outer" alias="systemuser_assigned">
      <attribute name="positionid" />
    </link-entity>
  </entity>
</fetch>"""

# -----------------------------------------------------------------------------
# PASO 4 — Extracción por ventanas de tiempo (sin paging cookies)
# Divide el periodo en chunks de CHUNK_DAYS días. Cada chunk devuelve
# menos de 5.000 registros y no necesita paginación.
# Si algún chunk llega a 5.000 → reduce CHUNK_DAYS.
# -----------------------------------------------------------------------------
CHUNK_DAYS = 5   # ← ajustar si algún chunk alcanza 5.000 registros

headers = {
    "Authorization": f"Bearer {token}",
    "Accept": "application/json",
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    "Prefer": 'odata.maxpagesize=5000,odata.include-annotations="*"'
}

# Parsear el FetchXML como árbol XML para modificar fechas por chunk
fetchxml_root = ET.fromstring(fetchxml)

all_records  = []
chunk_start  = dt_inicio_madrid.astimezone(timezone.utc)
end_limit    = dt_fin_madrid.astimezone(timezone.utc)
chunk_num    = 1
limit_warned = False

print("Extrayendo datos de Atenea...")

while chunk_start < end_limit:
    chunk_end = min(chunk_start + timedelta(days=CHUNK_DAYS), end_limit)
    start_str = chunk_start.strftime('%Y-%m-%dT%H:%M:%SZ')
    end_str   = chunk_end.strftime('%Y-%m-%dT%H:%M:%SZ')

    # Clonar árbol y actualizar condiciones de fecha para este chunk
    chunk_root   = copy.deepcopy(fetchxml_root)
    chunk_filter = chunk_root.find('entity').find('filter')
    for cond in chunk_filter.findall('condition'):
        if cond.get('attribute') == 'createdon':
            if cond.get('operator') == 'ge':
                cond.set('value', start_str)
            elif cond.get('operator') == 'lt':
                cond.set('value', end_str)

    url      = f"{ORG_URL}/api/data/v9.2/opportunities?fetchXml={urllib.parse.quote(ET.tostring(chunk_root, encoding='unicode'))}"
    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        raise Exception(f"Error {response.status_code}:\n{response.text}")

    records = response.json().get("value", [])
    all_records.extend(records)

    label = f"{chunk_start.strftime('%d/%m')} → {chunk_end.strftime('%d/%m')}"
    print(f"  Chunk {chunk_num} ({label}): {len(records):>5} registros  (total: {len(all_records)})")

    if len(records) >= 5000:
        print(f"  ⚠️  Chunk {chunk_num} alcanzó el límite — reduce CHUNK_DAYS a 2 o 3.")
        limit_warned = True

    chunk_start = chunk_end
    chunk_num  += 1

print(f"\n✓ Extracción completa: {len(all_records)} registros totales")
if limit_warned:
    print("⚠️  Algún chunk alcanzó 5.000 — puede haber datos faltantes. Reduce CHUNK_DAYS.\n")
else:
    print()

# -----------------------------------------------------------------------------
# PASO 5 — Convertir a DataFrame y renombrar columnas de valores formateados
#
# Dataverse devuelve las etiquetas como columnas con este sufijo:
#   "campo@OData.Community.Display.V1.FormattedValue"
#
# Las renombramos automáticamente:
#   statuscode@OData...FormattedValue       → statuscodename
#   _ownerid_value@OData...FormattedValue   → owneridname
#   alias._campo_value@OData...             → alias.campoidname
#   alias.campo@OData...                    → alias.camponame
# -----------------------------------------------------------------------------
ANNOTATION = "@OData.Community.Display.V1.FormattedValue"

df = pd.DataFrame(all_records)

rename_map = {}
for col in df.columns:
    if ANNOTATION not in col:
        continue
    base = col.replace(ANNOTATION, "")

    if "." in base:
        # Campo de link-entity: "alias.campo" o "alias._campo_value"
        alias, field = base.rsplit(".", 1)
        if field.startswith("_") and field.endswith("_value"):
            clean = field[1:-6]   # quita _ inicial y _value final
            rename_map[col] = f"{alias}.{clean}name"
        else:
            rename_map[col] = f"{alias}.{field}name"
    elif base.startswith("_") and base.endswith("_value"):
        # Lookup de entidad principal: "_ownerid_value" → "owneridname"
        clean = base[1:-6]
        rename_map[col] = f"{clean}name"
    else:
        # Option set de entidad principal: "statuscode" → "statuscodename"
        rename_map[col] = f"{base}name"

df = df.rename(columns=rename_map)

# Deduplicar por ID de oportunidad (por si la paginación introdujo duplicados)
if "opportunityid" in df.columns:
    before = len(df)
    df = df.drop_duplicates(subset=["opportunityid"])
    dupes = before - len(df)
    if dupes:
        print(f"⚠️  {dupes} duplicados eliminados por opportunityid")

# Eliminar columnas internas de OData (las que aún contienen "@")
df = df[[col for col in df.columns if "@" not in col]]

print(f"DataFrame listo: {df.shape[0]} filas × {df.shape[1]} columnas")

# Nombrar el DataFrame con la convención qb_cn_YYYYMMDD
fecha_hoy   = datetime.now(MADRID_TZ).strftime('%Y%m%d')
nombre_df   = f"qb_cn_{fecha_hoy}"
globals()[nombre_df] = df
print(f"DataFrame disponible como: {nombre_df}  ({df.shape[0]} filas × {df.shape[1]} columnas)")
print("-" * 50)
print("Columnas disponibles:")
for col in sorted(df.columns):
    print(f"  {col}")
print("-" * 50)

# =============================================================================
# EXTRACCIÓN 2 — LEADS ABIERTOS NO ASIGNADOS (op_no_asig)
# Sin filtro de fechas. Solo leads activos asignados a un equipo (no usuario).
# =============================================================================

# -----------------------------------------------------------------------------
# PASO 6 — FetchXML para leads abiertos no asignados
# -----------------------------------------------------------------------------
fetchxml2 = """<fetch>
  <entity name="opportunity">

    <attribute name="opportunityid" />
    <attribute name="createdon" />
    <attribute name="mcs_opportunityautonumber" />
    <attribute name="statecode" />
    <attribute name="mcs_region" />
    <attribute name="mcs_type" />
    <attribute name="mcs_pullpush" />
    <attribute name="mcs_suggestedadvisorid" />
    <attribute name="mcs_reopenedopportunity" />
    <attribute name="mcs_reopeningtype" />
    <attribute name="ownerid" />
    <attribute name="mcs_institutionid" />
    <attribute name="mcs_assignedteamid" />
    <attribute name="mcs_leadduplicatedname" />

    <filter type="and">
      <condition attribute="statecode"       operator="eq"  value="0" />
      <condition attribute="mcs_institutionid" operator="eq" value="7CA6E995-B198-EC11-B400-000D3A2E826F" />
    </filter>

    <!-- INNER JOIN Team: owner debe ser un equipo (lead no asignado a usuario) -->
    <link-entity name="team" from="teamid" to="ownerid" link-type="inner" alias="team_owner">
      <filter>
        <condition attribute="createdon" operator="not-null" />
      </filter>
    </link-entity>

    <!-- LEFT JOIN: Oportunidad de origen -->
    <link-entity name="opportunity" from="opportunityid" to="mcs_origintingopportunityid" link-type="outer" alias="opp_origin2">
      <attribute name="ownerid" />
    </link-entity>

    <!-- LEFT JOIN: Contacto principal (parentcontactid) -->
    <link-entity name="contact" from="contactid" to="parentcontactid" link-type="outer" alias="contact_parent2">
      <attribute name="fullname" />
      <attribute name="sis_address1countryid" />
      <attribute name="sis_address1_stateorprovince" />
      <attribute name="emailaddress1" />
    </link-entity>

    <!-- LEFT JOIN: Lead de origen -->
    <link-entity name="lead" from="leadid" to="originatingleadid" link-type="outer" alias="lead_origin2">
      <attribute name="mcs_address1countryid" />
      <attribute name="mcs_leadage" />
      <attribute name="campaignid" />
      <attribute name="mcs_programversionid" />
    </link-entity>

    <!-- INNER JOIN: Programa — solo qualifications específicas -->
    <link-entity name="mshied_program" from="mshied_programid" to="mcs_programid" link-type="inner" alias="program_interest2">
      <attribute name="sis_qualificationid" />
      <filter>
        <condition attribute="sis_qualificationid" operator="in">
          <value>2C6C0451-A7B1-EC11-9840-000D3A26C7B3</value>
          <value>09AAB12D-1CCE-ED11-B597-6045BD8C9CB4</value>
          <value>2B6C0451-A7B1-EC11-9840-000D3A26C7B3</value>
        </condition>
      </filter>
    </link-entity>

    <!-- INNER JOIN: Campaña — excluir un pillar específico -->
    <link-entity name="campaign" from="campaignid" to="campaignid" link-type="inner" alias="campaign_origin2">
      <attribute name="mcs_pillarid" />
      <attribute name="mcs_pillarname" />
      <attribute name="mcs_subpillarid" />
      <filter type="or">
        <condition attribute="mcs_pillarid" operator="ne"   value="5754583B-709C-ED11-AAD1-6045BD8C9877" />
        <condition attribute="mcs_pillarid" operator="null" />
      </filter>
    </link-entity>

    <!-- LEFT JOIN: Contacto cliente (customerid) -->
    <link-entity name="contact" from="contactid" to="customerid" link-type="outer" alias="contact_customer2">
      <attribute name="address1_telephone1" />
      <attribute name="sis_key_crmnet" />
      <attribute name="sis_key_migration" />
    </link-entity>

  </entity>
</fetch>"""

# -----------------------------------------------------------------------------
# PASO 7 — Llamada a la API (sin chunking de fechas — no hay date filter)
# Si el resultado llega a 5.000 exactos avisa, pero para leads no asignados
# el volumen suele ser muy inferior.
# -----------------------------------------------------------------------------
print("\n" + "=" * 50)
print("  Extracción 2 — Leads Abiertos No Asignados")
print("=" * 50)

url2     = f"{ORG_URL}/api/data/v9.2/opportunities?fetchXml={urllib.parse.quote(fetchxml2)}"
response2 = requests.get(url2, headers=headers)

if response2.status_code != 200:
    raise Exception(f"Error {response2.status_code} en op_no_asig:\n{response2.text}")

records2 = response2.json().get("value", [])
print(f"  Registros obtenidos: {len(records2)}")

if len(records2) >= 5000:
    print("  ⚠️  Se alcanzaron 5.000 registros — pueden faltar datos.")
    print("     Contacta a IT para implementar paginación en esta vista.")

# -----------------------------------------------------------------------------
# PASO 8 — DataFrame op_no_asig con el mismo renombrado de columnas
# -----------------------------------------------------------------------------
df2 = pd.DataFrame(records2)

rename_map2 = {}
for col in df2.columns:
    if ANNOTATION not in col:
        continue
    base = col.replace(ANNOTATION, "")
    if "." in base:
        alias, field = base.rsplit(".", 1)
        if field.startswith("_") and field.endswith("_value"):
            rename_map2[col] = f"{alias}.{field[1:-6]}name"
        else:
            rename_map2[col] = f"{alias}.{field}name"
    elif base.startswith("_") and base.endswith("_value"):
        rename_map2[col] = f"{base[1:-6]}name"
    else:
        rename_map2[col] = f"{base}name"

df2 = df2.rename(columns=rename_map2)

if "opportunityid" in df2.columns:
    before2 = len(df2)
    df2 = df2.drop_duplicates(subset=["opportunityid"])
    if before2 - len(df2):
        print(f"  ⚠️  {before2 - len(df2)} duplicados eliminados")

df2 = df2[[col for col in df2.columns if "@" not in col]]

nombre_df2        = f"op_no_asig_{fecha_hoy}"
globals()[nombre_df2] = df2
print(f"  DataFrame disponible como: {nombre_df2}  ({df2.shape[0]} filas × {df2.shape[1]} columnas)")

# =============================================================================
# EXPORTAR AMBOS DataFrames A EXCEL (una hoja por DataFrame)
# =============================================================================
output_file = f"leads_{fecha_inicio_input}_{fecha_fin_input}.xlsx"
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    globals()[nombre_df].to_excel(writer, sheet_name=nombre_df,  index=False)
    df2.to_excel(writer,               sheet_name=nombre_df2, index=False)

print(f"\n✓ Exportado a: {output_file}")
print(f"  · Hoja 1: {nombre_df}   ({globals()[nombre_df].shape[0]} filas)")
print(f"  · Hoja 2: {nombre_df2}  ({df2.shape[0]} filas)")

# =============================================================================
# A PARTIR DE AQUÍ SIGUE TU CÓDIGO EXISTENTE
# DataFrames disponibles:
#   globals()[nombre_df]   → qb_cn_YYYYMMDD
#   globals()[nombre_df2]  → op_no_asig_YYYYMMDD
# =============================================================================