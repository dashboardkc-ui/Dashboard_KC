import os
import json
import time
import requests
import builtins
import unicodedata

from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
import pytz

# Força flush imediato em todos os prints
_original_print = builtins.print
def print(*args, **kwargs):
    kwargs["flush"] = True
    _original_print(*args, **kwargs)

API_TIMEOUT = 60  # segundos

# ==============================
# CONFIGURAÇÕES
# ==============================

SOCIA_API_KEY = os.environ.get("SOCIAVAULT_API_KEY")

POSTS_LIMIT = 20  # quantos posts buscar por hashtag

SPREADSHEET_ID = "1cn68TA8_ajbbIOaMofE_7-Vc4_BWfQRMHehrO6SB_Q4"

SHEET_HASHTAGS = "Hashtag_Data"          # aba de entrada
SHEET_POSTS    = "Hashtag_posts"         # aba intermediária
SHEET_DETAIL   = "Hashtag_posts_detail"  # aba de saída final

tz_br = pytz.timezone("America/Sao_Paulo")

# Janela mínima (em dias) entre execuções da mesma hashtag
MIN_DAYS_BETWEEN_RUNS = 30

# Países da América Central e do Sul (ISO 2) — apenas posts dessas regiões são salvos
ALLOWED_COUNTRIES = {
    "BZ", "CR", "SV", "GT", "HN", "NI", "PA",
    "AR", "BO", "BR", "CL", "CO", "EC", "GY", "PY", "PE", "SR", "UY", "VE",
}


# ==============================
# GOOGLE SERVICES
# ==============================

def get_google_services():
    creds_json = json.loads(os.environ.get("GDRIVE_CREDENTIALS"))
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = service_account.Credentials.from_service_account_info(
        creds_json, scopes=scopes
    )
    sheets_service = build("sheets", "v4", credentials=creds)
    return sheets_service


# ==============================
# ETAPA 1 — LER HASHTAGS
# ==============================

def parse_run_datetime(raw_date):
    """Tenta parsear uma data/hora de execução em vários formatos. Retorna datetime tz-aware ou None."""
    raw_date = (raw_date or "").strip()
    if not raw_date:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y"):
        try:
            parsed = datetime.strptime(raw_date, fmt)
            return tz_br.localize(parsed)
        except ValueError:
            continue
    return None


def read_hashtags(sheets_service):
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_HASHTAGS}!A:Z"
    ).execute()

    rows = result.get("values", [])
    if len(rows) <= 1:
        print("Nenhuma hashtag encontrada na planilha.")
        return []

    headers = [h.strip().lower() for h in rows[0]]

    if "hashtag" not in headers:
        print(f"Coluna 'Hashtag' não encontrada. Colunas disponíveis: {headers}")
        return []

    col_hashtag = headers.index("hashtag")

    # FIX 3: validação segura das colunas opcionais antes de usar como índice
    col_country    = headers.index("country")    if "country"    in headers else None
    col_marca_kc   = headers.index("marca_kc")   if "marca_kc"   in headers else None
    col_competidor = headers.index("competidor") if "competidor" in headers else None
    col_pais       = headers.index("pais")       if "pais"       in headers else None
    col_rundate    = headers.index("rundatetime") if "rundatetime" in headers else None

    if col_country is None:
        print("  Aviso: coluna 'Country' não encontrada. Sem filtro de país.")
    if col_rundate is None:
        print("  Aviso: coluna 'rundatetime' não encontrada em Hashtag_Data. Será criada ao gravar.")

    entries = []
    seen = set()

    # row_idx começa em 2 pois rows[0] é o header (linha 1 do Sheets)
    for row_idx, row in enumerate(rows[1:], start=2):
        if len(row) <= col_hashtag:
            continue

        tag = row[col_hashtag].strip().lstrip("#")

        if not tag:
            continue

        country    = (row[col_country].strip().upper()  if col_country    is not None and len(row) > col_country    else "")
        marca_kc   = (row[col_marca_kc].strip()         if col_marca_kc   is not None and len(row) > col_marca_kc   else "")
        competidor = (row[col_competidor].strip()        if col_competidor is not None and len(row) > col_competidor else "")
        pais       = (row[col_pais].strip()              if col_pais       is not None and len(row) > col_pais       else "")
        last_run   = (parse_run_datetime(row[col_rundate]) if col_rundate is not None and len(row) > col_rundate else None)

        key = (tag.lower(), country, marca_kc, competidor, pais)

        if key not in seen:
            seen.add(key)

            entries.append({
                "hashtag": tag,
                "country": country,
                "marca_kc": marca_kc,
                "competidor": competidor,
                "pais": pais,
                "row_index": row_idx,
                "last_run": last_run
            })

    print(f"  {len(entries)} entrada(s) única(s) encontrada(s):")
    for e in entries:
        pais_label = e['country'] if e['country'] else "todos os países"
        print(f"    #{e['hashtag']} → {pais_label}")

    return entries


# ==============================
# ETAPA 2 — BUSCAR POSTS POR HASHTAG
# ==============================

def fetch_posts_by_hashtag(hashtag):
    url = "https://api.sociavault.com/v1/scrape/tiktok/search/hashtag"
    headers = {"X-API-Key": SOCIA_API_KEY}
    params = {"hashtag": hashtag, "count": POSTS_LIMIT}

    response = requests.get(url, headers=headers, params=params, timeout=API_TIMEOUT)
    print(f"    Status ({hashtag}): {response.status_code}")
    response.raise_for_status()

    data = response.json()
    aweme_list = data.get("data", {}).get("aweme_list", {})

    if isinstance(aweme_list, dict):
        items = list(aweme_list.values())
    elif isinstance(aweme_list, list):
        items = aweme_list
    else:
        print(f"    Aviso: estrutura inesperada de aweme_list: {type(aweme_list)}")
        return []

    results = []

    for item in items:
        share_url = (
            item.get("share_info", {}).get("share_url")
            or item.get("share_url")
            or ""
        )
        if not share_url:
            continue

        region = (
            item.get("author", {}).get("region", "")
            or item.get("region", "")
        ).upper()

        results.append({"share_url": share_url, "region": region})

    print(f"    Posts encontrados: {len(results)}")
    return results


def get_existing_urls_posts(sheets_service):
    """Retorna set de share_urls já salvas na aba Hashtag_posts."""
    try:
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_POSTS}!A:Z"
        ).execute()
        rows = result.get("values", [])
        if len(rows) <= 1:
            return set()
        headers = [h.strip().lower() for h in rows[0]]
        if "share_url" not in headers:
            return set()
        url_col = headers.index("share_url")
        existing = set()
        for row in rows[1:]:
            if len(row) > url_col and row[url_col].strip():
                existing.add(row[url_col].strip())
        print(f"  URLs já salvas em Hashtag_posts: {len(existing)}")
        return existing
    except Exception as e:
        print(f"  Aviso ao ler Hashtag_posts: {e}")
        return set()


def ensure_rundatetime_column(sheets_service):
    """Garante que a aba Hashtag_Data tenha uma coluna 'rundatetime'. Retorna o índice (0-based) da coluna."""
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_HASHTAGS}!1:1"
    ).execute()
    header_row = result.get("values", [[]])
    headers = header_row[0] if header_row else []
    headers_lower = [h.strip().lower() for h in headers]

    if "rundatetime" in headers_lower:
        return headers_lower.index("rundatetime")

    new_col_idx = len(headers)
    new_col_letter = column_index_to_letter(new_col_idx)
    sheets_service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_HASHTAGS}!{new_col_letter}1",
        valueInputOption="RAW",
        body={"values": [["rundatetime"]]}
    ).execute()
    print(f"  Coluna 'rundatetime' criada em Hashtag_Data (coluna {new_col_letter}).")
    return new_col_idx


def update_rundatetime_in_hashtag_data(sheets_service, row_indices, run_datetime):
    """Grava run_datetime na coluna 'rundatetime' de Hashtag_Data, para as linhas processadas nesta rodada."""
    if not row_indices:
        print("  Nenhuma hashtag processada nesta rodada. Nada para atualizar em Hashtag_Data.")
        return

    rundate_col_idx = ensure_rundatetime_column(sheets_service)
    rundate_col_letter = column_index_to_letter(rundate_col_idx)

    data_updates = [
        {
            "range": f"{SHEET_HASHTAGS}!{rundate_col_letter}{row_idx}",
            "values": [[run_datetime]]
        }
        for row_idx in row_indices
    ]

    BATCH_SIZE = 500
    for start in range(0, len(data_updates), BATCH_SIZE):
        chunk = data_updates[start:start + BATCH_SIZE]
        sheets_service.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={
                "valueInputOption": "RAW",
                "data": chunk
            }
        ).execute()

    print(f"  Hashtag_Data: rundatetime atualizado em {len(row_indices)} linha(s).")


def save_posts_to_sheets(sheets_service, rows_to_add):
    if not rows_to_add:
        print("  Nenhuma linha nova para salvar em Hashtag_posts.")
        return

    existing_data = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_POSTS}!A:A"
    ).execute()
    existing_rows = existing_data.get("values", [])

    if not existing_rows:
        header = [[
            "hashtag",
            "share_url",
            "country",
            "marca_kc",
            "competidor",
            "pais",
            "region",
            "run_datetime"
        ]]
        values = header + rows_to_add
        sheets_service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_POSTS}!A1",
            valueInputOption="RAW",
            body={"values": values}
        ).execute()
        print(f"  Hashtag_posts: {len(rows_to_add)} linhas inseridas com cabeçalho.")
    else:
        sheets_service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_POSTS}!A:A",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": rows_to_add}
        ).execute()
        print(f"  Hashtag_posts: {len(rows_to_add)} novas linhas adicionadas.")


# ==============================
# ETAPA 3 — BUSCAR VIDEO INFO
# ==============================

def fetch_video_info(share_url):
    url = "https://api.sociavault.com/v1/scrape/tiktok/video-info"
    headers = {"X-API-Key": SOCIA_API_KEY}
    params = {"url": share_url}

    response = requests.get(url, headers=headers, params=params, timeout=API_TIMEOUT)
    print(f"    Status video-info: {response.status_code}")
    response.raise_for_status()

    data = response.json()
    return data.get("data", {}).get("aweme_detail", {})


def extract_fields(detail, share_url, hashtag, country, marca_kc, competidor, pais, run_datetime):
    author = detail.get("author", {})
    stats  = detail.get("statistics", {})

    create_time_raw = detail.get("create_time", "")
    try:
        create_time = datetime.fromtimestamp(create_time_raw, tz=tz_br).strftime("%Y-%m-%d %H:%M:%S") if create_time_raw else ""
    except Exception:
        create_time = ""

    return [
        "hashtag",  # type_post — origem do post sempre é busca por hashtag neste script
        share_url, hashtag, country, marca_kc, competidor,
        detail.get("region", ""),  # video_region — ocupa o lugar de 'pais'
        run_datetime,
        detail.get("aweme_id", ""),
        detail.get("desc", ""),
        create_time,
        author.get("unique_id", ""),
        author.get("nickname", ""),
        author.get("follower_count", ""),
        stats.get("play_count", ""),
        stats.get("digg_count", ""),
        stats.get("comment_count", ""),
        stats.get("share_count", ""),
        stats.get("collect_count", ""),
        stats.get("download_count", ""),
        stats.get("repost_count", ""),
    ]


DETAIL_HEADER = [
    "type_post", "share_url", "hashtag", "country", "marca_kc", "competidor", "video_region", "run_datetime",
    "aweme_id", "description", "create_time",
    "author_username", "author_nickname", "author_followers",
    "play_count", "like_count", "comment_count", "share_count",
    "save_count", "download_count", "repost_count",
]


def get_processed_aweme_ids(sheets_service):
    """Retorna set de aweme_ids já existentes em Hashtag_posts_detail."""
    try:
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_DETAIL}!A:Z"
        ).execute()
        rows = result.get("values", [])
        if len(rows) <= 1:
            return set()
        headers = [h.strip().lower() for h in rows[0]]
        if "aweme_id" not in headers:
            print("  Aviso: coluna 'aweme_id' não encontrada em Hashtag_posts_detail. Sem checagem por aweme_id.")
            return set()
        id_col = headers.index("aweme_id")
        processed = set()
        for row in rows[1:]:
            if len(row) > id_col and row[id_col].strip():
                processed.add(row[id_col].strip())
        print(f"  aweme_ids já existentes em Hashtag_posts_detail: {len(processed)}")
        return processed
    except Exception as e:
        print(f"  Aviso ao ler aweme_ids de Hashtag_posts_detail: {e}")
        return set()


def get_processed_urls(sheets_service):
    """Retorna set de share_urls já processadas em Hashtag_posts_detail."""
    try:
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_DETAIL}!A:Z"
        ).execute()
        rows = result.get("values", [])
        if len(rows) <= 1:
            return set()
        headers = [h.strip().lower() for h in rows[0]]
        if "share_url" not in headers:
            return set()
        url_col = headers.index("share_url")
        processed = set()
        for row in rows[1:]:
            if len(row) > url_col and row[url_col].strip():
                processed.add(row[url_col].strip())
        print(f"  URLs já processadas em Hashtag_posts_detail: {len(processed)}")
        return processed
    except Exception as e:
        print(f"  Aviso ao ler Hashtag_posts_detail: {e}")
        return set()


# ==============================
# ETAPA 4 — PREENCHER DESCRIPTION EM HASHTAG_POSTS
# ==============================

def column_index_to_letter(idx):
    """Converte índice de coluna (0-based) para letra do Google Sheets (A, B, ..., Z, AA, ...)."""
    letter = ""
    idx += 1
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter


def ensure_description_column(sheets_service):
    """Garante que a aba Hashtag_posts tenha uma coluna 'description'. Retorna o índice (0-based) da coluna."""
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_POSTS}!1:1"
    ).execute()
    header_row = result.get("values", [[]])
    headers = header_row[0] if header_row else []
    headers_lower = [h.strip().lower() for h in headers]

    if "description" in headers_lower:
        return headers_lower.index("description")

    new_col_idx = len(headers)
    new_col_letter = column_index_to_letter(new_col_idx)
    sheets_service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_POSTS}!{new_col_letter}1",
        valueInputOption="RAW",
        body={"values": [["description"]]}
    ).execute()
    print(f"  Coluna 'description' criada em Hashtag_posts (coluna {new_col_letter}).")
    return new_col_idx


# ==============================
# CHECAGEM HASHTAG NA DESCRIPTION
# ==============================

def normalize_text(text):
    """Lowercase + remove acentos/diacríticos (NFKD), igual à fórmula do Sheets generalizada."""
    if not text:
        return ""
    text = text.strip().lower()
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in nfkd if not unicodedata.combining(ch))


def hashtag_in_description(hashtag, description):
    """Retorna True se a hashtag (normalizada) estiver contida na description (normalizada)."""
    norm_hashtag = normalize_text(hashtag)
    norm_desc = normalize_text(description)
    if not norm_hashtag:
        return False
    return norm_hashtag in norm_desc


def get_sheet_id(sheets_service, sheet_name):
    """Retorna o sheetId (gid) numérico de uma aba pelo nome."""
    metadata = sheets_service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for sheet in metadata.get("sheets", []):
        props = sheet.get("properties", {})
        if props.get("title") == sheet_name:
            return props.get("sheetId")
    return None


def remove_rows_from_hashtag_posts(sheets_service, share_urls_to_remove):
    """Remove as linhas da aba Hashtag_posts cujo share_url esteja em share_urls_to_remove."""
    if not share_urls_to_remove:
        return

    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_POSTS}!A:Z"
    ).execute()
    rows = result.get("values", [])
    if len(rows) <= 1:
        return

    headers = [h.strip().lower() for h in rows[0]]
    if "share_url" not in headers:
        print("  Aviso: coluna 'share_url' não encontrada em Hashtag_posts. Não foi possível remover linhas.")
        return
    col_url = headers.index("share_url")

    sheet_id = get_sheet_id(sheets_service, SHEET_POSTS)
    if sheet_id is None:
        print(f"  Aviso: não foi possível localizar o sheetId de '{SHEET_POSTS}'.")
        return

    row_indices_to_delete = []
    for i, row in enumerate(rows[1:], start=2):  # linha 2 = primeira linha de dados (1-indexed no Sheets)
        if len(row) <= col_url:
            continue
        url = row[col_url].strip()
        if url in share_urls_to_remove:
            row_indices_to_delete.append(i)

    if not row_indices_to_delete:
        return

    # Deleta de baixo para cima para não desalinhar os índices das linhas restantes
    row_indices_to_delete.sort(reverse=True)
    requests_batch = []
    for row_num in row_indices_to_delete:
        requests_batch.append({
            "deleteDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": row_num - 1,  # 0-based
                    "endIndex": row_num
                }
            }
        })

    BATCH_SIZE = 500
    for start in range(0, len(requests_batch), BATCH_SIZE):
        chunk = requests_batch[start:start + BATCH_SIZE]
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": chunk}
        ).execute()

    print(f"  Hashtag_posts: {len(row_indices_to_delete)} linha(s) removida(s) (hashtag não encontrada na description).")


def fill_descriptions_in_posts(sheets_service, processed_rows):
    """Preenche a coluna 'description' em Hashtag_posts apenas para os posts
    processados nesta rodada (sem backfill de posts antigos)."""

    if not processed_rows:
        print("  Nenhum post processado nesta rodada. Nada para preencher.")
        return

    # processed_rows: lista de tuplas (share_url, description)
    desc_by_url = {url: desc for url, desc in processed_rows if url and desc}
    if not desc_by_url:
        print("  Nenhuma description disponível nos posts processados.")
        return

    desc_col_idx = ensure_description_column(sheets_service)
    desc_col_letter = column_index_to_letter(desc_col_idx)

    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_POSTS}!A:Z"
    ).execute()
    rows = result.get("values", [])
    if len(rows) <= 1:
        print("  Hashtag_posts vazia, nada para preencher.")
        return

    headers = [h.strip().lower() for h in rows[0]]
    if "share_url" not in headers:
        print("  Aviso: coluna 'share_url' não encontrada em Hashtag_posts.")
        return
    col_url = headers.index("share_url")

    data_updates = []
    filled = 0

    for i, row in enumerate(rows[1:], start=2):  # linha 2 = primeira linha de dados
        if len(row) <= col_url:
            continue
        url = row[col_url].strip()
        if not url or url not in desc_by_url:
            continue

        data_updates.append({
            "range": f"{SHEET_POSTS}!{desc_col_letter}{i}",
            "values": [[desc_by_url[url]]]
        })
        filled += 1

    if not data_updates:
        print("  Nenhuma description nova para preencher em Hashtag_posts.")
        return

    BATCH_SIZE = 500
    for start in range(0, len(data_updates), BATCH_SIZE):
        chunk = data_updates[start:start + BATCH_SIZE]
        sheets_service.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={
                "valueInputOption": "RAW",
                "data": chunk
            }
        ).execute()

    print(f"  Hashtag_posts: {filled} description(s) preenchida(s) (apenas posts processados nesta rodada).")


def save_details_to_sheets(sheets_service, rows_to_add):
    if not rows_to_add:
        print("  Nenhuma linha nova para salvar em Hashtag_posts_detail.")
        return

    existing_data = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_DETAIL}!A:A"
    ).execute()
    existing_rows = existing_data.get("values", [])

    if not existing_rows:
        values = [DETAIL_HEADER] + rows_to_add
        sheets_service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_DETAIL}!A1",
            valueInputOption="RAW",
            body={"values": values}
        ).execute()
        print(f"  Hashtag_posts_detail: {len(rows_to_add)} linhas inseridas com cabeçalho.")
    else:
        sheets_service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_DETAIL}!A:A",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": rows_to_add}
        ).execute()
        print(f"  Hashtag_posts_detail: {len(rows_to_add)} novas linhas adicionadas.")


# ==============================
# EXECUÇÃO PRINCIPAL
# ==============================

def main():
    print("=" * 60)
    print("INICIANDO PIPELINE HASHTAG TIKTOK")
    print("=" * 60)

    print("\n[CONFIG] Verificando variáveis de ambiente...")
    missing = []
    for var in ["SOCIAVAULT_API_KEY", "GDRIVE_CREDENTIALS"]:
        val = os.environ.get(var)
        if not val:
            missing.append(var)
            print(f"  ERRO: {var} não encontrada!")
        else:
            print(f"  OK: {var} ({len(val)} chars)")

    if missing:
        print(f"\nVariáveis faltando: {missing}. Encerrando.")
        return

    print("\n[CONFIG] Inicializando Google Services...")
    sheets_service = get_google_services()
    print("  Google Services OK")

    # ── ETAPA 1 — Ler hashtags ──────────────────────────────────
    print(f"\n{'=' * 60}")
    print(f"[ETAPA 1] Lendo hashtags de '{SHEET_HASHTAGS}'...")
    print(f"{'=' * 60}")
    entries = read_hashtags(sheets_service)

    if not entries:
        print("Nenhuma hashtag para processar. Encerrando.")
        return

    # ── ETAPA 2 — Buscar posts por hashtag ─────────────────────
    print(f"\n{'=' * 60}")
    print(f"[ETAPA 2] Buscando posts por hashtag...")
    print(f"{'=' * 60}")

    # FIX 2: indentação corrigida — fora do for, no nível correto
    existing_urls = get_existing_urls_posts(sheets_service)
    now_br = datetime.now(tz_br)
    run_datetime = now_br.strftime("%Y-%m-%d %H:%M:%S")

    new_post_rows = []
    total_filtered_by_country = 0
    rows_to_mark_rundatetime = []  # row_index das hashtags processadas com sucesso (p/ Hashtag_Data)

    # Acumula todos os posts para etapa 3
    all_posts = []

    for entry in entries:

        hashtag    = entry["hashtag"]
        country    = entry["country"]
        marca_kc   = entry["marca_kc"]
        competidor = entry["competidor"]
        pais       = entry["pais"]
        row_index  = entry["row_index"]

        print(
            f"\n  HASHTAG: #{hashtag}" +
            (f" | PAÍS: {country}" if country else " | PAÍS: todos")
        )

        # ── Checagem de janela mínima entre execuções ──────────
        last_run = entry["last_run"]
        if last_run is not None:
            days_since_last_run = (now_br - last_run).days
            if days_since_last_run <= MIN_DAYS_BETWEEN_RUNS:
                print(
                    f"    Pulando #{hashtag}: última execução há {days_since_last_run} dia(s) "
                    f"(<= {MIN_DAYS_BETWEEN_RUNS} dias). Última run: {last_run.strftime('%Y-%m-%d %H:%M:%S')}"
                )
                continue

        try:
            posts = fetch_posts_by_hashtag(hashtag)

        except Exception as e:
            print(f"    ERRO ao buscar #{hashtag}: {e}. Pulando.")
            continue

        # Hashtag processada com sucesso nesta rodada -> marcar rundatetime em Hashtag_Data
        rows_to_mark_rundatetime.append(row_index)

        filtered_out = sum(1 for p in posts if p["region"] not in ALLOWED_COUNTRIES)
        total_filtered_by_country += filtered_out
        if filtered_out:
            print(f"    Posts descartados (fora da América Central/Sul): {filtered_out}")

        for post in posts:

            share_url = post["share_url"]
            region    = post["region"]

            if region not in ALLOWED_COUNTRIES:
                continue

            all_posts.append({
                "share_url": share_url,
                "hashtag": hashtag,
                "country": country,
                "marca_kc": marca_kc,
                "competidor": competidor,
                "pais": pais
            })

            if share_url in existing_urls:
                continue

            new_post_rows.append([
                hashtag,
                share_url,
                country,
                marca_kc,
                competidor,
                pais,
                region,
                run_datetime
            ])

            existing_urls.add(share_url)

        print(
            f"    Novos posts para #{hashtag}: "
            f"{sum(1 for r in new_post_rows if r[0] == hashtag)}"
        )

        time.sleep(1)

    # Salva posts novos em Hashtag_posts
    sheets_service = get_google_services()
    save_posts_to_sheets(sheets_service, new_post_rows)

    # Atualiza rundatetime em Hashtag_Data para as hashtags processadas nesta rodada
    sheets_service = get_google_services()
    update_rundatetime_in_hashtag_data(sheets_service, rows_to_mark_rundatetime, run_datetime)

    # ── ETAPA 3 — Buscar video-info para todos os posts ────────
    print(f"\n{'=' * 60}")
    print(f"[ETAPA 3] Buscando video-info para {len(all_posts)} post(s)...")
    print(f"{'=' * 60}")

    sheets_service     = get_google_services()
    processed_urls     = get_processed_urls(sheets_service)
    processed_aweme_ids = get_processed_aweme_ids(sheets_service)
    run_datetime    = datetime.now(tz_br).strftime("%Y-%m-%d %H:%M:%S")
    new_detail_rows = []
    urls_to_remove  = []  # posts sem a hashtag na description -> remover de Hashtag_posts
    errors          = 0
    skipped_duplicate_aweme_id = 0

    for i, post in enumerate(all_posts, start=1):
        share_url  = post["share_url"]
        hashtag    = post["hashtag"]
        country    = post["country"]
        marca_kc   = post["marca_kc"]
        competidor = post["competidor"]
        pais       = post["pais"]

        if share_url in processed_urls:
            print(f"  [{i}/{len(all_posts)}] Já processado, pulando.")
            continue

        print(f"\n  [{i}/{len(all_posts)}] #{hashtag} | {share_url}")

        try:
            detail = fetch_video_info(share_url)
            if not detail:
                print(f"    Aviso: resposta vazia. Pulando.")
                continue

            aweme_id = str(detail.get("aweme_id", "")).strip()

            if aweme_id and aweme_id in processed_aweme_ids:
                print(f"    Descartado: aweme_id '{aweme_id}' já existe em Hashtag_posts_detail. Pulando.")
                processed_urls.add(share_url)
                skipped_duplicate_aweme_id += 1
                continue

            description = detail.get("desc", "")

            if not hashtag_in_description(hashtag, description):
                print(f"    Descartado: hashtag '#{hashtag}' não encontrada na description. Removendo de Hashtag_posts.")
                urls_to_remove.append(share_url)
                continue

            row = extract_fields(detail, share_url, hashtag, country, marca_kc, competidor, pais, run_datetime)
            row = [str(v) if v is not None else "" for v in row]
            new_detail_rows.append(row)
            processed_urls.add(share_url)
            if aweme_id:
                processed_aweme_ids.add(aweme_id)
            print(f"    OK — @{detail.get('author', {}).get('unique_id', '?')} | plays: {detail.get('statistics', {}).get('play_count', '?')}")

        except Exception as e:
            print(f"    ERRO: {e}. Pulando.")
            errors += 1

        time.sleep(1)

    # Salva detalhes em Hashtag_posts_detail
    sheets_service = get_google_services()
    save_details_to_sheets(sheets_service, new_detail_rows)

    # ── Remove de Hashtag_posts os posts cuja hashtag não apareceu na description ──
    if urls_to_remove:
        print(f"\n{'=' * 60}")
        print(f"Removendo {len(urls_to_remove)} post(s) de '{SHEET_POSTS}' (hashtag não encontrada na description)...")
        print(f"{'=' * 60}")
        sheets_service = get_google_services()
        remove_rows_from_hashtag_posts(sheets_service, set(urls_to_remove))

    # ── ETAPA 4 — Preencher description em Hashtag_posts (apenas posts processados nesta rodada) ───
    print(f"\n{'=' * 60}")
    print(f"[ETAPA 4] Preenchendo description em '{SHEET_POSTS}' (posts desta rodada)...")
    print(f"{'=' * 60}")
    # share_url está no índice 1 e description no índice 9 de cada linha (ver DETAIL_HEADER, após adição de type_post)
    processed_descriptions = [(row[1], row[9]) for row in new_detail_rows]
    sheets_service = get_google_services()
    fill_descriptions_in_posts(sheets_service, processed_descriptions)

    print(f"\n{'=' * 60}")
    print(f"PIPELINE FINALIZADO")
    print(f"  Posts novos salvos em Hashtag_posts:           {len(new_post_rows)}")
    print(f"  Posts não salvos (fora da América Central/Sul): {total_filtered_by_country}")
    print(f"  Detalhes novos salvos em Hashtag_posts_detail: {len(new_detail_rows)}")
    print(f"  Posts pulados (aweme_id já existente):         {skipped_duplicate_aweme_id}")
    print(f"  Posts removidos (hashtag fora da description):{len(urls_to_remove)}")
    print(f"  Erros no video-info:                           {errors}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()

