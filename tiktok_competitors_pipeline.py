import os
import json
import time
import requests
import pandas as pd
from datetime import datetime, timezone, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ==============================
# CONFIG
# ==============================
SOCIAVAULT_API_KEY = os.environ.get("SOCIAVAULT_API_KEY", "")
GDRIVE_CREDENTIALS = os.environ.get("GDRIVE_CREDENTIALS", "")

SHEET_TIKTOK_PROFILE_ID   = "1cn68TA8_ajbbIOaMofE_7-Vc4_BWfQRMHehrO6SB_Q4"
SHEET_TT_DATA_POST_ID     = "1cn68TA8_ajbbIOaMofE_7-Vc4_BWfQRMHehrO6SB_Q4"

TAB_TIKTOK_PROFILE   = "tt_competitors_data"
TAB_TT_DATA_POST     = "Hashtag_posts_detail"

API_BASE         = "https://api.sociavault.com/v1/scrape/tiktok"
MAX_POSTS        = 10
POST_MAX_DAYS    = 14
PROFILE_REFRESH_DAYS = 30  # só reprocessa um perfil se já passaram esses dias desde o último run

# ==============================
# GOOGLE SHEETS HELPERS
# ==============================

def get_google_service():
    creds_json = json.loads(GDRIVE_CREDENTIALS)
    creds = service_account.Credentials.from_service_account_info(
        creds_json,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return build("sheets", "v4", credentials=creds)


def read_sheet(service, spreadsheet_id, tab):
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{tab}!A1:ZZ"
    ).execute()
    values = result.get("values", [])
    if not values:
        return pd.DataFrame()
    headers = values[0]
    rows = values[1:]
    rows = [r + [""] * (len(headers) - len(r)) for r in rows]
    return pd.DataFrame(rows, columns=headers)


def append_to_sheet(service, spreadsheet_id, tab, df):
    if df.empty:
        return
    values = df.values.tolist()
    service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=f"{tab}!A1",
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body={"values": values}
    ).execute()


def ensure_header(service, spreadsheet_id, tab, columns):
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{tab}!A1:1"
    ).execute()
    existing = result.get("values", [])
    if not existing:
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{tab}!A1",
            valueInputOption="RAW",
            body={"values": [columns]}
        ).execute()


def epoch_to_datetime_str(epoch_value):
    """Converte um epoch (segundos) para string 'YYYY-MM-DD HH:MM:SS' em UTC.
    Retorna string vazia se o valor não for válido."""
    try:
        epoch_int = int(epoch_value)
        if epoch_int <= 0:
            return ""
        return datetime.fromtimestamp(epoch_int, tz=timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    except (ValueError, TypeError, OverflowError):
        return ""


def get_last_run_by_profile(service):
    """Lê a aba de posts e retorna um dicionário {username: último run_datetime (datetime)}."""
    df = read_sheet(service, SHEET_TT_DATA_POST_ID, TAB_TT_DATA_POST)
    last_run = {}
    if df.empty or "username" not in df.columns or "run_datetime" not in df.columns:
        return last_run

    for _, row in df.iterrows():
        username = str(row.get("username", "")).strip()
        run_str  = str(row.get("run_datetime", "")).strip()
        if not username or not run_str:
            continue
        try:
            run_dt = datetime.strptime(run_str, "%Y-%m-%d %H:%M:%S").replace(tzinfo=timezone.utc)
        except ValueError:
            continue
        if username not in last_run or run_dt > last_run[username]:
            last_run[username] = run_dt

    return last_run


def deve_processar_perfil(username, last_run_map):
    """Retorna True se o perfil nunca foi processado ou se já passaram
    PROFILE_REFRESH_DAYS dias desde a última execução."""
    last_run = last_run_map.get(username)
    if last_run is None:
        return True
    elapsed = datetime.now(timezone.utc) - last_run
    return elapsed >= timedelta(days=PROFILE_REFRESH_DAYS)

# ==============================
# SOCIAVAULT HELPERS
# ==============================

def sv_get(endpoint, params, timeout=60):
    headers = {"X-API-Key": SOCIAVAULT_API_KEY}
    resp = requests.get(
        f"{API_BASE}/{endpoint}",
        headers=headers,
        params=params,
        timeout=timeout
    )
    resp.raise_for_status()
    return resp.json()

# ==============================
# ETAPA 1 — LER PERFIS
# ==============================

def ler_perfis(service):
    print("[ETAPA 1] Lendo perfis do tt_competitors_data...", flush=True)
    df = read_sheet(service, SHEET_TIKTOK_PROFILE_ID, TAB_TIKTOK_PROFILE)

    if df.empty:
        print("  Nenhum perfil encontrado.", flush=True)
        return []

    df.columns = [c.strip().lower() for c in df.columns]

    if "username" not in df.columns:
        print(f"  Coluna 'Username' não encontrada. Colunas disponíveis: {list(df.columns)}", flush=True)
        return []

    cols_to_keep = ["username"]
    if "type" in df.columns:
        cols_to_keep.append("type")
    if "country" in df.columns:
        cols_to_keep.append("country")

    perfis = (
        df[cols_to_keep]
        .rename(columns={"username": "profile"})
        .dropna(subset=["profile"])
        .drop_duplicates(subset=["profile"])  # ← evita processar o mesmo perfil mais de uma vez
        .to_dict("records")
    )
    perfis = [p for p in perfis if p["profile"].strip()]
    for p in perfis:
        p.setdefault("type", "")
        p.setdefault("country", "")

    print(f"  {len(perfis)} perfil(is) encontrado(s).", flush=True)
    return perfis

# ==============================
# ETAPA 2 — VÍDEOS / POSTS
# ==============================

POST_COLS = [
    "type_post", "video_url", "hashtag", "country", "marca_kc", "Competidor", "Pais",
    "run_datetime", "video_id", "description", "create_time",
    "author", "username", "followers",
    "play_count", "digg_count", "comment_count", "share_count",
    "collect_count", "download_count", "repost_count"
]

def buscar_video_info(video_url, video_id):
    try:
        data = sv_get("video-info", {"url": video_url})
        aweme = data.get("data", {}).get("aweme_detail", {})
        stats = aweme.get("statistics", {})
        return {
            "video_region":   aweme.get("region", ""),
            "digg_count":     stats.get("digg_count", ""),
            "comment_count":  stats.get("comment_count", ""),
            "share_count":    stats.get("share_count", ""),
            "play_count":     stats.get("play_count", ""),
            "collect_count":  stats.get("collect_count", ""),
            "download_count": stats.get("download_count", ""),
            "repost_count":   stats.get("repost_count", ""),
        }
    except Exception as e:
        print(f"      Erro ao buscar video-info de {video_id}: {e}", flush=True)
        return {
            "video_region": "", "digg_count": "", "comment_count": "",
            "share_count": "", "play_count": "", "collect_count": "",
            "download_count": "", "repost_count": ""
        }


def processar_videos(service, username, type_val="", country_val=""):
    print(f"  [2] Buscando vídeos de: {username}", flush=True)
    try:
        data = sv_get("videos", {"handle": username, "limit": MAX_POSTS})
    except Exception as e:
        print(f"    Erro ao buscar vídeos de {username}: {e}", flush=True)
        return []

    raw_list = None
    if isinstance(data, list):
        raw_list = data
    else:
        inner = data.get("data", data)
        aweme_list = inner.get("aweme_list", None)
        if aweme_list is not None:
            if isinstance(aweme_list, dict):
                raw_list = list(aweme_list.values())
            else:
                raw_list = aweme_list
        else:
            raw_list = inner.get("videos", inner.get("items", []))

    videos = raw_list[:MAX_POSTS] if raw_list else []

    if not videos:
        print(f"    Nenhum vídeo encontrado para {username}.", flush=True)
        return []

    ensure_header(service, SHEET_TT_DATA_POST_ID, TAB_TT_DATA_POST, POST_COLS)

    existing_df = read_sheet(service, SHEET_TT_DATA_POST_ID, TAB_TT_DATA_POST)
    existing_ids = (
        set(existing_df["aweme_id"].astype(str).tolist())
        if not existing_df.empty and "aweme_id" in existing_df.columns
        else set()
    )

    run_datetime_str = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    novos = []

    for v in videos:
        video_id = str(v.get("aweme_id", v.get("video_id", v.get("id", ""))))
        video_url = f"https://www.tiktok.com/@{username}/video/{video_id}"

        if video_id in existing_ids:
            continue

        author_obj = v.get("author", {})
        if isinstance(author_obj, dict):
            author_name    = author_obj.get("nickname", "")
            follower_count = author_obj.get("follower_count", "")
        else:
            author_name    = str(author_obj)
            follower_count = v.get("followers", "")

        # create_time vem como epoch (segundos) e representa a data de PUBLICAÇÃO do vídeo
        raw_create_time = v.get("create_time", v.get("createTime", ""))
        create_time_str = epoch_to_datetime_str(raw_create_time)

        print(f"      Buscando video-info para {video_id}...", flush=True)
        video_info = buscar_video_info(video_url, video_id)

        row = {
            "type_post":      "Competitors",
            "video_url":      video_url,
            "hashtag":        "",
            "country":        country_val,
            "marca_kc":       "",
            "Competidor":     username,
            "Pais":           country_val,
            "run_datetime":   run_datetime_str,   # data/hora em que o pipeline rodou
            "video_id":       video_id,
            "description":    v.get("desc", v.get("description", "")),
            "create_time":    create_time_str,    # data de publicação do post, já formatada
            "author":         author_name,
            "username":       username,
            "followers":      follower_count,
            "play_count":     video_info["play_count"],
            "digg_count":     video_info["digg_count"],
            "comment_count":  video_info["comment_count"],
            "share_count":    video_info["share_count"],
            "collect_count":  video_info["collect_count"],
            "download_count": video_info["download_count"],
            "repost_count":   video_info["repost_count"],
        }
        novos.append(row)
        existing_ids.add(video_id)

    if novos:
        df_new = pd.DataFrame(novos)[POST_COLS]
        append_to_sheet(service, SHEET_TT_DATA_POST_ID, TAB_TT_DATA_POST, df_new)
        print(f"    {len(novos)} vídeo(s) novo(s) salvos para {username}.", flush=True)
    else:
        print(f"    Nenhum vídeo novo para {username}.", flush=True)

    all_df = read_sheet(service, SHEET_TT_DATA_POST_ID, TAB_TT_DATA_POST)
    ids_perfil = [str(v.get("aweme_id", v.get("video_id", v.get("id", "")))) for v in videos]
    if not all_df.empty and "video_id" in all_df.columns:
        return all_df[all_df["video_id"].isin(ids_perfil)].to_dict("records")
    return []

# ==============================
# MAIN
# ==============================

def main():
    print("=== TikTok Pipeline (Competitors) - Posts only ===", flush=True)

    print(f"SOCIAVAULT_API_KEY: {'OK' if SOCIAVAULT_API_KEY else 'FALTANDO'}", flush=True)
    print(f"GDRIVE_CREDENTIALS: {'OK' if GDRIVE_CREDENTIALS else 'FALTANDO'}", flush=True)

    if not all([SOCIAVAULT_API_KEY, GDRIVE_CREDENTIALS]):
        print("ERRO: Variáveis de ambiente faltando. Abortando.", flush=True)
        return

    print("[INIT] Autenticando no Google Sheets...", flush=True)
    service = get_google_service()

    perfis = ler_perfis(service)
    if not perfis:
        return

    print("[INIT] Verificando últimas execuções por perfil...", flush=True)
    last_run_map = get_last_run_by_profile(service)

    for perfil in perfis:
        username = perfil["profile"].lstrip("@")
        print(f"\n{'='*40}", flush=True)
        print(f"PERFIL: @{username}", flush=True)
        print(f"{'='*40}", flush=True)

        if not deve_processar_perfil(username, last_run_map):
            last_run = last_run_map.get(username)
            print(f"  Pulando @{username}: último run em {last_run} (< {PROFILE_REFRESH_DAYS} dias).", flush=True)
            continue

        try:
            posts = processar_videos(
                service,
                username,
                type_val=perfil.get("type", ""),
                country_val=perfil.get("country", "")
            )
        except Exception as e:
            print(f"  Erro em 2 para {username}: {e}. Pulando.", flush=True)
            continue

        if not posts:
            print(f"  Sem posts processados para {username}.", flush=True)
            continue

    print("\n=== Pipeline finalizado ===", flush=True)


if __name__ == "__main__":
    main()
