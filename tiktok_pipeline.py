import os
import re
import json
import time
import requests
import pandas as pd
from datetime import datetime, timezone
from google import genai
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ==============================
# CONFIG
# ==============================
SOCIAVAULT_API_KEY = os.environ.get("SOCIAVAULT_API_KEY", "")
GEMINI_API_KEY     = os.environ.get("GEMINI_API_KEY", "")
GDRIVE_CREDENTIALS = os.environ.get("GDRIVE_CREDENTIALS", "")

SHEET_TIKTOK_PROFILE_ID   = "1947Wx86ZtNWQSaqcYVSXv_3WLvIA0p6u_Ol1DZ8GmX8"
SHEET_TT_DATA_PROFILE_ID  = "1roDSHeO9-O_DKfTwUKAQv3euCUyfipKq_KxkpKpf3r4"
SHEET_TT_DATA_POST_ID     = "1o96u5EXkqhtxGdEqaUGYX4Us2HGnHfkVBLJIobqcma8"
SHEET_TT_DATA_COMMENTS_ID = "1shH8-PpUBTEuS7Izy4uTgmEcOHF-tdk_DbJR1ifXqJA"

TAB_TIKTOK_PROFILE   = "tiktok_profile"
TAB_TT_DATA_PROFILE  = "tt_data_profile"
TAB_TT_DATA_POST     = "tt_data_post"
TAB_TT_DATA_COMMENTS = "tt_data_comments"

API_BASE         = "https://api.sociavault.com/v1/scrape/tiktok"
MAX_POSTS        = 5
POST_MAX_DAYS    = 14
GEMINI_BATCH     = 20
GEMINI_MAX_RETRY = 2
COMMENTS_LIMIT   = 100

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
# GEMINI HELPERS
# ==============================

def extrair_retry_seconds(error_str):
    match = re.search(r"retry in ([0-9.]+)s", error_str)
    if match:
        return float(match.group(1)) + 2
    return 60.0


def classify_comments_batch(client, comments_text):
    prompt = (
        "Você é um analista de redes sociais. Classifique cada comentário abaixo como "
        "'promotor' (positivo, elogio, apoio) ou 'detrator' (negativo, crítica, reclamação).\n"
        "Para cada comentário, retorne um JSON com os campos 'classification' e 'classification_reason'.\n"
        "Retorne APENAS uma lista JSON, sem markdown, sem texto extra.\n\n"
        "Comentários:\n"
    )
    for i, text in enumerate(comments_text):
        prompt += f"{i+1}. {text}\n"

    for attempt in range(1, GEMINI_MAX_RETRY + 1):
        try:
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt
            )
            raw = response.text.strip()
            raw = re.sub(r"^```json|^```|```$", "", raw, flags=re.MULTILINE).strip()
            return json.loads(raw)
        except Exception as e:
            err_str = str(e)
            if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                wait = extrair_retry_seconds(err_str)
                if attempt < GEMINI_MAX_RETRY:
                    print(f"    Rate limit atingido. Aguardando {wait:.0f}s antes de tentar novamente (tentativa {attempt}/{GEMINI_MAX_RETRY})...", flush=True)
                    time.sleep(wait)
                else:
                    print(f"    Rate limit após {GEMINI_MAX_RETRY} tentativas. Marcando lote como FALHA_API.", flush=True)
                    return [{"classification": "FALHA_API", "classification_reason": "rate limit"} for _ in comments_text]
            else:
                print(f"    Erro no Gemini: {e}", flush=True)
                return [{"classification": "ERRO", "classification_reason": str(e)} for _ in comments_text]

# ==============================
# ETAPA 1 — LER PERFIS
# ==============================

def ler_perfis(service):
    print("[ETAPA 1] Lendo perfis do tiktok_profile...", flush=True)
    df = read_sheet(service, SHEET_TIKTOK_PROFILE_ID, TAB_TIKTOK_PROFILE)

    if df.empty:
        print("  Nenhum perfil encontrado.", flush=True)
        return []

    df.columns = [c.strip().lower() for c in df.columns]

    if "username" not in df.columns:
        print(f"  Coluna 'Username' não encontrada. Colunas disponíveis: {list(df.columns)}", flush=True)
        return []

    perfis = (
        df[["username"]]
        .rename(columns={"username": "profile"})
        .dropna(subset=["profile"])
        .assign(date_added="")
        .drop_duplicates(subset=["profile"])  # ← CORREÇÃO: evita processar o mesmo perfil mais de uma vez
        .to_dict("records")
    )
    perfis = [p for p in perfis if p["profile"].strip()]

    print(f"  {len(perfis)} perfil(is) encontrado(s).", flush=True)
    return perfis

# ==============================
# ETAPA 2.0 — DADOS DO PERFIL
# ==============================

PROFILE_COLS = [
    "user_id", "username", "nickname", "verified",
    "followers", "following", "likes", "videos",
    "bio", "language", "is_organization", "run_datetime"
]

def processar_perfil(service, username):
    print(f"  [2.0] Buscando dados do perfil: {username}", flush=True)
    try:
        data = sv_get("profile", {"handle": username})
    except Exception as e:
        print(f"    Erro ao buscar perfil {username}: {e}", flush=True)
        return None

    ensure_header(service, SHEET_TT_DATA_PROFILE_ID, TAB_TT_DATA_PROFILE, PROFILE_COLS)

    inner = data.get("data", data)
    user  = inner.get("user", {})
    stats = inner.get("statsV2", inner.get("stats", {}))

    row = {
        "user_id":         str(user.get("id", "")),
        "username":        user.get("uniqueId", username),
        "nickname":        user.get("nickname", ""),
        "verified":        user.get("verified", ""),
        "followers":       stats.get("followerCount", ""),
        "following":       stats.get("followingCount", ""),
        "likes":           stats.get("heartCount", ""),
        "videos":          stats.get("videoCount", ""),
        "bio":             user.get("signature", ""),
        "language":        user.get("language", ""),
        "is_organization": user.get("isOrganization", ""),
        "run_datetime":    datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    }
    df_row = pd.DataFrame([row])[PROFILE_COLS]
    append_to_sheet(service, SHEET_TT_DATA_PROFILE_ID, TAB_TT_DATA_PROFILE, df_row)

    real_handle = user.get("uniqueId", username)
    print(f"    Perfil {username} salvo no tt_data_profile. Handle real: {real_handle}", flush=True)

    return data, real_handle

# ==============================
# ETAPA 2.1 — VÍDEOS / POSTS
# ==============================

POST_COLS = [
    "video_id", "description", "create_time", "author",
    "username", "followers", "likes", "comments",
    "views", "shares", "first_extracted_at", "video_url",
    "digg_count", "comment_count", "share_count", "play_count",
    "collect_count", "download_count", "whatsapp_share_count",
    "forward_count", "repost_count"
]

def buscar_video_info(video_url, video_id):
    try:
        data = sv_get("video-info", {"url": video_url})
        aweme = data.get("data", {}).get("aweme_detail", {})
        stats = aweme.get("statistics", {})
        return {
            "digg_count":           stats.get("digg_count", ""),
            "comment_count":        stats.get("comment_count", ""),
            "share_count":          stats.get("share_count", ""),
            "play_count":           stats.get("play_count", ""),
            "collect_count":        stats.get("collect_count", ""),
            "download_count":       stats.get("download_count", ""),
            "whatsapp_share_count": stats.get("whatsapp_share_count", ""),
            "forward_count":        stats.get("forward_count", ""),
            "repost_count":         stats.get("repost_count", ""),
        }
    except Exception as e:
        print(f"      Erro ao buscar video-info de {video_id}: {e}", flush=True)
        return {
            "digg_count": "", "comment_count": "", "share_count": "",
            "play_count": "", "collect_count": "", "download_count": "",
            "whatsapp_share_count": "", "forward_count": "", "repost_count": ""
        }


def processar_videos(service, username):
    print(f"  [2.1] Buscando vídeos de: {username}", flush=True)
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
        set(existing_df["video_id"].astype(str).tolist())
        if not existing_df.empty and "video_id" in existing_df.columns
        else set()
    )

    now_str = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    novos = []

    for v in videos:
        video_id = str(v.get("aweme_id", v.get("video_id", v.get("id", ""))))
        if video_id in existing_ids:
            continue

        author_obj = v.get("author", {})
        if isinstance(author_obj, dict):
            author_name    = author_obj.get("nickname", "")
            follower_count = author_obj.get("follower_count", "")
        else:
            author_name    = str(author_obj)
            follower_count = v.get("followers", "")

        stats    = v.get("statistics", {})
        likes    = stats.get("digg_count", v.get("likes", ""))
        comments = stats.get("comment_count", v.get("comments", ""))
        views    = stats.get("play_count", v.get("views", ""))
        shares   = stats.get("share_count", v.get("shares", ""))

        video_url = f"https://www.tiktok.com/@{username}/video/{video_id}"

        print(f"      Buscando video-info para {video_id}...", flush=True)
        video_info = buscar_video_info(video_url, video_id)

        row = {
            "video_id":             video_id,
            "description":          v.get("desc", v.get("description", "")),
            "create_time":          v.get("create_time", v.get("createTime", "")),
            "author":               author_name,
            "username":             username,
            "followers":            follower_count,
            "likes":                likes,
            "comments":             comments,
            "views":                views,
            "shares":               shares,
            "first_extracted_at":   now_str,
            "video_url":            video_url,
            "digg_count":           video_info["digg_count"],
            "comment_count":        video_info["comment_count"],
            "share_count":          video_info["share_count"],
            "play_count":           video_info["play_count"],
            "collect_count":        video_info["collect_count"],
            "download_count":       video_info["download_count"],
            "whatsapp_share_count": video_info["whatsapp_share_count"],
            "forward_count":        video_info["forward_count"],
            "repost_count":         video_info["repost_count"],
        }
        novos.append(row)

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
# ETAPA 2.2 — COMENTÁRIOS
# ==============================

COMMENT_COLS = [
    "comment_id", "video_id", "video_url", "text", "create_time",
    "likes", "replies_count", "purchase_intent",
    "user_name", "username", "language",
    "classification", "classification_reason"
]

def processar_comentarios(service, client, post):
    video_id  = str(post.get("video_id", ""))
    video_url = post.get("video_url", "")

    print(f"    [2.2] Buscando comentários do vídeo: {video_url}", flush=True)

    existing_df = read_sheet(service, SHEET_TT_DATA_COMMENTS_ID, TAB_TT_DATA_COMMENTS)
    existing_ids = (
        set(existing_df["comment_id"].astype(str).tolist())
        if not existing_df.empty and "comment_id" in existing_df.columns
        else set()
    )

    ensure_header(service, SHEET_TT_DATA_COMMENTS_ID, TAB_TT_DATA_COMMENTS, COMMENT_COLS)

    novos  = []
    cursor = None
    pagina = 1

    while len(novos) < COMMENTS_LIMIT:
        params = {"url": video_url}
        if cursor is not None:
            params["cursor"] = cursor

        try:
            data = sv_get("comments", params)
        except Exception as e:
            print(f"      Erro ao buscar comentários (página {pagina}) do vídeo {video_id}: {e}", flush=True)
            break

        inner = data.get("data", data)
        raw   = inner.get("comments", {})
        if isinstance(raw, dict):
            comments = list(raw.values())
        elif isinstance(raw, list):
            comments = raw
        else:
            comments = []

        print(f"      Página {pagina}: {len(comments)} comentários recebidos.", flush=True)

        novos_pagina = [
            c for c in comments
            if str(c.get("cid", c.get("comment_id", c.get("id", "")))) not in existing_ids
        ]
        novos.extend(novos_pagina)

        has_more = inner.get("has_more", 0)
        cursor   = inner.get("cursor", None)
        pagina  += 1

        if not has_more or cursor is None:
            break

        time.sleep(1)

    if not novos:
        print(f"      Sem comentários novos para vídeo {video_id}.", flush=True)
        return

    if len(novos) > COMMENTS_LIMIT:
        print(f"      Limitando de {len(novos)} para {COMMENTS_LIMIT} comentários.", flush=True)
        novos = novos[:COMMENTS_LIMIT]

    print(f"      {len(novos)} comentário(s) novo(s) para classificar.", flush=True)

    all_rows = []
    for i in range(0, len(novos), GEMINI_BATCH):
        lote   = novos[i:i + GEMINI_BATCH]
        textos = [c.get("text", c.get("comment", "")) for c in lote]
        print(f"      Classificando lote {i // GEMINI_BATCH + 1}...", flush=True)
        classificacoes = classify_comments_batch(client, textos)

        for j, c in enumerate(lote):
            clf    = classificacoes[j] if j < len(classificacoes) else {"classification": "ERRO", "classification_reason": "sem resposta"}
            c_user = c.get("user", {})
            row = {
                "comment_id":            str(c.get("cid", c.get("comment_id", c.get("id", "")))),
                "video_id":              video_id,
                "video_url":             video_url,
                "text":                  c.get("text", ""),
                "create_time":           c.get("create_time", ""),
                "likes":                 c.get("digg_count", c.get("likes", "")),
                "replies_count":         c.get("reply_comment_total", c.get("replies_count", "")),
                "purchase_intent":       c.get("is_high_purchase_intent", ""),
                "user_name":             c_user.get("nickname", c.get("user_name", "")),
                "username":              c_user.get("unique_id", c.get("username", "")),
                "language":              c.get("comment_language", c.get("language", "")),
                "classification":        clf.get("classification", ""),
                "classification_reason": clf.get("classification_reason", "")
            }
            all_rows.append(row)
        time.sleep(2)

    if all_rows:
        df_comments = pd.DataFrame(all_rows)[COMMENT_COLS]
        append_to_sheet(service, SHEET_TT_DATA_COMMENTS_ID, TAB_TT_DATA_COMMENTS, df_comments)
        print(f"      {len(all_rows)} comentário(s) salvo(s) para vídeo {video_id}.", flush=True)

# ==============================
# MAIN
# ==============================

def main():
    print("=== TikTok Pipeline ===", flush=True)

    print(f"SOCIAVAULT_API_KEY: {'OK' if SOCIAVAULT_API_KEY else 'FALTANDO'}", flush=True)
    print(f"GEMINI_API_KEY:     {'OK' if GEMINI_API_KEY else 'FALTANDO'}", flush=True)
    print(f"GDRIVE_CREDENTIALS: {'OK' if GDRIVE_CREDENTIALS else 'FALTANDO'}", flush=True)

    if not all([SOCIAVAULT_API_KEY, GEMINI_API_KEY, GDRIVE_CREDENTIALS]):
        print("ERRO: Variáveis de ambiente faltando. Abortando.", flush=True)
        return

    print("[INIT] Autenticando no Google Sheets...", flush=True)
    service = get_google_service()

    print("[INIT] Inicializando cliente Gemini...", flush=True)
    client = genai.Client(api_key=GEMINI_API_KEY)

    perfis = ler_perfis(service)
    if not perfis:
        return

    for perfil in perfis:
        username = perfil["profile"].lstrip("@")
        print(f"\n{'='*40}", flush=True)
        print(f"PERFIL: @{username}", flush=True)
        print(f"{'='*40}", flush=True)

        try:
            result = processar_perfil(service, username)
            if result is None:
                print(f"  Perfil {username} não retornou dados. Pulando.", flush=True)
                continue
            _, real_handle = result
        except Exception as e:
            print(f"  Erro em 2.0 para {username}: {e}. Pulando.", flush=True)
            continue

        try:
            posts = processar_videos(service, real_handle)
        except Exception as e:
            print(f"  Erro em 2.1 para {real_handle}: {e}. Pulando.", flush=True)
            continue

        if not posts:
            print(f"  Sem posts para processar comentários de {real_handle}.", flush=True)
            continue

        for post in posts:
            try:
                processar_comentarios(service, client, post)
            except Exception as e:
                print(f"  Erro em 2.2 para vídeo {post.get('video_id', '?')}: {e}. Pulando.", flush=True)
                continue

    print("\n=== Pipeline finalizado ===", flush=True)


if __name__ == "__main__":
    main()
