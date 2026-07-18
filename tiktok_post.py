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

SHEET_INPUT_ID            = "1947Wx86ZtNWQSaqcYVSXv_3WLvIA0p6u_Ol1DZ8GmX8"
SHEET_TT_DATA_COMMENTS_ID = "1BD4OoVfXZHI6p5kJ6KmLAMsPfpQ86MjdNdVoPPWhgkg"
SHEET_TT_DATA_POST_ID     = "1CtvNfYM5Jp_kuriycsYMAMzCQYW0pFxqvGmOD0O4n80"

TAB_INPUT            = "tiktok_profile"
TAB_TT_DATA_COMMENTS = "tt_data_comments_post"
TAB_TT_DATA_POST     = "tt_data_post_post"

API_BASE         = "https://api.sociavault.com/v1/scrape/tiktok"
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
    """Você é um especialista em análise de sentimentos para redes sociais.
    Sua tarefa é classificar comentários em 'promotor', 'neutro' ou 'detrator'.

    REGRAS CRÍTICAS:
    1. Existe comentário neutro, então caso você acredite que não seja nem detrator e nem promotor pode usar essa classificação.
    2. Se o comentário for positivo, elogio ou neutro-positivo (ex: "ok", "gostei", emojis), classifique como 'promotor'.
    3. Se houver qualquer reclamação, dúvida técnica, ironia ou crítica, classifique como 'detrator'.

    Comentários para análise:
    """
    )
    for i, text in enumerate(comments_text):
        prompt += f"{i+1}. {text}\n"

    # Criando o esquema estrito idêntico à estratégia do seu código que funciona
    schema = {
        "type": "object",
        "properties": {
            "results": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "classification": {
                            "type": "string",
                            "enum": ["promotor", "neutro", "detrator"]
                        },
                        "classification_reason": {"type": "string"}
                    },
                    "required": ["classification", "classification_reason"]
                }
            }
        },
        "required": ["results"]
    }

    for attempt in range(1, GEMINI_MAX_RETRY + 1):
        try:
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_json_schema=schema,
                    temperature=0.1
                )
            )
            # Retorna diretamente a lista extraída do objeto estruturado
            return json.loads(response.text)["results"]

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
                # Garante que o erro real seja capturado sem quebrar o pipeline do dataframe
                return [{"classification": "ERRO", "classification_reason": str(e)} for _ in comments_text]


# ETAPA 1 — LER POSTS DA PLANILHA
# ==============================

TIKTOK_VIDEO_URL_PATTERN = re.compile(
    r"https?://(?:www\.)?tiktok\.com/@[\w.]+/video/(\d+)"
)

def extrair_video_id(link):
    """Extrai o video_id de um link do TikTok. Retorna (video_id, erro)."""
    if not link or not link.strip():
        return None, "Link vazio"
    link = link.strip()
    match = TIKTOK_VIDEO_URL_PATTERN.match(link)
    if not match:
        return None, f"Link inválido ou fora do padrão TikTok: '{link}'"
    return match.group(1), None


def parse_date(date_str):
    """Tenta parsear a data da planilha. Retorna datetime com UTC ou None."""
    if not date_str or not date_str.strip():
        return None
    for fmt in ("%m/%d/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S"):
        try:
            return datetime.strptime(date_str.strip(), fmt).replace(tzinfo=timezone.utc)
        except ValueError:
            continue
    return None


def ler_posts(service):
    print("[ETAPA 1] Lendo posts da planilha de entrada...", flush=True)
    df = read_sheet(service, SHEET_INPUT_ID, TAB_INPUT)

    if df.empty:
        print("  Planilha vazia ou sem dados.", flush=True)
        return []

    # Normaliza cabeçalhos
    df.columns = [c.strip().lower() for c in df.columns]

    required_cols = {"date", "link of post"}
    missing = required_cols - set(df.columns)
    if missing:
        print(f"  ERRO: Colunas obrigatórias não encontradas: {missing}. Colunas disponíveis: {list(df.columns)}", flush=True)
        return []

    posts_validos = []
    now = datetime.now(timezone.utc)

    for i, row in df.iterrows():
        link      = str(row.get("link of post", "")).strip()
        date_str  = str(row.get("date", "")).strip()
        username  = str(row.get("username", "")).strip()
        plataform = str(row.get("plataform", "")).strip()

        # Extrai video_id
        video_id, erro = extrair_video_id(link)
        if erro:
            print(f"  [LINHA {i+2}] PULANDO — {erro}", flush=True)
            continue

        # Parseia a data
        post_date = parse_date(date_str)
        if post_date is None:
            print(f"  [LINHA {i+2}] PULANDO — Data inválida ou ausente: '{date_str}' (video_id={video_id})", flush=True)
            continue

        # Verifica janela de 14 dias
        dias = (now - post_date).days
        if dias > POST_MAX_DAYS:
            print(f"  [LINHA {i+2}] IGNORANDO — Post {video_id} de @{username} tem {dias} dias (limite: {POST_MAX_DAYS}).", flush=True)
            continue

        posts_validos.append({
            "video_id":  video_id,
            "video_url": link,
            "username":  username,
            "post_date": post_date,
            "dias":      dias,
        })

    print(f"  {len(posts_validos)} post(s) dentro da janela de {POST_MAX_DAYS} dias.", flush=True)
    return posts_validos

# ==============================
# ETAPA 2.1 — VIDEO INFO / STATISTICS
# ==============================

POST_COLS = [
    "video_url", "username", "run_datetime",
    "aweme_id", "digg_count", "comment_count", "share_count",
    "play_count", "collect_count", "download_count", "whatsapp_share_count",
    "forward_count", "repost_count", "desc", "create_time"
]

def processar_video_info(service, post):
    video_url = post["video_url"]
    video_id  = post["video_id"]
    username  = post["username"]  # FIX: obtido corretamente do dict post
    print(f"  [2.1] Buscando video-info: {video_url}", flush=True)

    try:
        data = sv_get("video-info", {"url": video_url})
    except Exception as e:
        print(f"    Erro ao buscar video-info de {video_id}: {e}", flush=True)
        return

    ensure_header(service, SHEET_TT_DATA_POST_ID, TAB_TT_DATA_POST, POST_COLS)

    aweme = data.get("data", {}).get("aweme_detail", {})
    stats = aweme.get("statistics", {})

    row = {
        "video_url":            video_url,
        "username":             username,                              # FIX: campo adicionado ao POST_COLS
        "run_datetime":         datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S"),
        "aweme_id":             stats.get("aweme_id", video_id),      # FIX: campo faltante adicionado
        "digg_count":           stats.get("digg_count", ""),          # FIX: chaves alinhadas com POST_COLS
        "comment_count":        stats.get("comment_count", ""),
        "share_count":          stats.get("share_count", ""),
        "play_count":           stats.get("play_count", ""),          # FIX: era "view"
        "collect_count":        stats.get("collect_count", ""),       # FIX: era "saves"
        "download_count":       stats.get("download_count", ""),
        "whatsapp_share_count": stats.get("whatsapp_share_count", ""),
        "forward_count":        stats.get("forward_count", ""),
        "repost_count":         stats.get("repost_count", ""),
        "desc":                 aweme.get("desc", ""),                # FIX: buscado em aweme, não em stats
        "create_time":          aweme.get("create_time", "")         # FIX: era v.get() — variável inexistente
    }

    df_row = pd.DataFrame([row])[POST_COLS]
    append_to_sheet(service, SHEET_TT_DATA_POST_ID, TAB_TT_DATA_POST, df_row)
    print(f"    video-info salvo para {video_id}.", flush=True)


# ==============================
# ETAPA 2.2 — COMENTÁRIOS
# ==============================

COMMENT_COLS = [
    "comment_id", "video_id", "text", "create_time",
    "likes", "replies_count", "purchase_intent",
    "user_name", "username", "language",
    "classification", "classification_reason", "video_url"
]

def processar_comentarios(service, client, post, existing_ids):
    video_id  = post["video_id"]
    video_url = post["video_url"]
    username  = post["username"]

    print(f"\n  [2.2] Buscando comentários do vídeo: {video_url}", flush=True)

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
            print(f"    Erro ao buscar comentários (página {pagina}) do vídeo {video_id}: {e}", flush=True)
            break

        inner = data.get("data", data)
        raw   = inner.get("comments", {})
        if isinstance(raw, dict):
            comments = list(raw.values())
        elif isinstance(raw, list):
            comments = raw
        else:
            comments = []

        print(f"    Página {pagina}: {len(comments)} comentários recebidos.", flush=True)

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
        print(f"    Sem comentários novos para vídeo {video_id}.", flush=True)
        return 0

    if len(novos) > COMMENTS_LIMIT:
        print(f"    Limitando de {len(novos)} para {COMMENTS_LIMIT} comentários.", flush=True)
        novos = novos[:COMMENTS_LIMIT]

    print(f"    {len(novos)} comentário(s) novo(s) para classificar.", flush=True)

    all_rows = []
    for i in range(0, len(novos), GEMINI_BATCH):
        lote   = novos[i:i + GEMINI_BATCH]
        textos = [c.get("text", c.get("comment", "")) for c in lote]
        print(f"    Classificando lote {i // GEMINI_BATCH + 1}...", flush=True)
        classificacoes = classify_comments_batch(client, textos)

        for j, c in enumerate(lote):
            clf    = classificacoes[j] if j < len(classificacoes) else {"classification": "ERRO", "classification_reason": "sem resposta"}
            c_user = c.get("user", {})
            cid    = str(c.get("cid", c.get("comment_id", c.get("id", ""))))

            row = {
                "comment_id":            cid,
                "video_id":              video_id,
                "text":                  c.get("text", ""),
                "create_time":           c.get("create_time", ""),
                "likes":                 c.get("digg_count", c.get("likes", "")),
                "replies_count":         c.get("reply_comment_total", c.get("replies_count", "")),
                "purchase_intent":       c.get("is_high_purchase_intent", ""),
                "user_name":             c_user.get("nickname", c.get("user_name", "")),
                "username":              c_user.get("unique_id", c.get("username", username)),
                "language":              c.get("comment_language", c.get("language", "")),
                "classification":        clf.get("classification", ""),
                "classification_reason": clf.get("classification_reason", ""),
                "video_url":             video_url
            }
            all_rows.append(row)
            existing_ids.add(cid)  # atualiza deduplicação em memória

        time.sleep(2)

    if all_rows:
        df_comments = pd.DataFrame(all_rows)[COMMENT_COLS]
        append_to_sheet(service, SHEET_TT_DATA_COMMENTS_ID, TAB_TT_DATA_COMMENTS, df_comments)
        print(f"    {len(all_rows)} comentário(s) salvo(s) para vídeo {video_id}.", flush=True)

    return len(all_rows)

# ==============================
# MAIN
# ==============================

def main():
    print("=== TikTok Comments Pipeline ===", flush=True)

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

    # Carrega IDs de comentários já salvos UMA VEZ para toda a execução
    print("[INIT] Carregando comment_ids já salvos...", flush=True)
    existing_df = read_sheet(service, SHEET_TT_DATA_COMMENTS_ID, TAB_TT_DATA_COMMENTS)
    existing_ids = (
        set(existing_df["comment_id"].astype(str).tolist())
        if not existing_df.empty and "comment_id" in existing_df.columns
        else set()
    )
    print(f"  {len(existing_ids)} comment_id(s) já existentes carregados.", flush=True)

    # ETAPA 1 — Ler posts válidos
    posts = ler_posts(service)
    if not posts:
        print("Nenhum post para processar. Encerrando.", flush=True)
        return

    total_salvos = 0

    for post in posts:
        print(f"\n{'='*40}", flush=True)
        print(f"POST: {post['video_url']} | @{post['username']} | {post['dias']} dia(s) desde publicação", flush=True)
        print(f"{'='*40}", flush=True)

        # ETAPA 2.1 — Video info + statistics
        try:
            processar_video_info(service, post)
        except Exception as e:
            print(f"  Erro em 2.1 para vídeo {post['video_id']}: {e}. Continuando para comentários.", flush=True)

        # ETAPA 2.2 — Comentários
        try:
            salvos = processar_comentarios(service, client, post, existing_ids)
            total_salvos += salvos
        except Exception as e:
            print(f"  Erro ao processar vídeo {post['video_id']}: {e}. Pulando.", flush=True)
            continue

    print(f"\n=== Pipeline finalizado. Total de comentários salvos: {total_salvos} ===", flush=True)


if __name__ == "__main__":
    main()
