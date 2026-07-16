import os
import re
import json
import time
import requests
import pandas as pd
import sys
import builtins

from datetime import datetime
from google import genai
from google.genai import types
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
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")

POSTS_LIMIT = 10
COMMENTS_LIMIT = 100
BATCH_SIZE = 20
PROFILE_REFRESH_DAYS = 30

# Spreadsheet IDs (planilha única unificada)
SPREADSHEET_PROFILES_ID = "1sMcUNkDVtuhL51f2BVVwuSim45d8rfiug5Sb6Hcb-Qk"
SPREADSHEET_DATA_PROFILE_ID = "1sMcUNkDVtuhL51f2BVVwuSim45d8rfiug5Sb6Hcb-Qk"
SPREADSHEET_DATA_COMMENTS_ID = "1sMcUNkDVtuhL51f2BVVwuSim45d8rfiug5Sb6Hcb-Qk"

# Sheet names
SHEET_PROFILES = "instagram_competitors_data"
SHEET_DATA_PROFILE = "instagram_competitors_data_profile"
SHEET_DATA_COMMENTS = "instagram_competitors_data_comments"

tz_br = pytz.timezone("America/Sao_Paulo")


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
    drive_service = build("drive", "v3", credentials=creds)
    sheets_service = build("sheets", "v4", credentials=creds)
    return drive_service, sheets_service


# ==============================
# ETAPA 1 — LER PERFIS
# ==============================

def read_profiles(sheets_service):
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_PROFILES_ID,
        range=f"{SHEET_PROFILES}!A:F"
    ).execute()

    rows = result.get("values", [])
    if len(rows) <= 1:
        print("Nenhum perfil encontrado na planilha instagram_profile.")
        return []

    headers = [h.strip().lower() for h in rows[0]]

    if "username" not in headers:
        print(f"Coluna 'Username' não encontrada. Colunas disponíveis: {headers}")
        return []

    username_col = headers.index("username")
    profiles = []

    for row in rows[1:]:
        if len(row) > username_col:
            username = row[username_col].strip()
            if username:
                profiles.append({"profile": username, "date_added": ""})

    print(f"{len(profiles)} perfil(is) encontrado(s) na planilha.")

    # FIX: Deduplicação de perfis — evita reprocessar o mesmo handle múltiplas vezes
    seen = set()
    profiles_unique = []
    for p in profiles:
        if p["profile"] not in seen:
            seen.add(p["profile"])
            profiles_unique.append(p)

    print(f"Após deduplicação: {len(profiles_unique)} perfil(is) único(s): {[p['profile'] for p in profiles_unique]}")
    return profiles_unique


def get_last_run_by_profile(sheets_service):
    """Retorna dict {username: last_run_datetime (datetime)} com o último run_datetime
    salvo para cada perfil na aba data_profile."""
    try:
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_DATA_PROFILE_ID,
            range=f"{SHEET_DATA_PROFILE}!A:Z"
        ).execute()
        rows = result.get("values", [])
        if len(rows) <= 1:
            return {}
        headers = rows[0]
        if "username" not in headers or "run_datetime" not in headers:
            return {}
        username_col = headers.index("username")
        run_col = headers.index("run_datetime")

        last_run = {}
        for row in rows[1:]:
            if len(row) > max(username_col, run_col):
                username = row[username_col].strip()
                run_str = row[run_col].strip()
                if not username or not run_str:
                    continue
                try:
                    run_dt = tz_br.localize(datetime.strptime(run_str, "%Y-%m-%d %H:%M:%S"))
                except Exception:
                    continue
                if username not in last_run or run_dt > last_run[username]:
                    last_run[username] = run_dt
        return last_run
    except Exception as e:
        print(f"  Aviso ao ler último run por perfil: {e}")
        return {}


def should_skip_profile(handle, last_run_by_profile):
    """Retorna True se o perfil já foi processado há menos de PROFILE_REFRESH_DAYS dias."""
    last_run = last_run_by_profile.get(handle)
    if not last_run:
        return False
    dias_desde_ultimo_run = (datetime.now(tz_br) - last_run).days
    if dias_desde_ultimo_run < PROFILE_REFRESH_DAYS:
        print(f"  Perfil @{handle} processado há {dias_desde_ultimo_run} dia(s) (< {PROFILE_REFRESH_DAYS}), pulando.")
        return True
    return False


def fetch_profile(handle):
    url = "https://api.sociavault.com/v1/scrape/instagram/profile"
    headers = {"X-API-Key": SOCIA_API_KEY}
    params = {"handle": handle}

    response = requests.get(url, headers=headers, params=params, timeout=API_TIMEOUT)
    print(f"  Status profile ({handle}): {response.status_code}")
    response.raise_for_status()

    data = response.json()

    user = (
    data.get("data", {})
        .get("data", {})
        .get("user")
    or data.get("user")
    or {}
    )

    return {
        "username": user.get("username", handle),
        "followers_count": user.get("edge_followed_by", {}).get("count", ""),
        "following_count": user.get("edge_follow", {}).get("count", ""),
        "total_posts_count": user.get("edge_owner_to_timeline_media", {}).get("count", "")
    }


# ==============================
# ETAPA 2 — EXTRAIR POSTS
# ==============================

def fetch_posts(handle):
    api_key = SOCIA_API_KEY
    headers = {"X-API-Key": api_key}
    url = "https://api.sociavault.com/v1/scrape/instagram/posts"
    params = {"handle": handle, "limit": POSTS_LIMIT}

    response = requests.get(url, params=params, headers=headers, timeout=API_TIMEOUT)
    print(f"  Status posts ({handle}): {response.status_code}")
    response.raise_for_status()
    json_data = response.json()

    items = json_data.get("data", {}).get("items", {})
    profile_data = fetch_profile(handle)
    top_username = profile_data.get("username")

    run_datetime = datetime.now(tz_br).strftime("%Y-%m-%d %H:%M:%S")

    rows = []
    if isinstance(items, dict):
        iterable = list(items.values())
    elif isinstance(items, list):
        iterable = items
    else:
        print(f"  Aviso: estrutura inesperada de 'items': {type(items)} — keys disponíveis: {list(json_data.get('data', {}).keys())}")
        iterable = []

    iterable = iterable[:POSTS_LIMIT]

    media_type_map = {
    1: "Image",
    2: "Video",
    8: "Carousel"
    }

    for item in iterable:
        username_shared = item.get("user", {}).get("username", top_username)
        code = item.get("code")
        taken_at_raw = item.get("taken_at")
        create_time = datetime.fromtimestamp(taken_at_raw, tz=tz_br).strftime("%Y-%m-%d %H:%M:%S") if taken_at_raw else ""
        post_url = item.get("url")

        if not post_url and code:
            post_url = f"https://www.instagram.com/p/{code}/"

        media_type = item.get("media_type")
        media_type_label = media_type_map.get(media_type, "Other")
        comment_count = item.get("comment_count")
        like_count = item.get("like_count")
        play_count = item.get("play_count")

        image_versions2 = item.get("image_versions2", {})
        candidates = image_versions2.get("candidates", {})
        preview_image_url = None

        if isinstance(candidates, dict):
            best_candidate = candidates.get("0") or next(iter(candidates.values()), {})
            preview_image_url = best_candidate.get("url")
        elif isinstance(candidates, list) and len(candidates) > 0:
            preview_image_url = candidates[0].get("url")

        additional_candidates = image_versions2.get("additional_candidates", {})
        first_frame_url = None
        if additional_candidates and isinstance(additional_candidates, dict):
            first_frame_url = additional_candidates.get("first_frame", {}).get("url")

        rows.append({
            "Plataform": "Instagram",
            "username": handle,
            "username_shared": username_shared,
            "followers_count": profile_data.get("followers_count"),
            "following_count": profile_data.get("following_count"),
            "total_posts_count": profile_data.get("total_posts_count"),
            "code": code,
            "create_time": create_time,
            "url": post_url,
            "media_type": media_type,
            "comment_count": comment_count,
            "like_count": like_count,
            "play_count": play_count,
            "preview_image_url": preview_image_url,
            "first_frame_url": first_frame_url,
            "post_caption": "",  # será preenchido na etapa 3
            "run_datetime": run_datetime
        })

    df = pd.DataFrame(rows)
    df = df.fillna("")
    print(f"  Posts extraídos: {len(df)}")
    return df


def save_posts_to_sheets(sheets_service, df):
    """Salva todos os posts sempre, criando snapshot histórico por run_datetime."""
    df = df.fillna("")

    existing_data = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_DATA_PROFILE_ID,
        range=f"{SHEET_DATA_PROFILE}!A:A"
    ).execute()
    existing_rows = existing_data.get("values", [])

    if not existing_rows:
        values = [df.columns.tolist()] + df.astype(str).values.tolist()
        sheets_service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_DATA_PROFILE_ID,
            range=f"{SHEET_DATA_PROFILE}!A1",
            valueInputOption="RAW",
            body={"values": values}
        ).execute()
        print(f"  data_profile: {len(df)} linhas inseridas com cabeçalho.")
    else:
        append_values = df.astype(str).values.tolist()
        sheets_service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_DATA_PROFILE_ID,
            range=f"{SHEET_DATA_PROFILE}!A:A",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": append_values}
        ).execute()
        print(f"  data_profile: {len(append_values)} linhas adicionadas (snapshot {df['run_datetime'].iloc[0]}).")


# ==============================
# ETAPA 3 — POST INFO (comentários + descrição)
# ==============================

def fetch_caption(shortcode):
    """
    Chama o endpoint /post-info apenas para extrair a legenda do post.
    Não usa esse endpoint para paginar comentários — o cursor que ele retorna
    (server_cursor / is_server_cursor_inverse) é a estrutura crua do GraphQL
    do Instagram e não funciona para buscar páginas seguintes neste endpoint.
    """
    url = "https://api.sociavault.com/v1/scrape/instagram/post-info"
    headers = {"X-API-Key": SOCIA_API_KEY}
    post_url_param = f"https://www.instagram.com/p/{shortcode}/"
    params = {"url": post_url_param}

    response = requests.get(url, params=params, headers=headers, timeout=API_TIMEOUT)
    response.raise_for_status()
    data = response.json()

    media = (
        data.get("data", {})
            .get("data", {})
            .get("xdt_shortcode_media", {})
    )

    caption_edges = (
        media.get("edge_media_to_caption", {})
             .get("edges", {})
    )
    if isinstance(caption_edges, dict):
        first = caption_edges.get("0", {})
    elif isinstance(caption_edges, list) and len(caption_edges) > 0:
        first = caption_edges[0]
    else:
        first = {}

    return first.get("node", {}).get("text", "")


def fetch_comments(shortcode):
    """
    Chama o endpoint dedicado /v1/scrape/instagram/comments, que é o endpoint
    feito pela SociaVault especificamente para paginar comentários (documentação:
    "Use the cursor parameter from the response to get the next page of comments").
    Retorna até COMMENTS_LIMIT comentários.
    """
    url = "https://api.sociavault.com/v1/scrape/instagram/comments"
    headers = {"X-API-Key": SOCIA_API_KEY}
    post_url_param = f"https://www.instagram.com/p/{shortcode}/"

    all_comments = []
    cursor = None
    page = 1
    seen_ids = set()  # controle de IDs já vistos para detectar páginas duplicadas

    while True:
        params = {"url": post_url_param}
        if cursor:
            params["cursor"] = cursor
            print(f"    [DEBUG] Enviando cursor na página {page}: {cursor}")

        response = requests.get(url, params=params, headers=headers, timeout=API_TIMEOUT)

        # 404 durante paginação = cursor expirado, trata como fim dos comentários
        if response.status_code == 404 and page > 1:
            print(f"    Página {page}: cursor expirado (404), encerrando paginação.")
            break

        response.raise_for_status()
        data = response.json().get("data", {})

        comments_raw = data.get("comments", {})
        if isinstance(comments_raw, dict):
            comment_nodes = list(comments_raw.values())
        elif isinstance(comments_raw, list):
            comment_nodes = comments_raw
        else:
            comment_nodes = []

        if not comment_nodes:
            print(f"    Página {page}: sem comentários, encerrando paginação.")
            break

        # Detecta página duplicada: se todos os IDs já foram vistos, para imediatamente
        page_ids = {str(c.get("id", "")) for c in comment_nodes if c.get("id")}
        if page_ids and page_ids.issubset(seen_ids):
            print(f"    Página {page}: todos os IDs já vistos (API retornou página duplicada), encerrando paginação.")
            break
        seen_ids.update(page_ids)

        page_comments = normalize_comments(comment_nodes, page)
        all_comments.extend(page_comments)

        next_cursor = data.get("cursor")
        print(f"    Página {page}: {len(page_comments)} comentários | next_cursor={'sim' if next_cursor else 'não'}")

        # Para de paginar se já atingiu o limite de comentários
        if len(all_comments) >= COMMENTS_LIMIT:
            print(f"    Limite de {COMMENTS_LIMIT} comentários atingido, encerrando paginação.")
            all_comments = all_comments[:COMMENTS_LIMIT]
            break

        if not next_cursor:
            print(f"    Sem cursor na resposta, encerrando paginação (fim dos comentários).")
            break

        cursor = next_cursor
        page += 1
        time.sleep(1)

    return all_comments


def normalize_comments(comment_nodes, page):
    comments = []
    for idx, node in enumerate(comment_nodes, start=1):
        node["_page"] = page
        node["_comment_number"] = idx
        node["_custom_comment_id"] = f"{page}_{idx}"
        comments.append(node)
    return comments


def has_emoji(text):
    emoji_pattern = re.compile(
        "["
        "\U0001F600-\U0001F64F"
        "\U0001F300-\U0001F5FF"
        "\U0001F680-\U0001F6FF"
        "\U0001F1E0-\U0001F1FF"
        "\U00002700-\U000027BF"
        "\U0001F900-\U0001F9FF"
        "\U00002600-\U000026FF"
        "]+",
        flags=re.UNICODE,
    )
    return bool(emoji_pattern.search(text))


def get_saved_comment_ids(sheets_service, post_url):
    try:
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_DATA_COMMENTS_ID,
            range=f"{SHEET_DATA_COMMENTS}!A:Z"
        ).execute()
        rows = result.get("values", [])
        if len(rows) <= 1:
            return set()
        headers = rows[0]
        if "id" not in headers or "post_url" not in headers:
            return set()
        id_col = headers.index("id")
        url_col = headers.index("post_url")
        saved_ids = set()
        for row in rows[1:]:
            if len(row) > max(id_col, url_col):
                if row[url_col].strip() == post_url.strip():
                    saved_ids.add(row[id_col].strip())
        print(f"    IDs já salvos para {post_url}: {len(saved_ids)}")
        return saved_ids
    except Exception as e:
        print(f"    Aviso ao ler data_comments: {e}")
        return set()


def comments_to_dataframe(comments, post_url, perfil, saved_ids):
    rows = []
    skipped = 0

    for item in comments:
        comment_id = str(item.get("id", ""))
        if comment_id in saved_ids:
            skipped += 1
            continue

        # Estrutura do endpoint /v1/scrape/instagram/comments: dados do autor
        # vêm em "user" (não mais em "owner"), e não há contagem de likes/replies
        # do comentário neste endpoint.
        user = item.get("user", {})
        rows.append({
            "post_url": post_url,
            "perfil": perfil,
            "Id Comentário": item.get("_custom_comment_id"),
            "id": comment_id,
            "text": item.get("text"),
            "comment_like_count": "",
            "child_comment_count": "",
            "created_at": item.get("created_at"),
            "user": json.dumps(user, ensure_ascii=False),
            "username": user.get("username"),
            "id_user": user.get("id"),
            "is_unpublished": user.get("is_unpublished"),
            "pk": user.get("pk"),
            "is_verified": user.get("is_verified")
        })

    print(f"    Comentários novos: {len(rows)} | Já salvos (ignorados): {skipped}")

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    df = df.fillna("")
    df["text"] = df["text"].astype(str)
    df["Id Comentário"] = df["Id Comentário"].astype(str)
    df["text_debug"] = df["text"].apply(repr)
    df["tem_emoji"] = df["text"].apply(has_emoji)

    return df


# ==============================
# CLASSIFICAÇÃO GEMINI
# ==============================

def extrair_retry_seconds(error_message):
    match = re.search(r"retry in ([0-9.]+)s", str(error_message))
    if match:
        return float(match.group(1)) + 2
    return 60


def classificar_lote_comentarios(comentarios, tentativa=1, max_tentativas=2):
    client = genai.Client(api_key=GEMINI_API_KEY)
    prompt = f"""
Você é um especialista em análise de sentimentos para redes sociais.
Sua tarefa é classificar comentários em 'promotor', 'neutro' ou 'detrator'.

REGRAS CRÍTICAS:
1. Se o comentário for claramente positivo, elogio, entusiasmo ou recomendação, classifique como 'promotor'.
2. Se houver qualquer reclamação, dúvida técnica, ironia ou crítica, classifique como 'detrator'.
3. Se o comentário for puramente informativo, ambíguo, irrelevante ao produto/marca, ou não expressar opinião clara (ex: apenas marcação de outro usuário, pergunta neutra sem tom negativo, comentário genérico tipo "ok"), classifique como 'neutro'.
4. Não force um comentário para 'promotor' ou 'detrator' apenas para evitar usar 'neutro' — use 'neutro' sempre que não houver sinal claro de sentimento positivo ou negativo.

Comentários para análise:
{json.dumps(comentarios, ensure_ascii=False)}
"""

    schema = {
        "type": "object",
        "properties": {
            "results": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "Id Comentário": {"type": "string"},
                        "sentimento_nps": {
                            "type": "string",
                            "enum": ["promotor", "neutro", "detrator"]
                        },
                        "justificativa": {"type": "string"}
                    },
                    "required": ["Id Comentário", "sentimento_nps", "justificativa"]
                }
            }
        },
        "required": ["results"]
    }

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
        return json.loads(response.text)["results"]

    except Exception as e:
        error_str = str(e)
        if "429" in error_str or "RESOURCE_EXHAUSTED" in error_str:
            wait_seconds = extrair_retry_seconds(error_str)
            if tentativa <= max_tentativas:
                print(f"    Rate limit atingido. Aguardando {wait_seconds:.0f}s antes de tentar novamente (tentativa {tentativa}/{max_tentativas})...")
                time.sleep(wait_seconds)
                return classificar_lote_comentarios(comentarios, tentativa=tentativa + 1, max_tentativas=max_tentativas)
            else:
                print(f"    Máximo de tentativas atingido para este lote.")
                raise
        raise


def classificar_dataframe(df):
    resultados = []
    print(f"    Classificando {len(df)} comentários...")

    for i in range(0, len(df), BATCH_SIZE):
        lote = df.iloc[i:i + BATCH_SIZE]
        comentarios_lote = [
            {"Id Comentário": row["Id Comentário"], "text": row["text"]}
            for _, row in lote.iterrows()
        ]

        try:
            classificados = classificar_lote_comentarios(comentarios_lote)
            resultados.extend(classificados)
            print(f"    Lote {i // BATCH_SIZE + 1} OK")
        except Exception as e:
            print(f"    Erro no lote {i // BATCH_SIZE + 1}: {e}")
            for item in comentarios_lote:
                resultados.append({
                    "Id Comentário": item["Id Comentário"],
                    "sentimento_nps": "FALHA_API",
                    "justificativa": str(e)
                })

        time.sleep(2)

    df_result = pd.DataFrame(resultados)
    df_result["Id Comentário"] = df_result["Id Comentário"].astype(str)
    df = df.drop(columns=["sentimento_nps", "justificativa"], errors="ignore")
    df = df.merge(df_result, on="Id Comentário", how="left")

    return df


def save_comments_to_sheets(sheets_service, df):
    df = df.fillna("")

    existing_data = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_DATA_COMMENTS_ID,
        range=f"{SHEET_DATA_COMMENTS}!A:A"
    ).execute()
    existing_rows = existing_data.get("values", [])

    if not existing_rows:
        values = [df.columns.tolist()] + df.astype(str).values.tolist()
        sheets_service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_DATA_COMMENTS_ID,
            range=f"{SHEET_DATA_COMMENTS}!A1",
            valueInputOption="RAW",
            body={"values": values}
        ).execute()
        print(f"    data_comments: {len(df)} linhas inseridas com cabeçalho.")
    else:
        append_values = df.astype(str).values.tolist()
        sheets_service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_DATA_COMMENTS_ID,
            range=f"{SHEET_DATA_COMMENTS}!A:A",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": append_values}
        ).execute()
        print(f"    data_comments: {len(append_values)} linhas adicionadas.")


# ==============================
# EXECUÇÃO PRINCIPAL
# ==============================

def main():
    print("=" * 60)
    print("INICIANDO PIPELINE INSTAGRAM")
    print("=" * 60)

    print("\n[CONFIG] Verificando variáveis de ambiente...")
    missing = []
    for var in ["SOCIAVAULT_API_KEY", "GEMINI_API_KEY", "GDRIVE_CREDENTIALS"]:
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
    drive_service, sheets_service = get_google_services()
    print("  Google Services OK")

    # ETAPA 1 — Ler perfis
    print("\n[ETAPA 1] Lendo perfis...")
    profiles = read_profiles(sheets_service)

    if not profiles:
        print("Nenhum perfil para processar. Encerrando.")
        return

    # Carrega último run_datetime conhecido por perfil (controle de 30 dias)
    last_run_by_profile = get_last_run_by_profile(sheets_service)
    print(f"Perfis com histórico de execução: {len(last_run_by_profile)}")

    for profile_entry in profiles:
        handle = profile_entry["profile"]

        if should_skip_profile(handle, last_run_by_profile):
            continue

        print(f"\n{'=' * 60}")
        print(f"PERFIL: {handle}")
        print(f"{'=' * 60}")

        # ETAPA 2 — Extrair posts
        print(f"\n[ETAPA 2] Extraindo posts de @{handle}...")
        try:
            df_posts = fetch_posts(handle)
        except Exception as e:
            print(f"  ERRO ao buscar posts de @{handle}: {e}. Pulando perfil.")
            continue

        if df_posts.empty:
            print(f"  Nenhum post encontrado para @{handle}. Pulando.")
            continue

        # ETAPA 3 — Para cada post: busca post-info (caption + comentários)
        print(f"\n[ETAPA 3] Processando post-info dos posts de @{handle}...")

        for idx, post_row in df_posts.iterrows():
            post_url = post_row.get("url", "")
            post_code = post_row.get("code", "")

            if not post_code:
                print(f"  Post sem code (url={post_url}), pulando.")
                continue

            if not post_url:
                post_url = f"https://www.instagram.com/p/{post_code}/"

            print(f"\n  Post: {post_url}")

            try:
                # Busca legenda (post-info) e comentários (endpoint dedicado de paginação)
                caption = fetch_caption(post_code)
                all_comments = fetch_comments(post_code)

                # Atualiza caption no dataframe de posts
                df_posts.at[idx, "post_caption"] = caption
                if caption:
                    print(f"    Caption extraída: {caption[:80]}{'...' if len(caption) > 80 else ''}")
                else:
                    print(f"    Caption: (vazia)")

                # Processa comentários
                saved_ids = get_saved_comment_ids(sheets_service, post_url)
                # Limita a classificação aos últimos 100 comentários novos
                comments_to_classify = all_comments[-COMMENTS_LIMIT:] if len(all_comments) > COMMENTS_LIMIT else all_comments
                df_comments = comments_to_dataframe(comments_to_classify, post_url, handle, saved_ids)

                if df_comments.empty:
                    print("    Nenhum comentário novo. Pulando classificação.")
                else:
                    df_comments = classificar_dataframe(df_comments)
                    df_comments["data_execucao"] = datetime.now(tz_br).strftime("%Y-%m-%d %H:%M:%S")
                    save_comments_to_sheets(sheets_service, df_comments)

            except Exception as e:
                print(f"    ERRO ao processar post-info de {post_url}: {e}. Pulando post.")
                continue

        # Reconecta Google Services antes de salvar (evita SSLEOFError por inatividade)
        _, sheets_service = get_google_services()

        # Salva posts (com captions preenchidas) no data_profile
        save_posts_to_sheets(sheets_service, df_posts)

    print(f"\n{'=' * 60}")
    print("PIPELINE FINALIZADO COM SUCESSO")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
