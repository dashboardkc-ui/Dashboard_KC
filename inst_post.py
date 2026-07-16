import os
import re
import json
import time
import requests
import pandas as pd
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

COMMENTS_LIMIT = 100
BATCH_SIZE = 20
POST_EXPIRY_DAYS = 14

# Spreadsheet IDs
SPREADSHEET_PROFILES_ID = "1VK7_oyA3boJaudPaAiwk7xYl6sxReed63eOYBP9ahxo"
SPREADSHEET_DATA_PROFILE_ID = "1TvqNwg2GeCpBRNKBw87-qcbsG1yK2BaI8UQEbbZgCBw"
SPREADSHEET_DATA_COMMENTS_ID = "1dD69AANExCtYQG0g3MX8q5Sv794J0ajkRJ2GLGhaqLI"

# Sheet names
SHEET_PROFILES = "instagram_profile"
SHEET_DATA_PROFILE = "data_profile_post"
SHEET_DATA_COMMENTS = "data_comments_post"

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
# ETAPA 1 — LER PERFIS E LINKS
# ==============================

def extract_shortcode_from_url(url):
    """
    Extrai o shortcode de uma URL no formato:
    https://www.instagram.com/p/SHORTCODE/
    Retorna None se a URL for inválida.
    """
    if not url or not isinstance(url, str):
        return None
    match = re.search(r"instagram\.com/(?:p|reel)/([A-Za-z0-9_\-]+)", url.strip())
    if match:
        return match.group(1)
    return None


def parse_date(date_str):
    """
    Tenta converter a string de data da planilha para datetime com timezone.
    Suporta formatos comuns do Google Sheets.
    """
    if not date_str or not isinstance(date_str, str):
        return None

    formats = [
        "%m/%d/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%m/%d/%Y",
        "%d/%m/%Y",
        "%Y-%m-%d",
    ]
    for fmt in formats:
        try:
            dt = datetime.strptime(date_str.strip(), fmt)
            return tz_br.localize(dt)
        except ValueError:
            continue
    return None


def is_post_expired_by_date(date_added):
    """Retorna True se o post foi adicionado há mais de 14 dias."""
    if not date_added:
        return False
    hoje = datetime.now(tz_br)
    return (hoje - date_added).days > POST_EXPIRY_DAYS


def read_profiles_and_links(sheets_service):
    """
    Lê a planilha instagram_profile e retorna lista de dicts com:
    - username, link_of_post, shortcode, date_added, plataform, country, type
    Loga erros para links inválidos e pula posts expirados.
    """
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_PROFILES_ID,
        range=f"{SHEET_PROFILES}!A:Z"
    ).execute()

    rows = result.get("values", [])
    if len(rows) <= 1:
        print("Nenhuma linha encontrada na planilha instagram_profile.")
        return []

    headers = [h.strip().lower() for h in rows[0]]
    print(f"Colunas encontradas: {headers}")

    # Mapeia colunas pelo nome (case-insensitive)
    col_map = {}
    for expected, variants in {
        "date": ["date"],
        "plataform": ["plataform", "platform"],
        "username": ["username"],
        "country": ["country"],
        "type": ["type"],
        "link_of_post": ["link of post", "link_of_post", "linkofpost"],
    }.items():
        for variant in variants:
            if variant in headers:
                col_map[expected] = headers.index(variant)
                break

    required = ["username", "link_of_post", "date"]
    for col in required:
        if col not in col_map:
            print(f"ERRO: Coluna obrigatória '{col}' não encontrada. Colunas disponíveis: {headers}")
            return []

    entries = []
    for i, row in enumerate(rows[1:], start=2):  # linha 2 em diante (1-indexed)
        def get_col(key, default=""):
            idx = col_map.get(key)
            if idx is None:
                return default
            return row[idx].strip() if len(row) > idx else default

        username = get_col("username")
        link = get_col("link_of_post")
        date_str = get_col("date")
        plataform = get_col("plataform", "Instagram")
        country = get_col("country")
        post_type = get_col("type")

        # Valida link
        shortcode = extract_shortcode_from_url(link)
        if not shortcode:
            print(f"  [LINHA {i}] Link inválido ou vazio para @{username}: '{link}' — pulando.")
            continue

        # Valida e parseia data
        date_added = parse_date(date_str)
        if not date_added:
            print(f"  [LINHA {i}] Data inválida para @{username} (link={link}): '{date_str}' — pulando.")
            continue

        # Verifica expiração
        if is_post_expired_by_date(date_added):
            print(f"  [LINHA {i}] Post expirado (>14 dias desde {date_str}) para @{username}: {link} — pulando.")
            continue

        entries.append({
            "username": username,
            "link_of_post": link,
            "shortcode": shortcode,
            "date_added": date_added,
            "plataform": plataform,
            "country": country,
            "type": post_type,
        })

    print(f"\n{len(entries)} post(s) válido(s) e dentro do prazo encontrado(s).")
    return entries


# ==============================
# ETAPA 2 — PERFIL
# ==============================

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
# ETAPA 3 — POST INFO (caption + comentários + métricas)
# ==============================

# Mapeamento de __typename para label legível
MEDIA_TYPE_MAP = {
    "XDTGraphImage": "Image",
    "XDTGraphVideo": "Video",
    "XDTGraphSidecar": "Carousel",
}

def fetch_post_info(shortcode):
    """
    Chama o endpoint /post-info e retorna:
    - caption (str): texto de descrição do post
    - comments (list): todos os comentários paginados
    - post_meta (dict): métricas e metadados do post
    """
    url = "https://api.sociavault.com/v1/scrape/instagram/post-info"
    headers = {"X-API-Key": SOCIA_API_KEY}

    all_comments = []
    caption = ""
    post_meta = {}
    cursor = None
    page = 1
    seen_ids = set()  # controle de IDs já vistos para detectar páginas duplicadas

    while True:
        post_url_param = f"https://www.instagram.com/p/{shortcode}/"
        params = {"url": post_url_param}
        if cursor:
            params["cursor"] = cursor

        response = requests.get(url, params=params, headers=headers, timeout=API_TIMEOUT)

        # 404 durante paginação = cursor expirado, trata como fim dos comentários
        if response.status_code == 404 and page > 1:
            print(f"    Página {page}: cursor expirado (404), encerrando paginação.")
            break

        response.raise_for_status()
        data = response.json()

        media = (
            data.get("data", {})
                .get("data", {})
                .get("xdt_shortcode_media", {})
        )

        # Extrai metadados apenas na primeira página
        if page == 1:
            # Caption
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
            caption = first.get("node", {}).get("text", "")

            # Métricas e metadados do post
            typename = media.get("__typename", "")
            taken_at_raw = media.get("taken_at_timestamp")
            taken_at = (
                datetime.fromtimestamp(taken_at_raw, tz=tz_br).strftime("%Y-%m-%d %H:%M:%S")
                if taken_at_raw else ""
            )

            owner = media.get("owner", {})
            username_shared = owner.get("username", "")

            like_count = media.get("edge_media_preview_like", {}).get("count", "")
            comment_count = media.get("edge_media_preview_comment", {}).get("count", "")
            play_count = media.get("video_play_count", "")

            preview_image_url = media.get("thumbnail_src", "") or media.get("display_url", "")

            display_resources = media.get("display_resources", {})
            first_frame_url = ""
            if isinstance(display_resources, dict) and display_resources:
                last_key = str(max(int(k) for k in display_resources.keys()))
                first_frame_url = display_resources.get(last_key, {}).get("src", "")
            elif isinstance(display_resources, list) and display_resources:
                first_frame_url = display_resources[-1].get("src", "")

            post_meta = {
                "username_shared": username_shared,
                "taken_at": taken_at,
                "media_type": MEDIA_TYPE_MAP.get(typename, typename),
                "like_count": like_count,
                "comment_count": comment_count,
                "play_count": play_count,
                "preview_image_url": preview_image_url,
                "first_frame_url": first_frame_url,
            }

        # Extrai comentários
        comment_data = media.get("edge_media_to_parent_comment", {})
        edges = comment_data.get("edges", {})

        if isinstance(edges, dict):
            comment_nodes = [v.get("node", {}) for v in edges.values()]
        elif isinstance(edges, list):
            comment_nodes = [item.get("node", {}) for item in edges]
        else:
            comment_nodes = []

        # Detecta página duplicada: se todos os IDs já foram vistos, para imediatamente
        page_ids = {str(node.get("id", "")) for node in comment_nodes if node.get("id")}
        if page_ids and page_ids.issubset(seen_ids):
            print(f"    Página {page}: todos os IDs já vistos (API retornou página duplicada), encerrando paginação.")
            break
        seen_ids.update(page_ids)

        page_comments = normalize_comments(comment_nodes, page)
        all_comments.extend(page_comments)
        print(f"    Página {page}: {len(page_comments)} comentários")

        # Para de paginar se já atingiu o limite
        if len(all_comments) >= COMMENTS_LIMIT:
            print(f"    Limite de {COMMENTS_LIMIT} comentários atingido, encerrando paginação.")
            all_comments = all_comments[:COMMENTS_LIMIT]
            break

        # Paginação
        page_info = comment_data.get("page_info", {})
        has_next = page_info.get("has_next_page", False)

        if not has_next:
            break

        raw_cursor = page_info.get("end_cursor")
        if raw_cursor:
            try:
                cursor_obj = json.loads(raw_cursor)
                cursor = cursor_obj.get("server_cursor", raw_cursor)
            except (json.JSONDecodeError, TypeError):
                cursor = raw_cursor
        else:
            break

        page += 1
        time.sleep(1)

    return caption, all_comments, post_meta


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

        user = item.get("owner", {})
        rows.append({
            "post_url": post_url,
            "perfil": perfil,
            "Id Comentário": item.get("_custom_comment_id"),
            "id": comment_id,
            "text": item.get("text"),
            "comment_like_count": item.get("edge_liked_by", {}).get("count"),
            "child_comment_count": item.get("edge_threaded_comments", {}).get("count", 0),
            "created_at": item.get("created_at"),
            "user": json.dumps(user, ensure_ascii=False),
            "username": user.get("username"),
            "id_user": user.get("id"),
            "is_unpublished": item.get("is_restricted_pending"),
            "pk": user.get("id"),
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
1. Existe comentário neutro, então caso você acredite que não seja nem detrator e nem promotor pode usar essa classificação.
2. Se o comentário for positivo, elogio ou neutro-positivo (ex: "ok", "gostei", emojis), classifique como 'promotor'.
3. Se houver qualquer reclamação, dúvida técnica, ironia ou crítica, classifique como 'detrator'.

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
                            "enum": ["promotor", "detrator"]
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
                print(f"    Rate limit atingido. Aguardando {wait_seconds:.0f}s (tentativa {tentativa}/{max_tentativas})...")
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


# ==============================
# SALVAMENTO
# ==============================

def save_post_snapshot_to_sheets(sheets_service, post_entry, caption, post_meta, profile_data, run_datetime):
    """Salva snapshot do post (com todas as métricas) no data_profile_post."""
    row_data = {
        "run_datetime": run_datetime,
        "Plataform": post_entry.get("plataform", "Instagram"),
        "username": post_entry.get("username", ""),
        "username_shared": post_meta.get("username_shared", ""),
        "followers_count": profile_data.get("followers_count", ""),
        "following_count": profile_data.get("following_count", ""),
        "total_posts_count": profile_data.get("total_posts_count", ""),
        "code": post_entry.get("shortcode", ""),
        "taken_at": post_meta.get("taken_at", ""),
        "url": post_entry.get("link_of_post", ""),
        "media_type": post_meta.get("media_type", ""),
        "comment_count": post_meta.get("comment_count", ""),
        "like_count": post_meta.get("like_count", ""),
        "play_count": post_meta.get("play_count", ""),
        "preview_image_url": post_meta.get("preview_image_url", ""),
        "first_frame_url": post_meta.get("first_frame_url", ""),
        "post_caption": caption,
        "first_extracted_at": run_datetime,
    }

    df = pd.DataFrame([row_data])
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
        print(f"  data_profile_post: 1 linha inserida com cabeçalho.")
    else:
        append_values = df.astype(str).values.tolist()
        sheets_service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_DATA_PROFILE_ID,
            range=f"{SHEET_DATA_PROFILE}!A:A",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": append_values}
        ).execute()
        print(f"  data_profile_post: snapshot salvo ({run_datetime}).")


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
    print("INICIANDO PIPELINE INSTAGRAM (MODO LINKS)")
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

    # ETAPA 1 — Ler perfis e links
    print("\n[ETAPA 1] Lendo perfis e links de posts...")
    entries = read_profiles_and_links(sheets_service)

    if not entries:
        print("Nenhum post válido para processar. Encerrando.")
        return

    # ETAPA 3 — Para cada post: busca post-info (caption + comentários + métricas)
    print(f"\n[ETAPA 3] Processando {len(entries)} post(s)...")

    for entry in entries:
        shortcode = entry["shortcode"]
        post_url = entry["link_of_post"]
        username = entry["username"]
        run_datetime = datetime.now(tz_br).strftime("%Y-%m-%d %H:%M:%S")

        print(f"\n{'=' * 60}")
        print(f"POST: {post_url} (@{username})")
        print(f"{'=' * 60}")

        try:
            # Busca dados do perfil
            print(f"  Buscando perfil de @{username}...")
            profile_data = fetch_profile(username)

            # Busca caption, comentários e métricas do post
            caption, all_comments, post_meta = fetch_post_info(shortcode)

            if caption:
                print(f"  Caption: {caption[:100]}{'...' if len(caption) > 100 else ''}")
            else:
                print(f"  Caption: (vazia)")

            print(f"  username_shared: {post_meta.get('username_shared')} | "
                  f"media_type: {post_meta.get('media_type')} | "
                  f"likes: {post_meta.get('like_count')} | "
                  f"comments: {post_meta.get('comment_count')} | "
                  f"plays: {post_meta.get('play_count')}")

            # Salva snapshot do post no data_profile_post
            _, sheets_service = get_google_services()  # reconecta para evitar timeout
            save_post_snapshot_to_sheets(sheets_service, entry, caption, post_meta, profile_data, run_datetime)

            # Processa comentários
            saved_ids = get_saved_comment_ids(sheets_service, post_url)
            comments_to_classify = all_comments[-COMMENTS_LIMIT:] if len(all_comments) > COMMENTS_LIMIT else all_comments
            df_comments = comments_to_dataframe(comments_to_classify, post_url, username, saved_ids)

            if df_comments.empty:
                print("  Nenhum comentário novo. Pulando classificação.")
            else:
                df_comments = classificar_dataframe(df_comments)
                df_comments["data_execucao"] = run_datetime
                save_comments_to_sheets(sheets_service, df_comments)

        except Exception as e:
            print(f"  ERRO ao processar post {post_url}: {e}. Pulando.")
            continue

    print(f"\n{'=' * 60}")
    print("PIPELINE FINALIZADO COM SUCESSO")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
