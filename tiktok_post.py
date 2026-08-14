import os
import re
import json
import time
import requests
import pandas as pd
from datetime import datetime, timezone
from google import genai
from google.genai import types
from google.oauth2 import service_account
from googleapiclient.discovery import build
# ==============================
# CONFIG
# ==============================
SOCIAVAULT_API_KEY = os.environ.get("SOCIAVAULT_API_KEY", "")
GEMINI_API_KEY     = os.environ.get("GEMINI_API_KEY_KC", "")
GDRIVE_CREDENTIALS = os.environ.get("GDRIVE_CREDENTIALS_KC", "")
SHEET_INPUT_ID            = "1947Wx86ZtNWQSaqcYVSXv_3WLvIA0p6u_Ol1DZ8GmX8"
SHEET_TT_DATA_COMMENTS_ID = "1BD4OoVfXZHI6p5kJ6KmLAMsPfpQ86MjdNdVoPPWhgkg"
SHEET_TT_DATA_POST_ID     = "1CtvNfYM5Jp_kuriycsYMAMzCQYW0pFxqvGmOD0O4n80"
TAB_INPUT            = "tiktok_profile"
TAB_TT_DATA_COMMENTS = "tt_data_comments_post"
TAB_TT_DATA_POST     = "tt_data_post_post"
TAB_TT_DATA_POST_MAX = "tt_data_post_post_max"
API_BASE         = "https://api.sociavault.com/v1/scrape/tiktok"
POST_MAX_DAYS    = 14
GEMINI_BATCH     = 20
GEMINI_MAX_RETRY = 3
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
def ts_to_datetime(value):
    """Converte um unix timestamp (segundos) em string datetime UTC.
    Retorna '' se o valor não for um número válido (ex.: vazio, '#VALUE!')."""
    if value is None:
        return ""
    s = str(value).strip()
    if not s:
        return ""
    try:
        ts = float(s)
    except ValueError:
        return ""
    if ts <= 0:
        return ""
    try:
        return datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    except (ValueError, OverflowError, OSError):
        return ""
 
 
def _update_row(service, spreadsheet_id, tab, row_number, values_list):
    """Sobrescreve UMA linha inteira (a partir da coluna A) da aba destino."""
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{tab}!A{row_number}",
        valueInputOption="RAW",
        body={"values": [values_list]}
    ).execute()
 
 
def _align_row(series, columns):
    """Devolve os valores da Series na ordem exata de `columns` (como strings)."""
    out = []
    for c in columns:
        v = series.get(c, "")
        out.append("" if (v is None or pd.isna(v)) else str(v))
    return out
 
 
def atualizar_post_max(service):
    """ETAPA 3 — mantém em tt_data_post_post_max a versão mais recente de cada post.
 
    1. Lê tt_data_post_post.
    2. Preenche 'Published date' com create_time (unix ts) em formato datetime.
    3. Para cada aweme_id mantém só a linha com run_datetime mais recente.
    4. Faz upsert em tt_data_post_post_max por aweme_id:
         - se já existe e o run_datetime da origem é MAIOR -> substitui a linha;
         - se não existe -> anexa como nova linha.
    """
    print("\n[ETAPA 3] Atualizando tab de máximos (tt_data_post_post_max)...", flush=True)
 
    # --- 1. Ler origem -------------------------------------------------------
    df_src = read_sheet(service, SHEET_TT_DATA_POST_ID, TAB_TT_DATA_POST)
    if df_src.empty:
        print("  tt_data_post_post vazia. Nada a fazer.", flush=True)
        return
    if "aweme_id" not in df_src.columns or "run_datetime" not in df_src.columns:
        print(f"  ERRO: colunas obrigatórias ausentes em tt_data_post_post. "
              f"Colunas: {list(df_src.columns)}", flush=True)
        return
 
    # --- 2. Preencher 'Published date' a partir de create_time ---------------
    if "create_time" not in df_src.columns:
        print("  ERRO: coluna 'create_time' não encontrada em tt_data_post_post.", flush=True)
        return
    df_src["Published date"] = df_src["create_time"].apply(ts_to_datetime)
 
    # --- 3. Manter, por aweme_id, a linha com run_datetime mais recente ------
    df_src["_run_dt"] = pd.to_datetime(df_src["run_datetime"], errors="coerce")
    # na_position='first' garante que uma data inválida (NaT) nunca "vença"
    # uma data válida ao usar keep='last'.
    df_src = (
        df_src.sort_values("_run_dt", na_position="first")
              .drop_duplicates(subset="aweme_id", keep="last")
    )
    print(f"  {len(df_src)} aweme_id(s) únicos após deduplicação.", flush=True)
 
    # --- 4. Ler destino (máximos) e garantir cabeçalho -----------------------
    df_max = read_sheet(service, SHEET_TT_DATA_POST_ID, TAB_TT_DATA_POST_MAX)
    if df_max.empty or not list(df_max.columns):
        # aba vazia: usa como cabeçalho as colunas da origem (menos o helper)
        max_cols = [c for c in df_src.columns if c != "_run_dt"]
        ensure_header(service, SHEET_TT_DATA_POST_ID, TAB_TT_DATA_POST_MAX, max_cols)
        df_max = pd.DataFrame(columns=max_cols)
    else:
        max_cols = list(df_max.columns)
 
    # Índice aweme_id -> posição da linha (para saber o nº da linha na planilha)
    max_pos = {}
    max_rundt = {}
    if "aweme_id" in df_max.columns:
        rundt_series = pd.to_datetime(df_max.get("run_datetime"), errors="coerce")
        for pos, (_, r) in enumerate(df_max.iterrows()):
            aid = str(r["aweme_id"]).strip()
            if aid:
                max_pos[aid] = pos              # 0-based dentro dos dados
                max_rundt[aid] = rundt_series.iloc[pos]
 
    # --- 4a/4b. Upsert -------------------------------------------------------
    novos = []          # linhas a anexar
    atualizados = 0
    for _, srow in df_src.iterrows():
        aid = str(srow["aweme_id"]).strip()
        if not aid:
            continue
        values = _align_row(srow, max_cols)
 
        if aid in max_pos:
            # já existe -> só atualiza se a origem for mais recente
            src_dt = srow["_run_dt"]
            dst_dt = max_rundt.get(aid)
            mais_recente = (
                pd.notna(src_dt) and (pd.isna(dst_dt) or src_dt > dst_dt)
            )
            if mais_recente:
                sheet_row = max_pos[aid] + 2      # +1 header, +1 base-1
                _update_row(service, SHEET_TT_DATA_POST_ID,
                            TAB_TT_DATA_POST_MAX, sheet_row, values)
                atualizados += 1
        else:
            # não existe -> anexa
            novos.append(values)
 
    if novos:
        df_novos = pd.DataFrame(novos, columns=max_cols)
        append_to_sheet(service, SHEET_TT_DATA_POST_ID, TAB_TT_DATA_POST_MAX, df_novos)
 
    print(f"  {atualizados} linha(s) atualizada(s) e {len(novos)} nova(s) anexada(s) "
          f"em tt_data_post_post_max.", flush=True)

# ==============================
# ETAPA 4 — NORMALIZAR TIPOS EM tt_data_post_post_max
# ==============================

# Aceita os dois nomes possíveis (código x planilha) — só processa os que existirem.
NUM_COLS_MAX = {
    "likes", "digg_count", "comment_count", "share_count",
    "views", "play_count", "saves", "collect_count",
    "download_count", "whatsapp_share_count", "forward_count",
    "repost_count", "create_time",
}
DT_COLS_MAX = {"run_datetime", "Published date"}


def _col_letter(idx):
    """0-based -> letra de coluna A1 (A, B, ..., Z, AA...)."""
    letters = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _to_number(v):
    s = str(v).strip()
    if s == "" or s.lower() in ("nan", "none"):
        return ""
    try:
        f = float(s)
        return int(f) if f.is_integer() else f
    except ValueError:
        return s  # deixa como está se não for número


def _to_datetime_str(v):
    s = str(v).strip()
    if not s:
        return ""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%m/%d/%Y %H:%M:%S", "%d/%m/%Y %H:%M:%S",
                "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            continue
    return s  # deixa como está se não parsear


def normalizar_tipos_post_max(service):
    """ETAPA 4 — relê tt_data_post_post_max e garante:
       - colunas numéricas gravadas como número;
       - run_datetime / Published date gravadas como datetime.
       Só reescreve as colunas afetadas (IDs e texto ficam intactos)."""
    print("\n[ETAPA 4] Normalizando tipos em tt_data_post_post_max...", flush=True)

    df = read_sheet(service, SHEET_TT_DATA_POST_ID, TAB_TT_DATA_POST_MAX)
    if df.empty:
        print("  Tab vazia. Nada a fazer.", flush=True)
        return

    n_rows = len(df)
    for idx, col in enumerate(df.columns):
        if col in NUM_COLS_MAX:
            conv, opt = _to_number, "RAW"            # número nativo -> fica número
        elif col in DT_COLS_MAX:
            conv, opt = _to_datetime_str, "USER_ENTERED"  # string ISO -> Sheets vira data
        else:
            continue

        col_values = [[conv(v)] for v in df[col].tolist()]
        letter = _col_letter(idx)
        service.spreadsheets().values().update(
            spreadsheetId=SHEET_TT_DATA_POST_ID,
            range=f"{TAB_TT_DATA_POST_MAX}!{letter}2:{letter}{n_rows + 1}",
            valueInputOption=opt,
            body={"values": col_values}
        ).execute()
        print(f"  Coluna '{col}' normalizada ({opt}).", flush=True)

    print("  Tipos normalizados.", flush=True)

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
    try:
        return resp.json()
    except ValueError:
        raise ValueError(
            f"SociaVault /{endpoint} retornou corpo não-JSON "
            f"(status={resp.status_code}, content-type={resp.headers.get('Content-Type')}, "
            f"body[:200]={resp.text[:200]!r})"
        )

# ==============================
#  HELPERS
# ==============================
def extrair_retry_seconds(error_str):
    match = re.search(r"retry in ([0-9.]+)s", error_str)
    if match:
        return float(match.group(1)) + 2
    return 60.0
def _extract_response_text(response):
    """FIX 2: nunca chamar json.loads direto em response.text.
    Valida se o  realmente retornou conteúdo e, se não, explica o porquê."""
    text = (response.text or "").strip() if response is not None else ""
    if text:
        return text
    finish_reason = None
    block_reason = None
    if getattr(response, "candidates", None):
        finish_reason = getattr(response.candidates[0], "finish_reason", None)
    if getattr(response, "prompt_feedback", None):
        block_reason = getattr(response.prompt_feedback, "block_reason", None)
    raise ValueError(
        f" retornou resposta vazia (finish_reason={finish_reason}, "
        f"block_reason={block_reason})"
    )


def classify_comments_batch(client, comments):
    """Recebe lista de dicts {"n": int, "text": str}.
    FIX 4: envia IDs explícitos (como no script 1) em vez de depender da ordem."""
    prompt = f"""Você é um especialista em análise de sentimentos para redes sociais.
Sua tarefa é classificar comentários em 'promotor', 'neutro' ou 'detrator'.

REGRAS CRÍTICAS:
1. Existe comentário neutro, então caso você acredite que não seja nem detrator e nem promotor pode usar essa classificação.
2. Se o comentário for positivo, elogio ou neutro-positivo (ex: "ok", "gostei", emojis), classifique como 'promotor'.
3. Se houver qualquer reclamação, dúvida técnica, ironia ou crítica, classifique como 'detrator'.
4. Retorne exatamente um resultado por comentário, com o mesmo "n" recebido.

Comentários para análise:
{json.dumps(comments, ensure_ascii=False)}
"""

    schema = {
        "type": "object",
        "properties": {
            "results": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "n": {"type": "integer"},
                        "classification": {
                            "type": "string",
                            "enum": ["promotor", "neutro", "detrator"]
                        },
                        "classification_reason": {"type": "string"}
                    },
                    "required": ["n", "classification", "classification_reason"]
                }
            }
        },
        "required": ["results"]
    }

    last_err = None
    for attempt in range(1, GEMINI_MAX_RETRY + 1):
        try:
            response = client.models.generate_content(
                model="gemini-3.1-flash-lite",
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_json_schema=schema,
                    temperature=0.1,
                    # FIX 5: gemini-2.5-flash é modelo "thinking" — sem limite,
                    # ele pode gastar todo o budget pensando e devolver texto
                    # vazio => "Expecting value: line 1 column 1 (char 0)".
                    max_output_tokens=8192,
                    thinking_config=types.ThinkingConfig(thinking_budget=0),
                )
            )
            raw = _extract_response_text(response)
            results = json.loads(raw)["results"]
            # Reindexa por "n" para garantir alinhamento
            return {int(r["n"]): r for r in results}

        except Exception as e:
            err_str = str(e)
            last_err = err_str
            if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                wait = extrair_retry_seconds(err_str)
                if attempt < GEMINI_MAX_RETRY:
                    print(f"    Rate limit atingido. Aguardando {wait:.0f}s (tentativa {attempt}/{GEMINI_MAX_RETRY})...", flush=True)
                    time.sleep(wait)
                    continue
                print(f"    Rate limit após {GEMINI_MAX_RETRY} tentativas. Marcando lote como FALHA_API.", flush=True)
                return {c["n"]: {"n": c["n"], "classification": "FALHA_API", "classification_reason": "rate limit"} for c in comments}
            # FIX 6: resposta vazia / JSON truncado agora também é retryable
            if isinstance(e, (ValueError, json.JSONDecodeError, KeyError)) and attempt < GEMINI_MAX_RETRY:
                print(f"    Resposta inválida do Gemini ({e}). Repetindo em 5s (tentativa {attempt}/{GEMINI_MAX_RETRY})...", flush=True)
                time.sleep(5)
                continue
            print(f"    Erro no Gemini: {e}", flush=True)
            return {c["n"]: {"n": c["n"], "classification": "ERRO", "classification_reason": err_str} for c in comments}

    return {c["n"]: {"n": c["n"], "classification": "ERRO", "classification_reason": last_err or "sem resposta"} for c in comments}
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
    total_examinados = 0  # FIX: conta comentários trazidos da API (novos ou não), não só os novos
    cursor = None
    pagina = 1
    while total_examinados < COMMENTS_LIMIT:
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
        if not comments:
            break
        total_examinados += len(comments)  # FIX: para de paginar com base no total examinado
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
        lote = novos[i:i + GEMINI_BATCH]
        # FIX 4: manda id explícito por comentário, igual à estratégia do script 1
        payload = [
            {"n": j, "text": str(c.get("text", c.get("comment", "")))}
            for j, c in enumerate(lote)
        ]
        print(f"    Classificando lote {i // GEMINI_BATCH + 1}...", flush=True)
        classificacoes = classify_comments_batch(client, payload)  # dict {n: resultado}

        for j, c in enumerate(lote):
            clf    = classificacoes.get(j, {"classification": "ERRO", "classification_reason": "sem resposta"})
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
            existing_ids.add(cid)

        time.sleep(2)

    if all_rows:
        df_comments = pd.DataFrame(all_rows)[COMMENT_COLS]
        df_comments = df_comments.fillna("").astype(str)
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
        try:
            atualizar_post_max(service)
        except Exception as e:
            print(f"  Erro na ETAPA 3 (post_max): {e}", flush=True)

    
            
    print(f"\n=== Pipeline finalizado. Total de comentários salvos: {total_salvos} ===", flush=True)
    # ETAPA 4 — Normalizar tipos na versão final de tt_data_post_post_max
    try:
        normalizar_tipos_post_max(service)
    except Exception as e:
        print(f"  Erro na ETAPA 4 (normalizar tipos): {e}", flush=True)
    
if __name__ == "__main__":
    main()
