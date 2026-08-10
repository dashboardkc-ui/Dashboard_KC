import os
import io
import re
import json
import sys
import builtins
from datetime import datetime
import numpy as np
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import pytz
# Força flush imediato em todos os prints
_original_print = builtins.print
def print(*args, **kwargs):
    kwargs["flush"] = True
    _original_print(*args, **kwargs)
tz_br = pytz.timezone("America/Sao_Paulo")
# ==============================
# CONFIGURAÇÕES
# ==============================
# --- Organic (Winclap) ---
DRIVE_FOLDER_ID = "1k8NDz3qxQ9ffZzkS2EsWT1tNSk3ghklU"
SPREADSHEET_ID = "1FnauIqLuTe1c2N8Z-HQPy8wambQzBhpbLJY24JMCMNY"  # planilha organic (SAÍDA etapa 1)
SHEET_NAME = "Hoja 1"
BASELINE_SPREADSHEET_ID = "1qMl_c5KgCDb0QCr4wOPnjnhdnmrHrDw3Fqs9T5F1wyg"
BASELINE_SHEET_RANGE = "A:Z"
SOURCE_TAB = "Winclap_Organic"
HEADER_SKIPROWS = 2
FILENAME_PATTERN = re.compile(r"^(\d{2})(\d{2})(\d{4})\.xlsx$")  # DDMMAAAA.xlsx
 
# ID orgânico extraído do permalink (mantido como coluna de dados).
KEY_COLUMN = "Organic_ID"
 
# Nome da coluna de permalink no relatório (mantido exatamente como vem no
# arquivo; repare que há um espaço inicial em " (EXTERNAL_VALUE)"). Se o
# cabeçalho real for "Permalink (EXTERNAL_VALUE)", basta trocar aqui.
PERMALINK_COLUMN = "Permalink (EXTERNAL_VALUE)"
 
# Chave COMPOSTA para identificar linhas únicas (deduplicação + upsert/update):
# Permalink + Country of Origin. É criada em read_organic_sheet e gravada na
# planilha como a coluna abaixo, para permitir o match nas próximas execuções.
UPSERT_KEY_COLUMN = "Unique_Key"
 
RAW_COLUMNS_NEEDED = [
    "Published Date",
    "Social Network",
    "Brand (Account)",
    "Account",
    "Country of Origin (Account)",
    "Outbound Post",
    "Outbound Post Id",
    "Outbound Message Category",
    PERMALINK_COLUMN,
    "Video Views (SUM)",
    "TikTok Video Saves (SUM)",
    "Instagram Business Post Saved (SUM)",
    "Post Comments (SUM)",
    "Count of Neutral Comments (SUM)",
    "Count of Positive Comments (SUM)",
    "Count of Negative Comments (SUM)",
    "Post Shares (SUM)",
    "Post Likes And Reactions (SUM)",
    "Is Sponsored",
    "TikTok Video Views (SUM)",
    "Post Reach (SUM)"
]
BASELINE_COLUMNS_MAP = {
    "Engagement Rate": "Baseline Engagement Rate",
    "Eng. Rate Neg. Com.": "Baseline Neg. Com.",
    "Video Views": "Baseline Video Views",
    "Shares": "Baseline Shares",
    "Post Likes And Reactions": "Baseline Post Likes And Reactions",
    "Post Comments / X Replies (SUM)": "Baseline Post Comments",
}
# --- min_views_boost (limite mínimo de views por Account para permitir Boost) ---
MIN_VIEWS_SPREADSHEET_ID = "1HwNv2AJVgULBNmZbqiPYDzZY7zPNLKhZz1DLtGJkOpM"
MIN_VIEWS_SHEET_RANGE = "Sheet1!A:Z"
MIN_VIEWS_COLUMN = "min_views_boost"
MIN_DAYS_FOR_BOOST = 3

# --- Log de execução (data de início da automação) ---
RUN_LOG_SPREADSHEET_ID = "17_D9mcA3ZvwrHZevC-eCCBJjg47q2kIKi3trLeD6ZiA"
RUN_LOG_SHEET_NAME = "Run_Datetime"
# ==============================
# HELPER — CONVERSÃO SEGURA DE VALORES PARA O SHEETS
# ==============================
def _to_sheet_value(v):
    """
    Converte um valor para o tipo mais adequado antes de mandar pro Sheets API.
    - Números continuam como int/float nativos (não viram string), então o Sheets
      grava como número de verdade, independente do locale (evita o problema de
      '.' vs ',' como separador decimal).
    - NaN/NaT/None viram string vazia.
    - Qualquer outra coisa vira string (texto mesmo).
    """
    if v is None:
        return ""
    if isinstance(v, (pd.Timestamp,)) or (hasattr(pd, "isna") and pd.isna(v) is True):
        return ""
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        if np.isnan(v):
            return ""
        return float(v)
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str) and v.strip() == "":
        return ""
    return str(v)
def _row_to_sheet_values(values):
    return [_to_sheet_value(v) for v in values]
# ==============================
# GOOGLE SERVICES
# ==============================
def get_google_services():
    creds_json = json.loads(os.environ.get("GDRIVE_CREDENTIALS_KC"))
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = service_account.Credentials.from_service_account_info(
        creds_json, scopes=scopes
    )
    drive_service = build("drive", "v3", credentials=creds)
    sheets_service = build("sheets", "v4", credentials=creds)
    return drive_service, sheets_service
# ==============================
# ETAPA 1 — ACHAR O ARQUIVO MAIS RECENTE NO DRIVE
# ==============================
def find_latest_file(drive_service):
    query = f"'{DRIVE_FOLDER_ID}' in parents and trashed = false"
    results = drive_service.files().list(
        q=query,
        fields="files(id, name)",
        pageSize=1000,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        corpora="allDrives",
    ).execute()
    files = results.get("files", [])
    print(f"[DEBUG] Arquivos visíveis na pasta {DRIVE_FOLDER_ID}: {len(files)}")
    for f in files:
        print(f"[DEBUG]  - {f['name']} (id={f['id']})")
    candidates = []
    for f in files:
        m = FILENAME_PATTERN.match(f["name"])
        if not m:
            continue
        dd, mm, yyyy = m.groups()
        try:
            file_date = datetime(int(yyyy), int(mm), int(dd))
        except ValueError:
            continue
        candidates.append((file_date, f))
    if not files:
        raise RuntimeError(
            "Nenhum arquivo foi retornado para essa pasta. Provavelmente a service "
            "account não tem acesso a essa pasta, ou o ID da pasta está errado."
        )
    if not candidates:
        raise RuntimeError(
            "Nenhum arquivo no padrão DDMMAAAA.xlsx foi encontrado na pasta do Drive. "
            "Veja a lista de arquivos no log [DEBUG] acima para conferir os nomes reais."
        )
    candidates.sort(key=lambda x: x[0])
    latest_date, latest_file = candidates[-1]
    print(f"Arquivo mais recente encontrado: {latest_file['name']} (data {latest_date.date()})")
    return latest_file["id"], latest_file["name"]
def download_file(drive_service, file_id, local_path):
    request = drive_service.files().get_media(fileId=file_id, supportsAllDrives=True)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    with open(local_path, "wb") as f:
        f.write(fh.getvalue())
    print(f"Arquivo baixado em: {local_path}")
# ==============================
# ETAPA 2 — LER E TRATAR A ABA ORGANIC
# ==============================
def extract_organic_id(row):
    url = row[PERMALINK_COLUMN]
    network = row["Social Network"]
    category = row["Outbound Message Category"]
    if not isinstance(url, str):
        return None
    if network == "Instagram" and category == "Reels":
        match = re.search(r"/reel/([^/]+)/", url)
    elif network == "Instagram" and category == "Story":
        match = re.search(r"/stories/[^/]+/(\d+)", url)
    elif network == "Instagram" and category == "Update":
        match = re.search(r"/p/([^/]+)/", url)
    elif network == "TikTok" and category == "Update":
        match = re.search(r"/video/(\d+)", url)
    else:
        return None
    return match.group(1) if match else None
def read_baseline(sheets_service):
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=BASELINE_SPREADSHEET_ID,
        range=BASELINE_SHEET_RANGE,
    ).execute()
    rows = result.get("values", [])
    if not rows:
        raise RuntimeError("Planilha de baseline está vazia.")
    headers = rows[0]
    data_rows = rows[1:]
    data_rows = [r + [""] * (len(headers) - len(r)) for r in data_rows]
    baseline = pd.DataFrame(data_rows, columns=headers)
    needed_cols = ["Concatenate"] + list(BASELINE_COLUMNS_MAP.keys())
    missing = [c for c in needed_cols if c not in baseline.columns]
    if missing:
        raise RuntimeError(f"Colunas faltando na planilha de baseline: {missing}")
    baseline = baseline[needed_cols].rename(columns=BASELINE_COLUMNS_MAP)
    numeric_cols = list(BASELINE_COLUMNS_MAP.values())
    for col in numeric_cols:
        baseline[col] = pd.to_numeric(
            baseline[col].astype(str).str.replace(",", ".").str.replace("%", ""),
            errors="coerce",
        )/100
    print(f"Baseline lido: {len(baseline)} linhas.")
    return baseline
    
def read_min_views_boost(sheets_service):
    """
    Lê a planilha de min_views_boost e retorna um DataFrame com apenas as colunas
    'Account' e 'min_views_boost'. Usada num left join por 'Account' para trazer,
    para cada conta, o mínimo de views necessário para permitir um Boost.
    """
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=MIN_VIEWS_SPREADSHEET_ID,
        range=MIN_VIEWS_SHEET_RANGE,
    ).execute()
    rows = result.get("values", [])
    if not rows:
        raise RuntimeError("Planilha de min_views_boost está vazia.")
    headers = rows[0]
    data_rows = rows[1:]
    data_rows = [r + [""] * (len(headers) - len(r)) for r in data_rows]
    min_views = pd.DataFrame(data_rows, columns=headers)
    needed_cols = ["Account", MIN_VIEWS_COLUMN]
    missing = [c for c in needed_cols if c not in min_views.columns]
    if missing:
        raise RuntimeError(
            f"Colunas faltando na planilha de min_views_boost: {missing}. "
            f"Colunas disponíveis: {list(min_views.columns)}"
        )
    min_views = min_views[needed_cols].copy()
    # Converte o limite para número (aceitando ',' como separador decimal).
    min_views[MIN_VIEWS_COLUMN] = pd.to_numeric(
        min_views[MIN_VIEWS_COLUMN].astype(str).str.replace(",", "."),
        errors="coerce",
    ).fillna(0)
    # Evita multiplicar linhas no merge caso a mesma Account apareça repetida.
    min_views["Account"] = min_views["Account"].astype(str).str.strip()
    min_views = min_views.drop_duplicates(subset="Account").copy()
    print(f"min_views_boost lido: {len(min_views)} linhas.")
    return min_views
    
def classify_boost(row):
    # Prioridade 1: Stories nunca entram na regra de Boost.
    if row["Outbound Message Category"] == "Story":
        return "Not Applies"
    boost = (
        row["Delta Eng. Rate (p.p.)"] > 0
        and row["Delta Neg. Sent. (p.p.)"] <= 0
        and row["Video Views (SUM)"] >= row[MIN_VIEWS_COLUMN]
        and row["Days Since Published"] >= MIN_DAYS_FOR_BOOST
    )
    if boost:
        return "Boost"
    elif row["Delta Eng. Rate (p.p.)"] > 0 and row["Delta Neg. Sent. (p.p.)"] > 0:
        return "Review"
    else:
        return "No Boost"
def read_organic_sheet(local_path, baseline_df, min_views_df):
    df = pd.read_excel(local_path, sheet_name=SOURCE_TAB, header=HEADER_SKIPROWS)
    missing_needed = [c for c in RAW_COLUMNS_NEEDED if c not in df.columns]
    if missing_needed:
        raise RuntimeError(
            f"Colunas esperadas não encontradas no arquivo Excel: {missing_needed}. "
            "O layout do relatório da Winclap pode ter mudado."
        )
    extra_cols = [c for c in df.columns if c not in RAW_COLUMNS_NEEDED]
    if extra_cols:
        print(f"[INFO] Colunas novas no relatório ignoradas de propósito: {extra_cols}")
    df = df[RAW_COLUMNS_NEEDED].copy()
    df = df[df["Outbound Message Category"] != "Reply"].copy()
    df["Organic_ID"] = df.apply(extract_organic_id, axis=1)
    df = df[df["Organic_ID"].notna()].copy()
 
    # ------------------------------------------------------------------
    # CHAVE COMPOSTA + DEDUPLICAÇÃO
    # Identidade única da linha = Permalink + Country of Origin.
    # ------------------------------------------------------------------
    df[UPSERT_KEY_COLUMN] = (
        df[PERMALINK_COLUMN].fillna("").astype(str).str.strip()
        + "||"
        + df["Country of Origin (Account)"].fillna("").astype(str).str.strip()
    )
    # Remove duplicatas com base na chave composta (mantém a 1ª ocorrência).
    df = df.drop_duplicates(subset=UPSERT_KEY_COLUMN).copy()
 
    interaction_cols = [
        "Post Comments (SUM)",
        "Post Shares (SUM)",
        "Post Likes And Reactions (SUM)",
        "TikTok Video Saves (SUM)",
        "Instagram Business Post Saved (SUM)"
    ]
    # .fillna(0) ensures that missing values don't break the addition
    df["Total Interactions"] = df[interaction_cols].fillna(0).sum(axis=1)
    def calculate_custom_engagement(row):
        total_int = row["Total Interactions"]
        video_views = row["Video Views (SUM)"]
        tiktok_views = row["TikTok Video Views (SUM)"]
        post_reach = row["Post Reach (SUM)"]
        # 1) If Video Views (SUM) > 0
        if video_views > 0:
            return total_int / video_views
        # 2) If Video Views (SUM) == 0 AND TikTok Video Views > 0
        # (Placed first to handle specific TikTok logic before general fallback)
        elif tiktok_views > 0:
            return total_int / tiktok_views
        # 3) If Video Views (SUM) == 0 fallback to Post Reach
        elif post_reach > 0:
            return total_int / post_reach
        # 4) Alternate fallback if all denominators are 0
        else:
            return 0
    df["Engagement Rate"] = df.apply(calculate_custom_engagement, axis=1)
    df["Engagement Rate"] = df["Engagement Rate"]
    def safe_pct(row, numerator_col):
        total = (
            row["Count of Negative Comments (SUM)"]
            + row["Count of Positive Comments (SUM)"]
            + row["Count of Neutral Comments (SUM)"]
        )
        if row["Post Comments (SUM)"] == 0 or total == 0:
            return 0
        return row[numerator_col] / total
    df["Sent. Negativo (%)"] = df.apply(lambda r: safe_pct(r, "Count of Negative Comments (SUM)"), axis=1)
    df["Sent. Positivo (%)"] = df.apply(lambda r: safe_pct(r, "Count of Positive Comments (SUM)"), axis=1)
    df["Sent. Neutro (%)"] = df.apply(lambda r: safe_pct(r, "Count of Neutral Comments (SUM)"), axis=1)
    df["Sent. Negativo (%)"] = df["Sent. Negativo (%)"]
    df["Sent. Positivo (%)"] = df["Sent. Positivo (%)"]
    df["Sent. Neutro (%)"] = df["Sent. Neutro (%)"]
    df["Concatenate"] = (
        df[["Account", "Country of Origin (Account)", "Outbound Message Category"]]
        .fillna("")
        .agg("".join, axis=1)
    )
    df = df.merge(baseline_df, on="Concatenate", how="left")
    df["Delta Eng. Rate (p.p.)"] = df["Engagement Rate"] - df["Baseline Engagement Rate"]
    df["Delta Neg. Sent. (p.p.)"] = df["Sent. Negativo (%)"] - df["Baseline Neg. Com."]

    # Left join min_views_boost por Account (sem match => 0), antes de classify.
    df["Account"] = df["Account"].astype(str).str.strip()
    df = df.merge(min_views_df, on="Account", how="left")
    df[MIN_VIEWS_COLUMN] = pd.to_numeric(df[MIN_VIEWS_COLUMN], errors="coerce").fillna(0)

    # Dias desde a publicação (data atual - Published Date), antes de virar string.
    published_dt = pd.to_datetime(df["Published Date"], errors="coerce")
    if getattr(published_dt.dt, "tz", None) is not None:
        published_dt = published_dt.dt.tz_localize(None)
    now_naive = datetime.now(tz_br).replace(tzinfo=None)
    df["Days Since Published"] = (pd.Timestamp(now_naive) - published_dt).dt.days

    df["Accionable"] = df.apply(classify_boost, axis=1)

    # Remove colunas auxiliares para não alterar o schema da planilha de destino.
    df = df.drop(columns=[MIN_VIEWS_COLUMN, "Days Since Published"])
    if "Published Date" in df.columns:
        df["Published Date"] = pd.to_datetime(df["Published Date"]).dt.strftime("%Y-%m-%d %H:%M:%S")
    df = df.fillna("")
    print(
        f"Linhas tratadas da aba {SOURCE_TAB} "
        f"(sem Reply, deduplicadas por {UPSERT_KEY_COLUMN} = Permalink + Country): {len(df)}"
    )
    return df
# ==============================
# ETAPA 3 — UPSERT NO GOOGLE SHEETS (ORGANIC)
# ==============================
def read_existing_sheet(sheets_service):
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!A:ZZ",
    ).execute()
    rows = result.get("values", [])
    if not rows:
        return [], {}
    headers = rows[0]
    key_col = headers.index(UPSERT_KEY_COLUMN) if UPSERT_KEY_COLUMN in headers else 0
    key_to_row_idx = {}
    for i, row in enumerate(rows[1:], start=2):
        if len(row) > key_col and row[key_col]:
            key_to_row_idx[row[key_col]] = i
    return headers, key_to_row_idx
def ensure_header(sheets_service, existing_headers, sheet_columns):
    if existing_headers:
        return existing_headers
    sheets_service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!A1",
        valueInputOption="RAW",
        body={"values": [sheet_columns]},
    ).execute()
    print("Cabeçalho criado na planilha.")
    return sheet_columns
def upsert_rows(sheets_service, df):
    other_cols = [c for c in df.columns if c != UPSERT_KEY_COLUMN]
    sheet_columns = [UPSERT_KEY_COLUMN] + other_cols + ["Last Updated At"]
    existing_headers, key_to_row_idx = read_existing_sheet(sheets_service)
    headers = ensure_header(sheets_service, existing_headers, sheet_columns)
    if set(sheet_columns) != set(headers):
        only_in_df = [c for c in sheet_columns if c not in headers]
        only_in_sheet = [c for c in headers if c not in sheet_columns]
        raise RuntimeError(
            "O conjunto de colunas do df não bate com o cabeçalho da planilha de "
            f"destino. Colunas só no df (novas): {only_in_df}. "
            f"Colunas só na planilha (faltando no df): {only_in_sheet}. "
            "Ajuste RAW_COLUMNS_NEEDED ou o cabeçalho da planilha antes de continuar."
        )
    now_str = datetime.now(tz_br).strftime("%Y-%m-%d %H:%M:%S")
    update_data = []
    append_rows = []
    data_headers = [h for h in headers if h != "Last Updated At"]
    end_col_letter = _col_letter(len(headers))
    for _, row in df.iterrows():
        key = str(row[UPSERT_KEY_COLUMN]).strip()
        raw_values = [row.get(col, "") for col in data_headers] + [now_str]
        # Mantém números como int/float nativos (em vez de forçar str),
        # para o Sheets gravar como número de verdade e não como texto.
        row_values = _row_to_sheet_values(raw_values)
        if key in key_to_row_idx:
            sheet_row = key_to_row_idx[key]
            update_data.append({
                "range": f"{SHEET_NAME}!A{sheet_row}:{end_col_letter}{sheet_row}",
                "values": [row_values],
            })
        else:
            append_rows.append(row_values)
    if update_data:
        sheets_service.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"valueInputOption": "USER_ENTERED", "data": update_data},
        ).execute()
        print(f"Linhas atualizadas: {len(update_data)}")
    if append_rows:
        sheets_service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_NAME}!A:A",
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body={"values": append_rows},
        ).execute()
        print(f"Linhas novas inseridas: {len(append_rows)}")
    if not update_data and not append_rows:
        print("Nenhuma linha para atualizar ou inserir.")
def _col_letter(n):
    letters = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters
# ==============================
# ETAPA 5 — REGISTRAR DATA DE INÍCIO DA AUTOMAÇÃO (após sucesso total)
# ==============================
def log_run_start(sheets_service, start_time_str):
    """
    Adiciona uma linha na aba Run_Datetime com a data/hora de conclusão da
    automação (coluna A) e o nome do script (coluna B). É chamada só depois
    que todo o pipeline rodou com sucesso.
    """
    sheets_service.spreadsheets().values().append(
        spreadsheetId=RUN_LOG_SPREADSHEET_ID,
        range=f"{RUN_LOG_SHEET_NAME}!A:B",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": [[start_time_str, "Sprinklr_Organic_Script"]]},
    ).execute()
    print(f"Data de início da automação registrada em '{RUN_LOG_SHEET_NAME}': {start_time_str}")
# ==============================
# EXECUÇÃO PRINCIPAL
# ==============================
def main():
    print("=" * 60)
    print("INICIANDO PIPELINE COMPLETO")
    print("=" * 60)
    missing = [v for v in ["GDRIVE_CREDENTIALS_KC"] if not os.environ.get(v)]
    if missing:
        print(f"Variáveis de ambiente faltando: {missing}. Encerrando.")
        sys.exit(1)
    drive_service, sheets_service = get_google_services()
    # --- Etapa 1 & 2: Organic (Winclap) ---
    print("\n[ETAPA 1/1] Processando dados orgânicos (Winclap)...")
    file_id, file_name = find_latest_file(drive_service)
    local_path = f"/tmp/{file_name}"
    download_file(drive_service, file_id, local_path)
    baseline_df = read_baseline(sheets_service)
    min_views_df = read_min_views_boost(sheets_service)
    organic_df = read_organic_sheet(local_path, baseline_df, min_views_df)
    upsert_rows(sheets_service, organic_df)
    # --- Etapa 5: registra a data/hora de conclusão da automação, só após sucesso total ---
    run_end_str = datetime.now(tz_br).strftime("%Y-%m-%d %H:%M:%S")
    log_run_start(sheets_service, run_end_str)
    print("\n" + "=" * 60)
    print("PIPELINE FINALIZADO COM SUCESSO")
    print("=" * 60)
if __name__ == "__main__":
    main()
 
