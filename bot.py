import os, re, io, time, json, base64, logging, threading, requests
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from http.server import HTTPServer, BaseHTTPRequestHandler
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, filters, ContextTypes

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    DRIVE_API_AVAILABLE = True
except ImportError:
    DRIVE_API_AVAILABLE = False

logging.basicConfig(format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "8822872793:AAEn0p_FLlH1CaB41cHhwHV7ADh7q6o6VGU").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "AQ.Ab8RN6Ji0H4-YVfeTsoYlo62qCVfulC0nYpDSalIf9MXI-EisA").strip()
RENDER_EXTERNAL_URL = os.getenv("RENDER_EXTERNAL_URL", "https://arslanlar-marble-bot.onrender.com").strip()
DEFAULT_DRIVE_LINK = os.getenv("GOOGLE_DRIVE_LINK", "https://drive.google.com/drive/folders/1bnZH4cy2zOdZWhvuN79DURjabDJehOjs")

GOOGLE_SERVICE_ACCOUNT_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
SERVICE_ACCOUNT_FILE = "./service_account.json"
PERSISTENT_STOCK_PATH = "./master_stok_arslanlar.xlsx"
PERSISTENT_PRICE_PATH = "./master_fiyat_arslanlar.xlsx"
HYPERLINKS_DB_PATH = "./hyperlinks_db.json"
CHAT_MEMORY_PATH = "./chat_memory.json"

bot_lock = threading.Lock()
user_context_memory = {}


# ==============================================================================
# 1. YARDIMCI ARAÇLAR VE METİN NORMALİZASYONU
# ==============================================================================
def tr_lower(text: str) -> str:
    if not text:
        return ""
    res = str(text)
    for k, v in {"İ": "i", "I": "ı", "Ğ": "ğ", "Ü": "ü", "Ş": "ş", "Ö": "ö", "Ç": "ç"}.items():
        res = res.replace(k, v)
    return res.lower().strip()


def tr_norm(text: str) -> str:
    return tr_lower(text).translate(str.maketrans("ıiğüöşç", "iiguosc"))


def pdf_ascii(text: str) -> str:
    if not text:
        return ""
    res = str(text)
    for k, v in {"İ": "I", "ı": "i", "Ğ": "G", "ğ": "g", "Ü": "U", "ü": "u", "Ş": "S", "ş": "s", "Ö": "O", "ö": "o", "Ç": "C", "ç": "c"}.items():
        res = res.replace(k, v)
    return res


def normalize_lot_code(text: str) -> str:
    clean = re.sub(r"[^a-zA-Z0-9]", "", str(text)).upper()
    m = re.match(r"^([A-Z]{1,3})(\d+[A-Z]?)$", clean)
    return f"{m.group(1)}-{m.group(2)}" if m else clean


async def safe_reply(update: Update, text: str):
    if not update.message or not text:
        return
    try:
        await update.message.reply_text(text, parse_mode="Markdown", disable_web_page_preview=True)
    except Exception:
        clean_text = text.replace("**", "").replace("*", "").replace("`", "")
        await update.message.reply_text(clean_text, disable_web_page_preview=True)


# ==============================================================================
# 2. VERİTABANI VE EXCEL YÖNETİMİ
# ==============================================================================
def load_json_db(path: str) -> dict:
    with bot_lock:
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                logger.error(f"JSON okuma hatası: {e}")
        return {}


def save_json_db(path: str, data: dict):
    with bot_lock:
        try:
            current = {}
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    current = json.load(f)
            current.update(data)
            with open(path, "w", encoding="utf-8") as f:
                json.dump(current, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.error(f"JSON kaydetme hatası: {e}")


def append_chat_history(chat_id: int, role: str, content: str):
    cid = str(chat_id)
    mem = load_json_db(CHAT_MEMORY_PATH)
    history = mem.get(cid, [])
    history.append({"role": role, "text": content[:500], "time": datetime.now().strftime("%H:%M")})
    mem[cid] = history[-15:]
    save_json_db(CHAT_MEMORY_PATH, mem)


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df.columns = [str(c).strip() for c in df.columns]
    valid_cols = [c for c in df.columns if not c.lower().startswith("unnamed")]
    return df[valid_cols].dropna(how="all")


def get_master_df() -> pd.DataFrame:
    with bot_lock:
        if os.path.exists(PERSISTENT_STOCK_PATH):
            try:
                return clean_dataframe(pd.read_excel(PERSISTENT_STOCK_PATH, engine="openpyxl"))
            except Exception as e:
                logger.error(f"Stok okuma hatası: {e}")
        return pd.DataFrame()


def get_price_df() -> pd.DataFrame:
    with bot_lock:
        if os.path.exists(PERSISTENT_PRICE_PATH):
            try:
                return clean_dataframe(pd.read_excel(PERSISTENT_PRICE_PATH, engine="openpyxl"))
            except Exception as e:
                logger.error(f"Fiyat okuma hatası: {e}")
        return pd.DataFrame()


def extract_excel_hyperlinks(excel_path: str) -> dict:
    extracted = {}
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=False)
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            for row in ws.iter_rows():
                row_vals = [str(c.value).strip() if c.value is not None else "" for c in row]
                for cell in row:
                    url = None
                    if cell.hyperlink and cell.hyperlink.target:
                        url = cell.hyperlink.target
                    elif isinstance(cell.value, str) and "HYPERLINK(" in cell.value.upper():
                        m = re.search(r'HYPERLINK\(\s*"([^"]+)"', cell.value, re.IGNORECASE)
                        if m:
                            url = m.group(1)
                    if url and "http" in url:
                        c_val = str(cell.value or "").strip()
                        if c_val and not c_val.upper().startswith("=HYPERLINK"):
                            extracted[tr_norm(c_val)] = url
                            extracted[normalize_lot_code(c_val)] = url
                        for v in row_vals:
                            if v and len(v) > 2 and not v.upper().startswith("=HYPERLINK"):
                                extracted[tr_norm(v)] = url
                                extracted[normalize_lot_code(v)] = url
    except Exception as e:
        logger.warning(f"Köprü çıkarma uyarısı: {e}")
    return extracted


# ==============================================================================
# 3. OTONOM GOOGLE DRIVE GEZGİNİ VE CANLI FOTOĞRAF ÇEKİCİ
# ==============================================================================
def get_drive_service():
    if not DRIVE_API_AVAILABLE:
        return None
    try:
        scopes = ["https://www.googleapis.com/auth/drive.readonly"]
        if GOOGLE_SERVICE_ACCOUNT_JSON:
            info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
            creds = service_account.Credentials.from_service_account_info(info, scopes=scopes)
            return build("drive", "v3", credentials=creds)
        elif os.path.exists(SERVICE_ACCOUNT_FILE):
            creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
            return build("drive", "v3", credentials=creds)
    except Exception as e:
        logger.warning(f"Drive API başlatma uyarısı: {e}")
    return None


def get_best_drive_link(lot_code: str = None, stone_name: str = None) -> str:
    links_db = load_json_db(HYPERLINKS_DB_PATH)
    if lot_code:
        nl, nk = normalize_lot_code(lot_code), tr_norm(lot_code)
        if nl in links_db:
            return links_db[nl]
        if nk in links_db:
            return links_db[nk]
        for k, v in links_db.items():
            if nk in tr_norm(k) or tr_norm(k) in nk:
                return v
    if stone_name:
        ns = tr_norm(stone_name)
        if ns in links_db:
            return links_db[ns]
        for k, v in links_db.items():
            if len(ns) > 3 and (ns in tr_norm(k) or tr_norm(k) in ns):
                return v

    service = get_drive_service()
    if service and (lot_code or stone_name):
        q_key = lot_code or stone_name
        norm_q, lot_q = tr_norm(q_key), normalize_lot_code(q_key)
        try:
            q_str = f"mimeType = 'application/vnd.google-apps.folder' and trashed = false and (name contains '{q_key}' or name contains '{lot_q}')"
            res = service.files().list(q=q_str, pageSize=3, fields="files(id, name, webViewLink)").execute()
            files = res.get("files", [])
            if files:
                url = files[0].get("webViewLink") or f"https://drive.google.com/drive/folders/{files[0]['id']}"
                save_json_db(HYPERLINKS_DB_PATH, {lot_q: url, norm_q: url})
                return url
        except Exception:
            pass
    return DEFAULT_DRIVE_LINK


def search_drive_folder_and_photos(query_term: str) -> dict:
    result = {"folder_url": None, "photo_bytes": None, "photo_name": None}
    if not query_term:
        return result

    service = get_drive_service()
    norm_q, lot_q = tr_norm(query_term), normalize_lot_code(query_term)

    if service:
        try:
            q_str = f"mimeType = 'application/vnd.google-apps.folder' and trashed = false and (name contains '{query_term}' or name contains '{lot_q}')"
            res = service.files().list(q=q_str, pageSize=10, fields="files(id, name, webViewLink)").execute()
            folders = res.get("files", [])

            if not folders:
                res_all = service.files().list(q="mimeType = 'application/vnd.google-apps.folder' and trashed = false", pageSize=200, fields="files(id, name, webViewLink)").execute()
                for f in res_all.get("files", []):
                    if norm_q in tr_norm(f["name"]) or lot_q == normalize_lot_code(f["name"]):
                        folders.append(f)

            if folders:
                target_folder = folders[0]
                f_id = target_folder["id"]
                result["folder_url"] = target_folder.get("webViewLink") or f"https://drive.google.com/drive/folders/{f_id}"
                save_json_db(HYPERLINKS_DB_PATH, {lot_q: result["folder_url"], norm_q: result["folder_url"]})

                img_q = f"'{f_id}' in parents and mimeType contains 'image/' and trashed = false"
                images = service.files().list(q=img_q, pageSize=5, fields="files(id, name)").execute().get("files", [])

                if not images:
                    sub_q = f"'{f_id}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
                    for sf in service.files().list(q=sub_q, pageSize=5, fields="files(id, name)").execute().get("files", []):
                        img_q2 = f"'{sf['id']}' in parents and mimeType contains 'image/' and trashed = false"
                        imgs2 = service.files().list(q=img_q2, pageSize=3, fields="files(id, name)").execute().get("files", [])
                        if imgs2:
                            images = imgs2
                            break

                if images:
                    img_file = images[0]
                    req = service.files().get_media(fileId=img_file["id"])
                    fh = io.BytesIO()
                    downloader = MediaIoBaseDownload(fh, req)
                    done = False
                    while not done:
                        _, done = downloader.next_chunk()
                    fh.seek(0)
                    result["photo_bytes"] = fh.read()
                    result["photo_name"] = img_file["name"]
                    return result

            direct_img_q = f"mimeType contains 'image/' and trashed = false and (name contains '{query_term}' or name contains '{lot_q}')"
            d_imgs = service.files().list(q=direct_img_q, pageSize=3, fields="files(id, name, webViewLink)").execute().get("files", [])
            if d_imgs:
                img_file = d_imgs[0]
                result["folder_url"] = img_file.get("webViewLink", DEFAULT_DRIVE_LINK)
                req = service.files().get_media(fileId=img_file["id"])
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, req)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
                fh.seek(0)
                result["photo_bytes"] = fh.read()
                result["photo_name"] = img_file["name"]
                return result
        except Exception as e:
            logger.warning(f"Drive API arama hatası: {e}")

    saved_link = get_best_drive_link(lot_q, query_term)
    result["folder_url"] = saved_link
    if saved_link and "drive.google.com" in saved_link:
        try:
            m_file = re.search(r"/file/d/([a-zA-Z0-9_-]+)", saved_link)
            if m_file:
                dl_url = f"https://drive.google.com/uc?export=download&id={m_file.group(1)}"
                r = requests.get(dl_url, timeout=10)
                if r.status_code == 200 and "image" in r.headers.get("Content-Type", ""):
                    result["photo_bytes"] = r.content
                    result["photo_name"] = f"{lot_q}.jpg"
        except Exception:
            pass
    return result


# ==============================================================================
# 4. SÜTUN TANIMA, KIRIK/FİRE ANALİZİ VE LOJİSTİK MOTORU
# ==============================================================================
def identify_columns(df: pd.DataFrame) -> dict:
    cols = {"lot": [], "stone": None, "status": None, "location": None, "reserve": None, "m2": None, "surface": None, "thickness": None, "width": None, "length": None, "slab_no": None, "company": None}
    for c in df.columns:
        cn = tr_norm(c)
        if any(k in cn for k in ["fabrika", "blok", "lot", "kod"]):
            cols["lot"].append(c)
        if any(k in cn for k in ["seleksiyon", "tas", "cins", "urun", "stone", "material"]):
            cols["stone"] = c
        elif cn in ["m2", "m²", "metraj", "alan", "miktar"]:
            cols["m2"] = c
        elif "durum" in cn and "rezerv" not in cn:
            cols["status"] = c
        elif "konum" in cn or "hol" in cn:
            cols["location"] = c
        elif "rezerv" in cn or "satis" in cn:
            cols["reserve"] = c
        elif "yuzey" in cn or "islem" in cn:
            cols["surface"] = c
        elif "kalinlik" in cn:
            cols["thickness"] = c
        elif "uzunluk" in cn or "boy" in cn:
            cols["length"] = c
        elif "yukseklik" in cn or "en" in cn:
            cols["width"] = c
        elif "plaka" in cn:
            cols["slab_no"] = c
        elif "firma" in cn:
            cols["company"] = c
    if not cols["m2"]:
        cols["m2"] = next((c for c in df.columns if "m2" in tr_norm(c) or "metraj" in tr_norm(c)), None)
    return cols


def get_broken_mask(df: pd.DataFrame) -> pd.Series:
    if df.empty:
        return pd.Series(dtype=bool)
    broken_kw = r"kirik|kırık|kirildi|fire|catlak|çatlak|hasar|iptal|iskarta|cop|çöp|imha"
    return df.astype(str).apply(lambda col: col.apply(tr_norm).str.contains(broken_kw, regex=True, na=False)).any(axis=1)


def calculate_weight_and_containers(df: pd.DataFrame, cols: dict) -> tuple:
    if df.empty or not cols["m2"]:
        return 0.0, 0.0
    m2_s = pd.to_numeric(df[cols["m2"]], errors="coerce").fillna(0.0)
    if cols["thickness"] and cols["thickness"] in df.columns:
        th_s = pd.to_numeric(df[cols["thickness"]], errors="coerce").fillna(20.0)
        th_mm = th_s.apply(lambda x: x * 10 if x < 10 else x)
    else:
        th_mm = pd.Series(20.0, index=df.index)
    tons = round(float((m2_s * (th_mm / 1000.0) * 2.75).sum()), 2)
    return tons, round(tons / 27.0, 2)


def find_stone_price(stone_name: str) -> str:
    price_df = get_price_df()
    if price_df.empty or not stone_name:
        return ""
    target = tr_norm(stone_name.split(",")[0])
    mask = price_df.astype(str).apply(lambda x: x.apply(tr_norm).str.contains(target, na=False)).any(axis=1)
    matched = price_df[mask]
    if not matched.empty:
        row = matched.iloc[0]
        return " | ".join([f"{c}: *{row[c]}*" for c in matched.columns[:5] if pd.notna(row[c]) and str(row[c]).strip() != ""])
    return ""


# ==============================================================================
# 5. GELİŞMİŞ SORGU ÇÖZÜMLEYİCİ
# ==============================================================================
def parse_user_query_and_filter(text: str, df: pd.DataFrame, chat_id: int = 0) -> dict:
    t_norm = tr_norm(text)
    cols = identify_columns(df)

    lot_candidates = re.findall(r"\b([a-zA-Z]{1,3}[- ]?\d{2,5}[a-zA-Z]?)\b", text)
    ignored = {"M2", "A4", "A3", "3D", "2D", "724", "20MM", "30MM", "2CM", "3CM", "PDF"}
    detected_lots = [normalize_lot_code(c) for c in lot_candidates if normalize_lot_code(c) not in ignored]

    detected_stone = None
    if not df.empty and cols["stone"]:
        unique_stones = sorted(df[cols["stone"]].dropna().astype(str).unique(), key=lambda x: len(str(x)), reverse=True)
        for st in unique_stones:
            st_clean = tr_norm(st)
            if len(st_clean) > 2 and st_clean in t_norm:
                detected_stone = st.strip()
                break

    ctx = user_context_memory.get(chat_id, {})
    if not detected_lots and not detected_stone:
        if any(w in t_norm for w in ["bunun", "onun", "bu lotun", "bu tasin", "ayni", "peki", "ceki", "link", "fotograf", "foto", "proforma", "kirik", "fire"]):
            if ctx.get("lot"):
                detected_lots = [ctx["lot"]]
            elif ctx.get("stone"):
                detected_stone = ctx["stone"]

    if detected_lots or detected_stone:
        user_context_memory[chat_id] = {"lot": detected_lots[0] if detected_lots else None, "stone": detected_stone}

    status_filter = "all"
    if any(k in t_norm for k in ["kirik", "kirilan", "fire", "catlak", "hasarli", "iptal", "iskarta"]):
        status_filter = "kirik"
    elif any(k in t_norm for k in ["serbest", "musait", "rezervsiz"]):
        status_filter = "serbest"
    elif any(k in t_norm for k in ["rezerv"]):
        status_filter = "rezerv"
    elif any(k in t_norm for k in ["yuklenen", "yuklendi", "sevk", "giden", "satilan"]):
        status_filter = "yuklenen"
    elif any(k in t_norm for k in ["stokta", "stoktaki", "depoda", "mevcut", "eldeki", "stok listesi", "stok"]):
        if "uretilen" not in t_norm and "toplam" not in t_norm:
            status_filter = "stok"

    unit_price, currency = None, "USD"
    price_match = re.search(r"(\d+(?:[.,]\d+)?)\s*(?:\$|dolar|usd|€|euro|eur|tl)", t_norm)
    if price_match:
        unit_price = float(price_match.group(1).replace(",", "."))
        if any(c in t_norm for c in ["€", "euro", "eur"]):
            currency = "EUR"

    filtered = df.copy()
    lot_matched = True

    if not df.empty:
        if detected_lots and cols["lot"]:
            target_lot = detected_lots[0]
            mask = pd.Series(False, index=filtered.index)
            for lcol in cols["lot"]:
                mask = mask | (filtered[lcol].astype(str).apply(normalize_lot_code) == target_lot)
            if not mask.any():
                for lcol in cols["lot"]:
                    mask = mask | filtered[lcol].astype(str).apply(normalize_lot_code).str.contains(target_lot, na=False)
            filtered = filtered[mask]
            if filtered.empty:
                lot_matched = False

        if detected_stone and cols["stone"] and not filtered.empty:
            mask = filtered[cols["stone"]].astype(str).apply(tr_norm).str.contains(tr_norm(detected_stone), na=False)
            filtered = filtered[mask]

        filtered_by_status = filtered.copy()
        if not filtered.empty and status_filter != "all":
            b_mask = get_broken_mask(filtered)
            if status_filter == "kirik":
                filtered_by_status = filtered[b_mask]
            elif status_filter == "stok" and cols["status"]:
                s_mask = filtered[cols["status"]].astype(str).apply(tr_norm).str.contains("stok", na=False)
                filtered_by_status = filtered[s_mask & (~b_mask)]
            elif status_filter == "yuklenen" and cols["status"]:
                y_mask = filtered[cols["status"]].astype(str).apply(tr_norm).str.contains("yuklen", na=False)
                filtered_by_status = filtered[y_mask]
            elif status_filter == "rezerv" and cols["reserve"]:
                r_mask = filtered[cols["reserve"]].astype(str).apply(tr_norm).str.contains("rezerv", na=False)
                filtered_by_status = filtered[r_mask & (~b_mask)]
            elif status_filter == "serbest" and cols["status"]:
                s_mask = filtered[cols["status"]].astype(str).apply(tr_norm).str.contains("stok", na=False)
                if cols["reserve"]:
                    r_mask = filtered[cols["reserve"]].astype(str).apply(tr_norm).str.contains("rezerv", na=False)
                    filtered_by_status = filtered[s_mask & (~r_mask) & (~b_mask)]
                else:
                    filtered_by_status = filtered[s_mask & (~b_mask)]
    else:
        filtered_by_status = pd.DataFrame()

    return {
        "lots": detected_lots, "stone": detected_stone, "status_filter": status_filter,
        "unit_price": unit_price, "currency": currency, "base_df": filtered,
        "final_df": filtered_by_status, "lot_matched": lot_matched, "cols": cols
    }


# ==============================================================================
# 6. KURUMSAL EXCEL ÇEKİ LİSTESİ & PDF PROFORMA INVOICE ÜRETİCİLERİ
# ==============================================================================
def generate_styled_excel(filtered_df: pd.DataFrame, file_name: str) -> str:
    df_out = filtered_df.copy()
    cols_norm_map = {tr_norm(c): c for c in df_out.columns}
    if "stok durumu" in cols_norm_map and "stok konumu" in cols_norm_map:
        drop_c = [cols_norm_map["stok durumu"]] + [cols_norm_map[x] for x in ["firma", "uretim tarihi", "yuklenme tarihi"] if x in cols_norm_map]
        df_out = df_out.drop(columns=[c for c in drop_c if c in df_out.columns], errors="ignore")

    cols_meta = identify_columns(df_out)
    m2_col = cols_meta["m2"]
    if m2_col and m2_col in df_out.columns:
        df_out[m2_col] = pd.to_numeric(df_out[m2_col], errors="coerce").round(2)

    df_out.to_excel(file_name, index=False, engine="openpyxl")
    wb = openpyxl.load_workbook(file_name)
    ws = wb.active
    ws.title = "Arslanlar_Ceki_Listesi"

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    zebra_fill = PatternFill(start_color="F2F6FA", end_color="F2F6FA", fill_type="solid")
    total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10, color="000000")
    link_font = Font(name="Calibri", size=10, color="0563C1", underline="single")
    total_font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
    thin_border = Border(left=Side(style="thin", color="BFBFBF"), right=Side(style="thin", color="BFBFBF"), top=Side(style="thin", color="BFBFBF"), bottom=Side(style="thin", color="BFBFBF"))
    align_center = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[1].height = 25
    for col_num in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill, cell.font, cell.alignment, cell.border = header_fill, header_font, align_center, thin_border

    links_db = load_json_db(HYPERLINKS_DB_PATH)
    max_data_row = ws.max_row

    for row_num in range(2, max_data_row + 1):
        ws.row_dimensions[row_num].height = 19
        is_even = (row_num % 2 == 0)
        for col_num in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_num, column=col_num)
            val_str = str(cell.value or "").strip()
            if "00:00:00" in val_str:
                cell.value = val_str.split(" ")[0]
            cell.font, cell.alignment, cell.border = data_font, align_center, thin_border
            if is_even:
                cell.fill = zebra_fill
            norm_v, norm_l = tr_norm(val_str), normalize_lot_code(val_str)
            if norm_l in links_db and links_db[norm_l] != DEFAULT_DRIVE_LINK:
                cell.hyperlink, cell.font = links_db[norm_l], link_font
            elif norm_v in links_db and links_db[norm_v] != DEFAULT_DRIVE_LINK:
                cell.hyperlink, cell.font = links_db[norm_v], link_font

    total_row = max_data_row + 1
    ws.row_dimensions[total_row].height = 24
    m2_col_idx = next((col_num for col_num in range(1, ws.max_column + 1) if m2_col and str(ws.cell(row=1, column=col_num).value or "") == m2_col), None)

    tons, _ = calculate_weight_and_containers(filtered_df, identify_columns(filtered_df))
    ws.cell(row=total_row, column=1, value="TOPLAM").font = total_font
    ws.cell(row=total_row, column=2, value=f"{max_data_row - 1} Plaka (~{tons} Ton)").font = total_font

    if m2_col_idx and m2_col in df_out.columns:
        ws.cell(row=total_row, column=m2_col_idx, value=f"{round(float(df_out[m2_col].sum()), 2)} m²").font = total_font

    for col_num in range(1, ws.max_column + 1):
        c = ws.cell(row=total_row, column=col_num)
        c.fill, c.alignment, c.border = total_fill, align_center, thin_border

    for col in ws.columns:
        max_len = max((len(str(cell.value or "")) for cell in col), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 4, 13)

    wb.save(file_name)
    return file_name


def generate_proforma_pdf(df_target: pd.DataFrame, parsed: dict, file_name: str) -> str:
    if not REPORTLAB_AVAILABLE:
        return None
    cols = parsed["cols"]
    unit_price = parsed["unit_price"] or 65.0
    currency = parsed["currency"] or "USD"

    doc = SimpleDocTemplate(file_name, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("DocTitle", parent=styles["Heading1"], fontSize=18, textColor=colors.HexColor("#1F4E78"), spaceAfter=6)
    sub_style = ParagraphStyle("DocSub", parent=styles["Normal"], fontSize=10, textColor=colors.HexColor("#333333"))

    elements.append(Paragraph("<b>ARSLANLAR MARBLE - PROFORMA INVOICE</b>", title_style))
    elements.append(Paragraph(f"<b>Date:</b> {datetime.now().strftime('%d.%m.%Y')} | <b>Invoice No:</b> PI-{datetime.now().strftime('%Y%m%d-%H%M')}", sub_style))
    elements.append(Paragraph("<b>Exporter:</b> Arslanlar Mermer San. ve Tic. | Izmir, Turkiye", sub_style))
    elements.append(Paragraph("<b>Representative:</b> Mehmet Morkoc", sub_style))
    elements.append(Spacer(1, 15))

    m2_col, stone_col, thick_col, surf_col = cols["m2"], cols["stone"], cols["thickness"], cols["surface"]
    lot_col = cols["lot"][0] if cols["lot"] else None
    table_data = [["No", "Lot / Block", "Stone Selection", "Finish / Thick.", "Slabs", "Quantity (m2)", f"Unit Price ({currency})", f"Total ({currency})"]]

    temp = df_target.copy()
    temp["_m2"] = pd.to_numeric(temp[m2_col], errors="coerce").fillna(0.0) if m2_col else 0.0
    group_col = lot_col if lot_col else stone_col
    total_slabs_all, total_m2_all, grand_total_val = 0, 0.0, 0.0

    if group_col and group_col in temp.columns:
        idx = 1
        for g_name, g_df in temp.groupby(group_col):
            slabs_cnt = len(g_df)
            m2_sum = round(float(g_df["_m2"].sum()), 2)
            st_name = str(g_df[stone_col].iloc[0]) if (stone_col and stone_col in g_df.columns) else str(g_name)
            surf_val = str(g_df[surf_col].iloc[0]) if (surf_col and surf_col in g_df.columns) else "Polished"
            th_val = str(g_df[thick_col].iloc[0]) if (thick_col and thick_col in g_df.columns) else "2cm"
            line_tot = round(m2_sum * unit_price, 2)
            total_slabs_all += slabs_cnt
            total_m2_all += m2_sum
            grand_total_val += line_tot
            table_data.append([str(idx), pdf_ascii(str(g_name)), pdf_ascii(st_name[:22]), pdf_ascii(f"{surf_val} / {th_val}"), str(slabs_cnt), f"{m2_sum:,.2f}", f"{unit_price:,.2f}", f"{line_tot:,.2f}"])
            idx += 1
            if idx > 25:
                break
    else:
        total_slabs_all = len(temp)
        total_m2_all = round(float(temp["_m2"].sum()), 2)
        grand_total_val = round(total_m2_all * unit_price, 2)
        table_data.append(["1", "ALL", "Marble Selections", "Polished / 2cm", str(total_slabs_all), f"{total_m2_all:,.2f}", f"{unit_price:,.2f}", f"{grand_total_val:,.2f}"])

    tons, containers = calculate_weight_and_containers(df_target, cols)
    table_data.append(["TOTAL", "", f"Est. Weight: ~{tons} Tons", f"(~{containers} Cont.)", str(total_slabs_all), f"{total_m2_all:,.2f} m2", "GRAND TOTAL:", f"{grand_total_val:,.2f} {currency}"])

    t = Table(table_data, colWidths=[28, 60, 105, 80, 40, 68, 70, 80])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("FONTSIZE", (0, 1), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#D9E1F2")),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
    ]))
    elements.append(t)
    elements.append(Spacer(1, 20))
    elements.append(Paragraph("<b>Delivery & Payment Terms:</b>", sub_style))
    elements.append(Paragraph("- Delivery Terms: FOB Izmir Port / EXW Factory (Torbali, Izmir)", sub_style))
    elements.append(Paragraph("- Payment Terms: 30% Advance Payment, 70% Balance Before Shipment", sub_style))
    elements.append(Paragraph("- Packing: Fumigated strong wooden bundles suitable for ocean freight.", sub_style))

    doc.build(elements)
    return file_name


# ==============================================================================
# 7. DETAYLI METİN RAPORU (KIRIK/FİRE + STOK + LOJİSTİK + FİYAT)
# ==============================================================================
def build_detailed_analysis_text(parsed: dict, full_df: pd.DataFrame) -> str:
    cols = parsed["cols"]
    base_df, final_df = parsed["base_df"], parsed["final_df"]
    lots, stone, status_filter = parsed["lots"], parsed["stone"], parsed["status_filter"]
    unit_price, currency = parsed["unit_price"], parsed["currency"]

    if not parsed["lot_matched"] and lots:
        return f"⚠️ Hafızadaki stok listesinde **{lots[0]}** kodlu lot/blok kaydı bulunamadı."
    if base_df.empty:
        return "⚠️ Aradığınız kriterlere uygun plaka kaydı bulunamadı."

    m2_col, stat_col, res_col, loc_col, stone_col, surf_col = cols["m2"], cols["status"], cols["reserve"], cols["location"], cols["stone"], cols["surface"]

    total_slabs = len(base_df)
    total_m2 = float(pd.to_numeric(base_df[m2_col], errors="coerce").sum()) if m2_col else 0.0

    b_mask = get_broken_mask(base_df)
    broken_df = base_df[b_mask]
    broken_slabs = len(broken_df)
    broken_m2 = float(pd.to_numeric(broken_df[m2_col], errors="coerce").sum()) if (m2_col and not broken_df.empty) else 0.0

    healthy_df = base_df[~b_mask]
    if stat_col:
        stok_df = healthy_df[healthy_df[stat_col].astype(str).apply(tr_norm).str.contains("stok", na=False)]
        yuk_df = healthy_df[healthy_df[stat_col].astype(str).apply(tr_norm).str.contains("yuklen", na=False)]
    else:
        stok_df, yuk_df = healthy_df, pd.DataFrame()

    stok_slabs = len(stok_df)
    stok_m2 = float(pd.to_numeric(stok_df[m2_col], errors="coerce").sum()) if (m2_col and not stok_df.empty) else 0.0
    yuk_slabs = len(yuk_df)
    yuk_m2 = float(pd.to_numeric(yuk_df[m2_col], errors="coerce").sum()) if (m2_col and not yuk_df.empty) else 0.0

    res_slabs, res_m2, free_slabs, free_m2 = 0, 0.0, stok_slabs, stok_m2
    if res_col and not stok_df.empty:
        res_mask = stok_df[res_col].astype(str).apply(tr_norm).str.contains("rezerv", na=False)
        res_df, free_df = stok_df[res_mask], stok_df[~res_mask]
        res_slabs = len(res_df)
        res_m2 = float(pd.to_numeric(res_df[m2_col], errors="coerce").sum()) if m2_col else 0.0
        free_slabs = len(free_df)
        free_m2 = float(pd.to_numeric(free_df[m2_col], errors="coerce").sum()) if m2_col else 0.0

    stone_label = stone
    if not stone_label and stone_col and not base_df.empty:
        top_stones = base_df[stone_col].dropna().astype(str).unique()
        stone_label = ", ".join(top_stones[:3]) if len(top_stones) <= 3 else f"{len(top_stones)} Farklı Seleksiyon"

    if status_filter == "kirik":
        if broken_df.empty:
            return f"✅ **{lots[0] if lots else (stone_label or 'Fabrika Stoğu')}** kapsamında kırık, çatlak veya fire plaka bulunmamaktadır."
        b_lots = broken_df[cols["lot"][0]].dropna().astype(str).unique().tolist()[:10] if cols["lot"] else []
        lot_line = f"\n🔢 *Kırık Bulunan Lotlar:* {', '.join(b_lots)}" if b_lots else ""
        return (
            f"⚠️ **ARSLANLAR MARBLE - KIRIK / FİRE PLAKA RAPORU**\n"
            f"🪨 *Kapsam:* {lots[0] if lots else (stone_label or 'Genel Fabrika')}\n"
            f"• **Kırık / Fire Plaka Sayısı:** {broken_slabs} Adet\n"
            f"• **Toplam Fire Metrajı:** **{round(broken_m2, 2)} m²**\n"
            f"• **Üretime Oranı:** %{round((broken_m2 / total_m2) * 100, 2) if total_m2 > 0 else 0}{lot_line}"
        )

    loc_info = ""
    if loc_col and not stok_df.empty:
        locs = [l for l in stok_df[loc_col].dropna().astype(str).unique() if "YÜKLENDİ" not in l.upper()]
        if locs:
            loc_info = f"\n📍 *Stok Konumu:* {', '.join(locs[:4])}"

    surf_info = ""
    if surf_col and not stok_df.empty:
        surfs = stok_df[surf_col].dropna().astype(str).unique()
        if surfs:
            surf_info = f"\n✨ *Yüzey İşlemi:* {', '.join(surfs[:4])}"

    tons, containers = calculate_weight_and_containers(stok_df, cols)
    logistics_info = f"\n⚖️ *Tahmini Stok Ağırlığı:* ~{tons} Ton (~{containers} Konteyner)" if stok_slabs > 0 else ""

    price_line = find_stone_price(stone_label)
    proforma_section = ""
    if unit_price and stok_m2 > 0:
        proforma_section = (
            f"\n\n💵 **PROFORMA FİYAT TEKLİFİ ({unit_price} {currency}/m²):**\n"
            f"• Tüm Stok Tutarı ({round(stok_m2, 2)} m²): **{round(stok_m2 * unit_price, 2):,.2f} {currency}**\n"
            f"• Serbest Stok Tutarı ({round(free_m2, 2)} m²): **{round(free_m2 * unit_price, 2):,.2f} {currency}**"
        )
    elif price_line:
        proforma_section = f"\n💰 *Kayıtlı Liste Fiyatı:* {price_line}"

    broken_section = f"\n⚠️ **Kırık / Fire:** {broken_slabs} Plaka ({round(broken_m2, 2)} m²)" if broken_slabs > 0 else ""
    target_lot = lots[0] if lots else None
    drive_url = get_best_drive_link(target_lot, stone_label)
    header_title = f"LOT {target_lot}" if target_lot else (stone_label.upper() if stone else "GENEL FABRİKA STOĞU")

    slab_range_info = ""
    if target_lot and cols["slab_no"] and not final_df.empty:
        slab_nums = pd.to_numeric(final_df[cols["slab_no"]], errors="coerce").dropna().astype(int).tolist()
        if slab_nums:
            slab_range_info = f"\n🔢 *Plaka No Aralığı:* #{min(slab_nums)} - #{max(slab_nums)} ({len(slab_nums)} Adet)"

    return (
        f"📊 **ARSLANLAR MARBLE - {header_title} RAPORU**\n"
        f"🪨 *Taş / Seleksiyon:* {stone_label or 'Genel'}{loc_info}{surf_info}{slab_range_info}{logistics_info}\n\n"
        f"🏭 **Toplam Üretilen:** {total_slabs} Plaka | **{round(total_m2, 2)} m²**\n\n"
        f"📦 **Depoda / Stokta Sağlam Mevcut:**\n"
        f"• Toplam Stok: **{stok_slabs} Plaka** | **{round(stok_m2, 2)} m²**\n"
        f"  ├─ 🟢 *Serbest (Müsait):* {free_slabs} Plaka ({round(free_m2, 2)} m²)\n"
        f"  └─ 🔒 *Rezervli:* {res_slabs} Plaka ({round(res_m2, 2)} m²)\n\n"
        f"🚛 **Yüklenen / Sevk Edilen:** {yuk_slabs} Plaka | **{round(yuk_m2, 2)} m²**"
        f"{broken_section}{proforma_section}\n\n"
        f"🔗 [İlgili Google Drive Klasörünü Aç]({drive_url})"
    )


def build_top_stones_summary(df: pd.DataFrame) -> str:
    cols = identify_columns(df)
    if df.empty or not cols["stone"] or not cols["m2"]:
        return "⚠️ Sıralama için gerekli sütunlar bulunamadı."
    temp = df.copy()
    temp["_m2_num"] = pd.to_numeric(temp[cols["m2"]], errors="coerce").fillna(0.0)
    healthy = temp[~get_broken_mask(temp)]
    stok_temp = healthy[healthy[cols["status"]].astype(str).apply(tr_norm).str.contains("stok", na=False)] if cols["status"] else healthy
    grouped_all = temp.groupby(cols["stone"])["_m2_num"].agg(["count", "sum"]).sort_values(by="sum", ascending=False).head(10)
    grouped_stok = stok_temp.groupby(cols["stone"])["_m2_num"].sum().to_dict()

    lines = ["🏆 **ARSLANLAR MARBLE - EN YÜKSEK METRAJLI 10 TAŞ (ÜRETİM & SAĞLAM STOK)**\n"]
    for idx, (stone_name, row) in enumerate(grouped_all.iterrows(), 1):
        lines.append(f"{idx}. **{stone_name}**\n   • Toplam: {round(row['sum'], 2)} m² ({int(row['count'])} Plaka) | 📦 **Stokta:** {round(grouped_stok.get(stone_name, 0.0), 2)} m²")
    return "\n".join(lines)


# ==============================================================================
# 8. GÖRSEL TAŞ EKSPERİ & SOHBET MOTORU
# ==============================================================================
def analyze_marble_photo_expert(image_bytes: bytes, caption: str, df: pd.DataFrame) -> str:
    cols = identify_columns(df)
    available_stones = df[cols["stone"]].dropna().astype(str).unique().tolist() if (not df.empty and cols["stone"]) else []
    stones_list_str = ", ".join(available_stones[:30]) if available_stones else "Silver Roots, Violet Breccia, Silver River, Golden Spider, Galapagos, Grey Philharmonic"

    prompt = (
        f"Sen 30 yıllık kıdemli bir Mermer Eksperisin (Arslanlar Marble). Kullanıcı fotoğraf gönderdi. Notu: '{caption}'.\n"
        f"Stoklarımızdaki taşlar: {stones_list_str}.\n"
        f"1. Fotoğraftaki taşın zemin rengini ve damar yapısını analiz et.\n"
        f"2. Stoklarımızdaki taşlardan hangisine en çok benzediğini tam isim vererek belirt."
    )

    ai_vision_result = None
    if GEMINI_API_KEY:
        try:
            b64_img = base64.b64encode(image_bytes).decode("utf-8")
            url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent"
            headers = {"Content-Type": "application/json"}
            if GEMINI_API_KEY.startswith("AIza"):
                url = f"{url}?key={GEMINI_API_KEY}"
            else:
                headers["Authorization"] = f"Bearer {GEMINI_API_KEY}"
                headers["x-goog-api-key"] = GEMINI_API_KEY
            payload = {"contents": [{"parts": [{"text": prompt}, {"inline_data": {"mime_type": "image/jpeg", "data": b64_img}}]}]}
            resp = requests.post(url, headers=headers, json=payload, timeout=15)
            if resp.status_code == 200:
                ai_vision_result = resp.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
        except Exception as e:
            logger.warning(f"Vision API uyarısı: {e}")

    matched_stock_text = ""
    if not df.empty and cols["stone"] and cols["m2"]:
        search_target = ai_vision_result or caption or ""
        found_st = next((st for st in available_stones if tr_norm(st) in tr_norm(search_target)), None)
        if found_st:
            matched_stock_text = "\n\n" + build_detailed_analysis_text(parse_user_query_and_filter(found_st, df), df)
        else:
            matched_stock_text = "\n\n" + build_top_stones_summary(df)

    if ai_vision_result:
        return f"🔍 **ARSLANLAR MARBLE - GÖRSEL TAŞ EKSPERTİZ RAPORU**\n\n{ai_vision_result}{matched_stock_text}"
    return f"🔍 **ARSLANLAR MARBLE - GÖRSEL EKSPERTİZ VE STOK EŞLEŞTİRME**\nGönderdiğiniz görsel incelendi. Mevcut koleksiyonlarımızla karşılaştırma tablosu:{matched_stock_text}"


def generate_expert_ai_response(user_text: str, df: pd.DataFrame, chat_id: int) -> str:
    cols = identify_columns(df)
    total_rows = len(df)
    m2_col = cols["m2"]
    total_m2 = round(float(pd.to_numeric(df[m2_col], errors="coerce").sum()), 2) if (not df.empty and m2_col) else 0.0
    mem = load_json_db(CHAT_MEMORY_PATH).get(str(chat_id), [])
    history_str = "\n".join([f"{m['role']}: {m['text']}" for m in mem[-6:]])

    system_prompt = (
        f"Sen Arslanlar Mermer'in üst düzey Fabrika ve İhracat Yapay Zeka Uzmanısın. "
        f"Kullanıcın Mehmet Morkoç. Fabrika veritabanında {total_rows} plaka ({total_m2} m²) kayıtlı.\n"
        f"Önceki konuşmalar:\n{history_str}\n\nKullanıcı mesajı: {user_text}\n"
        f"Doğrudan, zeki, profesyonel ve çözüm odaklı bir yanıt ver."
    )

    if GEMINI_API_KEY:
        try:
            url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent"
            headers = {"Content-Type": "application/json"}
            if GEMINI_API_KEY.startswith("AIza"):
                url = f"{url}?key={GEMINI_API_KEY}"
            else:
                headers["Authorization"] = f"Bearer {GEMINI_API_KEY}"
                headers["x-goog-api-key"] = GEMINI_API_KEY
            resp = requests.post(url, headers=headers, json={"contents": [{"parts": [{"text": system_prompt}]}]}, timeout=12)
            if resp.status_code == 200:
                return resp.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
        except Exception as e:
            logger.warning(f"AI REST uyarısı: {e}")

    t_norm = tr_norm(user_text)
    if any(k in t_norm for k in ["sunum", "tanitim", "mail", "teklif metni", "musteri", "ingilizce"]):
        top_stones_info = build_top_stones_summary(df) if not df.empty else ""
        return (
            f"📑 **ARSLANLAR MARBLE - KURUMSAL MÜŞTERİ SUNUM & TEKLİF TASLAĞI**\n\n"
            f"🇬🇧 **English Sales Pitch:**\n"
            f"*\"Dear Partner,*\n"
            f"*At Arslanlar Marble, we combine high-precision Italian Breton processing technology with our exclusive natural stone quarries. "
            f"Currently, we hold over {total_m2} m² ({total_rows} slabs) of premium selections in our warehouse, ready for global shipment. "
            f"Our signature selections—including Silver Roots, Violet Breccia, and Grey Philharmonic—are available in 2cm and 3cm thicknesses.\"*\n\n"
            f"{top_stones_info}"
        )
    return build_top_stones_summary(df) if not df.empty else "Merhaba Mehmet hocam! Argon Super-Master 7/24 aktif."


# ==============================================================================
# 9. TELEGRAM HANDLER FONKSİYONLARI (DOSYA, FOTOĞRAF, METİN)
# ==============================================================================
async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    file_name_lower = doc.file_name.lower()

    if file_name_lower.endswith(".json"):
        file = await doc.get_file()
        await file.download_to_drive(SERVICE_ACCOUNT_FILE)
        await safe_reply(
            update,
            "🔑 **Google Drive Service Account (JSON) Anahtarı Yüklendi!**\n"
            "Bot artık şirketinizin Google Drive klasörlerinde otonom olarak gezinebilir ve klasör içinden doğrudan fotoğraf çekebilir!"
        )
        return

    if not file_name_lower.endswith((".xlsx", ".xls", ".csv", ".xlsb")):
        await safe_reply(update, "⚠️ Lütfen geçerli bir Excel (.xlsx, .xlsb, .csv) veya Service Account (.json) dosyası gönderin.")
        return

    await safe_reply(update, f"⏳ `{doc.file_name}` inceleniyor; plaka, kırık/fire verileri ve Drive köprüleri hafızaya alınıyor...")
    file = await doc.get_file()
    local_path = f"./temp_{doc.file_name}"
    await file.download_to_drive(local_path)

    try:
        extracted_links = extract_excel_hyperlinks(local_path) if file_name_lower.endswith(".xlsx") else {}
        if extracted_links:
            save_json_db(HYPERLINKS_DB_PATH, extracted_links)

        if file_name_lower.endswith(".csv"):
            df = pd.read_csv(local_path)
        elif file_name_lower.endswith(".xlsb"):
            df = pd.read_excel(local_path, engine="pyxlsb")
        else:
            df = pd.read_excel(local_path, engine="openpyxl")

        df = clean_dataframe(df)
        if df.empty:
            await safe_reply(update, "⚠️ Dosya okundu ancak geçerli veri bulunamadı.")
            return

        cols_joined = " ".join([tr_norm(c) for c in df.columns])
        is_price_list = any(k in tr_norm(doc.file_name) for k in ["price", "fiyat"]) or (("fiyat" in cols_joined or "price" in cols_joined) and "plaka" not in cols_joined)

        if is_price_list:
            with bot_lock:
                df.to_excel(PERSISTENT_PRICE_PATH, index=False, engine="openpyxl")
            await safe_reply(update, f"💰 **Fiyat Listesi Hafızaya İşlendi!**\n• **Satır:** {len(df)} | **Drive Köprüsü:** {len(extracted_links)} Adet")
        else:
            with bot_lock:
                df.to_excel(PERSISTENT_STOCK_PATH, index=False, engine="openpyxl")
            cols_meta = identify_columns(df)
            m2_col = cols_meta["m2"]
            total_m2 = float(pd.to_numeric(df[m2_col], errors="coerce").sum()) if m2_col else 0.0
            b_mask = get_broken_mask(df)
            tons, containers = calculate_weight_and_containers(df[~b_mask], cols_meta)
            await safe_reply(
                update,
                f"✅ **Fabrika Stok Listesi Kalıcı Hafızaya Alındı!**\n"
                f"• **Toplam Kayıt:** {len(df)} Plaka | **{round(total_m2, 2)} m²**\n"
                f"• **Tespit Edilen Kırık/Fire:** {int(b_mask.sum())} Plaka\n"
                f"• **Sağlam Tonaj:** ~{tons} Ton (~{containers} Konteyner)\n"
                f"• **Yakalanan Drive Köprüsü:** {len(extracted_links)} Adet"
            )
    except Exception as e:
        logger.error(f"Dosya hatası: {e}")
        await safe_reply(update, f"⚠️ Dosya işlenirken hata oluştu: {e}")


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.message or not update.message.photo:
        return
    await safe_reply(update, "🔍 Fotoğraf alındı, damar ve renk dokusu stoklarımızla karşılaştırılıyor...")
    try:
        photo_file = await update.message.photo[-1].get_file()
        bio = io.BytesIO()
        await photo_file.download_to_memory(out=bio)
        bio.seek(0)
        caption = update.message.caption or "Bizde buna benzer hangi taş var?"
        await safe_reply(update, analyze_marble_photo_expert(bio.read(), caption, get_master_df()))
    except Exception as e:
        logger.error(f"Fotoğraf analiz hatası: {e}")
        await safe_reply(update, f"⚠️ Görsel incelenirken hata oluştu: {e}")


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.message or not update.message.text:
        return

    chat_id = update.effective_chat.id if update.effective_chat else 0
    text = update.message.text.strip()
    t_norm = tr_norm(text)

    append_chat_history(chat_id, "Kullanıcı", text)
    df_local = get_master_df()

    url_match = re.search(r"(https?://drive\.google\.com/\S+)", text)
    if url_match:
        new_url = url_match.group(1)
        lot_m = re.findall(r"\b([a-zA-Z]{1,3}[- ]?\d{2,5}[a-zA-Z]?)\b", text.replace(new_url, ""))
        if lot_m:
            l_code = normalize_lot_code(lot_m[0])
            save_json_db(HYPERLINKS_DB_PATH, {l_code: new_url, tr_norm(l_code): new_url})
            await safe_reply(update, f"✅ **{l_code}** için özel Google Drive bağlantısı kaydedildi!\n🔗 [Klasörü Aç]({new_url})")
            return

    if t_norm in ["merhaba", "selam", "orda misin", "orada misin", "hey", "bot", "test", "alo", "nasilsin"]:
        await safe_reply(update, f"Buradayım Mehmet hocam, Argon Super-Master 7/24 aktif! ⚡\n📦 Hafızada **{len(df_local)} plaka** hazır.")
        return

    if any(k in t_norm for k in ["en cok", "en yuksek", "sirala", "hangi taslar var", "tum taslar", "koleksiyonlar"]):
        if not df_local.empty:
            await safe_reply(update, build_top_stones_summary(df_local))
            return

    parsed = parse_user_query_and_filter(text, df_local, chat_id)
    lots, stone, status_filter = parsed["lots"], parsed["stone"], parsed["status_filter"]

    wants_excel = any(k in t_norm for k in ["ceki", "packing", "excel", "tablo", "liste"])
    wants_pdf_proforma = any(k in t_norm for k in ["proforma", "fatura", "pdf", "resmi teklif", "invoice"])
    wants_photo = any(k in t_norm for k in ["fotograf", "foto", "gorsel", "resim"])
    wants_link_only = ("link" in t_norm or "drive" in t_norm) and not wants_excel and not wants_photo
    is_presentation_or_ai = any(k in t_norm for k in ["sunum", "mail", "ingilizce", "tanitim", "strateji", "yaz", "hazirla"]) and not wants_pdf_proforma

    is_stock_question = (bool(lots or stone) or status_filter == "kirik" or any(
        k in t_norm for k in ["stok", "m2", "m²", "metraj", "plaka", "kac", "ne kadar", "analiz", "uretilen", "yuklenen", "rezerv", "serbest", "musait", "urun", "tas", "rapor", "ozet", "ton", "konteyner", "kirik", "fire"]
    )) and not is_presentation_or_ai

    if wants_photo or wants_link_only:
        target_key = lots[0] if lots else (stone or "")
        if not target_key:
            await safe_reply(update, f"🔗 **Arslanlar Marble Ana Drive Klasörü:**\n👉 [Klasörü Aç]({DEFAULT_DRIVE_LINK})")
            return

        drive_data = search_drive_folder_and_photos(target_key)
        f_url = drive_data.get("folder_url") or get_best_drive_link(lots[0] if lots else None, stone)
        await safe_reply(update, f"🔗 **{target_key} Doğrudan Google Drive Bağlantısı:**\n👉 [Klasörü ve Plaka Görsellerini Aç]({f_url})")

        if wants_photo:
            if drive_data.get("photo_bytes"):
                bio = io.BytesIO(drive_data["photo_bytes"])
                bio.name = drive_data.get("photo_name") or f"{target_key}.jpg"
                await update.message.reply_photo(photo=bio, caption=f"📸 **{target_key}** klasöründen çekilen plaka görseli (`{bio.name}`)")
            else:
                await safe_reply(update, f"ℹ️ **{target_key}** klasör linki paylaşıldı. Klasör içinden doğrudan `.jpg` çekebilmem için Google Cloud `.json` (Service Account) dosyasını bir kez bu sohbete göndermeniz yeterlidir.")
        return

    if wants_pdf_proforma:
        if df_local.empty:
            await safe_reply(update, "⚠️ Proforma üretmek için önce stok Excel dosyanızı yükleyin.")
            return
        target_df = parsed["final_df"] if not parsed["final_df"].empty else parsed["base_df"]
        if target_df.empty:
            target_df = df_local
        tag = lots[0] if lots else (stone.replace(" ", "_") if stone else "Genel")
        pdf_name = f"Arslanlar_Marble_Proforma_{tag}.pdf"
        res_pdf = generate_proforma_pdf(target_df, parsed, pdf_name)
        if res_pdf and os.path.exists(res_pdf):
            with open(res_pdf, "rb") as f:
                await update.message.reply_document(document=f, caption=f"📑 **Arslanlar Marble Resmi Proforma Invoice ({tag})**\nBirim Fiyat: {parsed['unit_price'] or 65.0} {parsed['currency']}/m²")
        else:
            await safe_reply(update, "⚠️ PDF oluşturulamadı (`reportlab` kütüphanesinin kurulu olduğundan emin olun).")
        return

    if is_stock_question or wants_excel:
        if df_local.empty:
            await safe_reply(update, "⚠️ Hafızada henüz stok tablosu yok. Lütfen önce Excel stok dosyanızı gönderin.")
            return

        report_text = build_detailed_analysis_text(parsed, df_local)
        await safe_reply(update, report_text)
        append_chat_history(chat_id, "Asistan", report_text)

        if wants_excel or lots or status_filter == "kirik":
            target_df = parsed["final_df"] if not parsed["final_df"].empty else parsed["base_df"]
            if not target_df.empty and len(target_df) <= 5000:
                tag_lot = lots[0] if lots else (stone.replace(" ", "_") if stone else "Genel")
                tag_stat = status_filter.upper()
                file_name = f"Arslanlar_Marble_Ceki_Listesi_{tag_lot}_{tag_stat}.xlsx"
                generate_styled_excel(target_df, file_name)
                m2_col = parsed["cols"]["m2"]
                sub_m2 = float(pd.to_numeric(target_df[m2_col], errors="coerce").sum()) if m2_col else 0.0
                tons, _ = calculate_weight_and_containers(target_df, parsed["cols"])
                with open(file_name, "rb") as f:
                    await update.message.reply_document(document=f, caption=f"📋 **{tag_lot} ({tag_stat}) Kurumsal Liste**\n📦 Plaka: {len(target_df)} Adet | 📐 Metraj: {round(sub_m2, 2)} m² | ⚖️ ~{tons} Ton")
        return

    ai_reply = generate_expert_ai_response(text, df_local, chat_id)
    await safe_reply(update, ai_reply)
    append_chat_history(chat_id, "Asistan", ai_reply)


# ==============================================================================
# 10. RENDER 7/24 KEEP-ALIVE VE ANA DÖNGÜ
# ==============================================================================
class HealthCheckHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Arslanlar Marble Argon Super-Master Bot 7/24 Aktif!")

    def log_message(self, format, *args):
        return


def keep_alive_pinger():
    while True:
        time.sleep(240)
        try:
            if RENDER_EXTERNAL_URL:
                requests.get(RENDER_EXTERNAL_URL, timeout=10)
        except Exception:
            pass


def main():
    port = int(os.environ.get("PORT", 8080))
    threading.Thread(target=lambda: HTTPServer(("0.0.0.0", port), HealthCheckHandler).serve_forever(), daemon=True).start()
    threading.Thread(target=keep_alive_pinger, daemon=True).start()

    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))

    logger.info("Arslanlar Mermer Argon Super-Master Bot başlatılıyor...")
    app.run_polling(drop_pending_updates=True, allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
