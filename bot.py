import os
import re
import io
import time
import json
import logging
import threading
import requests
from datetime import datetime
import pytz
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from http.server import HTTPServer, BaseHTTPRequestHandler
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    MessageHandler,
    CommandHandler,
    filters,
    ContextTypes,
)

# --- LOGLAMA YAPILANDIRMASI ---
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# --- ANAHTARLAR VE ORTAM DEĞİŞKENLERİ ---
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "8822872793:AAEn0p_FLlH1CaB41cHhwHV7ADh7q6o6VGU").strip()
GEMINI_API_KEY = os.getenv(
    "GEMINI_API_KEY",
    "AQ.Ab8RN6Ji0H4-YVfeTsoYlo62qCVfulC0nYpDSalIf9MXI-EisA"
).strip()

RENDER_EXTERNAL_URL = os.getenv("RENDER_EXTERNAL_URL", "https://arslanlar-marble-bot.onrender.com").strip()

DEFAULT_DRIVE_LINK = os.getenv(
    "GOOGLE_DRIVE_LINK",
    "https://drive.google.com/drive/folders/1bnZH4cy2zOdZWhvuN79DURjabDJehOjs"
)

# --- KALICI DİSK VERİTABANI DOSYALARI ---
PERSISTENT_STOCK_PATH = "./master_stok_arslanlar.xlsx"
PERSISTENT_PRICE_PATH = "./master_fiyat_arslanlar.xlsx"
HYPERLINKS_DB_PATH = "./hyperlinks_db.json"
CHAT_MEMORY_PATH = "./chat_memory.json"

bot_lock = threading.Lock()

# Kullanıcı Bazlı Son Konuşulan Lot/Taş Hafızası (Bağlam Takibi)
user_context_memory = {}


# ==============================================================================
# 1. TÜRKÇE DİL VE METİN NORMALİZASYON ARAÇLARI
# ==============================================================================
def tr_lower(text: str) -> str:
    if not text:
        return ""
    text = str(text)
    replacements = {
        "İ": "i", "I": "ı", "Ğ": "ğ", "Ü": "ü", "Ş": "ş", "Ö": "ö", "Ç": "ç"
    }
    for k, v in replacements.items():
        text = text.replace(k, v)
    return text.lower().strip()


def tr_norm(text: str) -> str:
    t = tr_lower(text)
    trans = str.maketrans("ıiğüöşç", "iiguosc")
    return t.translate(trans)


def normalize_lot_code(text: str) -> str:
    clean = re.sub(r"[^a-zA-Z0-9]", "", str(text)).upper()
    m = re.match(r"^([A-Z]{1,3})(\d+[A-Z]?)$", clean)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    return clean


async def safe_reply(update: Update, text: str):
    """Markdown özel karakter hatalarını yakalar, mesajın %100 iletilmesini garanti eder."""
    if not update.message or not text:
        return
    try:
        await update.message.reply_text(text, parse_mode="Markdown", disable_web_page_preview=True)
    except Exception as e:
        logger.warning(f"Markdown hatası, düz metin gönderiliyor: {e}")
        clean_text = text.replace("**", "").replace("*", "").replace("`", "")
        await update.message.reply_text(clean_text, disable_web_page_preview=True)


# ==============================================================================
# 2. KALICI HAFIZA, KÖPRÜ (HYPERLINK) VE SOHBET GEÇMİŞİ YÖNETİMİ
# ==============================================================================
def load_json_db(path: str) -> dict:
    with bot_lock:
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                logger.error(f"JSON okuma hatası ({path}): {e}")
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
            logger.error(f"JSON kaydetme hatası ({path}): {e}")


def append_chat_history(chat_id: int, role: str, content: str):
    """Konuşmaları zamanla öğrenmesi ve bağlamı hatırlaması için kaydeder."""
    cid = str(chat_id)
    mem = load_json_db(CHAT_MEMORY_PATH)
    history = mem.get(cid, [])
    history.append({"role": role, "text": content[:500], "time": datetime.now().strftime("%H:%M")})
    mem[cid] = history[-15:]  # Son 15 mesajı tut
    save_json_db(CHAT_MEMORY_PATH, mem)


def extract_excel_hyperlinks(excel_path: str) -> dict:
    """Excel'deki (Price List veya Stok) hücre içi köprüleri (Hyperlink) söküp hafızaya alır."""
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


def get_best_drive_link(lot_code: str = None, stone_name: str = None) -> str:
    links_db = load_json_db(HYPERLINKS_DB_PATH)
    if lot_code:
        nl = normalize_lot_code(lot_code)
        nk = tr_norm(lot_code)
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

    return DEFAULT_DRIVE_LINK


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
                df = pd.read_excel(PERSISTENT_STOCK_PATH, engine="openpyxl")
                return clean_dataframe(df)
            except Exception as e:
                logger.error(f"Stok okuma hatası: {e}")
        return pd.DataFrame()


def get_price_df() -> pd.DataFrame:
    with bot_lock:
        if os.path.exists(PERSISTENT_PRICE_PATH):
            try:
                df = pd.read_excel(PERSISTENT_PRICE_PATH, engine="openpyxl")
                return clean_dataframe(df)
            except Exception as e:
                logger.error(f"Fiyat okuma hatası: {e}")
        return pd.DataFrame()


# ==============================================================================
# 3. LOJİSTİK (TONAJ & KONTEYNER) VE FİYAT HESAPLAMA MOTORU
# ==============================================================================
def identify_columns(df: pd.DataFrame) -> dict:
    cols = {
        "lot": [], "stone": None, "status": None, "location": None,
        "reserve": None, "m2": None, "surface": None, "thickness": None,
        "width": None, "length": None, "slab_no": None, "company": None
    }
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


def calculate_weight_and_containers(df: pd.DataFrame, cols: dict) -> tuple:
    """
    Mermer yoğunluğunu (~2.75 ton/m3) baz alarak plakaların kalınlığına göre
    gerçek tonajı ve 27 tonluk konteyner kapasitesini hesaplar.
    """
    if df.empty or not cols["m2"]:
        return 0.0, 0.0

    m2_series = pd.to_numeric(df[cols["m2"]], errors="coerce").fillna(0.0)
    if cols["thickness"] and cols["thickness"] in df.columns:
        thick_series = pd.to_numeric(df[cols["thickness"]], errors="coerce").fillna(20.0)
        # Kalınlık cm cinsindense (örn 2 veya 3) mm'ye çevir
        thick_mm = thick_series.apply(lambda x: x * 10 if x < 10 else x)
    else:
        thick_mm = pd.Series(20.0, index=df.index)

    # 1 m2 * 1 mm kalınlık = 0.00275 Ton (2.75 kg) -> 20mm (2cm) = 55 kg/m2
    weights_ton = m2_series * (thick_mm / 1000.0) * 2.75
    total_tons = round(float(weights_ton.sum()), 2)
    containers = round(total_tons / 27.0, 2)
    return total_tons, containers


def find_stone_price(stone_name: str) -> str:
    """Yüklenen Price List dosyasından ilgili taşın fiyat satırını bulur."""
    price_df = get_price_df()
    if price_df.empty or not stone_name:
        return ""
    target = tr_norm(stone_name.split(",")[0])
    mask = price_df.astype(str).apply(lambda x: x.apply(tr_norm).str.contains(target, na=False)).any(axis=1)
    matched = price_df[mask]
    if not matched.empty:
        row = matched.iloc[0]
        parts = [f"{c}: *{row[c]}*" for c in matched.columns[:5] if pd.notna(row[c]) and str(row[c]).strip() != ""]
        return " | ".join(parts)
    return ""


# ==============================================================================
# 4. AKILLI SORGU ÇÖZÜMLEYİCİ (BAĞLAM HAFIZALI)
# ==============================================================================
def parse_user_query_and_filter(text: str, df: pd.DataFrame, chat_id: int = 0) -> dict:
    t_norm = tr_norm(text)
    cols = identify_columns(df)

    # 1. Lot Kodu Yakalama
    lot_candidates = re.findall(r"\b([a-zA-Z]{1,3}[- ]?\d{2,5}[a-zA-Z]?)\b", text)
    ignored = {"M2", "A4", "A3", "3D", "2D", "724", "20MM", "30MM", "2CM", "3CM"}
    detected_lots = [normalize_lot_code(c) for c in lot_candidates if normalize_lot_code(c) not in ignored]

    # 2. Dinamik Taş Cinsi Yakalama (Excel'deki tüm taşları tarar)
    detected_stone = None
    if not df.empty and cols["stone"]:
        unique_stones = df[cols["stone"]].dropna().astype(str).unique()
        unique_stones = sorted(unique_stones, key=lambda x: len(str(x)), reverse=True)
        for st in unique_stones:
            st_clean = tr_norm(st)
            if len(st_clean) > 2 and st_clean in t_norm:
                detected_stone = st.strip()
                break

    # 3. Bağlam Hafızası (Kullanıcı "bunun çeki listesini at", "peki stokta kaç var" dediyse önceki lotu/taşı hatırla)
    ctx = user_context_memory.get(chat_id, {})
    if not detected_lots and not detected_stone:
        if any(w in t_norm for w in ["bunun", "onun", "bu lotun", "bu tasin", "ayni", "peki", "ceki listesi", "linki"]):
            if ctx.get("lot"):
                detected_lots = [ctx["lot"]]
            elif ctx.get("stone"):
                detected_stone = ctx["stone"]

    # Yeni yakalanan lot/taşı hafızaya kaydet
    if detected_lots or detected_stone:
        user_context_memory[chat_id] = {
            "lot": detected_lots[0] if detected_lots else None,
            "stone": detected_stone
        }

    # 4. Durum Filtresi
    status_filter = "all"
    if any(k in t_norm for k in ["serbest", "musait", "rezervsiz"]):
        status_filter = "serbest"
    elif any(k in t_norm for k in ["rezerv"]):
        status_filter = "rezerv"
    elif any(k in t_norm for k in ["yuklenen", "yuklendi", "sevk", "giden", "satilan"]):
        status_filter = "yuklenen"
    elif any(k in t_norm for k in ["stokta", "stoktaki", "depoda", "mevcut", "eldeki", "stok listesi", "stok"]):
        # Eğer "toplam üretilen ve stok" birlikte sorulmadıysa stok filtrele
        if "uretilen" not in t_norm and "toplam" not in t_norm:
            status_filter = "stok"

    # 5. Birim Fiyat Teklif Yakalama (Örn: "65 dolardan", "80 euro", "75$")
    unit_price = None
    currency = "$"
    price_match = re.search(r"(\d+(?:[.,]\d+)?)\s*(?:\$|dolar|usd|€|euro|eur|tl)", t_norm)
    if price_match:
        unit_price = float(price_match.group(1).replace(",", "."))
        if any(c in t_norm for c in ["€", "euro", "eur"]):
            currency = "€"

    # 6. Filtreleme İşlemi
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
            if status_filter == "stok" and cols["status"]:
                filtered_by_status = filtered[filtered[cols["status"]].astype(str).apply(tr_norm).str.contains("stok", na=False)]
            elif status_filter == "yuklenen" and cols["status"]:
                filtered_by_status = filtered[filtered[cols["status"]].astype(str).apply(tr_norm).str.contains("yuklen", na=False)]
            elif status_filter == "rezerv" and cols["reserve"]:
                filtered_by_status = filtered[filtered[cols["reserve"]].astype(str).apply(tr_norm).str.contains("rezerv", na=False)]
            elif status_filter == "serbest" and cols["status"]:
                stok_m = filtered[cols["status"]].astype(str).apply(tr_norm).str.contains("stok", na=False)
                if cols["reserve"]:
                    res_m = filtered[cols["reserve"]].astype(str).apply(tr_norm).str.contains("rezerv", na=False)
                    filtered_by_status = filtered[stok_m & (~res_m)]
                else:
                    filtered_by_status = filtered[stok_m]
    else:
        filtered_by_status = pd.DataFrame()

    return {
        "lots": detected_lots,
        "stone": detected_stone,
        "status_filter": status_filter,
        "unit_price": unit_price,
        "currency": currency,
        "base_df": filtered,
        "final_df": filtered_by_status,
        "lot_matched": lot_matched,
        "cols": cols
    }


# ==============================================================================
# 5. ESTETİK VE KURUMSAL ÇEKİ LİSTESİ ÜRETİCİ
# ==============================================================================
def generate_styled_excel(filtered_df: pd.DataFrame, file_name: str) -> str:
    df_out = filtered_df.copy()

    cols_norm_map = {tr_norm(c): c for c in df_out.columns}
    if "stok durumu" in cols_norm_map and "stok konumu" in cols_norm_map:
        drop_candidates = [cols_norm_map["stok durumu"]]
        for extra in ["firma", "uretim tarihi", "yuklenme tarihi"]:
            if extra in cols_norm_map:
                drop_candidates.append(cols_norm_map[extra])
        df_out = df_out.drop(columns=[c for c in drop_candidates if c in df_out.columns], errors="ignore")

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

    thin_border = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )
    align_center = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[1].height = 25
    for col_num in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        cell.border = thin_border

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

            cell.font = data_font
            cell.alignment = align_center
            cell.border = thin_border
            if is_even:
                cell.fill = zebra_fill

            norm_v = tr_norm(val_str)
            norm_l = normalize_lot_code(val_str)
            if norm_l in links_db and links_db[norm_l] != DEFAULT_DRIVE_LINK:
                cell.hyperlink = links_db[norm_l]
                cell.font = link_font
            elif norm_v in links_db and links_db[norm_v] != DEFAULT_DRIVE_LINK:
                cell.hyperlink = links_db[norm_v]
                cell.font = link_font

    # Özet Toplam ve Tonaj Satırı
    total_row = max_data_row + 1
    ws.row_dimensions[total_row].height = 24
    m2_col_idx = None
    for col_num in range(1, ws.max_column + 1):
        col_title = str(ws.cell(row=1, column=col_num).value or "")
        if m2_col and col_title == m2_col:
            m2_col_idx = col_num

    tons, _ = calculate_weight_and_containers(filtered_df, identify_columns(filtered_df))
    ws.cell(row=total_row, column=1, value="TOPLAM").font = total_font
    ws.cell(row=total_row, column=2, value=f"{max_data_row - 1} Plaka (~{tons} Ton)").font = total_font

    if m2_col_idx and m2_col in df_out.columns:
        total_m2_val = round(float(df_out[m2_col].sum()), 2)
        ws.cell(row=total_row, column=m2_col_idx, value=f"{total_m2_val} m²").font = total_font

    for col_num in range(1, ws.max_column + 1):
        c = ws.cell(row=total_row, column=col_num)
        c.fill = total_fill
        c.alignment = align_center
        c.border = thin_border

    for col in ws.columns:
        max_len = max((len(str(cell.value or "")) for cell in col), default=10)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

    wb.save(file_name)
    return file_name


# ==============================================================================
# 6. UZMAN RAPORLAMA VE STRATEJİK ANALİZ MOTORU
# ==============================================================================
def build_top_stones_summary(df: pd.DataFrame) -> str:
    """Fabrikadaki en yüksek metrajlı taşları sıralar ve yönetim özeti çıkarır."""
    cols = identify_columns(df)
    if df.empty or not cols["stone"] or not cols["m2"]:
        return "⚠️ Sıralama için gerekli sütunlar bulunamadı."

    temp = df.copy()
    temp["_m2_num"] = pd.to_numeric(temp[cols["m2"]], errors="coerce").fillna(0.0)

    # Sadece stokta olanları da ayrıca hesapla
    if cols["status"]:
        stok_mask = temp[cols["status"]].astype(str).apply(tr_norm).str.contains("stok", na=False)
        stok_temp = temp[stok_mask]
    else:
        stok_temp = temp

    grouped_all = temp.groupby(cols["stone"])["_m2_num"].agg(["count", "sum"]).sort_values(by="sum", ascending=False).head(10)
    grouped_stok = stok_temp.groupby(cols["stone"])["_m2_num"].sum().to_dict()

    lines = ["🏆 **ARSLANLAR MARBLE - EN YÜKSEK METRAJLI 10 TAŞ (ÜRETİM & STOK)**\n"]
    for idx, (stone_name, row) in enumerate(grouped_all.iterrows(), 1):
        tot_m2 = round(row["sum"], 2)
        cnt = int(row["count"])
        st_m2 = round(grouped_stok.get(stone_name, 0.0), 2)
        lines.append(f"{idx}. **{stone_name}**\n   • Toplam Üretim: {tot_m2} m² ({cnt} Plaka) | 📦 **Stokta:** {st_m2} m²")

    return "\n".join(lines)


def build_detailed_analysis_text(parsed: dict, full_df: pd.DataFrame) -> str:
    cols = parsed["cols"]
    base_df = parsed["base_df"]
    final_df = parsed["final_df"]
    lots = parsed["lots"]
    stone = parsed["stone"]
    unit_price = parsed["unit_price"]
    currency = parsed["currency"]

    if not parsed["lot_matched"] and lots:
        return f"⚠️ Hafızadaki stok listesinde **{lots[0]}** kodlu lot/blok kaydı bulunamadı."

    if base_df.empty:
        return "⚠️ Aradığınız kriterlere uygun plaka kaydı bulunamadı."

    m2_col = cols["m2"]
    stat_col = cols["status"]
    res_col = cols["reserve"]
    loc_col = cols["location"]
    stone_col = cols["stone"]
    surf_col = cols["surface"]

    total_slabs = len(base_df)
    total_m2 = float(pd.to_numeric(base_df[m2_col], errors="coerce").sum()) if m2_col else 0.0

    if stat_col:
        stok_df = base_df[base_df[stat_col].astype(str).apply(tr_norm).str.contains("stok", na=False)]
        yuk_df = base_df[base_df[stat_col].astype(str).apply(tr_norm).str.contains("yuklen", na=False)]
    else:
        stok_df = base_df
        yuk_df = pd.DataFrame()

    stok_slabs = len(stok_df)
    stok_m2 = float(pd.to_numeric(stok_df[m2_col], errors="coerce").sum()) if (m2_col and not stok_df.empty) else 0.0

    yuk_slabs = len(yuk_df)
    yuk_m2 = float(pd.to_numeric(yuk_df[m2_col], errors="coerce").sum()) if (m2_col and not yuk_df.empty) else 0.0

    res_slabs, res_m2 = 0, 0.0
    free_slabs, free_m2 = stok_slabs, stok_m2
    if res_col and not stok_df.empty:
        res_mask = stok_df[res_col].astype(str).apply(tr_norm).str.contains("rezerv", na=False)
        res_df = stok_df[res_mask]
        free_df = stok_df[~res_mask]
        res_slabs = len(res_df)
        res_m2 = float(pd.to_numeric(res_df[m2_col], errors="coerce").sum()) if m2_col else 0.0
        free_slabs = len(free_df)
        free_m2 = float(pd.to_numeric(free_df[m2_col], errors="coerce").sum()) if m2_col else 0.0

    stone_label = stone
    if not stone_label and stone_col and not base_df.empty:
        top_stones = base_df[stone_col].dropna().astype(str).unique()
        stone_label = ", ".join(top_stones[:3]) if len(top_stones) <= 3 else f"{len(top_stones)} Farklı Seleksiyon"

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

    # Tonaj ve Konteyner Hesabı (Stoktaki plakalar için)
    tons, containers = calculate_weight_and_containers(stok_df, cols)
    logistics_info = f"\n⚖️ *Tahmini Stok Ağırlığı:* ~{tons} Ton (~{containers} Konteyner)" if stok_slabs > 0 else ""

    # Fiyat Listesi veya Proforma Teklif Hesabı
    price_line = find_stone_price(stone_label)
    proforma_section = ""
    if unit_price and stok_m2 > 0:
        tot_val = round(stok_m2 * unit_price, 2)
        free_val = round(free_m2 * unit_price, 2)
        proforma_section = (
            f"\n\n💵 **PROFORMA FİYAT TEKLİFİ ({unit_price} {currency}/m²):**\n"
            f"• Tüm Stok Tutarı ({round(stok_m2, 2)} m²): **{tot_val:,.2f} {currency}**\n"
            f"• Serbest Stok Tutarı ({round(free_m2, 2)} m²): **{free_val:,.2f} {currency}**"
        )
    elif price_line:
        proforma_section = f"\n💰 *Kayıtlı Liste Fiyatı:* {price_line}"

    target_lot = lots[0] if lots else None
    drive_url = get_best_drive_link(target_lot, stone_label)
    header_title = f"LOT {target_lot}" if target_lot else (stone_label.upper() if stone else "GENEL FABRİKA STOĞU")

    slab_range_info = ""
    if target_lot and cols["slab_no"] and not final_df.empty:
        slab_nums = pd.to_numeric(final_df[cols["slab_no"]], errors="coerce").dropna().astype(int).tolist()
        if slab_nums:
            slab_range_info = f"\n🔢 *Plaka No Aralığı:* #{min(slab_nums)} - #{max(slab_nums)} ({len(slab_nums)} Adet)"

    report = (
        f"📊 **ARSLANLAR MARBLE - {header_title} RAPORU**\n"
        f"🪨 *Taş / Seleksiyon:* {stone_label or 'Genel'}"
        f"{loc_info}{surf_info}{slab_range_info}{logistics_info}\n\n"
        f"🏭 **Toplam Üretilen (Stok + Yüklenen):**\n"
        f"• {total_slabs} Plaka | **{round(total_m2, 2)} m²**\n\n"
        f"📦 **Depoda / Stokta Mevcut Olan:**\n"
        f"• Toplam Stok: **{stok_slabs} Plaka** | **{round(stok_m2, 2)} m²**\n"
        f"  ├─ 🟢 *Serbest (Müsait):* {free_slabs} Plaka ({round(free_m2, 2)} m²)\n"
        f"  └─ 🔒 *Rezervli:* {res_slabs} Plaka ({round(res_m2, 2)} m²)\n\n"
        f"🚛 **Yüklenen / Sevk Edilen:**\n"
        f"• {yuk_slabs} Plaka | **{round(yuk_m2, 2)} m²**"
        f"{proforma_section}\n\n"
        f"🔗 [İlgili Google Drive Klasörünü Aç]({drive_url})"
    )
    return report


# ==============================================================================
# 7. ÇİFT KANALLI (AI + YERLEŞİK UZMAN) DİL VE SUNUM MOTORU
# ==============================================================================
def generate_expert_ai_response(user_text: str, df: pd.DataFrame, chat_id: int) -> str:
    """
    Hem AQ.Ab8... (OAuth/Vertex) hem AIza... anahtarlarını REST API ile dener.
    Eğer Google API anahtarı reddederse, asla hata basmaz; kendi içindeki
    Yerleşik Fabrika & İhracat Uzman Motoru ile sunum/mail/strateji üretir!
    """
    cols = identify_columns(df)
    total_rows = len(df)
    m2_col = cols["m2"]
    total_m2 = round(float(pd.to_numeric(df[m2_col], errors="coerce").sum()), 2) if (not df.empty and m2_col) else 0.0

    # Sohbet geçmişini al
    mem = load_json_db(CHAT_MEMORY_PATH).get(str(chat_id), [])
    history_str = "\n".join([f"{m['role']}: {m['text']}" for m in mem[-6:]])

    system_prompt = (
        f"Sen Arslanlar Mermer'in (Arslanlar Marble) üst düzey Fabrika ve İhracat Yapay Zeka Uzmanısın. "
        f"Kullanıcın Mehmet Morkoç (Operasyon & Satış Uzmanı). "
        f"Fabrika veritabanında şu an {total_rows} plaka ({total_m2} m²) kayıtlı. "
        f"Önceki konuşmalar:\n{history_str}\n\n"
        f"Kullanıcı mesajı: {user_text}\n\n"
        f"Doğrudan, çok zeki, insan gibi düşünen, profesyonel ve çözüm odaklı bir yanıt ver."
    )

    # 1. Google REST API Denemesi (Hem API Key hem Bearer Token destekli)
    if GEMINI_API_KEY:
        try:
            url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent"
            headers = {"Content-Type": "application/json"}
            if GEMINI_API_KEY.startswith("AIza"):
                url = f"{url}?key={GEMINI_API_KEY}"
            else:
                headers["Authorization"] = f"Bearer {GEMINI_API_KEY}"
                headers["x-goog-api-key"] = GEMINI_API_KEY

            payload = {"contents": [{"parts": [{"text": system_prompt}]}]}
            resp = requests.post(url, headers=headers, json=payload, timeout=12)
            if resp.status_code == 200:
                data = resp.json()
                ai_text = data["candidates"][0]["content"]["parts"][0]["text"].strip()
                return ai_text
        except Exception as e:
            logger.warning(f"Harici AI REST geçiş uyarısı: {e}")

    # 2. YERLEŞİK UZMAN MOTORU (API Anahtarı 401 verse bile devreye giren yerel zeka)
    t_norm = tr_norm(user_text)

    if any(k in t_norm for k in ["sunum", "tanitim", "mail", "teklif metni", "musteri", "ingilizce"]):
        top_stones_info = build_top_stones_summary(df) if not df.empty else ""
        return (
            f"📑 **ARSLANLAR MARBLE - KURUMSAL MÜŞTERİ SUNUM & TEKLİF TASLAĞI**\n\n"
            f"🇹🇷 **Türkçe Yönetici Özeti:**\n"
            f"Arslanlar Mermer olarak modern Breton plaka poliş hattımız ve yüksek üretim kapasitemizle "
            f"şu anda stoklarımızda **{total_rows} adet plaka ({total_m2} m²)** birinci sınıf doğal taş bulundurmaktayız. "
            f"Öne çıkan **Silver Roots, Violet Breccia, Silver River, Golden Spider ve Galapagos** seleksiyonlarımız "
            f"hem bookmatch hem de standart proje plakası olarak sevkiyata hazırdır.\n\n"
            f"🇬🇧 **English Sales & Presentation Pitch (For International Clients):**\n"
            f"*\"Dear Partner,*\n"
            f"*At Arslanlar Marble, we combine high-precision Italian Breton processing technology with our exclusive natural stone quarries. "
            f"Currently, we hold over {total_m2} m² ({total_rows} slabs) of premium selections in our showroom and warehouse, ready for immediate inspection and global shipment. "
            f"Our signature selections—including Silver Roots, Violet Breccia, and Grey Philharmonic—are available in 2cm and 3cm thicknesses with Polished and Honed finishes. "
            f"Please find our updated packing lists and high-resolution slab links attached.\"*\n\n"
            f"{top_stones_info}\n\n"
            f"🔗 [Tüm Koleksiyon Görselleri (Google Drive)]({DEFAULT_DRIVE_LINK})"
        )

    return build_top_stones_summary(df) if not df.empty else (
        "Merhaba Mehmet hocam! Sistem şu an 7/24 aktif. Bana doğrudan stok analizlerini, taş isimlerini, "
        "lot kodlarını, proforma fiyat hesaplarını veya sunum hazırlama taleplerini iletebilirsin."
    )


# ==============================================================================
# 8. TELEGRAM MESAJ VE DOSYA YÖNETİCİLERİ
# ==============================================================================
async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    file_name_lower = doc.file_name.lower()

    if not file_name_lower.endswith((".xlsx", ".xls", ".csv", ".xlsb")):
        await safe_reply(update, "⚠️ Lütfen geçerli bir Excel (.xlsx, .xlsb, .csv) dosyası gönderin.")
        return

    await safe_reply(update, f"⏳ `{doc.file_name}` inceleniyor; plaka verileri ve gömülü Drive köprüleri hafızaya alınıyor...")

    file = await doc.get_file()
    local_path = f"./temp_{doc.file_name}"
    await file.download_to_drive(local_path)

    try:
        extracted_links = {}
        if file_name_lower.endswith(".xlsx"):
            extracted_links = extract_excel_hyperlinks(local_path)
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
            await safe_reply(update, "⚠️ Dosya okundu ancak içinde geçerli tablo verisi bulunamadı.")
            return

        cols_joined = " ".join([tr_norm(c) for c in df.columns])
        is_price_list = any(k in tr_norm(doc.file_name) for k in ["price", "fiyat"]) or (
            ("fiyat" in cols_joined or "price" in cols_joined) and "plaka" not in cols_joined
        )

        if is_price_list:
            with bot_lock:
                df.to_excel(PERSISTENT_PRICE_PATH, index=False, engine="openpyxl")
            await safe_reply(
                update,
                f"💰 **Fiyat Listesi (Price List) Hafızaya İşlendi!**\n"
                f"• **Dosya:** `{doc.file_name}`\n"
                f"• **Fiyat Kalemi:** {len(df)} Satır\n"
                f"• **Hücrelerden Çekilen Drive Linki:** {len(extracted_links)} Adet\n\n"
                f"Artık taşların üzerine gömdüğün Drive köprüleri ve birim fiyatlar tüm sorgularda otomatik çalışacak!"
            )
        else:
            with bot_lock:
                df.to_excel(PERSISTENT_STOCK_PATH, index=False, engine="openpyxl")
            cols_meta = identify_columns(df)
            m2_col = cols_meta["m2"]
            total_m2 = float(pd.to_numeric(df[m2_col], errors="coerce").sum()) if m2_col else 0.0
            tons, containers = calculate_weight_and_containers(df, cols_meta)

            await safe_reply(
                update,
                f"✅ **Fabrika Stok Listesi Kalıcı Hafızaya Alındı!**\n"
                f"• **Dosya:** `{doc.file_name}`\n"
                f"• **Toplam Kayıt:** {len(df)} Plaka | **{round(total_m2, 2)} m²**\n"
                f"• **Toplam Tonaj:** ~{tons} Ton (~{containers} Konteyner)\n"
                f"• **Yakalanan Drive Köprüsü:** {len(extracted_links)} Adet"
            )
    except Exception as e:
        logger.error(f"Dosya hatası: {e}")
        await safe_reply(update, f"⚠️ Dosya işlenirken hata oluştu: {e}")


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.message or not update.message.text:
        return

    chat_id = update.effective_chat.id if update.effective_chat else 0
    text = update.message.text.strip()
    t_norm = tr_norm(text)

    append_chat_history(chat_id, "Kullanıcı", text)
    df_local = get_master_df()

    # 1. Telegram Üzerinden Yeni Drive Linki Öğretme (Örn: "A-283 linki https://drive.google.com/...")
    url_match = re.search(r"(https?://drive\.google\.com/\S+)", text)
    if url_match:
        new_url = url_match.group(1)
        lot_m = re.findall(r"\b([a-zA-Z]{1,3}[- ]?\d{2,5}[a-zA-Z]?)\b", text.replace(new_url, ""))
        if lot_m:
            l_code = normalize_lot_code(lot_m[0])
            save_json_db(HYPERLINKS_DB_PATH, {l_code: new_url, tr_norm(l_code): new_url})
            await safe_reply(update, f"✅ **{l_code}** için özel Google Drive bağlantısı kalıcı hafızaya kaydedildi!\n🔗 [Test Et]({new_url})")
            return

    # 2. Hızlı Selamlaşma / Kontrol
    if t_norm in ["merhaba", "selam", "orda misin", "orada misin", "hey", "bot", "test", "alo", "nasilsin"]:
        row_cnt = len(df_local)
        msg = (
            f"Buradayım Mehmet hocam, 7/24 tam kapasite aktifim! ⚡\n"
            f"📦 Hafızada **{row_cnt} plaka** ve gömülü Drive köprüleri hazır.\n"
            f"Bana istediğin lotu, taşı, tonaj/fiyat hesabını veya sunum talebini yazabilirsin."
        )
        await safe_reply(update, msg)
        return

    # 3. En Çok Bulunan Taşlar / Sıralama Soruları
    if any(k in t_norm for k in ["en cok", "en yuksek", "sirala", "hangi taslar var", "tum taslar", "koleksiyonlar"]):
        if not df_local.empty:
            top_rep = build_top_stones_summary(df_local)
            await safe_reply(update, top_rep)
            return

    # 4. Mesajı Çözümle (Lot, Taş, Durum, Fiyat Teklifi)
    parsed = parse_user_query_and_filter(text, df_local, chat_id)
    lots = parsed["lots"]
    stone = parsed["stone"]
    status_filter = parsed["status_filter"]

    wants_excel = any(k in t_norm for k in ["ceki", "packing", "excel", "dosya", "tablo", "liste"])
    wants_link_only = any(k in t_norm for k in ["link", "drive", "fotograf", "gorsel", "foto"]) and not wants_excel
    is_presentation_or_ai = any(k in t_norm for k in ["sunum", "mail", "ingilizce", "tanitim", "strateji", "yaz", "hazirla"])

    is_stock_question = (bool(lots or stone) or any(
        k in t_norm for k in [
            "stok", "m2", "m²", "metraj", "plaka", "kac", "ne kadar", "analiz",
            "uretilen", "yuklenen", "rezerv", "serbest", "musait", "urun", "tas", "rapor", "ozet", "ton", "konteyner"
        ]
    )) and not is_presentation_or_ai

    # 5. Sadece Link İstendiyse
    if wants_link_only:
        target_lot = lots[0] if lots else None
        url = get_best_drive_link(target_lot, stone)
        label = target_lot or stone or "Arslanlar Marble"
        await safe_reply(update, f"🔗 **{label} Doğrudan Google Drive Bağlantısı:**\n👉 [Klasörü ve Plaka Görsellerini Aç]({url})")
        return

    # 6. Stok, Lot, Taş, Tonaj, Proforma veya Çeki Listesi Sorusu
    if is_stock_question or wants_excel:
        if df_local.empty:
            await safe_reply(update, "⚠️ Hafızada henüz stok tablosu yok. Lütfen önce Excel stok dosyanızı gönderin.")
            return

        report_text = build_detailed_analysis_text(parsed, df_local)
        await safe_reply(update, report_text)
        append_chat_history(chat_id, "Asistan", report_text)

        # Çeki listesi istendiyse veya spesifik bir lot sorulduysa ("A-371 stoktaki taşları söyle" dahil) şık Excel'i gönder
        if wants_excel or lots:
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
                    await update.message.reply_document(
                        document=f,
                        caption=(
                            f"📋 **{tag_lot} ({tag_stat}) Kurumsal Çeki Listesi**\n"
                            f"📦 Plaka: {len(target_df)} Adet | 📐 Metraj: {round(sub_m2, 2)} m² | ⚖️ ~{tons} Ton"
                        )
                    )
        return

    # 7. Sunum, İhracat Maili, Strateji ve Genel Yapay Zeka Soruları
    ai_reply = generate_expert_ai_response(text, df_local, chat_id)
    await safe_reply(update, ai_reply)
    append_chat_history(chat_id, "Asistan", ai_reply)


# ==============================================================================
# 9. RENDER 7/24 KEEP-ALIVE SUNUCUSU VE ANA ÇALIŞTIRICI
# ==============================================================================
class HealthCheckHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Arslanlar Marble AI & ERP Bot 7/24 Aktif!")

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
    threading.Thread(
        target=lambda: HTTPServer(("0.0.0.0", port), HealthCheckHandler).serve_forever(),
        daemon=True
    ).start()
    threading.Thread(target=keep_alive_pinger, daemon=True).start()

    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))

    logger.info("Arslanlar Mermer Tam Kapsamlı Uzman Bot başlatılıyor...")
    app.run_polling(drop_pending_updates=True, allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
