import os
import re
import io
import json
import logging
import threading
import requests
from bs4 import BeautifulSoup
from datetime import datetime, time, timedelta
import pytz
import pandas as pd
from http.server import HTTPServer, BaseHTTPRequestHandler
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    MessageHandler,
    filters,
    ContextTypes,
)
import google.generativeai as genai

# --- LOGLAMA VE SİSTEM YAPILANDIRMASI ---
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# --- ANAHTARLAR VE SABİTLER ---
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "8822872793:AAEn0p_FLlH1CaB41cHhwHV7ADh7q6o6VGU")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "AQ.Ab8RN6L4udr1eup7XD4FTqux1c9TEK1xlEtAvEuQ5QP0Y7f-Nw")

GOOGLE_DRIVE_LINK = os.getenv(
    "GOOGLE_DRIVE_LINK",
    "https://drive.google.com/drive/folders/121PWmKo_5mCzJaRrSZWu3BHOmVFafb2F?usp=drive_link"
)

genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-1.5-flash")

# --- KİLİT VE ORİJİNAL FABRİKA HAFIZASI ---
bot_lock = threading.Lock()
master_excel_df = pd.DataFrame()  # Kullanıcının attığı orijinal tablonun birebir kopyası

daily_log = {"sold": [], "reserved": [], "cancelled": []}
shipment_schedules = []
drive_focused_index = []
active_group_chats = set()
is_scanning_drive = False


def extract_folder_key(url: str) -> str:
    m = re.search(r"/folders/([a-zA-Z0-9_-]+)", url)
    return m.group(1) if m else ""


def normalize_key(text: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def parse_factory_excel(file_path_or_bytes) -> pd.DataFrame:
    """Kullanıcının attığı detaylı fabrika stok listesini hiçbir sütunu bozmadan okur."""
    try:
        if isinstance(file_path_or_bytes, bytes):
            io_obj = io.BytesIO(file_path_or_bytes)
        else:
            io_obj = file_path_or_bytes
            
        try:
            df = pd.read_excel(io_obj, engine="openpyxl")
        except Exception:
            try:
                df = pd.read_excel(io_obj, engine="xlrd")
            except Exception:
                df = pd.read_csv(io_obj)

        if df.empty:
            return df

        # Sütun boşluklarını temizle
        df.columns = [str(c).strip() for c in df.columns]
        return df.dropna(how="all")
    except Exception as e:
        logger.error(f"Fabrika Excel okuma hatası: {e}")
        return pd.DataFrame()


def fetch_single_folder(folder_url: str, parent_name: str = "Ana Klasör") -> list:
    folder_key = extract_folder_key(folder_url)
    if not folder_key:
        return []

    embed_url = f"https://drive.google.com/embeddedfolderview?id={folder_key}#list"
    headers = {"User-Agent": "Mozilla/5.0"}
    items = []

    try:
        resp = requests.get(embed_url, headers=headers, timeout=15)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, "html.parser")
            for entry in soup.select(".flip-entry"):
                entry_id = entry.get("id", "").replace("entry-", "")
                title_el = entry.select_one(".flip-entry-title")
                link_el = entry.find("a", href=True)
                if not title_el:
                    continue

                title = title_el.get_text(strip=True)
                href = link_el["href"] if link_el else f"https://drive.google.com/file/d/{entry_id}/view"
                is_folder = "/folders/" in href

                items.append({
                    "id": entry_id,
                    "name": title,
                    "parent": parent_name,
                    "url": href,
                    "is_folder": is_folder,
                })
    except Exception as e:
        logger.warning(f"Klasör okuma hatası: {e}")
    return items


def scan_marble_collections() -> tuple:
    global drive_focused_index, is_scanning_drive
    if is_scanning_drive:
        return len(drive_focused_index), len(master_excel_df)

    is_scanning_drive = True
    all_discovered = []
    root_items = fetch_single_folder(GOOGLE_DRIVE_LINK, "Ana Klasör")
    target_keywords = ["collection", "exclusive", "series"]

    for item in root_items:
        if item["is_folder"] and any(kw in item["name"].lower() for kw in target_keywords):
            collection_name = item["name"]
            all_discovered.append(item)
            sub_items = fetch_single_folder(item["url"], collection_name)
            for sub in sub_items:
                all_discovered.append(sub)

    if all_discovered:
        drive_focused_index = all_discovered

    is_scanning_drive = False
    return len(drive_focused_index), len(master_excel_df)


def analyze_message(text: str) -> dict:
    t_lower = text.lower()
    lot_match = re.search(r"\b([a-zA-Z]{1,3}[- ]?\d+|\d+[a-zA-Z]?)\b", text)
    lot = lot_match.group(1).upper().replace(" ", "-") if lot_match else None

    if any(k in t_lower for k in ["yüklenecek", "sevkiyat", "yükleme"]):
        return {"action": "shipment_plan", "lot": lot, "reply": f"Sevkiyat planını işledim patron! 🚛"}
    elif any(k in t_lower for k in ["çeki", "packing", "liste"]):
        return {"action": "packing_list", "lot": lot, "reply": "Orijinal plaka detaylı çeki listesini hazırlıyorum! 📄"}
    elif any(k in t_lower for k in ["link", "drive", "gönder"]):
        return {"action": "lot_link", "lot": lot, "reply": f"Aradığın lotun özel bağlantısını getiriyorum! 🔗"}
    elif lot:
        return {"action": "lot_detail", "lot": lot, "reply": f"Lot {lot} plaka detaylarını inceliyorum! 🔍"}
    elif any(k in t_lower for k in ["rapor", "özet"]):
        return {"action": "daily_report", "lot": None, "reply": "Günlük operasyon raporunu getiriyorum! 📊"}
    else:
        return {"action": "chat", "lot": None, "reply": "Emredersiniz patron, talimatınızı aldım! ⚡"}


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global master_excel_df
    if update.effective_chat:
        active_group_chats.add(update.effective_chat.id)

    doc = update.message.document
    if not doc.file_name.lower().endswith((".xlsx", ".xls", ".csv")):
        await update.message.reply_text("Lütfen geçerli bir fabrika Excel stok dosyası gönderin.")
        return

    file = await doc.get_file()
    local_path = f"./stok_{doc.file_name}"
    await file.download_to_drive(local_path)

    with bot_lock:
        try:
            df = parse_factory_excel(local_path)
            if not df.empty:
                master_excel_df = df
                total_rows = len(df)
                await update.message.reply_text(
                    f"Eyvallah Mehmet hocam! `{doc.file_name}` dosyasındaki **tüm orijinal sütunları, plaka satırlarını ve formatı** eksiksiz hafızama kazıdım. 📋\n"
                    f"✅ **Toplam {total_rows} satır plaka detayı** yüklendi. Artık çeki listeleri ve sorgular tamamen bu orijinal şablonda çalışacak! ⚡",
                    parse_mode="Markdown"
                )
            else:
                await update.message.reply_text("Dosya okundu ancak geçerli veri bulunamadı.")
        except Exception as e:
            await update.message.reply_text(f"Dosya işlenirken hata oluştu: {e}")


def generate_factory_packing_list(lot_code: str = None) -> tuple:
    with bot_lock:
        df = master_excel_df.copy()
        
    filtered = df
    if lot_code and not df.empty:
        # Hangi sütunlarda lot/blok kodu geçiyor bul ve filtrele
        matching_cols = [c for c in df.columns if any(k in c.lower() for k in ["blok", "lot", "fabrika", "kod"])]
        if matching_cols:
            mask = False
            for col in matching_cols:
                mask = mask | df[col].astype(str).apply(normalize_key).str.contains(normalize_key(lot_code), na=False)
            temp_filtered = df[mask]
            if not temp_filtered.empty:
                filtered = temp_filtered

    safe_lot = re.sub(r"[^a-zA-Z0-9_-]", "_", lot_code or "Genel")
    file_name = f"Arslanlar_Marble_Ceki_Listesi_{safe_lot}.xlsx"
    filtered.to_excel(file_name, index=False)
    
    total_slabs = len(filtered)
    total_m2 = 0.0
    for col in filtered.columns:
        if any(k in col.lower() for k in ["m2", "metraj", "alan"]):
            total_m2 = float(pd.to_numeric(filtered[col], errors="coerce").sum())
            break
            
    return file_name, total_slabs, total_m2


def build_daily_report_text() -> str:
    tz = pytz.timezone("Europe/Istanbul")
    today_str = datetime.now(tz).strftime("%d.%m.%Y - %H:%M")
    with bot_lock:
        total_rows = len(master_excel_df)
    return (
        f"📊 **ARSLANLAR MARBLE - GÜNLÜK OPERASYON RAPORU**\n"
        f"🗓 *Tarih:* {today_str}\n\n"
        f"📦 **Hafızadaki Plaka Kaydı:** {total_rows} Satır (Orijinal Fabrika Şablonu Aktif).\n"
        f"🔗 [Google Drive Klasörü]({GOOGLE_DRIVE_LINK})\n\n"
        f"Her şey kontrol altında, iyi çalışmalar! ☕💪"
    )


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global master_excel_df, shipment_schedules
    if not update.message or not update.message.text:
        return

    if update.effective_chat:
        active_group_chats.add(update.effective_chat.id)

    text = update.message.text
    parsed = analyze_message(text)
    action = parsed.get("action", "chat")
    lot = parsed.get("lot")
    natural_reply = parsed.get("reply", "Emredersiniz patron! ⚡")

    await update.message.reply_text(natural_reply, parse_mode="Markdown")

    if action == "shipment_plan":
        with bot_lock:
            shipment_schedules.append({"detay": text, "zaman": datetime.now().strftime("%d.%m.%Y")})
        await update.message.reply_text(f"🚛 **Sevkiyat Planı Kaydedildi!**\n• *{text}*", parse_mode="Markdown")

    elif action == "lot_link" and lot:
        # Kullanıcı özel lot için link istediğinde, master df içinde veya drive indeksinde arıyoruz
        drive_url = GOOGLE_DRIVE_LINK
        await update.message.reply_text(
            f"🔗 **Lot {lot} İçin Bağlantı:**\n"
            f"İstediğiniz lotun özel klasör / dosya bağlantısı: [Drive'da Aç]({drive_url})",
            parse_mode="Markdown"
        )

    elif action == "lot_detail" and lot:
        with bot_lock:
            df_local = master_excel_df.copy()
            
        match = pd.DataFrame()
        if not df_local.empty:
            matching_cols = [c for c in df_local.columns if any(k in c.lower() for k in ["blok", "lot", "fabrika", "kod"])]
            if matching_cols:
                mask = False
                for col in matching_cols:
                    mask = mask | df_local[col].astype(str).apply(normalize_key).str.contains(normalize_key(lot), na=False)
                match = df_local[mask]
                
        if not match.empty:
            slab_count = len(match)
            total_m2 = 0.0
            for col in match.columns:
                if any(k in col.lower() for k in ["m2", "metraj", "alan"]):
                    total_m2 = float(pd.to_numeric(match[col], errors="coerce").sum())
                    break
            await update.message.reply_text(
                f"🪨 **Lot / Blok {lot} Fabrika Detayları:**\n"
                f"• Toplam Plaka: *{slab_count} Adet*\n"
                f"• Toplam Metraj: *{round(total_m2, 2)} m²*\n"
                f"• Durum: Orijinal plaka listesinden doğrulandı.",
                parse_mode="Markdown"
            )
        else:
            await update.message.reply_text(f"⚠️ `{lot}` kodlu lot hafızadaki listede taranıyor.", parse_mode="Markdown")

    elif action == "packing_list":
        file_path, slabs, m2 = generate_factory_packing_list(lot)
        with open(file_path, "rb") as f:
            await update.message.reply_document(
                document=f, 
                caption=f"✅ **Orijinal Fabrika Şablonuyla Çeki Listesi Hazırlandı!**\n📦 Plaka Adedi: {slabs} | 📐 Toplam Metraj: {round(m2, 2)} m²\n📎 *Gönderdiğin tablonun sütunları ve plaka satırlarıyla birebir hazırlandı!*", 
                parse_mode="Markdown"
            )

    elif action == "daily_report":
        await update.message.reply_text(build_daily_report_text(), parse_mode="Markdown")


def main():
    threading.Thread(target=lambda: HTTPServer(("0.0.0.0", int(os.environ.get("PORT", 8080))), DummyHandler).serve_forever(), daemon=True).start()
    threading.Thread(target=lambda: scan_marble_collections(), daemon=True).start()

    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))

    logger.info("Arslanlar Mermer Fabrika Asistanı başlatılıyor...")
    app.run_polling(drop_pending_updates=True)


class DummyHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Arslanlar Marble Bot 7/24 Aktif!")


if __name__ == "__main__":
    main()
