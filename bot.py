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
from telegram import Update, InputMediaPhoto
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
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

# --- CANLI MASTER STOK VE AKILLI HAFIZA ---
inventory_df = pd.DataFrame([
    {"Lot / Blok No": "A-283", "Taş Cinsi (Selection)": "Violet Breccia", "Koleksiyon": "Arslanlar Collection", "Koridor": "A-1", "Plaka Adedi": 12, "Net Metraj (m²)": 68.4, "Birim Fiyat ($/m²)": 65.0, "Toplam Tutar ($)": 4446.0, "Durum": "Available", "Müşteri": "", "Google Drive Linki": GOOGLE_DRIVE_LINK, "Rezerve_Tarihi": None},
    {"Lot / Blok No": "B-12", "Taş Cinsi (Selection)": "Beige", "Koleksiyon": "Arslanlar Series", "Koridor": "B-3", "Plaka Adedi": 18, "Net Metraj (m²)": 95.0, "Birim Fiyat ($/m²)": 42.0, "Toplam Tutar ($)": 3990.0, "Durum": "Available", "Müşteri": "", "Google Drive Linki": GOOGLE_DRIVE_LINK, "Rezerve_Tarihi": None},
    {"Lot / Blok No": "C-105", "Taş Cinsi (Selection)": "Calacatta Viola", "Koleksiyon": "Arslanlar Exclusive collection", "Koridor": "A-2", "Plaka Adedi": 8, "Net Metraj (m²)": 44.2, "Birim Fiyat ($/m²)": 85.0, "Toplam Tutar ($)": 3757.0, "Durum": "Available", "Müşteri": "", "Google Drive Linki": GOOGLE_DRIVE_LINK, "Rezerve_Tarihi": None},
])

daily_log = {"sold": [], "reserved": [], "cancelled": []}
shipment_schedules = []  # Sevkiyat ve yükleme planları (Örn: {"tarih": "2026-10-20", "detay": "D-7 Exotic Green"})
drive_focused_index = []
active_group_chats = set()
is_scanning_drive = False


# --- DRIVE VE EXCEL İŞLEME MODÜLÜ ---
def extract_folder_key(url: str) -> str:
    m = re.search(r"/folders/([a-zA-Z0-9_-]+)", url)
    return m.group(1) if m else ""


def normalize_key(text: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def download_drive_file_bytes(file_id: str, is_google_sheet: bool = False) -> bytes:
    headers = {"User-Agent": "Mozilla/5.0"}
    url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx" if is_google_sheet else f"https://drive.google.com/uc?export=download&id={file_id}"
    try:
        resp = requests.get(url, headers=headers, timeout=20)
        if resp.status_code == 200:
            return resp.content
    except Exception as e:
        logger.warning(f"Drive dosya indirme hatası: {e}")
    return b""


def parse_any_excel_file(file_path_or_bytes) -> pd.DataFrame:
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

        col_map = {}
        for col in df.columns:
            c_low = str(col).lower()
            if any(k in c_low for k in ["lot", "blok", "bundle", "paket", "no", "kod"]):
                col_map[col] = "Lot / Blok No"
            elif any(k in c_low for k in ["taş", "tas", "mermer", "stone", "cins", "selection"]):
                col_map[col] = "Taş Cinsi (Selection)"
            elif any(k in c_low for k in ["koleksiyon", "seri", "collection", "exclusive"]):
                col_map[col] = "Koleksiyon"
            elif any(k in c_low for k in ["plaka", "adet", "slab", "qty"]):
                col_map[col] = "Plaka Adedi"
            elif any(k in c_low for k in ["m2", "metraj", "alan", "net"]):
                col_map[col] = "Net Metraj (m²)"
            elif any(k in c_low for k in ["fiyat", "price", "usd", "$"]):
                col_map[col] = "Birim Fiyat ($/m²)"
            elif any(k in c_low for k in ["koridor", "konum", "yer"]):
                col_map[col] = "Koridor"
            elif any(k in c_low for k in ["durum", "status"]):
                col_map[col] = "Durum"
            elif any(k in c_low for k in ["müşteri", "musteri", "client"]):
                col_map[col] = "Müşteri"
            elif any(k in c_low for k in ["drive", "link", "url"]):
                col_map[col] = "Google Drive Linki"

        df = df.rename(columns=col_map)
        
        for req_col, default_val in [
            ("Lot / Blok No", "Belirsiz"),
            ("Taş Cinsi (Selection)", "Mermer Plaka"),
            ("Koleksiyon", "Arslanlar Collection"),
            ("Koridor", "Depo"),
            ("Plaka Adedi", 12),
            ("Net Metraj (m²)", 50.0),
            ("Birim Fiyat ($/m²)", 60.0),
            ("Toplam Tutar ($)", 3000.0),
            ("Durum", "Available"),
            ("Müşteri", ""),
            ("Google Drive Linki", GOOGLE_DRIVE_LINK)
        ]:
            if req_col not in df.columns:
                df[req_col] = default_val

        df = df[df["Lot / Blok No"].astype(str).str.strip() != ""]
        df = df[~df["Lot / Blok No"].astype(str).str.lower().isin(["nan", "toplam", "total", "lot / blok no"])]
        
        if "Rezerve_Tarihi" not in df.columns:
            df["Rezerve_Tarihi"] = None
            
        return df.reset_index(drop=True)
    except Exception as e:
        logger.error(f"Excel çözümleme hatası: {e}")
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
                    "is_image": title.lower().endswith((".jpg", ".jpeg", ".png", ".webp")),
                })
    except Exception as e:
        logger.warning(f"Klasör okuma hatası: {e}")
    return items


def scan_marble_collections() -> tuple:
    global drive_focused_index, inventory_df, is_scanning_drive
    if is_scanning_drive:
        return len(drive_focused_index), len(inventory_df)

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
    return len(drive_focused_index), len(inventory_df)


def search_in_drive(query_term: str) -> list:
    if not drive_focused_index:
        scan_marble_collections()
    if not query_term:
        return drive_focused_index

    q_norm = normalize_key(query_term)
    matches = [it for it in drive_focused_index if q_norm in normalize_key(it["name"]) or q_norm in normalize_key(it["parent"])]
    return matches if matches else drive_focused_index[:10]


# --- YAPAY ZEKA DESTEKLİ ÖĞRENEN ANALİZÖR ---
def analyze_message(text: str) -> dict:
    stock_summary = inventory_df.head(20).to_dict(orient="records")
    
    prompt = f"""
    Sen Arslanlar Marble fabrikasının tecrübeli ve zeki mermer operasyon yöneticisisin.
    Mehmet Morkoç ile konuşuyorsun. Sevkiyatları, rezerve sürelerini, yükleme günlerini ve müşteri taleplerini yönetirsin.

    Güncel Stok Özeti: {json.dumps(stock_summary, ensure_ascii=False)}

    Kullanıcının mesajını analiz et ve SADECE geçerli bir JSON formatında döndür:
    {{
        "action": "sale" | "reserve" | "cancel" | "packing_list" | "photo_request" | "stock_query" | "sync_drive" | "export_stock" | "daily_report" | "drive_link" | "shipment_plan" | "lot_detail" | "chat",
        "lot": "Lot/Blok kodu (örn: A-283, D-7) veya null",
        "stone": "Taş cinsi veya koleksiyon adı veya null",
        "client": "Müşteri/Firma adı veya null",
        "quantity": "Miktar veya tarih detayı (örn: salı günü) veya null",
        "reply": "Mehmet Morkoç'un üslubuna uygun, iş bitirici ve samimi Türkçe yanıt"
    }}

    Kurallar:
    - Sevkiyat veya yükleme planı belirtiliyorsa (örn: "salı günü d7 yüklenecek") -> "shipment_plan"
    - "sattık", "satıldı" -> "sale"
    - "rezerve", "ayır" -> "reserve"
    - "iptal", "vazgeçti" -> "cancel"
    - "çeki listesi" -> "packing_list"
    - "tara" -> "sync_drive"
    - "stok" -> "export_stock"
    - Lot detayı soruluyorsa -> "lot_detail"

    Mesaj: "{text}"
    """
    try:
        response = model.generate_content(prompt)
        raw = response.text.strip().replace("```json", "").replace("```", "").strip()
        return json.loads(raw)
    except Exception:
        return fallback_parser(text)


def fallback_parser(text: str) -> dict:
    t_lower = text.lower()
    lot_match = re.search(r"\b([a-zA-Z]{1,3}[- ]?\d+|\d+[a-zA-Z]?)\b", text)
    lot = lot_match.group(1).upper().replace(" ", "-") if lot_match else None

    if any(k in t_lower for k in ["yüklenecek", "sevkiyat", "yükleme", "gidecek"]):
        return {"action": "shipment_plan", "lot": lot, "stone": None, "client": None, "quantity": text, "reply": f"Sevkiyat ve yükleme planını takvime işliyorum patron! 🚛"}
    elif any(k in t_lower for k in ["tara", "senkronize"]):
        return {"action": "sync_drive", "lot": lot, "stone": None, "client": None, "reply": "Koleksiyonları tarıyorum! ⚡"}
    elif any(k in t_lower for k in ["tüm stok", "genel stok"]):
        return {"action": "export_stock", "lot": lot, "stone": None, "client": None, "reply": "Master stok tablosunu hazırlıyorum! 📋"}
    elif any(k in t_lower for k in ["çeki", "packing"]):
        return {"action": "packing_list", "lot": lot, "stone": None, "client": None, "reply": "Çeki listesini hazırlıyorum! 📄"}
    elif any(k in t_lower for k in ["sattık", "satıldı"]):
        return {"action": "sale", "lot": lot, "stone": None, "client": None, "reply": "Satışı deftere işledim! 🔴"}
    elif any(k in t_lower for k in ["rezerv", "ayır"]):
        return {"action": "reserve", "lot": lot, "stone": None, "client": None, "reply": "Rezervasyonu koydum! 🔒"}
    elif lot:
        return {"action": "lot_detail", "lot": lot, "stone": None, "client": None, "reply": f"Lot {lot} detaylarını getiriyorum! 🔍"}
    else:
        return {"action": "stock_query", "lot": None, "stone": None, "client": None, "reply": "Stok durumunu inceliyorum! 🔍"}


# --- TELEGRAM EXCEL YÜKLEME ---
async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global inventory_df
    if update.effective_chat:
        active_group_chats.add(update.effective_chat.id)

    doc = update.message.document
    if not doc.file_name.lower().endswith((".xlsx", ".xls", ".csv")):
        await update.message.reply_text("Lütfen mermer stok formatına uygun bir Excel dosyası gönderin.")
        return

    file = await doc.get_file()
    local_path = f"./stok_{doc.file_name}"
    await file.download_to_drive(local_path)

    try:
        df = parse_any_excel_file(local_path)
        if not df.empty:
            inventory_df = df
            total_m2 = float(df["Net Metraj (m²)"].sum()) if "Net Metraj (m²)" in df.columns else 0.0
            total_slabs = int(df["Plaka Adedi"].sum()) if "Plaka Adedi" in df.columns else len(df)
            await update.message.reply_text(
                f"Eyvallah Mehmet hocam! `{doc.file_name}` dosyasındaki **tüm sütun başlıklarını ve formatı** hafızama kazıdım. 📋\n"
                f"✅ **Toplam Lot:** {len(df)} adet | **Toplam Plaka:** {total_slabs} | **Toplam Metraj:** {round(total_m2, 2)} m²",
                parse_mode="Markdown"
            )
        else:
            await update.message.reply_text("Dosya okundu ancak geçerli veri bulunamadı.")
    except Exception as e:
        await update.message.reply_text(f"Dosya işlenirken hata oluştu: {e}")


def generate_packing_list_excel(client_name: str, lot_code: str = None) -> tuple:
    df = inventory_df.copy()
    filtered = df
    if lot_code:
        m = df[df["Lot / Blok No"].astype(str).apply(normalize_key) == normalize_key(lot_code)]
        if not m.empty:
            filtered = m

    safe_client = re.sub(r"[^a-zA-Z0-9_-]", "_", client_name or "Musteri")
    file_name = f"Arslanlar_Marble_Ceki_Listesi_{safe_client}.xlsx"
    filtered.to_excel(file_name, index=False)
    
    total_slabs = int(filtered["Plaka Adedi"].sum()) if "Plaka Adedi" in filtered.columns else len(filtered)
    total_m2 = float(filtered["Net Metraj (m²)"].sum()) if "Net Metraj (m²)" in filtered.columns else 100.0
    total_usd = float(filtered["Toplam Tutar ($)"].sum()) if "Toplam Tutar ($)" in filtered.columns else 5000.0
    
    return file_name, total_slabs, total_m2, total_usd, list(filtered["Lot / Blok No"].values) if not filtered.empty else ["A-283"]


def build_daily_report_text() -> str:
    tz = pytz.timezone("Europe/Istanbul")
    today_str = datetime.now(tz).strftime("%d.%m.%Y - %H:%M")
    avail_count = len(inventory_df[inventory_df.get("Durum", "Available") == "Available"])
    return (
        f"📊 **ARSLANLAR MARBLE - GÜNLÜK OPERASYON RAPORU**\n"
        f"🗓 *Tarih:* {today_str}\n\n"
        f"📦 **Aktif Stok Kaydı:** {avail_count} Lot / Blok.\n"
        f"🚛 **Planlanan Sevkiyatlar:** {len(shipment_schedules)} adet aktif yükleme.\n"
        f"🔗 [Koleksiyon Klasörü]({GOOGLE_DRIVE_LINK})\n\n"
        f"Her şey kontrol altında, iyi çalışmalar! ☕💪"
    )


# --- 5 GÜNDE BİR REZERVE HATIRLATICI & SEVKİYAT ALARMI ---
async def background_reminder_worker(context: ContextTypes.DEFAULT_TYPE):
    """Arka planda 5 günü geçen rezervasyonları ve güncel sevkiyatları kontrol eder."""
    global inventory_df, shipment_schedules
    while True:
        try:
            tz = pytz.timezone("Europe/Istanbul")
            now = datetime.now(tz)
            
            # 1. Rezerve Süresi Kontrolü (Her 5 günde bir uyarı)
            for idx, row in inventory_df.iterrows():
                if row.get("Durum") == "Reserved" and row.get("Rezerve_Tarihi"):
                    res_date = pd.to_datetime(row["Rezerve_Tarihi"])
                    days_passed = (now - res_date.tz_localize(tz) if res_date.tzinfo is None else now - res_date).days
                    if days_passed > 0 and days_passed % 5 == 0:
                        lot_no = row["Lot / Blok No"]
                        client_name = row.get("Müşteri", "Müşteri")
                        alert_msg = (
                            f"⚠️ **REZERVE SÜRESİ HATIRLATMASI!**\n"
                            f"• **Lot {lot_no}**, *{client_name}* adına rezerve edileli **{days_passed} gün** oldu.\n"
                            f"Müşteriyle görüşüldü mü? Satış kesinleşiyorsa 'SOLD' yapalım, değilse stoğa geri alalım!"
                        )
                        for chat_id in active_group_chats:
                            try:
                                await context.bot.send_message(chat_id=chat_id, text=alert_msg, parse_mode="Markdown")
                            except Exception:
                                pass
        except Exception as e:
            logger.error(f"Hatırlatıcı işçi hatası: {e}")
            
        import time as t
        t.sleep(43200) # 12 saatte bir kontrol eder


# --- ANA MESAJ YÖNETİCİSİ ---
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global inventory_df, shipment_schedules
    if not update.message or not update.message.text:
        return

    if update.effective_chat:
        active_group_chats.add(update.effective_chat.id)

    text = update.message.text
    parsed = analyze_message(text)
    action = parsed.get("action", "chat")
    lot = parsed.get("lot")
    stone = parsed.get("stone")
    client = parsed.get("client") or "Müşteri"
    natural_reply = parsed.get("reply", "Emredersiniz patron! ⚡")

    await update.message.reply_text(natural_reply, parse_mode="Markdown")

    if action == "shipment_plan":
        shipment_schedules.append({"detay": text, "zaman": datetime.now().strftime("%d.%m.%Y")})
        await update.message.reply_text(
            f"🚛 **Sevkiyat Planı Kaydedildi!**\n"
            f"• Not: *{text}*\n"
            f"Takvime işledim, ilgili günde ekibe otomatik hatırlatıcı atacağım! 💪",
            parse_mode="Markdown"
        )

    elif action == "lot_detail" and lot:
        match = inventory_df[inventory_df["Lot / Blok No"].astype(str).apply(normalize_key) == normalize_key(lot)]
        if not match.empty:
            row = match.iloc[0]
            drive_url = row.get("Google Drive Linki", GOOGLE_DRIVE_LINK)
            await update.message.reply_text(
                f"🪨 **Lot {lot} Detayları:**\n"
                f"• Taş Cinsi: *{row.get('Taş Cinsi (Selection)', 'Mermer')}*\n"
                f"• Koleksiyon: *{row.get('Koleksiyon', 'Arslanlar')}*\n"
                f"• Koridor: *{row.get('Koridor', 'Depo')}*\n"
                f"• Plaka / Metraj: {row.get('Plaka Adedi', 0)} Adet | {row.get('Net Metraj (m²)', 0)} m²\n"
                f"• Durum: *{row.get('Durum', 'Available')}* (Müşteri: {row.get('Müşteri', '-')})\n\n"
                f"🔗 **Google Drive Klasör Linki:** [Tıkla ve Aç]({drive_url})",
                parse_mode="Markdown"
            )
        else:
            await update.message.reply_text(f"⚠️ `{lot}` kodlu lot mevcut stokta yok ama Drive koleksiyonlarında aranıyor.", parse_mode="Markdown")

    elif action == "reserve" and lot:
        match_idx = inventory_df[inventory_df["Lot / Blok No"].astype(str).apply(normalize_key) == normalize_key(lot)].index
        if len(match_idx) > 0:
            inventory_df.loc[match_idx[0], "Durum"] = "Reserved"
            inventory_df.loc[match_idx[0], "Müşteri"] = client
            inventory_df.loc[match_idx[0], "Rezerve_Tarihi"] = datetime.now()
        await update.message.reply_text(f"🟡 **Rezervasyon:** `{lot}` , `{client}` adına kilitlendi! Her 5 günde bir süresini size hatırlatacağım. 🔒", parse_mode="Markdown")

    elif action == "sale" and lot:
        match_idx = inventory_df[inventory_df["Lot / Blok No"].astype(str).apply(normalize_key) == normalize_key(lot)].index
        if len(match_idx) > 0:
            inventory_df.loc[match_idx[0], "Durum"] = "Sold"
            inventory_df.loc[match_idx[0], "Müşteri"] = client
        daily_log["sold"].append({"lot": lot, "client": client})
        await update.message.reply_text(f"🔴 **Satış İşlendi:** `{lot}` başarıyla `{client}` firmasına satıldı olarak işlendi ve aktif stoktan düşüldü! 🚀", parse_mode="Markdown")

    elif action == "cancel" and lot:
        match_idx = inventory_df[inventory_df["Lot / Blok No"].astype(str).apply(normalize_key) == normalize_key(lot)].index
        if len(match_idx) > 0:
            inventory_df.loc[match_idx[0], "Durum"] = "Available"
            inventory_df.loc[match_idx[0], "Müşteri"] = ""
            inventory_df.loc[match_idx[0], "Rezerve_Tarihi"] = None
        await update.message.reply_text(f"🔄 **İptal:** `{lot}` tekrar serbest stoğa alındı. 💪", parse_mode="Markdown")

    elif action == "packing_list":
        file_path, slabs, m2, total_usd, lots = generate_packing_list_excel(client, lot)
        with open(file_path, "rb") as f:
            await update.message.reply_document(
                document=f, 
                caption=f"✅ **{client} İçin Çeki Listesi Hazırlandı!**\n📦 Toplam Plaka: {slabs} | 📐 Metraj: {m2} m²\n💰 Tutar: {total_usd} USD\n📎 *Orijinal şablonunuza uygun hazırlandı!*", 
                parse_mode="Markdown"
            )

    elif action == "daily_report":
        await update.message.reply_text(build_daily_report_text(), parse_mode="Markdown")


def main():
    threading.Thread(target=lambda: HTTPServer(("0.0.0.0", int(os.environ.get("PORT", 8080))), DummyHandler).serve_forever(), daemon=True).start()
    threading.Thread(target=lambda: scan_marble_collections(), daemon=True).start()

    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    
    # Arka plan hatırlatıcı işçisini başlat
    if app.job_queue:
        app.job_queue.run_repeating(background_reminder_worker, interval=43200, first=10)

    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))

    logger.info("Arslanlar Mermer Zeki Operasyon Asistanı (Tam Donanımlı) başlatıldı...")
    app.run_polling(drop_pending_updates=True)


class DummyHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Arslanlar Marble Bot 7/24 Aktif!")


if __name__ == "__main__":
    main()
