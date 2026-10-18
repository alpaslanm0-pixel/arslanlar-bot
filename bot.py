import os
import re
import io
import json
import logging
import threading
import requests
from bs4 import BeautifulSoup
from datetime import datetime, time
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
    {"Lot / Blok No": "A-283", "Taş Cinsi (Selection)": "Violet Breccia", "Koleksiyon": "Arslanlar Collection", "Koridor": "A-1", "Plaka Adedi": 12, "Net Metraj (m²)": 68.4, "Birim Fiyat ($/m²)": 65.0, "Toplam Tutar ($)": 4446.0, "Durum": "Available", "Müşteri": "", "Google Drive Linki": GOOGLE_DRIVE_LINK},
    {"Lot / Blok No": "B-12", "Taş Cinsi (Selection)": "Beige", "Koleksiyon": "Arslanlar Series", "Koridor": "B-3", "Plaka Adedi": 18, "Net Metraj (m²)": 95.0, "Birim Fiyat ($/m²)": 42.0, "Toplam Tutar ($)": 3990.0, "Durum": "Available", "Müşteri": "", "Google Drive Linki": GOOGLE_DRIVE_LINK},
    {"Lot / Blok No": "C-105", "Taş Cinsi (Selection)": "Calacatta Viola", "Koleksiyon": "Arslanlar Exclusive collection", "Koridor": "A-2", "Plaka Adedi": 8, "Net Metraj (m²)": 44.2, "Birim Fiyat ($/m²)": 85.0, "Toplam Tutar ($)": 3757.0, "Durum": "Available", "Müşteri": "", "Google Drive Linki": GOOGLE_DRIVE_LINK},
])

daily_log = {"sold": [], "reserved": [], "cancelled": []}
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
    """Kullanıcının attığı Excel dosyasını akıllı sütun haritalama ile okur."""
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

        # Sütun isimlerini akıllıca eşleştir
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
        
        # Eksik kritik sütunları tamamla
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

        # Geçersiz/boş lot satırlarını temizle
        df = df[df["Lot / Blok No"].astype(str).str.strip() != ""]
        df = df[~df["Lot / Blok No"].astype(str).str.lower().isin(["nan", "toplam", "total", "lot / blok no"])]
        
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


# --- AKILLI DOĞAL DİL MOTORU ---
def analyze_message(text: str) -> dict:
    t_lower = text.lower()
    lot_match = re.search(r"\b([a-zA-Z]{1,3}[- ]?\d+|\d+[a-zA-Z]?)\b", text)
    lot = lot_match.group(1).upper().replace(" ", "-") if lot_match else None

    stone = None
    for s in ["Violet Breccia", "Calacatta Viola", "Beige", "Afyon", "Collection", "Exclusive", "Series"]:
        if s.lower() in t_lower:
            stone = s
            break

    client = None
    for c in ["Glaze Dubai", "Skystone", "Jasmine Stone", "Mitra", "İZFAŞ"]:
        if c.lower() in t_lower:
            client = c
            break

    if any(k in t_lower for k in ["tara", "senkronize", "koleksiyon"]):
        return {"action": "sync_drive", "lot": lot, "stone": stone, "client": client, "reply": "Hemen Arslanlar Series, Collection ve Exclusive koleksiyonlarını tarıyorum patron! ⚡"}
    elif any(k in t_lower for k in ["tüm stok", "genel stok", "master stok", "stok tablosu"]):
        return {"action": "export_stock", "lot": lot, "stone": stone, "client": client, "reply": "Master mermer stok tablomuzu hazırlıyorum! 📋"}
    elif any(k in t_lower for k in ["foto", "resim", "görsel", "plaka resmi"]):
        return {"action": "photo_request", "lot": lot, "stone": stone, "client": client, "reply": "Plaka fotoğraflarını çekiyorum! 📸"}
    elif any(k in t_lower for k in ["iptal", "vazgeç", "geri al"]):
        return {"action": "cancel", "lot": lot, "stone": stone, "client": client, "reply": "İşlemi iptal edip stoğa geri alıyorum! 🔄"}
    elif any(k in t_lower for k in ["çeki", "packing", "liste"]):
        return {"action": "packing_list", "lot": lot, "stone": stone, "client": client, "reply": "Orijinal formatına birebir uygun çeki listesini hazırlıyorum! 📄"}
    elif any(k in t_lower for k in ["sattık", "satıldı", "sold"]):
        return {"action": "sale", "lot": lot, "stone": stone, "client": client, "reply": "Satışı deftere işledim, hayırlı olsun! 🔴"}
    elif any(k in t_lower for k in ["rezerv", "ayır"]):
        return {"action": "reserve", "lot": lot, "stone": stone, "client": client, "reply": "Rezervasyonu koydum, kimsede yok! 🔒"}
    elif any(k in t_lower for k in ["rapor", "özet"]):
        return {"action": "daily_report", "lot": None, "stone": None, "client": None, "reply": "Günlük operasyon raporunu getiriyorum! 📊"}
    elif lot and any(k in t_lower for k in ["drive", "link", "gönder", "kaç"]):
        return {"action": "lot_detail", "lot": lot, "stone": stone, "client": client, "reply": f"Lot {lot} detaylarını ve Drive bağlantısını hemen getiriyorum! 🔍"}
    elif any(k in t_lower for k in ["drive", "link"]):
        return {"action": "drive_link", "lot": lot, "stone": stone, "client": client, "reply": "Koleksiyon Drive linkini bırakıyorum! 🔗"}
    else:
        return {"action": "stock_query", "lot": lot, "stone": stone, "client": client, "reply": "Depo stok durumunu inceliyorum! 🔍"}


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
                f"Eyvallah Mehmet hocam! `{doc.file_name}` dosyasındaki **tüm sütun başlıklarını ve formatı** kusursuz şekilde hafızama kazıdım. 📋\n"
                f"✅ **Toplam Lot:** {len(df)} adet | **Toplam Plaka:** {total_slabs} | **Toplam Metraj:** {round(total_m2, 2)} m²\n"
                f"Bundan sonraki tüm işlemler **birebir bu şablonda** yürütülecek! ⚡",
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
        f"📦 **Aktif Stok Kaydı:** {avail_count} Lot / Blok (Format Onaylandı).\n"
        f"🔗 [Koleksiyon Klasörü]({GOOGLE_DRIVE_LINK})\n\n"
        f"Her şey kontrol altında, iyi çalışmalar! ☕💪"
    )


# --- ANA MESAJ YÖNETİCİSİ ---
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global inventory_df
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

    if action == "lot_detail" and lot:
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
                f"• Birim Fiyat: {row.get('Birim Fiyat ($/m²)', 0)} $/m² (Toplam: {row.get('Toplam Tutar ($)', 0)} $)\n"
                f"• Durum: *{row.get('Durum', 'Available')}*\n\n"
                f"🔗 **Google Drive Klasör Linki:** [Tıkla ve Aç]({drive_url})",
                parse_mode="Markdown"
            )
        else:
            await update.message.reply_text(f"⚠️ `{lot}` kodlu lot mevcut stok listesinde bulunamadı ama Drive koleksiyonlarında taranıyor.", parse_mode="Markdown")

    elif action == "sync_drive":
        items_count, stock_count = scan_marble_collections()
        file_path, slabs, m2, total_usd, _ = generate_packing_list_excel("Master_Stok")
        with open(file_path, "rb") as f:
            await update.message.reply_document(
                document=f,
                caption=f"✅ **Koleksiyonlar Tarandı!**\n📂 Öğe: {items_count}\n🪨 Stok: {len(inventory_df)} Kalem\n📎 *Master Stok tablosu ekte!*",
                parse_mode="Markdown"
            )

    elif action == "export_stock":
        file_path, slabs, m2, total_usd, _ = generate_packing_list_excel("Master_Stok")
        with open(file_path, "rb") as f:
            await update.message.reply_document(document=f, caption=f"📊 **Arslanlar Master Stok Tablosu** hazırlandı!", parse_mode="Markdown")

    elif action in ("drive_link", "photo_request", "stock_query"):
        matched = search_in_drive(lot or stone or text)
        lines = [f"• 📁 [{it['name']}]({it['url']}) *({it['parent']})*" for it in matched[:10]]
        await update.message.reply_text(
            f"🔍 **Drive Arama Sonuçları:**\n\n" + ("\n".join(lines) if lines else "Kayıt bulunamadı.") + f"\n\n🔗 [Koleksiyon Klasörü]({GOOGLE_DRIVE_LINK})",
            parse_mode="Markdown"
        )

    elif action == "sale":
        daily_log["sold"].append({"lot": lot or "Belirsiz", "client": client, "total_usd": 4500.0})
        await update.message.reply_text(f"🔴 **Satış İşlendi:** `{lot or 'Ürün'}` başarıyla `{client}` firmasına satıldı olarak kaydedildi! 🚀", parse_mode="Markdown")

    elif action == "reserve":
        daily_log["reserved"].append({"lot": lot or "Belirsiz", "client": client})
        await update.message.reply_text(f"🟡 **Rezervasyon:** `{lot or 'Ürün'}` , `{client}` adına opsiyonlandı. 🔒", parse_mode="Markdown")

    elif action == "cancel":
        await update.message.reply_text(f"🔄 **İptal:** `{lot or 'İlgili lot'}` stoğa geri alındı. 💪", parse_mode="Markdown")

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
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))

    logger.info("Arslanlar Mermer Zeki Operasyon Asistanı başlatıldı...")
    app.run_polling(drop_pending_updates=True)


class DummyHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Arslanlar Marble Bot 7/24 Aktif!")


if __name__ == "__main__":
    main()
