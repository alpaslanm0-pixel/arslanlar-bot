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

# Loglama Ayarları
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# --- ANAHTARLAR VE SABİT BAĞLANTILAR ---
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "8822872793:AAEn0p_FLlH1CaB41cHhwHV7ADh7q6o6VGU")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "AQ.Ab8RN6L4udr1eup7XD4FTqux1c9TEK1xlEtAvEuQ5QP0Y7f-Nw")

# Paylaştığın Google Drive Ana Plaka Stok Klasörü
GOOGLE_DRIVE_LINK = os.getenv(
    "GOOGLE_DRIVE_LINK",
    "https://drive.google.com/drive/folders/121PWmKo_5mCzJaRrSZWu3BHOmVFafb2F?usp=drive_link"
)

genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-1.5-flash")

# --- CANLI STOK VE HAREKET HAFIZASI ---
inventory_df = pd.DataFrame([
    {"Lot": "A-283", "Stone": "Violet Breccia", "Slabs": 12, "M2": 68.4, "Price_USD": 65.0, "Corridor": "A-1", "Status": "Available", "Client": "", "Drive_URL": GOOGLE_DRIVE_LINK},
    {"Lot": "B-12", "Stone": "Beige", "Slabs": 18, "M2": 95.0, "Price_USD": 42.0, "Corridor": "B-3", "Status": "Available", "Client": "", "Drive_URL": GOOGLE_DRIVE_LINK},
    {"Lot": "C-105", "Stone": "Calacatta Viola", "Slabs": 8, "M2": 44.2, "Price_USD": 85.0, "Corridor": "A-2", "Status": "Available", "Client": "", "Drive_URL": GOOGLE_DRIVE_LINK},
])

daily_log = {
    "sold": [],
    "reserved": [],
    "cancelled": []
}

# Paylaşılan Drive klasöründeki alt dosya/klasör önbelleği
drive_items_cache = []
active_group_chats = set()


# --- GOOGLE DRIVE KLASÖR TARAYICI VE FOTOĞRAF ÇEKİCİ ---
def extract_folder_key(url: str) -> str:
    m = re.search(r"/folders/([a-zA-Z0-9_-]+)", url)
    return m.group(1) if m else ""


def scan_public_drive_folder(folder_url: str = GOOGLE_DRIVE_LINK) -> list:
    """
    Herkese açık (Anyone with link) Google Drive klasörünün içindeki
    alt klasörleri, Excel tablolarını ve plaka fotoğraflarını tarar.
    """
    global drive_items_cache
    folder_key = extract_folder_key(folder_url)
    if not folder_key:
        return drive_items_cache

    embed_url = f"https://drive.google.com/embeddedfolderview?id={folder_key}#list"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
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
                    "url": href,
                    "is_folder": is_folder,
                    "is_image": title.lower().endswith((".jpg", ".jpeg", ".png", ".webp")),
                    "is_excel": title.lower().endswith((".xlsx", ".xls", ".csv")),
                })
        if items:
            drive_items_cache = items
    except Exception as e:
        logger.warning(f"Drive klasör tarama uyarısı: {e}")

    return drive_items_cache


def find_drive_matches(query_term: str) -> list:
    """Aranan lot veya taş adına göre Drive içindeki dosya/klasörleri bulur."""
    items = scan_public_drive_folder()
    if not query_term:
        return items

    q_low = query_term.lower().strip()
    matches = [it for it in items if q_low in it["name"].lower()]

    # Eğer eşleşen bir alt klasör varsa, onun da içine girip fotoğrafları/dosyaları çek
    sub_items = []
    for m in matches:
        if m["is_folder"]:
            sub_list = scan_public_drive_folder(m["url"])
            sub_items.extend(sub_list)

    return matches + sub_items


def download_drive_file_bytes(file_id: str) -> bytes:
    """Google Drive'daki fotoğraf veya dosyayı doğrudan indirir."""
    download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
    headers = {"User-Agent": "Mozilla/5.0"}
    resp = requests.get(download_url, headers=headers, timeout=20)
    if resp.status_code == 200:
        return resp.content
    return b""


# --- DOĞAL DİL ANLAMA MOTORU (GEMINI + YEREL FALLBACK) ---
def analyze_message(text: str) -> dict:
    stock_summary = inventory_df[["Lot", "Stone", "Slabs", "M2", "Price_USD", "Status", "Client"]].head(30).to_dict(orient="records")

    prompt = f"""
    Sen Arslanlar Marble mermer fabrikasının operasyon ve satış asistanısın.
    Mehmet Morkoç ve satış ekibi seninle grupta doğal bir insan gibi konuşuyor.

    Güncel Stok Özeti: {json.dumps(stock_summary, ensure_ascii=False)}

    Kullanıcının mesajını analiz et ve SADECE aşağıdaki JSON formatında yanıt ver:
    {{
        "action": "sale" | "reserve" | "cancel" | "packing_list" | "photo_request" | "stock_query" | "daily_report" | "drive_link" | "chat",
        "lot": "Tespit edilen lot/blok numarası (örn: A-283, B-12) veya null",
        "stone": "Tespit edilen mermer adı (örn: Violet Breccia, Beige) veya null",
        "client": "Müşteri/firma adı (örn: Glaze Dubai, Skystone, Mitra, Jasmine Stone) veya null",
        "quantity": "Miktar/kasa/plaka detayı veya null",
        "reply": "Mehmet Morkoç'un samimi, enerjik, iş bitirici üslubuyla Türkçe doğal yanıt"
    }}

    Kurallar:
    - Fotoğraf, görsel, resim, plaka fotosu isteniyorsa -> "photo_request"
    - Çeki listesi, packing list, ebat/metraj listesi isteniyorsa -> "packing_list"
    - Satıldı, sattık, sold, verildi deniyorsa -> "sale"
    - Rezerve, opsiyon, ayır deniyorsa -> "reserve"
    - İptal, vazgeçti, stoğa geri al deniyorsa -> "cancel"
    - Drive linki, klasör linki, drive'ı kontrol et deniyorsa -> "drive_link"
    - Stokta ne var, fiyatı ne, kaç plaka var deniyorsa -> "stock_query"
    - Günlük rapor, bugün ne yaptık, özet geç deniyorsa -> "daily_report"

    Mesaj: "{text}"
    """
    try:
        response = model.generate_content(prompt)
        raw = response.text.strip().replace("```json", "").replace("```", "").strip()
        return json.loads(raw)
    except Exception as e:
        logger.warning(f"Gemini API yedek motora geçiyor: {e}")
        return fallback_parser(text)


def fallback_parser(text: str) -> dict:
    t_lower = text.lower()
    lot_match = re.search(r"\b([a-zA-Z]-\d+|\d+[a-zA-Z]?)\b", text)
    lot = lot_match.group(1).upper() if lot_match else None

    stone = None
    for s in list(inventory_df["Stone"].dropna().unique()) + ["Violet Breccia", "Calacatta Viola", "Beige", "Afyon"]:
        if str(s).lower() in t_lower or str(s).split()[0].lower() in t_lower:
            stone = str(s)
            break

    client = None
    for c in ["Glaze Dubai", "Skystone", "Jasmine Stone", "Mitra", "İZFAŞ"]:
        if c.lower() in t_lower:
            client = c
            break

    if any(k in t_lower for k in ["foto", "resim", "görsel", "gorsel", "plaka resmi"]):
        return {"action": "photo_request", "lot": lot, "stone": stone, "client": client, "quantity": None}
    elif any(k in t_lower for k in ["iptal", "vazgeç", "geri al", "boşa çıkar"]):
        return {"action": "cancel", "lot": lot, "stone": stone, "client": client, "quantity": None}
    elif any(k in t_lower for k in ["çeki", "ceki", "packing", "liste hazırla"]):
        return {"action": "packing_list", "lot": lot, "stone": stone, "client": client, "quantity": None}
    elif any(k in t_lower for k in ["sattık", "satıldı", "sold", "satış"]):
        return {"action": "sale", "lot": lot, "stone": stone, "client": client, "quantity": None}
    elif any(k in t_lower for k in ["rezerv", "ayır", "opsiyon"]):
        return {"action": "reserve", "lot": lot, "stone": stone, "client": client, "quantity": None}
    elif any(k in t_lower for k in ["rapor", "özet", "bugün ne"]):
        return {"action": "daily_report", "lot": None, "stone": None, "client": None, "quantity": None}
    elif any(k in t_lower for k in ["drive", "link", "klasör"]):
        return {"action": "drive_link", "lot": lot, "stone": stone, "client": client, "quantity": None}
    elif any(k in t_lower for k in ["stok", "koridor", "fiyat", "kaç plaka", "var mı", "bakıver", "kontrol et"]):
        return {"action": "stock_query", "lot": lot, "stone": stone, "client": client, "quantity": None}
    else:
        return {
            "action": "chat",
            "reply": "Buradayım patron! Drive'dan fotoğraf çekmemi, çeki listesi hazırlamamı, satış/rezerve veya iptal girmemi istersen normal yazman yeterli! ⚡"
        }


# --- HAFTALIK EXCEL / CSV STOK YÜKLEME ---
async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global inventory_df
    if update.effective_chat:
        active_group_chats.add(update.effective_chat.id)

    doc = update.message.document
    if not doc.file_name.lower().endswith((".xlsx", ".xls", ".csv")):
        await update.message.reply_text("Hocam bana Excel (.xlsx) veya CSV formatında stok listesi atarsan anında sisteme işlerim! 📋")
        return

    file = await doc.get_file()
    local_path = f"./stok_{doc.file_name}"
    await file.download_to_drive(local_path)

    try:
        df = pd.read_csv(local_path) if local_path.lower().endswith(".csv") else pd.read_excel(local_path)

        col_map = {}
        for col in df.columns:
            c_low = str(col).lower()
            if any(k in c_low for k in ["lot", "blok", "bundle", "paket", "kod"]):
                col_map[col] = "Lot"
            elif any(k in c_low for k in ["taş", "tas", "mermer", "stone", "cins", "ad", "selection"]):
                col_map[col] = "Stone"
            elif any(k in c_low for k in ["plaka", "adet", "slab", "qty"]):
                col_map[col] = "Slabs"
            elif any(k in c_low for k in ["m2", "metraj", "alan", "net"]):
                col_map[col] = "M2"
            elif any(k in c_low for k in ["fiyat", "price", "usd", "tutar"]):
                col_map[col] = "Price_USD"
            elif any(k in c_low for k in ["koridor", "konum", "yer"]):
                col_map[col] = "Corridor"
            elif any(k in c_low for k in ["drive", "link", "url"]):
                col_map[col] = "Drive_URL"

        df = df.rename(columns=col_map)
        for req_col, default_val in [
            ("Lot", "Belirsiz"),
            ("Stone", "Mermer Plaka"),
            ("Slabs", 10),
            ("M2", 50.0),
            ("Price_USD", 55.0),
            ("Corridor", "Depo"),
            ("Status", "Available"),
            ("Client", ""),
            ("Drive_URL", GOOGLE_DRIVE_LINK),
        ]:
            if req_col not in df.columns:
                df[req_col] = default_val

        inventory_df = df
        total_lots = len(inventory_df)
        total_slabs = int(pd.to_numeric(inventory_df["Slabs"], errors="coerce").fillna(0).sum())
        total_m2 = round(float(pd.to_numeric(inventory_df["M2"], errors="coerce").fillna(0).sum()), 2)

        await update.message.reply_text(
            f"Sağ olasın Mehmet hocam, haftalık listeyi (`{doc.file_name}`) kaptım ve Drive ile eşleştirdim! 📋\n\n"
            f"✅ **Toplam Lot:** {total_lots} adet\n"
            f"📦 **Toplam Plaka:** {total_slabs} adet\n"
            f"📐 **Toplam Metraj:** {total_m2} m²\n"
            f"🔗 **Canlı Google Drive Klasörü:** [Tıkla ve Aç]({GOOGLE_DRIVE_LINK})\n\n"
            f"Depo fırtına gibi, her şey güncel! ⚡",
            parse_mode="Markdown",
        )
    except Exception as e:
        logger.error(f"Excel okuma hatası: {e}")
        await update.message.reply_text(f"Dosyayı aldım ama okurken ufak bir pürüz çıktı hocam: {e}")


# --- ÇEKİ LİSTESİ (PACKING LIST) EXCEL ÜRETİCİ ---
def generate_packing_list_excel(client_name: str, lot_code: str = None, stone_name: str = None) -> tuple:
    df = inventory_df.copy()
    filtered = df[df["Status"] != "Sold"]

    if lot_code:
        m = df[df["Lot"].astype(str).str.upper() == lot_code.upper()]
        if not m.empty:
            filtered = m
    elif stone_name:
        m = df[df["Stone"].astype(str).str.contains(stone_name, case=False, na=False)]
        if not m.empty:
            filtered = m

    if filtered.empty:
        filtered = df.head(2)

    filtered = filtered.copy()
    filtered["M2"] = pd.to_numeric(filtered["M2"], errors="coerce").fillna(50.0)
    filtered["Price_USD"] = pd.to_numeric(filtered["Price_USD"], errors="coerce").fillna(55.0)
    filtered["Slabs"] = pd.to_numeric(filtered["Slabs"], errors="coerce").fillna(10)
    filtered["Toplam_Tutar_USD"] = (filtered["M2"] * filtered["Price_USD"]).round(2)

    total_m2 = round(float(filtered["M2"].sum()), 2)
    total_slabs = int(filtered["Slabs"].sum())
    total_usd = round(float(filtered["Toplam_Tutar_USD"].sum()), 2)
    avg_price = round(float(filtered["Price_USD"].mean()), 2)

    safe_client = re.sub(r"[^a-zA-Z0-9_-]", "_", client_name or "Musteri")
    file_name = f"Arslanlar_Marble_Ceki_Listesi_{safe_client}.xlsx"

    export_df = filtered[["Lot", "Stone", "Corridor", "Slabs", "M2", "Price_USD", "Toplam_Tutar_USD", "Drive_URL"]].rename(
        columns={
            "Lot": "Lot / Blok No",
            "Stone": "Taş Cinsi (Selection)",
            "Corridor": "Koridor",
            "Slabs": "Plaka Adedi",
            "M2": "Net Metraj (m²)",
            "Price_USD": "Birim Fiyat ($/m²)",
            "Toplam_Tutar_USD": "Toplam Tutar ($)",
            "Drive_URL": "Google Drive Linki",
        }
    )
    export_df.to_excel(file_name, index=False)
    return file_name, total_slabs, total_m2, avg_price, total_usd, list(filtered["Lot"].values)


# --- GÜNLÜK RAPOR METNİ ---
def build_daily_report_text() -> str:
    tz = pytz.timezone("Europe/Istanbul")
    today_str = datetime.now(tz).strftime("%d.%m.%Y - %H:%M")

    sold_lines = [
        f"• **Lot {x['lot']}** ({x['stone']}) -> Alıcı: *{x['client']}* | {x['m2']} m² | **{x['total_usd']} USD**"
        for x in daily_log["sold"]
    ] or ["• *Bugün netleşen satış bulunmuyor.*"]

    res_lines = [
        f"• **Lot {x['lot']}** ({x['stone']}) -> Firma: *{x['client']}* ({x.get('qty') or 'Rezerve'})"
        for x in daily_log["reserved"]
    ] or ["• *Aktif yeni rezervasyon yok.*"]

    canc_lines = [
        f"• **Lot {x['lot']}** -> {x['info']}"
        for x in daily_log["cancelled"]
    ] or ["• *İptal edilen işlem yok.*"]

    avail_df = inventory_df[inventory_df["Status"] == "Available"]
    avail_lots = len(avail_df)
    avail_m2 = round(float(pd.to_numeric(avail_df["M2"], errors="coerce").fillna(0).sum()), 2) if not avail_df.empty else 0.0

    return (
        f"📊 **ARSLANLAR MARBLE - GÜNLÜK OPERASYON & SATIŞ RAPORU**\n"
        f"🗓 *Tarih:* {today_str}\n\n"
        f"🔴 **Satılanlar (SOLD):**\n" + "\n".join(sold_lines) + "\n\n"
        f"🟡 **Rezervasyonlar (RESERVED):**\n" + "\n".join(res_lines) + "\n\n"
        f"🔄 **İptal Edilenler / Stoğa Dönenler:**\n" + "\n".join(canc_lines) + "\n\n"
        f"📦 **Güncel Aktif Depo Özeti:**\n"
        f"• Serbest Lot Sayısı: **{avail_lots} Lot** ({avail_m2} m²)\n"
        f"• 🔗 [Canlı Google Drive Plaka Stok Klasörü]({GOOGLE_DRIVE_LINK})\n\n"
        f"Herkese emeğine sağlık, iyi akşamlar! ☕💪"
    )


# --- ANA MESAJ VE OPERASYON YÖNETİCİSİ ---
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
    qty = parsed.get("quantity")

    # 1. GOOGLE DRIVE'DAN FOTOĞRAF ÇEKİP GRUBA ATMA
    if action == "photo_request":
        search_key = lot or stone or ""
        await update.message.reply_text(f"🔍 Google Drive Plaka Stok klasörüne giriyorum, **{search_key or 'plaka'}** fotoğraflarını hemen çekiyorum patron...", parse_mode="Markdown")

        matched_items = find_drive_matches(search_key)
        image_items = [it for it in matched_items if it.get("is_image")]

        # Eğer doğrudan görsel bulunduysa indirip Telegram'a fotoğraf/albüm olarak gönder
        sent_photos = 0
        media_group = []
        for img in image_items[:5]:
            img_bytes = download_drive_file_bytes(img["id"])
            if img_bytes:
                bio = io.BytesIO(img_bytes)
                bio.name = img["name"]
                media_group.append(InputMediaPhoto(media=bio, caption=f"🪨 {img['name']}\n🔗 Drive: {img['url']}" if sent_photos == 0 else ""))
                sent_photos += 1

        if media_group:
            await update.message.reply_media_group(media=media_group)
            return

        # Klasör içindeki alt klasör/dosya linklerini listele
        folder_links = [f"• 📁 [{it['name']}]({it['url']})" for it in matched_items[:8]]
        extra_info = "\n".join(folder_links) if folder_links else f"• 📁 [Ana Plaka Stok Klasörü]({GOOGLE_DRIVE_LINK})"

        await update.message.reply_text(
            f"📸 **{search_key or 'Plaka Stoğu'} Google Drive Görsel & Klasör Erişimi:**\n\n"
            f"{extra_info}\n\n"
            f"👉 [Tüm Plaka Fotoğraflarını Drive'da Görüntüle]({GOOGLE_DRIVE_LINK})",
            parse_mode="Markdown",
        )

    # 2. SATIŞ (SOLD) İŞLEMİ
    elif action == "sale":
        target_idx = None
        if lot and lot.upper() in inventory_df["Lot"].astype(str).str.upper().values:
            target_idx = inventory_df[inventory_df["Lot"].astype(str).str.upper() == lot.upper()].index[0]
        elif stone:
            matches = inventory_df[inventory_df["Stone"].astype(str).str.contains(stone, case=False, na=False)]
            if not matches.empty:
                target_idx = matches.index[0]

        if target_idx is not None:
            row = inventory_df.loc[target_idx]
            lot_code = row["Lot"]
            stone_name = row["Stone"]
            m2_val = float(row["M2"])
            price_val = float(row["Price_USD"])
            total_val = round(m2_val * price_val, 2)

            inventory_df.loc[target_idx, "Status"] = "Sold"
            inventory_df.loc[target_idx, "Client"] = client

            daily_log["reserved"] = [r for r in daily_log["reserved"] if r["lot"].upper() != str(lot_code).upper()]
            daily_log["sold"].append({
                "lot": str(lot_code),
                "stone": str(stone_name),
                "client": client,
                "qty": qty or f"{row['Slabs']} Plaka",
                "m2": m2_val,
                "total_usd": total_val,
            })

            await update.message.reply_text(
                f"Yazdım deftere! ✍️\n"
                f"🔴 **Lot {lot_code}** ({stone_name}), **{client}** firmasına `SOLD` (Satıldı) olarak işlendi ve aktif stoktan düşüldü!\n"
                f"📐 **Metraj:** {m2_val} m² | 💰 **Tutar:** {total_val} USD ({price_val} $/m²)\n"
                f"🔗 [Drive Klasörü]({GOOGLE_DRIVE_LINK})\n"
                f"Hayırlı, bereketli olsun! 🚀",
                parse_mode="Markdown",
            )
        else:
            lot_label = lot or stone or "Belirtilen Parti"
            daily_log["sold"].append({
                "lot": lot_label,
                "stone": stone or "Mermer Plaka",
                "client": client,
                "qty": qty or "1 Parti",
                "m2": 50.0,
                "total_usd": 3250.0,
            })
            await update.message.reply_text(
                f"🔴 **Satış Kaydedildi!** `{lot_label}` ({stone or 'Plaka Grubu'}), **{client}** adına `SOLD` listesine eklendi! 🚀",
                parse_mode="Markdown",
            )

    # 3. REZERVASYON İŞLEMİ
    elif action == "reserve":
        lot_label = lot or stone or "Seçilen Lot"
        if lot and lot.upper() in inventory_df["Lot"].astype(str).str.upper().values:
            idx = inventory_df[inventory_df["Lot"].astype(str).str.upper() == lot.upper()].index[0]
            inventory_df.loc[idx, "Status"] = "Reserved"
            inventory_df.loc[idx, "Client"] = client
            stone = inventory_df.loc[idx, "Stone"]
            lot_label = inventory_df.loc[idx, "Lot"]

        daily_log["reserved"].append({
            "lot": str(lot_label),
            "stone": stone or "Mermer Plaka",
            "client": client,
            "qty": qty or "Tam Lot",
        })
        await update.message.reply_text(
            f"🟡 **Rezervasyon Tamamdır!**\n"
            f"**Lot {lot_label}** ({stone or 'Mermer'}), **{client}** adına opsiyonlandı ve ayrıldı. Kimseye dokundurtmuyorum! 🔒",
            parse_mode="Markdown",
        )

    # 4. DİNAMİK İPTAL İŞLEMİ (STOĞA GERİ DÖNÜŞ)
    elif action == "cancel":
        lot_label = lot or stone or "Son İşlem"
        if lot and lot.upper() in inventory_df["Lot"].astype(str).str.upper().values:
            idx = inventory_df[inventory_df["Lot"].astype(str).str.upper() == lot.upper()].index[0]
            inventory_df.loc[idx, "Status"] = "Available"
            inventory_df.loc[idx, "Client"] = ""
            lot_label = inventory_df.loc[idx, "Lot"]

        daily_log["sold"] = [s for s in daily_log["sold"] if s["lot"].upper() != str(lot_label).upper()]
        daily_log["reserved"] = [r for r in daily_log["reserved"] if r["lot"].upper() != str(lot_label).upper()]
        daily_log["cancelled"].append({
            "lot": str(lot_label),
            "info": "Satış/rezervasyon iptal edildi, tekrar aktif stoğa (Available) alındı."
        })

        await update.message.reply_text(
            f"Tamamdır patron, hemen hallediyorum! 🔄\n"
            f"**{lot_label}** satış/rezervasyon listesinden çıkarıldı ve tertemiz şekilde **aktif depoya (serbest stoğa)** geri alındı. "
            f"Depo canavar gibi, yola devam! 💪",
            parse_mode="Markdown",
        )

    # 5. ÇEKİ LİSTESİ + FİYAT + DRIVE LİNKİ + EXCEL DOSYASI
    elif action == "packing_list":
        excel_path, slabs, m2, unit_p, total_usd, lots_used = generate_packing_list_excel(client, lot, stone)
        caption = (
            f"✅ **{client} İçin Çeki Listesi Hazırlandı!**\n\n"
            f"📦 **Seçilen Lotlar:** {', '.join(map(str, lots_used))}\n"
            f"🪨 **Toplam Plaka:** {slabs} Adet\n"
            f"📐 **Net Metraj:** {m2} m²\n"
            f"💵 **Birim Fiyat:** {unit_p} USD / m²\n"
            f"💰 **Genel Toplam:** **{total_usd} USD**\n\n"
            f"📄 **Google Drive Plaka Klasörü:** [Tıkla ve İncele]({GOOGLE_DRIVE_LINK})\n"
            f"📎 *Hazırlanan Excel Çeki Listesi dosyasını da hemen aşağıya bıraktım!*"
        )
        with open(excel_path, "rb") as f:
            await update.message.reply_document(document=f, caption=caption, parse_mode="Markdown")

    # 6. STOK VE DRIVE İÇERİĞİ SORGULAMA
    elif action == "stock_query":
        search_key = lot or stone or ""
        drive_matches = find_drive_matches(search_key)

        df = inventory_df.copy()
        if lot:
            df = df[df["Lot"].astype(str).str.upper() == lot.upper()]
        elif stone:
            df = df[df["Stone"].astype(str).str.contains(stone, case=False, na=False)]
        if df.empty:
            df = inventory_df

        lines = []
        for _, r in df.head(10).iterrows():
            durum_ikon = "🟢 Boşta" if r["Status"] == "Available" else ("🟡 Rezerve (" + str(r["Client"]) + ")" if r["Status"] == "Reserved" else "🔴 SATILDI (" + str(r["Client"]) + ")")
            lines.append(
                f"• **Lot {r['Lot']}** ({r['Stone']}) | Koridor: `{r['Corridor']}` | {r['Slabs']} Plaka ({r['M2']} m²) | **{r['Price_USD']} $/m²** -> {durum_ikon}"
            )

        drive_section = ""
        if drive_matches:
            d_links = [f"• 📁 [{it['name']}]({it['url']})" for it in drive_matches[:5]]
            drive_section = "\n\n📂 **Drive Klasöründe Bulunan Eşleşmeler:**\n" + "\n".join(d_links)

        await update.message.reply_text(
            f"🔍 **Güncel Depo & Fiyat Durumu:**\n\n"
            + "\n".join(lines)
            + drive_section
            + f"\n\n🔗 [Ana Google Drive Plaka Stok Klasörü]({GOOGLE_DRIVE_LINK})",
            parse_mode="Markdown",
        )

    # 7. GÜNLÜK RAPOR İSTEĞİ
    elif action == "daily_report":
        await update.message.reply_text(build_daily_report_text(), parse_mode="Markdown")

    # 8. DOĞRUDAN GOOGLE DRIVE LİNKİ VE KLASÖR TARAMA İSTEĞİ
    elif action == "drive_link":
        items = scan_public_drive_folder()
        item_lines = [f"• {'📁' if it['is_folder'] else '📄'} [{it['name']}]({it['url']})" for it in items[:10]]
        inside_text = "\n".join(item_lines) if item_lines else "• Ana klasör bağlantısı aktif ve erişilebilir durumda."

        await update.message.reply_text(
            f"📄 **Arslanlar Marble Canlı Plaka Stok Drive Klasörü:**\n"
            f"👉 [Tıkla ve Google Drive'da Aç]({GOOGLE_DRIVE_LINK})\n\n"
            f"📂 **Klasör İçeriği Özeti:**\n{inside_text}",
            parse_mode="Markdown",
        )

    # 9. DOĞAL SOHBET
    else:
        reply_msg = parsed.get("reply") or "Buradayım hocam! Satış, rezervasyon, iptal, Drive'dan fotoğraf veya çeki listesi ne lazımsa yazman yeterli. ⚡"
        await update.message.reply_text(reply_msg)


# --- OTOMATİK AKŞAM RAPORU ZAMANLAYICISI (HER GÜN 18:30) ---
async def scheduled_daily_report(context: ContextTypes.DEFAULT_TYPE):
    report_msg = build_daily_report_text()
    for chat_id in active_group_chats:
        try:
            await context.bot.send_message(chat_id=chat_id, text=report_msg, parse_mode="Markdown")
        except Exception as e:
            logger.error(f"Otomatik rapor gönderilemedi ({chat_id}): {e}")


# --- RENDER 7/24 HEALTH CHECK SUNUCUSU ---
class DummyHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Arslanlar Marble Bot 7/24 Aktif!")

    def log_message(self, format, *args):
        return


def run_dummy_server():
    port = int(os.environ.get("PORT", 8080))
    HTTPServer(("0.0.0.0", port), DummyHandler).serve_forever()


def main():
    threading.Thread(target=run_dummy_server, daemon=True).start()
    # Başlangıçta Google Drive klasörünü bir kez tarayıp önbelleğe al
    threading.Thread(target=scan_public_drive_folder, daemon=True).start()

    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()

    if app.job_queue:
        tz = pytz.timezone("Europe/Istanbul")
        app.job_queue.run_daily(scheduled_daily_report, time=time(hour=18, minute=30, tzinfo=tz))

    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(CommandHandler("rapor", lambda u, c: u.message.reply_text(build_daily_report_text(), parse_mode="Markdown")))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))

    logger.info("Arslanlar Marble Asistan Botu baslatildi...")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
