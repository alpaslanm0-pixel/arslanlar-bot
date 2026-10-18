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

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# --- ANAHTARLAR VE SABİT BAĞLANTILAR ---
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "8822872793:AAEn0p_FLlH1CaB41cHhwHV7ADh7q6o6VGU")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "AQ.Ab8RN6L4udr1eup7XD4FTqux1c9TEK1xlEtAvEuQ5QP0Y7f-Nw")

GOOGLE_DRIVE_LINK = os.getenv(
    "GOOGLE_DRIVE_LINK",
    "https://drive.google.com/drive/folders/121PWmKo_5mCzJaRrSZWu3BHOmVFafb2F?usp=drive_link"
)

genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-1.5-flash")

# --- CANLI STOK VE HAREKET HAFIZASI ---
inventory_df = pd.DataFrame([
    {"Lot": "A-283", "Stone": "Violet Breccia", "Category": "Arslanlar Collection", "Slabs": 12, "M2": 68.4, "Price_USD": 65.0, "Corridor": "A-1", "Status": "Available", "Client": "", "Drive_URL": GOOGLE_DRIVE_LINK},
    {"Lot": "B-12", "Stone": "Beige", "Category": "Arslanlar Series", "Slabs": 18, "M2": 95.0, "Price_USD": 42.0, "Corridor": "B-3", "Status": "Available", "Client": "", "Drive_URL": GOOGLE_DRIVE_LINK},
    {"Lot": "C-105", "Stone": "Calacatta Viola", "Category": "Arslanlar Exclusive collection", "Slabs": 8, "M2": 44.2, "Price_USD": 85.0, "Corridor": "A-2", "Status": "Available", "Client": "", "Drive_URL": GOOGLE_DRIVE_LINK},
])

daily_log = {"sold": [], "reserved": [], "cancelled": []}
drive_deep_index = []
active_group_chats = set()
is_scanning_drive = False


# --- YARDIMCI FONKSİYONLAR ---
def extract_folder_key(url: str) -> str:
    m = re.search(r"/folders/([a-zA-Z0-9_-]+)", url)
    return m.group(1) if m else ""


def normalize_key(text: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def download_drive_file_bytes(file_id: str, is_google_sheet: bool = False) -> bytes:
    """Drive'daki Excel, Google Sheets veya fotoğraf dosyasını indirir."""
    headers = {"User-Agent": "Mozilla/5.0"}
    if is_google_sheet:
        url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
    else:
        url = f"https://drive.google.com/uc?export=download&id={file_id}"
    try:
        resp = requests.get(url, headers=headers, timeout=20)
        if resp.status_code == 200:
            return resp.content
    except Exception as e:
        logger.warning(f"Dosya indirme uyarısı ({file_id}): {e}")
    return b""


def parse_excel_bytes_to_summary(excel_bytes: bytes, fallback_lot: str, fallback_stone: str, category: str, drive_url: str) -> list:
    """
    Drive içindeki bir Çeki Listesi / Stok Excel dosyasını okur.
    İster plaka plaka çeki listesi olsun, ister genel liste olsun metraj ve plaka sayısını çıkarır.
    """
    records = []
    try:
        df = pd.read_excel(io.BytesIO(excel_bytes))
        if df.empty:
            return records

        # Sütun isimlerini normalize ederek kontrol et
        col_map = {}
        for col in df.columns:
            c_low = str(col).lower()
            if any(k in c_low for k in ["lot", "blok", "bundle", "paket", "no", "kod"]):
                col_map[col] = "Lot"
            elif any(k in c_low for k in ["taş", "tas", "mermer", "stone", "cins", "selection", "ürün"]):
                col_map[col] = "Stone"
            elif any(k in c_low for k in ["plaka", "adet", "slab", "qty", "quantity", "pcs"]):
                col_map[col] = "Slabs"
            elif any(k in c_low for k in ["m2", "metraj", "alan", "sqm", "net", "toplam m2"]):
                col_map[col] = "M2"
            elif any(k in c_low for k in ["fiyat", "price", "usd", "$"]):
                col_map[col] = "Price_USD"
            elif any(k in c_low for k in ["koridor", "konum", "yer"]):
                col_map[col] = "Corridor"

        renamed = df.rename(columns=col_map)

        # Eğer çok satırlı bir genel stok tablosuysa satır satır al
        if "Lot" in renamed.columns and "M2" in renamed.columns and len(renamed) > 1:
            for _, r in renamed.iterrows():
                lot_val = str(r.get("Lot", "")).strip()
                if not lot_val or lot_val.lower() in ("nan", "toplam", "total", "none"):
                    continue
                m2_val = pd.to_numeric(r.get("M2", 0), errors="coerce")
                slabs_val = pd.to_numeric(r.get("Slabs", 1), errors="coerce")
                price_val = pd.to_numeric(r.get("Price_USD", 60.0), errors="coerce")
                records.append({
                    "Lot": lot_val,
                    "Stone": str(r.get("Stone", fallback_stone)).strip() if "Stone" in renamed.columns else fallback_stone,
                    "Category": category,
                    "Slabs": int(slabs_val) if pd.notna(slabs_val) and slabs_val > 0 else 1,
                    "M2": round(float(m2_val), 2) if pd.notna(m2_val) and m2_val > 0 else 50.0,
                    "Price_USD": round(float(price_val), 2) if pd.notna(price_val) and price_val > 0 else 60.0,
                    "Corridor": str(r.get("Corridor", "Drive")),
                    "Status": "Available",
                    "Client": "",
                    "Drive_URL": drive_url,
                })
        else:
            # Eğer tek bir lotun plaka bazlı çeki listesiyse (her satır 1 plaka ise)
            numeric_cols = df.select_dtypes(include="number")
            total_m2 = 50.0
            if "M2" in renamed.columns:
                m2_series = pd.to_numeric(renamed["M2"], errors="coerce").dropna()
                # Son satır toplam satırıysa onu veya toplamın yarısını kontrol et
                total_m2 = round(float(m2_series.sum()), 2) if not m2_series.empty else 50.0
            elif not numeric_cols.empty:
                total_m2 = round(float(numeric_cols.iloc[:, -1].dropna().sum()), 2)

            slab_count = max(1, len(df.dropna(how="all")) - 1)
            records.append({
                "Lot": fallback_lot,
                "Stone": fallback_stone,
                "Category": category,
                "Slabs": slab_count,
                "M2": total_m2 if total_m2 > 0 else 50.0,
                "Price_USD": 60.0,
                "Corridor": "Drive",
                "Status": "Available",
                "Client": "",
                "Drive_URL": drive_url,
            })
    except Exception as e:
        logger.warning(f"Excel çözümleme uyarısı ({fallback_lot}): {e}")

    return records


# --- DERİNLEMESİNE GOOGLE DRIVE TARAYICI VE STOK OLUŞTURUCU ---
def fetch_single_folder(folder_url: str, parent_name: str = "Ana Klasör") -> list:
    folder_key = extract_folder_key(folder_url)
    if not folder_key:
        return []

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
                is_gsheet = "spreadsheets" in href

                items.append({
                    "id": entry_id,
                    "name": title,
                    "parent": parent_name,
                    "url": href,
                    "is_folder": is_folder,
                    "is_image": title.lower().endswith((".jpg", ".jpeg", ".png", ".webp")),
                    "is_excel": title.lower().endswith((".xlsx", ".xls", ".csv")) or is_gsheet,
                    "is_gsheet": is_gsheet,
                })
    except Exception as e:
        logger.warning(f"Klasör okuma uyarısı ({parent_name}): {e}")

    return items


def build_deep_drive_index_and_inventory(max_depth: int = 4) -> tuple:
    """
    Arslanlar Series, Arslanlar Collection, Arslanlar Exclusive collection ve # Inventory Fast preview
    klasörlerinin tamamını 4 katman derinliğe kadar tarar, içlerindeki taş klasörlerini ve
    Excel çeki listelerini okuyarak otomatik stok tablosuna ekler.
    """
    global drive_deep_index, inventory_df, is_scanning_drive
    if is_scanning_drive:
        return len(drive_deep_index), len(inventory_df)

    is_scanning_drive = True
    logger.info("Google Drive koleksiyonları ve Excel çeki listeleri derinlemesine taranıyor...")
    all_discovered = []
    visited_urls = set()
    discovered_stock_rows = []

    queue = [(GOOGLE_DRIVE_LINK, "Ana Klasör", 1)]

    while queue:
        current_url, parent_label, depth = queue.pop(0)
        if current_url in visited_urls or depth > max_depth:
            continue
        visited_urls.add(current_url)

        children = fetch_single_folder(current_url, parent_label)
        for child in children:
            all_discovered.append(child)
            path_parts = [p.strip() for p in parent_label.split(">")]
            category_name = path_parts[1] if len(path_parts) > 1 else child["name"]

            # Alt klasörse kuyruğa ekle ve taş/lot klasörü olarak kaydet
            if child["is_folder"]:
                if depth < max_depth:
                    queue.append((child["url"], f"{parent_label} > {child['name']}", depth + 1))

                # Eğer 2. veya 3. katmandaki bir taş/lot klasörüyse (Örn: Arslanlar Collection > Violet Breccia > Lot A-283)
                if depth >= 2:
                    stone_guess = path_parts[2] if len(path_parts) > 2 else child["name"]
                    lot_guess = child["name"]
                    discovered_stock_rows.append({
                        "Lot": lot_guess,
                        "Stone": stone_guess,
                        "Category": category_name,
                        "Slabs": 10,
                        "M2": 50.0,
                        "Price_USD": 60.0,
                        "Corridor": category_name,
                        "Status": "Available",
                        "Client": "",
                        "Drive_URL": child["url"],
                    })

            # Eğer klasörün içinde Excel veya Çeki Listesi dosyası varsa indirip oku!
            elif child["is_excel"]:
                excel_bytes = download_drive_file_bytes(child["id"], is_google_sheet=child.get("is_gsheet", False))
                if excel_bytes:
                    stone_guess = path_parts[-1] if len(path_parts) > 1 else "Mermer"
                    lot_guess = os.path.splitext(child["name"])[0]
                    parsed_rows = parse_excel_bytes_to_summary(
                        excel_bytes,
                        fallback_lot=lot_guess,
                        fallback_stone=stone_guess,
                        category=category_name,
                        drive_url=current_url,
                    )
                    discovered_stock_rows.extend(parsed_rows)

    if all_discovered:
        drive_deep_index = all_discovered

    # Drive'dan çıkan stok verilerini mevcut (veya Telegram'dan atılan) stok listesiyle birleştir
    if discovered_stock_rows:
        drive_df = pd.DataFrame(discovered_stock_rows)
        combined = pd.concat([inventory_df, drive_df], ignore_index=True)
        # Aynı Lot kodu varsa mevcut (Telegram'dan yüklenen veya işlem gören) kaydı koru, Drive URL'sini güncelle
        combined["_norm_lot"] = combined["Lot"].astype(str).apply(normalize_key)
        combined = combined.drop_duplicates(subset=["_norm_lot"], keep="first").drop(columns=["_norm_lot"])
        inventory_df = combined.reset_index(drop=True)

    is_scanning_drive = False
    logger.info(f"Drive taraması bitti: {len(drive_deep_index)} öğe, {len(inventory_df)} stok kalemi hazır.")
    return len(drive_deep_index), len(inventory_df)


def search_in_drive(query_term: str) -> list:
    if not drive_deep_index:
        build_deep_drive_index_and_inventory(max_depth=3)

    if not query_term:
        return drive_deep_index

    q_norm = normalize_key(query_term)
    matches = [
        item for item in drive_deep_index
        if q_norm in normalize_key(item["name"]) or q_norm in normalize_key(item["parent"])
    ]

    extra_children = []
    for m in matches[:6]:
        if m["is_folder"]:
            sub = fetch_single_folder(m["url"], f"{m['parent']} > {m['name']}")
            extra_children.extend(sub)

    seen = set()
    final_results = []
    for it in matches + extra_children:
        if it["url"] not in seen:
            seen.add(it["url"])
            final_results.append(it)

    return final_results


# --- DOĞAL DİL ANLAMA MOTORU ---
def analyze_message(text: str) -> dict:
    stock_summary = inventory_df[["Lot", "Stone", "Slabs", "M2", "Price_USD", "Status", "Client"]].head(35).to_dict(orient="records")

    prompt = f"""
    Sen Arslanlar Marble mermer fabrikasının operasyon ve satış asistanısın.
    Mehmet Morkoç ve satış ekibi seninle grupta doğal bir insan gibi konuşuyor.

    Güncel Stok Özeti: {json.dumps(stock_summary, ensure_ascii=False)}

    Kullanıcının mesajını analiz et ve SADECE aşağıdaki JSON formatında yanıt ver:
    {{
        "action": "sale" | "reserve" | "cancel" | "packing_list" | "photo_request" | "stock_query" | "sync_drive" | "export_stock" | "daily_report" | "drive_link" | "chat",
        "lot": "Tespit edilen lot/blok/kod numarası (örn: A-283, B-12) veya null",
        "stone": "Tespit edilen mermer/koleksiyon adı (örn: Violet Breccia, Beige, Exclusive, Series, Collection) veya null",
        "client": "Müşteri/firma adı (örn: Glaze Dubai, Skystone, Mitra, Jasmine Stone) veya null",
        "quantity": "Miktar/kasa/plaka detayı veya null",
        "reply": "Mehmet Morkoç'un samimi, enerjik, iş bitirici üslubuyla Türkçe doğal yanıt"
    }}

    Kurallar:
    - "drive'ı tara", "tüm klasörleri tara", "drive'dan stok çıkar", "klasörleri güncelle" -> "sync_drive"
    - "tüm stok listesini excel at", "genel stok tablosunu gönder", "master stok" -> "export_stock"
    - Fotoğraf, görsel, resim isteniyorsa -> "photo_request"
    - Müşteri için çeki listesi, packing list isteniyorsa -> "packing_list"
    - Satıldı, sattık, sold deniyorsa -> "sale"
    - Rezerve, opsiyon, ayır deniyorsa -> "reserve"
    - İptal, vazgeçti, stoğa geri al deniyorsa -> "cancel"
    - Drive linki, klasör linki isteniyorsa -> "drive_link"
    - Stokta ne var, fiyatı ne, kaç plaka var deniyorsa -> "stock_query"
    - Günlük rapor, bugün ne yaptık deniyorsa -> "daily_report"

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

    stone = None
    known_terms = list(inventory_df["Stone"].dropna().unique()) + [
        "Violet Breccia", "Calacatta Viola", "Beige", "Afyon",
        "Collection", "Exclusive", "Series", "Inventory", "Workshop"
    ]
    for s in known_terms:
        if str(s).lower() in t_lower or str(s).split()[0].lower() in t_lower:
            stone = str(s)
            break

    client = None
    for c in ["Glaze Dubai", "Skystone", "Jasmine Stone", "Mitra", "İZFAŞ"]:
        if c.lower() in t_lower:
            client = c
            break

    if any(k in t_lower for k in ["tara", "senkronize", "drive'dan stok", "klasörleri oku"]):
        return {"action": "sync_drive", "lot": lot, "stone": stone, "client": client, "quantity": None}
    elif any(k in t_lower for k in ["tüm stok", "genel stok", "stok listesi yap", "stok excel"]):
        return {"action": "export_stock", "lot": lot, "stone": stone, "client": client, "quantity": None}
    elif any(k in t_lower for k in ["foto", "resim", "görsel", "gorsel", "plaka resmi"]):
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
    elif any(k in t_lower for k in ["drive", "link", "klasör", "gönder", "at"]):
        return {"action": "drive_link", "lot": lot, "stone": stone, "client": client, "quantity": None}
    elif any(k in t_lower for k in ["stok", "koridor", "fiyat", "kaç plaka", "var mı", "bakıver", "kontrol et"]):
        return {"action": "stock_query", "lot": lot, "stone": stone, "client": client, "quantity": None}
    else:
        return {
            "action": "chat",
            "reply": "Buradayım patron! İster Telegram'dan genel stok Excel'ini at, ister 'Drive klasörlerini tara ve stok listesi çıkar' de, hemen halledelim! ⚡"
        }


# --- TELEGRAM'DAN GENEL STOK EXCEL / CSV YÜKLEME VE BİRLEŞTİRME ---
async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global inventory_df
    if update.effective_chat:
        active_group_chats.add(update.effective_chat.id)

    doc = update.message.document
    if not doc.file_name.lower().endswith((".xlsx", ".xls", ".csv")):
        await update.message.reply_text("Hocam bana Excel (.xlsx) veya CSV formatında genel stok listesi atarsan anında sisteme işlerim! 📋")
        return

    file = await doc.get_file()
    local_path = f"./stok_{doc.file_name}"
    await file.download_to_drive(local_path)

    try:
        df = pd.read_csv(local_path) if local_path.lower().endswith(".csv") else pd.read_excel(local_path)

        col_map = {}
        for col in df.columns:
            c_low = str(col).lower()
            if any(k in c_low for k in ["lot", "blok", "bundle", "paket", "kod", "no"]):
                col_map[col] = "Lot"
            elif any(k in c_low for k in ["taş", "tas", "mermer", "stone", "cins", "ad", "selection"]):
                col_map[col] = "Stone"
            elif any(k in c_low for k in ["kategori", "koleksiyon", "collection", "seri"]):
                col_map[col] = "Category"
            elif any(k in c_low for k in ["plaka", "adet", "slab", "qty"]):
                col_map[col] = "Slabs"
            elif any(k in c_low for k in ["m2", "metraj", "alan", "net"]):
                col_map[col] = "M2"
            elif any(k in c_low for k in ["fiyat", "price", "usd", "tutar"]):
                col_map[col] = "Price_USD"
            elif any(k in c_low for k in ["koridor", "konum", "yer"]):
                col_map[col] = "Corridor"
            elif any(k in c_low for k in ["durum", "status"]):
                col_map[col] = "Status"
            elif any(k in c_low for k in ["müşteri", "musteri", "client", "firma"]):
                col_map[col] = "Client"
            elif any(k in c_low for k in ["drive", "link", "url"]):
                col_map[col] = "Drive_URL"

        df = df.rename(columns=col_map)
        for req_col, default_val in [
            ("Lot", "Belirsiz"),
            ("Stone", "Mermer Plaka"),
            ("Category", "Genel Stok"),
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

        # Drive'daki alt klasör linkleriyle otomatik eşleştir
        for idx, r in df.iterrows():
            lot_str = str(r["Lot"])
            stone_str = str(r["Stone"])
            matches = search_in_drive(lot_str) or search_in_drive(stone_str)
            if matches:
                df.loc[idx, "Drive_URL"] = matches[0]["url"]

        inventory_df = df
        total_lots = len(inventory_df)
        total_slabs = int(pd.to_numeric(inventory_df["Slabs"], errors="coerce").fillna(0).sum())
        total_m2 = round(float(pd.to_numeric(inventory_df["M2"], errors="coerce").fillna(0).sum()), 2)

        await update.message.reply_text(
            f"Sağ olasın Mehmet hocam, genel stok listesini (`{doc.file_name}`) aldım ve Drive klasörleriyle eşleştirdim! 📋\n\n"
            f"✅ **Toplam Lot:** {total_lots} adet\n"
            f"📦 **Toplam Plaka:** {total_slabs} adet\n"
            f"📐 **Toplam Metraj:** {total_m2} m²\n"
            f"🔗 **Canlı Google Drive Klasörü:** [Tıkla ve Aç]({GOOGLE_DRIVE_LINK})\n\n"
            f"Tüm satış, rezervasyon, iptal ve çeki listesi takibi bu güncel liste üzerinden aktif! ⚡",
            parse_mode="Markdown",
        )
    except Exception as e:
        logger.error(f"Excel okuma hatası: {e}")
        await update.message.reply_text(f"Dosyayı aldım ama okurken ufak bir pürüz çıktı hocam: {e}")


# --- ÇEKİ LİSTESİ VE MASTER STOK EXCEL ÜRETİCİ ---
def generate_packing_list_excel(client_name: str, lot_code: str = None, stone_name: str = None) -> tuple:
    df = inventory_df.copy()
    filtered = df[df["Status"] != "Sold"]

    if lot_code:
        m = df[df["Lot"].astype(str).apply(normalize_key) == normalize_key(lot_code)]
        if not m.empty:
            filtered = m
    elif stone_name:
        m = df[
            df["Stone"].astype(str).str.contains(stone_name, case=False, na=False)
            | df.get("Category", pd.Series([""] * len(df))).astype(str).str.contains(stone_name, case=False, na=False)
        ]
        if not m.empty:
            filtered = m

    if filtered.empty:
        filtered = df.head(5)

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

    cols_to_use = [c for c in ["Lot", "Stone", "Category", "Corridor", "Slabs", "M2", "Price_USD", "Toplam_Tutar_USD", "Status", "Client", "Drive_URL"] if c in filtered.columns]
    export_df = filtered[cols_to_use].rename(
        columns={
            "Lot": "Lot / Blok No",
            "Stone": "Taş Cinsi (Selection)",
            "Category": "Koleksiyon / Seri",
            "Corridor": "Koridor",
            "Slabs": "Plaka Adedi",
            "M2": "Net Metraj (m²)",
            "Price_USD": "Birim Fiyat ($/m²)",
            "Toplam_Tutar_USD": "Toplam Tutar ($)",
            "Status": "Durum",
            "Client": "Müşteri",
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
    search_key = lot or stone or ""

    # 1. DRIVE KLASÖRLERİNİ VE ÇEKİ LİSTELERİNİ DERİNLEMESİNE TARAYIP STOK ÇIKARMA
    if action == "sync_drive":
        await update.message.reply_text(
            "🔄 **Arslanlar Series, Arslanlar Collection, Arslanlar Exclusive collection ve # Inventory Fast preview** "
            "alt klasörleri ile içlerindeki Excel çeki listeleri tek tek taranıyor... Lütfen 15-20 saniye bekle patron! ⚡",
            parse_mode="Markdown",
        )
        items_count, stock_count = build_deep_drive_index_and_inventory(max_depth=4)
        excel_path, slabs, m2, _, total_usd, _ = generate_packing_list_excel("Genel_Drive_Stok_Listesi")

        caption = (
            f"✅ **Google Drive Klasörleri & Çeki Listeleri Tarandı!**\n\n"
            f"📂 **Taranan Klasör/Dosya/Fotoğraf:** {items_count} adet\n"
            f"🪨 **Birleştirilen Toplam Stok Kalemi:** {stock_count} Lot\n"
            f"📦 **Toplam Plaka:** {slabs} Adet | **Toplam Metraj:** {m2} m²\n\n"
            f"📎 *Drive'dan ve mevcut stoktan oluşturulan güncel Master Stok Excel tablosunu aşağıya bıraktım!*"
        )
        with open(excel_path, "rb") as f:
            await update.message.reply_document(document=f, caption=caption, parse_mode="Markdown")

    # 2. GÜNCEL MASTER STOK LİSTESİNİ EXCEL OLARAK İNDİRME
    elif action == "export_stock":
        excel_path, slabs, m2, _, total_usd, _ = generate_packing_list_excel("Arslanlar_Guncel_Master_Stok", lot, stone)
        caption = (
            f"📊 **Arslanlar Marble Güncel Master Stok Listesi**\n"
            f"🪨 **Toplam Plaka:** {slabs} Adet | 📐 **Toplam Metraj:** {m2} m²\n"
            f"💰 **Toplam Stok Değeri:** {total_usd} USD\n"
            f"🔗 [Google Drive Ana Klasör]({GOOGLE_DRIVE_LINK})"
        )
        with open(excel_path, "rb") as f:
            await update.message.reply_document(document=f, caption=caption, parse_mode="Markdown")

    # 3. GOOGLE DRIVE LİNKİ VEYA FOTOĞRAF İSTEĞİ
    elif action in ("drive_link", "photo_request"):
        if search_key:
            await update.message.reply_text(
                f"🔍 Drive koleksiyonlarında (`{search_key}`) klasör ve fotoğraflar aranıyor patron...",
                parse_mode="Markdown"
            )
            matched_items = search_in_drive(search_key)
        else:
            matched_items = fetch_single_folder(GOOGLE_DRIVE_LINK, "Ana Klasör")

        image_items = [it for it in matched_items if it.get("is_image")]
        if image_items and action == "photo_request":
            media_group = []
            for idx, img in enumerate(image_items[:5]):
                img_bytes = download_drive_file_bytes(img["id"])
                if img_bytes:
                    bio = io.BytesIO(img_bytes)
                    bio.name = img["name"]
                    caption_txt = f"🪨 {img['name']} ({img['parent']})\n🔗 Drive: {img['url']}" if idx == 0 else ""
                    media_group.append(InputMediaPhoto(media=bio, caption=caption_txt))
            if media_group:
                await update.message.reply_media_group(media=media_group)
                return

        if matched_items:
            item_lines = [
                f"• {'📁' if it['is_folder'] else ('🖼️' if it['is_image'] else '📄')} [{it['name']}]({it['url']})  *({it['parent']})*"
                for it in matched_items[:12]
            ]
            await update.message.reply_text(
                f"✅ **`{search_key or 'Ana Klasör'}` İçin Google Drive Sonuçları:**\n\n"
                + "\n".join(item_lines)
                + f"\n\n👉 [Ana Plaka Stok Klasörünü Aç]({GOOGLE_DRIVE_LINK})",
                parse_mode="Markdown",
            )
        else:
            sub_folders = [it for it in drive_deep_index if it["is_folder"]][:12]
            sub_lines = [f"• 📁 [{it['name']}]({it['url']}) *({it['parent']})*" for it in sub_folders]
            await update.message.reply_text(
                f"⚠️ Drive alt klasörlerinde `{search_key}` adıyla birebir eşleşen bir kayıt bulunamadı.\n\n"
                f"📂 **Taranan Koleksiyon Alt Klasörleri:**\n"
                + ("\n".join(sub_lines) if sub_lines else f"• [Ana Klasör]({GOOGLE_DRIVE_LINK})")
                + f"\n\n💡 *İpucu: Yeni yüklediysen 'Drive klasörlerini tara' yazarak tüm alt klasörleri yeniden taratabilirsin.*",
                parse_mode="Markdown",
            )

    # 4. SATIŞ (SOLD) İŞLEMİ
    elif action == "sale":
        target_idx = None
        if lot:
            m = inventory_df[inventory_df["Lot"].astype(str).apply(normalize_key) == normalize_key(lot)]
            if not m.empty:
                target_idx = m.index[0]
        elif stone:
            m = inventory_df[inventory_df["Stone"].astype(str).str.contains(stone, case=False, na=False)]
            if not m.empty:
                target_idx = m.index[0]

        if target_idx is not None:
            row = inventory_df.loc[target_idx]
            lot_code = row["Lot"]
            stone_name = row["Stone"]
            m2_val = float(row["M2"])
            price_val = float(row["Price_USD"])
            total_val = round(m2_val * price_val, 2)

            inventory_df.loc[target_idx, "Status"] = "Sold"
            inventory_df.loc[target_idx, "Client"] = client

            daily_log["reserved"] = [r for r in daily_log["reserved"] if normalize_key(r["lot"]) != normalize_key(lot_code)]
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

    # 5. REZERVASYON İŞLEMİ
    elif action == "reserve":
        lot_label = lot or stone or "Seçilen Lot"
        if lot:
            m = inventory_df[inventory_df["Lot"].astype(str).apply(normalize_key) == normalize_key(lot)]
            if not m.empty:
                idx = m.index[0]
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

    # 6. DİNAMİK İPTAL İŞLEMİ (STOĞA GERİ DÖNÜŞ)
    elif action == "cancel":
        lot_label = lot or stone or "Son İşlem"
        if lot:
            m = inventory_df[inventory_df["Lot"].astype(str).apply(normalize_key) == normalize_key(lot)]
            if not m.empty:
                idx = m.index[0]
                inventory_df.loc[idx, "Status"] = "Available"
                inventory_df.loc[idx, "Client"] = ""
                lot_label = inventory_df.loc[idx, "Lot"]

        daily_log["sold"] = [s for s in daily_log["sold"] if normalize_key(s["lot"]) != normalize_key(lot_label)]
        daily_log["reserved"] = [r for r in daily_log["reserved"] if normalize_key(r["lot"]) != normalize_key(lot_label)]
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

    # 7. ÇEKİ LİSTESİ + FİYAT + DRIVE LİNKİ + EXCEL DOSYASI
    elif action == "packing_list":
        excel_path, slabs, m2, unit_p, total_usd, lots_used = generate_packing_list_excel(client, lot, stone)
        drive_matches = search_in_drive(search_key)
        specific_drive_url = drive_matches[0]["url"] if drive_matches else GOOGLE_DRIVE_LINK

        caption = (
            f"✅ **{client} İçin Çeki Listesi Hazırlandı!**\n\n"
            f"📦 **Seçilen Lotlar:** {', '.join(map(str, lots_used[:8]))}\n"
            f"🪨 **Toplam Plaka:** {slabs} Adet\n"
            f"📐 **Net Metraj:** {m2} m²\n"
            f"💵 **Birim Fiyat:** {unit_p} USD / m²\n"
            f"💰 **Genel Toplam:** **{total_usd} USD**\n\n"
            f"📄 **Google Drive Bağlantısı:** [İlgili Klasörü / Dosyayı Aç]({specific_drive_url})\n"
            f"📎 *Hazırlanan Excel Çeki Listesi dosyasını da hemen aşağıya bıraktım!*"
        )
        with open(excel_path, "rb") as f:
            await update.message.reply_document(document=f, caption=caption, parse_mode="Markdown")

    # 8. STOK VE DRIVE İÇERİĞİ SORGULAMA
    elif action == "stock_query":
        drive_matches = search_in_drive(search_key)

        df = inventory_df.copy()
        if lot:
            df = df[df["Lot"].astype(str).apply(normalize_key) == normalize_key(lot)]
        elif stone:
            df = df[
                df["Stone"].astype(str).str.contains(stone, case=False, na=False)
                | df.get("Category", pd.Series([""] * len(df))).astype(str).str.contains(stone, case=False, na=False)
            ]
        if df.empty:
            df = inventory_df

        lines = []
        for _, r in df.head(12).iterrows():
            durum_ikon = "🟢 Boşta" if r["Status"] == "Available" else ("🟡 Rezerve (" + str(r["Client"]) + ")" if r["Status"] == "Reserved" else "🔴 SATILDI (" + str(r["Client"]) + ")")
            lines.append(
                f"• **Lot {r['Lot']}** ({r['Stone']}) | {r['Slabs']} Plaka ({r['M2']} m²) | **{r['Price_USD']} $/m²** -> {durum_ikon}"
            )

        drive_section = ""
        if drive_matches:
            d_links = [f"• {'📁' if it['is_folder'] else ('🖼️' if it['is_image'] else '📄')} [{it['name']}]({it['url']}) *({it['parent']})*" for it in drive_matches[:6]]
            drive_section = "\n\n📂 **Drive Alt Klasörlerinde Bulunanlar:**\n" + "\n".join(d_links)

        await update.message.reply_text(
            f"🔍 **Güncel Depo & Fiyat Durumu:**\n\n"
            + "\n".join(lines)
            + drive_section
            + f"\n\n🔗 [Ana Google Drive Plaka Stok Klasörü]({GOOGLE_DRIVE_LINK})",
            parse_mode="Markdown",
        )

    # 9. GÜNLÜK RAPOR İSTEĞİ
    elif action == "daily_report":
        await update.message.reply_text(build_daily_report_text(), parse_mode="Markdown")

    # 10. DOĞAL SOHBET
    else:
        reply_msg = parsed.get("reply") or "Buradayım hocam! Satış, rezervasyon, iptal, Drive taraması, fotoğraf veya çeki listesi ne lazımsa yazman yeterli. ⚡"
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
    # Başlangıçta tüm koleksiyon klasörlerini ve içlerindeki Excel'leri arka planda indeksle
    threading.Thread(target=build_deep_drive_index_and_inventory, daemon=True).start()

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
