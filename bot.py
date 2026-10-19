import os
import re
import io
import json
import logging
import threading
import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
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
    "https://drive.google.com/drive/folders/1bnZH4cy2zOdZWhvuN79DURjabDJehOjs"
)

genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-1.5-flash")

# --- KALICI DİSK HAFIZA DOSYASI ---
PERSISTENT_EXCEL_PATH = "./master_stok_arslanlar.xlsx"

bot_lock = threading.Lock()
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


def get_master_df() -> pd.DataFrame:
    with bot_lock:
        if os.path.exists(PERSISTENT_EXCEL_PATH):
            try:
                df = pd.read_excel(PERSISTENT_EXCEL_PATH, engine="openpyxl")
                df.columns = [str(c).strip() for c in df.columns]
                return df.dropna(how="all")
            except Exception as e:
                logger.error(f"Kalıcı dosya okuma hatası: {e}")
        return pd.DataFrame()


def save_master_df(df: pd.DataFrame):
    with bot_lock:
        df.to_excel(PERSISTENT_EXCEL_PATH, index=False, engine='openpyxl')


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
        return len(drive_focused_index), len(get_master_df())

    is_scanning_drive = True
    all_discovered = []
    root_items = fetch_single_folder(GOOGLE_DRIVE_LINK, "Ana Klasör")
    for item in root_items:
        all_discovered.append(item)
        if item["is_folder"]:
            sub_items = fetch_single_folder(item["url"], item["name"])
            for sub in sub_items:
                all_discovered.append(sub)

    if all_discovered:
        drive_focused_index = all_discovered

    is_scanning_drive = False
    return len(drive_focused_index), len(get_master_df())


def analyze_message(text: str) -> dict:
    t_lower = text.lower()
    lot_match = re.search(r"\b([a-zA-Z]{1,3}[- ]?\d+|\d+[a-zA-Z]?)\b", text)
    lot = lot_match.group(1).upper().replace(" ", "-") if lot_match else None

    if any(k in t_lower for k in ["kaç m2", "ne kadar", "toplam m2", "stokta ne var", "stok", "m²", "metraj"]):
        if lot or any(k in t_lower for k in ["stok listesi", "serbest", "müsait"]):
            filter_val = "stok" if "stok" in t_lower or "serbest" in t_lower else "all"
            return {"action": "packing_list", "lot": lot, "filter": filter_val, "reply": f"Stok ve çeki listesini hazırlıyorum patron! 📋"}
        return {"action": "total_stock_query", "lot": lot, "filter": "all", "reply": "Tüm fabrika stok metraj ve plaka analizini çıkarıyorum! 📊"}
    elif any(k in t_lower for k in ["yüklenecek", "sevkiyat", "yükleme"]):
        return {"action": "shipment_plan", "lot": lot, "filter": "all", "reply": f"Sevkiyat planını işledim patron! 🚛"}
    elif any(k in t_lower for k in ["yüklenen", "sevk"]):
        return {"action": "packing_list", "lot": lot, "filter": "yuklenen", "reply": f"Yüklenen plakaları hazırlıyorum! 🚛"}
    elif any(k in t_lower for k in ["rezerv"]):
        return {"action": "packing_list", "lot": lot, "filter": "rezerv", "reply": f"Rezervli plakaları hazırlıyorum! 🔒"}
    elif any(k in t_lower for k in ["çeki", "packing", "liste"]):
        return {"action": "packing_list", "lot": lot, "filter": "all", "reply": "Profesyonel çeki listesini hazırlıyorum! 📄"}
    elif any(k in t_lower for k in ["link", "drive", "gönder"]):
        return {"action": "lot_link", "lot": lot, "filter": "all", "reply": f"Nokta atışı Drive bağlantısını getiriyorum! 🔗"}
    elif lot:
        return {"action": "lot_detail", "lot": lot, "filter": "all", "reply": f"Lot {lot} detaylarını inceliyorum! 🔍"}
    elif any(k in t_lower for k in ["rapor", "özet"]):
        return {"action": "daily_report", "lot": None, "filter": "all", "reply": "Günlük operasyon raporunu getiriyorum! 📊"}
    else:
        return {"action": "total_stock_query", "lot": None, "filter": "all", "reply": "Genel stok durumunu ve metrajları analiz ediyorum! 📊"}


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat:
        active_group_chats.add(update.effective_chat.id)

    doc = update.message.document
    file_name_lower = doc.file_name.lower()
    
    if not file_name_lower.endswith((".xlsx", ".xls", ".csv")):
        await update.message.reply_text("Lütfen geçerli bir Excel veya CSV dosyası gönderin.")
        return

    file = await doc.get_file()
    local_path = f"./{doc.file_name}"
    await file.download_to_drive(local_path)

    try:
        if file_name_lower.endswith(".csv"):
            df = pd.read_csv(local_path)
        else:
            df = pd.read_excel(local_path, engine="openpyxl")

        if not df.empty:
            df.columns = [str(c).strip() for c in df.columns]
            df = df.dropna(how="all")
            save_master_df(df)
            
            total_rows = len(df)
            total_m2 = 0.0
            for col in df.columns:
                if any(k in col.lower() for k in ["m2", "metraj", "alan"]):
                    total_m2 = float(pd.to_numeric(df[col], errors="coerce").sum())
                    break
                    
            await update.message.reply_text(
                f"Eyvallah Mehmet hocam! `{doc.file_name}` başarıyla kaydedildi. 📋\n"
                f"✅ **Toplam Plaka:** {total_rows} Adet | **Toplam Metraj:** {round(total_m2, 2)} m²\n"
                f"Kalıcı hafıza güncellendi ve ekibe hazır! ⚡",
                parse_mode="Markdown"
            )
        else:
            await update.message.reply_text("Dosya okundu ancak geçerli veri bulunamadı.")
    except Exception as e:
        await update.message.reply_text(f"Dosya işlenirken hata oluştu: {e}")


def generate_styled_excel(filtered_df: pd.DataFrame, file_name: str) -> str:
    """Gönderdiğin fabrika formatına uygun, estetik ve profesyonel kurumsal tablo tasarımı."""
    filtered_df.to_excel(file_name, index=False, engine='openpyxl')
    wb = openpyxl.load_workbook(file_name)
    ws = wb.active
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Arial", size=10)
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )
    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')

    for col_num in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        cell.border = thin_border

    for row_num in range(2, ws.max_row + 1):
        for col_num in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_num, column=col_num)
            cell.font = data_font
            cell.border = thin_border
            cell.alignment = align_center if col_num > 1 else align_left

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(file_name)
    return file_name


def generate_filtered_packing_list(lot_code: str = None, filter_type: str = "all") -> tuple:
    df = get_master_df()
    if df.empty:
        return "", 0, 0.0
        
    filtered = df.copy()
    if lot_code:
        matching_cols = [c for c in df.columns if any(k in c.lower() for k in ["blok", "lot", "fabrika", "kod"])]
        if matching_cols:
            mask = False
            for col in matching_cols:
                mask = mask | df[col].astype(str).apply(normalize_key).str.contains(normalize_key(lot_code), na=False)
            temp_filtered = df[mask]
            if not temp_filtered.empty:
                filtered = temp_filtered

    if filter_type != "all":
        status_cols = [c for c in filtered.columns if any(k in c.lower() for k in ["durum", "stok"])]
        res_cols = [c for c in filtered.columns if any(k in c.lower() for k in ["rezerv", "satiş", "satis"])]
        
        sub_filtered = filtered
        if filter_type == "stok" and status_cols:
            sub_filtered = filtered[filtered[status_cols[0]].astype(str).str.lower().str.contains("stok", na=False)]
        elif filter_type == "yuklenen" and status_cols:
            sub_filtered = filtered[filtered[status_cols[0]].astype(str).str.lower().str.contains("yüklen|yuklen", na=False)]
        elif filter_type == "rezerv" and res_cols:
            sub_filtered = filtered[filtered[res_cols[0]].astype(str).str.lower().str.contains("rezerv", na=False)]
        
        if not sub_filtered.empty:
            filtered = sub_filtered

    if filtered.empty:
        filtered = df

    safe_tag = f"{lot_code or 'Genel'}_{filter_type.upper()}"
    file_name = f"Arslanlar_Marble_Ceki_Listesi_{safe_tag}.xlsx"
    generate_styled_excel(filtered, file_name)
    
    total_slabs = len(filtered)
    total_m2 = 0.0
    for col in filtered.columns:
        if any(k in col.lower() for k in ["m2", "metraj", "alan"]):
            total_m2 = float(pd.to_numeric(filtered[col], errors="coerce").sum())
            break
            
    return file_name, total_slabs, total_m2


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat:
        active_group_chats.add(update.effective_chat.id)

    text = update.message.text
    parsed = analyze_message(text)
    action = parsed.get("action", "chat")
    lot = parsed.get("lot")
    filter_type = parsed.get("filter", "all")
    natural_reply = parsed.get("reply", "Emredersiniz patron! ⚡")

    await update.message.reply_text(natural_reply, parse_mode="Markdown")

    if action == "total_stock_query":
        df_local = get_master_df()
        if df_local.empty:
            await update.message.reply_text("⚠️ Henüz hafızaya bir stok Excel dosyası yüklemediniz! Lütfen önce dosya gönderin.", parse_mode="Markdown")
            return
            
        total_rows = len(df_local)
        total_m2 = 0.0
        m2_col = next((c for c in df_local.columns if any(k in c.lower() for k in ["m2", "metraj", "alan"])), None)
        if m2_col:
            total_m2 = float(pd.to_numeric(df_local[m2_col], errors="coerce").sum())
                
        stok_m2 = 0.0
        yuklenen_m2 = 0.0
        status_cols = [c for c in df_local.columns if any(k in c.lower() for k in ["durum", "stok"])]
        
        if status_cols and m2_col:
            for idx, row in df_local.iterrows():
                val = str(row[status_cols[0]]).lower()
                m2_val = pd.to_numeric(row[m2_col], errors="coerce") or 0.0
                if "stok" in val:
                    stok_m2 += m2_val
                elif "yüklen" in val or "yuklen" in val:
                    yuklenen_m2 += m2_val

        await update.message.reply_text(
            f"📊 **ARSLANLAR MARBLE - GENEL STOK ANALİZİ**\n\n"
            f"• **Toplam Plaka Sayısı:** {total_rows} Adet\n"
            f"• **Toplam Metraj:** {round(total_m2, 2)} m²\n"
            f"• **Depodaki / Stoktaki Miktar:** {round(stok_m2, 2)} m²\n"
            f"• **Yüklenen / Sevk Edilen Miktar:** {round(yuklenen_m2, 2)} m²\n\n"
            f"İstediğiniz özel bir lot veya filtre için 'A-283 stok listesi' yazabilirsiniz.",
            parse_mode="Markdown"
        )

    elif action == "lot_link" and lot:
        target_url = GOOGLE_DRIVE_LINK
        for item in drive_focused_index:
            if lot and normalize_key(lot) in normalize_key(item["name"]):
                target_url = item["url"]
                break
        await update.message.reply_text(
            f"🔗 **Lot {lot or 'Genel'} İçin Nokta Atışı Bağlantı:**\n"
            f"Doğrudan ilgili klasör/dosya bağlantısı: [Drive'da Aç]({target_url})",
            parse_mode="Markdown"
        )

    elif action == "lot_detail" and lot:
        df_local = get_master_df()
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
            stok_count = 0
            yuklenen_count = 0
            total_m2 = 0.0
            
            status_cols = [c for c in match.columns if any(k in c.lower() for k in ["durum", "stok"])]
            for idx, row in match.iterrows():
                if status_cols:
                    val = str(row[status_cols[0]]).lower()
                    if "stok" in val:
                        stok_count += 1
                    elif "yüklen" in val or "yuklen" in val:
                        yuklenen_count += 1
            for col in match.columns:
                if any(k in col.lower() for k in ["m2", "metraj", "alan"]):
                    total_m2 = float(pd.to_numeric(match[col], errors="coerce").sum())
                    break

            target_url = GOOGLE_DRIVE_LINK
            for item in drive_focused_index:
                if normalize_key(lot) in normalize_key(item["name"]):
                    target_url = item["url"]
                    break

            await update.message.reply_text(
                f"🪨 **Lot / Blok {lot} Detaylı Döküm:**\n"
                f"• Toplam Plaka: *{slab_count} Adet*\n"
                f"• Stokta / Serbest: *{stok_count} Adet*\n"
                f"• Yüklenen: *{yuklenen_count} Adet*\n"
                f"• Toplam Metraj: *{round(total_m2, 2)} m²*\n\n"
                f"🔗 [Drive'da İlgili Klasörü Aç]({target_url})\n\n"
                f"Çeki listesi almak için 'A-283 stok listesi' yazabilirsiniz.",
                parse_mode="Markdown"
            )
        else:
            await update.message.reply_text(f"⚠️ `{lot}` kodlu lot hafızada taranıyor.", parse_mode="Markdown")

    elif action == "packing_list":
        file_path, slabs, m2 = generate_filtered_packing_list(lot, filter_type)
        if slabs > 0 and file_path:
            with open(file_path, "rb") as f:
                await update.message.reply_document(
                    document=f, 
                    caption=f"✅ **Kurumsal ve Estetik Çeki Listesi Hazırlandı!** ({filter_type.upper()})\n📦 Plaka Adedi: {slabs} | 📐 Toplam Metraj: {round(m2, 2)} m²\n📎 *Gönderdiğiniz fabrika formatıyla birebir ve estetik şablonda üretilmiştir!*", 
                    parse_mode="Markdown"
                )
        else:
            await update.message.reply_text("⚠️ Seçilen kriterlere uygun plaka bulunamadı veya henüz Excel dosyası yüklenmedi.", parse_mode="Markdown")

    elif action == "daily_report":
        df_local = get_master_df()
        await update.message.reply_text(
            f"📊 **ARSLANLAR MARBLE - GÜNLÜK OPERASYON RAPORU**\n\n"
            f"📦 **Hafızadaki Plaka Kaydı:** {len(df_local)} Satır.\n"
            f"🔗 [Google Drive Klasörü]({GOOGLE_DRIVE_LINK})\n\n"
            f"Her şey kontrol altında, iyi çalışmalar! ☕💪",
            parse_mode="Markdown"
        )


def main():
    threading.Thread(target=lambda: HTTPServer(("0.0.0.0", int(os.environ.get("PORT", 8080))), DummyHandler).serve_forever(), daemon=True).start()
    threading.Thread(target=lambda: scan_marble_collections(), daemon=True).start()

    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))

    logger.info("Arslanlar Mermer Fabrika Asistanı başlatılıyor...")
    app.run_polling(drop_pending_updates=True, allowed_updates=Update.ALL_TYPES)


class DummyHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Arslanlar Marble Bot 7/24 Aktif!")


if __name__ == "__main__":
    main()
