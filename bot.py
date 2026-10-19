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

# --- SİSTEM YAPILANDIRMASI ---
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "8822872793:AAEn0p_FLlH1CaB41cHhwHV7ADh7q6o6VGU")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "AQ.Ab8RN6L4udr1eup7XD4FTqux1c9TEK1xlEtAvEuQ5QP0Y7f-Nw").strip()

GOOGLE_DRIVE_LINK = os.getenv(
    "GOOGLE_DRIVE_LINK",
    "https://drive.google.com/drive/folders/1bnZH4cy2zOdZWhvuN79DURjabDJehOjs"
)

genai.configure(api_key=GEMINI_API_KEY)
# Gemini Flash modeli ile akıllı dil işleme ve veri analizi
ai_model = genai.GenerativeModel("gemini-1.5-flash")

PERSISTENT_EXCEL_PATH = "./master_stok_arslanlar.xlsx"

bot_lock = threading.Lock()
active_group_chats = set()


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
                logger.error(f"Stok okuma hatası: {e}")
        return pd.DataFrame()


def save_master_df(df: pd.DataFrame):
    with bot_lock:
        df.to_excel(PERSISTENT_EXCEL_PATH, index=False, engine='openpyxl')


def generate_styled_excel(filtered_df: pd.DataFrame, file_name: str, source_file_path: str = PERSISTENT_EXCEL_PATH) -> str:
    filtered_df.to_excel(file_name, index=False, engine='openpyxl')
    
    try:
        wb_source = openpyxl.load_workbook(source_file_path, data_only=False)
        ws_source = wb_source.active
        wb_target = openpyxl.load_workbook(file_name)
        ws_target = wb_target.active
        
        for row in ws_source.iter_rows():
            for cell in row:
                if cell.hyperlink:
                    cell_value = str(cell.value)
                    for target_row in ws_target.iter_rows():
                        for target_cell in target_row:
                            if str(target_cell.value) == cell_value and cell_value != "None":
                                target_cell.hyperlink = cell.hyperlink
                                target_cell.font = Font(name="Arial", size=10, color="0563C1", underline="single")
        wb_target.save(file_name)
    except Exception as e:
        logger.warning(f"Köprü aktarım uyarısı: {e}")

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
            if not cell.hyperlink:
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


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat:
        active_group_chats.add(update.effective_chat.id)

    doc = update.message.document
    file_name_lower = doc.file_name.lower()
    
    if not file_name_lower.endswith((".xlsx", ".xls", ".csv", ".xlsb")):
        await update.message.reply_text("Lütfen geçerli bir Excel dosyası gönderin.")
        return

    file = await doc.get_file()
    local_path = f"./{doc.file_name}"
    await file.download_to_drive(local_path)

    try:
        if file_name_lower.endswith(".csv"):
            df = pd.read_csv(local_path)
        elif file_name_lower.endswith(".xlsb"):
            df = pd.read_excel(local_path, engine="pyxlsb")
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
                f"✅ `{doc.file_name}` başarıyla analiz motoruna yüklendi.\n"
                f"• **Toplam Plaka:** {total_rows} Adet\n"
                f"• **Toplam Metraj:** {round(total_m2, 2)} m²\n"
                f"Yapay zeka asistanı tam kapasite aktif!",
                parse_mode="Markdown"
            )
        else:
            await update.message.reply_text("Dosya okundu ancak veri bulunamadı.")
    except Exception as e:
        await update.message.reply_text(f"Dosya yükleme hatası: {e}")


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat:
        active_group_chats.add(update.effective_chat.id)

    text = update.message.text
    t_lower = text.lower()
    df_local = get_master_df()

    # 1. Çeki Listesi / Excel Dosyası Talebi
    if any(k in t_lower for k in ["çeki listesi", "packing list", "excel ver", "listesini gönder", "stok listesi"]):
        lot_match = re.search(r"\b([a-zA-Z]{1,3}[- ]?\d+|\d+[a-zA-Z]?)\b", text)
        lot = lot_match.group(1).upper().replace(" ", "-") if lot_match else None
        
        filter_type = "stok" if "stok" in t_lower or "serbest" in t_lower else ("yuklenen" if "yüklen" in t_lower else "all")
        
        filtered = df_local.copy()
        if lot and not df_local.empty:
            matching_cols = [c for c in df_local.columns if any(k in c.lower() for k in ["blok", "lot", "fabrika", "kod"])]
            if matching_cols:
                mask = False
                for col in matching_cols:
                    mask = mask | df_local[col].astype(str).apply(normalize_key).str.contains(normalize_key(lot), na=False)
                filtered = df_local[mask]

        if filtered.empty:
            filtered = df_local

        file_name = f"Arslanlar_Marble_Ceki_Listesi_{lot or 'Genel'}.xlsx"
        generate_styled_excel(filtered, file_name)
        
        total_m2 = 0.0
        m2_col = next((c for c in filtered.columns if any(k in c.lower() for k in ["m2", "metraj", "alan"])), None)
        if m2_col:
            total_m2 = float(pd.to_numeric(filtered[m2_col], errors="coerce").sum())

        with open(file_name, "rb") as f:
            await update.message.reply_document(
                document=f, 
                caption=f"✅ **Kurumsal Çeki Listesi Hazırlandı**\n📦 Adet: {len(filtered)} | 📐 Toplam: {round(total_m2, 2)} m²\n🔗 [Drive'da Aç]({GOOGLE_DRIVE_LINK})", 
                parse_mode="Markdown"
            )
        return

    # 2. Saf Python Veri Filtreleme Motoru (Silver Roots, metrajlar, stok sorguları için kusursuz eşleşme)
    stones = ["silver roots", "violet breccia", "brown armoni", "silver river", "golden spider", "galapagos", "grey philharmonic", "burno perla"]
    detected_stone = next((s for s in stones if s in t_lower), None)

    if detected_stone or any(k in t_lower for k in ["kaç m2", "ne kadar", "stokta", "üretilen", "analiz", "toplam"]):
        if not df_local.empty:
            match_df = df_local.copy()
            if detected_stone:
                mask = match_df.astype(str).apply(lambda x: x.str.lower().str.contains(detected_stone, na=False)).any(axis=1)
                match_df = match_df[mask]
            
            m2_col = next((c for c in match_df.columns if any(k in c.lower() for k in ["m2", "metraj", "alan"])), None)
            total_m2 = float(pd.to_numeric(match_df[m2_col], errors="coerce").sum()) if m2_col else 0.0
            
            stok_m2 = 0.0
            yuklenen_m2 = 0.0
            stat_cols = [c for c in match_df.columns if any(k in c.lower() for k in ["durum", "stok"])]
            if stat_cols and m2_col:
                for idx, row in match_df.iterrows():
                    val = str(row[stat_cols[0]]).lower()
                    m2_val = pd.to_numeric(row[m2_col], errors="coerce") or 0.0
                    if "stok" in val:
                        stok_m2 += m2_val
                    elif "yüklen" in val or "yuklen" in val:
                        yuklenen_m2 += m2_val

            await update.message.reply_text(
                f"📊 **ARSLANLAR MARBLE - DETAYLI ANALİZ RAPORU**\n"
                f"• **Sorgu Kriteri:** {detected_stone.title() if detected_stone else 'Genel Fabrika Stoğu'}\n"
                f"• **Toplam Plaka:** {len(match_df)} Adet\n"
                f"• **Toplam Metraj:** {round(total_m2, 2)} m²\n"
                f"• **Stoktaki Miktar:** {round(stok_m2, 2)} m²\n"
                f"• **Yüklenen Miktar:** {round(yuklenen_m2, 2)} m²\n\n"
                f"🔗 [Drive'da Aç]({GOOGLE_DRIVE_LINK})",
                parse_mode="Markdown"
            )
            return

    # 3. Gerçek Gemini AI Zekası (Sunum hazırlama, genel sohbet, strateji, mermer sektörü vb. her soruya akıllı yanıt)
    try:
        summary_ctx = f"Fabrika stok veritabanında toplam {len(df_local)} kayıtlı plaka var." if not df_local.empty else "Henüz stok dosyası yüklenmedi."
        
        prompt = f"""
        Sen Arslanlar Mermer fabrika yönetim asistanı ve üst düzey bir yapay zeka uzmansın. 
        Mevcut Sistem Durumu: {summary_ctx}
        Kullanıcının Mesajı: "{text}"
        
        Görevin: Kullanıcının sorduğu soruya (sunum hazırlama, genel strateji, mermer ihracatı, teknik bilgi veya serbest sohbet) son derece akıllı, profesyonel, kurumsal, akıcı ve detaylı bir Türkçe ile yanıt vermek. Asla kalıp sözler veya tekrara düşen robotik cümleler kullanma. Doğrudan konuya gir, bir yapay zeka uzantısı gibi en üst düzeyde yardımcı ol.
        """
        response = ai_model.generate_content(prompt)
        await update.message.reply_text(response.text.strip(), parse_mode="Markdown")
    except Exception as e:
        logger.error(f"AI Yanıt Hatası: {e}")
        total_rows = len(df_local)
        await update.message.reply_text(
            f"📊 **Fabrika Durum Özeti**\n• Hafızada kayıtlı plaka satırı: {total_rows}\n🔗 [Drive'da Aç]({GOOGLE_DRIVE_LINK})",
            parse_mode="Markdown"
        )


def main():
    threading.Thread(target=lambda: HTTPServer(("0.0.0.0", int(os.environ.get("PORT", 8080))), DummyHandler).serve_forever(), daemon=True).start()

    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))

    logger.info("Arslanlar Mermer Fabrika AI Asistanı başlatılıyor...")
    app.run_polling(drop_pending_updates=True, allowed_updates=Update.ALL_TYPES)


class DummyHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Arslanlar Marble AI Bot 7/24 Aktif!")


if __name__ == "__main__":
    main()
