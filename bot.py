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
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "AQ.Ab8RN6L4udr1eup7XD4FTqux1c9TEK1xlEtAvEuQ5QP0Y7f-Nw")

GOOGLE_DRIVE_LINK = os.getenv(
    "GOOGLE_DRIVE_LINK",
    "https://drive.google.com/drive/folders/1bnZH4cy2zOdZWhvuN79DURjabDJehOjs"
)

genai.configure(api_key=GEMINI_API_KEY)
# Gemini 1.5 Flash modeli ile sınırsız ve akıllı doğal dil işleme
ai_model = genai.GenerativeModel("gemini-1.5-flash")

PERSISTENT_EXCEL_PATH = "./master_stok_arslanlar.xlsx"
PERSISTENT_PRICE_PATH = "./master_fiyat_arslanlar.xlsx"

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


def smart_ai_analyze(user_text: str, df_summary: str) -> dict:
    """Gemini yapay zekasını kullanarak kullanıcının niyetini ve sorduğu lot/kriteri nokta atışı çözer."""
    prompt = f"""
    Sen Arslanlar Mermer fabrika yönetim botusun. Kullanıcı sana Telegram'dan bir mesaj yazdı.
    Hafızadaki Excel verisinin özeti: {df_summary}
    
    Kullanıcının mesajı: "{user_text}"
    
    Bu mesaja göre kullanıcının ne yapmak istediğini analiz et ve sadece şu JSON formatında yanıt ver (başka hiçbir şey yazma):
    {{
      "intent": "packing_list" (çeki listesi, stok listesi, serbest plakalar isteniyorsa), "total_stock" (genel stok, toplam m2 soruluyorsa), "stone_query" (belli bir taş cinsi soruluyorsa), "lot_detail" (belli bir lot detayı soruluyorsa), "chat" (genel sohbet veya diğer durumlar),
      "lot": "Bulunduğu lot kodu (örn: A-283, A-371) yoksa null",
      "filter_type": "stok" (stokta/serbest olanlar), "yuklenen" (yüklenenler), "rezerv" (rezervliler), "all" (tümü),
      "stone_name": "Taş cinsi adı (örn: silver roots, violet breccia) yoksa null",
      "reply_text": "Kullanıcıya verilecek profesyonel ve net Türkçe yanıt"
    }}
    """
    try:
        response = ai_model.generate_content(prompt)
        text_resp = response.text.strip()
        # JSON temizleme
        cleaned = re.sub(r"^```json\s*|\s*```$", "", text_resp, flags=re.MULTILINE).strip()
        return json.loads(cleaned)
    except Exception as e:
        logger.error(f"AI Analiz Hatası: {e}")
        # Hata durumunda akıllı varsayılan
        return {
            "intent": "total_stock",
            "lot": None,
            "filter_type": "all",
            "stone_name": None,
            "reply_text": "Emredersiniz patron, hemen inceliyorum! ⚡"
        }


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat:
        active_group_chats.add(update.effective_chat.id)

    doc = update.message.document
    file_name_lower = doc.file_name.lower()
    
    if not file_name_lower.endswith((".xlsx", ".xls", ".csv")):
        await update.message.reply_text("Lütfen geçerli bir Excel dosyası gönderin.")
        return

    file = await doc.get_file()
    local_path = f"./{doc.file_name}"
    await file.download_to_drive(local_path)

    try:
        df = pd.read_csv(local_path) if file_name_lower.endswith(".csv") else pd.read_excel(local_path, engine="openpyxl")
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
                f"Eyvallah Mehmet hocam! `{doc.file_name}` başarıyla işlendi. 📋\n"
                f"✅ **Toplam Plaka:** {total_rows} Adet | **Toplam Metraj:** {round(total_m2, 2)} m²\n"
                f"Yapay zeka motoru ve kalıcı hafıza aktif! Artık dilediğin her şeyi sorabilirsin. ⚡",
                parse_mode="Markdown"
            )
        else:
            await update.message.reply_text("Dosya okundu ancak geçerli veri bulunamadı.")
    except Exception as e:
        await update.message.reply_text(f"Dosya işlenirken hata oluştu: {e}")


def generate_styled_excel(filtered_df: pd.DataFrame, file_name: str) -> str:
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
    df_local = get_master_df()
    
    df_summary = f"Toplam {len(df_local)} satır plaka verisi mevcut." if not df_local.empty else "Henüz veri yüklenmedi."
    
    # Yapay zeka ile mesajı analiz et
    analysis = smart_ai_analyze(text, df_summary)
    
    intent = analysis.get("intent", "total_stock")
    lot = analysis.get("lot")
    filter_type = analysis.get("filter_type", "all")
    stone_name = analysis.get("stone_name")
    reply_text = analysis.get("reply_text", "Emredersiniz patron! ⚡")

    await update.message.reply_text(reply_text, parse_mode="Markdown")

    if intent == "total_stock" or df_local.empty:
        if df_local.empty:
            await update.message.reply_text("⚠️ Henüz hafızaya bir stok Excel dosyası yüklemediniz!", parse_mode="Markdown")
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
            f"• **Yüklenen / Sevk Edilen Miktar:** {round(yuklenen_m2, 2)} m²",
            parse_mode="Markdown"
        )

    elif intent == "stone_query" and stone_name:
        selection_cols = [c for c in df_local.columns if any(k in c.lower() for k in ["seleksiyon", "cins", "taş"])]
        match_df = df_local
        if selection_cols:
            mask = df_local[selection_cols[0]].astype(str).str.lower().str.contains(stone_name, na=False)
            match_df = df_local[mask]
            
        total_m2 = 0.0
        m2_col = next((c for c in match_df.columns if any(k in c.lower() for k in ["m2", "metraj", "alan"])), None)
        if m2_col:
            total_m2 = float(pd.to_numeric(match_df[m2_col], errors="coerce").sum())
            
        await update.message.reply_text(
            f"🪨 **{stone_name.title()} Stok Analizi:**\n"
            f"• **Toplam Plaka:** {len(match_df)} Adet\n"
            f"• **Toplam Metraj:** {round(total_m2, 2)} m²\n\n"
            f"🔗 [Drive'da İlgili Klasörü Aç]({GOOGLE_DRIVE_LINK})",
            parse_mode="Markdown"
        )

    elif intent == "packing_list" or intent == "lot_detail":
        file_path, slabs, m2 = generate_filtered_packing_list(lot, filter_type)
        if slabs > 0 and file_path:
            with open(file_path, "rb") as f:
                await update.message.reply_document(
                    document=f, 
                    caption=f"✅ **Kurumsal ve Estetik Çeki Listesi Hazırlandı!** (Lot: {lot or 'Genel'} | {filter_type.upper()})\n📦 Plaka Adedi: {slabs} | 📐 Toplam Metraj: {round(m2, 2)} m²\n🔗 [Drive'da İlgili Klasörü Aç]({GOOGLE_DRIVE_LINK})", 
                    parse_mode="Markdown"
                )
        else:
            await update.message.reply_text("⚠️ Seçilen kriterlere uygun plaka bulunamadı.", parse_mode="Markdown")


def main():
    threading.Thread(target=lambda: HTTPServer(("0.0.0.0", int(os.environ.get("PORT", 8080))), DummyHandler).serve_forever(), daemon=True).start()

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
