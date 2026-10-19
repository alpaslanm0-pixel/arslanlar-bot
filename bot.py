import os
import telebot
from telebot.types import Message
import pandas as pd

# Telegram Bot Token'ınızı buraya girin
TOKEN = "BURAYA_TELEGRAM_BOT_TOKEN_GIRINIZ"
bot = telebot.TeleBot(TOKEN)

# Bellekte tutulacak veri yolları (Geçici veya kalıcı veritabanı / Excel yolları)
EXCEL_PATH = "arslanlar_stok.xlsx"  # Ana stok Excel dosyanız
PRICE_LIST_PATH = "price_list.xlsx" # Telegram'dan güncelleyeceğiniz fiyat listesi

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message: Message):
    welcome_text = (
        "👋 **Arslanlar Marble Akıllı Asistanına Hoş Geldiniz!**\n\n"
        "Komut Örnekleri:\n"
        "🔹 `A-283 stok` -> Sadece stokta olan plakaları estetik şablonda listeler.\n"
        "🔹 `A-283 tüm liste` -> Yüklenen ve stok tüm plakaları gösterir.\n"
        "🔹 `A-283 yüklenenler` -> Yüklenmiş plakaları listeler.\n"
        "🔹 `A-283 drive` -> İlgili lotun özel Drive klasör linkini verir.\n"
        "🔹 Fiyat listesi güncellemek için dosyayı doğrudan bota 'price_list' başlığıyla gönderebilirsiniz."
    )
    bot.reply_to(message, welcome_text, parse_mode="Markdown")

@bot.message_handler(content_types=['document'])
def handle_docs(message: Message):
    """Telegram üzerinden gönderilen güncel Price List veya Stok dosyalarını yakalar."""
    try:
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        
        file_name = message.document.file_name.lower()
        if "price" * 2 in file_name or "fiyat" in file_name or "list" in file_name:
            save_path = PRICE_LIST_PATH
            doc_type = "Fiyat Listesi (Price List)"
        else:
            save_path = EXCEL_PATH
            doc_type = "Stok / Çeki Listesi"

        with open(save_path, 'wb') as new_file:
            new_file.write(downloaded_file)
            
        bot.reply_to(message, f"✅ Başarılı! Güncel **{doc_type}** sisteme kaydedildi ve işlem için aktif edildi.")
    except Exception as e:
        bot.reply_to(message, f"❌ Dosya kaydedilirken hata oluştu: {str(e)}")

@bot.message_handler(func=lambda message: True)
def handle_query(message: Message):
    text = message.text.lower()
    
    # Örnek olarak A-283 sorgusu işleniyor
    if "a-283" in text or "violet" in text:
        try:
            if not os.path.exists(EXCEL_PATH):
                bot.reply_to(message, "⚠️ Sistemde aktif stok Excel dosyası bulunamadı. Lütfen önce Excel yükleyin.")
                return

            df = pd.read_excel(EXCEL_PATH)
            # Sütun adlarını normalize etme
            df.columns = [str(c).strip().upper() for c in df.columns]
            
            # A-283 filtresi
            lot_df = df[df['KOD'].astype(str).str.contains("A-283", na=False)]
            
            if lot_df.empty:
                bot.reply_to(message, "🔍 Aradığınız kriterlere uygun A-283 kaydı bulunamadı.")
                return

            # Filtreleme Mantığı (Stok mu, Yüklenen mi, Tümü mü?)
            filter_type = "stok" # varsayılan
            if "tüm" in text or "hepsi" in text:
                filter_type = "all"
            elif "yüklenen" in text or "yuklenen" in text:
                filter_type = "loaded"
            elif "stok" in text:
                filter_type = "stock"

            # Durum sütununa göre filtreleme (Örn: 'STOK DURUMU' veya 'STOK')
            stok_col = next((col for col in lot_df.columns if 'STOK' in col and 'DURUM' in col), 'STOK')
            
            if filter_type == "stock":
                filtered_df = lot_df[lot_df[stok_col].astype(str).str.contains("STOK", case=False, na=False)]
                title_suffix = "🟢 STOKTA OLAN PLAKALAR"
            elif filter_type == "loaded":
                filtered_df = lot_df[lot_df[stok_col].astype(str).str.contains("YÜKLEN", case=False, na=False)]
                title_suffix = "🚚 YÜKLENMİŞ PLAKALAR"
            else:
                filtered_df = lot_df
                title_suffix = "📋 TÜM ÇEKİ LİSTESİ"

            # Drive Linki İsteği Kontrolü
            if "drive" in text or "link" in text:
                # Spesifik A-283 Drive klasör linki
                drive_link = "https://drive.google.com/drive/folders/ORNEK_A283_OZEF_KLASOR_ID"
                bot.reply_to(message, f"🔗 **Violet Breccia (3 A-283)** Özel Drive Klasör Linki:\n{drive_link}", parse_mode="Markdown")
                return

            # Fiyat Okuma (Telegram üzerinden yüklenen price_list.xlsx'ten)
            unit_price = 0.0
            if os.path.exists(PRICE_LIST_PATH):
                try:
                    price_df = pd.read_excel(PRICE_LIST_PATH)
                    # Fiyat eşleştirme mantığı (Violet Breccia için)
                    match_price = price_df[price_df.astype(str).apply(lambda x: x.str.contains("Violet Breccia|A-283", case=False)).any(axis=1)]
                    if not match_price.empty:
                        # Fiyat sütununu bulup ilk eşleşeni alalım
                        unit_price = float(match_price.iloc[0].filter(like='FİYAT').values[0]) if 'FİYAT' in match_price.columns else 259.87 # Varsayılan test verisi
                except Exception:
                    unit_price = 259.87 # Hata durumunda varsayılan birim fiyat
            else:
                unit_price = 259.87

            # Estetik Şablon Oluşturma (Markdown Monospace blok içinde tablo)
            response_lines = [
                f"💎 **ARSLANLAR MARBLE - ÇEKİ LİSTESİ**",
                f"📌 **Cinsi:** Violet Breccia ({title_suffix})",
                f"📊 **Toplam Kalem:** {len(filtered_df)} Adet Plaka",
                f"💵 **Birim Fiyat (Price List):** ${unit_price} / M2",
                "----------------------------------------",
                "```text",
                f"{'Plaka':<6} {'Boy':<6} {'Yük':<6} {'Kal':<4} {'M2':<6} {'Durum'}"
            ]
            response_lines.append("-" * 40)

            for _, row in filtered_df.iterrows():
                plaka_no = str(row.get('PLAKA NO', row.get('PLAKA', '-'))).strip()
                uzunluk = str(row.get('UZUNLUK', '-')).strip()
                yukseklik = str(row.get('YÜKSEKLİK', '-')).strip()
                kalinlik = str(row.get('KALINLIK', '-')).strip()
                m2 = str(row.get('M2', '-')).strip()
                durum = str(row.get(stok_col, '-')).strip()
                
                response_lines.append(f"{plaka_no:<6} {uzunluk:<6} {yukseklik:<6} {kalinlik:<4} {m2:<6} {durum}")

            response_lines.append("```")
            response_lines.append("✨ *Rapor Arslanlar Marble Operasyon Sistemi tarafından oluşturulmuştur.*")

            bot.reply_to(message, "\n".join(response_lines), parse_mode="Markdown")

        except Exception as e:
            bot.reply_to(message, f"❌ İşlem sırasında bir hata oluştu: {str(e)}")
    else:
        bot.reply_to(message, "🤖 Mesajınız anlaşıldı. Stok veya çeki listesi sorgulamak için lütfen lot kodu (Örn: A-283) belirtin.")

if __name__ == "__main__":
    print("Bot çalıştırılıyor...")
    bot.infinity_polling()
