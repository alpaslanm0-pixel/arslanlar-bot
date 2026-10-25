import os
import io
import re
import json
import time
import base64
import logging
import asyncio
import threading
from datetime import datetime
from collections import deque
from typing import Dict, List, Optional, Tuple, Any

import requests
import pandas as pd
import openpyxl
from flask import Flask
from PIL import Image, ImageDraw

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image as RLImage,
    KeepTogether,
    HRFlowable,
)

from telegram import Update, InputMediaPhoto, InputMediaDocument
from telegram.constants import ParseMode, ChatAction
from telegram.ext import (
    ApplicationBuilder,
    ContextTypes,
    MessageHandler,
    filters,
)

from google.oauth2 import service_account
import google.auth.transport.requests
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# =====================================================================
# 1. LOGLAMA VE SİSTEM YAPILANDIRMASI
# =====================================================================
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("ArgonMasterPerfect_v4")

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GOOGLE_SA_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()

DEFAULT_DRIVE_LINK = (
    os.getenv("GOOGLE_DRIVE_LINK")
    or os.getenv("DEFAULT_DRIVE_LINK")
    or "https://drive.google.com"
).strip()

HYPERLINKS_DB_FILE = "hyperlinks_db.json"
VISUAL_CACHE_FILE = "visual_cache.json"
STONE_DNA_DB_FILE = "stone_dna_db.json"


# =====================================================================
# 2. YARDIMCI FONKSİYONLAR VE VERİ NORMALİZASYONU
# =====================================================================
def load_json_file(filepath: str, default_data: Any) -> Any:
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logger.warning(f"{filepath} JSON okuma hatası: {e}")
    return default_data


def save_json_file(filepath: str, data: Any) -> None:
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.error(f"{filepath} JSON yazma hatası: {e}")


def normalize_turkish_key(text: Any) -> str:
    """Türkçe karakterleri ASCII'ye çevirip büyük harf yapar, boşlukları temizler."""
    if text is None or pd.isna(text):
        return ""
    s = str(text).strip()
    tr_map = str.maketrans("İıiŞşĞğÜüÖöÇç", "IIISSGGUUOOCC")
    s = s.translate(tr_map).upper()
    return re.sub(r"\s+", " ", s).strip()


def normalize_lot_code(code: Any) -> str:
    """A 371, A-371, a371 gibi formatları standart A-371'e dönüştürür."""
    raw = normalize_turkish_key(code)
    match = re.match(r"^([A-Z]+)[-\s]*(\d+)$", raw)
    return f"{match.group(1)}-{match.group(2)}" if match else raw


def safe_float(val: Any) -> float:
    """Hücrelerdeki metin/sayı karışımını güvenli bir float değerine çevirir."""
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = re.sub(r"[^\d,.-]", "", str(val).strip())
    if not s:
        return 0.0
    if "," in s and "." in s:
        s = (
            s.replace(".", "").replace(",", ".")
            if s.rfind(",") > s.rfind(".")
            else s.replace(",", "")
        )
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


# =====================================================================
# 3. GOOGLE DRIVE OTONOM GEZGİN VE GÖRSEL/DOSYA AYIKLAMA MOTORU
# =====================================================================
class DriveExplorer:
    """
    Arslanlar Mermer'in devasa Google Drive ağacında otonom olarak gezinen,
    istenen taşı/lotu bulan, içindeki Excel ve görselleri süzen kompleks motor.
    """
    def __init__(self, sa_json_str: str):
        self.service = None
        self.sa_json_str = sa_json_str
        self.hyperlinks_db = load_json_file(HYPERLINKS_DB_FILE, {})
        self.visual_cache = load_json_file(VISUAL_CACHE_FILE, {})
        self._init_service()

    def _init_service(self) -> None:
        if not self.sa_json_str:
            logger.warning("CRITICAL: GOOGLE_SERVICE_ACCOUNT_JSON eksik!")
            return
        try:
            info = json.loads(self.sa_json_str)
            creds = service_account.Credentials.from_service_account_info(
                info, scopes=["https://www.googleapis.com/auth/drive.readonly"]
            )
            self.service = build("drive", "v3", credentials=creds, cache_discovery=False)
            logger.info("Drive API Bağlantısı (Deep Scan) Aktif.")
        except Exception as e:
            logger.error(f"Drive API Başlatma Hatası: {e}")

    def find_folder(self, query_term: str) -> Tuple[Optional[str], Optional[str]]:
        """Sorgulanan Lot veya Taş adını Drive'da arar, Klasör ID ve Link döner."""
        norm_key = normalize_lot_code(query_term)
        if norm_key in self.hyperlinks_db:
            url = self.hyperlinks_db[norm_key]
            fid = self._extract_folder_id(url)
            if fid:
                return fid, url

        if not self.service:
            return None, None

        # Olası klasör adı varyasyonları
        cands = [
            query_term.strip().replace("'", ""),
            norm_key,
            norm_key.replace("-", ""),
            norm_key.replace("-", " "),
        ]
        # Türkçe varyasyonlar (Örn: SİLVER RİVER)
        tr_var = norm_key.replace("SILVER", "SİLVER").replace("RIVER", "RİVER")
        if tr_var != norm_key:
            cands.append(tr_var)
            
        cands = list(dict.fromkeys([c for c in cands if len(c) > 2]))

        for cand in cands:
            try:
                q = (
                    f"mimeType = 'application/vnd.google-apps.folder' "
                    f"and trashed = false and name contains '{cand}'"
                )
                res = self.service.files().list(q=q, spaces="drive", fields="files(id, name, webViewLink)", pageSize=5).execute()
                files = res.get("files", [])
                if files:
                    for f in files:
                        # Kesin eşleşme önceliği
                        if normalize_turkish_key(f["name"]) == normalize_turkish_key(cand):
                            fid = f["id"]
                            link = f.get("webViewLink", f"https://drive.google.com/drive/folders/{fid}")
                            self.hyperlinks_db[norm_key] = link
                            save_json_file(HYPERLINKS_DB_FILE, self.hyperlinks_db)
                            return fid, link
                    # Kesin eşleşme yoksa ilk bulduğunu al
                    fid = files[0]["id"]
                    link = files[0].get("webViewLink", f"https://drive.google.com/drive/folders/{fid}")
                    self.hyperlinks_db[norm_key] = link
                    save_json_file(HYPERLINKS_DB_FILE, self.hyperlinks_db)
                    return fid, link
            except Exception as e:
                logger.error(f"Drive klasör arama hatası ({cand}): {e}")
                continue
        return None, None

    def list_child_folders(self, parent_folder_id: str) -> List[Dict[str, str]]:
        """Deep Scan ve alt klasör taramaları için klasörleri listeler."""
        if not self.service or not parent_folder_id:
            return []
        try:
            q = (
                f"'{parent_folder_id}' in parents and trashed = false "
                f"and mimeType = 'application/vnd.google-apps.folder'"
            )
            res = self.service.files().list(q=q, spaces="drive", fields="files(id, name)", pageSize=100).execute()
            return [{"id": f["id"], "name": f["name"]} for f in res.get("files", [])]
        except Exception:
            return []

    def get_excel_from_folder(self, folder_id: str) -> Tuple[Optional[bytes], Optional[str]]:
        """Klasörün (veya alt klasörlerinin) içindeki ilk orijinal Excel dosyasını bulur ve indirir."""
        if not self.service or not folder_id:
            return None, None
        try:
            # Önce ana klasöre bak
            q = (
                f"'{folder_id}' in parents and trashed = false and "
                f"(mimeType contains 'spreadsheet' or mimeType contains 'excel' or name contains '.xls')"
            )
            res = self.service.files().list(q=q, spaces="drive", fields="files(id, name)", pageSize=2).execute()
            files = res.get("files", [])
            
            # Eğer ana klasörde yoksa, 1 seviye alt klasörlere bak (Örn: A-371 -> Packing Lists alt klasörü)
            if not files:
                sub_folders = self.list_child_folders(folder_id)[:5]
                for sf in sub_folders:
                    sq = (
                        f"'{sf['id']}' in parents and trashed = false and "
                        f"(mimeType contains 'spreadsheet' or mimeType contains 'excel' or name contains '.xls')"
                    )
                    sres = self.service.files().list(q=sq, spaces="drive", fields="files(id, name)", pageSize=2).execute()
                    files.extend(sres.get("files", []))
                    if files:
                        break

            if not files:
                return None, None

            file_meta = files[0]
            req = self.service.files().get_media(fileId=file_meta["id"])
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, req)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            fh.seek(0)
            return fh.read(), file_meta["name"]
        except Exception as e:
            logger.error(f"Excel indirme hatası (Klasör: {folder_id}): {e}")
            return None, None

    def get_images_structured(self, folder_id: str, query_term: str) -> Tuple[Optional[bytes], Optional[bytes]]:
        """
        Telegram Flash-Kart için:
        Resim 1: Standart Plaka (Bookmatch olmayan)
        Resim 2: Bookmatch (Varsa) veya 2. Plaka.
        """
        all_imgs = self._scan_images(folder_id, query_term)
        if not all_imgs:
            return None, None

        bookmatch_file = None
        standard_files = []

        for f in all_imgs:
            fname = f.get("name", "").lower()
            pname = f.get("parent_name", "").lower()
            if "bookmatch" in fname or "bm" in fname or "bookmatch" in pname:
                if not bookmatch_file:
                    bookmatch_file = f
            else:
                standard_files.append(f)

        img1_bytes, img2_bytes = None, None

        # 1. Görsel (Öncelik Standart Plaka)
        if standard_files:
            img1_bytes = self._download_file(standard_files[0]["id"])
        elif bookmatch_file:
            img1_bytes = self._download_file(bookmatch_file["id"])

        # 2. Görsel (Öncelik Bookmatch, yoksa 2. Plaka)
        if bookmatch_file and bookmatch_file != (standard_files[0] if standard_files else None):
            img2_bytes = self._download_file(bookmatch_file["id"])
        elif len(standard_files) > 1:
            img2_bytes = self._download_file(standard_files[1]["id"])
        elif img1_bytes:
            img2_bytes = img1_bytes

        return img1_bytes, img2_bytes

    def get_best_slab_image(self, folder_id: str, query_term: str) -> Optional[bytes]:
        """
        PDF Sunumu için: SADECE en temiz plaka fotoğrafını seçer.
        'bookmatch' veya 'bm' içeren hiçbir fotoğrafı PDF'e dahil etmez!
        """
        all_imgs = self._scan_images(folder_id, query_term)
        if not all_imgs:
            return None

        # Sadece standart plaka fotoğraflarını filtrele
        clean_slabs = []
        for f in all_imgs:
            fname = f.get("name", "").lower()
            pname = f.get("parent_name", "").lower()
            if not ("bookmatch" in fname or "bm" in fname or "bookmatch" in pname):
                clean_slabs.append(f)
        
        # Temiz plaka varsa onu kullan, mecbursak ilk bulduğumuzu al
        target_file = clean_slabs[0] if clean_slabs else all_imgs[0]
        return self._download_file(target_file["id"])

    def _scan_images(self, folder_id: str, query_term: str) -> List[Dict[str, str]]:
        """Klasör ve alt klasörlerdeki tüm JPEG/PNG dosyalarını tarar."""
        if not self.service:
            return []
        all_imgs = []
        try:
            if folder_id:
                # Ana klasör tarama
                q = f"'{folder_id}' in parents and trashed = false and (mimeType contains 'image/jpeg' or mimeType contains 'image/png')"
                res = self.service.files().list(q=q, spaces="drive", fields="files(id, name)", pageSize=10).execute()
                all_imgs.extend(res.get("files", []))
                
                # Alt klasör tarama
                sub_folders = self.list_child_folders(folder_id)[:5]
                for sf in sub_folders:
                    sq = f"'{sf['id']}' in parents and trashed = false and (mimeType contains 'image/jpeg' or mimeType contains 'image/png')"
                    sres = self.service.files().list(q=sq, spaces="drive", fields="files(id, name)", pageSize=5).execute()
                    for img in sres.get("files", []):
                        img["parent_name"] = sf["name"]
                        all_imgs.append(img)
            else:
                # Klasör bulunamadıysa isimle genel tarama yap
                cands = self._build_candidates(query_term)
                for cand in cands[:2]:
                    q = f"trashed = false and name contains '{cand}' and (mimeType contains 'image/jpeg' or mimeType contains 'image/png')"
                    res = self.service.files().list(q=q, spaces="drive", fields="files(id, name)", pageSize=5).execute()
                    all_imgs.extend(res.get("files", []))
            return all_imgs
        except Exception as e:
            logger.error(f"Görsel tarama hatası: {e}")
            return []

    def _download_file(self, file_id: str) -> Optional[bytes]:
        try:
            req = self.service.files().get_media(fileId=file_id)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, req)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            fh.seek(0)
            return fh.read()
        except Exception:
            return None

    @staticmethod
    def _extract_folder_id(url: str) -> Optional[str]:
        if not url: return None
        m = re.search(r"/folders/([a-zA-Z0-9_-]+)", url)
        if m: return m.group(1)
        m2 = re.search(r"id=([a-zA-Z0-9_-]+)", url)
        return m2.group(1) if m2 else None


# =====================================================================
# 4. DECENTRALIZED INVENTORY ENGINE (DOĞRUDAN DRIVE EXCEL OKUMA)
# =====================================================================
class InventoryEngine:
    """
    Merkezi bir stok Excel'ine ihtiyaç duymadan, sorgulanan taşın Drive'daki 
    klasörüne sızıp orijinal çeki listesini anlık olarak okuyan otonom motor.
    """
    def __init__(self, drive: DriveExplorer):
        self.drive = drive

    def parse_packing_list_bytes(self, excel_bytes: bytes) -> Dict[str, Any]:
        """
        Drive'dan indirilen herhangi bir formattaki Excel dosyasını akıllı
        sütun eşleştirici ile okur, sıfır hatayla stok toplamlarını döner.
        """
        try:
            # Tüm sayfaları oku, ilk dolu olanı al
            xls = pd.ExcelFile(io.BytesIO(excel_bytes))
            df = pd.DataFrame()
            for sheet in xls.sheet_names:
                temp_df = pd.read_excel(xls, sheet_name=sheet)
                if not temp_df.empty and len(temp_df.columns) > 3:
                    df = temp_df
                    break
            
            if df.empty:
                raise ValueError("Excel dosyası boş veya formatı desteklenmiyor.")

            # Sütunları normalize et (Örn: 'NET M2 (m²)' -> 'NET M2 M')
            col_map = {c: normalize_turkish_key(c) for c in df.columns}
            
            # Kapsamlı Sütun Arama Algoritması
            m2_col = next((c for c, n in col_map.items() if any(k in n for k in ["M2", "ALAN", "NET", "METRAJ"])), None)
            slab_col = next((c for c, n in col_map.items() if any(k in n for k in ["ADET", "PLAKA", "PCS", "MIKTAR", "SLAB"])), None)
            thick_col = next((c for c, n in col_map.items() if any(k in n for k in ["KALIN", "THICK", "CM"])), None)
            len_col = next((c for c, n in col_map.items() if any(k in n for k in ["BOY", "UZUNLUK", "LENG"])), None)
            height_col = next((c for c, n in col_map.items() if any(k in n for k in ["EN", "YUKSEKLIK", "HEIG", "WID"])), None)
            surf_col = next((c for c, n in col_map.items() if any(k in n for k in ["YUZEY", "FINISH", "SURF"])), None)
            
            total_m2 = df[m2_col].apply(safe_float).sum() if m2_col else 0.0
            
            if slab_col:
                total_slabs = int(df[slab_col].apply(safe_float).sum())
            else:
                # Plaka sayısı yoksa dolu satır sayısını plaka kabul et
                total_slabs = len(df.dropna(how='all'))
                
            thick_val = df[thick_col].dropna().iloc[0] if thick_col and not df[thick_col].dropna().empty else 2
            surf_val = df[surf_col].dropna().iloc[0] if surf_col and not df[surf_col].dropna().empty else "Polished"
            
            avg_l, avg_h = 0, 0
            if len_col and height_col:
                valid_l = df[len_col].apply(safe_float).replace(0.0, pd.NA).dropna()
                valid_h = df[height_col].apply(safe_float).replace(0.0, pd.NA).dropna()
                if not valid_l.empty and not valid_h.empty:
                    avg_l = int(valid_l.mean())
                    avg_h = int(valid_h.mean())
                
            return {
                "total_m2": total_m2,
                "total_slabs": total_slabs,
                "thickness": f"{thick_val} cm" if str(thick_val).isdigit() else str(thick_val),
                "surface": str(surf_val).title(),
                "dimensions": f"{avg_l} x {avg_h} cm" if avg_l > 0 else "Standart Ebat",
                "raw_df": df,
                "mapped_cols": {
                    "m2": m2_col, "len": len_col, "height": height_col, 
                    "slab": slab_col, "thick": thick_col, "surf": surf_col
                }
            }
        except Exception as e:
            logger.error(f"Excel Parse Hatası: {e}")
            return {"total_m2": 0.0, "total_slabs": 0, "thickness": "2 cm", "surface": "Polished", "dimensions": "-", "raw_df": None, "mapped_cols": {}}

    def get_lot_info(self, query: str) -> Dict[str, Any]:
        """Taşın klasörüne girer, orijinal çeki listesini indirir ve anında okur."""
        folder_id, folder_link = self.drive.find_folder(query)
        if not folder_id:
            return {"status": "error", "message": "Klasör bulunamadı."}
            
        excel_bytes, excel_name = self.drive.get_excel_from_folder(folder_id)
        if not excel_bytes:
            return {
                "status": "error", 
                "message": "Klasörde orijinal Excel (Çeki Listesi) bulunamadı.", 
                "drive_link": folder_link, 
                "folder_id": folder_id
            }
            
        data = self.parse_packing_list_bytes(excel_bytes)
        return {
            "status": "success",
            "query": query.upper(),
            "drive_link": folder_link,
            "folder_id": folder_id,
            "excel_name": excel_name,
            "excel_bytes": excel_bytes,
            "data": data
        }

    async def deep_scan_collections(self) -> str:
        """
        Argon Master'ın en güçlü özelliği. 3 ana koleksiyon klasörüne girer,
        tüm alt taş klasörlerini bulur, içlerindeki Excel'leri okur ve devasa
        bir 'Tüm Stok Raporu' oluşturur.
        """
        collections = ["Arslanlar Collection", "Arslanlar Series", "Arslanlar Exclusive Collection"]
        report_lines = ["🔍 **ARGON MASTER - DEEP SCAN EXECUTIVE REPORT** 🔍\n"]
        total_factory_m2 = 0
        total_factory_slabs = 0
        
        for coll in collections:
            fid, _ = self.drive.find_folder(coll)
            if not fid: continue
            
            report_lines.append(f"📁 **{coll.upper()}**")
            sub_folders = self.drive.list_child_folders(fid)
            coll_m2 = 0
            
            for sf in sub_folders:
                ex_bytes, _ = self.drive.get_excel_from_folder(sf["id"])
                if ex_bytes:
                    data = self.parse_packing_list_bytes(ex_bytes)
                    if data["total_m2"] > 0:
                        report_lines.append(f"  ├ {sf['name']}: {data['total_m2']:,.2f} m² | {data['total_slabs']} Plaka")
                        coll_m2 += data["total_m2"]
                        total_factory_m2 += data["total_m2"]
                        total_factory_slabs += data["total_slabs"]
                await asyncio.sleep(0.1) # Rate limit koruması
            
            report_lines.append(f"  └ *Koleksiyon Toplamı: {coll_m2:,.2f} m²*\n")
            
        report_lines.append(f"🌐 **FABRİKA GENEL TOPLAMI: {total_factory_m2:,.2f} m² | {total_factory_slabs} Plaka**")
        return "\n".join(report_lines)


# =====================================================================
# 5. PDF ÜRETİM MOTORU (SUNUM VE ÇEKİ LİSTESİ)
# =====================================================================
class PDFPresentationEngine:
    def __init__(self, inventory: InventoryEngine, drive: DriveExplorer):
        self.inventory = inventory
        self.drive = drive

    @staticmethod
    def _create_fallback_image() -> bytes:
        """Kayıp görseller için şık bir Arslanlar logolu yedek görsel."""
        img = Image.new("RGB", (600, 400), color=(20, 33, 61))
        draw = ImageDraw.Draw(img)
        draw.rectangle([12, 12, 588, 388], outline=(197, 160, 89), width=2)
        draw.text((180, 185), "ARSLANLAR MARBLE PREMIER", fill=(220, 225, 230))
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=90)
        return buf.getvalue()

    def generate_presentation_pdf(self, queries: List[str]) -> Tuple[bytes, str]:
        """
        %100 İngilizce, Arslanlar logolu, Bookmatch GÖRSELSİZ, sadece 
        temiz plaka fotoğraflı Lüks PDF Sunumu.
        """
        cards = []
        for q in queries:
            info = self.inventory.get_lot_info(q)
            if info["status"] != "success": continue
            
            # Sadece tek ve en iyi plaka görselini al (Bookmatch elendi)
            img1 = self.drive.get_best_slab_image(info["folder_id"], q)
            cards.append({
                "title": q.upper(),
                "data": info["data"],
                "drive_link": info["drive_link"],
                "image_bytes": img1
            })

        if not cards:
            raise ValueError("Belirtilen lotlar için veri veya klasör bulunamadı.")

        pdf_io = io.BytesIO()
        doc = SimpleDocTemplate(pdf_io, pagesize=A4, rightMargin=32, leftMargin=32, topMargin=32, bottomMargin=32)
        styles = getSampleStyleSheet()
        NAVY, GOLD, SLATE, LIGHT_BG, BORD = colors.HexColor("#14213D"), colors.HexColor("#C5A059"), colors.HexColor("#2B2D42"), colors.HexColor("#F8F9FA"), colors.HexColor("#DEE2E6")

        title_s = ParagraphStyle("T", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=21, textColor=NAVY, alignment=1)
        sub_s = ParagraphStyle("S", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=10, textColor=GOLD, alignment=1, spaceAfter=8)
        ch_s = ParagraphStyle("CH", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=12, textColor=NAVY, spaceAfter=6)
        th_s = ParagraphStyle("TH", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=8.5, textColor=colors.whitesmoke)
        tdl_s = ParagraphStyle("TDL", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=8.5, textColor=SLATE)
        tdv_s = ParagraphStyle("TDV", parent=styles["Normal"], fontName="Helvetica", fontSize=8.5, textColor=SLATE)

        story = [
            Paragraph("ARSLANLAR MARBLE", title_s), 
            Paragraph("EXECUTIVE STOCK & SLAB PRESENTATION", sub_s), 
            HRFlowable(width="100%", thickness=1.5, color=GOLD, spaceAfter=12)
        ]

        for idx, card in enumerate(cards, start=1):
            d = card["data"]
            m2 = d["total_m2"]
            slabs = d["total_slabs"]
            
            # İstenmeyen "Total Production" ve "Avg. Slab Size" boş/sıfırsa kaldırıldı.
            stock_str = f"{m2:,.2f} m² ({slabs} Slabs)" if m2 > 0 else "Contact for Availability"
            surf_thick_str = f"{d['surface']} | {d['thickness']}"
            
            raw_img = card["image_bytes"] or self._create_fallback_image()
            try:
                pil_im = Image.open(io.BytesIO(raw_img)).convert("RGB")
                img_io = io.BytesIO()
                pil_im.save(img_io, format="JPEG", quality=88)
                img_io.seek(0)
                rl_img = RLImage(img_io, width=170, height=120)
            except:
                rl_img = RLImage(io.BytesIO(self._create_fallback_image()), width=170, height=120)

            spec_rows = [
                [Paragraph("Specification", th_s), Paragraph("Technical Data", th_s)],
                [Paragraph("Available Stock", tdl_s), Paragraph(stock_str, tdv_s)],
                [Paragraph("Surface & Thickness", tdl_s), Paragraph(surf_thick_str, tdv_s)],
                [Paragraph("Active Code", tdl_s), Paragraph(card["title"], tdv_s)],
                [Paragraph("High-Res Photos", tdl_s), Paragraph(f"<link href='{card['drive_link']}' color='#0056b3'><u>Open Google Drive Folder</u></link>", tdv_s)],
            ]

            st = Table(spec_rows, colWidths=[115, 235])
            st.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), NAVY), 
                ("BACKGROUND", (0,1), (-1,-1), LIGHT_BG), 
                ("GRID", (0,0), (-1,-1), 0.5, BORD), 
                ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                ("TOPPADDING", (0,0), (-1,-1), 4),
                ("BOTTOMPADDING", (0,0), (-1,-1), 4),
            ]))
            mr = Table([[rl_img, st]], colWidths=[180, 360])
            mr.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("ALIGN", (0,0), (0,0), "CENTER")]))
            
            story.append(KeepTogether([Paragraph(f"{idx}. {card['title']}", ch_s), mr, Spacer(1, 15)]))

        doc.build(story)
        pdf_io.seek(0)
        return pdf_io.read(), f"Arslanlar_Presentation_{datetime.now().strftime('%Y%m%d')}.pdf"

    def generate_packing_list_pdf(self, raw_df: pd.DataFrame, lot_name: str, mapped_cols: Dict[str, str]) -> bytes:
        """
        Orijinal Excel'den (raw_df) okunan satırları, A4 yatay (Landscape) formatında
        Arslanlar logolu, son derece resmi ve şık bir PDF Çeki Listesine dönüştürür.
        """
        pdf_io = io.BytesIO()
        doc = SimpleDocTemplate(pdf_io, pagesize=landscape(A4), rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
        styles = getSampleStyleSheet()
        
        NAVY = colors.HexColor("#14213D")
        GOLD = colors.HexColor("#C5A059")
        SLATE = colors.HexColor("#2B2D42")
        
        title_p = ParagraphStyle("PlT", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=18, textColor=NAVY, alignment=1)
        sub_p = ParagraphStyle("PlS", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=10, textColor=GOLD, alignment=1, spaceAfter=15)
        th_p = ParagraphStyle("TH", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9, textColor=colors.whitesmoke, alignment=1)
        td_p = ParagraphStyle("TD", parent=styles["Normal"], fontName="Helvetica", fontSize=8.5, textColor=SLATE, alignment=1)

        story = [
            Paragraph("ARSLANLAR MARBLE", title_p),
            Paragraph(f"OFFICIAL PACKING LIST — {lot_name.upper()}", sub_p),
            HRFlowable(width="100%", thickness=1, color=GOLD, spaceAfter=10),
        ]

        table_data = [[
            Paragraph("Row #", th_p),
            Paragraph("Slab No", th_p),
            Paragraph("Length", th_p),
            Paragraph("Height", th_p),
            Paragraph("Thick", th_p),
            Paragraph("Net M2", th_p),
        ]]

        c_len = mapped_cols.get("len")
        c_h = mapped_cols.get("height")
        c_m2 = mapped_cols.get("m2")
        c_slab = mapped_cols.get("slab")
        c_th = mapped_cols.get("thick")

        total_m2 = 0.0
        row_count = 0
        
        for idx, row in raw_df.iterrows():
            # Tamamen boş satırları atla
            if pd.isna(row.get(c_m2)) or safe_float(row.get(c_m2)) <= 0:
                continue
                
            row_count += 1
            m2_val = safe_float(row.get(c_m2))
            total_m2 += m2_val
            
            table_data.append([
                Paragraph(str(row_count), td_p),
                Paragraph(str(row.get(c_slab, "-")), td_p),
                Paragraph(str(row.get(c_len, "-")), td_p),
                Paragraph(str(row.get(c_h, "-")), td_p),
                Paragraph(str(row.get(c_th, "-")), td_p),
                Paragraph(f"{m2_val:.2f}", td_p),
            ])
            
            # PDF'in şişmemesi için ilk 150 satırı al
            if row_count >= 150:
                break

        pl_table = Table(table_data, colWidths=[60, 100, 100, 100, 80, 100])
        pl_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), NAVY),
            ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F8F9FA")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#DEE2E6")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(pl_table)
        story.append(Spacer(1, 15))

        sum_p = Paragraph(
            f"<b>Total Evaluated Slabs:</b> {row_count} pcs | <b>Total Net Area:</b> {total_m2:,.2f} m²<br/>"
            f"<i>Arslanlar Marble Export Department — Packing List Generated Autonomously by Argon Master.</i>",
            ParagraphStyle("Sum", parent=styles["Normal"], fontName="Helvetica-Oblique", fontSize=9, textColor=SLATE)
        )
        story.append(sum_p)

        doc.build(story)
        pdf_io.seek(0)
        return pdf_io.read()

# =====================================================================
# 6. ARGON AI ORKESTRATÖRÜ (KIRILMAZ ÇİFT MOTOR)
# =====================================================================
class ArgonBrain:
    def __init__(self, api_key: str, sa_json_str: str):
        self.api_key = api_key
        self.sa_json_str = sa_json_str
        self.endpoint = "https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent"

    async def evaluate_intent(self, user_text: str, is_group: bool, bot_mentioned: bool) -> Dict[str, Any]:
        """
        Argon'un Zekası: Yüksek özgüvenli, hata yapmayan, direkt işlem odaklı.
        Metni analiz eder, lot kodlarını çıkarır ve kullanıcının ne istediğini (PDF, Çeki, Foto, Deep Scan) anlar.
        """
        raw = (user_text or "").strip().upper()
        if not raw: return {"should_reply": False}
        
        # M, CM, MM, PDF gibi sahte lotları engelleyen akıllı Regex
        lot_regex = re.compile(r"\b([A-Z]{1,4})[-\s]?(\d{1,5})\b")
        matched = []
        for m in lot_regex.finditer(raw):
            prefix, num = m.group(1), m.group(2)
            if prefix not in ["M", "CM", "MM", "KG", "DC", "TON", "PDF"]:
                matched.append(f"{prefix}-{num}")
                
        # Niyet (Intent) Tespitleri
        want_deep_scan = any(kw in raw for kw in ["TÜM STOĞUMU", "TUM STOGUMU", "GENEL TARAMA", "DEEP SCAN", "TÜM KOLEKSIYON", "RAPORLA"])
        want_pdf = any(kw in raw for kw in ["PDF", "SUNUM", "KATALOG"])
        want_packing = any(kw in raw for kw in ["ÇEKİ", "CEKI", "PACKING", "SEVK", "LISTESI", "LIST"])
        want_photo = any(kw in raw for kw in ["FOTO", "RESIM", "GORSEL"])
        is_greeting = any(kw in raw for kw in ["MERHABA", "SELAM", "ORDA MISIN", "HELLO"])

        has_intent = matched or want_deep_scan or want_pdf or want_packing or want_photo
        
        # Grup içinde gereksiz konuşmalara susma yeteneği
        if is_group and not bot_mentioned and not has_intent and not is_greeting:
            return {"should_reply": False}

        return {
            "should_reply": True,
            "matched_queries": matched[:3], # En fazla 3 lot işleme alınır
            "want_deep_scan": want_deep_scan,
            "want_pdf": want_pdf,
            "want_packing": want_packing,
            "want_photo": want_photo or bool(matched) # Lot sorulursa default fotoğrafları at
        }

# =====================================================================
# 7. TELEGRAM SOHBET YÖNETİCİSİ (ASIL SİHİR BURADA)
# =====================================================================
drive_explorer = DriveExplorer(GOOGLE_SA_JSON)
inventory_engine = InventoryEngine(drive_explorer)
pdf_engine = PDFPresentationEngine(inventory_engine, drive_explorer)
argon_brain = ArgonBrain(GEMINI_API_KEY, GOOGLE_SA_JSON)

PROCESSED_MESSAGES = deque(maxlen=500)

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.effective_message
    if not msg or not update.effective_chat: return
    
    # Çift mesaj engelleme (Anti-Spam)
    msg_key = f"{update.effective_chat.id}:{msg.message_id}"
    if msg_key in PROCESSED_MESSAGES: return
    PROCESSED_MESSAGES.append(msg_key)

    chat_id = update.effective_chat.id
    is_group = update.effective_chat.type in ["group", "supergroup"]
    bot_mentioned = bool(context.bot.username and f"@{context.bot.username.lower()}" in (msg.text or "").lower())
    
    intent = await argon_brain.evaluate_intent(msg.text or "", is_group, bot_mentioned)
    if not intent.get("should_reply"): return

    # ---------------------------------------------------------
    # 1. BÜYÜK VERİ TARAMASI (DEEP SCAN) İSTEĞİ
    # ---------------------------------------------------------
    if intent.get("want_deep_scan"):
        status = await msg.reply_text("🔎 **Argon Deep Scan Başlatıldı.**\nTüm Drive klasörlerinize ve Excel dosyalarınıza eşzamanlı iniliyor. Lütfen bekleyin...", parse_mode=ParseMode.MARKDOWN)
        report = await inventory_engine.deep_scan_collections()
        await status.edit_text(report, parse_mode=ParseMode.MARKDOWN)
        return

    queries = intent.get("matched_queries", [])
    
    # EĞER LOT SORGUSU YOKSA VE SADECE SELAM VERİLDİYSE
    if not queries:
        await msg.reply_text("👋 Ben Argon Master. Merkezi Drive veri ağınıza bağlıyım, operasyon raporlamasına hazırım.\nBir Lot kodu (Örn: A-371), çeki listesi, PDF sunum veya 'Tüm stoğumu raporla' diyerek işlem başlatabilirsiniz.")
        return

    await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.TYPING)
    
    # ---------------------------------------------------------
    # 2. LOT BAZLI İŞLEMLER (DÖNGÜ)
    # ---------------------------------------------------------
    for q in queries:
        # A. Drive'dan Canlı Excel Okuma
        info = await asyncio.to_thread(inventory_engine.get_lot_info, q)
        
        if info["status"] != "success":
            await msg.reply_text(f"⚠️ **{q}** için veritabanı okunamadı.\nSebep: {info['message']}\nKlasör Linki: {info.get('drive_link','Bulunamadı')}", parse_mode=ParseMode.MARKDOWN)
            continue

        d = info["data"]
        fid = info["folder_id"]
        flink = info["drive_link"]

        # B. ÇEKİ LİSTESİ İSTENDİYSE (DUAL EXPORT: EXCEL + PDF)
        if intent.get("want_packing"):
            await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.UPLOAD_DOCUMENT)
            # Orijinal Excel'i gönder
            await msg.reply_document(
                document=io.BytesIO(info["excel_bytes"]), 
                filename=info["excel_name"], 
                caption=f"📦 **{q}** Orijinal Excel Çeki Listesi (Drive'dan Çekildi)", 
                parse_mode=ParseMode.MARKDOWN
            )
            
            # Pandas DataFrame'den PDF Çeki Listesi üret ve gönder
            if d.get("raw_df") is not None:
                pdf_pack_bytes = await asyncio.to_thread(
                    pdf_engine.generate_packing_list_pdf, d["raw_df"], q, d["mapped_cols"]
                )
                await msg.reply_document(
                    document=io.BytesIO(pdf_pack_bytes), 
                    filename=f"PackingList_{q}.pdf", 
                    caption=f"📄 **{q}** Resmi PDF Çeki Listesi", 
                    parse_mode=ParseMode.MARKDOWN
                )
            continue
            
        # C. PDF SUNUM İSTENDİYSE (BOOKMATCH'SİZ, PREMIUM)
        if intent.get("want_pdf"):
            await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.UPLOAD_DOCUMENT)
            try:
                pdf_b, pdf_n = await asyncio.to_thread(pdf_engine.generate_presentation_pdf, [q])
                await msg.reply_document(document=io.BytesIO(pdf_b), filename=pdf_n, caption=f"📄 **{q}** Premium İngilizce PDF Sunumu", parse_mode=ParseMode.MARKDOWN)
            except Exception as e:
                await msg.reply_text(f"PDF oluşturulurken hata: {e}")
            continue

        # D. NORMAL SORGULAMA (TELEGRAM FLASH-KART VE 2'Lİ ALBÜM)
        # Mesaj metni (Flash-Kart formatı)
        text_response = (
            f"💎 **{q.upper()} (Live Drive Data)**\n"
            f"• Mevcut Stok: `{d['total_m2']:,.2f} m²` (`{d['total_slabs']} Plaka`)\n"
            f"• Ortalama Ebat: `{d['dimensions']}`\n"
            f"• Yüzey & Kalınlık: `{d['surface']} ({d['thickness']})`\n"
            f"• Drive Linki: {flink}"
        )
        
        # 1. Fotoğraf (Plaka) ve 2. Fotoğraf (Varsa Bookmatch) Çek
        img1, img2 = await asyncio.to_thread(drive_explorer.get_images_structured, fid, q)
        
        media = []
        if img1: 
            # İlk fotoğrafa caption (alt yazı) ekle
            media.append(InputMediaPhoto(media=img1, caption=text_response, parse_mode=ParseMode.MARKDOWN))
        if img2: 
            media.append(InputMediaPhoto(media=img2))
        
        # Albümü gönder
        if media:
            await msg.reply_media_group(media=media)
        else:
            await msg.reply_text(text_response, parse_mode=ParseMode.MARKDOWN)


# =====================================================================
# 8. RENDER SIFIR-BEKLEME (ZERO-BLOCKING) BAŞLATICI
# =====================================================================
flask_app = Flask(__name__)

@flask_app.route("/")
def health_check(): 
    return "Argon Master Perfect Version v4 (Deep Scan) is Live!", 200

def run_flask():
    flask_app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)), use_reloader=False)

def main():
    if not TELEGRAM_TOKEN:
        logger.error("CRITICAL: TELEGRAM_TOKEN eksik!")
        return

    threading.Thread(target=run_flask, daemon=True).start()
    application = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    
    # Sadece metin ve fotoğraf mesajlarını dinle (Komutları yoksay)
    application.add_handler(MessageHandler((filters.TEXT | filters.PHOTO) & (~filters.COMMAND), handle_message))
    
    logger.info("🚀 Argon Master Perfect Version v4 başlatılıyor...")
    application.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
