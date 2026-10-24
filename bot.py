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

from reportlab.lib.pagesizes import A4
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

from telegram import Update, InputMediaPhoto
from telegram.constants import ParseMode, ChatAction
from telegram.ext import (
    ApplicationBuilder,
    ContextTypes,
    MessageHandler,
    CommandHandler,
    filters,
)

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# =====================================================================
# 1. LOGLAMA VE ORTAM DEĞİŞKENLERİ (ENVIRONMENT VARIABLES)
# =====================================================================
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("ArslanlarArgonMaster")

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GOOGLE_SA_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()

DEFAULT_DRIVE_LINK = (
    os.getenv("GOOGLE_DRIVE_LINK")
    or os.getenv("DEFAULT_DRIVE_LINK")
    or "https://drive.google.com"
).strip()

EXCEL_FILE_PATH = os.getenv("EXCEL_FILE_PATH", "stok.xlsx").strip()

HYPERLINKS_DB_FILE = "hyperlinks_db.json"
VISUAL_CACHE_FILE = "visual_cache.json"
STONE_DNA_DB_FILE = "stone_dna_db.json"

# =====================================================================
# 2. VERİ TEMİZLİĞİ: KARA LİSTE VE EŞ ANLAMLI TAŞ BİRLEŞTİRME HARİTASI
# =====================================================================
BLACKLIST_PATTERNS = [
    "SERKAN CELIK",
    "PIETRA(BUSE",
    "STONE LINE",
    "DERE MERMER",
    "ERKOCLAR",
    "DENIZLI OCAK",
    "IRAK PROJE",
    "SIMSEKLER",
    "KARISIK TASLAR",
    "TOPLAM",
    "GENEL TOPLAM",
    "ARA TOPLAM",
    "SUBTOTAL",
    "TOTAL",
    "IADE",
    "FIRE",
    "BELIRSIZ",
    "TEST",
    "ORNEK",
    "UNNAMED",
]

# Yazım hatalarını tek bir ana taş isminde birleştiren sözlük
STONE_SYNONYM_MAP = {
    "SILVER ROODS": "SILVER ROOTS",
    "SILVER ROOT": "SILVER ROOTS",
    "DECART": "DEKART",
    "ALEXDANDER BLACK": "ALEXANDER BLACK",
    "GRIGIO-GREY": "GRIGIO GREY",
    "CEPPO BEIGE": "CEPPO BEJ",
    "TUNDURA": "TUNDRA GREY",
    "SILVER ARMONI": "SILVER ROOTS ARMONI",
    "SILVER ROOTS BRECIA": "SILVER ROOTS BRECCIA",
    "VOILET GREY": "VIOLET GREY",
    "EXOTIC GREEN": "EXOTIC GREEN",
}

DEFAULT_STONE_DNA = {
    "A-283": {
        "stone_name": "VIOLET BRECCIA",
        "visual_dna": "White/cream broken breccia clasts with deep violet, purple, and dark grey veining.",
        "keywords": ["violet", "breccia", "mor", "kırıklı", "a-283", "a283"],
    },
    "SILVER ROOTS": {
        "stone_name": "SILVER ROOTS",
        "visual_dna": "Silver-grey and smoked background with fine root-like white, gold, and charcoal veins.",
        "keywords": ["silver roots", "gümüş", "gri", "kök damar", "füme"],
    },
    "BLUE ROOTS": {
        "stone_name": "BLUE ROOTS",
        "visual_dna": "Cool blue-grey background with root-like dark grey and white mineral veining.",
        "keywords": ["blue roots", "mavi", "gri", "damarlı"],
    },
    "ROSSO LEVANTO": {
        "stone_name": "ROSSO LEVANTO",
        "visual_dna": "Deep burgundy and cherry-red marble with striking contrasting white calcite veins.",
        "keywords": ["rosso levanto", "bordo", "kırmızı", "vişne", "beyaz damar"],
    },
    "TERRA MATER": {
        "stone_name": "TERRA MATER",
        "visual_dna": "Earthy brown, beige, and warm grey wavy brecciated natural stone texture.",
        "keywords": ["terra mater", "toprak", "kahve", "dalgalı"],
    },
    "SILVER RIVER": {
        "stone_name": "SILVER RIVER",
        "visual_dna": "Flowing parallel and wavy silver-grey, white, and black river-like striated layers.",
        "keywords": ["silver river", "nehir", "çizgisel", "dalgalı gri"],
    },
    "DEKART": {
        "stone_name": "DEKART",
        "visual_dna": "Modern anthracite and medium grey marble with subtle cloudy and linear movement.",
        "keywords": ["dekart", "decart", "antrasit", "gri"],
    },
}


# =====================================================================
# 3. YARDIMCI FONKSİYONLAR VE NORMALİZASYON MOTORU
# =====================================================================
def load_json_file(filepath: str, default_data: Any) -> Any:
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logger.warning(f"{filepath} okunamadı: {e}")
    return default_data


def save_json_file(filepath: str, data: Any) -> None:
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.error(f"{filepath} kaydedilemedi: {e}")


def normalize_turkish_key(text: Any) -> str:
    """
    Türkçe karakterleri İngilizce ASCII karşılıklarına çevirir, fazla boşlukları
    temizler ve STONE_SYNONYM_MAP ile yazım farklarını tekilleştirir.
    """
    if text is None or pd.isna(text):
        return ""
    s = str(text).strip()
    if not s or s.lower() == "nan":
        return ""
    tr_map = str.maketrans(
        {
            "İ": "I",
            "ı": "I",
            "i": "I",
            "Ş": "S",
            "ş": "S",
            "Ğ": "G",
            "ğ": "G",
            "Ü": "U",
            "ü": "U",
            "Ö": "O",
            "ö": "O",
            "Ç": "C",
            "ç": "C",
        }
    )
    s = s.translate(tr_map).upper()
    s = re.sub(r"\s+", " ", s).strip()
    if s in STONE_SYNONYM_MAP:
        s = STONE_SYNONYM_MAP[s]
    return s


def normalize_lot_code(code: Any) -> str:
    """'a 283', 'A283', 'a-283' -> 'A-283' standardına çevirir."""
    raw = normalize_turkish_key(code)
    if not raw:
        return ""
    match = re.match(r"^([A-Z]+)[-\s]*(\d+)$", raw)
    if match:
        return f"{match.group(1)}-{match.group(2)}"
    return raw


def is_blacklisted_stone(norm_name: str) -> bool:
    if not norm_name or len(norm_name) < 2:
        return True
    for pat in BLACKLIST_PATTERNS:
        if pat in norm_name:
            return True
    return False


def safe_float(val: Any) -> float:
    if val is None or pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    s = re.sub(r"[^\d,.-]", "", s)
    if not s:
        return 0.0
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def format_thickness_en(raw_val: Any) -> str:
    """20.0, 21.0, 22.0 -> '2 cm' | 30.0 -> '3 cm' standardına çevirir."""
    val = safe_float(raw_val)
    if val <= 0:
        return "2 cm"
    if 17 <= val <= 24 or val == 2:
        return "2 cm"
    if 27 <= val <= 34 or val == 3:
        return "3 cm"
    if val > 10:
        return f"{val / 10:.1f} cm"
    return f"{val:g} cm"


def format_surface_en(raw_surface: Any) -> str:
    norm = normalize_turkish_key(raw_surface)
    if not norm or norm == "NAN":
        return "Polished"
    if "POLISHED" in norm or "PARLAK" in norm or "CILALI" in norm:
        return "Polished"
    if "HONED" in norm or "HONLU" in norm or "MAT" in norm:
        return "Honed"
    if "LEATHER" in norm or "BRUSHED" in norm or "ESKITME" in norm:
        return "Leather / Brushed"
    return str(raw_surface).strip().title()


# =====================================================================
# 4. GOOGLE DRIVE OTONOM GEZGİN, ALT KLASÖR TARAYICI VE GÖRSEL MOTORU
# =====================================================================
class DriveExplorer:
    def __init__(self, sa_json_str: str):
        self.service = None
        self.sa_json_str = sa_json_str
        self.hyperlinks_db: Dict[str, str] = load_json_file(HYPERLINKS_DB_FILE, {})
        self.visual_cache: Dict[str, List[str]] = load_json_file(VISUAL_CACHE_FILE, {})
        self._init_service()

    def _init_service(self) -> None:
        if not self.sa_json_str:
            logger.warning("GOOGLE_SERVICE_ACCOUNT_JSON tanımlı değil!")
            return
        try:
            info = json.loads(self.sa_json_str)
            creds = service_account.Credentials.from_service_account_info(
                info, scopes=["https://www.googleapis.com/auth/drive.readonly"]
            )
            self.service = build("drive", "v3", credentials=creds, cache_discovery=False)
            logger.info("Google Drive API bağlantısı başarıyla kuruldu.")
        except Exception as e:
            logger.error(f"Google Drive API başlatma hatası: {e}")

    def save_custom_link(self, key: str, url: str) -> None:
        norm_key = normalize_lot_code(key)
        if norm_key and url:
            self.hyperlinks_db[norm_key] = url.strip()
            save_json_file(HYPERLINKS_DB_FILE, self.hyperlinks_db)

    def find_folder(self, query_term: str) -> Tuple[Optional[str], Optional[str]]:
        """
        Lot kodu ('A-283'), taş adı ('SILVER ROOTS') veya ana koleksiyon adı
        ('Arslanlar Collection', 'Arslanlar Series') için Drive klasörünü bulur.
        """
        norm_key = normalize_lot_code(query_term)
        raw_clean = query_term.strip().replace("'", "")

        if norm_key in self.hyperlinks_db:
            cached_url = self.hyperlinks_db[norm_key]
            folder_id = self._extract_folder_id_from_url(cached_url)
            if cached_url and cached_url != DEFAULT_DRIVE_LINK and folder_id:
                return folder_id, cached_url

        if not self.service:
            return None, None

        candidates = [raw_clean, norm_key]
        if "-" in norm_key:
            candidates.append(norm_key.replace("-", " "))
            candidates.append(norm_key.replace("-", ""))
        candidates = list(dict.fromkeys([c for c in candidates if c]))

        try:
            for cand in candidates:
                q = (
                    f"mimeType = 'application/vnd.google-apps.folder' "
                    f"and trashed = false "
                    f"and name contains '{cand}'"
                )
                res = (
                    self.service.files()
                    .list(
                        q=q,
                        spaces="drive",
                        fields="files(id, name, webViewLink)",
                        includeItemsFromAllDrives=True,
                        supportsAllDrives=True,
                        pageSize=15,
                    )
                    .execute()
                )
                files = res.get("files", [])
                if files:
                    best = files[0]
                    for f in files:
                        if normalize_lot_code(f["name"]) == norm_key:
                            best = f
                            break
                    folder_id = best["id"]
                    link = best.get(
                        "webViewLink",
                        f"https://drive.google.com/drive/folders/{folder_id}",
                    )
                    self.hyperlinks_db[norm_key] = link
                    save_json_file(HYPERLINKS_DB_FILE, self.hyperlinks_db)
                    return folder_id, link
        except Exception as e:
            logger.error(f"Drive klasör arama hatası ({query_term}): {e}")

        return None, None

    def list_child_folders(self, parent_folder_id: str, max_folders: int = 20) -> List[Dict[str, str]]:
        """
        'Arslanlar Collection' veya 'Arslanlar Series' gibi bir ana klasörün
        içindeki tüm alt klasörleri (taş/lot klasörlerini) listeler.
        """
        if not self.service or not parent_folder_id:
            return []
        try:
            q = (
                f"'{parent_folder_id}' in parents and trashed = false "
                f"and mimeType = 'application/vnd.google-apps.folder'"
            )
            res = (
                self.service.files()
                .list(
                    q=q,
                    spaces="drive",
                    fields="files(id, name, webViewLink)",
                    includeItemsFromAllDrives=True,
                    supportsAllDrives=True,
                    pageSize=max_folders,
                )
                .execute()
            )
            folders = []
            for f in res.get("files", []):
                fid = f["id"]
                flink = f.get(
                    "webViewLink", f"https://drive.google.com/drive/folders/{fid}"
                )
                folders.append({"id": fid, "name": f["name"], "link": flink})
            return folders
        except Exception as e:
            logger.error(f"Alt klasör tarama hatası ({parent_folder_id}): {e}")
            return []

    def get_folder_images(self, folder_id: str, max_images: int = 2) -> List[bytes]:
        """
        Belirtilen Drive klasörünün (veya içindeki alt klasörlerin) içindeki
        gerçek plaka fotoğraflarını indirir.
        """
        if not self.service or not folder_id:
            return []

        image_bytes_list: List[bytes] = []
        try:
            q = (
                f"'{folder_id}' in parents and trashed = false "
                f"and (mimeType contains 'image/jpeg' or mimeType contains 'image/png')"
            )
            res = (
                self.service.files()
                .list(
                    q=q,
                    spaces="drive",
                    fields="files(id, name, mimeType, size)",
                    includeItemsFromAllDrives=True,
                    supportsAllDrives=True,
                    pageSize=max_images,
                )
                .execute()
            )
            files = res.get("files", [])

            # Doğrudan bu klasörde fotoğraf yoksa 1 seviye alt klasörlere bak
            if not files:
                sub_folders = self.list_child_folders(folder_id, max_folders=4)
                for sub in sub_folders:
                    img_q = (
                        f"'{sub['id']}' in parents and trashed = false "
                        f"and (mimeType contains 'image/jpeg' or mimeType contains 'image/png')"
                    )
                    img_res = (
                        self.service.files()
                        .list(
                            q=img_q,
                            spaces="drive",
                            fields="files(id, name, mimeType, size)",
                            includeItemsFromAllDrives=True,
                            supportsAllDrives=True,
                            pageSize=max_images,
                        )
                        .execute()
                    )
                    files.extend(img_res.get("files", []))
                    if len(files) >= max_images:
                        break

            for f in files[:max_images]:
                request = self.service.files().get_media(
                    fileId=f["id"], supportsAllDrives=True
                )
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
                fh.seek(0)
                image_bytes_list.append(fh.read())

        except Exception as e:
            logger.error(f"Drive görsel indirme hatası (folder_id={folder_id}): {e}")

        return image_bytes_list

    @staticmethod
    def _extract_folder_id_from_url(url: str) -> Optional[str]:
        if not url:
            return None
        m = re.search(r"/folders/([a-zA-Z0-9_-]+)", url)
        if m:
            return m.group(1)
        m2 = re.search(r"id=([a-zA-Z0-9_-]+)", url)
        return m2.group(1) if m2 else None


# =====================================================================
# 5. AKILLI EXCEL STOK MOTORU (16.630 SATIRLIK GERÇEK TABLO UYUMLU)
# =====================================================================
class InventoryEngine:
    def __init__(self, excel_path: str, drive_explorer: DriveExplorer):
        self.excel_path = excel_path
        self.drive = drive_explorer
        self.stones_summary: Dict[str, Dict[str, Any]] = {}
        self.lots_index: Dict[str, Dict[str, Any]] = {}
        self.stone_dna_db: Dict[str, Dict[str, Any]] = load_json_file(
            STONE_DNA_DB_FILE, DEFAULT_STONE_DNA
        )
        self.last_loaded_time: Optional[str] = None

    def locate_any_excel_file(self) -> Optional[str]:
        """
        Sunucuda 'stok.xlsx' yoksa bile dizindeki herhangi bir '.xlsx'
        (Örn: 'PLAKA STOK LİSTESİ.xlsb (1) (1).xlsx') dosyasını otomatik bulur.
        """
        if os.path.exists(self.excel_path):
            return self.excel_path
        for fname in os.listdir("."):
            if fname.lower().endswith((".xlsx", ".xls")) and not fname.startswith("~$"):
                logger.info(f"Alternatif Excel dosyası bulundu: {fname}")
                return fname
        return None

    def reload_excel(self) -> Tuple[int, int]:
        target_path = self.locate_any_excel_file()
        if not target_path:
            logger.warning(f"Excel dosyası henüz bulunamadı ({self.excel_path}). Telegram'dan yüklenmesi bekleniyor.")
            return 0, 0

        try:
            # 1. Hücre içi köprüleri (hyperlinks) yakala
            try:
                wb = openpyxl.load_workbook(target_path, data_only=True)
                for sheet in wb.sheetnames:
                    ws = wb[sheet]
                    for row in ws.iter_rows():
                        for cell in row:
                            if cell.hyperlink and cell.hyperlink.target and cell.value:
                                key = normalize_lot_code(str(cell.value))
                                self.drive.save_custom_link(key, cell.hyperlink.target)
            except Exception as he:
                logger.debug(f"Hyperlink taraması atlandı: {he}")

            # 2. Pandas ile tüm sayfaları oku
            xls = pd.ExcelFile(target_path)
            stones_agg: Dict[str, Dict[str, Any]] = {}
            lots_agg: Dict[str, Dict[str, Any]] = {}

            for sheet_name in xls.sheet_names:
                df = pd.read_excel(target_path, sheet_name=sheet_name)
                if df.empty:
                    continue

                col_map = {col: normalize_turkish_key(col) for col in df.columns}

                stone_col = self._find_col(
                    col_map, ["SELEKSIYON", "TAS CINSI", "URUN ADI", "MERMER", "CINSI", "STONE"]
                )
                fab_lot_col = self._find_col(col_map, ["FABRIKA KODU", "LOT", "PARTI", "BUNDLE"])
                blok_col = self._find_col(col_map, ["BLOK NO", "BLOK"])
                status_col = self._find_col(col_map, ["STOK DURUMU", "DURUM"])
                loc_col = self._find_col(col_map, ["STOK KONUMU", "KONUM", "KORIDOR", "HOL"])
                res_col = self._find_col(col_map, ["SATIS VE REZERV", "REZERV"])
                m2_col = self._find_col(col_map, ["M2", "NET M2", "METRAJ", "ALAN"])
                len_col = self._find_col(col_map, ["UZUNLUK", "BOY", "LENGTH"])
                height_col = self._find_col(col_map, ["YUKSEKLIK", "EN", "HEIGHT", "WIDTH"])
                thick_col = self._find_col(col_map, ["KALINLIK", "THICKNESS", "CM"])
                surf_col = self._find_col(col_map, ["YUZEY ISLEMI", "YUZEY", "SURFACE"])
                adet_col = self._find_col(col_map, ["ADET", "PCS", "SLAB COUNT"])

                if not stone_col or not m2_col:
                    continue

                for _, row in df.iterrows():
                    raw_stone = row.get(stone_col)
                    norm_stone = normalize_turkish_key(raw_stone)

                    # Fabrika Kodu (Örn: A-283) ve Blok No okuma
                    raw_fab = row.get(fab_lot_col) if fab_lot_col else None
                    raw_blok = row.get(blok_col) if blok_col else None
                    lot_code = normalize_lot_code(raw_fab) or normalize_lot_code(raw_blok)

                    if lot_code in self.stone_dna_db and (not norm_stone or norm_stone == "BELIRSIZ"):
                        norm_stone = self.stone_dna_db[lot_code]["stone_name"]

                    if is_blacklisted_stone(norm_stone):
                        continue

                    m2_val = safe_float(row.get(m2_col))
                    if m2_val <= 0:
                        continue

                    # Eğer 'ADET' sütunu varsa onu kullan, yoksa (Stok sayfasındaki gibi) her satır 1 plakadır
                    slab_count = int(safe_float(row.get(adet_col))) if adet_col and pd.notna(row.get(adet_col)) else 1
                    if slab_count <= 0:
                        slab_count = 1

                    # Stok durumu kontrolü ('STOK' vs 'YÜKLENDİ'/'EBATLANDI'/'KIRILDI')
                    raw_status = normalize_turkish_key(row.get(status_col)) if status_col else "STOK"
                    is_in_stock = (raw_status == "STOK") or ("SHOWROOM" in raw_status) or (status_col is None)

                    raw_loc = normalize_turkish_key(row.get(loc_col)) if loc_col else ""
                    raw_res = normalize_turkish_key(row.get(res_col)) if res_col else ""
                    is_reserved = "REZERV" in raw_res

                    length_cm = safe_float(row.get(len_col)) if len_col else 0.0
                    height_cm = safe_float(row.get(height_col)) if height_col else 0.0
                    thick_str = format_thickness_en(row.get(thick_col)) if thick_col else "2 cm"
                    surf_str = format_surface_en(row.get(surf_col)) if surf_col else "Polished"

                    # 1. Taş Bazlı Toplama
                    if norm_stone not in stones_agg:
                        stones_agg[norm_stone] = {
                            "stone_name": norm_stone,
                            "stock_m2": 0.0,
                            "stock_slabs": 0,
                            "reserved_m2": 0.0,
                            "reserved_slabs": 0,
                            "production_m2": 0.0,
                            "production_slabs": 0,
                            "active_lots": [],
                            "all_lots": [],
                            "surfaces": set(),
                            "thicknesses": set(),
                            "locations": set(),
                            "lengths": [],
                            "heights": [],
                        }

                    st_entry = stones_agg[norm_stone]
                    st_entry["production_m2"] += m2_val
                    st_entry["production_slabs"] += slab_count
                    st_entry["surfaces"].add(surf_str)
                    st_entry["thicknesses"].add(thick_str)

                    if is_in_stock:
                        st_entry["stock_m2"] += m2_val
                        st_entry["stock_slabs"] += slab_count
                        if is_reserved:
                            st_entry["reserved_m2"] += m2_val
                            st_entry["reserved_slabs"] += slab_count
                        if raw_loc and raw_loc not in ["NAN", "YUKLENDI"]:
                            st_entry["locations"].add(raw_loc)
                        if length_cm > 0 and height_cm > 0:
                            st_entry["lengths"].append(length_cm)
                            st_entry["heights"].append(height_cm)

                    if lot_code and lot_code not in ["NAN", "NONE"]:
                        if lot_code not in st_entry["all_lots"]:
                            st_entry["all_lots"].append(lot_code)
                        if is_in_stock and lot_code not in st_entry["active_lots"]:
                            st_entry["active_lots"].append(lot_code)

                        # 2. Lot (Fabrika Kodu / Blok No) Bazlı Toplama
                        if lot_code not in lots_agg:
                            lots_agg[lot_code] = {
                                "lot": lot_code,
                                "stone_name": norm_stone,
                                "stock_m2": 0.0,
                                "stock_slabs": 0,
                                "reserved_m2": 0.0,
                                "reserved_slabs": 0,
                                "production_m2": 0.0,
                                "production_slabs": 0,
                                "surfaces": set(),
                                "thicknesses": set(),
                                "locations": set(),
                                "lengths": [],
                                "heights": [],
                            }
                        lt_entry = lots_agg[lot_code]
                        lt_entry["production_m2"] += m2_val
                        lt_entry["production_slabs"] += slab_count
                        lt_entry["surfaces"].add(surf_str)
                        lt_entry["thicknesses"].add(thick_str)

                        if is_in_stock:
                            lt_entry["stock_m2"] += m2_val
                            lt_entry["stock_slabs"] += slab_count
                            if is_reserved:
                                lt_entry["reserved_m2"] += m2_val
                                lt_entry["reserved_slabs"] += slab_count
                            if raw_loc and raw_loc not in ["NAN", "YUKLENDI"]:
                                lt_entry["locations"].add(raw_loc)
                            if length_cm > 0 and height_cm > 0:
                                lt_entry["lengths"].append(length_cm)
                                lt_entry["heights"].append(height_cm)

            # Set ve boyut özetlerini son haline getir
            for dct in (stones_agg, lots_agg):
                for k, v in dct.items():
                    v["surface_str"] = " / ".join(sorted(v["surfaces"])) if v["surfaces"] else "Polished"
                    v["thickness_str"] = " / ".join(sorted(v["thicknesses"])) if v["thicknesses"] else "2 cm"
                    v["location_str"] = ", ".join(sorted(v["locations"])) if v["locations"] else "Main Warehouse"
                    if v["lengths"] and v["heights"]:
                        avg_l = int(round(sum(v["lengths"]) / len(v["lengths"])))
                        avg_h = int(round(sum(v["heights"]) / len(v["heights"])))
                        v["dimensions_str"] = f"{avg_l} x {avg_h} cm"
                    else:
                        v["dimensions_str"] = "Standard Slab Sizes"
                    # Gereksiz geçici listeleri temizle
                    v.pop("surfaces", None)
                    v.pop("thicknesses", None)
                    v.pop("locations", None)
                    v.pop("lengths", None)
                    v.pop("heights", None)

            self.stones_summary = stones_agg
            self.lots_index = lots_agg
            self.last_loaded_time = datetime.now().strftime("%d.%m.%Y %H:%M")
            logger.info(
                f"Excel başarıyla yüklendi: {len(self.stones_summary)} tekil taş, "
                f"{len(self.lots_index)} tekil lot/blok."
            )
            return len(self.stones_summary), len(self.lots_index)
        except Exception as e:
            logger.error(f"Excel okuma hatası: {e}")
            return 0, 0

    @staticmethod
    def _find_col(col_map: Dict[Any, str], keywords: List[str]) -> Optional[Any]:
        for kw in keywords:
            for orig_col, norm_col in col_map.items():
                if kw == norm_col or kw in norm_col:
                    return orig_col
        return None

    def find_item_data(self, query: str) -> Dict[str, Any]:
        """
        Verilen sorgu bir Lot koduysa ('A-283') lot verisini, taş adıysa ('SILVER ROOTS')
        taş verisini döner. Bulamazsa kısmi eşleşme arar.
        """
        norm_lot = normalize_lot_code(query)
        norm_stone = normalize_turkish_key(query)

        if norm_lot in self.lots_index:
            item = self.lots_index[norm_lot].copy()
            item["display_title"] = f"{item['stone_name']} (Lot: {norm_lot})"
            item["query_key"] = norm_lot
            return item

        if norm_stone in self.stones_summary:
            item = self.stones_summary[norm_stone].copy()
            item["display_title"] = norm_stone
            item["query_key"] = norm_stone
            return item

        # Kısmi arama
        for k, v in self.stones_summary.items():
            if norm_stone in k or k in norm_stone:
                item = v.copy()
                item["display_title"] = k
                item["query_key"] = k
                return item

        for k, v in self.lots_index.items():
            if norm_lot in k:
                item = v.copy()
                item["display_title"] = f"{v['stone_name']} (Lot: {k})"
                item["query_key"] = k
                return item

        return {
            "display_title": norm_stone or query.upper(),
            "query_key": norm_lot or norm_stone,
            "stone_name": norm_stone or query.upper(),
            "stock_m2": 0.0,
            "stock_slabs": 0,
            "production_m2": 0.0,
            "production_slabs": 0,
            "dimensions_str": "Contact for Sizes",
            "thickness_str": "2 cm",
            "surface_str": "Polished",
            "location_str": "Factory Showroom",
            "active_lots": [],
        }

    def get_context_snapshot(self) -> str:
        lines = [
            f"=== ARSLANLAR MARBLE CANLI STOK VERİTABANI (Son Güncelleme: {self.last_loaded_time or 'Bekleniyor'}) ==="
        ]
        if self.stones_summary:
            sorted_stones = sorted(
                self.stones_summary.values(),
                key=lambda x: (x["stock_m2"], x["production_m2"]),
                reverse=True,
            )
            lines.append("EN ÇOK STOĞU BULUNAN İLK 25 TAŞ CİNSİ (Aktif Stok vs Toplam Üretim):")
            for idx, item in enumerate(sorted_stones[:25], start=1):
                active_lots_preview = ", ".join(item["active_lots"][:6]) if item.get("active_lots") else "Yok"
                lines.append(
                    f"{idx}. {item['stone_name']} -> "
                    f"AKTİF STOK: {item['stock_m2']:.2f} m² ({item['stock_slabs']} Plaka) | "
                    f"TOPLAM ÜRETİM: {item['production_m2']:.2f} m² ({item['production_slabs']} Plaka) | "
                    f"Ort. Ebat: {item['dimensions_str']} | Yüzey: {item['surface_str']} | "
                    f"Aktif Lotlar: {active_lots_preview}"
                )
        else:
            lines.append("UYARI: Excel stok dosyası henüz yüklenmedi.")

        lines.append("\n=== ÖZEL LOT VE GÖRSEL DNA KATALOĞU ===")
        for code, info in self.stone_dna_db.items():
            lot_data = self.lots_index.get(code) or self.stones_summary.get(info["stone_name"], {})
            st_m2 = lot_data.get("stock_m2", 0.0)
            st_slabs = lot_data.get("stock_slabs", 0)
            pr_m2 = lot_data.get("production_m2", 0.0)
            pr_slabs = lot_data.get("production_slabs", 0)
            loc = lot_data.get("location_str", "-")
            dims = lot_data.get("dimensions_str", "-")
            lines.append(
                f"- {code} ({info.get('stone_name')}): {info.get('visual_dna')} "
                f"[Aktif Stok: {st_m2:.2f} m² ({st_slabs} Plaka) | Toplam Üretim: {pr_m2:.2f} m² ({pr_slabs} Plaka) | "
                f"Ebat: {dims} | Konum: {loc}]"
            )
        return "\n".join(lines)


# =====================================================================
# 6. %100 İNGİLİZCE, FOTOĞRAFLI PREMIUM PDF SUNUM MOTORU
# =====================================================================
class PDFPresentationEngine:
    """
    Müşterilere gönderilecek %100 İngilizce, gerçek plaka fotoğraflı,
    gereksiz sözel açıklamalardan arındırılmış kurumsal PDF katalog üreticisi.
    """

    def __init__(self, inventory: InventoryEngine, drive: DriveExplorer):
        self.inventory = inventory
        self.drive = drive

    @staticmethod
    def _create_fallback_slab_image() -> bytes:
        """Drive klasöründe henüz fotoğraf yoksa şık bir mermer plaka görseli üretir."""
        img = Image.new("RGB", (600, 400), color=(42, 48, 56))
        draw = ImageDraw.Draw(img)
        # Zarif çerçeve ve mermer dokusu çizgileri
        draw.rectangle([12, 12, 588, 388], outline=(197, 160, 89), width=2)
        draw.text((195, 185), "ARSLANLAR MARBLE SLAB", fill=(220, 225, 230))
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=90)
        return buf.getvalue()

    def generate_pdf(
        self,
        queries: List[str],
        collection_folder_name: Optional[str] = None,
    ) -> Tuple[bytes, str]:
        """
        Belirtilen taş/lot listesi VEYA Drive'daki bir ana koleksiyon klasörü
        ('Arslanlar Collection', 'Arslanlar Series') için İngilizce PDF üretir.
        """
        cards_to_render: List[Dict[str, Any]] = []
        subtitle_en = "EXECUTIVE STOCK & SLAB PRESENTATION"

        # 1. Eğer 'Arslanlar Collection' veya 'Arslanlar Series' gibi bir klasör taraması istendiyse
        if collection_folder_name:
            parent_id, parent_link = self.drive.find_folder(collection_folder_name)
            subtitle_en = f"{normalize_turkish_key(collection_folder_name)} - CATALOG PRESENTATION"
            if parent_id:
                subfolders = self.drive.list_child_folders(parent_id, max_folders=15)
                for sf in subfolders:
                    item_data = self.inventory.find_item_data(sf["name"])
                    imgs = self.drive.get_folder_images(sf["id"], max_images=1)
                    cards_to_render.append(
                        {
                            "title": item_data["display_title"],
                            "data": item_data,
                            "drive_link": sf["link"],
                            "image_bytes": imgs[0] if imgs else None,
                        }
                    )

        # 2. Taş veya Lot bazlı PDF istendiyse (veya klasör boş döndüyse)
        if not cards_to_render:
            target_queries = queries if queries else ["A-283", "SILVER ROOTS"]
            for q in target_queries[:12]:
                item_data = self.inventory.find_item_data(q)
                q_key = item_data["query_key"]
                folder_id, folder_link = self.drive.find_folder(q_key)
                if not folder_id and item_data.get("stone_name"):
                    folder_id, folder_link = self.drive.find_folder(item_data["stone_name"])
                imgs = self.drive.get_folder_images(folder_id, max_images=1) if folder_id else []
                cards_to_render.append(
                    {
                        "title": item_data["display_title"],
                        "data": item_data,
                        "drive_link": folder_link or DEFAULT_DRIVE_LINK,
                        "image_bytes": imgs[0] if imgs else None,
                    }
                )

        # 3. ReportLab ile %100 İngilizce Lüks PDF İnşası
        pdf_buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            pdf_buffer,
            pagesize=A4,
            rightMargin=32,
            leftMargin=32,
            topMargin=32,
            bottomMargin=32,
        )

        styles = getSampleStyleSheet()
        NAVY = colors.HexColor("#14213D")
        GOLD = colors.HexColor("#C5A059")
        SLATE = colors.HexColor("#2B2D42")
        LIGHT_BG = colors.HexColor("#F8F9FA")
        BORDER_GRAY = colors.HexColor("#DEE2E6")

        title_style = ParagraphStyle(
            "ArslanlarTitle",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=21,
            leading=25,
            textColor=NAVY,
            alignment=1,
            spaceAfter=3,
        )
        subtitle_style = ParagraphStyle(
            "ArslanlarSub",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=10,
            leading=13,
            textColor=GOLD,
            alignment=1,
            spaceAfter=10,
        )
        card_heading_style = ParagraphStyle(
            "CardHeading",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=12,
            leading=16,
            textColor=NAVY,
            spaceBefore=6,
            spaceAfter=6,
        )
        th_style = ParagraphStyle(
            "TH",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=8.5,
            leading=11,
            textColor=colors.whitesmoke,
            alignment=0,
        )
        td_label_style = ParagraphStyle(
            "TDLabel",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=8.5,
            leading=11,
            textColor=SLATE,
            alignment=0,
        )
        td_val_style = ParagraphStyle(
            "TDVal",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=8.5,
            leading=11,
            textColor=SLATE,
            alignment=0,
        )
        footer_style = ParagraphStyle(
            "FooterEN",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=8,
            leading=11,
            textColor=colors.HexColor("#6C757D"),
            alignment=1,
        )

        story: List[Any] = []

        # Kurumsal Başlık Bloğu
        story.append(Paragraph("ARSLANLAR MARBLE", title_style))
        story.append(Paragraph(subtitle_en, subtitle_style))
        story.append(HRFlowable(width="100%", thickness=1.2, color=GOLD, spaceAfter=12))

        for idx, card in enumerate(cards_to_render, start=1):
            d = card["data"]
            stock_m2 = d.get("stock_m2", 0.0)
            stock_slabs = d.get("stock_slabs", 0)
            prod_m2 = d.get("production_m2", 0.0)
            prod_slabs = d.get("production_slabs", 0)

            stock_str = (
                f"{stock_m2:,.2f} m2 ({stock_slabs:,} Slabs)"
                if stock_m2 > 0
                else "Available on Order / Check Production"
            )
            prod_str = (
                f"{prod_m2:,.2f} m2 ({prod_slabs:,} Slabs)"
                if prod_m2 > 0
                else "Certified Factory Selection"
            )
            dims_str = d.get("dimensions_str", "Standard Slab Sizes")
            surf_thick_str = f"{d.get('surface_str', 'Polished')} | {d.get('thickness_str', '2 cm')}"
            lots_list = d.get("active_lots") or d.get("all_lots") or []
            lots_str = ", ".join(lots_list[:6]) if lots_list else str(d.get("query_key", "-"))
            drive_url = card["drive_link"]

            # Görseli hazırla (Sol Taraf: 175x125 pt)
            raw_img = card["image_bytes"] or self._create_fallback_slab_image()
            try:
                pil_im = Image.open(io.BytesIO(raw_img)).convert("RGB")
                img_io = io.BytesIO()
                pil_im.save(img_io, format="JPEG", quality=88)
                img_io.seek(0)
                rl_img = RLImage(img_io, width=175, height=122)
            except Exception:
                fb_io = io.BytesIO(self._create_fallback_slab_image())
                rl_img = RLImage(fb_io, width=175, height=122)

            # Sağ Taraf: Saf Teknik Veri Tablosu (%100 İngilizce, Doku/Karakter Yok)
            spec_rows = [
                [
                    Paragraph("Specification", th_style),
                    Paragraph("Technical Data", th_style),
                ],
                [
                    Paragraph("Available Stock", td_label_style),
                    Paragraph(stock_str, td_val_style),
                ],
                [
                    Paragraph("Total Production", td_label_style),
                    Paragraph(prod_str, td_val_style),
                ],
                [
                    Paragraph("Avg. Slab Size", td_label_style),
                    Paragraph(dims_str, td_val_style),
                ],
                [
                    Paragraph("Surface & Thickness", td_label_style),
                    Paragraph(surf_thick_str, td_val_style),
                ],
                [
                    Paragraph("Active Lots / Code", td_label_style),
                    Paragraph(lots_str, td_val_style),
                ],
                [
                    Paragraph("High-Res Slab Photos", td_label_style),
                    Paragraph(
                        f"<link href='{drive_url}' color='#0056b3'><u>Click to Open Google Drive Folder</u></link>",
                        td_val_style,
                    ),
                ],
            ]

            spec_table = Table(spec_rows, colWidths=[115, 225])
            spec_table.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                        ("BACKGROUND", (0, 1), (-1, -1), LIGHT_BG),
                        ("GRID", (0, 0), (-1, -1), 0.5, BORDER_GRAY),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("TOPPADDING", (0, 0), (-1, -1), 4),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                        ("LEFTPADDING", (0, 0), (-1, -1), 6),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ]
                )
            )

            master_row = Table([[rl_img, spec_table]], colWidths=[185, 345])
            master_row.setStyle(
                TableStyle(
                    [
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("ALIGN", (0, 0), (0, 0), "CENTER"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 0),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ]
                )
            )

            block_elements = [
                Paragraph(f"{idx}. {card['title']}", card_heading_style),
                master_row,
                Spacer(1, 14),
            ]
            story.append(KeepTogether(block_elements))

        story.append(Spacer(1, 8))
        story.append(HRFlowable(width="100%", thickness=0.6, color=BORDER_GRAY, spaceAfter=6))
        story.append(
            Paragraph(
                "<b>ARSLANLAR MARBLE</b> | International Sales & Export Department | Izmir, Turkey",
                footer_style,
            )
        )

        doc.build(story)
        pdf_buffer.seek(0)
        clean_filename = f"Arslanlar_Marble_Presentation_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
        return pdf_buffer.read(), clean_filename


# =====================================================================
# 7. YAPAY ZEKA ORKESTRATÖRÜ (DOĞRUDAN REST API - AQ. AUTH KEY UYUMLU)
# =====================================================================
class ArgonBrain:
    def __init__(self, api_key: str, inventory: InventoryEngine, drive: DriveExplorer):
        self.api_key = api_key
        self.inventory = inventory
        self.drive = drive
        self.chat_memory: Dict[int, deque] = {}
        self.endpoint = (
            "https://generativelanguage.googleapis.com/v1beta/models/"
            "gemini-flash-latest:generateContent"
        )

    def add_memory(self, chat_id: int, sender: str, text: str) -> None:
        if chat_id not in self.chat_memory:
            self.chat_memory[chat_id] = deque(maxlen=12)
        self.chat_memory[chat_id].append(f"{sender}: {text}")

    def get_memory_text(self, chat_id: int) -> str:
        if chat_id not in self.chat_memory or not self.chat_memory[chat_id]:
            return "Önceki sohbet kaydı yok."
        return "\n".join(self.chat_memory[chat_id])

    def _call_gemini_rest(
        self, prompt_text: str, image_bytes: Optional[bytes] = None
    ) -> Dict[str, Any]:
        parts: List[Dict[str, Any]] = [{"text": prompt_text}]

        if image_bytes:
            try:
                pil_img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
                buf = io.BytesIO()
                pil_img.save(buf, format="JPEG", quality=85)
                b64_data = base64.b64encode(buf.getvalue()).decode("utf-8")
                parts.append(
                    {
                        "inline_data": {
                            "mime_type": "image/jpeg",
                            "data": b64_data,
                        }
                    }
                )
            except Exception as e:
                logger.warning(f"Görsel base64'e çevrilemedi: {e}")

        payload = {
            "contents": [{"parts": parts}],
            "generationConfig": {
                "temperature": 0.2,
                "responseMimeType": "application/json",
            },
        }

        headers = {
            "Content-Type": "application/json",
            "X-goog-api-key": self.api_key,
        }

        resp = requests.post(
            self.endpoint, headers=headers, json=payload, timeout=40
        )
        if resp.status_code != 200:
            raise RuntimeError(f"HTTP {resp.status_code}: {resp.text}")

        data = resp.json()
        raw_text = data["candidates"][0]["content"]["parts"][0]["text"].strip()
        raw_text = re.sub(r"^```json\s*", "", raw_text)
        raw_text = re.sub(r"\s*```$", "", raw_text)
        return json.loads(raw_text)

    async def evaluate_and_respond(
        self,
        chat_id: int,
        user_name: str,
        user_text: str,
        is_group: bool,
        bot_mentioned: bool,
        image_bytes: Optional[bytes] = None,
    ) -> Dict[str, Any]:
        catalog_context = self.inventory.get_context_snapshot()
        recent_history = self.get_memory_text(chat_id)

        full_prompt = f"""
Sen Arslanlar Marble (Arslanlar Mermer) şirketinin kıdemli dijital satış ve operasyon uzmanısın.
Adın Argon. Grupta veya özel sohbette gerçek bir insan uzman gibi davranırsın. Komutlarla değil, doğal konuşmayla çalışırsın.

GÖREV VE DAVRANIŞ KURALLARIN:
1. SUSMA YETENEĞİ (GATEKEEPER):
   - Eğer mesaj bir grup sohbetinden geliyorsa (`is_group=True`) ve insanlar kendi aralarında sohbet ediyorsa, "Bravo", "Teşekkürler", "Tamam", "😂😂😂", "Harika", "Eline sağlık" gibi ifadeler yazmışsa ve mermer/stok/lot/fiyat/görsel/PDF sunum/nakliye ile ilgili hiçbir soru yoksa KESİNLİKLE ARAYA GİRME (`"should_reply": false` döndür).
   - Mesajda bir taş adı, lot kodu (A-283 vb.), stok/metraj sorusu, PDF sunum talebi, konteyner hesabı, Drive linki talebi varsa VEYA kullanıcı bir mermer fotoğrafı göndermişse `"should_reply": true` döndür.

2. ÇOK DİLLİ UZMANLIK (TR / UK / EN):
   - Kullanıcı Türkçe yazdıysa Türkçe, Ukraynaca (Українська) yazdıysa akıcı Ukraynaca, İngilizce yazdıysa İngilizce yanıt ver.
   - Ancak PDF sunumu istendiğinde PDF dosyasının içi her zaman %100 İngilizce olarak üretilir.

3. AKTİF STOK vs TOPLAM ÜRETİM AYRIMI:
   - Tabloda her taş ve lot için hem `AKTİF STOK` (şu an fabrikada mevcut olan) hem de `TOPLAM ÜRETİM` (yüklenenler dahil genel toplam) ayrı ayrı verilmiştir. Sorularda her iki veriyi de net olarak belirt.
   - Asla "SERKAN ÇELİK" gibi kişi/tedarikçi isimlerini mermer çeşidi olarak listeleme.

4. PDF SUNUM TALEPLERİ (`generate_pdf`):
   - Eğer kullanıcı "PDF sunum hazırla", "PDF katalog at", "sunum yap", "send PDF presentation" gibi bir istekte bulunursa `"generate_pdf": true` yap.
   - Eğer kullanıcı belirli bir ana Drive koleksiyon klasörünü taramayı istediyse (Örn: "Arslanlar Collection klasörünü tarayarak PDF yap" veya "Arslanlar Series klasöründen PDF sunum hazırla"), `"collection_folder"` alanına tam olarak o klasör adını yaz (`"Arslanlar Collection"` veya `"Arslanlar Series"`).
   - Eğer belirli taşlar veya lotlar için PDF istediyse (Örn: "A-283 ve Silver Roots için PDF sunum at"), `"matched_queries"` listesine bu kodları/taşları yaz (`["A-283", "SILVER ROOTS"]`).

5. GÖRSEL DNA EŞLEŞTİRME (FOTOĞRAF GELDİĞİNDE):
   - Kullanıcı bir mermer fotoğrafı gönderdiyse: Taşın zemin rengini ve damar yapısını analiz et, katalogdan en çok benzeyen 1-3 taşı/lotu `matched_queries` içine yaz ve `"send_photos": true` yap.

6. B2B KONTEYNER VE TONAJ HESABI:
   - 2cm mermer plaka ≈ 54 kg/m², 3cm mermer plaka ≈ 81 kg/m². 1 adet 20 DC konteyner maksimum ~27.000 kg (27 ton) alır (2cm'den ~500 m², 3cm'den ~330 m²).

SON SOHBET GEÇMİŞİ:
{recent_history}

{catalog_context}

GELEN MESAJ BİLGİSİ:
- Grup Sohbeti mi: {is_group}
- Bot Etiketlendi mi: {bot_mentioned}
- Gönderen: {user_name}
- Mesaj: {user_text or '[Görsel Gönderildi]'}

ÇIKTI FORMATI (SADECE GEÇERLİ BİR JSON DÖNDÜR):
{{
  "should_reply": true,
  "language": "tr",
  "reply_text": "Kullanıcıya verilecek doğal, profesyonel ve nokta atışı yanıt...",
  "matched_queries": ["A-283", "SILVER ROOTS"],
  "send_photos": false,
  "generate_pdf": false,
  "collection_folder": null
}}
"""
        try:
            result = await asyncio.to_thread(
                self._call_gemini_rest, full_prompt, image_bytes
            )
            return result
        except Exception as e:
            logger.error(f"Gemini REST analiz hatası: {e}")
            return {
                "should_reply": False,
                "language": "tr",
                "reply_text": "",
                "matched_queries": [],
                "send_photos": False,
                "generate_pdf": False,
                "collection_folder": None,
            }


# =====================================================================
# 8. TELEGRAM HANDLER'LARI (EXCEL AUTO-SYNC + ALBÜM + PDF SUNUM)
# =====================================================================
drive_explorer = DriveExplorer(GOOGLE_SA_JSON)
inventory_engine = InventoryEngine(EXCEL_FILE_PATH, drive_explorer)
pdf_engine = PDFPresentationEngine(inventory_engine, drive_explorer)
argon_brain = ArgonBrain(GEMINI_API_KEY, inventory_engine, drive_explorer)

PROCESSED_MESSAGES: deque = deque(maxlen=500)


async def handle_document_upload(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    """
    Telegram üzerinden doğrudan gönderilen Excel (.xlsx / .xls) stok listesini
    yakalar, 'stok.xlsx' olarak kaydeder ve hafızayı anında günceller.
    """
    message = update.effective_message
    if not message or not message.document:
        return

    doc = message.document
    file_name = doc.file_name or ""
    if not file_name.lower().endswith((".xlsx", ".xls")):
        return

    status_msg = await message.reply_text(
        f"📥 **{file_name}** alındı, plaka ve lot verileri analiz ediliyor...",
        parse_mode=ParseMode.MARKDOWN,
    )

    try:
        tg_file = await context.bot.get_file(doc.file_id)
        await tg_file.download_to_drive(EXCEL_FILE_PATH)

        stones_count, lots_count = await asyncio.to_thread(
            inventory_engine.reload_excel
        )

        await status_msg.edit_text(
            f"✅ **Stok Listesi Başarıyla Güncellendi!** 🚀\n\n"
            f"📊 **Tekilleştirilmiş Mermer Çeşidi:** `{stones_count}`\n"
            f"📦 **Aktif Lot / Fabrika Kodu:** `{lots_count}`\n"
            f"🕒 **Güncelleme Saati:** `{inventory_engine.last_loaded_time}`\n\n"
            f"Artık tüm stok sorgularını ve İngilizce PDF sunumlarını bu güncel veriyle yapabilirsiniz.",
            parse_mode=ParseMode.MARKDOWN,
        )
    except Exception as e:
        logger.error(f"Excel yükleme hatası: {e}")
        await status_msg.edit_text(f"❌ Dosya işlenirken bir hata oluştu: {e}")


async def handle_incoming_message(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    message = update.effective_message
    if not message or not update.effective_chat:
        return

    # 1. Çift Mesaj Kilidi
    unique_msg_key = f"{update.effective_chat.id}:{message.message_id}"
    if unique_msg_key in PROCESSED_MESSAGES:
        return
    PROCESSED_MESSAGES.append(unique_msg_key)

    if message.from_user and message.from_user.is_bot:
        return

    chat_id = update.effective_chat.id
    is_group = update.effective_chat.type in ["group", "supergroup"]
    user_name = message.from_user.first_name if message.from_user else "Kullanıcı"
    user_text = (message.text or message.caption or "").strip()

    # Manuel Drive linki öğretme kısayolu ("A-283 linki: https://...")
    link_teach_match = re.search(
        r"([A-Za-z0-9-_]+)\s+linki\s*:\s*(https?://\S+)", user_text, re.IGNORECASE
    )
    if link_teach_match:
        lot_key = normalize_lot_code(link_teach_match.group(1))
        url_val = link_teach_match.group(2)
        drive_explorer.save_custom_link(lot_key, url_val)
        await message.reply_text(
            f"✅ **{lot_key}** klasör bağlantısı hafızaya kaydedildi:\n{url_val}",
            parse_mode=ParseMode.MARKDOWN,
        )
        return

    bot_username = context.bot.username or ""
    bot_mentioned = False
    if bot_username and f"@{bot_username.lower()}" in user_text.lower():
        bot_mentioned = True
    if (
        message.reply_to_message
        and message.reply_to_message.from_user
        and message.reply_to_message.from_user.id == context.bot.id
    ):
        bot_mentioned = True

    image_bytes: Optional[bytes] = None
    if message.photo:
        try:
            photo_file = await message.photo[-1].get_file()
            byte_arr = await photo_file.download_as_bytearray()
            image_bytes = bytes(byte_arr)
        except Exception as e:
            logger.error(f"Telegram fotoğraf indirme hatası: {e}")

    if not user_text and not image_bytes:
        return

    argon_brain.add_memory(
        chat_id, user_name, user_text if user_text else "[Mermer Fotoğrafı Gönderdi]"
    )

    decision = await argon_brain.evaluate_and_respond(
        chat_id=chat_id,
        user_name=user_name,
        user_text=user_text,
        is_group=is_group,
        bot_mentioned=bot_mentioned,
        image_bytes=image_bytes,
    )

    if not decision.get("should_reply", False):
        return

    reply_text = decision.get("reply_text", "").strip()
    matched_queries: List[str] = decision.get("matched_queries", [])
    send_photos: bool = decision.get("send_photos", False) or bool(image_bytes)
    want_pdf: bool = decision.get("generate_pdf", False)
    collection_folder: Optional[str] = decision.get("collection_folder")

    # Kullanıcı mesajında açıkça 'pdf' ve ('sunum' / 'katalog' / 'hazırla' / 'presentation') geçtiyse garantiye al
    if re.search(r"\bpdf\b", user_text, re.IGNORECASE):
        want_pdf = True

    if not reply_text:
        return

    # 2. EĞER PDF SUNUM İSTENDİYSE: İngilizce Fotoğraflı PDF Üret ve Gönder
    if want_pdf:
        await context.bot.send_chat_action(
            chat_id=chat_id, action=ChatAction.UPLOAD_DOCUMENT
        )
        try:
            pdf_bytes, pdf_filename = await asyncio.to_thread(
                pdf_engine.generate_pdf,
                matched_queries,
                collection_folder,
            )
            argon_brain.add_memory(chat_id, "Argon (Bot)", f"[PDF Sunum Gönderildi] {reply_text}")
            await message.reply_document(
                document=io.BytesIO(pdf_bytes),
                filename=pdf_filename,
                caption=reply_text[:1020],
            )
            return
        except Exception as pdf_err:
            logger.error(f"PDF oluşturma hatası: {pdf_err}")

    # 3. STANDART YANIT VEYA DRIVE FOTOĞRAF ALBÜMÜ
    drive_links_section: List[str] = []
    media_group_items: List[InputMediaPhoto] = []
    newly_cached_keys: List[Tuple[str, int]] = []
    last_folder_link: Optional[str] = None

    for query in matched_queries[:3]:
        norm_q = normalize_lot_code(query)
        folder_id, folder_link = drive_explorer.find_folder(norm_q)

        if folder_link:
            last_folder_link = folder_link
            drive_links_section.append(f"📂 {norm_q} Drive: {folder_link}")
        elif norm_q in inventory_engine.lots_index:
            drive_links_section.append(
                f"📂 {norm_q} Genel Drive: {DEFAULT_DRIVE_LINK}"
            )

        if send_photos:
            cached_file_ids = drive_explorer.visual_cache.get(norm_q, [])
            if cached_file_ids:
                for fid in cached_file_ids[:2]:
                    media_group_items.append(InputMediaPhoto(media=fid))
            elif folder_id:
                await context.bot.send_chat_action(
                    chat_id=chat_id, action=ChatAction.UPLOAD_PHOTO
                )
                raw_imgs = await asyncio.to_thread(
                    drive_explorer.get_folder_images, folder_id, 2
                )
                for img_b in raw_imgs:
                    media_group_items.append(InputMediaPhoto(media=img_b))
                    newly_cached_keys.append((norm_q, len(media_group_items) - 1))

    if drive_links_section:
        links_block = "\n".join(drive_links_section)
        if not last_folder_link or last_folder_link not in reply_text:
            reply_text = f"{reply_text}\n\n{links_block}"

    argon_brain.add_memory(chat_id, "Argon (Bot)", reply_text)

    if media_group_items:
        try:
            short_caption = (
                reply_text[:1020] + "..." if len(reply_text) > 1024 else reply_text
            )
            media_group_items[0] = InputMediaPhoto(
                media=media_group_items[0].media,
                caption=short_caption,
            )
            sent_messages = await message.reply_media_group(
                media=media_group_items[:10]
            )

            for norm_q, idx in newly_cached_keys:
                if idx < len(sent_messages) and sent_messages[idx].photo:
                    fid = sent_messages[idx].photo[-1].file_id
                    drive_explorer.visual_cache.setdefault(norm_q, [])
                    if fid not in drive_explorer.visual_cache[norm_q]:
                        drive_explorer.visual_cache[norm_q].append(fid)
            save_json_file(VISUAL_CACHE_FILE, drive_explorer.visual_cache)

            if len(reply_text) > 1024:
                await message.reply_text(reply_text)
            return
        except Exception as e:
            logger.error(f"Albüm gönderme hatası, metne dönülüyor: {e}")

    await message.reply_text(reply_text, disable_web_page_preview=False)


# =====================================================================
# 9. ZERO-BLOCKING RENDER BAŞLATICI (15 DK ASILI KALMAYI ÖNLER)
# =====================================================================
flask_app = Flask(__name__)


@flask_app.route("/")
def health_check():
    return (
        f"Arslanlar Marble Argon Master Bot is Live! "
        f"(Loaded Stones: {len(inventory_engine.stones_summary)}, "
        f"Lots: {len(inventory_engine.lots_index)})",
        200,
    )


def run_flask():
    port = int(os.environ.get("PORT", 10000))
    flask_app.run(host="0.0.0.0", port=port, use_reloader=False)


def background_initial_load():
    """Render portunun anında açılması için ilk Excel okumasını arka planda yapar."""
    time.sleep(1)
    inventory_engine.reload_excel()


def main():
    if not TELEGRAM_TOKEN:
        raise ValueError("TELEGRAM_TOKEN ortam değişkeni eksik!")

    # 1. Önce Flask Web Sunucusunu başlat (Render 1 saniyede 'Live' görsün)
    threading.Thread(target=run_flask, daemon=True).start()

    # 2. Arka planda varsa mevcut Excel'i yükle (Ana akışı asla bloklamaz)
    threading.Thread(target=background_initial_load, daemon=True).start()

    application = ApplicationBuilder().token(TELEGRAM_TOKEN).build()

    # Telegram'dan atılan Excel (.xlsx / .xls) dosyalarını yakalayan handler
    application.add_handler(
        MessageHandler(filters.Document.ALL, handle_document_upload)
    )

    # Metin ve fotoğraf mesajlarını yakalayan ana orkestratör
    application.add_handler(
        MessageHandler(
            (filters.TEXT | filters.PHOTO) & (~filters.COMMAND),
            handle_incoming_message,
        )
    )

    logger.info("🚀 Arslanlar Marble Argon Master Ultimate başlatılıyor...")
    application.run_polling(drop_pending_updates=True, allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
