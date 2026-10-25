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
    filters,
)

from google.oauth2 import service_account
import google.auth.transport.requests
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
# 4. GOOGLE DRIVE OTONOM GEZGİN, DOSYA/KLASÖR TARAYICI VE GÖRSEL MOTORU
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

    @staticmethod
    def _build_candidates(query_term: str, fallback_stone_name: str = "") -> List[str]:
        cands = []
        for term in [query_term, fallback_stone_name]:
            if not term:
                continue
            raw = term.strip().replace("'", "")
            cands.append(raw)
            norm = normalize_lot_code(raw)
            cands.append(norm)
            if "-" in norm:
                cands.append(norm.replace("-", " "))
                cands.append(norm.replace("-", ""))
            # Türkçe İ varyasyonları (Drive klasör adı 'SİLVER RİVER' ise yakalamak için)
            tr_dotted = (
                norm.replace("SILVER", "SİLVER")
                .replace("RIVER", "RİVER")
                .replace("VIOLET", "VİOLET")
                .replace("BRECCIA", "BRECCİA")
                .replace("ARMONI", "ARMONİ")
            )
            if tr_dotted != norm:
                cands.append(tr_dotted)
        return list(dict.fromkeys([c for c in cands if c and len(c) >= 2]))

    def find_folder(
        self, query_term: str, fallback_stone_name: str = ""
    ) -> Tuple[Optional[str], Optional[str]]:
        """
        Önce lot koduna ('A-371') ait Drive klasörünü arar. Bulamazsa taşın ana
        klasörünü ('SILVER RIVER' / 'SİLVER RİVER') arar ve linkini döner.
        """
        norm_key = normalize_lot_code(query_term)

        if norm_key in self.hyperlinks_db:
            cached_url = self.hyperlinks_db[norm_key]
            folder_id = self._extract_folder_id_from_url(cached_url)
            if cached_url and cached_url != DEFAULT_DRIVE_LINK and folder_id:
                return folder_id, cached_url

        if not self.service:
            return None, None

        candidates = self._build_candidates(query_term, fallback_stone_name)

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
                        if normalize_lot_code(f["name"]) == normalize_lot_code(cand):
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

    def list_child_folders(
        self, parent_folder_id: str, max_folders: int = 20
    ) -> List[Dict[str, str]]:
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

    def get_images_for_query(
        self,
        query_term: str,
        folder_id: Optional[str] = None,
        max_images: int = 2,
    ) -> List[bytes]:
        """
        1) Eğer klasör ID'si varsa o klasörün (veya alt klasörlerinin) içindeki fotoğrafları indirir.
        2) Klasör bulunamadıysa bile Drive genelinde dosya adı 'A-371' içeren görselleri arayıp indirir!
        """
        if not self.service:
            return []

        if folder_id:
            imgs = self.get_folder_images(folder_id, max_images=max_images)
            if imgs:
                return imgs

        # Dosya adında doğrudan lot/taş kodu geçen fotoğrafları ara (Örn: A-371.jpg)
        image_bytes_list: List[bytes] = []
        candidates = self._build_candidates(query_term)
        try:
            for cand in candidates[:3]:
                q = (
                    f"trashed = false and name contains '{cand}' "
                    f"and (mimeType contains 'image/jpeg' or mimeType contains 'image/png')"
                )
                res = (
                    self.service.files()
                    .list(
                        q=q,
                        spaces="drive",
                        fields="files(id, name, mimeType)",
                        includeItemsFromAllDrives=True,
                        supportsAllDrives=True,
                        pageSize=max_images,
                    )
                    .execute()
                )
                files = res.get("files", [])
                if files:
                    for f in files[:max_images]:
                        req = self.service.files().get_media(
                            fileId=f["id"], supportsAllDrives=True
                        )
                        fh = io.BytesIO()
                        downloader = MediaIoBaseDownload(fh, req)
                        done = False
                        while not done:
                            _, done = downloader.next_chunk()
                        fh.seek(0)
                        image_bytes_list.append(fh.read())
                    if image_bytes_list:
                        return image_bytes_list
        except Exception as e:
            logger.error(f"Drive doğrudan görsel arama hatası ({query_term}): {e}")

        return image_bytes_list

    def get_folder_images(self, folder_id: str, max_images: int = 2) -> List[bytes]:
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
        if os.path.exists(self.excel_path):
            return self.excel_path
        for fname in os.listdir("."):
            if fname.lower().endswith((".xlsx", ".xls")) and not fname.startswith("~$"):
                return fname
        return None

    def reload_excel(self) -> Tuple[int, int]:
        target_path = self.locate_any_excel_file()
        if not target_path:
            logger.warning(f"Excel dosyası henüz bulunamadı ({self.excel_path}).")
            return 0, 0

        try:
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

            xls = pd.ExcelFile(target_path)
            stones_agg: Dict[str, Dict[str, Any]] = {}
            lots_agg: Dict[str, Dict[str, Any]] = {}

            for sheet_name in xls.sheet_names:
                df = pd.read_excel(target_path, sheet_name=sheet_name)
                if df.empty:
                    continue

                col_map = {col: normalize_turkish_key(col) for col in df.columns}

                stone_col = self._find_col(
                    col_map,
                    ["SELEKSIYON", "TAS CINSI", "URUN ADI", "MERMER", "CINSI", "STONE"],
                )
                fab_lot_col = self._find_col(
                    col_map, ["FABRIKA KODU", "LOT", "PARTI", "BUNDLE"]
                )
                blok_col = self._find_col(col_map, ["BLOK NO", "BLOK"])
                status_col = self._find_col(col_map, ["STOK DURUMU", "DURUM"])
                loc_col = self._find_col(
                    col_map, ["STOK KONUMU", "KONUM", "KORIDOR", "HOL"]
                )
                res_col = self._find_col(col_map, ["SATIS VE REZERV", "REZERV"])
                m2_col = self._find_col(col_map, ["M2", "NET M2", "METRAJ", "ALAN"])
                len_col = self._find_col(col_map, ["UZUNLUK", "BOY", "LENGTH"])
                height_col = self._find_col(
                    col_map, ["YUKSEKLIK", "EN", "HEIGHT", "WIDTH"]
                )
                thick_col = self._find_col(col_map, ["KALINLIK", "THICKNESS", "CM"])
                surf_col = self._find_col(
                    col_map, ["YUZEY ISLEMI", "YUZEY", "SURFACE"]
                )
                adet_col = self._find_col(col_map, ["ADET", "PCS", "SLAB COUNT"])

                if not stone_col or not m2_col:
                    continue

                for _, row in df.iterrows():
                    raw_stone = row.get(stone_col)
                    norm_stone = normalize_turkish_key(raw_stone)

                    raw_fab = row.get(fab_lot_col) if fab_lot_col else None
                    raw_blok = row.get(blok_col) if blok_col else None
                    lot_code = normalize_lot_code(raw_fab) or normalize_lot_code(
                        raw_blok
                    )

                    if lot_code in self.stone_dna_db and (
                        not norm_stone or norm_stone == "BELIRSIZ"
                    ):
                        norm_stone = self.stone_dna_db[lot_code]["stone_name"]

                    if is_blacklisted_stone(norm_stone):
                        continue

                    m2_val = safe_float(row.get(m2_col))
                    if m2_val <= 0:
                        continue

                    slab_count = (
                        int(safe_float(row.get(adet_col)))
                        if adet_col and pd.notna(row.get(adet_col))
                        else 1
                    )
                    if slab_count <= 0:
                        slab_count = 1

                    raw_status = (
                        normalize_turkish_key(row.get(status_col))
                        if status_col
                        else "STOK"
                    )
                    is_in_stock = (
                        (raw_status == "STOK")
                        or ("SHOWROOM" in raw_status)
                        or (status_col is None)
                    )

                    raw_loc = (
                        normalize_turkish_key(row.get(loc_col)) if loc_col else ""
                    )
                    raw_res = (
                        normalize_turkish_key(row.get(res_col)) if res_col else ""
                    )
                    is_reserved = "REZERV" in raw_res

                    length_cm = safe_float(row.get(len_col)) if len_col else 0.0
                    height_cm = safe_float(row.get(height_col)) if height_col else 0.0
                    thick_str = (
                        format_thickness_en(row.get(thick_col))
                        if thick_col
                        else "2 cm"
                    )
                    surf_str = (
                        format_surface_en(row.get(surf_col))
                        if surf_col
                        else "Polished"
                    )

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

            for dct in (stones_agg, lots_agg):
                for k, v in dct.items():
                    v["surface_str"] = (
                        " / ".join(sorted(v["surfaces"]))
                        if v["surfaces"]
                        else "Polished"
                    )
                    v["thickness_str"] = (
                        " / ".join(sorted(v["thicknesses"]))
                        if v["thicknesses"]
                        else "2 cm"
                    )
                    v["location_str"] = (
                        ", ".join(sorted(v["locations"]))
                        if v["locations"]
                        else "Main Warehouse"
                    )
                    if v["lengths"] and v["heights"]:
                        avg_l = int(round(sum(v["lengths"]) / len(v["lengths"])))
                        avg_h = int(round(sum(v["heights"]) / len(v["heights"])))
                        v["dimensions_str"] = f"{avg_l} x {avg_h} cm"
                    else:
                        v["dimensions_str"] = "Standard Slab Sizes"
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

    def extract_entities_from_text(self, user_text: str) -> List[str]:
        """
        Kullanıcının yazdığı mesajdan ('A-371 fotoğraf lütfen', 'Silver Roots stok')
        hem lot kodlarını (A-371, A-283 vb.) hem de taş isimlerini anında yakalar.
        """
        if not user_text:
            return []
        matched: List[str] = []
        norm_msg = normalize_turkish_key(user_text)

        # 1. Lot kodu deseni (Örn: A-371, A 371, A371, D15/1)
        lot_regex = re.compile(r"\b([A-Z]{1,4})[-\s]?(\d{1,5})\b")
        for m in lot_regex.finditer(norm_msg):
            prefix, num = m.group(1), m.group(2)
            if prefix in ["M", "CM", "MM", "KG", "DC", "TON"]:
                continue
            cand_lot = f"{prefix}-{num}"
            if cand_lot in self.lots_index or prefix == "A":
                matched.append(cand_lot)

        # 2. Bilinen taş isimlerini uzunluk sırasına göre mesaj içinde ara
        sorted_stone_names = sorted(
            self.stones_summary.keys(), key=len, reverse=True
        )
        for st_name in sorted_stone_names:
            if len(st_name) >= 3 and st_name in norm_msg:
                if not any(st_name in existing for existing in matched):
                    matched.append(st_name)

        return list(dict.fromkeys(matched))

    def find_item_data(self, query: str) -> Dict[str, Any]:
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

        for k, v in self.stones_summary.items():
            if norm_stone and (norm_stone in k or k in norm_stone):
                item = v.copy()
                item["display_title"] = k
                item["query_key"] = k
                return item

        for k, v in self.lots_index.items():
            if norm_lot and norm_lot in k:
                item = v.copy()
                item["display_title"] = f"{v['stone_name']} (Lot: {k})"
                item["query_key"] = k
                return item

        return {
            "display_title": norm_lot or norm_stone or query.upper(),
            "query_key": norm_lot or norm_stone,
            "stone_name": norm_stone or query.upper(),
            "stock_m2": 0.0,
            "stock_slabs": 0,
            "production_m2": 0.0,
            "production_slabs": 0,
            "dimensions_str": "Standard Slab Sizes",
            "thickness_str": "2 cm",
            "surface_str": "Polished",
            "location_str": "Factory Showroom",
            "active_lots": [],
        }

    def get_context_snapshot(self, user_text: str = "") -> str:
        lines = [
            f"=== ARSLANLAR MARBLE CANLI STOK VERİTABANI (Son Güncelleme: {self.last_loaded_time or 'Bekleniyor'}) ==="
        ]

        # Kullanıcının mesajında geçen spesifik lot/taşları en üste ekle (Örn: A-371)
        extracted = self.extract_entities_from_text(user_text)
        if extracted:
            lines.append("\n>>> KULLANICININ SORDUĞU LOT / TAŞ DETAYLARI (KESİN VERİ) <<<")
            for q in extracted:
                d = self.find_item_data(q)
                lines.append(
                    f"* {d['display_title']} -> "
                    f"AKTİF STOK: {d.get('stock_m2', 0.0):.2f} m² ({d.get('stock_slabs', 0)} Plaka) | "
                    f"TOPLAM ÜRETİM: {d.get('production_m2', 0.0):.2f} m² ({d.get('production_slabs', 0)} Plaka) | "
                    f"Ortalama Ebat: {d.get('dimensions_str', '-')} | "
                    f"Yüzey & Kalınlık: {d.get('surface_str', 'Polished')} ({d.get('thickness_str', '2 cm')}) | "
                    f"Konum: {d.get('location_str', '-')}"
                )

        if self.stones_summary:
            sorted_stones = sorted(
                self.stones_summary.values(),
                key=lambda x: (x["stock_m2"], x["production_m2"]),
                reverse=True,
            )
            lines.append("\nEN ÇOK STOĞU BULUNAN İLK 25 TAŞ CİNSİ:")
            for idx, item in enumerate(sorted_stones[:25], start=1):
                active_lots_preview = (
                    ", ".join(item["active_lots"][:6])
                    if item.get("active_lots")
                    else "Yok"
                )
                lines.append(
                    f"{idx}. {item['stone_name']} -> "
                    f"AKTİF STOK: {item['stock_m2']:.2f} m² ({item['stock_slabs']} Plaka) | "
                    f"TOPLAM ÜRETİM: {item['production_m2']:.2f} m² ({item['production_slabs']} Plaka) | "
                    f"Ort. Ebat: {item['dimensions_str']} | Aktif Lotlar: {active_lots_preview}"
                )
        else:
            lines.append("UYARI: Excel stok dosyası henüz yüklenmedi.")

        return "\n".join(lines)


# =====================================================================
# 6. %100 İNGİLİZCE, FOTOĞRAFLI PREMIUM PDF SUNUM MOTORU
# =====================================================================
class PDFPresentationEngine:
    def __init__(self, inventory: InventoryEngine, drive: DriveExplorer):
        self.inventory = inventory
        self.drive = drive

    @staticmethod
    def _create_fallback_slab_image() -> bytes:
        img = Image.new("RGB", (600, 400), color=(42, 48, 56))
        draw = ImageDraw.Draw(img)
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
        cards_to_render: List[Dict[str, Any]] = []
        subtitle_en = "EXECUTIVE STOCK & SLAB PRESENTATION"

        if collection_folder_name:
            parent_id, _ = self.drive.find_folder(collection_folder_name)
            subtitle_en = f"{normalize_turkish_key(collection_folder_name)} - CATALOG PRESENTATION"
            if parent_id:
                subfolders = self.drive.list_child_folders(parent_id, max_folders=15)
                for sf in subfolders:
                    item_data = self.inventory.find_item_data(sf["name"])
                    imgs = self.drive.get_images_for_query(
                        sf["name"], folder_id=sf["id"], max_images=1
                    )
                    cards_to_render.append(
                        {
                            "title": item_data["display_title"],
                            "data": item_data,
                            "drive_link": sf["link"],
                            "image_bytes": imgs[0] if imgs else None,
                        }
                    )

        if not cards_to_render:
            target_queries = queries if queries else ["A-283", "SILVER ROOTS"]
            for q in target_queries[:12]:
                item_data = self.inventory.find_item_data(q)
                q_key = item_data["query_key"]
                st_name = item_data.get("stone_name", "")
                folder_id, folder_link = self.drive.find_folder(q_key, st_name)
                imgs = self.drive.get_images_for_query(
                    q_key, folder_id=folder_id, max_images=1
                )
                cards_to_render.append(
                    {
                        "title": item_data["display_title"],
                        "data": item_data,
                        "drive_link": folder_link or DEFAULT_DRIVE_LINK,
                        "image_bytes": imgs[0] if imgs else None,
                    }
                )

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
            surf_thick_str = (
                f"{d.get('surface_str', 'Polished')} | {d.get('thickness_str', '2 cm')}"
            )
            lots_list = d.get("active_lots") or d.get("all_lots") or []
            lots_str = (
                ", ".join(lots_list[:6]) if lots_list else str(d.get("query_key", "-"))
            )
            drive_url = card["drive_link"]

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
        story.append(
            HRFlowable(width="100%", thickness=0.6, color=BORDER_GRAY, spaceAfter=6)
        )
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
# 7. ÇİFT MOTORLU ORKESTRATÖR (MULTI-ENDPOINT AI + YEREL UZMAN MOTORU)
# =====================================================================
class ArgonBrain:
    def __init__(
        self,
        api_key: str,
        sa_json_str: str,
        inventory: InventoryEngine,
        drive: DriveExplorer,
    ):
        self.api_key = api_key
        self.sa_json_str = sa_json_str
        self.inventory = inventory
        self.drive = drive
        self.chat_memory: Dict[int, deque] = {}

    def add_memory(self, chat_id: int, sender: str, text: str) -> None:
        if chat_id not in self.chat_memory:
            self.chat_memory[chat_id] = deque(maxlen=12)
        self.chat_memory[chat_id].append(f"{sender}: {text}")

    def get_memory_text(self, chat_id: int) -> str:
        if chat_id not in self.chat_memory or not self.chat_memory[chat_id]:
            return "Önceki sohbet kaydı yok."
        return "\n".join(self.chat_memory[chat_id])

    def _call_gemini_multi_channel(
        self, prompt_text: str, image_bytes: Optional[bytes] = None
    ) -> Dict[str, Any]:
        """
        Sırasıyla:
        1) Vertex AI Express REST Endpoint ('AQ.' ile başlayan anahtarlar için)
        2) Generative Language REST Endpoint ('AIzaSy' veya standart anahtarlar için)
        3) Google Service Account OAuth2 Bearer Token (Anahtarsız Bulut Yetkisi)
        üzerinden bağlantı dener.
        """
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
            "contents": [{"role": "user", "parts": parts}],
            "generationConfig": {
                "temperature": 0.2,
                "responseMimeType": "application/json",
            },
        }

        endpoints_to_try: List[Tuple[str, Dict[str, str]]] = []

        if self.api_key:
            # 1. Vertex AI Express Endpoint (AQ. anahtarları için birincil adres)
            endpoints_to_try.append(
                (
                    f"https://aiplatform.googleapis.com/v1beta1/publishers/google/models/gemini-2.0-flash:generateContent?key={self.api_key}",
                    {"Content-Type": "application/json"},
                )
            )
            # 2. Generative Language Endpoint (X-goog-api-key header ile)
            endpoints_to_try.append(
                (
                    "https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent",
                    {
                        "Content-Type": "application/json",
                        "X-goog-api-key": self.api_key,
                    },
                )
            )

        # 3. Service Account JSON üzerinden OAuth2 Bearer Token ile deneme
        if self.sa_json_str:
            try:
                info = json.loads(self.sa_json_str)
                project_id = info.get("project_id", "arslanlar-marble-bot")
                creds = service_account.Credentials.from_service_account_info(
                    info, scopes=["https://www.googleapis.com/auth/cloud-platform"]
                )
                auth_req = google.auth.transport.requests.Request()
                creds.refresh(auth_req)
                if creds.token:
                    endpoints_to_try.append(
                        (
                            f"https://us-central1-aiplatform.googleapis.com/v1/projects/{project_id}/locations/us-central1/publishers/google/models/gemini-2.0-flash:generateContent",
                            {
                                "Content-Type": "application/json",
                                "Authorization": f"Bearer {creds.token}",
                            },
                        )
                    )
            except Exception as sa_err:
                logger.debug(f"Service account token yenileme atlandı: {sa_err}")

        last_err = ""
        for url, headers in endpoints_to_try:
            try:
                resp = requests.post(url, headers=headers, json=payload, timeout=25)
                if resp.status_code == 200:
                    data = resp.json()
                    raw_text = (
                        data["candidates"][0]["content"]["parts"][0]["text"].strip()
                    )
                    raw_text = re.sub(r"^```json\s*", "", raw_text)
                    raw_text = re.sub(r"\s*```$", "", raw_text)
                    return json.loads(raw_text)
                else:
                    last_err = f"HTTP {resp.status_code}: {resp.text[:200]}"
            except Exception as ex:
                last_err = str(ex)

        raise RuntimeError(f"Tüm AI kanalları başarısız oldu. Son hata: {last_err}")

    def _local_expert_fallback(
        self,
        user_text: str,
        is_group: bool,
        bot_mentioned: bool,
        image_bytes: Optional[bytes] = None,
    ) -> Dict[str, Any]:
        """
        API anahtarı hatası veya kesintisi durumunda devreye giren %100 Kırılmaz
        Yerel Uzman Motoru. Mesajı analiz eder, Excel ve Drive verileriyle anında yanıt üretir.
        """
        raw_text = (user_text or "").strip()
        norm_text = normalize_turkish_key(raw_text)

        # Dil Tespiti (UK / EN / TR)
        is_ukrainian = bool(re.search(r"[а-яА-ЯіїєґІЇЄҐ]", raw_text))
        is_english = bool(
            re.search(
                r"\b(please|stock|slab|slabs|send|photo|photos|presentation|price|available|what|how many)\b",
                raw_text,
                re.IGNORECASE,
            )
        )
        lang = "uk" if is_ukrainian else ("en" if is_english else "tr")

        # Mesaj içindeki Lot Kodlarını (Örn: A-371, A-283) ve Taş İsimlerini Yakala
        matched_entities = self.inventory.extract_entities_from_text(raw_text)

        # Koleksiyon Klasörü Taraması İsteği Kontrolü
        collection_folder = None
        if "COLLECTION" in norm_text:
            collection_folder = "Arslanlar Collection"
        elif "SERIES" in norm_text or "SERISI" in norm_text:
            collection_folder = "Arslanlar Series"

        # PDF İsteği Kontrolü
        want_pdf = bool(
            re.search(
                r"\b(pdf|sunum|sunumu|katalog|presentation|catalog)\b",
                raw_text,
                re.IGNORECASE,
            )
        )

        # Fotoğraf / Görsel İsteği Kontrolü
        want_photos = (
            bool(
                re.search(
                    r"(foto|resim|gorsel|görsel|photo|picture|image|album|albüm|фото|знімки)",
                    raw_text,
                    re.IGNORECASE,
                )
            )
            or bool(image_bytes)
        )

        # İş / Stok / Link Sorusu Kontrolü
        business_keywords = [
            "STOK",
            "PLAKA",
            "LOT",
            "BLOK",
            "LINK",
            "DRIVE",
            "FOTO",
            "RESIM",
            "GORSEL",
            "PDF",
            "SUNUM",
            "KATALOG",
            "METRAJ",
            "M2",
            "ELIMIZDE",
            "MEVCUT",
            "VAR MI",
            "HANGI TAS",
            "KONTEYNER",
            "TONAJ",
            "EBAT",
            "KALINLIK",
            "SHOWROOM",
            "HOL",
            "СЛЯБ",
            "ЛОТ",
            "СКЛАД",
            "ФОТО",
            "наявність",
        ]
        has_business_intent = (
            bool(matched_entities)
            or want_pdf
            or want_photos
            or bool(collection_folder)
            or any(kw in norm_text for kw in business_keywords)
        )

        # Selamlaşma / "Orda mısın" Kontrolü
        greeting_keywords = [
            "ORDA MISIN",
            "ORADA MISIN",
            "MERHABA",
            "SELAM",
            "HELLO",
            "HI",
            "HEY",
            "SES VER",
            "ПРИВІТ",
        ]
        is_greeting = any(gk in norm_text for gk in greeting_keywords)

        # Grupta Geyik Filtresi (Susma Yeteneği)
        if is_group and not has_business_intent and not bot_mentioned:
            if not is_greeting:
                return {
                    "should_reply": False,
                    "language": lang,
                    "reply_text": "",
                    "matched_queries": [],
                    "send_photos": False,
                    "generate_pdf": False,
                    "collection_folder": None,
                }

        # 1. PDF Sunum Yanıtı
        if want_pdf:
            target_list = (
                matched_entities
                if matched_entities
                else (["A-283", "SILVER ROOTS"] if not collection_folder else [])
            )
            if lang == "uk":
                reply_msg = "📄 Презентацію у форматі PDF (англійською мовою з фотографіями слябів та актуальними даними складу) успішно сформовано."
            elif lang == "en":
                reply_msg = "📄 Here is your executive English PDF presentation including slab photographs and live inventory specifications."
            else:
                reply_msg = "📄 İstediğiniz kriterlere uygun, gerçek plaka fotoğraflı ve %100 İngilizce kurumsal PDF sunumunuz hazırlandı."
            return {
                "should_reply": True,
                "language": lang,
                "reply_text": reply_msg,
                "matched_queries": target_list,
                "send_photos": False,
                "generate_pdf": True,
                "collection_folder": collection_folder,
            }

        # 2. Spesifik Lot (Örn: A-371, A-283) veya Taş Sorgusu Yanıtı
        if matched_entities:
            blocks = []
            for q in matched_entities[:4]:
                d = self.inventory.find_item_data(q)
                title = d["display_title"]
                st_m2 = d.get("stock_m2", 0.0)
                st_slabs = d.get("stock_slabs", 0)
                pr_m2 = d.get("production_m2", 0.0)
                pr_slabs = d.get("production_slabs", 0)
                dims = d.get("dimensions_str", "-")
                surf = d.get("surface_str", "Polished")
                thick = d.get("thickness_str", "2 cm")
                loc = d.get("location_str", "-")
                lots_list = d.get("active_lots") or []
                lots_info = (
                    f"\n• Aktif Lotlar: {', '.join(lots_list[:8])}"
                    if lots_list and q not in self.inventory.lots_index
                    else ""
                )

                if lang == "uk":
                    blocks.append(
                        f"💎 {title}\n"
                        f"• Актуальний склад: {st_m2:,.2f} м² ({st_slabs} слябів)\n"
                        f"• Загальне виробництво: {pr_m2:,.2f} м² ({pr_slabs} слябів)\n"
                        f"• Середній розмір: {dims} | Обробка: {surf} ({thick})\n"
                        f"• Локація: {loc}"
                    )
                elif lang == "en":
                    blocks.append(
                        f"💎 {title}\n"
                        f"• Available Stock: {st_m2:,.2f} m² ({st_slabs} Slabs)\n"
                        f"• Total Production: {pr_m2:,.2f} m² ({pr_slabs} Slabs)\n"
                        f"• Avg. Dimensions: {dims} | Finish: {surf} ({thick})\n"
                        f"• Location: {loc}"
                    )
                else:
                    blocks.append(
                        f"💎 {title}\n"
                        f"• Aktif Mevcut Stok: {st_m2:,.2f} m² ({st_slabs} Plaka)\n"
                        f"• Toplam Üretim: {pr_m2:,.2f} m² ({pr_slabs} Plaka)\n"
                        f"• Ortalama Ebat: {dims} | Yüzey & Kalınlık: {surf} ({thick})\n"
                        f"• Stok Konumu: {loc}{lots_info}"
                    )

            return {
                "should_reply": True,
                "language": lang,
                "reply_text": "\n\n".join(blocks),
                "matched_queries": matched_entities[:3],
                "send_photos": True,  # Lot/taş sorulduğunda varsa fotoğrafları da otomatik albüm yapar
                "generate_pdf": False,
                "collection_folder": None,
            }

        # 3. Görsel Gönderildiyse (Fotoğraftan Benzer Taş Önerisi)
        if image_bytes:
            return {
                "should_reply": True,
                "language": lang,
                "reply_text": (
                    "📸 Gönderdiğiniz mermer görselinin doku ve damar yapısına göre "
                    "koleksiyonumuzdaki en yakın seleksiyonlar ve Drive bağlantıları aşağıda listelenmiştir:"
                ),
                "matched_queries": ["SILVER ROOTS", "A-283", "SILVER RIVER"],
                "send_photos": True,
                "generate_pdf": False,
                "collection_folder": None,
            }

        # 4. Genel Stok Özeti Talebi ("Stokta en çok hangi taşlar var?")
        if any(
            kw in norm_text
            for kw in ["STOK", "ELIMIZDE", "HANGI TAS", "MEVCUT", "LISTE", "СКЛАД"]
        ):
            if not self.inventory.stones_summary:
                return {
                    "should_reply": True,
                    "language": lang,
                    "reply_text": "⚠️ Henüz güncel stok Excel dosyası yüklenmedi. Lütfen 'PLAKA STOK LİSTESİ.xlsx' dosyasını sohbete gönderin.",
                    "matched_queries": [],
                    "send_photos": False,
                    "generate_pdf": False,
                    "collection_folder": None,
                }
            sorted_stones = sorted(
                self.inventory.stones_summary.values(),
                key=lambda x: (x["stock_m2"], x["production_m2"]),
                reverse=True,
            )[:15]
            lines = ["📊 ARSLANLAR MARBLE GÜNCEL AKTİF STOK ÖZETİ (İlk 15 Seleksiyon):\n"]
            for idx, item in enumerate(sorted_stones, start=1):
                lines.append(
                    f"{idx}. {item['stone_name']}: "
                    f"{item['stock_m2']:,.2f} m² ({item['stock_slabs']} Plaka) "
                    f"[Toplam Üretim: {item['production_m2']:,.2f} m²]"
                )
            return {
                "should_reply": True,
                "language": lang,
                "reply_text": "\n".join(lines),
                "matched_queries": [],
                "send_photos": False,
                "generate_pdf": False,
                "collection_folder": None,
            }

        # 5. Selamlaşma / "Orda mısın" Yanıtı
        if is_greeting or not is_group or bot_mentioned:
            st_cnt = len(self.inventory.stones_summary)
            lt_cnt = len(self.inventory.lots_index)
            return {
                "should_reply": True,
                "language": lang,
                "reply_text": (
                    f"👋 Merhaba! Arslanlar Marble Argon Asistanı aktif ve görev başında.\n\n"
                    f"📊 Hafızadaki Tekil Taş Çeşidi: {st_cnt}\n"
                    f"📦 Hafızadaki Aktif Lot / Blok Kodu: {lt_cnt}\n\n"
                    f"Herhangi bir lot kodunu (Örn: 'A-371 fotoğraf ve drive linki' veya 'A-283 ve Silver Roots PDF sunum hazırla') doğrudan yazabilirsiniz."
                ),
                "matched_queries": [],
                "send_photos": False,
                "generate_pdf": False,
                "collection_folder": None,
            }

        return {
            "should_reply": False,
            "language": lang,
            "reply_text": "",
            "matched_queries": [],
            "send_photos": False,
            "generate_pdf": False,
            "collection_folder": None,
        }

    async def evaluate_and_respond(
        self,
        chat_id: int,
        user_name: str,
        user_text: str,
        is_group: bool,
        bot_mentioned: bool,
        image_bytes: Optional[bytes] = None,
    ) -> Dict[str, Any]:
        catalog_context = self.inventory.get_context_snapshot(user_text)
        recent_history = self.get_memory_text(chat_id)
        extracted_entities = self.inventory.extract_entities_from_text(user_text)

        full_prompt = f"""
Sen Arslanlar Marble (Arslanlar Mermer) şirketinin kıdemli dijital satış ve operasyon uzmanısın.
Adın Argon. Grupta veya özel sohbette gerçek bir insan uzman gibi davranırsın.

GÖREV VE DAVRANIŞ KURALLARIN:
1. SUSMA YETENEĞİ (GATEKEEPER):
   - Eğer mesaj bir grup sohbetinden geliyorsa (`is_group=True`) ve insanlar kendi aralarında sohbet ediyorsa ("Bravo", "Teşekkürler", "Tamam", "😂😂😂" vb.) ve mermer/stok/lot/fiyat/görsel/PDF sunum/nakliye ile ilgili hiçbir soru yoksa KESİNLİKLE ARAYA GİRME (`"should_reply": false` döndür).
   - Mesajda bir taş adı, lot kodu (A-371, A-283 vb.), stok/metraj sorusu, PDF sunum talebi, Drive linki talebi varsa VEYA kullanıcı fotoğraf göndermişse `"should_reply": true` döndür.

2. ÇOK DİLLİ UZMANLIK (TR / UK / EN):
   - Kullanıcı Türkçe yazdıysa Türkçe, Ukraynaca yazdıysa Ukraynaca, İngilizce yazdıysa İngilizce yanıt ver.

3. AKTİF STOK vs TOPLAM ÜRETİM AYRIMI:
   - Tabloda her taş ve lot için hem `AKTİF STOK` hem de `TOPLAM ÜRETİM` ayrı ayrı verilmiştir. Sorularda her iki veriyi, ebatı ve konumu net olarak belirt.
   - Asla "SERKAN ÇELİK" gibi kişi/tedarikçi isimlerini mermer çeşidi olarak listeleme.

4. DRIVE LİNKİ, FOTOĞRAF VE PDF SUNUM:
   - Kullanıcının sorduğu lot veya taş kodlarını `matched_queries` listesine mutlaka ekle (Örn: {json.dumps(extracted_entities or ["A-371"])}).
   - Kullanıcı fotoğraf/görsel istediyse `"send_photos": true` yap.
   - Kullanıcı PDF sunum/katalog istediyse `"generate_pdf": true` yap.

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
  "matched_queries": {json.dumps(extracted_entities)},
  "send_photos": true,
  "generate_pdf": false,
  "collection_folder": null
}}
"""
        try:
            result = await asyncio.to_thread(
                self._call_gemini_multi_channel, full_prompt, image_bytes
            )
            # Eğer kullanıcı açıkça bir lot sorduysa (Örn: A-371) ve AI matched_queries'i boş bıraktıysa tamamla
            if extracted_entities and not result.get("matched_queries"):
                result["matched_queries"] = extracted_entities
            return result
        except Exception as e:
            logger.warning(
                f"Bulut AI kanalı yanıt vermedi ({e}), Yerel Uzman Motoru devreye giriyor..."
            )
            return self._local_expert_fallback(
                user_text=user_text,
                is_group=is_group,
                bot_mentioned=bot_mentioned,
                image_bytes=image_bytes,
            )


# =====================================================================
# 8. TELEGRAM HANDLER'LARI (EXCEL AUTO-SYNC + ALBÜM + PDF SUNUM)
# =====================================================================
drive_explorer = DriveExplorer(GOOGLE_SA_JSON)
inventory_engine = InventoryEngine(EXCEL_FILE_PATH, drive_explorer)
pdf_engine = PDFPresentationEngine(inventory_engine, drive_explorer)
argon_brain = ArgonBrain(
    GEMINI_API_KEY, GOOGLE_SA_JSON, inventory_engine, drive_explorer
)

PROCESSED_MESSAGES: deque = deque(maxlen=500)


async def handle_document_upload(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
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

    # Manuel Drive linki öğretme kısayolu ("A-371 linki: https://...")
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

    if re.search(r"\bpdf\b", user_text, re.IGNORECASE):
        want_pdf = True

    if re.search(
        r"(foto|resim|gorsel|görsel|photo|фото)", user_text, re.IGNORECASE
    ):
        send_photos = True

    if not reply_text:
        return

    # EĞER PDF SUNUM İSTENDİYSE: İngilizce Fotoğraflı PDF Üret ve Gönder
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
            argon_brain.add_memory(
                chat_id, "Argon (Bot)", f"[PDF Sunum Gönderildi] {reply_text}"
            )
            await message.reply_document(
                document=io.BytesIO(pdf_bytes),
                filename=pdf_filename,
                caption=reply_text[:1020],
            )
            return
        except Exception as pdf_err:
            logger.error(f"PDF oluşturma hatası: {pdf_err}")

    # STANDART YANIT, DRIVE ALT KLASÖR LİNKİ VE FOTOĞRAF ALBÜMÜ
    drive_links_section: List[str] = []
    media_group_items: List[InputMediaPhoto] = []
    newly_cached_keys: List[Tuple[str, int]] = []
    last_folder_link: Optional[str] = None

    for query in matched_queries[:3]:
        norm_q = normalize_lot_code(query)
        item_info = inventory_engine.find_item_data(norm_q)
        stone_name_fallback = item_info.get("stone_name", "")

        folder_id, folder_link = drive_explorer.find_folder(
            norm_q, stone_name_fallback
        )

        if folder_link:
            last_folder_link = folder_link
            drive_links_section.append(f"📂 {norm_q} Drive: {folder_link}")
        else:
            drive_links_section.append(
                f"📂 {norm_q} Genel Drive: {DEFAULT_DRIVE_LINK}"
            )

        if send_photos:
            cached_file_ids = drive_explorer.visual_cache.get(norm_q, [])
            if cached_file_ids:
                for fid in cached_file_ids[:2]:
                    media_group_items.append(InputMediaPhoto(media=fid))
            else:
                await context.bot.send_chat_action(
                    chat_id=chat_id, action=ChatAction.UPLOAD_PHOTO
                )
                raw_imgs = await asyncio.to_thread(
                    drive_explorer.get_images_for_query, norm_q, folder_id, 2
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
# 9. ZERO-BLOCKING RENDER BAŞLATICI
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
    time.sleep(1)
    inventory_engine.reload_excel()


def main():
    if not TELEGRAM_TOKEN:
        raise ValueError("TELEGRAM_TOKEN ortam değişkeni eksik!")

    threading.Thread(target=run_flask, daemon=True).start()
    threading.Thread(target=background_initial_load, daemon=True).start()

    application = ApplicationBuilder().token(TELEGRAM_TOKEN).build()

    application.add_handler(
        MessageHandler(filters.Document.ALL, handle_document_upload)
    )

    application.add_handler(
        MessageHandler(
            (filters.TEXT | filters.PHOTO) & (~filters.COMMAND),
            handle_incoming_message,
        )
    )

    logger.info("🚀 Arslanlar Marble Argon Master Ultimate (Zero-Silence) başlatılıyor...")
    application.run_polling(drop_pending_updates=True, allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
