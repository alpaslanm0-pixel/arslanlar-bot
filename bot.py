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
# 1. LOGLAMA VE ORTAM DEĞİŞKENLERİ
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
# 2. YARDIMCI FONKSİYONLAR VE NORMALİZASYON MOTORU
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
# 3. GOOGLE DRIVE OTONOM GEZGİN VE GÖRSEL MOTORU
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

    def get_images_structured(
        self, folder_id: Optional[str], query_term: str
    ) -> Tuple[Optional[bytes], Optional[bytes]]:
        """
        Klasörden veya Drive'dan 2 adet görsel seçer:
        1. Görsel: Standart plaka fotoğrafı.
        2. Görsel: Bookmatch fotoğrafı (dosya veya alt klasör adında 'bookmatch' geçen).
        Eğer bookmatch yoksa 2. plaka fotoğrafını döner.
        """
        if not self.service:
            return None, None

        all_image_files = []

        try:
            search_folders = [folder_id] if folder_id else []
            # Eğer klasör ID verilmediyse veya içindekiler azsa ada göre klasör bul
            if not folder_id:
                fid, _ = self.find_folder(query_term)
                if fid:
                    search_folders.append(fid)

            for fid in search_folders:
                # Ana klasördeki görseller
                q = (
                    f"'{fid}' in parents and trashed = false "
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
                        pageSize=10,
                    )
                    .execute()
                )
                all_image_files.extend(res.get("files", []))

                # Alt klasörlerdeki görseller (Bookmatch klasörleri dahil)
                sub_folders = self.list_child_folders(fid, max_folders=5)
                for sf in sub_folders:
                    sq = (
                        f"'{sf['id']}' in parents and trashed = false "
                        f"and (mimeType contains 'image/jpeg' or mimeType contains 'image/png')"
                    )
                    sres = (
                        self.service.files()
                        .list(
                            q=sq,
                            spaces="drive",
                            fields="files(id, name, mimeType)",
                            includeItemsFromAllDrives=True,
                            supportsAllDrives=True,
                            pageSize=10,
                        )
                        .execute()
                    )
                    # Alt klasör adını dosya objesine ekle (bookmatch tespiti için)
                    for sf_img in sres.get("files", []):
                        sf_img["parent_folder_name"] = sf["name"]
                        all_image_files.append(sf_img)

            # Eğer klasörden bulunamadıysa doğrudan genel Drive'da ada göre ara
            if not all_image_files:
                for cand in self._build_candidates(query_term)[:2]:
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
                            pageSize=10,
                        )
                        .execute()
                    )
                    all_image_files.extend(res.get("files", []))

            if not all_image_files:
                return None, None

            # Bookmatch tespiti
            bookmatch_file = None
            standard_files = []

            for f in all_image_files:
                fname_lower = f.get("name", "").lower()
                parent_name_lower = f.get("parent_folder_name", "").lower()
                if "bookmatch" in fname_lower or "bookmatch" in parent_name_lower or "bm" in fname_lower:
                    if not bookmatch_file:
                        bookmatch_file = f
                else:
                    standard_files.append(f)

            # 1. Görsel (Plaka)
            img1_bytes = None
            if standard_files:
                img1_bytes = self._download_file_bytes(standard_files[0]["id"])
            elif bookmatch_file:
                img1_bytes = self._download_file_bytes(bookmatch_file["id"])

            # 2. Görsel (Bookmatch veya 2. Plaka)
            img2_bytes = None
            if bookmatch_file and bookmatch_file != (standard_files[0] if standard_files else None):
                img2_bytes = self._download_file_bytes(bookmatch_file["id"])
            elif len(standard_files) > 1:
                img2_bytes = self._download_file_bytes(standard_files[1]["id"])
            elif img1_bytes:
                img2_bytes = img1_bytes # İkinci yoksa aynı plaka fotoğrafı

            return img1_bytes, img2_bytes

        except Exception as e:
            logger.error(f"Görsel arama/indirme hatası ({query_term}): {e}")
            return None, None

    def _download_file_bytes(self, file_id: str) -> Optional[bytes]:
        try:
            req = self.service.files().get_media(fileId=file_id, supportsAllDrives=True)
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
    def _extract_folder_id_from_url(url: str) -> Optional[str]:
        if not url:
            return None
        m = re.search(r"/folders/([a-zA-Z0-9_-]+)", url)
        if m:
            return m.group(1)
        m2 = re.search(r"id=([a-zA-Z0-9_-]+)", url)
        return m2.group(1) if m2 else None


# =====================================================================
# 4. AKILLI EXCEL STOK VE ÇEKİ LİSTESİ MOTORU
# =====================================================================
class InventoryEngine:
    def __init__(self, excel_path: str, drive_explorer: DriveExplorer):
        self.excel_path = excel_path
        self.drive = drive_explorer
        self.stones_summary: Dict[str, Dict[str, Any]] = {}
        self.lots_index: Dict[str, Dict[str, Any]] = {}
        self.raw_stock_rows: List[Dict[str, Any]] = []
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
            logger.warning(f"Excel dosyası bulunamadı ({self.excel_path}).")
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
            all_raw_rows: List[Dict[str, Any]] = []

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
                plaka_no_col = self._find_col(col_map, ["PLAKA NO", "PLAKA"])
                firma_col = self._find_col(col_map, ["FIRMA", "COMPANY"])

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
                        normalize_turkish_key(row.get(loc_col)) if loc_col else "Factory Showroom"
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
                    plaka_no_val = row.get(plaka_no_col) if plaka_no_col else "-"
                    firma_val = row.get(firma_col) if firma_col else "ARSLANLAR"

                    # Çeki listesi ve detay raporları için satırı kaydet
                    all_raw_rows.append(
                        {
                            "stone_name": norm_stone,
                            "lot": lot_code or "GENERAL",
                            "slab_no": plaka_no_val,
                            "company": firma_val,
                            "status": raw_status,
                            "location": raw_loc,
                            "length": length_cm,
                            "height": height_cm,
                            "thickness": thick_str,
                            "surface": surf_str,
                            "m2": m2_val,
                            "is_in_stock": is_in_stock,
                        }
                    )

                    # Taş bazlı toplama
                    if norm_stone not in stones_agg:
                        stones_agg[norm_stone] = {
                            "stone_name": norm_stone,
                            "stock_m2": 0.0,
                            "stock_slabs": 0,
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
                        else "Factory Showroom"
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
            self.raw_stock_rows = all_raw_rows
            self.last_loaded_time = datetime.now().strftime("%d.%m.%Y %H:%M")
            logger.info(
                f"Excel başarıyla yüklendi: {len(self.stones_summary)} tekil taş, "
                f"{len(self.lots_index)} lot, {len(self.raw_stock_rows)} plaka satırı."
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
        if not user_text:
            return []
        matched: List[str] = []
        norm_msg = normalize_turkish_key(user_text)

        lot_regex = re.compile(r"\b([A-Z]{1,4})[-\s]?(\d{1,5})\b")
        for m in lot_regex.finditer(norm_msg):
            prefix, num = m.group(1), m.group(2)
            if prefix in ["M", "CM", "MM", "KG", "DC", "TON", "PDF"]:
                continue
            cand_lot = f"{prefix}-{num}"
            if cand_lot in self.lots_index or prefix == "A":
                matched.append(cand_lot)

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

    def generate_packing_list_files(self, query: str) -> Tuple[Optional[bytes], Optional[bytes], str]:
        """
        Belirtilen lot veya taş için hem Premium Excel Çeki Listesi (.xlsx)
        hem de Kurumsal PDF Çeki Listesi (.pdf) üretir.
        """
        norm_key = normalize_lot_code(query) or normalize_turkish_key(query)
        matching_rows = [
            r for r in self.raw_stock_rows
            if norm_key in normalize_lot_code(r["lot"]) or norm_key in r["stone_name"]
        ]
        if not matching_rows:
            # Bulunamazsa tüm aktif stokları baz al
            matching_rows = [r for r in self.raw_stock_rows if r["is_in_stock"]][:50]

        stone_name_title = matching_rows[0]["stone_name"] if matching_rows else query.upper()
        lot_title = matching_rows[0]["lot"] if matching_rows else query

        # 1. EXCEL ÇEKİ LİSTESİ ÜRETİMİ (Pandas / Openpyxl)
        excel_io = io.BytesIO()
        df_export = pd.DataFrame(matching_rows)
        # Sütun isimlerini ve sıralamasını güzelleştir
        rename_cols = {
            "stone_name": "SELECTION / STONE",
            "lot": "LOT / BLOCK",
            "slab_no": "SLAB NO",
            "company": "COMPANY",
            "status": "STATUS",
            "location": "WAREHOUSE LOCATION",
            "length": "LENGTH (cm)",
            "height": "HEIGHT (cm)",
            "thickness": "THICKNESS",
            "surface": "SURFACE FINISH",
            "m2": "NET M2",
        }
        df_export = df_export.rename(columns=rename_cols)
        with pd.ExcelWriter(excel_io, engine="openpyxl") as writer:
            df_export.to_excel(writer, index=False, sheet_name="Packing List")
        excel_io.seek(0)
        excel_bytes = excel_io.read()

        # 2. PDF ÇEKİ LİSTESİ ÜRETİMİ (ReportLab)
        pdf_io = io.BytesIO()
        doc = SimpleDocTemplate(
            pdf_io, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30
        )
        styles = getSampleStyleSheet()
        NAVY = colors.HexColor("#14213D")
        GOLD = colors.HexColor("#C5A059")
        SLATE = colors.HexColor("#2B2D42")
        LIGHT_BG = colors.HexColor("#F8F9FA")

        title_p = ParagraphStyle("PlTitle", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=18, leading=22, textColor=NAVY, alignment=1)
        sub_p = ParagraphStyle("PlSub", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=10, leading=13, textColor=GOLD, alignment=1)
        th_p = ParagraphStyle("PlTH", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=8, leading=10, textColor=colors.whitesmoke, alignment=1)
        td_p = ParagraphStyle("PlTD", parent=styles["Normal"], fontName="Helvetica", fontSize=7.5, leading=10, textColor=SLATE, alignment=1)

        story = [
            Paragraph("ARSLANLAR MARBLE", title_p),
            Paragraph(f"OFFICIAL PACKING LIST — LOT: {lot_title} ({stone_name_title})", sub_p),
            HRFlowable(width="100%", thickness=1, color=GOLD, spaceAfter=10),
        ]

        table_data = [[
            Paragraph("Slab #", th_p),
            Paragraph("Lot/Block", th_p),
            Paragraph("Selection", th_p),
            Paragraph("Length", th_p),
            Paragraph("Height", th_p),
            Paragraph("Thick", th_p),
            Paragraph("Net M2", th_p),
            Paragraph("Location", th_p),
        ]]

        total_m2 = 0.0
        for idx, r in enumerate(matching_rows[:100], start=1):
            total_m2 += r["m2"]
            table_data.append([
                Paragraph(str(r["slab_no"]), td_p),
                Paragraph(str(r["lot"]), td_p),
                Paragraph(str(r["stone_name"]), td_p),
                Paragraph(f"{r['length']}", td_p),
                Paragraph(f"{r['height']}", td_p),
                Paragraph(str(r["thickness"]), td_p),
                Paragraph(f"{r['m2']:.2f}", td_p),
                Paragraph(str(r["location"]), td_p),
            ])

        pl_table = Table(table_data, colWidths=[45, 60, 110, 50, 50, 45, 55, 120])
        pl_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), NAVY),
            ("BACKGROUND", (0, 1), (-1, -1), LIGHT_BG),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#DEE2E6")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(pl_table)
        story.append(Spacer(1, 10))

        summary_p = Paragraph(
            f"<b>Total Slabs:</b> {len(matching_rows)} pcs | <b>Total Net Area:</b> {total_m2:,.2f} m²<br/>"
            f"<i>Arslanlar Marble Export Department — Packing List Generated Autonomously.</i>",
            ParagraphStyle("SumSt", parent=styles["Normal"], fontName="Helvetica-Oblique", fontSize=8, textColor=SLATE)
        )
        story.append(summary_p)

        doc.build(story)
        pdf_io.seek(0)
        pdf_bytes = pdf_io.read()

        file_prefix = f"Arslanlar_PackingList_{lot_title.replace('-', '')}"
        return excel_bytes, pdf_bytes, file_prefix

    def get_context_snapshot(self, user_text: str = "") -> str:
        lines = [
            f"=== ARSLANLAR MARBLE CANLI STOK VERİTABANI (Son Güncelleme: {self.last_loaded_time or 'Bekleniyor'}) ==="
        ]
        extracted = self.extract_entities_from_text(user_text)
        if extracted:
            lines.append("\n>>> KULLANICININ SORDUĞU LOT / TAŞ DETAYLARI (KESİN VERİ) <<<")
            for q in extracted:
                d = self.find_item_data(q)
                lines.append(
                    f"* {d['display_title']} -> "
                    f"AKTİF STOK: {d.get('stock_m2', 0.0):.2f} m² ({d.get('stock_slabs', 0)} Plaka) | "
                    f"Yüzey & Kalınlık: {d.get('surface_str', 'Polished')} ({d.get('thickness_str', '2 cm')}) | "
                    f"Konum: {d.get('location_str', '-')}"
                )
        return "\n".join(lines)


# =====================================================================
# 5. %100 İNGİLİZCE, LOGOLU VE FOTOĞRAFLI PREMIUM PDF SUNUM MOTORU
# =====================================================================
class PDFPresentationEngine:
    def __init__(self, inventory: InventoryEngine, drive: DriveExplorer):
        self.inventory = inventory
        self.drive = drive

    @staticmethod
    def _create_fallback_slab_image() -> bytes:
        img = Image.new("RGB", (600, 400), color=(20, 33, 61))
        draw = ImageDraw.Draw(img)
        draw.rectangle([12, 12, 588, 388], outline=(197, 160, 89), width=2)
        draw.text((180, 185), "ARSLANLAR MARBLE PREMIER", fill=(220, 225, 230))
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
                    img1, img2 = self.drive.get_images_structured(sf["id"], sf["name"])
                    cards_to_render.append(
                        {
                            "title": item_data["display_title"],
                            "data": item_data,
                            "drive_link": sf["link"],
                            "image_bytes": img1,
                            "image_bytes_2": img2,
                        }
                    )

        if not cards_to_render:
            target_queries = queries if queries else ["A-283", "SILVER ROOTS"]
            for q in target_queries[:12]:
                item_data = self.inventory.find_item_data(q)
                q_key = item_data["query_key"]
                st_name = item_data.get("stone_name", "")
                folder_id, folder_link = self.drive.find_folder(q_key, st_name)
                img1, img2 = self.drive.get_images_structured(folder_id, q_key)
                cards_to_render.append(
                    {
                        "title": item_data["display_title"],
                        "data": item_data,
                        "drive_link": folder_link or DEFAULT_DRIVE_LINK,
                        "image_bytes": img1,
                        "image_bytes_2": img2,
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
            fontSize=20,
            leading=24,
            textColor=NAVY,
            alignment=1,
            spaceAfter=2,
        )
        subtitle_style = ParagraphStyle(
            "ArslanlarSub",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=9.5,
            leading=12,
            textColor=GOLD,
            alignment=1,
            spaceAfter=8,
        )
        card_heading_style = ParagraphStyle(
            "CardHeading",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=11,
            leading=15,
            textColor=NAVY,
            spaceBefore=4,
            spaceAfter=4,
        )
        th_style = ParagraphStyle(
            "TH",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=10,
            textColor=colors.whitesmoke,
            alignment=0,
        )
        td_label_style = ParagraphStyle(
            "TDLabel",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=10,
            textColor=SLATE,
            alignment=0,
        )
        td_val_style = ParagraphStyle(
            "TDVal",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=8,
            leading=10,
            textColor=SLATE,
            alignment=0,
        )
        footer_style = ParagraphStyle(
            "FooterEN",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=7.5,
            leading=10,
            textColor=colors.HexColor("#6C757D"),
            alignment=1,
        )

        story: List[Any] = []
        story.append(Paragraph("ARSLANLAR MARBLE", title_style))
        story.append(Paragraph(subtitle_en, subtitle_style))
        story.append(HRFlowable(width="100%", thickness=1.5, color=GOLD, spaceAfter=10))

        for idx, card in enumerate(cards_to_render, start=1):
            d = card["data"]
            stock_m2 = d.get("stock_m2", 0.0)
            stock_slabs = d.get("stock_slabs", 0)

            stock_str = (
                f"{stock_m2:,.2f} m² ({stock_slabs:,} Slabs)"
                if stock_m2 > 0
                else "Available on Order / Check Production"
            )
            surf_thick_str = (
                f"{d.get('surface_str', 'Polished')} | {d.get('thickness_str', '2 cm')}"
            )
            lots_list = d.get("active_lots") or d.get("all_lots") or []
            lots_str = (
                ", ".join(lots_list[:6]) if lots_list else str(d.get("query_key", "-"))
            )
            drive_url = card["drive_link"]

            # Sadece plaka fotoğrafı (Bookmatch kaldırıldı, en iyi çekim seçildi)
            raw_img = card["image_bytes"] or self._create_fallback_slab_image()
            try:
                pil_im = Image.open(io.BytesIO(raw_img)).convert("RGB")
                img_io = io.BytesIO()
                pil_im.save(img_io, format="JPEG", quality=88)
                img_io.seek(0)
                rl_img = RLImage(img_io, width=160, height=110)
            except Exception:
                fb_io = io.BytesIO(self._create_fallback_slab_image())
                rl_img = RLImage(fb_io, width=160, height=110)

            # İstenmeyen 'Total Production' ve boş kalan 'Avg. Slab Size' kaldırıldı!
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

            spec_table = Table(spec_rows, colWidths=[110, 240])
            spec_table.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                        ("BACKGROUND", (0, 1), (-1, -1), LIGHT_BG),
                        ("GRID", (0, 0), (-1, -1), 0.5, BORDER_GRAY),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("TOPPADDING", (0, 0), (-1, -1), 3),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                        ("LEFTPADDING", (0, 0), (-1, -1), 5),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                    ]
                )
            )

            master_row = Table([[rl_img, spec_table]], colWidths=[170, 350])
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
                Spacer(1, 10),
            ]
            story.append(KeepTogether(block_elements))

        story.append(Spacer(1, 4))
        story.append(
            HRFlowable(width="100%", thickness=0.5, color=BORDER_GRAY, spaceAfter=4)
        )
        story.append(
            Paragraph(
                "<b>ARSLANLAR MARBLE</b> | International Sales & Export Department | Izmir, Turkey | info@arslanlarmarble.com",
                footer_style,
            )
        )

        doc.build(story)
        pdf_buffer.seek(0)
        clean_filename = f"Arslanlar_Marble_Presentation_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
        return pdf_buffer.read(), clean_filename


# =====================================================================
# 6. ÇİFT MOTORLU ORKESTRATÖR (AI + YEREL UZMAN MOTORU)
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
            endpoints_to_try.append(
                (
                    f"https://aiplatform.googleapis.com/v1beta1/publishers/google/models/gemini-2.0-flash:generateContent?key={self.api_key}",
                    {"Content-Type": "application/json"},
                )
            )
            endpoints_to_try.append(
                (
                    "https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent",
                    {
                        "Content-Type": "application/json",
                        "X-goog-api-key": self.api_key,
                    },
                )
            )

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
                logger.debug(f"Service account token atlandı: {sa_err}")

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
        raw_text = (user_text or "").strip()
        norm_text = normalize_turkish_key(raw_text)

        is_ukrainian = bool(re.search(r"[а-яА-ЯіїєґІЇЄҐ]", raw_text))
        is_english = bool(
            re.search(
                r"\b(please|stock|slab|slabs|send|photo|photos|presentation|price|available|what|how many|packing list)\b",
                raw_text,
                rre := re.IGNORECASE,
            )
        )
        lang = "uk" if is_ukrainian else ("en" if is_english else "tr")

        matched_entities = self.inventory.extract_entities_from_text(raw_text)

        collection_folder = None
        if "COLLECTION" in norm_text:
            collection_folder = "Arslanlar Collection"
        elif "SERIES" in norm_text or "SERISI" in norm_text:
            collection_folder = "Arslanlar Series"

        want_pdf = bool(
            re.search(r"\b(pdf|sunum|sunumu|katalog|presentation|catalog)\b", raw_text, re.IGNORECASE)
        )
        want_packing_list = bool(
            re.search(r"(çeki\s*listesi|packing\s*list|sevk\s*listesi|çeki)", raw_text, re.IGNORECASE)
        )
        want_photos = (
            bool(re.search(r"(foto|resim|gorsel|görsel|photo|picture|image|album|albüm|фото)", raw_text, re.IGNORECASE))
            or bool(image_bytes)
        )

        business_keywords = [
            "STOK", "PLAKA", "LOT", "BLOK", "LINK", "DRIVE", "FOTO", "RESIM", "GORSEL",
            "PDF", "SUNUM", "KATALOG", "METRAJ", "M2", "ELIMIZDE", "MEVCUT", "VAR MI",
            "ÇEKİ", "PACKING", "LISTESI", "LIST"
        ]
        has_business_intent = (
            bool(matched_entities) or want_pdf or want_packing_list or want_photos or bool(collection_folder)
            or any(kw in norm_text for kw in business_keywords)
        )

        greeting_keywords = ["ORDA MISIN", "ORADA MISIN", "MERHABA", "SELAM", "HELLO", "HI", "HEY"]
        is_greeting = any(gk in norm_text for gk in greeting_keywords)

        if is_group and not has_business_intent and not bot_mentioned:
            if not is_greeting:
                return {
                    "should_reply": False, "language": lang, "reply_text": "",
                    "matched_queries": [], "send_photos": False, "generate_pdf": False,
                    "generate_packing_list": False, "collection_folder": None,
                }

        # 1. Çeki Listesi İsteği (Excel + PDF)
        if want_packing_list:
            target_q = matched_entities[0] if matched_entities else "GENERAL"
            msg = f"📦 {target_q} lotu / seleksiyonu için resmi Çeki Listesi (Excel ve PDF) hazırlanarak gönderiliyor."
            return {
                "should_reply": True, "language": lang, "reply_text": msg,
                "matched_queries": matched_entities[:1], "send_photos": False,
                "generate_pdf": False, "generate_packing_list": True, "collection_folder": None,
            }

        # 2. PDF Sunum İsteği
        if want_pdf:
            target_list = matched_entities if matched_entities else (["A-283", "SILVER ROOTS"] if not collection_folder else [])
            msg = "📄 Kurumsal logolu ve %100 İngilizce lüks PDF sunumunuz hazırlanıyor."
            return {
                "should_reply": True, "language": lang, "reply_text": msg,
                "matched_queries": target_list, "send_photos": False,
                "generate_pdf": True, "generate_packing_list": False, "collection_folder": collection_folder,
            }

        # 3. Spesifik Lot veya Taş Sorgusu (Flash Kart / Bilgi + Fotoğraflar)
        if matched_entities:
            blocks = []
            for q in matched_entities[:3]:
                d = self.inventory.find_item_data(q)
                title = d["display_title"]
                st_m2 = d.get("stock_m2", 0.0)
                st_slabs = d.get("stock_slabs", 0)
                dims = d.get("dimensions_str", "-")
                surf = d.get("surface_str", "Polished")
                thick = d.get("thickness_str", "2 cm")
                loc = d.get("location_str", "-")

                blocks.append(
                    f"💎 **{title}**\n"
                    f"• Aktif Mevcut Stok: `{st_m2:,.2f} m²` (`{st_slabs} Plaka`)\n"
                    f"• Ortalama Ebat: `{dims}`\n"
                    f"• Yüzey & Kalınlık: `{surf} ({thick})`\n"
                    f"• Stok Konumu: `{loc}`"
                )
            return {
                "should_reply": True, "language": lang,
                "reply_text": "\n\n".join(blocks),
                "matched_queries": matched_entities[:3],
                "send_photos": True, # 1. Plaka ve 2. Bookmatch / Plaka albümü tetiklenir
                "generate_pdf": False, "generate_packing_list": False, "collection_folder": None,
            }

        # 4. Genel Stok Özeti
        if any(kw in norm_text for kw in ["STOK", "ELIMIZDE", "MEVCUT", "LISTE"]):
            if not self.inventory.stones_summary:
                return {
                    "should_reply": True, "language": lang,
                    "reply_text": "⚠️ Güncel stok Excel dosyası henüz yüklenmedi.",
                    "matched_queries": [], "send_photos": False, "generate_pdf": False,
                    "generate_packing_list": False, "collection_folder": None,
                }
            sorted_stones = sorted(
                self.inventory.stones_summary.values(),
                key=lambda x: (x["stock_m2"], x["production_m2"]),
                reverse=True,
            )[:12]
            lines = ["📊 **ARSLANLAR MARBLE GÜNCEL AKTİF STOK ÖZETİ:**\n"]
            for idx, item in enumerate(sorted_stones, start=1):
                lines.append(f"{idx}. `{item['stone_name']}`: **{item['stock_m2']:,.2f} m²** ({item['stock_slabs']} Plaka)")
            return {
                "should_reply": True, "language": lang, "reply_text": "\n".join(lines),
                "matched_queries": [], "send_photos": False, "generate_pdf": False,
                "generate_packing_list": False, "collection_folder": None,
            }

        return {
            "should_reply": False, "language": lang, "reply_text": "",
            "matched_queries": [], "send_photos": False, "generate_pdf": False,
            "generate_packing_list": False, "collection_folder": None,
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
Sen Arslanlar Marble şirketinin kıdemli dijital satış ve operasyon uzmanısın. Adın Argon.
GÖREVLER:
1. Kullanıcının mesajında geçen lotları/taşları bul (`matched_queries`: {json.dumps(extracted_entities)}).
2. "çeki listesi" istendiyse `"generate_packing_list": true` yap.
3. "PDF sunum/katalog" istendiyse `"generate_pdf": true` yap.
4. Fotoğraf veya lot detayları istendiyse `"send_photos": true` yap.

SOHBET GEÇMİŞİ:
{recent_history}
{catalog_context}
GELEN MESAJ: {user_text}

JSON ÇIKTI:
{{
  "should_reply": true,
  "language": "tr",
  "reply_text": "Yanıt...",
  "matched_queries": {json.dumps(extracted_entities)},
  "send_photos": true,
  "generate_pdf": false,
  "generate_packing_list": false,
  "collection_folder": null
}}
"""
        try:
            result = await asyncio.to_thread(
                self._call_gemini_multi_channel, full_prompt, image_bytes
            )
            if extracted_entities and not result.get("matched_queries"):
                result["matched_queries"] = extracted_entities
            if re.search(r"(çeki\s*listesi|packing\s*list|sevk)", user_text, re.IGNORECASE):
                result["generate_packing_list"] = True
            if re.search(r"\bpdf\b", user_text, re.IGNORECASE):
                result["generate_pdf"] = True
            return result
        except Exception:
            return self._local_expert_fallback(
                user_text=user_text, is_group=is_group,
                bot_mentioned=bot_mentioned, image_bytes=image_bytes,
            )


# =====================================================================
# 7. TELEGRAM HANDLER'LARI (EXCEL + PDF ÇEKİ LİSTESİ + ALBÜM)
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
        f"📥 **{file_name}** alındı, stok verileri işleniyor...", parse_mode=ParseMode.MARKDOWN
    )

    try:
        tg_file = await context.bot.get_file(doc.file_id)
        await tg_file.download_to_drive(EXCEL_FILE_PATH)
        stones_count, lots_count = await asyncio.to_thread(
            inventory_engine.reload_excel
        )
        await status_msg.edit_text(
            f"✅ **Stok Listesi Güncellendi!** 🚀\n"
            f"📊 Taş Ceşidi: `{stones_count}` | Lot/Blok: `{lots_count}`",
            parse_mode=ParseMode.MARKDOWN,
        )
    except Exception as e:
        await status_msg.edit_text(f"❌ Hata: {e}")


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

    link_teach_match = re.search(
        r"([A-Za-z0-9-_]+)\s+linki\s*:\s*(https?://\S+)", user_text, re.IGNORECASE
    )
    if link_teach_match:
        lot_key = normalize_lot_code(link_teach_match.group(1))
        url_val = link_teach_match.group(2)
        drive_explorer.save_custom_link(lot_key, url_val)
        await message.reply_text(f"✅ **{lot_key}** bağlantısı kaydedildi:\n{url_val}", parse_mode=ParseMode.MARKDOWN)
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
            logger.error(f"Fotoğraf indirme hatası: {e}")

    if not user_text and not image_bytes:
        return

    argon_brain.add_memory(
        chat_id, user_name, user_text if user_text else "[Mermer Fotoğrafı]"
    )

    decision = await argon_brain.evaluate_and_respond(
        chat_id=chat_id, user_name=user_name, user_text=user_text,
        is_group=is_group, bot_mentioned=bot_mentioned, image_bytes=image_bytes,
    )

    if not decision.get("should_reply", False):
        return

    reply_text = decision.get("reply_text", "").strip()
    matched_queries: List[str] = decision.get("matched_queries", [])
    send_photos: bool = decision.get("send_photos", False) or bool(image_bytes)
    want_pdf: bool = decision.get("generate_pdf", False)
    want_packing_list: bool = decision.get("generate_packing_list", False)
    collection_folder: Optional[str] = decision.get("collection_folder")

    if re.search(r"\bpdf\b", user_text, re.IGNORECASE):
        want_pdf = True
    if re.search(r"(çeki\s*listesi|packing\s*list|sevk)", user_text, re.IGNORECASE):
        want_packing_list = True

    if not reply_text:
        return

    # 1. ÇEKİ LİSTESİ İSTENDİYSE: Excel ve PDF Dosyasını Birlikte Gönder
    if want_packing_list:
        await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.UPLOAD_DOCUMENT)
        try:
            target_q = matched_queries[0] if matched_queries else "GENERAL"
            excel_b, pdf_b, prefix_name = await asyncio.to_thread(
                inventory_engine.generate_packing_list_files, target_q
            )
            # Excel gönder
            await message.reply_document(
                document=io.BytesIO(excel_b),
                filename=f"{prefix_name}_PackingList.xlsx",
                caption=f"📊 Resmi Excel Çeki Listesi — {target_q}",
            )
            # PDF gönder
            await message.reply_document(
                document=io.BytesIO(pdf_b),
                filename=f"{prefix_name}_PackingList.pdf",
                caption=f"📄 Resmi PDF Çeki Listesi — {target_q}",
            )
            return
        except Exception as pl_err:
            logger.error(f"Çeki listesi oluşturma hatası: {pl_err}")

    # 2. PDF SUNUM İSTENDİYSE
    if want_pdf:
        await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.UPLOAD_DOCUMENT)
        try:
            pdf_bytes, pdf_filename = await asyncio.to_thread(
                pdf_engine.generate_pdf, matched_queries, collection_folder
            )
            await message.reply_document(
                document=io.BytesIO(pdf_bytes), filename=pdf_filename, caption=reply_text[:1020]
            )
            return
        except Exception as pdf_err:
            logger.error(f"PDF sunum hatası: {pdf_err}")

    # 3. FLASH KART MESAJI + 2 FOTOĞRAF (1. Plaka, 2. Bookmatch / 2. Plaka)
    drive_links_section: List[str] = []
    media_group_items: List[InputMediaPhoto] = []
    newly_cached_keys: List[Tuple[str, int]] = []
    last_folder_link: Optional[str] = None

    for query in matched_queries[:2]:
        norm_q = normalize_lot_code(query)
        item_info = inventory_engine.find_item_data(norm_q)
        stone_name_fallback = item_info.get("stone_name", "")

        folder_id, folder_link = drive_explorer.find_folder(norm_q, stone_name_fallback)

        if folder_link:
            last_folder_link = folder_link
            drive_links_section.append(f"📂 {norm_q} Drive: {folder_link}")
        else:
            drive_links_section.append(f"📂 {norm_q} Genel Drive: {DEFAULT_DRIVE_LINK}")

        if send_photos:
            await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.UPLOAD_PHOTO)
            img1, img2 = await asyncio.to_thread(
                drive_explorer.get_images_structured, folder_id, norm_q
            )
            if img1:
                media_group_items.append(InputMediaPhoto(media=img1))
            if img2:
                media_group_items.append(InputMediaPhoto(media=img2))

    if drive_links_section:
        links_block = "\n".join(drive_links_section)
        if not last_folder_link or last_folder_link not in reply_text:
            reply_text = f"{reply_text}\n\n{links_block}"

    argon_brain.add_memory(chat_id, "Argon (Bot)", reply_text)

    if media_group_items:
        try:
            short_caption = reply_text[:1020] + "..." if len(reply_text) > 1024 else reply_text
            media_group_items[0] = InputMediaPhoto(
                media=media_group_items[0].media, caption=short_caption
            )
            await message.reply_media_group(media=media_group_items[:10])
            if len(reply_text) > 1024:
                await message.reply_text(reply_text)
            return
        except Exception as e:
            logger.error(f"Albüm gönderme hatası: {e}")

    await message.reply_text(reply_text, disable_web_page_preview=False)


# =====================================================================
# 8. RENDER BAŞLATICI
# =====================================================================
flask_app = Flask(__name__)


@flask_app.route("/")
def health_check():
    return f"Arslanlar Marble Argon Master Ultimate v3 is Live!", 200


def run_flask():
    port = int(os.environ.get("PORT", 10000))
    flask_app.run(host="0.0.0.0", port=port, use_reloader=False)


def background_initial_load():
    time.sleep(1)
    inventory_engine.reload_excel()


def main():
    if not TELEGRAM_TOKEN:
        raise ValueError("TELEGRAM_TOKEN eksik!")

    threading.Thread(target=run_flask, daemon=True).start()
    threading.Thread(target=background_initial_load, daemon=True).start()

    application = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    application.add_handler(MessageHandler(filters.Document.ALL, handle_document_upload))
    application.add_handler(
        MessageHandler((filters.TEXT | filters.PHOTO) & (~filters.COMMAND), handle_incoming_message)
    )

    logger.info("🚀 Arslanlar Marble Argon Master Ultimate v3 başlatılıyor...")
    application.run_polling(drop_pending_updates=True, allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
