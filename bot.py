import os
import io
import re
import gc
import json
import time
import math
import base64
import logging
import asyncio
import threading
from datetime import datetime
from collections import deque
from typing import Dict, List, Optional, Tuple, Any

# =====================================================================
# 0. RENDER PORTUNU 0. SANİYEDE AÇ (DEPLOY TAKILMASINI %100 ÖNLER)
# =====================================================================
from flask import Flask
import requests

flask_app = Flask(__name__)

@flask_app.route("/")
def health_check():
    return "Argon Master V12 Corporate Engine is Live!", 200

@flask_app.route("/healthz")
def healthz():
    return "OK", 200

def _start_flask_immediately():
    port = int(os.environ.get("PORT", 10000))
    flask_app.run(host="0.0.0.0", port=port, threaded=True, use_reloader=False)

threading.Thread(target=_start_flask_immediately, daemon=True).start()

# =====================================================================
# 1. KÜTÜPHANELER VE GÖMÜLÜ ORİJİNAL ARSLANLAR MARBLE LOGOSU
# =====================================================================
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.drawing.image import Image as XLImage
from PIL import Image

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

from telegram import Update, InputMediaPhoto
from telegram.error import Conflict, NetworkError, TimedOut
from telegram.constants import ParseMode
from telegram.ext import (
    ApplicationBuilder,
    ContextTypes,
    MessageHandler,
    filters,
)

from google.oauth2 import service_account
import google.auth.transport.requests

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("ArgonMaster_V12")

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GOOGLE_SA_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
DEFAULT_DRIVE_LINK = (
    os.getenv("GOOGLE_DRIVE_LINK")
    or os.getenv("DEFAULT_DRIVE_LINK")
    or "https://drive.google.com"
).strip()

# GÖMÜLÜ ORİJİNAL ARSLANLAR MARBLE BRONZ ASLAN LOGOSU (logo_bronz)
EMBEDDED_LOGO_B64 = (
    "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"
)

# Alt klasör kara listesi (Bunlar asla ayrı bir taş veya lot olarak listelenmez!)
SUBFOLDER_BLACKLIST = {
    "CUSTOM",
    "WEB",
    "LARGE",
    "SMALL",
    "MEDIUM",
    "EMAIL",
    "E-MAIL",
    "MAIL",
    "PHOTOS",
    "PHOTO",
    "FOTO",
    "FOTOGRAF",
    "FOTOGRAFLAR",
    "PICTURES",
    "IMAGES",
    "BOOKMATCH",
    "BM",
    "PACKING",
    "PACKING LIST",
    "CEKI",
    "CEKI LISTESI",
    "VIDEO",
    "VIDEOS",
    "SLAB",
    "SLABS",
    "PLAKA",
    "PLAKALAR",
    "RAW",
    "ORJINAL",
    "ORIJINAL",
    "HD",
    "HIGH RES",
    "LOW RES",
}

DEFAULT_STONE_DNA = {
    "A-283": {
        "stone_name": "VIOLET BRECCIA",
        "keywords": ["violet", "breccia", "mor", "kırıklı", "beyaz", "purple"],
    },
    "SILVER ROOTS": {
        "stone_name": "SILVER ROOTS",
        "keywords": ["silver", "roots", "gümüş", "gri", "kök", "füme", "grey"],
    },
    "SILVER RIVER": {
        "stone_name": "SILVER RIVER",
        "keywords": ["silver river", "nehir", "çizgisel", "dalgalı", "linear", "vein"],
    },
    "ROSSO LEVANTO": {
        "stone_name": "ROSSO LEVANTO",
        "keywords": ["rosso", "levanto", "bordo", "kırmızı", "vişne", "red", "burgundy"],
    },
    "TERRA MATER": {
        "stone_name": "TERRA MATER",
        "keywords": ["terra", "mater", "toprak", "kahve", "bej", "brown", "earth"],
    },
    "BLUE ROOTS": {
        "stone_name": "BLUE ROOTS",
        "keywords": ["blue", "mavi", "gri", "soğuk"],
    },
    "DEKART": {
        "stone_name": "DEKART",
        "keywords": ["dekart", "decart", "antrasit", "koyu gri", "dark grey"],
    },
}


# =====================================================================
# 2. YARDIMCI FONKSİYONLAR, İSİM TEMİZLEME VE BUNDLE HESAPLAMA
# =====================================================================
def normalize_turkish_key(text: Any) -> str:
    if text is None or pd.isna(text):
        return ""
    tr_map = str.maketrans("İıiŞşĞğÜüÖöÇç", "IIISSGGUUOOCC")
    s = str(text).strip().translate(tr_map).upper()
    return re.sub(r"\s+", " ", s).strip()


def normalize_lot_code(code: Any) -> str:
    raw = normalize_turkish_key(code)
    match = re.match(r"^([A-Z]+)[-\s_.]*0*(\d+)$", raw)
    return f"{match.group(1)}-{match.group(2)}" if match else raw


def is_utility_subfolder(folder_name: str) -> bool:
    """
    Klasör adı 'CUSTOM', 'WEB', 'LARGE', 'EMAIL', 'PHOTOS', 'BOOKMATCH' gibi
    bir medya alt klasörü mü kontrol eder.
    """
    norm = normalize_turkish_key(folder_name)
    norm_clean = re.sub(r"[^A-Z0-9\s]", " ", norm).strip()
    if norm_clean in SUBFOLDER_BLACKLIST:
        return True
    # Eğer içinde hiç lot kodu (örn: A-468) yoksa ve tamamen kara liste kelimelerinden oluşuyorsa
    if not re.search(r"\b[A-Z]{1,4}[-\s_.]*\d{1,5}\b", norm):
        words = norm_clean.split()
        if words and all(w in SUBFOLDER_BLACKLIST for w in words):
            return True
    return False


def clean_stone_and_lot_title(
    raw_title: str, parent_stone: str = ""
) -> Tuple[str, str, str]:
    """
    'CEPPO BRECCIA A-468[FULL LOT] - CUSTOM' gibi kirli klasör adlarını temizler.
    Dönüş: (clean_display_title, clean_stone_name, clean_lot_code)
    Örnek: ('CEPPO BRECCIA — A-468', 'CEPPO BRECCIA', 'A-468')
    """
    s = normalize_turkish_key(raw_title)
    # Köşeli ve normal parantez içindeki [FULL LOT] vb. ifadeleri sil
    s = re.sub(r"\[.*?\]|\(.*?\)", " ", s)

    # Lot kodunu yakala (Örn: A-468, A-1009, A-91)
    lot_m = re.search(r"\b([A-Z]{1,4})[-\s_.]*0*(\d{1,5})\b", s)
    lot_code = f"{lot_m.group(1)}-{lot_m.group(2)}" if lot_m else ""

    stone_part = s
    if lot_m:
        stone_part = s[: lot_m.start()] + " " + s[lot_m.end() :]

    # Gürültü kelimelerini temizle
    noise_words = list(SUBFOLDER_BLACKLIST) + ["FULL", "LOT", "BLOCK", "BLOK"]
    for nw in noise_words:
        stone_part = re.sub(
            rf"(?<![A-Z0-9]){re.escape(nw)}(?![A-Z0-9])", " ", stone_part
        )

    stone_part = re.sub(r"[-_/:,.]+", " ", stone_part)
    stone_part = re.sub(r"\s+", " ", stone_part).strip()

    if not stone_part and parent_stone:
        p_clean = normalize_turkish_key(parent_stone)
        p_clean = re.sub(r"\[.*?\]|\(.*?\)", " ", p_clean)
        for nw in noise_words:
            p_clean = re.sub(
                rf"(?<![A-Z0-9]){re.escape(nw)}(?![A-Z0-9])", " ", p_clean
            )
        p_clean = re.sub(r"[-_/:,.]+", " ", p_clean)
        stone_part = re.sub(r"\s+", " ", p_clean).strip()

    if (
        stone_part in ("ARSLANLAR COLLECTION", "ARSLANLAR SERIES", "ARSLANLAR EXCLUSIVE COLLECTION")
        and lot_code
    ):
        stone_part = ""

    if not stone_part and lot_code in DEFAULT_STONE_DNA:
        stone_part = DEFAULT_STONE_DNA[lot_code]["stone_name"]

    if stone_part and lot_code:
        display_title = f"{stone_part} — {lot_code}"
    elif lot_code:
        display_title = lot_code
    else:
        display_title = stone_part or "ARSLANLAR MARBLE SELECTION"

    return (
        display_title,
        (stone_part or "ARSLANLAR MARBLE"),
        (lot_code or display_title),
    )


def exact_lot_match_in_text(folder_or_file_name: str, target_lot: str) -> int:
    fn_norm = normalize_turkish_key(folder_or_file_name)
    target_norm = normalize_lot_code(target_lot)

    if normalize_lot_code(fn_norm) == target_norm:
        return 100

    lot_m = re.match(r"^([A-Z]+)-(\d+)$", target_norm)
    if lot_m:
        prefix, num = lot_m.group(1), lot_m.group(2)
        pattern = rf"(?<![A-Z0-9]){prefix}[-\s_.]*0*{num}(?![0-9])"
        if re.search(pattern, fn_norm):
            return 90
        return 0

    if target_norm == fn_norm:
        return 100
    if re.search(rf"(?<![A-Z0-9]){re.escape(target_norm)}(?![A-Z0-9])", fn_norm):
        return 85
    return 0


def safe_float(val: Any) -> float:
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return 0.0
    s = re.sub(r"(?i)\b(m2|m²|sqm|cm|mm|pcs|adet)\b", "", s)
    m = re.search(r"[-+]?\d+(?:[.,]\d+)*", s)
    if not m:
        return 0.0
    num_str = m.group(0)
    if "," in num_str and "." in num_str:
        if num_str.rfind(",") > num_str.rfind("."):
            num_str = num_str.replace(".", "").replace(",", ".")
        else:
            num_str = num_str.replace(",", "")
    elif "," in num_str:
        num_str = num_str.replace(",", ".")
    try:
        return float(num_str)
    except Exception:
        return 0.0


def normalize_cm(val: float) -> float:
    if val <= 0:
        return 0.0
    if val > 600:
        return val / 10.0
    if 0 < val < 10:
        return val * 100.0
    return val


def split_slabs_into_bundles(
    clean_rows: List[Dict[str, Any]], min_bundle_size: int = 12
) -> List[List[Dict[str, Any]]]:
    """
    Plakaları minimum 12'şerli dengeli ahşap bundle'lara (wooden bundles) böler.
    (Örn: 24 plaka -> 12+12 (2 bundle), 33 plaka -> 17+16 (2 bundle), 40 plaka -> 14+13+13 (3 bundle))
    """
    n = len(clean_rows)
    if n == 0:
        return []
    if n < min_bundle_size * 2:
        return [clean_rows]
    num_bundles = max(1, n // min_bundle_size)
    base_sz = n // num_bundles
    rem = n % num_bundles
    bundles = []
    idx = 0
    for b in range(num_bundles):
        sz = base_sz + (1 if b < rem else 0)
        bundles.append(clean_rows[idx : idx + sz])
        idx += sz
    return bundles


def prepare_telegram_image_bytes(raw_bytes: Optional[bytes]) -> Optional[bytes]:
    if not raw_bytes:
        return None
    try:
        with Image.open(io.BytesIO(raw_bytes)) as im:
            img = im.convert("RGB")
            img.thumbnail((1600, 1600), Image.Resampling.LANCZOS)
            out = io.BytesIO()
            img.save(out, format="JPEG", quality=88)
            return out.getvalue()
    except Exception:
        return None


def prepare_pdf_card_image_bytes(raw_bytes: Optional[bytes]) -> Optional[bytes]:
    """
    PDF sunumları için 900x600 px (Yüksek Çözünürlük, %92 Kalite) net görsel üretir.
    """
    if not raw_bytes:
        return None
    try:
        with Image.open(io.BytesIO(raw_bytes)) as im:
            img = im.convert("RGB")
            img.thumbnail((900, 600), Image.Resampling.LANCZOS)
            out = io.BytesIO()
            img.save(out, format="JPEG", quality=92)
            return out.getvalue()
    except Exception:
        return None


# =====================================================================
# 3. SAF REST API TABANLI GOOGLE DRIVE MOTORU (3 KATMAN DERİN ARAMA)
# =====================================================================
class DriveExplorer:
    DRIVE_API_URL = "https://www.googleapis.com/drive/v3/files"

    def __init__(self, sa_json_str: str):
        self.sa_json_str = sa_json_str
        self._creds = None
        self._token_lock = threading.Lock()
        self._logo_jpeg_bytes: Optional[bytes] = None
        self._logo_png_bytes: Optional[bytes] = None
        self._lot_result_cache: Dict[str, Tuple[float, Dict[str, Any]]] = {}

        if self.sa_json_str:
            try:
                info = json.loads(self.sa_json_str)
                self._creds = service_account.Credentials.from_service_account_info(
                    info, scopes=["https://www.googleapis.com/auth/drive.readonly"]
                )
                logger.info("Google Drive REST Kimlik Bilgileri Hazır.")
            except Exception as e:
                logger.error(f"Google Service Account JSON Hatası: {e}")

    def _get_headers(self) -> Optional[Dict[str, str]]:
        if not self._creds:
            return None
        with self._token_lock:
            if not self._creds.valid or self._creds.expired or not self._creds.token:
                req = google.auth.transport.requests.Request()
                self._creds.refresh(req)
            return {"Authorization": f"Bearer {self._creds.token}"}

    def api_list_files(
        self,
        q: str,
        fields: str = "files(id,name,mimeType,parents,webViewLink,thumbnailLink,shortcutDetails,size)",
        page_size: int = 150,
    ) -> List[Dict[str, Any]]:
        headers = self._get_headers()
        if not headers:
            return []
        params = {
            "q": q,
            "spaces": "drive",
            "fields": fields,
            "includeItemsFromAllDrives": "true",
            "supportsAllDrives": "true",
            "pageSize": page_size,
        }
        try:
            resp = requests.get(
                self.DRIVE_API_URL, headers=headers, params=params, timeout=12
            )
            if resp.status_code == 200:
                return resp.json().get("files", [])
        except Exception as e:
            logger.warning(f"Drive REST List uyarısı: {e}")
        return []

    def get_folder_metadata(self, folder_id: str) -> Optional[Dict[str, Any]]:
        headers = self._get_headers()
        if not headers or not folder_id:
            return None
        try:
            url = f"{self.DRIVE_API_URL}/{folder_id}"
            params = {
                "fields": "id,name,parents,webViewLink",
                "supportsAllDrives": "true",
            }
            resp = requests.get(url, headers=headers, params=params, timeout=8)
            if resp.status_code == 200:
                return resp.json()
        except Exception:
            pass
        return None

    def download_file_bytes(
        self, file_id: str, mime_type: str = ""
    ) -> Optional[bytes]:
        headers = self._get_headers()
        if not headers or not file_id:
            return None
        try:
            if mime_type == "application/vnd.google-apps.spreadsheet":
                url = f"{self.DRIVE_API_URL}/{file_id}/export"
                params = {
                    "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                }
            else:
                url = f"{self.DRIVE_API_URL}/{file_id}"
                params = {"alt": "media", "supportsAllDrives": "true"}

            resp = requests.get(url, headers=headers, params=params, timeout=18)
            if resp.status_code == 200 and len(resp.content) > 50:
                return resp.content
        except Exception as e:
            logger.warning(f"Dosya indirme uyarısı ({file_id}): {e}")
        return None

    def get_official_logo_jpeg(self) -> bytes:
        if self._logo_jpeg_bytes:
            return self._logo_jpeg_bytes
        raw_png = base64.b64decode(EMBEDDED_LOGO_B64)
        with Image.open(io.BytesIO(raw_png)) as im:
            rgb_im = im.convert("RGB")
            out = io.BytesIO()
            rgb_im.save(out, format="JPEG", quality=96)
            self._logo_jpeg_bytes = out.getvalue()
        return self._logo_jpeg_bytes

    def get_official_logo_png(self) -> bytes:
        if self._logo_png_bytes:
            return self._logo_png_bytes
        self._logo_png_bytes = base64.b64decode(EMBEDDED_LOGO_B64)
        return self._logo_png_bytes

    def find_folder_exact(
        self, query_term: str
    ) -> Tuple[Optional[str], Optional[str], Optional[str], str]:
        """
        Aranan lot veya koleksiyon klasörünü TAM EŞLEŞME ile bulur.
        'CUSTOM', 'WEB', 'LARGE', 'EMAIL', 'PHOTOS' gibi alt medya klasörlerini eler,
        doğrudan ANA LOT klasörünü seçer!
        """
        norm_key = normalize_lot_code(query_term)
        if not norm_key:
            return None, None, None, query_term

        raw_clean = query_term.strip().replace("'", "")
        search_terms = [norm_key, raw_clean]

        lot_m = re.match(r"^([A-Z]+)-(\d+)$", norm_key)
        if lot_m:
            prefix, num = lot_m.group(1), lot_m.group(2)
            search_terms = [
                f"{prefix}-{num}",
                f"{prefix} {num}",
                f"{prefix}{num}",
                num,
            ]
        else:
            tr_var = norm_key.replace("SILVER", "SİLVER").replace("RIVER", "RİVER")
            if tr_var != norm_key:
                search_terms.append(tr_var)

        search_terms = list(dict.fromkeys([t for t in search_terms if len(t) >= 2]))

        matched_folders = []
        seen_ids = set()

        for term in search_terms:
            q = (
                f"mimeType = 'application/vnd.google-apps.folder' "
                f"and trashed = false and name contains '{term}'"
            )
            files = self.api_list_files(q, page_size=50)
            for f in files:
                fid = f["id"]
                if fid in seen_ids:
                    continue
                fname = f.get("name", "")
                score = exact_lot_match_in_text(fname, norm_key)
                if score > 0:
                    # Eğer klasör adında "- CUSTOM", "- WEB", "- LARGE", "- EMAIL", "- PHOTOS" varsa puanını düşür ki Ana Klasör seçilsin!
                    fn_up = normalize_turkish_key(fname)
                    if any(
                        re.search(rf"(?<![A-Z0-9]){bw}(?![A-Z0-9])", fn_up)
                        for bw in ["CUSTOM", "WEB", "LARGE", "EMAIL", "PHOTOS", "BOOKMATCH"]
                    ):
                        score -= 35
                    seen_ids.add(fid)
                    matched_folders.append((score, f))
            if any(s >= 90 for s, _ in matched_folders):
                break

        if not matched_folders:
            return None, None, None, norm_key

        matched_folders.sort(key=lambda x: x[0], reverse=True)
        best = matched_folders[0][1]

        # Eğer seçilen klasör hâlâ bir alt klasörse (örn: "A-468 - CUSTOM") ve üst klasörü ana A-468 klasörüyse üst klasöre çık!
        parents = best.get("parents", [])
        parent_id = parents[0] if parents else None
        if parent_id and matched_folders[0][0] < 80:
            p_meta = self.get_folder_metadata(parent_id)
            if p_meta and exact_lot_match_in_text(p_meta.get("name", ""), norm_key) > 0:
                best = p_meta
                p_parents = best.get("parents", [])
                parent_id = p_parents[0] if p_parents else None

        fid = best["id"]
        link = best.get(
            "webViewLink", f"https://drive.google.com/drive/folders/{fid}"
        )
        return fid, link, parent_id, best.get("name", norm_key)

    def batch_list_children(
        self, parent_ids: List[str]
    ) -> Dict[str, List[Dict[str, Any]]]:
        result_map: Dict[str, List[Dict[str, Any]]] = {
            pid: [] for pid in parent_ids if pid
        }
        if not result_map:
            return result_map

        valid_pids = list(result_map.keys())
        chunk_size = 15

        for i in range(0, len(valid_pids), chunk_size):
            chunk = valid_pids[i : i + chunk_size]
            or_clause = " or ".join([f"'{pid}' in parents" for pid in chunk])
            q = f"({or_clause}) and trashed = false"
            files = self.api_list_files(q, page_size=600)
            for f in files:
                for p in f.get("parents", []):
                    if p in result_map:
                        result_map[p].append(f)

        return result_map

    @staticmethod
    def _extract_excel_candidate(item: Dict[str, Any]) -> Optional[Dict[str, str]]:
        name = item.get("name", "")
        if name.startswith("~$"):
            return None
        mime = item.get("mimeType", "")
        if mime == "application/vnd.google-apps.folder":
            return None

        target_id = item["id"]
        target_mime = mime
        if mime == "application/vnd.google-apps.shortcut":
            sc = item.get("shortcutDetails", {})
            target_id = sc.get("targetId", target_id)
            target_mime = sc.get("targetMimeType", mime)

        low_name = name.lower()
        is_excel = (
            low_name.endswith((".xlsx", ".xls", ".xlsm", ".csv"))
            or "spreadsheet" in target_mime
            or "excel" in target_mime
            or target_mime == "application/vnd.google-apps.spreadsheet"
        )
        if is_excel:
            return {"id": target_id, "name": name, "mimeType": target_mime}
        return None

    def get_excel_from_hierarchy(
        self,
        folder_id: Optional[str],
        query_hint: str = "",
        parent_folder_id: Optional[str] = None,
    ) -> Tuple[Optional[bytes], Optional[str]]:
        """
        Excel çeki listesini bulmak için:
        1. Ana Lot klasörünün içine bakar,
        2. Alt klasörlerin (2 katman derine kadar) tamamına bakar,
        3. Üst Taş klasörüne (parent_folder_id) bakar,
        4. Tüm Drive'da lot koduyla veya taş adıyla Excel/Google Sheet arar.
        """
        norm_q = normalize_lot_code(query_hint) if query_hint else ""

        if folder_id:
            lvl1 = self.batch_list_children([folder_id]).get(folder_id, [])
            sub_ids_1 = []
            for item in lvl1:
                cand = self._extract_excel_candidate(item)
                if cand:
                    b_data = self.download_file_bytes(cand["id"], cand["mimeType"])
                    if b_data:
                        fname = cand["name"]
                        if not fname.lower().endswith((".xlsx", ".xls", ".csv")):
                            fname += ".xlsx"
                        return b_data, fname
                if item.get("mimeType") == "application/vnd.google-apps.folder":
                    sub_ids_1.append(item["id"])

            if sub_ids_1:
                lvl2_map = self.batch_list_children(sub_ids_1[:15])
                sub_ids_2 = []
                for _, s_items in lvl2_map.items():
                    for item in s_items:
                        cand = self._extract_excel_candidate(item)
                        if cand:
                            b_data = self.download_file_bytes(
                                cand["id"], cand["mimeType"]
                            )
                            if b_data:
                                fname = cand["name"]
                                if not fname.lower().endswith(
                                    (".xlsx", ".xls", ".csv")
                                ):
                                    fname += ".xlsx"
                                return b_data, fname
                        if item.get("mimeType") == "application/vnd.google-apps.folder":
                            sub_ids_2.append(item["id"])

                if sub_ids_2:
                    lvl3_map = self.batch_list_children(sub_ids_2[:15])
                    for _, s3_items in lvl3_map.items():
                        for item in s3_items:
                            cand = self._extract_excel_candidate(item)
                            if cand:
                                b_data = self.download_file_bytes(
                                    cand["id"], cand["mimeType"]
                                )
                                if b_data:
                                    fname = cand["name"]
                                    if not fname.lower().endswith(
                                        (".xlsx", ".xls", ".csv")
                                    ):
                                        fname += ".xlsx"
                                    return b_data, fname

        # Üst klasörde (Taş klasöründe, örn: SILVER RIVER) Excel var mı kontrol et
        if parent_folder_id:
            p_items = self.batch_list_children([parent_folder_id]).get(
                parent_folder_id, []
            )
            for item in p_items:
                cand = self._extract_excel_candidate(item)
                if cand:
                    # Eğer birden fazla Excel varsa lot numarası eşleşeni tercih et
                    if not norm_q or exact_lot_match_in_text(cand["name"], norm_q) > 0:
                        b_data = self.download_file_bytes(
                            cand["id"], cand["mimeType"]
                        )
                        if b_data:
                            fname = cand["name"]
                            if not fname.lower().endswith(
                                (".xlsx", ".xls", ".csv")
                            ):
                                fname += ".xlsx"
                            return b_data, fname
            # Eşleşen yoksa üst klasördeki ilk genel stok Excel'ini al
            for item in p_items:
                cand = self._extract_excel_candidate(item)
                if cand:
                    b_data = self.download_file_bytes(cand["id"], cand["mimeType"])
                    if b_data:
                        fname = cand["name"]
                        if not fname.lower().endswith((".xlsx", ".xls", ".csv")):
                            fname += ".xlsx"
                        return b_data, fname

        # Tüm Drive'da bu lot koduyla kayıtlı Excel veya Google Sheet ara
        if norm_q:
            lot_m = re.match(r"^([A-Z]+)-(\d+)$", norm_q)
            terms = (
                [
                    f"{lot_m.group(1)}-{lot_m.group(2)}",
                    f"{lot_m.group(1)} {lot_m.group(2)}",
                    f"{lot_m.group(1)}{lot_m.group(2)}",
                    lot_m.group(2),
                ]
                if lot_m
                else [norm_q]
            )
            for term in terms:
                gq = (
                    f"trashed = false and name contains '{term}' and "
                    f"mimeType != 'application/vnd.google-apps.folder' and "
                    f"not mimeType contains 'image/'"
                )
                gfiles = self.api_list_files(gq, page_size=25)
                for gf in gfiles:
                    if exact_lot_match_in_text(gf.get("name", ""), norm_q) >= 85:
                        cand = self._extract_excel_candidate(gf)
                        if cand:
                            b_data = self.download_file_bytes(
                                cand["id"], cand["mimeType"]
                            )
                            if b_data:
                                fname = cand["name"]
                                if not fname.lower().endswith(
                                    (".xlsx", ".xls", ".csv")
                                ):
                                    fname += ".xlsx"
                                return b_data, fname

        return None, None

    def scan_images_deep(
        self, folder_id: Optional[str], query_term: str
    ) -> List[Dict[str, Any]]:
        """
        ASH GREY A-91 gibi fotoğrafların 'PHOTOS -> LARGE / CUSTOM' alt klasörlerinde
        bulunduğu durumlar için 3 katman derinliğe kadar tüm görselleri toplar!
        """
        all_imgs: List[Dict[str, Any]] = []
        if folder_id:
            lvl1 = self.batch_list_children([folder_id]).get(folder_id, [])
            sub_1 = []
            for it in lvl1:
                mime = it.get("mimeType", "")
                if mime == "application/vnd.google-apps.folder":
                    sub_1.append(it)
                elif mime.startswith("image/"):
                    all_imgs.append(it)

            if sub_1:
                sf1_map = {sf["id"]: sf["name"] for sf in sub_1[:12]}
                lvl2_batch = self.batch_list_children(list(sf1_map.keys()))
                sub_2 = []
                for sf_id, sf_items in lvl2_batch.items():
                    p_name = sf1_map.get(sf_id, "")
                    for it in sf_items:
                        mime = it.get("mimeType", "")
                        if mime == "application/vnd.google-apps.folder":
                            it_copy = dict(it)
                            it_copy["parent_chain"] = f"{p_name}/{it.get('name', '')}"
                            sub_2.append(it_copy)
                        elif mime.startswith("image/"):
                            img_copy = dict(it)
                            img_copy["parent_name"] = p_name
                            all_imgs.append(img_copy)

                if sub_2:
                    sf2_map = {
                        sf["id"]: sf.get("parent_chain", sf.get("name", ""))
                        for sf in sub_2[:15]
                    }
                    lvl3_batch = self.batch_list_children(list(sf2_map.keys()))
                    for sf2_id, sf2_items in lvl3_batch.items():
                        p_chain = sf2_map.get(sf2_id, "")
                        for img in sf2_items:
                            if img.get("mimeType", "").startswith("image/"):
                                img_copy = dict(img)
                                img_copy["parent_name"] = p_chain
                                all_imgs.append(img_copy)

        if not all_imgs and query_term:
            norm_q = normalize_lot_code(query_term)
            lot_m = re.match(r"^([A-Z]+)-(\d+)$", norm_q)
            cands = (
                [f"{lot_m.group(1)}-{lot_m.group(2)}", f"{lot_m.group(1)}{lot_m.group(2)}"]
                if lot_m
                else [norm_q]
            )
            for cand in cands:
                q = f"trashed = false and name contains '{cand}' and mimeType contains 'image/'"
                found = self.api_list_files(q, page_size=20)
                for img in found:
                    if exact_lot_match_in_text(img.get("name", ""), norm_q) > 0:
                        all_imgs.append(img)
                if all_imgs:
                    break

        return all_imgs

    @staticmethod
    def _score_slab_image_quality(img_meta: Dict[str, Any]) -> int:
        """
        En kaliteli plaka fotoğrafını seçmek için puanlama yapar:
        - 'CUSTOM' veya 'LARGE' klasöründeki (Arslanlar logolu stüdyo çekimi) fotoğraflara +50 puan verir.
        - Dosya adında 'Area', 'Block', 'Slab' geçen gerçek çerçeveli fotoğraflara +40 puan verir.
        - 'WEB' veya 'EMAIL' gibi kırpılmış/düşük kaliteli fotoğraflardan -25 puan kırar.
        """
        fname = img_meta.get("name", "").upper()
        pname = img_meta.get("parent_name", "").upper()
        score = 50

        if "CUSTOM" in pname or "CUSTOM" in fname:
            score += 45
        elif "LARGE" in pname or "LARGE" in fname or "ORJ" in pname:
            score += 35
        elif "PHOTOS" in pname and "WEB" not in pname and "EMAIL" not in pname:
            score += 25

        if any(k in fname for k in ["AREA", "BLOCK", "SLAB", "CM"]):
            score += 30

        if "WEB" in pname or "EMAIL" in pname or "SMALL" in pname:
            score -= 25

        size_bytes = int(img_meta.get("size", 0) or 0)
        if size_bytes > 500_000:
            score += 15

        return score

    def get_images_for_telegram(
        self, folder_id: Optional[str], query_term: str
    ) -> Tuple[Optional[bytes], Optional[bytes]]:
        all_imgs = self.scan_images_deep(folder_id, query_term)
        if not all_imgs:
            return None, None

        bm_files = []
        std_files = []
        for f in all_imgs:
            fn = f.get("name", "").lower()
            pn = f.get("parent_name", "").lower()
            if "bookmatch" in fn or "bm" in fn or "bookmatch" in pn:
                bm_files.append(f)
            else:
                std_files.append(f)

        std_files.sort(key=self._score_slab_image_quality, reverse=True)
        bm_files.sort(key=self._score_slab_image_quality, reverse=True)

        f1 = std_files[0] if std_files else (bm_files[0] if bm_files else None)
        f2 = (
            bm_files[0]
            if (bm_files and bm_files[0] != f1)
            else (std_files[1] if len(std_files) > 1 else f1)
        )

        img1 = (
            self.download_fast_image(
                f1["id"], f1.get("thumbnailLink"), for_pdf=False
            )
            if f1
            else None
        )
        img2 = (
            self.download_fast_image(
                f2["id"], f2.get("thumbnailLink"), for_pdf=False
            )
            if f2
            else None
        )
        return img1, img2

    def get_best_slab_image_for_pdf(
        self,
        folder_id: Optional[str],
        query_term: str,
        pre_imgs: Optional[List[Dict[str, Any]]] = None,
    ) -> Optional[bytes]:
        """
        PDF Sunumu İçin: Bookmatch Görüntülerini ASLA Almaz.
        CUSTOM / LARGE klasöründeki en kaliteli, gerçek plaka fotoğrafını seçer!
        """
        all_imgs = (
            pre_imgs
            if pre_imgs is not None and len(pre_imgs) > 0
            else self.scan_images_deep(folder_id, query_term)
        )
        if not all_imgs:
            return None

        clean_slabs = [
            f
            for f in all_imgs
            if "bookmatch" not in f.get("name", "").lower()
            and "bm" not in f.get("name", "").lower()
            and "bookmatch" not in f.get("parent_name", "").lower()
        ]
        candidates = clean_slabs if clean_slabs else all_imgs
        candidates.sort(key=self._score_slab_image_quality, reverse=True)
        target = candidates[0]
        return self.download_fast_image(
            target["id"], target.get("thumbnailLink"), for_pdf=True
        )

    def download_fast_image(
        self,
        file_id: Optional[str],
        thumb_url: Optional[str] = None,
        for_pdf: bool = False,
    ) -> Optional[bytes]:
        raw = None
        if thumb_url:
            try:
                # Yüksek kalite için =s1600 indir
                hi_res_url = re.sub(r"=s\d+", "=s1600", thumb_url)
                resp = requests.get(hi_res_url, timeout=8)
                if resp.status_code == 200 and len(resp.content) > 2000:
                    raw = resp.content
            except Exception:
                pass

        if not raw and file_id:
            raw = self.download_file_bytes(file_id)

        if not raw:
            return None
        return (
            prepare_pdf_card_image_bytes(raw)
            if for_pdf
            else prepare_telegram_image_bytes(raw)
        )


# =====================================================================
# 4. AKILLI ENVANTER MOTORU (TEKİL LOT BİRLEŞTİRME & KESİN METRAJ)
# =====================================================================
class InventoryEngine:
    def __init__(self, drive: DriveExplorer):
        self.drive = drive

    @staticmethod
    def _detect_columns(col_map: Dict[Any, str]) -> Dict[str, Any]:
        m2_col = next(
            (
                c
                for c, n in col_map.items()
                if "BRUT" not in n
                and "GROSS" not in n
                and "AGIRLIK" not in n
                and "WEIGHT" not in n
                and any(
                    k in n
                    for k in [
                        "NET M2",
                        "NET ALAN",
                        "M2",
                        "M²",
                        "ALAN",
                        "METRAJ",
                        "METREKARE",
                        "AREA",
                        "SQM",
                    ]
                )
            ),
            None,
        )
        if m2_col is None:
            m2_col = next(
                (
                    c
                    for c, n in col_map.items()
                    if "AGIRLIK" not in n
                    and "WEIGHT" not in n
                    and any(
                        k in n
                        for k in ["M2", "M²", "ALAN", "METRAJ", "AREA", "SQM"]
                    )
                ),
                None,
            )

        len_col = next(
            (
                c
                for c, n in col_map.items()
                if c != m2_col
                and any(
                    k in n
                    for k in ["UZUNLUK", "LENGHT", "LENGTH", "LONG", "EN CM LENGHT"]
                )
            ),
            None,
        )
        if len_col is None:
            len_col = next(
                (
                    c
                    for c, n in col_map.items()
                    if c != m2_col and "BOY" in n and "WID" not in n
                ),
                None,
            )

        height_col = next(
            (
                c
                for c, n in col_map.items()
                if c not in (m2_col, len_col)
                and (
                    any(
                        k in n
                        for k in [
                            "YUKSEKLIK",
                            "GENISLIK",
                            "WIDTH",
                            "WIDHT",
                            "HEIGHT",
                            "BOY",
                        ]
                    )
                    or bool(re.search(r"\bEN\b", n))
                )
            ),
            None,
        )

        dim_col = next(
            (
                c
                for c, n in col_map.items()
                if c not in (m2_col, len_col, height_col)
                and any(k in n for k in ["EBAT", "OLCU", "SIZE", "DIM"])
            ),
            None,
        )

        slab_col = next(
            (
                c
                for c, n in col_map.items()
                if c not in (m2_col, len_col, height_col)
                and any(
                    k in n
                    for k in ["PLAKA NO", "SLAB NUMBER", "SLAB NO", "PLAKA", "SLAB"]
                )
            ),
            None,
        )

        adet_col = next(
            (
                c
                for c, n in col_map.items()
                if c not in (m2_col, len_col, height_col, slab_col)
                and "M2" not in n
                and any(k in n for k in ["ADET", "PCS", "QTY", "QUANTITY"])
            ),
            None,
        )

        thick_col = next(
            (
                c
                for c, n in col_map.items()
                if c not in (m2_col, len_col, height_col)
                and any(k in n for k in ["KALIN", "THICK"])
            ),
            None,
        )

        surf_col = next(
            (
                c
                for c, n in col_map.items()
                if any(k in n for k in ["YUZEY", "FINISH", "SURF", "ISLEM"])
            ),
            None,
        )

        lot_col = next(
            (
                c
                for c, n in col_map.items()
                if c not in (m2_col, len_col, height_col, slab_col)
                and any(
                    k in n
                    for k in ["BLOK NO", "BLOCK NUMBER", "LOT", "BLOK", "BLOCK", "PARTI"]
                )
            ),
            None,
        )

        desc_col = next(
            (
                c
                for c, n in col_map.items()
                if any(
                    k in n
                    for k in ["URUN ADI", "PRODUCT DESCRIPTION", "CINSI", "STONE"]
                )
            ),
            None,
        )

        return {
            "m2": m2_col,
            "len": len_col,
            "height": height_col,
            "dim": dim_col,
            "slab": slab_col,
            "adet": adet_col,
            "thick": thick_col,
            "surf": surf_col,
            "lot": lot_col,
            "desc": desc_col,
        }

    def parse_packing_list_bytes(
        self, excel_bytes: bytes, target_lot: str = ""
    ) -> Optional[Dict[str, Any]]:
        norm_target_lot = normalize_lot_code(target_lot) if target_lot else ""
        try:
            wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
            sheet_names = list(wb.sheetnames)
            if norm_target_lot:
                sheet_names.sort(
                    key=lambda s: 0
                    if exact_lot_match_in_text(s, norm_target_lot) >= 85
                    else 1
                )

            for s_name in sheet_names:
                ws = wb[s_name]
                # Birleştirilmiş hücreleri (Merged Cells - Örn: Blok No C14:C25) çöz ve tüm alt hücrelere kopyala!
                for m_range in list(ws.merged_cells.ranges):
                    top_val = ws.cell(m_range.min_row, m_range.min_col).value
                    ws.unmerge_cells(str(m_range))
                    for r_idx in range(m_range.min_row, m_range.max_row + 1):
                        for c_idx in range(m_range.min_col, m_range.max_col + 1):
                            ws.cell(r_idx, c_idx).value = top_val

                rows_matrix = list(ws.iter_rows(values_only=True))
                if not rows_matrix or len(rows_matrix) < 2:
                    continue

                raw_df = pd.DataFrame(rows_matrix)
                raw_df.dropna(how="all", inplace=True)
                raw_df.reset_index(drop=True, inplace=True)
                if raw_df.empty or len(raw_df) < 2:
                    continue

                header_idx = 0
                best_score = 0
                header_keywords = [
                    "M2",
                    "M²",
                    "SQM",
                    "ALAN",
                    "AREA",
                    "METRAJ",
                    "QUANTITY",
                    "BOY",
                    "UZUNLUK",
                    "LENGTH",
                    "LENGHT",
                    "YUKSEKLIK",
                    "GENISLIK",
                    "WIDTH",
                    "WIDHT",
                    "HEIGHT",
                    "PLAKA",
                    "SLAB",
                    "ADET",
                    "BLOK",
                    "BLOCK",
                ]

                for idx in range(min(30, len(raw_df))):
                    row_vals = [
                        normalize_turkish_key(v)
                        for v in raw_df.iloc[idx].values
                        if pd.notna(v)
                    ]
                    score = sum(
                        1
                        for v in row_vals
                        if any(kw == v or kw in v for kw in header_keywords)
                        or bool(re.search(r"\bEN\b", v))
                    )
                    if score > best_score:
                        best_score = score
                        header_idx = idx

                composite_headers = {}
                for col_i in range(raw_df.shape[1]):
                    h_val = normalize_turkish_key(raw_df.iloc[header_idx, col_i])
                    if (
                        header_idx + 1 < len(raw_df)
                        and any(
                            k in h_val for k in ["EBAT", "OLCU", "DIMENSION", "SIZE"]
                        )
                    ):
                        sub_h = normalize_turkish_key(
                            raw_df.iloc[header_idx + 1, col_i]
                        )
                        if sub_h and not re.match(r"^\d+([.,]\d+)?$", sub_h):
                            h_val = f"{h_val} {sub_h}"
                    composite_headers[col_i] = h_val if h_val else f"COL_{col_i}"

                mapped = self._detect_columns(composite_headers)

                data_start = header_idx + 1
                df = raw_df.iloc[data_start:].copy()
                df.columns = list(range(raw_df.shape[1]))

                if (
                    mapped["m2"] is None
                    or mapped["len"] is None
                    or mapped["height"] is None
                ):
                    dim_cands = []
                    m2_cands = []
                    for col_i in df.columns:
                        nums = [
                            safe_float(v)
                            for v in df[col_i].dropna().values
                            if 0 < safe_float(v) < 4500
                        ]
                        if len(nums) >= 2:
                            sample_nums = nums[:-1] if len(nums) > 3 else nums
                            avg_n = sum(sample_nums) / len(sample_nums)
                            if (80 <= avg_n <= 380) or (800 <= avg_n <= 3800):
                                dim_cands.append((avg_n, col_i))
                            elif 1.0 <= avg_n <= 12.0 and any(
                                (n % 1) != 0 for n in sample_nums
                            ):
                                m2_cands.append(col_i)

                    if len(dim_cands) >= 2 and (
                        mapped["len"] is None or mapped["height"] is None
                    ):
                        dim_cands.sort(key=lambda x: x[0], reverse=True)
                        if mapped["len"] is None:
                            mapped["len"] = dim_cands[0][1]
                        if mapped["height"] is None:
                            for _, c_i in dim_cands[1:]:
                                if c_i != mapped["len"]:
                                    mapped["height"] = c_i
                                    break
                    if mapped["m2"] is None and m2_cands:
                        mapped["m2"] = m2_cands[-1]

                m2_col = mapped["m2"]
                len_col = mapped["len"]
                height_col = mapped["height"]
                dim_col = mapped["dim"]
                slab_col = mapped["slab"]
                adet_col = mapped["adet"]
                thick_col = mapped["thick"]
                surf_col = mapped["surf"]
                lot_col = mapped["lot"]
                desc_col = mapped["desc"]

                # Eğer çok bloklu bir çeki listesi ise (Örn: ARS-48-02 içindeki A-105, A-265) ve aranan bir lot varsa filtrele
                rows_to_iter = df
                filtered_by_lot = False
                if norm_target_lot and lot_col is not None:
                    lot_filtered = df[
                        df[lot_col].apply(
                            lambda v: exact_lot_match_in_text(
                                str(v), norm_target_lot
                            )
                            >= 85
                        )
                    ]
                    if not lot_filtered.empty:
                        rows_to_iter = lot_filtered
                        filtered_by_lot = True

                valid_rows = []
                sum_slab_m2 = 0.0
                sum_slabs = 0
                explicit_total_row_m2 = 0.0
                lengths = []
                heights = []
                detected_stone_desc = ""

                for _, row in rows_to_iter.iterrows():
                    row_cells_norm = [
                        normalize_turkish_key(v) for v in row.values if pd.notna(v)
                    ]
                    row_str = " ".join(row_cells_norm)

                    if any(
                        tot_w in row_str
                        for tot_w in [
                            "TOPLAM",
                            "TOTAL",
                            "GENEL TOPLAM",
                            "SUBTOTAL",
                            "PORT OF",
                            "AMBALAJ",
                            "PACKAGE",
                            "KAP SAYISI",
                        ]
                    ):
                        if not filtered_by_lot and m2_col is not None and (
                            "TOPLAM" in row_str or "TOTAL" in row_str
                        ):
                            t_val = safe_float(row.get(m2_col))
                            if t_val > 1000:
                                t_val = t_val / 10000.0
                            if t_val > explicit_total_row_m2:
                                explicit_total_row_m2 = t_val
                        continue

                    if (
                        not detected_stone_desc
                        and desc_col is not None
                        and pd.notna(row.get(desc_col))
                    ):
                        detected_stone_desc = normalize_turkish_key(
                            row.get(desc_col)
                        )

                    raw_l = (
                        safe_float(row.get(len_col)) if len_col is not None else 0.0
                    )
                    raw_h = (
                        safe_float(row.get(height_col))
                        if height_col is not None
                        else 0.0
                    )

                    if (
                        (raw_l <= 0 or raw_h <= 0)
                        and dim_col is not None
                        and pd.notna(row.get(dim_col))
                    ):
                        dm = re.search(
                            r"(\d+(?:[.,]\d+)?)\s*[xX*]\s*(\d+(?:[.,]\d+)?)",
                            str(row.get(dim_col)),
                        )
                        if dm:
                            raw_l = safe_float(dm.group(1))
                            raw_h = safe_float(dm.group(2))

                    l_cm = normalize_cm(raw_l)
                    h_cm = normalize_cm(raw_h)

                    # Uzun kenar her zaman Boy (Length), kısa kenar En (Width/Height) olsun
                    if l_cm > 0 and h_cm > l_cm:
                        l_cm, h_cm = h_cm, l_cm

                    cell_m2 = (
                        safe_float(row.get(m2_col)) if m2_col is not None else 0.0
                    )
                    if cell_m2 > 1000:
                        cell_m2 = cell_m2 / 10000.0

                    calc_m2 = (
                        (l_cm * h_cm) / 10000.0 if (l_cm > 20 and h_cm > 20) else 0.0
                    )

                    if (
                        not filtered_by_lot
                        and l_cm == 0
                        and h_cm == 0
                        and len(valid_rows) >= 2
                        and cell_m2 > sum_slab_m2 * 0.75
                    ):
                        explicit_total_row_m2 = cell_m2
                        continue

                    if 0.15 <= cell_m2 <= 25.0:
                        m2_val = cell_m2
                    elif 0.15 <= calc_m2 <= 25.0:
                        m2_val = calc_m2
                    else:
                        continue

                    adet_val = 1
                    if adet_col is not None:
                        cand_adet = int(safe_float(row.get(adet_col)))
                        if 1 <= cand_adet <= 10:
                            adet_val = cand_adet

                    raw_slab_no = (
                        row.get(slab_col, len(valid_rows) + 1)
                        if slab_col is not None and pd.notna(row.get(slab_col))
                        else len(valid_rows) + 1
                    )
                    if isinstance(raw_slab_no, float) and raw_slab_no.is_integer():
                        slab_no_str = str(int(raw_slab_no))
                    else:
                        slab_no_str = str(raw_slab_no).strip()
                        if slab_no_str.endswith(".0"):
                            slab_no_str = slab_no_str[:-2]

                    sum_slab_m2 += m2_val
                    sum_slabs += adet_val
                    if l_cm > 20:
                        lengths.append(l_cm)
                    if h_cm > 20:
                        heights.append(h_cm)

                    valid_rows.append(
                        {
                            "slab_no": slab_no_str,
                            "length_cm": int(round(l_cm)) if l_cm > 0 else 0,
                            "height_cm": int(round(h_cm)) if h_cm > 0 else 0,
                            "thickness": "2 cm",
                            "m2": round(m2_val, 4),
                        }
                    )

                final_m2 = sum_slab_m2
                if not filtered_by_lot and explicit_total_row_m2 > 0:
                    if sum_slab_m2 == 0 or (
                        0.85 <= (explicit_total_row_m2 / sum_slab_m2) <= 1.15
                    ):
                        final_m2 = explicit_total_row_m2

                if final_m2 <= 0:
                    continue

                thick_val = "2 cm"
                if thick_col is not None and not rows_to_iter[thick_col].dropna().empty:
                    raw_th = safe_float(rows_to_iter[thick_col].dropna().iloc[0])
                    if 17 <= raw_th <= 24 or raw_th == 2:
                        thick_val = "2 cm"
                    elif 27 <= raw_th <= 34 or raw_th == 3:
                        thick_val = "3 cm"
                    elif 0 < raw_th <= 10:
                        thick_val = f"{raw_th:g} cm"
                elif "3CM" in detected_stone_desc or "3 CM" in detected_stone_desc:
                    thick_val = "3 cm"

                surf_val = "Polished"
                if surf_col is not None and not rows_to_iter[surf_col].dropna().empty:
                    surf_val = (
                        str(rows_to_iter[surf_col].dropna().iloc[0]).strip().title()
                    )

                if lengths and heights:
                    avg_l = int(round(sum(lengths) / len(lengths)))
                    avg_h = int(round(sum(heights) / len(heights)))
                    dims_str = f"{avg_l} x {avg_h} cm"
                elif lengths:
                    avg_l = int(round(sum(lengths) / len(lengths)))
                    dims_str = f"{avg_l} cm"
                else:
                    dims_str = "Standard Slab Sizes"

                return {
                    "total_m2": round(final_m2, 2),
                    "total_slabs": sum_slabs if sum_slabs > 0 else len(valid_rows),
                    "thickness": thick_val,
                    "surface": surf_val,
                    "dimensions": dims_str,
                    "stone_desc": detected_stone_desc,
                    "clean_rows": valid_rows,
                }
        except Exception as e:
            logger.warning(f"Excel okuma uyarısı: {e}")
        return None

    def _build_stock_from_folder_photos(
        self, all_imgs: List[Dict[str, Any]]
    ) -> Optional[Dict[str, Any]]:
        """
        Eğer bir lot klasöründe henüz Excel dosyası yoksa, klasördeki plaka
        fotoğraflarından (CUSTOM / LARGE alt klasörlerinde aynı plakanın birden fazla
        kopyası varsa tekilleştirerek!) gerçek plaka sayısını ve ebat/m2'yi çıkarır.
        """
        if not all_imgs:
            return None

        clean_slabs = [
            f
            for f in all_imgs
            if "bookmatch" not in f.get("name", "").lower()
            and "bm" not in f.get("name", "").lower()
            and "bookmatch" not in f.get("parent_name", "").lower()
        ]
        target_imgs = clean_slabs if clean_slabs else all_imgs

        # Eğer hem CUSTOM hem WEB hem LARGE klasörü taranmışsa aynı plaka 3 kez sayılmasın!
        # En çok plaka barındıran tek bir alt klasörü (örn: CUSTOM veya LARGE) baz al.
        by_parent: Dict[str, List[Dict[str, Any]]] = {}
        for img in target_imgs:
            p_key = img.get("parent_name", "ROOT")
            by_parent.setdefault(p_key, []).append(img)

        best_group = max(by_parent.values(), key=len) if by_parent else target_imgs
        slab_count = len(best_group)
        if slab_count <= 0:
            return None

        parsed_l = []
        parsed_h = []
        clean_rows = []

        for idx, f in enumerate(best_group, start=1):
            fname = f.get("name", "")
            dim_m = re.search(r"(\d{2,4})\s*[xX*]\s*(\d{2,4})", fname)
            area_m = re.search(
                r"(\d+[.,]\d+)\s*(?:m2|m²|sqm)", fname, re.IGNORECASE
            )
            l_cm = normalize_cm(safe_float(dim_m.group(1))) if dim_m else 295.0
            h_cm = normalize_cm(safe_float(dim_m.group(2))) if dim_m else 182.0
            if l_cm > 0 and h_cm > l_cm:
                l_cm, h_cm = h_cm, l_cm

            m2_v = (
                safe_float(area_m.group(1))
                if area_m
                else round((l_cm * h_cm) / 10000.0, 2)
            )
            if m2_v <= 0 or m2_v > 20:
                m2_v = round((l_cm * h_cm) / 10000.0, 2)

            parsed_l.append(l_cm)
            parsed_h.append(h_cm)
            clean_rows.append(
                {
                    "slab_no": str(idx),
                    "length_cm": int(round(l_cm)),
                    "height_cm": int(round(h_cm)),
                    "thickness": "2 cm",
                    "m2": m2_v,
                }
            )

        total_m2 = round(sum(r["m2"] for r in clean_rows), 2)
        avg_l = int(round(sum(parsed_l) / len(parsed_l)))
        avg_h = int(round(sum(parsed_h) / len(parsed_h)))

        return {
            "total_m2": total_m2,
            "total_slabs": slab_count,
            "thickness": "2 cm",
            "surface": "Polished",
            "dimensions": f"{avg_l} x {avg_h} cm",
            "stone_desc": "",
            "clean_rows": clean_rows,
        }

    def get_lot_info(self, query: str) -> Dict[str, Any]:
        norm_key = normalize_lot_code(query)
        now = time.time()

        if norm_key in self.drive._lot_result_cache:
            ts, cached_res = self.drive._lot_result_cache[norm_key]
            if now - ts < 600:
                return cached_res

        folder_id, folder_link, parent_id, actual_name = (
            self.drive.find_folder_exact(norm_key)
        )

        parent_stone_name = ""
        if parent_id:
            p_meta = self.drive.get_folder_metadata(parent_id)
            if p_meta:
                parent_stone_name = p_meta.get("name", "")

        clean_disp, stone_name, lot_code = clean_stone_and_lot_title(
            actual_name or norm_key, parent_stone_name
        )

        excel_bytes, excel_name = self.drive.get_excel_from_hierarchy(
            folder_id, query_hint=norm_key, parent_folder_id=parent_id
        )
        data = None
        if excel_bytes:
            data = self.parse_packing_list_bytes(excel_bytes, target_lot=norm_key)

        all_imgs = self.drive.scan_images_deep(folder_id, norm_key)

        if not data or data.get("total_m2", 0.0) <= 0:
            data = self._build_stock_from_folder_photos(all_imgs)

        if not data:
            data = {
                "total_m2": 0.0,
                "total_slabs": 0,
                "thickness": "2 cm",
                "surface": "Polished",
                "dimensions": "Standard Slab Sizes",
                "stone_desc": "",
                "clean_rows": [],
            }

        if data.get("stone_desc") and stone_name == "ARSLANLAR MARBLE":
            _, parsed_st, _ = clean_stone_and_lot_title(data["stone_desc"])
            if parsed_st:
                stone_name = parsed_st
                clean_disp = f"{stone_name} — {lot_code}"

        res = {
            "status": "success",
            "query": norm_key,
            "title": clean_disp,
            "stone_name": stone_name,
            "lot_code": lot_code,
            "drive_link": folder_link,
            "folder_id": folder_id,
            "excel_name": excel_name or f"{lot_code}_Original_PackingList.xlsx",
            "excel_bytes": excel_bytes,
            "data": data,
            "imgs": all_imgs,
        }
        self.drive._lot_result_cache[norm_key] = (now, res)
        return res

    def scan_collection_lots(
        self, collection_name: str, max_lots: int = 20
    ) -> List[Dict[str, Any]]:
        """
        Koleksiyon klasörünü tararken 'CUSTOM', 'WEB', 'LARGE', 'EMAIL', 'PHOTOS'
        alt klasörlerini kesin olarak filtreler ve aynı lot kodlu (örn: A-468, A-91)
        klasörleri TEK BİR KART olarak birleştirir!
        """
        fid, _, _, _ = self.drive.find_folder_exact(collection_name)
        if not fid:
            return []

        lvl1_items = self.drive.batch_list_children([fid]).get(fid, [])
        lvl1_folders = [
            f
            for f in lvl1_items
            if f.get("mimeType") == "application/vnd.google-apps.folder"
            and not is_utility_subfolder(f.get("name", ""))
        ]
        if not lvl1_folders:
            return []

        lvl1_map = {f["id"]: f for f in lvl1_folders}
        lvl2_batch = self.drive.batch_list_children(list(lvl1_map.keys()))

        candidate_lot_folders = []

        for f1_id, children in lvl2_batch.items():
            f1_meta = lvl1_map[f1_id]
            f1_name = f1_meta.get("name", "")

            # Alt klasörlerden 'CUSTOM', 'WEB', 'LARGE', 'EMAIL', 'PHOTOS' OLMAYAN gerçek lot klasörlerini ayır
            real_lot_subfolders = [
                c
                for c in children
                if c.get("mimeType") == "application/vnd.google-apps.folder"
                and not is_utility_subfolder(c.get("name", ""))
            ]

            if real_lot_subfolders:
                for sf in real_lot_subfolders:
                    candidate_lot_folders.append(
                        {
                            "id": sf["id"],
                            "raw_name": sf.get("name", ""),
                            "parent_stone": f1_name,
                            "parent_id": f1_id,
                            "link": sf.get(
                                "webViewLink",
                                f"https://drive.google.com/drive/folders/{sf['id']}",
                            ),
                        }
                    )
            else:
                # Alt klasörlerin tamamı PHOTOS / CUSTOM / WEB / LARGE ise, asıl lot klasörü f1'in kendisidir!
                candidate_lot_folders.append(
                    {
                        "id": f1_id,
                        "raw_name": f1_name,
                        "parent_stone": "",
                        "parent_id": fid,
                        "link": f1_meta.get(
                            "webViewLink",
                            f"https://drive.google.com/drive/folders/{f1_id}",
                        ),
                    }
                )

        # Tekilleştirme (Deduplication): Aynı Lot Koduna (Örn: A-468, A-91) sahip birden fazla kayıt asla oluşmasın!
        unique_lots: Dict[str, Dict[str, Any]] = {}
        for cf in candidate_lot_folders:
            disp_title, st_name, l_code = clean_stone_and_lot_title(
                cf["raw_name"], cf["parent_stone"]
            )
            dedup_key = l_code if re.match(r"^[A-Z]+-\d+$", l_code) else disp_title
            if dedup_key not in unique_lots:
                cf["clean_title"] = disp_title
                cf["stone_name"] = st_name
                cf["lot_code"] = l_code
                unique_lots[dedup_key] = cf

        results = []
        for dedup_key, lf in list(unique_lots.items())[:max_lots]:
            ex_bytes, _ = self.drive.get_excel_from_hierarchy(
                lf["id"],
                query_hint=lf["lot_code"],
                parent_folder_id=lf["parent_id"],
            )
            parsed_data = None
            if ex_bytes:
                parsed_data = self.parse_packing_list_bytes(
                    ex_bytes, target_lot=lf["lot_code"]
                )

            imgs = self.drive.scan_images_deep(lf["id"], lf["lot_code"])

            if not parsed_data or parsed_data.get("total_m2", 0.0) <= 0:
                parsed_data = self._build_stock_from_folder_photos(imgs)

            if not parsed_data:
                parsed_data = {
                    "total_m2": 0.0,
                    "total_slabs": 0,
                    "thickness": "2 cm",
                    "surface": "Polished",
                    "dimensions": "Standard Slab Sizes",
                    "stone_desc": "",
                    "clean_rows": [],
                }

            results.append(
                {
                    "title": lf["clean_title"],
                    "stone_name": lf["stone_name"],
                    "lot_code": lf["lot_code"],
                    "id": lf["id"],
                    "link": lf["link"],
                    "data": parsed_data,
                    "imgs": imgs,
                }
            )
        return results

    def generate_deep_scan_report(self) -> str:
        collections = [
            "Arslanlar Collection",
            "Arslanlar Series",
            "Arslanlar Exclusive Collection",
        ]
        report_lines = ["🔍 <b>ARGON MASTER - DEEP SCAN STOK RAPORU</b> 🔍\n"]
        total_factory_m2 = 0.0
        total_factory_slabs = 0

        for coll in collections:
            items = self.scan_collection_lots(coll, max_lots=35)
            if not items:
                continue

            report_lines.append(f"📁 <b>{coll.upper()}</b>")
            coll_m2 = 0.0
            coll_slabs = 0

            for it in items:
                d = it["data"]
                if d["total_m2"] > 0:
                    report_lines.append(
                        f"  ├ <b>{it['title']}</b>: <code>{d['total_m2']:,.2f} m²</code> ({d['total_slabs']} Plaka)"
                    )
                    coll_m2 += d["total_m2"]
                    coll_slabs += d["total_slabs"]
                    total_factory_m2 += d["total_m2"]
                    total_factory_slabs += d["total_slabs"]

            report_lines.append(
                f"  └ <i>Koleksiyon Toplamı: {coll_m2:,.2f} m² ({coll_slabs} Plaka)</i>\n"
            )

        report_lines.append(
            f"🌐 <b>GENEL FABRİKA TOPLAMI: {total_factory_m2:,.2f} m² | {total_factory_slabs} Plaka</b>"
        )
        return "\n".join(report_lines)


# =====================================================================
# 5. KURUMSAL SUNUM PDF & BUNDLE'LI PACKING LIST (EXCEL + PDF) MOTORU
# =====================================================================
class CorporateDocumentEngine:
    def __init__(self, inventory: InventoryEngine, drive: DriveExplorer):
        self.inventory = inventory
        self.drive = drive

    def _build_header_flowables(self, pdf_title: str, styles: Any) -> List[Any]:
        BRONZE = colors.HexColor("#9C5933")
        SLATE = colors.HexColor("#2B2D42")

        logo_bytes = self.drive.get_official_logo_jpeg()
        rl_logo = RLImage(io.BytesIO(logo_bytes), width=145, height=86.6)
        rl_logo.hAlign = "CENTER"

        sub_s = ParagraphStyle(
            "SubTitleLogo",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=9.5,
            leading=13,
            textColor=SLATE,
            alignment=1,
            spaceBefore=4,
            spaceAfter=6,
        )
        return [
            rl_logo,
            Paragraph(pdf_title, sub_s),
            HRFlowable(width="100%", thickness=1.5, color=BRONZE, spaceAfter=14),
        ]

    def generate_presentation_pdf(
        self,
        queries: List[str],
        collection_folder_names: Optional[List[str]] = None,
    ) -> Tuple[bytes, str]:
        cards: List[Dict[str, Any]] = []
        seen_lot_keys = set()

        if collection_folder_names:
            titles_norm = [
                normalize_turkish_key(c) for c in collection_folder_names
            ]
            pdf_title = f"{' & '.join(titles_norm)} — EXECUTIVE CATALOG PRESENTATION"
            per_coll = max(8, int(24 / len(collection_folder_names)))
            for coll_name in collection_folder_names:
                scanned = self.inventory.scan_collection_lots(
                    coll_name, max_lots=per_coll
                )
                for item in scanned:
                    dedup_k = item["lot_code"]
                    if dedup_k in seen_lot_keys:
                        continue
                    seen_lot_keys.add(dedup_k)

                    img = self.drive.get_best_slab_image_for_pdf(
                        item["id"], item["lot_code"], pre_imgs=item.get("imgs")
                    )
                    cards.append(
                        {
                            "title": item["title"],
                            "stone_name": item["stone_name"],
                            "lot_code": item["lot_code"],
                            "data": item["data"],
                            "drive_link": item["link"],
                            "image_bytes": img,
                        }
                    )
        else:
            pdf_title = "EXECUTIVE STOCK & SLAB PRESENTATION"
            for q in queries:
                info = self.inventory.get_lot_info(q)
                dedup_k = info["lot_code"]
                if dedup_k in seen_lot_keys:
                    continue
                seen_lot_keys.add(dedup_k)

                img = self.drive.get_best_slab_image_for_pdf(
                    info["folder_id"], q, pre_imgs=info.get("imgs")
                )
                cards.append(
                    {
                        "title": info["title"],
                        "stone_name": info["stone_name"],
                        "lot_code": info["lot_code"],
                        "data": info["data"],
                        "drive_link": info["drive_link"],
                        "image_bytes": img,
                    }
                )

        if not cards:
            raise ValueError("Belirtilen kriterler için klasör bulunamadı.")

        pdf_io = io.BytesIO()
        doc = SimpleDocTemplate(
            pdf_io,
            pagesize=A4,
            rightMargin=32,
            leftMargin=32,
            topMargin=22,
            bottomMargin=26,
        )
        styles = getSampleStyleSheet()
        NAVY = colors.HexColor("#14213D")
        BRONZE = colors.HexColor("#9C5933")
        SLATE = colors.HexColor("#2B2D42")
        LIGHT_BG = colors.HexColor("#F8F9FA")
        BORD = colors.HexColor("#DEE2E6")

        ch_s = ParagraphStyle(
            "CH",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=11.5,
            leading=15,
            textColor=NAVY,
            spaceAfter=5,
        )
        th_s = ParagraphStyle(
            "TH",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=8.5,
            leading=11,
            textColor=colors.whitesmoke,
        )
        tdl_s = ParagraphStyle(
            "TDL",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=8.5,
            leading=11,
            textColor=SLATE,
        )
        tdv_s = ParagraphStyle(
            "TDV",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=8.5,
            leading=11,
            textColor=SLATE,
        )
        footer_s = ParagraphStyle(
            "FT",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=7.5,
            leading=10,
            textColor=colors.HexColor("#6C757D"),
            alignment=1,
        )

        story: List[Any] = self._build_header_flowables(pdf_title, styles)

        for idx, card in enumerate(cards, start=1):
            d = card["data"]
            m2 = d.get("total_m2", 0.0)
            slabs = d.get("total_slabs", 0)
            bundles_cnt = max(1, slabs // 12) if slabs >= 12 else (1 if slabs > 0 else 0)

            if m2 > 0:
                stock_str = f"<b>{m2:,.2f} m²  ({slabs} Slabs / {bundles_cnt} Bundles)</b>"
                heading_badge = (
                    f" — <font color='#9C5933'>{m2:,.2f} m² | {slabs} Slabs</font>"
                )
            else:
                stock_str = "<b>Available in Factory Stock</b>"
                heading_badge = ""

            surf_str = d.get("surface", "Polished") or "Polished"
            thick_str = d.get("thickness", "2 cm") or "2 cm"
            surf_thick_str = f"{surf_str} | {thick_str}"
            dims_str = d.get("dimensions", "Standard Slab Sizes")
            if not dims_str or dims_str == "-":
                dims_str = "Standard Slab Sizes"

            stone_block_str = f"<b>{card['stone_name']}</b> (Block / Lot: <b>{card['lot_code']}</b>)"

            rl_img = Paragraph("<b>ARSLANLAR MARBLE</b><br/>High-Res Slab Selection", tdv_s)
            if card.get("image_bytes"):
                try:
                    rl_img = RLImage(
                        io.BytesIO(card["image_bytes"]), width=185, height=118
                    )
                except Exception:
                    pass

            dlink = card.get("drive_link")
            link_html = (
                f"<link href='{dlink}' color='#0056b3'><u>Click to Open High-Res Slab Gallery</u></link>"
                if dlink
                else "Available on Request"
            )

            # Active Code yerine net ve anlaşılır Stone & Block Number ve Avg. Slab Size eklendi
            spec_rows = [
                [Paragraph("Specification", th_s), Paragraph("Technical Data", th_s)],
                [
                    Paragraph("Stone & Block No", tdl_s),
                    Paragraph(stone_block_str, tdv_s),
                ],
                [
                    Paragraph("Available Stock", tdl_s),
                    Paragraph(stock_str, tdv_s),
                ],
                [
                    Paragraph("Average Slab Size", tdl_s),
                    Paragraph(f"{dims_str}  (Min. 12 Slabs / Bundle)", tdv_s),
                ],
                [
                    Paragraph("Surface & Thickness", tdl_s),
                    Paragraph(surf_thick_str, tdv_s),
                ],
                [
                    Paragraph("High-Res Slab Photos", tdl_s),
                    Paragraph(link_html, tdv_s),
                ],
            ]

            st = Table(spec_rows, colWidths=[125, 205])
            st.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                        ("BACKGROUND", (0, 1), (-1, -1), LIGHT_BG),
                        ("GRID", (0, 0), (-1, -1), 0.5, BORD),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("TOPPADDING", (0, 0), (-1, -1), 4.5),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5),
                        ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ]
                )
            )

            mr = Table([[rl_img, st]], colWidths=[195, 335])
            mr.setStyle(
                TableStyle(
                    [
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("ALIGN", (0, 0), (0, 0), "CENTER"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 0),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ]
                )
            )

            story.append(
                KeepTogether(
                    [
                        Paragraph(f"{idx}. {card['title']}{heading_badge}", ch_s),
                        mr,
                        Spacer(1, 14),
                    ]
                )
            )

        story.append(HRFlowable(width="100%", thickness=0.8, color=BRONZE, spaceAfter=6))
        story.append(
            Paragraph(
                "<b>ARSLANLAR MARBLE</b> | International Sales & Export Department | Menderes / Izmir, Turkey",
                footer_s,
            )
        )

        doc.build(story)
        pdf_io.seek(0)
        pdf_bytes = pdf_io.read()
        gc.collect()

        if collection_folder_names:
            fname = f"Arslanlar_Collections_{datetime.now().strftime('%H%M')}.pdf"
        else:
            fname = f"Arslanlar_Presentation_{'_'.join(queries[:3])}.pdf"
        return pdf_bytes, fname

    def generate_draft_format_packing_excel(
        self,
        lot_infos: List[Dict[str, Any]],
        customer_code: str = "EXPORT SELECTION",
    ) -> bytes:
        """
        Kullanıcının gönderdiği 'PACKING LIST DRAFT 29.07.2026 ARS-48-02.xlsx'
        dosyasının BİREBİR AYNISI olan (Minimum 12'şerli ahşap bundle birleştirmeli,
        =H*57 kg ağırlık formüllü, bronz logolu, ALIAGA PORT alt bilgili) Excel üretir!
        """
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sayfa1"

        col_widths = {
            "A": 3.5,
            "B": 44.5,
            "C": 15.0,
            "D": 20.0,
            "E": 14.5,
            "F": 14.0,
            "G": 12.5,
            "H": 13.5,
            "I": 13.5,
            "J": 15.5,
            "K": 19.0,
        }
        for col_l, w in col_widths.items():
            ws.column_dimensions[col_l].width = w

        # Sol üst köşeye (B2) gömülü Arslanlar Marble Bronz Aslan Logosunu ekle
        try:
            logo_png = self.drive.get_official_logo_png()
            xl_img = XLImage(io.BytesIO(logo_png))
            xl_img.width = 240
            xl_img.height = 143
            ws.add_image(xl_img, "B2")
        except Exception as e:
            logger.warning(f"Excel logo ekleme uyarısı: {e}")

        font_company = Font(name="Calibri", size=11, bold=True, color="14213D")
        font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        font_bold = Font(name="Calibri", size=11, bold=True)
        font_regular = Font(name="Calibri", size=11, bold=False)

        fill_header = PatternFill(
            start_color="14213D", end_color="14213D", fill_type="solid"
        )
        fill_total = PatternFill(
            start_color="F2E6D8", end_color="F2E6D8", fill_type="solid"
        )
        fill_zebra = PatternFill(
            start_color="F8F9FA", end_color="F8F9FA", fill_type="solid"
        )

        thin_side = Side(border_style="thin", color="BFBFBF")
        med_side = Side(border_style="medium", color="14213D")
        border_data = Border(
            left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
        )
        border_box = Border(
            left=med_side, right=med_side, top=med_side, bottom=med_side
        )

        align_center = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        align_left = Alignment(
            horizontal="left", vertical="center", wrap_text=True
        )

        # Sağ Üst Şirket Anteti (Satır 6 - 11, H Sütunu - Birebir Şablon)
        ws["H6"] = "ARSLANLAR TAAHHUT INS. HAFR.MAD.SAN.TİC.LTD.ŞTİ"
        ws["H7"] = "ADDRESS: KasımPasa mah. Ataturk cad. No:99/11"
        ws["H8"] = "Menderes/Izmir/Turkey"
        ws["H9"] = "MENDERES VD.0840060648 TURKEY"
        ws["B10"] = f"Customer Name: {customer_code}"
        ws.merge_cells("B10:C11")
        ws["B10"].font = font_bold
        ws["B10"].alignment = align_left

        ws["H10"] = f"INVOICE NO: EXP-{datetime.now().strftime('%Y%m%d')}"
        ws["H11"] = f"DATE: {datetime.now().strftime('%d.%m.%Y')}"
        for r_i in range(6, 12):
            ws[f"H{r_i}"].font = font_company

        # Satır 13: Tablo Başlıkları (Birebir Şablon)
        headers = [
            "ÜRÜN ADI / PRODUCT DESCRIPTION",
            "BLOK NO / BLOCK NUMBER",
            "PLAKA NO / SLAB NUMBER",
            "EN (cm) / LENGHT (cm)",
            "BOY (cm) / WIDHT (cm)",
            "MİKTAR (adet) / QUANTITY (pcs)",
            "MİKTAR (m2) / QUANTITY (m2)",
            "NET AĞIRLIK (kg) / NET WEIGHT (kg)",
            "BRÜT AĞIRLIK (kg) / GROSS WEIGHT (kg)",
            "NUMBER OF BUNDLE",
        ]
        ws.row_dimensions[13].height = 32
        for idx, h_text in enumerate(headers, start=2):
            cell = ws.cell(13, idx, h_text)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = border_box

        current_row = 14
        first_data_row = 14
        total_bundles_count = 0

        for info in lot_infos:
            d = info["data"]
            clean_rows = d.get("clean_rows", [])
            if not clean_rows:
                continue

            stone_name = info.get("stone_name", "ARSLANLAR MARBLE")
            lot_code = info.get("lot_code", info.get("query", "A-100"))
            thick_str = (
                d.get("thickness", "2 cm").upper().replace(" ", "")
            )  # "2CM"
            surf_str = d.get("surface", "POLISHED").upper()
            prod_desc = f"{stone_name} {thick_str} {surf_str} MARBLE SLABS"

            # Kalınlığa göre 1 m² ağırlık çarpanı (2cm = 57 kg/m², 3cm = 85.5 kg/m²)
            kg_per_m2 = 85.5 if "3" in thick_str else 57.0

            # Plakaları minimum 12'şerli ahşap bundle'lara böl
            bundles = split_slabs_into_bundles(clean_rows, min_bundle_size=12)

            for b_rows in bundles:
                total_bundles_count += 1
                b_start = current_row
                b_end = current_row + len(b_rows) - 1

                bundle_m2_sum = sum(r.get("m2", 0.0) for r in b_rows)
                bundle_net_kg = round(bundle_m2_sum * kg_per_m2, 2)
                bundle_gross_kg = round(bundle_net_kg + 100.0, 2)  # +100 kg ahşap bundle darası

                for s_row in b_rows:
                    ws.row_dimensions[current_row].height = 20
                    ws.cell(current_row, 4, s_row.get("slab_no", ""))
                    ws.cell(current_row, 5, s_row.get("length_cm", 0))
                    ws.cell(current_row, 6, s_row.get("height_cm", 0))
                    ws.cell(current_row, 7, 1)

                    c_m2 = ws.cell(current_row, 8, round(s_row.get("m2", 0.0), 2))
                    c_m2.number_format = "0.00"

                    c_net = ws.cell(
                        current_row, 9, f"=H{current_row}*{kg_per_m2:g}"
                    )
                    c_net.number_format = "0.00"

                    for c_i in range(2, 12):
                        cell = ws.cell(current_row, c_i)
                        cell.font = font_bold if c_i in (2, 3, 8, 11) else font_regular
                        cell.alignment = align_center
                        cell.border = border_data
                        if total_bundles_count % 2 == 0 and c_i in (4, 5, 6, 7, 8, 9):
                            cell.fill = fill_zebra

                    current_row += 1

                # Bundle Hücrelerini Dikey Birleştir (B: Ürün Adı, C: Blok No, J: Brüt Ağırlık, K: Bundle=1)
                ws.cell(b_start, 2, prod_desc)
                ws.cell(b_start, 3, lot_code)
                c_gross = ws.cell(b_start, 10, bundle_gross_kg)
                c_gross.number_format = "0.00"
                ws.cell(b_start, 11, 1)

                if b_end > b_start:
                    ws.merge_cells(
                        start_row=b_start, start_column=2, end_row=b_end, end_column=2
                    )
                    ws.merge_cells(
                        start_row=b_start, start_column=3, end_row=b_end, end_column=3
                    )
                    ws.merge_cells(
                        start_row=b_start, start_column=10, end_row=b_end, end_column=10
                    )
                    ws.merge_cells(
                        start_row=b_start, start_column=11, end_row=b_end, end_column=11
                    )

        last_data_row = max(first_data_row, current_row - 1)
        total_row = current_row

        # TOTAL Satırı (Birebir Şablon)
        ws.row_dimensions[total_row].height = 24
        ws.cell(total_row, 2, "TOTAL")
        ws.merge_cells(
            start_row=total_row, start_column=2, end_row=total_row, end_column=6
        )
        ws.cell(total_row, 7, f"=SUM(G{first_data_row}:G{last_data_row})")
        c_tot_m2 = ws.cell(
            total_row, 8, f"=SUM(H{first_data_row}:H{last_data_row})"
        )
        c_tot_m2.number_format = "0.00"
        c_tot_net = ws.cell(
            total_row, 9, f"=SUM(I{first_data_row}:I{last_data_row})"
        )
        c_tot_net.number_format = "0.00"
        c_tot_gross = ws.cell(
            total_row,
            10,
            f"=I{total_row}+{total_bundles_count * 100}",
        )
        c_tot_gross.number_format = "0.00"
        ws.cell(total_row, 11, f"=SUM(K{first_data_row}:K{last_data_row})")

        for c_i in range(2, 12):
            c = ws.cell(total_row, c_i)
            c.font = font_bold
            c.fill = fill_total
            c.alignment = align_center
            c.border = border_box

        # Alt Sevkiyat & Liman Bilgileri (Satır +2 .. +8 - Birebir Şablon)
        r_foot = total_row + 2
        total_slabs_cnt = (last_data_row - first_data_row + 1)
        footer_rows = [
            ("PORT OF LOADING:", "ALIAGA PORT OF TURKEY"),
            ("PORT OF DISCHARGE:", ""),
            (
                "AMBALAJ ŞEKLİ / PACKAGE TYPE:",
                f"{total_bundles_count} AHŞAP BANDIL / {total_bundles_count} WOODEN BUNDLES",
            ),
            ("KAP SAYISI / NUMBER OF PACKAGES:", f"{total_slabs_cnt} SLABS ({total_bundles_count} BUNDLES)"),
            ("NET AĞIRLIK / NET WEIGHT (KG):", f"=I{total_row}"),
            ("BRÜT AĞIRLIK / GROSS WEIGHT (KG):", f"=J{total_row}"),
            ("GTİP KODU / HS CODE:", "6802.91.00.00.00"),
        ]
        for idx_f, (lbl, val) in enumerate(footer_rows):
            curr_f = r_foot + idx_f
            c_lbl = ws.cell(curr_f, 2, lbl)
            c_lbl.font = font_bold
            c_val = ws.cell(curr_f, 3, val)
            c_val.font = font_bold
            c_val.alignment = align_left
            if str(val).startswith("="):
                c_val.number_format = "#,##0.00"

        out_io = io.BytesIO()
        wb.save(out_io)
        out_io.seek(0)
        return out_io.read()

    def generate_packing_list_pdf(
        self, info: Dict[str, Any]
    ) -> bytes:
        """
        Kurumsal Bundle'lı Excel şablonunun birebir PDF karşılığını (Yatay A4,
        Bronz Aslan Logolu, Bundle No, Net/Brüt Ağırlık ve Liman Detaylı) üretir.
        """
        d = info["data"]
        clean_rows = d.get("clean_rows", [])
        stone_name = info.get("stone_name", "ARSLANLAR MARBLE")
        lot_code = info.get("lot_code", info.get("query", "A-100"))
        thick_str = d.get("thickness", "2 cm")
        surf_str = d.get("surface", "Polished")
        prod_desc = f"{stone_name} ({thick_str} {surf_str})"
        kg_per_m2 = 85.5 if "3" in thick_str else 57.0

        pdf_io = io.BytesIO()
        doc = SimpleDocTemplate(
            pdf_io,
            pagesize=landscape(A4),
            rightMargin=20,
            leftMargin=20,
            topMargin=18,
            bottomMargin=18,
        )
        styles = getSampleStyleSheet()
        NAVY = colors.HexColor("#14213D")
        BRONZE = colors.HexColor("#9C5933")
        SLATE = colors.HexColor("#2B2D42")

        logo_bytes = self.drive.get_official_logo_jpeg()
        rl_logo = RLImage(io.BytesIO(logo_bytes), width=115, height=68.7)

        comp_info_style = ParagraphStyle(
            "CompInfo",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=8.5,
            leading=11.5,
            textColor=NAVY,
            alignment=2,
        )
        comp_html = (
            "<b>ARSLANLAR TAAHHUT INS. HAFR.MAD.SAN.TİC.LTD.ŞTİ</b><br/>"
            "KasımPasa Mah. Ataturk Cad. No:99/11 Menderes / Izmir / Turkey<br/>"
            "Tax Office: MENDERES VD. 0840060648<br/>"
            f"<b>DATE:</b> {datetime.now().strftime('%d.%m.%Y')} | <b>BLOCK / LOT:</b> {lot_code}"
        )
        top_header_table = Table(
            [[rl_logo, Paragraph(comp_html, comp_info_style)]],
            colWidths=[200, 580],
        )
        top_header_table.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ALIGN", (0, 0), (0, 0), "LEFT"),
                    ("ALIGN", (1, 0), (1, 0), "RIGHT"),
                ]
            )
        )

        th_p = ParagraphStyle(
            "TH",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=7.5,
            leading=9.5,
            textColor=colors.whitesmoke,
            alignment=1,
        )
        td_p = ParagraphStyle(
            "TD",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=7.5,
            leading=9.5,
            textColor=SLATE,
            alignment=1,
        )
        td_b = ParagraphStyle(
            "TDB",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=7.5,
            leading=9.5,
            textColor=NAVY,
            alignment=1,
        )

        story: List[Any] = [
            top_header_table,
            Spacer(1, 6),
            HRFlowable(width="100%", thickness=1.4, color=BRONZE, spaceAfter=8),
        ]

        table_data = [
            [
                Paragraph("PRODUCT DESCRIPTION", th_p),
                Paragraph("BLOCK NO", th_p),
                Paragraph("SLAB NO", th_p),
                Paragraph("LENGTH<br/>(cm)", th_p),
                Paragraph("WIDTH<br/>(cm)", th_p),
                Paragraph("QTY<br/>(pcs)", th_p),
                Paragraph("AREA<br/>(m²)", th_p),
                Paragraph("NET WT<br/>(kg)", th_p),
                Paragraph("BUNDLE #", th_p),
            ]
        ]

        bundles = split_slabs_into_bundles(clean_rows[:180], min_bundle_size=12)
        total_m2 = 0.0
        total_net_kg = 0.0
        total_slabs = 0

        for b_idx, b_rows in enumerate(bundles, start=1):
            for s_idx, r in enumerate(b_rows):
                m2_val = r.get("m2", 0.0)
                l_cm = r.get("length_cm", 0)
                h_cm = r.get("height_cm", 0)
                net_kg = m2_val * kg_per_m2
                total_m2 += m2_val
                total_net_kg += net_kg
                total_slabs += 1

                table_data.append(
                    [
                        Paragraph(prod_desc if s_idx == 0 else '""', td_b if s_idx == 0 else td_p),
                        Paragraph(lot_code if s_idx == 0 else '""', td_b if s_idx == 0 else td_p),
                        Paragraph(str(r.get("slab_no", total_slabs)), td_p),
                        Paragraph(f"{l_cm:g}" if l_cm > 0 else "-", td_p),
                        Paragraph(f"{h_cm:g}" if h_cm > 0 else "-", td_p),
                        Paragraph("1", td_p),
                        Paragraph(f"{m2_val:.2f}", td_b),
                        Paragraph(f"{net_kg:,.1f}", td_p),
                        Paragraph(f"Bundle #{b_idx}" if s_idx == 0 else "-", td_b if s_idx == 0 else td_p),
                    ]
                )

        if d.get("total_m2", 0.0) > 0:
            total_m2 = d["total_m2"]
            total_net_kg = total_m2 * kg_per_m2

        num_bundles = len(bundles)
        total_gross_kg = total_net_kg + (num_bundles * 100.0)

        # TOTAL Satırı
        table_data.append(
            [
                Paragraph("<b>TOTAL</b>", td_b),
                Paragraph(f"<b>{lot_code}</b>", td_b),
                Paragraph("-", td_b),
                Paragraph("-", td_b),
                Paragraph("-", td_b),
                Paragraph(f"<b>{total_slabs} pcs</b>", td_b),
                Paragraph(f"<b>{total_m2:,.2f} m²</b>", td_b),
                Paragraph(f"<b>{total_net_kg:,.1f} kg</b>", td_b),
                Paragraph(f"<b>{num_bundles} Bundles</b>", td_b),
            ]
        )

        pl_table = Table(
            table_data, colWidths=[180, 65, 55, 65, 65, 55, 75, 80, 80], repeatRows=1
        )
        pl_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                    ("BACKGROUND", (0, 1), (-1, -2), colors.HexColor("#F8F9FA")),
                    ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#F2E6D8")),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#DEE2E6")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 3),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ]
            )
        )
        story.append(pl_table)
        story.append(Spacer(1, 10))

        summary_html = (
            f"<b>PORT OF LOADING:</b> ALIAGA PORT OF TURKEY &nbsp;&nbsp;|&nbsp;&nbsp; "
            f"<b>PACKAGE TYPE:</b> {num_bundles} WOODEN BUNDLES ({total_slabs} SLABS)<br/>"
            f"<b>TOTAL NET AREA:</b> {total_m2:,.2f} m² &nbsp;&nbsp;|&nbsp;&nbsp; "
            f"<b>NET WEIGHT:</b> {total_net_kg:,.2f} kg &nbsp;&nbsp;|&nbsp;&nbsp; "
            f"<b>GROSS WEIGHT:</b> {total_gross_kg:,.2f} kg &nbsp;&nbsp;|&nbsp;&nbsp; "
            f"<b>HS CODE:</b> 6802.91.00.00.00"
        )
        story.append(
            Paragraph(
                summary_html,
                ParagraphStyle(
                    "SumFoot",
                    parent=styles["Normal"],
                    fontName="Helvetica",
                    fontSize=8.5,
                    leading=12,
                    textColor=NAVY,
                ),
            )
        )

        doc.build(story)
        pdf_io.seek(0)
        return pdf_io.read()


# =====================================================================
# 6. ARGON AI & GÖRSEL TAŞ EŞLEŞTİRME MOTORU
# =====================================================================
class ArgonBrain:
    def __init__(self, api_key: str):
        self.api_key = api_key

    def match_visual_stone(self, image_bytes: bytes) -> List[str]:
        if self.api_key:
            try:
                b64_img = base64.b64encode(image_bytes).decode("utf-8")
                url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={self.api_key}"
                prompt = (
                    "Bu mermer fotoğrafının zemin rengini ve damar yapısını analiz et. "
                    "Şu seleksiyonlarımızdan en çok benzeyen 2 tanesini seç ve sadece JSON listesi olarak döndür: "
                    '["A-283", "SILVER ROOTS", "SILVER RIVER", "ROSSO LEVANTO", "TERRA MATER", "BLUE ROOTS", "DEKART"]'
                )
                payload = {
                    "contents": [
                        {
                            "parts": [
                                {"text": prompt},
                                {
                                    "inline_data": {
                                        "mime_type": "image/jpeg",
                                        "data": b64_img,
                                    }
                                },
                            ]
                        }
                    ]
                }
                resp = requests.post(url, json=payload, timeout=10)
                if resp.status_code == 200:
                    txt = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
                    found = [
                        k
                        for k in DEFAULT_STONE_DNA.keys()
                        if k in normalize_turkish_key(txt)
                    ]
                    if found:
                        return found[:2]
            except Exception:
                pass
        return ["SILVER ROOTS", "SILVER RIVER"]

    def evaluate_intent(
        self,
        user_text: str,
        is_group: bool,
        bot_mentioned: bool,
        has_photo: bool = False,
    ) -> Dict[str, Any]:
        raw = normalize_turkish_key(user_text)
        if has_photo:
            return {"should_reply": True, "is_visual_search": True}

        if not raw:
            return {"should_reply": False}

        lot_regex = re.compile(r"\b([A-Z]{1,4})[-\s]?(\d{1,5})\b")
        matched = [
            f"{m.group(1)}-{int(m.group(2))}"
            for m in lot_regex.finditer(raw)
            if m.group(1) not in ["M", "CM", "MM", "PDF", "TON", "KG"]
        ]

        for st_key in DEFAULT_STONE_DNA.keys():
            if st_key in raw and st_key not in matched:
                matched.append(st_key)

        deep_scan_keywords = [
            "TUM STOK",
            "TUM STOGUM",
            "GENEL TARAMA",
            "DEEP SCAN",
            "RAPORLA",
            "NE KADAR STOGUM",
            "TOPLAM STOK",
            "TOPLAM NE KADAR",
            "TOPLAM STOGUM",
        ]
        want_deep_scan = any(kw in raw for kw in deep_scan_keywords)

        collection_folders: List[str] = []
        if re.search(r"ARSLANLAR\s+COLLECTION(?!\s*EXCLUSIVE)", raw) or (
            "COLLECTION" in raw and "EXCLUSIVE" not in raw
        ):
            collection_folders.append("Arslanlar Collection")
        if "SERIES" in raw or "SERISI" in raw:
            collection_folders.append("Arslanlar Series")
        if "EXCLUSIVE" in raw:
            collection_folders.append("Arslanlar Exclusive Collection")

        want_pdf = any(kw in raw for kw in ["PDF", "SUNUM", "KATALOG"])
        want_packing = any(kw in raw for kw in ["CEKI", "PACKING", "SEVK"])
        want_only_link = any(kw in raw for kw in ["LINK", "DRIVE"]) and not any(
            kw in raw for kw in ["FOTO", "RESIM", "GORSEL", "PDF", "SUNUM", "CEKI"]
        )
        want_photo = (
            any(kw in raw for kw in ["FOTO", "RESIM", "GORSEL"])
            or (
                bool(matched)
                and not want_only_link
                and not want_pdf
                and not want_packing
            )
        )

        if (
            is_group
            and not bot_mentioned
            and not matched
            and not want_pdf
            and not want_packing
            and not want_photo
            and not want_only_link
            and not collection_folders
            and not want_deep_scan
        ):
            return {"should_reply": False}

        return {
            "should_reply": True,
            "is_visual_search": False,
            "matched_queries": list(dict.fromkeys(matched))[:10],
            "collection_folders": collection_folders,
            "want_deep_scan": want_deep_scan,
            "want_pdf": want_pdf,
            "want_packing": want_packing,
            "want_only_link": want_only_link,
            "want_photo": want_photo,
        }


# =====================================================================
# 7. TELEGRAM SOHBET YÖNETİCİSİ
# =====================================================================
drive_explorer = DriveExplorer(GOOGLE_SA_JSON)
inventory_engine = InventoryEngine(drive_explorer)
doc_engine = CorporateDocumentEngine(inventory_engine, drive_explorer)
argon_brain = ArgonBrain(GEMINI_API_KEY)

PROCESSED_MESSAGES = deque(maxlen=500)


async def telegram_error_handler(
    update: object, context: ContextTypes.DEFAULT_TYPE
) -> None:
    err = context.error
    if isinstance(err, (Conflict, NetworkError, TimedOut)):
        return
    logger.error(f"Telegram hatası: {err}")


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.effective_message
    if not msg or not update.effective_chat:
        return

    msg_key = f"{update.effective_chat.id}:{msg.message_id}"
    if msg_key in PROCESSED_MESSAGES:
        return
    PROCESSED_MESSAGES.append(msg_key)

    is_group = update.effective_chat.type in ["group", "supergroup"]
    user_text = (msg.text or msg.caption or "").strip()
    bot_mentioned = bool(
        context.bot.username
        and f"@{context.bot.username.lower()}" in user_text.lower()
    )

    has_photo = bool(msg.photo)
    intent = argon_brain.evaluate_intent(
        user_text, is_group, bot_mentioned, has_photo=has_photo
    )
    if not intent.get("should_reply"):
        return

    # 0. FOTOĞRAF ATILDIĞINDA BENZER TAŞ BULMA
    if intent.get("is_visual_search"):
        status = await msg.reply_text(
            "🔍 <b>Gönderdiğiniz mermer fotoğrafı analiz ediliyor...</b>",
            parse_mode=ParseMode.HTML,
        )
        try:
            photo_file = await msg.photo[-1].get_file()
            img_bytes = bytes(await photo_file.download_as_bytearray())
            matched_stones = await asyncio.to_thread(
                argon_brain.match_visual_stone, img_bytes
            )
            await status.delete()
            intent["matched_queries"] = matched_stones
            intent["want_photo"] = True
        except Exception as e:
            await status.edit_text(f"⚠️ Görsel analiz edilemedi: {e}")
            return

    queries = intent.get("matched_queries", [])
    collection_folders: List[str] = intent.get("collection_folders", [])

    # 1. DEEP SCAN (TOPLAM STOK RAPORU)
    if intent.get("want_deep_scan") and not queries and not collection_folders:
        status = await msg.reply_text(
            "🔎 <b>Argon Deep Scan Raporu Hazırlanıyor...</b>",
            parse_mode=ParseMode.HTML,
        )
        try:
            report = await asyncio.to_thread(
                inventory_engine.generate_deep_scan_report
            )
            await status.edit_text(report, parse_mode=ParseMode.HTML)
        except Exception as e:
            await status.edit_text(f"⚠️ Rapor uyarısı: {e}")
        return

    if not queries and not collection_folders:
        if not is_group or bot_mentioned:
            await msg.reply_text(
                "👋 <b>Argon Master V12 Corporate Aktif.</b>\n"
                "• Lot & Fotoğraf: <code>A-371</code> veya <code>A-1009</code>\n"
                "• Drive Linki: <code>A-371 drive linki lütfen</code>\n"
                "• Çoklu Tek PDF Sunum: <code>A-96 ve A-383 sunum hazırla</code>\n"
                "• Koleksiyon PDF: <code>Arslanlar Collection sunum lütfen</code>\n"
                "• 3'lü Çeki Listesi (Orijinal + Draft Bundle Excel + PDF): <code>A-371 çeki listesi</code>\n"
                "• Genel Stok: <code>Toplam stoğum ne kadar</code>",
                parse_mode=ParseMode.HTML,
            )
        return

    # 2. KOLEKSİYON PDF SUNUMU (Tekilleştirilmiş, Alt Klasör Filtreli)
    if collection_folders and intent.get("want_pdf") and not intent.get("want_packing"):
        coll_label = ", ".join(collection_folders)
        status_msg = await msg.reply_text(
            f"📁 <b>{coll_label}</b> için Tekilleştirilmiş ve Yüksek Çözünürlüklü PDF Kataloğu hazırlanıyor...",
            parse_mode=ParseMode.HTML,
        )
        try:
            pdf_b, pdf_n = await asyncio.to_thread(
                doc_engine.generate_presentation_pdf,
                [],
                collection_folder_names=collection_folders,
            )
            await msg.reply_document(
                document=io.BytesIO(pdf_b),
                filename=pdf_n,
                caption=f"📄 <b>{coll_label}</b> — Executive Stock & Slab Presentation",
                parse_mode=ParseMode.HTML,
            )
            await status_msg.delete()
        except Exception as e:
            await status_msg.edit_text(f"⚠️ Koleksiyon PDF uyarısı: {e}")
        return

    # 3. ÇOKLU VEYA TEKLİ LOT PDF SUNUMU
    if queries and intent.get("want_pdf") and not intent.get("want_packing"):
        status_msg = await msg.reply_text(
            f"📄 <b>{', '.join(queries)}</b> için Bronz Aslan Logolu Premium PDF Sunumu hazırlanıyor...",
            parse_mode=ParseMode.HTML,
        )
        try:
            pdf_b, pdf_n = await asyncio.to_thread(
                doc_engine.generate_presentation_pdf, queries
            )
            await msg.reply_document(
                document=io.BytesIO(pdf_b),
                filename=pdf_n,
                caption=f"📄 <b>{', '.join(queries)}</b> — Executive Stock & Slab Presentation",
                parse_mode=ParseMode.HTML,
            )
            await status_msg.delete()
        except Exception as e:
            await status_msg.edit_text(f"⚠️ PDF oluşturulamadı: {e}")
        return

    # 4. ÇEKİ LİSTESİ (1. DRIVE ORİJİNAL EXCEL + 2. BUNDLE DRAFT EXCEL + 3. KURUMSAL PDF)
    if queries and intent.get("want_packing"):
        status_msg = await msg.reply_text(
            f"📦 <b>{', '.join(queries)}</b> için 12'şerli Bundle Formatında Kurumsal Çeki Listeleri (Orijinal + Draft Excel + PDF) hazırlanıyor...",
            parse_mode=ParseMode.HTML,
        )
        collected_infos = []
        for q in queries:
            info = await asyncio.to_thread(inventory_engine.get_lot_info, q)
            collected_infos.append(info)
            # 1) Drive'daki Orijinal Çeki Listesi Excel'i (varsa)
            if info.get("excel_bytes"):
                await msg.reply_document(
                    document=io.BytesIO(info["excel_bytes"]),
                    filename=info["excel_name"],
                    caption=f"📁 <b>{info['title']}</b> — Drive Orijinal Çeki Listesi",
                    parse_mode=ParseMode.HTML,
                )

        # 2) Kullanıcının gönderdiği PACKING LIST DRAFT formatında (Minimum 12'şerli Bundle) Excel
        valid_infos = [
            inf for inf in collected_infos if inf["data"].get("clean_rows")
        ]
        if valid_infos:
            draft_excel_bytes = await asyncio.to_thread(
                doc_engine.generate_draft_format_packing_excel,
                valid_infos,
                customer_code="-".join(queries[:3]),
            )
            date_str = datetime.now().strftime("%d.%m.%Y")
            draft_fname = f"PACKING LIST DRAFT {date_str} {'-'.join(queries[:3])}.xlsx"
            await msg.reply_document(
                document=io.BytesIO(draft_excel_bytes),
                filename=draft_fname,
                caption=f"📊 <b>{', '.join(queries)}</b> — Kurumsal Bundle Packing List (Min. 12'li Bundle Formatı)",
                parse_mode=ParseMode.HTML,
            )

            # 3) Kurumsal Yatay PDF Çeki Listesi
            for inf in valid_infos:
                pdf_pack_bytes = await asyncio.to_thread(
                    doc_engine.generate_packing_list_pdf, inf
                )
                d_inf = inf["data"]
                await msg.reply_document(
                    document=io.BytesIO(pdf_pack_bytes),
                    filename=f"OFFICIAL_PACKING_LIST_{inf['lot_code']}.pdf",
                    caption=(
                        f"📄 <b>{inf['title']}</b> — Resmi PDF Çeki Listesi "
                        f"(<code>{d_inf['total_m2']:,.2f} m²</code> | <code>{d_inf['total_slabs']} Plaka</code>)"
                    ),
                    parse_mode=ParseMode.HTML,
                )
        await status_msg.delete()
        return

    # 5. STANDART LOT SORGUSU (FLASH-KART + 2 YÜKSEK ÇÖZÜNÜRLÜKLÜ FOTOĞRAF)
    for q in queries:
        try:
            info = await asyncio.to_thread(inventory_engine.get_lot_info, q)
            d = info["data"]
            fid = info["folder_id"]
            flink = info["drive_link"]

            m2_val = d.get("total_m2", 0.0)
            slabs_val = d.get("total_slabs", 0)
            bundles_cnt = max(1, slabs_val // 12) if slabs_val >= 12 else (1 if slabs_val > 0 else 0)

            stock_str = (
                f"{m2_val:,.2f} m² ({slabs_val} Plaka / {bundles_cnt} Bundle)"
                if m2_val > 0
                else "Stok Kontrolü İçin Klasörü Açın"
            )
            link_html = (
                f"<a href='{flink}'>Klasöre Git ({info['lot_code']})</a>"
                if flink
                else "Drive'da Klasör Bulunamadı"
            )

            text_response = (
                f"💎 <b>{info['title']}</b>\n"
                f"• Mevcut Stok: <code>{stock_str}</code>\n"
                f"• Ortalama Ebat: <code>{d.get('dimensions', 'Standard Slab Sizes')}</code>\n"
                f"• Yüzey & Kalınlık: <code>{d.get('surface', 'Polished')} ({d.get('thickness', '2 cm')})</code>\n"
                f"• Drive Linki: {link_html}"
            )

            if intent.get("want_only_link"):
                await msg.reply_text(text_response, parse_mode=ParseMode.HTML)
                continue

            img1, img2 = await asyncio.to_thread(
                drive_explorer.get_images_for_telegram, fid, q
            )

            media = []
            if img1:
                media.append(
                    InputMediaPhoto(
                        media=img1, caption=text_response, parse_mode=ParseMode.HTML
                    )
                )
            if img2:
                media.append(InputMediaPhoto(media=img2))

            if media:
                try:
                    await msg.reply_media_group(media=media)
                except Exception:
                    await msg.reply_text(text_response, parse_mode=ParseMode.HTML)
            else:
                await msg.reply_text(text_response, parse_mode=ParseMode.HTML)
        except Exception as ex:
            logger.error(f"Sorgu hatası ({q}): {ex}")
            await msg.reply_text(
                f"⚠️ <b>{q}</b> sorgulanırken gecikme yaşandı, lütfen tekrar deneyin.",
                parse_mode=ParseMode.HTML,
            )


# =====================================================================
# 8. ANA BAŞLATICI
# =====================================================================
def main():
    if not TELEGRAM_TOKEN:
        logger.error("CRITICAL: TELEGRAM_TOKEN eksik!")
        return

    application = (
        ApplicationBuilder()
        .token(TELEGRAM_TOKEN)
        .concurrent_updates(True)
        .build()
    )

    application.add_error_handler(telegram_error_handler)
    application.add_handler(
        MessageHandler(
            (filters.TEXT | filters.PHOTO) & (~filters.COMMAND), handle_message
        )
    )

    logger.info("🚀 Argon Master V12 Corporate başlatılıyor...")
    application.run_polling(
        drop_pending_updates=True, poll_interval=1.0, timeout=15
    )


if __name__ == "__main__":
    main()
