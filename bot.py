import os
import io
import re
import gc
import sys
import json
import time
import base64
import logging
import asyncio
import threading
from datetime import datetime
from collections import deque
from typing import Dict, List, Optional, Tuple, Any

# =====================================================================
# 0. RENDER PORTUNU ANINDA AÇ (DEPLOY TAKILMASINI %100 ÖNLER)
# =====================================================================
from flask import Flask
import requests

flask_app = Flask(__name__)

@flask_app.route("/")
def health_check():
    return "Argon Master V10 Pre-Indexed Zero-Wait is Live!", 200

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

import httplib2
from google.oauth2 import service_account
from google_auth_httplib2 import AuthorizedHttp
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("ArgonMaster_V10")

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GOOGLE_SA_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
DEFAULT_DRIVE_LINK = (
    os.getenv("GOOGLE_DRIVE_LINK")
    or os.getenv("DEFAULT_DRIVE_LINK")
    or "https://drive.google.com"
).strip()

MASTER_INDEX_FILE = "drive_master_index_v10.json"

# GÖMÜLÜ ORİJİNAL ARSLANLAR MARBLE BRONZ ASLAN LOGOSU (logo_bronz)
EMBEDDED_LOGO_B64 = (
    "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"
)

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
# 2. YARDIMCI FONKSİYONLAR VE NORMALİZASYON
# =====================================================================
def load_json_file(filepath: str, default_data: Any) -> Any:
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return default_data


def save_json_file(filepath: str, data: Any) -> None:
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def normalize_turkish_key(text: Any) -> str:
    if text is None or pd.isna(text):
        return ""
    tr_map = str.maketrans("İıiŞşĞğÜüÖöÇç", "IIISSGGUUOOCC")
    s = str(text).strip().translate(tr_map).upper()
    return re.sub(r"\s+", " ", s).strip()


def normalize_lot_code(code: Any) -> str:
    raw = normalize_turkish_key(code)
    match = re.match(r"^([A-Z]+)[-\s_.]*(\d+)$", raw)
    return f"{match.group(1)}-{match.group(2)}" if match else raw


def safe_float(val: Any) -> float:
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return 0.0
    s_clean = re.sub(r"[^\d,.-]", "", s)
    if not s_clean:
        return 0.0
    if "," in s_clean and "." in s_clean:
        if s_clean.rfind(",") > s_clean.rfind("."):
            s_clean = s_clean.replace(".", "").replace(",", ".")
        else:
            s_clean = s_clean.replace(",", "")
    elif "," in s_clean:
        s_clean = s_clean.replace(",", ".")
    try:
        return float(s_clean)
    except Exception:
        m = re.search(r"[-+]?\d+(?:\.\d+)?", s_clean)
        return float(m.group(0)) if m else 0.0


def compute_slab_m2(
    l_val: float, h_val: float, raw_m2: float
) -> Tuple[float, float, float]:
    if l_val > 600:
        l_cm = l_val / 10.0
    elif 0 < l_val < 10:
        l_cm = l_val * 100.0
    else:
        l_cm = l_val

    if h_val > 600:
        h_cm = h_val / 10.0
    elif 0 < h_val < 10:
        h_cm = h_val * 100.0
    else:
        h_cm = h_val

    calc_m2 = (l_cm * h_cm) / 10000.0 if (l_cm > 20 and h_cm > 20) else 0.0

    if raw_m2 > 1000:
        raw_m2 = raw_m2 / 10000.0

    if 0.2 <= raw_m2 <= 25.0:
        return raw_m2, l_cm, h_cm
    if 0.2 <= calc_m2 <= 25.0:
        return calc_m2, l_cm, h_cm
    return 0.0, l_cm, h_cm


def prepare_telegram_image_bytes(raw_bytes: Optional[bytes]) -> Optional[bytes]:
    if not raw_bytes:
        return None
    try:
        with Image.open(io.BytesIO(raw_bytes)) as im:
            img = im.convert("RGB")
            img.thumbnail((1500, 1500), Image.Resampling.LANCZOS)
            out = io.BytesIO()
            img.save(out, format="JPEG", quality=85)
            return out.getvalue()
    except Exception:
        return None


def prepare_pdf_card_image_bytes(raw_bytes: Optional[bytes]) -> Optional[bytes]:
    if not raw_bytes:
        return None
    try:
        with Image.open(io.BytesIO(raw_bytes)) as im:
            img = im.convert("RGB")
            img.thumbnail((540, 360), Image.Resampling.LANCZOS)
            out = io.BytesIO()
            img.save(out, format="JPEG", quality=85)
            return out.getvalue()
    except Exception:
        return None


# =====================================================================
# 3. HIZLI THUMBNAIL & BATCH DESTEKLİ DRIVE MOTORU
# =====================================================================
class DriveExplorer:
    def __init__(self, sa_json_str: str):
        self.sa_json_str = sa_json_str
        self._sa_info = None
        self._local = threading.local()
        self._logo_jpeg_bytes: Optional[bytes] = None

        if self.sa_json_str:
            try:
                self._sa_info = json.loads(self.sa_json_str)
                logger.info("Google Drive Kimlik Bilgileri Hazır.")
            except Exception as e:
                logger.error(f"Google Service Account JSON Hatası: {e}")

    def _get_service(self):
        if not self._sa_info:
            return None
        if getattr(self._local, "service", None) is None:
            creds = service_account.Credentials.from_service_account_info(
                self._sa_info,
                scopes=["https://www.googleapis.com/auth/drive.readonly"],
            )
            http = httplib2.Http(timeout=12)
            authed_http = AuthorizedHttp(creds, http=http)
            self._local.service = build(
                "drive", "v3", http=authed_http, cache_discovery=False
            )
        return self._local.service

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

    @staticmethod
    def _matches_target(folder_name: str, norm_target: str) -> int:
        fn_norm = normalize_turkish_key(folder_name)
        fn_lot = normalize_lot_code(folder_name)
        if fn_lot == norm_target or fn_norm == norm_target:
            return 100

        lot_m = re.match(r"^([A-Z]+)-(\d+)$", norm_target)
        if lot_m:
            prefix, num = lot_m.group(1), lot_m.group(2)
            pat = rf"(?<![A-Z0-9]){prefix}[-\s_.]*0*{num}(?![0-9])"
            if re.search(pat, fn_norm):
                return 85
            return 0

        if norm_target in fn_norm:
            return 75
        return 0

    def find_folder_live(
        self, query_term: str
    ) -> Tuple[Optional[str], Optional[str], Optional[str]]:
        norm_key = normalize_lot_code(query_term)
        svc = self._get_service()
        if not svc or not norm_key:
            return None, None, None

        raw_clean = query_term.strip().replace("'", "")
        search_terms = [norm_key, raw_clean]

        lot_m = re.match(r"^([A-Z]+)-(\d+)$", norm_key)
        if lot_m:
            prefix, num = lot_m.group(1), lot_m.group(2)
            search_terms.extend([f"{prefix} {num}", f"{prefix}{num}", num])
        else:
            tr_var = norm_key.replace("SILVER", "SİLVER").replace("RIVER", "RİVER")
            if tr_var != norm_key:
                search_terms.append(tr_var)

        search_terms = list(dict.fromkeys([t for t in search_terms if len(t) >= 2]))

        matched_folders = []
        seen_ids = set()

        for term in search_terms:
            try:
                q = (
                    f"mimeType = 'application/vnd.google-apps.folder' "
                    f"and trashed = false and name contains '{term}'"
                )
                res = (
                    svc.files()
                    .list(
                        q=q,
                        spaces="drive",
                        fields="files(id, name, webViewLink, parents)",
                        includeItemsFromAllDrives=True,
                        supportsAllDrives=True,
                        pageSize=25,
                    )
                    .execute()
                )
                for f in res.get("files", []):
                    if f["id"] in seen_ids:
                        continue
                    score = self._matches_target(f.get("name", ""), norm_key)
                    if score > 0:
                        seen_ids.add(f["id"])
                        matched_folders.append((score, f))
                if any(s >= 90 for s, _ in matched_folders):
                    break
            except Exception:
                continue

        if not matched_folders:
            return None, None, None

        matched_folders.sort(key=lambda x: x[0], reverse=True)
        best = matched_folders[0][1]
        fid = best["id"]
        link = best.get(
            "webViewLink", f"https://drive.google.com/drive/folders/{fid}"
        )
        parents = best.get("parents", [])
        return fid, link, (parents[0] if parents else None)

    def batch_list_children(
        self, parent_ids: List[str]
    ) -> Dict[str, List[Dict[str, Any]]]:
        result_map: Dict[str, List[Dict[str, Any]]] = {
            pid: [] for pid in parent_ids if pid
        }
        svc = self._get_service()
        if not svc or not result_map:
            return result_map

        valid_pids = list(result_map.keys())
        chunk_size = 15

        for i in range(0, len(valid_pids), chunk_size):
            chunk = valid_pids[i : i + chunk_size]
            or_clause = " or ".join([f"'{pid}' in parents" for pid in chunk])
            q = f"({or_clause}) and trashed = false"
            try:
                res = (
                    svc.files()
                    .list(
                        q=q,
                        spaces="drive",
                        fields="files(id, name, mimeType, parents, webViewLink, thumbnailLink, shortcutDetails)",
                        includeItemsFromAllDrives=True,
                        supportsAllDrives=True,
                        pageSize=800,
                    )
                    .execute()
                )
                for f in res.get("files", []):
                    for p in f.get("parents", []):
                        if p in result_map:
                            result_map[p].append(f)
            except Exception as e:
                logger.warning(f"Batch listeleme uyarısı: {e}")

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

    def download_excel_or_sheet(
        self, file_id: str, mime_type: str
    ) -> Optional[bytes]:
        svc = self._get_service()
        if not svc or not file_id:
            return None
        try:
            if mime_type == "application/vnd.google-apps.spreadsheet":
                req = svc.files().export_media(
                    fileId=file_id,
                    mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                req = svc.files().get_media(
                    fileId=file_id, supportsAllDrives=True
                )
            fh = io.BytesIO()
            dl = MediaIoBaseDownload(fh, req)
            done = False
            while not done:
                _, done = dl.next_chunk()
            fh.seek(0)
            return fh.read()
        except Exception as e:
            logger.warning(f"Tablo indirme uyarısı ({file_id}): {e}")
            return None

    def download_fast_image(
        self,
        file_id: Optional[str],
        thumb_url: Optional[str] = None,
        for_pdf: bool = False,
    ) -> Optional[bytes]:
        """
        Önce Google Drive'ın hızlı yüksek çözünürlüklü önizleme sunucusunu (=s1000)
        dener (15 MB yerine 80 KB indirir, 10 kat hızlıdır ve CPU'yu hiç yormaz!).
        Eğer önizleme yoksa doğrudan dosyayı indirir.
        """
        raw = None
        if thumb_url:
            try:
                hi_res_url = re.sub(r"=s\d+", "=s1200", thumb_url)
                resp = requests.get(hi_res_url, timeout=6)
                if resp.status_code == 200 and len(resp.content) > 1500:
                    raw = resp.content
            except Exception:
                pass

        if not raw and file_id:
            svc = self._get_service()
            if svc:
                try:
                    req = svc.files().get_media(
                        fileId=file_id, supportsAllDrives=True
                    )
                    fh = io.BytesIO()
                    dl = MediaIoBaseDownload(fh, req)
                    done = False
                    while not done:
                        _, done = dl.next_chunk()
                    fh.seek(0)
                    raw = fh.read()
                except Exception:
                    pass

        if not raw:
            return None
        return (
            prepare_pdf_card_image_bytes(raw)
            if for_pdf
            else prepare_telegram_image_bytes(raw)
        )


# =====================================================================
# 4. ARKA PLAN ÖNBELLEKLEYİCİ (PRE-INDEXED) ENVANTER MOTORU
# =====================================================================
class InventoryEngine:
    """
    Bot başlarken arka planda (CPU'yu hiç yormadan) tüm koleksiyonları ve lotları
    indeksler. Kullanıcı soru sorduğunda veya PDF istediğinde Drive'da arama yapmak
    yerine 0.1 saniyede hazır veriyi kullanır!
    """
    def __init__(self, drive: DriveExplorer):
        self.drive = drive
        self.index_db: Dict[str, Any] = load_json_file(
            MASTER_INDEX_FILE,
            {"lots": {}, "collections": {}, "last_updated": None},
        )
        self._lock = threading.Lock()
        self.is_indexing = False

    @staticmethod
    def _detect_columns(col_map: Dict[Any, str]) -> Dict[str, Any]:
        m2_col = next(
            (
                c
                for c, n in col_map.items()
                if any(
                    k in n
                    for k in [
                        "M2",
                        "M²",
                        "ALAN",
                        "NET",
                        "METRAJ",
                        "METREKARE",
                        "AREA",
                        "SQM",
                    ]
                )
            ),
            None,
        )

        len_col = next(
            (
                c
                for c, n in col_map.items()
                if c != m2_col
                and any(k in n for k in ["UZUNLUK", "BOY", "LENGTH", "LONG"])
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
                        for k in ["YUKSEKLIK", "GENISLIK", "WIDTH", "HEIGHT"]
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
                and any(k in n for k in ["PLAKA", "SLAB", "SIRA", "NO"])
            ),
            None,
        )

        adet_col = next(
            (
                c
                for c, n in col_map.items()
                if c not in (m2_col, len_col, height_col, slab_col)
                and any(k in n for k in ["ADET", "PCS", "QTY", "MIKTAR"])
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
                if any(
                    k in n for k in ["LOT", "BLOK", "BLOCK", "BUNDLE", "PARTI", "KOD"]
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
        }

    def parse_packing_list_bytes(
        self, excel_bytes: bytes, target_lot: str = ""
    ) -> Optional[Dict[str, Any]]:
        norm_target_lot = normalize_lot_code(target_lot) if target_lot else ""
        try:
            xls = pd.ExcelFile(io.BytesIO(excel_bytes))
            ordered_sheets = list(xls.sheet_names)
            if norm_target_lot:
                ordered_sheets.sort(
                    key=lambda s: 0
                    if DriveExplorer._matches_target(s, norm_target_lot) > 0
                    else 1
                )

            for sheet in ordered_sheets:
                raw_df = pd.read_excel(xls, sheet_name=sheet, header=None)
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
                    "METREKARE",
                    "BOY",
                    "UZUNLUK",
                    "LENGTH",
                    "YUKSEKLIK",
                    "GENISLIK",
                    "WIDTH",
                    "HEIGHT",
                    "PLAKA",
                    "SLAB",
                    "ADET",
                    "KALINLIK",
                    "THICKNESS",
                    "EBAT",
                    "OLCU",
                ]
                for idx in range(min(35, len(raw_df))):
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

                start_h = max(0, header_idx - 1) if best_score >= 1 else 0
                end_h = min(len(raw_df) - 1, header_idx + 1) if best_score >= 1 else 0

                composite_headers = {}
                for col_i in range(raw_df.shape[1]):
                    parts = []
                    for r_i in range(start_h, end_h + 1):
                        val = normalize_turkish_key(raw_df.iloc[r_i, col_i])
                        if val and not re.match(r"^\d+([.,]\d+)?$", val):
                            if val not in parts:
                                parts.append(val)
                    composite_headers[col_i] = (
                        " ".join(parts) if parts else f"COL_{col_i}"
                    )

                data_start_row = header_idx + 1 if best_score >= 1 else 0
                df = raw_df.iloc[data_start_row:].copy()
                df.columns = list(range(raw_df.shape[1]))

                mapped = self._detect_columns(composite_headers)

                if (
                    mapped["len"] is None
                    or mapped["height"] is None
                    or mapped["m2"] is None
                ):
                    dim_candidates = []
                    m2_candidates = []
                    for col_i in df.columns:
                        nums = [
                            safe_float(v)
                            for v in df[col_i].dropna().values
                            if 0 < safe_float(v) < 5000
                        ]
                        if len(nums) >= 2:
                            avg_n = sum(nums) / len(nums)
                            if (70 <= avg_n <= 380) or (700 <= avg_n <= 3800):
                                dim_candidates.append((avg_n, col_i))
                            elif 1.2 <= avg_n <= 12.0 and any(
                                (n % 1) != 0 for n in nums
                            ):
                                m2_candidates.append(col_i)

                    if len(dim_candidates) >= 2 and (
                        mapped["len"] is None or mapped["height"] is None
                    ):
                        dim_candidates.sort(key=lambda x: x[0], reverse=True)
                        if mapped["len"] is None:
                            mapped["len"] = dim_candidates[0][1]
                        if mapped["height"] is None:
                            for _, c_idx in dim_candidates[1:]:
                                if c_idx != mapped["len"]:
                                    mapped["height"] = c_idx
                                    break
                    if mapped["m2"] is None and m2_candidates:
                        mapped["m2"] = m2_candidates[-1]

                m2_col = mapped["m2"]
                len_col = mapped["len"]
                height_col = mapped["height"]
                dim_col = mapped["dim"]
                slab_col = mapped["slab"]
                adet_col = mapped["adet"]
                thick_col = mapped["thick"]
                surf_col = mapped["surf"]
                lot_col = mapped["lot"]

                rows_to_iter = df
                if norm_target_lot and lot_col is not None:
                    lot_filtered = df[
                        df[lot_col].apply(
                            lambda v: DriveExplorer._matches_target(
                                str(v), norm_target_lot
                            )
                            > 0
                        )
                    ]
                    if not lot_filtered.empty:
                        rows_to_iter = lot_filtered

                valid_rows = []
                total_m2 = 0.0
                total_slabs = 0
                lengths = []
                heights = []

                for _, row in rows_to_iter.iterrows():
                    row_str = " ".join(
                        [normalize_turkish_key(v) for v in row.values if pd.notna(v)]
                    )
                    if any(
                        tot_w in row_str
                        for tot_w in ["TOPLAM", "TOTAL", "GENEL TOPLAM", "SUBTOTAL"]
                    ):
                        continue

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

                    raw_m2 = (
                        safe_float(row.get(m2_col)) if m2_col is not None else 0.0
                    )

                    m2_val, l_cm, h_cm = compute_slab_m2(raw_l, raw_h, raw_m2)
                    if m2_val <= 0:
                        continue

                    adet_val = 1
                    if adet_col is not None:
                        cand_adet = int(safe_float(row.get(adet_col)))
                        if 1 <= cand_adet <= 5:
                            adet_val = cand_adet

                    total_m2 += m2_val
                    total_slabs += adet_val
                    if l_cm > 20:
                        lengths.append(l_cm)
                    if h_cm > 20:
                        heights.append(h_cm)

                    valid_rows.append(
                        {
                            "slab_no": str(
                                row.get(slab_col, len(valid_rows) + 1)
                                if slab_col is not None
                                else len(valid_rows) + 1
                            ),
                            "length_cm": l_cm,
                            "height_cm": h_cm,
                            "thickness": "2 cm",
                            "m2": m2_val,
                        }
                    )

                if total_m2 <= 0 and m2_col is not None:
                    col_vals = [
                        safe_float(v)
                        for v in rows_to_iter[m2_col].dropna().values
                        if 0.3 <= safe_float(v) <= 500.0
                    ]
                    if col_vals:
                        slab_like = [v for v in col_vals if 0.4 <= v <= 15.0]
                        if len(slab_like) >= 2:
                            total_m2 = sum(slab_like)
                            total_slabs = len(slab_like)
                        elif max(col_vals) > 2.0:
                            total_m2 = max(col_vals)
                            total_slabs = max(1, len(col_vals) - 1)

                if total_m2 <= 0:
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
                    dims_str = f"{avg_l} cm (Avg. Length)"
                else:
                    dims_str = "Standard Slab Sizes"

                return {
                    "total_m2": round(total_m2, 2),
                    "total_slabs": total_slabs,
                    "thickness": thick_val,
                    "surface": surf_val,
                    "dimensions": dims_str,
                    "clean_rows": valid_rows,
                }
        except Exception as e:
            logger.warning(f"Excel parse uyarısı: {e}")
        return None

    def _estimate_stock_from_images_list(
        self, all_imgs: List[Dict[str, Any]]
    ) -> Optional[Dict[str, Any]]:
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
        slab_count = len(target_imgs)
        if slab_count <= 0:
            return None

        parsed_m2_list = []
        parsed_l_list = []
        parsed_h_list = []
        clean_rows = []

        for idx, f in enumerate(target_imgs, start=1):
            fname = f.get("name", "")
            dim_m = re.search(r"(\d{2,4})\s*[xX*]\s*(\d{2,4})", fname)
            area_m = re.search(
                r"(\d+[.,]\d+)\s*(?:m2|m²|sqm)", fname, re.IGNORECASE
            )
            l_v = safe_float(dim_m.group(1)) if dim_m else 0.0
            h_v = safe_float(dim_m.group(2)) if dim_m else 0.0
            a_v = safe_float(area_m.group(1)) if area_m else 0.0
            m2_v, l_cm, h_cm = compute_slab_m2(l_v, h_v, a_v)
            if m2_v <= 0:
                m2_v = 5.40
                l_cm = 290.0
                h_cm = 186.0
            parsed_m2_list.append(m2_v)
            if l_cm > 20 and h_cm > 20:
                parsed_l_list.append(l_cm)
                parsed_h_list.append(h_cm)
            clean_rows.append(
                {
                    "slab_no": str(idx),
                    "length_cm": l_cm,
                    "height_cm": h_cm,
                    "thickness": "2 cm",
                    "m2": m2_v,
                }
            )

        total_m2 = round(sum(parsed_m2_list), 2)
        if parsed_l_list and parsed_h_list:
            avg_l = int(round(sum(parsed_l_list) / len(parsed_l_list)))
            avg_h = int(round(sum(parsed_h_list) / len(parsed_h_list)))
            dims_str = f"{avg_l} x {avg_h} cm"
        else:
            dims_str = "290 x 186 cm"

        return {
            "total_m2": total_m2,
            "total_slabs": slab_count,
            "thickness": "2 cm",
            "surface": "Polished",
            "dimensions": dims_str,
            "clean_rows": clean_rows,
        }

    def _build_entry_from_folder_items(
        self,
        folder_id: str,
        display_name: str,
        raw_name: str,
        folder_link: str,
        items: List[Dict[str, Any]],
        sub_batch_map: Dict[str, List[Dict[str, Any]]],
    ) -> Dict[str, Any]:
        excel_cand = None
        imgs = []

        for it in items:
            mime = it.get("mimeType", "")
            if not excel_cand:
                excel_cand = DriveExplorer._extract_excel_candidate(it)
            if mime.startswith("image/"):
                imgs.append(it)
            elif mime == "application/vnd.google-apps.folder":
                sf_items = sub_batch_map.get(it["id"], [])
                for s_it in sf_items:
                    if not excel_cand:
                        excel_cand = DriveExplorer._extract_excel_candidate(s_it)
                    if s_it.get("mimeType", "").startswith("image/"):
                        s_copy = dict(s_it)
                        s_copy["parent_name"] = it.get("name", "")
                        imgs.append(s_copy)

        parsed_data = None
        if excel_cand:
            ex_bytes = self.drive.download_excel_or_sheet(
                excel_cand["id"], excel_cand["mimeType"]
            )
            if ex_bytes:
                parsed_data = self.parse_packing_list_bytes(
                    ex_bytes, target_lot=raw_name
                )
            time.sleep(0.08)  # CPU ve GIL nefes alma payı

        if not parsed_data or parsed_data.get("total_m2", 0.0) <= 0:
            parsed_data = self._estimate_stock_from_images_list(imgs)

        if not parsed_data:
            parsed_data = {
                "total_m2": 0.0,
                "total_slabs": 0,
                "thickness": "2 cm",
                "surface": "Polished",
                "dimensions": "Standard Slab Sizes",
                "clean_rows": [],
            }

        # Standart plaka ve Bookmatch görsel ID'lerini ayır
        bm_img = None
        std_imgs = []
        for f in imgs:
            fn = f.get("name", "").lower()
            pn = f.get("parent_name", "").lower()
            if "bookmatch" in fn or "bm" in fn or "bookmatch" in pn:
                if not bm_img:
                    bm_img = f
            else:
                std_imgs.append(f)

        slab_1 = std_imgs[0] if std_imgs else bm_img
        slab_2 = (
            bm_img
            if (bm_img and bm_img != slab_1)
            else (std_imgs[1] if len(std_imgs) > 1 else slab_1)
        )

        return {
            "title": display_name.upper(),
            "raw_name": raw_name,
            "folder_id": folder_id,
            "drive_link": folder_link,
            "excel_meta": excel_cand,
            "slab_img": {
                "id": slab_1["id"],
                "thumb": slab_1.get("thumbnailLink"),
            }
            if slab_1
            else None,
            "second_img": {
                "id": slab_2["id"],
                "thumb": slab_2.get("thumbnailLink"),
            }
            if slab_2
            else None,
            "data": parsed_data,
        }

    def sync_all_collections_background(self) -> None:
        """
        Arka planda sessizce çalışır; 3 ana koleksiyonun tüm taş ve lotlarını
        indeksleyip RAM'e ve JSON'a kaydeder.
        """
        if self.is_indexing:
            return
        self.is_indexing = True
        logger.info("🔄 Arka plan Google Drive İndeksleme (Pre-Index) başladı...")

        collections = [
            "Arslanlar Collection",
            "Arslanlar Series",
            "Arslanlar Exclusive Collection",
        ]
        new_lots: Dict[str, Any] = dict(self.index_db.get("lots", {}))
        new_colls: Dict[str, List[str]] = dict(
            self.index_db.get("collections", {})
        )

        try:
            for coll_name in collections:
                fid, _, _ = self.drive.find_folder_live(coll_name)
                if not fid:
                    continue

                lvl1_items = self.drive.batch_list_children([fid]).get(fid, [])
                lvl1_folders = [
                    f
                    for f in lvl1_items
                    if f.get("mimeType") == "application/vnd.google-apps.folder"
                ]
                if not lvl1_folders:
                    continue

                lvl1_map = {f["id"]: f for f in lvl1_folders}
                lvl2_batch = self.drive.batch_list_children(list(lvl1_map.keys()))

                leaf_list = []
                lvl2_subfolders = {}

                for f1_id, children in lvl2_batch.items():
                    f1_meta = lvl1_map[f1_id]
                    sub_f = [
                        c
                        for c in children
                        if c.get("mimeType") == "application/vnd.google-apps.folder"
                        and "bookmatch" not in c.get("name", "").lower()
                        and "bm" not in c.get("name", "").lower()
                    ]
                    if sub_f:
                        for sf in sub_f:
                            sf["stone_parent_name"] = f1_meta.get("name", "")
                            lvl2_subfolders[sf["id"]] = sf
                    else:
                        leaf_list.append(
                            {
                                "id": f1_id,
                                "display_name": f1_meta.get("name", ""),
                                "raw_name": f1_meta.get("name", ""),
                                "link": f1_meta.get(
                                    "webViewLink",
                                    f"https://drive.google.com/drive/folders/{f1_id}",
                                ),
                                "items": children,
                            }
                        )

                if lvl2_subfolders:
                    lvl3_batch = self.drive.batch_list_children(
                        list(lvl2_subfolders.keys())
                    )
                    for f2_id, children in lvl3_batch.items():
                        f2_meta = lvl2_subfolders[f2_id]
                        p_stone = f2_meta.get("stone_parent_name", "")
                        f2_name = f2_meta.get("name", "")
                        disp = (
                            f"{p_stone} - {f2_name}"
                            if p_stone
                            and normalize_turkish_key(p_stone)
                            not in normalize_turkish_key(f2_name)
                            else f2_name
                        )
                        leaf_list.append(
                            {
                                "id": f2_id,
                                "display_name": disp,
                                "raw_name": f2_name,
                                "link": f2_meta.get(
                                    "webViewLink",
                                    f"https://drive.google.com/drive/folders/{f2_id}",
                                ),
                                "items": children,
                            }
                        )

                # Alt klasörlerin içindeki Bookmatch klasörlerini de tek seferde çek
                bookmatch_sub_ids = []
                for lf in leaf_list:
                    for it in lf["items"]:
                        if it.get("mimeType") == "application/vnd.google-apps.folder":
                            bookmatch_sub_ids.append(it["id"])
                bm_batch = (
                    self.drive.batch_list_children(bookmatch_sub_ids)
                    if bookmatch_sub_ids
                    else {}
                )

                coll_keys = []
                for lf in leaf_list:
                    entry = self._build_entry_from_folder_items(
                        folder_id=lf["id"],
                        display_name=lf["display_name"],
                        raw_name=lf["raw_name"],
                        folder_link=lf["link"],
                        items=lf["items"],
                        sub_batch_map=bm_batch,
                    )
                    norm_k = normalize_lot_code(lf["raw_name"])
                    # Lot kodunu da anahtar olarak kaydet (Örn: "A-383 SILVER ROOTS" -> "A-383")
                    lot_m = re.search(r"\b([A-Z]{1,4})[-\s]?(\d{1,5})\b", norm_k)
                    primary_key = (
                        f"{lot_m.group(1)}-{lot_m.group(2)}" if lot_m else norm_k
                    )
                    new_lots[primary_key] = entry
                    new_lots[norm_k] = entry
                    if primary_key not in coll_keys:
                        coll_keys.append(primary_key)
                    time.sleep(0.05)

                new_colls[coll_name] = coll_keys
                with self._lock:
                    self.index_db["lots"] = new_lots
                    self.index_db["collections"] = new_colls
                    self.index_db["last_updated"] = datetime.now().strftime(
                        "%d.%m.%Y %H:%M"
                    )
                    save_json_file(MASTER_INDEX_FILE, self.index_db)

            logger.info(
                f"✅ Drive İndeksleme Tamamlandı! Toplam İndekslenen Kayıt: {len(new_lots)}"
            )
        except Exception as e:
            logger.error(f"Arka plan indeksleme uyarısı: {e}")
        finally:
            self.is_indexing = False
            gc.collect()

    def get_lot_entry(self, query: str) -> Dict[str, Any]:
        """
        Önce RAM indeksine bakar (0.01 sn!). İndekste yoksa canlı olarak Drive'dan
        çekip indekse ekler.
        """
        norm_q = normalize_lot_code(query)
        lots_map = self.index_db.get("lots", {})

        if norm_q in lots_map:
            return lots_map[norm_q]

        for k, v in lots_map.items():
            if norm_q in k or k in norm_q:
                return v

        # İndekste henüz yoksa canlı bul ve kaydet
        fid, flink, _ = self.drive.find_folder_live(norm_q)
        if fid:
            items = self.drive.batch_list_children([fid]).get(fid, [])
            sub_ids = [
                it["id"]
                for it in items
                if it.get("mimeType") == "application/vnd.google-apps.folder"
            ]
            sub_batch = (
                self.drive.batch_list_children(sub_ids) if sub_ids else {}
            )
            entry = self._build_entry_from_folder_items(
                folder_id=fid,
                display_name=norm_q,
                raw_name=norm_q,
                folder_link=flink or DEFAULT_DRIVE_LINK,
                items=items,
                sub_batch_map=sub_batch,
            )
            with self._lock:
                self.index_db.setdefault("lots", {})[norm_q] = entry
                save_json_file(MASTER_INDEX_FILE, self.index_db)
            return entry

        return {
            "title": norm_q,
            "raw_name": norm_q,
            "folder_id": None,
            "drive_link": None,
            "excel_meta": None,
            "slab_img": None,
            "second_img": None,
            "data": {
                "total_m2": 0.0,
                "total_slabs": 0,
                "thickness": "2 cm",
                "surface": "Polished",
                "dimensions": "Standard Slab Sizes",
                "clean_rows": [],
            },
        }

    def get_collection_entries(self, coll_name: str) -> List[Dict[str, Any]]:
        colls_map = self.index_db.get("collections", {})
        lots_map = self.index_db.get("lots", {})
        keys = colls_map.get(coll_name, [])
        if keys:
            return [lots_map[k] for k in keys if k in lots_map]

        # Henüz indekslenmemişse hızlıca sadece 1. katmanı çek
        fid, _, _ = self.drive.find_folder_live(coll_name)
        if not fid:
            return []
        lvl1_items = self.drive.batch_list_children([fid]).get(fid, [])
        lvl1_folders = [
            f
            for f in lvl1_items
            if f.get("mimeType") == "application/vnd.google-apps.folder"
        ][:12]
        f_ids = [f["id"] for f in lvl1_folders]
        lvl2_batch = self.drive.batch_list_children(f_ids)

        entries = []
        for f in lvl1_folders:
            entry = self._build_entry_from_folder_items(
                folder_id=f["id"],
                display_name=f["name"],
                raw_name=f["name"],
                folder_link=f.get(
                    "webViewLink",
                    f"https://drive.google.com/drive/folders/{f['id']}",
                ),
                items=lvl2_batch.get(f["id"], []),
                sub_batch_map={},
            )
            entries.append(entry)
        return entries

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
            entries = self.get_collection_entries(coll)
            if not entries:
                continue

            report_lines.append(f"📁 <b>{coll.upper()}</b>")
            coll_m2 = 0.0
            coll_slabs = 0

            for ent in entries:
                d = ent.get("data", {})
                m2 = d.get("total_m2", 0.0)
                slabs = d.get("total_slabs", 0)
                if m2 > 0:
                    report_lines.append(
                        f"  ├ <b>{ent['title']}</b>: <code>{m2:,.2f} m²</code> ({slabs} Plaka)"
                    )
                    coll_m2 += m2
                    coll_slabs += slabs
                    total_factory_m2 += m2
                    total_factory_slabs += slabs

            report_lines.append(
                f"  └ <i>Koleksiyon Toplamı: {coll_m2:,.2f} m² ({coll_slabs} Plaka)</i>\n"
            )

        if total_factory_m2 == 0 and self.is_indexing:
            report_lines.append(
                "⏳ <i>Arka plan Drive indekslemesi devam ediyor, 30 saniye sonra tüm rakamlar tam listelenecektir.</i>\n"
            )

        report_lines.append(
            f"🌐 <b>GENEL FABRİKA TOPLAMI: {total_factory_m2:,.2f} m² | {total_factory_slabs} Plaka</b>"
        )
        return "\n".join(report_lines)


# =====================================================================
# 5. GÖMÜLÜ ORİJİNAL LOGOLU & STOK DETAYLI PREMIUM PDF MOTORU
# =====================================================================
class PDFPresentationEngine:
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
        entries_to_render: List[Dict[str, Any]] = []

        if collection_folder_names:
            titles_norm = [
                normalize_turkish_key(c) for c in collection_folder_names
            ]
            pdf_title = f"{' & '.join(titles_norm)} — EXECUTIVE CATALOG PRESENTATION"
            for coll_name in collection_folder_names:
                coll_entries = self.inventory.get_collection_entries(coll_name)
                entries_to_render.extend(coll_entries[:12])
        else:
            pdf_title = "EXECUTIVE STOCK & SLAB PRESENTATION"
            for q in queries:
                entries_to_render.append(self.inventory.get_lot_entry(q))

        if not entries_to_render:
            raise ValueError("Belirtilen kriterler için klasör veya veri bulunamadı.")

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

        for idx, ent in enumerate(entries_to_render[:24], start=1):
            d = ent.get("data", {})
            m2 = d.get("total_m2", 0.0)
            slabs = d.get("total_slabs", 0)

            if m2 > 0 or slabs > 0:
                stock_str = f"<b>{m2:,.2f} m²  ({slabs} Slabs)</b>"
                heading_badge = (
                    f" — <font color='#9C5933'>{m2:,.2f} m² | {slabs} Slabs</font>"
                )
            else:
                stock_str = "<b>0.00 m² (0 Slabs)</b>"
                heading_badge = ""

            dims_str = d.get("dimensions", "Standard Slab Sizes")
            if not dims_str or dims_str == "-":
                dims_str = "Standard Slab Sizes"

            surf_str = d.get("surface", "Polished")
            if not surf_str or surf_str == "-":
                surf_str = "Polished"

            thick_str = d.get("thickness", "2 cm")
            if not thick_str or thick_str == "-":
                thick_str = "2 cm"

            surf_thick_str = f"{surf_str} | {thick_str}"

            # Hızlı Thumbnail (=s1200) üzerinden sadece düz plaka görselini indir
            rl_img = Paragraph("High-Res Slab Image in Drive", tdv_s)
            slab_meta = ent.get("slab_img")
            if slab_meta:
                img_bytes = self.drive.download_fast_image(
                    slab_meta.get("id"),
                    slab_meta.get("thumb"),
                    for_pdf=True,
                )
                if img_bytes:
                    try:
                        rl_img = RLImage(
                            io.BytesIO(img_bytes), width=185, height=118
                        )
                    except Exception:
                        pass

            dlink = ent.get("drive_link")
            link_html = (
                f"<link href='{dlink}' color='#0056b3'><u>Click to Open Google Drive Folder</u></link>"
                if dlink
                else "Available on Request"
            )

            spec_rows = [
                [Paragraph("Specification", th_s), Paragraph("Technical Data", th_s)],
                [
                    Paragraph("Available Stock (m² / Slabs)", tdl_s),
                    Paragraph(stock_str, tdv_s),
                ],
                [Paragraph("Avg. Slab Size", tdl_s), Paragraph(dims_str, tdv_s)],
                [
                    Paragraph("Surface & Thickness", tdl_s),
                    Paragraph(surf_thick_str, tdv_s),
                ],
                [Paragraph("Active Code", tdl_s), Paragraph(ent["title"], tdv_s)],
                [
                    Paragraph("High-Res Slab Photos", tdl_s),
                    Paragraph(link_html, tdv_s),
                ],
            ]

            st = Table(spec_rows, colWidths=[135, 195])
            st.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                        ("BACKGROUND", (0, 1), (-1, -1), LIGHT_BG),
                        ("GRID", (0, 0), (-1, -1), 0.5, BORD),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("TOPPADDING", (0, 0), (-1, -1), 5),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
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
                        Paragraph(f"{idx}. {ent['title']}{heading_badge}", ch_s),
                        mr,
                        Spacer(1, 13),
                    ]
                )
            )

        story.append(HRFlowable(width="100%", thickness=0.8, color=BRONZE, spaceAfter=6))
        story.append(
            Paragraph(
                "<b>ARSLANLAR MARBLE</b> | International Sales & Export Department | Izmir, Turkey",
                footer_s,
            )
        )

        doc.build(story)
        pdf_io.seek(0)
        pdf_bytes = pdf_io.read()
        gc.collect()

        if collection_folder_names:
            fname = f"Arslanlar_Collections_Presentation_{datetime.now().strftime('%H%M')}.pdf"
        else:
            fname = f"Arslanlar_Presentation_{'_'.join(queries[:3])}.pdf"
        return pdf_bytes, fname

    def generate_packing_list_pdf(
        self, clean_rows: List[Dict[str, Any]], lot_name: str
    ) -> bytes:
        pdf_io = io.BytesIO()
        doc = SimpleDocTemplate(
            pdf_io,
            pagesize=landscape(A4),
            rightMargin=25,
            leftMargin=25,
            topMargin=22,
            bottomMargin=22,
        )
        styles = getSampleStyleSheet()
        NAVY = colors.HexColor("#14213D")
        BRONZE = colors.HexColor("#9C5933")
        SLATE = colors.HexColor("#2B2D42")

        logo_bytes = self.drive.get_official_logo_jpeg()
        rl_logo = RLImage(io.BytesIO(logo_bytes), width=120, height=71.7)
        rl_logo.hAlign = "CENTER"

        sub_p = ParagraphStyle(
            "PlS",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=10.5,
            leading=14,
            textColor=BRONZE,
            alignment=1,
            spaceBefore=4,
            spaceAfter=10,
        )
        th_p = ParagraphStyle(
            "TH",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=9,
            textColor=colors.whitesmoke,
            alignment=1,
        )
        td_p = ParagraphStyle(
            "TD",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=8.5,
            textColor=SLATE,
            alignment=1,
        )

        story = [
            rl_logo,
            Paragraph(f"OFFICIAL PACKING LIST — {lot_name.upper()}", sub_p),
            HRFlowable(width="100%", thickness=1.2, color=BRONZE, spaceAfter=12),
        ]
        table_data = [
            [
                Paragraph("Row #", th_p),
                Paragraph("Slab No", th_p),
                Paragraph("Length (cm)", th_p),
                Paragraph("Height (cm)", th_p),
                Paragraph("Thickness", th_p),
                Paragraph("Net Area (m²)", th_p),
            ]
        ]

        total_m2 = 0.0
        for idx, r in enumerate(clean_rows[:180], start=1):
            m2_val = r.get("m2", 0.0)
            l_cm = r.get("length_cm", 0.0)
            h_cm = r.get("height_cm", 0.0)
            total_m2 += m2_val
            table_data.append(
                [
                    Paragraph(str(idx), td_p),
                    Paragraph(str(r.get("slab_no", idx)), td_p),
                    Paragraph(f"{l_cm:g}" if l_cm > 0 else "-", td_p),
                    Paragraph(f"{h_cm:g}" if h_cm > 0 else "-", td_p),
                    Paragraph(str(r.get("thickness", "2 cm")), td_p),
                    Paragraph(f"{m2_val:.2f}", td_p),
                ]
            )

        pl_table = Table(table_data, colWidths=[65, 110, 110, 110, 90, 110])
        pl_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                    ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F8F9FA")),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#DEE2E6")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ]
            )
        )
        story.append(pl_table)
        story.append(Spacer(1, 14))
        story.append(
            Paragraph(
                f"<b>Total Slabs:</b> {len(clean_rows)} pcs | <b>Total Net Area:</b> {total_m2:,.2f} m²",
                ParagraphStyle(
                    "Sum",
                    parent=styles["Normal"],
                    fontName="Helvetica-Bold",
                    fontSize=10,
                    textColor=NAVY,
                    alignment=1,
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

    def match_visual_stone(self, image_bytes: bytes, caption: str = "") -> List[str]:
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

        try:
            with Image.open(io.BytesIO(image_bytes)) as im:
                rgb_small = im.convert("RGB").resize((64, 64))
                pixels = list(rgb_small.getdata())
            r_avg = sum(p[0] for p in pixels) / len(pixels)
            g_avg = sum(p[1] for p in pixels) / len(pixels)
            b_avg = sum(p[2] for p in pixels) / len(pixels)
            if r_avg > g_avg + 25 and r_avg > b_avg + 25:
                return ["ROSSO LEVANTO", "A-283"]
            if b_avg > r_avg + 10:
                return ["BLUE ROOTS", "SILVER ROOTS"]
            if r_avg > 165 and g_avg > 155:
                return ["A-283", "TERRA MATER"]
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
            f"{m.group(1)}-{m.group(2)}"
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
# 7. EŞZAMANLI (NON-BLOCKING) TELEGRAM SOHBET YÖNETİCİSİ
# =====================================================================
drive_explorer = DriveExplorer(GOOGLE_SA_JSON)
inventory_engine = InventoryEngine(drive_explorer)
pdf_engine = PDFPresentationEngine(inventory_engine, drive_explorer)
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

    chat_id = update.effective_chat.id
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

    # 0. FOTOĞRAF ATILDIĞINDA BENZER TAŞ BULMA ÖZELLİĞİ
    if intent.get("is_visual_search"):
        status = await msg.reply_text(
            "🔍 <b>Gönderdiğiniz mermer fotoğrafı analiz ediliyor...</b>",
            parse_mode=ParseMode.HTML,
        )
        try:
            photo_file = await msg.photo[-1].get_file()
            img_bytes = bytes(await photo_file.download_as_bytearray())
            matched_stones = await asyncio.to_thread(
                argon_brain.match_visual_stone, img_bytes, user_text
            )
            await status.delete()
            intent["matched_queries"] = matched_stones
            intent["want_photo"] = True
        except Exception as e:
            await status.edit_text(f"⚠️ Görsel analiz edilemedi: {e}")
            return

    queries = intent.get("matched_queries", [])
    collection_folders: List[str] = intent.get("collection_folders", [])

    # 1. DEEP SCAN (TOPLAM STOK RAPORU - RAM İNDEKSİNDEN ANINDA DÖNER)
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
                "👋 <b>Argon Master V10 (Pre-Indexed Zero-Wait) Aktif.</b>\n"
                "• Lot & Fotoğraf: <code>A-371</code> veya <code>A-371 fotoğraf</code>\n"
                "• Drive Linki: <code>A-371 drive linki lütfen</code>\n"
                "• Çoklu Tek PDF Sunum: <code>A-96 ve A-383 sunum hazırla</code>\n"
                "• Koleksiyon PDF: <code>Arslanlar Collection sunum lütfen</code>\n"
                "• Çeki Listesi (Excel + PDF): <code>A-371 çeki listesi</code>\n"
                "• Genel Stok: <code>Toplam stoğum ne kadar</code>",
                parse_mode=ParseMode.HTML,
            )
        return

    # 2. KOLEKSİYON (TEKLİ VEYA 3 KOLEKSİYON BİRDEN) PDF SUNUMU
    if collection_folders and intent.get("want_pdf"):
        coll_label = ", ".join(collection_folders)
        status_msg = await msg.reply_text(
            f"📁 <b>{coll_label}</b> için Bronz Aslan Logolu ve Stok Detaylı Premium PDF Kataloğu hazırlanıyor...",
            parse_mode=ParseMode.HTML,
        )
        try:
            pdf_b, pdf_n = await asyncio.to_thread(
                pdf_engine.generate_presentation_pdf,
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

    # 3. ÇOKLU VEYA TEKLİ LOT PDF SUNUMU (TEK PDF'TE BİRLEŞTİRİR)
    if queries and intent.get("want_pdf"):
        status_msg = await msg.reply_text(
            f"📄 <b>{', '.join(queries)}</b> için Bronz Aslan Logolu ve Stok Detaylı Premium PDF Sunumu hazırlanıyor...",
            parse_mode=ParseMode.HTML,
        )
        try:
            pdf_b, pdf_n = await asyncio.to_thread(
                pdf_engine.generate_presentation_pdf, queries
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

    # 4. STANDART LOT SORGUSU, LİNK VE ÇEKİ LİSTESİ (1 EXCEL + 1 PDF)
    for q in queries:
        try:
            ent = await asyncio.to_thread(inventory_engine.get_lot_entry, q)
            d = ent.get("data", {})
            flink = ent.get("drive_link")
            ex_meta = ent.get("excel_meta")

            if intent.get("want_packing"):
                ex_bytes = None
                ex_name = f"PackingList_{q}.xlsx"
                if ex_meta:
                    ex_bytes = await asyncio.to_thread(
                        drive_explorer.download_excel_or_sheet,
                        ex_meta.get("id"),
                        ex_meta.get("mimeType", ""),
                    )
                    ex_name = ex_meta.get("name", ex_name)
                    if not ex_name.lower().endswith((".xlsx", ".xls")):
                        ex_name += ".xlsx"

                if ex_bytes:
                    await msg.reply_document(
                        document=io.BytesIO(ex_bytes),
                        filename=ex_name,
                        caption=f"📦 <b>{q}</b> Orijinal Excel Çeki Listesi",
                        parse_mode=ParseMode.HTML,
                    )
                elif d.get("clean_rows"):
                    df_gen = pd.DataFrame(d["clean_rows"])
                    ex_io = io.BytesIO()
                    df_gen.to_excel(ex_io, index=False)
                    ex_io.seek(0)
                    await msg.reply_document(
                        document=ex_io,
                        filename=ex_name,
                        caption=f"📦 <b>{q}</b> Excel Çeki Listesi",
                        parse_mode=ParseMode.HTML,
                    )

                if d.get("clean_rows"):
                    pdf_pack_bytes = await asyncio.to_thread(
                        pdf_engine.generate_packing_list_pdf,
                        d["clean_rows"],
                        q,
                    )
                    await msg.reply_document(
                        document=io.BytesIO(pdf_pack_bytes),
                        filename=f"PackingList_{q}.pdf",
                        caption=f"📄 <b>{q}</b> Resmi PDF Çeki Listesi",
                        parse_mode=ParseMode.HTML,
                    )
                elif not ex_bytes:
                    link_note = (
                        f"\n📂 <a href='{flink}'>Drive Klasörünü Aç ({q})</a>"
                        if flink
                        else ""
                    )
                    await msg.reply_text(
                        f"⚠️ <b>{q}</b> klasöründe henüz çeki listesi bulunamadı.{link_note}",
                        parse_mode=ParseMode.HTML,
                    )
                continue

            m2_val = d.get("total_m2", 0.0)
            slabs_val = d.get("total_slabs", 0)
            stock_str = (
                f"{m2_val:,.2f} m² ({slabs_val} Plaka)"
                if (m2_val > 0 or slabs_val > 0)
                else "0.00 m² (0 Plaka)"
            )
            link_html = (
                f"<a href='{flink}'>Klasöre Git ({q})</a>"
                if flink
                else "Drive'da Klasör Bulunamadı"
            )

            text_response = (
                f"💎 <b>{ent.get('title', q.upper())}</b>\n"
                f"• Mevcut Stok: <code>{stock_str}</code>\n"
                f"• Ortalama Ebat: <code>{d.get('dimensions', 'Standard Slab Sizes')}</code>\n"
                f"• Yüzey & Kalınlık: <code>{d.get('surface', 'Polished')} ({d.get('thickness', '2 cm')})</code>\n"
                f"• Drive Linki: {link_html}"
            )

            if intent.get("want_only_link"):
                await msg.reply_text(text_response, parse_mode=ParseMode.HTML)
                continue

            slab_meta = ent.get("slab_img")
            sec_meta = ent.get("second_img")

            img1 = (
                await asyncio.to_thread(
                    drive_explorer.download_fast_image,
                    slab_meta.get("id"),
                    slab_meta.get("thumb"),
                    False,
                )
                if slab_meta
                else None
            )
            img2 = (
                await asyncio.to_thread(
                    drive_explorer.download_fast_image,
                    sec_meta.get("id"),
                    sec_meta.get("thumb"),
                    False,
                )
                if sec_meta
                else None
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
def _start_background_indexer():
    time.sleep(3)
    inventory_engine.sync_all_collections_background()

def main():
    if not TELEGRAM_TOKEN:
        logger.error("CRITICAL: TELEGRAM_TOKEN eksik!")
        return

    threading.Thread(target=_start_background_indexer, daemon=True).start()

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

    logger.info("🚀 Argon Master V10 (Pre-Indexed Zero-Wait) başlatılıyor...")
    application.run_polling(
        drop_pending_updates=True, poll_interval=1.0, timeout=15
    )


if __name__ == "__main__":
    main()
