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

# =====================================================================
# 0. RENDER PORTUNU 0. SANİYEDE AÇ & UYKU MODUNU ÖNLE (KEEP-ALIVE)
# =====================================================================
from flask import Flask
import requests

flask_app = Flask(__name__)

@flask_app.route("/")
def health_check():
    return "Argon Master V8 Official Logo & Exact Stock is Live!", 200

@flask_app.route("/healthz")
def healthz():
    return "OK", 200

def _start_flask_immediately():
    port = int(os.environ.get("PORT", 10000))
    flask_app.run(host="0.0.0.0", port=port, use_reloader=False)

def _keep_render_awake():
    render_url = os.environ.get("RENDER_EXTERNAL_URL", "").strip()
    while True:
        time.sleep(600)
        if render_url:
            try:
                requests.get(render_url, timeout=10)
            except Exception:
                pass

threading.Thread(target=_start_flask_immediately, daemon=True).start()
threading.Thread(target=_keep_render_awake, daemon=True).start()

# =====================================================================
# 1. KÜTÜPHANELER VE SİSTEM AYARLARI
# =====================================================================
import pandas as pd
import openpyxl
from PIL import Image, ImageDraw

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.graphics.shapes import Drawing, Rect, String, Line, Polygon, Path
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

import httplib2
from google.oauth2 import service_account
from google_auth_httplib2 import AuthorizedHttp
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("ArgonMaster_V8")

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GOOGLE_SA_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
DEFAULT_DRIVE_LINK = (
    os.getenv("GOOGLE_DRIVE_LINK")
    or os.getenv("DEFAULT_DRIVE_LINK")
    or "https://drive.google.com"
).strip()

CACHE_DB_FILE = "argon_v8_folder_cache.json"
CUSTOM_LOGO_PATH = "company_logo.png"

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
    """
    Boy ve En değerlerini cm'ye normalize eder ve M2 değerini doğrular/hesaplar.
    """
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
        img = Image.open(io.BytesIO(raw_bytes)).convert("RGB")
        img.thumbnail((1800, 1800), Image.Resampling.LANCZOS)
        out = io.BytesIO()
        img.save(out, format="JPEG", quality=86)
        return out.getvalue()
    except Exception as e:
        logger.warning(f"Görsel optimize edilemedi: {e}")
        return None


# =====================================================================
# 3. THREAD-SAFE & AKILLI GOOGLE DRIVE MOTORU
# =====================================================================
class DriveExplorer:
    def __init__(self, sa_json_str: str):
        self.sa_json_str = sa_json_str
        self._sa_info = None
        self._local = threading.local()
        self.folder_cache: Dict[str, Dict[str, str]] = load_json_file(
            CACHE_DB_FILE, {}
        )
        self._db_lock = threading.Lock()
        self._cached_logo_bytes: Optional[bytes] = None

        if self.sa_json_str:
            try:
                self._sa_info = json.loads(self.sa_json_str)
                logger.info("Google Drive Service Account JSON Yüklendi.")
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
            http = httplib2.Http(timeout=15)
            authed_http = AuthorizedHttp(creds, http=http)
            self._local.service = build(
                "drive", "v3", http=authed_http, cache_discovery=False
            )
        return self._local.service

    def get_company_logo_bytes(self) -> Optional[bytes]:
        """
        1. Yerel dizinde 'company_logo.png' veya 'logo.jpg' varsa onu yükler.
        2. Yoksa Google Drive'da 'logo_bronz' veya 'arslanlar logo' dosyasını arayıp indirir.
        """
        if self._cached_logo_bytes:
            return self._cached_logo_bytes

        for local_name in [
            CUSTOM_LOGO_PATH,
            "logo.png",
            "logo.jpg",
            "logo.jpeg",
            "logo_bronz.jpg",
            "logo_bronz.png",
        ]:
            if os.path.exists(local_name):
                try:
                    with open(local_name, "rb") as f:
                        self._cached_logo_bytes = f.read()
                        return self._cached_logo_bytes
                except Exception:
                    pass

        svc = self._get_service()
        if not svc:
            return None

        try:
            q = (
                "trashed = false and mimeType contains 'image/' and "
                "(name contains 'logo_bronz' or name contains 'LOGO' or name contains 'logo')"
            )
            res = (
                svc.files()
                .list(
                    q=q,
                    spaces="drive",
                    fields="files(id, name)",
                    includeItemsFromAllDrives=True,
                    supportsAllDrives=True,
                    pageSize=5,
                )
                .execute()
            )
            files = res.get("files", [])
            if files:
                raw = self._download_file(files[0]["id"])
                if raw:
                    self._cached_logo_bytes = raw
                    try:
                        with open(CUSTOM_LOGO_PATH, "wb") as f:
                            f.write(raw)
                    except Exception:
                        pass
                    return raw
        except Exception as e:
            logger.debug(f"Drive logo arama: {e}")

        return None

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

    def find_folder(self, query_term: str) -> Tuple[Optional[str], Optional[str]]:
        norm_key = normalize_lot_code(query_term)
        if not norm_key:
            return None, None

        if norm_key in self.folder_cache:
            cached = self.folder_cache[norm_key]
            if cached.get("id") and cached.get("link"):
                return cached["id"], cached["link"]

        svc = self._get_service()
        if not svc:
            return None, None

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
                        fields="files(id, name, webViewLink)",
                        includeItemsFromAllDrives=True,
                        supportsAllDrives=True,
                        pageSize=30,
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
            except Exception as e:
                logger.warning(f"Klasör arama uyarısı ({term}): {e}")
                continue

        if not matched_folders:
            return None, None

        # En yüksek puanlı klasörleri sırala; birden fazla varsa içinde Excel olanı tercih et
        matched_folders.sort(key=lambda x: x[0], reverse=True)
        best_candidate = matched_folders[0][1]

        if len(matched_folders) > 1:
            for _, cand_f in matched_folders[:4]:
                ex_b, _ = self.get_excel_from_folder(cand_f["id"], max_depth=1)
                if ex_b:
                    best_candidate = cand_f
                    break

        fid = best_candidate["id"]
        link = best_candidate.get(
            "webViewLink", f"https://drive.google.com/drive/folders/{fid}"
        )
        with self._db_lock:
            self.folder_cache[norm_key] = {
                "id": fid,
                "link": link,
                "name": best_candidate.get("name", norm_key),
            }
            save_json_file(CACHE_DB_FILE, self.folder_cache)
        return fid, link

    def list_child_folders(self, parent_folder_id: str) -> List[Dict[str, str]]:
        svc = self._get_service()
        if not svc or not parent_folder_id:
            return []
        try:
            q = (
                f"'{parent_folder_id}' in parents and trashed = false "
                f"and mimeType = 'application/vnd.google-apps.folder'"
            )
            res = (
                svc.files()
                .list(
                    q=q,
                    spaces="drive",
                    fields="files(id, name, webViewLink)",
                    includeItemsFromAllDrives=True,
                    supportsAllDrives=True,
                    pageSize=100,
                )
                .execute()
            )
            return [
                {
                    "id": f["id"],
                    "name": f["name"],
                    "link": f.get(
                        "webViewLink",
                        f"https://drive.google.com/drive/folders/{f['id']}",
                    ),
                }
                for f in res.get("files", [])
            ]
        except Exception:
            return []

    def get_excel_from_folder(
        self, folder_id: Optional[str], query_hint: str = "", max_depth: int = 2
    ) -> Tuple[Optional[bytes], Optional[str]]:
        svc = self._get_service()
        if not svc:
            return None, None

        try:
            if folder_id:
                q = f"'{folder_id}' in parents and trashed = false"
                res = (
                    svc.files()
                    .list(
                        q=q,
                        spaces="drive",
                        fields="files(id, name, mimeType, shortcutDetails)",
                        includeItemsFromAllDrives=True,
                        supportsAllDrives=True,
                        pageSize=100,
                    )
                    .execute()
                )
                all_items = res.get("files", [])

                excel_candidates = []
                sub_folders = []

                for item in all_items:
                    name = item.get("name", "")
                    if name.startswith("~$"):
                        continue
                    mime = item.get("mimeType", "")
                    low_name = name.lower()

                    if mime == "application/vnd.google-apps.folder":
                        sub_folders.append(item)
                        continue

                    target_id = item["id"]
                    target_mime = mime
                    if mime == "application/vnd.google-apps.shortcut":
                        sc = item.get("shortcutDetails", {})
                        target_id = sc.get("targetId", target_id)
                        target_mime = sc.get("targetMimeType", mime)

                    is_excel = (
                        low_name.endswith((".xlsx", ".xls", ".xlsm", ".csv"))
                        or "spreadsheet" in target_mime
                        or "excel" in target_mime
                        or target_mime == "application/vnd.google-apps.spreadsheet"
                    )
                    if is_excel:
                        excel_candidates.append(
                            {"id": target_id, "name": name, "mimeType": target_mime}
                        )

                if excel_candidates:
                    target = excel_candidates[0]
                    b_data = self._download_excel_or_sheet(
                        target["id"], target["mimeType"]
                    )
                    fname = target["name"]
                    if not fname.lower().endswith((".xlsx", ".xls", ".csv")):
                        fname += ".xlsx"
                    if b_data:
                        return b_data, fname

                if max_depth > 0:
                    for sf in sub_folders[:12]:
                        b_data, fname = self.get_excel_from_folder(
                            sf["id"], query_hint="", max_depth=max_depth - 1
                        )
                        if b_data:
                            return b_data, fname

            # Klasörde bulunamadıysa doğrudan lot koduyla tüm Drive'da Excel/Sheet ara
            if query_hint and max_depth >= 1:
                norm_q = normalize_lot_code(query_hint)
                cands = [norm_q, norm_q.replace("-", " "), norm_q.replace("-", "")]
                for cand in cands:
                    gq = (
                        f"trashed = false and name contains '{cand}' and "
                        f"mimeType != 'application/vnd.google-apps.folder' and "
                        f"not mimeType contains 'image/'"
                    )
                    gres = (
                        svc.files()
                        .list(
                            q=gq,
                            spaces="drive",
                            fields="files(id, name, mimeType)",
                            includeItemsFromAllDrives=True,
                            supportsAllDrives=True,
                            pageSize=10,
                        )
                        .execute()
                    )
                    for gf in gres.get("files", []):
                        if self._matches_target(gf.get("name", ""), norm_q) > 0:
                            gn = gf.get("name", "").lower()
                            gm = gf.get("mimeType", "")
                            if (
                                gn.endswith((".xlsx", ".xls", ".csv"))
                                or "spreadsheet" in gm
                                or "excel" in gm
                            ):
                                b_data = self._download_excel_or_sheet(gf["id"], gm)
                                if b_data:
                                    fname = gf["name"]
                                    if not fname.lower().endswith((".xlsx", ".xls")):
                                        fname += ".xlsx"
                                    return b_data, fname
        except Exception as e:
            logger.error(f"Excel arama hatası ({folder_id}): {e}")

        return None, None

    def _download_excel_or_sheet(
        self, file_id: str, mime_type: str
    ) -> Optional[bytes]:
        svc = self._get_service()
        if not svc:
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
            logger.error(f"Tablo indirme hatası ({file_id}): {e}")
            return None

    def get_images_structured(
        self, folder_id: Optional[str], query_term: str
    ) -> Tuple[Optional[bytes], Optional[bytes]]:
        all_imgs = self._scan_images(folder_id, query_term)
        if not all_imgs:
            return None, None

        bm_file = None
        std_files = []
        for f in all_imgs:
            fn = f.get("name", "").lower()
            pn = f.get("parent_name", "").lower()
            if "bookmatch" in fn or "bm" in fn or "bookmatch" in pn:
                if not bm_file:
                    bm_file = f
            else:
                std_files.append(f)

        raw1 = (
            self._download_file(std_files[0]["id"])
            if std_files
            else (self._download_file(bm_file["id"]) if bm_file else None)
        )
        raw2 = None
        if bm_file and bm_file != (std_files[0] if std_files else None):
            raw2 = self._download_file(bm_file["id"])
        elif len(std_files) > 1:
            raw2 = self._download_file(std_files[1]["id"])
        elif raw1:
            raw2 = raw1

        return prepare_telegram_image_bytes(raw1), prepare_telegram_image_bytes(raw2)

    def get_best_slab_image(
        self, folder_id: Optional[str], query_term: str
    ) -> Optional[bytes]:
        all_imgs = self._scan_images(folder_id, query_term)
        if not all_imgs:
            return None
        clean_slabs = [
            f
            for f in all_imgs
            if "bookmatch" not in f.get("name", "").lower()
            and "bm" not in f.get("name", "").lower()
            and "bookmatch" not in f.get("parent_name", "").lower()
        ]
        target = clean_slabs[0] if clean_slabs else all_imgs[0]
        return prepare_telegram_image_bytes(self._download_file(target["id"]))

    def _scan_images(
        self, folder_id: Optional[str], query_term: str
    ) -> List[Dict[str, str]]:
        svc = self._get_service()
        if not svc:
            return []
        all_imgs: List[Dict[str, str]] = []
        try:
            if folder_id:
                q = (
                    f"'{folder_id}' in parents and trashed = false and ("
                    f"mimeType contains 'image/jpeg' or mimeType contains 'image/png' or mimeType contains 'image/webp'"
                    f")"
                )
                res = (
                    svc.files()
                    .list(
                        q=q,
                        spaces="drive",
                        fields="files(id, name, mimeType)",
                        includeItemsFromAllDrives=True,
                        supportsAllDrives=True,
                        pageSize=40,
                    )
                    .execute()
                )
                all_imgs.extend(res.get("files", []))

                for sf in self.list_child_folders(folder_id)[:6]:
                    sq = (
                        f"'{sf['id']}' in parents and trashed = false and ("
                        f"mimeType contains 'image/jpeg' or mimeType contains 'image/png' or mimeType contains 'image/webp'"
                        f")"
                    )
                    sres = (
                        svc.files()
                        .list(
                            q=sq,
                            spaces="drive",
                            fields="files(id, name, mimeType)",
                            includeItemsFromAllDrives=True,
                            supportsAllDrives=True,
                            pageSize=20,
                        )
                        .execute()
                    )
                    for img in sres.get("files", []):
                        img["parent_name"] = sf["name"]
                        all_imgs.append(img)

            if not all_imgs and query_term:
                norm_q = normalize_lot_code(query_term)
                for cand in [norm_q, norm_q.replace("-", " ")]:
                    q = (
                        f"trashed = false and name contains '{cand}' and ("
                        f"mimeType contains 'image/jpeg' or mimeType contains 'image/png'"
                        f")"
                    )
                    res = (
                        svc.files()
                        .list(
                            q=q,
                            spaces="drive",
                            fields="files(id, name, mimeType)",
                            includeItemsFromAllDrives=True,
                            supportsAllDrives=True,
                            pageSize=15,
                        )
                        .execute()
                    )
                    all_imgs.extend(res.get("files", []))
                    if all_imgs:
                        break
            return all_imgs
        except Exception as e:
            logger.warning(f"Görsel tarama uyarısı: {e}")
            return []

    def _download_file(self, file_id: str) -> Optional[bytes]:
        svc = self._get_service()
        if not svc:
            return None
        try:
            req = svc.files().get_media(fileId=file_id, supportsAllDrives=True)
            fh = io.BytesIO()
            dl = MediaIoBaseDownload(fh, req)
            done = False
            while not done:
                _, done = dl.next_chunk()
            fh.seek(0)
            return fh.read()
        except Exception:
            return None


# =====================================================================
# 4. ÇİFT KATMANLI BAŞLIK & FOTOĞRAF METRAJ ÇÖZÜMLEYİCİ ENVANTER MOTORU
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

        return {
            "m2": m2_col,
            "len": len_col,
            "height": height_col,
            "dim": dim_col,
            "slab": slab_col,
            "adet": adet_col,
            "thick": thick_col,
            "surf": surf_col,
        }

    def parse_packing_list_bytes(self, excel_bytes: bytes) -> Optional[Dict[str, Any]]:
        """
        2 satırlı birleştirilmiş başlıkları (Merged Headers) birleştirerek okur.
        Sütun isimleri standart dışıysa sayısal değer dağılımından (Boy, En, M2)
        sütunları otomatik tespit eder.
        """
        try:
            xls = pd.ExcelFile(io.BytesIO(excel_bytes))
            for sheet in xls.sheet_names:
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

                # Birleştirilmiş 2 katmanlı başlıkları (üst satır + alt satır) birleştir!
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
                    composite_headers[col_i] = " ".join(parts) if parts else f"COL_{col_i}"

                data_start_row = header_idx + 1 if best_score >= 1 else 0
                df = raw_df.iloc[data_start_row:].copy()
                df.columns = list(range(raw_df.shape[1]))

                mapped = self._detect_columns(composite_headers)

                # Sayısal Sütun Profilleme Yedeği (Başlıklar farklı dilde veya boşsa)
                if mapped["len"] is None or mapped["height"] is None or mapped["m2"] is None:
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

                valid_rows = []
                total_m2 = 0.0
                total_slabs = 0
                lengths = []
                heights = []

                for _, row in df.iterrows():
                    row_str = " ".join(
                        [normalize_turkish_key(v) for v in row.values if pd.notna(v)]
                    )
                    if any(
                        tot_w in row_str
                        for tot_w in ["TOPLAM", "TOTAL", "GENEL TOPLAM", "SUBTOTAL"]
                    ):
                        continue

                    raw_l = safe_float(row.get(len_col)) if len_col is not None else 0.0
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

                    raw_m2 = safe_float(row.get(m2_col)) if m2_col is not None else 0.0

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
                            "slab_no": row.get(slab_col, len(valid_rows) + 1)
                            if slab_col is not None
                            else len(valid_rows) + 1,
                            "length_cm": l_cm,
                            "height_cm": h_cm,
                            "thickness": "2 cm",
                            "m2": m2_val,
                        }
                    )

                if total_m2 <= 0 and m2_col is not None:
                    col_vals = [
                        safe_float(v)
                        for v in df[m2_col].dropna().values
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
                if thick_col is not None and not df[thick_col].dropna().empty:
                    raw_th = safe_float(df[thick_col].dropna().iloc[0])
                    if 17 <= raw_th <= 24 or raw_th == 2:
                        thick_val = "2 cm"
                    elif 27 <= raw_th <= 34 or raw_th == 3:
                        thick_val = "3 cm"
                    elif 0 < raw_th <= 10:
                        thick_val = f"{raw_th:g} cm"

                surf_val = "Polished"
                if surf_col is not None and not df[surf_col].dropna().empty:
                    surf_val = str(df[surf_col].dropna().iloc[0]).strip().title()

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
                    "total_m2": total_m2,
                    "total_slabs": total_slabs,
                    "thickness": thick_val,
                    "surface": surf_val,
                    "dimensions": dims_str,
                    "clean_rows": valid_rows,
                }
        except Exception as e:
            logger.error(f"Excel parse hatası: {e}")
        return None

    def _estimate_stock_from_folder_images(
        self, folder_id: Optional[str], query_term: str
    ) -> Optional[Dict[str, Any]]:
        """
        Eğer klasörde henüz Excel dosyası yoksa, klasördeki plaka fotoğraflarını
        (Bookmatch hariç) sayar; dosya isimlerinde ebat veya m2 yazıyorsa (Örn: 329 x 185 - 6,09 m2)
        doğrudan okur, yazmıyorsa plaka sayısından stok detayını çıkarır.
        """
        all_imgs = self.drive._scan_images(folder_id, query_term)
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

        for f in target_imgs:
            fname = f.get("name", "")
            dim_m = re.search(r"(\d{2,4})\s*[xX*]\s*(\d{2,4})", fname)
            area_m = re.search(
                r"(\d+[.,]\d+)\s*(?:m2|m²|sqm)", fname, re.IGNORECASE
            )
            l_v = safe_float(dim_m.group(1)) if dim_m else 0.0
            h_v = safe_float(dim_m.group(2)) if dim_m else 0.0
            a_v = safe_float(area_m.group(1)) if area_m else 0.0
            m2_v, l_cm, h_cm = compute_slab_m2(l_v, h_v, a_v)
            if m2_v > 0:
                parsed_m2_list.append(m2_v)
            if l_cm > 20 and h_cm > 20:
                parsed_l_list.append(l_cm)
                parsed_h_list.append(h_cm)

        if parsed_m2_list:
            avg_m2 = sum(parsed_m2_list) / len(parsed_m2_list)
            total_m2 = avg_m2 * slab_count
        else:
            # Klasörde sadece plaka fotoğrafları varsa standart plaka ortalaması (5.40 m²) ile hesapla
            total_m2 = round(slab_count * 5.40, 2)

        if parsed_l_list and parsed_h_list:
            avg_l = int(round(sum(parsed_l_list) / len(parsed_l_list)))
            avg_h = int(round(sum(parsed_h_list) / len(parsed_h_list)))
            dims_str = f"{avg_l} x {avg_h} cm"
        else:
            dims_str = "290 x 185 cm (Avg.)"

        return {
            "total_m2": total_m2,
            "total_slabs": slab_count,
            "thickness": "2 cm",
            "surface": "Polished",
            "dimensions": dims_str,
            "clean_rows": [],
        }

    def get_lot_info(
        self,
        query: str,
        pre_folder_id: Optional[str] = None,
        pre_folder_link: Optional[str] = None,
    ) -> Dict[str, Any]:
        norm_key = normalize_lot_code(query)
        folder_id = pre_folder_id
        folder_link = pre_folder_link

        if not folder_id:
            folder_id, folder_link = self.drive.find_folder(norm_key)

        excel_bytes, excel_name, data = None, None, None

        excel_bytes, excel_name = self.drive.get_excel_from_folder(
            folder_id, query_hint=norm_key, max_depth=2
        )
        if excel_bytes:
            data = self.parse_packing_list_bytes(excel_bytes)

        # Klasörde Excel yoksa plaka fotoğraflarından stok adet ve m2 bilgisini çıkar
        if not data or data.get("total_m2", 0.0) <= 0:
            img_est = self._estimate_stock_from_folder_images(folder_id, norm_key)
            if img_est:
                data = img_est

        if not data:
            data = {
                "total_m2": 0.0,
                "total_slabs": 0,
                "thickness": "2 cm",
                "surface": "Polished",
                "dimensions": "Standard Slab Sizes",
                "clean_rows": [],
            }

        return {
            "status": "success",
            "query": norm_key,
            "drive_link": folder_link,
            "folder_id": folder_id,
            "excel_name": excel_name or f"{norm_key}_PackingList.xlsx",
            "excel_bytes": excel_bytes,
            "data": data,
        }

    def scan_folder_recursively_for_lots(
        self, parent_id: str, depth: int = 0, max_depth: int = 2
    ) -> List[Dict[str, Any]]:
        results = []
        if depth > max_depth or not parent_id:
            return results

        sub_folders = self.drive.list_child_folders(parent_id)
        for sf in sub_folders:
            ex_bytes, _ = self.drive.get_excel_from_folder(
                sf["id"], query_hint="", max_depth=0
            )
            if ex_bytes:
                parsed = self.parse_packing_list_bytes(ex_bytes)
                if parsed and parsed["total_m2"] > 0:
                    results.append(
                        {
                            "name": sf["name"],
                            "id": sf["id"],
                            "link": sf["link"],
                            "data": parsed,
                        }
                    )
                    continue

            deeper = self.scan_folder_recursively_for_lots(
                sf["id"], depth=depth + 1, max_depth=max_depth
            )
            if deeper:
                results.extend(deeper)
            else:
                est = self._estimate_stock_from_folder_images(sf["id"], sf["name"])
                results.append(
                    {
                        "name": sf["name"],
                        "id": sf["id"],
                        "link": sf["link"],
                        "data": est
                        if est
                        else {
                            "total_m2": 0.0,
                            "total_slabs": 0,
                            "thickness": "2 cm",
                            "surface": "Polished",
                            "dimensions": "Standard Slab Sizes",
                            "clean_rows": [],
                        },
                    }
                )
        return results

    def deep_scan_collections(self) -> str:
        collections = [
            "Arslanlar Collection",
            "Arslanlar Series",
            "Arslanlar Exclusive Collection",
        ]
        report_lines = ["🔍 <b>ARGON MASTER - DEEP SCAN STOK RAPORU</b> 🔍\n"]
        total_factory_m2 = 0.0
        total_factory_slabs = 0

        for coll in collections:
            fid, _ = self.drive.find_folder(coll)
            if not fid:
                continue

            report_lines.append(f"📁 <b>{coll.upper()}</b>")
            lot_items = self.scan_folder_recursively_for_lots(fid, max_depth=2)
            coll_m2 = 0.0
            coll_slabs = 0

            for item in lot_items:
                d = item["data"]
                if d["total_m2"] > 0:
                    report_lines.append(
                        f"  ├ <b>{item['name']}</b>: <code>{d['total_m2']:,.2f} m²</code> ({d['total_slabs']} Plaka)"
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
# 5. BRONZ ASLAN LOGOLU & NET STOK DETAYLI PREMIUM PDF MOTORU
# =====================================================================
class PDFPresentationEngine:
    def __init__(self, inventory: InventoryEngine, drive: DriveExplorer):
        self.inventory = inventory
        self.drive = drive

    @staticmethod
    def _build_bronze_lion_vector_header(subtitle_text: str) -> Drawing:
        """
        Arslanlar Marble'ın kurumsal Bronz Aslan ('logo_bronz') amblemini ve
        bronz-kahve gradyanlı ARSLANLAR MARBLE tipografisini vektörel olarak çizer.
        """
        d = Drawing(530, 108)
        dark_bronze = colors.HexColor("#431C0E")
        mid_bronze = colors.HexColor("#8E4B2A")
        warm_bronze = colors.HexColor("#C27D4E")
        light_bronze = colors.HexColor("#D4956A")
        slate = colors.HexColor("#4A4E69")

        # Vektörel kükreyen aslan başı silüeti (sağa bakan profil ve pençe)
        lion_mane_outer = Polygon(
            [
                222, 58,
                228, 70,
                238, 82,
                234, 83,
                247, 92,
                243, 93,
                258, 100,
                274, 103,
                290, 99,
                302, 89,
                309, 80,
                305, 76,
                299, 77,
                302, 71,
                296, 68,
                289, 74,
                283, 66,
                279, 59,
                274, 66,
                266, 67,
                272, 60,
                272, 54,
                292, 62,
                308, 62,
                316, 58,
                312, 53,
                296, 52,
                278, 44,
                265, 45,
                252, 54,
                222, 58,
            ],
            fillColor=mid_bronze,
            strokeColor=dark_bronze,
            strokeWidth=0.6,
        )
        lion_highlight = Polygon(
            [
                256, 98,
                274, 102,
                290, 98,
                302, 88,
                308, 80,
                299, 76,
                288, 66,
                272, 54,
                294, 61,
                312, 59,
                309, 54,
                292, 51,
                275, 45,
                260, 56,
                256, 98,
            ],
            fillColor=warm_bronze,
            strokeColor=None,
        )
        lion_paw_lower = Polygon(
            [
                288, 49,
                298, 53,
                310, 49,
                307, 45,
                295, 46,
                288, 49,
            ],
            fillColor=light_bronze,
            strokeColor=None,
        )
        d.add(lion_mane_outer)
        d.add(lion_highlight)
        d.add(lion_paw_lower)

        # ARSLANLAR (Bronz tonlu kurumsal başlık)
        d.add(
            String(
                265,
                26,
                "ARSLANLAR",
                fontName="Helvetica-Bold",
                fontSize=22,
                fillColor=dark_bronze,
                textAnchor="middle",
            )
        )
        d.add(
            String(
                265,
                14,
                "M   A   R   B   L   E",
                fontName="Helvetica",
                fontSize=10,
                fillColor=warm_bronze,
                textAnchor="middle",
            )
        )
        d.add(
            String(
                265,
                3,
                subtitle_text,
                fontName="Helvetica-Bold",
                fontSize=8,
                fillColor=slate,
                textAnchor="middle",
            )
        )
        d.add(Line(0, 0, 530, 0, strokeColor=warm_bronze, strokeWidth=1.2))
        return d

    def _build_header_flowables(self, pdf_title: str, styles: Any) -> List[Any]:
        """
        Önce orijinal yüksek çözünürlüklü 'logo_bronz' görselini kontrol eder,
        varsa doğrudan kullanır; yoksa vektörel Bronz Aslan logosunu çizer.
        """
        BRONZE = colors.HexColor("#A4633A")
        SLATE = colors.HexColor("#2B2D42")

        logo_bytes = self.drive.get_company_logo_bytes()
        if logo_bytes:
            try:
                im = Image.open(io.BytesIO(logo_bytes)).convert("RGB")
                w, h = im.size
                target_w = 155.0
                target_h = target_w * (h / float(w))
                img_io = io.BytesIO()
                im.save(img_io, format="JPEG", quality=95)
                img_io.seek(0)
                rl_logo = RLImage(img_io, width=target_w, height=target_h)
                rl_logo.hAlign = "CENTER"

                sub_s = ParagraphStyle(
                    "SubTitleLogo",
                    parent=styles["Normal"],
                    fontName="Helvetica-Bold",
                    fontSize=9,
                    leading=12,
                    textColor=SLATE,
                    alignment=1,
                    spaceBefore=4,
                    spaceAfter=6,
                )
                return [
                    rl_logo,
                    Paragraph(pdf_title, sub_s),
                    HRFlowable(
                        width="100%", thickness=1.4, color=BRONZE, spaceAfter=14
                    ),
                ]
            except Exception as e:
                logger.warning(f"Logo görseli işlenemedi, vektörel kullanılıyor: {e}")

        return [
            self._build_bronze_lion_vector_header(pdf_title),
            Spacer(1, 14),
        ]

    def generate_presentation_pdf(
        self, queries: List[str], collection_folder_name: Optional[str] = None
    ) -> Tuple[bytes, str]:
        cards = []
        pdf_title = "EXECUTIVE STOCK & SLAB PRESENTATION"

        if collection_folder_name:
            pdf_title = f"{normalize_turkish_key(collection_folder_name)} — CATALOG PRESENTATION"
            parent_id, _ = self.drive.find_folder(collection_folder_name)
            if parent_id:
                scanned_items = self.inventory.scan_folder_recursively_for_lots(
                    parent_id, max_depth=2
                )
                for item in scanned_items[:24]:
                    img = self.drive.get_best_slab_image(item["id"], item["name"])
                    cards.append(
                        {
                            "title": item["name"].upper(),
                            "data": item["data"],
                            "drive_link": item["link"],
                            "image_bytes": img,
                        }
                    )
        else:
            for q in queries:
                info = self.inventory.get_lot_info(q)
                img = self.drive.get_best_slab_image(info["folder_id"], q)
                cards.append(
                    {
                        "title": q.upper(),
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
            topMargin=24,
            bottomMargin=28,
        )
        styles = getSampleStyleSheet()
        NAVY = colors.HexColor("#14213D")
        BRONZE = colors.HexColor("#A4633A")
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

            # Toplam m² ve Plaka Adedi her zaman net ve kalın gösterilir
            if m2 > 0 or slabs > 0:
                stock_str = f"<b>{m2:,.2f} m²  ({slabs} Slabs)</b>"
                heading_stock_badge = f" — <font color='#A4633A'>{m2:,.2f} m² | {slabs} Slabs</font>"
            else:
                stock_str = "<b>Check Current Lot Packing List</b>"
                heading_stock_badge = ""

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

            rl_img = Paragraph("High-Res Slab Image in Drive", tdv_s)
            if card.get("image_bytes"):
                try:
                    img_io = io.BytesIO(card["image_bytes"])
                    rl_img = RLImage(img_io, width=185, height=118)
                except Exception:
                    pass

            dlink = card.get("drive_link")
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
                [Paragraph("Active Code", tdl_s), Paragraph(card["title"], tdv_s)],
                [
                    Paragraph("High-Res Slab Photos", tdl_s),
                    Paragraph(link_html, tdv_s),
                ],
            ]

            st = Table(spec_rows, colWidths=[130, 200])
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
                        Paragraph(
                            f"{idx}. {card['title']}{heading_stock_badge}", ch_s
                        ),
                        mr,
                        Spacer(1, 14),
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
        fname = (
            f"Arslanlar_{normalize_turkish_key(collection_folder_name).replace(' ', '_')}.pdf"
            if collection_folder_name
            else f"Arslanlar_Presentation_{'_'.join(queries[:3])}.pdf"
        )
        return pdf_io.read(), fname

    def generate_packing_list_pdf(
        self, clean_rows: List[Dict[str, Any]], lot_name: str
    ) -> bytes:
        pdf_io = io.BytesIO()
        doc = SimpleDocTemplate(
            pdf_io,
            pagesize=landscape(A4),
            rightMargin=25,
            leftMargin=25,
            topMargin=25,
            bottomMargin=25,
        )
        styles = getSampleStyleSheet()
        NAVY = colors.HexColor("#14213D")
        BRONZE = colors.HexColor("#A4633A")
        SLATE = colors.HexColor("#2B2D42")

        title_p = ParagraphStyle(
            "PlT",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=18,
            leading=22,
            textColor=NAVY,
            alignment=1,
        )
        sub_p = ParagraphStyle(
            "PlS",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=10,
            leading=14,
            textColor=BRONZE,
            alignment=1,
            spaceAfter=12,
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
            Paragraph("ARSLANLAR MARBLE", title_p),
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
                Paragraph("Net M2", th_p),
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
                resp = requests.post(url, json=payload, timeout=12)
                if resp.status_code == 200:
                    txt = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
                    found = [
                        k
                        for k in DEFAULT_STONE_DNA.keys()
                        if k in normalize_turkish_key(txt)
                    ]
                    if found:
                        return found[:2]
            except Exception as e:
                logger.warning(f"Gemini Vision uyarısı: {e}")

        try:
            im = Image.open(io.BytesIO(image_bytes)).convert("RGB").resize((64, 64))
            pixels = list(im.getdata())
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
            if "LOGO" in raw:
                return {"should_reply": True, "is_logo_upload": True}
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

        collection_folder = None
        if "COLLECTION" in raw or "SERIES" in raw or "EXCLUSIVE" in raw:
            if "EXCLUSIVE" in raw:
                collection_folder = "Arslanlar Exclusive Collection"
            elif "SERIES" in raw:
                collection_folder = "Arslanlar Series"
            else:
                collection_folder = "Arslanlar Collection"

        want_pdf = any(kw in raw for kw in ["PDF", "SUNUM", "KATALOG"])
        want_packing = any(kw in raw for kw in ["CEKI", "PACKING", "SEVK"])
        want_only_link = any(kw in raw for kw in ["LINK", "DRIVE"]) and not any(
            kw in raw for kw in ["FOTO", "RESIM", "GORSEL", "PDF", "SUNUM", "CEKI"]
        )
        want_photo = (
            any(kw in raw for kw in ["FOTO", "RESIM", "GORSEL"])
            or (bool(matched) and not want_only_link and not want_pdf and not want_packing)
        )

        if (
            is_group
            and not bot_mentioned
            and not matched
            and not want_pdf
            and not want_packing
            and not want_photo
            and not want_only_link
            and not collection_folder
            and not want_deep_scan
        ):
            return {"should_reply": False}

        return {
            "should_reply": True,
            "is_visual_search": False,
            "is_logo_upload": False,
            "matched_queries": list(dict.fromkeys(matched))[:10],
            "collection_folder": collection_folder,
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

    # 0-A. KULLANICI TELEGRAM'DAN LOGO FOTOĞRAFI GÖNDERDİYSE KAYDET
    if intent.get("is_logo_upload"):
        try:
            photo_file = await msg.photo[-1].get_file()
            img_bytes = bytes(await photo_file.download_as_bytearray())
            with open(CUSTOM_LOGO_PATH, "wb") as f:
                f.write(img_bytes)
            drive_explorer._cached_logo_bytes = img_bytes
            await msg.reply_text(
                "✅ <b>Arslanlar Marble Kurumsal Logosu</b> yüksek çözünürlüklü olarak kaydedildi ve tüm PDF sunumlarına entegre edildi.",
                parse_mode=ParseMode.HTML,
            )
        except Exception as e:
            await msg.reply_text(f"⚠️ Logo kaydedilemedi: {e}")
        return

    # 0-B. FOTOĞRAF ATILDIĞINDA BENZER TAŞ BULMA ÖZELLİĞİ
    if intent.get("is_visual_search"):
        status = await msg.reply_text(
            "🔍 <b>Gönderdiğiniz mermer fotoğrafı analiz ediliyor...</b>\n"
            "Drive'daki seleksiyonlarımızla doku ve damar eşleştirmesi yapılıyor.",
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

    # 1. DEEP SCAN (TOPLAM STOK TARAMASI)
    if (
        intent.get("want_deep_scan")
        and not intent.get("matched_queries")
        and not intent.get("collection_folder")
    ):
        status = await msg.reply_text(
            "🔎 <b>Argon Deep Scan Başlatıldı.</b>\n"
            "Arslanlar Collection, Series ve Exclusive klasörlerindeki tüm alt lotların Excel tabloları taranıyor...",
            parse_mode=ParseMode.HTML,
        )
        try:
            report = await asyncio.to_thread(inventory_engine.deep_scan_collections)
            await status.edit_text(report, parse_mode=ParseMode.HTML)
        except Exception as e:
            await status.edit_text(f"⚠️ Tarama uyarısı: {e}")
        return

    queries = intent.get("matched_queries", [])
    collection_folder = intent.get("collection_folder")

    if not queries and not collection_folder:
        if not is_group or bot_mentioned:
            await msg.reply_text(
                "👋 <b>Argon Master V8 Aktif.</b>\n"
                "• Lot & Fotoğraf: <code>A-371</code> veya <code>A-371 fotoğraf</code>\n"
                "• Drive Linki: <code>A-371 drive linki lütfen</code>\n"
                "• Çoklu Tek PDF Sunum: <code>A-96 ve A-383 sunum hazırla</code>\n"
                "• Koleksiyon PDF: <code>Arslanlar Collection sunum hazırla</code>\n"
                "• Çeki Listesi (Excel + PDF): <code>A-371 çeki listesi</code>\n"
                "• Genel Stok: <code>Toplam stoğum ne kadar</code>\n"
                "• Görsel Eşleştirme: Doğrudan mermer fotoğrafı gönderin.",
                parse_mode=ParseMode.HTML,
            )
        return

    # 2. KOLEKSİYON BAZLI PDF SUNUMU
    if collection_folder and intent.get("want_pdf"):
        status_msg = await msg.reply_text(
            f"📁 <b>{collection_folder}</b> altındaki taş ve lot klasörleri taranıp tek bir Premium PDF'te birleştiriliyor...",
            parse_mode=ParseMode.HTML,
        )
        try:
            pdf_b, pdf_n = await asyncio.wait_for(
                asyncio.to_thread(
                    pdf_engine.generate_presentation_pdf,
                    [],
                    collection_folder_name=collection_folder,
                ),
                timeout=90,
            )
            await msg.reply_document(
                document=io.BytesIO(pdf_b),
                filename=pdf_n,
                caption=f"📄 <b>{collection_folder}</b> — Executive PDF Presentation",
                parse_mode=ParseMode.HTML,
            )
            await status_msg.delete()
        except Exception as e:
            await status_msg.edit_text(f"⚠️ Koleksiyon PDF uyarısı: {e}")
        return

    # 3. ÇOKLU VEYA TEKLİ LOT PDF SUNUMU (TEK PDF'TE BİRLEŞTİRİR)
    if queries and intent.get("want_pdf"):
        status_msg = await msg.reply_text(
            f"📄 <b>{', '.join(queries)}</b> için Kurumsal Logolu ve Stok Detaylı Premium PDF Sunumu hazırlanıyor...",
            parse_mode=ParseMode.HTML,
        )
        try:
            pdf_b, pdf_n = await asyncio.wait_for(
                asyncio.to_thread(pdf_engine.generate_presentation_pdf, queries),
                timeout=60,
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

    # 4. STANDART LOT SORGUSU, LİNK VE ÇEKİ LİSTESİ (EXCEL + PDF)
    for q in queries:
        try:
            info = await asyncio.wait_for(
                asyncio.to_thread(inventory_engine.get_lot_info, q), timeout=35
            )
            d = info["data"]
            fid = info["folder_id"]
            flink = info["drive_link"]

            if intent.get("want_packing"):
                if info["excel_bytes"]:
                    await msg.reply_document(
                        document=io.BytesIO(info["excel_bytes"]),
                        filename=info["excel_name"],
                        caption=f"📦 <b>{q}</b> Orijinal Excel Çeki Listesi",
                        parse_mode=ParseMode.HTML,
                    )
                    if d.get("clean_rows"):
                        try:
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
                        except Exception as pe:
                            logger.warning(f"PDF Çeki listesi uyarısı: {pe}")
                else:
                    link_note = (
                        f"\n📂 <a href='{flink}'>Drive Klasörünü Aç ({q})</a>"
                        if flink
                        else ""
                    )
                    await msg.reply_text(
                        f"⚠️ <b>{q}</b> klasöründe henüz Excel çeki listesi bulunamadı.{link_note}",
                        parse_mode=ParseMode.HTML,
                    )
                continue

            m2_val = d.get("total_m2", 0.0)
            slabs_val = d.get("total_slabs", 0)
            stock_str = (
                f"{m2_val:,.2f} m² ({slabs_val} Plaka)"
                if (m2_val > 0 or slabs_val > 0)
                else "Klasörde Excel Bekleniyor"
            )
            link_html = (
                f"<a href='{flink}'>Klasöre Git ({q})</a>"
                if flink
                else "Drive'da Klasör Bulunamadı"
            )

            text_response = (
                f"💎 <b>{q.upper()}</b>\n"
                f"• Mevcut Stok: <code>{stock_str}</code>\n"
                f"• Ortalama Ebat: <code>{d.get('dimensions', 'Standard Slab Sizes')}</code>\n"
                f"• Yüzey & Kalınlık: <code>{d.get('surface', 'Polished')} ({d.get('thickness', '2 cm')})</code>\n"
                f"• Drive Linki: {link_html}"
            )

            if intent.get("want_only_link"):
                await msg.reply_text(text_response, parse_mode=ParseMode.HTML)
                continue

            img1, img2 = await asyncio.wait_for(
                asyncio.to_thread(drive_explorer.get_images_structured, fid, q),
                timeout=30,
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
                except Exception as me:
                    logger.warning(f"Albüm gönderilemedi, tekli gönderiliyor: {me}")
                    await msg.reply_text(text_response, parse_mode=ParseMode.HTML)
            else:
                await msg.reply_text(text_response, parse_mode=ParseMode.HTML)
        except Exception as ex:
            logger.error(f"Sorgu hatası ({q}): {ex}")
            await msg.reply_text(
                f"⚠️ <b>{q}</b> sorgulanırken bağlantı zaman aşımına uğradı, lütfen tekrar deneyin.",
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

    application.add_handler(
        MessageHandler(
            (filters.TEXT | filters.PHOTO) & (~filters.COMMAND), handle_message
        )
    )

    logger.info("🚀 Argon Master V8 Official Logo & Exact Stock başlatılıyor...")
    application.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
