import os
import io
import re
import json
import time
import logging
import asyncio
import threading
from datetime import datetime
from collections import deque
from typing import Dict, List, Optional, Tuple, Any

# =====================================================================
# 0. RENDER PORTUNU 0. SANİYEDE AÇ (DEPLOY TAKILMASINI ÖNLER)
# =====================================================================
from flask import Flask

flask_app = Flask(__name__)

@flask_app.route("/")
def health_check():
    return "Argon Master V6.2 Thread-Safe is Live!", 200

@flask_app.route("/healthz")
def healthz():
    return "OK", 200

def _start_flask_immediately():
    port = int(os.environ.get("PORT", 10000))
    flask_app.run(host="0.0.0.0", port=port, use_reloader=False)

threading.Thread(target=_start_flask_immediately, daemon=True).start()

# =====================================================================
# 1. KÜTÜPHANELER VE SİSTEM AYARLARI
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
logger = logging.getLogger("ArgonMaster_V6_2")

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GOOGLE_SA_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
DEFAULT_DRIVE_LINK = (
    os.getenv("GOOGLE_DRIVE_LINK")
    or os.getenv("DEFAULT_DRIVE_LINK")
    or "https://drive.google.com"
).strip()

HYPERLINKS_DB_FILE = "hyperlinks_db.json"


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
    match = re.match(r"^([A-Z]+)[-\s]*(\d+)$", raw)
    return f"{match.group(1)}-{match.group(2)}" if match else raw


def safe_float(val: Any) -> float:
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    # Metin içindeki ilk anlamlı ondalık sayıyı yakala (örn: "6,09 m2" -> "6,09")
    m = re.search(r"[-+]?\d+(?:[.,]\d+)?", s)
    if not m:
        return 0.0
    num_str = m.group(0).replace(",", ".")
    try:
        return float(num_str)
    except Exception:
        return 0.0


def compute_slab_m2(l_val: float, h_val: float, raw_m2: float) -> Tuple[float, float, float]:
    """
    Boy ve En değerlerini cm'ye normalize eder ve M2 değerini doğrular/hesaplar.
    (mm, cm veya m cinsinden yazılmış tabloları otomatik algılar)
    """
    # Boy ve En birimlerini cm'ye çevir
    if l_val > 600:  # mm yazılmış (örn: 2850)
        l_cm = l_val / 10.0
    elif 0 < l_val < 10:  # metre yazılmış (örn: 2.85)
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

    # Eğer raw_m2 cm2 veya mm2 olarak gelmişse düzelt
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
# 3. THREAD-SAFE GOOGLE DRIVE MOTORU (ASLA KİLİTLENMEZ)
# =====================================================================
class DriveExplorer:
    """
    Her thread için kendi httplib2/Drive istemcisini (thread-local) oluşturur.
    Böylece arka arkaya 10 mesaj geldiğinde SSL kilitlenmesi (deadlock) yaşanmaz.
    """
    def __init__(self, sa_json_str: str):
        self.sa_json_str = sa_json_str
        self._creds = None
        self._local = threading.local()
        self.hyperlinks_db = load_json_file(HYPERLINKS_DB_FILE, {})
        self._db_lock = threading.Lock()

        if self.sa_json_str:
            try:
                info = json.loads(self.sa_json_str)
                self._creds = service_account.Credentials.from_service_account_info(
                    info, scopes=["https://www.googleapis.com/auth/drive.readonly"]
                )
                logger.info("Google Drive Kimlik Bilgileri Hazır.")
            except Exception as e:
                logger.error(f"Google Service Account JSON Hatası: {e}")

    def _get_service(self):
        if not self._creds:
            return None
        if getattr(self._local, "service", None) is None:
            http = httplib2.Http(timeout=18)
            authed_http = AuthorizedHttp(self._creds, http=http)
            self._local.service = build(
                "drive", "v3", http=authed_http, cache_discovery=False
            )
        return self._local.service

    def find_folder(self, query_term: str) -> Tuple[Optional[str], Optional[str]]:
        norm_key = normalize_lot_code(query_term)
        svc = self._get_service()
        if not svc:
            return None, None

        # Önbellekte varsa bile geçerli olup olmadığını kontrol et
        if norm_key in self.hyperlinks_db:
            url = self.hyperlinks_db[norm_key]
            fid = self._extract_folder_id(url)
            if fid:
                return fid, url

        raw_clean = query_term.strip().replace("'", "")
        cands = [
            norm_key,
            raw_clean,
            norm_key.replace("-", " "),
            norm_key.replace("-", ""),
            norm_key.replace("SILVER", "SİLVER").replace("RIVER", "RİVER"),
        ]
        cands = list(dict.fromkeys([c for c in cands if len(c) >= 2]))

        for cand in cands:
            try:
                q = (
                    f"mimeType = 'application/vnd.google-apps.folder' "
                    f"and trashed = false and name contains '{cand}'"
                )
                res = (
                    svc.files()
                    .list(
                        q=q,
                        spaces="drive",
                        fields="files(id, name, webViewLink)",
                        includeItemsFromAllDrives=True,
                        supportsAllDrives=True,
                        pageSize=25,
                    )
                    .execute()
                )
                files = res.get("files", [])
                if files:
                    # En uygun klasörü seç (Tam eşleşen veya içinde Excel barındıran)
                    best = files[0]
                    for f in files:
                        fname_norm = normalize_lot_code(f["name"])
                        if fname_norm == norm_key or norm_key in fname_norm.split():
                            best = f
                            break
                    fid = best["id"]
                    link = best.get(
                        "webViewLink", f"https://drive.google.com/drive/folders/{fid}"
                    )
                    with self._db_lock:
                        self.hyperlinks_db[norm_key] = link
                        save_json_file(HYPERLINKS_DB_FILE, self.hyperlinks_db)
                    return fid, link
            except Exception as e:
                logger.warning(f"Klasör arama uyarısı ({cand}): {e}")
                continue
        return None, None

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
        except Exception as e:
            logger.warning(f"Alt klasör listeleme uyarısı: {e}")
            return []

    def get_excel_from_folder(
        self, folder_id: str, query_hint: str = "", max_depth: int = 2
    ) -> Tuple[Optional[bytes], Optional[str]]:
        """
        Drive API'nin 'name contains .xls' kısıtlamasına takılmamak için
        klasördeki TÜM dosyaları listeler ve Python tarafında Excel / Google Sheet
        veya kısayol (shortcut) dosyalarını yakalar.
        """
        svc = self._get_service()
        if not svc or not folder_id:
            return None, None

        try:
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

                # Kısayol (Shortcut) ise hedef dosyayı kontrol et
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

            # Alt klasörlere in
            if max_depth > 0:
                for sf in sub_folders[:15]:
                    b_data, fname = self.get_excel_from_folder(
                        sf["id"], query_hint="", max_depth=max_depth - 1
                    )
                    if b_data:
                        return b_data, fname

            # Hâlâ bulunamadıysa ve query_hint (örn: A-383) varsa tüm Drive'da bu lot adıyla Excel ara
            if query_hint and max_depth == 2:
                norm_q = normalize_lot_code(query_hint)
                for cand in [norm_q, norm_q.replace("-", " "), norm_q.replace("-", "")]:
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
                q = f"'{folder_id}' in parents and trashed = false and mimeType contains 'image/'"
                res = (
                    svc.files()
                    .list(
                        q=q,
                        spaces="drive",
                        fields="files(id, name, mimeType)",
                        includeItemsFromAllDrives=True,
                        supportsAllDrives=True,
                        pageSize=20,
                    )
                    .execute()
                )
                all_imgs.extend(res.get("files", []))

                for sf in self.list_child_folders(folder_id)[:8]:
                    sq = f"'{sf['id']}' in parents and trashed = false and mimeType contains 'image/'"
                    sres = (
                        svc.files()
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
                    for img in sres.get("files", []):
                        img["parent_name"] = sf["name"]
                        all_imgs.append(img)

            if not all_imgs and query_term:
                norm_q = normalize_lot_code(query_term)
                for cand in [norm_q, norm_q.replace("-", " ")]:
                    q = f"trashed = false and name contains '{cand}' and mimeType contains 'image/'"
                    res = (
                        svc.files()
                        .list(
                            q=q,
                            spaces="drive",
                            fields="files(id, name, mimeType)",
                            includeItemsFromAllDrives=True,
                            supportsAllDrives=True,
                            pageSize=6,
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
        except Exception as e:
            logger.warning(f"Görsel indirme uyarısı ({file_id}): {e}")
            return None

    @staticmethod
    def _extract_folder_id(url: str) -> Optional[str]:
        if not url:
            return None
        m = re.search(r"/folders/([a-zA-Z0-9_-]+)", url)
        if m:
            return m.group(1)
        m2 = re.search(r"id=([a-zA-Z0-9_-]+)", url)
        return m2.group(1) if m2 else None


# =====================================================================
# 4. AKILLI TABLO ÇÖZÜMLEYİCİ VE ENVANTER MOTORU
# =====================================================================
class InventoryEngine:
    def __init__(self, drive: DriveExplorer):
        self.drive = drive

    def parse_packing_list_bytes(self, excel_bytes: bytes) -> Optional[Dict[str, Any]]:
        """
        Her türlü çeki listesi formatını (üstte firma anteti olanlar, birleştirilmiş
        hücreliler, mm/cm/m farklılıkları, formüllü M2 sütunları) okur.
        """
        try:
            xls = pd.ExcelFile(io.BytesIO(excel_bytes))
            for sheet in xls.sheet_names:
                raw_df = pd.read_excel(xls, sheet_name=sheet, header=None)
                if raw_df.empty or len(raw_df) < 2:
                    continue

                header_idx = None
                best_score = 0
                header_keywords = [
                    "M2",
                    "M²",
                    "SQM",
                    "ALAN",
                    "AREA",
                    "METRAJ",
                    "BOY",
                    "UZUNLUK",
                    "LENGTH",
                    "EN",
                    "YUKSEKLIK",
                    "WIDTH",
                    "HEIGHT",
                    "PLAKA",
                    "SLAB",
                    "ADET",
                    "KALINLIK",
                    "THICKNESS",
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
                    )
                    if score > best_score:
                        best_score = score
                        header_idx = idx

                if header_idx is not None and best_score >= 2:
                    df = pd.read_excel(xls, sheet_name=sheet, header=header_idx)
                else:
                    df = pd.read_excel(xls, sheet_name=sheet)

                if df.empty:
                    continue

                col_map = {c: normalize_turkish_key(c) for c in df.columns}

                m2_col = next(
                    (
                        c
                        for c, n in col_map.items()
                        if any(
                            k in n
                            for k in ["M2", "M²", "ALAN", "NET", "METRAJ", "AREA", "SQM"]
                        )
                    ),
                    None,
                )
                len_col = next(
                    (
                        c
                        for c, n in col_map.items()
                        if any(k in n for k in ["BOY", "UZUNLUK", "LENG"])
                    ),
                    None,
                )
                height_col = next(
                    (
                        c
                        for c, n in col_map.items()
                        if any(k in n for k in ["EN", "YUKSEKLIK", "HEIG", "WID"])
                    ),
                    None,
                )
                slab_col = next(
                    (
                        c
                        for c, n in col_map.items()
                        if any(k in n for k in ["PLAKA", "SLAB", "NO", "SIRA"])
                    ),
                    None,
                )
                adet_col = next(
                    (
                        c
                        for c, n in col_map.items()
                        if any(k in n for k in ["ADET", "PCS", "QTY", "MIKTAR"])
                    ),
                    None,
                )
                thick_col = next(
                    (
                        c
                        for c, n in col_map.items()
                        if any(k in n for k in ["KALIN", "THICK", "CM"])
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

                    raw_l = safe_float(row.get(len_col)) if len_col else 0.0
                    raw_h = safe_float(row.get(height_col)) if height_col else 0.0
                    raw_m2 = safe_float(row.get(m2_col)) if m2_col else 0.0

                    m2_val, l_cm, h_cm = compute_slab_m2(raw_l, raw_h, raw_m2)
                    if m2_val <= 0:
                        continue

                    adet_val = (
                        int(safe_float(row.get(adet_col)))
                        if adet_col and 0 < safe_float(row.get(adet_col)) <= 50
                        else 1
                    )

                    total_m2 += m2_val
                    total_slabs += adet_val
                    if l_cm > 0 and h_cm > 0:
                        lengths.append(l_cm)
                        heights.append(h_cm)
                    valid_rows.append(row)

                # Eğer satır bazlı okuma sonuç vermediyse tüm tablodaki TOPLAM satırını ara
                if total_m2 <= 0 and m2_col:
                    col_vals = [
                        safe_float(v) for v in df[m2_col].dropna().values if safe_float(v) > 0
                    ]
                    if col_vals:
                        max_v = max(col_vals)
                        if max_v > 2.0:
                            total_m2 = max_v
                            total_slabs = max(1, len(col_vals) - 1)

                if total_m2 <= 0:
                    continue

                thick_val = "2 cm"
                if thick_col and not df[thick_col].dropna().empty:
                    raw_th = safe_float(df[thick_col].dropna().iloc[0])
                    if 17 <= raw_th <= 24 or raw_th == 2:
                        thick_val = "2 cm"
                    elif 27 <= raw_th <= 34 or raw_th == 3:
                        thick_val = "3 cm"
                    elif raw_th > 0:
                        thick_val = f"{raw_th:g} cm"

                surf_val = "Polished"
                if surf_col and not df[surf_col].dropna().empty:
                    surf_val = str(df[surf_col].dropna().iloc[0]).strip().title()

                dims_str = (
                    f"{int(sum(lengths)/len(lengths))} x {int(sum(heights)/len(heights))} cm"
                    if lengths and heights
                    else "Standard Sizes"
                )

                clean_df = pd.DataFrame(valid_rows) if valid_rows else df
                return {
                    "total_m2": total_m2,
                    "total_slabs": total_slabs,
                    "thickness": thick_val,
                    "surface": surf_val,
                    "dimensions": dims_str,
                    "raw_df": clean_df,
                    "mapped_cols": {
                        "m2": m2_col,
                        "len": len_col,
                        "height": height_col,
                        "slab": slab_col,
                        "thick": thick_col,
                        "surf": surf_col,
                    },
                }
        except Exception as e:
            logger.error(f"Excel parse hatası: {e}")
        return None

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

        if folder_id:
            excel_bytes, excel_name = self.drive.get_excel_from_folder(
                folder_id, query_hint=norm_key, max_depth=2
            )
            if excel_bytes:
                data = self.parse_packing_list_bytes(excel_bytes)

        if not data:
            data = {
                "total_m2": 0.0,
                "total_slabs": 0,
                "thickness": "2 cm",
                "surface": "Polished",
                "dimensions": "-",
                "raw_df": None,
                "mapped_cols": {},
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
        self, parent_id: str, depth: int = 0, max_depth: int = 3
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
                results.append(
                    {
                        "name": sf["name"],
                        "id": sf["id"],
                        "link": sf["link"],
                        "data": {
                            "total_m2": 0.0,
                            "total_slabs": 0,
                            "thickness": "2 cm",
                            "surface": "Polished",
                            "dimensions": "-",
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
            lot_items = self.scan_folder_recursively_for_lots(fid, max_depth=3)
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
# 5. KUSURSUZ VEKTÖREL LÜKS PDF MOTORU
# =====================================================================
class PDFPresentationEngine:
    def __init__(self, inventory: InventoryEngine, drive: DriveExplorer):
        self.inventory = inventory
        self.drive = drive

    def generate_presentation_pdf(
        self, queries: List[str], collection_folder_name: Optional[str] = None
    ) -> Tuple[bytes, str]:
        cards = []
        pdf_title = "EXECUTIVE STOCK & SLAB PRESENTATION"

        if collection_folder_name:
            pdf_title = f"{normalize_turkish_key(collection_folder_name)} - CATALOG PRESENTATION"
            parent_id, _ = self.drive.find_folder(collection_folder_name)
            if parent_id:
                scanned_items = self.inventory.scan_folder_recursively_for_lots(
                    parent_id, max_depth=2
                )
                for item in scanned_items[:25]:
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
            topMargin=36,
            bottomMargin=36,
        )
        styles = getSampleStyleSheet()
        NAVY = colors.HexColor("#14213D")
        GOLD = colors.HexColor("#C5A059")
        SLATE = colors.HexColor("#2B2D42")
        LIGHT_BG = colors.HexColor("#F8F9FA")
        BORD = colors.HexColor("#DEE2E6")

        brand_title = ParagraphStyle(
            "BrandT",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=24,
            leading=28,
            textColor=NAVY,
            alignment=1,
            spaceAfter=4,
        )
        brand_sub = ParagraphStyle(
            "BrandS",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=10,
            leading=14,
            textColor=GOLD,
            alignment=1,
            spaceAfter=6,
        )
        doc_sub = ParagraphStyle(
            "DocS",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=9,
            leading=12,
            textColor=SLATE,
            alignment=1,
            spaceAfter=14,
        )
        ch_s = ParagraphStyle(
            "CH",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=13,
            leading=17,
            textColor=NAVY,
            spaceAfter=8,
        )
        th_s = ParagraphStyle(
            "TH",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=9,
            leading=12,
            textColor=colors.whitesmoke,
        )
        tdl_s = ParagraphStyle(
            "TDL",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=9,
            leading=12,
            textColor=SLATE,
        )
        tdv_s = ParagraphStyle(
            "TDV",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=9,
            leading=12,
            textColor=SLATE,
        )

        story = [
            Paragraph("ARSLANLAR MARBLE", brand_title),
            Paragraph("NATURAL STONE & SLAB COLLECTION", brand_sub),
            Paragraph(pdf_title, doc_sub),
            HRFlowable(width="100%", thickness=1.5, color=GOLD, spaceAfter=20),
        ]

        for idx, card in enumerate(cards, start=1):
            d = card["data"]
            m2 = d.get("total_m2", 0.0)
            slabs = d.get("total_slabs", 0)

            stock_str = (
                f"{m2:,.2f} m² ({slabs} Slabs)"
                if m2 > 0
                else "Available in Showroom / Ask for Packing List"
            )
            surf_thick_str = (
                f"{d.get('surface', 'Polished')} | {d.get('thickness', '2 cm')}"
            )

            rl_img = Paragraph("Image Pending", tdv_s)
            if card.get("image_bytes"):
                try:
                    img_io = io.BytesIO(card["image_bytes"])
                    rl_img = RLImage(img_io, width=210, height=135)
                except Exception:
                    pass

            dlink = card.get("drive_link")
            link_html = (
                f"<link href='{dlink}' color='#0056b3'><u>Open Google Drive Folder</u></link>"
                if dlink
                else "Contact Sales Representative"
            )

            spec_rows = [
                [Paragraph("Specification", th_s), Paragraph("Technical Data", th_s)],
                [Paragraph("Available Stock", tdl_s), Paragraph(stock_str, tdv_s)],
                [
                    Paragraph("Surface & Thickness", tdl_s),
                    Paragraph(surf_thick_str, tdv_s),
                ],
                [Paragraph("Active Code", tdl_s), Paragraph(card["title"], tdv_s)],
                [Paragraph("High-Res Photos", tdl_s), Paragraph(link_html, tdv_s)],
            ]

            st = Table(spec_rows, colWidths=[115, 190])
            st.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                        ("BACKGROUND", (0, 1), (-1, -1), LIGHT_BG),
                        ("GRID", (0, 0), (-1, -1), 0.5, BORD),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("TOPPADDING", (0, 0), (-1, -1), 7),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                        ("LEFTPADDING", (0, 0), (-1, -1), 7),
                    ]
                )
            )

            mr = Table([[rl_img, st]], colWidths=[225, 305])
            mr.setStyle(
                TableStyle(
                    [
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("ALIGN", (0, 0), (0, 0), "CENTER"),
                    ]
                )
            )

            story.append(
                KeepTogether(
                    [Paragraph(f"{idx}. {card['title']}", ch_s), mr, Spacer(1, 22)]
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
        self, raw_df: pd.DataFrame, lot_name: str, mapped_cols: Dict[str, str]
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
        GOLD = colors.HexColor("#C5A059")
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
            textColor=GOLD,
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
            HRFlowable(width="100%", thickness=1.2, color=GOLD, spaceAfter=12),
        ]
        table_data = [
            [
                Paragraph("Row #", th_p),
                Paragraph("Slab No", th_p),
                Paragraph("Length", th_p),
                Paragraph("Height", th_p),
                Paragraph("Thick", th_p),
                Paragraph("Net M2", th_p),
            ]
        ]

        total_m2, row_count = 0.0, 0
        c_len = mapped_cols.get("len")
        c_h = mapped_cols.get("height")
        c_m2 = mapped_cols.get("m2")
        c_slab = mapped_cols.get("slab")
        c_th = mapped_cols.get("thick")

        for _, row in raw_df.iterrows():
            raw_m2 = safe_float(row.get(c_m2)) if c_m2 else 0.0
            raw_l = safe_float(row.get(c_len)) if c_len else 0.0
            raw_h = safe_float(row.get(c_h)) if c_h else 0.0
            m2_val, l_cm, h_cm = compute_slab_m2(raw_l, raw_h, raw_m2)
            if m2_val <= 0:
                continue

            row_count += 1
            total_m2 += m2_val
            table_data.append(
                [
                    Paragraph(str(row_count), td_p),
                    Paragraph(str(row.get(c_slab, row_count)), td_p),
                    Paragraph(f"{l_cm:g}" if l_cm > 0 else "-", td_p),
                    Paragraph(f"{h_cm:g}" if h_cm > 0 else "-", td_p),
                    Paragraph(str(row.get(c_th, "2 cm")), td_p),
                    Paragraph(f"{m2_val:.2f}", td_p),
                ]
            )
            if row_count >= 150:
                break

        pl_table = Table(table_data, colWidths=[60, 100, 100, 100, 80, 100])
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
                f"<b>Total Slabs:</b> {row_count} pcs | <b>Total Net Area:</b> {total_m2:,.2f} m²",
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
# 6. ARGON NLP (Niyet ve Komut Çözümleyici)
# =====================================================================
class ArgonBrain:
    def evaluate_intent(
        self, user_text: str, is_group: bool, bot_mentioned: bool
    ) -> Dict[str, Any]:
        raw = normalize_turkish_key(user_text)
        if not raw:
            return {"should_reply": False}

        lot_regex = re.compile(r"\b([A-Z]{1,4})[-\s]?(\d{1,5})\b")
        matched = [
            f"{m.group(1)}-{m.group(2)}"
            for m in lot_regex.finditer(raw)
            if m.group(1) not in ["M", "CM", "MM", "PDF", "TON", "KG"]
        ]

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
            "matched_queries": list(dict.fromkeys(matched))[:10],
            "collection_folder": collection_folder,
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
pdf_engine = PDFPresentationEngine(inventory_engine, drive_explorer)
argon_brain = ArgonBrain()

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
    bot_mentioned = bool(
        context.bot.username
        and f"@{context.bot.username.lower()}" in (msg.text or "").lower()
    )

    intent = argon_brain.evaluate_intent(msg.text or "", is_group, bot_mentioned)
    if not intent.get("should_reply"):
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
            await status.edit_text(f"Tarama uyarısı: {e}")
        return

    queries = intent.get("matched_queries", [])
    collection_folder = intent.get("collection_folder")

    if not queries and not collection_folder:
        if not is_group or bot_mentioned:
            await msg.reply_text(
                "👋 <b>Argon Master V6.2 Aktif.</b>\n"
                "• Lot sorgusu: <code>A-371</code>\n"
                "• Çoklu Tek PDF Sunum: <code>A-96 ve A-383 sunum hazırla</code>\n"
                "• Koleksiyon PDF: <code>Arslanlar Collection sunum hazırla</code>\n"
                "• Çeki Listesi: <code>A-371 çeki listesi</code>\n"
                "• Genel Stok: <code>Toplam stoğum ne kadar</code>",
                parse_mode=ParseMode.HTML,
            )
        return

    await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.TYPING)

    # 2. KOLEKSİYON BAZLI PDF SUNUMU
    if collection_folder and intent.get("want_pdf"):
        status_msg = await msg.reply_text(
            f"📁 <b>{collection_folder}</b> altındaki tüm taş ve lot klasörleri taranıp tek bir Premium PDF'te birleştiriliyor...",
            parse_mode=ParseMode.HTML,
        )
        await context.bot.send_chat_action(
            chat_id=chat_id, action=ChatAction.UPLOAD_DOCUMENT
        )
        try:
            pdf_b, pdf_n = await asyncio.to_thread(
                pdf_engine.generate_presentation_pdf,
                [],
                collection_folder_name=collection_folder,
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

    # 3. ÇOKLU LOT TEK PDF SUNUMU
    if queries and intent.get("want_pdf"):
        status_msg = await msg.reply_text(
            f"📄 Seçilen Lotlar (<b>{', '.join(queries)}</b>) tek bir Premium PDF'te birleştiriliyor...",
            parse_mode=ParseMode.HTML,
        )
        await context.bot.send_chat_action(
            chat_id=chat_id, action=ChatAction.UPLOAD_DOCUMENT
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

    # 4. STANDART LOT SORGUSU, LİNK VE ÇEKİ LİSTESİ (EXCEL + PDF)
    for q in queries:
        try:
            info = await asyncio.to_thread(inventory_engine.get_lot_info, q)
            d = info["data"]
            fid = info["folder_id"]
            flink = info["drive_link"]

            if intent.get("want_packing"):
                if info["excel_bytes"]:
                    await context.bot.send_chat_action(
                        chat_id=chat_id, action=ChatAction.UPLOAD_DOCUMENT
                    )
                    await msg.reply_document(
                        document=io.BytesIO(info["excel_bytes"]),
                        filename=info["excel_name"],
                        caption=f"📦 <b>{q}</b> Orijinal Excel Çeki Listesi",
                        parse_mode=ParseMode.HTML,
                    )
                    if d.get("raw_df") is not None and not d["raw_df"].empty:
                        try:
                            pdf_pack_bytes = await asyncio.to_thread(
                                pdf_engine.generate_packing_list_pdf,
                                d["raw_df"],
                                q,
                                d["mapped_cols"],
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
                    link_note = f"\n📂 <a href='{flink}'>Drive Klasörünü Aç</a>" if flink else ""
                    await msg.reply_text(
                        f"⚠️ <b>{q}</b> klasöründe henüz Excel çeki listesi bulunamadı.{link_note}",
                        parse_mode=ParseMode.HTML,
                    )
                continue

            m2_val = d.get("total_m2", 0.0)
            slabs_val = d.get("total_slabs", 0)
            stock_str = (
                f"{m2_val:,.2f} m² ({slabs_val} Plaka)"
                if m2_val > 0
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
                f"• Ortalama Ebat: <code>{d.get('dimensions', '-')}</code>\n"
                f"• Yüzey & Kalınlık: <code>{d.get('surface', 'Polished')} ({d.get('thickness', '2 cm')})</code>\n"
                f"• Drive Linki: {link_html}"
            )

            # Eğer kullanıcı sadece "drive linki lütfen" dediyse fotoğrafları indirmekle vakit kaybetmeden anında yanıtla
            if intent.get("want_only_link"):
                await msg.reply_text(text_response, parse_mode=ParseMode.HTML)
                continue

            img1, img2 = await asyncio.to_thread(
                drive_explorer.get_images_structured, fid, q
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

    application = ApplicationBuilder().token(TELEGRAM_TOKEN).build()

    application.add_handler(
        MessageHandler(
            (filters.TEXT | filters.PHOTO) & (~filters.COMMAND), handle_message
        )
    )

    logger.info("🚀 Argon Master V6.2 Thread-Safe başlatılıyor...")
    application.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
