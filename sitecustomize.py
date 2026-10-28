# -*- coding: utf-8 -*-
"""
ARSLANLAR MARBLE - V22 MASTER ARCHITECTURE (sitecustomize.py)
1. 'custom' klasorundeki fotograflari ASLA kullanmaz (PDF, Telegram ve Render'da tamamen yasaklandi).
2. Her lotun icindeki Excel (.xlsx / .xls / Google Sheet) dosyasini derinlemesine okuyarak
   m2, Plaka (Slabs), Ortalama Ebat (Avg. Dimensions) ve Kalinlik bilgilerini %100 cikarir.
3. Sunumda (PDF) ve Telegram listesinde m2, plaka, ebat ve kalinlik eksiksiz yer alir.
4. Render motoru arka plan/sehpa/fabrika goruntusunu otomatik kirpar (Auto-Slab Crop),
   saf mermer dokusundan kusursuz Bookmatch & Mimari Ic Mekan Render uretir.
5. 'stok tablosu' veya 'detayli stok' komutuyla 3 koleksiyonun Aktif (In Stock) ve
   Satilan (Sold Out) tum lotlarini iceren Premium PDF + Excel Master Stok Tablosu uretir.
6. bot.py dosyasina %0 mudahale eder, mevcut calisan hicbir ozelligi bozmaz.
"""

import os
import re
import gc
import io
import sys
import json
import time
import base64
import asyncio
import logging
import threading
import subprocess
from datetime import datetime, timezone, timedelta
from typing import Dict, List, Any, Optional, Tuple
from concurrent.futures import ThreadPoolExecutor

import requests
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from PIL import Image, ImageDraw, ImageFilter, ImageEnhance, ImageOps

from google.oauth2 import service_account
import google.auth.transport.requests

from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors

import telegram
from telegram import Update, InputMediaPhoto
from telegram.ext import (
    Application,
    MessageHandler,
    ContextTypes,
    ApplicationHandlerStop,
    filters,
)

logger = logging.getLogger("ArgonMaster_V17.Plugin")

STATE_FILE = "arslanlar_watcher_state.json"
BOT_CACHE_FILE = "argon_v17_master_cache.json"

TARGET_COLLECTIONS = [
    "Arslanlar Series",
    "Arslanlar Collection",
    "Arslanlar Exclusive Collection",
]

INTERNAL_SUBFOLDER_NAMES = {
    "custom", "email", "mail", "photos", "photo", "foto", "fotograf", "fotograflar",
    "images", "image", "img", "slabs", "slab", "plaka", "plakalar", "raw", "original",
    "orijinal", "video", "videos", "pdf", "pdfs", "katalog", "catalog", "whatsapp",
    "web", "social", "instagram", "render", "renders", "bookmatch", "excel", "list",
    "packing", "etiket", "label", "qr", "temp", "tmp", "backup", "arsiv", "archive",
    "secilenler", "selected", "musteri", "client"
}

RENDER_EXECUTOR = ThreadPoolExecutor(max_workers=1, thread_name_prefix="RAMSafeWorker")
TR_TZ = timezone(timedelta(hours=3))


def normalize_compare(s: str) -> str:
    if not s:
        return ""
    s = s.replace("İ", "i").replace("I", "i").replace("ı", "i").lower()
    tr_map = str.maketrans("çğöşü", "cgosu")
    s = s.translate(tr_map)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def extract_stone_query_name(text: str) -> str:
    t = text.strip()
    t_low = t.replace("İ", "i").replace("I", "ı").lower()

    patterns_to_strip = [
        r"^ne\s+kadar\s+",
        r"^kaç\s+m2\s+",
        r"^kac\s+m2\s+",
        r"^kaç\s+metrekare\s+",
        r"^kac\s+metrekare\s+",
        r"^elimde\s+ne\s+kadar\s+",
        r"^elimizde\s+ne\s+kadar\s+",
        r"^stokta\s+ne\s+kadar\s+",
        r"^stoklarda\s+ne\s+kadar\s+",
        r"^bana\s+",
        r"^lütfen\s+",
        r"^lutfen\s+",
        r"^how\s+much\s+",
        r"^how\s+many\s+",
        r"\s+sunumu?\s+lütfen.*$",
        r"\s+sunumu?\s+lutfen.*$",
        r"\s+sunumu?\s+hazırla.*$",
        r"\s+sunumu?\s+yap.*$",
        r"\s+sunumu?.*$",
        r"\s+presentation\s*.*$",
        r"\s+kataloğu?\s*.*$",
        r"\s+katalog\s*.*$",
        r"\s+catalog\s*.*$",
        r"\s+pdf\s*.*$",
        r"\s+çeşitleri\s*.*$",
        r"\s+cesitleri\s*.*$",
        r"\s+varieties\s*.*$",
        r"\s+nelerdir\s*.*$",
        r"\s+neler\s+var.*$",
        r"\s+neler\s*.*$",
        r"\s+var\s+mı.*$",
        r"\s+varmı.*$",
        r"\s+var\s*\??$",
        r"\s+mevcut\s+mu.*$",
        r"\s+mevcut\s*$",
        r"\s+stok\s+listesi.*$",
        r"\s+stokları.*$",
        r"\s+stoklari.*$",
        r"\s+stok\s*$",
        r"\s+stock\s*$",
        r"\s+lütfen\s*$",
        r"\s+lutfen\s*$",
        r"\s+please\s*$",
    ]
    for pat in patterns_to_strip:
        t_low = re.sub(pat, "", t_low, flags=re.IGNORECASE).strip()

    words = t_low.split()
    cleaned_words = []
    for w in words:
        w_clean = re.sub(r"(ları|leri|lar|ler|dan|den|tan|ten|nın|nin|nun|nün)$", "", w)
        if len(w_clean) >= 2:
            cleaned_words.append(w_clean)
        else:
            cleaned_words.append(w)

    return " ".join(cleaned_words).strip(" ?.,!:-")


def is_image_file_obj(file_obj: Dict[str, Any]) -> bool:
    mt = str(file_obj.get("mimeType", "")).lower()
    nm = str(file_obj.get("name", "")).lower()
    return mt.startswith("image/") or nm.endswith((".jpg", ".jpeg", ".png", ".webp"))


def to_float_safe(val: Any) -> Optional[float]:
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return None
    # '1.250,45' -> '1250.45' veya '3,64' -> '3.64'
    s = re.sub(r"[^\d,.-]", "", s)
    if not s or s in ("-", ".", ","):
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts) == 2 and len(parts[1]) in (1, 2, 3):
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    try:
        return float(s)
    except Exception:
        return None


class PersistentWatcherState:
    def __init__(self, filepath: str = STATE_FILE):
        self.filepath = filepath
        self.lock = threading.Lock()
        self.data: Dict[str, Any] = {
            "chat_ids": [],
            "initialized": False,
            "active_lots": {},
            "sold_out_lots": {},
            "metrics_cache": {},
            "last_summary_report_key": "",
            "updated_at": "",
        }
        self.load()

    def load(self):
        with self.lock:
            if os.path.exists(self.filepath):
                try:
                    with open(self.filepath, "r", encoding="utf-8") as f:
                        loaded = json.load(f)
                        if isinstance(loaded, dict):
                            self.data.update(loaded)
                except Exception as e:
                    logger.warning(f"State load warning: {e}")

            env_chat = os.getenv("ADMIN_CHAT_ID", "") or os.getenv("REPORT_CHAT_ID", "")
            if env_chat.strip():
                for cid in env_chat.split(","):
                    cid = cid.strip()
                    if cid.lstrip("-").isdigit():
                        val = int(cid)
                        if val not in self.data["chat_ids"]:
                            self.data["chat_ids"].append(val)

    def save(self, git_sync: bool = False):
        with self.lock:
            self.data["updated_at"] = datetime.now(TR_TZ).isoformat()
            try:
                with open(self.filepath, "w", encoding="utf-8") as f:
                    json.dump(self.data, f, ensure_ascii=False, indent=2)
            except Exception as e:
                logger.warning(f"State save warning: {e}")

        if git_sync:
            threading.Thread(target=self.git_persist_async, daemon=True).start()

    def git_persist_async(self):
        try:
            subprocess.run(["git", "config", "user.name", "ArslanlarBot"], timeout=5, check=False)
            subprocess.run(["git", "config", "user.email", "bot@arslanlarmarble.com"], timeout=5, check=False)
            subprocess.run(["git", "add", self.filepath], timeout=5, check=False)
            res = subprocess.run(
                ["git", "commit", "-m", "chore: auto-sync watcher state [skip ci]"],
                capture_output=True,
                text=True,
                timeout=10,
                check=False,
            )
            if res.returncode == 0:
                subprocess.run(["git", "push"], timeout=15, check=False)
        except Exception:
            pass

    def register_chat(self, chat_id: int) -> bool:
        if chat_id not in self.data["chat_ids"]:
            self.data["chat_ids"].append(chat_id)
            self.save(git_sync=True)
            return True
        return False


class ArslanlarDriveIntelligence:
    LOT_CODE_REGEX = re.compile(r"\b([A-Z]{1,3}\s*[-_]\s*\d{1,5}[A-Z]?|[A-Z]{1,2}\d{2,5}[A-Z]?)\b", re.IGNORECASE)
    SOLD_KEYWORD_REGEX = re.compile(r"\b(sold\s*out|soldout|sold|satildi|satıldı|pasif)\b", re.IGNORECASE)

    def __init__(self, state: PersistentWatcherState):
        self.state = state
        self.sa_json_str = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
        self.creds = None
        self.token_lock = threading.Lock()
        self.session = requests.Session()
        self.inventory_lock = threading.Lock()
        self.active_lots: Dict[str, Dict[str, Any]] = {}
        self.sold_out_lots: Dict[str, Dict[str, Any]] = {}
        self.last_scan_time: float = 0.0
        self.cached_logo_bytes: Optional[bytes] = None
        self.init_credentials()

    def init_credentials(self):
        if not self.sa_json_str:
            return
        try:
            info = json.loads(self.sa_json_str)
            self.creds = service_account.Credentials.from_service_account_info(
                info, scopes=["https://www.googleapis.com/auth/drive.readonly"]
            )
        except Exception as e:
            logger.error(f"Plugin Drive Auth Error: {e}")

    def get_headers(self) -> Optional[Dict[str, str]]:
        if not self.creds:
            return None
        with self.token_lock:
            if not self.creds.valid or self.creds.expired or not self.creds.token:
                req = google.auth.transport.requests.Request()
                self.creds.refresh(req)
            return {"Authorization": f"Bearer {self.creds.token}"}

    def get_arslanlar_logo_bytes(self) -> Optional[bytes]:
        if self.cached_logo_bytes:
            return self.cached_logo_bytes

        try:
            main_mod = sys.modules.get("__main__")
            if main_mod:
                for attr_name in dir(main_mod):
                    low_attr = attr_name.lower()
                    if "logo" in low_attr or "b64" in low_attr or "icon" in low_attr or "header" in low_attr:
                        val = getattr(main_mod, attr_name, None)
                        if isinstance(val, bytes) and len(val) > 200:
                            self.cached_logo_bytes = val
                            return val
                        if isinstance(val, str) and len(val) > 250:
                            clean_b64 = re.sub(r"^data:image/[a-zA-Z]+;base64,", "", val.strip())
                            clean_b64 = re.sub(r"\s+", "", clean_b64)
                            try:
                                decoded = base64.b64decode(clean_b64)
                                with Image.open(io.BytesIO(decoded)) as im:
                                    im.verify()
                                self.cached_logo_bytes = decoded
                                return decoded
                            except Exception:
                                pass
        except Exception:
            pass

        for fname in ("logo.png", "logo.jpg", "arslanlar_logo.png", "header_logo.png"):
            if os.path.exists(fname):
                try:
                    with open(fname, "rb") as f:
                        self.cached_logo_bytes = f.read()
                        return self.cached_logo_bytes
                except Exception:
                    pass

        try:
            w, h = 640, 160
            img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            copper_gold = (196, 132, 78, 255)
            light_gold = (228, 178, 128, 255)

            draw.pieslice([25, 25, 135, 135], start=135, end=315, fill=copper_gold)
            draw.polygon([(55, 130), (95, 55), (145, 55), (105, 130)], fill=light_gold)
            draw.rectangle([75, 98, 145, 112], fill=copper_gold)
            draw.rectangle([45, 122, 155, 132], fill=copper_gold)

            out = io.BytesIO()
            img.save(out, format="PNG")
            self.cached_logo_bytes = out.getvalue()
            return self.cached_logo_bytes
        except Exception:
            return None

    def find_bot_drive_explorer(self) -> Optional[Any]:
        try:
            for obj in gc.get_objects():
                if type(obj).__name__ == "DriveExplorer":
                    return obj
        except Exception:
            pass
        return None

    def load_bot_existing_caches(self) -> Dict[str, Dict[str, Any]]:
        combined_meta: Dict[str, Dict[str, Any]] = {}

        def extract_entries(d: Any):
            if not isinstance(d, dict):
                return
            for k, v in d.items():
                if isinstance(v, dict):
                    has_metric = any(
                        mk in v
                        for mk in (
                            "m2", "total_m2", "sqm", "area", "plaka", "slabs",
                            "slab_count", "count", "ebat", "avg_dim", "dimensions",
                            "mevcut_stok", "ortalama_ebat"
                        )
                    )
                    if has_metric:
                        combined_meta[str(k).upper().strip()] = v
                        fid = v.get("folder_id") or v.get("id")
                        if fid:
                            combined_meta[str(fid).strip()] = v
                        code = v.get("code") or v.get("lot_code") or v.get("lot")
                        if code:
                            combined_meta[str(code).upper().strip()] = v
                    else:
                        extract_entries(v)
                elif isinstance(v, list):
                    for item in v:
                        if isinstance(item, dict):
                            fid = item.get("folder_id") or item.get("id")
                            code = item.get("code") or item.get("lot_code")
                            if fid:
                                combined_meta[str(fid).strip()] = item
                            if code:
                                combined_meta[str(code).upper().strip()] = item

        if os.path.exists(BOT_CACHE_FILE):
            try:
                with open(BOT_CACHE_FILE, "r", encoding="utf-8") as f:
                    extract_entries(json.load(f))
            except Exception:
                pass

        explorer = self.find_bot_drive_explorer()
        if explorer is not None:
            try:
                for attr_val in list(explorer.__dict__.values()):
                    if isinstance(attr_val, dict):
                        extract_entries(attr_val)
            except Exception:
                pass

        return combined_meta

    def list_all_drive_items(self) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        headers = self.get_headers()
        if not headers:
            return [], []

        all_items: List[Dict[str, Any]] = []
        page_token = None
        url = "https://www.googleapis.com/drive/v3/files"

        while True:
            params = {
                "q": "trashed = false",
                "pageSize": 1000,
                "fields": "nextPageToken, files(id, name, mimeType, parents, webViewLink, modifiedTime, size, description)",
                "supportsAllDrives": "true",
                "includeItemsFromAllDrives": "true",
            }
            if page_token:
                params["pageToken"] = page_token

            try:
                resp = self.session.get(url, headers=headers, params=params, timeout=30)
                if resp.status_code != 200:
                    break
                data = resp.json()
                all_items.extend(data.get("files", []))
                page_token = data.get("nextPageToken")
                if not page_token:
                    break
            except Exception as e:
                logger.warning(f"Drive listing error: {e}")
                break

        folders = [x for x in all_items if x.get("mimeType") == "application/vnd.google-apps.folder"]
        files = [x for x in all_items if x.get("mimeType") != "application/vnd.google-apps.folder"]
        return folders, files

    def download_file_bytes(self, file_id: str, mime_type: str = "") -> Optional[bytes]:
        headers = self.get_headers()
        if not headers:
            return None
        if mime_type == "application/vnd.google-apps.spreadsheet":
            url = f"https://www.googleapis.com/drive/v3/files/{file_id}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        else:
            url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media&supportsAllDrives=true"
        try:
            resp = self.session.get(url, headers=headers, timeout=25)
            if resp.status_code == 200:
                return resp.content
        except Exception:
            pass
        return None

    def parse_excel_metrics_deep(self, file_bytes: bytes, file_name: str = "") -> Optional[Dict[str, Any]]:
        """
        Her lot klasörünün içindeki Excel (.xlsx / Google Sheet / .xls) dosyasını
        sütun başlıklarına (EN, BOY, M2, AREA, KALINLIK, TOPLAM) ve plaka satırlarına göre
        derinlemesine analiz eder. Virgüllü/noktalı tüm formatları çözer.
        """
        try:
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        except Exception:
            return None

        best_m2 = 0.0
        best_slabs = 0
        best_avg_dim = "-"
        best_thickness = "2 cm"

        for sheet in wb.worksheets:
            rows = list(sheet.iter_rows(values_only=True))
            if not rows:
                continue

            col_w_idx = None
            col_h_idx = None
            col_th_idx = None
            col_m2_idx = None
            header_row_idx = None

            # 1. Sütun başlıklarını bul (İlk 20 satırda ara)
            for r_i, r in enumerate(rows[:20]):
                norm_cells = [normalize_compare(str(c)) if c is not None else "" for c in r]
                w_cand = None
                h_cand = None
                m2_cand = None
                th_cand = None

                for c_i, nc in enumerate(norm_cells):
                    if not nc:
                        continue
                    if nc in ("boy", "length", "uzunluk", "len", "l", "boy cm", "length cm", "uzunluk cm"):
                        w_cand = c_i
                    elif nc in ("en", "width", "yukseklik", "height", "genislik", "w", "h", "en cm", "width cm"):
                        h_cand = c_i
                    elif any(k in nc for k in ("m2", "m²", "sqm", "alan", "area", "metraj", "miktar", "quantity")):
                        m2_cand = c_i
                    elif any(k in nc for k in ("kalinlik", "thickness", "thick", "kal")):
                        th_cand = c_i

                if (w_cand is not None and h_cand is not None) or m2_cand is not None:
                    col_w_idx = w_cand
                    col_h_idx = h_cand
                    col_m2_idx = m2_cand
                    col_th_idx = th_cand
                    header_row_idx = r_i
                    break

            widths: List[float] = []
            heights: List[float] = []
            row_areas: List[float] = []
            summary_totals: List[float] = []
            slab_rows_count = 0

            start_row = (header_row_idx + 1) if header_row_idx is not None else 0

            for r in rows[start_row:]:
                raw_strs = [str(c).strip() if c is not None else "" for c in r]
                joined_low = normalize_compare(" ".join(raw_strs))
                if not joined_low:
                    continue

                is_total_row = any(
                    kw in joined_low
                    for kw in ("toplam", "total", "genel toplam", "grand total", "sum", "yekun", "alt toplam")
                )

                if is_total_row:
                    # Toplam satırındaki m2 değerini yakala
                    if col_m2_idx is not None and col_m2_idx < len(r):
                        v_tot = to_float_safe(r[col_m2_idx])
                        if v_tot and 1.0 <= v_tot <= 5000.0:
                            summary_totals.append(v_tot)
                    for c in r:
                        v_num = to_float_safe(c)
                        if v_num and 2.5 <= v_num <= 5000.0:
                            # Eğer bu sayı plaka adedi değilse (ondalıklıysa veya satırdaki en büyük makul m2 ise)
                            if isinstance(c, float) or ("," in str(c)) or ("." in str(c)) or v_num > 15.0:
                                summary_totals.append(v_num)
                    continue

                # Başlıklı sütunlar bulunduysa doğrudan ilgili sütunlardan oku
                w_val = None
                h_val = None
                a_val = None

                if col_w_idx is not None and col_h_idx is not None and col_w_idx < len(r) and col_h_idx < len(r):
                    w_val = to_float_safe(r[col_w_idx])
                    h_val = to_float_safe(r[col_h_idx])
                    # Metre cinsinden yazılmışsa (örn: 3.14 x 1.70) cm'ye çevir
                    if w_val and h_val and 0.5 <= w_val <= 4.5 and 0.5 <= h_val <= 4.5:
                        w_val *= 100.0
                        h_val *= 100.0

                if col_m2_idx is not None and col_m2_idx < len(r):
                    a_val = to_float_safe(r[col_m2_idx])

                if col_th_idx is not None and col_th_idx < len(r):
                    th_v = to_float_safe(r[col_th_idx])
                    if th_v in (2.0, 20.0):
                        best_thickness = "2 cm"
                    elif th_v in (3.0, 30.0):
                        best_thickness = "3 cm"

                # Eğer sütun başlığı standart değilse satırdaki sayıları akıllıca çözümle
                if w_val is None or h_val is None:
                    nums = [to_float_safe(c) for c in r]
                    nums = [x for x in nums if x is not None and x > 0]
                    cm_cands = [x for x in nums if 55.0 <= x <= 440.0]
                    m_cands = [x * 100.0 for x in nums if 0.65 <= x <= 4.20]

                    if len(cm_cands) >= 2:
                        w_val, h_val = cm_cands[0], cm_cands[1]
                    elif len(m_cands) >= 2 and len(nums) >= 3:
                        w_val, h_val = m_cands[0], m_cands[1]

                    if a_val is None and w_val and h_val:
                        calc_a = (w_val * h_val) / 10000.0
                        area_cands = [x for x in nums if 0.5 <= x <= 14.0 and abs(x - calc_a) <= 1.5]
                        if area_cands:
                            a_val = area_cands[-1]
                        else:
                            a_val = calc_a

                if w_val and h_val and 45.0 <= w_val <= 450.0 and 45.0 <= h_val <= 450.0:
                    widths.append(w_val)
                    heights.append(h_val)
                    slab_rows_count += 1
                    if a_val and 0.4 <= a_val <= 15.0:
                        row_areas.append(a_val)
                    else:
                        row_areas.append((w_val * h_val) / 10000.0)
                elif a_val and 0.5 <= a_val <= 15.0:
                    row_areas.append(a_val)
                    slab_rows_count += 1

            sheet_m2 = 0.0
            if row_areas:
                sheet_m2 = sum(row_areas)
            if summary_totals:
                max_tot = max(summary_totals)
                if sheet_m2 <= 0 or abs(max_tot - sheet_m2) < (sheet_m2 * 0.35):
                    sheet_m2 = max_tot

            if sheet_m2 > best_m2:
                best_m2 = sheet_m2
                best_slabs = slab_rows_count if slab_rows_count > 0 else best_slabs
                if widths and heights:
                    avg_w = int(round(sum(widths) / len(widths)))
                    avg_h = int(round(sum(heights) / len(heights)))
                    best_avg_dim = f"{avg_w} x {avg_h} cm"

        if best_m2 > 0 or best_slabs > 0:
            return {
                "m2": round(best_m2, 2),
                "slabs": best_slabs,
                "avg_dim": best_avg_dim,
                "surface": f"Polished ({best_thickness})",
            }
        return None

    def call_bot_native_folder_analyzer(self, folder_id: str, folder_name: str) -> Optional[Dict[str, Any]]:
        """
        bot.py içindeki DriveExplorer nesnesinin kendi klasör/excel analiz metodunu
        doğrudan çağırır (K-3179 ve K-3968'i hesaplayan orijinal motor).
        """
        explorer = self.find_bot_drive_explorer()
        if not explorer:
            return None

        try:
            for attr_name in dir(explorer):
                if attr_name.startswith("__"):
                    continue
                fn = getattr(explorer, attr_name, None)
                if not callable(fn):
                    continue
                low_name = attr_name.lower()
                if any(k in low_name for k in ("meta", "metric", "analyze", "folder", "lot", "excel", "detail", "calc")):
                    for args in ((folder_id,), (folder_id, folder_name)):
                        try:
                            res = fn(*args)
                            if isinstance(res, dict):
                                m2_v = to_float_safe(
                                    res.get("m2") or res.get("total_m2") or res.get("sqm") or res.get("area")
                                )
                                if m2_v and m2_v > 0:
                                    return res
                        except Exception:
                            pass
        except Exception:
            pass
        return None

    def compute_lot_metrics(
        self,
        folder_id: str,
        lot_code: str,
        folder_name: str,
        modified_time: str,
        all_lot_files: List[Dict[str, Any]],
        non_custom_images: List[Dict[str, Any]],
        custom_images: List[Dict[str, Any]],
        bot_cache: Dict[str, Dict[str, Any]],
    ) -> Dict[str, Any]:
        m2 = 0.0
        slabs = 0
        avg_dim = "-"
        surface = "Polished (2 cm)"

        # 1. bot.py önbelleğinden kontrol et
        for key in (folder_id, lot_code.upper(), lot_code.replace("-", "").upper(), folder_name.upper()):
            c = bot_cache.get(key)
            if isinstance(c, dict):
                for k_m2 in ("m2", "total_m2", "sqm", "area", "total_area", "mevcut_stok"):
                    v_f = to_float_safe(c.get(k_m2))
                    if v_f and v_f > 0:
                        m2 = v_f
                        break
                for k_pl in ("plaka", "slabs", "slab_count", "count", "pcs", "plaka_sayisi"):
                    v_i = to_float_safe(c.get(k_pl))
                    if v_i and v_i > 0:
                        slabs = int(v_i)
                        break
                for k_d in ("avg_dim", "ebat", "ortalama_ebat", "dimensions", "dim"):
                    val = c.get(k_d)
                    if val and str(val).strip() not in ("", "-", "None", "N/A"):
                        avg_dim = str(val).strip()
                        break
                for k_s in ("surface_thickness", "surface", "finish", "yuzey"):
                    val = c.get(k_s)
                    if val and str(val).strip() not in ("", "-", "None"):
                        surface = str(val).strip()
                        break

        # 2. Kalıcı state önbelleğinden kontrol et (Sadece m2 > 0 ise kabul et!)
        persisted = self.state.data.setdefault("metrics_cache", {}).get(folder_id)
        if m2 <= 0 and isinstance(persisted, dict) and persisted.get("m2", 0) > 0:
            m2 = float(persisted["m2"])
            if persisted.get("slabs", 0) > 0:
                slabs = int(persisted["slabs"])
            if persisted.get("avg_dim", "-") != "-":
                avg_dim = persisted["avg_dim"]
            if persisted.get("surface"):
                surface = persisted["surface"]

        # 3. Lotun içindeki Excel (.xlsx, .xls, Google Sheets) dosyalarını indir ve derinlemesine oku!
        if m2 <= 0 or avg_dim == "-":
            excel_files = [
                f for f in all_lot_files
                if f.get("mimeType") in (
                    "application/vnd.google-apps.spreadsheet",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "application/vnd.ms-excel",
                )
                or str(f.get("name", "")).lower().endswith((".xlsx", ".xls", ".xlsm"))
            ]
            for ef in excel_files:
                eb = self.download_file_bytes(ef["id"], ef.get("mimeType", ""))
                if eb:
                    parsed = self.parse_excel_metrics_deep(eb, ef.get("name", ""))
                    if parsed and parsed.get("m2", 0) > 0:
                        m2 = parsed["m2"]
                        if parsed.get("slabs", 0) > 0:
                            slabs = parsed["slabs"]
                        if parsed.get("avg_dim", "-") != "-":
                            avg_dim = parsed["avg_dim"]
                        if parsed.get("surface"):
                            surface = parsed["surface"]
                        break

        # 4. bot.py'nin kendi DriveExplorer analizörünü çağır
        if m2 <= 0:
            native_res = self.call_bot_native_folder_analyzer(folder_id, folder_name)
            if isinstance(native_res, dict):
                v_m2 = to_float_safe(native_res.get("m2") or native_res.get("total_m2") or native_res.get("sqm"))
                if v_m2 and v_m2 > 0:
                    m2 = v_m2
                v_sl = to_float_safe(native_res.get("slabs") or native_res.get("plaka") or native_res.get("count"))
                if v_sl and v_sl > 0:
                    slabs = int(v_sl)
                v_dm = native_res.get("avg_dim") or native_res.get("ebat")
                if v_dm and str(v_dm).strip() not in ("", "-", "None"):
                    avg_dim = str(v_dm).strip()

        # 5. Dosya adlarındaki ebat ve metraj kalıplarını oku (Örn: "314 x 170 x 2 cm - Area: 5,34 m2")
        widths_from_names: List[int] = []
        heights_from_names: List[int] = []
        areas_from_names: List[float] = []

        for f in all_lot_files:
            fn = f.get("name", "") + " " + str(f.get("description", ""))
            m_dim = re.search(r"\b(\d{2,3})\s*[xX*]\s*(\d{2,3})(?:\s*[xX*]\s*(\d(?:[.,]\d)?))?\s*(?:cm|mm)?\b", fn)
            if m_dim:
                w_v, h_v = int(m_dim.group(1)), int(m_dim.group(2))
                if 50 <= w_v <= 450 and 50 <= h_v <= 450:
                    widths_from_names.append(w_v)
                    heights_from_names.append(h_v)
                if m_dim.group(3):
                    th = m_dim.group(3).replace(",", ".")
                    surface = f"Polished ({th} cm)"

            m_ar = re.search(r"(?:area\s*[:=]?\s*)?(\d+[.,]\d+|\d+)\s*(?:m2|m²|sqm)", fn, flags=re.I)
            if m_ar:
                a_v = to_float_safe(m_ar.group(1))
                if a_v and 0.5 <= a_v <= 2000.0:
                    areas_from_names.append(a_v)

        if slabs <= 0:
            m_between = re.search(
                r"\b(?:between\s+)?(\d{1,3})\s*[-_\s]\s*(\d{1,3})\s*(?:stock|stok)?\b", folder_name, flags=re.I
            )
            if m_between:
                b1, b2 = int(m_between.group(1)), int(m_between.group(2))
                if 1 <= b1 < b2 <= 160:
                    slabs = (b2 - b1) + 1

        if slabs <= 0:
            if non_custom_images:
                slabs = len(non_custom_images)
            elif custom_images:
                slabs = len(custom_images)

        if avg_dim == "-" and widths_from_names and heights_from_names:
            avg_w = int(round(sum(widths_from_names) / len(widths_from_names)))
            avg_h = int(round(sum(heights_from_names) / len(heights_from_names)))
            avg_dim = f"{avg_w} x {avg_h} cm"

        if m2 <= 0 and areas_from_names:
            if len(areas_from_names) > 1 and all(a < 15.0 for a in areas_from_names):
                if slabs > len(areas_from_names):
                    m2 = (sum(areas_from_names) / len(areas_from_names)) * slabs
                else:
                    m2 = sum(areas_from_names)
            else:
                single_a = max(areas_from_names)
                if single_a < 12.0 and slabs > 1:
                    m2 = single_a * slabs
                else:
                    m2 = single_a

        if m2 <= 0 and avg_dim != "-" and slabs > 0:
            m_d = re.search(r"(\d+)\s*x\s*(\d+)", avg_dim)
            if m_d:
                m2 = ((int(m_d.group(1)) * int(m_d.group(2))) / 10000.0) * slabs

        result = {
            "m2": round(m2, 2),
            "slabs": slabs,
            "avg_dim": avg_dim,
            "surface": surface,
            "modified_time": modified_time,
        }
        # Sadece m2 > 0 bulunduğunda kalıcı önbelleğe yaz ki 0.00 m2 kilitlenmesin!
        if m2 > 0:
            self.state.data.setdefault("metrics_cache", {})[folder_id] = result

        return result

    @classmethod
    def parse_lot_folder_title(cls, folder_name: str, parent_names: List[str]) -> Tuple[str, str]:
        raw = folder_name.strip()
        m = cls.LOT_CODE_REGEX.search(raw)
        lot_code = ""
        if m:
            lot_code = re.sub(r"\s+", "", m.group(1)).upper()
            if "-" not in lot_code:
                lot_code = re.sub(r"^([A-Z]+)(\d+.*)$", r"\1-\2", lot_code)

        stone_part = raw
        if m:
            stone_part = stone_part.replace(m.group(0), " ")

        stone_part = re.sub(
            r"\b(full\s+lot|full|lot|stock|stok|between\s*\d+\s*[-_]?\s*\d*|between|\d+\s+\d+\s+stock|\d+([.,]\d+)?\s*(m2|m²|plaka|slabs?|cm|mm|pcs)|sold\s*out|sold)\b",
            " ",
            stone_part,
            flags=re.IGNORECASE,
        )
        stone_part = re.sub(r"[-_()\[\]/\\|]+", " ", stone_part)
        stone_part = re.sub(r"\s+", " ", stone_part).strip()

        if len(stone_part) < 2:
            for pname in reversed(parent_names):
                p_norm = normalize_compare(pname)
                if p_norm not in (
                    "arslanlar series",
                    "arslanlar collection",
                    "arslanlar exclusive collection",
                    "sold out",
                    "soldout",
                ) and p_norm not in INTERNAL_SUBFOLDER_NAMES:
                    stone_part = pname.strip()
                    break

        if not lot_code:
            lot_code = raw.upper()
        if not stone_part:
            stone_part = raw

        return lot_code, stone_part.title()

    def scan_drive_collections(self, force: bool = False) -> Tuple[Dict[str, Dict[str, Any]], Dict[str, Dict[str, Any]]]:
        now = time.time()
        if not force and self.active_lots and (now - self.last_scan_time < 45):
            return self.active_lots, self.sold_out_lots

        folders, files = self.list_all_drive_items()
        if not folders:
            return self.active_lots, self.sold_out_lots

        bot_cache = self.load_bot_existing_caches()

        folder_by_id: Dict[str, Dict[str, Any]] = {f["id"]: f for f in folders if "id" in f}
        children_folders: Dict[str, List[Dict[str, Any]]] = {}
        children_files: Dict[str, List[Dict[str, Any]]] = {}

        for f in folders:
            for p in f.get("parents", []):
                children_folders.setdefault(p, []).append(f)

        for fl in files:
            for p in fl.get("parents", []):
                children_files.setdefault(p, []).append(fl)

        target_norm_map = {normalize_compare(c): c for c in TARGET_COLLECTIONS}
        root_collection_folders: List[Tuple[str, str]] = []

        for f in folders:
            fn_norm = normalize_compare(f.get("name", ""))
            if fn_norm in target_norm_map:
                root_collection_folders.append((f["id"], target_norm_map[fn_norm]))

        new_active: Dict[str, Dict[str, Any]] = {}
        new_sold_out: Dict[str, Dict[str, Any]] = {}

        def collect_all_descendant_files(fid: str) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, Any]]]:
            """
            KESİN KURAL 1: 'custom' klasöründeki fotoğrafları GÖRSEL LİSTESİNE ASLA EKLEME!
            Görselleri sadece lotun ana klasöründen veya 'email'/'photos' gibi custom olmayan klasörlerden al.
            (Ancak 'custom' içindeki dosya isimlerinde metraj/ebat yazabileceği için metin analizinde kullan).
            """
            all_f: List[Dict[str, Any]] = list(children_files.get(fid, []))
            direct_imgs: List[Dict[str, Any]] = []
            sub_non_custom_imgs: List[Dict[str, Any]] = []
            custom_imgs_for_count_only: List[Dict[str, Any]] = []

            for x in children_files.get(fid, []):
                if is_image_file_obj(x):
                    direct_imgs.append(x)

            stack = [(sf, normalize_compare(sf.get("name", ""))) for sf in children_folders.get(fid, [])]
            while stack:
                sub_f, sub_norm = stack.pop(0)
                sub_id = sub_f["id"]
                s_files = children_files.get(sub_id, [])
                all_f.extend(s_files)
                is_custom_dir = "custom" in sub_norm.split() or "custom" in sub_norm

                for x in s_files:
                    if is_image_file_obj(x):
                        if is_custom_dir:
                            custom_imgs_for_count_only.append(x)
                        else:
                            sub_non_custom_imgs.append(x)

                for deeper_sf in children_folders.get(sub_id, []):
                    deeper_norm = normalize_compare(deeper_sf.get("name", ""))
                    stack.append((deeper_sf, f"{sub_norm} {deeper_norm}"))

            direct_imgs.sort(key=lambda item: item.get("name", ""))
            sub_non_custom_imgs.sort(key=lambda item: item.get("name", ""))
            # SADECE custom OLMAYAN görseller!
            pure_images = direct_imgs + sub_non_custom_imgs
            return all_f, pure_images, custom_imgs_for_count_only

        def walk_hierarchy(current_fid: str, collection_name: str, path_names: List[str], in_sold_out: bool):
            cur_folder = folder_by_id.get(current_fid, {})
            cur_name = cur_folder.get("name", "").strip()
            cur_norm = normalize_compare(cur_name)

            if cur_norm in INTERNAL_SUBFOLDER_NAMES:
                return

            is_sold_container = cur_norm in ("sold out", "soldout", "satildi", "pasif")
            has_sold_in_title = bool(self.SOLD_KEYWORD_REGEX.search(cur_name))
            now_sold_out = in_sold_out or is_sold_container or has_sold_in_title

            subfolders = children_folders.get(current_fid, [])
            non_internal_subfolders = [
                sf for sf in subfolders
                if normalize_compare(sf.get("name", "")) not in INTERNAL_SUBFOLDER_NAMES
                and normalize_compare(sf.get("name", "")) not in ("sold out", "soldout")
            ]

            has_lot_code = bool(self.LOT_CODE_REGEX.search(cur_name))

            if len(path_names) > 1 and not is_sold_container and (has_lot_code or len(non_internal_subfolders) == 0):
                all_lot_files, pure_imgs, custom_imgs = collect_all_descendant_files(current_fid)
                if not has_lot_code and not all_lot_files and not pure_imgs:
                    return

                lot_code, stone_name = self.parse_lot_folder_title(cur_name, path_names[:-1])
                modified_time = cur_folder.get("modifiedTime", "")

                metrics = self.compute_lot_metrics(
                    current_fid,
                    lot_code,
                    cur_name,
                    modified_time,
                    all_lot_files,
                    pure_imgs,
                    custom_imgs,
                    bot_cache,
                )

                drive_link = cur_folder.get("webViewLink") or f"https://drive.google.com/drive/folders/{current_fid}"

                record = {
                    "folder_id": current_fid,
                    "folder_name": cur_name,
                    "lot_code": lot_code,
                    "stone_name": stone_name,
                    "collection": collection_name,
                    "path": " / ".join(path_names),
                    "m2": metrics["m2"],
                    "slabs": metrics["slabs"],
                    "avg_dim": metrics["avg_dim"],
                    "surface": metrics["surface"],
                    "drive_link": drive_link,
                    "images": pure_imgs,  # 'custom' fotoğrafları kesinlikle hariç!
                    "all_files": all_lot_files,
                    "modified_time": modified_time,
                    "status": "SOLD OUT" if now_sold_out else "IN STOCK",
                }

                if now_sold_out:
                    new_sold_out[current_fid] = record
                else:
                    new_active[current_fid] = record

                return

            for sf in subfolders:
                sf_name = sf.get("name", "").strip()
                walk_hierarchy(sf["id"], collection_name, path_names + [sf_name], now_sold_out)

        for root_fid, coll_name in root_collection_folders:
            walk_hierarchy(root_fid, coll_name, [coll_name], False)

        with self.inventory_lock:
            self.active_lots = new_active
            self.sold_out_lots = new_sold_out
            self.last_scan_time = now

        self.state.save(git_sync=False)
        return new_active, self.sold_out_lots

    def ensure_deep_metrics_for_lots(self, lots: List[Dict[str, Any]]):
        """
        Sorgulanan veya sunumu yapılan lotların içinde hala m2 == 0 olan varsa,
        doğrudan o lotun Drive alt klasörlerindeki Excel dosyasını canlı sorgulayıp indirerek
        metrajını anında tamamlar.
        """
        headers = self.get_headers()
        if not headers:
            return

        for item in lots:
            if item.get("m2", 0.0) > 0 and item.get("avg_dim", "-") != "-":
                continue

            fid = item["folder_id"]
            all_files = item.get("all_files", [])

            # Excel dosyasını bul ve indir
            for f in all_files:
                mt = str(f.get("mimeType", ""))
                nm = str(f.get("name", "")).lower()
                if mt in (
                    "application/vnd.google-apps.spreadsheet",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "application/vnd.ms-excel",
                ) or nm.endswith((".xlsx", ".xls", ".xlsm")):
                    eb = self.download_file_bytes(f["id"], mt)
                    if eb:
                        parsed = self.parse_excel_metrics_deep(eb, nm)
                        if parsed and parsed.get("m2", 0) > 0:
                            item["m2"] = parsed["m2"]
                            if parsed.get("slabs", 0) > 0:
                                item["slabs"] = parsed["slabs"]
                            if parsed.get("avg_dim", "-") != "-":
                                item["avg_dim"] = parsed["avg_dim"]
                            if parsed.get("surface"):
                                item["surface"] = parsed["surface"]
                            self.state.data.setdefault("metrics_cache", {})[fid] = {
                                "m2": item["m2"],
                                "slabs": item["slabs"],
                                "avg_dim": item["avg_dim"],
                                "surface": item["surface"],
                                "modified_time": item.get("modified_time", ""),
                            }
                            break

    def search_active_by_stone(self, raw_query: str) -> Tuple[str, List[Dict[str, Any]]]:
        stone_query = extract_stone_query_name(raw_query)
        norm_query = normalize_compare(stone_query)
        if not norm_query:
            return stone_query, []

        active_map, _ = self.scan_drive_collections(force=False)
        query_tokens = [t for t in norm_query.split() if len(t) >= 2]

        matched: List[Dict[str, Any]] = []
        for fid, item in active_map.items():
            if self.SOLD_KEYWORD_REGEX.search(item.get("folder_name", "")):
                continue
            if normalize_compare(item.get("lot_code", "")) in INTERNAL_SUBFOLDER_NAMES:
                continue

            searchable = normalize_compare(
                f"{item.get('stone_name', '')} {item.get('folder_name', '')} {item.get('path', '')}"
            )
            if query_tokens and all(tok in searchable for tok in query_tokens):
                matched.append(item)

        def sort_key(x: Dict[str, Any]):
            lc = x.get("lot_code", "")
            m_num = re.search(r"(\d+)", lc)
            num = int(m_num.group(1)) if m_num else 999999
            return (x.get("collection", ""), num, lc)

        matched.sort(key=sort_key)
        # Eşleşen tüm lotların Excel metrajlarını garantiye al
        self.ensure_deep_metrics_for_lots(matched)
        return stone_query.title(), matched

    def find_lot_by_code_or_name(self, query_str: str) -> Optional[Dict[str, Any]]:
        active_map, _ = self.scan_drive_collections(force=False)
        q_norm = normalize_compare(query_str)
        q_compact = re.sub(r"[^a-z0-9]", "", q_norm)

        for item in active_map.values():
            lc_compact = re.sub(r"[^a-z0-9]", "", item.get("lot_code", "").lower())
            if lc_compact and lc_compact == q_compact:
                self.ensure_deep_metrics_for_lots([item])
                return item

        for item in active_map.values():
            blob = normalize_compare(f"{item.get('lot_code', '')} {item.get('stone_name', '')} {item.get('folder_name', '')}")
            if q_norm and q_norm in blob:
                self.ensure_deep_metrics_for_lots([item])
                return item

        return None


def extract_pure_marble_texture(raw_img: Image.Image) -> Image.Image:
    """
    Fotoğrafın kenarlarındaki fabrika duvarlarını, demir sehpaları, karanlık boşlukları
    ve etiket şeritlerini otomatik tespit edip keserek SADECE SAF MERMER DOKUSUNU çıkarır.
    Böylece Bookmatch yapıldığında ortada fabrika penceresi veya boşluk ASLA kalmaz!
    """
    img = raw_img.convert("RGB")
    w, h = img.size

    # Önce dış %14'lük fabrika/sehpa/alt-bant çerçevesini güvenli şekilde kırp
    # (Mermer plaka her zaman fotoğrafın merkezindedir)
    arr = np.asarray(img, dtype=np.float32)

    # Merkez bölgenin (%45x%45) ortalama renk ve parlaklık profilini çıkar (Saf mermer referansı)
    cy1, cy2 = int(h * 0.28), int(h * 0.72)
    cx1, cx2 = int(w * 0.25), int(w * 0.75)
    center_patch = arr[cy1:cy2, cx1:cx2]
    mean_rgb = np.mean(center_patch, axis=(0, 1))
    std_rgb = np.std(center_patch, axis=(0, 1)) + 28.0

    # Her pikselin mermer dokusuna uyum skoru
    diff = np.abs(arr - mean_rgb) / std_rgb
    is_slab_pixel = np.mean(diff, axis=2) < 2.15

    # Çok karanlık (alt siyah yazı bandı) veya aşırı doygun (turuncu/mavi demir sehpa) pikselleri dışla
    brightness = np.mean(arr, axis=2)
    saturation = np.max(arr, axis=2) - np.min(arr, axis=2)
    center_sat = np.mean(np.max(center_patch, axis=2) - np.min(center_patch, axis=2))

    valid_mask = (
        is_slab_pixel
        & (brightness > 28.0)
        & (saturation < (center_sat + 55.0))
    )

    # Satır ve sütun yoğunluğuna göre plakanın iç sınırlarını bul
    row_ratio = np.mean(valid_mask[:, int(w * 0.18) : int(w * 0.82)], axis=1)
    col_ratio = np.mean(valid_mask[int(h * 0.18) : int(h * 0.82), :], axis=0)

    valid_rows = np.where(row_ratio > 0.62)[0]
    valid_cols = np.where(col_ratio > 0.62)[0]

    if len(valid_rows) > int(h * 0.35) and len(valid_cols) > int(w * 0.35):
        y1, y2 = int(valid_rows[0]), int(valid_rows[-1])
        x1, x2 = int(valid_cols[0]), int(valid_cols[-1])
        # Kenar pürüzlerini (kırık taş kenarlarını) sıfırlamak için %5 daha içeri gir
        pad_x = int((x2 - x1) * 0.06)
        pad_y = int((y2 - y1) * 0.07)
        x1 = max(int(w * 0.12), x1 + pad_x)
        x2 = min(int(w * 0.88), x2 - pad_x)
        y1 = max(int(h * 0.14), y1 + pad_y)
        y2 = min(int(h * 0.84), y2 - pad_y)
    else:
        # Garantili merkez mermer kesimi
        x1, x2 = int(w * 0.16), int(w * 0.84)
        y1, y2 = int(h * 0.18), int(h * 0.82)

    if (x2 - x1) < 150 or (y2 - y1) < 120:
        x1, x2 = int(w * 0.15), int(w * 0.85)
        y1, y2 = int(h * 0.16), int(h * 0.82)

    return img.crop((x1, y1, x2, y2))


def build_ram_safe_architectural_render(
    img_bytes: bytes,
    lot_code: str,
    stone_name: str,
    collection_name: str,
    m2: float = 0.0,
    slabs: int = 0,
    avg_dim: str = "-",
) -> bytes:
    """
    ULTRA-GERÇEKÇİ MİMARİ BOOKMATCH & LUXURY SHOWROOM RENDER MOTORU (V22)
    - 'custom' fotoğrafı kullanmaz, arka planı/sehpayı tamamen keser, sadece saf mermer dokusunu alır.
    - İki mermer plakasını sıfıra sıfır (kesintisiz kelebek damar takibiyle) birleştirir.
    - Tavan gizli LED aydınlatması, bronz mimari çıtalar, cilalı zemin yansıması ve net tipografi içerir.
    """
    try:
        with Image.open(io.BytesIO(img_bytes)) as raw_img:
            raw_img = raw_img.convert("RGB")
            raw_img.thumbnail((1400, 1000), Image.Resampling.LANCZOS)
            pure_slab = extract_pure_marble_texture(raw_img)

        # Doğal mermer parlaklık ve damar netliği
        pure_slab = ImageEnhance.Contrast(pure_slab).enhance(1.05)
        pure_slab = ImageEnhance.Sharpness(pure_slab).enhance(1.15)

        # 2'li Kusursuz Bookmatch (Sol Plaka + Sıfıra Sıfır Aynalanmış Sağ Plaka)
        single_w, single_h = 520, 560
        left_panel = pure_slab.resize((single_w, single_h), Image.Resampling.LANCZOS)
        right_panel = ImageOps.mirror(left_panel)

        bm_w, bm_h = single_w * 2, single_h  # 1040 x 560
        bookmatch = Image.new("RGB", (bm_w, bm_h))
        bookmatch.paste(left_panel, (0, 0))
        bookmatch.paste(right_panel, (single_w, 0))

        # Ana Sahne Tuvali: 1400 x 980
        cw, ch = 1400, 980
        y_idx, x_idx = np.mgrid[0:ch, 0:cw].astype(np.float32)
        scene_arr = np.zeros((ch, cw, 3), dtype=np.float32)

        # Lüks Antrasit/Obsidyen Galeri Duvarı
        scene_arr[:, :, 0] = 20.0
        scene_arr[:, :, 1] = 20.0
        scene_arr[:, :, 2] = 22.0

        # Tavan (Y: 0..95) ve Sıcak Gizli LED Covesi
        ceil_mask = y_idx < 95
        scene_arr[ceil_mask] *= (0.45 + 0.55 * (y_idx[ceil_mask] / 95.0))[:, None]

        # Mermer Duvar Yerleşimi (X: 180..1220, Y: 95..655 -> 1040 x 560)
        wall_x1, wall_y1 = 180, 95
        wall_w, wall_h = bm_w, bm_h
        wall_arr = np.asarray(bookmatch, dtype=np.float32)

        # Çok Yumuşak Galeri Tavan Aydınlatması (Doğal cila parlaması)
        wy, wx = np.mgrid[0:wall_h, 0:wall_w].astype(np.float32)
        light_map = np.ones((wall_h, wall_w), dtype=np.float32) * 0.88
        for sx in (wall_w * 0.25, wall_w * 0.50, wall_w * 0.75):
            spread = 160.0 + wy * 0.35
            dist_x = np.abs(wx - sx)
            v_fade = np.clip(1.0 - (wy / (wall_h * 1.35)), 0.35, 1.0)
            light_map += np.exp(-0.5 * (dist_x / spread) ** 2) * v_fade * 0.20

        light_map = np.clip(light_map, 0.75, 1.12)
        wall_lit = np.zeros_like(wall_arr)
        wall_lit[:, :, 0] = np.clip(wall_arr[:, :, 0] * light_map * 1.02, 0, 255)
        wall_lit[:, :, 1] = np.clip(wall_arr[:, :, 1] * light_map * 1.00, 0, 255)
        wall_lit[:, :, 2] = np.clip(wall_arr[:, :, 2] * light_map * 0.97, 0, 255)

        scene_arr[wall_y1 : wall_y1 + wall_h, wall_x1 : wall_x1 + wall_w] = wall_lit

        # Cilalı Koyu Mermer Zemin (Y: 655..885) & Gerçekçi Perspektif Yansıma
        floor_y1, floor_y2 = 655, 885
        floor_h = floor_y2 - floor_y1
        flipped_wall = ImageOps.flip(Image.fromarray(wall_lit.astype(np.uint8))).resize(
            (wall_w, floor_h), Image.Resampling.LANCZOS
        )
        flipped_wall = flipped_wall.filter(ImageFilter.GaussianBlur(radius=3.0))
        refl_arr = np.asarray(flipped_wall, dtype=np.float32)

        fy = np.linspace(0.0, 1.0, floor_h, dtype=np.float32)[:, None, None]
        floor_base = np.array([16.0, 16.0, 18.0], dtype=np.float32) * (1.0 - 0.35 * fy)
        refl_alpha = 0.48 * np.exp(-2.6 * fy)

        scene_arr[floor_y1:floor_y2, :] = floor_base
        scene_arr[floor_y1:floor_y2, wall_x1 : wall_x1 + wall_w] = (
            refl_arr * refl_alpha + floor_base * (1.0 - refl_alpha)
        )

        canvas = Image.fromarray(np.clip(scene_arr, 0, 255).astype(np.uint8))
        draw = ImageDraw.Draw(canvas)

        # İnce Şampanya Altın Mimari Çerçeve & Yan Dekoratif Niş Çizgileri
        gold = (197, 155, 95)
        dark_gold = (110, 86, 52)
        draw.rectangle([wall_x1 - 3, wall_y1 - 3, wall_x1 + wall_w + 2, wall_y1 + wall_h + 2], outline=gold, width=3)
        # Ortadaki kelebek birleşim çizgisi (çok zarif 1px derz)
        draw.line([(wall_x1 + single_w, wall_y1), (wall_x1 + single_w, wall_y1 + wall_h)], fill=(60, 58, 55), width=1)
        # Sol ve sağ yan duvar dikey mimari bronz çıtalar
        draw.line([(90, 95), (90, floor_y1)], fill=dark_gold, width=2)
        draw.line([(cw - 90, 95), (cw - 90, floor_y1)], fill=dark_gold, width=2)
        draw.line([(0, floor_y1), (cw, floor_y1)], fill=dark_gold, width=2)

        # Alt Kurumsal Bilgi Bandı (Y: 885..980)
        draw.rectangle([0, 885, cw, ch], fill=(10, 10, 12))
        draw.line([(0, 885), (cw, 885)], fill=gold, width=2)

        stock_txt = f"{m2:,.2f} m2 ({slabs} Slabs)" if m2 > 0 else f"{slabs} Slabs"
        dim_txt = f"  |  Avg. Size: {avg_dim}" if avg_dim and avg_dim != "-" else ""

        header_line = f"ARSLANLAR MARBLE   |   LOT: {lot_code} — {stone_name.upper()}   ({collection_name.upper()})"
        spec_line = f"Seamless 2x Bookmatch Architectural Wall Application   |   Available Stock: {stock_txt}{dim_txt}"

        draw.text((45, 906), header_line, fill=(235, 205, 155))
        draw.text((45, 940), spec_line, fill=(175, 175, 182))

        out_buf = io.BytesIO()
        canvas.save(out_buf, format="JPEG", quality=92, optimize=True)
        result_bytes = out_buf.getvalue()

        del raw_img, pure_slab, left_panel, right_panel, bookmatch, wall_arr, wall_lit, scene_arr, canvas
        gc.collect()
        return result_bytes
    except Exception as e:
        gc.collect()
        raise RuntimeError(f"Render error: {e}")


def build_english_luxury_pdf(
    stone_title: str, lots: List[Dict[str, Any]], drive_intel: ArslanlarDriveIntelligence
) -> bytes:
    """
    %100 İNGİLİZCE, ORİJİNAL LOGOLU, 'CUSTOM' FOTOĞRAFI KULLANMAYAN,
    METRAJ + PLAKA + ORTALAMA EBAT + YÜZEY DETAYLI PREMIUM SUNUM KATALOĞU.
    """
    buf = io.BytesIO()
    page_w, page_h = landscape(A4)
    c = pdf_canvas.Canvas(buf, pagesize=(page_w, page_h))

    logo_bytes = drive_intel.get_arslanlar_logo_bytes()
    total_m2 = sum(x.get("m2", 0.0) for x in lots)
    total_slabs = sum(x.get("slabs", 0) for x in lots)

    bg_dark = colors.HexColor("#0E0E10")
    card_bg = colors.HexColor("#16161A")
    gold_primary = colors.HexColor("#C59B5F")
    gold_light = colors.HexColor("#E5C99B")
    text_white = colors.HexColor("#F7F7F7")
    text_muted = colors.HexColor("#A0A0A8")

    # 1. KAPAK SAYFASI (COVER PAGE)
    c.setFillColor(bg_dark)
    c.rect(0, 0, page_w, page_h, fill=1, stroke=0)

    c.setStrokeColor(gold_primary)
    c.setLineWidth(2.2)
    c.rect(24, 24, page_w - 48, page_h - 48, fill=0, stroke=1)
    c.setStrokeColor(colors.HexColor("#5A462B"))
    c.setLineWidth(0.8)
    c.rect(30, 30, page_w - 60, page_h - 60, fill=0, stroke=1)

    if logo_bytes:
        try:
            logo_io = io.BytesIO(logo_bytes)
            c.drawImage(
                ImageReader(logo_io),
                (page_w - 230) / 2,
                page_h - 175,
                width=230,
                height=88,
                preserveAspectRatio=True,
                mask="auto",
                anchor="c",
            )
        except Exception:
            pass

    c.setFillColor(gold_primary)
    c.setFont("Helvetica-Bold", 30)
    c.drawCentredString(page_w / 2, page_h - 215, "ARSLANLAR MARBLE")

    c.setFillColor(gold_light)
    c.setFont("Helvetica", 11)
    c.drawCentredString(page_w / 2, page_h - 235, "NATURAL STONE  •  EXCLUSIVE SLAB SELECTION")

    c.setStrokeColor(gold_primary)
    c.setLineWidth(1.5)
    c.line(page_w / 2 - 140, page_h - 255, page_w / 2 + 140, page_h - 255)

    c.setFillColor(text_white)
    c.setFont("Helvetica-Bold", 26)
    c.drawCentredString(page_w / 2, page_h - 305, f"{stone_title.upper()} COLLECTION")

    c.setFillColor(gold_light)
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(page_w / 2, page_h - 332, "LIVE INVENTORY & TECHNICAL SLAB PRESENTATION")

    box_w, box_h = 520, 68
    box_x = (page_w - box_w) / 2
    box_y = page_h - 425
    c.setFillColor(card_bg)
    c.setStrokeColor(gold_primary)
    c.setLineWidth(1.0)
    c.roundRect(box_x, box_y, box_w, box_h, 6, fill=1, stroke=1)

    c.setFillColor(gold_primary)
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(box_x + box_w * 0.20, box_y + 42, "AVAILABLE LOTS")
    c.drawCentredString(box_x + box_w * 0.52, box_y + 42, "TOTAL QUANTITY (M²)")
    c.drawCentredString(box_x + box_w * 0.83, box_y + 42, "TOTAL SLABS")

    c.setFillColor(text_white)
    c.setFont("Helvetica-Bold", 15)
    c.drawCentredString(box_x + box_w * 0.20, box_y + 18, f"{len(lots)} Lots")
    c.drawCentredString(box_x + box_w * 0.52, box_y + 18, f"{total_m2:,.2f} m²")
    c.drawCentredString(box_x + box_w * 0.83, box_y + 18, f"{total_slabs} Slabs")

    c.setFillColor(text_muted)
    c.setFont("Helvetica", 9.5)
    c.drawCentredString(
        page_w / 2,
        52,
        f"Prepared on {datetime.now(TR_TZ).strftime('%B %d, %Y')}  |  Arslanlar Marble Export Department  |  Izmir, Turkiye",
    )
    c.showPage()

    # 2. LOT SAYFALARI (HER LOTUN METRAJI, PLAKA ADEDİ, EBADI VE SAF PLAKA FOTOĞRAFI)
    for idx, item in enumerate(lots[:40], start=1):
        c.setFillColor(bg_dark)
        c.rect(0, 0, page_w, page_h, fill=1, stroke=0)

        c.setStrokeColor(gold_primary)
        c.setLineWidth(1.4)
        c.rect(20, 20, page_w - 40, page_h - 40, fill=0, stroke=1)

        c.setFillColor(card_bg)
        c.rect(21, page_h - 78, page_w - 42, 57, fill=1, stroke=0)
        c.setStrokeColor(gold_primary)
        c.setLineWidth(1.0)
        c.line(20, page_h - 78, page_w - 20, page_h - 78)

        lot_code = item.get("lot_code", "-")
        stone_name = item.get("stone_name", stone_title)
        coll = item.get("collection", "Arslanlar Collection")
        m2 = item.get("m2", 0.0)
        slabs = item.get("slabs", 0)
        ebat = item.get("avg_dim", "-")
        surf = item.get("surface", "Polished (2 cm)")
        link = item.get("drive_link", "")

        c.setFillColor(gold_primary)
        c.setFont("Helvetica-Bold", 19)
        c.drawString(38, page_h - 48, f"LOT: {lot_code}")

        c.setFillColor(text_white)
        c.setFont("Helvetica-Bold", 16)
        c.drawString(170, page_h - 48, f"—  {stone_name.upper()}")

        c.setFillColor(text_muted)
        c.setFont("Helvetica", 10)
        c.drawString(38, page_h - 66, f"COLLECTION: {coll.upper()}   |   ORIGIN: TURKIYE")

        if logo_bytes:
            try:
                c.drawImage(
                    ImageReader(io.BytesIO(logo_bytes)),
                    page_w - 165,
                    page_h - 73,
                    width=130,
                    height=46,
                    preserveAspectRatio=True,
                    mask="auto",
                )
            except Exception:
                pass

        # Teknik Detaylar Şeridi (Metraj, Plaka, Ebat, Yüzey & Kalınlık, Durum)
        spec_y = page_h - 120
        spec_h = 34
        c.setFillColor(colors.HexColor("#1B1B20"))
        c.rect(35, spec_y, page_w - 70, spec_h, fill=1, stroke=0)
        c.setStrokeColor(colors.HexColor("#4A3E2D"))
        c.setLineWidth(0.8)
        c.rect(35, spec_y, page_w - 70, spec_h, fill=0, stroke=1)

        qty_str = f"{m2:,.2f} m²  ({slabs} Slabs)" if m2 > 0 else f"{slabs} Slabs"
        dim_str = ebat if ebat and ebat != "-" else "Standard Export Size"

        metrics_cols = [
            ("TOTAL QUANTITY & SLABS:", qty_str, 46),
            ("AVERAGE DIMENSIONS:", dim_str, 260),
            ("FINISH & THICKNESS:", surf, 470),
            ("BUNDLE STATUS:", "IN STOCK / READY", 675),
        ]
        for lbl, val, x_pos in metrics_cols:
            c.setFillColor(gold_light)
            c.setFont("Helvetica-Bold", 8.5)
            c.drawString(x_pos, spec_y + 20, lbl)
            c.setFillColor(text_white)
            c.setFont("Helvetica-Bold", 10.5)
            c.drawString(x_pos, spec_y + 6, str(val))

        # SADECE custom OLMAYAN görseller kullanılır (item["images"])
        img_candidates = item.get("images") or []
        img_drawn = False
        for img_meta in img_candidates[:2]:
            raw_b = drive_intel.download_file_bytes(img_meta["id"], img_meta.get("mimeType", ""))
            if raw_b:
                try:
                    with Image.open(io.BytesIO(raw_b)) as im:
                        im = im.convert("RGB")
                        im.thumbnail((1200, 750), Image.Resampling.LANCZOS)
                        img_io = io.BytesIO()
                        im.save(img_io, format="JPEG", quality=88)
                        img_io.seek(0)

                        c.drawImage(
                            ImageReader(img_io),
                            38,
                            58,
                            width=page_w - 76,
                            height=page_h - 188,
                            preserveAspectRatio=True,
                            anchor="c",
                        )
                        img_drawn = True
                        break
                except Exception:
                    continue

        if not img_drawn:
            c.setFillColor(text_muted)
            c.setFont("Helvetica-Oblique", 12)
            c.drawCentredString(
                page_w / 2,
                page_h / 2 - 20,
                "High-Resolution Slab Photos & Packing List Available via Google Drive Link Below",
            )

        c.setFillColor(card_bg)
        c.rect(21, 21, page_w - 42, 30, fill=1, stroke=0)
        c.setStrokeColor(gold_primary)
        c.setLineWidth(0.8)
        c.line(20, 51, page_w - 20, 51)

        if link:
            c.setFillColor(gold_light)
            c.setFont("Helvetica-Bold", 10)
            c.drawString(38, 32, "CLICK HERE TO OPEN HIGH-RESOLUTION SLAB PHOTOS & EXCEL PACKING LIST (GOOGLE DRIVE)")
            c.linkURL(link, (35, 24, 610, 48), relative=0)

        c.setFillColor(text_muted)
        c.setFont("Helvetica", 9.5)
        c.drawRightString(page_w - 38, 32, f"ARSLANLAR MARBLE  |  Page {idx} of {min(len(lots), 40)}")

        c.showPage()

    c.save()
    pdf_bytes = buf.getvalue()
    gc.collect()
    return pdf_bytes


def build_master_stock_table_pdf_and_excel(
    active_map: Dict[str, Dict[str, Any]],
    sold_map: Dict[str, Dict[str, Any]],
    drive_intel: ArslanlarDriveIntelligence,
) -> Tuple[bytes, bytes]:
    """
    MADDE 5: Arslanlar Series, Arslanlar Collection ve Arslanlar Exclusive Collection
    içindeki TÜM AKTİF (Stokta Olanlar) ve TÜM SOLD OUT (Satılanlar) lotları içeren
    Premium Kurumsal PDF Katalog Tablosu + Çok Sayfalı Renkli Excel Tablosu üretir.
    """
    def sort_lots(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        def sk(x: Dict[str, Any]):
            lc = x.get("lot_code", "")
            m_n = re.search(r"(\d+)", lc)
            return (x.get("collection", ""), x.get("stone_name", ""), int(m_n.group(1)) if m_n else 999999, lc)
        return sorted(items, key=sk)

    active_list = sort_lots(list(active_map.values()))
    sold_list = sort_lots(list(sold_map.values()))

    # =========================================================================
    # A) EXCEL (.XLSX) MASTER TABLO ÜRETİMİ
    # =========================================================================
    wb = openpyxl.Workbook()
    ws_summary = wb.active
    ws_summary.title = "EXECUTIVE SUMMARY"

    header_fill = PatternFill(start_color="16161A", end_color="16161A", fill_type="solid")
    gold_fill = PatternFill(start_color="C59B5F", end_color="C59B5F", fill_type="solid")
    active_fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
    sold_fill = PatternFill(start_color="FFEBEE", end_color="FFEBEE", fill_type="solid")

    font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=10, bold=True)
    font_regular = Font(name="Calibri", size=10)
    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )

    # 1. Özet Sayfası
    ws_summary.merge_cells("A1:G1")
    ws_summary["A1"] = f"ARSLANLAR MARBLE — MASTER INVENTORY & SOLD OUT REPORT ({datetime.now(TR_TZ).strftime('%d.%m.%Y %H:%M')})"
    ws_summary["A1"].font = font_title
    ws_summary["A1"].fill = header_fill
    ws_summary["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_summary.row_dimensions[1].height = 28

    sum_headers = ["COLLECTION NAME", "ACTIVE LOTS (IN STOCK)", "ACTIVE M²", "ACTIVE SLABS", "SOLD OUT LOTS", "SOLD OUT M²", "TOTAL LOTS"]
    ws_summary.append([])
    ws_summary.append(sum_headers)
    for col_idx in range(1, 8):
        cell = ws_summary.cell(row=3, column=col_idx)
        cell.font = font_header
        cell.fill = gold_fill
        cell.alignment = Alignment(horizontal="center")

    for coll in TARGET_COLLECTIONS:
        c_act = [x for x in active_list if x.get("collection") == coll]
        c_sld = [x for x in sold_list if x.get("collection") == coll]
        act_m2 = sum(x.get("m2", 0.0) for x in c_act)
        act_sl = sum(x.get("slabs", 0) for x in c_act)
        sld_m2 = sum(x.get("m2", 0.0) for x in c_sld)
        row_vals = [coll, len(c_act), round(act_m2, 2), act_sl, len(c_sld), round(sld_m2, 2), len(c_act) + len(c_sld)]
        ws_summary.append(row_vals)
        r_num = ws_summary.max_row
        for c_i in range(1, 8):
            c_cell = ws_summary.cell(row=r_num, column=c_i)
            c_cell.font = font_regular
            c_cell.border = thin_border

    tot_act_m2 = sum(x.get("m2", 0.0) for x in active_list)
    tot_act_sl = sum(x.get("slabs", 0) for x in active_list)
    tot_sld_m2 = sum(x.get("m2", 0.0) for x in sold_list)
    ws_summary.append([
        "GRAND TOTAL", len(active_list), round(tot_act_m2, 2), tot_act_sl,
        len(sold_list), round(tot_sld_m2, 2), len(active_list) + len(sold_list)
    ])
    for c_i in range(1, 8):
        c_cell = ws_summary.cell(row=ws_summary.max_row, column=c_i)
        c_cell.font = font_bold
        c_cell.border = thin_border

    for col in ws_summary.columns:
        ws_summary.column_dimensions[col[0].column_letter].width = 24

    # 2. Aktif Stok Sayfası & 3. Sold Out Sayfası
    def populate_sheet(ws, title_text: str, items: List[Dict[str, Any]], is_sold: bool):
        ws.merge_cells("A1:H1")
        ws["A1"] = title_text
        ws["A1"].font = font_title
        ws["A1"].fill = header_fill
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 26

        cols = ["NO", "COLLECTION", "LOT CODE", "STONE VARIETY", "QUANTITY (M²)", "SLABS (PCS)", "AVG. DIMENSIONS", "STATUS / DRIVE LINK"]
        ws.append(cols)
        for c_i in range(1, 9):
            cell = ws.cell(row=2, column=c_i)
            cell.font = font_header
            cell.fill = gold_fill
            cell.alignment = Alignment(horizontal="center")

        row_fill = sold_fill if is_sold else active_fill
        for idx_i, it in enumerate(items, start=1):
            ws.append([
                idx_i,
                it.get("collection", ""),
                it.get("lot_code", ""),
                it.get("stone_name", ""),
                round(it.get("m2", 0.0), 2),
                it.get("slabs", 0),
                it.get("avg_dim", "-"),
                it.get("drive_link", ""),
            ])
            rn = ws.max_row
            for c_i in range(1, 9):
                c_cell = ws.cell(row=rn, column=c_i)
                c_cell.font = font_regular
                c_cell.fill = row_fill
                c_cell.border = thin_border

        widths_map = [8, 28, 16, 28, 16, 14, 20, 45]
        for idx_w, w_val in enumerate(widths_map, start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(idx_w)].width = w_val

    ws_active = wb.create_sheet(title="IN STOCK (ACTIVE)")
    populate_sheet(ws_active, "ARSLANLAR MARBLE — ACTIVE INVENTORY (IN STOCK)", active_list, is_sold=False)

    ws_sold = wb.create_sheet(title="SOLD OUT (ARCHIVED)")
    populate_sheet(ws_sold, "ARSLANLAR MARBLE — SOLD OUT & ARCHIVED LOTS", sold_list, is_sold=True)

    excel_buf = io.BytesIO()
    wb.save(excel_buf)
    excel_bytes = excel_buf.getvalue()

    # =========================================================================
    # B) PREMIUM PDF MASTER STOK TABLOSU ÜRETİMİ
    # =========================================================================
    pdf_buf = io.BytesIO()
    page_w, page_h = landscape(A4)
    c = pdf_canvas.Canvas(pdf_buf, pagesize=(page_w, page_h))
    logo_bytes = drive_intel.get_arslanlar_logo_bytes()

    bg_dark = colors.HexColor("#0E0E10")
    card_bg = colors.HexColor("#18181C")
    gold_primary = colors.HexColor("#C59B5F")
    gold_light = colors.HexColor("#E5C99B")
    text_white = colors.HexColor("#F5F5F7")
    text_muted = colors.HexColor("#A0A0A8")
    green_badge = colors.HexColor("#2E7D32")
    red_badge = colors.HexColor("#C62828")

    # Kapak & Yönetici Özeti Sayfası
    c.setFillColor(bg_dark)
    c.rect(0, 0, page_w, page_h, fill=1, stroke=0)
    c.setStrokeColor(gold_primary)
    c.setLineWidth(2)
    c.rect(22, 22, page_w - 44, page_h - 44, fill=0, stroke=1)

    if logo_bytes:
        try:
            c.drawImage(
                ImageReader(io.BytesIO(logo_bytes)),
                (page_w - 200) / 2,
                page_h - 140,
                width=200,
                height=75,
                preserveAspectRatio=True,
                mask="auto",
            )
        except Exception:
            pass

    c.setFillColor(gold_primary)
    c.setFont("Helvetica-Bold", 24)
    c.drawCentredString(page_w / 2, page_h - 175, "ARSLANLAR MARBLE — MASTER INVENTORY TABLE")
    c.setFillColor(gold_light)
    c.setFont("Helvetica", 11)
    c.drawCentredString(
        page_w / 2,
        page_h - 195,
        f"COMPLETE STOCK & SOLD OUT LEDGER  |  GENERATED: {datetime.now(TR_TZ).strftime('%d.%m.%Y %H:%M')}",
    )

    # Koleksiyon Özet Tablosu
    y_cursor = page_h - 245
    c.setFillColor(card_bg)
    c.rect(55, y_cursor, page_w - 110, 26, fill=1, stroke=0)
    c.setFillColor(gold_light)
    c.setFont("Helvetica-Bold", 9.5)
    c.drawString(70, y_cursor + 8, "COLLECTION")
    c.drawString(270, y_cursor + 8, "IN STOCK (LOTS)")
    c.drawString(400, y_cursor + 8, "IN STOCK (M²)")
    c.drawString(525, y_cursor + 8, "IN STOCK (SLABS)")
    c.drawString(660, y_cursor + 8, "SOLD OUT (LOTS)")

    y_cursor -= 26
    for coll in TARGET_COLLECTIONS:
        c_act = [x for x in active_list if x.get("collection") == coll]
        c_sld = [x for x in sold_list if x.get("collection") == coll]
        act_m2 = sum(x.get("m2", 0.0) for x in c_act)
        act_sl = sum(x.get("slabs", 0) for x in c_act)

        c.setStrokeColor(colors.HexColor("#2A2A30"))
        c.rect(55, y_cursor, page_w - 110, 24, fill=0, stroke=1)
        c.setFillColor(text_white)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(70, y_cursor + 7, coll)
        c.setFont("Helvetica", 10)
        c.drawString(270, y_cursor + 7, f"{len(c_act)} Lots")
        c.drawString(400, y_cursor + 7, f"{act_m2:,.2f} m²")
        c.drawString(525, y_cursor + 7, f"{act_sl} Slabs")
        c.setFillColor(colors.HexColor("#EF9A9A"))
        c.drawString(660, y_cursor + 7, f"{len(c_sld)} Sold Out")
        y_cursor -= 24

    # Genel Toplam Satırı
    c.setFillColor(colors.HexColor("#242018"))
    c.rect(55, y_cursor, page_w - 110, 26, fill=1, stroke=1)
    c.setFillColor(gold_light)
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(70, y_cursor + 8, "TOTAL INVENTORY")
    c.drawString(270, y_cursor + 8, f"{len(active_list)} Active Lots")
    c.drawString(400, y_cursor + 8, f"{tot_act_m2:,.2f} m²")
    c.drawString(525, y_cursor + 8, f"{tot_act_sl} Slabs")
    c.drawString(660, y_cursor + 8, f"{len(sold_list)} Sold Out")

    c.showPage()

    # Detaylı Satır Satır Tablo Sayfaları (Önce IN STOCK, Sonra SOLD OUT)
    combined_rows = [("IN STOCK", x) for x in active_list] + [("SOLD OUT", x) for x in sold_list]
    rows_per_page = 22
    total_pages = max(1, (len(combined_rows) + rows_per_page - 1) // rows_per_page)

    for p_idx in range(total_pages):
        c.setFillColor(bg_dark)
        c.rect(0, 0, page_w, page_h, fill=1, stroke=0)
        c.setStrokeColor(gold_primary)
        c.setLineWidth(1.4)
        c.rect(20, 20, page_w - 40, page_h - 40, fill=0, stroke=1)

        # Üst Başlık
        c.setFillColor(gold_primary)
        c.setFont("Helvetica-Bold", 14)
        c.drawString(35, page_h - 45, "ARSLANLAR MARBLE — DETAILED INVENTORY & SOLD OUT TABLE")
        c.setFillColor(text_muted)
        c.setFont("Helvetica", 9.5)
        c.drawRightString(page_w - 35, page_h - 45, f"Page {p_idx + 1} of {total_pages}")

        # Tablo Sütun Başlıkları
        table_top = page_h - 75
        c.setFillColor(card_bg)
        c.rect(32, table_top, page_w - 64, 22, fill=1, stroke=0)
        c.setFillColor(gold_light)
        c.setFont("Helvetica-Bold", 8.5)
        c.drawString(40, table_top + 7, "#")
        c.drawString(68, table_top + 7, "LOT CODE")
        c.drawString(150, table_top + 7, "STONE VARIETY")
        c.drawString(330, table_top + 7, "COLLECTION")
        c.drawString(495, table_top + 7, "QUANTITY (M²)")
        c.drawString(585, table_top + 7, "SLABS")
        c.drawString(640, table_top + 7, "AVG. SIZE")
        c.drawString(735, table_top + 7, "STATUS")

        row_y = table_top - 20
        page_slice = combined_rows[p_idx * rows_per_page : (p_idx + 1) * rows_per_page]

        for r_offset, (status_label, item) in enumerate(page_slice, start=p_idx * rows_per_page + 1):
            if r_offset % 2 == 0:
                c.setFillColor(colors.HexColor("#141418"))
                c.rect(32, row_y, page_w - 64, 20, fill=1, stroke=0)

            c.setFillColor(text_muted)
            c.setFont("Helvetica", 8.5)
            c.drawString(40, row_y + 6, str(r_offset))

            c.setFillColor(gold_light if status_label == "IN STOCK" else text_muted)
            c.setFont("Helvetica-Bold", 9)
            c.drawString(68, row_y + 6, str(item.get("lot_code", "-"))[:14])

            c.setFillColor(text_white)
            c.setFont("Helvetica", 9)
            c.drawString(150, row_y + 6, str(item.get("stone_name", "-"))[:28])

            c.setFillColor(text_muted)
            c.drawString(330, row_y + 6, str(item.get("collection", "-"))[:26])

            m2_v = item.get("m2", 0.0)
            sl_v = item.get("slabs", 0)
            c.setFillColor(text_white)
            c.setFont("Helvetica-Bold", 9)
            c.drawString(495, row_y + 6, f"{m2_v:,.2f} m²" if m2_v > 0 else "-")
            c.drawString(585, row_y + 6, f"{sl_v} Pcs" if sl_v > 0 else "-")

            c.setFont("Helvetica", 8.5)
            c.drawString(640, row_y + 6, str(item.get("avg_dim", "-"))[:14])

            # Status Badge
            badge_color = green_badge if status_label == "IN STOCK" else red_badge
            c.setFillColor(badge_color)
            c.roundRect(730, row_y + 3, 68, 14, 3, fill=1, stroke=0)
            c.setFillColor(text_white)
            c.setFont("Helvetica-Bold", 7.5)
            c.drawCentredString(764, row_y + 7, status_label)

            link = item.get("drive_link", "")
            if link:
                c.linkURL(link, (65, row_y + 2, 725, row_y + 18), relative=0)

            row_y -= 20

        c.showPage()

    c.save()
    pdf_bytes = pdf_buf.getvalue()
    gc.collect()
    return pdf_bytes, excel_bytes


class ArslanlarPluginController:
    def __init__(self, app: Application):
        self.app = app
        self.state = PersistentWatcherState()
        self.intel = ArslanlarDriveIntelligence(self.state)
        self.bot_token = os.getenv("TELEGRAM_TOKEN", "").strip()

    def send_telegram_sync(self, text: str, parse_mode: str = "HTML"):
        if not self.bot_token or not self.state.data.get("chat_ids"):
            return
        url = f"https://api.telegram.org/bot{self.bot_token}/sendMessage"
        for cid in list(self.state.data["chat_ids"]):
            try:
                requests.post(
                    url,
                    json={
                        "chat_id": cid,
                        "text": text,
                        "parse_mode": parse_mode,
                        "disable_web_page_preview": True,
                    },
                    timeout=15,
                )
            except Exception as e:
                logger.warning(f"Telegram notification warning ({cid}): {e}")

    def send_telegram_document_sync(self, file_bytes: bytes, filename: str, caption: str):
        if not self.bot_token or not self.state.data.get("chat_ids"):
            return
        url = f"https://api.telegram.org/bot{self.bot_token}/sendDocument"
        for cid in list(self.state.data["chat_ids"]):
            try:
                files = {"document": (filename, io.BytesIO(file_bytes))}
                data = {"chat_id": cid, "caption": caption, "parse_mode": "HTML"}
                requests.post(url, data=data, files=files, timeout=30)
            except Exception as e:
                logger.warning(f"Telegram doc warning ({cid}): {e}")

    def format_executive_summary_report(self, title_prefix: str = "WEEKLY") -> str:
        active_map, sold_map = self.intel.scan_drive_collections(force=True)

        total_lots = len(active_map)
        total_m2 = sum(x.get("m2", 0.0) for x in active_map.values())
        total_slabs = sum(x.get("slabs", 0) for x in active_map.values())

        lines = [
            f"📊 <b>ARSLANLAR MARBLE — {title_prefix} INVENTORY & SOLD OUT REPORT</b>",
            f"🗓 <i>{datetime.now(TR_TZ).strftime('%d.%m.%Y %H:%M')}</i>\n",
            f"🟢 <b>Total In Stock (Active):</b> <code>{total_m2:,.2f} m²</code> (<code>{total_slabs} Slabs</code> / <code>{total_lots} Lots</code>)",
            f"🔴 <b>Total Sold Out (Archived):</b> <code>{len(sold_map)} Lots</code>\n",
            "📂 <b>Breakdown by Collection:</b>",
        ]

        for coll in TARGET_COLLECTIONS:
            c_active = [x for x in active_map.values() if x.get("collection") == coll]
            c_sold = [x for x in sold_map.values() if x.get("collection") == coll]
            c_m2 = sum(x.get("m2", 0.0) for x in c_active)
            c_slabs = sum(x.get("slabs", 0) for x in c_active)
            lines.append(
                f"• <b>{coll}:</b>\n"
                f"   🟢 In Stock: <code>{len(c_active)} Lots</code> | <code>{c_m2:,.2f} m²</code> (<code>{c_slabs} Slabs</code>)\n"
                f"   🔴 Sold Out: <code>{len(c_sold)} Lots</code>"
            )

        variety_totals: Dict[str, Dict[str, Any]] = {}
        for x in active_map.values():
            st = x.get("stone_name", "Other").strip().title()
            entry = variety_totals.setdefault(st, {"lots": 0, "m2": 0.0, "slabs": 0})
            entry["lots"] += 1
            entry["m2"] += x.get("m2", 0.0)
            entry["slabs"] += x.get("slabs", 0)

        top_varieties = sorted(
            variety_totals.items(), key=lambda kv: (kv[1]["m2"], kv[1]["lots"]), reverse=True
        )[:12]

        if top_varieties:
            lines.append("\n💎 <b>Top Active Stone Varieties:</b>")
            for st_name, st_info in top_varieties:
                lines.append(
                    f"  ▫️ <b>{st_name}:</b> {st_info['lots']} Lots — "
                    f"<code>{st_info['m2']:,.2f} m²</code> ({st_info['slabs']} Slabs)"
                )

        return "\n".join(lines)

    def background_watcher_loop(self):
        time.sleep(12)
        logger.info("🔍 Arslanlar V22 Drive Watcher started.")

        while True:
            try:
                active_map, sold_map = self.intel.scan_drive_collections(force=True)

                if active_map or sold_map:
                    if not self.state.data.get("initialized"):
                        self.state.data["active_lots"] = {
                            fid: {
                                "lot_code": v["lot_code"],
                                "stone_name": v["stone_name"],
                                "collection": v["collection"],
                                "m2": v["m2"],
                                "slabs": v["slabs"],
                                "drive_link": v["drive_link"],
                            }
                            for fid, v in active_map.items()
                        }
                        self.state.data["sold_out_lots"] = {
                            fid: {
                                "lot_code": v["lot_code"],
                                "stone_name": v["stone_name"],
                                "collection": v["collection"],
                            }
                            for fid, v in sold_map.items()
                        }
                        self.state.data["initialized"] = True
                        self.state.save(git_sync=True)
                    else:
                        prev_active = self.state.data.get("active_lots", {})
                        prev_sold = self.state.data.get("sold_out_lots", {})
                        prev_active_codes = {
                            v.get("lot_code"): fid for fid, v in prev_active.items() if v.get("lot_code")
                        }

                        changed = False

                        for s_fid, s_item in sold_map.items():
                            s_code = s_item.get("lot_code", "")
                            was_in_sold = s_fid in prev_sold or any(
                                pv.get("lot_code") == s_code for pv in prev_sold.values()
                            )
                            was_in_active = s_fid in prev_active or (s_code in prev_active_codes)

                            if not was_in_sold and (was_in_active or s_fid not in prev_sold):
                                changed = True
                                stone_label = s_item.get("stone_name", "")
                                coll_label = s_item.get("collection", "")
                                msg = (
                                    f"📢 <b>{s_code} ({stone_label}) Sold Out klasörüne taşındı, stoktan düşüldü.</b>\n"
                                    f"📂 Koleksiyon: <i>{coll_label}</i>"
                                )
                                self.send_telegram_sync(msg)

                        for a_fid, a_item in active_map.items():
                            a_code = a_item.get("lot_code", "")
                            if a_fid not in prev_active and a_code not in prev_active_codes:
                                changed = True
                                self.intel.ensure_deep_metrics_for_lots([a_item])
                                stone_label = a_item.get("stone_name", "")
                                coll_label = a_item.get("collection", "")
                                m2_val = a_item.get("m2", 0.0)
                                slabs_val = a_item.get("slabs", 0)
                                ebat_val = a_item.get("avg_dim", "-")
                                link_val = a_item.get("drive_link", "")
                                msg = (
                                    f"🆕 <b>Yeni Lot Eklendi: {a_code} ({stone_label})</b>\n"
                                    f"📂 Koleksiyon: <b>{coll_label}</b>\n"
                                    f"• Available Stock: <code>{m2_val:.2f} m² ({slabs_val} Slabs)</code>\n"
                                    f"• Avg. Dimensions: <code>{ebat_val}</code>\n"
                                    f"• Drive Folder: <a href='{link_val}'>Open Folder</a>"
                                )
                                self.send_telegram_sync(msg)

                        self.state.data["active_lots"] = {
                            fid: {
                                "lot_code": v["lot_code"],
                                "stone_name": v["stone_name"],
                                "collection": v["collection"],
                                "m2": v["m2"],
                                "slabs": v["slabs"],
                                "drive_link": v["drive_link"],
                            }
                            for fid, v in active_map.items()
                        }
                        self.state.data["sold_out_lots"] = {
                            fid: {
                                "lot_code": v["lot_code"],
                                "stone_name": v["stone_name"],
                                "collection": v["collection"],
                            }
                            for fid, v in sold_map.items()
                        }
                        if changed:
                            self.state.save(git_sync=True)

                now_tr = datetime.now(TR_TZ)
                if now_tr.weekday() in (0, 4) and now_tr.hour >= 9:
                    report_key = now_tr.strftime("%Y-%m-%d")
                    if self.state.data.get("last_summary_report_key") != report_key:
                        day_title = "MONDAY" if now_tr.weekday() == 0 else "FRIDAY"
                        report_html = self.format_executive_summary_report(title_prefix=day_title)
                        self.send_telegram_sync(report_html)
                        try:
                            pdf_b, xlsx_b = build_master_stock_table_pdf_and_excel(active_map, sold_map, self.intel)
                            self.send_telegram_document_sync(
                                pdf_b,
                                "ARSLANLAR_MARBLE_MASTER_INVENTORY_TABLE.pdf",
                                f"📊 <b>{day_title} MASTER INVENTORY TABLE (PDF)</b> — In Stock & Sold Out",
                            )
                            self.send_telegram_document_sync(
                                xlsx_b,
                                "ARSLANLAR_MARBLE_MASTER_STOCK_TABLE.xlsx",
                                f"📗 <b>{day_title} MASTER INVENTORY TABLE (EXCEL)</b> — In Stock & Sold Out",
                            )
                        except Exception:
                            pass
                        self.state.data["last_summary_report_key"] = report_key
                        self.state.save(git_sync=True)

            except Exception as e:
                logger.warning(f"Drive watcher loop warning: {e}")

            time.sleep(120)

    async def handle_priority_message(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        if not update.effective_message or not update.effective_message.text:
            return

        if update.effective_chat:
            self.state.register_chat(update.effective_chat.id)

        raw_text = update.effective_message.text.strip()
        norm_text = normalize_compare(raw_text)

        # ---------------------------------------------------------------------
        # 1. DETAYLI MASTER STOK TABLOSU & ÖZET RAPOR İSTEĞİ
        # ("stok tablosu", "detaylı stok", "stok listesi", "özet rapor", "satılanlar")
        # ---------------------------------------------------------------------
        is_master_table_req = any(
            kw in norm_text
            for kw in (
                "stok tablosu", "detayli stok", "detayli liste", "master stok",
                "tum stok", "genel stok", "ozet rapor", "stok ozet", "haftalik ozet",
                "stokta olanlar", "satilanlar", "sold out listesi"
            )
        ) or norm_text in ("ozet", "tablo", "summary") or raw_text.lower().startswith(("/ozet", "/tablo", "/stok"))

        if is_master_table_req:
            status_msg = await update.effective_message.reply_text(
                "⏳ <b>Arslanlar Series, Collection ve Exclusive Collection</b> (Aktif + Sold Out) taranıyor, "
                "Master Stok Tablosu (PDF & Excel) hazırlanıyor...",
                parse_mode="HTML",
            )
            loop = asyncio.get_running_loop()

            def generate_all():
                act_m, sld_m = self.intel.scan_drive_collections(force=True)
                # Aktif lotların metrajlarını garantiye al
                self.intel.ensure_deep_metrics_for_lots(list(act_m.values()))
                rep_html = self.format_executive_summary_report(title_prefix="MASTER")
                pdf_b, xlsx_b = build_master_stock_table_pdf_and_excel(act_m, sld_m, self.intel)
                return rep_html, pdf_b, xlsx_b, len(act_m), len(sld_m)

            report_html, pdf_bytes, excel_bytes, act_cnt, sld_cnt = await loop.run_in_executor(
                RENDER_EXECUTOR, generate_all
            )

            await status_msg.edit_text(report_html, parse_mode="HTML", disable_web_page_preview=True)

            await update.effective_message.reply_document(
                document=io.BytesIO(pdf_bytes),
                filename="ARSLANLAR_MARBLE_MASTER_INVENTORY_TABLE.pdf",
                caption=(
                    f"📕 <b>ARSLANLAR MARBLE — MASTER INVENTORY TABLE (PDF)</b>\n"
                    f"• 🟢 <b>In Stock:</b> <code>{act_cnt} Active Lots</code>\n"
                    f"• 🔴 <b>Sold Out:</b> <code>{sld_cnt} Archived Lots</code>"
                ),
                parse_mode="HTML",
            )
            await update.effective_message.reply_document(
                document=io.BytesIO(excel_bytes),
                filename="ARSLANLAR_MARBLE_MASTER_STOCK_TABLE.xlsx",
                caption=(
                    f"📗 <b>ARSLANLAR MARBLE — MASTER INVENTORY TABLE (EXCEL)</b>\n"
                    f"• 3 Sheets: <i>Executive Summary, In Stock (Active), Sold Out (Archived)</i>"
                ),
                parse_mode="HTML",
            )
            raise ApplicationHandlerStop

        # ---------------------------------------------------------------------
        # 2. RAM-KORUMALI MİMARİ RENDER & BOOKMATCH (CUSTOM FOTOĞRAFSIZ, SAF MERMER DOKUSU)
        # ---------------------------------------------------------------------
        if re.search(r"\b(render|bookmatch|kelebek)\b", norm_text):
            target_q = re.sub(
                r"\b(render|bookmatch|kelebek|yap|lütfen|lutfen|hazırla|hazirla|stüdyo|studyo)\b",
                " ",
                raw_text,
                flags=re.I,
            ).strip(" /.,:-")

            if not target_q:
                await update.effective_message.reply_text(
                    "🎨 Please specify a lot code for rendering.\nExample: <code>render A-371</code>",
                    parse_mode="HTML",
                )
                raise ApplicationHandlerStop

            status_msg = await update.effective_message.reply_text(
                f"🎨 Preparing Seamless Bookmatch & Architectural Showroom Render for <b>{target_q.upper()}</b>...",
                parse_mode="HTML",
            )
            loop = asyncio.get_running_loop()
            lot_item = await loop.run_in_executor(None, self.intel.find_lot_by_code_or_name, target_q)

            # KESİN KURAL: Sadece custom olmayan saf plaka fotoğrafları (lot_item["images"])
            imgs = lot_item.get("images", []) if lot_item else []
            if not lot_item or not imgs:
                await status_msg.edit_text(
                    f"⚠️ No non-custom slab photo found for <b>{target_q}</b>.",
                    parse_mode="HTML",
                )
                raise ApplicationHandlerStop

            img_bytes = await loop.run_in_executor(
                None, self.intel.download_file_bytes, imgs[0]["id"], imgs[0].get("mimeType", "")
            )
            if not img_bytes:
                await status_msg.edit_text("⚠️ Could not download slab image from Google Drive.")
                raise ApplicationHandlerStop

            try:
                rendered_bytes = await loop.run_in_executor(
                    RENDER_EXECUTOR,
                    build_ram_safe_architectural_render,
                    img_bytes,
                    lot_item["lot_code"],
                    lot_item["stone_name"],
                    lot_item["collection"],
                    lot_item.get("m2", 0.0),
                    lot_item.get("slabs", 0),
                    lot_item.get("avg_dim", "-"),
                )
                m2_v = lot_item.get("m2", 0.0)
                sl_v = lot_item.get("slabs", 0)
                eb_v = lot_item.get("avg_dim", "-")
                sf_v = lot_item.get("surface", "Polished (2 cm)")
                caption = (
                    f"🏛 <b>{lot_item['lot_code']} — {lot_item['stone_name']}</b>\n"
                    f"• <b>Collection:</b> {lot_item['collection']}\n"
                    f"• <b>Available Stock:</b> <code>{m2_v:,.2f} m² ({sl_v} Slabs)</code>\n"
                    f"• <b>Avg. Dimensions:</b> <code>{eb_v}</code> | {sf_v}\n"
                    f"• <b>Studio:</b> Seamless 2x Bookmatch Architectural Wall Simulation"
                )
                await update.effective_message.reply_photo(
                    photo=io.BytesIO(rendered_bytes),
                    caption=caption,
                    parse_mode="HTML",
                )
                await status_msg.delete()
            except Exception as e:
                await status_msg.edit_text(f"⚠️ Render error: {e}")

            raise ApplicationHandlerStop

        # ---------------------------------------------------------------------
        # 3. TAŞ ÇEŞİDİNE ÖZEL İNGİLİZCE PREMIUM SUNUM ("Silver Rootslar sunum lütfen", "Silver Riverlar sunum")
        # ---------------------------------------------------------------------
        has_sunum_kw = bool(re.search(r"\b(sunum|sunumu|katalog|kataloğu|presentation|catalog)\b", norm_text))
        is_main_collection_sunum = any(
            normalize_compare(c) in norm_text
            or norm_text.startswith("series")
            or norm_text.startswith("collection")
            or norm_text.startswith("exclusive")
            for c in TARGET_COLLECTIONS
        )

        if has_sunum_kw and not is_main_collection_sunum:
            status_msg = await update.effective_message.reply_text(
                "📑 Scanning active lots, reading Excel metrics & generating English Luxury Presentation PDF...",
                parse_mode="HTML",
            )
            stone_title, matched_lots = await asyncio.get_running_loop().run_in_executor(
                None, self.intel.search_active_by_stone, raw_text
            )
            if matched_lots:
                await self.send_stone_variety_summary(update, stone_title, matched_lots, send_photos=False)

                try:
                    pdf_bytes = await asyncio.get_running_loop().run_in_executor(
                        RENDER_EXECUTOR, build_english_luxury_pdf, stone_title, matched_lots, self.intel
                    )
                    safe_fn = re.sub(r"[^a-zA-Z0-9_]", "_", stone_title.upper()) + "_ARSLANLAR_MARBLE_PRESENTATION.pdf"
                    total_m2 = sum(x.get("m2", 0.0) for x in matched_lots)
                    total_slabs = sum(x.get("slabs", 0) for x in matched_lots)
                    await update.effective_message.reply_document(
                        document=io.BytesIO(pdf_bytes),
                        filename=safe_fn,
                        caption=(
                            f"🏛 <b>ARSLANLAR MARBLE — {stone_title.upper()} PRESENTATION</b>\n"
                            f"• <b>Available Lots:</b> <code>{len(matched_lots)} Lots</code>\n"
                            f"• <b>Total Quantity:</b> <code>{total_m2:,.2f} m² ({total_slabs} Slabs)</code>\n"
                            f"• <b>Edition:</b> International Export Standard (English)"
                        ),
                        parse_mode="HTML",
                    )
                    await status_msg.delete()
                except Exception as e:
                    await status_msg.edit_text(f"⚠️ PDF error: {e}")

                raise ApplicationHandlerStop
            else:
                await status_msg.delete()

        # ---------------------------------------------------------------------
        # 4. TAŞ ÇEŞİDİNE GÖRE STOK SORGULAMA ("Ne kadar Silver Roots var?", "Ne kadar Silver River var?")
        # ---------------------------------------------------------------------
        is_variety_query = bool(
            re.search(
                r"(^ne\s+kadar\s+|var\s+m[ıi]|kaç\s+m2|kac\s+m2|\bçeşitleri\b|\bcesitleri\b|\bneler\s+var\b|\bvar\s*\??$|^how\s+much\s+)",
                raw_text,
                flags=re.I,
            )
        )
        is_single_lot_code = bool(re.match(r"^[A-Za-z]{1,3}\s*[-_]?\s*\d{1,5}[A-Za-z]?$", raw_text))

        if is_variety_query and not is_single_lot_code:
            status_msg = await update.effective_message.reply_text("🔍 Scanning active collections & reading Excel metrics...")
            stone_title, matched_lots = await asyncio.get_running_loop().run_in_executor(
                None, self.intel.search_active_by_stone, raw_text
            )
            await status_msg.delete()

            if not matched_lots:
                await update.effective_message.reply_text(
                    f"🔍 No active stock found for <b>{stone_title}</b> in active collections "
                    f"(<i>Arslanlar Series, Collection, Exclusive</i>).\n"
                    f"<i>(Sold Out folders and SOLD lots are excluded.)</i>",
                    parse_mode="HTML",
                )
                raise ApplicationHandlerStop

            await self.send_stone_variety_summary(update, stone_title, matched_lots, send_photos=True)
            raise ApplicationHandlerStop

    async def send_stone_variety_summary(
        self, update: Update, stone_title: str, matched_lots: List[Dict[str, Any]], send_photos: bool = True
    ):
        total_m2 = sum(x.get("m2", 0.0) for x in matched_lots)
        total_slabs = sum(x.get("slabs", 0) for x in matched_lots)

        lines = [
            f"💎 <b>{stone_title.upper()} — ACTIVE INVENTORY & SELECTION</b>",
            f"📊 <b>Total Available Stock:</b> <code>{total_m2:,.2f} m² ({total_slabs} Slabs)</code> — <b>{len(matched_lots)} Active Lots</b>",
            "<i>(Sold Out folders & SOLD lots are strictly excluded)</i>\n",
        ]

        for item in matched_lots:
            l_code = item.get("lot_code", "-")
            s_name = item.get("stone_name", stone_title)
            coll = item.get("collection", "")
            m2 = item.get("m2", 0.0)
            slabs = item.get("slabs", 0)
            ebat = item.get("avg_dim", "-")
            surf = item.get("surface", "Polished (2 cm)")
            link = item.get("drive_link", "")

            stock_disp = f"{m2:,.2f} m² ({slabs} Slabs)" if m2 > 0 else f"{slabs} Slabs"
            lines.append(
                f"🔹 <b>{l_code}</b> — {s_name} <i>({coll})</i>\n"
                f"  • Available Stock: <code>{stock_disp}</code>\n"
                f"  • Avg. Dimensions: <code>{ebat}</code> | {surf}\n"
                f"  • Drive Folder: <a href='{link}'>Open High-Res Folder</a>\n"
            )

        full_text = "\n".join(lines)
        if len(full_text) <= 4000:
            await update.effective_message.reply_text(
                full_text, parse_mode="HTML", disable_web_page_preview=True
            )
        else:
            chunk = ""
            for ln in lines:
                if len(chunk) + len(ln) + 2 > 3900:
                    await update.effective_message.reply_text(
                        chunk, parse_mode="HTML", disable_web_page_preview=True
                    )
                    chunk = ln + "\n"
                else:
                    chunk += ln + "\n"
            if chunk.strip():
                await update.effective_message.reply_text(
                    chunk, parse_mode="HTML", disable_web_page_preview=True
                )

        # SADECE custom OLMAYAN görseller (item["images"]) gönderilir
        if send_photos:
            media_group = []
            loop = asyncio.get_running_loop()
            for item in matched_lots[:4]:
                imgs = item.get("images") or []
                if not imgs:
                    continue
                b = await loop.run_in_executor(
                    None, self.intel.download_file_bytes, imgs[0]["id"], imgs[0].get("mimeType", "")
                )
                if b:
                    m2_v = item.get("m2", 0.0)
                    sl_v = item.get("slabs", 0)
                    eb_v = item.get("avg_dim", "-")
                    st_txt = f"{m2_v:,.2f} m² ({sl_v} Slabs) | {eb_v}" if m2_v > 0 else f"{sl_v} Slabs"
                    cap = f"💎 {item['lot_code']} — {item['stone_name']} | {st_txt}"
                    media_group.append(InputMediaPhoto(media=io.BytesIO(b), caption=cap))

            if media_group:
                try:
                    await update.effective_message.reply_media_group(media=media_group)
                except Exception as e:
                    logger.warning(f"Media group warning: {e}")


ORIGINAL_RUN_POLLING = Application.run_polling
PLUGIN_ATTACHED = False


def hooked_run_polling(self: Application, *args, **kwargs):
    global PLUGIN_ATTACHED
    if not PLUGIN_ATTACHED:
        PLUGIN_ATTACHED = True
        try:
            controller = ArslanlarPluginController(self)
            self.add_handler(
                MessageHandler(filters.TEXT, controller.handle_priority_message),
                group=-1,
            )
            watcher_thread = threading.Thread(
                target=controller.background_watcher_loop,
                name="ArslanlarDriveWatcherV22",
                daemon=True,
            )
            watcher_thread.start()
            logger.info("✅ Arslanlar V22 Master Architecture Plugin successfully attached!")
        except Exception as e:
            logger.error(f"Plugin startup error: {e}")

    return ORIGINAL_RUN_POLLING(self, *args, **kwargs)


Application.run_polling = hooked_run_polling
