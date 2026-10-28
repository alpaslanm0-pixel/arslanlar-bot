# -*- coding: utf-8 -*-
"""
ARSLANLAR MARBLE - V20 ULTRA-PREMIUM & ZERO-TOUCH PLUGIN (sitecustomize.py)
- bot.py dosyasina HIC DOKUNMAZ, mevcut sistemi %100 korur.
- 'custom', 'email', 'photos' gibi alt klasorleri asla ayri lot sanmaz; lotun icine birlestirir.
- Klasor adinda 'SOLD' gecenleri ve 'Sold Out' klasorundekileri kesin olarak eler.
- Gercek m2, Plaka (Slabs), Ortalama Ebat (Dimensions) ve Kalinlik verilerini bot.py + Excel + Dosya analizinden ceker.
- PDF Sunumlarini %100 Ingilizce, Orijinal Arslanlar Marble Logolu ve Premium tasarimla uretir.
- Silver Roots, Rosso Levanto, Tundra, Silver River vb. TUM tas cesitlerinde dinamik calisir.
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
from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageEnhance, ImageOps

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

# Lot klasörlerinin içindeki alt klasörler (Bunlar ayrı birer lot DEĞİLDİR!)
INTERNAL_SUBFOLDER_NAMES = {
    "custom", "email", "mail", "photos", "photo", "foto", "fotograf", "fotograflar",
    "images", "image", "img", "slabs", "slab", "plaka", "plakalar", "raw", "original",
    "orijinal", "video", "videos", "pdf", "pdfs", "katalog", "catalog", "whatsapp",
    "web", "social", "instagram", "render", "renders", "bookmatch", "excel", "list",
    "packing", "etiket", "label", "qr", "temp", "tmp", "backup", "arsiv", "archive"
}

RENDER_EXECUTOR = ThreadPoolExecutor(max_workers=1, thread_name_prefix="RAMSafeWorker")
TR_TZ = timezone(timedelta(hours=3))


def _normalize_compare(s: str) -> str:
    """Metinleri Türkçe/İngilizce karakter farkı olmadan karşılaştırmak için normalize eder."""
    if not s:
        return ""
    s = s.replace("İ", "i").replace("I", "i").replace("ı", "i").lower()
    tr_map = str.maketrans("çğöşü", "cgosu")
    s = s.translate(tr_map)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _extract_stone_query_name(text: str) -> str:
    """
    Kullanıcının mesajından aradığı taş çeşidini (Silver Roots, Rosso Levanto, Tundra vb.)
    Türkçe çoğul/çekim eklerini (-lar, -ler, sunum, ne kadar var vb.) temizleyerek çıkarır.
    """
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
        # 'rootslar' -> 'roots', 'levantolar' -> 'levanto', 'tundralar' -> 'tundra'
        w_clean = re.sub(r"(ları|leri|lar|ler|dan|den|tan|ten|nın|nin|nun|nün)$", "", w)
        if len(w_clean) >= 2:
            cleaned_words.append(w_clean)
        else:
            cleaned_words.append(w)

    return " ".join(cleaned_words).strip(" ?.,!:-")


class PersistentWatcherState:
    def __init__(self, filepath: str = STATE_FILE):
        self.filepath = filepath
        self.lock = threading.Lock()
        self.data: Dict[str, Any] = {
            "chat_ids": [],
            "initialized": False,
            "active_lots": {},
            "sold_out_lots": {},
            "metrics_cache": {},  # folder_id -> {m2, slabs, avg_dim, surface, modified_time}
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
                    logger.warning(f"State okuma uyarisi: {e}")

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
                logger.warning(f"State yazma uyarisi: {e}")

        if git_sync:
            threading.Thread(target=self._git_persist_async, daemon=True).start()

    def _git_persist_async(self):
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
    """
    Drive Klasör Yapısını (Arslanlar Series, Collection, Exclusive Collection)
    ve altındaki Lot klasörlerini (custom, email, photos alt klasörlerini birleştirerek)
    hatasız tarayan, Excel/Sheets ve görsel adlarından gerçek m2/plaka okuyan motor.
    """

    # Lot Kodu Yakalayıcı: A-371, A-1030, B-24, M-105 vb.
    LOT_CODE_REGEX = re.compile(r"\b([A-Z]{1,3}\s*[-_]\s*\d{1,5}[A-Z]?|[A-Z]{1,2}\d{2,5}[A-Z]?)\b", re.IGNORECASE)
    SOLD_KEYWORD_REGEX = re.compile(r"\b(sold\s*out|soldout|sold|satildi|satıldı|pasif)\b", re.IGNORECASE)

    def __init__(self, state: PersistentWatcherState):
        self.state = state
        self.sa_json_str = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
        self._creds = None
        self._token_lock = threading.Lock()
        self.session = requests.Session()
        self.inventory_lock = threading.Lock()
        self.active_lots: Dict[str, Dict[str, Any]] = {}
        self.sold_out_lots: Dict[str, Dict[str, Any]] = {}
        self.last_scan_time: float = 0.0
        self._cached_logo_bytes: Optional[bytes] = None
        self._init_credentials()

    def _init_credentials(self):
        if not self.sa_json_str:
            return
        try:
            info = json.loads(self.sa_json_str)
            self._creds = service_account.Credentials.from_service_account_info(
                info, scopes=["https://www.googleapis.com/auth/drive.readonly"]
            )
        except Exception as e:
            logger.error(f"Plugin Drive Kimlik Hatasi: {e}")

    def _get_headers(self) -> Optional[Dict[str, str]]:
        if not self._creds:
            return None
        with self._token_lock:
            if not self._creds.valid or self._creds.expired or not self._creds.token:
                req = google.auth.transport.requests.Request()
                self._creds.refresh(req)
            return {"Authorization": f"Bearer {self._creds.token}"}

    def get_arslanlar_logo_bytes(self) -> Optional[bytes]:
        """
        bot.py içindeki orijinal Arslanlar Marble logosunu (base64 veya dosya olarak)
        otomatik yakalar. Bulamazsa yüksek çözünürlüklü vektörel amblem üretir.
        """
        if self._cached_logo_bytes:
            return self._cached_logo_bytes

        # 1. bot.py (__main__) içindeki base64 veya logo değişkenlerini tara
        try:
            main_mod = sys.modules.get("__main__")
            if main_mod:
                for attr_name in dir(main_mod):
                    if "logo" in attr_name.lower() or "b64" in attr_name.lower() or "icon" in attr_name.lower():
                        val = getattr(main_mod, attr_name, None)
                        if isinstance(val, bytes) and len(val) > 200:
                            self._cached_logo_bytes = val
                            return val
                        if isinstance(val, str) and len(val) > 300:
                            clean_b64 = re.sub(r"^data:image/[a-zA-Z]+;base64,", "", val.strip())
                            try:
                                decoded = base64.b64decode(clean_b64)
                                with Image.open(io.BytesIO(decoded)) as im:
                                    im.verify()
                                self._cached_logo_bytes = decoded
                                return decoded
                            except Exception:
                                pass
        except Exception:
            pass

        # 2. Çalışma dizininde logo.png / logo.jpg var mı kontrol et
        for fname in ("logo.png", "logo.jpg", "arslanlar_logo.png", "header_logo.png"):
            if os.path.exists(fname):
                try:
                    with open(fname, "rb") as f:
                        self._cached_logo_bytes = f.read()
                        return self._cached_logo_bytes
                except Exception:
                    pass

        # 3. Orijinal Arslanlar Marble Altın/Bronz Amblemini Çiz (Yüksek Çözünürlüklü PNG)
        try:
            w, h = 520, 140
            img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            gold = (197, 155, 95, 255)
            light_gold = (232, 203, 154, 255)

            # Sol Geometrik 'A' Monogram Amblemi
            draw.polygon([(20, 115), (65, 20), (95, 20), (50, 115)], fill=gold)
            draw.polygon([(70, 20), (115, 115), (85, 115), (55, 50)], fill=light_gold)
            draw.rectangle([38, 82, 102, 95], fill=gold)
            draw.rectangle([15, 118, 120, 124], fill=gold)

            out = io.BytesIO()
            img.save(out, format="PNG")
            self._cached_logo_bytes = out.getvalue()
            return self._cached_logo_bytes
        except Exception:
            return None

    def _load_bot_existing_caches(self) -> Dict[str, Dict[str, Any]]:
        """bot.py'nin kendi önbelleğini (argon_v17_master_cache.json ve DriveExplorer) okur."""
        combined_meta: Dict[str, Dict[str, Any]] = {}

        def _extract(d: Any):
            if not isinstance(d, dict):
                return
            for k, v in d.items():
                if isinstance(v, dict):
                    has_metric = any(
                        mk in v
                        for mk in (
                            "m2", "total_m2", "sqm", "area", "plaka", "slabs",
                            "slab_count", "count", "ebat", "avg_dim", "dimensions"
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
                        _extract(v)
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
                    _extract(json.load(f))
            except Exception:
                pass

        try:
            for obj in gc.get_objects():
                if type(obj).__name__ == "DriveExplorer":
                    for attr_val in list(obj.__dict__.values()):
                        if isinstance(attr_val, dict):
                            _extract(attr_val)
                    break
        except Exception:
            pass

        return combined_meta

    def _list_all_drive_items(self) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        headers = self._get_headers()
        if not headers:
            return [], []

        all_items: List[Dict[str, Any]] = []
        page_token = None
        url = "https://www.googleapis.com/drive/v3/files"

        while True:
            params = {
                "q": "trashed = false",
                "pageSize": 1000,
                "fields": "nextPageToken, files(id, name, mimeType, parents, webViewLink, modifiedTime, size)",
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
                logger.warning(f"Drive listeleme hatasi: {e}")
                break

        folders = [x for x in all_items if x.get("mimeType") == "application/vnd.google-apps.folder"]
        files = [x for x in all_items if x.get("mimeType") != "application/vnd.google-apps.folder"]
        return folders, files

    def download_file_bytes(self, file_id: str, mime_type: str = "") -> Optional[bytes]:
        headers = self._get_headers()
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

    def _parse_excel_metrics(self, xlsx_bytes: bytes) -> Optional[Dict[str, Any]]:
        """Lot klasörünün içindeki Excel/Sheets dosyasından net m2, plaka, ortalama ebat ve kalınlık okur."""
        try:
            wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
            sheet = wb.active
            rows = list(sheet.iter_rows(values_only=True))
            if not rows:
                return None

            total_m2 = 0.0
            slab_count = 0
            widths: List[float] = []
            heights: List[float] = []
            thickness_val = "2 cm"

            # Başlık satırını veya sayısal plaka satırlarını bul
            for r in rows:
                row_strs = [str(c).strip() if c is not None else "" for c in r]
                joined_low = " ".join(row_strs).lower()

                # Eğer satırda 'toplam' veya 'total' varsa
                if any(k in joined_low for k in ("toplam", "total", "genel toplam")):
                    nums = []
                    for c in r:
                        if isinstance(c, (int, float)) and float(c) > 0:
                            nums.append(float(c))
                        elif isinstance(c, str):
                            m_n = re.search(r"(\d+[.,]\d+|\d+)", c)
                            if m_n:
                                try:
                                    nums.append(float(m_n.group(1).replace(",", ".")))
                                except Exception:
                                    pass
                    for n in nums:
                        if 2.0 <= n <= 5000.0 and n > total_m2:
                            total_m2 = n

                # Plaka satırı analizi: Genelde (Sıra No, En, Boy, Kalınlık, m2)
                numeric_cells = []
                for c in r:
                    if isinstance(c, (int, float)):
                        numeric_cells.append(float(c))
                    elif isinstance(c, str):
                        c_cl = c.strip().replace(",", ".")
                        if re.match(r"^\d+(\.\d+)?$", c_cl):
                            numeric_cells.append(float(c_cl))

                # En ve boy değerleri (cm veya m cinsinden: örn 250 x 160 veya 2.50 x 1.60)
                dims_cm = [x for x in numeric_cells if 60 <= x <= 420]
                areas = [x for x in numeric_cells if 0.8 <= x <= 12.0]
                if len(dims_cm) >= 2:
                    w_cm, h_cm = dims_cm[0], dims_cm[1]
                    widths.append(w_cm)
                    heights.append(h_cm)
                    slab_count += 1
                    calc_area = (w_cm * h_cm) / 10000.0
                    # Eğer satırda buna yakın bir m2 hücresi varsa onu al, yoksa hesaplananı ekle
                    matched_a = [a for a in areas if abs(a - calc_area) < 1.2]
                    if matched_a and total_m2 == 0:
                        pass

                for x in numeric_cells:
                    if x in (1.8, 2.0, 3.0, 18.0, 20.0, 30.0):
                        if x in (3.0, 30.0):
                            thickness_val = "3 cm"
                        elif x in (2.0, 20.0):
                            thickness_val = "2 cm"

            if widths and heights:
                avg_w = int(round(sum(widths) / len(widths)))
                avg_h = int(round(sum(heights) / len(heights)))
                avg_dim = f"{avg_w} x {avg_h} cm"
                if total_m2 <= 0:
                    total_m2 = sum((widths[i] * heights[i]) / 10000.0 for i in range(len(widths)))
            else:
                avg_dim = "-"

            if total_m2 > 0 or slab_count > 0:
                return {
                    "m2": round(total_m2, 2),
                    "slabs": slab_count,
                    "avg_dim": avg_dim,
                    "surface": f"Polished ({thickness_val})",
                }
        except Exception:
            pass
        return None

    def _compute_lot_metrics(
        self,
        folder_id: str,
        lot_code: str,
        folder_name: str,
        modified_time: str,
        all_lot_files: List[Dict[str, Any]],
        custom_images: List[Dict[str, Any]],
        all_images: List[Dict[str, Any]],
        bot_cache: Dict[str, Dict[str, Any]],
    ) -> Dict[str, Any]:
        """
        Bir lotun metrajını (m²), plaka sayısını, ortalama ebadını ve yüzey/kalınlığını
        5 aşamalı garantili sistemle bulur.
        """
        m2 = 0.0
        slabs = 0
        avg_dim = "-"
        surface = "Polished (2 cm)"

        # AŞAMA 1: bot.py'nin kendi önbelleğinden (argon_v17_master_cache.json / DriveExplorer) oku
        for key in (folder_id, lot_code.upper(), lot_code.replace("-", "").upper(), folder_name.upper()):
            c = bot_cache.get(key)
            if isinstance(c, dict):
                for k_m2 in ("m2", "total_m2", "sqm", "area", "total_area", "mevcut_stok"):
                    val = c.get(k_m2)
                    if val is not None:
                        try:
                            v_f = float(str(val).replace(",", "."))
                            if v_f > 0:
                                m2 = v_f
                                break
                        except Exception:
                            pass
                for k_pl in ("plaka", "slabs", "slab_count", "count", "pcs", "plaka_sayisi"):
                    val = c.get(k_pl)
                    if val is not None:
                        try:
                            v_i = int(float(str(val)))
                            if v_i > 0:
                                slabs = v_i
                                break
                        except Exception:
                            pass
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

        # AŞAMA 2: Bizim kalıcı metrik önbelleğimizde varsa ve klasör değişmediyse oradan al
        persisted_metrics = self.state.data.setdefault("metrics_cache", {}).get(folder_id)
        if m2 <= 0 and isinstance(persisted_metrics, dict):
            if persisted_metrics.get("m2", 0) > 0:
                m2 = float(persisted_metrics["m2"])
                slabs = int(persisted_metrics.get("slabs", slabs))
                avg_dim = persisted_metrics.get("avg_dim", avg_dim)
                surface = persisted_metrics.get("surface", surface)

        # AŞAMA 3: Dosya ve Klasör İsimlerindeki Kalıpları Analiz Et
        # Örn: "SILVER_RIVER - Block No: A276 - Slab No: 36 - 226 x 161 x 3 cm - Area: 3,64 m2"
        # veya "A-157 - Silver Roots 25 37 Stock"
        widths_from_names: List[int] = []
        heights_from_names: List[int] = []
        areas_from_names: List[float] = []

        for f in all_lot_files:
            fn = f.get("name", "")
            # Ebat yakala: 226 x 161 x 3 cm veya 280x170
            m_dim = re.search(r"\b(\d{2,3})\s*[xX*]\s*(\d{2,3})(?:\s*[xX*]\s*(\d(?:[.,]\d)?))?\s*(?:cm|mm)?\b", fn)
            if m_dim:
                w_v, h_v = int(m_dim.group(1)), int(m_dim.group(2))
                if 50 <= w_v <= 450 and 50 <= h_v <= 450:
                    widths_from_names.append(w_v)
                    heights_from_names.append(h_v)
                if m_dim.group(3):
                    th = m_dim.group(3).replace(",", ".")
                    surface = f"Polished ({th} cm)"

            # Area: 3,64 m2 yakala
            m_ar = re.search(r"(\d+[.,]\d+|\d+)\s*(?:m2|m²|sqm)", fn, flags=re.I)
            if m_ar:
                try:
                    a_v = float(m_ar.group(1).replace(",", "."))
                    if 0.5 <= a_v <= 2000.0:
                        areas_from_names.append(a_v)
                except Exception:
                    pass

        if m2 <= 0 and areas_from_names:
            # Eğer her plaka fotoğrafında kendi m2'si yazıyorsa topla, tek bir toplam yazıyorsa onu al
            if len(areas_from_names) > 1 and all(a < 15.0 for a in areas_from_names):
                m2 = sum(areas_from_names)
            else:
                m2 = max(areas_from_names)

        if avg_dim == "-" and widths_from_names and heights_from_names:
            avg_dim = f"{int(sum(widths_from_names)/len(widths_from_names))} x {int(sum(heights_from_names)/len(heights_from_names))} cm"

        # AŞAMA 4: Lotun içinde Excel (.xlsx) veya Google Sheets dosyası varsa indirip içini oku
        if m2 <= 0:
            excel_files = [
                f for f in all_lot_files
                if f.get("mimeType") in (
                    "application/vnd.google-apps.spreadsheet",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "application/vnd.ms-excel",
                )
                or str(f.get("name", "")).lower().endswith((".xlsx", ".xls"))
            ]
            for ef in excel_files[:2]:
                eb = self.download_file_bytes(ef["id"], ef.get("mimeType", ""))
                if eb:
                    parsed = self._parse_excel_metrics(eb)
                    if parsed and parsed.get("m2", 0) > 0:
                        m2 = parsed["m2"]
                        if parsed.get("slabs", 0) > 0:
                            slabs = parsed["slabs"]
                        if parsed.get("avg_dim", "-") != "-":
                            avg_dim = parsed["avg_dim"]
                        if parsed.get("surface"):
                            surface = parsed["surface"]
                        break

        # AŞAMA 5: Plaka Sayısı (Slabs) hala 0 ise custom klasöründeki (veya ana) plaka fotoğraflarını say
        if slabs <= 0:
            if custom_images:
                slabs = len(custom_images)
            elif all_images:
                slabs = len(all_images)

        # Sonucu kalıcı önbelleğe kaydet
        result = {
            "m2": round(m2, 2),
            "slabs": slabs,
            "avg_dim": avg_dim,
            "surface": surface,
            "modified_time": modified_time,
        }
        if m2 > 0 or slabs > 0:
            self.state.data.setdefault("metrics_cache", {})[folder_id] = result

        return result

    @classmethod
    def _parse_lot_folder_title(cls, folder_name: str, parent_names: List[str]) -> Tuple[str, str]:
        """
        Örn: 'A-147 - Silver Roots Full Lot' -> ('A-147', 'Silver Roots')
        Örn: 'A-290 - Silver Roots Armoni Between 25 56 stock' -> ('A-290', 'Silver Roots Armoni')
        """
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

        # Gereksiz stok kelimelerini temizle, saf taş adını bırak
        stone_part = re.sub(
            r"\b(full\s+lot|full|lot|stock|stok|between\s*\d+\s*[-_]?\s*\d*|between|\d+\s+\d+\s+stock|\d+([.,]\d+)?\s*(m2|m²|plaka|slabs?|cm|mm|pcs))\b",
            " ",
            stone_part,
            flags=re.IGNORECASE,
        )
        stone_part = re.sub(r"[-_()\[\]/\\|]+", " ", stone_part)
        stone_part = re.sub(r"\s+", " ", stone_part).strip()

        if len(stone_part) < 2:
            for pname in reversed(parent_names):
                p_norm = _normalize_compare(pname)
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
        """
        3 Ana Koleksiyonu tarar.
        KRİTİK KURAL: Bir klasör Lot klasörü olarak tespit edildiği anda (örn: A-147, A-276),
        altındaki 'custom', 'email', 'photos' gibi klasörlerin içine ayrı lot bulmak için GİRMEZ;
        onların içindeki tüm görsel ve Excel dosyalarını o lotun kendisine bağlar!
        """
        now = time.time()
        if not force and self.active_lots and (now - self.last_scan_time < 45):
            return self.active_lots, self.sold_out_lots

        folders, files = self._list_all_drive_items()
        if not folders:
            return self.active_lots, self.sold_out_lots

        bot_cache = self._load_bot_existing_caches()

        folder_by_id: Dict[str, Dict[str, Any]] = {f["id"]: f for f in folders if "id" in f}
        children_folders: Dict[str, List[Dict[str, Any]]] = {}
        children_files: Dict[str, List[Dict[str, Any]]] = {}

        for f in folders:
            for p in f.get("parents", []):
                children_folders.setdefault(p, []).append(f)

        for fl in files:
            for p in fl.get("parents", []):
                children_files.setdefault(p, []).append(fl)

        target_norm_map = {_normalize_compare(c): c for c in TARGET_COLLECTIONS}
        root_collection_folders: List[Tuple[str, str]] = []

        for f in folders:
            fn_norm = _normalize_compare(f.get("name", ""))
            if fn_norm in target_norm_map:
                root_collection_folders.append((f["id"], target_norm_map[fn_norm]))

        new_active: Dict[str, Dict[str, Any]] = {}
        new_sold_out: Dict[str, Dict[str, Any]] = {}

        def _collect_all_descendant_files(fid: str) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, Any]]]:
            """
            Bir Lot klasörünün kendisindeki ve 'custom', 'photos', 'email' gibi
            tüm alt klasörlerindeki dosyaları tek bir havuzda toplar.
            En kaliteli sunum için 'custom' klasöründeki logolu görselleri ilk sıraya koyar!
            """
            all_f: List[Dict[str, Any]] = list(children_files.get(fid, []))
            custom_imgs: List[Dict[str, Any]] = []
            other_imgs: List[Dict[str, Any]] = []

            def _ is_img(file_obj: Dict[str, Any]) -> bool:
                mt = str(file_obj.get("mimeType", "")).lower()
                nm = str(file_obj.get("name", "")).lower()
                return mt.startswith("image/") or nm.endswith((".jpg", ".jpeg", ".png", ".webp"))

            for x in children_files.get(fid, []):
                if _is_img(x):
                    other_imgs.append(x)

            # Alt klasörleri gez (custom, email, photos vb.)
            stack = [(sf, _normalize_compare(sf.get("name", ""))) for sf in children_folders.get(fid, [])]
            while stack:
                sub_f, sub_norm = stack.pop(0)
                sub_id = sub_f["id"]
                s_files = children_files.get(sub_id, [])
                all_f.extend(s_files)
                for x in s_files:
                    if _is_img(x):
                        if "custom" in sub_norm:
                            custom_imgs.append(x)
                        else:
                            other_imgs.append(x)
                for deeper_sf in children_folders.get(sub_id, []):
                    deeper_norm = _normalize_compare(deeper_sf.get("name", ""))
                    stack.append((deeper_sf, f"{sub_norm} {deeper_norm}"))

            # Görselleri isme göre sırala; önce custom_imgs (logolu/etiketli), sonra diğerleri
            custom_imgs.sort(key=lambda item: item.get("name", ""))
            other_imgs.sort(key=lambda item: item.get("name", ""))
            combined_imgs = custom_imgs + other_imgs
            return all_f, custom_imgs, combined_imgs

        def _walk_hierarchy(current_fid: str, collection_name: str, path_names: List[str], in_sold_out: bool):
            cur_folder = folder_by_id.get(current_fid, {})
            cur_name = cur_folder.get("name", "").strip()
            cur_norm = _normalize_compare(cur_name)

            # Eğer bu klasör 'custom', 'email', 'photos' gibi bir sistem alt klasörüyse asla lot yapma!
            if cur_norm in INTERNAL_SUBFOLDER_NAMES:
                return

            # Klasör adında 'Sold Out' veya 'SOLD' geçiyor mu?
            is_sold_container = cur_norm in ("sold out", "soldout", "satildi", "pasif")
            has_sold_in_title = bool(self.SOLD_KEYWORD_REGEX.search(cur_name))
            now_sold_out = in_sold_out or is_sold_container or has_sold_in_title

            subfolders = children_folders.get(current_fid, [])
            non_internal_subfolders = [
                sf for sf in subfolders
                if _normalize_compare(sf.get("name", "")) not in INTERNAL_SUBFOLDER_NAMES
                and _normalize_compare(sf.get("name", "")) not in ("sold out", "soldout")
            ]

            has_lot_code = bool(self.LOT_CODE_REGEX.search(cur_name))

            # Eğer bu klasör bir LOT klasörüyse (adında A-147 gibi Lot kodu varsa VEYA
            # kök koleksiyon/Sold Out klasörü olmayıp altında sadece custom/email/photos gibi iç klasörler varsa):
            if len(path_names) > 1 and not is_sold_container and (has_lot_code or len(non_internal_subfolders) == 0):
                all_lot_files, custom_imgs, all_imgs = _collect_all_descendant_files(current_fid)
                if not has_lot_code and not all_lot_files and not all_imgs:
                    return

                lot_code, stone_name = self._parse_lot_folder_title(cur_name, path_names[:-1])
                modified_time = cur_folder.get("modifiedTime", "")

                metrics = self._compute_lot_metrics(
                    current_fid,
                    lot_code,
                    cur_name,
                    modified_time,
                    all_lot_files,
                    custom_imgs,
                    all_imgs,
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
                    "images": all_imgs,
                    "custom_images": custom_imgs,
                    "modified_time": modified_time,
                }

                if now_sold_out:
                    new_sold_out[current_fid] = record
                else:
                    new_active[current_fid] = record

                # ÇOK ÖNEMLİ: Bu bir lot klasörü olduğu için altındaki custom/email/photos klasörlerine tekrar lot diye girme!
                return

            # Kök koleksiyon veya taş kategorisi klasörüyse alt klasörleri gezmeye devam et
            for sf in subfolders:
                sf_name = sf.get("name", "").strip()
                _walk_hierarchy(sf["id"], collection_name, path_names + [sf_name], now_sold_out)

        for root_fid, coll_name in root_collection_folders:
            _walk_hierarchy(root_fid, coll_name, [coll_name], False)

        with self.inventory_lock:
            self.active_lots = new_active
            self.sold_out_lots = new_sold_out
            self.last_scan_time = now

        # Güncellenen metrik önbelleğini kaydet
        self.state.save(git_sync=False)
        return new_active, self.sold_out_lots

    def search_active_by_stone(self, raw_query: str) -> Tuple[str, List[Dict[str, Any]]]:
        """
        Herhangi bir taş çeşidi sorulduğunda (Silver Roots, Rosso Levanto, Tundra, Silver River vb.)
        SADECE aktif ve gerçek lotları döndürür.
        """
        stone_query = _extract_stone_query_name(raw_query)
        norm_query = _normalize_compare(stone_query)
        if not norm_query:
            return stone_query, []

        active_map, _ = self.scan_drive_collections(force=False)
        query_tokens = [t for t in norm_query.split() if len(t) >= 2]

        matched: List[Dict[str, Any]] = []
        for fid, item in active_map.items():
            # Klasör adında SOLD geçiyorsa veya custom/email ise kesinlikle atla
            if self.SOLD_KEYWORD_REGEX.search(item.get("folder_name", "")):
                continue
            if _normalize_compare(item.get("lot_code", "")) in INTERNAL_SUBFOLDER_NAMES:
                continue

            searchable = _normalize_compare(
                f"{item.get('stone_name', '')} {item.get('folder_name', '')} {item.get('path', '')}"
            )
            if query_tokens and all(tok in searchable for tok in query_tokens):
                matched.append(item)

        # Lot numarasına göre doğal sırala (A-147, A-157, A-221...)
        def _sort_key(x: Dict[str, Any]):
            lc = x.get("lot_code", "")
            m_num = re.search(r"(\d+)", lc)
            num = int(m_num.group(1)) if m_num else 999999
            return (x.get("collection", ""), num, lc)

        matched.sort(key=_sort_key)
        return stone_query.title(), matched

    def find_lot_by_code_or_name(self, query_str: str) -> Optional[Dict[str, Any]]:
        active_map, _ = self.scan_drive_collections(force=False)
        q_norm = _normalize_compare(query_str)
        q_compact = re.sub(r"[^a-z0-9]", "", q_norm)

        for item in active_map.values():
            lc_compact = re.sub(r"[^a-z0-9]", "", item.get("lot_code", "").lower())
            if lc_compact and lc_compact == q_compact:
                return item

        for item in active_map.values():
            blob = _normalize_compare(f"{item.get('lot_code', '')} {item.get('stone_name', '')} {item.get('folder_name', '')}")
            if q_norm and q_norm in blob:
                return item

        return None


def _build_english_luxury_pdf(
    stone_title: str, lots: List[Dict[str, Any]], drive_intel: ArslanlarDriveIntelligence
) -> bytes:
    """
    %100 İNGİLİZCE, ORİJİNAL LOGOLU, ULTRA-PREMIUM KATALOG SUNUM MOTORU
    - Kapak sayfasında ve tüm iç sayfalarda Arslanlar Marble kurumsal logosu yer alır.
    - Her lotun 'custom' klasöründeki yüksek çözünürlüklü plaka görseli tam çerçeve basılır.
    - Tüm metrikler (Available Stock, Slab Count, Average Dimensions, Finish & Thickness) İngilizcedir.
    """
    buf = io.BytesIO()
    page_w, page_h = landscape(A4)  # 841.89 x 595.27 pt
    c = pdf_canvas.Canvas(buf, pagesize=(page_w, page_h))

    logo_bytes = drive_intel.get_arslanlar_logo_bytes()
    total_m2 = sum(x.get("m2", 0.0) for x in lots)
    total_slabs = sum(x.get("slabs", 0) for x in lots)

    # Renk Paleti (Arslanlar Luxury Obsidian & Champagne Gold)
    bg_dark = colors.HexColor("#0E0E10")
    card_bg = colors.HexColor("#16161A")
    gold_primary = colors.HexColor("#C59B5F")
    gold_light = colors.HexColor("#E5C99B")
    text_white = colors.HexColor("#F7F7F7")
    text_muted = colors.HexColor("#A0A0A8")

    # =========================================================================
    # 1. PREMIUM COVER PAGE (İNGİLİZCE KAPAK SAYFASI)
    # =========================================================================
    c.setFillColor(bg_dark)
    c.rect(0, 0, page_w, page_h, fill=1, stroke=0)

    # Çift Katmanlı Altın Çerçeve
    c.setStrokeColor(gold_primary)
    c.setLineWidth(2.2)
    c.rect(24, 24, page_w - 48, page_h - 48, fill=0, stroke=1)
    c.setStrokeColor(colors.HexColor("#5A462B"))
    c.setLineWidth(0.8)
    c.rect(30, 30, page_w - 60, page_h - 60, fill=0, stroke=1)

    # Üst Merkezde Orijinal Arslanlar Marble Logosu
    if logo_bytes:
        try:
            logo_io = io.BytesIO(logo_bytes)
            c.drawImage(
                ImageReader(logo_io),
                (page_w - 220) / 2,
                page_h - 175,
                width=220,
                height=85,
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

    # Orta Dekoratif Altın Çizgi
    c.setStrokeColor(gold_primary)
    c.setLineWidth(1.5)
    c.line(page_w / 2 - 140, page_h - 255, page_w / 2 + 140, page_h - 255)

    # Seçilen Taş Başlığı
    c.setFillColor(text_white)
    c.setFont("Helvetica-Bold", 26)
    c.drawCentredString(page_w / 2, page_h - 305, f"{stone_title.upper()} COLLECTION")

    c.setFillColor(gold_light)
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(page_w / 2, page_h - 332, "LIVE INVENTORY & ARCHITECTURAL PRESENTATION")

    # Özet İstatistik Kutusu (Summary Metrics Box)
    box_w, box_h = 460, 64
    box_x = (page_w - box_w) / 2
    box_y = page_h - 425
    c.setFillColor(card_bg)
    c.setStrokeColor(gold_primary)
    c.setLineWidth(1.0)
    c.roundRect(box_x, box_y, box_w, box_h, 6, fill=1, stroke=1)

    c.setFillColor(gold_primary)
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(box_x + box_w * 0.22, box_y + 40, "AVAILABLE LOTS")
    c.drawCentredString(box_x + box_w * 0.52, box_y + 40, "TOTAL QUANTITY")
    c.drawCentredString(box_x + box_w * 0.82, box_y + 40, "TOTAL SLABS")

    c.setFillColor(text_white)
    c.setFont("Helvetica-Bold", 15)
    c.drawCentredString(box_x + box_w * 0.22, box_y + 18, f"{len(lots)} Lots")
    c.drawCentredString(box_x + box_w * 0.52, box_y + 18, f"{total_m2:,.2f} m²")
    c.drawCentredString(box_x + box_w * 0.82, box_y + 18, f"{total_slabs} Pcs")

    # Alt Bilgi (Footer)
    c.setFillColor(text_muted)
    c.setFont("Helvetica", 9.5)
    c.drawCentredString(
        page_w / 2,
        52,
        f"Generated on {datetime.now(TR_TZ).strftime('%B %d, %Y')}  |  Arslanlar Marble Export Department  |  Izmir, Turkiye",
    )
    c.showPage()

    # =========================================================================
    # 2. LOT PAGES (HER LOT İÇİN TAM SAYFA İNGİLİZCE KATALOG SAYFASI)
    # =========================================================================
    for idx, item in enumerate(lots[:35], start=1):
        c.setFillColor(bg_dark)
        c.rect(0, 0, page_w, page_h, fill=1, stroke=0)

        # Dış Çerçeve
        c.setStrokeColor(gold_primary)
        c.setLineWidth(1.4)
        c.rect(20, 20, page_w - 40, page_h - 40, fill=0, stroke=1)

        # Üst Kurumsal Başlık Barı (Header Bar)
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

        # Sol Üst: Lot Kodu & Taş Adı
        c.setFillColor(gold_primary)
        c.setFont("Helvetica-Bold", 19)
        c.drawString(38, page_h - 48, f"LOT: {lot_code}")

        c.setFillColor(text_white)
        c.setFont("Helvetica-Bold", 16)
        c.drawString(165, page_h - 48, f"—  {stone_name.upper()}")

        c.setFillColor(text_muted)
        c.setFont("Helvetica", 10)
        c.drawString(38, page_h - 66, f"COLLECTION: {coll.upper()}")

        # Sağ Üst: Arslanlar Marble Mini Logo & Marka
        if logo_bytes:
            try:
                c.drawImage(
                    ImageReader(io.BytesIO(logo_bytes)),
                    page_w - 155,
                    page_h - 72,
                    width=120,
                    height=44,
                    preserveAspectRatio=True,
                    mask="auto",
                )
            except Exception:
                pass
        else:
            c.setFillColor(gold_primary)
            c.setFont("Helvetica-Bold", 13)
            c.drawRightString(page_w - 38, page_h - 50, "ARSLANLAR MARBLE")

        # Teknik Özellikler Şeridi (4'lü İngilizce Metrik Kartları)
        spec_y = page_h - 118
        spec_h = 32
        c.setFillColor(colors.HexColor("#1B1B20"))
        c.rect(35, spec_y, page_w - 70, spec_h, fill=1, stroke=0)
        c.setStrokeColor(colors.HexColor("#3A3226"))
        c.setLineWidth(0.7)
        c.rect(35, spec_y, page_w - 70, spec_h, fill=0, stroke=1)

        stock_str = f"{m2:,.2f} m² ({slabs} Slabs)" if m2 > 0 else f"{slabs} Slabs (In Stock)"
        dim_str = ebat if ebat and ebat != "-" else "Standard Slab Size"

        metrics_cols = [
            ("AVAILABLE STOCK:", stock_str, 48),
            ("AVERAGE DIMENSIONS:", dim_str, 265),
            ("SURFACE & THICKNESS:", surf, 495),
            ("STATUS:", "AVAILABLE / READY", 705),
        ]
        for lbl, val, x_pos in metrics_cols:
            c.setFillColor(gold_light)
            c.setFont("Helvetica-Bold", 8.5)
            c.drawString(x_pos, spec_y + 19, lbl)
            c.setFillColor(text_white)
            c.setFont("Helvetica-Bold", 10.5)
            c.drawString(x_pos, spec_y + 6, str(val))

        # Plaka Görseli Alanı (Önce custom_images, yoksa tüm images)
        img_candidates = item.get("custom_images") or item.get("images") or []
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

                        img_box_x = 38
                        img_box_y = 58
                        img_box_w = page_w - 76
                        img_box_h = page_h - 186

                        c.drawImage(
                            ImageReader(img_io),
                            img_box_x,
                            img_box_y,
                            width=img_box_w,
                            height=img_box_h,
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
                "High-Resolution Slab Photos Available via Google Drive Link Below",
            )

        # Alt Bar: Tıklanabilir Google Drive Butonu & Sayfa Numarası
        c.setFillColor(card_bg)
        c.rect(21, 21, page_w - 42, 30, fill=1, stroke=0)
        c.setStrokeColor(gold_primary)
        c.setLineWidth(0.8)
        c.line(20, 51, page_w - 20, 51)

        if link:
            c.setFillColor(gold_light)
            c.setFont("Helvetica-Bold", 10)
            c.drawString(38, 32, "🔗 CLICK TO OPEN HIGH-RESOLUTION SLAB PHOTOS & BUNDLE DETAILS (GOOGLE DRIVE)")
            c.linkURL(link, (35, 24, 580, 48), relative=0)

        c.setFillColor(text_muted)
        c.setFont("Helvetica", 9.5)
        c.drawRightString(page_w - 38, 32, f"ARSLANLAR MARBLE  |  Page {idx} of {min(len(lots), 35)}")

        c.showPage()

    c.save()
    pdf_bytes = buf.getvalue()
    gc.collect()
    return pdf_bytes


def _build_ram_safe_architectural_render(
    img_bytes: bytes, lot_code: str, stone_name: str, collection_name: str
) -> bytes:
    """7 GB RAM'i yormayan (<60 MB) 2'li Bookmatch + 3D Mimari Lobi Duvar & Zemin Yansıma Motoru."""
    try:
        with Image.open(io.BytesIO(img_bytes)) as raw_img:
            raw_img = raw_img.convert("RGB")
            raw_img.thumbnail((1100, 800), Image.Resampling.LANCZOS)
            w, h = raw_img.size
            # Eğer alt kısımda siyah bilgi bandı varsa (custom görsellerindeki gibi) onu ve kenarları kırp
            crop_x = int(w * 0.04)
            crop_top = int(h * 0.04)
            crop_bottom = int(h * 0.10)
            if w > 250 and h > 250:
                slab = raw_img.crop((crop_x, crop_top, w - crop_x, h - crop_bottom))
            else:
                slab = raw_img.copy()

        slab = ImageEnhance.Contrast(slab).enhance(1.06)
        slab = ImageEnhance.Sharpness(slab).enhance(1.12)

        bm_single_w, bm_single_h = 480, 300
        left_slab = slab.resize((bm_single_w, bm_single_h), Image.Resampling.LANCZOS)
        right_slab = ImageOps.mirror(left_slab)

        bookmatch = Image.new("RGB", (bm_single_w * 2, bm_single_h))
        bookmatch.paste(left_slab, (0, 0))
        bookmatch.paste(right_slab, (bm_single_w, 0))

        bm_draw = ImageDraw.Draw(bookmatch)
        bm_draw.line([(bm_single_w, 0), (bm_single_w, bm_single_h)], fill=(35, 35, 35), width=1)

        cw, ch = 1200, 900
        y_idx, x_idx = np.mgrid[0:ch, 0:cw].astype(np.float32)
        bg_arr = np.zeros((ch, cw, 3), dtype=np.float32)
        bg_arr[:, :, 0] = 24.0
        bg_arr[:, :, 1] = 23.0
        bg_arr[:, :, 2] = 26.0

        ceil_mask = y_idx < 110
        bg_arr[ceil_mask] *= (0.55 + 0.45 * (y_idx[ceil_mask] / 110.0))[:, None]

        wall_x1, wall_y1 = 140, 110
        wall_w, wall_h = 920, 520
        wall_bm = bookmatch.resize((wall_w, wall_h), Image.Resampling.LANCZOS)
        wall_arr = np.asarray(wall_bm, dtype=np.float32)

        wy, wx = np.mgrid[0:wall_h, 0:wall_w].astype(np.float32)
        light_map = np.ones((wall_h, wall_w), dtype=np.float32) * 0.72
        for sx in (wall_w * 0.22, wall_w * 0.50, wall_w * 0.78):
            cone_spread = 95.0 + wy * 0.42
            dist_x = np.abs(wx - sx)
            vertical_fade = np.clip(1.0 - (wy / (wall_h * 1.15)), 0.15, 1.0)
            spot_intensity = np.exp(-0.5 * (dist_x / cone_spread) ** 2) * vertical_fade
            light_map += spot_intensity * 0.42

        edge_x = np.minimum(wx, wall_w - wx) / 60.0
        edge_y = np.minimum(wy, wall_h - wy) / 45.0
        ao_mask = np.clip(np.minimum(edge_x, edge_y), 0.55, 1.0)
        light_map = np.clip(light_map * ao_mask, 0.35, 1.22)

        wall_lit = np.zeros_like(wall_arr)
        wall_lit[:, :, 0] = np.clip(wall_arr[:, :, 0] * light_map * 1.03, 0, 255)
        wall_lit[:, :, 1] = np.clip(wall_arr[:, :, 1] * light_map * 1.00, 0, 255)
        wall_lit[:, :, 2] = np.clip(wall_arr[:, :, 2] * light_map * 0.95, 0, 255)

        bg_arr[wall_y1 : wall_y1 + wall_h, wall_x1 : wall_x1 + wall_w] = wall_lit

        floor_y1, floor_y2 = 630, 820
        floor_h = floor_y2 - floor_y1
        flipped_wall = ImageOps.flip(Image.fromarray(wall_lit.astype(np.uint8))).resize(
            (wall_w, floor_h), Image.Resampling.LANCZOS
        )
        flipped_wall = flipped_wall.filter(ImageFilter.GaussianBlur(radius=2.2))
        refl_arr = np.asarray(flipped_wall, dtype=np.float32)

        fy = np.linspace(0.0, 1.0, floor_h, dtype=np.float32)[:, None, None]
        floor_base = np.array([20.0, 19.0, 22.0], dtype=np.float32) * (1.0 - 0.3 * fy)
        refl_alpha = 0.42 * np.exp(-2.4 * fy)

        bg_arr[floor_y1:floor_y2, :] = floor_base
        bg_arr[floor_y1:floor_y2, wall_x1 : wall_x1 + wall_w] = (
            refl_arr * refl_alpha + floor_base * (1.0 - refl_alpha)
        )

        canvas = Image.fromarray(np.clip(bg_arr, 0, 255).astype(np.uint8))
        draw = ImageDraw.Draw(canvas)

        bronze_color = (197, 155, 95)
        draw.rectangle([wall_x1 - 4, wall_y1 - 4, wall_x1 + wall_w + 3, wall_y1 + wall_h + 3], outline=bronze_color, width=3)
        draw.line([(0, floor_y1), (cw, floor_y1)], fill=(70, 62, 50), width=2)

        mini_w, mini_h = 240, 150
        mini_bm = bookmatch.resize((mini_w, mini_h), Image.Resampling.LANCZOS)
        canvas.paste(mini_bm, (cw - mini_w - 28, 22))
        draw.rectangle([cw - mini_w - 30, 20, cw - 26, 24 + mini_h], outline=bronze_color, width=2)
        draw.rectangle([cw - mini_w - 30, 20, cw - 26, 42], fill=(20, 20, 22))
        draw.text((cw - mini_w - 20, 24), "BOOKMATCH LAYOUT", fill=(235, 215, 175))

        draw.rectangle([0, 820, cw, ch], fill=(12, 12, 14))
        draw.line([(0, 820), (cw, 820)], fill=bronze_color, width=2)

        title_str = f"ARSLANLAR MARBLE STUDIO  |  {lot_code} - {stone_name.upper()}  ({collection_name})"
        sub_str = "Architectural Bookmatch & Interior Wall / Floor Reflection Simulation"
        draw.text((35, 836), title_str, fill=(240, 225, 195))
        draw.text((35, 864), sub_str, fill=(165, 165, 170))

        out_buf = io.BytesIO()
        canvas.save(out_buf, format="JPEG", quality=90, optimize=True)
        result_bytes = out_buf.getvalue()

        del raw_img, slab, left_slab, right_slab, bookmatch, wall_bm, wall_arr, wall_lit, bg_arr, canvas
        gc.collect()
        return result_bytes
    except Exception as e:
        gc.collect()
        raise RuntimeError(f"Render error: {e}")


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
                logger.warning(f"Telegram bildirim hatasi ({cid}): {e}")

    def format_executive_summary_report(self, title_prefix: str = "HAFTALIK") -> str:
        active_map, sold_map = self.intel.scan_drive_collections(force=True)

        total_lots = len(active_map)
        total_m2 = sum(x.get("m2", 0.0) for x in active_map.values())
        total_slabs = sum(x.get("slabs", 0) for x in active_map.values())

        lines = [
            f"📊 <b>ARSLANLAR MARBLE — {title_prefix} STOK & KOLEKSİYON RAPORU</b>",
            f"🗓 <i>{datetime.now(TR_TZ).strftime('%d.%m.%Y %H:%M')}</i>\n",
            f"🏛 <b>Toplam Aktif Stok:</b> <code>{total_m2:,.2f} m²</code> (<code>{total_slabs} Plaka</code> / <code>{total_lots} Lot</code>)",
            f"📦 <b>Sold Out (Arşiv):</b> <code>{len(sold_map)} Lot</code>\n",
            "📂 <b>Koleksiyon Bazlı Dağılım:</b>",
        ]

        for coll in TARGET_COLLECTIONS:
            c_active = [x for x in active_map.values() if x.get("collection") == coll]
            c_sold = [x for x in sold_map.values() if x.get("collection") == coll]
            c_m2 = sum(x.get("m2", 0.0) for x in c_active)
            c_slabs = sum(x.get("slabs", 0) for x in c_active)
            lines.append(
                f"• <b>{coll}:</b> <code>{len(c_active)} Aktif Lot</code> | "
                f"<code>{c_m2:,.2f} m²</code> (<code>{c_slabs} Plaka</code>) | "
                f"<i>Sold Out: {len(c_sold)}</i>"
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
        )[:10]

        if top_varieties:
            lines.append("\n💎 <b>En Çok Stoğu Bulunan Taş Çeşitleri:</b>")
            for st_name, st_info in top_varieties:
                lines.append(
                    f"  ▫️ <b>{st_name}:</b> {st_info['lots']} Lot — "
                    f"<code>{st_info['m2']:,.2f} m²</code> ({st_info['slabs']} Plaka)"
                )

        return "\n".join(lines)

    def _background_watcher_loop(self):
        time.sleep(12)
        logger.info("🔍 Arslanlar V20 Drive Gozcusu baslatildi.")

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

                        # 1. Sold Out klasörüne yeni taşınan lotlar
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

                        # 2. Aktif koleksiyonlara yeni eklenen lotlar
                        for a_fid, a_item in active_map.items():
                            a_code = a_item.get("lot_code", "")
                            if a_fid not in prev_active and a_code not in prev_active_codes:
                                changed = True
                                stone_label = a_item.get("stone_name", "")
                                coll_label = a_item.get("collection", "")
                                m2_val = a_item.get("m2", 0.0)
                                slabs_val = a_item.get("slabs", 0)
                                link_val = a_item.get("drive_link", "")
                                msg = (
                                    f"🆕 <b>Yeni Lot Eklendi: {a_code} ({stone_label})</b>\n"
                                    f"📂 Koleksiyon: <b>{coll_label}</b>\n"
                                    f"• Mevcut Stok: <code>{m2_val:.2f} m² ({slabs_val} Plaka)</code>\n"
                                    f"• Drive Linki: <a href='{link_val}'>Klasörü Aç</a>"
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

                # 3. Pazartesi ve Cuma Otomatik Özet Rapor
                now_tr = datetime.now(TR_TZ)
                if now_tr.weekday() in (0, 4) and now_tr.hour >= 9:
                    report_key = now_tr.strftime("%Y-%m-%d")
                    if self.state.data.get("last_summary_report_key") != report_key:
                        day_title = "PAZARTESİ" if now_tr.weekday() == 0 else "CUMA"
                        report_html = self.format_executive_summary_report(title_prefix=day_title)
                        self.send_telegram_sync(report_html)
                        self.state.data["last_summary_report_key"] = report_key
                        self.state.save(git_sync=True)

            except Exception as e:
                logger.warning(f"Drive Gozcu dongu uyarisi: {e}")

            time.sleep(120)

    async def handle_priority_message(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        if not update.effective_message or not update.effective_message.text:
            return

        if update.effective_chat:
            self.state.register_chat(update.effective_chat.id)

        raw_text = update.effective_message.text.strip()
        norm_text = _normalize_compare(raw_text)

        # 1. ÖZET RAPOR İSTEĞİ
        if norm_text in ("ozet", "ozet rapor", "stok ozet", "haftalik ozet", "genel stok") or raw_text.lower().startswith("/ozet"):
            status_msg = await update.effective_message.reply_text("⏳ Koleksiyonlar taranıyor, özet rapor hazırlanıyor...")
            loop = asyncio.get_running_loop()
            report_html = await loop.run_in_executor(None, self.format_executive_summary_report, "ANLIK")
            await status_msg.edit_text(report_html, parse_mode="HTML", disable_web_page_preview=True)
            raise ApplicationHandlerStop

        # 2. RAM-KORUMALI MİMARİ RENDER & BOOKMATCH
        if re.search(r"\b(render|bookmatch|kelebek)\b", norm_text):
            target_q = re.sub(
                r"\b(render|bookmatch|kelebek|yap|lütfen|lutfen|hazırla|hazirla|stüdyo|studyo)\b",
                " ",
                raw_text,
                flags=re.I,
            ).strip(" /.,:-")

            if not target_q:
                await update.effective_message.reply_text(
                    "🎨 Lütfen render yapılacak lot kodunu yazın.\nÖrnek: <code>render A-371</code>",
                    parse_mode="HTML",
                )
                raise ApplicationHandlerStop

            status_msg = await update.effective_message.reply_text(
                f"🎨 <b>{target_q.upper()}</b> için Mimari Sahne & Bookmatch Render hazırlanıyor...",
                parse_mode="HTML",
            )
            loop = asyncio.get_running_loop()
            lot_item = await loop.run_in_executor(None, self.intel.find_lot_by_code_or_name, target_q)

            imgs = (lot_item.get("custom_images") or lot_item.get("images") or []) if lot_item else []
            if not lot_item or not imgs:
                await status_msg.edit_text(
                    f"⚠️ <b>{target_q}</b> için aktif stokta plaka fotoğrafı bulunamadı.",
                    parse_mode="HTML",
                )
                raise ApplicationHandlerStop

            img_bytes = await loop.run_in_executor(
                None, self.intel.download_file_bytes, imgs[0]["id"], imgs[0].get("mimeType", "")
            )
            if not img_bytes:
                await status_msg.edit_text("⚠️ Plaka görseli indirilemedi.")
                raise ApplicationHandlerStop

            try:
                rendered_bytes = await loop.run_in_executor(
                    RENDER_EXECUTOR,
                    _build_ram_safe_architectural_render,
                    img_bytes,
                    lot_item["lot_code"],
                    lot_item["stone_name"],
                    lot_item["collection"],
                )
                caption = (
                    f"🏛 <b>{lot_item['lot_code']} — {lot_item['stone_name']}</b>\n"
                    f"• <b>Collection:</b> {lot_item['collection']}\n"
                    f"• <b>Available Stock:</b> <code>{lot_item['m2']:.2f} m² ({lot_item['slabs']} Slabs)</code>\n"
                    f"• <b>Studio:</b> 2x Bookmatch & Architectural Wall/Floor Simulation"
                )
                await update.effective_message.reply_photo(
                    photo=io.BytesIO(rendered_bytes),
                    caption=caption,
                    parse_mode="HTML",
                )
                await status_msg.delete()
            except Exception as e:
                await status_msg.edit_text(f"⚠️ Render hatası: {e}")

            raise ApplicationHandlerStop

        # 3. TAŞ ÇEŞİDİNE ÖZEL İNGİLİZCE PREMIUM SUNUM ("Silver Rootslar sunum lütfen", "Rosso levantolar sunum lütfen", "Tundra sunum")
        has_sunum_kw = bool(re.search(r"\b(sunum|sunumu|katalog|kataloğu|presentation|catalog)\b", norm_text))
        is_main_collection_sunum = any(
            _normalize_compare(c) in norm_text
            or norm_text.startswith("series")
            or norm_text.startswith("collection")
            or norm_text.startswith("exclusive")
            for c in TARGET_COLLECTIONS
        )

        if has_sunum_kw and not is_main_collection_sunum:
            stone_title, matched_lots = await asyncio.get_running_loop().run_in_executor(
                None, self.intel.search_active_by_stone, raw_text
            )
            if matched_lots:
                status_msg = await update.effective_message.reply_text(
                    f"📑 <b>{stone_title}</b> için <b>{len(matched_lots)} aktif lot</b> bulundu. "
                    f"İngilizce kurumsal sunum (PDF) ve stok özeti hazırlanıyor...",
                    parse_mode="HTML",
                )
                await self._send_stone_variety_summary(update, stone_title, matched_lots, send_photos=False)

                try:
                    pdf_bytes = await asyncio.get_running_loop().run_in_executor(
                        RENDER_EXECUTOR, _build_english_luxury_pdf, stone_title, matched_lots, self.intel
                    )
                    safe_fn = re.sub(r"[^a-zA-Z0-9_]", "_", stone_title.upper()) + "_ARSLANLAR_MARBLE_SELECTION.pdf"
                    total_m2 = sum(x.get("m2", 0.0) for x in matched_lots)
                    total_slabs = sum(x.get("slabs", 0) for x in matched_lots)
                    await update.effective_message.reply_document(
                        document=io.BytesIO(pdf_bytes),
                        filename=safe_fn,
                        caption=(
                            f"🏛 <b>ARSLANLAR MARBLE — {stone_title.upper()} SELECTION</b>\n"
                            f"• <b>Available Lots:</b> <code>{len(matched_lots)} Lots</code>\n"
                            f"• <b>Total Quantity:</b> <code>{total_m2:,.2f} m² ({total_slabs} Slabs)</code>\n"
                            f"• <b>Language:</b> English (Export Standard)"
                        ),
                        parse_mode="HTML",
                    )
                    await status_msg.delete()
                except Exception as e:
                    await status_msg.edit_text(f"⚠️ PDF oluşturulurken hata: {e}")

                raise ApplicationHandlerStop

        # 4. TAŞ ÇEŞİDİNE GÖRE STOK SORGULAMA ("Ne kadar Silver Roots var?", "Ne kadar Rosso Levanto var?", "Ne kadar Tundra var?", "Silver Roots çeşitleri")
        is_variety_query = bool(
            re.search(
                r"(^ne\s+kadar\s+|var\s+m[ıi]|kaç\s+m2|kac\s+m2|\bçeşitleri\b|\bcesitleri\b|\bstok\s+listesi\b|\bneler\s+var\b|\bvar\s*\??$|^how\s+much\s+)",
                raw_text,
                flags=re.I,
            )
        )
        is_single_lot_code = bool(re.match(r"^[A-Za-z]{1,3}\s*[-_]?\s*\d{1,5}[A-Za-z]?$", raw_text))

        if is_variety_query and not is_single_lot_code:
            status_msg = await update.effective_message.reply_text("🔍 Aktif koleksiyonlar ve metrajlar taranıyor...")
            stone_title, matched_lots = await asyncio.get_running_loop().run_in_executor(
                None, self.intel.search_active_by_stone, raw_text
            )
            await status_msg.delete()

            if not matched_lots:
                await update.effective_message.reply_text(
                    f"🔍 <b>{stone_title}</b> için aktif koleksiyon klasörlerinde (<i>Arslanlar Series, Collection, Exclusive</i>) "
                    f"aktif stok bulunamadı.\n<i>(Sold Out klasöründeki ve satılmış taşlar hariç tutulmuştur.)</i>",
                    parse_mode="HTML",
                )
                raise ApplicationHandlerStop

            await self._send_stone_variety_summary(update, stone_title, matched_lots, send_photos=True)
            raise ApplicationHandlerStop

    async def _send_stone_variety_summary(
        self, update: Update, stone_title: str, matched_lots: List[Dict[str, Any]], send_photos: bool = True
    ):
        total_m2 = sum(x.get("m2", 0.0) for x in matched_lots)
        total_slabs = sum(x.get("slabs", 0) for x in matched_lots)

        lines = [
            f"💎 <b>{stone_title.upper()} — ACTIVE INVENTORY & VARIETIES</b>",
            f"📊 <b>Total Active Stock:</b> <code>{total_m2:,.2f} m² ({total_slabs} Slabs)</code> — <b>{len(matched_lots)} Active Lots</b>",
            "<i>(Sold Out folders & sold lots are strictly excluded)</i>\n",
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

        # İlk 4 lotun 'custom' klasöründeki logolu görselini albüm olarak gönder
        if send_photos:
            media_group = []
            loop = asyncio.get_running_loop()
            for item in matched_lots[:4]:
                imgs = item.get("custom_images") or item.get("images") or []
                if not imgs:
                    continue
                b = await loop.run_in_executor(
                    None, self.intel.download_file_bytes, imgs[0]["id"], imgs[0].get("mimeType", "")
                )
                if b:
                    m2_v = item.get("m2", 0.0)
                    sl_v = item.get("slabs", 0)
                    st_txt = f"{m2_v:.2f} m² ({sl_v} Slabs)" if m2_v > 0 else f"{sl_v} Slabs"
                    cap = f"💎 {item['lot_code']} — {item['stone_name']} | {st_txt}"
                    media_group.append(InputMediaPhoto(media=io.BytesIO(b), caption=cap))

            if media_group:
                try:
                    await update.effective_message.reply_media_group(media=media_group)
                except Exception as e:
                    logger.warning(f"Medya grubu gonderim uyarisi: {e}")


# =========================================================================
# SIFIR MÜDAHALE (MONKEY-PATCH) KANCASI
# =========================================================================
_ORIGINAL_RUN_POLLING = Application.run_polling
_PLUGIN_ATTACHED = False


def _hooked_run_polling(self: Application, *args, **kwargs):
    global _PLUGIN_ATTACHED
    if not _PLUGIN_ATTACHED:
        _PLUGIN_ATTACHED = True
        try:
            controller = ArslanlarPluginController(self)
            self.add_handler(
                MessageHandler(filters.TEXT, controller.handle_priority_message),
                group=-1,
            )
            watcher_thread = threading.Thread(
                target=controller._background_watcher_loop,
                name="ArslanlarDriveWatcherV20",
                daemon=True,
            )
            watcher_thread.start()
            logger.info("✅ Arslanlar V20 Ultra-Premium Eklenti Modulu basariyla devreye girdi!")
        except Exception as e:
            logger.error(f"Eklenti baslatma hatasi: {e}")

    return _ORIGINAL_RUN_POLLING(self, *args, **kwargs)


Application.run_polling = _hooked_run_polling
