# -*- coding: utf-8 -*-
"""
ARSLANLAR MARBLE - V23 UNIVERSAL EXCEL & ZERO-TOUCH ENGINE (sitecustomize.py)
- ASLA alt klasorlerdeki kopya fotograflari toplayip '276 Slabs' gibi uydurma plaka sayisi yazmaz!
- Her lot klasorunun icindeki Excel tablosunu (.xlsx, .xls, Google Sheets, .csv, .ods, HTML-xls)
  5 katmanli Evrensel Excel Okuyucu + bot.py'nin kendi orijinal DriveExplorer motoru ile %100 okur.
- 'custom' klasorundeki fotograflari hicbir yerde (PDF, Telegram, Render) kullanmaz; sadece saf plaka gorseli kullanir.
- %100 Ingilizce, Orijinal Arslanlar Marble Logolu Premium PDF Sunum & Master Stok Tablosu uretir.
- bot.py dosyasina %0 mudahale eder.
"""

import os
import re
import gc
import io
import sys
import csv
import json
import time
import base64
import asyncio
import logging
import inspect
import zipfile
import threading
import subprocess
import xml.etree.ElementTree as ET
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

# Eski tip .xls dosyalarını okumak için xlrd kütüphanesi yoksa arka planda sessizce hazırla
try:
    import xlrd
except ImportError:
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "xlrd", "-q"], timeout=25, check=False)
        import xlrd
    except Exception:
        xlrd = None

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


def is_spreadsheet_or_table_file(file_obj: Dict[str, Any]) -> bool:
    """Bir dosyanın Excel / Tablo dosyası olup olmadığını geniş kapsamlı tespit eder."""
    mt = str(file_obj.get("mimeType", "")).lower()
    nm = str(file_obj.get("name", "")).lower()
    if is_image_file_obj(file_obj):
        return False
    if nm.endswith((".pdf", ".mp4", ".mov", ".avi", ".rar", ".zip", ".7z")):
        return False
    if any(
        k in mt
        for k in ("spreadsheet", "excel", "csv", "tab-separated", "opendocument.spreadsheet", "ms-excel", "officedocument")
    ):
        return True
    if nm.endswith((".xlsx", ".xls", ".xlsm", ".xlsb", ".ods", ".csv", ".tsv")):
        return True
    # Bazı Excel dosyaları Drive'a uzantısız veya 'octet-stream' olarak yüklenmiş olabilir
    if mt == "application/octet-stream" and not nm.endswith((".jpg", ".jpeg", ".png", ".pdf", ".mp4")):
        return True
    return False


def to_float_safe(val: Any) -> Optional[float]:
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return None
    # İçinde birim varsa temizle (örn: '3,64 m2' -> '3,64', '314 cm' -> '314')
    s = re.sub(r"(?:m2|m²|sqm|cm|mm|pcs|adet|plaka|slabs?)", "", s, flags=re.I).strip()
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
            "version": 23,
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
                            # Eski sürümden kalan bozuk metrik önbelleğini sıfırla
                            if loaded.get("version") != 23:
                                loaded["metrics_cache"] = {}
                                loaded["version"] = 23
                            else:
                                # İçinde m2 <= 0 olan hatalı kayıtları temizle
                                mc = loaded.get("metrics_cache", {})
                                loaded["metrics_cache"] = {
                                    k: v for k, v in mc.items()
                                    if isinstance(v, dict) and v.get("m2", 0.0) > 0 and v.get("avg_dim", "-") != "-"
                                }
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
            self.data["version"] = 23
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
                ["git", "commit", "-m", "chore: auto-sync watcher state v23 [skip ci]"],
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


class UniversalExcelParser:
    """
    Her türlü Excel / Tablo formatını (.xlsx, .xls BIFF8, XML/HTML .xls, .ods, .csv)
    2D satır-sütun matrisine çeviren ve içinden gerçek Mermer Metrajı (m²),
    Gerçek Plaka Sayısı (Slabs), Ortalama Ebat (cm) ve Kalınlık çıkaran evrensel motor.
    """

    @staticmethod
    def extract_sheets_as_rows(file_bytes: bytes, filename: str = "") -> List[List[List[Any]]]:
        sheets: List[List[List[Any]]] = []

        # 1. YÖNTEM: openpyxl (.xlsx / .xlsm)
        try:
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            for ws in wb.worksheets:
                rows = [list(r) for r in ws.iter_rows(values_only=True)]
                if rows:
                    sheets.append(rows)
            if sheets:
                return sheets
        except Exception:
            pass

        # 2. YÖNTEM: openpyxl (data_only=False, eğer formüller önbelleksiz kaydedildiyse)
        try:
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
            for ws in wb.worksheets:
                rows = [list(r) for r in ws.iter_rows(values_only=True)]
                if rows:
                    sheets.append(rows)
            if sheets:
                return sheets
        except Exception:
            pass

        # 3. YÖNTEM: xlrd (Eski tip .xls Excel 97-2003 dosyaları)
        if xlrd is not None:
            try:
                book = xlrd.open_workbook(file_contents=file_bytes)
                for s_idx in range(book.nsheets):
                    sh = book.sheet_by_index(s_idx)
                    rows = []
                    for r_idx in range(sh.nrows):
                        rows.append([sh.cell_value(r_idx, c_idx) for c_idx in range(sh.ncols)])
                    if rows:
                        sheets.append(rows)
                if sheets:
                    return sheets
            except Exception:
                pass

        # 4. YÖNTEM: ZIP tabanlı XML / OpenDocument (.ods veya bozuk .xlsx)
        try:
            if zipfile.is_zipfile(io.BytesIO(file_bytes)):
                with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
                    namelist = zf.namelist()
                    # OpenDocument (.ods) content.xml
                    if "content.xml" in namelist:
                        xml_data = zf.read("content.xml")
                        root = ET.fromstring(xml_data)
                        rows = []
                        for tr in root.iter():
                            if tr.tag.endswith("table-row"):
                                row_vals = []
                                for tc in tr:
                                    if tc.tag.endswith("table-cell"):
                                        txt = "".join(tc.itertext()).strip()
                                        row_vals.append(txt)
                                if any(row_vals):
                                    rows.append(row_vals)
                        if rows:
                            sheets.append(rows)
                            return sheets
        except Exception:
            pass

        # 5. YÖNTEM: CSV / TSV / HTML Tablo olarak kaydedilmiş .xls dosyaları
        for enc in ("utf-8-sig", "utf-8", "cp1254", "latin-1", "utf-16"):
            try:
                text = file_bytes.decode(enc)
                if "<table" in text.lower() or "<tr" in text.lower():
                    tr_matches = re.findall(r"<tr[^>]*>(.*?)</tr>", text, flags=re.I | re.S)
                    rows = []
                    for tr_m in tr_matches:
                        cells = re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr_m, flags=re.I | re.S)
                        clean_cells = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
                        if any(clean_cells):
                            rows.append(clean_cells)
                    if rows:
                        sheets.append(rows)
                        return sheets

                lines = [ln for ln in text.splitlines() if ln.strip()]
                if len(lines) >= 2:
                    delim = ";" if lines[0].count(";") >= lines[0].count(",") else ","
                    if "\t" in lines[0]:
                        delim = "\t"
                    reader = csv.reader(lines, delimiter=delim)
                    rows = [list(r) for r in reader if any(r)]
                    if len(rows) >= 2 and any(len(r) >= 2 for r in rows):
                        sheets.append(rows)
                        return sheets
            except Exception:
                continue

        return sheets

    @classmethod
    def analyze_table_rows(cls, sheets: List[List[List[Any]]]) -> Optional[Dict[str, Any]]:
        """
        Çıkarılan Excel sayfasındaki satır ve sütunları tarayarak:
        - Toplam Metraj (m²)
        - Gerçek Plaka Adedi (Slabs)
        - Ortalama Ebat (En x Boy cm)
        - Kalınlık (2 cm / 3 cm)
        bilgilerini sıfır hatayla hesaplar.
        """
        best_m2 = 0.0
        best_slabs = 0
        best_avg_dim = "-"
        best_thickness = "2 cm"

        for rows in sheets:
            if not rows:
                continue

            col_w_idx = None
            col_h_idx = None
            col_th_idx = None
            col_m2_idx = None
            col_no_idx = None
            header_row_idx = None

            # İlk 25 satırda başlık satırını ara
            for r_i, r in enumerate(rows[:25]):
                norm_cells = [normalize_compare(str(c)) if c is not None else "" for c in r]
                w_c, h_c, m2_c, th_c, no_c = None, None, None, None, None

                for c_i, nc in enumerate(norm_cells):
                    if not nc:
                        continue
                    if nc in ("no", "sira", "sira no", "slab no", "plaka no", "plaka", "slab", "bundle", "id", "#"):
                        no_c = c_i
                    elif any(w_kw == nc or nc.startswith(w_kw + " ") for w_kw in ("boy", "length", "uzunluk", "len", "l")):
                        w_c = c_i
                    elif any(h_kw == nc or nc.startswith(h_kw + " ") for h_kw in ("en", "width", "yukseklik", "height", "genislik", "w", "h")):
                        h_c = c_i
                    elif any(m_kw in nc for m_kw in ("m2", "m²", "sqm", "alan", "area", "metraj", "miktar", "quantity", "net m2", "toplam m2", "total m2")):
                        m2_c = c_i
                    elif any(t_kw in nc for t_kw in ("kalinlik", "thickness", "thick", "kal")):
                        th_c = c_i

                # Eğer sadece 'ebat' veya 'dimensions' veya 'size' adında iki yan yana sütun varsa
                if w_c is None and h_c is None:
                    dim_cols = [
                        idx for idx, nc in enumerate(norm_cells)
                        if any(k in nc for k in ("ebat", "olcu", "dimension", "size", "cm"))
                    ]
                    if len(dim_cols) >= 2:
                        w_c, h_c = dim_cols[0], dim_cols[1]

                if (w_c is not None and h_c is not None) or m2_c is not None:
                    col_w_idx = w_c
                    col_h_idx = h_c
                    col_m2_idx = m2_c
                    col_th_idx = th_c
                    col_no_idx = no_c
                    header_row_idx = r_i
                    break

            widths: List[float] = []
            heights: List[float] = []
            row_areas: List[float] = []
            explicit_total_m2: List[float] = []
            explicit_total_slabs: List[int] = []
            valid_slab_rows = 0

            start_r = (header_row_idx + 1) if header_row_idx is not None else 0

            for r in rows[start_r:]:
                if not r or all(c is None or str(c).strip() == "" for c in r):
                    continue

                raw_strs = [str(c).strip() if c is not None else "" for c in r]
                joined_norm = normalize_compare(" ".join(raw_strs))
                if not joined_norm:
                    continue

                # Hücre içinde '310x165' gibi birleşik ebat yazılmış mı kontrol et
                inline_w, inline_h = None, None
                for rs in raw_strs:
                    m_in = re.search(r"\b(\d{2,3})\s*[xX*]\s*(\d{2,3})\b", rs)
                    if m_in:
                        iw, ih = float(m_in.group(1)), float(m_in.group(2))
                        if 50 <= iw <= 440 and 50 <= ih <= 440:
                            inline_w, inline_h = iw, ih
                            break

                is_total_line = any(
                    kw in joined_norm
                    for kw in ("toplam", "total", "genel toplam", "grand total", "sum", "yekun", "alt toplam")
                )

                if is_total_line:
                    if col_m2_idx is not None and col_m2_idx < len(r):
                        v_m = to_float_safe(r[col_m2_idx])
                        if v_m and 1.0 <= v_m <= 5000.0:
                            explicit_total_m2.append(v_m)

                    for c in r:
                        if c is None:
                            continue
                        s_c = str(c).strip()
                        # Eğer hücrede '45,64 m2' gibi ifade varsa
                        m_u = re.search(r"(\d+[.,]\d+|\d+)\s*(?:m2|m²|sqm)", s_c, flags=re.I)
                        if m_u:
                            vf = to_float_safe(m_u.group(1))
                            if vf and 1.0 <= vf <= 5000.0:
                                explicit_total_m2.append(vf)
                        m_p = re.search(r"(\d+)\s*(?:plaka|slabs?|pcs|adet)", s_c, flags=re.I)
                        if m_p:
                             explicit_total_slabs.append(int(m_p.group(1)))

                        vf = to_float_safe(c)
                        if vf and 2.0 <= vf <= 5000.0:
                            # Ondalıklıysa m2 olma ihtimali çok yüksektir
                            if isinstance(c, float) or ("," in s_c) or ("." in s_c):
                                explicit_total_m2.append(vf)
                    continue

                # Normal veri satırı analizi
                w_val, h_val, a_val = inline_w, inline_h, None

                if col_w_idx is not None and col_h_idx is not None and col_w_idx < len(r) and col_h_idx < len(r):
                    w_c_val = to_float_safe(r[col_w_idx])
                    h_c_val = to_float_safe(r[col_h_idx])
                    if w_c_val and h_c_val:
                        # Metre olarak girildiyse (örn: 2.90 x 1.35) cm'ye çevir
                        if 0.45 <= w_c_val <= 4.50 and 0.45 <= h_c_val <= 4.50:
                            w_c_val *= 100.0
                            h_c_val *= 100.0
                        if 45.0 <= w_c_val <= 450.0 and 45.0 <= h_c_val <= 450.0:
                            w_val, h_val = w_c_val, h_c_val

                if col_m2_idx is not None and col_m2_idx < len(r):
                    a_c_val = to_float_safe(r[col_m2_idx])
                    if a_c_val and 0.3 <= a_c_val <= 20.0:
                        a_val = a_c_val
                    elif a_c_val and 3000.0 <= a_c_val <= 150000.0:
                        # cm2 olarak yazıldıysa m2'ye çevir
                        a_val = a_c_val / 10000.0

                if col_th_idx is not None and col_th_idx < len(r):
                    th_v = to_float_safe(r[col_th_idx])
                    if th_v in (2.0, 20.0):
                        best_thickness = "2 cm"
                    elif th_v in (3.0, 30.0):
                        best_thickness = "3 cm"

                # Eğer sütun başlıklarından bulunamadıysa satırdaki sayı dizilimini incele
                if w_val is None or h_val is None:
                    nums = []
                    for c in r:
                        fv = to_float_safe(c)
                        if fv is not None and fv > 0:
                            nums.append(fv)

                    # Plaka boyut adayları (cm: 50..440 veya m: 0.55..4.30)
                    cm_dims = [x for x in nums if 50.0 <= x <= 440.0]
                    m_dims = [x * 100.0 for x in nums if 0.55 <= x <= 4.30]

                    if len(cm_dims) >= 2:
                        w_val, h_val = cm_dims[0], cm_dims[1]
                    elif len(m_dims) >= 2 and len(nums) >= 3:
                        w_val, h_val = m_dims[0], m_dims[1]

                    if w_val and h_val and a_val is None:
                        calc_a = (w_val * h_val) / 10000.0
                        # Satırda bu çarpıma yakın bir m2 değeri var mı?
                        close_areas = [x for x in nums if 0.4 <= x <= 16.0 and abs(x - calc_a) <= 1.5]
                        if close_areas:
                            a_val = close_areas[-1]
                        else:
                            a_val = calc_a

                if w_val and h_val and 45.0 <= w_val <= 450.0 and 45.0 <= h_val <= 450.0:
                    widths.append(w_val)
                    heights.append(h_val)
                    valid_slab_rows += 1
                    if a_val and 0.3 <= a_val <= 18.0:
                        row_areas.append(a_val)
                    else:
                        row_areas.append((w_val * h_val) / 10000.0)
                elif a_val and 0.4 <= a_val <= 16.0:
                    row_areas.append(a_val)
                    valid_slab_rows += 1

            # Bu sayfa için toplamları değerlendir
            sheet_m2 = sum(row_areas) if row_areas else 0.0
            sheet_slabs = valid_slab_rows

            if explicit_total_m2:
                # Eğer satırlardan hesaplanan m2 varsa ve toplam satırıyla uyumluysa toplam satırını al
                if sheet_m2 > 0:
                    matching_tots = [t for t in explicit_total_m2 if abs(t - sheet_m2) <= (sheet_m2 * 0.25 + 2.0)]
                    if matching_tots:
                        sheet_m2 = matching_tots[0]
                else:
                    sheet_m2 = max(explicit_total_m2)

            if explicit_total_slabs and sheet_slabs == 0:
                sheet_slabs = max(explicit_total_slabs)

            if sheet_m2 > best_m2:
                best_m2 = sheet_m2
                best_slabs = sheet_slabs
                if widths and heights:
                    avg_w = int(round(sum(widths) / len(widths)))
                    avg_h = int(round(sum(heights) / len(heights)))
                    best_avg_dim = f"{avg_w} x {avg_h} cm"

        if best_m2 > 0:
            return {
                "m2": round(best_m2, 2),
                "slabs": best_slabs,
                "avg_dim": best_avg_dim,
                "surface": f"Polished ({best_thickness})",
            }
        return None


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
            main_mod = sys.modules.get("__main__")
            if main_mod:
                for attr_name in dir(main_mod):
                    val = getattr(main_mod, attr_name, None)
                    if type(val).__name__ == "DriveExplorer":
                        return val
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

    def fetch_folder_children_direct(self, folder_id: str) -> List[Dict[str, Any]]:
        """
        Belirli bir Lot klasörünün içindeki tüm dosyaları doğrudan Drive API'den sorgular
        (Genel taramada gözden kaçmış hiçbir Excel dosyası kalmaması için!).
        """
        headers = self.get_headers()
        if not headers:
            return []
        url = "https://www.googleapis.com/drive/v3/files"
        params = {
            "q": f"'{folder_id}' in parents and trashed = false",
            "pageSize": 500,
            "fields": "files(id, name, mimeType, parents, webViewLink, modifiedTime, size, description)",
            "supportsAllDrives": "true",
            "includeItemsFromAllDrives": "true",
        }
        try:
            resp = self.session.get(url, headers=headers, params=params, timeout=20)
            if resp.status_code == 200:
                return resp.json().get("files", [])
        except Exception:
            pass
        return []

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

    def invoke_bot_py_native_parser(
        self, folder_id: str, folder_name: str, lot_code: str, excel_bytes: Optional[bytes] = None
    ) -> Optional[Dict[str, Any]]:
        """
        bot.py (__main__) içindeki tüm fonksiyonları ve DriveExplorer metodlarını
        inceleyerek K-3179 ve K-3968'i hesaplayan orijinal fonksiyonu çağırır.
        """
        def normalize_result(res: Any) -> Optional[Dict[str, Any]]:
            if isinstance(res, dict):
                m2_v = to_float_safe(
                    res.get("m2") or res.get("total_m2") or res.get("sqm") or res.get("area") or res.get("stok_m2")
                )
                sl_v = to_float_safe(
                    res.get("slabs") or res.get("plaka") or res.get("slab_count") or res.get("count") or res.get("pcs")
                )
                dm_v = (
                    res.get("avg_dim") or res.get("ebat") or res.get("ortalama_ebat") or res.get("dimensions") or "-"
                )
                sf_v = res.get("surface_thickness") or res.get("surface") or res.get("yuzey") or "Polished (2 cm)"
                if m2_v and m2_v > 0:
                    return {
                        "m2": round(m2_v, 2),
                        "slabs": int(sl_v) if sl_v and sl_v > 0 else 0,
                        "avg_dim": str(dm_v).strip() if dm_v else "-",
                        "surface": str(sf_v).strip() if sf_v else "Polished (2 cm)",
                    }
            elif isinstance(res, (tuple, list)) and len(res) >= 2:
                # Örn: (m2, slabs, avg_dim) veya (m2, slabs) döndüren fonksiyonlar
                v0 = to_float_safe(res[0])
                v1 = to_float_safe(res[1])
                if v0 and v0 > 0 and v1 is not None:
                    dm = str(res[2]).strip() if len(res) >= 3 and res[2] else "-"
                    return {
                        "m2": round(v0, 2),
                        "slabs": int(v1),
                        "avg_dim": dm,
                        "surface": "Polished (2 cm)",
                    }
            return None

        # 1. DriveExplorer nesnesinin metodlarını dene
        explorer = self.find_bot_drive_explorer()
        if explorer is not None:
            for attr_name in dir(explorer):
                if attr_name.startswith("__"):
                    continue
                fn = getattr(explorer, attr_name, None)
                if not callable(fn):
                    continue
                low = attr_name.lower()
                if any(k in low for k in ("excel", "meta", "metric", "lot", "folder", "parse", "calc", "detail", "info", "stock")):
                    # Eğer excel_bytes alan bir metodsa
                    if excel_bytes and ("excel" in low or "parse" in low or "sheet" in low):
                        for arg_ try in ((excel_bytes,), (io.BytesIO(excel_bytes),)):
                            try:
                                r = fn(*arg_try)
                                nr = normalize_result(r)
                                if nr:
                                    return nr
                            except Exception:
                                pass
                    # folder_id veya lot_code alan bir metodsa
                    for arg_try in ((folder_id,), (folder_id, folder_name), (lot_code,)):
                        try:
                            r = fn(*arg_try)
                            nr = normalize_result(r)
                            if nr:
                                return nr
                        except Exception:
                            pass

            # Çربیدن sonra DriveExplorer'ın kendi iç önbellek sözlüklerini tekrar kontrol et
            bot_cache_now = self.load_bot_existing_caches()
            for key in (folder_id, lot_code.upper(), folder_name.upper()):
                nr = normalize_result(bot_cache_now.get(key))
                if nr:
                    return nr

        # 2. bot.py (__main__) içindeki bağımsız Excel okuma fonksiyonlarını dene
        if excel_bytes:
            main_mod = sys.modules.get("__main__")
            if main_mod:
                for attr_name in dir(main_mod):
                    if attr_name.startswith("__"):
                        continue
                    fn = getattr(main_mod, attr_name, None)
                    if callable(fn) and any(k in attr_name.lower() for k in ("excel", "parse", "sheet", "metric", "lot")):
                        for arg_try in ((excel_bytes,), (io.BytesIO(excel_bytes),)):
                            try:
                                r = fn(*arg_try)
                                nr = normalize_result(r)
                                if nr:
                                    return nr
                            except Exception:
                                pass

        return None

    def extract_metrics_from_lot_files(
        self,
        folder_id: str,
        lot_code: str,
        folder_name: str,
        modified_time: str,
        all_lot_files: List[Dict[str, Any]],
        direct_images: List[Dict[str, Any]],
        bot_cache: Dict[str, Dict[str, Any]],
        deep_fetch_if_needed: bool = False,
    ) -> Dict[str, Any]:
        """
        Bir lotun metrajını (m²), plaka adedini (Slabs), ortalama ebadını (cm) ve kalınlığını
        KESİN OLARAK Excel tablosundan ve bot.py motorundan çıkarır.
        Asla alt klasörlerdeki kopya fotoğrafları toplayıp 276 Slabs gibi sayılar üretmez!
        """
        m2 = 0.0
        slabs = 0
        avg_dim = "-"
        surface = "Polished (2 cm)"

        # 1. ADIM: bot.py'nin kendi önbelleğinden (argon_v17_master_cache.json) kontrol et
        for key in (folder_id, lot_code.upper(), lot_code.replace("-", "").upper(), folder_name.upper()):
            c = bot_cache.get(key)
            if isinstance(c, dict):
                v_m2 = to_float_safe(
                    c.get("m2") or c.get("total_m2") or c.get("sqm") or c.get("area") or c.get("mevcut_stok")
                )
                v_sl = to_float_safe(
                    c.get("slabs") or c.get("plaka") or c.get("slab_count") or c.get("count") or c.get("pcs")
                )
                if v_m2 and v_m2 > 0:
                    m2 = v_m2
                    if v_sl and v_sl > 0:
                        slabs = int(v_sl)
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
                    break

        # 2. ADIM: V23 kalıcı önbelleğinden kontrol et (Sadece m2 > 0 ve avg_dim != '-' ise!)
        persisted = self.state.data.setdefault("metrics_cache", {}).get(folder_id)
        if m2 <= 0 and isinstance(persisted, dict) and persisted.get("m2", 0.0) > 0:
            m2 = float(persisted["m2"])
            slabs = int(persisted.get("slabs", 0))
            avg_dim = persisted.get("avg_dim", "-")
            surface = persisted.get("surface", surface)

        # 3. ADIM: Lotun içindeki Excel / Tablo dosyalarını tespit et ve oku!
        candidate_tables = [f for f in all_lot_files if is_spreadsheet_or_table_file(f)]

        # Eğer genel listede tablo görünmüyorsa ve deep_fetch aktifse doğrudan klasörün içini API ile sorgula
        if not candidate_tables and deep_fetch_if_needed and m2 <= 0:
            fresh_children = self.fetch_folder_children_direct(folder_id)
            for fc in fresh_children:
                if fc.get("mimeType") == "application/vnd.google-apps.folder":
                    sub_children = self.fetch_folder_children_direct(fc["id"])
                    all_lot_files.extend(sub_children)
                else:
                    all_lot_files.append(fc)
            candidate_tables = [f for f in all_lot_files if is_spreadsheet_or_table_file(f)]

        if m2 <= 0 and (candidate_tables or deep_fetch_if_needed):
            for tf in candidate_tables:
                raw_bytes = self.download_file_bytes(tf["id"], tf.get("mimeType", ""))
                if not raw_bytes:
                    continue

                # A) Evrensel Excel Okuyucu (openpyxl + xlrd + XML/ODS + CSV/HTML)
                sheets = UniversalExcelParser.extract_sheets_as_rows(raw_bytes, tf.get("name", ""))
                if sheets:
                    parsed = UniversalExcelParser.analyze_table_rows(sheets)
                    if parsed and parsed.get("m2", 0.0) > 0:
                        m2 = parsed["m2"]
                        if parsed.get("slabs", 0) > 0:
                            slabs = parsed["slabs"]
                        if parsed.get("avg_dim", "-") != "-":
                            avg_dim = parsed["avg_dim"]
                        if parsed.get("surface"):
                            surface = parsed["surface"]
                        break

                # B) bot.py'nin kendi yerel fonksiyonunu bu dosya byte'larıyla dene
                native_parsed = self.invoke_bot_py_native_parser(folder_id, folder_name, lot_code, raw_bytes)
                if native_parsed and native_parsed.get("m2", 0.0) > 0:
                    m2 = native_parsed["m2"]
                    if native_parsed.get("slabs", 0) > 0:
                        slabs = native_parsed["slabs"]
                    if native_parsed.get("avg_dim", "-") != "-":
                        avg_dim = native_parsed["avg_dim"]
                    if native_parsed.get("surface"):
                        surface = native_parsed["surface"]
                    break

        # 4. ADIM: Hala m2 <= 0 ise ve deep_fetch_if_needed aktifse bot.py'nin klasör analizörünü çağır
        if m2 <= 0 and deep_fetch_if_needed:
            native_folder_res = self.invoke_bot_py_native_parser(folder_id, folder_name, lot_code, None)
            if native_folder_res and native_folder_res.get("m2", 0.0) > 0:
                m2 = native_folder_res["m2"]
                if native_folder_res.get("slabs", 0) > 0:
                    slabs = native_folder_res["slabs"]
                if native_folder_res.get("avg_dim", "-") != "-":
                    avg_dim = native_folder_res["avg_dim"]
                if native_folder_res.get("surface"):
                    surface = native_folder_res["surface"]

        # 5. ADIM: Klasör adında plaka aralığı varsa (örn: "A-1037 - Silver River 1 38" -> 38 Plaka)
        if slabs <= 0:
            m_range = re.search(
                r"\b(?:between\s+)?(\d{1,3})\s*[-_\s]\s*(\d{1,3})\s*(?:stock|stok)?\s*$",
                folder_name.strip(),
                flags=re.I,
            )
            if m_range:
                b1, b2 = int(m_range.group(1)), int(m_range.group(2))
                if 1 <= b1 < b2 <= 150:
                    slabs = (b2 - b1) + 1

        # Eğer hala plaka sayısı 0 ise SADECE ana klasördeki (direct_images) tekil plaka fotoğraflarını say
        # ASLA alt klasörlerin (custom, email, photos) toplamını alma!
        if slabs <= 0 and direct_images:
            # Benzersiz plaka numaralarını çıkar (örn: A258-1, A258-2)
            slab_nums = set()
            for img_f in direct_images:
                nm = os.path.splitext(img_f.get("name", ""))[0]
                m_sn = re.search(r"[-_\s](\d{1,3})$", nm)
                if m_sn:
                    slab_nums.add(int(m_sn.group(1)))
            if slab_nums:
                slabs = len(slab_nums)
            else:
                slabs = len(direct_images)

        result = {
            "m2": round(m2, 2),
            "slabs": slabs,
            "avg_dim": avg_dim,
            "surface": surface,
            "modified_time": modified_time,
        }

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
            r"\b(full\s+lot|full|lot|stock|stok|between\s*\d+\s*[-_]?\s*\d*|between|\d+\s+\d+\s*(?:stock|stok)?|\d+([.,]\d+)?\s*(m2|m²|plaka|slabs?|cm|mm|pcs)|sold\s*out|sold)\b",
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

        def collect_lot_files_and_pure_images(fid: str) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
            """
            1) Lotun kendi içindeki ve alt klasörlerindeki TÜM dosyaları (Excel tablolarını bulmak için) toplar.
            2) Görselleri seçerken 'custom' klasörünü KESİNLİKLE YASAKLAR!
               Eğer ana klasörde fotoğraf varsa SADECE ana klasördeki fotoğrafları alır (alt klasörlerdeki
               kopyalarla asla karıştırmaz!). Ana klasörde fotoğraf yoksa sadece bir tane custom-olmayan
               alt klasörün fotoğraflarını alır.
            """
            all_f: List[Dict[str, Any]] = list(children_files.get(fid, []))
            direct_imgs: List[Dict[str, Any]] = []
            fallback_sub_imgs: List[Dict[str, Any]] = []

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

                if not is_custom_dir and not direct_imgs and not fallback_sub_imgs:
                    for x in s_files:
                        if is_image_file_obj(x):
                            fallback_sub_imgs.append(x)

                for deeper_sf in children_folders.get(sub_id, []):
                    deeper_norm = normalize_compare(deeper_sf.get("name", ""))
                    stack.append((deeper_sf, f"{sub_norm} {deeper_norm}"))

            direct_imgs.sort(key=lambda item: item.get("name", ""))
            fallback_sub_imgs.sort(key=lambda item: item.get("name", ""))
            chosen_images = direct_imgs if direct_imgs else fallback_sub_imgs
            return all_f, chosen_images

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
                all_lot_files, pure_imgs = collect_lot_files_and_pure_images(current_fid)
                if not has_lot_code and not all_lot_files and not pure_imgs:
                    return

                lot_code, stone_name = self.parse_lot_folder_title(cur_name, path_names[:-1])
                modified_time = cur_folder.get("modifiedTime", "")

                metrics = self.extract_metrics_from_lot_files(
                    current_fid,
                    lot_code,
                    cur_name,
                    modified_time,
                    all_lot_files,
                    pure_imgs,
                    bot_cache,
                    deep_fetch_if_needed=False,
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
                    "images": pure_imgs,
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

    def ensure_deep_excel_metrics_for_lots(self, lots: List[Dict[str, Any]]):
        """
        Sorgulanan lotların her birinin içindeki Excel dosyasını (eğer henüz okunmadıysa)
        doğrudan Drive klasörünün içine girip indirir ve gerçek m2, plaka ve ebadı çıkarır.
        """
        bot_cache = self.load_bot_existing_caches()
        updated_any = False

        for item in lots:
            if item.get("m2", 0.0) > 0 and item.get("avg_dim", "-") != "-":
                continue

            metrics = self.extract_metrics_from_lot_files(
                item["folder_id"],
                item["lot_code"],
                item["folder_name"],
                item.get("modified_time", ""),
                item.get("all_files", []),
                item.get("images", []),
                bot_cache,
                deep_fetch_if_needed=True,
            )
            if metrics.get("m2", 0.0) > 0:
                item["m2"] = metrics["m2"]
                item["slabs"] = metrics["slabs"]
                item["avg_dim"] = metrics["avg_dim"]
                item["surface"] = metrics["surface"]
                updated_any = True

        if updated_any:
            self.state.save(git_sync=True)

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
        # Eşleşen tüm lotların içindeki Excel dosyalarını derinlemesine oku!
        self.ensure_deep_excel_metrics_for_lots(matched)
        return stone_query.title(), matched

    def find_lot_by_code_or_name(self, query_str: str) -> Optional[Dict[str, Any]]:
        active_map, _ = self.scan_drive_collections(force=False)
        q_norm = normalize_compare(query_str)
        q_compact = re.sub(r"[^a-z0-9]", "", q_norm)

        for item in active_map.values():
            lc_compact = re.sub(r"[^a-z0-9]", "", item.get("lot_code", "").lower())
            if lc_compact and lc_compact == q_compact:
                self.ensure_deep_excel_metrics_for_lots([item])
                return item

        for item in active_map.values():
            blob = normalize_compare(f"{item.get('lot_code', '')} {item.get('stone_name', '')} {item.get('folder_name', '')}")
            if q_norm and q_norm in blob:
                self.ensure_deep_excel_metrics_for_lots([item])
                return item

        return None


def extract_pure_marble_texture(raw_img: Image.Image) -> Image.Image:
    """
    Fotoğrafın üstündeki vinç kıskacını (turuncu/sarı aparat), fabrika tavanını ve kenar boşluklarını
    otomatik keserek SADECE SAF MERMER YÜZEYİNİ çıkarır.
    """
    img = raw_img.convert("RGB")
    w, h = img.size
    arr = np.asarray(img, dtype=np.float32)

    # Plakanın göbek bölgesi (%32..%75 yükseklik, %22..%78 genişlik)
    cy1, cy2 = int(h * 0.32), int(h * 0.75)
    cx1, cx2 = int(w * 0.22), int(w * 0.78)
    center_patch = arr[cy1:cy2, cx1:cx2]
    mean_rgb = np.mean(center_patch, axis=(0, 1))
    std_rgb = np.std(center_patch, axis=(0, 1)) + 26.0

    diff = np.abs(arr - mean_rgb) / std_rgb
    is_slab_pixel = np.mean(diff, axis=2) < 2.10

    brightness = np.mean(arr, axis=2)
    saturation = np.max(arr, axis=2) - np.min(arr, axis=2)
    center_sat = np.mean(np.max(center_patch, axis=2) - np.min(center_patch, axis=2))

    valid_mask = (
        is_slab_pixel
        & (brightness > 30.0)
        & (saturation < (center_sat + 48.0))
    )

    row_ratio = np.mean(valid_mask[:, int(w * 0.20) : int(w * 0.80)], axis=1)
    col_ratio = np.mean(valid_mask[int(h * 0.25) : int(h * 0.80), :], axis=0)

    valid_rows = np.where(row_ratio > 0.65)[0]
    valid_cols = np.where(col_ratio > 0.65)[0]

    if len(valid_rows) > int(h * 0.35) and len(valid_cols) > int(w * 0.35):
        y1, y2 = int(valid_rows[0]), int(valid_rows[-1])
        x1, x2 = int(valid_cols[0]), int(valid_cols[-1])
        pad_x = int((x2 - x1) * 0.06)
        pad_y = int((y2 - y1) * 0.08)
        # Üstteki vinç kıskacını (clamp) kesinlikle dışarıda bırakmak için üstten en az %22 kırp
        x1 = max(int(w * 0.14), x1 + pad_x)
        x2 = min(int(w * 0.86), x2 - pad_x)
        y1 = max(int(h * 0.22), y1 + pad_y)
        y2 = min(int(h * 0.84), y2 - pad_y)
    else:
        x1, x2 = int(w * 0.16), int(w * 0.84)
        y1, y2 = int(h * 0.24), int(h * 0.82)

    if (x2 - x1) < 150 or (y2 - y1) < 120:
        x1, x2 = int(w * 0.15), int(w * 0.85)
        y1, y2 = int(h * 0.22), int(h * 0.82)

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
    try:
        with Image.open(io.BytesIO(img_bytes)) as raw_img:
            raw_img = raw_img.convert("RGB")
            raw_img.thumbnail((1400, 1000), Image.Resampling.LANCZOS)
            pure_slab = extract_pure_marble_texture(raw_img)

        pure_slab = ImageEnhance.Contrast(pure_slab).enhance(1.05)
        pure_slab = ImageEnhance.Sharpness(pure_slab).enhance(1.15)

        single_w, single_h = 520, 560
        left_panel = pure_slab.resize((single_w, single_h), Image.Resampling.LANCZOS)
        right_panel = ImageOps.mirror(left_panel)

        bm_w, bm_h = single_w * 2, single_h
        bookmatch = Image.new("RGB", (bm_w, bm_h))
        bookmatch.paste(left_panel, (0, 0))
        bookmatch.paste(right_panel, (single_w, 0))

        cw, ch = 1400, 980
        y_idx, x_idx = np.mgrid[0:ch, 0:cw].astype(np.float32)
        scene_arr = np.zeros((ch, cw, 3), dtype=np.float32)

        scene_arr[:, :, 0] = 20.0
        scene_arr[:, :, 1] = 20.0
        scene_arr[:, :, 2] = 22.0

        ceil_mask = y_idx < 95
        scene_arr[ceil_mask] *= (0.45 + 0.55 * (y_idx[ceil_mask] / 95.0))[:, None]

        wall_x1, wall_y1 = 180, 95
        wall_w, wall_h = bm_w, bm_h
        wall_arr = np.asarray(bookmatch, dtype=np.float32)

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

        gold = (197, 155, 95)
        dark_gold = (110, 86, 52)
        draw.rectangle([wall_x1 - 3, wall_y1 - 3, wall_x1 + wall_w + 2, wall_y1 + wall_h + 2], outline=gold, width=3)
        draw.line([(wall_x1 + single_w, wall_y1), (wall_x1 + single_w, wall_y1 + wall_h)], fill=(60, 58, 55), width=1)
        draw.line([(90, 95), (90, floor_y1)], fill=dark_gold, width=2)
        draw.line([(cw - 90, 95), (cw - 90, floor_y1)], fill=dark_gold, width=2)
        draw.line([(0, floor_y1), (cw, floor_y1)], fill=dark_gold, width=2)

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

    # 1. COVER PAGE
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

    # 2. LOT PAGES
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
    def sort_lots(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        def sk(x: Dict[str, Any]):
            lc = x.get("lot_code", "")
            m_n = re.search(r"(\d+)", lc)
            return (x.get("collection", ""), x.get("stone_name", ""), int(m_n.group(1)) if m_n else 999999, lc)
        return sorted(items, key=sk)

    active_list = sort_lots(list(active_map.values()))
    sold_list = sort_lots(list(sold_map.values()))

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

    combined_rows = [("IN STOCK", x) for x in active_list] + [("SOLD OUT", x) for x in sold_list]
    rows_per_page = 22
    total_pages = max(1, (len(combined_rows) + rows_per_page - 1) // rows_per_page)

    for p_idx in range(total_pages):
        c.setFillColor(bg_dark)
        c.rect(0, 0, page_w, page_h, fill=1, stroke=0)
        c.setStrokeColor(gold_primary)
        c.setLineWidth(1.4)
        c.rect(20, 20, page_w - 40, page_h - 40, fill=0, stroke=1)

        c.setFillColor(gold_primary)
        c.setFont("Helvetica-Bold", 14)
        c.drawString(35, page_h - 45, "ARSLANLAR MARBLE — DETAILED INVENTORY & SOLD OUT TABLE")
        c.setFillColor(text_muted)
        c.setFont("Helvetica", 9.5)
        c.drawRightString(page_w - 35, page_h - 45, f"Page {p_idx + 1} of {total_pages}")

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
        logger.info("🔍 Arslanlar V23 Drive Watcher started.")

        while True:
            try:
                active_map, sold_map = self.intel.scan_drive_collections(force=True)

                # Arka planda m2'si henüz okunmamış aktif lotların Excel'lerini partiler halinde sessizce oku
                unparsed = [x for x in active_map.values() if x.get("m2", 0.0) <= 0]
                if unparsed:
                    self.intel.ensure_deep_excel_metrics_for_lots(unparsed[:15])

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
                                self.intel.ensure_deep_excel_metrics_for_lots([a_item])
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

        # 1. DETAYLI MASTER STOK TABLOSU & ÖZET RAPOR İSTEĞİ
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
                "⏳ <b>Arslanlar Series, Collection ve Exclusive Collection</b> (In Stock + Sold Out) taranıyor, "
                "Excel tabloları okunuyor ve Master Stok Tablosu (PDF & Excel) hazırlanıyor...",
                parse_mode="HTML",
            )
            loop = asyncio.get_running_loop()

            def generate_all():
                act_m, sld_m = self.intel.scan_drive_collections(force=True)
                self.intel.ensure_deep_excel_metrics_for_lots(list(act_m.values()))
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

        # 3. TAŞ ÇEŞİDİNE ÖZEL İNGİLİZCE PREMIUM SUNUM
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
                "📑 Scanning active lots, reading Excel files inside each lot & generating English Luxury Presentation PDF...",
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

        # 4. TAŞ ÇEŞİDİNE GÖRE STOK SORGULAMA
        is_variety_query = bool(
            re.search(
                r"(^ne\s+kadar\s+|var\s+m[ıi]|kaç\s+m2|kac\s+m2|\bçeşitleri\b|\bcesitleri\b|\bneler\s+var\b|\bvar\s*\??$|^how\s+much\s+)",
                raw_text,
                flags=re.I,
            )
        )
        is_single_lot_code = bool(re.match(r"^[A-Za-z]{1,3}\s*[-_]?\s*\d{1,5}[A-Za-z]?$", raw_text))

        if is_variety_query and not is_single_lot_code:
            status_msg = await update.effective_message.reply_text("🔍 Scanning active collections & reading Excel files inside each lot...")
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
                name="ArslanlarDriveWatcherV23",
                daemon=True,
            )
            watcher_thread.start()
            logger.info("✅ Arslanlar V23 Universal Excel & Master Plugin successfully attached!")
        except Exception as e:
            logger.error(f"Plugin startup error: {e}")

    return ORIGINAL_RUN_POLLING(self, *args, **kwargs)


Application.run_polling = hooked_run_polling
