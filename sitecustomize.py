# -*- coding: utf-8 -*-
"""
ARSLANLAR MARBLE - V26 NATIVE BOT.PY ENGINE BRIDGE & RANGE-AWARE EXCEL PARSER (sitecustomize.py)
- Kendi basina uydurma plaka sayisi (75 Slabs, 34 Slabs) veya birlesik sayi (738371030) ASLA uretmez!
- 1. Oncelik: Dogrudan bot.py (__main__) icindeki orijinal DriveExplorer / Excel metraj fonksiyonunu
  (A-62, A-96, A-248, A-276'yi kusursuz hesaplayan orijinal kodu) yakalar ve calistirir.
- 2. Oncelik: Klasor adindaki plaka araliklarini ('25 37' -> 25..37 arasi 13 plaka, '2 33' -> 2..33 arasi,
  '1 + 20 44' -> 26 plaka, 'Without No:56') tipki bot.py gibi filtreleyerek Excel'den tam m2 ve ebat hesaplar.
- 'custom' klasorundeki fotograflari hicbir yerde kullanmaz.
- Render motorunda alt siyah kararti, ustteki vinc kiskaci ve 'DIKKAT' tabelasi tamamen kirpilmistir.
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
from typing import Dict, List, Any, Optional, Tuple, Set
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

RENDER_EXECUTOR = ThreadPoolExecutor(max_workers=2, thread_name_prefix="RAMSafeWorker")
TR_TZ = timezone(timedelta(hours=3))


def normalize_compare(s: str) -> str:
    if not s:
        return ""
    s = s.replace("İ", "i").replace("I", "i").replace("ı", "i").lower()
    tr_map = str.maketrans("çğöşü", "cgosu")
    s = s.translate(tr_map)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def is_sane_lot_metrics(m2: float, slabs: int, avg_dim: str = "-") -> bool:
    if m2 <= 1.2 or m2 > 2500.0:
        return False
    if slabs <= 0 or slabs > 350:
        return False
    avg_slab_area = m2 / float(slabs)
    if avg_slab_area < 0.80 or avg_slab_area > 12.5:
        return False
    if avg_dim and avg_dim != "-":
        m_d = re.search(r"(\d+)\s*[xX*]\s*(\d+)", str(avg_dim))
        if m_d:
            w, h = int(m_d.group(1)), int(m_d.group(2))
            if max(w, h) < 100 or min(w, h) < 70 or max(w, h) > 450:
                return False
    return True


def safe_float_cell(val: Any) -> Optional[float]:
    """
    Hücredeki sayıyı ASLA noktaları silip birleştirmeden (738371030 hatası yapmadan!)
    güvenli bir şekilde float'a çevirir.
    """
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        f = float(val)
        return f if not np.isnan(f) and not np.isinf(f) else None

    s = str(val).strip()
    if not s:
        return None

    # Tarih veya saat ifadelerini eler
    if re.search(r"\b\d{1,4}[/:-]\d{1,2}[/:-]\d{1,4}\b", s):
        return None
    # Lot kodunu (A-157 vb.) eler
    if re.match(r"^[A-Za-z]{1,3}\s*[-_]\s*\d+", s):
        return None

    # Birim ifadelerini temizle
    s_clean = re.sub(r"\b(m2|m²|sqm|cm|mm|mt|m|pcs|adet|plaka|slabs?)\b", "", s, flags=re.I).strip()

    # Sadece tek bir sayısal ifadeyi yakala
    m = re.search(r"[-+]?\d+(?:[.,]\d+)?", s_clean)
    if not m:
        return None

    num_str = m.group(0).replace(",", ".")
    try:
        return float(num_str)
    except Exception:
        return None


def parse_slab_filter_from_folder_name(folder_name: str, lot_code: str) -> Tuple[Optional[Set[int]], Set[int]]:
    """
    Klasör adındaki kısmi stok kurallarını tıpkı A-62 ('1 + 20 44' -> 26 plaka) ve
    A-96 ('Without No:56') ve A-157 ('25 37' -> 25..37 arası 13 plaka) ve
    A-221 ('2 33' -> 2..33 arası 32 plaka) gibi çözümler!
    Döndürdüğü değer: (allowed_slab_numbers_or_None, excluded_slab_numbers)
    """
    raw = folder_name.strip()
    # Lot kodunu (örn: A-157, A-62) metinden çıkar ki lot numarasını plaka numarası sanmasın!
    if lot_code:
        raw = re.sub(re.escape(lot_code), " ", raw, flags=re.I)
        raw = re.sub(re.escape(lot_code.replace("-", "")), " ", raw, flags=re.I)

    # Kalınlık veya m2 ifadelerini çıkar
    raw = re.sub(r"\b\d+\s*(?:cm|mm|m2|m²|sqm)\b", " ", raw, flags=re.I)

    excluded: Set[int] = set()
    allowed: Set[int] = set()

    # 1. 'Without No:56' veya 'Without 12, 15' kontrolü
    m_without = re.search(r"(?:without|haric|hariç|except)\s*(?:no\s*[:.]?\s*)?([\d\s,+&-]+)", raw, flags=re.I)
    if m_without:
        for n_str in re.findall(r"\d+", m_without.group(1)):
            v = int(n_str)
            if 1 <= v <= 200:
                excluded.add(v)
        raw = raw[:m_without.start()] + " " + raw[m_without.end():]

    # 2. Klasör adının sonundaki aralık veya '+' kombinasyonları:
    # Örn: "Silver Roots 1 + 20 44", "Silver Roots 25 37 Stock", "Silver Roots Between 2 33"
    # Önce 'a - b' veya 'a b' aralıklarını ve '+ c' tekil plakalarını bul
    tokens = re.findall(r"\+|\d{1,3}", raw)
    if tokens:
        nums_only = [int(x) for x in tokens if x.isdigit() and 1 <= int(x) <= 200]
        # Eğer en az 2 sayı varsa (örn: [25, 37] veya [1, 20, 44] veya [2, 33])
        if len(nums_only) >= 2:
            idx = 0
            while idx < len(tokens):
                tok = tokens[idx]
                if tok == "+":
                    idx += 1
                    continue
                if tok.isdigit():
                    n1 = int(tok)
                    # Yanında doğrudan bir sayı daha varsa (arada '+' yoksa) bu bir aralıktır (örn: 20 44 veya 25 37)
                    if idx + 1 < len(tokens) and tokens[idx + 1].isdigit():
                        n2 = int(tokens[idx + 1])
                        if 1 <= n1 < n2 <= 200:
                            for k in range(n1, n2 + 1):
                                allowed.add(k)
                            idx += 2
                            continue
                    # Tek başına duruyorsa (örn: '1 + 20 44' içindeki '1')
                    if 1 <= n1 <= 200:
                        allowed.add(n1)
                idx += 1

    final_allowed = (allowed - excluded) if allowed else None
    return final_allowed, excluded


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
    mt = str(file_obj.get("mimeType", "")).lower()
    nm = str(file_obj.get("name", "")).lower()
    if is_image_file_obj(file_obj):
        return False
    if nm.endswith((".pdf", ".mp4", ".mov", ".avi", ".rar", ".zip", ".7z", ".ds_store", ".ini", ".db")):
        return False
    if any(
        k in mt
        for k in ("spreadsheet", "excel", "csv", "tab-separated", "opendocument.spreadsheet", "ms-excel", "officedocument")
    ):
        return True
    if nm.endswith((".xlsx", ".xls", ".xlsm", ".xlsb", ".ods", ".csv", ".tsv")):
        return True
    if mt in ("application/octet-stream", "application/x-ole-storage", "application/vnd.ms-office"):
        return True
    return False


class PersistentWatcherState:
    def __init__(self, filepath: str = STATE_FILE):
        self.filepath = filepath
        self.lock = threading.Lock()
        self.data: Dict[str, Any] = {
            "version": 26,
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
                            if loaded.get("version") != 26:
                                loaded["metrics_cache"] = {}
                                loaded["version"] = 26
                            else:
                                mc = loaded.get("metrics_cache", {})
                                clean_mc = {}
                                for k, v in mc.items():
                                    if isinstance(v, dict) and is_sane_lot_metrics(
                                        float(v.get("m2", 0.0)),
                                        int(v.get("slabs", 0)),
                                        str(v.get("avg_dim", "-")),
                                    ):
                                        clean_mc[k] = v
                                loaded["metrics_cache"] = clean_mc
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
            self.data["version"] = 26
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
                ["git", "commit", "-m", "chore: auto-sync watcher state v26 [skip ci]"],
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


class RangeAwareMarbleExcelParser:
    """
    Hem tam lot Excel'lerini hem de 'A-157 Silver Roots 25 37', 'A-221 Silver Roots 2 33',
    'A-62 Silver Roots 1 + 20 44', 'A-96 Without No:56' gibi aralıklı/kısmi stok Excel'lerini
    milimetrik olarak okuyan motor.
    """

    @staticmethod
    def extract_sheets_as_rows(file_bytes: bytes, filename: str = "") -> List[List[List[Any]]]:
        sheets: List[List[List[Any]]] = []

        # 1. openpyxl (data_only=True)
        try:
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            for ws in wb.worksheets:
                rows = [list(r) for r in ws.iter_rows(values_only=True)]
                if rows:
                    sheets.append(rows)
        except Exception:
            pass

        # 2. openpyxl (data_only=False - Formül metinlerini görmek için)
        try:
            wb_f = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
            for ws in wb_f.worksheets:
                rows = [list(r) for r in ws.iter_rows(values_only=True)]
                if rows:
                    sheets.append(rows)
        except Exception:
            pass

        if sheets:
            return sheets

        # 3. xlrd (Eski .xls dosyaları)
        if xlrd is not None:
            try:
                book = xlrd.open_workbook(file_contents=file_bytes)
                for s_idx in range(book.nsheets):
                    sh = book.sheet_by_index(s_idx)
                    rows = []
                    for r_idx in range(sh.nrows):
                        row_vals = []
                        for c_idx in range(sh.ncols):
                            ctype = sh.cell_type(r_idx, c_idx)
                            cval = sh.cell_value(r_idx, c_idx)
                            if ctype == 3:
                                row_vals.append("")
                            else:
                                row_vals.append(cval)
                        rows.append(row_vals)
                    if rows:
                        sheets.append(rows)
                if sheets:
                    return sheets
            except Exception:
                pass

        # 4. OpenDocument (.ods)
        try:
            if zipfile.is_zipfile(io.BytesIO(file_bytes)):
                with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
                    if "content.xml" in zf.namelist():
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

        # 5. CSV / TSV / HTML Table
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

    @staticmethod
    def normalize_dim_to_cm(val: Optional[float]) -> Optional[float]:
        """Bir ebat değerini (metre: 2.95, cm: 295, mm: 2950) cm'ye standartlaştırır."""
        if val is None or val <= 0:
            return None
        if 0.60 <= val <= 4.50:
            return round(val * 100.0, 1)
        if 60.0 <= val <= 450.0:
            return round(val, 1)
        if 600.0 <= val <= 4500.0:
            return round(val / 10.0, 1)
        return None

    @classmethod
    def analyze_sheets_with_folder_filter(
        cls, sheets: List[List[List[Any]]], folder_name: str = "", lot_code: str = ""
    ) -> Optional[Dict[str, Any]]:
        allowed_slabs, excluded_slabs = parse_slab_filter_from_folder_name(folder_name, lot_code)
        best_result: Optional[Dict[str, Any]] = None

        for rows in sheets:
            if not rows:
                continue

            # Excel'in herhangi bir hücresinde '3cm' veya '2cm' kalınlık ibaresi var mı?
            thickness_str = "2 cm"
            for r in rows[:15]:
                row_txt = " ".join(str(c) for c in r if c is not None).lower()
                if re.search(r"\b(3\s*cm|30\s*mm)\b", row_txt):
                    thickness_str = "3 cm"
                    break
                if re.search(r"\b(2\s*cm|20\s*mm)\b", row_txt):
                    thickness_str = "2 cm"
                    break

            # Yan yana iki tablo (sol blok / sağ blok) olma ihtimaline karşı tüm sütun gruplarını tara
            col_no_idx = None
            col_w_idx = None
            col_h_idx = None
            col_m2_idx = None
            header_row_idx = None

            for r_i, r in enumerate(rows[:35]):
                norm_cells = [normalize_compare(str(c)) if c is not None else "" for c in r]
                w_c, h_c, m2_c, no_c = None, None, None, None

                for c_i, nc in enumerate(norm_cells):
                    if not nc:
                        continue
                    if nc in ("no", "sira", "sira no", "slab no", "plaka no", "plaka", "slab", "id", "#", "adet", "pcs"):
                        if no_c is None:
                            no_c = c_i
                    elif any(w_kw == nc or nc.startswith(w_kw + " ") for w_kw in ("boy", "length", "uzunluk", "len", "l", "en", "width")):
                        if w_c is None:
                            w_c = c_i
                        elif h_c is None and c_i != w_c:
                            h_c = c_i
                    elif any(h_kw == nc or nc.startswith(h_kw + " ") for h_kw in ("yukseklik", "height", "genislik", "h", "w", "boy", "en")):
                        if h_c is None and c_i != w_c:
                            h_c = c_i
                        elif w_c is None:
                            w_c = c_i
                    elif any(m_kw in nc for m_kw in ("m2", "sqm", "alan", "area", "metraj", "miktar", "quantity", "net", "toplam", "total")):
                        if m2_c is None:
                            m2_c = c_i

                if w_c is not None and h_c is not None:
                    col_no_idx = no_c
                    col_w_idx = w_c
                    col_h_idx = h_c
                    col_m2_idx = m2_c
                    header_row_idx = r_i
                    break

            slab_records: List[Tuple[int, float, float, float]] = []  # (slab_no, w_cm, h_cm, area_m2)
            summary_m2_list: List[float] = []
            summary_slabs_list: List[int] = []

            start_r = (header_row_idx + 1) if header_row_idx is not None else 0
            auto_seq = 0

            for r in rows[start_r:]:
                if not r or all(c is None or str(c).strip() == "" for c in r):
                    continue

                raw_strs = [str(c).strip() if c is not None else "" for c in r]
                joined_norm = normalize_compare(" ".join(raw_strs))
                if not joined_norm:
                    continue

                # Toplam satırı kontrolü
                if any(kw in joined_norm for kw in ("toplam", "total", "genel toplam", "grand total", "sum", "yekun")):
                    for c in r:
                        vn = safe_float_cell(c)
                        if vn is not None:
                            if 2.0 <= vn <= 2500.0 and (isinstance(c, float) or "," in str(c) or "." in str(c)):
                                summary_m2_list.append(vn)
                            elif 1.0 <= vn <= 250.0 and float(vn).is_integer():
                                summary_slabs_list.append(int(vn))
                    continue

                w_cm, h_cm, area_m2, slab_no = None, None, None, None

                # 1. Başlık sütunlarından oku
                if col_w_idx is not None and col_h_idx is not None and col_w_idx < len(r) and col_h_idx < len(r):
                    w_cm = cls.normalize_dim_to_cm(safe_float_cell(r[col_w_idx]))
                    h_cm = cls.normalize_dim_to_cm(safe_float_cell(r[col_h_idx]))
                    if col_no_idx is not None and col_no_idx < len(r):
                        sn_v = safe_float_cell(r[col_no_idx])
                        if sn_v and 1 <= int(sn_v) <= 250 and float(sn_v).is_integer():
                            slab_no = int(sn_v)
                    if col_m2_idx is not None and col_m2_idx < len(r):
                        am_v = safe_float_cell(r[col_m2_idx])
                        if am_v and 0.6 <= am_v <= 14.0:
                            area_m2 = am_v

                # 2. Başlık sütunları yoksa veya o hücreler boşsa satırı otomatik tara
                if w_cm is None or h_cm is None:
                    # Örn: "292 x 148" tek hücrede yazılmış mı?
                    for rs in raw_strs:
                        m_in = re.search(r"\b(\d{2,4}(?:[.,]\d{1,2})?)\s*[xX*]\s*(\d{2,4}(?:[.,]\d{1,2})?)\b", rs)
                        if m_in:
                            d1 = cls.normalize_dim_to_cm(safe_float_cell(m_in.group(1)))
                            d2 = cls.normalize_dim_to_cm(safe_float_cell(m_in.group(2)))
                            if d1 and d2 and max(d1, d2) >= 100.0 and min(d1, d2) >= 70.0:
                                w_cm, h_cm = max(d1, d2), min(d1, d2)
                                break

                if w_cm is None or h_cm is None:
                    # Satırdaki tüm hücrelerin sayısal değerlerini çıkar
                    cell_nums: List[Tuple[int, float]] = []
                    for c_i, c in enumerate(r):
                        vn = safe_float_cell(c)
                        if vn is not None and vn > 0:
                            cell_nums.append((c_i, vn))

                    # Yan yana veya yakın iki hücrede geçerli plaka boyutu (w_cm, h_cm) var mı?
                    for idx_a in range(len(cell_nums)):
                        for idx_b in range(idx_a + 1, len(cell_nums)):
                            c_i1, v1 = cell_nums[idx_a]
                            c_i2, v2 = cell_nums[idx_b]
                            d1 = cls.normalize_dim_to_cm(v1)
                            d2 = cls.normalize_dim_to_cm(v2)
                            if d1 and d2 and max(d1, d2) >= 100.0 and min(d1, d2) >= 70.0:
                                calc_a = (d1 * d2) / 10000.0
                                # Diğer hücrelerde bu alana yakın bir m2 değeri VEYA başta sıra no var mı?
                                other_vals = [v for k_i, (_, v) in enumerate(cell_nums) if k_i not in (idx_a, idx_b)]
                                matching_a = [ov for ov in other_vals if abs(ov - calc_a) <= 0.45]
                                seq_cands = [
                                    int(v) for k_i, (_, v) in enumerate(cell_nums[:idx_a])
                                    if 1 <= v <= 220 and float(v).is_integer()
                                ]
                                if matching_a or seq_cands or len(cell_nums) >= 2:
                                    w_cm, h_cm = max(d1, d2), min(d1, d2)
                                    if matching_a:
                                        area_m2 = matching_a[0]
                                    if seq_cands:
                                        slab_no = seq_cands[0]
                                    break
                        if w_cm and h_cm:
                            break

                if w_cm and h_cm and max(w_cm, h_cm) >= 100.0 and min(w_cm, h_cm) >= 70.0:
                    auto_seq += 1
                    final_no = slab_no if slab_no is not None else auto_seq
                    calc_area = (w_cm * h_cm) / 10000.0
                    final_area = area_m2 if (area_m2 and abs(area_m2 - calc_area) <= 1.2) else calc_area
                    slab_records.append((final_no, max(w_cm, h_cm), min(w_cm, h_cm), final_area))

            if slab_records:
                # Klasör adında plaka filtresi varsa (örn: '25 37' veya '2 33' veya 'Without No:56') uygula!
                filtered_records = []
                for s_no, w_v, h_v, a_v in slab_records:
                    if s_no in excluded_slabs:
                        continue
                    if allowed_slabs is not None and s_no not in allowed_slabs:
                        continue
                    filtered_records.append((s_no, w_v, h_v, a_v))

                # Eğer plaka numaraları 1'den başlamadığı için filtre boş kaldıysa, sıra indeksine göre filtrele
                if not filtered_records and allowed_slabs is not None:
                    for idx_rec, (s_no, w_v, h_v, a_v) in enumerate(slab_records, start=1):
                        if idx_rec in allowed_slabs and idx_rec not in excluded_slabs:
                            filtered_records.append((s_no, w_v, h_v, a_v))

                target_recs = filtered_records if filtered_records else slab_records
                tot_m2 = round(sum(x[3] for x in target_recs), 2)
                tot_slabs = len(target_recs)
                avg_w = int(round(sum(x[1] for x in target_recs) / tot_slabs))
                avg_h = int(round(sum(x[2] for x in target_recs) / tot_slabs))
                avg_dim_str = f"{avg_w} x {avg_h} cm"

                # Eğer klasör filtresi yoksa ve toplam satırı varsa
                if allowed_slabs is None and not excluded_slabs and summary_m2_list:
                    for sm in summary_m2_list:
                        if abs(sm - tot_m2) <= (tot_m2 * 0.15 + 1.5):
                            tot_m2 = round(sm, 2)
                            break

                if is_sane_lot_metrics(tot_m2, tot_slabs, avg_dim_str):
                    cand = {
                        "m2": tot_m2,
                        "slabs": tot_slabs,
                        "avg_dim": avg_dim_str,
                        "surface": f"Polished ({thickness_str})",
                    }
                    if best_result is None or tot_slabs > best_result["slabs"]:
                        best_result = cand

            elif summary_m2_list and summary_slabs_list:
                sm2 = max(summary_m2_list)
                ssl = max(summary_slabs_list)
                if is_sane_lot_metrics(sm2, ssl, "-"):
                    cand = {
                        "m2": round(sm2, 2),
                        "slabs": ssl,
                        "avg_dim": "-",
                        "surface": f"Polished ({thickness_str})",
                    }
                    if best_result is None:
                        best_result = cand

        return best_result


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

    def call_bot_py_native_lot_engine(
        self, folder_id: str, folder_name: str, lot_code: str, excel_bytes: Optional[bytes] = None
    ) -> Optional[Dict[str, Any]]:
        """
        bot.py'nin A-62, A-96, A-248, A-276'yı hesaplayan kendi orijinal metodunu
        imzasına (signature) bakarak doğru argümanlarla çağırır.
        """
        def parse_any_bot_return(res: Any) -> Optional[Dict[str, Any]]:
            if isinstance(res, dict):
                m2_v = safe_float_cell(
                    res.get("m2") or res.get("total_m2") or res.get("sqm") or res.get("area") or res.get("stok_m2")
                )
                sl_v = safe_float_cell(
                    res.get("slabs") or res.get("plaka") or res.get("slab_count") or res.get("count") or res.get("pcs")
                )
                dm_v = (
                    res.get("avg_dim") or res.get("ebat") or res.get("ortalama_ebat") or res.get("dimensions") or "-"
                )
                sf_v = res.get("surface_thickness") or res.get("surface") or res.get("yuzey") or "Polished (2 cm)"
                if m2_v and sl_v and is_sane_lot_metrics(float(m2_v), int(sl_v), str(dm_v)):
                    return {
                        "m2": round(float(m2_v), 2),
                        "slabs": int(sl_v),
                        "avg_dim": str(dm_v).strip() if dm_v else "-",
                        "surface": str(sf_v).strip() if sf_v else "Polished (2 cm)",
                    }
            return None

        explorer = self.find_bot_drive_explorer()
        if explorer is not None:
            for attr_name in dir(explorer):
                if attr_name.startswith("__"):
                    continue
                fn = getattr(explorer, attr_name, None)
                if not callable(fn) or inspect.iscoroutinefunction(fn):
                    continue
                try:
                    sig = inspect.signature(fn)
                    param_names = [p.lower() for p in sig.parameters.keys()]
                except Exception:
                    param_names = []

                low_fn = attr_name.lower()
                if any(k in low_fn for k in ("meta", "metric", "excel", "lot", "folder", "analyze", "parse", "calc", "stock")):
                    # Parametre isimlerine göre nokta atışı çağır
                    call_args_list = [
                        (folder_id, folder_name),
                        (folder_id,),
                        ({"id": folder_id, "name": folder_name},),
                    ]
                    if excel_bytes:
                        call_args_list.insert(0, (excel_bytes, folder_name))
                        call_args_list.insert(1, (excel_bytes,))

                    for c_args in call_args_list:
                        try:
                            res = fn(*c_args)
                            parsed = parse_any_bot_return(res)
                            if parsed:
                                return parsed
                        except Exception:
                            pass

            # Metod çağrısı DriveExplorer'ın iç önbelleğini (_folder_meta_cache vb.) doldurduysa oradan al
            fresh_cache = self.load_bot_existing_caches()
            for key in (folder_id, lot_code.upper(), folder_name.upper()):
                parsed = parse_any_bot_return(fresh_cache.get(key))
                if parsed:
                    return parsed

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
        m2 = 0.0
        slabs = 0
        avg_dim = "-"
        surface = "Polished (2 cm)"

        # 1. ADIM: bot.py'nin hazır önbelleğinden (argon_v17_master_cache.json) oku
        for key in (folder_id, lot_code.upper(), lot_code.replace("-", "").upper(), folder_name.upper()):
            c = bot_cache.get(key)
            if isinstance(c, dict):
                v_m2 = safe_float_cell(
                    c.get("m2") or c.get("total_m2") or c.get("sqm") or c.get("area") or c.get("mevcut_stok")
                )
                v_sl = safe_float_cell(
                    c.get("slabs") or c.get("plaka") or c.get("slab_count") or c.get("count") or c.get("pcs")
                )
                v_dm = "-"
                for k_d in ("avg_dim", "ebat", "ortalama_ebat", "dimensions", "dim"):
                    val = c.get(k_d)
                    if val and str(val).strip() not in ("", "-", "None", "N/A"):
                        v_dm = str(val).strip()
                        break
                if v_m2 and v_sl and is_sane_lot_metrics(float(v_m2), int(v_sl), v_dm):
                    m2 = float(v_m2)
                    slabs = int(v_sl)
                    avg_dim = v_dm
                    for k_s in ("surface_thickness", "surface", "finish", "yuzey"):
                        val = c.get(k_s)
                        if val and str(val).strip() not in ("", "-", "None"):
                            surface = str(val).strip()
                            break
                    break

        # 2. ADIM: V26 kalıcı önbelleğinden oku
        persisted = self.state.data.setdefault("metrics_cache", {}).get(folder_id)
        if m2 <= 0 and isinstance(persisted, dict):
            pm2 = float(persisted.get("m2", 0.0))
            psl = int(persisted.get("slabs", 0))
            pdm = str(persisted.get("avg_dim", "-"))
            if is_sane_lot_metrics(pm2, psl, pdm):
                m2 = pm2
                slabs = psl
                avg_dim = pdm
                surface = persisted.get("surface", surface)

        # 3. ADIM: Derin okuma aktifse önce bot.py'nin kendi orijinal motorunu, sonra RangeAwareExcelParser'ı çalıştır!
        if m2 <= 0 and deep_fetch_if_needed:
            # A) Doğrudan bot.py'nin kendi klasör analizörünü çağır
            native_res = self.call_bot_py_native_lot_engine(folder_id, folder_name, lot_code, None)
            if native_res:
                m2 = native_res["m2"]
                slabs = native_res["slabs"]
                avg_dim = native_res["avg_dim"]
                surface = native_res["surface"]

            # B) Hala bulunmadıysa klasördeki tüm dosyaları (ve alt klasörleri) doğrudan Drive API'den çekip Excel'i oku
            if m2 <= 0:
                fresh_children = self.fetch_folder_children_direct(folder_id)
                combined_files = list(all_lot_files)
                seen_ids = {x.get("id") for x in combined_files if x.get("id")}
                for fc in fresh_children:
                    if fc.get("id") not in seen_ids:
                        combined_files.append(fc)
                        seen_ids.add(fc.get("id"))
                    if fc.get("mimeType") == "application/vnd.google-apps.folder":
                        for sc in self.fetch_folder_children_direct(fc["id"]):
                            if sc.get("id") not in seen_ids:
                                combined_files.append(sc)
                                seen_ids.add(sc.get("id"))

                candidate_tables = [f for f in combined_files if is_spreadsheet_or_table_file(f)]
                for tf in candidate_tables:
                    raw_bytes = self.download_file_bytes(tf["id"], tf.get("mimeType", ""))
                    if not raw_bytes:
                        continue

                    # Önce bot.py'nin kendi Excel fonksiyonuna byte'ları ver
                    native_excel_res = self.call_bot_py_native_lot_engine(folder_id, folder_name, lot_code, raw_bytes)
                    if native_excel_res:
                        m2 = native_excel_res["m2"]
                        slabs = native_excel_res["slabs"]
                        avg_dim = native_excel_res["avg_dim"]
                        surface = native_excel_res["surface"]
                        break

                    # RangeAwareMarbleExcelParser ('25 37', '2 33', '1 + 20 44', 'Without No:56' duyarlı!)
                    sheets = RangeAwareMarbleExcelParser.extract_sheets_as_rows(raw_bytes, tf.get("name", ""))
                    if sheets:
                        parsed = RangeAwareMarbleExcelParser.analyze_sheets_with_folder_filter(
                            sheets, folder_name=folder_name, lot_code=lot_code
                        )
                        if parsed and is_sane_lot_metrics(parsed["m2"], parsed["slabs"], parsed["avg_dim"]):
                            m2 = parsed["m2"]
                            slabs = parsed["slabs"]
                            avg_dim = parsed["avg_dim"]
                            surface = parsed["surface"]
                            break

        # 4. ADIM: Klasör adında plaka aralığı varsa (örn: "A-157 - Silver Roots 25 37" -> 13 Plaka, "2 33" -> 32 Plaka)
        if slabs <= 0:
            allowed_set, _ = parse_slab_filter_from_folder_name(folder_name, lot_code)
            if allowed_set:
                slabs = len(allowed_set)

        if slabs <= 0 and direct_images:
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

        if is_sane_lot_metrics(result["m2"], result["slabs"], result["avg_dim"]):
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
            r"\b(full\s+lot|full|lot|stock|stok|between\s*\d+\s*[-_]?\s*\d*|between|\d+([.,]\d+)?\s*(m2|m²|plaka|slabs?|cm|mm|pcs)|sold\s*out|sold)\b",
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

    def ensure_deep_excel_metrics_for_lots(self, lots: List[Dict[str, Any]], max_fetch: int = 35):
        bot_cache = self.load_bot_existing_caches()
        updated_any = False
        fetched_count = 0

        for item in lots:
            if is_sane_lot_metrics(item.get("m2", 0.0), item.get("slabs", 0), item.get("avg_dim", "-")):
                continue
            if fetched_count >= max_fetch:
                break

            fetched_count += 1
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
        self.ensure_deep_excel_metrics_for_lots(matched, max_fetch=35)
        return stone_query.title(), matched

    def find_lot_by_code_or_name(self, query_str: str) -> Optional[Dict[str, Any]]:
        active_map, _ = self.scan_drive_collections(force=False)
        q_norm = normalize_compare(query_str)
        q_compact = re.sub(r"[^a-z0-9]", "", q_norm)

        for item in active_map.values():
            lc_compact = re.sub(r"[^a-z0-9]", "", item.get("lot_code", "").lower())
            if lc_compact and lc_compact == q_compact:
                self.ensure_deep_excel_metrics_for_lots([item], max_fetch=2)
                return item

        for item in active_map.values():
            blob = normalize_compare(f"{item.get('lot_code', '')} {item.get('stone_name', '')} {item.get('folder_name', '')}")
            if q_norm and q_norm in blob:
                self.ensure_deep_excel_metrics_for_lots([item], max_fetch=2)
                return item

        return None


def extract_pure_marble_texture(raw_img: Image.Image) -> Image.Image:
    img = raw_img.convert("RGB")
    w, h = img.size
    x1 = int(w * 0.20)
    x2 = int(w * 0.80)
    y1 = int(h * 0.34)
    y2 = int(h * 0.84)
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

        cw, ch = 1400, 880
        canvas = Image.new("RGB", (cw, ch), (242, 239, 234))
        draw = ImageDraw.Draw(canvas)

        single_w, single_h = 640, 720
        left_panel = pure_slab.resize((single_w, single_h), Image.Resampling.LANCZOS)
        right_panel = ImageOps.mirror(left_panel)

        bm_w, bm_h = single_w * 2, single_h
        bookmatch = Image.new("RGB", (bm_w, bm_h))
        bookmatch.paste(left_panel, (0, 0))
        bookmatch.paste(right_panel, (single_w, 0))

        wall_x1, wall_y1 = 60, 40
        draw.rectangle([wall_x1 - 6, wall_y1 - 6, wall_x1 + bm_w + 5, wall_y1 + bm_h + 5], fill=(215, 210, 202))
        canvas.paste(bookmatch, (wall_x1, wall_y1))

        gold = (188, 145, 85)
        draw.rectangle([wall_x1 - 3, wall_y1 - 3, wall_x1 + bm_w + 2, wall_y1 + bm_h + 2], outline=gold, width=3)
        draw.line([(wall_x1 + single_w, wall_y1), (wall_x1 + single_w, wall_y1 + bm_h)], fill=(130, 125, 120), width=1)

        draw.rectangle([0, 795, cw, ch], fill=(22, 22, 25))
        draw.line([(0, 795), (cw, 795)], fill=gold, width=3)

        stock_txt = f"{m2:,.2f} m2 ({slabs} Slabs)" if m2 > 0 else f"{slabs} Slabs"
        dim_txt = f"   |   Avg. Size: {avg_dim}" if avg_dim and avg_dim != "-" else ""

        header_line = f"ARSLANLAR MARBLE   |   LOT: {lot_code} — {stone_name.upper()}   ({collection_name.upper()})"
        spec_line = f"Seamless 2x Bookmatch (Butterfly) Architectural Wall   |   Available Stock: {stock_txt}{dim_txt}"

        draw.text((60, 812), header_line, fill=(235, 205, 155))
        draw.text((60, 842), spec_line, fill=(205, 205, 210))

        out_buf = io.BytesIO()
        canvas.save(out_buf, format="JPEG", quality=93, optimize=True)
        result_bytes = out_buf.getvalue()

        del raw_img, pure_slab, left_panel, right_panel, bookmatch, canvas
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
    c.drawCentredString(page_w / 2, page_h - 175, "ARSLANLAR MARBLE — MASTER INVENTORY PRESENTATION")
    c.setFillColor(gold_light)
    c.setFont("Helvetica", 11)
    c.drawCentredString(
        page_w / 2,
        page_h - 195,
        f"COMPLETE IN STOCK & SOLD OUT LEDGER  |  GENERATED: {datetime.now(TR_TZ).strftime('%d.%m.%Y %H:%M')}",
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


def format_welcome_help_guide() -> str:
    return (
        "🏢 <b>Arslanlar Marble | Argon Kurumsal Dijital Asistan (V26)</b>\n\n"
        "Merhaba! Sistem üzerinden aşağıdaki tüm işlemleri anında gerçekleştirebilirsiniz. "
        "Mavi renkli örnek komutlara dokunarak kopyalayabilir veya benzer şekilde yazabilirsiniz:\n\n"
        "📊 <b>1. Detaylı Master Stok Tablosu (Aktif + Sold Out PDF & Excel):</b>\n"
        "<i>Arslanlar Series, Collection ve Exclusive Collection içindeki tüm stokta olan ve satılan taşların tam listesini verir:</i>\n"
        "• <code>Stok tablosu</code>\n"
        "• <code>Detaylı stok</code> veya <code>Toplam stoğum ne kadar</code>\n\n"
        "💎 <b>2. Taş Çeşidine Göre Aktif Stok Sorgulama:</b>\n"
        "<i>Sadece aktif klasörlerdeki taşları tarar (Sold Out hariç), metraj, plaka, ebat ve fotoğrafları getirir:</i>\n"
        "• <code>Ne kadar Silver Roots var?</code>\n"
        "• <code>Ne kadar Silver River var?</code>\n"
        "• <code>Ne kadar Rosso Levanto var?</code> veya <code>Tundra çeşitleri</code>\n\n"
        "📑 <b>3. Taş Çeşidine Özel İngilizce Lüks PDF Sunum:</b>\n"
        "<i>İstediğiniz taşın aktif lotlarını tarayıp Arslanlar Marble logolu, %100 İngilizce PDF katalog hazırlar:</i>\n"
        "• <code>Silver Rootslar sunum lütfen</code>\n"
        "• <code>Silver Riverlar sunum lütfen</code>\n"
        "• <code>Rosso Levantolar sunum lütfen</code>\n\n"
        "🏛 <b>4. Koleksiyon Bazlı Sunumlar (Aktif & Sold Out):</b>\n"
        "• <code>Arslanlar Series sunum lütfen</code>\n"
        "• <code>Arslanlar Collection sunum lütfen</code>\n"
        "• <code>Arslanlar Exclusive Collection sunum lütfen</code>\n"
        "• <code>Arslanlar Series satılan taşların sunumu</code>\n\n"
        "🔍 <b>5. Tekil Lot Sorgulama & Resmi Çeki Listesi (Packing List):</b>\n"
        "• <code>A-371</code> veya <code>A-258 fotoğraf</code>\n"
        "• <code>A-371 çeki listesi</code>\n\n"
        "🎨 <b>6. Mimari Bookmatch (Kelebek) Duvar Render Stüdyosu:</b>\n"
        "• <code>render A-371</code> veya <code>render A-258</code>\n\n"
        "🔔 <b>7. Otomatik Drive Gözcüsü & Haftalık Rapor (7/24 Aktif):</b>\n"
        "• Drive'da bir lot <b>Sold Out</b> klasörüne taşındığında veya <b>Yeni Lot</b> eklendiğinde otomatik bildirim düşer.\n"
        "• Her <b>Pazartesi ve Cuma 09:30'da</b> otomatik özet stok raporu gönderilir."
    )


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

    def format_executive_summary_report(self, active_map: Dict[str, Dict[str, Any]], sold_map: Dict[str, Dict[str, Any]], title_prefix: str = "WEEKLY") -> str:
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
        logger.info("🔍 Arslanlar V26 Drive Watcher started.")

        while True:
            try:
                active_map, sold_map = self.intel.scan_drive_collections(force=True)

                unparsed = [
                    x for x in active_map.values()
                    if not is_sane_lot_metrics(x.get("m2", 0.0), x.get("slabs", 0), x.get("avg_dim", "-"))
                ]
                if unparsed:
                    self.intel.ensure_deep_excel_metrics_for_lots(unparsed, max_fetch=15)

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
                                self.intel.ensure_deep_excel_metrics_for_lots([a_item], max_fetch=2)
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
                        report_html = self.format_executive_summary_report(active_map, sold_map, title_prefix=day_title)
                        self.send_telegram_sync(report_html)
                        try:
                            pdf_b, xlsx_b = build_master_stock_table_pdf_and_excel(active_map, sold_map, self.intel)
                            self.send_telegram_document_sync(
                                pdf_b,
                                "ARSLANLAR_MARBLE_MASTER_INVENTORY_TABLE.pdf",
                                f"📊 <b>{day_title} MASTER INVENTORY PRESENTATION (PDF)</b> — In Stock & Sold Out",
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

            time.sleep(90)

    async def handle_priority_message(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        if not update.effective_message or not update.effective_message.text:
            return

        if update.effective_chat:
            self.state.register_chat(update.effective_chat.id)

        raw_text = update.effective_message.text.strip()
        norm_text = normalize_compare(raw_text)

        # 0. MERHABA / KARŞILAMA VE YETENEK REHBERİ
        if norm_text in (
            "merhaba", "selam", "slm", "merhabalar", "iyi gunler", "gunaydin",
            "iyi aksamlar", "yardim", "help", "komutlar", "neler yapabilirsin",
            "ne yapabilirsin", "menu", "nasilsin"
        ) or raw_text.lower().startswith(("/start", "/help", "/menu")):
            await update.effective_message.reply_text(
                format_welcome_help_guide(),
                parse_mode="HTML",
                disable_web_page_preview=True,
            )
            raise ApplicationHandlerStop

        # 1. DETAYLI MASTER STOK TABLOSU & SUNUMU (EXCEL + PDF)
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
                "Premium PDF Sunum & Excel Stok Tablosu hazırlanıyor...",
                parse_mode="HTML",
            )
            loop = asyncio.get_running_loop()

            def generate_master_fast():
                act_m, sld_m = self.intel.scan_drive_collections(force=False)
                if not act_m and not sld_m:
                    act_m, sld_m = self.intel.scan_drive_collections(force=True)
                unparsed_lots = [
                    x for x in act_m.values()
                    if not is_sane_lot_metrics(x.get("m2", 0.0), x.get("slabs", 0), x.get("avg_dim", "-"))
                ]
                if unparsed_lots:
                    self.intel.ensure_deep_excel_metrics_for_lots(unparsed_lots, max_fetch=12)
                rep_html = self.format_executive_summary_report(act_m, sld_m, title_prefix="MASTER")
                pdf_b, xlsx_b = build_master_stock_table_pdf_and_excel(act_m, sld_m, self.intel)
                return rep_html, pdf_b, xlsx_b, len(act_m), len(sld_m)

            try:
                report_html, pdf_bytes, excel_bytes, act_cnt, sld_cnt = await loop.run_in_executor(
                    RENDER_EXECUTOR, generate_master_fast
                )
                await status_msg.edit_text(report_html, parse_mode="HTML", disable_web_page_preview=True)

                await update.effective_message.reply_document(
                    document=io.BytesIO(pdf_bytes),
                    filename="ARSLANLAR_MARBLE_MASTER_INVENTORY_PRESENTATION.pdf",
                    caption=(
                        f"📕 <b>ARSLANLAR MARBLE — MASTER INVENTORY PRESENTATION (PDF)</b>\n"
                        f"• 🟢 <b>In Stock (Active):</b> <code>{act_cnt} Lots</code>\n"
                        f"• 🔴 <b>Sold Out (Archived):</b> <code>{sld_cnt} Lots</code>\n"
                        f"• <b>Collections:</b> Series, Collection, Exclusive Collection"
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
            except Exception as e:
                await status_msg.edit_text(f"⚠️ Master tablo oluşturulurken hata: {e}")

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
                f"🎨 Preparing Bright Full-Screen Bookmatch Architectural Wall Render for <b>{target_q.upper()}</b>...",
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
                    f"• <b>Studio:</b> Bright Full-Screen 2x Bookmatch Architectural Wall"
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
                "📑 Scanning active lots, reading Excel files & generating English Luxury Presentation PDF...",
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
        is_general_stock_question = any(
            gk in norm_text for gk in ("stogumda", "toplam stogum", "genel stogum", "tum stogum")
        )
        is_variety_query = bool(
            re.search(
                r"(^ne\s+kadar\s+|var\s+m[ıi]|kaç\s+m2|kac\s+m2|\bçeşitleri\b|\bcesitleri\b|\bneler\s+var\b|\bvar\s*\??$|^how\s+much\s+)",
                raw_text,
                flags=re.I,
            )
        )
        is_single_lot_code = bool(re.match(r"^[A-Za-z]{1,3}\s*[-_]?\s*\d{1,5}[A-Za-z]?$", raw_text))

        if is_variety_query and not is_single_lot_code and not is_general_stock_question:
            status_msg = await update.effective_message.reply_text("🔍 Scanning active collections & reading Excel files...")
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
                name="ArslanlarDriveWatcherV26",
                daemon=True,
            )
            watcher_thread.start()
            logger.info("✅ Arslanlar V26 Native Bridge & Range-Aware Excel Plugin successfully attached!")
        except Exception as e:
            logger.error(f"Plugin startup error: {e}")

    return ORIGINAL_RUN_POLLING(self, *args, **kwargs)


Application.run_polling = hooked_run_polling
