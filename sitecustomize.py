# -*- coding: utf-8 -*-
"""
ARSLANLAR MARBLE - V18 ZERO-TOUCH EKLENTI MODULU (sitecustomize.py)
Bu dosya bot.py dosyasina HIC DOKUNMADAN yeni ozellikleri sisteme entegre eder.
Mevcut 4640 satirlik bot.py kodunun tek bir satirini bile degistirmez veya bozmaz.
"""

import os
import re
import gc
import io
import json
import time
import base64
import logging
import threading
import subprocess
from datetime import datetime, timezone, timedelta
from typing import Dict, List, Any, Optional, Tuple
from concurrent.futures import ThreadPoolExecutor

import requests
import numpy as np
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
    CommandHandler,
    ContextTypes,
    ApplicationHandlerStop,
    filters,
)

logger = logging.getLogger("ArgonMaster_V17.Plugin")

# Sabit Dosya ve Koleksiyon İsimleri
STATE_FILE = "arslanlar_watcher_state.json"
BOT_CACHE_FILE = "argon_v17_master_cache.json"

TARGET_COLLECTIONS = [
    "Arslanlar Series",
    "Arslanlar Collection",
    "Arslanlar Exclusive Collection",
]

# Render ve PDF işlemleri için RAM kilitlemeyen tek kanallı havuz (Max 1 işlem aynı anda)
RENDER_EXECUTOR = ThreadPoolExecutor(max_workers=1, thread_name_prefix="RAMSafeRender")

# Türkiye Saati (UTC+3)
TR_TZ = timezone(timedelta(hours=3))


def _strip_turkish_plural_and_clean(text: str) -> str:
    """
    'Ne kadar silver Rootslar sunum lütfen' -> 'silver roots'
    'Rosso levantolar sunum lütfen' -> 'rosso levanto'
    Türkçe çoğul ve çekim eklerini taş isimlerinden güvenli şekilde temizler.
    """
    t = text.strip()
    # Küçük harfe çevirirken Türkçe karakterleri koru / normalize et
    t_low = t.replace("İ", "i").replace("I", "ı").lower()

    # Sorgu kalıplarını çıkar
    remove_phrases = [
        r"^ne\s+kadar\s+",
        r"^kaç\s+m2\s+",
        r"^kac\s+m2\s+",
        r"^kaç\s+metrekare\s+",
        r"^elimde\s+ne\s+kadar\s+",
        r"^elimizde\s+ne\s+kadar\s+",
        r"^stokta\s+ne\s+kadar\s+",
        r"^bana\s+",
        r"^lütfen\s+",
        r"^lutfen\s+",
        r"\s+sunumu?\s+lütfen.*$",
        r"\s+sunumu?\s+lutfen.*$",
        r"\s+sunumu?\s+hazırla.*$",
        r"\s+sunumu?\s+yap.*$",
        r"\s+sunumu?$",
        r"\s+kataloğu?\s*.*$",
        r"\s+katalog\s*.*$",
        r"\s+pdf\s*.*$",
        r"\s+çeşitleri\s*.*$",
        r"\s+cesitleri\s*.*$",
        r"\s+nelerdir\s*.*$",
        r"\s+neler\s*.*$",
        r"\s+var\s+mı.*$",
        r"\s+varmı.*$",
        r"\s+var\s*$",
        r"\s+stok\s+listesi.*$",
        r"\s+stokları.*$",
        r"\s+stoklari.*$",
        r"\s+stok\s*$",
        r"\s+lütfen\s*$",
        r"\s+lutfen\s*$",
    ]
    for pat in remove_phrases:
        t_low = re.sub(pat, "", t_low, flags=re.IGNORECASE).strip()

    # Kelime sonlarındaki Türkçe çoğul eklerini (-lar, -ler, -ları, -leri) temizle
    words = t_low.split()
    cleaned_words = []
    for w in words:
        w_clean = re.sub(r"(ları|leri|lar|ler|dan|den|탄|ten|da|de|ta|te)$", "", w)
        if len(w_clean) >= 2:
            cleaned_words.append(w_clean)
        else:
            cleaned_words.append(w)

    return " ".join(cleaned_words).strip(" ?.,!:")


def _normalize_compare(s: str) -> str:
    """Karşılaştırma için metni standartlaştırır."""
    if not s:
        return ""
    s = s.replace("İ", "i").replace("I", "ı").lower()
    tr_map = str.maketrans("çğıöşü", "cgiosu")
    s = s.translate(tr_map)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


class PersistentWatcherState:
    """
    GitHub Actions 5.5 saatte bir yeniden başlasa bile chat_id, aktif lotlar,
    Sold Out lotlar ve Pazartesi/Cuma rapor bilgisini koruyan durum yöneticisi.
    """

    def __init__(self, filepath: str = STATE_FILE):
        self.filepath = filepath
        self.lock = threading.Lock()
        self.data: Dict[str, Any] = {
            "chat_ids": [],
            "initialized": False,
            "active_lots": {},      # folder_id -> {code, stone, collection, name, link, m2, slabs}
            "sold_out_lots": {},    # folder_id -> {code, stone, collection, name, link}
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

            # Ortam değişkeninden tanımlı chat_id varsa ekle
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
        """GitHub Actions üzerinde state dosyasını depoya sessizce yedekler."""
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
    Arslanlar Series, Arslanlar Collection ve Arslanlar Exclusive Collection
    klasörlerini ve içlerindeki Sold Out klasörlerini doğrudan tarayan,
    hem bot.py'nin mevcut önbelleğiyle konuşan hem de Google Drive API'den
    canlı veri çeken akıllı motor.
    """

    LOT_CODE_REGEX = re.compile(r"\b([A-Z]{1,3}\s*[-_]?\s*\d{2,5}[A-Z]?)\b", re.IGNORECASE)

    def __init__(self):
        self.sa_json_str = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
        self._creds = None
        self._token_lock = threading.Lock()
        self.session = requests.Session()
        self.inventory_lock = threading.Lock()
        self.active_lots: Dict[str, Dict[str, Any]] = {}
        self.sold_out_lots: Dict[str, Dict[str, Any]] = {}
        self.last_scan_time: float = 0.0
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

    def _find_bot_drive_explorer(self) -> Optional[Any]:
        """bot.py içindeki aktif DriveExplorer nesnesini RAM'den yakalar."""
        try:
            for obj in gc.get_objects():
                if type(obj).__name__ == "DriveExplorer":
                    return obj
        except Exception:
            pass
        return None

    def _load_bot_existing_caches(self) -> Dict[str, Dict[str, Any]]:
        """
        bot.py'nin oluşturduğu argon_v17_master_cache.json ve RAM önbelleğindeki
        metrekare, plaka, ebat ve yüzey bilgilerini okur.
        """
        combined_meta: Dict[str, Dict[str, Any]] = {}

        def _extract_dict_entries(d: Any):
            if not isinstance(d, dict):
                return
            for k, v in d.items():
                if isinstance(v, dict):
                    # Eğer içinde metraj/plaka veya klasör bilgisi varsa kaydet
                    has_metric = any(
                        mk in v
                        for mk in (
                            "m2", "total_m2", "sqm", "area", "mevcut_stok",
                            "plaka", "slabs", "slab_count", "count",
                            "ebat", "avg_dim", "dimensions", "folder_id", "id", "code"
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
                        _extract_dict_entries(v)
                elif isinstance(v, list):
                    for item in v:
                        if isinstance(item, dict):
                            fid = item.get("folder_id") or item.get("id")
                            code = item.get("code") or item.get("lot_code") or item.get("name")
                            if fid:
                                combined_meta[str(fid).strip()] = item
                            if code:
                                combined_meta[str(code).upper().strip()] = item

        # 1. Disk üzerindeki argon_v17_master_cache.json dosyasını oku
        if os.path.exists(BOT_CACHE_FILE):
            try:
                with open(BOT_CACHE_FILE, "r", encoding="utf-8") as f:
                    disk_cache = json.load(f)
                    _extract_dict_entries(disk_cache)
            except Exception:
                pass

        # 2. RAM'deki DriveExplorer nesnesinin iç önbelleklerini oku
        explorer = self._find_bot_drive_explorer()
        if explorer is not None:
            try:
                for attr_name, attr_val in list(explorer.__dict__.items()):
                    if isinstance(attr_val, dict):
                        _extract_dict_entries(attr_val)
            except Exception:
                pass

        return combined_meta

    def _list_all_drive_folders_and_files(self) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        """Google Drive'daki klasör ve dosyaları hızlı sayfalamayla çeker."""
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
                logger.warning(f"Drive tarama hatasi: {e}")
                break

        folders = [x for x in all_items if x.get("mimeType") == "application/vnd.google-apps.folder"]
        files = [x for x in all_items if x.get("mimeType") != "application/vnd.google-apps.folder"]
        return folders, files

    @classmethod
    def _extract_lot_and_stone_name(
        cls, folder_name: str, parent_names: List[str], child_files: List[Dict[str, Any]]
    ) -> Tuple[str, str]:
        """Klasör adı, üst klasör yolu ve içindeki dosyalardan Lot Kodu ve Taş Çeşidini çıkarır."""
        raw_name = folder_name.strip()
        m = cls.LOT_CODE_REGEX.search(raw_name)
        lot_code = ""
        if m:
            lot_code = re.sub(r"\s+", "", m.group(1)).upper()
            if "-" not in lot_code and any(c.isalpha() for c in lot_code) and any(c.isdigit() for c in lot_code):
                lot_code = re.sub(r"^([A-Z]+)(\d+.*)$", r"\1-\2", lot_code)

        # Klasör adından lot kodunu ve ölçü ibarelerini çıkarıp taş adını bul
        cleaned = raw_name
        if m:
            cleaned = cleaned.replace(m.group(0), " ")
        cleaned = re.sub(r"\b\d+([.,]\d+)?\s*(m2|m²|plaka|slab|slabs|cm|mm|pcs|adet)\b", " ", cleaned, flags=re.I)
        cleaned = re.sub(r"[-_()\[\]/\\|]+", " ", cleaned)
        cleaned = re.sub(r"\s+", " ", cleaned).strip()

        stone_name = cleaned

        # Eğer klasör adı sadece 'A-371' ise taş adını üst klasörden veya içindeki dosyalardan çek
        if len(stone_name) < 3:
            for pname in reversed(parent_names):
                p_norm = _normalize_compare(pname)
                if p_norm not in (
                    "arslanlar series",
                    "arslanlar collection",
                    "arslanlar exclusive collection",
                    "sold out",
                    "soldout",
                ):
                    stone_name = pname.strip()
                    break

        if len(stone_name) < 3 and child_files:
            for cf in child_files:
                cf_name = os.path.splitext(cf.get("name", ""))[0]
                cf_clean = cls.LOT_CODE_REGEX.sub(" ", cf_name)
                cf_clean = re.sub(r"\b(whatsapp|image|img|dsc|photo|foto|plaka|slab|lot)\b", " ", cf_clean, flags=re.I)
                cf_clean = re.sub(r"[\d\-_()\[\].]+", " ", cf_clean)
                cf_clean = re.sub(r"\s+", " ", cf_clean).strip()
                if len(cf_clean) >= 4:
                    stone_name = cf_clean.title()
                    break

        if not lot_code:
            lot_code = raw_name.upper()
        if not stone_name:
            stone_name = raw_name

        return lot_code, stone_name

    def _resolve_metrics(
        self,
        folder_id: str,
        lot_code: str,
        folder_name: str,
        child_files: List[Dict[str, Any]],
        bot_cache: Dict[str, Dict[str, Any]],
    ) -> Dict[str, Any]:
        """
        Taşın m², plaka sayısı, ortalama ebat ve yüzey/kalınlık bilgilerini
        önce bot.py'nin kendi önbelleğinden, yoksa klasör/dosya verilerinden çıkarır.
        """
        m2 = 0.0
        slabs = 0
        avg_dim = "-"
        surface = "Polished (2 cm)"
        stone_from_cache = ""

        # 1. bot.py önbelleğinden kontrol et
        candidates = [
            bot_cache.get(folder_id),
            bot_cache.get(lot_code.upper()),
            bot_cache.get(lot_code.replace("-", "").upper()),
            bot_cache.get(folder_name.upper()),
        ]
        for c in candidates:
            if isinstance(c, dict):
                for k_m2 in ("m2", "total_m2", "sqm", "area", "total_area", "stok_m2"):
                    if c.get(k_m2) is not None:
                        try:
                            val = float(str(c[k_m2]).replace(",", "."))
                            if val > 0:
                                m2 = val
                                break
                        except Exception:
                            pass
                for k_pl in ("plaka", "slabs", "slab_count", "count", "plaka_sayisi", "pcs"):
                    if c.get(k_pl) is not None:
                        try:
                            val = int(float(str(c[k_pl])))
                            if val > 0:
                                slabs = val
                                break
                        except Exception:
                            pass
                for k_dim in ("avg_dim", "ebat", "ortalama_ebat", "dimensions", "dim", "size"):
                    if c.get(k_dim) and str(c.get(k_dim)).strip() not in ("", "-", "None"):
                        avg_dim = str(c.get(k_dim)).strip()
                        break
                for k_surf in ("surface_thickness", "surface", "yuzey", "finish"):
                    if c.get(k_surf) and str(c.get(k_surf)).strip() not in ("", "-", "None"):
                        surface = str(c.get(k_surf)).strip()
                        break
                for k_st in ("stone_name", "stone", "variety", "material", "tas_adi", "title"):
                    if c.get(k_st) and len(str(c.get(k_st)).strip()) >= 3:
                        stone_from_cache = str(c.get(k_st)).strip()
                        break

        # 2. Klasör veya dosya adlarında m2 / plaka yazıyorsa oku
        combined_text = folder_name + " " + " ".join(f.get("name", "") for f in child_files)
        if m2 <= 0:
            m_m2 = re.search(r"(\d+(?:[.,]\d+)?)\s*(?:m2|m²|sqm)", combined_text, flags=re.I)
            if m_m2:
                try:
                    m2 = float(m_m2.group(1).replace(",", "."))
                except Exception:
                    pass
        if slabs <= 0:
            m_pl = re.search(r"(\d+)\s*(?:plaka|slab|slabs|pcs|adet)", combined_text, flags=re.I)
            if m_pl:
                try:
                    slabs = int(m_pl.group(1))
                except Exception:
                    pass

        return {
            "m2": round(m2, 2),
            "slabs": slabs,
            "avg_dim": avg_dim,
            "surface": surface,
            "stone_from_cache": stone_from_cache,
        }

    def scan_drive_collections(self, force: bool = False) -> Tuple[Dict[str, Dict[str, Any]], Dict[str, Dict[str, Any]]]:
        """
        Arslanlar Series, Arslanlar Collection ve Arslanlar Exclusive Collection
        klasörlerini ve içlerindeki Sold Out klasörlerini tarar.
        """
        now = time.time()
        if not force and self.active_lots and (now - self.last_scan_time < 45):
            return self.active_lots, self.sold_out_lots

        folders, files = self._list_all_drive_folders_and_files()
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

        # 3 Ana Koleksiyon kök klasörünü tespit et
        target_norm_map = {_normalize_compare(c): c for c in TARGET_COLLECTIONS}
        root_collection_folders: List[Tuple[str, str]] = []  # (folder_id, canonical_collection_name)

        for f in folders:
            fn_norm = _normalize_compare(f.get("name", ""))
            if fn_norm in target_norm_map:
                root_collection_folders.append((f["id"], target_norm_map[fn_norm]))

        new_active: Dict[str, Dict[str, Any]] = {}
        new_sold_out: Dict[str, Dict[str, Any]] = {}

        def _walk(current_fid: str, collection_name: str, path_names: List[str], in_sold_out: bool):
            subfolders = children_folders.get(current_fid, [])
            c_files = children_files.get(current_fid, [])
            cur_folder = folder_by_id.get(current_fid, {})
            cur_name = cur_folder.get("name", "").strip()
            cur_norm = _normalize_compare(cur_name)

            is_sold_out_dir = cur_norm in ("sold out", "soldout", "satildi", "pasif")
            now_sold_out = in_sold_out or is_sold_out_dir

            # Alt klasörleri gez
            for sf in subfolders:
                sf_name = sf.get("name", "").strip()
                _walk(sf["id"], collection_name, path_names + [sf_name], now_sold_out)

            # Eğer bu klasör kök koleksiyon klasörü veya doğrudan 'Sold Out' kapsayıcı klasörüyse lot değildir
            if len(path_names) <= 1 or is_sold_out_dir:
                return

            # Bir klasörün Lot klasörü olma kriteri:
            # Ya içinde dosyalar (görsel/excel/pdf) vardır, ya yaprak (en alt) klasördür, ya da adında Lot kodu vardır
            has_lot_code = bool(self.LOT_CODE_REGEX.search(cur_name))
            is_leaf_folder = len(subfolders) == 0

            if has_lot_code or is_leaf_folder or len(c_files) > 0:
                # Eğer alt klasörleri de lot olan bir ara kategori klasörüyse (örn: 'Silver Roots' klasörü altında 'A-371' varsa)
                # ve kendi içinde plaka görseli yoksa onu kategori olarak bırak, altındaki lotları al
                image_files = [
                    x for x in c_files
                    if str(x.get("mimeType", "")).startswith("image/")
                    or str(x.get("name", "")).lower().endswith((".jpg", ".jpeg", ".png", ".webp"))
                ]
                if not is_leaf_folder and not has_lot_code and not image_files:
                    return

                lot_code, stone_name = self._extract_lot_and_stone_name(cur_name, path_names[:-1], c_files)
                metrics = self._resolve_metrics(current_fid, lot_code, cur_name, c_files, bot_cache)

                if metrics["stone_from_cache"] and (
                    stone_name.upper() == lot_code.upper() or len(stone_name) < 3
                ):
                    stone_name = metrics["stone_from_cache"]

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
                    "images": image_files,
                    "modified_time": cur_folder.get("modifiedTime", ""),
                }

                if now_sold_out:
                    new_sold_out[current_fid] = record
                else:
                    new_active[current_fid] = record

        for root_fid, coll_name in root_collection_folders:
            _walk(root_fid, coll_name, [coll_name], False)

        with self.inventory_lock:
            self.active_lots = new_active
            self.sold_out_lots = new_sold_out
            self.last_scan_time = now

        return new_active, new_sold_out

    def search_active_by_stone(self, query_text: str) -> Tuple[str, List[Dict[str, Any]]]:
        """
        Taş çeşidine göre (örn: 'Silver Roots', 'Silver River', 'Rosso Levanto')
        SADECE aktif klasörlerdeki lotları bulur. Sold Out klasörlerini kesinlikle dahil etmez.
        """
        cleaned_query = _strip_turkish_plural_and_clean(query_text)
        norm_query = _normalize_compare(cleaned_query)
        if not norm_query:
            return cleaned_query, []

        active_map, _ = self.scan_drive_collections(force=False)
        query_tokens = [t for t in norm_query.split() if len(t) >= 2]

        matched: List[Dict[str, Any]] = []
        for fid, item in active_map.items():
            searchable_blob = _normalize_compare(
                f"{item.get('stone_name', '')} {item.get('folder_name', '')} {item.get('path', '')}"
            )
            # Tüm arama kelimeleri (örn: 'silver' ve 'roots') eşleşmeli
            if query_tokens and all(tok in searchable_blob for tok in query_tokens):
                matched.append(item)

        # Lot koduna göre sırala
        matched.sort(key=lambda x: (x.get("collection", ""), x.get("lot_code", "")))
        return cleaned_query.title(), matched

    def find_lot_by_code_or_name(self, query_str: str) -> Optional[Dict[str, Any]]:
        """Render veya tekil işlem için lotu aktif stoktan bulur."""
        active_map, _ = self.scan_drive_collections(force=False)
        q_norm = _normalize_compare(query_str)
        q_code_compact = re.sub(r"[^a-z0-9]", "", q_norm)

        # 1. Tam lot kodu eşleşmesi (Örn: A-371)
        for item in active_map.values():
            lc_compact = re.sub(r"[^a-z0-9]", "", item.get("lot_code", "").lower())
            if lc_compact and lc_compact == q_code_compact:
                return item

        # 2. Alt metin / taş adı eşleşmesi
        for item in active_map.values():
            blob = _normalize_compare(f"{item.get('lot_code', '')} {item.get('stone_name', '')} {item.get('folder_name', '')}")
            if q_norm and q_norm in blob:
                return item

        return None

    def download_file_bytes(self, file_id: str) -> Optional[bytes]:
        """Google Drive'dan görsel dosyasını güvenli şekilde indirir."""
        headers = self._get_headers()
        if not headers:
            return None
        url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media&supportsAllDrives=true"
        try:
            resp = self.session.get(url, headers=headers, timeout=25)
            if resp.status_code == 200:
                return resp.content
        except Exception as e:
            logger.warning(f"Dosya indirme hatasi ({file_id}): {e}")
        return None


def _build_ram_safe_architectural_render(
    img_bytes: bytes, lot_code: str, stone_name: str, collection_name: str, logo_b64: str = ""
) -> bytes:
    """
    7 GB RAM'i ASLA KİLİTLEMEYEN (<60 MB RAM ayak izi) Mimari İç Mekan + Bookmatch Stüdyo Motoru.
    - Üst Panel: 2'li Bookmatch (Kelebek) Kusursuz Ayna Simetrisi
    - Alt/Ana Panel: 3D Perspektif Lüks Lobi Duvarı, Spot Işık Kırılması ve Mermer Zemin Yansıması
    """
    try:
        with Image.open(io.BytesIO(img_bytes)) as raw_img:
            raw_img = raw_img.convert("RGB")
            # RAM Güvenliği: Plaka görselini çalışma boyutuna küçült
            raw_img.thumbnail((1100, 800), Image.Resampling.LANCZOS)
            w, h = raw_img.size

            # Kenarlardaki sehpa/demir paylarını (%3) temizle
            crop_x = int(w * 0.03)
            crop_y = int(h * 0.03)
            if w > 200 and h > 200:
                slab = raw_img.crop((crop_x, crop_y, w - crop_x, h - crop_y))
            else:
                slab = raw_img.copy()

        # Hafif kontrast ve cila parlaklığı
        slab = ImageEnhance.Contrast(slab).enhance(1.06)
        slab = ImageEnhance.Sharpness(slab).enhance(1.12)

        # 2'li Bookmatch (Sol Plaka + Aynalanmış Sağ Plaka) oluştur
        slab_w, slab_h = slab.size
        bm_single_w = 480
        bm_single_h = 300
        left_slab = slab.resize((bm_single_w, bm_single_h), Image.Resampling.LANCZOS)
        right_slab = ImageOps.mirror(left_slab)

        bookmatch = Image.new("RGB", (bm_single_w * 2, bm_single_h))
        bookmatch.paste(left_slab, (0, 0))
        bookmatch.paste(right_slab, (bm_single_w, 0))

        # Ortadaki kelebek birleşim çizgisine ince derz izi
        bm_draw = ImageDraw.Draw(bookmatch)
        bm_draw.line([(bm_single_w, 0), (bm_single_w, bm_single_h)], fill=(35, 35, 35), width=1)

        # Ana Tuval: 1200 x 900 Lüks Mimari Sahne
        cw, ch = 1200, 900
        canvas = Image.new("RGB", (cw, ch), (18, 18, 20))

        # Arka Plan Duvar & Tavan Tonlaması (Numpy ile hızlı gradyan)
        y_idx, x_idx = np.mgrid[0:ch, 0:cw].astype(np.float32)
        bg_arr = np.zeros((ch, cw, 3), dtype=np.float32)
        bg_arr[:, :, 0] = 24.0
        bg_arr[:, :, 1] = 23.0
        bg_arr[:, :, 2] = 26.0

        # Tavan (0 - 110px)
        ceil_mask = y_idx < 110
        bg_arr[ceil_mask] *= (0.55 + 0.45 * (y_idx[ceil_mask] / 110.0))[:, None]

        # Duvar Mermer Yerleşimi (X: 140..1060, Y: 110..630 -> 920 x 520)
        wall_x1, wall_y1 = 140, 110
        wall_w, wall_h = 920, 520
        wall_bm = bookmatch.resize((wall_w, wall_h), Image.Resampling.LANCZOS)
        wall_arr = np.asarray(wall_bm, dtype=np.float32)

        # 3 Adet Sıcak Mimari Spot Işığı (3000K tavan aydınlatma konileri)
        wy, wx = np.mgrid[0:wall_h, 0:wall_w].astype(np.float32)
        light_map = np.ones((wall_h, wall_w), dtype=np.float32) * 0.72
        spot_centers = [wall_w * 0.22, wall_w * 0.50, wall_w * 0.78]
        for sx in spot_centers:
            cone_spread = 95.0 + wy * 0.42
            dist_x = np.abs(wx - sx)
            vertical_fade = np.clip(1.0 - (wy / (wall_h * 1.15)), 0.15, 1.0)
            spot_intensity = np.exp(-0.5 * (dist_x / cone_spread) ** 2) * vertical_fade
            light_map += spot_intensity * 0.42

        # Kenar Ambient Occlusion (Derinlik gölgesi)
        edge_x = np.minimum(wx, wall_w - wx) / 60.0
        edge_y = np.minimum(wy, wall_h - wy) / 45.0
        ao_mask = np.clip(np.minimum(edge_x, edge_y), 0.55, 1.0)
        light_map = np.clip(light_map * ao_mask, 0.35, 1.22)

        # Sıcak ışık tonu (R, G, B çarpanları)
        wall_lit = np.zeros_like(wall_arr)
        wall_lit[:, :, 0] = np.clip(wall_arr[:, :, 0] * light_map * 1.03, 0, 255)
        wall_lit[:, :, 1] = np.clip(wall_arr[:, :, 1] * light_map * 1.00, 0, 255)
        wall_lit[:, :, 2] = np.clip(wall_arr[:, :, 2] * light_map * 0.95, 0, 255)

        bg_arr[wall_y1 : wall_y1 + wall_h, wall_x1 : wall_x1 + wall_w] = wall_lit

        # Zemin (Y: 630..820) ve Gerçek Zamanlı Mermer Ayna Yansıması
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

        # Görüntüyü PIL formatına aktar
        canvas = Image.fromarray(np.clip(bg_arr, 0, 255).astype(np.uint8))
        draw = ImageDraw.Draw(canvas)

        # Mimari Bronz Çerçeve ve Süpürgelik Çizgileri
        bronze_color = (176, 141, 87)
        draw.rectangle([wall_x1 - 4, wall_y1 - 4, wall_x1 + wall_w + 3, wall_y1 + wall_h + 3], outline=bronze_color, width=3)
        draw.line([(0, floor_y1), (cw, floor_y1)], fill=(70, 62, 50), width=2)

        # Sağ Üst Köşeye "Net Bookmatch Planı" Mini Pano Gösterimi
        mini_w, mini_h = 240, 150
        mini_bm = bookmatch.resize((mini_w, mini_h), Image.Resampling.LANCZOS)
        canvas.paste(mini_bm, (cw - mini_w - 28, 22))
        draw.rectangle([cw - mini_w - 30, 20, cw - 26, 24 + mini_h], outline=bronze_color, width=2)
        draw.rectangle([cw - mini_w - 30, 20, cw - 26, 42], fill=(20, 20, 22))
        draw.text((cw - mini_w - 20, 24), "2'LI BOOKMATCH PLANI", fill=(235, 215, 175))

        # Alt Bilgi Bandı (Footer)
        draw.rectangle([0, 820, cw, ch], fill=(12, 12, 14))
        draw.line([(0, 820), (cw, 820)], fill=bronze_color, width=2)

        title_str = f"ARSLANLAR MARBLE STUDIO  |  {lot_code} - {stone_name.upper()}  ({collection_name})"
        sub_str = "Mimari Bookmatch (Kelebek) & Icim Mekan Duvar / Zemin Yansima Simulasyonu"
        draw.text((35, 836), title_str, fill=(240, 225, 195))
        draw.text((35, 864), sub_str, fill=(165, 165, 170))

        out_buf = io.BytesIO()
        canvas.save(out_buf, format="JPEG", quality=90, optimize=True)
        result_bytes = out_buf.getvalue()

        # Bellek temizliği
        del raw_img, slab, left_slab, right_slab, bookmatch, wall_bm, wall_arr, wall_lit, bg_arr, canvas
        gc.collect()
        return result_bytes
    except Exception as e:
        gc.collect()
        raise RuntimeError(f"Render motoru hatasi: {e}")


def _build_stone_variety_pdf(
    stone_title: str, lots: List[Dict[str, Any]], drive_intel: ArslanlarDriveIntelligence
) -> bytes:
    """
    'Silver Rootslar sunum lütfen' veya 'Rosso levantolar sunum lütfen' denildiğinde
    sadece o taşa ait aktif lotları içeren şık bir PDF sunum kataloğu üretir.
    """
    buf = io.BytesIO()
    page_w, page_h = landscape(A4)
    c = pdf_canvas.Canvas(buf, pagesize=(page_w, page_h))

    total_m2 = sum(x.get("m2", 0.0) for x in lots)
    total_slabs = sum(x.get("slabs", 0) for x in lots)

    # KAPAK SAYFASI
    c.setFillColor(colors.HexColor("#121214"))
    c.rect(0, 0, page_w, page_h, fill=1, stroke=0)

    c.setStrokeColor(colors.HexColor("#B08D57"))
    c.setLineWidth(2)
    c.rect(25, 25, page_w - 50, page_h - 50, fill=0, stroke=1)

    c.setFillColor(colors.HexColor("#B08D57"))
    c.setFont("Helvetica-Bold", 28)
    c.drawCentredString(page_w / 2, page_h - 140, "ARSLANLAR MARBLE")

    c.setFillColor(colors.HexColor("#FFFFFF"))
    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(page_w / 2, page_h - 195, f"{stone_title.upper()} - AKTIF STOK SUNUMU")

    c.setFillColor(colors.HexColor("#CCCCCC"))
    c.setFont("Helvetica", 13)
    c.drawCentredString(
        page_w / 2,
        page_h - 235,
        f"Toplam Aktif Lot: {len(lots)}   |   Toplam Metraj: {total_m2:,.2f} m2 ({total_slabs} Plaka)",
    )
    c.drawCentredString(
        page_w / 2,
        page_h - 260,
        f"Tarih: {datetime.now(TR_TZ).strftime('%d.%m.%Y %H:%M')}  (Sold Out klasorleri harictir)",
    )
    c.showPage()

    # HER LOT İÇİN SUNUM SAYFASI (En fazla 25 lot, RAM korumalı)
    for item in lots[:25]:
        c.setFillColor(colors.HexColor("#161618"))
        c.rect(0, 0, page_w, page_h, fill=1, stroke=0)

        c.setStrokeColor(colors.HexColor("#B08D57"))
        c.setLineWidth(1.5)
        c.rect(20, 20, page_w - 40, page_h - 40, fill=0, stroke=1)

        lot_code = item.get("lot_code", "-")
        stone_name = item.get("stone_name", stone_title)
        coll = item.get("collection", "")
        m2 = item.get("m2", 0.0)
        slabs = item.get("slabs", 0)
        ebat = item.get("avg_dim", "-")
        surf = item.get("surface", "Polished (2 cm)")
        link = item.get("drive_link", "")

        # Üst Başlık
        c.setFillColor(colors.HexColor("#B08D57"))
        c.setFont("Helvetica-Bold", 18)
        c.drawString(40, page_h - 55, f"{lot_code}  |  {stone_name}")

        c.setFillColor(colors.HexColor("#AAAAAA"))
        c.setFont("Helvetica", 12)
        c.drawRightString(page_w - 40, page_h - 55, f"Koleksiyon: {coll}")

        # Sol Metrik Kutusu
        c.setFillColor(colors.HexColor("#FFFFFF"))
        c.setFont("Helvetica-Bold", 12)
        c.drawString(40, page_h - 90, f"Mevcut Stok: {m2:,.2f} m2 ({slabs} Plaka)")
        c.drawString(300, page_h - 90, f"Ortalama Ebat: {ebat}")
        c.drawString(530, page_h - 90, f"Yuzey & Kalinlik: {surf}")

        # Görsel Varsa İndir ve Küçülterek PDF'e Bas
        images = item.get("images", [])
        if images:
            raw_b = drive_intel.download_file_bytes(images[0]["id"])
            if raw_b:
                try:
                    with Image.open(io.BytesIO(raw_b)) as im:
                        im = im.convert("RGB")
                        im.thumbnail((1000, 650), Image.Resampling.LANCZOS)
                        img_io = io.BytesIO()
                        im.save(img_io, format="JPEG", quality=85)
                        img_io.seek(0)
                        c.drawImage(
                            ImageReader(img_io),
                            40,
                            65,
                            width=page_w - 80,
                            height=page_h - 175,
                            preserveAspectRatio=True,
                            anchor="c",
                        )
                except Exception:
                    pass

        if link:
            c.setFillColor(colors.HexColor("#66B2FF"))
            c.setFont("Helvetica-Bold", 11)
            c.drawString(40, 38, f"Google Drive Klasoru: {link}")
            c.linkURL(link, (40, 32, page_w - 40, 52), relative=0)

        c.showPage()

    c.save()
    pdf_bytes = buf.getvalue()
    gc.collect()
    return pdf_bytes


class ArslanlarPluginController:
    """
    Tüm yeni özellikleri (Taş Çeşidi Sorgu, Taşa Özel Sunum, Sold Out & Yeni Lot
    Anlık Bildirimi, Pazartesi/Cuma Özet Rapor ve RAM-Korumalı Render) yöneten ana sınıf.
    """

    def __init__(self, app: Application):
        self.app = app
        self.state = PersistentWatcherState()
        self.intel = ArslanlarDriveIntelligence()
        self.bot_token = os.getenv("TELEGRAM_TOKEN", "").strip()

    def send_telegram_sync(self, text: str, parse_mode: str = "HTML"):
        """Arka plan thread'lerinden kayıtlı tüm sohbetlere güvenli mesaj gönderir."""
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
        """Pazartesi ve Cuma günleri (veya /ozet komutuyla) gönderilecek detaylı stok raporu."""
        active_map, sold_map = self.intel.scan_drive_collections(force=True)

        total_lots = len(active_map)
        total_m2 = sum(x.get("m2", 0.0) for x in active_map.values())
        total_slabs = sum(x.get("slabs", 0) for x in active_map.values())

        lines = [
            f"📊 <b>ARSLANLAR MARBLE — {title_prefix} STOK & KOLEKSİYON ÖZETİ</b>",
            f"🗓 <i>{datetime.now(TR_TZ).strftime('%d.%m.%Y %H:%M')}</i>\n",
            f"🏛 <b>Genel Aktif Stok:</b> <code>{total_m2:,.2f} m²</code> (<code>{total_slabs} Plaka</code> / <code>{total_lots} Lot</code>)",
            f"📦 <b>Toplam Arşivlenen (Sold Out):</b> <code>{len(sold_map)} Lot</code>\n",
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

        # En çok stoğu olan ilk 8 taş çeşidi
        variety_totals: Dict[str, Dict[str, Any]] = {}
        for x in active_map.values():
            st = x.get("stone_name", "Diğer").strip().title()
            entry = variety_totals.setdefault(st, {"lots": 0, "m2": 0.0, "slabs": 0})
            entry["lots"] += 1
            entry["m2"] += x.get("m2", 0.0)
            entry["slabs"] += x.get("slabs", 0)

        top_varieties = sorted(
            variety_totals.items(), key=lambda kv: (kv[1]["m2"], kv[1]["lots"]), reverse=True
        )[:8]

        if top_varieties:
            lines.append("\n💎 <b>Öne Çıkan Taş Çeşitleri (Aktif Stok):</b>")
            for st_name, st_info in top_varieties:
                lines.append(
                    f"  ▫️ <b>{st_name}:</b> {st_info['lots']} Lot — "
                    f"<code>{st_info['m2']:,.2f} m²</code> ({st_info['slabs']} Plaka)"
                )

        return "\n".join(lines)

    def _background_watcher_loop(self):
        """
        Arka planda çalışan:
        1) Sold Out klasörüne taşınan taşları yakalayıp anında bildiren
        2) 3 ana koleksiyona yeni eklenen lotları bildiren
        3) Her Pazartesi ve Cuma otomatik özet rapor geçen gözcü döngüsü.
        """
        time.sleep(12)  # bot.py'nin ilk açılışını tamamlamasını bekle
        logger.info("🔍 Arslanlar Drive Gozcusu (Sold Out & Yeni Lot & Pzt/Cuma Rapor) baslatildi.")

        while True:
            try:
                active_map, sold_map = self.intel.scan_drive_collections(force=True)

                if active_map or sold_map:
                    if not self.state.data.get("initialized"):
                        # İlk kurulum: mevcut durumu kaydet (yüzlerce eski lotu yeniymiş gibi spamlamamak için)
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

                        # 1. KONTROL: Sold Out klasörüne yeni taşınan lotlar
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

                        # 2. KONTROL: Aktif Koleksiyonlara yeni eklenen lotlar
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

                        # Güncel durumu kaydet
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

                # 3. KONTROL: Her Pazartesi (0) ve Cuma (4) saat 09:00 sonrası Otomatik Özet Rapor
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

            time.sleep(120)  # 2 dakikada bir kontrol et

    async def handle_priority_message(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """
        group=-1 üzerinde çalışır.
        Eğer mesaj yeni özelliklerden biriyse (Taş çeşidi sorgu, Render, Özet Rapor, Taşa özel sunum)
        yanıtlar ve ApplicationHandlerStop ile sonlandırır.
        Eğer mesaj mevcut bot.py'nin komutuysa (A-371, Arslanlar Series sunum vb.) HİÇ DOKUNMAZ,
        böylece bot.py kendi orijinal akışıyla çalışmaya devam eder!
        """
        if not update.effective_message or not update.effective_message.text:
            return

        # Otomatik bildirimler için sohbet ID'sini kaydet
        if update.effective_chat:
            self.state.register_chat(update.effective_chat.id)

        raw_text = update.effective_message.text.strip()
        norm_text = _normalize_compare(raw_text)

        # ---------------------------------------------------------------------
        # 1. ÖZET RAPOR İSTEĞİ ("özet rapor", "/ozet", "haftalık özet")
        # ---------------------------------------------------------------------
        if norm_text in ("ozet", "ozet rapor", "stok ozet", "haftalik ozet", "genel stok") or raw_text.lower().startswith("/ozet"):
            status_msg = await update.effective_message.reply_text("⏳ Koleksiyonlar taranıyor, özet rapor hazırlanıyor...")
            loop = asyncio.get_running_loop()
            report_html = await loop.run_in_executor(None, self.format_executive_summary_report, "ANLIK")
            await status_msg.edit_text(report_html, parse_mode="HTML", disable_web_page_preview=True)
            raise ApplicationHandlerStop

        # ---------------------------------------------------------------------
        # 2. RAM-KORUMALI RENDER & BOOKMATCH İSTEĞİ ("render A-371", "A-371 render yap")
        # ---------------------------------------------------------------------
        is_render_req = bool(
            re.search(r"\b(render|bookmatch|kelebek)\b", norm_text)
        )
        if is_render_req:
            target_q = re.sub(
                r"\b(render|bookmatch|kelebek|yap|lütfen|lutfen|hazırla|hazirla|stüdyo|studyo)\b",
                " ",
                raw_text,
                flags=re.I,
            ).strip(" /.,:-")

            if not target_q:
                await update.effective_message.reply_text(
                    "🎨 Lütfen render yapılacak lot kodunu veya taş adını yazın.\nÖrnek: <code>render A-371</code>",
                    parse_mode="HTML",
                )
                raise ApplicationHandlerStop

            status_msg = await update.effective_message.reply_text(
                f"🎨 <b>{target_q.upper()}</b> için RAM korumalı Mimari Sahne & Bookmatch Render hazırlanıyor...",
                parse_mode="HTML",
            )

            loop = asyncio.get_running_loop()
            lot_item = await loop.run_in_executor(None, self.intel.find_lot_by_code_or_name, target_q)

            if not lot_item or not lot_item.get("images"):
                await status_msg.edit_text(
                    f"⚠️ <b>{target_q}</b> için aktif stok klasörlerinde uygun plaka fotoğrafı bulunamadı.",
                    parse_mode="HTML",
                )
                raise ApplicationHandlerStop

            first_img_id = lot_item["images"][0]["id"]
            img_bytes = await loop.run_in_executor(None, self.intel.download_file_bytes, first_img_id)
            if not img_bytes:
                await status_msg.edit_text("⚠️ Plaka görseli Google Drive'dan indirilemedi.")
                raise ApplicationHandlerStop

            try:
                rendered_bytes = await loop.run_in_executor(
                    RENDER_EXECUTOR,
                    _build_ram_safe_architectural_render,
                    img_bytes,
                    lot_item["lot_code"],
                    lot_item["stone_name"],
                    lot_item["collection"],
                    "",
                )
                caption = (
                    f"🏛 <b>{lot_item['lot_code']} — {lot_item['stone_name']}</b>\n"
                    f"• <b>Koleksiyon:</b> {lot_item['collection']}\n"
                    f"• <b>Mevcut Stok:</b> <code>{lot_item['m2']:.2f} m² ({lot_item['slabs']} Plaka)</code>\n"
                    f"• <b>Tasarım:</b> 2'li Bookmatch (Kelebek) & Mimari Sahne Yansıma Simülasyonu"
                )
                await update.effective_message.reply_photo(
                    photo=io.BytesIO(rendered_bytes),
                    caption=caption,
                    parse_mode="HTML",
                )
                await status_msg.delete()
            except Exception as e:
                await status_msg.edit_text(f"⚠️ Render oluşturulurken hata: {e}")

            raise ApplicationHandlerStop

        # ---------------------------------------------------------------------
        # 3. TAŞ ÇEŞİDİNE ÖZEL SUNUM ("Silver Rootslar sunum lütfen", "Rosso levantolar sunum lütfen")
        # Not: Eğer kullanıcı 3 ana koleksiyon adını yazdıysa ("Arslanlar Series sunum"),
        # hiç dokunmadan mevcut bot.py'ye bırak!
        # ---------------------------------------------------------------------
        has_sunum_kw = bool(re.search(r"\b(sunum|sunumu|katalog|kataloğu)\b", norm_text))
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
                    f"📑 <b>{stone_title}</b> için aktif stoklar taranıyor ({len(matched_lots)} Lot bulundu), özel sunum PDF'i hazırlanıyor...",
                    parse_mode="HTML",
                )
                await self._send_stone_variety_summary(update, stone_title, matched_lots, send_photos=False)

                try:
                    pdf_bytes = await asyncio.get_running_loop().run_in_executor(
                        RENDER_EXECUTOR, _build_stone_variety_pdf, stone_title, matched_lots, self.intel
                    )
                    safe_fn = re.sub(r"[^a-zA-Z0-9_]", "_", stone_title) + "_Aktif_Stok_Sunumu.pdf"
                    await update.effective_message.reply_document(
                        document=io.BytesIO(pdf_bytes),
                        filename=safe_fn,
                        caption=f"💎 <b>{stone_title}</b> — Özel Seçki Aktif Stok Sunumu ({len(matched_lots)} Lot)",
                        parse_mode="HTML",
                    )
                    await status_msg.delete()
                except Exception as e:
                    await status_msg.edit_text(f"⚠️ PDF oluşturulurken hata: {e}")

                raise ApplicationHandlerStop
            # Eşleşme yoksa bot.py'nin kendi sunum motorunun denemesi için serbest bırak

        # ---------------------------------------------------------------------
        # 4. TAŞ ÇEŞİDİ STOK & ÇEŞİT SORGUSU
        # ("Ne kadar Silver Roots var?", "Ne kadar Silver River var?", "Silver Roots çeşitleri")
        # ---------------------------------------------------------------------
        is_variety_query = bool(
            re.search(
                r"(^ne\s+kadar\s+|var\s+m[ıi]|kaç\s+m2|kac\s+m2|\bçeşitleri\b|\bcesitleri\b|\bstok\s+listesi\b|\bneler\s+var\b|\bvar\s*\??$)",
                raw_text,
                flags=re.I,
            )
        )
        # Tekil lot kodu sorgusuysa (örn: 'A-371') kesinlikle mevcut bot.py'ye bırak!
        is_single_lot_code = bool(re.match(r"^[A-Za-z]{1,3}\s*[-_]?\s*\d{2,5}[A-Za-z]?$", raw_text))

        if is_variety_query and not is_single_lot_code:
            stone_title, matched_lots = await asyncio.get_running_loop().run_in_executor(
                None, self.intel.search_active_by_stone, raw_text
            )
            if not matched_lots:
                await update.effective_message.reply_text(
                    f"🔍 <b>{stone_title}</b> için aktif koleksiyon klasörlerinde (<i>Arslanlar Series, Collection, Exclusive</i>) "
                    f"aktif stok bulunamadı.\n<i>(Sold Out klasöründeki taşlar listeye dahil edilmemiştir.)</i>",
                    parse_mode="HTML",
                )
                raise ApplicationHandlerStop

            await self._send_stone_variety_summary(update, stone_title, matched_lots, send_photos=True)
            raise ApplicationHandlerStop

    async def _send_stone_variety_summary(
        self, update: Update, stone_title: str, matched_lots: List[Dict[str, Any]], send_photos: bool = True
    ):
        """Bulunan aktif taş çeşitlerinin metrajlarını, Drive linklerini ve önizleme görsellerini gönderir."""
        total_m2 = sum(x.get("m2", 0.0) for x in matched_lots)
        total_slabs = sum(x.get("slabs", 0) for x in matched_lots)

        lines = [
            f"💎 <b>{stone_title.upper()} — AKTİF STOK & ÇEŞİTLER LİSTESİ</b>",
            f"📊 <b>Toplam Aktif Stok:</b> <code>{total_m2:,.2f} m² ({total_slabs} Plaka)</code> — <b>{len(matched_lots)} Aktif Lot</b>",
            "<i>(Sold Out klasöründeki taşlar hariç tutulmuştur)</i>\n",
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

            lines.append(
                f"🔹 <b>{l_code}</b> — {s_name} <i>({coll})</i>\n"
                f"  • Mevcut Stok: <code>{m2:.2f} m² ({slabs} Plaka)</code>\n"
                f"  • Ortalama Ebat: <code>{ebat}</code> | {surf}\n"
                f"  • Drive Linki: <a href='{link}'>Klasörü Aç</a>\n"
            )

        full_text = "\n".join(lines)
        # Telegram 4096 karakter sınırına karşı güvenli bölme
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

        # İlk 4 aktif lotun plaka görselini albüm olarak gönder
        if send_photos:
            media_group = []
            loop = asyncio.get_running_loop()
            for item in matched_lots[:4]:
                imgs = item.get("images", [])
                if not imgs:
                    continue
                b = await loop.run_in_executor(None, self.intel.download_file_bytes, imgs[0]["id"])
                if b:
                    cap = f"💎 {item['lot_code']} ({item['stone_name']}) — {item['m2']:.2f} m² ({item['slabs']} Plaka)"
                    media_group.append(InputMediaPhoto(media=io.BytesIO(b), caption=cap))

            if media_group:
                try:
                    await update.effective_message.reply_media_group(media=media_group)
                except Exception as e:
                    logger.warning(f"Medya grubu gonderim uyarisi: {e}")


import asyncio

# =========================================================================
# SIFIR MÜDAHALE (MONKEY-PATCH) KANCASI:
# bot.py kendi içinde Application.run_polling() çağırdığı an devreye girer!
# =========================================================================
_ORIGINAL_RUN_POLLING = Application.run_polling
_PLUGIN_ATTACHED = False


def _hooked_run_polling(self: Application, *args, **kwargs):
    global _PLUGIN_ATTACHED
    if not _PLUGIN_ATTACHED:
        _PLUGIN_ATTACHED = True
        try:
            controller = ArslanlarPluginController(self)
            # group=-1 sayesinde mevcut bot.py handler'larını bozmadan önce filtreler
            self.add_handler(
                MessageHandler(filters.TEXT, controller.handle_priority_message),
                group=-1,
            )
            # Arka plan Drive Gözcüsü ve Pazartesi/Cuma Raporlayıcısını başlat
            watcher_thread = threading.Thread(
                target=controller._background_watcher_loop,
                name="ArslanlarDriveWatcher",
                daemon=True,
            )
            watcher_thread.start()
            logger.info("✅ Arslanlar V18 Eklenti Modulu (Sifir Mudahale) basariyla devreye girdi!")
        except Exception as e:
            logger.error(f"Eklenti baslatma hatasi: {e}")

    return _ORIGINAL_RUN_POLLING(self, *args, **kwargs)


Application.run_polling = _hooked_run_polling
