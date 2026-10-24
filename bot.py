import os
import io
import re
import json
import base64
import logging
import asyncio
import threading
from collections import deque
from typing import Dict, List, Optional, Tuple, Any

import requests
import pandas as pd
import openpyxl
from flask import Flask
from PIL import Image

from telegram import Update, InputMediaPhoto
from telegram.constants import ParseMode, ChatAction
from telegram.ext import (
    ApplicationBuilder,
    ContextTypes,
    MessageHandler,
    filters,
)

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# =====================================================================
# 1. LOGLAMA VE TEMEL AYARLAR
# =====================================================================
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("ArslanlarArgonBot")

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GOOGLE_SA_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()

# Render'daki GOOGLE_DRIVE_LINK veya DEFAULT_DRIVE_LINK değişkenini otomatik tanır
DEFAULT_DRIVE_LINK = (
    os.getenv("GOOGLE_DRIVE_LINK")
    or os.getenv("DEFAULT_DRIVE_LINK")
    or "https://drive.google.com"
).strip()

EXCEL_FILE_PATH = os.getenv("EXCEL_FILE_PATH", "stok.xlsx").strip()

HYPERLINKS_DB_FILE = "hyperlinks_db.json"
VISUAL_CACHE_FILE = "visual_cache.json"
STONE_DNA_DB_FILE = "stone_dna_db.json"

# Kişi/Tedarikçi isimleri ve özet satırları için Kara Liste (Veri Körlüğü Koruması)
BLACKLIST_NAMES = {
    "SERKAN CELIK",
    "SERKAN ÇELİK",
    "TOPLAM",
    "GENEL TOPLAM",
    "ARA TOPLAM",
    "SUBTOTAL",
    "TOTAL",
    "IADE",
    "İADE",
    "FIRE",
    "FİRE",
    "BELIRSIZ",
    "BELİRSİZ",
    "TEST",
    "ORNEK",
    "ÖRNEK",
    "NONE",
    "NAN",
}

# Bilinen Lot -> Taş Cinsi ve Görsel DNA Eşleşmeleri (Başlangıç Hafızası)
DEFAULT_STONE_DNA = {
    "A-283": {
        "stone_name": "VIOLET BRECCIA",
        "visual_dna": "Beyaz ve krem parçalı zemin üzerinde yoğun mor, viyole ve koyu gri breccia (parçalı/kırıklı) damar yapısı.",
        "keywords": ["violet", "breccia", "mor", "kırıklı", "a-283", "a283"],
    },
    "SILVER ROOTS": {
        "stone_name": "SILVER ROOTS",
        "visual_dna": "Gümüş gri ve füme zemin üzerinde kök şeklinde dağılan ince beyaz, altın ve koyu gri çizgisel/ağsı damarlar.",
        "keywords": ["silver roots", "gümüş", "gri", "kök damar", "füme"],
    },
    "BLUE ROOTS": {
        "stone_name": "BLUE ROOTS",
        "visual_dna": "Mavi-gri soğuk tonlu zemin üzerinde koyu gri ve beyaz kök damar geçişleri.",
        "keywords": ["blue roots", "mavi", "gri", "damarlı"],
    },
    "ROSSO LEVANTO": {
        "stone_name": "ROSSO LEVANTO",
        "visual_dna": "Bordo, vişne çürüğü ve koyu kırmızı zemin üzerinde belirgin beyaz kalsit damar ağları.",
        "keywords": ["rosso levanto", "bordo", "kırmızı", "vişne", "beyaz damar"],
    },
    "TERRA MATER": {
        "stone_name": "TERRA MATER",
        "visual_dna": "Toprak tonları, kahverengi, bej ve sıcak gri geçişli, doğal dalgalı ve hareketli yapı.",
        "keywords": ["terra mater", "toprak", "kahve", "dalgalı"],
    },
    "SILVER RIVER": {
        "stone_name": "SILVER RIVER",
        "visual_dna": "Nehir akışını andıran paralel ve dalgalı gümüş-gri, beyaz ve siyah katmanlı damar yapısı.",
        "keywords": ["silver river", "nehir", "çizgisel", "dalgalı gri"],
    },
    "DEKART": {
        "stone_name": "DEKART",
        "visual_dna": "Modern gri ve antrasit tonlarda homojen ve hafif bulutlu/çizgisel mermer dokusu.",
        "keywords": ["dekart", "antrasit", "gri"],
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
    """
    'SİLVER ROOTS' ve 'SILVER ROOTS' gibi yazım farklarını tek bir standart
    anahtara ('SILVER ROOTS') dönüştürür.
    """
    if text is None:
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
    return s


def normalize_lot_code(code: str) -> str:
    """'a 283', 'A283', 'a-283' -> 'A-283' standardına çevirir."""
    raw = normalize_turkish_key(code)
    match = re.match(r"^([A-Z]+)[-\s]*(\d+)$", raw)
    if match:
        return f"{match.group(1)}-{match.group(2)}"
    return raw


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


# =====================================================================
# 3. GOOGLE DRIVE OTONOM GEZGİN VE GÖRSEL İNDİRİCİ
# =====================================================================
class DriveExplorer:
    def __init__(self, sa_json_str: str):
        self.service = None
        self.hyperlinks_db: Dict[str, str] = load_json_file(HYPERLINKS_DB_FILE, {})
        self.visual_cache: Dict[str, List[str]] = load_json_file(VISUAL_CACHE_FILE, {})
        self._init_service(sa_json_str)

    def _init_service(self, sa_json_str: str) -> None:
        if not sa_json_str:
            logger.warning("GOOGLE_SERVICE_ACCOUNT_JSON tanımlı değil!")
            return
        try:
            info = json.loads(sa_json_str)
            creds = service_account.Credentials.from_service_account_info(
                info, scopes=["https://www.googleapis.com/auth/drive.readonly"]
            )
            self.service = build("drive", "v3", credentials=creds, cache_discovery=False)
            logger.info("Google Drive API bağlantısı başarıyla kuruldu.")
        except Exception as e:
            logger.error(f"Google Drive API başlatma hatası: {e}")

    def save_custom_link(self, key: str, url: str) -> None:
        norm_key = normalize_lot_code(key)
        self.hyperlinks_db[norm_key] = url.strip()
        save_json_file(HYPERLINKS_DB_FILE, self.hyperlinks_db)

    def find_folder(self, query_term: str) -> Tuple[Optional[str], Optional[str]]:
        norm_key = normalize_lot_code(query_term)
        raw_clean = query_term.strip().replace("'", "")

        if norm_key in self.hyperlinks_db:
            cached_url = self.hyperlinks_db[norm_key]
            folder_id = self._extract_folder_id_from_url(cached_url)
            if cached_url and cached_url != DEFAULT_DRIVE_LINK:
                return folder_id, cached_url

        if not self.service:
            return None, None

        candidates = [raw_clean, norm_key]
        if "-" in norm_key:
            candidates.append(norm_key.replace("-", " "))
            candidates.append(norm_key.replace("-", ""))
        candidates = list(dict.fromkeys([c for c in candidates if c]))

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
                        pageSize=10,
                    )
                    .execute()
                )
                files = res.get("files", [])
                if files:
                    best = files[0]
                    for f in files:
                        if normalize_lot_code(f["name"]) == norm_key:
                            best = f
                            break
                    folder_id = best["id"]
                    link = best.get(
                        "webViewLink", f"https://drive.google.com/drive/folders/{folder_id}"
                    )
                    self.hyperlinks_db[norm_key] = link
                    save_json_file(HYPERLINKS_DB_FILE, self.hyperlinks_db)
                    return folder_id, link
        except Exception as e:
            logger.error(f"Drive klasör arama hatası ({query_term}): {e}")

        return None, None

    def get_folder_images(self, folder_id: str, max_images: int = 3) -> List[bytes]:
        if not self.service or not folder_id:
            return []

        image_bytes_list: List[bytes] = []
        try:
            q = (
                f"'{folder_id}' in parents and trashed = false "
                f"and (mimeType contains 'image/jpeg' or mimeType contains 'image/png')"
            )
            res = (
                self.service.files()
                .list(
                    q=q,
                    spaces="drive",
                    fields="files(id, name, mimeType, size)",
                    includeItemsFromAllDrives=True,
                    supportsAllDrives=True,
                    pageSize=max_images,
                )
                .execute()
            )
            files = res.get("files", [])

            if not files:
                sub_q = (
                    f"'{folder_id}' in parents and trashed = false "
                    f"and mimeType = 'application/vnd.google-apps.folder'"
                )
                sub_res = (
                    self.service.files()
                    .list(
                        q=sub_q,
                        spaces="drive",
                        fields="files(id, name)",
                        includeItemsFromAllDrives=True,
                        supportsAllDrives=True,
                        pageSize=3,
                    )
                    .execute()
                )
                for sub in sub_res.get("files", []):
                    img_q = (
                        f"'{sub['id']}' in parents and trashed = false "
                        f"and (mimeType contains 'image/jpeg' or mimeType contains 'image/png')"
                    )
                    img_res = (
                        self.service.files()
                        .list(
                            q=img_q,
                            spaces="drive",
                            fields="files(id, name, mimeType, size)",
                            includeItemsFromAllDrives=True,
                            supportsAllDrives=True,
                            pageSize=max_images,
                        )
                        .execute()
                    )
                    files.extend(img_res.get("files", []))
                    if len(files) >= max_images:
                        break

            for f in files[:max_images]:
                request = self.service.files().get_media(
                    fileId=f["id"], supportsAllDrives=True
                )
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
                fh.seek(0)
                image_bytes_list.append(fh.read())

        except Exception as e:
            logger.error(f"Drive görsel indirme hatası (folder_id={folder_id}): {e}")

        return image_bytes_list

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
# 4. AKILLI EXCEL STOK VE KATALOG MOTORU
# =====================================================================
class InventoryEngine:
    def __init__(self, excel_path: str, drive_explorer: DriveExplorer):
        self.excel_path = excel_path
        self.drive = drive_explorer
        self.stones_summary: Dict[str, Dict[str, Any]] = {}
        self.lots_index: Dict[str, Dict[str, Any]] = {}
        self.stone_dna_db: Dict[str, Dict[str, Any]] = load_json_file(
            STONE_DNA_DB_FILE, DEFAULT_STONE_DNA
        )
        self.reload_excel()

    def reload_excel(self) -> None:
        if not os.path.exists(self.excel_path):
            logger.warning(f"Excel dosyası bulunamadı: {self.excel_path}")
            return

        try:
            wb = openpyxl.load_workbook(self.excel_path, data_only=True)
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.hyperlink and cell.hyperlink.target and cell.value:
                            key = normalize_lot_code(str(cell.value))
                            self.drive.save_custom_link(key, cell.hyperlink.target)

            xls = pd.ExcelFile(self.excel_path)
            stones_agg: Dict[str, Dict[str, Any]] = {}
            lots_agg: Dict[str, Dict[str, Any]] = {}

            for sheet_name in xls.sheet_names:
                df = pd.read_excel(self.excel_path, sheet_name=sheet_name)
                if df.empty:
                    continue

                col_map = {col: normalize_turkish_key(col) for col in df.columns}
                stone_col = self._find_col(
                    col_map, ["TAS", "CINS", "URUN", "MERMER", "STONE", "MATERIAL", "AD"]
                )
                lot_col = self._find_col(
                    col_map, ["LOT", "PARTI", "BLOK", "BUNDLE", "KOD", "NO"]
                )
                prod_col = self._find_col(
                    col_map, ["URETIM", "TOPLAM M2", "KESIM", "PRODUCTION"]
                )
                stock_col = self._find_col(
                    col_map, ["STOK", "KALAN", "MEVCUT", "STOCK", "NET M2"]
                )
                slab_col = self._find_col(
                    col_map, ["PLAKA", "ADET", "SLAB", "PCS", "SAYI"]
                )
                thick_col = self._find_col(
                    col_map, ["KALINLIK", "THICKNESS", "CM", "MM"]
                )
                dim_col = self._find_col(col_map, ["EBAT", "OLCU", "BOYUT", "DIM"])

                if not stone_col and len(df.columns) > 0:
                    stone_col = df.columns[0]

                for _, row in df.iterrows():
                    raw_stone = row.get(stone_col) if stone_col else None
                    norm_stone = normalize_turkish_key(raw_stone)

                    if (
                        not norm_stone
                        or norm_stone in BLACKLIST_NAMES
                        or any(b in norm_stone for b in ["TOPLAM", "SERKAN CELIK"])
                        or len(norm_stone) < 2
                    ):
                        continue

                    raw_lot = row.get(lot_col) if lot_col else None
                    norm_lot = normalize_lot_code(raw_lot) if pd.notna(raw_lot) else ""

                    if norm_lot in self.stone_dna_db and norm_stone in ["", "BELIRSIZ"]:
                        norm_stone = self.stone_dna_db[norm_lot]["stone_name"]

                    prod_m2 = safe_float(row.get(prod_col)) if prod_col else 0.0
                    stock_m2 = safe_float(row.get(stock_col)) if stock_col else 0.0
                    slabs = int(safe_float(row.get(slab_col))) if slab_col else 0
                    thickness = (
                        str(row.get(thick_col)).strip()
                        if thick_col and pd.notna(row.get(thick_col))
                        else "2cm"
                    )
                    dims = (
                        str(row.get(dim_col)).strip()
                        if dim_col and pd.notna(row.get(dim_col))
                        else ""
                    )

                    if norm_stone not in stones_agg:
                        stones_agg[norm_stone] = {
                            "stone_name": norm_stone,
                            "production_m2": 0.0,
                            "stock_m2": 0.0,
                            "slabs": 0,
                            "lots": [],
                        }
                    stones_agg[norm_stone]["production_m2"] += prod_m2
                    stones_agg[norm_stone]["stock_m2"] += stock_m2
                    stones_agg[norm_stone]["slabs"] += slabs

                    if norm_lot and norm_lot not in ["NAN", "NONE"]:
                        if norm_lot not in stones_agg[norm_stone]["lots"]:
                            stones_agg[norm_stone]["lots"].append(norm_lot)
                        lots_agg[norm_lot] = {
                            "lot": norm_lot,
                            "stone_name": norm_stone,
                            "production_m2": prod_m2,
                            "stock_m2": stock_m2,
                            "slabs": slabs,
                            "thickness": thickness,
                            "dimensions": dims,
                        }

            self.stones_summary = stones_agg
            self.lots_index = lots_agg
            logger.info(
                f"Excel yüklendi: {len(self.stones_summary)} tekil taş türü, "
                f"{len(self.lots_index)} lot kaydı."
            )
        except Exception as e:
            logger.error(f"Excel okuma hatası: {e}")

    @staticmethod
    def _find_col(col_map: Dict[Any, str], keywords: List[str]) -> Optional[Any]:
        for orig_col, norm_col in col_map.items():
            for kw in keywords:
                if kw in norm_col:
                    return orig_col
        return None

    def get_context_snapshot(self) -> str:
        lines = ["=== ARSLANLAR MARBLE CANLI STOK VE ÜRETİM ÖZETİ ==="]
        if self.stones_summary:
            sorted_stones = sorted(
                self.stones_summary.values(),
                key=lambda x: (x["stock_m2"], x["production_m2"]),
                reverse=True,
            )
            for idx, item in enumerate(sorted_stones[:25], start=1):
                lots_str = (
                    f" | Lotlar: {', '.join(item['lots'][:8])}"
                    if item["lots"]
                    else ""
                )
                lines.append(
                    f"{idx}. {item['stone_name']} -> "
                    f"Stok: {item['stock_m2']:.2f} m² | "
                    f"Toplam Üretim: {item['production_m2']:.2f} m² | "
                    f"Plaka: {item['slabs']}{lots_str}"
                )
        else:
            lines.append("Excel tablosu henüz yüklenmedi veya okunamadı.")

        lines.append("\n=== BİLİNEN ÖZEL LOT VE GÖRSEL DNA KATALOĞU ===")
        for code, info in self.stone_dna_db.items():
            lot_data = self.lots_index.get(code, {})
            st_m2 = lot_data.get("stock_m2", "Bilinmiyor")
            slabs = lot_data.get("slabs", "Bilinmiyor")
            lines.append(
                f"- {code} ({info.get('stone_name')}): {info.get('visual_dna')} "
                f"[Stok: {st_m2} m², Plaka: {slabs}]"
            )
        return "\n".join(lines)


# =====================================================================
# 5. YAPAY ZEKA ORKESTRATÖRÜ (DOĞRUDAN REST API - AQ. AUTH KEY UYUMLU)
# =====================================================================
class ArgonBrain:
    def __init__(self, api_key: str, inventory: InventoryEngine, drive: DriveExplorer):
        self.api_key = api_key
        self.inventory = inventory
        self.drive = drive
        self.chat_memory: Dict[int, deque] = {}
        # Senin gönderdiğin cURL çıktısındaki resmi endpoint:
        self.endpoint = (
            "https://generativelanguage.googleapis.com/v1beta/models/"
            "gemini-flash-latest:generateContent"
        )

    def add_memory(self, chat_id: int, sender: str, text: str) -> None:
        if chat_id not in self.chat_memory:
            self.chat_memory[chat_id] = deque(maxlen=10)
        self.chat_memory[chat_id].append(f"{sender}: {text}")

    def get_memory_text(self, chat_id: int) -> str:
        if chat_id not in self.chat_memory or not self.chat_memory[chat_id]:
            return "Önceki sohbet kaydı yok."
        return "\n".join(self.chat_memory[chat_id])

    def _call_gemini_rest(
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
            "contents": [{"parts": parts}],
            "generationConfig": {
                "temperature": 0.25,
                "responseMimeType": "application/json",
            },
        }

        headers = {
            "Content-Type": "application/json",
            "X-goog-api-key": self.api_key,
        }

        resp = requests.post(
            self.endpoint, headers=headers, json=payload, timeout=45
        )
        if resp.status_code != 200:
            raise RuntimeError(f"HTTP {resp.status_code}: {resp.text}")

        data = resp.json()
        raw_text = (
            data["candidates"][0]["content"]["parts"][0]["text"].strip()
        )
        # Markdown JSON bloğu varsa temizle
        raw_text = re.sub(r"^```json\s*", "", raw_text)
        raw_text = re.sub(r"\s*```$", "", raw_text)
        return json.loads(raw_text)

    async def evaluate_and_respond(
        self,
        chat_id: int,
        user_name: str,
        user_text: str,
        is_group: bool,
        bot_mentioned: bool,
        image_bytes: Optional[bytes] = None,
    ) -> Dict[str, Any]:
        catalog_context = self.inventory.get_context_snapshot()
        recent_history = self.get_memory_text(chat_id)

        full_prompt = f"""
Sen Arslanlar Marble (Arslanlar Mermer) şirketinin kıdemli dijital satış ve operasyon uzmanısın.
Adın Argon. Grupta veya özel sohbette gerçek bir insan uzman gibi davranırsın. Komutlarla değil, doğal konuşmayla çalışırsın.

GÖREV VE DAVRANIŞ KURALLARIN:
1. SUSMA YETENEĞİ (GATEKEEPER):
   - Eğer mesaj bir grup sohbetinden geliyorsa (`is_group=True`) ve insanlar kendi aralarında geyik yapıyorsa, sadece "Bravo", "Teşekkürler", "Tamam", "😂😂😂", "Harika" gibi sohbet ifadeleri yazmışsa ve mermer/stok/lot/fiyat/görsel/nakliye ile ilgili hiçbir soru veya talep yoksa KESİNLİKLE ARAYA GİRME (`"should_reply": false` döndür).
   - Ancak mesajda bir taş adı, lot kodu (A-283 vb.), stok/metraj sorusu, konteyner hesabı, Drive linki talebi varsa VEYA kullanıcı bir mermer fotoğrafı göndermişse `"should_reply": true` döndür.

2. ÇOK DİLLİ UZMANLIK (TR / UK / EN):
   - Kullanıcı Türkçe yazdıysa Türkçe, Ukraynaca (Українська) yazdıysa akıcı ve profesyonel Ukraynaca, İngilizce yazdıysa İngilizce yanıt ver.
   - Ukraynaca yanıtlarda mermer terminolojisini doğru kullan (сляби, лот, квадратні метри, полірована поверхня, брекчія, наявність на складі).

3. VERİ DOĞRULUĞU VE KİŞİ FİLTRESİ:
   - Asla "SERKAN ÇELİK" gibi kişi/tedarikçi isimlerini mermer çeşidi olarak listeleme.
   - "SILVER ROOTS" ve "SİLVER ROOTS" tek bir taş olarak birleştirilmiştir, her zaman birleştirilmiş net veriyi kullan.
   - Durup dururken tüm stok listesini ezbere dökme! Kullanıcı hangi taşı veya konuyu sorduysa doğrudan ona odaklan. Sadece "elimizde en çok hangi taşlar var / genel stok nedir" diye sorulursa özet listeyi şık biçimde sun.

4. GÖRSEL DNA EŞLEŞTİRME (FOTOĞRAF GELDİĞİNDE):
   - Kullanıcı bir mermer fotoğrafı gönderdiyse: Önce taşın zemin rengini, damar yapısını (breccia/kırıklı, kök damarlı, çizgisel, bulutlu) ve damar rengini analiz et.
   - Ardından aşağıdaki ARSLANLAR MARBLE KATALOĞU içinden görsel olarak en çok benzeyen 1 ila 3 taşı/lotu seç.
   - `matched_queries` listesine bu eşleşen lot veya taş isimlerini yaz (Örn: `["A-283", "SILVER ROOTS"]`) ve `"send_photos": true` yap ki sistem Google Drive'dan o taşların gerçek plaka fotoğraflarını ve Drive alt klasör linklerini gruba albüm olarak atsın!

5. B2B KONTEYNER VE TONAJ HESABI:
   - Tonaj veya konteyner sorulursa: 2cm mermer plaka ≈ 54 kg/m², 3cm mermer plaka ≈ 81 kg/m². 1 adet 20 DC konteyner maksimum ~27.000 kg (27 ton) alır (2cm'den ~500 m², 3cm'den ~330 m²).

SON SOHBET GEÇMİŞİ (BAĞLAM):
{recent_history}

{catalog_context}

GELEN MESAJ BİLGİSİ:
- Grup Sohbeti mi: {is_group}
- Bot Etiketlendi mi: {bot_mentioned}
- Gönderen: {user_name}
- Mesaj: {user_text or '[Görsel Gönderildi]'}

ÇIKTI FORMATI (SADECE GEÇERLİ BİR JSON DÖNDÜR):
{{
  "should_reply": true,
  "language": "tr",
  "reply_text": "Kullanıcıya verilecek doğal, profesyonel ve nokta atışı yanıt...",
  "matched_queries": ["A-283"],
  "send_photos": true
}}
"""

        try:
            result = await asyncio.to_thread(
                self._call_gemini_rest, full_prompt, image_bytes
            )
            return result
        except Exception as e:
            logger.error(f"Gemini REST analiz hatası: {e}")
            return {
                "should_reply": False,
                "language": "tr",
                "reply_text": "",
                "matched_queries": [],
                "send_photos": False,
            }


# =====================================================================
# 6. TELEGRAM HANDLER VE MEDYA ALBÜM YÖNETİCİSİ
# =====================================================================
drive_explorer = DriveExplorer(GOOGLE_SA_JSON)
inventory_engine = InventoryEngine(EXCEL_FILE_PATH, drive_explorer)
argon_brain = ArgonBrain(GEMINI_API_KEY, inventory_engine, drive_explorer)

PROCESSED_MESSAGES: deque = deque(maxlen=500)


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
    user_name = (
        message.from_user.first_name if message.from_user else "Kullanıcı"
    )
    user_text = (message.text or message.caption or "").strip()

    link_teach_match = re.search(
        r"([A-Za-z0-9-_]+)\s+linki\s*:\s*(https?://\S+)", user_text, re.IGNORECASE
    )
    if link_teach_match:
        lot_key = normalize_lot_code(link_teach_match.group(1))
        url_val = link_teach_match.group(2)
        drive_explorer.save_custom_link(lot_key, url_val)
        await message.reply_text(
            f"✅ **{lot_key}** klasör bağlantısı hafızaya kaydedildi:\n{url_val}",
            parse_mode=ParseMode.MARKDOWN,
        )
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
            logger.error(f"Telegram fotoğraf indirme hatası: {e}")

    if not user_text and not image_bytes:
        return

    argon_brain.add_memory(
        chat_id, user_name, user_text if user_text else "[Mermer Fotoğrafı Gönderdi]"
    )

    decision = await argon_brain.evaluate_and_respond(
        chat_id=chat_id,
        user_name=user_name,
        user_text=user_text,
        is_group=is_group,
        bot_mentioned=bot_mentioned,
        image_bytes=image_bytes,
    )

    if not decision.get("should_reply", False):
        return

    reply_text = decision.get("reply_text", "").strip()
    matched_queries: List[str] = decision.get("matched_queries", [])
    send_photos: bool = decision.get("send_photos", False) or bool(image_bytes)

    if not reply_text:
        return

    drive_links_section: List[str] = []
    media_group_items: List[InputMediaPhoto] = []
    newly_cached_keys: List[Tuple[str, int]] = []
    last_folder_link: Optional[str] = None

    for query in matched_queries[:3]:
        norm_q = normalize_lot_code(query)
        folder_id, folder_link = drive_explorer.find_folder(norm_q)

        if folder_link:
            last_folder_link = folder_link
            drive_links_section.append(f"📂 {norm_q} Drive: {folder_link}")
        elif norm_q in inventory_engine.lots_index:
            drive_links_section.append(
                f"📂 {norm_q} Genel Drive: {DEFAULT_DRIVE_LINK}"
            )

        if send_photos:
            cached_file_ids = drive_explorer.visual_cache.get(norm_q, [])
            if cached_file_ids:
                for fid in cached_file_ids[:2]:
                    media_group_items.append(InputMediaPhoto(media=fid))
            elif folder_id:
                await context.bot.send_chat_action(
                    chat_id=chat_id, action=ChatAction.UPLOAD_PHOTO
                )
                raw_imgs = await asyncio.to_thread(
                    drive_explorer.get_folder_images, folder_id, 2
                )
                for img_b in raw_imgs:
                    media_group_items.append(InputMediaPhoto(media=img_b))
                    newly_cached_keys.append((norm_q, len(media_group_items) - 1))

    if drive_links_section:
        links_block = "\n".join(drive_links_section)
        if not last_folder_link or last_folder_link not in reply_text:
            reply_text = f"{reply_text}\n\n{links_block}"

    argon_brain.add_memory(chat_id, "Argon (Bot)", reply_text)

    if media_group_items:
        try:
            short_caption = (
                reply_text[:1020] + "..." if len(reply_text) > 1024 else reply_text
            )
            media_group_items[0] = InputMediaPhoto(
                media=media_group_items[0].media,
                caption=short_caption,
            )
            sent_messages = await message.reply_media_group(
                media=media_group_items[:10]
            )

            for norm_q, idx in newly_cached_keys:
                if idx < len(sent_messages) and sent_messages[idx].photo:
                    fid = sent_messages[idx].photo[-1].file_id
                    drive_explorer.visual_cache.setdefault(norm_q, [])
                    if fid not in drive_explorer.visual_cache[norm_q]:
                        drive_explorer.visual_cache[norm_q].append(fid)
            save_json_file(VISUAL_CACHE_FILE, drive_explorer.visual_cache)

            if len(reply_text) > 1024:
                await message.reply_text(reply_text)
            return
        except Exception as e:
            logger.error(f"Albüm gönderme hatası, metne dönülüyor: {e}")

    await message.reply_text(reply_text, disable_web_page_preview=False)


# =====================================================================
# 7. RENDER KEEP-ALIVE WEB SUNUCUSU VE BAŞLATICI
# =====================================================================
flask_app = Flask(__name__)


@flask_app.route("/")
def health_check():
    return "Arslanlar Marble Argon Master Bot is Alive and Running!", 200


def run_flask():
    port = int(os.environ.get("PORT", 10000))
    flask_app.run(host="0.0.0.0", port=port, use_reloader=False)


def main():
    if not TELEGRAM_TOKEN:
        raise ValueError("TELEGRAM_TOKEN ortam değişkeni eksik!")

    threading.Thread(target=run_flask, daemon=True).start()

    application = ApplicationBuilder().token(TELEGRAM_TOKEN).build()

    application.add_handler(
        MessageHandler(
            (filters.TEXT | filters.PHOTO) & (~filters.COMMAND),
            handle_incoming_message,
        )
    )

    logger.info("🚀 Arslanlar Marble Argon Master Bot (REST Mode) başlatılıyor...")
    application.run_polling(drop_pending_updates=True, allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
