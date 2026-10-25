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

import requests
import pandas as pd
import openpyxl
from flask import Flask
from PIL import Image, ImageDraw, ImageFont

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

from telegram import Update, InputMediaPhoto, InputMediaDocument
from telegram.constants import ParseMode, ChatAction
from telegram.ext import (
    ApplicationBuilder,
    ContextTypes,
    MessageHandler,
    filters,
)

from google.oauth2 import service_account
import google.auth.transport.requests
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# =====================================================================
# 1. LOGLAMA VE SİSTEM YAPILANDIRMASI
# =====================================================================
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("ArgonMasterHybridPerfect")

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GOOGLE_SA_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()

DEFAULT_DRIVE_LINK = (
    os.getenv("GOOGLE_DRIVE_LINK")
    or os.getenv("DEFAULT_DRIVE_LINK")
    or "https://drive.google.com"
).strip()

EXCEL_FILE_PATH = os.getenv("EXCEL_FILE_PATH", "stok.xlsx").strip()

HYPERLINKS_DB_FILE = "hyperlinks_db.json"
VISUAL_CACHE_FILE = "visual_cache.json"

# =====================================================================
# 2. YARDIMCI FONKSİYONLAR
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
    if text is None or pd.isna(text): return ""
    s = str(text).strip()
    tr_map = str.maketrans("İıiŞşĞğÜüÖöÇç", "IIISSGGUUOOCC")
    return re.sub(r"\s+", " ", s.translate(tr_map).upper()).strip()

def normalize_lot_code(code: Any) -> str:
    raw = normalize_turkish_key(code)
    match = re.match(r"^([A-Z]+)[-\s]*(\d+)$", raw)
    return f"{match.group(1)}-{match.group(2)}" if match else raw

def safe_float(val: Any) -> float:
    if pd.isna(val) or val is None: return 0.0
    if isinstance(val, (int, float)): return float(val)
    s = re.sub(r"[^\d,.-]", "", str(val).strip())
    if not s: return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try: return float(s)
    except ValueError: return 0.0

# =====================================================================
# 3. GOOGLE DRIVE MOTORU (İkili Fotoğraf & Bookmatch Zekası)
# =====================================================================
class DriveExplorer:
    def __init__(self, sa_json_str: str):
        self.service = None
        self.sa_json_str = sa_json_str
        self.hyperlinks_db = load_json_file(HYPERLINKS_DB_FILE, {})
        self.visual_cache = load_json_file(VISUAL_CACHE_FILE, {})
        self._init_service()

    def _init_service(self) -> None:
        if not self.sa_json_str: return
        try:
            info = json.loads(self.sa_json_str)
            creds = service_account.Credentials.from_service_account_info(info, scopes=["https://www.googleapis.com/auth/drive.readonly"])
            self.service = build("drive", "v3", credentials=creds, cache_discovery=False)
        except Exception as e:
            logger.error(f"Drive API Hatası: {e}")

    def find_folder(self, query_term: str) -> Tuple[Optional[str], Optional[str]]:
        norm_key = normalize_lot_code(query_term)
        if norm_key in self.hyperlinks_db:
            url = self.hyperlinks_db[norm_key]
            fid = self._extract_folder_id(url)
            if fid: return fid, url

        if not self.service: return None, None

        cands = list(dict.fromkeys([
            query_term.strip().replace("'", ""), norm_key, 
            norm_key.replace("-", ""), norm_key.replace("-", " "),
            norm_key.replace("SILVER", "SİLVER").replace("RIVER", "RİVER")
        ]))
        cands = [c for c in cands if len(c) > 2]

        for cand in cands:
            try:
                q = f"mimeType = 'application/vnd.google-apps.folder' and trashed = false and name contains '{cand}'"
                res = self.service.files().list(q=q, spaces="drive", fields="files(id, name, webViewLink)", pageSize=5).execute()
                files = res.get("files", [])
                if files:
                    fid, link = files[0]["id"], files[0].get("webViewLink", f"https://drive.google.com/drive/folders/{files[0]['id']}")
                    self.hyperlinks_db[norm_key] = link
                    save_json_file(HYPERLINKS_DB_FILE, self.hyperlinks_db)
                    return fid, link
            except: continue
        return None, None

    def list_child_folders(self, parent_folder_id: str) -> List[Dict[str, str]]:
        if not self.service or not parent_folder_id: return []
        try:
            q = f"'{parent_folder_id}' in parents and trashed = false and mimeType = 'application/vnd.google-apps.folder'"
            res = self.service.files().list(q=q, spaces="drive", fields="files(id, name)", pageSize=50).execute()
            return [{"id": f["id"], "name": f["name"]} for f in res.get("files", [])]
        except: return []

    def get_excel_from_folder(self, folder_id: str) -> Tuple[Optional[bytes], Optional[str]]:
        if not self.service or not folder_id: return None, None
        try:
            q = f"'{folder_id}' in parents and trashed = false and (mimeType contains 'spreadsheet' or mimeType contains 'excel' or name contains '.xls')"
            res = self.service.files().list(q=q, spaces="drive", fields="files(id, name)", pageSize=2).execute()
            files = res.get("files", [])
            
            if not files:
                subs = self.list_child_folders(folder_id)[:5]
                for sf in subs:
                    sq = f"'{sf['id']}' in parents and trashed = false and (mimeType contains 'spreadsheet' or mimeType contains 'excel' or name contains '.xls')"
                    sres = self.service.files().list(q=sq, spaces="drive", fields="files(id, name)", pageSize=2).execute()
                    files.extend(sres.get("files", []))
                    if files: break

            if not files: return None, None
            req = self.service.files().get_media(fileId=files[0]["id"])
            fh = io.BytesIO()
            dl = MediaIoBaseDownload(fh, req)
            done = False
            while not done: _, done = dl.next_chunk()
            fh.seek(0)
            return fh.read(), files[0]["name"]
        except: return None, None

    def get_images_structured(self, folder_id: str, query_term: str) -> Tuple[Optional[bytes], Optional[bytes]]:
        """1. Foto: Plaka, 2. Foto: Bookmatch (veya 2. plaka)"""
        all_imgs = self._scan_images(folder_id, query_term)
        if not all_imgs: return None, None

        bm_file = None
        std_files = []

        for f in all_imgs:
            fn, pn = f.get("name", "").lower(), f.get("parent_name", "").lower()
            if "bookmatch" in fn or "bm" in fn or "bookmatch" in pn:
                if not bm_file: bm_file = f
            else:
                std_files.append(f)

        img1 = self._download_file(std_files[0]["id"]) if std_files else (self._download_file(bm_file["id"]) if bm_file else None)
        img2 = None
        if bm_file and bm_file != (std_files[0] if std_files else None):
            img2 = self._download_file(bm_file["id"])
        elif len(std_files) > 1:
            img2 = self._download_file(std_files[1]["id"])
        elif img1:
            img2 = img1
            
        return img1, img2

    def get_best_slab_image(self, folder_id: str, query_term: str) -> Optional[bytes]:
        """PDF İçin: Sadece düz plaka, Bookmatch ASLA seçilmez."""
        all_imgs = self._scan_images(folder_id, query_term)
        if not all_imgs: return None

        clean = [f for f in all_imgs if "bookmatch" not in f.get("name","").lower() and "bm" not in f.get("name","").lower() and "bookmatch" not in f.get("parent_name","").lower()]
        target = clean[0] if clean else all_imgs[0]
        return self._download_file(target["id"])

    def _scan_images(self, folder_id: str, query_term: str) -> List[Dict[str, str]]:
        if not self.service: return []
        all_imgs = []
        try:
            if folder_id:
                q = f"'{folder_id}' in parents and trashed = false and (mimeType contains 'image/jpeg' or mimeType contains 'image/png')"
                res = self.service.files().list(q=q, spaces="drive", fields="files(id, name)", pageSize=15).execute()
                all_imgs.extend(res.get("files", []))
                
                for sf in self.list_child_folders(folder_id)[:4]:
                    sq = f"'{sf['id']}' in parents and trashed = false and (mimeType contains 'image/jpeg' or mimeType contains 'image/png')"
                    sres = self.service.files().list(q=sq, spaces="drive", fields="files(id, name)", pageSize=5).execute()
                    for img in sres.get("files", []):
                        img["parent_name"] = sf["name"]
                        all_imgs.append(img)
            else:
                cands = [normalize_lot_code(query_term), query_term]
                for cand in cands[:2]:
                    q = f"trashed = false and name contains '{cand}' and (mimeType contains 'image/jpeg' or mimeType contains 'image/png')"
                    res = self.service.files().list(q=q, spaces="drive", fields="files(id, name)", pageSize=5).execute()
                    all_imgs.extend(res.get("files", []))
            return all_imgs
        except: return []

    def _download_file(self, file_id: str) -> Optional[bytes]:
        try:
            req = self.service.files().get_media(fileId=file_id)
            fh = io.BytesIO()
            dl = MediaIoBaseDownload(fh, req)
            done = False
            while not done: _, done = dl.next_chunk()
            fh.seek(0)
            return fh.read()
        except: return None

    @staticmethod
    def _extract_folder_id(url: str) -> Optional[str]:
        if not url: return None
        m = re.search(r"/folders/([a-zA-Z0-9_-]+)", url)
        return m.group(1) if m else (re.search(r"id=([a-zA-Z0-9_-]+)", url).group(1) if re.search(r"id=([a-zA-Z0-9_-]+)", url) else None)


# =====================================================================
# 4. HİBRİT ENVANTER MOTORU (DRIVE EXCEL + MASTER STOK.XLSX FALLBACK)
# =====================================================================
class InventoryEngine:
    def __init__(self, drive: DriveExplorer):
        self.drive = drive
        self.master_lots_index: Dict[str, Dict[str, Any]] = {}
        self.master_raw_rows: List[Dict[str, Any]] = []
        self._load_master_excel()

    def _load_master_excel(self) -> None:
        """Sisteme yüklenen ana stok.xlsx dosyasını arka planda okur ve yedekte tutar."""
        target = EXCEL_FILE_PATH if os.path.exists(EXCEL_FILE_PATH) else None
        if not target:
            for fname in os.listdir("."):
                if fname.lower().endswith((".xlsx", ".xls")) and not fname.startswith("~$"):
                    target = fname
                    break
        if not target: return

        try:
            xls = pd.ExcelFile(target)
            df = pd.DataFrame()
            for sheet in xls.sheet_names:
                temp_df = pd.read_excel(xls, sheet_name=sheet)
                if not temp_df.empty and len(temp_df.columns) > 3:
                    df = temp_df
                    break

            if df.empty: return
            col_map = {c: normalize_turkish_key(c) for c in df.columns}
            
            m2_col = next((c for c, n in col_map.items() if "M2" in n or "ALAN" in n or "NET" in n), None)
            slab_col = next((c for c, n in col_map.items() if "ADET" in n or "PLAKA" in n or "PCS" in n), None)
            thick_col = next((c for c, n in col_map.items() if "KALINLIK" in n or "THICKNESS" in n or "CM" in n), None)
            len_col = next((c for c, n in col_map.items() if "BOY" in n or "UZUNLUK" in n or "LENG" in n), None)
            height_col = next((c for c, n in col_map.items() if "EN" in n or "YUKSEKLIK" in n or "HEIG" in n), None)
            surf_col = next((c for c, n in col_map.items() if "YUZEY" in n or "FINISH" in n), None)
            lot_col = next((c for c, n in col_map.items() if "FABRIKA KODU" in n or "LOT" in n or "PARTI" in n), None)
            blok_col = next((c for c, n in col_map.items() if "BLOK NO" in n or "BLOK" in n), None)
            status_col = next((c for c, n in col_map.items() if "STOK DURUMU" in n or "DURUM" in n), None)

            lots_agg = {}
            raw_rows = []

            for _, row in df.iterrows():
                raw_fab = row.get(lot_col) if lot_col else None
                raw_blok = row.get(blok_col) if blok_col else None
                lot_code = normalize_lot_code(raw_fab) or normalize_lot_code(raw_blok)
                if not lot_code or lot_code in ["NAN", "NONE"]: continue

                m2_val = safe_float(row.get(m2_col))
                if m2_val <= 0: continue

                sc = int(safe_float(row.get(slab_col))) if slab_col else 1
                if sc <= 0: sc = 1

                raw_stat = normalize_turkish_key(row.get(status_col)) if status_col else "STOK"
                is_in_stock = (raw_stat == "STOK") or ("SHOWROOM" in raw_stat) or (not status_col)

                raw_rows.append({
                    "lot": lot_code, "m2": m2_val, "slab_no": sc,
                    "length": safe_float(row.get(len_col)), "height": safe_float(row.get(height_col)),
                    "thickness": row.get(thick_col, "2 cm"), "surface": row.get(surf_col, "Polished"),
                    "is_in_stock": is_in_stock
                })

                if lot_code not in lots_agg:
                    lots_agg[lot_code] = {
                        "stock_m2": 0.0, "stock_slabs": 0,
                        "lengths": [], "heights": [],
                        "thickness": row.get(thick_col, "2 cm"),
                        "surface": str(row.get(surf_col, "Polished")).title()
                    }
                
                lt = lots_agg[lot_code]
                if is_in_stock:
                    lt["stock_m2"] += m2_val
                    lt["stock_slabs"] += sc
                    lt["lengths"].append(safe_float(row.get(len_col)))
                    lt["heights"].append(safe_float(row.get(height_col)))

            for k, v in lots_agg.items():
                if v["lengths"] and v["heights"]:
                    avg_l = int(sum(v["lengths"]) / len(v["lengths"]))
                    avg_h = int(sum(v["heights"]) / len(v["heights"]))
                    v["dimensions"] = f"{avg_l} x {avg_h} cm"
                else:
                    v["dimensions"] = "Standard Slab Sizes"

            self.master_lots_index = lots_agg
            self.master_raw_rows = raw_rows
            logger.info(f"Master Excel Yüklendi: {len(self.master_lots_index)} aktif lot.")
        except Exception as e:
            logger.error(f"Master Excel okuma hatası: {e}")

    def parse_packing_list_bytes(self, excel_bytes: bytes) -> Dict[str, Any]:
        try:
            xls = pd.ExcelFile(io.BytesIO(excel_bytes))
            df = pd.DataFrame()
            for sheet in xls.sheet_names:
                temp_df = pd.read_excel(xls, sheet_name=sheet)
                if not temp_df.empty and len(temp_df.columns) > 3:
                    df = temp_df
                    break
            
            if df.empty: raise ValueError("Boş Excel")
            col_map = {c: normalize_turkish_key(c) for c in df.columns}
            
            m2_col = next((c for c, n in col_map.items() if any(k in n for k in ["M2", "ALAN", "NET", "METRAJ"])), None)
            slab_col = next((c for c, n in col_map.items() if any(k in n for k in ["ADET", "PLAKA", "PCS", "MIKTAR", "SLAB"])), None)
            thick_col = next((c for c, n in col_map.items() if any(k in n for k in ["KALIN", "THICK", "CM"])), None)
            len_col = next((c for c, n in col_map.items() if any(k in n for k in ["BOY", "UZUNLUK", "LENG"])), None)
            height_col = next((c for c, n in col_map.items() if any(k in n for k in ["EN", "YUKSEKLIK", "HEIG", "WID"])), None)
            surf_col = next((c for c, n in col_map.items() if any(k in n for k in ["YUZEY", "FINISH", "SURF"])), None)
            
            total_m2 = df[m2_col].apply(safe_float).sum() if m2_col else 0.0
            total_slabs = int(df[slab_col].apply(safe_float).sum()) if slab_col else len(df.dropna(how='all'))
            thick_val = df[thick_col].dropna().iloc[0] if thick_col and not df[thick_col].dropna().empty else 2
            surf_val = df[surf_col].dropna().iloc[0] if surf_col and not df[surf_col].dropna().empty else "Polished"
            
            avg_l, avg_h = 0, 0
            if len_col and height_col:
                valid_l = df[len_col].apply(safe_float).replace(0.0, pd.NA).dropna()
                valid_h = df[height_col].apply(safe_float).replace(0.0, pd.NA).dropna()
                if not valid_l.empty and not valid_h.empty:
                    avg_l, avg_h = int(valid_l.mean()), int(valid_h.mean())
                
            return {
                "total_m2": total_m2, "total_slabs": total_slabs,
                "thickness": f"{thick_val} cm" if str(thick_val).isdigit() else str(thick_val),
                "surface": str(surf_val).title(),
                "dimensions": f"{avg_l} x {avg_h} cm" if avg_l > 0 else "Standart Ebat",
                "raw_df": df,
                "mapped_cols": {"m2": m2_col, "len": len_col, "height": height_col, "slab": slab_col, "thick": thick_col, "surf": surf_col}
            }
        except: return {"total_m2": 0.0, "total_slabs": 0, "thickness": "2 cm", "surface": "Polished", "dimensions": "-", "raw_df": None, "mapped_cols": {}}

    def get_lot_info_hybrid(self, query: str) -> Dict[str, Any]:
        """HIBIT MOTOR: Önce Drive klasörüne bakar. Excel yoksa MASTER Stok.xlsx'ten veriyi çeker! Asla hata vermez."""
        norm_key = normalize_lot_code(query)
        folder_id, folder_link = self.drive.find_folder(norm_key)
        
        excel_bytes, excel_name = None, None
        data = None

        # 1. Drive'dan Excel bulmayı dene
        if folder_id:
            excel_bytes, excel_name = self.drive.get_excel_from_folder(folder_id)
            if excel_bytes:
                data = self.parse_packing_list_bytes(excel_bytes)
        
        # 2. Drive'da Excel Yoksa -> Master Stok.xlsx Fallback (Yedek) Mekanizması
        if not data or data["total_m2"] == 0:
            if norm_key in self.master_lots_index:
                mst = self.master_lots_index[norm_key]
                data = {
                    "total_m2": mst["stock_m2"],
                    "total_slabs": mst["stock_slabs"],
                    "thickness": mst["thickness"],
                    "surface": mst["surface"],
                    "dimensions": mst["dimensions"],
                    "raw_df": None, "mapped_cols": {}
                }
                # Master Excel'deki satırları toplayıp sanal bir raw_df oluştur
                matched_rows = [r for r in self.master_raw_rows if r["lot"] == norm_key and r["is_in_stock"]]
                if matched_rows:
                    data["raw_df"] = pd.DataFrame(matched_rows)
                    data["mapped_cols"] = {"m2": "m2", "len": "length", "height": "height", "slab": "slab_no", "thick": "thickness", "surf": "surface"}

        if not data or data["total_m2"] == 0:
            return {
                "status": "error", 
                "message": "Stok verisi (Drive'da veya Master Excel'de) bulunamadı.", 
                "drive_link": folder_link, 
                "folder_id": folder_id
            }

        return {
            "status": "success",
            "query": norm_key,
            "drive_link": folder_link or DEFAULT_DRIVE_LINK,
            "folder_id": folder_id,
            "excel_name": excel_name or f"PackingList_{norm_key}.xlsx",
            "excel_bytes": excel_bytes,
            "data": data
        }


# =====================================================================
# 5. LÜKS PDF ÜRETİM MOTORU (SIFIR KAYMA, LOGOLU)
# =====================================================================
class PDFPresentationEngine:
    def __init__(self, inventory: InventoryEngine, drive: DriveExplorer):
        self.inventory = inventory
        self.drive = drive

    @staticmethod
    def _create_luxury_logo() -> bytes:
        img = Image.new("RGB", (400, 100), color=(255, 255, 255))
        draw = ImageDraw.Draw(img)
        draw.text((20, 30), "ARSLANLAR MARBLE", fill=(20, 33, 61))
        draw.text((20, 60), "P R E M I U M   C O L L E C T I O N", fill=(197, 160, 89))
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()

    def generate_presentation_pdf(self, queries: List[str]) -> Tuple[bytes, str]:
        """Boşlukları ayarlayarak başlık kaymasını önleyen, şık logolu Lüks PDF."""
        cards = []
        for q in queries:
            info = self.inventory.get_lot_info_hybrid(q)
            if info["status"] != "success": continue
            img1 = self.drive.get_best_slab_image(info["folder_id"], q)
            cards.append({
                "title": q.upper(),
                "data": info["data"],
                "drive_link": info["drive_link"],
                "image_bytes": img1
            })

        if not cards: raise ValueError("Belirtilen lotlar için veri bulunamadı.")

        pdf_io = io.BytesIO()
        doc = SimpleDocTemplate(pdf_io, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=40, bottomMargin=40)
        styles = getSampleStyleSheet()
        NAVY, GOLD, SLATE, LIGHT_BG, BORD = colors.HexColor("#14213D"), colors.HexColor("#C5A059"), colors.HexColor("#2B2D42"), colors.HexColor("#F8F9FA"), colors.HexColor("#DEE2E6")

        title_s = ParagraphStyle("T", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=24, textColor=NAVY, alignment=1, spaceAfter=6)
        sub_s = ParagraphStyle("S", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=10, textColor=GOLD, alignment=1, spaceAfter=20) # Spacer eklendi
        ch_s = ParagraphStyle("CH", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=13, textColor=NAVY, spaceAfter=8) # Spacer eklendi
        th_s = ParagraphStyle("TH", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9, textColor=colors.whitesmoke)
        tdl_s = ParagraphStyle("TDL", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9, textColor=SLATE)
        tdv_s = ParagraphStyle("TDV", parent=styles["Normal"], fontName="Helvetica", fontSize=9, textColor=SLATE)

        # Üst Kısım Tasarımı (Logolu)
        try:
            logo_img = RLImage(io.BytesIO(self._create_luxury_logo()), width=200, height=50)
            header_table = Table([[logo_img]], colWidths=[500])
            header_table.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER')]))
            story = [header_table, Spacer(1, 15)]
        except:
            story = [Paragraph("ARSLANLAR MARBLE", title_s), Paragraph("EXECUTIVE STOCK & SLAB PRESENTATION", sub_s)]

        story.append(HRFlowable(width="100%", thickness=1.5, color=GOLD, spaceAfter=25))

        for idx, card in enumerate(cards, start=1):
            d = card["data"]
            m2 = d["total_m2"]
            slabs = d["total_slabs"]
            
            # Available Stock karşısında metraj doluluğu sağlandı
            stock_str = f"{m2:,.2f} m² ({slabs} Slabs)" if m2 > 0 else "Contact for Availability"
            surf_thick_str = f"{d['surface']} | {d['thickness']}"
            
            raw_img = card["image_bytes"]
            try:
                pil_im = Image.open(io.BytesIO(raw_img)).convert("RGB")
                img_io = io.BytesIO()
                pil_im.save(img_io, format="JPEG", quality=88)
                img_io.seek(0)
                rl_img = RLImage(img_io, width=190, height=130)
            except:
                rl_img = Paragraph("No Image Available", tdv_s)

            # Temizlenmiş Tablo (Total Production ve boş Avg Slab Size yok)
            spec_rows = [
                [Paragraph("Specification", th_s), Paragraph("Technical Data", th_s)],
                [Paragraph("Available Stock", tdl_s), Paragraph(stock_str, tdv_s)],
                [Paragraph("Surface & Thickness", tdl_s), Paragraph(surf_thick_str, tdv_s)],
                [Paragraph("Active Code", tdl_s), Paragraph(card["title"], tdv_s)],
                [Paragraph("High-Res Photos", tdl_s), Paragraph(f"<link href='{card['drive_link']}' color='#0056b3'><u>Open Google Drive Folder</u></link>", tdv_s)],
            ]

            st = Table(spec_rows, colWidths=[130, 220])
            st.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), NAVY), 
                ("BACKGROUND", (0,1), (-1,-1), LIGHT_BG), 
                ("GRID", (0,0), (-1,-1), 0.5, BORD), 
                ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                ("TOPPADDING", (0,0), (-1,-1), 6),
                ("BOTTOMPADDING", (0,0), (-1,-1), 6),
            ]))
            
            mr = Table([[rl_img, st]], colWidths=[200, 360])
            mr.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("ALIGN", (0,0), (0,0), "CENTER")]))
            
            story.append(KeepTogether([Paragraph(f"{idx}. {card['title']}", ch_s), Spacer(1, 5), mr, Spacer(1, 25)]))

        doc.build(story)
        pdf_io.seek(0)
        return pdf_io.read(), f"Arslanlar_Presentation_{datetime.now().strftime('%Y%m%d')}.pdf"

    def generate_packing_list_pdf(self, raw_df: pd.DataFrame, lot_name: str, mapped_cols: Dict[str, str]) -> bytes:
        pdf_io = io.BytesIO()
        doc = SimpleDocTemplate(pdf_io, pagesize=landscape(A4), rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
        styles = getSampleStyleSheet()
        NAVY, GOLD, SLATE = colors.HexColor("#14213D"), colors.HexColor("#C5A059"), colors.HexColor("#2B2D42")
        
        title_p = ParagraphStyle("PlT", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=18, textColor=NAVY, alignment=1)
        sub_p = ParagraphStyle("PlS", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=10, textColor=GOLD, alignment=1, spaceAfter=15)
        th_p = ParagraphStyle("TH", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9, textColor=colors.whitesmoke, alignment=1)
        td_p = ParagraphStyle("TD", parent=styles["Normal"], fontName="Helvetica", fontSize=8.5, textColor=SLATE, alignment=1)

        story = [
            Paragraph("ARSLANLAR MARBLE", title_p),
            Paragraph(f"OFFICIAL PACKING LIST — {lot_name.upper()}", sub_p),
            HRFlowable(width="100%", thickness=1, color=GOLD, spaceAfter=10),
        ]

        table_data = [[Paragraph("Row #", th_p), Paragraph("Slab No", th_p), Paragraph("Length", th_p), Paragraph("Height", th_p), Paragraph("Thick", th_p), Paragraph("Net M2", th_p)]]
        
        total_m2, row_count = 0.0, 0
        c_len, c_h, c_m2, c_slab, c_th = mapped_cols.get("len"), mapped_cols.get("height"), mapped_cols.get("m2"), mapped_cols.get("slab"), mapped_cols.get("thick")

        for idx, row in raw_df.iterrows():
            if pd.isna(row.get(c_m2)) or safe_float(row.get(c_m2)) <= 0: continue
            row_count += 1
            m2_val = safe_float(row.get(c_m2))
            total_m2 += m2_val
            table_data.append([
                Paragraph(str(row_count), td_p), Paragraph(str(row.get(c_slab, "-")), td_p),
                Paragraph(str(row.get(c_len, "-")), td_p), Paragraph(str(row.get(c_h, "-")), td_p),
                Paragraph(str(row.get(c_th, "-")), td_p), Paragraph(f"{m2_val:.2f}", td_p)
            ])
            if row_count >= 150: break

        pl_table = Table(table_data, colWidths=[60, 100, 100, 100, 80, 100])
        pl_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), NAVY), ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F8F9FA")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#DEE2E6")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")
        ]))
        story.append(pl_table)
        story.append(Spacer(1, 15))
        story.append(Paragraph(f"<b>Total Slabs:</b> {row_count} pcs | <b>Total Net Area:</b> {total_m2:,.2f} m²", ParagraphStyle("Sum", parent=styles["Normal"], fontName="Helvetica-Oblique", fontSize=9, textColor=SLATE)))

        doc.build(story)
        pdf_io.seek(0)
        return pdf_io.read()

# =====================================================================
# 6. ARGON AI ORKESTRATÖRÜ
# =====================================================================
class ArgonBrain:
    def __init__(self, api_key: str, sa_json_str: str):
        self.api_key = api_key
        self.sa_json_str = sa_json_str

    async def evaluate_intent(self, user_text: str, is_group: bool, bot_mentioned: bool) -> Dict[str, Any]:
        raw = (user_text or "").strip().upper()
        if not raw: return {"should_reply": False}
        
        lot_regex = re.compile(r"\b([A-Z]{1,4})[-\s]?(\d{1,5})\b")
        matched = [f"{m.group(1)}-{m.group(2)}" for m in lot_regex.finditer(raw) if m.group(1) not in ["M","CM","MM","PDF"]]
        
        want_pdf = any(kw in raw for kw in ["PDF", "SUNUM", "KATALOG"])
        want_packing = any(kw in raw for kw in ["ÇEKİ", "CEKI", "PACKING", "SEVK"])
        want_photo = any(kw in raw for kw in ["FOTO", "RESIM", "GORSEL"])
        
        if is_group and not bot_mentioned and not matched and not want_pdf and not want_packing and not want_photo:
            return {"should_reply": False}

        return {
            "should_reply": True,
            "matched_queries": matched[:3],
            "want_pdf": want_pdf,
            "want_packing": want_packing,
            "want_photo": want_photo or bool(matched)
        }

# =====================================================================
# 7. TELEGRAM SOHBET YÖNETİCİSİ
# =====================================================================
drive_explorer = DriveExplorer(GOOGLE_SA_JSON)
inventory_engine = InventoryEngine(drive_explorer)
pdf_engine = PDFPresentationEngine(inventory_engine, drive_explorer)
argon_brain = ArgonBrain(GEMINI_API_KEY, GOOGLE_SA_JSON)

PROCESSED_MESSAGES = deque(maxlen=500)

async def handle_document_upload(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.effective_message
    if not msg or not msg.document: return
    if not msg.document.file_name.lower().endswith((".xlsx", ".xls")): return
    
    status = await msg.reply_text("📥 Master Excel alınıyor...", parse_mode=ParseMode.MARKDOWN)
    try:
        tg_file = await context.bot.get_file(msg.document.file_id)
        await tg_file.download_to_drive(EXCEL_FILE_PATH)
        await asyncio.to_thread(inventory_engine._load_master_excel)
        await status.edit_text("✅ Master Excel Yüklendi. Argon Backup Sistemi Aktif.")
    except Exception as e:
        await status.edit_text(f"❌ Hata: {e}")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.effective_message
    if not msg or not update.effective_chat: return
    
    msg_key = f"{update.effective_chat.id}:{msg.message_id}"
    if msg_key in PROCESSED_MESSAGES: return
    PROCESSED_MESSAGES.append(msg_key)

    chat_id = update.effective_chat.id
    is_group = update.effective_chat.type in ["group", "supergroup"]
    bot_mentioned = bool(context.bot.username and f"@{context.bot.username.lower()}" in (msg.text or "").lower())
    
    intent = await argon_brain.evaluate_intent(msg.text or "", is_group, bot_mentioned)
    if not intent.get("should_reply"): return

    queries = intent.get("matched_queries", [])
    if not queries:
        if not is_group or bot_mentioned:
            await msg.reply_text("👋 Argon Master Hybrid System Aktif.\nLot kodu (A-371), çeki listesi veya PDF isteyebilirsiniz.")
        return

    await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.TYPING)
    
    for q in queries:
        info = await asyncio.to_thread(inventory_engine.get_lot_info_hybrid, q)
        if info["status"] != "success":
            await msg.reply_text(f"⚠️ **{q}** Stok verisi bulunamadı.", parse_mode=ParseMode.MARKDOWN)
            continue

        d = info["data"]
        fid = info["folder_id"]
        flink = info["drive_link"]

        # ÇEKİ LİSTESİ İSTENDİYSE (DUAL EXPORT)
        if intent.get("want_packing"):
            await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.UPLOAD_DOCUMENT)
            if info["excel_bytes"]:
                await msg.reply_document(document=io.BytesIO(info["excel_bytes"]), filename=info["excel_name"], caption=f"📦 **{q}** Orijinal Excel Çeki Listesi", parse_mode=ParseMode.MARKDOWN)
            else:
                # Master Excel'den sanal excel üret
                excel_io = io.BytesIO()
                d["raw_df"].to_excel(excel_io, index=False)
                excel_io.seek(0)
                await msg.reply_document(document=excel_io, filename=f"PackingList_{q}.xlsx", caption=f"📦 **{q}** Master Excel'den Üretilmiş Çeki Listesi", parse_mode=ParseMode.MARKDOWN)

            if d.get("raw_df") is not None:
                pdf_pack_bytes = await asyncio.to_thread(pdf_engine.generate_packing_list_pdf, d["raw_df"], q, d["mapped_cols"])
                await msg.reply_document(document=io.BytesIO(pdf_pack_bytes), filename=f"PackingList_{q}.pdf", caption=f"📄 **{q}** Resmi PDF Çeki Listesi", parse_mode=ParseMode.MARKDOWN)
            continue
            
        # PDF SUNUM İSTENDİYSE
        if intent.get("want_pdf"):
            await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.UPLOAD_DOCUMENT)
            try:
                pdf_b, pdf_n = await asyncio.to_thread(pdf_engine.generate_presentation_pdf, [q])
                await msg.reply_document(document=io.BytesIO(pdf_b), filename=pdf_n, caption=f"📄 **{q}** Premium PDF Sunumu", parse_mode=ParseMode.MARKDOWN)
            except Exception as e:
                await msg.reply_text(f"PDF oluşturulurken hata: {e}")
            continue

        # NORMAL SORGULAMA (İKİLİ FOTO ALBÜM + BİLGİ KARTI)
        text_response = (
            f"💎 **{q.upper()}**\n"
            f"• Mevcut Stok: `{d['total_m2']:,.2f} m²` (`{d['total_slabs']} Plaka`)\n"
            f"• Ortalama Ebat: `{d['dimensions']}`\n"
            f"• Yüzey & Kalınlık: `{d['surface']} ({d['thickness']})`\n"
            f"• Drive Linki: {flink}"
        )
        
        img1, img2 = await asyncio.to_thread(drive_explorer.get_images_structured, fid, q)
        media = []
        if img1: media.append(InputMediaPhoto(media=img1, caption=text_response, parse_mode=ParseMode.MARKDOWN))
        if img2: media.append(InputMediaPhoto(media=img2))
        
        if media: await msg.reply_media_group(media=media)
        else: await msg.reply_text(text_response, parse_mode=ParseMode.MARKDOWN)

# =====================================================================
# 8. RENDER SIFIR-BEKLEME BAŞLATICI
# =====================================================================
flask_app = Flask(__name__)

@flask_app.route("/")
def health_check(): return "Argon Master Hybrid Perfect v4 is Live!", 200

def run_flask():
    flask_app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)), use_reloader=False)

def main():
    threading.Thread(target=run_flask, daemon=True).start()
    application = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    
    application.add_handler(MessageHandler(filters.Document.ALL, handle_document_upload))
    application.add_handler(MessageHandler((filters.TEXT | filters.PHOTO) & (~filters.COMMAND), handle_message))
    
    logger.info("🚀 Argon Master Hybrid Perfect v4 başlatılıyor...")
    application.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
