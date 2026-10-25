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
from flask import Flask
from PIL import Image, ImageDraw

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
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# =====================================================================
# 1. LOGLAMA VE SİSTEM YAPILANDIRMASI
# =====================================================================
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("ArgonMaster_UltimateFinal")

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()  # HATA BURADAYDI, EKLENDİ!
GOOGLE_SA_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
DEFAULT_DRIVE_LINK = (os.getenv("GOOGLE_DRIVE_LINK") or "https://drive.google.com").strip()

HYPERLINKS_DB_FILE = "hyperlinks_db.json"

# =====================================================================
# 2. YARDIMCI FONKSİYONLAR
# =====================================================================
def load_json_file(filepath: str, default_data: Any) -> Any:
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f: return json.load(f)
        except: pass
    return default_data

def save_json_file(filepath: str, data: Any) -> None:
    try:
        with open(filepath, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=2)
    except: pass

def normalize_turkish_key(text: Any) -> str:
    if text is None or pd.isna(text): return ""
    tr_map = str.maketrans("İıiŞşĞğÜüÖöÇç", "IIISSGGUUOOCC")
    return re.sub(r"\s+", " ", str(text).strip().translate(tr_map).upper()).strip()

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
    elif "," in s: s = s.replace(",", ".")
    try: return float(s)
    except ValueError: return 0.0

# =====================================================================
# 3. GOOGLE DRIVE MOTORU (Otonom Klasör & Dosya Tarayıcı)
# =====================================================================
class DriveExplorer:
    def __init__(self, sa_json_str: str):
        self.service = None
        self.sa_json_str = sa_json_str
        self.hyperlinks_db = load_json_file(HYPERLINKS_DB_FILE, {})
        self._init_service()

    def _init_service(self) -> None:
        if not self.sa_json_str: return
        try:
            info = json.loads(self.sa_json_str)
            creds = service_account.Credentials.from_service_account_info(info, scopes=["https://www.googleapis.com/auth/drive.readonly"])
            self.service = build("drive", "v3", credentials=creds, cache_discovery=False)
            logger.info("Drive API Bağlantısı Aktif.")
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
        cands = [c for c in cands if len(c) >= 2]

        for cand in cands:
            try:
                q = f"mimeType = 'application/vnd.google-apps.folder' and trashed = false and name contains '{cand}'"
                res = self.service.files().list(q=q, spaces="drive", fields="files(id, name, webViewLink)", pageSize=3).execute()
                files = res.get("files", [])
                if files:
                    for f in files:
                        if normalize_turkish_key(f["name"]) == normalize_turkish_key(cand):
                            fid, link = f["id"], f.get("webViewLink", f"https://drive.google.com/drive/folders/{f['id']}")
                            self.hyperlinks_db[norm_key] = link
                            save_json_file(HYPERLINKS_DB_FILE, self.hyperlinks_db)
                            return fid, link
                    fid = files[0]["id"]
                    link = files[0].get("webViewLink", f"https://drive.google.com/drive/folders/{fid}")
                    self.hyperlinks_db[norm_key] = link
                    save_json_file(HYPERLINKS_DB_FILE, self.hyperlinks_db)
                    return fid, link
            except: continue
        return None, None

    def list_child_folders(self, parent_folder_id: str) -> List[Dict[str, str]]:
        if not self.service or not parent_folder_id: return []
        try:
            q = f"'{parent_folder_id}' in parents and trashed = false and mimeType = 'application/vnd.google-apps.folder'"
            res = self.service.files().list(q=q, spaces="drive", fields="files(id, name, webViewLink)", pageSize=100).execute()
            return [{"id": f["id"], "name": f["name"], "link": f.get("webViewLink")} for f in res.get("files", [])]
        except: return []

    def get_excel_from_folder(self, folder_id: str) -> Tuple[Optional[bytes], Optional[str]]:
        if not self.service or not folder_id: return None, None
        try:
            q = f"'{folder_id}' in parents and trashed = false and (mimeType contains 'spreadsheet' or mimeType contains 'excel' or name contains '.xls')"
            res = self.service.files().list(q=q, spaces="drive", fields="files(id, name)", pageSize=2).execute()
            files = res.get("files", [])
            
            if not files:
                for sf in self.list_child_folders(folder_id)[:5]:
                    sq = f"'{sf['id']}' in parents and trashed = false and (mimeType contains 'spreadsheet' or mimeType contains 'excel' or name contains '.xls')"
                    sres = self.service.files().list(q=sq, spaces="drive", fields="files(id, name)", pageSize=2).execute()
                    files.extend(sres.get("files", []))
                    if files: break

            if not files: return None, None
            return self._download_file(files[0]["id"]), files[0]["name"]
        except: return None, None

    def get_images_structured(self, folder_id: str, query_term: str) -> Tuple[Optional[bytes], Optional[bytes]]:
        """Telegram İçin: 1. Foto (Plaka), 2. Foto (Bookmatch veya 2. Plaka)"""
        all_imgs = self._scan_images(folder_id, query_term)
        if not all_imgs: return None, None

        bm_file, std_files = None, []
        for f in all_imgs:
            fn, pn = f.get("name", "").lower(), f.get("parent_name", "").lower()
            if "bookmatch" in fn or "bm" in fn or "bookmatch" in pn:
                if not bm_file: bm_file = f
            else:
                std_files.append(f)

        img1_bytes = self._download_file(std_files[0]["id"]) if std_files else (self._download_file(bm_file["id"]) if bm_file else None)
        img2_bytes = None
        if bm_file and bm_file != (std_files[0] if std_files else None):
            img2_bytes = self._download_file(bm_file["id"])
        elif len(std_files) > 1:
            img2_bytes = self._download_file(std_files[1]["id"])
        elif img1_bytes:
            img2_bytes = img1_bytes

        return img1_bytes, img2_bytes

    def get_best_slab_image(self, folder_id: str, query_term: str) -> Optional[bytes]:
        """PDF İçin: Bookmatch ASLA seçilmez. En iyi düz plaka seçilir."""
        all_imgs = self._scan_images(folder_id, query_term)
        if not all_imgs: return None
        clean_slabs = [f for f in all_imgs if "bookmatch" not in f.get("name", "").lower() and "bm" not in f.get("name", "").lower() and "bookmatch" not in f.get("parent_name", "").lower()]
        target = clean_slabs[0] if clean_slabs else all_imgs[0]
        return self._download_file(target["id"])

    def _scan_images(self, folder_id: str, query_term: str) -> List[Dict[str, str]]:
        if not self.service: return []
        all_imgs = []
        try:
            if folder_id:
                q = f"'{folder_id}' in parents and trashed = false and (mimeType contains 'image/jpeg' or mimeType contains 'image/png')"
                res = self.service.files().list(q=q, spaces="drive", fields="files(id, name)", pageSize=10).execute()
                all_imgs.extend(res.get("files", []))
                for sf in self.list_child_folders(folder_id)[:5]:
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
        if m: return m.group(1)
        m2 = re.search(r"id=([a-zA-Z0-9_-]+)", url)
        return m2.group(1) if m2 else None


# =====================================================================
# 4. HİBRİT ENVANTER MOTORU (Master Excel Fallback Sistemi)
# =====================================================================
class InventoryEngine:
    def __init__(self, drive: DriveExplorer):
        self.drive = drive
        self.master_lots_index: Dict[str, Dict[str, Any]] = {}
        self.master_raw_rows: List[Dict[str, Any]] = []
        self._load_master_excel()

    def _load_master_excel(self) -> None:
        """Klasörde Excel yoksa sıfır yazmasın diye ana stok listesini hafızaya alır."""
        target = None
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
            m2_col = next((c for c, n in col_map.items() if any(k in n for k in ["M2", "ALAN", "NET", "METRAJ"])), None)
            slab_col = next((c for c, n in col_map.items() if any(k in n for k in ["ADET", "PLAKA", "PCS"])), None)
            thick_col = next((c for c, n in col_map.items() if any(k in n for k in ["KALIN", "THICK", "CM"])), None)
            len_col = next((c for c, n in col_map.items() if any(k in n for k in ["BOY", "UZUNLUK"])), None)
            height_col = next((c for c, n in col_map.items() if any(k in n for k in ["EN", "YUKSEKLIK"])), None)
            surf_col = next((c for c, n in col_map.items() if any(k in n for k in ["YUZEY", "FINISH"])), None)
            lot_col = next((c for c, n in col_map.items() if any(k in n for k in ["FABRIKA KODU", "LOT", "PARTI", "BLOK NO", "BLOK"])), None)
            status_col = next((c for c, n in col_map.items() if any(k in n for k in ["STOK DURUMU", "DURUM"])), None)

            lots_agg = {}
            raw_rows = []

            for _, row in df.iterrows():
                lot_code = normalize_lot_code(row.get(lot_col))
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
                    lots_agg[lot_code] = {"stock_m2": 0.0, "stock_slabs": 0, "lengths": [], "heights": [], "thickness": str(row.get(thick_col, "2 cm")), "surface": str(row.get(surf_col, "Polished")).title()}
                
                lt = lots_agg[lot_code]
                if is_in_stock:
                    lt["stock_m2"] += m2_val
                    lt["stock_slabs"] += sc
                    lt["lengths"].append(safe_float(row.get(len_col)))
                    lt["heights"].append(safe_float(row.get(height_col)))

            for k, v in lots_agg.items():
                if v["lengths"] and v["heights"]:
                    v["dimensions"] = f"{int(sum(v['lengths'])/len(v['lengths']))} x {int(sum(v['heights'])/len(v['heights']))} cm"
                else:
                    v["dimensions"] = "Standard Slab Sizes"

            self.master_lots_index = lots_agg
            self.master_raw_rows = raw_rows
            logger.info(f"Master Excel Yedeği Aktif. Toplam Lot: {len(self.master_lots_index)}")
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
            slab_col = next((c for c, n in col_map.items() if any(k in n for k in ["ADET", "PLAKA", "PCS"])), None)
            thick_col = next((c for c, n in col_map.items() if any(k in n for k in ["KALIN", "THICK", "CM"])), None)
            len_col = next((c for c, n in col_map.items() if any(k in n for k in ["BOY", "UZUNLUK"])), None)
            height_col = next((c for c, n in col_map.items() if any(k in n for k in ["EN", "YUKSEKLIK"])), None)
            surf_col = next((c for c, n in col_map.items() if any(k in n for k in ["YUZEY", "FINISH"])), None)
            
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
        except: 
            return None

    def get_lot_info_hybrid(self, query: str) -> Dict[str, Any]:
        """Klasörde Excel yoksa sıfır vermez, Master Excel'den çeker."""
        norm_key = normalize_lot_code(query)
        folder_id, folder_link = self.drive.find_folder(norm_key)
        
        excel_bytes, excel_name, data = None, None, None

        if folder_id:
            excel_bytes, excel_name = self.drive.get_excel_from_folder(folder_id)
            if excel_bytes:
                data = self.parse_packing_list_bytes(excel_bytes)
        
        # Drive Klasöründe Excel Yoksa Ana (Master) Stok Listesine Bak
        if not data or data.get("total_m2", 0) == 0:
            if norm_key in self.master_lots_index:
                mst = self.master_lots_index[norm_key]
                data = {
                    "total_m2": mst["stock_m2"], "total_slabs": mst["stock_slabs"],
                    "thickness": mst["thickness"], "surface": mst["surface"],
                    "dimensions": mst["dimensions"], "raw_df": None, "mapped_cols": {}
                }
                matched_rows = [r for r in self.master_raw_rows if r["lot"] == norm_key and r["is_in_stock"]]
                if matched_rows:
                    data["raw_df"] = pd.DataFrame(matched_rows)
                    data["mapped_cols"] = {"m2": "m2", "len": "length", "height": "height", "slab": "slab_no", "thick": "thickness", "surf": "surface"}

        if not data:
            data = {"total_m2": 0.0, "total_slabs": 0, "thickness": "-", "surface": "-", "dimensions": "Belirsiz", "raw_df": None, "mapped_cols": {}}

        return {
            "status": "success",
            "query": norm_key,
            "drive_link": folder_link or DEFAULT_DRIVE_LINK,
            "folder_id": folder_id,
            "excel_name": excel_name or f"PackingList_{norm_key}.xlsx",
            "excel_bytes": excel_bytes,
            "data": data
        }

    async def deep_scan_collections(self) -> str:
        """Tüm Stoğumu Raporla komutu için ana klasörleri asenkron tarar."""
        collections = ["Arslanlar Collection", "Arslanlar Series", "Arslanlar Exclusive Collection"]
        report_lines = ["🔍 <b>ARGON MASTER - GENEL STOK RAPORU</b> 🔍\n"]
        total_factory_m2 = 0
        total_factory_slabs = 0
        
        for coll in collections:
            fid, _ = self.drive.find_folder(coll)
            if not fid: continue
            
            report_lines.append(f"📁 <b>{coll.upper()}</b>")
            sub_folders = self.drive.list_child_folders(fid)
            coll_m2 = 0
            
            for sf in sub_folders:
                ex_bytes, _ = self.drive.get_excel_from_folder(sf["id"])
                if ex_bytes:
                    data = self.parse_packing_list_bytes(ex_bytes)
                    if data and data["total_m2"] > 0:
                        report_lines.append(f"  ├ {sf['name']}: {data['total_m2']:,.2f} m² | {data['total_slabs']} Plaka")
                        coll_m2 += data["total_m2"]
                        total_factory_m2 += data["total_m2"]
                        total_factory_slabs += data["total_slabs"]
                await asyncio.sleep(0.05)
            
            report_lines.append(f"  └ <i>Koleksiyon Toplamı: {coll_m2:,.2f} m²</i>\n")
            
        report_lines.append(f"🌐 <b>FABRİKA GENEL TOPLAMI: {total_factory_m2:,.2f} m² | {total_factory_slabs} Plaka</b>")
        return "\n".join(report_lines)


# =====================================================================
# 5. KUSURSUZ LÜKS PDF MOTORU (Tek PDF'te Çoklu Taş & Mükemmel Tasarım)
# =====================================================================
class PDFPresentationEngine:
    def __init__(self, inventory: InventoryEngine, drive: DriveExplorer):
        self.inventory = inventory
        self.drive = drive

    @staticmethod
    def _create_luxury_logo() -> bytes:
        img = Image.new("RGB", (600, 100), color=(255, 255, 255))
        draw = ImageDraw.Draw(img)
        draw.text((20, 20), "ARSLANLAR MARBLE", fill=(20, 33, 61))
        draw.text((20, 60), "P R E M I U M   C O L L E C T I O N", fill=(197, 160, 89))
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()

    def generate_presentation_pdf(self, queries: List[str], collection_folder_name: str = None) -> Tuple[bytes, str]:
        """
        Sen A-96, A-383 yazdığında hepsini TEK BİR PDF İÇİNDE alt alta birleştirir.
        Koleksiyon yazarsan klasördeki tüm taşları tek bir PDF yapar.
        """
        cards = []
        pdf_title = "EXECUTIVE STOCK & SLAB PRESENTATION"

        if collection_folder_name:
            pdf_title = f"{normalize_turkish_key(collection_folder_name)} - CATALOG PRESENTATION"
            parent_id, _ = self.drive.find_folder(collection_folder_name)
            if parent_id:
                subfolders = self.drive.list_child_folders(parent_id)
                for sf in subfolders:
                    info = self.inventory.get_lot_info_hybrid(sf["name"])
                    img = self.drive.get_best_slab_image(info["folder_id"], sf["name"])
                    cards.append({"title": sf["name"].upper(), "data": info["data"], "drive_link": info["drive_link"], "image_bytes": img})
        else:
            for q in queries:
                info = self.inventory.get_lot_info_hybrid(q)
                img = self.drive.get_best_slab_image(info["folder_id"], q)
                cards.append({"title": q.upper(), "data": info["data"], "drive_link": info["drive_link"], "image_bytes": img})

        if not cards: raise ValueError("Belirtilen lotlar için veri bulunamadı.")

        pdf_io = io.BytesIO()
        doc = SimpleDocTemplate(pdf_io, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=40, bottomMargin=40)
        styles = getSampleStyleSheet()
        NAVY, GOLD, SLATE, LIGHT_BG, BORD = colors.HexColor("#14213D"), colors.HexColor("#C5A059"), colors.HexColor("#2B2D42"), colors.HexColor("#F8F9FA"), colors.HexColor("#DEE2E6")

        title_s = ParagraphStyle("T", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=26, textColor=NAVY, alignment=1, spaceAfter=8)
        sub_s = ParagraphStyle("S", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=11, textColor=GOLD, alignment=1, spaceAfter=20)
        ch_s = ParagraphStyle("CH", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=14, textColor=NAVY, spaceAfter=12)
        th_s = ParagraphStyle("TH", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9, textColor=colors.whitesmoke)
        tdl_s = ParagraphStyle("TDL", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9, textColor=SLATE)
        tdv_s = ParagraphStyle("TDV", parent=styles["Normal"], fontName="Helvetica", fontSize=9, textColor=SLATE)

        story = []
        try:
            logo_img = RLImage(io.BytesIO(self._create_luxury_logo()), width=240, height=48)
            header_table = Table([[logo_img]], colWidths=[500])
            header_table.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER')]))
            story = [header_table, Spacer(1, 10), Paragraph(pdf_title, sub_s)]
        except:
            story = [Paragraph("ARSLANLAR MARBLE", title_s), Paragraph(pdf_title, sub_s)]

        story.append(HRFlowable(width="100%", thickness=1.5, color=GOLD, spaceAfter=35))

        for idx, card in enumerate(cards, start=1):
            d = card["data"]
            m2 = d.get("total_m2", 0.0)
            slabs = d.get("total_slabs", 0)
            
            stock_str = f"{m2:,.2f} m² ({slabs} Slabs)" if m2 > 0 else "0.00 m² (Check Production)"
            surf_thick_str = f"{d.get('surface', '-')} | {d.get('thickness', '-')}"
            
            try:
                pil_im = Image.open(io.BytesIO(card["image_bytes"])).convert("RGB")
                img_io = io.BytesIO()
                pil_im.save(img_io, format="JPEG", quality=88)
                img_io.seek(0)
                rl_img = RLImage(img_io, width=220, height=140)
            except:
                rl_img = Paragraph("No Image Available", tdv_s)

            spec_rows = [
                [Paragraph("Specification", th_s), Paragraph("Technical Data", th_s)],
                [Paragraph("Available Stock", tdl_s), Paragraph(stock_str, tdv_s)],
                [Paragraph("Surface & Thickness", tdl_s), Paragraph(surf_thick_str, tdv_s)],
                [Paragraph("Active Code", tdl_s), Paragraph(card["title"], tdv_s)],
                [Paragraph("High-Res Photos", tdl_s), Paragraph(f"<link href='{card['drive_link']}' color='#0056b3'><u>Open Google Drive Folder</u></link>", tdv_s)],
            ]

            st = Table(spec_rows, colWidths=[110, 180])
            st.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), NAVY), 
                ("BACKGROUND", (0,1), (-1,-1), LIGHT_BG), 
                ("GRID", (0,0), (-1,-1), 0.5, BORD), 
                ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                ("TOPPADDING", (0,0), (-1,-1), 7),
                ("BOTTOMPADDING", (0,0), (-1,-1), 7),
                ("LEFTPADDING", (0,0), (-1,-1), 6),
            ]))
            
            mr = Table([[rl_img, st]], colWidths=[230, 300])
            mr.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("ALIGN", (0,0), (0,0), "CENTER")]))
            
            story.append(KeepTogether([Paragraph(f"{idx}. {card['title']}", ch_s), Spacer(1, 5), mr]))
            story.append(Spacer(1, 35))

        doc.build(story)
        pdf_io.seek(0)
        fname = "Collection_Presentation.pdf" if collection_folder_name else f"Presentation_{'_'.join(queries[:2])}.pdf"
        return pdf_io.read(), fname

    def generate_packing_list_pdf(self, raw_df: pd.DataFrame, lot_name: str, mapped_cols: Dict[str, str]) -> bytes:
        pdf_io = io.BytesIO()
        doc = SimpleDocTemplate(pdf_io, pagesize=landscape(A4), rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
        styles = getSampleStyleSheet()
        NAVY, GOLD, SLATE = colors.HexColor("#14213D"), colors.HexColor("#C5A059"), colors.HexColor("#2B2D42")
        
        title_p = ParagraphStyle("PlT", fontName="Helvetica-Bold", fontSize=18, textColor=NAVY, alignment=1)
        sub_p = ParagraphStyle("PlS", fontName="Helvetica-Bold", fontSize=10, textColor=GOLD, alignment=1, spaceAfter=15)
        th_p = ParagraphStyle("TH", fontName="Helvetica-Bold", fontSize=9, textColor=colors.whitesmoke, alignment=1)
        td_p = ParagraphStyle("TD", fontName="Helvetica", fontSize=8.5, textColor=SLATE, alignment=1)

        story = [Paragraph("ARSLANLAR MARBLE", title_p), Paragraph(f"OFFICIAL PACKING LIST — {lot_name.upper()}", sub_p), HRFlowable(width="100%", thickness=1, color=GOLD, spaceAfter=10)]
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
        story.append(Paragraph(f"<b>Total Evaluated Slabs:</b> {row_count} pcs | <b>Total Net Area:</b> {total_m2:,.2f} m²", ParagraphStyle("Sum", fontName="Helvetica-Oblique", fontSize=9, textColor=SLATE)))
        doc.build(story)
        pdf_io.seek(0)
        return pdf_io.read()

# =====================================================================
# 6. ARGON AI NLP MOTORU (İnsan Gibi Akıllı Konuşma + Komut Yönlendirme)
# =====================================================================
class ArgonBrain:
    def __init__(self, api_key: str):
        self.api_key = api_key
        self.endpoint = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent"

    async def evaluate_intent(self, user_text: str, is_group: bool, bot_mentioned: bool) -> Dict[str, Any]:
        raw = (user_text or "").strip().upper()
        if not raw: return {"should_reply": False}
        
        lot_regex = re.compile(r"\b([A-Z]{1,4})[-\s]?(\d{1,5})\b")
        matched = [f"{m.group(1)}-{m.group(2)}" for m in lot_regex.finditer(raw) if m.group(1) not in ["M","CM","MM","PDF","TON","KG"]]
        
        want_deep_scan = any(kw in raw for kw in ["TÜM STOĞUMU", "TUM STOGUM", "GENEL TARAMA", "DEEP SCAN", "RAPORLA", "NE KADAR STOĞUM", "NE KADAR STOGUM", "TOPLAM STOK", "TOPLAM NE KADAR"])
        
        collection_folder = None
        if "COLLECTION" in raw or "SERIES" in raw or "EXCLUSIVE" in raw:
            if "ARSLANLAR COLLECTION" in raw: collection_folder = "Arslanlar Collection"
            elif "ARSLANLAR SERIES" in raw: collection_folder = "Arslanlar Series"
            elif "EXCLUSIVE" in raw: collection_folder = "Arslanlar Exclusive Collection"
            elif "TÜM" in raw or "TUM" in raw: collection_folder = "Arslanlar Collection"
        
        want_pdf = any(kw in raw for kw in ["PDF", "SUNUM", "KATALOG"])
        want_packing = any(kw in raw for kw in ["ÇEKİ", "CEKI", "PACKING", "SEVK"])
        want_photo = any(kw in raw for kw in ["FOTO", "RESIM", "GORSEL", "DRIVE", "LINK"])
        is_greeting = any(kw in raw for kw in ["MERHABA", "SELAM", "ORDA MISIN", "HELLO", "NASILSIN"])

        # Akıllı AI Yanıtı (Operasyon Dışı Sorular İçin İnsan Gibi Cevap)
        reply_text = ""
        has_command = matched or want_deep_scan or want_pdf or want_packing or want_photo or collection_folder
        
        if self.api_key and not has_command and (not is_group or bot_mentioned or is_greeting):
            try:
                payload = {
                    "contents": [{"parts": [{"text": f"Sen Arslanlar Marble dijital asistanı Argonsun. Kullanıcı sana şunu yazdı: '{user_text}'. Çok kısa, profesyonel, akıllı ve yardımsever bir cevap ver. Konu dışına çıkma."}]}],
                    "generationConfig": {"temperature": 0.3}
                }
                headers = {"Content-Type": "application/json", "X-goog-api-key": self.api_key}
                resp = requests.post(self.endpoint, json=payload, headers=headers, timeout=5)
                if resp.status_code == 200:
                    reply_text = resp.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
            except Exception as e:
                logger.error(f"Gemini AI Hatası: {e}")
        
        if is_group and not bot_mentioned and not has_command and not is_greeting:
            return {"should_reply": False}

        return {
            "should_reply": True,
            "reply_text": reply_text,
            "matched_queries": matched[:5],
            "collection_folder": collection_folder,
            "want_deep_scan": want_deep_scan,
            "want_pdf": want_pdf,
            "want_packing": want_packing,
            "want_photo": want_photo or bool(matched)
        }

# =====================================================================
# 7. TELEGRAM ANA SİSTEM (HTML MOTORU)
# =====================================================================
drive_explorer = DriveExplorer(GOOGLE_SA_JSON)
inventory_engine = InventoryEngine(drive_explorer)
pdf_engine = PDFPresentationEngine(inventory_engine, drive_explorer)
argon_brain = ArgonBrain(GEMINI_API_KEY)

PROCESSED_MESSAGES = deque(maxlen=500)

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

    # 1. DEEP SCAN (TOPLAM STOK)
    if intent.get("want_deep_scan") and not intent.get("matched_queries") and not intent.get("collection_folder"):
        status = await msg.reply_text("🔎 <b>Argon Deep Scan Başlatıldı.</b>\nTüm Drive klasörlerinizdeki stoklar taranıyor...", parse_mode=ParseMode.HTML)
        try:
            report = await asyncio.to_thread(inventory_engine.deep_scan_collections)
            await status.edit_text(report, parse_mode=ParseMode.HTML)
        except Exception as e:
            await status.edit_text(f"Tarama Hatası: {e}")
        return

    queries = intent.get("matched_queries", [])
    collection_folder = intent.get("collection_folder")
    
    # EĞER KOMUT YOKSA SADECE AI MESAJINI GÖNDER (İnsan Gibi Yanıt)
    if not queries and not collection_folder and not intent.get("want_packing") and not intent.get("want_pdf"):
        ai_reply = intent.get("reply_text")
        if ai_reply:
            await msg.reply_text(ai_reply, parse_mode=ParseMode.MARKDOWN)
        elif not is_group or bot_mentioned:
            await msg.reply_text("👋 Argon Master Final Pure Drive System Aktif.\nLot kodu (A-371, A-96), Koleksiyon Sunumu veya çeki listesi isteyebilirsiniz.")
        return

    await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.TYPING)

    # =========================================================
    # MULTI-LOT (A-96, A-383) VE KOLEKSİYON İÇİN TEK PDF MOTORU
    # =========================================================
    if intent.get("want_pdf"):
        if collection_folder:
            await msg.reply_text(f"📁 <b>{collection_folder}</b> koleksiyonu taranıp tek bir Premium PDF'te birleştiriliyor...", parse_mode=ParseMode.HTML)
        else:
            await msg.reply_text(f"📄 Seçilen Lotlar (<b>{', '.join(queries)}</b>) tek bir Premium PDF'te birleştiriliyor...", parse_mode=ParseMode.HTML)
        
        await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.UPLOAD_DOCUMENT)
        try:
            pdf_b, pdf_n = await asyncio.to_thread(pdf_engine.generate_presentation_pdf, queries, collection_folder_name=collection_folder)
            await msg.reply_document(document=io.BytesIO(pdf_b), filename=pdf_n, caption="📄 Lüks İngilizce Sunum (Bookmatch'siz)", parse_mode=ParseMode.HTML)
        except Exception as e:
            await msg.reply_text(f"PDF oluşturulurken hata: {e}")
        return

    # =========================================================
    # NORMAL SORGULAMA VE ÇEKİ LİSTESİ DÖNGÜSÜ
    # =========================================================
    for q in queries:
        try:
            info = await asyncio.to_thread(inventory_engine.get_lot_info_hybrid, q)
            d = info["data"]
            fid = info["folder_id"]
            flink = info["drive_link"]

            if intent.get("want_packing"):
                await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.UPLOAD_DOCUMENT)
                if info["excel_bytes"]:
                    await msg.reply_document(document=io.BytesIO(info["excel_bytes"]), filename=info["excel_name"], caption=f"📦 <b>{q}</b> Orijinal Excel Çeki Listesi", parse_mode=ParseMode.HTML)
                elif d.get("raw_df") is not None:
                    excel_io = io.BytesIO()
                    d["raw_df"].to_excel(excel_io, index=False)
                    excel_io.seek(0)
                    await msg.reply_document(document=excel_io, filename=f"PackingList_{q}.xlsx", caption=f"📦 <b>{q}</b> Master Excel'den Üretilmiş Çeki Listesi", parse_mode=ParseMode.HTML)
                else:
                    await msg.reply_text(f"⚠️ {q} için veritabanında çeki listesi (Excel) bulunamadı.")

                if d.get("raw_df") is not None:
                    pdf_pack_bytes = await asyncio.to_thread(pdf_engine.generate_packing_list_pdf, d["raw_df"], q, d["mapped_cols"])
                    await msg.reply_document(document=io.BytesIO(pdf_pack_bytes), filename=f"PackingList_{q}.pdf", caption=f"📄 <b>{q}</b> Resmi PDF Çeki Listesi", parse_mode=ParseMode.HTML)
                continue

            # TELEGRAM FLASH-KART VE İKİLİ FOTO ALBÜM
            text_response = (
                f"💎 <b>{q.upper()}</b>\n"
                f"• Mevcut Stok: <code>{d['total_m2']:,.2f} m²</code> (<code>{d['total_slabs']} Plaka</code>)\n"
                f"• Ortalama Ebat: <code>{d['dimensions']}</code>\n"
                f"• Yüzey & Kalınlık: <code>{d['surface']} ({d['thickness']})</code>\n"
                f"• Drive Linki: <a href='{flink}'>Klasöre Git</a>"
            )
            
            img1, img2 = await asyncio.to_thread(drive_explorer.get_images_structured, fid, q)
            
            media = []
            if img1: media.append(InputMediaPhoto(media=img1, caption=text_response, parse_mode=ParseMode.HTML))
            if img2: media.append(InputMediaPhoto(media=img2))
            
            if media:
                await msg.reply_media_group(media=media)
            else:
                await msg.reply_text(text_response, parse_mode=ParseMode.HTML)
        except Exception as ex:
            logger.error(f"İşlem Hatası ({q}): {ex}")
            await msg.reply_text(f"⚠️ {q} işlenirken beklenmeyen bir hata oluştu.")


# =====================================================================
# 8. RENDER SIFIR-BEKLEME BAŞLATICI
# =====================================================================
flask_app = Flask(__name__)

@flask_app.route("/")
def health_check(): return "Argon Master Ultimate Final is Live!", 200

def run_flask():
    flask_app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)), use_reloader=False)

def main():
    if not TELEGRAM_TOKEN:
        logger.error("CRITICAL: TELEGRAM_TOKEN eksik!")
        return

    threading.Thread(target=run_flask, daemon=True).start()
    application = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    
    application.add_handler(MessageHandler((filters.TEXT | filters.PHOTO) & (~filters.COMMAND), handle_message))
    
    logger.info("🚀 Argon Master Ultimate Final başlatılıyor...")
    application.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
