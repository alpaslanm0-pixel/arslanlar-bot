import os
import io
import re
import json
import logging
import asyncio
import threading
from datetime import datetime
from collections import deque
from typing import Dict, List, Optional, Tuple, Any

import pandas as pd
from flask import Flask
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

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# =====================================================================
# 1. LOGLAMA VE AYARLAR
# =====================================================================
logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger("ArgonMaster_PureDrive")

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "").strip()
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
    if pd.isna(text) or text is None: return ""
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
    except: return 0.0

# =====================================================================
# 3. PURE DRIVE MOTORU (Sadece Klasör İçine Odaklanır)
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
            logger.error(f"Drive Hatası: {e}")

    def find_folder(self, query_term: str) -> Tuple[Optional[str], Optional[str]]:
        norm_key = normalize_lot_code(query_term)
        if norm_key in self.hyperlinks_db:
            url = self.hyperlinks_db[norm_key]
            fid = self._extract_folder_id(url)
            if fid: return fid, url

        if not self.service: return None, None
        
        # Olası isimleri daha geniş tut
        cands = list(dict.fromkeys([
            norm_key, query_term.strip().upper(),
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
            res = self.service.files().list(q=q, spaces="drive", fields="files(id, name, webViewLink)", pageSize=100).execute()
            return [{"id": f["id"], "name": f["name"], "link": f.get("webViewLink")} for f in res.get("files", [])]
        except: return []

    def get_excel_from_folder(self, folder_id: str) -> Tuple[Optional[bytes], Optional[str]]:
        """SADECE belirtilen klasörün içindeki Excel dosyasını bulur."""
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
        """Telegram için: 1. Düz Plaka, 2. Bookmatch (yoksa 2. Plaka)"""
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
        """PDF için: Asla Bookmatch almaz. En iyi düz plakayı alır."""
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
                res = self.service.files().list(q=q, spaces="drive", fields="files(id, name)", pageSize=15).execute()
                all_imgs.extend(res.get("files", []))
                for sf in self.list_child_folders(folder_id)[:5]:
                    sq = f"'{sf['id']}' in parents and trashed = false and (mimeType contains 'image/jpeg' or mimeType contains 'image/png')"
                    sres = self.service.files().list(q=sq, spaces="drive", fields="files(id, name)", pageSize=5).execute()
                    for img in sres.get("files", []):
                        img["parent_name"] = sf["name"]
                        all_imgs.append(img)
            else:
                cands = [normalize_lot_code(query_term), query_term.upper()]
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
# 4. ENVANTER MOTORU (Sıfır Master Excel - Sadece Klasör İçi)
# =====================================================================
class InventoryEngine:
    def __init__(self, drive: DriveExplorer):
        self.drive = drive

    def parse_packing_list_bytes(self, excel_bytes: bytes) -> Dict[str, Any]:
        """Klasördeki Excel dosyasını okuyup toplamları döner."""
        try:
            xls = pd.ExcelFile(io.BytesIO(excel_bytes))
            df = pd.DataFrame()
            for sheet in xls.sheet_names:
                temp_df = pd.read_excel(xls, sheet_name=sheet)
                if not temp_df.empty and len(temp_df.columns) > 3:
                    df = temp_df
                    break
            if df.empty: raise ValueError("Boş Excel")
            
            col_map = {c: str(normalize_turkish_key(c)) for c in df.columns}
            m2_col = next((c for c, n in col_map.items() if any(k in n for k in ["M2", "ALAN", "NET", "METRAJ"])), None)
            slab_col = next((c for c, n in col_map.items() if any(k in n for k in ["ADET", "PLAKA", "PCS"])), None)
            thick_col = next((c for c, n in col_map.items() if any(k in n for k in ["KALIN", "THICK", "CM"])), None)
            len_col = next((c for c, n in col_map.items() if any(k in n for k in ["BOY", "UZUNLUK"])), None)
            height_col = next((c for c, n in col_map.items() if any(k in n for k in ["EN", "YUKSEKLIK"])), None)
            surf_col = next((c for c, n in col_map.items() if any(k in n for k in ["YUZEY", "FINISH"])), None)
            
            total_m2 = df[m2_col].apply(safe_float).sum() if m2_col else 0.0
            total_slabs = int(df[slab_col].apply(safe_float).sum()) if slab_col else len(df.dropna(how='all'))
            thick_val = df[thick_col].dropna().iloc[0] if thick_col and not df[thick_col].dropna().empty else "2"
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
                "dimensions": f"{avg_l} x {avg_h} cm" if avg_l > 0 else "-",
                "raw_df": df, "mapped_cols": {"m2": m2_col, "len": len_col, "height": height_col, "slab": slab_col, "thick": thick_col, "surf": surf_col}
            }
        except: return None

    def get_lot_info(self, query: str) -> Dict[str, Any]:
        """Lot klasörüne girer. Excel bulursa okur, bulamazsa hata vermeden sıfır değerlerle döner."""
        norm_key = normalize_lot_code(query)
        folder_id, folder_link = self.drive.find_folder(norm_key)
        
        excel_bytes, excel_name, data = None, None, None

        if folder_id:
            excel_bytes, excel_name = self.drive.get_excel_from_folder(folder_id)
            if excel_bytes:
                data = self.parse_packing_list_bytes(excel_bytes)
        
        # Eğer klasörde Excel yoksa (Sıfır hata prensibi)
        if not data:
            data = {"total_m2": 0.0, "total_slabs": 0, "thickness": "-", "surface": "-", "dimensions": "-", "raw_df": None, "mapped_cols": {}}

        return {
            "status": "success",
            "query": norm_key,
            "drive_link": folder_link or DEFAULT_DRIVE_LINK,
            "folder_id": folder_id,
            "excel_name": excel_name,
            "excel_bytes": excel_bytes,
            "data": data
        }

    def deep_scan_collections(self) -> str:
        """Belirtilen 3 ana koleksiyon klasörüne girer, tüm lotların Excel'ini okuyup raporlar."""
        collections = ["Arslanlar Collection", "Arslanlar Series", "Arslanlar Exclusive Collection"]
        report_lines = ["🔍 <b>ARGON MASTER - DEEP SCAN EXECUTIVE REPORT</b> 🔍\n"]
        total_factory_m2, total_factory_slabs = 0.0, 0
        
        for coll in collections:
            fid, _ = self.drive.find_folder(coll)
            if not fid: continue
            
            report_lines.append(f"📁 <b>{coll.upper()}</b>")
            sub_folders = self.drive.list_child_folders(fid)
            coll_m2 = 0.0
            
            for sf in sub_folders:
                ex_bytes, _ = self.drive.get_excel_from_folder(sf["id"])
                if ex_bytes:
                    data = self.parse_packing_list_bytes(ex_bytes)
                    if data and data["total_m2"] > 0:
                        report_lines.append(f"  ├ {sf['name']}: {data['total_m2']:,.2f} m² | {data['total_slabs']} Plaka")
                        coll_m2 += data["total_m2"]
                        total_factory_m2 += data["total_m2"]
                        total_factory_slabs += data["total_slabs"]
            
            report_lines.append(f"  └ <i>Koleksiyon Toplamı: {coll_m2:,.2f} m²</i>\n")
            
        report_lines.append(f"🌐 <b>FABRİKA GENEL TOPLAMI: {total_factory_m2:,.2f} m² | {total_factory_slabs} Plaka</b>")
        return "\n".join(report_lines)

# =====================================================================
# 5. KUSURSUZ LÜKS PDF MOTORU (Native ReportLab Tasarımı)
# =====================================================================
class PDFPresentationEngine:
    def __init__(self, inventory: InventoryEngine, drive: DriveExplorer):
        self.inventory = inventory
        self.drive = drive

    def generate_presentation_pdf(self, queries: List[str], collection_folder_name: str = None) -> Tuple[bytes, str]:
        """
        Görsel kalitesizliğini bitiren, tamamen ReportLab Native Metinleriyle çizilmiş 
        mükemmel tasarımlı Lüks PDF Sunumu. Tüm lotları tek PDF'te birleştirir.
        """
        cards = []
        pdf_title = "EXECUTIVE STOCK & SLAB PRESENTATION"

        if collection_folder_name:
            pdf_title = f"{normalize_turkish_key(collection_folder_name)} - CATALOG PRESENTATION"
            parent_id, _ = self.drive.find_folder(collection_folder_name)
            if parent_id:
                subfolders = self.drive.list_child_folders(parent_id)
                for sf in subfolders:
                    info = self.inventory.get_lot_info(sf["name"])
                    img = self.drive.get_best_slab_image(sf["id"], sf["name"])
                    cards.append({"title": sf["name"].upper(), "data": info["data"], "drive_link": info["drive_link"], "image_bytes": img})
        else:
            for q in queries:
                info = self.inventory.get_lot_info(q)
                img = self.drive.get_best_slab_image(info["folder_id"], q)
                cards.append({"title": q.upper(), "data": info["data"], "drive_link": info["drive_link"], "image_bytes": img})

        if not cards: raise ValueError("Belirtilen lotlar için veri bulunamadı.")

        pdf_io = io.BytesIO()
        doc = SimpleDocTemplate(pdf_io, pagesize=A4, rightMargin=35, leftMargin=35, topMargin=45, bottomMargin=40)
        styles = getSampleStyleSheet()
        NAVY, GOLD, SLATE, LIGHT_BG, BORD = colors.HexColor("#14213D"), colors.HexColor("#C5A059"), colors.HexColor("#2B2D42"), colors.HexColor("#F8F9FA"), colors.HexColor("#DEE2E6")

        # Native ReportLab Font Stilleri (Premium Görünüm)
        header_title = ParagraphStyle("HT", fontName="Helvetica-Bold", fontSize=28, textColor=NAVY, alignment=1, spaceAfter=2)
        header_sub = ParagraphStyle("HS", fontName="Helvetica-Bold", fontSize=11, textColor=GOLD, alignment=1, spaceAfter=25)
        pdf_title_style = ParagraphStyle("PTS", fontName="Helvetica-Bold", fontSize=16, textColor=NAVY, alignment=1, spaceAfter=5)
        
        ch_s = ParagraphStyle("CH", fontName="Helvetica-Bold", fontSize=15, textColor=NAVY, spaceAfter=15)
        th_s = ParagraphStyle("TH", fontName="Helvetica-Bold", fontSize=9.5, textColor=colors.whitesmoke)
        tdl_s = ParagraphStyle("TDL", fontName="Helvetica-Bold", fontSize=9.5, textColor=SLATE)
        tdv_s = ParagraphStyle("TDV", fontName="Helvetica", fontSize=9.5, textColor=SLATE)

        # Üst Başlık (Yapay PIL resmi silindi, gerçek şık metin yazılıyor)
        story = [
            Paragraph("ARSLANLAR MARBLE", header_title),
            Paragraph("P R E M I U M   C O L L E C T I O N", header_sub),
            Paragraph(pdf_title, pdf_title_style),
            HRFlowable(width="100%", thickness=1.5, color=GOLD, spaceAfter=35)
        ]

        for idx, card in enumerate(cards, start=1):
            d = card["data"]
            m2 = d.get("total_m2", 0.0)
            slabs = d.get("total_slabs", 0)
            
            # Klasörde Excel yoksa şık bir uyarı verir, 0.00 yazıp bırakmaz
            if m2 > 0:
                stock_str = f"{m2:,.2f} m² ({slabs} Slabs)"
            else:
                stock_str = "Excel Tablosu Bekleniyor"
                
            surf_thick_str = f"{d.get('surface', '-')} | {d.get('thickness', '-')}"
            
            try:
                # Fotoğrafın bozulmasını engellemek için yüksek kalite
                pil_im = Image.open(io.BytesIO(card["image_bytes"])).convert("RGB")
                img_io = io.BytesIO()
                pil_im.save(img_io, format="JPEG", quality=95)
                img_io.seek(0)
                rl_img = RLImage(img_io, width=220, height=145)
            except:
                rl_img = Paragraph("No Image Available", tdv_s)

            spec_rows = [
                [Paragraph("Specification", th_s), Paragraph("Technical Data", th_s)],
                [Paragraph("Available Stock", tdl_s), Paragraph(stock_str, tdv_s)],
                [Paragraph("Surface & Thickness", tdl_s), Paragraph(surf_thick_str, tdv_s)],
                [Paragraph("Active Code", tdl_s), Paragraph(card["title"], tdv_s)],
                [Paragraph("High-Res Photos", tdl_s), Paragraph(f"<link href='{card['drive_link']}' color='#0056b3'><u>Open Google Drive Folder</u></link>", tdv_s)],
            ]

            st = Table(spec_rows, colWidths=[125, 205])
            st.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), NAVY), ("BACKGROUND", (0,1), (-1,-1), LIGHT_BG), 
                ("GRID", (0,0), (-1,-1), 0.5, BORD), ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                ("TOPPADDING", (0,0), (-1,-1), 8), ("BOTTOMPADDING", (0,0), (-1,-1), 8),
                ("LEFTPADDING", (0,0), (-1,-1), 8),
            ]))
            
            mr = Table([[rl_img, st]], colWidths=[235, 335])
            mr.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("ALIGN", (0,0), (0,0), "CENTER")]))
            
            story.append(KeepTogether([Paragraph(f"{idx}. {card['title']}", ch_s), mr]))
            story.append(Spacer(1, 40))

        doc.build(story)
        pdf_io.seek(0)
        fname = "Collection_Presentation.pdf" if collection_folder_name else f"Presentation_{'_'.join(queries[:3])}.pdf"
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
# 6. ARGON NLP (Sabit, Hatasız Niyet Okuyucu)
# =====================================================================
class ArgonBrain:
    def evaluate_intent(self, user_text: str, is_group: bool, bot_mentioned: bool) -> Dict[str, Any]:
        raw = (user_text or "").strip().upper()
        if not raw: return {"should_reply": False}
        
        # Sadece Lot Kodlarını Çıkar
        lot_regex = re.compile(r"\b([A-Z]{1,4})[-\s]?(\d{1,5})\b")
        matched = [f"{m.group(1)}-{m.group(2)}" for m in lot_regex.finditer(raw) if m.group(1) not in ["M","CM","MM","PDF","TON","KG"]]
        
        # Deep Scan Tespiti (Toplam Stok Komutları)
        deep_scan_keywords = ["TÜM STOK", "TUM STOK", "DEEP SCAN", "RAPORLA", "NE KADAR STOĞUM", "NE KADAR STOGUM", "TOPLAM STOK", "TOPLAM NE KADAR", "TOPLAM STOĞUM", "TOPLAM STOGUM"]
        want_deep_scan = any(kw in raw for kw in deep_scan_keywords)
        
        collection_folder = None
        if "COLLECTION" in raw or "SERIES" in raw or "EXCLUSIVE" in raw:
            if "ARSLANLAR COLLECTION" in raw: collection_folder = "Arslanlar Collection"
            elif "ARSLANLAR SERIES" in raw: collection_folder = "Arslanlar Series"
            elif "EXCLUSIVE" in raw: collection_folder = "Arslanlar Exclusive Collection"
            elif "TÜM" in raw or "TUM" in raw: collection_folder = "Arslanlar Collection"
        
        want_pdf = any(kw in raw for kw in ["PDF", "SUNUM", "KATALOG"])
        want_packing = any(kw in raw for kw in ["ÇEKİ", "CEKI", "PACKING", "SEVK"])
        want_photo = any(kw in raw for kw in ["FOTO", "RESIM", "GORSEL", "DRIVE", "LINK"])
        
        if is_group and not bot_mentioned and not matched and not want_pdf and not want_packing and not want_photo and not collection_folder and not want_deep_scan:
            return {"should_reply": False}

        return {
            "should_reply": True,
            "matched_queries": list(dict.fromkeys(matched))[:10], # 10 lota kadar tek PDF'te birleştirir
            "collection_folder": collection_folder,
            "want_deep_scan": want_deep_scan,
            "want_pdf": want_pdf,
            "want_packing": want_packing,
            "want_photo": want_photo or bool(matched)
        }

# =====================================================================
# 7. TELEGRAM SOHBET YÖNETİCİSİ (TRY-EXCEPT ZIRHLI)
# =====================================================================
drive_explorer = DriveExplorer(GOOGLE_SA_JSON)
inventory_engine = InventoryEngine(drive_explorer)
pdf_engine = PDFPresentationEngine(inventory_engine, drive_explorer)
argon_brain = ArgonBrain()

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
    
    intent = argon_brain.evaluate_intent(msg.text or "", is_group, bot_mentioned)
    if not intent.get("should_reply"): return

    # 1. BÜYÜK VERİ TARAMASI (DEEP SCAN)
    if intent.get("want_deep_scan") and not intent.get("matched_queries") and not intent.get("collection_folder"):
        status = await msg.reply_text("🔎 <b>Argon Deep Scan Başlatıldı.</b>\nKlasörlerinizdeki Excel tabloları taranıyor...", parse_mode=ParseMode.HTML)
        try:
            report = await asyncio.to_thread(inventory_engine.deep_scan_collections)
            await status.edit_text(report, parse_mode=ParseMode.HTML)
        except Exception as e:
            await status.edit_text(f"Tarama Hatası: {e}")
        return

    queries = intent.get("matched_queries", [])
    collection_folder = intent.get("collection_folder")
    
    if not queries and not collection_folder:
        if not is_group or bot_mentioned:
            await msg.reply_text("👋 Argon Master Final Pure Drive System Aktif.\nLot kodu (A-371, A-96), Koleksiyon Sunumu veya çeki listesi isteyebilirsiniz.")
        return

    await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.TYPING)
    
    # 2. KOLEKSİYON BAZLI PDF SUNUMU (Tüm klasörü tek PDF yapar)
    if collection_folder and intent.get("want_pdf"):
        await msg.reply_text(f"📁 <b>{collection_folder}</b> taranıp tek bir Premium PDF'te birleştiriliyor...", parse_mode=ParseMode.HTML)
        await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.UPLOAD_DOCUMENT)
        try:
            pdf_b, pdf_n = await asyncio.to_thread(pdf_engine.generate_presentation_pdf, [], collection_folder_name=collection_folder)
            await msg.reply_document(document=io.BytesIO(pdf_b), filename=pdf_n, caption=f"📄 <b>{collection_folder}</b> Premium PDF Kataloğu", parse_mode=ParseMode.HTML)
        except Exception as e:
            await msg.reply_text(f"PDF oluşturulurken hata: {e}")
        return

    # 3. ÇOKLU LOT PDF BİRLEŞTİRİCİSİ (Örn: A-96 ve A-383 sunum hazırla)
    if queries and intent.get("want_pdf"):
        await msg.reply_text(f"📄 Seçilen Lotlar (<b>{', '.join(queries)}</b>) tek bir Premium PDF'te birleştiriliyor...", parse_mode=ParseMode.HTML)
        await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.UPLOAD_DOCUMENT)
        try:
            pdf_b, pdf_n = await asyncio.to_thread(pdf_engine.generate_presentation_pdf, queries)
            await msg.reply_document(document=io.BytesIO(pdf_b), filename=pdf_n, caption="📄 Lüks İngilizce Sunum (Bookmatch'siz)", parse_mode=ParseMode.HTML)
        except Exception as e:
            await msg.reply_text(f"PDF oluşturulurken hata: {e}")
        return # İşlem bitti, çık.

    # 4. NORMAL SORGULAMA VE ÇEKİ LİSTESİ DÖNGÜSÜ
    for q in queries:
        try:
            info = await asyncio.to_thread(inventory_engine.get_lot_info, q)
            d = info["data"]
            fid = info["folder_id"]
            flink = info["drive_link"]

            # ÇEKİ LİSTESİ İSTENDİYSE (Sadece klasörde varsa atar)
            if intent.get("want_packing"):
                await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.UPLOAD_DOCUMENT)
                if info["excel_bytes"]:
                    await msg.reply_document(document=io.BytesIO(info["excel_bytes"]), filename=info["excel_name"], caption=f"📦 <b>{q}</b> Orijinal Excel Çeki Listesi", parse_mode=ParseMode.HTML)
                    if d.get("raw_df") is not None:
                        pdf_pack_bytes = await asyncio.to_thread(pdf_engine.generate_packing_list_pdf, d["raw_df"], q, d["mapped_cols"])
                        await msg.reply_document(document=io.BytesIO(pdf_pack_bytes), filename=f"PackingList_{q}.pdf", caption=f"📄 <b>{q}</b> Resmi PDF Çeki Listesi", parse_mode=ParseMode.HTML)
                else:
                    await msg.reply_text(f"⚠️ {q} klasöründe orijinal çeki listesi (Excel) bulunamadı.")
                continue

            # TELEGRAM İKİLİ FOTO ALBÜM VE BİLGİ KARTI
            stock_str = f"{d.get('total_m2', 0.0):,.2f} m²" if d.get('total_m2', 0.0) > 0 else "Excel Verisi Bekleniyor"
            text_response = (
                f"💎 <b>{q.upper()}</b>\n"
                f"• Mevcut Stok: <code>{stock_str}</code> (<code>{d.get('total_slabs', 0)} Plaka</code>)\n"
                f"• Ortalama Ebat: <code>{d.get('dimensions', '-')}</code>\n"
                f"• Yüzey & Kalınlık: <code>{d.get('surface', '-')} ({d.get('thickness', '-')})</code>\n"
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
            # Botun çökmesini tamamen engelleyen zırh
            await msg.reply_text(f"⚠️ {q} işlenirken teknik bir uyarı oluştu, ancak bot çalışmaya devam ediyor.")

# =====================================================================
# 8. RENDER SIFIR-BEKLEME BAŞLATICI
# =====================================================================
flask_app = Flask(__name__)

@flask_app.route("/")
def health_check(): return "Argon Master Final is Live!", 200

def run_flask():
    flask_app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)), use_reloader=False)

def main():
    if not TELEGRAM_TOKEN:
        logger.error("CRITICAL: TELEGRAM_TOKEN eksik!")
        return

    threading.Thread(target=run_flask, daemon=True).start()
    application = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    
    application.add_handler(MessageHandler((filters.TEXT | filters.PHOTO) & (~filters.COMMAND), handle_message))
    
    logger.info("🚀 Argon Master Final (Native PDF, Deep Scan, 0 Hata) başlatılıyor...")
    application.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
