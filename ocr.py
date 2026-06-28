import streamlit as st
from streamlit.components.v1 import html as st_html
from PIL import Image
import pytesseract
import re
import pandas as pd
import io
import cv2
import numpy as np
import os
import base64
import json
import gspread
from google.oauth2.service_account import Credentials

def save_to_sheets(all_bills):
    import gspread
    try:
        # ดึงข้อมูลจาก Secrets ที่คุณตั้งไว้
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=[
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ])
        
        client = gspread.authorize(creds)
        # ใส่ ID ของไฟล์ที่คุณต้องการบันทึก
        sheet = client.open_by_key("1IhQFHxlK7vlAJ-xxgWNA_M2fcXp-9jZm8WQZGZC4MxQ").worksheet("Data")
        
        # เตรียมแถวข้อมูล (ปรับแก้หัวข้อคอลัมน์ให้ตรงกับที่ใช้จริง)
        rows = []
        for b in all_bills:
            bill = b.get('bill', {})
            rows.append([
                bill.get('date'), bill.get('time'), bill.get('branch_code'),
                bill.get('receipt_number'), bill.get('total_amount'), bill.get('discount')
            ])
        
        sheet.append_rows(rows)
        return True
    except Exception as e:
        st.error(f"บันทึกไม่สำเร็จ: {e}")
        return False
# ── Google Drive / Docs API ───────────────────────────────────────────────────
try:
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseUpload
    from google.oauth2.service_account import Credentials
    _GDRIVE_AVAILABLE = True
except ImportError:
    _GDRIVE_AVAILABLE = False

# ── Tesseract path (Windows) ──────────────────────────────────────────────────
_WIN_PATHS = [
    r'C:\Tesseract-OCR\tesseract.exe',
    r'C:\Program Files\Tesseract-OCR\tesseract.exe',
    r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
]
if os.name == 'nt':
    for _p in _WIN_PATHS:
        if os.path.exists(_p):
            pytesseract.pytesseract.tesseract_cmd = _p
            _td = os.path.join(os.path.dirname(_p), 'tessdata')
            if os.path.exists(_td):
                os.environ['TESSDATA_PREFIX'] = _td
            break

st.set_page_config(page_title="CJ Express Receipt OCR Pro", layout="wide", page_icon="🧾")

st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #f8f7ff; }
.block-container { padding-top: 1.5rem; max-width: 1100px; }
.step-bar { display:flex; align-items:center; margin-bottom:1.5rem; gap:0; }
.step-item { display:flex; align-items:center; gap:6px; }
.step-dot { width:28px; height:28px; border-radius:50%; display:flex; align-items:center;
            justify-content:center; font-size:13px; font-weight:600; flex-shrink:0; }
.step-dot.done { background:#534AB7; color:#fff; }
.step-dot.now  { background:#EEEDFE; color:#534AB7; border:2px solid #534AB7; }
.step-dot.pend { background:#f1f0f0; color:#aaa; border:1px solid #ddd; }
.step-label     { font-size:13px; color:#888; }
.step-label.now { color:#534AB7; font-weight:600; }
.step-line { flex:1; height:2px; background:#d5d3f5; margin:0 8px; min-width:20px; }
.sec-header { font-size:11px; font-weight:600; color:#534AB7;
              letter-spacing:.07em; text-transform:uppercase; margin-bottom:8px; }
.hint-box { background:#E1F5EE; border-radius:8px; padding:9px 14px;
            font-size:13px; color:#085041; margin-bottom:12px; }
div[data-testid="stButton"] button { border-radius:8px !important; }
</style>
""", unsafe_allow_html=True)

_DEFAULTS = dict(step=1, bill_count=0,
                 gallery_files=[], selected_idx=-1,
                 crop_result=None, all_bills=[], lang="th",
                 manual_splits_px=None,
                 crop_applied=False,
                 batch_files=[],
                 ocr_engine="tesseract")

for k, v in _DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v
S = st.session_state

TXT = {
    "th": dict(
        title="🧾 วิเคราะห์บิล CJ Express OCR Pro",
        steps=["อัปโหลด","ปรับภาพ","OCR"],
        count_label="มีกี่บิลในภาพ?",
        b1="1 บิล", b1s="บิลเดียว",
        b2="2 บิล", b2s="เคียงกัน",
        b3="3 บิล", b3s="เรียงกัน",
        upload_label="อัปโหลดรูปภาพ",
        upload_hint="📁 รองรับการเลือกหลายไฟล์พร้อมกัน · JPG PNG JPEG HEIC",
        gallery_label="เลือกรูปที่ต้องการ",
        gallery_count=lambda n: f"พบ {n} รูป — คลิกเพื่อเลือก",
        crop_label="ปรับและ Crop รูป",
        crop_free="✂️ Crop อิสระ",
        crop_a4="📄 สัดส่วน A4",
        crop_hint="ลากบนรูปเพื่อเลือกพื้นที่ · ปล่อยเพื่อยืนยัน · ลากเส้นแดงเพื่อปรับจุดตัดบิล",
        crop_confirm="✅ ยืนยัน Crop",
        crop_skip="⏭️ ข้ามขั้นตอนนี้ (ใช้รูปทั้งหมด)",
        analyze="🔍 เริ่ม OCR วิเคราะห์",
        reset="🔄 เริ่มใหม่",
        result_label="ผลการวิเคราะห์",
        download="📥 ดาวน์โหลด Excel",
        no_items="ไม่พบรายการสินค้า",
        raw_text="ข้อความดิบ (OCR)",
        found=lambda n: f"✅ พบ {n} ใบเสร็จ",
        split_mode_label="วิธีแยกบิล",
        split_manual="✋ ใช้จุดตัดที่ปรับเอง (แนะนำ)",
        split_auto="🤖 ให้ระบบตรวจจับอัตโนมัติ",
        mode_label="โหมดการอัปโหลด",
        mode_single="🖼️ เลือกทีละรูป",
        mode_single_desc="รูปเดียวมีหลายบิล ต้อง Crop/แยกเอง",
        mode_batch="📂 วิเคราะห์ทั้งโฟลเดอร์พร้อมกัน",
        mode_batch_desc="แต่ละไฟล์ = 1 บิล (1 รูป 1 ใบเสร็จ)",
        batch_upload_hint="📂 เลือกไฟล์ทั้งหมดในโฟลเดอร์พร้อมกันได้เลย "
                           "(เปิดหน้าต่างเลือกไฟล์แล้วกด Ctrl+A หรือ Cmd+A เพื่อเลือกทุกไฟล์)",
        batch_found=lambda n: f"📂 พบ {n} ไฟล์ — พร้อมวิเคราะห์ทุกไฟล์เป็นบิลแยกกัน",
        batch_analyze="🔍 วิเคราะห์ทุกรูปพร้อมกัน",
        batch_progress=lambda i, n, name: f"กำลัง OCR {i}/{n}: {name}",
        batch_done=lambda n: f"✅ วิเคราะห์เสร็จสิ้น {n} ไฟล์",
    ),
    "en": dict(
        title="🧾 CJ Express Receipt OCR Pro",
        steps=["Upload","Enhance","OCR"],
        count_label="How many bills in the image?",
        b1="1 bill", b1s="single",
        b2="2 bills", b2s="side by side",
        b3="3 bills", b3s="in a row",
        upload_label="Upload image(s)",
        upload_hint="📁 Select multiple files at once · JPG PNG JPEG HEIC",
        gallery_label="Select image",
        gallery_count=lambda n: f"{n} images — click to select",
        crop_label="Adjust & Crop",
        crop_free="✂️ Free crop",
        crop_a4="📄 A4 ratio",
        crop_hint="Drag on image to select area · release to confirm · drag red line to adjust split",
        crop_confirm="✅ Apply Crop",
        crop_skip="⏭️ Skip crop (use full image)",
        analyze="🔍 Run OCR Analysis",
        reset="🔄 Reset",
        result_label="Analysis results",
        download="📥 Download Excel",
        no_items="No items found",
        raw_text="Raw OCR text",
        found=lambda n: f"✅ Found {n} receipt(s)",
        split_mode_label="Bill split method",
        split_manual="✋ Use adjusted split lines (recommended)",
        split_auto="🤖 Auto-detect",
        mode_label="Upload mode",
        mode_single="🖼️ Select one by one",
        mode_single_desc="One image has multiple bills, needs crop/split",
        mode_batch="📂 Analyze whole folder at once",
        mode_batch_desc="Each file = 1 bill (1 image = 1 receipt)",
        batch_upload_hint="📂 Select all files in a folder at once "
                           "(open the file picker, then press Ctrl+A or Cmd+A to select all)",
        batch_found=lambda n: f"📂 Found {n} files — ready to analyze each as a separate bill",
        batch_analyze="🔍 Analyze all images at once",
        batch_progress=lambda i, n, name: f"OCR {i}/{n}: {name}",
        batch_done=lambda n: f"✅ Finished analyzing {n} files",
    ),
}
def t(key): return TXT[S.lang][key]

def render_steps():
    steps = t("steps")
    html = '<div class="step-bar">'
    for i, label in enumerate(steps, 1):
        cls = "done" if i < S.step else ("now" if i == S.step else "pend")
        dot = "✓" if i < S.step else str(i)
        lc  = "now" if i == S.step else ""
        html += f'<div class="step-item"><div class="step-dot {cls}">{dot}</div>'
        html += f'<span class="step-label {lc}">{label}</span></div>'
        if i < len(steps):
            html += '<div class="step-line"></div>'
    st.markdown(html + '</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# Image utils
# ─────────────────────────────────────────────────────────────────────────────
def pil_to_cv(img): return cv2.cvtColor(np.array(img.convert("RGB")), cv2.COLOR_RGB2BGR)
def cv_to_pil(img): return Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))

def img_to_b64(pil_img):
    buf = io.BytesIO(); pil_img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode()

def img_to_bytes_png(img_cv):
    try:
        buf = io.BytesIO(); cv_to_pil(img_cv).save(buf, format="PNG")
        return buf.getvalue()
    except Exception: return None

def find_paper_contours(img_cv, min_area_frac=0.015):
    H, W = img_cv.shape[:2]
    hsv = cv2.cvtColor(img_cv, cv2.COLOR_BGR2HSV)
    h, s, v = cv2.split(hsv)
    paper_mask = ((s < 60) & (v > 120)).astype(np.uint8) * 255
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (9, 9))
    mask = cv2.morphologyEx(paper_mask, cv2.MORPH_CLOSE, kernel)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel)
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    min_area = H * W * min_area_frac
    results = [(cv2.boundingRect(c), c) for c in contours if cv2.contourArea(c) > min_area]
    results.sort(key=lambda r: r[0][0])
    return results, mask

def deskew_paper(img_cv, contour, pad=10):
    H, W = img_cv.shape[:2]
    rect = cv2.minAreaRect(contour)
    (cx, cy), (rw, rh), angle = rect
    if rw > rh:
        rw, rh = rh, rw
        angle += 90
    box = cv2.boxPoints(rect).astype(np.float32)
    out_w, out_h = max(int(rw) + pad * 2, 10), max(int(rh) + pad * 2, 10)
    s = box.sum(axis=1)
    diff = np.diff(box, axis=1).flatten()
    tl, br = box[np.argmin(s)], box[np.argmax(s)]
    tr, bl = box[np.argmin(diff)], box[np.argmax(diff)]
    src_pts = np.array([tl, tr, br, bl], dtype=np.float32)
    dst_pts = np.array([
        [pad, pad], [out_w - pad, pad],
        [out_w - pad, out_h - pad], [pad, out_h - pad]
    ], dtype=np.float32)
    M = cv2.getPerspectiveTransform(src_pts, dst_pts)
    warped = cv2.warpPerspective(img_cv, M, (out_w, out_h), borderValue=(255, 255, 255))
    mask_full = np.zeros((H, W), dtype=np.uint8)
    cv2.drawContours(mask_full, [contour], -1, 255, thickness=cv2.FILLED)
    warped_mask = cv2.warpPerspective(mask_full, M, (out_w, out_h), borderValue=0)
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (7, 7))
    warped_mask = cv2.dilate(warped_mask, kernel, iterations=1)
    return warped, warped_mask

def correct_illumination(gray):
    kernel_size = max(gray.shape) // 15
    kernel_size = kernel_size if kernel_size % 2 == 1 else kernel_size + 1
    kernel_size = max(kernel_size, 3)
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (kernel_size, kernel_size))
    background = cv2.morphologyEx(gray, cv2.MORPH_CLOSE, kernel)
    gray_f = gray.astype(np.float32)
    bg_f = background.astype(np.float32) + 1e-6
    corrected = (gray_f / bg_f) * 255
    return np.clip(corrected, 0, 255).astype(np.uint8)

def whiten_background(img_cv, contour=None, bbox=None, pad=6, deskew=True):
    H, W = img_cv.shape[:2]
    if contour is not None and deskew:
        warped, warped_mask = deskew_paper(img_cv, contour, pad=pad + 4)
        gray = cv2.cvtColor(warped, cv2.COLOR_BGR2GRAY)
        gray = np.where(warped_mask > 0, gray, 255).astype(np.uint8)
    elif bbox is not None:
        x, y, w, h = bbox
        x0, y0 = max(x - pad, 0), max(y - pad, 0)
        x1, y1 = min(x + w + pad, W), min(y + h + pad, H)
        crop = img_cv[y0:y1, x0:x1]
        gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
        if contour is not None:
            mask_full = np.zeros((H, W), dtype=np.uint8)
            cv2.drawContours(mask_full, [contour], -1, 255, thickness=cv2.FILLED)
            kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (5, 5))
            mask_full = cv2.dilate(mask_full, kernel, iterations=1)
            mask_crop = mask_full[y0:y1, x0:x1]
            gray = np.where(mask_crop > 0, gray, 255).astype(np.uint8)
    else:
        gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
    gray = correct_illumination(gray)
    gray = cv2.bilateralFilter(gray, d=5, sigmaColor=30, sigmaSpace=30)
    clahe = cv2.createCLAHE(clipLimit=1.5, tileGridSize=(8, 8))
    return clahe.apply(gray)

def auto_crop_receipts(img_cv, n_expected=None, deskew=True):
    results, mask = find_paper_contours(img_cv)
    h, w = img_cv.shape[:2]
    if not results:
        return [img_cv]
    if n_expected and len(results) != n_expected:
        if abs(len(results) - n_expected) > 1:
            return [img_cv]
    crops = []
    for bbox, contour in results:
        if deskew:
            warped, warped_mask = deskew_paper(img_cv, contour, pad=10)
            white_bg = np.full_like(warped, 255)
            mask_3ch = cv2.cvtColor(warped_mask, cv2.COLOR_GRAY2BGR)
            composited = np.where(mask_3ch > 0, warped, white_bg)
            crops.append(composited)
        else:
            x, y, cw, ch = bbox
            pad = 6
            x0, y0 = max(x - pad, 0), max(y - pad, 0)
            x1, y1 = min(x + cw + pad, w), min(y + ch + pad, h)
            crops.append(img_cv[y0:y1, x0:x1])
    return crops if crops else [img_cv]

def preprocess_image(img_cv, whiten=True):
    if whiten:
        gray = whiten_background(img_cv)
    else:
        gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
    gray = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    gray = cv2.fastNlMeansDenoising(gray, h=10)
    blur = cv2.GaussianBlur(gray, (0, 0), 3)
    gray = cv2.addWeighted(gray, 1.5, blur, -0.5, 0)
    return cv2.adaptiveThreshold(gray, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 15)

# ─────────────────────────────────────────────────────────────────────────────
# Bill split utils
# ─────────────────────────────────────────────────────────────────────────────
def _find_bill_splits(col_ratio, w, min_bill_frac=0.15, n_expected=None):
    kernel = np.ones(5) / 5
    smooth = np.convolve(col_ratio, kernel, mode='same')
    thr_outer = max(0.03, smooth.max() * 0.05)
    content_cols = np.where(smooth > thr_outer)[0]
    if len(content_cols) == 0:
        return [], smooth, smooth.mean()
    x0 = max(content_cols[0]  - 5, 0)
    x1 = min(content_cols[-1] + 5, w-1)
    region_w = x1 - x0
    if region_w <= 0:
        return [], smooth, smooth.mean()
    seg = smooth[x0:x1]
    max_v = seg.max()

    def _candidates(gap_ratio):
        gap_thr = max_v * gap_ratio
        grad = np.gradient(seg)
        return [(i, float(seg[i])) for i in range(3, len(seg)-3)
                if grad[i-1] <= 0 and grad[i+1] >= 0 and seg[i] < gap_thr]

    def _select_best(candidates, min_bw, n_keep=None):
        ordered = sorted(candidates, key=lambda x: x[1])
        chosen = []
        for pos, val in ordered:
            if pos < min_bw or pos > region_w - min_bw: continue
            if all(abs(pos - p) >= min_bw for p in chosen):
                chosen.append(pos)
            if n_keep and len(chosen) >= n_keep: break
        return sorted(chosen)

    min_bw = max(int(region_w * min_bill_frac), 8)
    target_n = (n_expected - 1) if n_expected else None
    candidates = _candidates(0.55)
    splits_rel = _select_best(candidates, min_bw, n_keep=target_n)
    if n_expected and len(splits_rel) < target_n:
        for gap_ratio in (0.65, 0.75, 0.85, 0.92):
            wider_min_bw = max(int(region_w * max(min_bill_frac * 0.7, 0.10)), 8)
            attempt_candidates = _candidates(gap_ratio)
            attempt = _select_best(attempt_candidates, wider_min_bw, n_keep=target_n)
            if len(attempt) > len(splits_rel):
                splits_rel = attempt
            if len(splits_rel) >= target_n:
                break
    splits_abs = [x0 + p for p in splits_rel]
    mean_v = float(seg.mean())
    return splits_abs, smooth, mean_v

def split_receipts_image(img_cv, n_expected=None):
    h, w = img_cv.shape[:2]
    gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
    col_ratio = (gray > 180).astype(np.uint8).mean(axis=0)
    splits, smooth, mean_val = _find_bill_splits(col_ratio, w, n_expected=n_expected)
    thr_outer = max(0.03, smooth.max() * 0.05)
    content = np.where(smooth > thr_outer)[0]
    x_start = max(content[0]  - 5, 0)   if len(content) else 0
    x_end   = min(content[-1] + 5, w-1) if len(content) else w-1
    bounds = [x_start] + splits + [x_end]
    crops  = []
    min_cw = max(int(w * 0.08), 20)
    for i in range(len(bounds)-1):
        xa, xb = bounds[i], bounds[i+1]
        if (xb - xa) < min_cw: continue
        seg = smooth[xa:xb]
        content_seg = np.where(seg > thr_outer)[0]
        if len(content_seg) == 0: continue
        ca = max(xa + content_seg[0]  - 3, 0)
        cb = min(xa + content_seg[-1] + 3, w-1)
        if (cb - ca) >= min_cw:
            crops.append(img_cv[:, ca:cb+1])
    return crops if crops else [img_cv]

def split_by_positions(img_cv, split_px_list):
    h, w = img_cv.shape[:2]
    pts = sorted(set(int(p) for p in split_px_list if 0 < p < w))
    if not pts:
        return [img_cv]
    bounds = [0] + pts + [w]
    crops = []
    for i in range(len(bounds)-1):
        xa, xb = bounds[i], bounds[i+1]
        if xb - xa < 5: continue
        crops.append(img_cv[:, xa:xb])
    return crops if crops else [img_cv]

# ─────────────────────────────────────────────────────────────────────────────
# Google Drive OCR
# ─────────────────────────────────────────────────────────────────────────────
_GDRIVE_SCOPES = (
    "https://www.googleapis.com/auth/drive "
    "https://www.googleapis.com/auth/spreadsheets "
    "https://www.googleapis.com/auth/userinfo.email "
    "https://www.googleapis.com/auth/userinfo.profile"
)
_GDRIVE_AUTH_URL  = "https://accounts.google.com/o/oauth2/v2/auth"
_GDRIVE_TOKEN_URL = "https://oauth2.googleapis.com/token"

def _gdrive_client_config() -> dict:
    try:
        cfg = st.secrets.get("gdrive", {})
        return {
            "client_id":     cfg.get("client_id",     os.environ.get("GDRIVE_CLIENT_ID", "")).strip(),
            "client_secret": cfg.get("client_secret", os.environ.get("GDRIVE_CLIENT_SECRET", "")).strip(),
            "redirect_uri":  cfg.get("redirect_uri",  os.environ.get("GDRIVE_REDIRECT_URI",
                                                                       "http://localhost:8501/")).strip(),
        }
    except Exception:
        return {
            "client_id":     os.environ.get("GDRIVE_CLIENT_ID", "").strip(),
            "client_secret": os.environ.get("GDRIVE_CLIENT_SECRET", "").strip(),
            "redirect_uri":  os.environ.get("GDRIVE_REDIRECT_URI", "http://localhost:8501/").strip(),
        }

def is_gdrive_configured() -> bool:
    if not _GDRIVE_AVAILABLE:
        return False
    cfg = _gdrive_client_config()
    return bool(cfg["client_id"] and cfg["client_secret"])

def is_gdrive_token_ready() -> bool:
    token = st.session_state.get("gdrive_token")
    if not token:
        return False
    import time
    expires_at = st.session_state.get("gdrive_token_expires_at", 0)
    if expires_at and time.time() > expires_at - 60:
        return False
    return True

def gdrive_get_auth_url() -> str:
    import urllib.parse, secrets as _sec
    cfg = _gdrive_client_config()
    state = _sec.token_urlsafe(16)
    st.session_state["gdrive_oauth_state"] = state
    params = {
        "client_id":     cfg["client_id"],
        "redirect_uri":  cfg["redirect_uri"],
        "response_type": "code",
        "scope":         _GDRIVE_SCOPES,
        "access_type":   "offline",
        "prompt":        "consent",
        "state":         state,
    }
    return f"{_GDRIVE_AUTH_URL}?{urllib.parse.urlencode(params)}"

def gdrive_exchange_code(code: str) -> bool:
    import requests as _req, time
    cfg = _gdrive_client_config()
    try:
        resp = _req.post(_GDRIVE_TOKEN_URL, data={
            "code":          code,
            "client_id":     cfg["client_id"],
            "client_secret": cfg["client_secret"],
            "redirect_uri":  cfg["redirect_uri"],
            "grant_type":    "authorization_code",
        }, timeout=15)
        resp.raise_for_status()
        data = resp.json()
        st.session_state["gdrive_token"]            = data["access_token"]
        st.session_state["gdrive_refresh_token"]    = data.get("refresh_token", "")
        st.session_state["gdrive_token_expires_at"] = time.time() + data.get("expires_in", 3600)
        return True
    except Exception as e:
        st.session_state["gdrive_auth_error"] = str(e)
        return False

def gdrive_refresh_token_if_needed() -> bool:
    import requests as _req, time
    if is_gdrive_token_ready():
        return True
    refresh_token = st.session_state.get("gdrive_refresh_token", "")
    if not refresh_token:
        return False
    cfg = _gdrive_client_config()
    try:
        resp = _req.post(_GDRIVE_TOKEN_URL, data={
            "refresh_token": refresh_token,
            "client_id":     cfg["client_id"],
            "client_secret": cfg["client_secret"],
            "grant_type":    "refresh_token",
        }, timeout=15)
        resp.raise_for_status()
        data = resp.json()
        st.session_state["gdrive_token"]            = data["access_token"]
        st.session_state["gdrive_token_expires_at"] = time.time() + data.get("expires_in", 3600)
        return True
    except Exception:
        st.session_state.pop("gdrive_token", None)
        st.session_state.pop("gdrive_refresh_token", None)
        return False

def gdrive_get_user_info() -> dict:
    import requests as _req
    token = st.session_state.get("gdrive_token", "")
    if not token:
        return {}
    try:
        resp = _req.get(
            "https://www.googleapis.com/oauth2/v2/userinfo",
            headers={"Authorization": f"Bearer {token}"},
            timeout=10,
        )
        return resp.json() if resp.ok else {}
    except Exception:
        return {}

def run_ocr_google_drive(crop_cv) -> str:
    import requests as _req
    if not gdrive_refresh_token_if_needed():
        raise RuntimeError("กรุณา Login Google Drive ก่อนใช้งาน")
    token = st.session_state["gdrive_token"]
    gray = whiten_background(crop_cv)
    img_color = cv2.cvtColor(gray, cv2.COLOR_GRAY2BGR)
    ok, buf = cv2.imencode(".jpg", img_color, [cv2.IMWRITE_JPEG_QUALITY, 92])
    if not ok:
        raise RuntimeError("แปลงภาพเป็น JPEG ไม่สำเร็จ")
    jpg_bytes = buf.tobytes()
    headers = {"Authorization": f"Bearer {token}"}
    import email.mime.multipart, email.mime.base, email.mime.application
    metadata = json.dumps({
        "name": "_ocr_tmp",
        "mimeType": "application/vnd.google-apps.document",
    }).encode()
    boundary = "OCR_BOUNDARY_XYZ"
    body = (
        f"--{boundary}\r\n"
        f"Content-Type: application/json; charset=UTF-8\r\n\r\n"
    ).encode() + metadata + (
        f"\r\n--{boundary}\r\n"
        f"Content-Type: image/jpeg\r\n\r\n"
    ).encode() + jpg_bytes + f"\r\n--{boundary}--".encode()
    upload_resp = _req.post(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id",
        headers={**headers, "Content-Type": f"multipart/related; boundary={boundary}"},
        data=body,
        timeout=60,
    )
    upload_resp.raise_for_status()
    doc_id = upload_resp.json()["id"]
    raw_text = ""
    try:
        export_resp = _req.get(
            f"https://www.googleapis.com/drive/v3/files/{doc_id}/export",
            params={"mimeType": "text/plain"},
            headers=headers,
            timeout=30,
        )
        export_resp.raise_for_status()
        raw_text = export_resp.content.decode("utf-8")
        raw_text = raw_text.lstrip('\ufeff')
        raw_text = raw_text.replace('\f', '\n').replace('\x0c', '\n')
        lines = raw_text.split('\n')
        start = 0
        for i, line in enumerate(lines):
            if re.search(r'^_{5,}$', line.strip()):
                start = i + 1
            elif line.strip() and start > 0:
                break
        if start > 0:
            raw_text = '\n'.join(lines[start:])
        _lines = [l for l in raw_text.split('\n') if l.strip()]
        _half = len(_lines) // 2
        if _half > 5 and _lines[:_half] == _lines[_half:]:
            raw_text = '\n'.join(_lines[:_half])
    finally:
        try:
            _req.delete(f"https://www.googleapis.com/drive/v3/files/{doc_id}",
                        headers=headers, timeout=10)
        except Exception:
            pass
        try:
            _req.delete("https://www.googleapis.com/drive/v3/files/trash",
                        headers=headers, timeout=10)
        except Exception:
            pass
    if "_gdrive_raw_texts" not in st.session_state:
        st.session_state["_gdrive_raw_texts"] = []
    st.session_state["_gdrive_raw_texts"].append(raw_text)
    return clean_text(raw_text) if raw_text.strip() else ""

def render_gdrive_login_ui():
    params = st.query_params
    if "code" in params and "gdrive_token" not in st.session_state:
        code  = params["code"]
        state = params.get("state", "")
        if state == st.session_state.get("gdrive_oauth_state", ""):
            with st.spinner("กำลัง Login..."):
                ok = gdrive_exchange_code(code)
            if ok:
                st.query_params.clear()
                st.rerun()
            else:
                err = st.session_state.pop("gdrive_auth_error", "unknown error")
                st.error(f"Login ไม่สำเร็จ: {err}")
        else:
            st.warning("OAuth state ไม่ตรง — กรุณาลอง Login ใหม่")
        st.query_params.clear()
    if not is_gdrive_configured():
        st.warning(
            "⚠️ ยังไม่ได้ตั้งค่า OAuth Client\n\n"
            "**วิธีตั้งค่า:**\n"
            "1. https://console.cloud.google.com → APIs & Services → Credentials\n"
            "2. **+ CREATE CREDENTIALS** → **OAuth client ID** → Web application\n"
            "3. Authorized redirect URIs: ใส่ URL ของ Streamlit app\n"
            "4. ใส่ใน `.streamlit/secrets.toml`:\n"
            "```toml\n[gdrive]\nclient_id = \"xxx\"\nclient_secret = \"xxx\"\nredirect_uri = \"https://your-app.streamlit.app/\"\n```"
        )
        return
    cfg = _gdrive_client_config()
    with st.expander("🔍 Debug — ค่าที่อ่านได้จาก secrets", expanded=False):
        st.code(f'client_id = "{cfg["client_id"]}"\nclient_secret = "{cfg["client_secret"][:8]}..."\nredirect_uri = "{cfg["redirect_uri"]}"')
        auth_url = gdrive_get_auth_url()
        st.markdown(f"[คลิกทดสอบ Login ตรงๆ]({auth_url})")
    if is_gdrive_token_ready():
        info = gdrive_get_user_info()
        name  = info.get("name", "")
        email = info.get("email", "")
        pic   = info.get("picture", "")
        col_pic, col_info = st.columns([1, 4])
        if pic:
            col_pic.image(pic, width=48)
        col_info.success(f"✅ Login แล้ว: **{name}** ({email})")
        if st.button("🔓 Logout", key="gdrive_logout_btn"):
            for k in ["gdrive_token","gdrive_refresh_token","gdrive_token_expires_at","gdrive_oauth_state"]:
                st.session_state.pop(k, None)
            st.rerun()
    else:
        auth_url = gdrive_get_auth_url()
        st.info("🔑 กดปุ่มด้านล่างเพื่อ Login Google Drive")
        st.markdown(
            f"""<a href="{auth_url}" target="_self"
               style="display:inline-block;padding:10px 24px;background:#4285F4;
                      color:white;border-radius:8px;text-decoration:none;font-weight:600;font-size:15px">
               🔐 Login ด้วย Google Drive
            </a>""",
            unsafe_allow_html=True,
        )
        with st.expander("⚙️ Login ด้วยตัวเอง (ถ้าปุ่มด้านบนไม่ทำงาน)"):
            st.markdown(f"**[คลิกที่นี่เพื่อ Login]({auth_url})**\n\nCopy ค่าหลัง `?code=` มาวางด้านล่าง:")
            manual_code = st.text_input("วาง Authorization Code:", key="gdrive_manual_code", placeholder="4/0AX4XfWh...")
            if st.button("✅ ยืนยัน Code", key="gdrive_manual_submit", type="primary"):
                if manual_code.strip():
                    with st.spinner("กำลัง Login..."):
                        ok = gdrive_exchange_code(manual_code.strip())
                    if ok:
                        st.success("✅ Login สำเร็จ!")
                        st.rerun()
                    else:
                        err = st.session_state.pop("gdrive_auth_error", "unknown")
                        st.error(f"Login ไม่สำเร็จ: {err}")

# ─────────────────────────────────────────────────────────────────────────────
# Google Vision API
# ─────────────────────────────────────────────────────────────────────────────
_VISION_API_KEY = os.environ.get("GOOGLE_VISION_API_KEY", "")

def is_vision_api_configured() -> bool:
    return bool(_VISION_API_KEY.strip())

def run_ocr_google_vision(crop_cv) -> str:
    import requests
    if not is_vision_api_configured():
        raise RuntimeError("ยังไม่ได้ตั้งค่า GOOGLE_VISION_API_KEY")
    gray = whiten_background(crop_cv)
    img_for_api = cv2.cvtColor(gray, cv2.COLOR_GRAY2BGR)
    success, buf = cv2.imencode(".png", img_for_api)
    if not success:
        raise RuntimeError("แปลงภาพเป็น PNG ไม่สำเร็จ")
    image_b64 = base64.b64encode(buf.tobytes()).decode("utf-8")
    url = f"https://vision.googleapis.com/v1/images:annotate?key={_VISION_API_KEY}"
    payload = {"requests": [{"image": {"content": image_b64},
                              "features": [{"type": "DOCUMENT_TEXT_DETECTION"}],
                              "imageContext": {"languageHints": ["th", "en"]}}]}
    resp = requests.post(url, json=payload, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    resp0 = data.get("responses", [{}])[0]
    if "error" in resp0:
        raise RuntimeError(f"Google Vision API error: {resp0['error'].get('message', resp0['error'])}")
    full_text = resp0.get("fullTextAnnotation", {}).get("text", "")
    return clean_text(full_text) if full_text else ""

def run_ocr(crop_cv, engine: str = "tesseract"):
    if engine == "gdrive":
        try:
            text = run_ocr_google_drive(crop_cv)
            if text:
                return text
        except Exception as e:
            st.session_state.setdefault("_gdrive_warning", str(e))
    if engine == "vision":
        try:
            text = run_ocr_google_vision(crop_cv)
            if text:
                return text
        except Exception as e:
            st.session_state.setdefault("_vision_api_warning", str(e))
    try:
        text = pytesseract.image_to_string(
            preprocess_image(crop_cv), lang='tha+eng', config='--psm 6')
        return clean_text(text)
    except Exception as e:
        return f"[OCR ERROR] {e}"
# ── Cache ข้อมูลจาก Sheet ──────────────────────────────────────────────────
def _load_correct_items_from_sheet(token: str) -> list:
    """โหลดรายการที่ถูกต้องจาก Sheet มา cache"""
    import requests as _req
    try:
        url = (f"https://sheets.googleapis.com/v4/spreadsheets/{_SHEETS_FILE_ID}"
               f"/values/{_SHEET_CORRECT}!A:K")
        resp = _req.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=10)
        if not resp.ok: return []
        rows = resp.json().get("values", [])
        if len(rows) < 2: return []
        # header: วันที่บันทึก, ไฟล์บิล, วันที่บิล, รหัสสาขา,
        #         ชื่อ OCR (ดิบ), ชื่อที่ถูกต้อง, หมวดหมู่, จำนวน, ราคาต่อหน่วย, ยอดรวม, หมายเหตุ
        result = []
        for row in rows[1:]:  # skip header
            if len(row) >= 7:
                result.append({
                    "ocr_name":    row[4] if len(row) > 4 else "",
                    "correct_name": row[5] if len(row) > 5 else "",
                    "category":    row[6] if len(row) > 6 else "",
                    "unit_price":  float(row[8]) if len(row) > 8 and row[8] else 0.0,
                })
        return result
    except Exception:
        return []


def _find_similar_item(ocr_name: str, known_items: list, threshold: float = 0.6) -> dict | None:
    """หา item ที่ชื่อคล้ายกันจาก known_items"""
    if not ocr_name or not known_items:
        return None
    
    def _similarity(a: str, b: str) -> float:
        """คำนวณความคล้ายกันแบบง่ายๆ"""
        a = re.sub(r'\s+', '', a.lower())
        b = re.sub(r'\s+', '', b.lower())
        if not a or not b: return 0.0
        if a == b: return 1.0
        if a in b or b in a: return 0.85
        # นับตัวอักษรร่วมกัน
        common = sum(1 for c in a if c in b)
        return common / max(len(a), len(b))
    
    best_match = None
    best_score = 0.0
    
    for item in known_items:
        # เปรียบเทียบกับ ocr_name เดิม
        score1 = _similarity(ocr_name, item.get("ocr_name", ""))
        # เปรียบเทียบกับ correct_name ด้วย
        score2 = _similarity(ocr_name, item.get("correct_name", ""))
        score = max(score1, score2)
        
        if score > best_score and score >= threshold:
            best_score = score
            best_match = item
    
    return best_match if best_match else None


def enrich_items_from_sheet(items: list, token: str) -> list:
    """แทนที่ชื่อ/หมวดหมู่ที่ผิดด้วยข้อมูลจาก Sheet"""
    if not token or not items:
        return items
    
    # โหลด cache (ทำครั้งเดียวต่อ session)
    cache_key = "_sheet_correct_items_cache"
    if cache_key not in st.session_state:
        with st.spinner("🔍 โหลดข้อมูลอ้างอิงจาก Google Sheets..."):
            st.session_state[cache_key] = _load_correct_items_from_sheet(token)
    
    known_items = st.session_state[cache_key]
    if not known_items:
        return items
    
    enriched = []
    for it in items:
        name = it.get("ชื่อสินค้า", "")
        cat  = it.get("หมวดหมู่", "")
        
        # ถ้าชื่อดูเหมือน noise หรือหมวดเป็น เบ็ดเตล็ด → ลอง lookup
        should_lookup = (
            len(name) < 4 or
            cat == "สินค้าเบ็ดเตล็ดอื่นๆ" or
            re.search(r'[A-Za-z]{3,}(?:\d+)?$', name) or  # ชื่อเป็น Latin ล้วน
            not re.search(r'[ก-๙]', name)  # ไม่มีภาษาไทยเลย
        )
        
        if should_lookup:
            match = _find_similar_item(name, known_items)
            if match:
                new_name = match["correct_name"]
                new_cat  = match["category"] or cat
                st.session_state.setdefault("_sheet_enriched_log", []).append(
                    f'"{name}" → "{new_name}" [{new_cat}]'
                )
                enriched.append({
                    **it,
                    "ชื่อสินค้า":    new_name,
                    "หมวดหมู่":     new_cat,
                    "ชื่อ OCR เดิม": name,  # เก็บชื่อเดิมไว้
                    "from_sheet":    True,   # flag ว่ามาจาก Sheet
                })
                continue
        
        enriched.append(it)
    
    return enriched
# ─────────────────────────────────────────────────────────────────────────────
# Gemini API
# ─────────────────────────────────────────────────────────────────────────────
_GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent"

def _get_gemini_keys() -> list:
    """ดึง Gemini API keys ทั้งหมด (GEMINI_API_KEY_1, GEMINI_API_KEY_2, ... _11)"""
    keys = []
    try:
        # key หลัก
        k = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))
        if k.strip(): keys.append(k.strip())
        # key เพิ่มเติม _2 ถึง _11
        for i in range(1, 12):
            k = st.secrets.get(f"GEMINI_API_KEY_{i}", os.environ.get(f"GEMINI_API_KEY_{i}", ""))
            if k.strip(): keys.append(k.strip())
    except Exception:
        k = os.environ.get("GEMINI_API_KEY", "")
        if k.strip(): keys.append(k.strip())
    return keys

def is_gemini_configured() -> bool:
    return bool(_get_gemini_keys())

def _call_gemini(prompt: str, max_tokens: int = 1500) -> str:
    import requests as _req, time
    keys = _get_gemini_keys()
    if not keys:
        raise RuntimeError("ยังไม่ได้ตั้งค่า GEMINI_API_KEY")

    if "_gemini_key_idx" not in st.session_state:
        st.session_state["_gemini_key_idx"] = 0

    n = len(keys)
    start_idx = st.session_state["_gemini_key_idx"]

    for attempt in range(n * 2):
        idx = (start_idx + attempt) % n
        key = keys[idx]
        try:
            resp = _req.post(
                f"{_GEMINI_API_URL}?key={key}",
                json={
                    "contents": [{"parts": [{"text": prompt}]}],
                    "generationConfig": {"maxOutputTokens": max_tokens, "temperature": 0.1},
                },
                timeout=30,
            )
            if resp.status_code == 429:
                # rate limit → ลอง key ถัดไป
                continue
            if resp.status_code in (401, 403):
                # key ผิด/หมดอายุ → ลอง key ถัดไป
                continue
            if resp.status_code >= 500:
                # server error → ลอง key ถัดไป
                time.sleep(2)
                continue
            resp.raise_for_status()
            data = resp.json()
            st.session_state["_gemini_key_idx"] = (idx + 1) % n
            return data["candidates"][0]["content"]["parts"][0]["text"]
        except _req.exceptions.Timeout:
            # timeout → ลอง key ถัดไป
            continue
        except _req.exceptions.RequestException:
            continue

    raise RuntimeError(f"Gemini ไม่ตอบสนอง ({n} keys) — รอสักครู่แล้วลองใหม่")

# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────
BAO_CAFE_CATEGORY = "Bao Cafe"

CATEGORY_PROMPT = """หมวดหมู่ที่ใช้ได้ (เลือกได้เฉพาะ 8 หมวดนี้เท่านั้น):
1. Bao Cafe — Bao ลาเต้, Bao อเมริกาโน่, Bao เอสเปรสโซ่, Bao โคโค่ ฯลฯ
2. อาหารพร้อมทานและเบเกอรี่ — ข้าวกล่อง แซนด์วิช มาม่า ขนมปัง เค้ก ยูโรเค้ก
3. ขนมและของขบเคี้ยว — ลูกอม ช็อกโกแลต มันฝรั่ง ถั่ว เยลลี่ สแน็ค
4. เครื่องดื่ม — น้ำดื่ม น้ำอัดลม ชา กาแฟ นม ไมโล คาราบาว
5. ของใช้ส่วนตัว — สบู่ แชมพู ยาสีฟัน โลชั่น ผ้าอนามัย
6. ของใช้ในบ้าน — ผงซักฟอก น้ำยาล้างจาน กระดาษทิชชู ถุงขยะ
7. เวชภัณฑ์และอุปกรณ์ดูแลสุขภาพ — ยา พลาสเตอร์ หน้ากาก วิตามิน
8. สินค้าเบ็ดเตล็ดอื่นๆ — ถ่านไฟฉาย ปากกา บุหรี่ บัตรเติมเงิน
หมวดพิเศษ: ส่วนลด/โปรโมชั่น — สำหรับส่วนลดและโปรโมชั่นเท่านั้น"""

BAO_CAFE_MENU = """เมนู Bao Cafe: ลาเต้ | อเมริกาโน่ | เอสเปรสโซ่ | คาปูชิโน่ | โกโก้ | ชาเขียว | ชาไทย (ร้อน/เย็น)
OCR มักอ่านผิด: "1B30", "1B20", "lB30" → "Bao" เสมอ (1/l = B, 30/20 = ชื่อเมนู)
"1B30.อเมริกาไปเป็น" → "Bao อเมริกาโน่เย็น"
OCR มักอ่านผิด: "Bao_", "Beo", "B80", "Be0", "Ba0" → ล้วนหมายถึง "Bao"
"Bao_อเมริกาโน่เป็น" → "Bao อเมริกาโน่เย็น" (_ = space, เป็น = เย็น)"""


CJ_PRODUCT_KNOWLEDGE = """สินค้าในร้าน CJ Express:
อาหาร: มาม่า ยูโรเค้ก ขนมปัง แซนด์วิช ไส้กรอก ข้าวกล่อง
ขนม: เฮอร์ชีย์ โอโชะ มันฝรั่ง ป๊อปคอร์น เยลลี่ ลูกอม
เครื่องดื่ม: น้ำดื่ม สไปรท์ โคคา เป๊ปซี่ คาราบาว ไมโล โออิชิ อราวน์
ของใช้: สบู่ แชมพู ยาสีฟัน ผ้าอนามัย ทิชชู ผงซักฟอก
สินค้าเบ็ดเตล็ด: บุหรี่ ไฟแช็ก ปากกา สมุด บัตรเติมเงิน"""

def _is_bao_item(name: str) -> bool:
    """ตรวจว่าชื่อสินค้าเป็นเมนู Bao Cafe รวม OCR variants ทั้งหมด"""
    return bool(re.search(
        r'\b[Bb8][a8AeE๐4][oO0๐gGcCqQ][_\W]|'  # Bag_, Bac_, B4o_ ฯลฯ
        r'\b[Bb8][a8AeE๐4][oO0๐gGcCqQ]$|'        # Bag, Bac, B4o ท้ายคำ
        r'\b[Bb8][a8AeE๐4][oO0๐][\W_]|'          # Bao, BAO ตามด้วย _/space
        r'\b[Bb8][a8AeE๐4][oO0๐]$|'               # Bao ท้ายคำ
        r'\bbao\b|'                 # bao ตรงๆ
        r'\b1[Bb][23]0\b|'          # ← เพิ่ม: 1B30 = Bao
        r'\b[1Il][Bb]\d0\b|'        # ← เพิ่ม: variants                                
        r'เบา\s*คา|บาว\s*คา',                     # บาว คาเฟ่
        name, re.IGNORECASE
    ))

def _categorize_by_rule(name: str) -> str:
    if _is_bao_item(name):
        return BAO_CAFE_CATEGORY
    # ── FIX: โปรโมชั่นและส่วนลด ──
    if re.search(r'ส่วนลด|โปรโมชั่น|โปรโม|discount|promotion|แถม', name, re.IGNORECASE):
        return "ส่วนลด/โปรโมชั่น"
    n = name.lower()
    if re.search(r'มาม่า|ไวไว|ยำยำ|บะหมี่|ข้าว|แซนด์|ขนมปัง|เค้ก|ยูโร|ไส้กรอก|'
                 r'หมูแผ่น|เนื้อแผ่น|สลัด|โจ๊ก|ซาลาเปา|euro|bakery|bread', n):
        return "อาหารพร้อมทานและเบเกอรี่"
    if re.search(r'เฮอร์ช|ช็อก|chocolate|โอโช|ข้าวอบ|มันฝรั่ง|ป๊อปคอร์|ถั่ว|เยลลี่|'
                 r'ลูกอม|หมากฝรั่ง|สแน็ค|snack|คุกกี้|เวเฟอร์|เลย์|pringle|popcorn|candy', n):
        return "ขนมและของขบเคี้ยว"
    if re.search(r'น้ำ|นม|ชา|กาแฟ|coffee|tea|milk|โค้ก|coke|cola|เป๊ปซี่|pepsi|'
                 r'สไปรท์|sprite|แฟนต้า|คาราบาว|กระทิง|ไมโล|milo|อราวน์|โออิชิ|'
                 r'เกลือแร่|gatorade|เครื่องดื่ม|drink|juice|น้ำผล|น้ำส้ม', n):
        return "เครื่องดื่ม"
    if re.search(r'สบู่|แชมพู|shampoo|ยาสีฟัน|แปรงสีฟัน|ครีมอาบ|โลชั่น|lotion|'
                 r'ผ้าอนามัย|ดีโอ|โรลออน|แป้ง|คอนแทค|ผ้าเช็ด|สกิน|skin|ครีม|serum', n):
        return "ของใช้ส่วนตัว"
    if re.search(r'ทิชชู|tissue|ผงซักฟอก|น้ำยาซัก|น้ำยาปรับ|น้ำยาล้าง|'
                 r'ถุงขยะ|ถุงพลาสติก|ฟิล์มห่อ|cellox|comfort|downy|sunlight|vim|น้ำยา', n):
        return "ของใช้ในบ้าน"
    if re.search(r'ยา|พลาสเตอร์|แอลกอฮอล์|alcohol|หน้ากาก|mask|วิตามิน|vitamin|'
                 r'อาหารเสริม|supplement|ถุงยาง|ผ้าพันแผล|paracetamol|ibuprofen', n):
        return "เวชภัณฑ์และอุปกรณ์ดูแลสุขภาพ"
    return "สินค้าเบ็ดเตล็ดอื่นๆ"

def categorize_items_batch(items: list) -> list:
    if not items:
        return []
    names = [it.get("ชื่อสินค้า", "") for it in items]
    pre_assigned = [_is_bao_item(n) for n in names]
    if not is_gemini_configured():
        return [BAO_CAFE_CATEGORY if bao else "สินค้าเบ็ดเตล็ดอื่นๆ" for bao in pre_assigned]
    pending_indices = [i for i, bao in enumerate(pre_assigned) if not bao]
    pending_names   = [names[i] for i in pending_indices]
    result_cats     = [BAO_CAFE_CATEGORY if bao else "" for bao in pre_assigned]
    if not pending_names:
        return result_cats
    try:
        names_str = "\n".join(f"{i+1}. {n}" for i, n in enumerate(pending_names))
        prompt = f"""จัดหมวดหมู่สินค้า CJ Express:\n{names_str}\n{CATEGORY_PROMPT}\nตอบเป็น JSON array เท่านั้น"""
        raw = _call_gemini(prompt, max_tokens=500).strip()
        raw = re.sub(r"```(?:json)?\s*|\s*```", "", raw).strip()
        cats = json.loads(raw)
        if isinstance(cats, list) and len(cats) == len(pending_names):
            for list_pos, original_idx in enumerate(pending_indices):
                cat = cats[list_pos]
                if _is_bao_item(names[original_idx]):
                    cat = BAO_CAFE_CATEGORY
                result_cats[original_idx] = cat
            return result_cats
    except Exception:
        pass
    return [BAO_CAFE_CATEGORY if bao else "สินค้าเบ็ดเตล็ดอื่นๆ" for bao in pre_assigned]

# ── FIX: ลบ duplicate except block ──
def extract_multi_bills_with_gemini(raw_text: str, n_bills: int) -> list:
    prompt = f"""รูปภาพมี {n_bills} ใบเสร็จเคียงกัน OCR อ่านได้:

{raw_text}

กฎ:
- แต่ละบิลมี BNO: หรือ BNO. เป็นของตัวเอง
- OCR อาจอ่าน "$" แทน "S" ใน BNO
- pos_id: จาก BNO เช่น BNO:S26061416N03 → "1416"
- pos_machine: จาก BNO เช่น N03
- Bao_, Beo, B80, Bac → "Bao" เสมอ
- "ไม่หวาน" "หวานน้อย" คือ note ไม่ใช่สินค้า
- ส่วนลด "-40.00" → item "ส่วนลด" ราคาติดลบ
- โปรโมชั่น "1 แถม 1 Bao" → item "โปรโมชั่น 1 แถม 1 Bao" หมวด "ส่วนลด/โปรโมชั่น" ราคาติดลบ
- วันที่ผิด เช่น "2028" → แก้เป็น "2026"

{BAO_CAFE_MENU}
{CATEGORY_PROMPT}

ตอบเป็น JSON array {n_bills} บิล:
[{{"date":"","time":"","branch":"","pos_id":"","pos_machine":"","rcpt_no":"","total_amount":0.0,
  "items":[{{"ชื่อสินค้า":"","หมวดหมู่":"","จำนวน":1,"ราคาต่อหน่วย":0.0,"ยอดรวมสินค้า":0.0}}]}}]"""
    try:
        raw = _call_gemini(prompt, max_tokens=3000)
        raw = re.sub(r"```(?:json)?\s*|\s*```", "", raw).strip()
        bills_raw = json.loads(raw)
        if not isinstance(bills_raw, list):
            bills_raw = [bills_raw]
        results = []
        for bd in bills_raw[:n_bills]:
            date_str = str(bd.get("date", "ไม่พบ"))
            date_str = re.sub(r'/202[7-9]/', '/2026/', date_str)
            date_str = re.sub(r'/20[3-9]\d/', '/2026/', date_str)
            bill = {
                "date": date_str, "time": str(bd.get("time", "ไม่พบ")),
                "branch": str(bd.get("branch", "ไม่พบ")),
                "pos_id": str(bd.get("pos_id", "ไม่พบ")),
                "pos_machine": str(bd.get("pos_machine", "ไม่พบ")),
                "rcpt_no": str(bd.get("rcpt_no", "ไม่พบ")),
                "total_amount": float(bd.get("total_amount", 0)),
                "cash": 0.0, "change": 0.0, "tax_id": "ไม่พบ", "user": "ไม่พบ", "name": "ไม่พบ",
            }
            items = []
            for it in bd.get("items", []):
                name = str(it.get("ชื่อสินค้า", ""))
                cat  = BAO_CAFE_CATEGORY if _is_bao_item(name) else str(it.get("หมวดหมู่", ""))
                if not cat: cat = _categorize_by_rule(name)
                items.append({
                    "ชื่อสินค้า": name, "หมวดหมู่": cat,
                    "จำนวน": int(it.get("จำนวน", 1)),
                    "ราคาต่อหน่วย": float(it.get("ราคาต่อหน่วย", 0)),
                    "ยอดรวมสินค้า": float(it.get("ยอดรวมสินค้า", 0)),
                })
            results.append({"bill": bill, "items": items})
        while len(results) < n_bills:
            results.append({"bill": {"date":"ไม่พบ","time":"ไม่พบ","branch":"ไม่พบ",
                             "pos_id":"ไม่พบ","pos_machine":"ไม่พบ","rcpt_no":"ไม่พบ",
                             "total_amount":0.0,"cash":0.0,"change":0.0,
                             "tax_id":"ไม่พบ","user":"ไม่พบ","name":"ไม่พบ"}, "items": []})
        return results
    except Exception:
        # fallback: ใช้ extract_with_gemini ปกติ
        single = extract_with_gemini(raw_text, ocr_source="gdrive")
        result = [{"bill": single["bill"], "items": single["items"]}] if single["ok"] else []
        while len(result) < n_bills:
            result.append({"bill": {"date":"ไม่พบ","time":"ไม่พบ","branch":"ไม่พบ",
                             "pos_id":"ไม่พบ","pos_machine":"ไม่พบ","rcpt_no":"ไม่พบ",
                             "total_amount":0.0,"cash":0.0,"change":0.0,
                             "tax_id":"ไม่พบ","user":"ไม่พบ","name":"ไม่พบ"}, "items": []})
        return result

def extract_items_with_gemini(raw_text: str) -> list:
    result = extract_with_gemini(raw_text)
    return result.get("items", [])

def extract_with_gemini(raw_text: str, ocr_source: str = "gdrive") -> dict:
    # ── ดึงจำนวนสินค้ารวมจาก raw text ──
    _n_items_match = re.search(
        r'จ[าำ]?นวนสินค[้า]?\s*รวม\s*(\d+)\s*รายการ',
        raw_text, re.IGNORECASE)
    n_items_hint = ""
    if _n_items_match:
        n = _n_items_match.group(1)
        n_items_hint = f"\n⚠️ บิลนี้มีสินค้า {n} รายการ — ต้องได้ items ครบ {n} ชิ้น ไม่นับโปรโมชั่น/ส่วนลด"

    if ocr_source == "gdrive":
        format_hint = """คุณคือผู้เชี่ยวชาญวิเคราะห์ใบเสร็จ CJ Express
ข้อความนี้มาจาก OCR ลำดับบรรทัดอาจสลับ ราคาอาจอยู่ไม่ติดชื่อสินค้า

กฎการอ่านบรรทัดสินค้า:
- บรรทัดสินค้า: ขึ้นต้นด้วยตัวเลขจำนวน เช่น "1 มาม่า..." หรือ "2 ทีโก..."
- ราคาแยกบรรทัด: บรรทัด 1=ชื่อ, บรรทัด 2=ราคา/หน่วย, บรรทัด 3=ราคารวม V
- ราคา inline: "1 คาราบาว 12.00 12.00 V" → ชื่อ=คาราบาว, unit=12, total=12
- qty>1 inline: "2 คาราบาว 12.00 24.00 V" → จำนวน=2, unit=12, total=24
- "2 Bao_ชาเขียว 40.00 80.00 V" → จำนวน=2, unit=40, total=80
- ถ้ามี N สินค้า ราคา N ตัวถัดมาเรียงตามลำดับ (แม้มี User/BNO คั่น)
- "จำนวนสินค้ารวม N รายการ" → ต้องได้ items ครบ N ชิ้น (ไม่นับโปรโมชั่น)
- "1,285.00" หรือ "1,196.00" บรรทัดเดียว = ยอดรวมสินค้า ไม่ใช่ item
- ราคา 2 ตัวในบรรทัดเดียวเช่น "35.00 89.00" → ราคา item 1 และ item 2 ตามลำดับ
- ราคาที่อยู่ก่อน item name → เป็นราคาของ item ถัดมา
- โปรโมชั่นราคาบวก เช่น 48.00 หลังโปรโมชั่น → ต้องแปลงเป็น -48.00

ห้ามนำสิ่งเหล่านี้มาเป็นชื่อสินค้า (เด็ดขาด):
- ชื่อสาขา, เบอร์โทรศัพท์ (0xx-xxxxxxx), TAX ID, MORE TAX, VAT INCLUDED
- User..., BNO..., POS :Nxx, ID:E..., รหัสสมาชิก
- หวานน้อย, ไม่หวาน, ไปหวาน, หวานปกติ, หวานมาก → note เครื่องดื่ม ไม่ใช่สินค้า
- แต้มสะสม, เงินสด, เงินทอน, ยอดรวม, บอลราม, บอดราม, QR ธนาคาร
- Grab Mart, Order no. GM-xxx → ช่องทางชำระ ไม่ใช่สินค้า
- บาว คาเฟ คิวที่..., ร้องเรียน, ขอบคุณ, ลูกค้าสามารถ, สะสมแต้ม, ร่วมรายการ
- "จำนวนสินค้ารวม N รายการ XX.00" → ยอดรวม ไม่ใช่สินค้า
- ชื่อร้าน/สาขา เช่น "เจ มอสสาขาที่01351..." → ข้ามไป

กรองตัวอักษรขยะ (noise) ออก:
- ชื่อสินค้าต้องมีตัวอักษรไทยหรืออังกฤษที่อ่านออก ความยาว >= 3 ตัว
- ตัวอักษรเดี่ยว เช่น "V", "S", "D", "g", "Ba", "ง" → ข้ามไป
- Noise เช่น "autu", "ne-User", "Ba S" → ข้ามไป
- ถ้าไม่แน่ใจว่าเป็นสินค้าหรือ noise → ข้ามไป อย่าเดา

โปรโมชั่น/ส่วนลด:
- โปรโมชั่นคือส่วนลด → ราคาต้องติดลบเสมอ หมวด "ส่วนลด/โปรโมชั่น"
- Pattern บวก+ลบ: "โปรโมชั่น\n87.00\n-5.00" → ใช้เฉพาะ -5.00 (87.00=ยอดก่อนหัก)
- ห้ามนำ "จำนวนสินค้ารวม XX.00" มาเป็นมูลค่าส่วนลด

Bao Cafe:
- 1B30, lB30, Bao_, Beo, B80, Be0, Bac, BAO → "Bao" เสมอ หมวด "Bao Cafe"
- "Bao_อเมริกาไปเป็น" → "Bao อเมริกาโน่เย็น"
- "+1 วิปปิ้งครีม 15.00" → item เสริม Bao หมวด "Bao Cafe" ราคา +15

OCR แก้ชื่อผิด:
- "สวนลด" → "ส่วนลด", "บอลรวม/บอดรวม" → ยอดรวม ข้ามไป
- ตัวเลขหลังชื่อที่ดูเหมือน serial → ตัดออก"""
    else:
        format_hint = """คุณคือผู้เชี่ยวชาญวิเคราะห์ใบเสร็จ CJ Express
แต่ละสินค้าอยู่บรรทัดเดียว: "จำนวน ชื่อสินค้า ราคา/หน่วย ราคารวม"

กฎการอ่าน:
- "2 คาราบาวแดง 12.00 24.00" → จำนวน=2, unit=12, total=24
- 1B30, lB30, Bao_, Beo, B80 → "Bao" เสมอ หมวด "Bao Cafe"
- "+1 วิปปิ้งครีม 15.00" → addon Bao Cafe ราคา +15
- โปรโมชั่น → ราคาติดลบเสมอ Pattern บวก+ลบ: ใช้เฉพาะราคาลบ
- ห้าม: ชื่อสาขา, เบอร์โทร, User, BNO, หวานน้อย, Grab Mart, แต้มสะสม
- ตัวอักษรขยะ "V","S","autu" → ข้ามไป
- pos_id จาก BNO: "BNO:S26061745N04-xxx" → "1745", pos_machine="N04"
- วันที่ > 2026 → แก้เป็น 2026"""

    # ── Step 1: Header ──
    header_prompt = f"""จากใบเสร็จ CJ Express นี้ สกัดเฉพาะข้อมูล header/footer:
{raw_text}

ข้อมูลที่ต้องการ:
- date: วันที่ (dd/mm/yyyy)
- time: เวลา (HH:MM)
- branch: ชื่อสาขา (เช่น "สาขา 01351 เขื่อนขุนดาน")
- pos_id: รหัสสาขา 4 หลักจาก BNO (เช่น BNO:S26061416N03 → "1416")
- pos_machine: หมายเลข POS จาก BNO (เช่น N03)
- rcpt_no: เลขที่ใบเสร็จจาก BNO (เต็ม เช่น S26061416N03-004232)
- total_amount: ยอดรวมสุทธิ (บาท)
- cash: เงินสด
- change: เงินทอน

ตอบ JSON เท่านั้น ห้าม markdown:
{{"date":"","time":"","branch":"","pos_id":"","pos_machine":"","rcpt_no":"","total_amount":0.0,"cash":0.0,"change":0.0}}"""

    # ── Step 2: Items ──
    items_prompt = f"""จากใบเสร็จ CJ Express นี้ จงสกัดรายการสินค้าทั้งหมดให้ครบทุกรายการ:{n_items_hint}

{raw_text}

{format_hint}

{BAO_CAFE_MENU}
{CJ_PRODUCT_KNOWLEDGE}

{CATEGORY_PROMPT}

กฎสำคัญ:
- Bao_, Beo, B80, Be0, Ba0, Bao. → "Bao" + หมวด "Bao Cafe"
- โปรโมชั่น/ส่วนลด → ราคาติดลบ หมวด "ส่วนลด/โปรโมชั่น"
- "+1 addon" → item เสริม Bao Cafe ราคาบวก
- ถ้าได้ items ไม่ครบ → ตรวจ raw text อีกครั้ง
- ราคา 2 ตัวในบรรทัดเดียว "35.00 89.00" → item 1 และ item 2
- ราคาที่อยู่ก่อน item → เป็นราคาของ item ถัดมา
- จำนวนสินค้ารวม N → ต้องได้ items ครบ N (ไม่นับส่วนลด)
- ห้ามนำ: โฆษณา, คำขอบคุณ, สะสมแต้ม, ร่วมรายการ, ชื่อสาขา, เบอร์โทร มาเป็นสินค้า

ตอบ JSON array เท่านั้น ห้าม markdown:
[{{"ชื่อสินค้า":"","หมวดหมู่":"","จำนวน":1,"ราคาต่อหน่วย":0.0,"ยอดรวมสินค้า":0.0}}]"""

    try:
        # ── เรียก Gemini 2 ครั้ง ──
        header_raw = _call_gemini(header_prompt, max_tokens=300)
        header_raw = re.sub(r"```(?:json)?\s*|\s*```", "", header_raw).strip()
        header_data = json.loads(header_raw)

        items_raw = _call_gemini(items_prompt, max_tokens=2000)
        items_raw = re.sub(r"```(?:json)?\s*|\s*```", "", items_raw).strip()
        items_data = json.loads(items_raw)

        # ── แปลง header ──
        date_str = str(header_data.get("date", "ไม่พบ"))
        date_str = re.sub(r'/202[7-9]/', '/2026/', date_str)
        date_str = re.sub(r'/20[3-9]\d/', '/2026/', date_str)
        bill = {
            "date":         date_str,
            "time":         str(header_data.get("time", "ไม่พบ")),
            "branch":       str(header_data.get("branch", "ไม่พบ")),
            "pos_id":       str(header_data.get("pos_id", "ไม่พบ")),
            "pos_machine":  str(header_data.get("pos_machine", "ไม่พบ")),
            "rcpt_no":      str(header_data.get("rcpt_no", "ไม่พบ")),
            "total_amount": float(header_data.get("total_amount", 0)),
            "cash":         float(header_data.get("cash", 0)),
            "change":       float(header_data.get("change", 0)),
            "tax_id":       "ไม่พบ", "user": "ไม่พบ", "name": "ไม่พบ",
        }

        # ── แปลง items ──
        items = []
        for it in (items_data if isinstance(items_data, list) else []):
            name = str(it.get("ชื่อสินค้า", ""))
            if not name or len(name) < 2: continue
            cat = BAO_CAFE_CATEGORY if _is_bao_item(name) else str(it.get("หมวดหมู่", ""))
            if not cat or cat == "สินค้าเบ็ดเตล็ดอื่นๆ":
                rule_cat = _categorize_by_rule(name)
                if rule_cat != "สินค้าเบ็ดเตล็ดอื่นๆ":
                    cat = rule_cat
            if not cat:
                cat = "สินค้าเบ็ดเตล็ดอื่นๆ"
            items.append({
                "ชื่อสินค้า":    name,
                "หมวดหมู่":     cat,
                "จำนวน":        int(it.get("จำนวน", 1)),
                "ราคาต่อหน่วย": float(it.get("ราคาต่อหน่วย", 0)),
                "ยอดรวมสินค้า": float(it.get("ยอดรวมสินค้า", 0)),
            })

        if not items:
            items = universal_item_parser(raw_text)
        if is_gdrive_token_ready():
            items = enrich_items_from_sheet(items, st.session_state["gdrive_token"])
        return {"bill": bill, "items": items, "ok": True}

    except Exception as e:
        # fallback: single-step เดิม
        try:
            single_prompt = f"""ใบเสร็จ CJ Express:{n_items_hint}
{raw_text}
{format_hint}
{BAO_CAFE_MENU}
ตอบ JSON: {{"date":"","time":"","branch":"","pos_id":"","pos_machine":"","rcpt_no":"","total_amount":0.0,"cash":0.0,"change":0.0,"items":[{{"ชื่อสินค้า":"","หมวดหมู่":"","จำนวน":1,"ราคาต่อหน่วย":0.0,"ยอดรวมสินค้า":0.0}}]}}"""
            text_out = _call_gemini(single_prompt, max_tokens=2000)
            text_out = re.sub(r"```(?:json)?\s*|\s*```", "", text_out).strip()
            data = json.loads(text_out)
            bill = {
                "date": str(data.get("date","ไม่พบ")),
                "time": str(data.get("time","ไม่พบ")),
                "branch": str(data.get("branch","ไม่พบ")),
                "pos_id": str(data.get("pos_id","ไม่พบ")),
                "pos_machine": str(data.get("pos_machine","ไม่พบ")),
                "rcpt_no": str(data.get("rcpt_no","ไม่พบ")),
                "total_amount": float(data.get("total_amount",0)),
                "cash": 0.0, "change": 0.0,
                "tax_id": "ไม่พบ", "user": "ไม่พบ", "name": "ไม่พบ",
            }
            items = []
            for it in data.get("items", []):
                name = str(it.get("ชื่อสินค้า",""))
                if not name: continue
                cat = BAO_CAFE_CATEGORY if _is_bao_item(name) else str(it.get("หมวดหมู่",""))
                if not cat: cat = _categorize_by_rule(name)
                items.append({
                    "ชื่อสินค้า": name, "หมวดหมู่": cat,
                    "จำนวน": int(it.get("จำนวน",1)),
                    "ราคาต่อหน่วย": float(it.get("ราคาต่อหน่วย",0)),
                    "ยอดรวมสินค้า": float(it.get("ยอดรวมสินค้า",0)),
                })
            if not items:
                items = universal_item_parser(raw_text)
            return {"bill": bill, "items": items, "ok": True}
        except Exception as e2:
            return {"bill": None, "items": [], "ok": False, "error": str(e2)}
def _collapse(text: str) -> str:
    return re.sub(r'\s+', '', text)

def _th_spaced(word: str) -> str:
    return r'\s*'.join(re.escape(c) for c in word)

_KW_CANONICAL = {
    "ยอดรวม": ["ยอดรวม","ยอตรวม","ยอดราม","บอดรวม","มอดราม","นลดราม","นอดรวม",
               "รวมสุทธิ","รวมทั้งสิ้น","Total","NET TOTAL","UORTIN","UORT"],
    "เงินสด": ["เงินสด","เง็นสด","เม็นสด","เแงินสด","CASH","QR","ในสต","เงินเด"],
    "เงินทอน": ["เงินทอน","เง็นทอน","เงินทอม","เงินหอน","Change","CHANGE","เงินบน"],
    "สาขา": ["มอร์สาขา","CJมอร์","สาขา","สาขาที่","มอร์สาขาที่","มอร์ลาขา","ซีเจมอร์"],
}

def _build_kw_re(key: str) -> re.Pattern:
    variants = _KW_CANONICAL.get(key, [key])
    return re.compile('|'.join(_th_spaced(v) for v in variants), re.IGNORECASE)

_RE_TOTAL  = _build_kw_re("ยอดรวม")
_RE_CASH   = _build_kw_re("เงินสด")
_RE_CHANGE = _build_kw_re("เงินทอน")
_RE_BRANCH = _build_kw_re("สาขา")
_RE_PRICE  = re.compile(r'(\d{1,6}[.,]\d{2})')
_RE_DATE   = re.compile(r'(\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4})')
_RE_TIME   = re.compile(r'\b([01]\d|2[0-3])[:.:]([0-5]\d)\b')
_RE_TAX_ID = re.compile(r'TAX\s*(?:ID|1D|10|lD|id)?\s*[:\s]?\s*(\d{10,13})', re.IGNORECASE)

def parse_price(s: str) -> float:
    if not s: return 0.0
    s = str(s).strip()
    s = s.replace('O','0').replace('o','0').replace('l','1').replace('I','1')
    s = re.sub(r'^(\d+)\s+(\d{2})$', r'\1.\2', s.strip())
    s = s.replace(',', '.')
    cleaned = re.sub(r'[^\d.]', '', s)
    if not cleaned: return 0.0
    parts = cleaned.split('.')
    if len(parts) > 2: cleaned = parts[-2] + '.' + parts[-1]
    if '.' not in cleaned and len(cleaned) >= 3 and cleaned.endswith('00'):
        cleaned = cleaned[:-2] + '.' + cleaned[-2:]
    try: return float(cleaned)
    except: return 0.0

def _find_prices_in_line(line: str) -> list:
    prices = []
    for m in re.finditer(r'\b(\d{1,6}[.,]\d{2})\b', line):
        prices.append(parse_price(m.group(1)))
    if prices: return prices
    for m in re.finditer(r'\b(\d{1,5})\s+(\d{2})\b', line):
        prices.append(parse_price(f"{m.group(1)}.{m.group(2)}"))
    if prices: return prices
    c = _collapse(line)
    for m in re.finditer(r'\b(\d{3,6})\b', c):
        n = m.group(1)
        if n.endswith('00'): prices.append(parse_price(n))
    return prices

def _fix_date(d: str) -> str:
    parts = re.split(r'[-/.]', d)
    if len(parts) != 3: return d
    day, mon, yr = parts
    if len(yr) == 4:
        if yr[0] not in ('1','2'): yr = '2' + yr[1:]
        if int(yr) > 2500: yr = str(int(yr) - 543)
        yr_int = int(yr) if yr.isdigit() else 0
        if yr_int > 2026:
            yr = '2026'
        elif not (2020 <= yr_int <= 2026):
            candidate = '202' + yr[3]
            if candidate.isdigit() and 2020 <= int(candidate) <= 2026:
                yr = candidate
    elif len(yr) == 2:
        yr = '20' + yr
    if mon == '00': mon = '01'
    if day == '00': day = '01'
    sep = re.search(r'[-/.]', d).group(0)
    return f"{day}{sep}{mon}{sep}{yr}"

def _clean_bno(raw: str) -> str:
    raw = re.sub(r'(?<=[A-Za-z0-9]):(?=\d)', '-', raw)
    result = []
    for i, c in enumerate(raw):
        if c in ('$','€','£'):
            prev = raw[i-1] if i > 0 else ''
            nxt  = raw[i+1] if i < len(raw)-1 else ''
            result.append('5' if (prev.isdigit() and nxt.isdigit()) else 'S')
        elif c == '\u00a7': result.append('S')
        elif c in ("'", "\u2018", "\u2019", "`"): result.append('')
        else: result.append(c)
    return re.sub(r'S{2,}', 'S', ''.join(result))

def _find_rcpt_no(text: str, compact: str) -> str:
    for line in text.split('\n'):
        c = _collapse(line)
        m = re.search(r"(?:BNO|8NO|BN0)[:'\u2018\u2019`\-\.\s]*(.+)$", c, re.IGNORECASE)
        if m:
            raw_tail = m.group(1).strip()
            if raw_tail:
                cleaned = _clean_bno(raw_tail)
                return cleaned if cleaned else raw_tail
    m = re.search(r'([A-Z]\d{7,}[A-Z]\d{2}-\d{4,})', compact)
    if m: return m.group(1)
    for src in [text, compact]:
        m = re.search(r'(?:1D|ID|lD)\s*[:\s€£$]\s*([A-Za-z][A-Za-z0-9]{5,24})', src, re.IGNORECASE)
        if m:
            raw = m.group(1).replace('€','E').replace('£','E').replace('$','5').replace('O','0')
            cut = re.match(r'([A-Z][A-Z0-9]{5,23})', raw)
            if cut: return cut.group(1)
    m = re.search(r'(?:Rcpt|RCPT|Rcopth)[^\d]{0,6}(\d{6,})', compact, re.IGNORECASE)
    if m: return m.group(1)
    for line in text.split('\n'):
        c = _collapse(line)
        if not re.search(r'B[NL]O|8NO|BN0', c, re.IGNORECASE): continue
        all_digits = ''.join(re.findall(r'\d', c))
        n = len(all_digits)
        if n >= 16:
            dd = all_digits[:16]
            return f"S{dd[:8]}N{dd[8:10]}-{dd[10:16]}"
        elif n >= 10:
            return f"S{all_digits[:8]}N{all_digits[8:10]}-{all_digits[10:]}"
    return "ไม่พบ"

def _find_pos_machine_id(text: str, compact: str) -> str:
    for line in text.split('\n'):
        c = _collapse(line)
        m = re.search(r'(?:BNO|8NO|BN0)[:\s.]*[A-Z]\d{8}([A-Z]\d{2,3})-', c, re.IGNORECASE)
        if m: return m.group(1)
    # จาก User line: "Userxxxxx POS :N04" หรือ "POS N04"
    for line in text.split('\n'):
        m = re.search(r'POS\s*[:\s]*([A-Z]\d{2,3})(?:\s|$)', line, re.IGNORECASE)
        if m: return m.group(1)
    m = re.search(r'POS\s*[:\s]+([A-Z]\d{2,3})', compact, re.IGNORECASE)
    if m: return m.group(1)
    return "ไม่พบ"

def _find_pos_id(text: str, compact: str, lines: list) -> str:
    # 1. BNO — รองรับ BNO: และ BNO.
    for line in text.split('\n'):
        c = _collapse(line)
        m = re.search(r'(?:BNO|8NO|BN0)[:\s.]*[A-Z]\d{4}(\d{4})[A-Z]\d+', c, re.IGNORECASE)
        if m: return m.group(1)
    # 2. keyword สาขา — skip TAXID
    _BRANCH_KW = r'(?:สาขา|ลาขา|ฉาขา|ฬาขา|ขัสาชา|สาชาต|ชาขาต)'
    for line in lines[:12]:
        c = _collapse(line)
        if re.search(r'TAXID|TAX\s*ID|0105556', c, re.IGNORECASE): continue
        m = re.search(_BRANCH_KW + r'[^\d]{0,4}(0?\d{4,5})', c, re.IGNORECASE)
        if m: return m.group(1)
    for line in lines[:12]:
        c = _collapse(line)
        if re.search(r'TAXID|TAX\s*ID|0105556', c, re.IGNORECASE): continue
        if re.search(r'มอร์|มอร|เจเอ|CJ|MORE', line, re.IGNORECASE) or 'มอร์' in c:
            nums = re.findall(r'0?\d{4,5}(?!\d)', c)
            if nums: return nums[0]
    m = re.search(r'(?:สาขา|branch)[^\d]{0,5}(\d{4,5})', compact, re.IGNORECASE)
    if m: return m.group(1)
    for line in lines[:10]:
        c = _collapse(line)
        if re.search(r'TAXID|TAX\s*ID|0105556', c, re.IGNORECASE): continue
        m = re.search(r'POS\s*[.\s]*NO\s*S?\s*(\d{1,2})\b', c, re.IGNORECASE)
        if m: return f"NO{m.group(1)}"
        m = re.search(r'POS\s*(?:ID|:)?\s*#?\s*([A-Za-z]\d{1,3})(?!\d)', c, re.IGNORECASE)
        if m: return m.group(1)
    return "ไม่พบ"

def _find_branch(text: str, compact: str, lines: list) -> str:
    cj_pattern = r'(?:ซีเจ|CJ)?ม[อ][ร][ลซ]์?.*?สาขา(?:ที่|เลขที่)?\s*?(\d{2,}(?:\s*\d+)*)'
    bigc_pattern = r'(Big\s*C\s*(?:Mini|Extra)?|BCM|BCH)'
    for line in lines[:12]:
        c = _collapse(line)
        m_cj = re.search(cj_pattern, c, re.IGNORECASE)
        if m_cj:
            return f"สาขา {m_cj.group(1).replace(' ', '')}"
        m_bigc = re.search(bigc_pattern, c, re.IGNORECASE)
        if m_bigc:
            return line.strip()
    m_compact = re.search(r'(?:มอร์|สาขา)\s*?(\d{2,}(?:\s*\d+)*)', compact, re.IGNORECASE)
    if m_compact:
        return f"สาขา {m_compact.group(1).replace(' ', '')}"
    return "ไม่พบ"

_RE_POINTS_LINE = re.compile(r'แต\s*้?\s*ม', re.IGNORECASE)

def _find_amount(text_lines: list, kw_re: re.Pattern, allow_zero: bool = True) -> float:
    for line in text_lines:
        c = _collapse(line)
        if _RE_POINTS_LINE.search(c): continue
        if not (kw_re.search(c) or kw_re.search(line)): continue
        prices = _find_prices_in_line(line)
        if not prices: prices = _find_prices_in_line(c)
        if prices:
            val = prices[-1]
            if val > 0 or allow_zero: return val
    return 0.0

def _find_amounts_positional(lines: list) -> list:
    end_idx = 0
    for idx, line in enumerate(lines):
        if re.search(r'[รง]\s*[ก-๙]{0,2}\s*ย\s*ก\s*า\s*ร', line):
            end_idx = idx + 1
            break
    found = []
    for line in lines[end_idx:]:
        prices = _find_prices_in_line(line) or _find_prices_in_line(_collapse(line))
        if prices:
            found.append(parse_price(prices[-1]))
        if len(found) >= 3:
            break
    return found

def _find_tax_id(compact: str) -> str:
    m = _RE_TAX_ID.search(compact)
    return m.group(1) if m else "ไม่พบ"

def clean_text(text: str) -> str:
    if not text: return ""
    lines = [re.sub(r'[ \t]+', ' ', ln).strip() for ln in text.split('\n')]
    _SPACED = [
        (r'ย\s*อ\s*[ดต]\s*ร\s*ว\s*ม', 'ยอดรวม'),
        (r'น\s*ล\s*ด\s*ร\s*[า5]\s*ม', 'ยอดรวม'),
        (r'บ\s*อ\s*ด\s*ร\s*ว\s*ม',    'ยอดรวม'),
        (r'ม\s*อ\s*ด\s*ร\s*า\s*ม',    'ยอดรวม'),
        (r'เ\s*ง\s*ิ?\s*น\s*ส\s*ด',   'เงินสด'),
        (r'เ\s*ง\s*ิ?\s*น\s*ท\s*อ\s*น', 'เงินทอน'),
        (r'เ\s*ง\s*ิ?\s*น\s*ห\s*อ\s*[นแ]', 'เงินทอน'),
        (r'บ\s*า\s*ท', 'บาท'),
        (r'ม\s*อ\s*ร\s*์\s*[สลซ]\s*า\s*ข\s*า', 'มอร์สาขา'),
        (r'T\s*o\s*t\s*a\s*[l1i]', 'Total'),
        (r'C\s*A\s*S\s*H', 'CASH'),
        (r'C\s*h\s*a\s*n\s*g\s*e', 'Change'),
        (r'V\s*A\s*[TI]', 'VAT'),
        (r'T\s*A\s*X\s*[I1l]\s*D', 'TAXID'),
    ]
    out = []
    for line in lines:
        for pat, rep in _SPACED:
            line = re.sub(pat, rep, line, flags=re.IGNORECASE)
        out.append(line)
    text2 = '\n'.join(out)
    _FIXES = {
        "ยอดราม":"ยอดรวม","ยอตรวม":"ยอดรวม","รวมสุทธิ":"ยอดรวม",
        "เง็นสด":"เงินสด","เม็นสด":"เงินสด",
        "เงินทอม":"เงินทอน","เง็นทอน":"เงินทอน",
        "1D:":"ID:","lD:":"ID:","ID:£":"ID:E","ID:$":"ID:S","ID:€":"ID:E",
        "RECEIPI":"RECEIPT",
    }
    for old, new in _FIXES.items():
        text2 = text2.replace(old, new)
    def fix_date_m(m):
        return _fix_date(m.group(0))
    text2 = _RE_DATE.sub(fix_date_m, text2)
    amt_kw = ["Total","CASH","Change","VAT","เงินทอน","เงินสด","ยอดรวม","บาท","QR"]
    final = []
    for line in text2.split('\n'):
        low = line.lower()
        if any(k.lower() in low for k in amt_kw):
            line = re.sub(r'\b(\d{1,5})\s*[:]\s*(\d{2})\b', r'\1.\2', line)
            line = re.sub(r'\b(\d{1,5})\s+(\d{2})\b(?!\s*\d)', r'\1.\2', line)
            line = re.sub(r'(\d+),(\d{2})\b', r'\1.\2', line)
        final.append(line)
    # normalize OCR digit errors ในบรรทัดที่มีราคา
    price_kw = re.compile(r'\d|[Vv]\s*$|[Oo]{2}|\bl')
    result_lines = []
    for line in final:
        if price_kw.search(line):
            line = re.sub(r'\bl(\d)', r'1\1', line)
            line = re.sub(r'(\d+[.,])[Oo]{2}', r'\g<1>00', line)
            line = re.sub(r'(\d)[Oo](\d)', r'\g<1>0\2', line)
        result_lines.append(line)
    return '\n'.join(result_lines)

def extract_receipt(text: str) -> dict:
    lines   = text.split('\n')
    compact = _collapse(text)
    date_m = _RE_DATE.search(text)
    date_str = _fix_date(date_m.group(1)) if date_m else "ไม่พบ"
    time_val = "ไม่พบ"
    for line in lines:
        m = _RE_TIME.search(line)
        if m:
            time_val = f"{m.group(1)}:{m.group(2)}"
            break
    if time_val == "ไม่พบ":
        for line in lines:
            if _RE_DATE.search(line):
                after = _RE_DATE.sub('', line).strip()
                m = re.search(r'\b(\d{2})\s*[:. ]\s*(\d{2})\b', after)
                if m and 0 <= int(m.group(1)) <= 23 and 0 <= int(m.group(2)) <= 59:
                    time_val = f"{m.group(1)}:{m.group(2)}"
                    break
    branch  = _find_branch(text, compact, lines)
    pos_id  = _find_pos_id(text, compact, lines)
    rcpt_no = _find_rcpt_no(text, compact)
    tax_id  = _find_tax_id(compact)
    user_m = re.search(r'User\s*#?\s*(\w+)', compact, re.IGNORECASE)
    user   = user_m.group(1) if user_m else "ไม่พบ"
    total  = _find_amount(lines, _RE_TOTAL)
    cash   = _find_amount(lines, _RE_CASH)
    change = _find_amount(lines, _RE_CHANGE)
    if total == 0.0:
        _skip_total = re.compile(r'จำนวนสินค้า|จำนวนรายการ|จานวน|รายการ', re.IGNORECASE)
        for line in lines:
            c = _collapse(line)
            is_total_line = (re.search(r'ยอดรวม|นลดราม|มอดราม', c, re.IGNORECASE) and
                             not _skip_total.search(c))
            if is_total_line:
                prices = _find_prices_in_line(line) or _find_prices_in_line(c)
                if prices: total = prices[-1]; break
    if total == 0.0:
        positional = _find_amounts_positional(lines)
        if len(positional) >= 1: total  = positional[0]
        if len(positional) >= 2: cash   = positional[1]
        if len(positional) >= 3: change = positional[2]
    try:
        pos_machine = _find_pos_machine_id(text, compact)
    except Exception:
        pos_machine = "ไม่พบ"
    name = "ไม่พบ"
    skip_name_kw = ["Total","CASH","Change","Vat","Rcpt","POS","TAX","User",
                    "INCLUDED","RECEIPT","ขอบคุณ","ยอดรวม","เงินสด","เงินทอน","สาขา","บาท","ID:","BNO"]
    for line in lines:
        if any(k.lower() in line.lower() for k in skip_name_kw): continue
        m = _RE_PRICE.search(line)
        if m:
            candidate = re.sub(r'^\d+\s*[xXP]?\s*', '', line[:m.start()]).strip()
            if len(candidate) >= 2: name = candidate; break
    return {
        "date": date_str, "time": time_val, "branch": branch, "name": name,
        "total_amount": total, "cash": 0.0, "change": 0.0,
        "pos_id": pos_id, "pos_machine": pos_machine,
        "rcpt_no": rcpt_no, "tax_id": tax_id, "user": user,
    }

def _fix_spaced_price_core(s: str) -> str:
    def _dot4(m):
        n = m.group(0)
        return (n[:-2]+'.'+n[-2:]) if n.endswith('00') else n
    s = re.sub(r'(?<!\d)\d{3,4}(?!\d)', _dot4, s)
    s = re.sub(r'(?<![.\d])(\d+)\s+(\d{2})(?=\s|$)', r'\1.\2', s)
    s = re.sub(r'(?<![.\d])(\d+)\s+(\d)(?=\s|$)', r'\1.\g<2>0', s)
    return s

def _fix_spaced_price(s: str) -> str:
    m = re.match(r'^((?:\d+\s+){1,2})(?=\D)(.*)$', s)
    if m:
        qty_prefix, rest = m.groups()
        return qty_prefix + _fix_spaced_price_core(rest)
    return _fix_spaced_price_core(s)

def _clean_item_name(nm: str) -> str:
    nm = re.sub(r'\s+', ' ', nm).strip()
    nm = re.sub(r'^[!"\'|！,\-\.]+\s*', '', nm)
    nm = re.sub(r'\s*[!"\'|！,\-]+$', '', nm)
    return nm.strip()



def universal_item_parser(text: str) -> list:
    """
    Universal parser รองรับทุก OCR format — 100 test cases
    """
    _RE_PRICE = re.compile(
        r'(\d{1,6}[.,]\d{2})'           # 12.00 / 12,00
        r'|\d+\s*(?:[gG](?:m[lL]?)?|m[lL])\b'  # 12 00
    )
    _RE_NEG   = re.compile(r'^-\s*(\d+[.,]\d{2})\s*[Vv]?\s*$')
    _RE_DATE  = re.compile(r'^\s*\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}\s*(\d{2}:\d{2})?\s*$')
    # _SKIP: เฉพาะ standalone keywords ไม่ใช่ substring
    _SKIP_UP  = re.compile(
        r'(?:^|\s)(?:ยอดรวม|บอลรวม|บอดราม|รวมสุทธิ|รวมทั้งสิ้น|UORTIN|UORT)(?:\s|$)'
        r'|(?:^|\s)(?:เงินสด|เงินทอน|เงินเด|ในสต|เงินบน)(?:\s|$)'
        r'|^BNO|^8NO|^BN0|^TAX|^MORE|^TAXID|^INCLUDED|^MAT\s*INCLUDED|^ID:|^User|^ne-User'
        r'|^ID[-:]?[A-Z]\d{5,}'
        r'|^autu|^auto'
        r'|^POS\s*:|^PCS\s*:'
        r'|(?:QR|OR|0R|QRE|ORE)\s*(?:ธนาคาร|[อา][ลา]?คาร|นาค[\w]*)|Grab\s*Mart|^Order\s*no|^ธนาคาร'
        r'|^สมาชิก|^รหัสสมาชิก|^แต้ม|^เติม|^ขอบคุณ|^ร้องเรียน|^RECEIPT|^INVOICE'
        r'|^บาว\s*คาเฟ|^ลูกค้าสามารถ|^\*แต้ม|^\*\s*แต้ม'
        r'|ใบเสร็จรับเงิน|ใบกำกับภาษี|ใบกากับภาษี'
        r'|MORE\s*TAX|VAT\s*INCLUDED'
        r'|\d+[.,]\d{2}\s*บาท\s+\d+[.,]\d{2}\s*บาท'  # ← "38.00 บาท 100.00 บาท 62.00 บาท"
        r'|^\d+\s+แต้ม\s+\d+\s+แต้ม'
        r'|^\d{1,3}(?:,\d{3})+[.,]\d{2}\s*$'                 # ← "1 แต้ม 311 แต้ม"
        r'|เบอร์ร้าน|เบอร์\s*ร้าน|\d{3}-\d{7,}',
        re.IGNORECASE)
    _KCOUNT   = re.compile(
        r'(?:[สจ].{0,6}|นวน|แวน|แจน|ลาน|ชํา)สินค[้า]?.{0,3}ร[าว][มน]'
        r'|[นจสแลช].{0,5}(?:นค้า|ค้า|สค้า|สนค้า|ค่า|สนค่า)ร[าว][มน]'
        r'|ลานาน.{0,5}ร[าว][มน]'  # ลานาน น าราม
        r'|สานาน.{0,8}ราม' 
        r'|(?:^|\s)-?\s*(?:[สจนแลช]).{0,6}(?:สินค|นค|ค้า|สนค)ร[าว][มน]',
        re.IGNORECASE)
    _PROMO    = re.compile(r'โปรโมชั่น|promotion', re.IGNORECASE)
    _DISC     = re.compile(r'^ส่วนลด\s*$|^สวนลด\s*$', re.IGNORECASE)
    _NOTE = re.compile(r'^(?:หวานน้อย|ลดน้ำตาล|ไม่หวาน|ไปหวาน|หวานปกติ|หวานมาก|extra\s*shot|ใบหราม|\d+[gG](?:ml?)?\s*(?:ร้อน|เย็น)?)\s*$', re.IGNORECASE)

    def _norm_digits(s):
        # comma separator ในราคา เช่น "1,000.00" → "1000.00"
        s = re.sub(r'(\d{1,3}),(\d{3}[.,])', r'\1\2', s)
        s = re.sub(r'\b[lI]([Oo\d])', r'1\1', s)   # l/I + digit/O → 1x
        s = re.sub(r'(\d)[Oo]([.,])', r'\g<1>0\2', s) # 2O. → 20.
        s = re.sub(r'(\d+[.,])[Oo]{2}', r'\g<1>00', s) # 12.OO → 12.00
        s = re.sub(r'(\d+[.,])[Oo]', r'\g<1>0', s)    # 12.O → 12.0
        s = re.sub(r'(\d)[Oo](\d)', r'\g<1>0\2', s)   # 1O2 → 102
        s = re.sub(r'([.,]\d*)[Oo]', r'\g<1>0', s)
        return s

    def _pp(s):
        if not s: return 0.0
        s = _norm_digits(str(s)).replace(',','.')
        s = re.sub(r'[^\d.]','',s)
        if s.count('.') > 1: s = s[:s.rfind('.',-4)] + '.' + s[-2:]
        try: return float(s)
        except: return 0.0

    def _find_prices(s):
        """หาราคาจาก string รองรับหลาย format"""
        s = _norm_digits(s)
        results = []
        # 1. 12.00 / 12,00 (standard)
        for m in re.finditer(r'(\d{1,6}[.,]\d{2})', s):
            results.append(_pp(m.group(1)))
        if results: return results
        # 2. 12 00 (space แทน dot เช่น "12 00")
        for m in re.finditer(r'(?<![.\dก-๙A-Za-z])(\d{1,5})\s+(\d{2})(?=\s|[Vv]|$)', s):
            results.append(_pp(f"{m.group(1)}.{m.group(2)}"))
        if results: return results
        # 3. "6 6 V" (same number × 2, qty=1)
        for m in re.finditer(r'(?<![.\d])(\d{1,4})\s+\1(?=\s|[Vv]|$)', s):
            results.append(float(m.group(1)))
        if results: return results
        # 4. ตัวเลขเดียวที่ดูเหมือนราคา (≥5 บาท ไม่ใช่ qty)
        # เฉพาะเมื่อ line มี Thai text และตัวเลขน้อยกว่า 3 ตัว
        nums = re.findall(r'\b(\d{2,5})\b', re.sub(r'[Vv]\s*$','',s).strip())
        if nums and len(nums) == 1 and not re.search(r'[ก-๙A-Za-z]',s):
            try:
                v = float(nums[0])
                if 5 <= v <= 9999: results.append(v)
            except: pass
        return results

    def _clean(s):
        s = re.sub(r'\s*[สจ].{0,6}สินค[้า]?.{0,3}ร[าว][มน].*รายการ\s*$','',s,flags=re.IGNORECASE)
        s = re.sub(r'\s+(?:หวานน้อย|ลดน้ำตาล|ไม่หวาน|หวานปกติ|extra\s*shot)\s*$','',s,flags=re.IGNORECASE)
        s = re.sub(r'\s+[Vv]\s*$','',s)
        s = re.sub(r'\s+\d+[.,]\d{2}\s*$','',s)
        s = re.sub(r'\s+-\s*$','',s)
        s = re.sub(r'^[|!>"\x60\[\]]+\s*', '', s)
        s = re.sub(r'^\d+\s+','',s)  # ตัด qty
        return s.strip()

    def _nq(s): return len(re.sub(r'[^\u0E00-\u0E7FA-Za-z0-9]','',s))

    def _is_skip(line):
        s = line.strip()
        lc = re.sub(r'\s+','',s)
        return bool(
            _SKIP_UP.search(s) or
            _KCOUNT.search(lc) or
            _RE_DATE.match(s) or
            _NOTE.match(s) or
            re.search(r'แต้มสะสม|แต้มหมด|แต้มครั้ง|เติมสะสม|เติมครั้ง',s) or
            re.match(r'^\d+(?:แต้ม|เติม|แต่ม)',s) or
            re.match(r'^\d+\s+แต้ม\s+\d+\s+แต้ม', s) or   # ← เพิ่ม "1 แต้ม 311 แต้ม"
            re.match(r'^\d+[.,]\d{2}\s*บาท', s) or          # ← เพิ่ม "38.00 บาท..."
            re.search(r'(?:ne-)?User\w+.*(?:POS|PCS)|BNO[:\s.]',s,re.IGNORECASE) or  # User+BNO ปนกลาง
            re.match(r'^(?:QR|OR|0R|Grab)\s*$',s,re.IGNORECASE) or   # payment keyword เดี่ยว
            re.match(r'^\d+[.,]\d{2}\s*บาท\s*$',s) or              # "40.00 บาท" → ยอดชำระ
            re.match(r'^\d{5,}$',lc)
        )

    # normalize
    lines = []
    for l in text.split('\n'):
        l = re.sub(r'\bสวนลด\b','ส่วนลด',l.strip())
        l = re.sub(r'\bBag(?=[_\s])','Bao',l,flags=re.IGNORECASE)  # Bag → Bao
        l = re.sub(r'\bBAO(?=[_\s])','Bao',l)  # BAO → Bao
        l = re.sub(r'\bจานวนสินค[้า]?รวม','จำนวนสินค้ารวม',l)
        l = re.sub(r'\s*[สจ].{0,6}สินค[้า]?.{0,3}ร[าว][มน].*รายการ\s*$','',l,flags=re.IGNORECASE).strip()
        l = _norm_digits(l)
        lines.append(l)

    items = []
    used  = set()
    i = 0

    while i < len(lines):
        if i in used: i+=1; continue
        line = lines[i]
        if not line: i+=1; continue

        # ตัด "ยอดรวม..." inline ออกก่อน (F59: "1 น้ำดื่ม 7.00 V ยอดรวม 7.00")
        line_p = re.sub(r'\s*(?:ยอดรวม|บอลรวม|บอดราม|รวมทั้งสิ้น)[^\n]*$', '', line, flags=re.IGNORECASE).strip() or line
        skip = _is_skip(line_p)

        # ── ราคาลบ standalone
        m_neg = _RE_NEG.match(line)
        if m_neg:
            p = _pp(m_neg.group(1))
            if items and items[-1]["ยอดรวมสินค้า"] == 0.0:
                items[-1]["ราคาต่อหน่วย"] = -p
                items[-1]["ยอดรวมสินค้า"] = -p
            elif p > 0:
                items.append({"ชื่อสินค้า":"ส่วนลด","จำนวน":1,"ราคาต่อหน่วย":-p,"ยอดรวมสินค้า":-p})
            used.add(i); i+=1; continue

        # ── inline ลบ เช่น "โปรโมชั่น -4.00"
        m_neg_inline = re.search(r'-\s*(\d+[.,]\d{2})\s*[Vv]?\s*$', line)
        if m_neg_inline and _PROMO.search(line):
            nm = re.sub(r'\s*-\s*\d+[.,]\d{2}\s*[Vv]?\s*$','',line).strip()
            nm = _clean(nm)
            p  = _pp(m_neg_inline.group(1))
            if _nq(nm) >= 2:
                items.append({"ชื่อสินค้า":nm,"จำนวน":1,"ราคาต่อหน่วย":-p,"ยอดรวมสินค้า":-p})
                used.add(i); i+=1; continue

        if skip: i+=1; continue

        # ── "ส่วนลด" บรรทัดเดียว
        if _DISC.match(line):
            items.append({"ชื่อสินค้า":"ส่วนลด","จำนวน":1,"ราคาต่อหน่วย":0.0,"ยอดรวมสินค้า":0.0})
            used.add(i); i+=1; continue

        has_th = bool(re.search(r'[ก-๙A-Za-z]', line_p))
        pline  = _find_prices(line_p)
        is_p   = bool(_PROMO.search(line))
        # qty: ตัวเลขนำหน้า (ไม่ใช่ 0 หรือราคา)
        qty_m  = re.match(r'^[+\-]?\s*([1-9]\d{0,1})\s+', line_p)
        qty    = int(qty_m.group(1)) if qty_m else 1

        # Format B: ชื่อ+ราคาในบรรทัดเดียว
        is_qty_zero = bool(re.match(r'^\s*0\s+', line))
        if has_th and len(pline) >= 1 and not is_qty_zero:
            nm = re.sub(r'\s+\d{1,6}[.,]\d{2}(?:\s+\d{1,6}[.,]\d{2})*\s*[Vv \-]*$','',line_p).strip()
            nm = _clean(nm)
            # ป้องกัน: ชื่อต้องมี alphanumeric >=2 และไม่ใช่แค่ตัวเลข
            if _nq(nm) >= 2 and not re.match(r'^\d+\.?\d*$',nm):
                pv = pline; unit=pv[0]; total=pv[-1]
                if total==0 or total<unit*0.3: total=unit*qty
                if is_p: unit=-abs(unit); total=-abs(total)
                if total >= 5.0 or total < 0:  # filter ราคา 0 และ < 5 (ตัวเลขในชื่อ)
                    items.append({"ชื่อสินค้า":nm,"จำนวน":qty,"ราคาต่อหน่วย":unit,"ยอดรวมสินค้า":total})
                    used.add(i); i+=1; continue
                # total < 5 → อาจเป็นตัวเลขในชื่อ ไม่ใช่ราคา → ตกไป Format A

        # Format A/C: look-ahead
        if has_th:
            nc = _clean(line)
            if _nq(nc) < 2 or re.match(r'^\d+\.?\d*$',nc): i+=1; continue
            pf = []
            j  = i+1
            while j < min(i+15, len(lines)) and len(pf) < 2:
                if j in used: j+=1; continue
                nxt = lines[j].strip()
                if not nxt: j+=1; continue
                if _NOTE.match(nxt): j+=1; continue
                m_n2 = _RE_NEG.match(nxt)
                if m_n2:
                    if is_p:
                        pf.append('-'+m_n2.group(1)); used.add(j)
                    break
                ps = _find_prices(nxt)
                if ps and not _is_skip(nxt):
                    pf.extend(ps[:2-len(pf)]); used.add(j)
                elif _is_skip(nxt):
                    pass  # ข้าม skip แต่ look-ahead ต่อ
                j+=1
            if pf:
                pv=[_pp(p.lstrip('-'))*(-1 if str(p).startswith('-') else 1) for p in map(str,pf)]
                unit=pv[0]; total=pv[-1]
                if total>0 and total<unit*0.3: total=unit*qty
                if is_p and total > 0: unit=-abs(unit); total=-abs(total)
                elif is_p and total == 0:
                    m_pamt=re.search(r'(?:ราคาพิเศษ|ลด)\s*(\d+)',nc,re.IGNORECASE)
                    if m_pamt: total=-float(m_pamt.group(1)); unit=total
                nc = re.sub(r'^\d+\s+','',nc).strip()
                if total >= 5.0 or total < 0:  # filter ราคา 0 และ < 5
                    items.append({"ชื่อสินค้า":nc,"จำนวน":qty,"ราคาต่อหน่วย":unit,"ยอดรวมสินค้า":total})
                used.add(i)
            elif not pf:
                # backward scan: ราคาอาจอยู่ก่อน item line
                for bj in range(max(0,i-5), i):
                    if bj in used: continue
                    bp = _find_prices(lines[bj].strip())
                    if bp and not _is_skip(lines[bj]):
                        bunit=bp[0]; btotal=bp[-1]
                        if btotal<bunit*0.3: btotal=bunit*qty
                        nc2=re.sub(r'^\d+\s+','',nc).strip()
                        items.append({"ชื่อสินค้า":nc2,"จำนวน":qty,"ราคาต่อหน่วย":bunit,"ยอดรวมสินค้า":btotal})
                        used.add(i); used.add(bj)
                        break
        i+=1

    # หมวดหมู่
    for it in items:
        it["หมวดหมู่"] = _categorize_by_rule(it["ชื่อสินค้า"])
    return items
_multi_item = re.compile(
    r'^(\d+[.\s]+[ก-๙A-Za-z][ก-๙\s]{2,}?)\s+(\d+[.\s]*[ก-๙A-Za-z].{2,}?)(?:\s+จ.{0,15}รวม.*)?\s*$'
)

def _split_multi_items(lines: list) -> list:
    # ── FIX: loop จนไม่มีอะไรเปลี่ยน (รองรับ 5+ items ในบรรทัดเดียว) ──
    prev = None
    while prev != lines:
        prev = lines[:]
        out = []
        for l in lines:
            s = l.strip()
            s_clean = re.sub(r'\s*จ[าำ]?.{0,6}(?:นค้า|ค่า|สินค้า)?รวม.*$', '', s, flags=re.IGNORECASE).strip()
            if not re.search(r'^\d+\s+[ก-๙A-Za-z].*\s+\d+\s+[ก-๙A-Za-z]', s_clean):
                out.append(l); continue
            first_part = s_clean[:len(s_clean)//2]
            if re.search(r'\d+[.,]\d{2}', first_part):
                out.append(l); continue
            m = _multi_item.match(s_clean)
            if m:
                p1 = m.group(1).strip()
                p2 = m.group(2).strip()
                if len(re.findall(r'[ก-๙]', p1)) >= 2 and len(p2) >= 4:
                    out.append(p1)
                    out.append(p2)
                    continue
            out.append(l)
        lines = out
    return lines


def _merge_gdrive_lines(lines: list) -> list:
    _price_only = re.compile(r'^[\d,.\s]+[Vv]?\s*$')
    _v_only     = re.compile(r'^[Vv]\s*$')
    _has_thai   = re.compile(r'[ก-๙]')
    _item_start = re.compile(r'^(?:\d+[.\s]?|-)\s*\S')
    _kw_amount  = re.compile(
        r'ยอดรวม|ยอดราม|ยอดราเม|UORTIN|UORT|เงินสด|เงินเด|ในสต|เงินทอน|เงินบน|รวมทั้งสิ้น|บอดราม|บอลราม',
        re.IGNORECASE)
    _skip_note  = re.compile(r'^(?:หวานน้อย|ลดน้ำตาล|ไม่หวาน|ไปหวาน|หวานปกติ|หวานมาก|extra\s*shot|ใบหราม|\d+[gG](?:ml?)?\s*(?:ร้อน|เย็น)?)\s*$', re.IGNORECASE)
    _kw_count   = re.compile(
        r'จ[าำํ]?นวนสินค[้า]?[่า]?\s*รวม'
        r'|^[สจ].{0,4}สินค[้า]?.{0,3}ร[าว][มน]'
        r'|จ[าำํ]?นวนสินค[้า]?รวมรายการ',
        re.IGNORECASE)
    _re_price   = re.compile(r'(\d{1,6}[.,]\d{2})')

    lines = [l for l in lines if not _v_only.match(l.strip())]
    lines = _split_multi_items(lines)

    _normalized = []
    for l in lines:
        l = re.sub(r'\bสวนลด\b', 'ส่วนลด', l)
        l = re.sub(r'\bจานวนสินค[้า]?รวม', 'จำนวนสินค้ารวม', l)
        l = re.sub(r'(\d+),(\d{2})\b', r'\1.\2', l)   # 10,00 → 10.00
        l = re.sub(r'^\s*\.(\d)', r'\1', l)             # .69.00 → 69.00
        _normalized.append(l)
    lines = _normalized
    _hdr_kw = re.compile(
    r'BNO|TAX|VAT|POS|User|ID:|MORE|มอร์|สาขา|เบอร์|ซีเจ|CJ',
    re.IGNORECASE)
    # ── ประกาศ nested functions ก่อนเรียกใช้ ──

    def _merge_kw_price(lines):
        expanded = []
        for line in lines:
            s = line.strip()
            if not _has_thai.search(s) and not re.search(r'[Vv]\s*$', s):
                prices = re.findall(r'\d{1,3}(?:,\d{3})*[.,]\d{2}', s)
                rebuilt = re.sub(r'\d{1,3}(?:,\d{3})*[.,]\d{2}', '', s).strip()
                if len(prices) >= 2 and not rebuilt:
                    for p in prices:
                        expanded.append(p)
                    continue
            expanded.append(line)
        lines = expanded
        out = []
        i = 0
        _hdr_pattern = re.compile(
            r'(?:ซีเจ|เจ|CJ)\s*มอ[รส]?สาขา|สาขาที่\s*\d|เบอร์ร้าน'
            r'|MORE\s*TAX|VAT\s*INCLUDED|ใบเสร็จ|ใบกากับ',
            re.IGNORECASE)
        while i < len(lines):
            line = lines[i].strip()
            line = re.sub(
                r'\s*[สจ].{0,4}สินค[้า]?.{0,3}ร[าว][มน].*รายการ\s*$', '',
                line, flags=re.IGNORECASE).strip()
            line = re.sub(r'สา[นม]านสินค[้า]?ราม\S*', '', line).strip()
            line = re.sub(
                r'\s+(?:หวานน้อย|ลดน้ำตาล|ไม่หวาน|หวานปกติ|extra\s*shot)\s*$',
                '', line, flags=re.IGNORECASE).strip()

            if _kw_amount.search(line) and not _re_price.search(line):
                j = i + 1
                if j < len(lines):
                    nx = lines[j].strip()
                    if _re_price.search(nx) and not _has_thai.search(nx):
                        # ← เพิ่ม: next line เป็นราคาล้วน → merge
                        out.append(line + " " + nx)
                        i = j + 1
                        continue
                # ← FIX: ไม่ว่า next line จะเป็นอะไร ให้ยังคง append บรรทัดนี้
                # เพื่อให้ state machine จับ _kw_am_sm ได้
                if line:
                    out.append(line)
                i += 1
                continue  # ← เพิ่ม: ข้าม else branch

            if line:
                out.append(line)
            i += 1
        return out

    def _find_price_block(lines):
        consec = 0
        block_start = None
        for i, line in enumerate(lines):
            s = line.strip()
            if _kw_amount.search(s) and not _re_price.search(s):
                return None
            if _price_only.match(s) and s:
                if consec == 0:
                    block_start = i
                consec += 1
                if consec >= 2:
                    name_only_count = sum(
                        1 for l in lines[:block_start]
                        if _item_start.match(l.strip())
                        and _has_thai.search(l)
                        and not _re_price.search(l)
                        and not _hdr_kw.search(l)        # ← เพิ่มกรอง header
                    )
                    # ── FIX: นับเฉพาะ item ที่มีราคา inline (ไม่ใช่ GDrive format) ──
                    item_with_inline_price = sum(
                        1 for l in lines[i+1:i+10]
                        if _item_start.match(l.strip())
                        and _has_thai.search(l)
                        and not _price_only.match(l.strip())
                        and not _kw_count.search(l)
                        and not _kw_amount.search(l)
                        and not re.search(r'โปรโมชั่น|promotion', l, re.IGNORECASE)
                        and _re_price.search(l)           # ← เฉพาะ inline price
                    )
                    if name_only_count >= 1 and item_with_inline_price == 0:
                        return block_start
            else:
                consec = 0
                block_start = None
        return None

    def _reorder_inline_v(lines):
        """แก้ format: item1+price, item2+price, V1, V2 → item1+price, V1, item2+price, V2"""
        out = []
        i = 0
        while i < len(lines):
            if (_item_start.match(lines[i].strip()) and
                    _has_thai.search(lines[i]) and
                    re.search(r'\d+[.,]\d{2}', lines[i])):
                item_block = []
                j = i
                while (j < len(lines) and
                       _item_start.match(lines[j].strip()) and
                       _has_thai.search(lines[j]) and
                       re.search(r'\d+[.,]\d{2}', lines[j])):
                    item_block.append(lines[j])
                    j += 1
                v_block = []
                while (j < len(lines) and
                       _price_only.match(lines[j].strip()) and
                       not _has_thai.search(lines[j])):
                    v_block.append(lines[j])
                    j += 1
                if len(v_block) >= len(item_block) and len(item_block) > 1:
                    v_per = len(v_block) // len(item_block)
                    for k, item in enumerate(item_block):
                        out.append(item)
                        for vl in v_block[k*v_per:(k+1)*v_per]:
                            out.append(vl)
                    for vl in v_block[len(item_block)*v_per:]:
                        out.append(vl)
                    i = j
                    continue
            out.append(lines[i])
            i += 1
        return out

    # ── เรียกใช้ตามลำดับ ──
    lines = _reorder_inline_v(lines)
    lines = _merge_kw_price(lines)
    lines = [re.sub(r'\b(\d{1,2}):(\d{2})\b',
                    lambda m: f"{m.group(1)}.{m.group(2)}", l) for l in lines]

    price_block_start = _find_price_block(lines)

    if price_block_start and price_block_start > 0:
        _neg_price = re.compile(r'^-\s*[\d,.]+\s*[Vv]?\s*$')
        _DATE_RE3  = re.compile(r'\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}')
        _skip_hdr  = re.compile(
            r'BNO|8NO|TAX|VAT|POS|User|ID:|TAXID'
            r'|บอลราม|บอดราม|ยอดรวม|บอลรวม|UORTIN|UORT'
            r'|เงินสด|เงินทอน|เงินเด|ในสต|เงินบน|รวมทั้งสิ้น'
            r'|(?:ซีเจ|เจ|CJ)\s*(?:มอ[รส]?(?:ร์)?)\s*(?:สาขา|ลาขา|ฉาขา)'  # ← แทนที่เดิม
            r'|(?:มอ[รส]?(?:ร์)?)\s*(?:สาขา|ลาขา)'                          # ← แทนที่เดิม
            r'|มอร์สาขา|มอสสาขา|มอร์ลาขา|ซีเจ\s*มอ|CJ\s*มอ'  # ← เพิ่ม
            r'|สาขาที่\d|สาขา\s*\d|เขื่อน|ปราการ|นครนายก|องครักษ์'  # ← เพิ่ม
            r'|เบอร์ร้าน|เบอร์\s*ร้าน|\d{3}-\d{7}', 
            re.IGNORECASE)

        def _strip_note(s):
            return re.sub(
                r'\s+(?:หวานน้อย|ลดน้ำตาล|ไม่หวาน|หวานปกติ|extra\s*shot)\s*$',
                '', s, flags=re.IGNORECASE).strip()

        items_merged = []
        header_lines = []
        cur_name = None
        prices_buf = []
        in_items = False

        _kw_am_sm = re.compile(
            r'ยอดรวม|บอดราม|UORTIN|UORT|เงินสด|เงินเด|ในสต|เงินทอน|เงินบน|รวมทั้งสิ้น|บอลรวม|บอลราม|QR\s*ธนาคาร',
            re.IGNORECASE)

        for l in lines:
            s = l.strip()
            if not s: continue

            if _kw_am_sm.search(s):
                if cur_name and prices_buf:
                    items_merged.append(f"{cur_name} {prices_buf[-1]}")
                elif cur_name and not prices_buf:
    # นับ items ที่รอราคา (รวม cur_name)
                    pending = [x for x in items_merged if not re.search(r'\d+[.,]\d{2}', x)]
                    pending_count = len(pending) + 1

                    _lookahead_prices = []
                    _cur_idx = lines.index(l) if l in lines else -1
                    if _cur_idx >= 0:
                        for _j in range(_cur_idx + 1, min(_cur_idx + 25, len(lines))):
                            _ls = lines[_j].strip()
                            _lc = re.sub(r'\s*[Vv]\s*$', '', _ls).strip()
                            if not _lc: continue
                            if _skip_note.match(_lc): continue
                            if _kw_count.search(_lc): continue
                            if re.search(r'[ก-๙]', _lc) and not _price_only.match(_lc):
                                if re.match(r'^\d+\s+[ก-๙A-Za-z].{2,}', _lc):
                                    break
                                elif re.search(r'รหัส|แคม|แสน|นนทม|CR\s|LW$|บท$|บล่ม|บ\./น|User|BNO|POS', _lc, re.IGNORECASE):
                                    continue
                                elif len(re.sub(r'\s+', '', _lc)) <= 15:
                                    continue
                                else:
                                    break
                            if _price_only.match(_lc) and _lc and not re.search(r'[ก-๙]', _lc):
                                if not _lookahead_prices or _lc != _lookahead_prices[-1]:
                                    _lookahead_prices.append(_lc)
                                if len(_lookahead_prices) >= pending_count:
                                    break

                    if _lookahead_prices:
                        price_idx = 0
                        for mi, mitem in enumerate(items_merged):
                            if not re.search(r'\d+[.,]\d{2}', mitem) and price_idx < len(_lookahead_prices):
                                items_merged[mi] = f"{mitem} {_lookahead_prices[price_idx]} {_lookahead_prices[price_idx]}"
                                price_idx += 1
                        if price_idx < len(_lookahead_prices):
                            items_merged.append(f"{cur_name} {_lookahead_prices[price_idx]} {_lookahead_prices[price_idx]}")
                        else:
                            items_merged.append(cur_name)
                    else:
                        items_merged.append(cur_name)
                cur_name = None; prices_buf = []
                break

            clean_p = re.sub(r'\s*[Vv]\s*$', '', s).strip()
            is_neg  = bool(_neg_price.match(clean_p)) and bool(clean_p)
            is_pos  = bool(_price_only.match(clean_p)) and bool(clean_p) and not is_neg

            if is_neg:
                if cur_name and prices_buf:
                    items_merged.append(f"{cur_name} {prices_buf[-1]}")
                    cur_name = None; prices_buf = []
                neg_val = re.sub(r'^-\s*', '', clean_p).strip()
                items_merged.append(f"1 ส่วนลด -{neg_val}")
                in_items = True
            elif is_pos:
                prices_buf.append(clean_p)
                if len(prices_buf) == 2 and cur_name:
                    items_merged.append(f"{cur_name} {prices_buf[0]} {prices_buf[1]}")
                    cur_name = None; prices_buf = []
                in_items = True
            elif not _DATE_RE3.match(s) and (
                    _item_start.match(s) or
                    (_has_thai.search(s) and not _skip_hdr.search(s))):

                if cur_name and prices_buf:
                    items_merged.append(f"{cur_name} {prices_buf[-1]}")
                elif cur_name and not prices_buf:
                    # ── FIX: item ไม่มีราคา และ next item เป็นโปรโมชั่น
                    # → look-ahead ข้ามโปรโมชั่นไปหาราคาของ item นี้ก่อน ──
                    if re.search(r'โปรโมชั่น|promotion', s, re.IGNORECASE):
                        # s คือโปรโมชั่น, cur_name ยังไม่มีราคา
                        # หาราคาของ cur_name จาก prices ก่อนโปรโมชั่นจะดูด
                        _cur_line_idx = lines.index(l) if l in lines else -1
                        _found_price = None
                        for _fj in range(_cur_line_idx + 1, min(_cur_line_idx + 5, len(lines))):
                            _fl = lines[_fj].strip()
                            _flc = re.sub(r'\s*[Vv]\s*$', '', _fl).strip()
                            if _price_only.match(_flc) and not re.search(r'[ก-๙]', _flc) and not _neg_price.match(_flc):
                                _found_price = _flc
                                break
                            elif _has_thai.search(_fl):
                                break
                        if _found_price:
                            items_merged.append(f"{cur_name} {_found_price} {_found_price}")
                        else:
                            items_merged.append(cur_name)
                    else:
                        items_merged.append(cur_name)

                cur_name = None; prices_buf = []
                if re.search(r'โปรโมชั่น|promotion', s, re.IGNORECASE):
                    items_merged.append(s)
                    in_items = True
                elif re.match(r'^ส่วนลด\s*$', s):
                    in_items = True
                elif not in_items and _skip_hdr.search(s):
                    header_lines.append(s)
                else:
                    stripped = _strip_note(s)
                    if stripped:
                        cur_name = stripped
                        in_items = True
            else:
                if not in_items:
                    header_lines.append(s)

        if cur_name and prices_buf:
            items_merged.append(f"{cur_name} {prices_buf[-1]}")
        elif cur_name:
            items_merged.append(cur_name)

        return header_lines + items_merged

    else:
        merged = []
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if not line:
                merged.append(line); i += 1; continue
            if _item_start.match(line) and _has_thai.search(line):
                combined = line
                j = i + 1
                price_count = 0

                has_inline_price = bool(re.search(r'\d+[.,]\d{2}', line))
                if has_inline_price and j < len(lines):
                    nx = lines[j].strip()
                    if _item_start.match(nx) and _has_thai.search(nx):
                        merged.append(combined)
                        inline_vals = re.findall(r'(\d+[.,]\d{2})', line)
                        skip_j = j
                        while skip_j < len(lines):
                            sv = re.sub(r'\s*[Vv]\s*$', '', lines[skip_j].strip()).strip()
                            if sv in inline_vals:
                                skip_j += 1
                            else:
                                break
                        i = skip_j
                        continue
                    nxc = re.sub(r'\s*[Vv]\s*$', '', nx).strip()
                    inline_prices = re.findall(r'\d+[.,]\d{2}', line)
                    if inline_prices and nxc in inline_prices:
                        combined += ' ' + nxc
                        merged.append(combined)
                        j += 1
                        i = j
                        continue

                while j < len(lines) and price_count < 2:
                    nx = lines[j].strip()
                    if not nx: break
                    if _skip_note.match(nx):
                        j += 1; continue
                    if _price_only.match(nx):
                        part = re.sub(r'\s*[Vv]\s*$', '', nx).strip()
                        part = re.sub(r'^-\s*', '', part).strip()
                        if part: combined += " " + part
                        price_count += 1
                        j += 1
                        while j < len(lines) and _skip_note.match(lines[j].strip()):
                            j += 1
                    elif re.match(r'^-\s*\d+[.,]\d{2}\s*[Vv]?\s*$', nx):
                        part = re.sub(r'^-\s*', '', nx)
                        part = re.sub(r'\s*[Vv]\s*$', '', part).strip()
                        if part: combined += " " + part
                        price_count += 1
                        j += 1
                    elif price_count == 0 and _has_thai.search(nx) and not _kw_amount.search(nx):
                        if _item_start.match(nx):
                            merged.append(combined)
                            combined = nx
                            j += 1
                        else:
                            combined += " " + nx
                            j += 1
                    else:
                        break
                if price_count > 0:
                    merged.append(combined); i = j; continue
            if _skip_note.match(line):
                i += 1; continue
            merged.append(line); i += 1
        return merged
       
# ── FIX: extract_items_cj ครบทุกจุด ──
def extract_items_cj(text: str) -> list:
    items  = []
    lines  = text.split('\n')
    lines = _split_multi_items(lines) 
    lines = _merge_gdrive_lines(lines)

    start_idx = 0
    _HEADER_BLOCK = re.compile(
    r'สาขา|MORE\s*TAX|VAT\s*INCLUDED|ใบเสร็จ|ใบกากับ|เบอร์ร้าน'
    r'|^ID[-:]?[A-Z]\d{5,}|Userchanjira|User\w+\s*POS',
    re.IGNORECASE
    )
    _DATE_RE2 = re.compile(r'\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}')
    for idx, line in enumerate(lines):
            if _DATE_RE2.search(line) or _DATE_RE2.search(_collapse(line)):
                start_idx = idx + 1; break

    stop_kw = ["ยอดรวม","ยอดราเม","ยอดราม","บอดราม","UORTIN","UORT",
               "รวมทั้งสิ้น","เงินสด","เงินเด","ในสต","เงินทอน","เงินบน",
               "QR ธนาคาร","QRธนาคาร",
               "จำนวนสินค้า","จำนวนรายการ","จานวนสินค้า","จํานวนสินค้า",
               "จํานวนสินค้ารวม","จำนวนสินค้ารวม","จานวนสินค่ารวม"]

    # ── FIX: ลบ "โปรโม" และ "ส่วนลด" ออกจาก skip_kw เพื่อให้แสดงโปรโมชั่น ──
    skip_kw = ["BNO","8NO","POS","TAX","INCLUDED","ใบเสร็จ",
                "แต้ม","แต่ม","ขอบคุณ","สาขา","RECEIPT","INVOICE",
                "สมาชิก","ID:","หวานน้อย","ลดน้ำตาล",
                "www.","FB:","ร้องเรียน","สมัคร",  # ← เพิ่ม comma ตรงนี้
                "เบอร์ร้าน","เบอร์โทร","MORE TAX","VAT INCLUDED",
                "ใบกากับ","ใบกำกับ","ใบเสร็จรับ","บาท 100","บาท 62",   # ยอดชำระ patterns
                "ID:E", "บอลราม", "บอดราม", "บอลรวม","ID-E"]

    _SUFFIX = r'[\sA-Za-z\u0E00-\u0E7F"\u201c\u201d|!！Vv]*'
    _QTY    = r'(?:\d+|-)'
    _full = re.compile(r'^[\.]?\s*(' + _QTY + r')\s*(.+?)\s+(\d+[.,]\d{2})\s+(\d+[.,]\d{2})' + _SUFFIX + r'$')
    _fb_a  = re.compile(r'^(.+?)\s+(\d+[.,]\d{2})\s+(\d+[.,]\d{2})' + _SUFFIX + r'$')
    _fb_b = re.compile(r'^[\.]?\s*(' + _QTY + r')\s*(.+?)\s+(\d+[.,]\d{2})' + _SUFFIX + r'$')
    _fb_c = re.compile(r'^[\.]?\s*(' + _QTY + r')\s*(.+?)\s+(\d+[.,]\d{2})\s*$')

    def _make_item(nm: str, qty: int, unit_price: float, total: float) -> dict:
        return {
            "ชื่อสินค้า":    nm,
            "หมวดหมู่":     _categorize_by_rule(nm),
            "จำนวน":        qty,
            "ราคาต่อหน่วย": unit_price,
            "ยอดรวมสินค้า": total,
        }

    # ── FIX: ครอบ try/except ทุก line ──
    for line in lines[start_idx:]:
        try:
            line = line.strip()
            if not line: continue
            # normalize OCR ผิด: ไม่มีวรรณยุกต์
            line = re.sub(r'\bสวนลด\b', 'ส่วนลด', line)
            line = re.sub(r'\bจานวนสินค[้า]?รวม', 'จำนวนสินค้ารวม', line)
            # ตัด "สานานสินค้ารามรายการ" suffix ออกจากชื่อสินค้า
            line = re.sub(
                r'\s*[สจ].{0,4}สินค[้า]?.{0,3}ร[าว][มน].*รายการ\s*$', '',
                line, flags=re.IGNORECASE
            ).strip()
            compact = _collapse(line)

            if re.search(r'จ[าำ]?นวนสินค[้า]?[่า]?\s*รวม', line): continue  # จำนวนสินค้ารวม
            if re.search(r'จ.{0,3}นวนสินค.{0,3}รวม', compact): break
            if re.search(r'สา[นม]านสินค|สานานสินค', compact): break
            if re.search(r'[รง]\s*[ก-๙]{0,2}\s*ย\s*ก\s*า\s*ร', line): break
            if any(k.lower() in compact.lower() for k in skip_kw): continue
            if _RE_DATE.search(line): continue

            # ── ส่วนลด: "-40.00" ──
            m_discount = re.match(r'^-\s*(\d+[.,]\d{2})\s*$', line.strip())
            if m_discount:
                p = parse_price(m_discount.group(1))
                items.append(_make_item("ส่วนลด", 1, -p, -p))
                continue

            # ── FIX: โปรโมชั่น → แสดงเป็น item ──
            m_promo = re.search(
                r'โปรโมชั่น|promotion|แถม\s*\d*\s*bao|bao.*แถม',
                compact, re.IGNORECASE
            )
            if m_promo:
                prices = _find_prices_in_line(line)
                p = prices[0] if prices else 0.0
                promo_name = re.sub(r'^\d+\s*', '', re.sub(r'\s+', ' ', line).strip()).strip()
                # ราคาโปรโมชั่นจะเป็นลบ (ส่วนลด) ถ้าไม่มีราคาในบรรทัดจะรับจากบรรทัดถัดไป
                items.append(_make_item(promo_name, 1, -p if p else 0.0, -p if p else 0.0))
                continue

            if len(line) < 4: continue

            line_clean = re.sub(r'^[.\[\]>"\'`*]+\s*', '', line.strip())
            if not line_clean: line_clean = line.strip()
            line_clean = re.sub(r'\s+[Vv]\s*$', '', line_clean).strip()
            lf = _fix_spaced_price(re.sub(r'[\|！｜\[\]]+\s*$', '', line_clean).strip())

            has_stop = any(k in line or k in compact for k in stop_kw)
            has_item_pattern = bool(
                _full.match(lf) or _fb_a.match(lf) or
                _fb_b.match(lf) or _fb_c.match(lf)
            )
            if has_stop and not has_item_pattern:
                break

            m = _full.match(lf)
            if m:
                qty_raw, nm, up, tp = m.groups()
                qty = 1 if qty_raw == '-' else int(qty_raw)
                nm = _clean_item_name(nm)
                if len(re.sub(r'[^\u0E00-\u0E7FA-Za-z0-9]', '', nm)) >= 2:
                    u = parse_price(up); tt = parse_price(tp)
                    if tt < u * 0.5: tt = round(u * qty, 2)
                    items.append(_make_item(nm, qty, u, tt))
                continue

            m = _fb_a.match(lf)
            if m:
                nm, up, tp = m.groups()
                nm = _clean_item_name(nm)
                if len(re.sub(r'[^\u0E00-\u0E7FA-Za-z0-9]', '', nm)) >= 2:
                    items.append(_make_item(nm, 1, parse_price(up), parse_price(tp)))
                continue

            m = _fb_b.match(lf)  # "1 คูลคูล 12.00"
            if m:
                qty_raw, nm, price = m.groups()
                qty = 1 if qty_raw == '-' else int(qty_raw)
                nm = _clean_item_name(nm)
                # ลบ suffix "จํานวนสินค้ารวม..." ออกจาก nm ก่อน
                nm = re.sub(r'\s*จ.{0,2}นวนสินค.{0,3}รวม.*$', '', nm, flags=re.IGNORECASE).strip()
                if len(re.sub(r'[^\u0E00-\u0E7FA-Za-z0-9]', '', nm)) >= 2:
                    p = parse_price(price)
                    items.append(_make_item(nm, qty, p, round(p * qty, 2)))  # ✅ flush ทันที
                continue

            m = _fb_c.match(lf)
            if m:
                qty_raw, nm, price = m.groups()
                qty = 1 if qty_raw == '-' else int(qty_raw)
                nm = re.sub(r'\s+[0-9]+[.,][0-9]{2}\s*', ' ', nm)
                nm = _clean_item_name(nm)
                if len(re.sub(r'[^\u0E00-\u0E7FA-Za-z]', '', nm)) >= 3:
                    p = parse_price(price)
                    items.append(_make_item(nm, qty, p, round(p * qty, 2)))

        except Exception:
            # ── FIX: ข้าม line ที่ error แทนที่จะหยุดทั้งหมด ──
            continue

    # ── Fallback: ถ้าได้ 0 items ให้ลอง universal_item_parser ──
    if not items:
        items = universal_item_parser(text)

    return items

# ─────────────────────────────────────────────────────────────────────────────
# Excel export
# ─────────────────────────────────────────────────────────────────────────────
def build_excel(all_bills: list) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        rows = []
        for b in all_bills:
            d = b['bill']
            base = {
                "ไฟล์":           b['filename'],
                "วันที่":          d['date'],
                "เวลา":           d['time'],
                "รหัสสาขา":       d['pos_id'],
                "POS ID":         d.get('pos_machine',''),
                "เลขที่ใบเสร็จ":  d['rcpt_no'],
            }
            items = b['items'] or []
            for it in items:
                name = it.get("ชื่อสินค้า", "")
                cat  = BAO_CAFE_CATEGORY if _is_bao_item(name) \
                       else (it.get("หมวดหมู่") or "สินค้าเบ็ดเตล็ดอื่นๆ")
                unit_price = it.get("ราคาต่อหน่วย", 0)
                total_amt  = it.get("ยอดรวมสินค้า", 0)
                is_disc = (cat == "ส่วนลด/โปรโมชั่น"
                           or re.search(r'ส่วนลด|โปรโมชั่น|โปรโม|discount|promotion',
                                        name, re.IGNORECASE))
                row = base.copy()
                if is_disc:
                    row.update({
                        "ชื่อสินค้า":        "",
                        "หมวดหมู่":          "",
                        "จำนวน":             "",
                        "ราคาต่อหน่วย":      "",
                        "ยอดรวมสินค้า":      "",
                        "ยอดรวม":            "",
                        "ชื่อส่วนลด/โปรโม":  name,
                        "มูลค่าส่วนลด":      total_amt if total_amt <= 0 else -abs(total_amt),
                    })
                else:
                    row.update({
                        "ชื่อสินค้า":        name,
                        "หมวดหมู่":          cat,
                        "จำนวน":             it.get("จำนวน", 1),
                        "ราคาต่อหน่วย":      unit_price,
                        "ยอดรวมสินค้า":      total_amt,
                        "ยอดรวม":            "",
                        "ชื่อส่วนลด/โปรโม":  "",
                        "มูลค่าส่วนลด":      "",
                    })
                rows.append(row)
            # summary row
            summary = base.copy()
            summary.update({
                "ชื่อสินค้า":       "",
                "หมวดหมู่":         "",
                "จำนวน":            "",
                "ราคาต่อหน่วย":     "",
                "ยอดรวมสินค้า":     "",
                "ยอดรวม":           d['total_amount'],
                "ชื่อส่วนลด/โปรโม": "",
                "มูลค่าส่วนลด":     "",
            })
            rows.append(summary)
        # column order
        cols = ["ไฟล์","วันที่","เวลา","รหัสสาขา","POS ID","เลขที่ใบเสร็จ",
                "ชื่อสินค้า","หมวดหมู่","จำนวน","ราคาต่อหน่วย","ยอดรวมสินค้า",
                "ชื่อส่วนลด/โปรโม","มูลค่าส่วนลด","ยอดรวม"]
        df = pd.DataFrame(rows, columns=cols)
        df.to_excel(writer, index=False, sheet_name='ใบเสร็จ')
        # จัด column width
        ws = writer.sheets['ใบเสร็จ']
        widths = [20,12,8,10,8,22, 28,18,6,12,12, 28,12,10]
        for i, w in enumerate(widths, 1):
            from openpyxl.utils import get_column_letter
            ws.column_dimensions[get_column_letter(i)].width = w
    return output.getvalue()

# ─────────────────────────────────────────────────────────────────────────────
# Interactive crop canvas
# ─────────────────────────────────────────────────────────────────────────────
def _compute_split_positions(pil_img: Image.Image, n_bills: int) -> list:
    if n_bills <= 1:
        return []
    img_cv = pil_to_cv(pil_img)
    h, w = img_cv.shape[:2]
    gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
    col_ratio = (gray > 180).astype(np.uint8).mean(axis=0)
    splits, smooth, _ = _find_bill_splits(col_ratio, w, min_bill_frac=0.12, n_expected=n_bills)
    if len(splits) >= n_bills - 1:
        return [s / w for s in splits[:n_bills - 1]]
    return [i / n_bills for i in range(1, n_bills)]

def crop_component_html(pil_img: Image.Image, crop_mode: str = "free",
                        bill_count: int = 1, component_key: str = "crop1") -> str:
    b64 = img_to_b64(pil_img)
    orig_w, orig_h = pil_img.size
    is_a4 = "true" if crop_mode == "a4" else "false"
    split_fracs = _compute_split_positions(pil_img, bill_count) if bill_count > 1 else []
    split_js = json.dumps(split_fracs)
    n_bills_js = bill_count
    label_hint = {
        1: "ลากบนรูปเพื่อเลือกพื้นที่ที่ต้องการ Crop",
        2: "เส้นแดง = จุดตัดบิล · ลากปรับตำแหน่งได้",
        3: "เส้นแดง = จุดตัดบิล · ลากปรับได้",
    }.get(bill_count, "ลากบนรูปเพื่อเลือกพื้นที่")

    return f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:#f0eefc;font-family:sans-serif;padding:8px}}
#wrap{{position:relative;display:inline-block;border:2px solid #534AB7;border-radius:8px;overflow:hidden;cursor:crosshair}}
canvas{{display:block}}
#info{{margin-top:8px;font-size:13px;color:#534AB7;line-height:1.5}}
#result{{margin-top:6px;font-size:12px;background:#EEEDFE;border-radius:6px;padding:6px 10px;color:#3C3489;display:none}}
#split-info{{margin-top:6px;font-size:12px;background:#FEF2F2;border-radius:6px;padding:6px 10px;color:#991B1B;display:none;line-height:1.6}}
button{{margin-top:8px;margin-right:6px;padding:7px 16px;border-radius:6px;border:none;cursor:pointer;font-size:13px;font-weight:600}}
#btn-confirm{{background:#534AB7;color:#fff}}
#btn-confirm:disabled{{opacity:.4;cursor:not-allowed}}
#btn-reset{{background:#f1f0f0;color:#444}}
textarea#data-out{{width:100%;margin-top:8px;font-size:11px;font-family:monospace;height:60px;border-radius:6px;border:1px solid #ddd;padding:6px;display:none}}
</style></head><body>
<div id="wrap"><canvas id="c"></canvas></div>
<div id="info">{label_hint}</div>
<div id="result"></div>
<div id="split-info"></div><br>
<button id="btn-confirm" disabled>✅ ยืนยัน Crop</button>
<button id="btn-reset">🔄 วาดใหม่</button>
<br>
<div style="margin-top:10px;font-size:12px;color:#666">👇 คัดลอกข้อความนี้ไปวางในช่อง "ข้อมูล Crop" ด้านล่าง</div>
<textarea id="data-out" readonly onclick="this.select()"></textarea>
<script>
const IMG_W={orig_w},IMG_H={orig_h},A4={is_a4},N_BILLS={n_bills_js};
const SPLIT_FRACS={split_js};
const canvas=document.getElementById('c'),ctx=canvas.getContext('2d');
const MAX_W=420,MAX_H=300;
let scaleX,scaleY,sx,sy,ex,ey,drawing=false,hasCrop=false,cropRect=null;
let splitLines=SPLIT_FRACS.map(f=>({{frac:f,dragging:false}}));
let dragSplitIdx=-1,dragStartX=0,dragStartFrac=0;
const img=new Image();
img.onload=()=>{{
  let dw=Math.min(IMG_W,MAX_W),dh=dw*IMG_H/IMG_W;
  if(dh>MAX_H){{dh=MAX_H;dw=dh*IMG_W/IMG_H;}}
  canvas.width=Math.round(dw);canvas.height=Math.round(dh);
  scaleX=IMG_W/canvas.width;scaleY=IMG_H/canvas.height;
  draw();updateSplitInfo();
  cropRect={{x:0,y:0,w:canvas.width,h:canvas.height}};
  hasCrop=true;draw();finalise();
}};
img.src='data:image/png;base64,{b64}';
function draw(){{
  ctx.clearRect(0,0,canvas.width,canvas.height);
  ctx.drawImage(img,0,0,canvas.width,canvas.height);
  if(hasCrop&&cropRect)drawCropBox();
  if(N_BILLS>1)drawSplitLines();
}}
function drawCropBox(){{
  const{{x,y,w,h}}=cropRect;
  ctx.fillStyle='rgba(0,0,0,0.30)';
  ctx.fillRect(0,0,canvas.width,y);
  ctx.fillRect(0,y+h,canvas.width,canvas.height-y-h);
  ctx.fillRect(0,y,x,h);
  ctx.fillRect(x+w,y,canvas.width-x-w,h);
  ctx.strokeStyle='#534AB7';ctx.lineWidth=2;ctx.setLineDash([]);
  ctx.strokeRect(x,y,w,h);
  [[x,y],[x+w,y],[x,y+h],[x+w,y+h]].forEach(([cx,cy])=>{{
    ctx.fillStyle='#534AB7';ctx.fillRect(cx-5,cy-5,10,10);
  }});
}}
function drawSplitLines(){{
  splitLines.forEach((sl,i)=>{{
    const px=Math.round(sl.frac*canvas.width);
    ctx.shadowColor='rgba(220,0,0,0.4)';ctx.shadowBlur=6;
    ctx.strokeStyle='#EF4444';ctx.lineWidth=2.5;ctx.setLineDash([8,5]);
    ctx.beginPath();ctx.moveTo(px,0);ctx.lineTo(px,canvas.height);ctx.stroke();
    ctx.setLineDash([]);ctx.shadowBlur=0;
    const mid=canvas.height/2;
    ctx.fillStyle='#EF4444';
    ctx.beginPath();ctx.arc(px,mid,10,0,Math.PI*2);ctx.fill();
    ctx.fillStyle='#fff';ctx.font='bold 9px sans-serif';ctx.textAlign='center';ctx.textBaseline='middle';
    ctx.fillText('↔',px,mid);
    const prevPx=i===0?0:Math.round(splitLines[i-1].frac*canvas.width);
    const labelX=(prevPx+px)/2;
    const lbl=`บิล ${{i+1}}`;
    const tw=ctx.measureText(lbl).width+16;
    ctx.fillStyle='rgba(239,68,68,0.85)';
    ctx.beginPath();
    ctx.roundRect?ctx.roundRect(labelX-tw/2,6,tw,20,6):ctx.rect(labelX-tw/2,6,tw,20);
    ctx.fill();
    ctx.fillStyle='#fff';ctx.font='bold 11px sans-serif';ctx.textAlign='center';ctx.textBaseline='middle';
    ctx.fillText(lbl,labelX,16);
  }});
  const lastPx=splitLines.length>0?Math.round(splitLines[splitLines.length-1].frac*canvas.width):0;
  const lastLbl=`บิล ${{N_BILLS}}`;
  const tw2=ctx.measureText(lastLbl).width+16;
  const labelX2=(lastPx+canvas.width)/2;
  ctx.fillStyle='rgba(239,68,68,0.85)';
  ctx.beginPath();
  ctx.roundRect?ctx.roundRect(labelX2-tw2/2,6,tw2,20,6):ctx.rect(labelX2-tw2/2,6,tw2,20);
  ctx.fill();
  ctx.fillStyle='#fff';ctx.font='bold 11px sans-serif';ctx.textAlign='center';ctx.textBaseline='middle';
  ctx.fillText(lastLbl,labelX2,16);
}}
function getSplitHitIdx(px){{
  for(let i=0;i<splitLines.length;i++){{
    if(Math.abs(px-splitLines[i].frac*canvas.width)<=14)return i;
  }}
  return -1;
}}
function updateSplitInfo(){{
  if(N_BILLS<=1)return;
  const info=document.getElementById('split-info');
  info.style.display='block';
  let fracs=[0,...splitLines.map(s=>s.frac),1.0];
  let parts=[];
  for(let i=0;i<N_BILLS;i++){{
    const pct=Math.round((fracs[i+1]-fracs[i])*100);
    const px=Math.round(fracs[i+1]*IMG_W)-Math.round(fracs[i]*IMG_W);
    parts.push(`บิล ${{i+1}}: ${{pct}}% (${{px}}px)`);
  }}
  info.innerHTML='🔴 จุดตัด: '+splitLines.map((s,i)=>`<b>${{Math.round(s.frac*IMG_W)}}px</b>`).join(' · ')+'<br>'+parts.join(' &nbsp;|&nbsp; ');
}}
function getPos(e){{
  const r=canvas.getBoundingClientRect();
  const cx=(e.touches?e.touches[0].clientX:e.clientX)-r.left;
  const cy=(e.touches?e.touches[0].clientY:e.clientY)-r.top;
  return[Math.max(0,Math.min(cx,canvas.width)),Math.max(0,Math.min(cy,canvas.height))];
}}
canvas.addEventListener('mousedown',e=>{{
  const[px,py]=getPos(e);
  dragSplitIdx=getSplitHitIdx(px);
  if(dragSplitIdx>=0){{dragStartX=px;dragStartFrac=splitLines[dragSplitIdx].frac;canvas.style.cursor='ew-resize';return;}}
  drawing=true;sx=px;sy=py;hasCrop=false;
}});
canvas.addEventListener('mousemove',e=>{{
  const[px,py]=getPos(e);
  if(dragSplitIdx>=0){{
    let newFrac=dragStartFrac+(px-dragStartX)/canvas.width;
    const minF=dragSplitIdx>0?splitLines[dragSplitIdx-1].frac+0.03:0.03;
    const maxF=dragSplitIdx<splitLines.length-1?splitLines[dragSplitIdx+1].frac-0.03:0.97;
    splitLines[dragSplitIdx].frac=Math.max(minF,Math.min(maxF,newFrac));
    draw();updateSplitInfo();finalise();return;
  }}
  if(drawing){{
    ex=px;ey=py;
    let w=ex-sx,h=ey-sy;
    if(A4)h=Math.abs(w)*1.4142*Math.sign(h);
    cropRect={{x:Math.min(sx,ex),y:Math.min(sy,ey),w:Math.abs(w),h:Math.abs(h)}};
    hasCrop=true;draw();return;
  }}
  canvas.style.cursor=getSplitHitIdx(px)>=0?'ew-resize':'crosshair';
}});
canvas.addEventListener('mouseup',e=>{{
  if(dragSplitIdx>=0){{dragSplitIdx=-1;canvas.style.cursor='crosshair';draw();finalise();return;}}
  drawing=false;finalise();
}});
canvas.addEventListener('mouseleave',()=>{{
  if(dragSplitIdx>=0){{dragSplitIdx=-1;draw();finalise();}}
  drawing=false;
}});
canvas.addEventListener('touchstart',e=>{{
  e.preventDefault();
  const[px,py]=getPos(e);
  dragSplitIdx=getSplitHitIdx(px);
  if(dragSplitIdx>=0){{dragStartX=px;dragStartFrac=splitLines[dragSplitIdx].frac;return;}}
  drawing=true;sx=px;sy=py;hasCrop=false;
}},{{passive:false}});
canvas.addEventListener('touchmove',e=>{{
  e.preventDefault();
  const[px,py]=getPos(e);
  if(dragSplitIdx>=0){{
    let newFrac=dragStartFrac+(px-dragStartX)/canvas.width;
    const minF=dragSplitIdx>0?splitLines[dragSplitIdx-1].frac+0.03:0.03;
    const maxF=dragSplitIdx<splitLines.length-1?splitLines[dragSplitIdx+1].frac-0.03:0.97;
    splitLines[dragSplitIdx].frac=Math.max(minF,Math.min(maxF,newFrac));
    draw();updateSplitInfo();return;
  }}
  if(!drawing)return;
  ex=px;ey=py;
  let w=ex-sx,h=ey-sy;
  if(A4)h=Math.abs(w)*1.4142*Math.sign(h);
  cropRect={{x:Math.min(sx,ex),y:Math.min(sy,ey),w:Math.abs(w),h:Math.abs(h)}};
  hasCrop=true;draw();
}},{{passive:false}});
canvas.addEventListener('touchend',e=>{{
  e.preventDefault();
  if(dragSplitIdx>=0){{dragSplitIdx=-1;draw();finalise();return;}}
  drawing=false;finalise();
}},{{passive:false}});
function finalise(){{
  if(!hasCrop||!cropRect||cropRect.w<10||cropRect.h<10)return;
  const ox=Math.round(cropRect.x*scaleX),oy=Math.round(cropRect.y*scaleY);
  const ow=Math.round(cropRect.w*scaleX),oh=Math.round(cropRect.h*scaleY);
  const splitPx=splitLines.map(s=>Math.round(s.frac*IMG_W));
  document.getElementById('result').style.display='block';
  document.getElementById('result').textContent=`เลือก: ${{ow}}×${{oh}} px (ตำแหน่ง ${{ox}},${{oy}})`+(splitPx.length?`  |  จุดตัด: ${{splitPx.join(', ')}} px`:'');
  const btn=document.getElementById('btn-confirm');
  btn.disabled=false;
  const dataObj={{x:ox,y:oy,w:ow,h:oh,splits:splitPx}};
  btn.dataset.crop=JSON.stringify(dataObj);
  const out=document.getElementById('data-out');
  out.style.display='block';
  out.value=JSON.stringify(dataObj);
}}
document.getElementById('btn-confirm').onclick=()=>{{
  const out=document.getElementById('data-out');
  out.select();
  document.execCommand('copy');
  out.style.background='#d1fae5';
}};
document.getElementById('btn-reset').onclick=()=>{{
  hasCrop=false;cropRect=null;
  splitLines=SPLIT_FRACS.map(f=>({{frac:f,dragging:false}}));
  document.getElementById('result').style.display='none';
  document.getElementById('btn-confirm').disabled=true;
  document.getElementById('data-out').style.display='none';
  draw();updateSplitInfo();
}};
</script></body></html>"""

# ─────────────────────────────────────────────────────────────────────────────
# Batch mode
# ─────────────────────────────────────────────────────────────────────────────
def run_batch_analysis(files: list, progress_cb=None, auto_detect_multi: bool = False,
                        ocr_engine: str = "tesseract", bills_per_image: int = 1) -> list:
    results = []
    n = len(files)
    for i, (fname, fbytes) in enumerate(files, 1):
        if progress_cb: progress_cb(i, n, fname)
        try:
            pil = Image.open(io.BytesIO(fbytes)).convert("RGB")
            img_cv = pil_to_cv(pil)

            # ── ตรวจสอบความชัด ──
            _gray_b = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
            _blur_b = cv2.Laplacian(_gray_b, cv2.CV_64F).var()
            if _blur_b < 20:  # ไม่ชัดมาก → skip OCR
                results.append({"filename": fname,
                    "bill": {"date":"ไม่พบ","time":"ไม่พบ","branch":"ไม่พบ","name":"ไม่พบ",
                             "total_amount":0.0,"cash":0.0,"change":0.0,"pos_machine":"ไม่พบ",
                             "pos_id":"ไม่พบ","rcpt_no":"ไม่พบ","tax_id":"ไม่พบ","user":"ไม่พบ"},
                    "items": [], "raw_text": f"[รูปไม่ชัด blur={_blur_b:.0f}]",
                    "image": img_to_bytes_png(img_cv), "blurry": True})
                continue

            if ocr_engine == "gdrive" and bills_per_image > 1:
                thumb_crops = split_receipts_image(img_cv, n_expected=bills_per_image)
                if len(thumb_crops) < bills_per_image:
                    thumb_crops = thumb_crops + [img_cv] * (bills_per_image - len(thumb_crops))
                thumb_crops = thumb_crops[:bills_per_image]
                for ci, crop in enumerate(thumb_crops):
                    label = f"{fname} — บิล {ci+1}"
                    st.session_state["_gdrive_raw_texts"] = []
                    text       = run_ocr(crop, engine="gdrive")
                    gdrive_raw = (st.session_state.get("_gdrive_raw_texts") or [""])[0]
                    text_for_gemini = gdrive_raw if gdrive_raw else text
                    if is_gemini_configured() and text_for_gemini.strip():
                        gr    = extract_with_gemini(text_for_gemini, ocr_source="gdrive")
                        bill  = gr["bill"] if gr["ok"] else extract_receipt(text)
                        items = gr["items"] if gr["ok"] and gr["items"] else extract_items_cj(text)
                    else:
                        bill  = extract_receipt(text)
                        items = extract_items_cj(text)
                    results.append({"filename": label, "bill": bill, "items": items,
                                    "raw_text": text, "gdrive_raw": gdrive_raw,
                                    "image": img_to_bytes_png(crop)})
                continue

            if auto_detect_multi:
                sub_crops = auto_crop_receipts(img_cv)
            else:
                sub_crops = [img_cv]

            if len(sub_crops) == 1:
                st.session_state["_gdrive_raw_texts"] = []
                text       = run_ocr(sub_crops[0], engine=ocr_engine)
                gdrive_raw = (st.session_state.get("_gdrive_raw_texts") or [""])[0]
                text_for_gemini = gdrive_raw if (ocr_engine == "gdrive" and gdrive_raw) else text
                if is_gemini_configured() and text_for_gemini.strip():
                    gr    = extract_with_gemini(text_for_gemini,
                                                ocr_source="gdrive" if ocr_engine == "gdrive" else "tesseract")
                    bill  = gr["bill"] if gr["ok"] else extract_receipt(text)
                    items = gr["items"] if gr["ok"] and gr["items"] else extract_items_cj(text)
                else:
                    bill  = extract_receipt(text)
                    items = extract_items_cj(text)
                results.append({"filename": fname, "bill": bill, "items": items,
                                "raw_text": text, "gdrive_raw": gdrive_raw,
                                "image": img_to_bytes_png(sub_crops[0])})
            else:
                for ci, crop in enumerate(sub_crops, 1):
                    label = f"{fname} — บิล {ci}"
                    st.session_state["_gdrive_raw_texts"] = []
                    text       = run_ocr(crop, engine=ocr_engine)
                    gdrive_raw = (st.session_state.get("_gdrive_raw_texts") or [""])[0]
                    text_for_gemini = gdrive_raw if (ocr_engine == "gdrive" and gdrive_raw) else text
                    if is_gemini_configured() and text_for_gemini.strip():
                        gr    = extract_with_gemini(text_for_gemini,
                                                    ocr_source="gdrive" if ocr_engine == "gdrive" else "tesseract")
                        bill  = gr["bill"] if gr["ok"] else extract_receipt(text)
                        items = gr["items"] if gr["ok"] and gr["items"] else extract_items_cj(text)
                    else:
                        bill  = extract_receipt(text)
                        items = extract_items_cj(text)
                    results.append({"filename": label, "bill": bill, "items": items,
                                    "raw_text": text, "gdrive_raw": gdrive_raw,
                                    "image": img_to_bytes_png(crop)})
        except Exception as e:
            results.append({"filename": fname,
                            "bill": {"date":"ไม่พบ","time":"ไม่พบ","branch":"ไม่พบ","name":"ไม่พบ",
                                     "total_amount":0.0,"cash":0.0,"change":0.0,"pos_machine":"ไม่พบ",
                                     "pos_id":"ไม่พบ","rcpt_no":"ไม่พบ","tax_id":"ไม่พบ","user":"ไม่พบ"},
                            "items": [], "raw_text": f"[ERROR] {e}", "image": None})
    return results

def run_batch_mode_ui():
    st.markdown(f'<p class="sec-header">{t("upload_label")}</p>', unsafe_allow_html=True)
    st.markdown(f'<div class="hint-box">{t("batch_upload_hint")}</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader("batch_files", type=["png","jpg","jpeg","heic"],
                                accept_multiple_files=True, label_visibility="collapsed",
                                key="batch_uploader")
    if uploaded:
        new_files = [(f.name, f.read()) for f in uploaded]
        names_new = [x[0] for x in new_files]
        names_old = [x[0] for x in S.batch_files]
        if names_new != names_old:
            S.batch_files = new_files
            S.all_bills = []
    if not S.batch_files:
        st.info("👆 เลือกไฟล์ภาพทั้งหมดที่ต้องการวิเคราะห์")
        return
    st.success(t("batch_found")(len(S.batch_files)))
    cols_per_row = 6
    for chunk_start in range(0, len(S.batch_files), cols_per_row):
        chunk = S.batch_files[chunk_start:chunk_start+cols_per_row]
        cols  = st.columns(len(chunk))
        for ci, (fname, fbytes) in enumerate(chunk):
            with cols[ci]:
                try:
                    pil_t = Image.open(io.BytesIO(fbytes))
                    st.image(pil_t, use_container_width=True,
                              caption=fname[:16] + ('…' if len(fname)>16 else ''))
                    # blur check
                    cv_t = pil_to_cv(pil_t)
                    gray_t = cv2.cvtColor(cv_t, cv2.COLOR_BGR2GRAY)
                    blur_t = cv2.Laplacian(gray_t, cv2.CV_64F).var()
                    if blur_t < 30:
                        st.caption("⚠️ รูปไม่ชัด")
                except Exception:
                    st.warning(fname[:16])
    st.divider()
    if S.ocr_engine == "gdrive":
        bills_col, _ = st.columns([2, 2])
        with bills_col:
            bills_per_img = st.radio(
                "📄 แต่ละรูปมีกี่บิล", [1, 2, 3],
                index=S.get("batch_bills_per_image", 1) - 1,
                horizontal=True, key="batch_bills_per_image_radio")
            S["batch_bills_per_image"] = bills_per_img
            if bills_per_img > 1:
                st.info(f"✅ จะ split รูปแล้ว Google Drive OCR ทีละบิล → Gemini วิเคราะห์ทีละบิล")
    else:
        S["batch_bills_per_image"] = 1

    auto_detect = st.checkbox("🪄 ตรวจจับหลายใบเสร็จในภาพเดียวอัตโนมัติ + ทำพื้นหลังขาว",
                               value=False, key="batch_auto_detect")
    bc1, bc2 = st.columns([3,1])
    with bc1:
        run_clicked = st.button(t("batch_analyze"), use_container_width=True, type="primary")
    with bc2:
        if st.button(t("reset"), use_container_width=True, key="batch_reset"):
            for k, v in _DEFAULTS.items(): S[k] = v
            S.bill_count = -1
            st.rerun()
    if run_clicked:
        progress_bar = st.progress(0, text=t("batch_progress")(0, len(S.batch_files), ""))
        def _cb(i, n, fname):
            progress_bar.progress(i / n, text=t("batch_progress")(i, n, fname))
        results = run_batch_analysis(S.batch_files, progress_cb=_cb,
                                      auto_detect_multi=auto_detect,
                                      ocr_engine=S.ocr_engine,
                                      bills_per_image=S.get("batch_bills_per_image", 1))
        progress_bar.progress(1.0, text=t("batch_done")(len(results)))
        S.all_bills = results
        st.rerun()
    if S.all_bills:
        _render_bills_ui(S.all_bills, key_prefix="b")
        st.divider()
        dl_c, sheets_c, rs_c = st.columns([2, 2, 1])
        with dl_c:
            st.download_button(t("download"), data=build_excel(S.all_bills),
                               file_name="receipts.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)
        with sheets_c:                                  # ← เพิ่ม
            if st.button("📊 บันทึกลง Google Sheets",
                         use_container_width=True, type="secondary",
                         key="single_save_sheets"):
                save_to_sheets(S.all_bills)                   
        with rs_c:
            if st.button(t("reset"), use_container_width=True):
                for k,v in _DEFAULTS.items(): S[k]=v
                st.rerun()
# ── Google Sheets API ──────────────────────────────────────────────────────
_SHEETS_FILE_ID = "1IhQFHxlK7vlAJ-xxgWNA_M2fcXp-9jZm8WQZGZC4MxQ"
_SHEET_CORRECT  = "sheet1"   # ชื่อ sheet สำหรับชื่อดิบ → ชื่อถูก
_SHEET_DELETED  = "sheet2"    # ชื่อ sheet สำหรับรายการที่ถูกลบ

def _sheets_append(token: str, sheet_name: str, rows: list):
    """append rows ลง Google Sheet"""
    import requests as _req
    url = (f"https://sheets.googleapis.com/v4/spreadsheets/{_SHEETS_FILE_ID}"
           f"/values/{sheet_name}!A1:Z1000:append"
           f"?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS")
    body = {"values": rows}
    resp = _req.post(url, json=body,
                     headers={"Authorization": f"Bearer {token}"}, timeout=15)
    resp.raise_for_status()
    return resp.json()

def _sheets_ensure_headers(token: str):
    """สร้าง header row ถ้า sheet ยังว่างอยู่"""
    import requests as _req
    headers_map = {
        _SHEET_CORRECT: ["วันที่บันทึก","ไฟล์บิล","วันที่บิล","รหัสสาขา",
                         "ชื่อ OCR (ดิบ)","ชื่อที่ถูกต้อง","หมวดหมู่",
                         "จำนวน","ราคาต่อหน่วย","ยอดรวม","หมายเหตุ"],
        _SHEET_DELETED: ["วันที่บันทึก","ไฟล์บิล","วันที่บิล","รหัสสาขา",
                         "ชื่อ OCR (ที่ลบ)","หมวดหมู่เดิม","ราคาเดิม","สาเหตุที่ลบ"],
    }
    for sheet_name, header in headers_map.items():
        url = (f"https://sheets.googleapis.com/v4/spreadsheets/{_SHEETS_FILE_ID}"
               f"/values/{sheet_name}!A1")
        resp = _req.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=10)
        if resp.ok:
            data = resp.json().get("values", [])
            if not data:  # ว่างอยู่ → ใส่ header
                _sheets_append(token, sheet_name, [header])

def save_to_sheets(all_bills: list):
    """บันทึกรายการที่แก้/ลบ ลง Google Sheet"""
    if not is_gdrive_token_ready():
        st.error("❌ กรุณา Login Google Drive ก่อน")
        return False
    token = st.session_state["gdrive_token"]
    
    try:
        _sheets_ensure_headers(token)
        
        from datetime import datetime
        now = datetime.now().strftime("%d/%m/%Y %H:%M")
        
        correct_rows = []
        deleted_rows = []
        
        for b in all_bills:
            bill_date   = b['bill'].get('date', '')
            branch_id   = b['bill'].get('pos_id', '')
            filename    = b['filename']
            
            # ── รายการที่ถูก (เพิ่มเอง + แก้ชื่อ) ──
            for it in b.get('items', []):
                if it.get('is_manual') or it.get('ชื่อ OCR เดิม'):
                    correct_rows.append([
                        now, filename, bill_date, branch_id,
                        it.get('ชื่อ OCR เดิม', ''),   # ชื่อดิบจาก OCR
                        it.get('ชื่อสินค้า', ''),       # ชื่อที่ถูกต้อง
                        it.get('หมวดหมู่', ''),
                        it.get('จำนวน', 1),
                        it.get('ราคาต่อหน่วย', 0),
                        it.get('ยอดรวมสินค้า', 0),
                        'เพิ่มเอง' if not it.get('ชื่อ OCR เดิม') else 'Gemini แปลง',
                    ])
            
            # ── รายการที่ถูกลบ ──
            for it in b.get('deleted_items', []):
                deleted_rows.append([
                    now, filename, bill_date, branch_id,
                    it.get('deleted_name_ocr', ''),
                    it.get('หมวดหมู่', ''),
                    it.get('ยอดรวมสินค้า', 0),
                    'ลบโดยผู้ใช้',
                ])
        
        if correct_rows:
            _sheets_append(token, _SHEET_CORRECT, correct_rows)
        if deleted_rows:
            _sheets_append(token, _SHEET_DELETED, deleted_rows)
        
        total = len(correct_rows) + len(deleted_rows)
        st.success(f"✅ บันทึกลง Google Sheets แล้ว {total} รายการ "
                   f"({len(correct_rows)} แก้/เพิ่ม, {len(deleted_rows)} ลบ)")
        return True
        
    except Exception as e:
        st.error(f"❌ บันทึกไม่สำเร็จ: {e}")
        return False
def gemini_clean_item_name(raw_name: str, raw_text: str = "") -> dict:
    prompt = f"""ชื่อสินค้าจาก OCR ใบเสร็จ CJ Express อ่านผิด:
"{raw_name}"

{f'บริบทจากบิล:{chr(10)}{raw_text[:500]}' if raw_text else ''}

{CJ_PRODUCT_KNOWLEDGE}
{BAO_CAFE_MENU}

แปลงชื่อ OCR ผิดๆ เป็นชื่อสินค้าที่ถูกต้อง และเลือกหมวดหมู่

{CATEGORY_PROMPT}

ตอบ JSON เท่านั้น:
{{"ชื่อที่ถูกต้อง": "", "หมวดหมู่": ""}}"""
    try:
        raw = _call_gemini(prompt, max_tokens=200)
        raw = re.sub(r"```(?:json)?\s*|\s*```", "", raw).strip()
        data = json.loads(raw)
        return {
            "ชื่อที่ถูกต้อง": str(data.get("ชื่อที่ถูกต้อง", raw_name)),
            "หมวดหมู่":      str(data.get("หมวดหมู่", "สินค้าเบ็ดเตล็ดอื่นๆ")),
        }
    except Exception:
        return {"ชื่อที่ถูกต้อง": raw_name, "หมวดหมู่": "สินค้าเบ็ดเตล็ดอื่นๆ"}   
def _render_bills_ui(all_bills, key_prefix=""):
    st.success(t("found")(len(all_bills)))
    
    for idx, b in enumerate(all_bills):
        with st.expander(f"📄 {b['filename']}", expanded=True):
            ic, dc = st.columns([1,2])
            with ic:
                if b.get('image'): st.image(b['image'], use_container_width=True)
            with dc:
                d = b['bill']
                r1 = st.columns(5)
                d['date']        = r1[0].text_input("วันที่",         d['date'],               key=f"{key_prefix}dt{idx}")
                d['time']        = r1[1].text_input("เวลา",           d['time'],               key=f"{key_prefix}tm{idx}")
                d['pos_id']      = r1[2].text_input("รหัสสาขา",      d['pos_id'],             key=f"{key_prefix}ps{idx}")
                d['pos_machine'] = r1[3].text_input("POS ID",         d.get('pos_machine',''), key=f"{key_prefix}pm{idx}")
                d['rcpt_no']     = r1[4].text_input("เลขที่ใบเสร็จ", d['rcpt_no'],            key=f"{key_prefix}rc{idx}")
                tot_str = st.text_input("ยอดรวม", f"{float(d['total_amount']):.2f}", key=f"{key_prefix}tot{idx}")
                d['total_amount'] = parse_price(tot_str)

            st.metric("💰 ยอดรวม", f"{d['total_amount']:.2f} ฿")
            enriched_log = st.session_state.get("_sheet_enriched_log", [])
            if enriched_log:
                with st.expander(f"🔗 จับคู่จาก Sheet {len(enriched_log)} รายการ"):
                    for log in enriched_log:
                        st.caption(f"✅ {log}")
                st.session_state["_sheet_enriched_log"] = []  # clear หลังแสดง

            # ── init session state สำหรับ items และ deleted ──
            items_key   = f"items_{key_prefix}{idx}"
            deleted_key = f"deleted_{key_prefix}{idx}"
            if items_key not in st.session_state:
                st.session_state[items_key] = b['items'].copy()
            if deleted_key not in st.session_state:
                st.session_state[deleted_key] = []  # เก็บรายการที่ถูกลบ

            items = st.session_state[items_key]

            st.markdown("**🛒 รายการสินค้า**")
            
            # ── แสดง items พร้อมปุ่มลบ ──
            for ii, it in enumerate(items):
                col_name, col_cat, col_qty, col_price, col_total, col_del = st.columns([3,2,1,1.5,1.5,0.5])
                
                # แก้ไขชื่อ inline
                new_name = col_name.text_input(
                    "ชื่อ", it.get("ชื่อสินค้า",""),
                    key=f"{key_prefix}name{idx}_{ii}", label_visibility="collapsed")
                new_cat = col_cat.selectbox(
                    "หมวด", 
                    ["Bao Cafe","อาหารพร้อมทานและเบเกอรี่","ขนมและของขบเคี้ยว",
                     "เครื่องดื่ม","ของใช้ส่วนตัว","ของใช้ในบ้าน",
                     "เวชภัณฑ์และอุปกรณ์ดูแลสุขภาพ","สินค้าเบ็ดเตล็ดอื่นๆ","ส่วนลด/โปรโมชั่น"],
                    index=0 if it.get("หมวดหมู่") not in ["Bao Cafe","อาหารพร้อมทานและเบเกอรี่","ขนมและของขบเคี้ยว","เครื่องดื่ม","ของใช้ส่วนตัว","ของใช้ในบ้าน","เวชภัณฑ์และอุปกรณ์ดูแลสุขภาพ","สินค้าเบ็ดเตล็ดอื่นๆ","ส่วนลด/โปรโมชั่น"] else ["Bao Cafe","อาหารพร้อมทานและเบเกอรี่","ขนมและของขบเคี้ยว","เครื่องดื่ม","ของใช้ส่วนตัว","ของใช้ในบ้าน","เวชภัณฑ์และอุปกรณ์ดูแลสุขภาพ","สินค้าเบ็ดเตล็ดอื่นๆ","ส่วนลด/โปรโมชั่น"].index(it.get("หมวดหมู่","สินค้าเบ็ดเตล็ดอื่นๆ")),
                    key=f"{key_prefix}cat{idx}_{ii}", label_visibility="collapsed")
                new_qty = col_qty.number_input(
                    "จำนวน", min_value=1, value=int(it.get("จำนวน",1)),
                    key=f"{key_prefix}qty{idx}_{ii}", label_visibility="collapsed")
                new_price = col_price.number_input(
                    "ราคา/หน่วย", value=float(it.get("ราคาต่อหน่วย",0)),
                    key=f"{key_prefix}uprice{idx}_{ii}", label_visibility="collapsed")
                new_total = col_total.number_input(
                    "รวม", value=float(it.get("ยอดรวมสินค้า",0)),
                    key=f"{key_prefix}total{idx}_{ii}", label_visibility="collapsed")

                # อัปเดตค่า
                items[ii] = {
                    "ชื่อสินค้า": new_name, "หมวดหมู่": new_cat,
                    "จำนวน": new_qty, "ราคาต่อหน่วย": new_price,
                    "ยอดรวมสินค้า": new_total,
                    "is_manual": it.get("is_manual", False)
                }

                # ปุ่มลบ
                if col_del.button("🗑", key=f"{key_prefix}del{idx}_{ii}", help="ลบรายการนี้"):
                    deleted = it.copy()
                    deleted["deleted_from_bill"] = b['filename']
                    deleted["deleted_name_ocr"]  = it.get("ชื่อสินค้า","")
                    st.session_state[deleted_key].append(deleted)
                    st.session_state[items_key].pop(ii)
                    st.rerun()

            # ── เพิ่มรายการใหม่ ──
            st.markdown("---")
            st.markdown("**➕ เพิ่มรายการ**")
            add_cols = st.columns([3,2,1,1.5,1.5,1])
            new_item_name  = add_cols[0].text_input("ชื่อสินค้า", key=f"{key_prefix}add_name{idx}", placeholder="ชื่อสินค้า")
            new_item_cat   = add_cols[1].selectbox("หมวด", 
                ["Bao Cafe","อาหารพร้อมทานและเบเกอรี่","ขนมและของขบเคี้ยว",
                 "เครื่องดื่ม","ของใช้ส่วนตัว","ของใช้ในบ้าน",
                 "เวชภัณฑ์และอุปกรณ์ดูแลสุขภาพ","สินค้าเบ็ดเตล็ดอื่นๆ","ส่วนลด/โปรโมชั่น"],
                key=f"{key_prefix}add_cat{idx}")
            new_item_qty   = add_cols[2].number_input("จำนวน", min_value=1, value=1, key=f"{key_prefix}add_qty{idx}")
            new_item_price = add_cols[3].number_input("ราคา/หน่วย", value=0.0, key=f"{key_prefix}add_price{idx}")
            new_item_total = add_cols[4].number_input("รวม", value=0.0, key=f"{key_prefix}add_total{idx}")
            
            if add_cols[5].button("➕ เพิ่ม", key=f"{key_prefix}add_btn{idx}"):
                if new_item_name.strip():
                    with st.spinner("✨ Gemini กำลังแปลงชื่อ..."):
                        cleaned = gemini_clean_item_name(
                            raw_name=new_item_name.strip(),
                            raw_text=b.get('raw_text', '')
                        )
                    corrected_name = cleaned["ชื่อที่ถูกต้อง"]
                    corrected_cat  = cleaned["หมวดหมู่"]
                    final_cat = new_item_cat if new_item_cat != "สินค้าเบ็ดเตล็ดอื่นๆ" else corrected_cat
                    total_val = new_item_total if new_item_total != 0 else new_item_price * new_item_qty
                    st.session_state[items_key].append({
                        "ชื่อสินค้า":      corrected_name,
                        "ชื่อ OCR เดิม":  new_item_name.strip(),
                        "หมวดหมู่":       final_cat,
                        "จำนวน":          new_item_qty,
                        "ราคาต่อหน่วย":  new_item_price,
                        "ยอดรวมสินค้า":  total_val,
                        "is_manual":      True,
                    })
                    st.success(f'✅ "{new_item_name}" → **"{corrected_name}"** [{final_cat}]')
                    st.rerun()

            # sync กลับ
            b['items'] = st.session_state[items_key]
            b['deleted_items'] = st.session_state[deleted_key]

            # raw text
            with st.expander("🔬 ข้อความดิบ (คัดลอกชื่อสินค้าจากนี้)"):
                st.text_area("", b['raw_text'], height=200,
                             key=f"{key_prefix}raw{idx}", disabled=False)

# ─────────────────────────────────────────────────────────────────────────────
# MAIN UI
# ─────────────────────────────────────────────────────────────────────────────
def main():
    _qp = st.query_params
    if "code" in _qp and "gdrive_token" not in st.session_state:
        _code           = _qp["code"]
        _state_received = _qp.get("state", "")
        _state_expected = st.session_state.get("gdrive_oauth_state", "")
        _state_ok = (not _state_expected) or (_state_received == _state_expected)
        if _state_ok:
            with st.spinner("🔐 กำลัง Login Google Drive..."):
                _ok = gdrive_exchange_code(_code)
            st.query_params.clear()
            if _ok:
                st.success("✅ Login สำเร็จ!")
                st.rerun()
            else:
                _err = st.session_state.pop("gdrive_auth_error", "unknown")
                st.error(f"Login ไม่สำเร็จ: {_err}")
        else:
            st.query_params.clear()
            st.warning("OAuth state ไม่ตรง")

    col_t, col_l = st.columns([5,1])
    with col_t: st.markdown(f"## {t('title')}")
    with col_l:
        lc = st.radio("", ["ไทย","EN"], horizontal=True, label_visibility="collapsed",
                      index=0 if S.lang=="th" else 1)
        nl = "th" if lc=="ไทย" else "en"
        if nl != S.lang: S.lang = nl; st.rerun()

    with st.expander("⚙️ ตัวเลือก OCR Engine", expanded=False):
        if "_gdrive_warning" in S:
            st.warning(f"⚠️ Google Drive OCR error: {S._gdrive_warning}")
            del S["_gdrive_warning"]
        if "_vision_api_warning" in S:
            st.warning(f"⚠️ Google Vision API error: {S._vision_api_warning}")
            del S["_vision_api_warning"]

        gdrive_ready = is_gdrive_configured()
        vision_ready = is_vision_api_configured()
        engine_options = [
            "🆓 Tesseract (ฟรี, รันในเครื่อง)",
            f"📄 Google Drive OCR {'✅' if is_gdrive_token_ready() else ('🔑 ต้อง Login' if gdrive_ready else '⚙️ ต้องตั้งค่า')} (ฟรี ไม่จำกัด)",
            f"🎯 Google Cloud Vision API {'✅' if vision_ready else '⚙️ ต้องตั้งค่า'} (แม่นสุด)",
        ]
        engine_map = ["tesseract", "gdrive", "vision"]
        current_idx = engine_map.index(S.ocr_engine) if S.ocr_engine in engine_map else 0
        choice = st.radio("เลือก OCR Engine", engine_options, index=current_idx,
                          key="ocr_engine_radio", label_visibility="collapsed")
        new_engine = engine_map[engine_options.index(choice)]
        if new_engine != S.ocr_engine:
            S.ocr_engine = new_engine
        if S.ocr_engine == "gdrive":
            render_gdrive_login_ui()
        elif S.ocr_engine == "vision":
            if vision_ready:
                st.success("✅ ตั้งค่า Google Vision API key แล้ว")
            else:
                st.warning("⚠️ ยังไม่ได้ตั้งค่า GOOGLE_VISION_API_KEY")
        else:
            st.caption("💡 Tesseract ฟรีและรันในเครื่องทั้งหมด แต่แม่นน้อยกว่า")

        gemini_ready = is_gemini_configured()
        st.divider()
        st.markdown("**✨ Gemini API — วิเคราะห์ข้อมูลจาก raw text**")
        if gemini_ready:
            st.success("✅ ตั้งค่า Gemini API key แล้ว — ฟรี 1,500 requests/วัน")
        else:
            st.info("💡 เพิ่ม GEMINI_API_KEY ใน `.streamlit/secrets.toml` เพื่อแม่นขึ้น")

        with st.expander("📊 เปรียบเทียบ OCR Engine"):
            st.markdown("""| Engine | ค่าใช้จ่าย | ความแม่น | ภาษาไทย | Speed |
|--------|-----------|---------|---------|-------|
| 🆓 Tesseract | ฟรี | ⭐⭐ | พอใช้ | ⚡⚡⚡ |
| 📄 **Drive OCR + Gemini** | **ฟรี ไม่จำกัด** | **⭐⭐⭐⭐⭐** | **ดีมาก** | ⚡⚡ |
| 🎯 Google Vision API | ฟรี 1K/เดือน | ⭐⭐⭐⭐⭐ | ดีมาก | ⚡⚡ |""")

    st.markdown(f'<p class="sec-header">{t("mode_label")}</p>', unsafe_allow_html=True)
    m1, m2 = st.columns(2)
    is_single = S.bill_count != -1
    with m1:
        if st.button(f"{'✅ ' if is_single else ''}{t('mode_single')}\n_{t('mode_single_desc')}_",
                     use_container_width=True, key="mode_single_btn",
                     type="primary" if is_single else "secondary"):
            if not is_single:
                for k, v in _DEFAULTS.items(): S[k] = v
                st.rerun()
    with m2:
        is_batch = S.bill_count == -1
        if st.button(f"{'✅ ' if is_batch else ''}{t('mode_batch')}\n_{t('mode_batch_desc')}_",
                     use_container_width=True, key="mode_batch_btn",
                     type="primary" if is_batch else "secondary"):
            if not is_batch:
                for k, v in _DEFAULTS.items(): S[k] = v
                S.bill_count = -1
                st.rerun()
    st.divider()

    if S.bill_count == -1:
        run_batch_mode_ui()
        return

    render_steps()

    # 1 บิลต่อรูปเสมอ (ตัด 2/3 บิล ออก)
    if S.bill_count == 0:
        S.bill_count = 1
        S.step = 2
    st.divider()

    st.markdown(f'<p class="sec-header">{t("upload_label")}</p>', unsafe_allow_html=True)
    st.markdown(f'<div class="hint-box">{t("upload_hint")}</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader("files", type=["png","jpg","jpeg","heic"],
                                accept_multiple_files=True, label_visibility="collapsed",
                                key="uploader")
    if uploaded:
        new_files = [(f.name, f.read()) for f in uploaded]
        names_new = [x[0] for x in new_files]
        names_old = [x[0] for x in S.gallery_files]
        if names_new != names_old:
            S.gallery_files=new_files; S.selected_idx=-1
            S.crop_result=None; S.step=2
            S.manual_splits_px=None; S.crop_applied=False

    if S.gallery_files:
        st.markdown(f'<p class="sec-header">{t("gallery_label")}</p>', unsafe_allow_html=True)
        st.caption(t("gallery_count")(len(S.gallery_files)))
        cols_per_row = 5
        all_f = S.gallery_files
        for chunk_start in range(0, len(all_f), cols_per_row):
            chunk = all_f[chunk_start:chunk_start+cols_per_row]
            cols  = st.columns(len(chunk))
            for ci, (fname, fbytes) in enumerate(chunk):
                idx = chunk_start + ci
                with cols[ci]:
                    try:
                        pil = Image.open(io.BytesIO(fbytes))
                        st.image(pil, use_container_width=True)
                        is_sel = S.selected_idx == idx
                        label  = f"{'✅ ' if is_sel else ''}{fname[:14]}{'…' if len(fname)>14 else ''}"
                        if st.button(label, key=f"gal_{idx}", use_container_width=True,
                                     type="primary" if is_sel else "secondary"):
                            S.selected_idx=idx; S.crop_result=None; S.step=2
                            S.manual_splits_px=None; S.crop_applied=False
                            st.rerun()
                    except Exception:
                        st.warning(fname)
        if S.selected_idx < 0 and len(S.gallery_files) == 1:
            S.selected_idx = 0
        if S.selected_idx >= 0:
            fname = S.gallery_files[S.selected_idx][0]
            st.success(f"✅ เลือก: **{fname}**")

    active_bytes = None
    if 0 <= S.selected_idx < len(S.gallery_files):
        active_bytes = S.gallery_files[S.selected_idx][1]

    if active_bytes:
        st.divider()
        try:
            pil_orig = Image.open(io.BytesIO(active_bytes)).convert("RGB")
        except Exception as e:
            st.error(f"ไม่สามารถโหลดรูป: {e}"); return

        # ── ปรับภาพ ──────────────────────────────────────────────────────
        img_cv_raw = pil_to_cv(pil_orig)
        gray_raw   = cv2.cvtColor(img_cv_raw, cv2.COLOR_BGR2GRAY)

        enh_col, prev_col = st.columns([2, 1])
        with enh_col:
            st.markdown("**🖼️ ปรับภาพก่อน OCR**")
            sharp_val    = st.slider("ความคมชัด (Sharpen)", 0.0, 3.0, 1.0, 0.1, key="enh_sharp")
            contrast_val = st.slider("คอนทราสต์",           0.5, 3.0, 1.0, 0.1, key="enh_cont")
            bw_mode      = st.checkbox("ขาวดำ (Grayscale → ช่วย OCR)", value=True, key="enh_bw")

        # สร้างรูปที่ผ่านการปรับแล้ว
        from PIL import ImageEnhance, ImageFilter
        pil_enh = pil_orig.copy()
        if contrast_val != 1.0:
            pil_enh = ImageEnhance.Contrast(pil_enh).enhance(contrast_val)
        if sharp_val != 1.0:
            pil_enh = ImageEnhance.Sharpness(pil_enh).enhance(sharp_val)
        if bw_mode:
            pil_enh = pil_enh.convert("L").convert("RGB")

        with prev_col:
            st.markdown("**ตัวอย่าง**")
            st.image(pil_enh, use_container_width=True)

        S.crop_result = pil_enh
        S.crop_applied = True

    working_pil = S.crop_result
    if working_pil is None and active_bytes:
        working_pil = Image.open(io.BytesIO(active_bytes)).convert("RGB")

    if working_pil:
        st.divider()
        img_cv_preview = pil_to_cv(working_pil)

        # blur check
        gray_chk = cv2.cvtColor(img_cv_preview, cv2.COLOR_BGR2GRAY)
        blur_score = cv2.Laplacian(gray_chk, cv2.CV_64F).var()
        if blur_score < 30:
            st.warning(f"⚠️ รูปไม่ชัด (blur score: {blur_score:.0f}) — ผล OCR อาจไม่แม่น")

        engine_label = {"tesseract": "🆓 Tesseract", "gdrive": "📄 Google Drive OCR",
                        "vision": "🎯 Google Vision API"}.get(S.ocr_engine, S.ocr_engine)
        st.caption(f"Engine: **{engine_label}**")

        if st.button(t("analyze"), use_container_width=True, type="primary"):
                with st.spinner("กำลัง OCR..."):
                    img_cv = pil_to_cv(working_pil)
                    all_bills = []
                    fname = (S.gallery_files[S.selected_idx][0]
                             if 0 <= S.selected_idx < len(S.gallery_files) else "image")
                    progress = st.progress(0, text="เริ่มต้น OCR...")

                    # 1 บิลต่อรูปเสมอ
                    crops = [img_cv]

                    for ci, crop in enumerate(crops):
                        label = fname if len(crops)==1 else f"{fname} — บิล {ci+1}"
                        progress.progress((ci) / len(crops),
                            text=f"🔍 OCR บิล {ci+1}/{len(crops)}...")
                        st.session_state["_gdrive_raw_texts"] = []
                        text       = run_ocr(crop, engine=S.ocr_engine)
                        gdrive_raw = (st.session_state.get("_gdrive_raw_texts") or [""])[0]
                        text_for_gemini = gdrive_raw if (S.ocr_engine == "gdrive" and gdrive_raw) else text
                        if is_gemini_configured() and text_for_gemini.strip():
                            progress.progress((ci + 0.5) / len(crops),
                                text=f"✨ Gemini วิเคราะห์บิล {ci+1}/{len(crops)}...")
                            gemini_result = extract_with_gemini(
                                text_for_gemini,
                                ocr_source="gdrive" if S.ocr_engine == "gdrive" else "tesseract")
                            if gemini_result["ok"] and gemini_result["items"]:
                                bill  = gemini_result["bill"]
                                items = gemini_result["items"]
                            else:
                                bill  = extract_receipt(text)
                                items = extract_items_cj(text)
                        else:
                            bill  = extract_receipt(text)
                            items = extract_items_cj(text)
                        all_bills.append({"filename":label,"bill":bill,"items":items,
                                          "raw_text":text,"gdrive_raw":gdrive_raw,
                                          "image":img_to_bytes_png(crop)})
                    progress.progress(1.0, text=f"✅ OCR เสร็จสิ้น {len(crops)} บิล")

                    S.all_bills = all_bills; S.step=3; st.rerun()

    if S.all_bills:
        _render_bills_ui(S.all_bills, key_prefix="s")
        st.divider()
        dl_c, rs_c = st.columns([3,1])
        with dl_c:
            st.download_button(t("download"), data=build_excel(S.all_bills),
                           file_name="receipts.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
        with rs_c:
            if st.button(t("reset"), use_container_width=True):
                for k,v in _DEFAULTS.items(): S[k]=v
                st.rerun()
            if st.button("💾 บันทึกลง Google Sheets"):
        with st.spinner("กำลังบันทึกข้อมูล..."):
            if save_to_sheets(S.all_bills):
                st.success("บันทึกข้อมูลสำเร็จแล้ว!")
            if st.button("🔄 Refresh ข้อมูลจาก Sheet", key="refresh_sheet_cache"):
                st.session_state.pop("_sheet_correct_items_cache", None)
                st.success("✅ ล้าง cache แล้ว จะโหลดใหม่รอบหน้า")

if __name__ == "__main__":
    main()
