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
        steps=["เลือกจำนวน","อัปโหลด","Crop","OCR"],
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
        steps=["Count","Upload","Crop","OCR"],
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
_GDRIVE_SCOPES    = "https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/userinfo.email https://www.googleapis.com/auth/userinfo.profile"
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

# ─────────────────────────────────────────────────────────────────────────────
# Gemini API
# ─────────────────────────────────────────────────────────────────────────────
_GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent"

def _get_gemini_key() -> str:
    try:
        return st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))
    except Exception:
        return os.environ.get("GEMINI_API_KEY", "")

def is_gemini_configured() -> bool:
    return bool(_get_gemini_key().strip())

def _call_gemini(prompt: str, max_tokens: int = 1500) -> str:
    import requests as _req, time
    key = _get_gemini_key()
    if not key:
        raise RuntimeError("ยังไม่ได้ตั้งค่า GEMINI_API_KEY")
    for attempt in range(2):
        resp = _req.post(
            f"{_GEMINI_API_URL}?key={key}",
            json={
                "contents": [{"parts": [{"text": prompt}]}],
                "generationConfig": {"maxOutputTokens": max_tokens, "temperature": 0.1},
            },
            timeout=30,
        )
        if resp.status_code == 429:
            time.sleep(5)
            continue
        resp.raise_for_status()
        data = resp.json()
        return data["candidates"][0]["content"]["parts"][0]["text"]
    raise RuntimeError("Gemini rate limit — รอสักครู่แล้วลองใหม่")

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
OCR มักอ่านผิด: "Bao_", "Beo", "B80", "Be0", "Ba0","Bac." → ล้วนหมายถึง "Bao"
"Bao_อเมริกาโน่เป็น" → "Bao อเมริกาโน่เย็น" (_ = space, เป็น = เย็น)"""

CJ_PRODUCT_KNOWLEDGE = """สินค้าในร้าน CJ Express:
อาหาร: มาม่า ยูโรเค้ก ขนมปัง แซนด์วิช ไส้กรอก ข้าวกล่อง
ขนม: เฮอร์ชีย์ โอโชะ มันฝรั่ง ป๊อปคอร์น เยลลี่ ลูกอม
เครื่องดื่ม: น้ำดื่ม สไปรท์ โคคา เป๊ปซี่ คาราบาว ไมโล โออิชิ อราวน์
ของใช้: สบู่ แชมพู ยาสีฟัน ผ้าอนามัย ทิชชู ผงซักฟอก
สินค้าเบ็ดเตล็ด: บุหรี่ ไฟแช็ก ปากกา สมุด บัตรเติมเงิน"""

def _is_bao_item(name: str) -> bool:
    return bool(re.search(
        r'\bb[a8e๐o0][o๐0][\W_]|'
        r'\bb[a8e๐o0][o๐0]$|'
        r'\bbo[\W_]|'
        r'bao|'
        r'เบา\s*คา|บาว\s*คา',
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
- pos_id: ดึงจาก BNO format S+ปี(2)+เดือน(2)+สาขา(4)+N เช่น BNO:S26061326N03 → pos_id = "1326"
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
    if ocr_source == "gdrive":
        format_hint = """รูปแบบข้อมูล (Google Drive OCR แยกราคาเป็นบรรทัดใหม่):
- บรรทัด 1: จำนวน ชื่อสินค้า  (เช่น "1 Bao ลาเต้เย็น" หรือ "1 PFPเต้าหู้ 120g")
- บรรทัด 2: ราคาต่อหน่วย  (เช่น "40.00")
- บรรทัด 3: ราคารวม V  (เช่น "40.00 V" หรือ "40.00 - V" หรือ "80.00 V")

กรณีพิเศษที่พบบ่อย:
- "- 50.00" หรือ "40.00 - V" คือราคาของสินค้าบรรทัดก่อนหน้า ไม่ใช่ส่วนลด
- "หวานน้อย" "หวานปกติ" "ไม่หวาน" คือ note ข้ามไป ไม่ใช่สินค้า (อาจอยู่หลังราคาก็ได้)
- "-" หน้าชื่อสินค้าหมายถึงจำนวน 1
- บรรทัดที่มี "จํานวนสินค้ารวมXรายการ" ติดท้ายชื่อสินค้า ให้ตัดส่วนนั้นทิ้ง
  เช่น "1 Bao อาราบีก้า 45.00 จํานวนสินค้ารวมๆรายการ" → ชื่อ "Bao อาราบีก้า" ราคา 45.00
- OCR อาจรวมสินค้า 2 รายการไว้บรรทัดเดียว เช่น "1 สกินแม็บยาสีฟัน 1 Bigsmileแปรงคนจัดฟัน จานวนสินค้ารวม2รายการ"
  → แยกเป็น 2 items ราคาจับคู่ตามลำดับ: สินค้า 1 → ราคาชุดแรก, สินค้า 2 → ราคาชุดที่ 2
- OCR อ่านชื่อ Bao ผิดเป็น "Beo", "B80", "Be0", "Bo", "Bao_", "Bao.", "Bac" → แปลงเป็น "Bao" เสมอ
- ส่วนลด: บรรทัด "-40.00", "-20.00", "-16.00" ฯลฯ ให้ใส่เป็น item ชื่อ "ส่วนลด" ราคาติดลบ
  เช่น {{"ชื่อสินค้า":"ส่วนลด","หมวดหมู่":"ส่วนลด/โปรโมชั่น","จำนวน":1,"ราคาต่อหน่วย":-40.0,"ยอดรวมสินค้า":-40.0}}
- ถ้า raw text มี "จำนวนสินค้ารวม N รายการ" → ต้องได้ items ครบ N รายการ (ไม่นับส่วนลด)
- pos_id: ดึงจาก BNO เช่น BNO:S26061326N03 → pos_id = "1326", BNO:S26060707N01 → pos_id = "0707"
- ชื่อสินค้าที่ OCR อ่านผิดให้แก้เปรียบเทียบกับสินค้าในร้านสะดวกซื้อไทย:
  "ขาไทยเป็น/เย็น" → "ชาไทยเย็น", "ขาเขียวนมสด" → "ชาเขียวนมสด"
  "กาแฟบาวเป็น" → "Bao กาแฟ", "เอสเปรสโซ่เป็น" → "เอสเปรสโซ่เย็น"
  "เปาผงซักฟอก" → "เปา ผงซักฟอก", "เล คลาส คออริจินอล" → "เลย์คลาสสิค ออริจินอล"
  "เลยคลาสสดรสโนริ" → "เลย์คลาสสิค รสโนริสาหร่าย"
  "หากเก่นนมเปรี้ยว" → "ยาคูลท์นมเปรี้ยว", "เลิฟโทนเลิฟชาโยคา" → "เลิฟโพชั่น ชาโยคา"
  "TGMคอทเทจ เบค่อน" → "TGM คอทเทจ เบคอน", "เกาลูนส้นแก้วสาหร่าย" → "เกาลูน สาหร่ายทะเล"
  "ดราช่างเทปลบ" → "ดราช่าง เทปลบคำผิด", "MBLรองฟัน" → "MBL ฟลอสขัดฟัน"
  "เกลดเจลลาเวน" → "เจลลาเวนเดอร์", "สมูทอีเจลแต้มสิว" → "สมูทอี เจลแต้มสิว"
  "คูลคูลเฟรช" → "คูลคูล เฟรช น้ำมะพร้าว"
  "วอลสท็อปเท็นช็อก" → "Walltop ช็อกโกแลต"
  "เซนได กาจ่อปลวก" → "เซนได กำจัดปลวก"
  "PFPเต้าหู้" → "PFP เต้าหู้แข็ง"
  "อิชิตันน้าต่าง พีเอ" → "อิชิตัน น้ำตาง พลัส"
  "เอ ฟรีซดรายกล้วย" → "A-Freeze กล้วยอบกรอบ"
  "เรนเจอร์ยาจดกันยุงควัน" → "เรนเจอร์ ยากันยุงแบบจุด"
  "เคเอฟไข่ไก่" → "ไข่ไก่ขนาดกลาง"
  "นํ้าทิพยนาดิม" → "น้ำทิพย์ ขนาด 1500ml" """
    else:
        format_hint = """รูปแบบข้อมูล (แต่ละสินค้าอยู่บรรทัดเดียว):
- รูปแบบ: จำนวน ชื่อสินค้า ราคาต่อหน่วย ราคารวม  (เช่น "1 คาราบาวแดง 12.00 12.00")
- หรือ: จำนวน ชื่อสินค้า ราคา  (เช่น "2 น้ำดื่ม 10.00 20.00")
- OCR อาจรวมสินค้า 2 รายการและ "จำนวนสินค้ารวมXรายการ" ไว้บรรทัดเดียว → แยกเป็น 2 items
- "- 50.00" หรือ "40.00 - V" คือราคาของสินค้าบรรทัดก่อนหน้า ไม่ใช่ส่วนลด
- ส่วนลด: "-40.00", "-20.00" ฯลฯ ให้ใส่เป็น item ชื่อ "ส่วนลด" ราคาติดลบ
  เช่น {{"ชื่อสินค้า":"ส่วนลด","หมวดหมู่":"ส่วนลด/โปรโมชั่น","จำนวน":1,"ราคาต่อหน่วย":-40.0,"ยอดรวมสินค้า":-40.0}}
- OCR อ่านชื่อ Bao ผิดเป็น "Beo", "B80", "Be0", "Bo", "Bao_", "Bac" → แปลงเป็น "Bao" เสมอ
- ถ้า raw text มี "จำนวนสินค้ารวม N รายการ" → ต้องได้ items ครบ N รายการ (ไม่นับส่วนลด)
- pos_id: ดึงจาก BNO เช่น BNO:S26061326N03 → pos_id = "1326"
- ชื่อสินค้าที่ OCR อ่านผิดให้แก้เปรียบเทียบกับสินค้าในร้านสะดวกซื้อไทย:
  "ขาไทยเป็น" → "ชาไทยเย็น", "เอสเปรสโซ่เป็น" → "เอสเปรสโซ่เย็น"
  "เลย์คลาสลิค" → "เลย์คลาสสิค", "เลยคลาสสดรสโนริ" → "เลย์คลาสสิค รสโนริสาหร่าย"
  "เรนเจอร์ยาจดกันยุง" → "เรนเจอร์ ยากันยุง" """

    prompt = f"""ใบเสร็จ CJ Express:

{raw_text}

{format_hint}

{BAO_CAFE_MENU}
{CJ_PRODUCT_KNOWLEDGE}

ตอบ JSON เท่านั้น:
{{
  "date": "วัน/เดือน/ปี",
  "time": "HH:MM",
  "branch": "ชื่อสาขา",
  "pos_id": "รหัสสาขา 4 หลัก จาก BNO",
  "pos_machine": "NXX",
  "rcpt_no": "เลขที่ใบเสร็จ",
  "total_amount": 0.0,
  "cash": 0.0,
  "change": 0.0,
  "items": [{{"ชื่อสินค้า":"","หมวดหมู่":"","จำนวน":1,"ราคาต่อหน่วย":0.0,"ยอดรวมสินค้า":0.0}}]
}}

{CATEGORY_PROMPT}

กฎ:
- Bao_, Beo, B80, Be0, Ba0, Bao. → "Bao" + หมวด "Bao Cafe"
- โปรโมชั่น/ส่วนลด → ใส่เป็น item หมวด "ส่วนลด/โปรโมชั่น" ราคาติดลบ
- "1 โปรโมชั่นM 1 แถม 1 Bao" + "-40.00" → item "โปรโมชั่น 1 แถม 1 Bao" ราคา -40.0
- จำนวนสินค้ารวม N รายการ → ต้องได้ items ครบ N"""

    try:
        text_out = _call_gemini(prompt)
        text_out = re.sub(r"```(?:json)?\s*|\s*```", "", text_out).strip()
        data = json.loads(text_out)
        bill = {
            "date":         str(data.get("date", "ไม่พบ")),
            "time":         str(data.get("time", "ไม่พบ")),
            "branch":       str(data.get("branch", "ไม่พบ")),
            "pos_id":       str(data.get("pos_id", "ไม่พบ")),
            "pos_machine":  str(data.get("pos_machine", "ไม่พบ")),
            "rcpt_no":      str(data.get("rcpt_no", "ไม่พบ")),
            "total_amount": float(data.get("total_amount", 0)),
            "cash":         0.0, "change": 0.0,
            "tax_id":       "ไม่พบ", "user": "ไม่พบ", "name": "ไม่พบ",
        }
        items = []
        for it in data.get("items", []):
            name = str(it.get("ชื่อสินค้า", ""))
            if _is_bao_item(name):
                cat = BAO_CAFE_CATEGORY
            else:
                cat = str(it.get("หมวดหมู่", ""))
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
        return {"bill": bill, "items": items, "ok": True}
    except Exception as e:
        return {"bill": None, "items": [], "ok": False, "error": str(e)}

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
    return '\n'.join(final)

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

def _merge_gdrive_lines(lines: list) -> list:
    _price_only = re.compile(r'^[\d,.\s]+[Vv]?\s*$')
    _v_only     = re.compile(r'^[Vv]\s*$')
    _has_thai   = re.compile(r'[ก-๙]')
    _item_start = re.compile(r'^(?:\d+|-)\s*\S')
    _kw_amount  = re.compile(
        r'ยอดรวม|ยอดราม|ยอดราเม|UORTIN|UORT|เงินสด|เงินเด|ในสต|เงินทอน|เงินบน|รวมทั้งสิ้น|บอดราม',
        re.IGNORECASE)
    _skip_note  = re.compile(r'^(หวานน้อย|ลดน้ำตาล|ไม่หวาน|หวานปกติ|extra\s*shot)', re.IGNORECASE)
    _kw_count   = re.compile(r'จ[ํา]?นวนสินค[้า]?[่า]?\s*รวม', re.IGNORECASE)

    lines = [l for l in lines if not _v_only.match(l.strip())]

    def _merge_kw_price(lines):
        out = []
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            # ตัด "จำนวนสินค้ารวมXรายการ" ออกจากท้ายบรรทัด
            line = re.sub(
                r'จ[ํา]?นวนสินค[้า]?[่า]?\s*รวม[^บ]*รายการ', '', line
            ).strip()
            if _kw_amount.search(line) and not _RE_PRICE.search(line):
                j = i + 1
                if j < len(lines):
                    nx = lines[j].strip()
                    if _RE_PRICE.search(nx):
                        out.append(line + " " + nx)
                        i = j + 1
                        continue
            if line:
                out.append(line)
            i += 1
        return out

    def _find_price_block(lines):
        """ตรวจหา block ราคาที่แยกจากชื่อสินค้า — รองรับทั้ง 2+ และ 4+ บรรทัด"""
        consec = 0
        block_start = None
        for i, line in enumerate(lines):
            s = line.strip()
            if _price_only.match(s) and s:
                if consec == 0:
                    block_start = i
                consec += 1
                # threshold ต่ำ = 2 บรรทัดราคาติดกัน + มีชื่อสินค้าข้างบน
                if consec >= 2:
                    name_only_count = sum(
                        1 for l in lines[:block_start]
                        if _item_start.match(l.strip())
                        and _has_thai.search(l)
                        and not _RE_PRICE.search(l)
                    )
                    if name_only_count >= 1:
                        return block_start
            else:
                consec = 0
                block_start = None
        return None

    lines = _merge_kw_price(lines)
    lines = [re.sub(r'\b(\d{1,2}):(\d{2})\b',
                    lambda m: f"{m.group(1)}.{m.group(2)}", l) for l in lines]

    price_block_start = _find_price_block(lines)

    if price_block_start and price_block_start > 0:
        name_lines  = lines[:price_block_start]
        price_lines = lines[price_block_start:]

        item_names = [l.strip() for l in name_lines
                      if l.strip()
                      and not _price_only.match(l.strip())
                      and not _kw_amount.search(l.strip())
                      and not _kw_count.search(l.strip())
                      and (_item_start.match(l.strip()) or _has_thai.search(l.strip()))
                      and not re.search(r'BNO|TAX|VAT|POS|User|ID:', l, re.IGNORECASE)
                      and not re.match(r'\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}', l.strip())
                      and not _skip_note.match(l.strip())]

        raw_prices = []
        for l in price_lines:
            l = l.strip()
            if not l: continue
            if _kw_amount.search(l): break
            if _kw_count.search(l): continue
            clean = re.sub(r'\s*[Vv]\s*$', '', l).strip()
            # จับราคาปกติ และ "- 50.00" (ราคาที่มี - นำหน้า)
            if _price_only.match(l) and clean:
                # แปลง "- 50.00" → "50.00"
                clean = re.sub(r'^-\s*', '', clean).strip()
                if not raw_prices or raw_prices[-1] != clean:
                    raw_prices.append(clean)

        merged = []
        for i, name in enumerate(item_names):
            if i < len(raw_prices):
                merged.append(f"{name} {raw_prices[i]}")
            else:
                merged.append(name)

        # เพิ่มส่วนลดที่อยู่ใน price_lines
        for l in price_lines:
            l = l.strip()
            if not l: continue
            if _kw_amount.search(l): break
            m_disc = re.match(r'^-\s*(\d+[.,]\d{2})\s*$', l)
            if m_disc:
                merged.append(f"1 ส่วนลด -{m_disc.group(1)}")

        after_block = []
        found_kw = False
        for l in price_lines:
            if _kw_amount.search(l.strip()):
                found_kw = True
            if found_kw:
                after_block.append(l)
        merged += after_block

        header = [l for l in name_lines
                  if l.strip() and l.strip() not in item_names]
        return header + merged

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
                        # ข้าม note หลังราคา
                        while j < len(lines) and _skip_note.match(lines[j].strip()):
                            j += 1
                    elif re.match(r'^-\s*\d+[.,]\d{2}\s*[Vv]?\s*$', nx):
                        part = re.sub(r'^-\s*', '', nx)
                        part = re.sub(r'\s*[Vv]\s*$', '', part).strip()
                        if part: combined += " " + part
                        price_count += 1
                        j += 1
                    elif price_count == 0 and _has_thai.search(nx) and not _kw_amount.search(nx):
                        combined += " " + nx
                        j += 1
                    else:
                        break
                if price_count > 0:
                    merged.append(combined); i = j; continue
            merged.append(line); i += 1
        return merged

# ── FIX: extract_items_cj ครบทุกจุด ──
def extract_items_cj(text: str) -> list:
    items  = []
    lines  = text.split('\n')
    lines = _merge_gdrive_lines(lines)

    start_idx = 0
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
               "www.","FB:","ร้องเรียน","สมัคร"]

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
            compact = _collapse(line)

            if re.search(r'จ[าำ]?นวนสินค[้า]?[่า]?\s*รวม', line): break  # จำนวนสินค้ารวม
            if re.search(r'จ.{0,3}นวนสินค.{0,3}รวม', compact): break
            if re.search(r'สา[นม]านสินค', compact): break
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

            m = _fb_b.match(lf)
            if m:
                qty_raw, nm, price = m.groups()
                qty = 1 if qty_raw == '-' else int(qty_raw)
                nm = _clean_item_name(nm)
                if len(re.sub(r'[^\u0E00-\u0E7FA-Za-z0-9]', '', nm)) >= 2:
                    p = parse_price(price)
                    items.append(_make_item(nm, qty, p, round(p * qty, 2)))
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
            base = {"ไฟล์":b['filename'],"วันที่":d['date'],"เวลา":d['time'],
                    "สาขา":d['branch'],"รหัสสาขา":d['pos_id'],
                    "POS ID":d.get('pos_machine',''),
                    "เลขที่ใบเสร็จ":d['rcpt_no']}
            items = b['items'] or []
            for it in items:
                cat = BAO_CAFE_CATEGORY if _is_bao_item(it.get("ชื่อสินค้า","")) \
                      else (it.get("หมวดหมู่") or "สินค้าเบ็ดเตล็ดอื่นๆ")
                row = base.copy()
                row.update({"ชื่อสินค้า":   it.get('ชื่อสินค้า',''),
                            "หมวดหมู่":     cat,
                            "จำนวน":        it.get('จำนวน',1),
                            "ราคาต่อหน่วย": it.get('ราคาต่อหน่วย',0),
                            "ยอดรวมสินค้า": it.get('ยอดรวมสินค้า',0),
                            "ยอดรวม":       ""})
                rows.append(row)
            summary = base.copy()
            summary.update({"ชื่อสินค้า":"","หมวดหมู่":"","จำนวน":"",
                            "ราคาต่อหน่วย":"","ยอดรวมสินค้า":"",
                            "ยอดรวม":d['total_amount']})
            rows.append(summary)
        pd.DataFrame(rows).to_excel(writer, index=False, sheet_name='ใบเสร็จ')
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
                    st.image(Image.open(io.BytesIO(fbytes)), use_container_width=True,
                              caption=fname[:16] + ('…' if len(fname)>16 else ''))
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

def _render_bills_ui(all_bills, key_prefix=""):
    st.success(t("found")(len(all_bills)))
    for idx, b in enumerate(all_bills):
        with st.expander(f"📄 {b['filename']}", expanded=True):
            ic, dc = st.columns([1,2])
            with ic:
                if b.get('image'): st.image(b['image'], use_container_width=True)
            with dc:
                d = b['bill']
                st.markdown("**ตรวจสอบ / แก้ไข**")
                r1 = st.columns(5)
                d['date']        = r1[0].text_input("วันที่",         d['date'],               key=f"{key_prefix}dt{idx}")
                d['time']        = r1[1].text_input("เวลา",           d['time'],               key=f"{key_prefix}tm{idx}")
                d['pos_id']      = r1[2].text_input("รหัสสาขา",      d['pos_id'],             key=f"{key_prefix}ps{idx}")
                d['pos_machine'] = r1[3].text_input("POS ID",         d.get('pos_machine',''), key=f"{key_prefix}pm{idx}")
                d['rcpt_no']     = r1[4].text_input("เลขที่ใบเสร็จ", d['rcpt_no'],            key=f"{key_prefix}rc{idx}")
                tot_str = st.text_input("ยอดรวม", f"{float(d['total_amount']):.2f}", key=f"{key_prefix}tot{idx}")
                d['total_amount'] = parse_price(tot_str)
            st.metric("💰 ยอดรวม", f"{d['total_amount']:.2f} ฿")
            if b['items']:
                st.markdown("**🛒 รายการสินค้า**")
                items_display = b['items'].copy()
                for i, it in enumerate(items_display):
                    if _is_bao_item(it.get("ชื่อสินค้า", "")):
                        items_display[i] = {**it, "หมวดหมู่": BAO_CAFE_CATEGORY}
                    elif not it.get("หมวดหมู่"):
                        items_display[i] = {**it, "หมวดหมู่": "สินค้าเบ็ดเตล็ดอื่นๆ"}
                st.dataframe(pd.DataFrame(items_display), use_container_width=True, hide_index=True)
            else:
                st.info(t("no_items"))
            with st.expander(f"🔬 {t('raw_text')} + debug"):
                gdrive_raws = st.session_state.get("_gdrive_raw_texts", [])
                if S.ocr_engine == "gdrive" and idx < len(gdrive_raws):
                    st.markdown("**📄 ข้อความดิบจาก Google Doc (ก่อน clean)**")
                    st.text_area("Google Doc raw", gdrive_raws[idx], height=200,
                                 key=f"{key_prefix}gdraw{idx}", disabled=True)
                    st.markdown("---")
                    st.markdown("**🧹 หลัง clean_text**")
                st.text_area("Raw OCR (cleaned)", b['raw_text'], height=160,
                             key=f"{key_prefix}raw{idx}", disabled=True)
                st.json({k:v for k,v in b['bill'].items()})

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

    st.markdown(f'<p class="sec-header" style="margin-top:1rem">{t("count_label")}</p>',
                unsafe_allow_html=True)
    cc1, cc2, cc3 = st.columns(3)
    for n, col in [(1,cc1),(2,cc2),(3,cc3)]:
        with col:
            is_on = S.bill_count == n
            if st.button(
                f"{'✅ ' if is_on else ''}{t(f'b{n}')}\n_{t(f'b{n}s')}_",
                use_container_width=True, key=f"cnt{n}",
                type="primary" if is_on else "secondary"):
                S.bill_count=n; S.step=2
                S.gallery_files=[]; S.selected_idx=-1
                S.crop_result=None; S.all_bills=[]
                S.manual_splits_px=None; S.crop_applied=False
                st.rerun()
    if S.bill_count == 0:
        st.info("👆 กรุณาเลือกจำนวนบิลก่อน")
        return
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
                            S.selected_idx=idx; S.crop_result=None; S.step=3
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
        st.markdown(f'<p class="sec-header">{t("crop_label")}</p>', unsafe_allow_html=True)
        try:
            pil_orig = Image.open(io.BytesIO(active_bytes)).convert("RGB")
        except Exception as e:
            st.error(f"ไม่สามารถโหลดรูป: {e}"); return

        crop_col, ctrl_col = st.columns([3,1])
        with ctrl_col:
            mode_choice = st.radio("โหมด Crop", [t("crop_free"), t("crop_a4")], key="crop_mode_r")
            mode_key = "a4" if t("crop_a4") in mode_choice else "free"
            st.info(t("crop_hint"))
        with crop_col:
            st_html(crop_component_html(pil_orig, mode_key, bill_count=S.bill_count),
                height=440 if S.bill_count > 1 else 400, scrolling=False)

        st.markdown("**📋 วางข้อมูล Crop ที่คัดลอกจากด้านบน:**")
        paste_col, btn_col = st.columns([4,1])
        crop_json_str = paste_col.text_input(
            "JSON ข้อมูล Crop", value="", key="crop_json_input",
            label_visibility="collapsed",
            placeholder='วาง {"x":0,"y":0,"w":800,"h":600,"splits":[300,560]} ที่นี่')

        w0, h0 = pil_orig.size
        cx, cy, cw, ch, parsed_splits = 0, 0, w0, h0, []
        if crop_json_str.strip():
            try:
                data = json.loads(crop_json_str.strip())
                cx = int(data.get("x", 0)); cy = int(data.get("y", 0))
                cw = int(data.get("w", w0)); ch = int(data.get("h", h0))
                parsed_splits = [int(p) for p in data.get("splits", [])]
            except Exception:
                st.warning("⚠️ รูปแบบข้อมูลไม่ถูกต้อง")

        st.caption("หรือระบุพื้นที่ Crop ด้วยตัวเลขเอง:")
        mc = st.columns(4)
        cx = mc[0].number_input("X (px)", 0, w0, cx, key="cx")
        cy = mc[1].number_input("Y (px)", 0, h0, cy, key="cy")
        cw = mc[2].number_input("W (px)", 1, w0, cw if cw>0 else w0, key="cw")
        ch = mc[3].number_input("H (px)", 1, h0, ch if ch>0 else h0, key="ch")

        if S.bill_count > 1:
            st.caption("จุดตัดบิล (พิกเซล X — คั่นด้วยจุลภาค):")
            default_splits_str = ", ".join(str(p) for p in parsed_splits) if parsed_splits else ""
            splits_str = st.text_input("จุดตัดบิล", value=default_splits_str, key="splits_input",
                                        label_visibility="collapsed",
                                        placeholder="เช่น 300, 560  (เว้นว่าง = ให้ระบบเดาเอง)")
        else:
            splits_str = ""

        bc1, bc2 = st.columns(2)
        with bc1:
            if st.button(t("crop_confirm"), use_container_width=True, type="primary"):
                S.crop_result = pil_orig.crop((cx, cy, min(cx+cw,w0), min(cy+ch,h0)))
                S.crop_applied = True
                manual_splits = []
                if splits_str.strip():
                    try:
                        manual_splits = [int(x.strip()) - cx for x in splits_str.split(",") if x.strip()]
                        manual_splits = [p for p in manual_splits if 0 < p < cw]
                    except Exception:
                        manual_splits = []
                elif parsed_splits:
                    manual_splits = [p - cx for p in parsed_splits if 0 < (p - cx) < cw]
                S.manual_splits_px = manual_splits if manual_splits else None
                S.step=4; st.rerun()
        with bc2:
            if st.button(t("crop_skip"), use_container_width=True):
                S.crop_result = pil_orig
                S.crop_applied = False
                S.manual_splits_px = None
                S.step=4; st.rerun()

    working_pil = S.crop_result
    if working_pil is None and active_bytes and S.step == 4:
        working_pil = Image.open(io.BytesIO(active_bytes)).convert("RGB")

    if working_pil:
        st.divider()
        prev_c, ocr_c = st.columns([1,2])
        with prev_c:
            st.image(working_pil, caption="รูปที่จะวิเคราะห์", use_container_width=True)
        with ocr_c:
            engine_label = {"tesseract": "🆓 Tesseract", "gdrive": "📄 Google Drive OCR",
                            "vision": "🎯 Google Vision API"}.get(S.ocr_engine, S.ocr_engine)
            st.caption(f"OCR Engine: **{engine_label}**")

            if st.button(t("analyze"), use_container_width=True, type="primary"):
                with st.spinner("กำลัง OCR..."):
                    img_cv = pil_to_cv(working_pil)
                    all_bills = []
                    fname = (S.gallery_files[S.selected_idx][0]
                             if 0 <= S.selected_idx < len(S.gallery_files) else "image")
                    progress = st.progress(0, text="เริ่มต้น OCR...")

                    if S.ocr_engine == "gdrive" and S.bill_count > 1:
                        if S.manual_splits_px:
                            crops = split_by_positions(img_cv, S.manual_splits_px)
                        else:
                            crops = split_receipts_image(img_cv, n_expected=S.bill_count)
                        if len(crops) < S.bill_count:
                            crops = crops + [img_cv] * (S.bill_count - len(crops))
                        crops = crops[:S.bill_count]

                        for ci, crop in enumerate(crops):
                            label = f"{fname} — บิล {ci+1}"
                            progress.progress(ci / len(crops),
                                text=f"📤 Google Drive OCR บิล {ci+1}/{len(crops)}...")
                            st.session_state["_gdrive_raw_texts"] = []
                            text       = run_ocr(crop, engine="gdrive")
                            gdrive_raw = (st.session_state.get("_gdrive_raw_texts") or [""])[0]
                            text_for_gemini = gdrive_raw if gdrive_raw else text
                            if is_gemini_configured() and text_for_gemini.strip():
                                progress.progress((ci + 0.5) / len(crops),
                                    text=f"✨ Gemini วิเคราะห์บิล {ci+1}/{len(crops)}...")
                                gemini_result = extract_with_gemini(text_for_gemini, ocr_source="gdrive")
                                if gemini_result["ok"] and gemini_result["items"]:
                                    bill  = gemini_result["bill"]
                                    items = gemini_result["items"]
                                else:
                                    bill  = extract_receipt(text)
                                    items = extract_items_cj(text)
                            else:
                                bill  = extract_receipt(text)
                                items = extract_items_cj(text)
                            all_bills.append({"filename": label, "bill": bill, "items": items,
                                              "raw_text": text, "gdrive_raw": gdrive_raw,
                                              "image": img_to_bytes_png(crop)})
                        progress.progress(1.0, text=f"✅ OCR เสร็จสิ้น {len(crops)} บิล")

                    else:
                        if S.bill_count == 1:
                            crops = [img_cv]
                        elif S.manual_splits_px:
                            crops = split_by_positions(img_cv, S.manual_splits_px)
                        else:
                            crops = split_receipts_image(img_cv, n_expected=S.bill_count)
                        if len(crops) < S.bill_count:
                            crops = crops + [img_cv] * (S.bill_count - len(crops))
                        crops = crops[:S.bill_count]

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

                    S.all_bills = all_bills; S.step=4; st.rerun()

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

if __name__ == "__main__":
    main()
