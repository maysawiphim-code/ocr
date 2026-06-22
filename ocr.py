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

# ─────────────────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────────────────
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

# ─────────────────────────────────────────────────────────────────────────────
# Session state
# ─────────────────────────────────────────────────────────────────────────────
_DEFAULTS = dict(step=1, bill_count=0,
                 gallery_files=[], selected_idx=-1,
                 crop_result=None, all_bills=[], lang="th",
                 manual_splits_px=None,   # ตำแหน่งตัดบิลที่ user ปรับเอง (พิกเซลในรูปที่ครอปแล้ว)
                 crop_applied=False,      # True เมื่อ user เพิ่งกด "ยืนยัน Crop" (ไม่ใช่ข้าม)
                 batch_files=[],          # ไฟล์สำหรับโหมด Batch (1 รูป = 1 บิล วิเคราะห์ทุกรูปพร้อมกัน)
                 ocr_engine="tesseract")  # "tesseract" (ฟรี) หรือ "vision" (Google Cloud Vision API)

for k, v in _DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v
S = st.session_state

# ─────────────────────────────────────────────────────────────────────────────
# UI Text (bilingual)
# ─────────────────────────────────────────────────────────────────────────────
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

# ─────────────────────────────────────────────────────────────────────────────
# Step bar
# ─────────────────────────────────────────────────────────────────────────────
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
    """
    หาตำแหน่งกระดาษใบเสร็จในภาพ แยกออกจากพื้นหลัง (โต๊ะไม้/พื้นสี/เงา ฯลฯ)
    ใช้ HSV color space: กระดาษขาวมี saturation ต่ำ + brightness สูง
    ต่างจากพื้นหลังที่มักมีสี (saturation สูงกว่า) หรือมืดกว่า
    คืนค่า: list ของ (bbox, contour) เรียงจากซ้ายไปขวา, และ mask ทั้งภาพ
    """
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
    results.sort(key=lambda r: r[0][0])  # เรียงซ้าย→ขวา
    return results, mask


def deskew_paper(img_cv, contour, pad=10):
    """
    แก้มุมเอียงของภาพ (perspective correction): หา minAreaRect ของ contour
    กระดาษ แล้ว warp ให้กลายเป็นสี่เหลี่ยมตรง ไม่ว่าจะถ่ายเอียงแค่ไหน
    คืนค่า: (ภาพที่ deskew แล้ว BGR, mask ของกระดาษในพิกัดใหม่หลัง warp)
    """
    H, W = img_cv.shape[:2]
    rect = cv2.minAreaRect(contour)
    (cx, cy), (rw, rh), angle = rect
    # normalize: ให้ rw < rh เสมอ (ใบเสร็จสูงกว่ากว้างเป็นปกติ)
    if rw > rh:
        rw, rh = rh, rw
        angle += 90

    box = cv2.boxPoints(rect).astype(np.float32)
    out_w, out_h = max(int(rw) + pad * 2, 10), max(int(rh) + pad * 2, 10)

    # เรียงจุด 4 มุมให้ตรงลำดับ: top-left, top-right, bottom-right, bottom-left
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
    """
    แก้แสง/เงาที่ตกไม่สม่ำเสมอบนกระดาษ (เช่น เงามือถ่ายภาพ, แสงไฟด้านเดียว)
    หลักการ background-subtraction: ประมาณ "แนวโน้มแสง" ด้วย morphological
    closing ขนาดใหญ่ (มองข้ามตัวอักษร เห็นแค่ความสว่างพื้นหลังกว้างๆ)
    แล้วหารภาพต้นฉบับด้วยค่านั้น ทำให้แสงสม่ำเสมอทั้งภาพโดยไม่กระทบตัวอักษร
    """
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
    """
    เตรียมภาพใบเสร็จให้อ่านง่ายที่สุดก่อน OCR รวม 3 การแก้ไขในฟังก์ชันเดียว:
    1. แก้มุมเอียง (ถ้ามี contour และ deskew=True) — perspective correction
    2. ทำพื้นหลังขาวบริสุทธิ์ (ตัดส่วนนอกรูปทรงกระดาษทิ้ง)
    3. แก้แสง/เงาไม่สม่ำเสมอ + ลด noise จากรอยพับ/ยับเบาๆ (bilateral filter)
    คืนค่า: ภาพขาวดำ (grayscale) พร้อมส่งต่อให้ preprocess_image/OCR
    """
    H, W = img_cv.shape[:2]

    if contour is not None and deskew:
        # มีรูปทรงกระดาษชัดเจน → แก้มุมเอียง + ตัดพื้นหลังนอกขอบในขั้นตอนเดียว
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

    # แก้แสง/เงาก่อน (สำคัญที่สุดสำหรับภาพที่มีเงาทับ)
    gray = correct_illumination(gray)
    # ลด noise จากรอยพับ/ยับเบาๆ — รักษาขอบตัวอักษรไว้ ไม่เบลอเหมือน Gaussian blur
    gray = cv2.bilateralFilter(gray, d=5, sigmaColor=30, sigmaSpace=30)
    # เพิ่ม contrast เบาๆ ปิดท้าย
    clahe = cv2.createCLAHE(clipLimit=1.5, tileGridSize=(8, 8))
    return clahe.apply(gray)


def auto_crop_receipts(img_cv, n_expected=None, deskew=True):
    """
    ตรวจจับกระดาษใบเสร็จในภาพอัตโนมัติ + แก้มุมเอียง + ครอป ในขั้นตอนเดียว
    ใช้แทน split_receipts_image() เมื่อภาพมีพื้นหลังที่ไม่ใช่สีขาว/มีลวดลาย
    (เช่น ถ่ายบนโต๊ะไม้ พื้นกระเบื้อง ผ้าปูโต๊ะลาย ฯลฯ) หรือถ่ายเอียง
    คืนค่า: list ของภาพ BGR สี (deskew แล้วถ้า deskew=True) สำหรับส่งต่อ OCR
    """
    results, mask = find_paper_contours(img_cv)
    h, w = img_cv.shape[:2]

    if not results:
        return [img_cv]  # หาไม่เจอ → คืนภาพเดิมทั้งใบ ให้ fallback ไป split ปกติ

    # ถ้ารู้จำนวนบิลที่คาดไว้ แต่หาได้มาก/น้อยกว่า ให้ใช้ภาพเดิมแทน (ป้องกัน false positive)
    if n_expected and len(results) != n_expected:
        if abs(len(results) - n_expected) > 1:
            return [img_cv]

    crops = []
    for bbox, contour in results:
        if deskew:
            warped, warped_mask = deskew_paper(img_cv, contour, pad=10)
            # ทำพื้นหลังนอกขอบกระดาษเป็นขาว (composite บน 3-channel)
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
    """
    เตรียมภาพก่อนส่งเข้า OCR:
    1. (ถ้า whiten=True) แก้มุมเอียง + ทำพื้นหลังขาว + แก้แสงเงา + ลด noise รอยยับ
    2. ขยาย 2x + denoise + sharpen (เดิม)
    3. adaptive threshold → ขาวดำคมชัด พร้อม OCR

    หมายเหตุ: เมื่อเรียกจาก run_ocr() ตรงๆ (ไม่ผ่าน auto_crop_receipts ก่อน)
    จะไม่มี contour ให้ deskew ได้ — ฟังก์ชันนี้จะข้ามขั้นตอนแก้มุมเอียงอัตโนมัติ
    แต่ยังคงแก้แสงเงา/ลด noise ได้ตามปกติ
    """
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
# _find_bill_splits — v2: รองรับช่องว่างแคบมาก (1mm–1cm)
# ─────────────────────────────────────────────────────────────────────────────
def _find_bill_splits(col_ratio, w, min_bill_frac=0.15, n_expected=None):
    """
    หาตำแหน่งแบ่งบิล — 2 ขั้นตอน:
    1. หา outer bounds ของบริเวณที่มีเนื้อหา (ตัดพื้นหลังออก)
    2. หา split ภายใน outer bounds โดยมอง local-minima ของความสว่าง
       ใช้ smoothing kernel เล็กลง (5px) เพื่อให้ไวต่อช่องว่างแคบ ~1mm-1cm
       (ที่ DPI ทั่วไปของภาพมือถือ ~150-300dpi, 1mm ≈ 6-12px)
    n_expected: ถ้าระบุจำนวนบิลที่คาดไว้ (2 หรือ 3) จะพยายามหา split ให้ครบ
                จำนวนนั้นโดยลด threshold ลงเรื่อยๆ ถ้ายังไม่ครบ
    """
    # smoothing kernel เล็กลงจากเดิม (10px) เป็น 5px ให้ไวต่อช่องว่างแคบ
    kernel = np.ones(5) / 5
    smooth = np.convolve(col_ratio, kernel, mode='same')

    # ── ขั้น 1: หา outer bounds ──
    thr_outer = max(0.03, smooth.max() * 0.05)
    content_cols = np.where(smooth > thr_outer)[0]
    if len(content_cols) == 0:
        return [], smooth, smooth.mean()
    x0 = max(content_cols[0]  - 5, 0)
    x1 = min(content_cols[-1] + 5, w-1)
    region_w = x1 - x0
    if region_w <= 0:
        return [], smooth, smooth.mean()

    # ── ขั้น 2: หา splits ภายใน outer region ──
    seg = smooth[x0:x1]
    max_v = seg.max()

    def _candidates(gap_ratio):
        """หา local-minima candidates ทั้งหมดที่ต่ำกว่า threshold พร้อมค่าความสว่าง"""
        gap_thr = max_v * gap_ratio
        grad = np.gradient(seg)
        return [(i, float(seg[i])) for i in range(3, len(seg)-3)
                if grad[i-1] <= 0 and grad[i+1] >= 0 and seg[i] < gap_thr]

    def _select_best(candidates, min_bw, n_keep=None):
        """
        เลือก split points จาก candidates โดย greedy เริ่มจากจุดที่มืดที่สุด (gap ชัดที่สุด) ก่อน
        แล้วค่อยรับจุดถัดไปที่ห่างจากจุดที่เลือกแล้ว >= min_bw เท่านั้น
        วิธีนี้ป้องกันการเลือก noise ภายในบิลเดียวกันที่บังเอิญอยู่ใกล้ขอบภาพ
        แต่ไม่ใช่ gap จริงระหว่างบิล
        """
        ordered = sorted(candidates, key=lambda x: x[1])  # มืดสุดก่อน
        chosen = []
        for pos, val in ordered:
            if pos < min_bw or pos > region_w - min_bw: continue
            if all(abs(pos - p) >= min_bw for p in chosen):
                chosen.append(pos)
            if n_keep and len(chosen) >= n_keep: break
        return sorted(chosen)

    # พยายามหาด้วย threshold มาตรฐานก่อน
    min_bw = max(int(region_w * min_bill_frac), 8)
    target_n = (n_expected - 1) if n_expected else None
    candidates = _candidates(0.55)
    splits_rel = _select_best(candidates, min_bw, n_keep=target_n)

    # ถ้าระบุจำนวนบิลที่คาดไว้ และยังหาไม่ครบ → ลด threshold (gap_ratio) ขึ้นเรื่อยๆ
    # เพื่อรองรับช่องว่างที่แคบมากจนความสว่างไม่ลดลงมาก แต่ยังคงเลือกจาก
    # จุดที่มืดที่สุดก่อนเสมอ (ไม่ใช่จุดแรกที่เจอ) เพื่อไม่ให้หยิบ noise มาผิดจุด
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
    """แยกบิลจากภาพ รองรับบิลเดี่ยว/หลายบิล ช่องว่างตั้งแต่ ~1มม. และพื้นหลังทุกแบบ"""
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
    """
    ตัดภาพตามตำแหน่ง x (พิกเซล) ที่ระบุไว้ตรงๆ
    ใช้เมื่อ user ปรับเส้นตัดเองในขั้นตอน Crop แล้ว — ไม่ต้องเดาอีก
    """
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
# Google Cloud Vision API — OCR engine ทางเลือก (แม่นกว่า Tesseract มาก
# โดยเฉพาะกับภาพถ่ายมือถือที่เอียง/เงา/ยับ เพราะเทรนด้วยภาพถ่ายจริงจำนวนมาก
# ไม่ใช่แค่ font เคลียร์ๆ แบบที่ Tesseract ถนัด)
#
# วิธีตั้งค่า (ทำครั้งเดียว):
# 1. ไปที่ https://console.cloud.google.com/apis/library/vision.googleapis.com
#    เปิดใช้งาน "Cloud Vision API" สำหรับโปรเจกต์ของคุณ
# 2. สร้าง API key ที่ https://console.cloud.google.com/apis/credentials
#    (กด "Create Credentials" → "API key") — คัดลอกคีย์ที่ได้
# 3. ตั้งค่า environment variable ก่อนรันแอป:
#      Windows (PowerShell):  $env:GOOGLE_VISION_API_KEY="your-key-here"
#      Windows (cmd):         set GOOGLE_VISION_API_KEY=your-key-here
#      Mac/Linux:              export GOOGLE_VISION_API_KEY="your-key-here"
#    หรือใส่ตรงๆ ในตัวแปร _VISION_API_KEY ด้านล่างนี้แทนก็ได้ (ง่ายกว่าแต่
#    ไม่ควร commit ขึ้น git เพราะคีย์จะรั่วไหล)
#
# Free tier: 1,000 หน่วยแรกของทุกเดือนใช้ฟรี (DOCUMENT_TEXT_DETECTION นับ
# เป็น 1 หน่วยต่อภาพ) เกินจากนั้นคิดราคาถูกมาก ดูราคาล่าสุดที่
# https://cloud.google.com/vision/pricing
# ─────────────────────────────────────────────────────────────────────────────
_VISION_API_KEY = os.environ.get("GOOGLE_VISION_API_KEY", "AQ.Ab8RN6KP-bl4Z5o7v42NYu0CQgCf__Wzrpyx1vjvz_xcN8ttHQ")  # หรือใส่คีย์ตรงๆ ในเครื่องหมาย "" นี้

def is_vision_api_configured() -> bool:
    """เช็คว่าตั้งค่า Google Vision API key ไว้แล้วหรือยัง"""
    return bool(_VISION_API_KEY.strip())


def run_ocr_google_vision(crop_cv) -> str:
    """
    ส่งภาพไปอ่านด้วย Google Cloud Vision API (DOCUMENT_TEXT_DETECTION)
    ใช้ preprocess_image เดิม (whiten/deskew/แก้แสงเงา) ก่อนส่งเหมือน Tesseract
    เพราะภาพที่สะอาดขึ้นก็ยังช่วยให้ Vision API แม่นขึ้นไปอีก
    คืนค่า: ข้อความที่อ่านได้ (ผ่าน clean_text() แล้ว) หรือ raise Exception
    ถ้าเรียก API ไม่สำเร็จ (ให้ผู้เรียกจัดการ fallback เอง)
    """
    import requests  # import เฉพาะตอนใช้งานจริง กันแอป error ถ้าไม่ได้ติดตั้ง

    if not is_vision_api_configured():
        raise RuntimeError(
            "ยังไม่ได้ตั้งค่า GOOGLE_VISION_API_KEY — ดูวิธีตั้งค่าในคอมเมนต์ "
            "เหนือฟังก์ชันนี้ หรือสลับกลับไปใช้ Tesseract (ฟรี ไม่ต้องตั้งค่า)"
        )

    # Vision API เดิมต้องการภาพสี (BGR/RGB) ไม่ใช่ภาพขาวดำที่ threshold แล้ว
    # เหมือนที่ preprocess_image คืนมาให้ Tesseract เพราะ Vision API จัดการ
    # การปรับภาพเองได้ดีอยู่แล้ว และบางครั้งภาพขาวดำ-threshold กลับทำให้มัน
    # อ่านยากขึ้น (เส้นขอบ contrast จัดเกินไป) — จึงใช้ whiten_background
    # อย่างเดียว (ไม่ resize/threshold เพิ่ม) แล้วแปลงกลับเป็นภาพสีก่อนส่ง
    gray = whiten_background(crop_cv)
    img_for_api = cv2.cvtColor(gray, cv2.COLOR_GRAY2BGR)

    success, buf = cv2.imencode(".png", img_for_api)
    if not success:
        raise RuntimeError("แปลงภาพเป็น PNG ไม่สำเร็จ ก่อนส่งเข้า Vision API")
    image_b64 = base64.b64encode(buf.tobytes()).decode("utf-8")

    url = f"https://vision.googleapis.com/v1/images:annotate?key={_VISION_API_KEY}"
    payload = {
        "requests": [{
            "image": {"content": image_b64},
            "features": [{"type": "DOCUMENT_TEXT_DETECTION"}],
            "imageContext": {"languageHints": ["th", "en"]},
        }]
    }

    resp = requests.post(url, json=payload, timeout=30)
    resp.raise_for_status()
    data = resp.json()

    resp0 = data.get("responses", [{}])[0]
    if "error" in resp0:
        raise RuntimeError(f"Google Vision API error: {resp0['error'].get('message', resp0['error'])}")

    full_text = resp0.get("fullTextAnnotation", {}).get("text", "")
    if not full_text:
        # ไม่มีข้อความเลย (ภาพว่าง/อ่านไม่ออกจริงๆ) — ไม่ถือว่า error แต่คืนค่าว่าง
        return ""

    return clean_text(full_text)


def run_ocr(crop_cv, engine: str = "tesseract"):
    """
    engine: "tesseract" (ฟรี, local, ค่าเริ่มต้น) หรือ "vision" (Google Cloud
    Vision API — แม่นกว่ามาก แต่ต้องตั้งค่า API key และมีค่าใช้จ่ายเมื่อเกิน
    free tier 1,000 ภาพ/เดือน)

    ถ้าเลือก "vision" แต่เรียกไม่สำเร็จ (ไม่มี key, network error, quota
    หมด ฯลฯ) จะ fallback กลับไปใช้ Tesseract อัตโนมัติ พร้อมข้อความแจ้งเตือน
    นำหน้าผลลัพธ์ ไม่ทำให้แอปพังหรือ block การทำงาน
    """
    if engine == "vision":
        try:
            text = run_ocr_google_vision(crop_cv)
            if text:  # มีข้อความจริง ไม่ใช่ภาพว่าง
                return text
            # ภาพว่างจริงๆ (ไม่ใช่ error) — ลอง Tesseract ต่อเผื่ออ่านได้บ้าง
        except Exception as e:
            st.session_state.setdefault("_vision_api_warning", str(e))

    try:
        text = pytesseract.image_to_string(
            preprocess_image(crop_cv), lang='tha+eng', config='--psm 6')
        return clean_text(text)
    except Exception as e:
        return f"[OCR ERROR] {e}"

# ─────────────────────────────────────────────────────────────────────────────
# ██████████████  ENGINE v2 — SPACED CHARACTER PARSER  ████████████████████████
# ─────────────────────────────────────────────────────────────────────────────

def _collapse(text: str) -> str:
    """ยุบ whitespace ทั้งหมดในข้อความ (รองรับตัวอักษรกระจัดกระจาย)"""
    return re.sub(r'\s+', '', text)

def _th_spaced(word: str) -> str:
    """สร้าง regex ที่รองรับ space ระหว่างแต่ละตัวอักษร"""
    return r'\s*'.join(re.escape(c) for c in word)

# ─── keyword canonical map ────────────────────────────────────────────────────
_KW_CANONICAL = {
    "ยอดรวม": ["ยอดรวม","ยอตรวม","ยอดราม","บอดรวม","มอดราม","นลดราม","นอดรวม",
               "นลดร5าม","นลดร6าม","นลดรSาม","ยถอดรวม","ยถอดราม","มบลดาม",
               # "น ล ด ร 5 า ม" / "ย ถอด ร ว ม" / "ม บ ลดา ม" = ยอดรวม OCR เพี้ยน
               "รวมสุทธิ","รวมทั้งสิ้น","Total","NET TOTAL","net total"],
    "เงินสด": ["เงินสด","เง็นสด","เม็นสด","เแงินสด","CASH","QR","QR5","QRcode",
               "ฝัน","เงินสะด","เฉินสด"],
    "เงินทอน": ["เงินทอน","เง็นทอน","เงินทอม","เงินหอน","เงินหอแ",
                "Change","CHANGE","เงินทอร","เง็นทอร"],
    "สาขา": ["มอร์สาขา","CJมอร์","CJมอร์สาขา","สาขา","สาขาที่","มอร์สาขาที่",
             "มอร์ลาขา","ซีเจมอร์"],
}

def _build_kw_re(key: str) -> re.Pattern:
    variants = _KW_CANONICAL.get(key, [key])
    return re.compile('|'.join(_th_spaced(v) for v in variants), re.IGNORECASE)

_RE_TOTAL  = _build_kw_re("ยอดรวม")
_RE_CASH   = _build_kw_re("เงินสด")
_RE_CHANGE = _build_kw_re("เงินทอน")
_RE_BRANCH = _build_kw_re("สาขา")

# ราคา: รองรับ "39.00", "39,00", "3900" (ไม่มีจุด), "39 00" (space แทนจุด)
_RE_PRICE  = re.compile(r'(\d{1,6}[.,]\d{2})')
_RE_DATE   = re.compile(r'(\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4})')
# time: ต้องเป็น HH:MM จริง ไม่ใช่ราคา — ใช้ lookahead ว่าไม่ตามด้วยตัวเลข
_RE_TIME   = re.compile(r'\b([01]\d|2[0-3])[:.:]([0-5]\d)\b')
_RE_TAX_ID = re.compile(r'TAX\s*(?:ID|1D|10|lD|id)?\s*[:\s]?\s*(\d{10,13})', re.IGNORECASE)

# ─────────────────────────────────────────────────────────────────────────────
# parse_price — รองรับทุก format ที่ OCR อ่านผิด
# ─────────────────────────────────────────────────────────────────────────────
def parse_price(s: str) -> float:
    """
    แปลง string → float รองรับ:
    - comma แทนจุด: "45,00" → 45.00
    - space แทนจุด: "39 00" → 39.00 (เฉพาะเมื่อตัวท้าย 2 หลัก)
    - ไม่มีจุด: "3900" → 39.00, "11000" → 110.00
    - ตัวอักษรแปลก: O→0, l/I→1
    """
    if not s: return 0.0
    s = str(s).strip()
    # OCR char fixes
    s = s.replace('O','0').replace('o','0').replace('l','1').replace('I','1')
    # space แทนจุด (".39 00" → "39.00")
    s = re.sub(r'^(\d+)\s+(\d{2})$', r'\1.\2', s.strip())
    # comma → dot
    s = s.replace(',', '.')
    # เอาเฉพาะตัวเลขและจุด
    cleaned = re.sub(r'[^\d.]', '', s)
    if not cleaned: return 0.0
    # ถ้ามีจุดมากกว่า 1 ตัว ให้เก็บจุดสุดท้าย
    parts = cleaned.split('.')
    if len(parts) > 2: cleaned = parts[-2] + '.' + parts[-1]
    # ถ้าไม่มีจุดเลย และลงท้ายด้วย 00 ให้ใส่จุด: "3900"→"39.00"
    if '.' not in cleaned and len(cleaned) >= 3 and cleaned.endswith('00'):
        cleaned = cleaned[:-2] + '.' + cleaned[-2:]
    try: return float(cleaned)
    except: return 0.0

# ─────────────────────────────────────────────────────────────────────────────
# _find_prices_in_line — ดึงราคาจากบรรทัด รองรับ format หลากหลาย
# ─────────────────────────────────────────────────────────────────────────────
def _find_prices_in_line(line: str) -> list:
    """
    ดึงราคาจากบรรทัดเดียว รองรับ:
    "39.00", "39,00", "3900", "39 00", "50 0" (→500?), "110.00 บาท"
    """
    prices = []
    # 1. มีจุด/comma ปกติ
    for m in re.finditer(r'\b(\d{1,6}[.,]\d{2})\b', line):
        prices.append(parse_price(m.group(1)))
    if prices: return prices
    # 2. space แทนจุด "39 00"
    for m in re.finditer(r'\b(\d{1,5})\s+(\d{2})\b', line):
        prices.append(parse_price(f"{m.group(1)}.{m.group(2)}"))
    if prices: return prices
    # 3. ตัวอักษรไทย/อังกฤษเดี่ยว 1 ตัว แทรกกลางกลุ่มตัวเลข — เป็นเลข "0"
    #    ที่ OCR อ่านผิดเป็นตัวอักษร (ไม่ใช่ตัวคั่นทศนิยม) เช่น
    #    "10 ว 00" จริงๆ คือ "100.00" (เลข 0 ตัวที่ 3 ถูกอ่านเป็น ว แทรกกลาง
    #    ระหว่าง "10" กับ "00") แทนตัวอักษรด้วย "0" แล้วรวมกลุ่มตัวเลข
    #    เข้าด้วยกันก่อนตัดสินใจตำแหน่งทศนิยมแบบเดียวกับ fallback ถัดไป
    _merged_zero = re.sub(r'(\d)\s*[ก-๙a-zA-Z]\s*(\d)', r'\g<1>0\g<2>', line)
    if _merged_zero != line:
        _mc = _collapse(_merged_zero)
        for m in re.finditer(r'(?<!\d)(\d{3,6})(?!\d)', _mc):
            n = m.group(1)
            if n.endswith('00'): prices.append(parse_price(n))
        if prices: return prices
    # 4. ไม่มีจุด ลงท้าย 00: "3900", "11000"
    c = _collapse(line)
    for m in re.finditer(r'\b(\d{3,6})\b', c):
        n = m.group(1)
        if n.endswith('00'): prices.append(parse_price(n))
    if prices: return prices
    # 5. เลขติดกับตัวอักษรไทย/อังกฤษโดยไม่มีช่องว่างคั่น (\b ใช้ไม่ได้ในกรณีนี้
    #    เพราะตัวอักษรไทยไม่ใช่ word character ของ \b) เช่น "ส70000wiv"
    for m in re.finditer(r'(?<!\d)(\d{3,6})(?!\d)', c):
        n = m.group(1)
        if n.endswith('00'): prices.append(parse_price(n))
    return prices

# ─────────────────────────────────────────────────────────────────────────────
# _find_date_robust — รองรับ OCR อ่านปีผิด
# ─────────────────────────────────────────────────────────────────────────────
def _fix_date(d: str) -> str:
    """
    แก้วันที่ที่ OCR อ่านผิด:
    - ปีเป็น 4 หลักแต่ผิด: 7024→2024, 3024→2024, 9024→2024 (หลักแรกผิด)
    - ปีเป็น 4 หลักแต่หลักกลางผิด: 2076→2026 (อยู่นอกช่วงสมเหตุสมผล)
    - ปีเป็น 2 หลัก: 24→2024
    """
    parts = re.split(r'[-/.]', d)
    if len(parts) != 3: return d
    day, mon, yr = parts
    # fix year
    if len(yr) == 4:
        # ถ้าหลักแรกผิด (ไม่ใช่ 1 หรือ 2) แก้เป็น 2
        if yr[0] not in ('1','2'): yr = '2' + yr[1:]
        # ปีพ.ศ. > 2500 ให้แปลงเป็น ค.ศ.
        if int(yr) > 2500: yr = str(int(yr) - 543)
        # ตรวจช่วงปีสมเหตุสมผล (2020-2035) — ถ้าไม่ใช่ ลองแก้เป็น 202X
        # (รองรับกรณี OCR อ่านหลักกลางผิด เช่น 2076 ที่จริงคือ 2026)
        yr_int = int(yr) if yr.isdigit() else 0
        if not (2020 <= yr_int <= 2035) and len(yr) == 4:
            candidate = '202' + yr[3]
            if candidate.isdigit() and 2020 <= int(candidate) <= 2035:
                yr = candidate
    elif len(yr) == 2:
        yr = '20' + yr
    # fix month 00 → 01 fallback
    if mon == '00': mon = '01'
    # fix day 00 → 01 fallback
    if day == '00': day = '01'
    sep = re.search(r'[-/.]', d).group(0)
    return f"{day}{sep}{mon}{sep}{yr}"

# ─────────────────────────────────────────────────────────────────────────────
# Receipt number finder — v3 (รองรับ BNO'S..., BNO-S..., ID:€...)
# ─────────────────────────────────────────────────────────────────────────────
def _clean_bno(raw: str) -> str:
    """แปลง OCR noise ใน BNO string"""
    raw = re.sub(r'(?<=[A-Za-z0-9]):(?=\d)', '-', raw)
    result = []
    for i, c in enumerate(raw):
        if c in ('$','€','£'):
            prev = raw[i-1] if i > 0 else ''
            nxt  = raw[i+1] if i < len(raw)-1 else ''
            result.append('5' if (prev.isdigit() and nxt.isdigit()) else 'S')
        elif c == '\u00a7': result.append('S')  # § มักเป็น S ที่ OCR อ่านผิด
        elif c in ("'", "\u2018", "\u2019", "`"): result.append('')  # straight + smart quotes
        else: result.append(c)
    return re.sub(r'S{2,}', 'S', ''.join(result))

def _find_rcpt_no(text: str, compact: str) -> str:
    """
    เลขที่ใบเสร็จ = ทุกอย่างตั้งแต่ BNO จนจบบรรทัด (รวมอักษรขยะที่ติดมาด้วย)
    ไม่ตัดทิ้งกลางทาง เพื่อไม่ให้สูญเสียข้อมูลที่ OCR อ่านมา
    """
    # 1. BNO / 8NO / BN0 — จับทุกอย่างหลัง keyword จนจบบรรทัด ไม่หยุดกลางทาง
    #    separator รองรับทั้ง straight quote (') และ smart/curly quote (‘ ’) ที่ OCR มักใช้
    for line in text.split('\n'):
        c = _collapse(line)
        m = re.search(r"(?:BNO|8NO|BN0)[:'\u2018\u2019`\-\.\s]*(.+)$", c, re.IGNORECASE)
        if m:
            raw_tail = m.group(1).strip()
            if raw_tail:
                # พยายามล้าง OCR noise ปกติก่อน ($→S/5, colon→dash) แต่ไม่ตัดอักษรขยะทิ้ง
                cleaned = _clean_bno(raw_tail)
                return cleaned if cleaned else raw_tail

    # 2. standalone S\d+N\d+-\d+ (ไม่ต้องมี BNO นำหน้า)
    m = re.search(r'([A-Z]\d{7,}[A-Z]\d{2}-\d{4,})', compact)
    if m: return m.group(1)

    # 3. ID:E / 1D:€ / lD:E pattern (session ID — ใช้เป็น fallback)
    for src in [text, compact]:
        m = re.search(
            r'(?:1D|ID|lD)\s*[:\s€£$]\s*([A-Za-z][A-Za-z0-9]{5,24})',
            src, re.IGNORECASE)
        if m:
            raw = m.group(1).replace('€','E').replace('£','E').replace('$','5').replace('O','0')
            cut = re.match(r'([A-Z][A-Z0-9]{5,23})', raw)
            if cut: return cut.group(1)

    # 4. fallback: รูปแบบ Rcpt# (เผื่อ OCR อ่าน BNO ผิดเป็นรูปแบบอื่น)
    m = re.search(r'(?:Rcpt|RCPT|Rcopth)[^\d]{0,6}(\d{6,})', compact, re.IGNORECASE)
    if m: return m.group(1)

    # 5. fallback: OCR อ่านตัวอักษร BNO เพี้ยนจนไม่เหลือ B/N/O เลย
    #    (เช่น "ธ ม 0: ร 25110041 ม 02.000877" ที่จริงคือ "BNO:S25110041N02-000877")
    #    หา pattern ตัวเลข 8 หลัก + (แยก) + 2 หลัก + (แยก) + 6 หลัก ซึ่งเป็นรูปแบบ
    #    เฉพาะของเลขที่ใบเสร็จ CJ ไม่ปนกับวันที่/เวลาที่ไม่มี digit-run ยาวขนาดนี้
    for line in text.split('\n'):
        c = _collapse(line)
        m = re.search(r'(\d{8})[^0-9]{0,4}(\d{2})[^0-9]{0,4}(\d{6})', c)
        if m:
            return f"S{m.group(1)}N{m.group(2)}-{m.group(3)}"

    # 6. fallback: เลขกระจัดกระจายมาก จนแม้แต่ pattern ขั้น 5 ก็จับไม่ได้
    #    (เช่น "BLO ร 26 ู 051775 ผ น 2 ะ 66172") — รวมเลขทั้งหมดในบรรทัดที่มี
    #    BNO-variant เข้าด้วยกันเป็นก้อนเดียว แล้วประกอบเป็น BNO ตามสัดส่วน 8:2:6
    #    แม่นยำน้อยกว่าขั้นก่อน แต่ยังดีกว่า "ไม่พบ" สำหรับ OCR ที่เพี้ยนหนักมาก
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

    # 7. fallback: เลขยาวบนบรรทัดวันที่
    for line in text.split('\n'):
        if _RE_DATE.search(line):
            nums = re.findall(r'\d{6,}', _collapse(line))
            if nums: return max(nums, key=len)

    return "ไม่พบ"

def _find_pos_id(text: str, compact: str, lines: list) -> str:
    """
    รหัสสาขา/POS = เลขรหัสสาขาที่อยู่บรรทัดแรกๆ ของบิล (เช่น "มอร์สาขาที่01745")
    ให้ความสำคัญกับคำว่า "สาขา" (รวม OCR variant) ก่อนเสมอ ไม่ใช่ POS terminal number
    ค้นในหลายบรรทัดแรก เพราะบางภาพมีขยะ OCR จากพื้นหลัง/ขอบกระดาษปนมาก่อน
    เนื้อหาบิลจริงหลายบรรทัด ทำให้ "บรรทัดแรกตามตัวอักษร" ไม่ใช่บรรทัดที่มีรหัสสาขาเสมอไป
    """
    _BRANCH_KW = r'(?:สาขา|ลาขา|ฉาขา|ฬาขา|ขัสาชา|สาชาต|ชาขาต)'

    # 1. ค้นใน 12 บรรทัดแรก หาเลข 5 หลักที่อยู่ติดกับคำว่า "สาขา" (รวม OCR variant)
    for line in lines[:12]:
        c = _collapse(line)
        m = re.search(_BRANCH_KW + r'[^\d]{0,4}(\d{5})', c, re.IGNORECASE)
        if m: return m.group(1)

    # 2. ถ้าไม่เจอคำนำหน้าเลย ลองหาบรรทัดที่มีทั้งคำว่า "มอร์" และเลข 5 หลัก
    #    (ชื่อร้าน "มอร์" มักอยู่ติดกับรหัสสาขาเสมอ แม้คำว่า "สาขา" จะอ่านผิดไปไกล)
    for line in lines[:12]:
        c = _collapse(line)
        if re.search(r'มอร์|ม อ ร', line, re.IGNORECASE) or 'มอร์' in c:
            nums = re.findall(r'\d{5}', c)
            if nums: return nums[0]

    # 3. fallback: เลข 5 หลักตัวแรกที่เจอใน compact ทั้งก้อน (ก่อนถึง TAX ID ซึ่งยาว 13 หลัก)
    m = re.search(r'(?:สาขา|branch)[^\d]{0,5}(\d{5})', compact, re.IGNORECASE)
    if m: return m.group(1)
    # หมายเหตุ: ใช้ (?<!\d)...(?!\d) แทน \b เพราะ \b ใช้ไม่ได้ระหว่างตัวอักษร
    # ไทยกับตัวเลข (ตัวอักษรไทยไม่ถูกนับเป็น \w ในมาตรฐาน regex ของ Python)
    # ทำให้ \b\d{5}\b ไม่ match เลข 5 หลักที่อยู่ติดกับข้อความไทยโดยไม่มี
    # ช่องว่างคั่นเลย เช่น "...สาวาถี01745หนองแสนะ" ซึ่งเป็นรูปแบบที่ OCR
    # มักให้ผลออกมาบ่อยเมื่อความคมชัดของภาพต่ำ
    nums_all = re.findall(r'(?<!\d)\d{5}(?!\d)', compact)
    if nums_all: return nums_all[0]

    # 4. fallback สุดท้าย: POS terminal number (NO1/NO2/...) ถ้าไม่มีเลขสาขาเลย
    for line in lines[:10]:
        c = _collapse(line)
        m = re.search(r'POS\s*[.\s]*NO\s*S?\s*(\d{1,2})\b', c, re.IGNORECASE)
        if m: return f"NO{m.group(1)}"
        m = re.search(r'POS\s*(?:ID|:)?\s*#?\s*([A-Za-z]\d{1,3})(?!\d)', c, re.IGNORECASE)
        if m: return m.group(1)

    m = re.search(r'\((\d{4,})\)', compact)
    if m: return m.group(1)
    return "ไม่พบ"

def _find_branch(text: str, compact: str, lines: list) -> str:
    # 1. จัดการกรณี CJ / มอร์ (ใช้ Regex ที่ดึงตัวเลขแยกจากคำนำหน้า)
    cj_pattern = r'(?:ซีเจ|CJ)?ม[อ][ร][ลซ]์?.*?สาขา(?:ที่|เลขที่)?\s*?(\d{2,}(?:\s*\d+)*)'
    bigc_pattern = r'(Big\s*C\s*(?:Mini|Extra)?|BCM|BCH)'

    for line in lines[:12]:
        c = _collapse(line)
        m_cj = re.search(cj_pattern, c, re.IGNORECASE)
        if m_cj:
            raw_num = m_cj.group(1).replace(" ", "")
            return f"สาขา {raw_num}"
        m_bigc = re.search(bigc_pattern, c, re.IGNORECASE)
        if m_bigc:
            return line.strip()

    m_compact = re.search(r'(?:มอร์|สาขา)\s*?(\d{2,}(?:\s*\d+)*)', compact, re.IGNORECASE)
    if m_compact:
        return f"สาขา {m_compact.group(1).replace(' ', '')}"
    return "ไม่พบ"

_RE_POINTS_LINE = re.compile(r'แต\s*้?\s*ม', re.IGNORECASE)  # บรรทัดเกี่ยวกับแต้มสะสม

def _find_amount(text_lines: list, kw_re: re.Pattern, allow_zero: bool = True) -> float:
    """
    หาจำนวนเงินจาก keyword line — spaced-aware
    ค้นหาราคาในบรรทัดทั้งรูปแบบปกติและ compact
    allow_zero=True: คืน 0.0 ถ้า OCR เจอ "0.00" จริง

    ข้ามบรรทัดที่เกี่ยวกับ "แต้ม" สะสม (เช่น "5.00 แต้มแลกแทนเงินสด 50 บาท")
    เพราะบรรทัดพวกนี้มักมีคำว่า "เงินสด"/"บาท" ปนอยู่ด้วยทั้งที่ไม่ใช่ยอดเงิน
    จริงของบิล ทำให้จับผิดบรรทัดได้ถ้าไม่กรองออกก่อน
    """
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
    """
    fallback สุดท้ายสำหรับ ยอดรวม/เงินสด/เงินทอน เมื่อ OCR เพี้ยนหนักจน
    keyword (ยอดรวม/เงินสด/เงินทอน และทุก OCR-variant ที่รู้จัก) หาไม่เจอเลย
    เช่น "ยอดรวม" กลายเป็น "บลดราม", "เงินสด" กลายเป็น "เง น 11 า น"

    หลักการ: ใบเสร็จ CJ Express มีโครงสร้างคงที่เสมอ — หลังบรรทัด
    "...รายการ" (สรุปจำนวนสินค้า) จะตามด้วย ยอดรวม → เงินสด/QR → เงินทอน
    เรียงลำดับแบบนี้เสมอ ไม่ว่า keyword จะเพี้ยนแค่ไหน เราจึงดึงราคาจาก
    3 บรรทัดแรกที่มีตัวเลขถัดจากจุดนั้นมาใช้ตามตำแหน่งได้
    คืนค่า: list ของราคา (float) ตามลำดับ [ยอดรวม, เงินสด, เงินทอน] (อาจสั้นกว่า 3 ถ้าหาไม่ครบ)
    """
    end_idx = 0
    for idx, line in enumerate(lines):
        # รองรับ OCR สับสน ร↔ง (เช่น "รายการ" อ่านเป็น "งายการ") และไม่บังคับ
        # ต้องมีตัวเลขนำหน้าติดกันเป๊ะอีกต่อไป (เดิม \d\s*[รง]... พลาดเคสที่
        # ตัวเลขกับคำว่า "รายการ" ถูก OCR แยกออกจากกันไกล หรือมีตัวอักษรแทรก
        # ระหว่าง ร/ง กับ ย เช่น "ใ ร จ ย ก า ร" ที่จริงคือ "รายการ")
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

# ─────────────────────────────────────────────────────────────────────────────
# Text cleaning — v3 (เพิ่ม pass สำหรับ date/price/BNO fix)
# ─────────────────────────────────────────────────────────────────────────────
def clean_text(text: str) -> str:
    if not text: return ""

    # pass 1: normalize whitespace
    lines = [re.sub(r'[ \t]+', ' ', ln).strip() for ln in text.split('\n')]

    # pass 2: spaced-keyword collapse
    _SPACED = [
        (r'ย\s*อ\s*[ดต]\s*ร\s*ว\s*ม',          'ยอดรวม'),
        (r'น\s*ล\s*ด\s*ร\s*[า5]\s*ม',           'ยอดรวม'),   # "น ล ด ร า ม"
        (r'บ\s*อ\s*ด\s*ร\s*ว\s*ม',              'ยอดรวม'),
        (r'ม\s*อ\s*ด\s*ร\s*า\s*ม',              'ยอดรวม'),
        (r'เ\s*ง\s*ิ?\s*น\s*ส\s*ด',             'เงินสด'),
        (r'เ\s*ง\s*ิ?\s*น\s*ท\s*อ\s*น',         'เงินทอน'),
        (r'เ\s*ง\s*ิ?\s*น\s*ห\s*อ\s*[นแ]',      'เงินทอน'),
        (r'เ\s*ง\s*[ิี]\s*น\s*[7]\s*[ไ]',        'เงินทอน'),   # "เงิน 7 ไน"
        (r'บ\s*า\s*ท',                            'บาท'),
        (r'ม\s*อ\s*ร\s*์\s*[สลซ]\s*า\s*ข\s*า',  'มอร์สาขา'),
        (r'ส\s*า\s*ข\s*า\s*[หท]\s*[ิี]\s*[่]?',  'สาขาที่'),
        (r'ใ\s*บ\s*เ\s*[สล]\s*ร\s*[็]\s*จ',      'ใบเสร็จ'),
        (r'T\s*o\s*t\s*a\s*[l1i]',               'Total'),
        (r'C\s*A\s*S\s*H',                        'CASH'),
        (r'C\s*h\s*a\s*n\s*g\s*e',               'Change'),
        (r'V\s*A\s*[TI]',                         'VAT'),
        (r'R\s*E\s*C\s*E\s*I\s*P\s*T',           'RECEIPT'),
        (r'I\s*N\s*C\s*L\s*U\s*D\s*E\s*D',       'INCLUDED'),
        (r'T\s*A\s*X\s*[I1l]\s*D',               'TAXID'),
    ]
    out = []
    for line in lines:
        for pat, rep in _SPACED:
            line = re.sub(pat, rep, line, flags=re.IGNORECASE)
        out.append(line)

    # pass 3: string replacements
    text2 = '\n'.join(out)
    _FIXES = {
        # Total variants
        "Tota)":"Total","Tota1":"Total","Toial":"Total","Totai":"Total",
        # ยอดรวม
        "ยอดราม":"ยอดรวม","ยอตรวม":"ยอดรวม","รวมสุทธิ":"ยอดรวม","Net Total":"ยอดรวม",
        # เงินสด
        "เง็นสด":"เงินสด","เม็นสด":"เงินสด","เแงินสด":"เงินสด",
        "Quan":"เงินสด","QUA N":"เงินสด","ฝัน ต":"เงินสด","ฝันต":"เงินสด",
        # เงินทอน
        "เงินทอม":"เงินทอน","เง็นทอน":"เงินทอน",
        # ID
        "1D:":"ID:","lD:":"ID:","ID:£":"ID:E","ID:$":"ID:S","ID:€":"ID:E",
        "1D:€":"ID:E","lD:€":"ID:E",
        # receipt
        "RECEIPI":"RECEIPT",
        # บาท variants ที่ OCR อ่านเป็นอื่น
        "uw":"บาท","inn":"บาท","inv":"บาท","บนาท":"บาท","น บ นา ท":"บาท",
        # comma ทศนิยม — ไม่แทน (parse_price จัดการ)
    }
    for old, new in _FIXES.items():
        text2 = text2.replace(old, new)

    # pass 4: fix date year OCR errors (e.g. 7024→2024)
    def fix_date_m(m):
        return _fix_date(m.group(0))
    text2 = _RE_DATE.sub(fix_date_m, text2)

    # pass 5: fix price format in amount keyword lines
    amt_kw = ["Total","CASH","Change","VAT","เงินทอน","เงินสด","ยอดรวม","บาท","QR"]
    final = []
    for line in text2.split('\n'):
        low = line.lower()
        if any(k.lower() in low for k in amt_kw):
            # colon/space แทนจุด: "39:00" "39 00"
            line = re.sub(r'\b(\d{1,5})\s*[:]\s*(\d{2})\b', r'\1.\2', line)
            line = re.sub(r'\b(\d{1,5})\s+(\d{2})\b(?!\s*\d)', r'\1.\2', line)
            # ไม่มีจุด ลงท้าย 00: 11000→110.00
            line = re.sub(r'(?<!\d\.)\b(\d{3,6})\b(?!\.\d)',
                lambda m: (m.group(0)[:-2]+'.'+m.group(0)[-2:]
                           if m.group(0).endswith('00') and len(m.group(0))>=3
                           else m.group(0)), line)
            # comma→dot
            line = re.sub(r'(\d+),(\d{2})\b', r'\1.\2', line)
        final.append(line)

    return '\n'.join(final)

# ─────────────────────────────────────────────────────────────────────────────
# Unified extractor (CJ Express only — engine v2)
# ─────────────────────────────────────────────────────────────────────────────
def extract_receipt(text: str) -> dict:
    """
    ดึงข้อมูลหลักจากบิล CJ Express
    ทำงานกับ OCR ที่มีช่องว่างกระจัดกระจาย
    """
    lines   = text.split('\n')
    compact = _collapse(text)

    # --- date ---
    date_m = _RE_DATE.search(text)
    date_str = _fix_date(date_m.group(1)) if date_m else "ไม่พบ"

    # --- time: ค้นหา HH:MM จริง ไม่ใช่ราคา ---
    time_val = "ไม่พบ"
    for line in lines:
        m = _RE_TIME.search(line)
        if m:
            h, mn = m.group(1), m.group(2)
            time_val = f"{h}:{mn}"
            break
    if time_val == "ไม่พบ":
        for line in lines:
            if _RE_DATE.search(line):
                after = _RE_DATE.sub('', line).strip()
                m = re.search(r'\b(\d{2})\s*[:. ]\s*(\d{2})\b', after)
                if m:
                    h, mn = m.group(1), m.group(2)
                    if 0 <= int(h) <= 23 and 0 <= int(mn) <= 59:
                        time_val = f"{h}:{mn}"
                        break
                m2 = re.search(r'\b([01]\d|2[0-3])([0-5]\d)\b', after)
                if m2: time_val = f"{m2.group(1)}:{m2.group(2)}"; break

    # --- branch / pos / rcpt / tax ---
    branch  = _find_branch(text, compact, lines)
    pos_id  = _find_pos_id(text, compact, lines)
    rcpt_no = _find_rcpt_no(text, compact)
    tax_id  = _find_tax_id(compact)

    # --- user ---
    user_m = re.search(r'User\s*#?\s*(\w+)', compact, re.IGNORECASE)
    user   = user_m.group(1) if user_m else "ไม่พบ"

    # --- amounts (spaced-aware) ---
    total  = _find_amount(lines, _RE_TOTAL)
    cash   = _find_amount(lines, _RE_CASH)
    change = _find_amount(lines, _RE_CHANGE)

    if total == 0.0:
        _skip_total = re.compile(r'จำนวนสินค้า|จำนวนรายการ|จานวน|ร า ย ก า ร|รายการ', re.IGNORECASE)
        for line in lines:
            c = _collapse(line)
            is_total_line = (re.search(r'ยอดรวม|นลดราม|มอดราม', c, re.IGNORECASE) and
                             not _skip_total.search(c))
            if is_total_line:
                prices = _find_prices_in_line(line) or _find_prices_in_line(c)
                if prices: total = prices[-1]; break

    if change == 0.0:
        for line in lines:
            c = _collapse(line)
            if re.search(r'เงิน.{0,4}[าท][าน]', c, re.IGNORECASE):
                prices = _find_prices_in_line(line) or _find_prices_in_line(c)
                if prices and prices[-1] > 0: change = prices[-1]; break
            if re.search(r'เง\s*ิ\s*น\s*[1ไ7]\s*[าท]', line, re.IGNORECASE):
                prices = _find_prices_in_line(line)
                if prices and prices[-1] > 0: change = prices[-1]; break
            # เคสเพี้ยนหนัก: "เงินทอน" เหลือแค่ "เง น 11 า น" (ิ หาย, ท→1,
            # อ→า) — ไม่มี "ท" เหลือให้ pattern ข้างบนจับได้เลย ต้องเทียบ
            # จาก compact string โดยตรงว่าขึ้นต้นด้วย "เงน" ตามด้วยเลข 1-2
            # ตัวซ้ำกัน (1 หรือ l ที่ OCR สับสนกับ 1) แล้วตามด้วย "าน/านม"
            if re.match(r'เง[นม]\s*[1lI]{1,2}\s*[าห][นม]', c, re.IGNORECASE):
                prices = _find_prices_in_line(line) or _find_prices_in_line(c)
                if prices and prices[-1] > 0: change = prices[-1]; break

    # fallback สุดท้าย: ถ้าหา "ยอดรวม" ด้วย keyword ทุกแบบไม่เจอเลย
    # (total ยังเป็น 0 อยู่ทั้งที่ผ่าน 2 รอบ keyword-based ด้านบนแล้ว)
    # แปลว่า OCR เพี้ยนหนักมากจน keyword ของยอดรวมไม่เหลือเค้าโครงเดิมเลย
    # ในกรณีนี้ค่า cash/change ที่ "เจอ" จาก keyword อื่นก็อาจสุ่มมั่วได้เช่นกัน
    # (เช่น keyword ของเงินสดบังเอิญไปจับบรรทัดอื่นที่ไม่ใช่เงินสดจริง)
    # จึงใช้ตำแหน่งสัมพัทธ์แทนทั้งชุด — ดูคำอธิบายใน _find_amounts_positional()
    #
    # หมายเหตุ: เงื่อนไขนี้ตั้งใจให้ผูกกับ total==0.0 เท่านั้น (ไม่ใช่ cash/change
    # แยกอิสระ) เพราะตำแหน่งสัมพัทธ์ (positional) คำนวณจาก "ลำดับบรรทัดที่มี
    # ราคา" นับจากจุดจบรายการสินค้า ถ้า total ถูกเจอแล้วจาก keyword ที่แม่นยำ
    # (อยู่คนละตำแหน่งกับลำดับที่ positional namedคาดไว้ เช่น มีบรรทัดส่วนลด
    # แทรกอยู่ก่อนยอดรวมจริง) การเอา positional[1]/[2] มาเขียนทับ cash/change
    # จะยิ่งทำให้ผิดมากขึ้น เพราะตำแหน่งเลื่อนไม่ตรงกับที่ keyword เจอจริงแล้ว
    if total == 0.0:
        positional = _find_amounts_positional(lines)
        if len(positional) >= 1: total  = positional[0]
        if len(positional) >= 2: cash   = positional[1]
        if len(positional) >= 3: change = positional[2]

    # --- name (first product line heuristic) ---
    skip_name_kw = ["Total","CASH","Change","Vat","Rcpt","POS","TAX","User",
                    "INCLUDED","RECEIPT","INVOICE","ขอบคุณ",
                    "ยอดรวม","เงินสด","เงินทอน","สาขา","บาท","ID:","BNO"]
    name = "ไม่พบ"
    for line in lines:
        if any(k.lower() in line.lower() for k in skip_name_kw): continue
        m = _RE_PRICE.search(line)
        if m:
            candidate = re.sub(r'^\d+\s*[xXP]?\s*', '', line[:m.start()]).strip()
            if len(candidate) >= 2: name = candidate; break

    return {
        "date": date_str,
        "time": time_val,
        "branch": branch,
        "name":   name,
        "total_amount": total,
        "cash":   cash,
        "change": change,
        "pos_id": pos_id,
        "rcpt_no":rcpt_no,
        "tax_id": tax_id,
        "user":   user,
    }

# ─────────────────────────────────────────────────────────────────────────────
# Item extractor (CJ) — engine v2 with spaced-line support
# ─────────────────────────────────────────────────────────────────────────────
def _fix_spaced_price_core(s: str) -> str:
    # ทำ "เลข 3-4 หลักลงท้าย 00" ให้เป็นทศนิยมก่อนเสมอ (เช่น "1000" -> "10.00")
    # ต้องทำขั้นนี้ก่อน "25 00" -> "25.00" เสมอ ไม่งั้นถ้าบรรทัดมีทั้งสองรูปแบบ
    # ติดกัน (เช่น "1000 10 00") การรันสลับลำดับจะทำให้ได้จุดทศนิยมซ้อนกัน
    # ผิดรูปแบบ (เช่น "10.00.10 00") เพราะ "1000" ถูกแปลงเป็น "10.00" ไปแล้ว
    # แล้ว regex "(\d+)\s+(\d{2})" ตัวถัดมาดันไปจับ "00" ที่เหลือจาก "10" ก่อนหน้าซ้ำ
    def _dot4(m):
        n = m.group(0)
        return (n[:-2]+'.'+n[-2:]) if n.endswith('00') else n
    s = re.sub(r'(?<!\d)\d{3,4}(?!\d)', _dot4, s)
    # "25 00" → "25.00" (ช่องว่างแทนจุดทศนิยม 2 หลัก) — เฉพาะที่ยังไม่มีจุด
    # นำหน้าอยู่แล้ว (กัน double-dot ซ้ำกับขั้นตอนด้านบน)
    s = re.sub(r'(?<![.\d])(\d+)\s+(\d{2})(?=\s|$)', r'\1.\2', s)
    # "50 0" → "50.00" (OCR อ่านเลขท้ายตกหายไป 1 หลัก จาก "50.00" เดิม)
    s = re.sub(r'(?<![.\d])(\d+)\s+(\d)(?=\s|$)', r'\1.\g<2>0', s)
    return s

def _fix_spaced_price(s: str) -> str:
    # ตัดเลขจำนวน (qty) นำหน้าออกก่อนเสมอ ไม่ให้ logic แก้ราคาด้านบนไปจับ
    # ผิด เช่น "0 1 ชื่อสินค้า..." (เลข item index 2 ตัวติดกันต้นบรรทัด
    # ที่ OCR เผลอใส่ "0" นำหน้า "1") จะถูกตีความผิดเป็นรูปแบบราคา
    # "ช่องว่างแทนจุดทศนิยม" กลายเป็น "0.10" ทั้งที่ไม่ใช่ราคาเลย
    # รองรับ qty นำหน้าได้สูงสุด 2 ตัวเลขติดกัน ตามด้วยอักขระที่ไม่ใช่ตัวเลข
    # (จุดเริ่มต้นของชื่อสินค้า) จึงจะถือว่าเป็นส่วน qty ที่ต้องเว้นไว้
    m = re.match(r'^((?:\d+\s+){1,2})(?=\D)(.*)$', s)
    if m:
        qty_prefix, rest = m.groups()
        return qty_prefix + _fix_spaced_price_core(rest)
    return _fix_spaced_price_core(s)

def _clean_item_name(nm: str) -> str:
    """ยุบช่องว่างซ้ำ + ตัดสัญลักษณ์ขยะที่ติดหัว/ท้ายชื่อสินค้า (เช่น !, ", ', |)"""
    nm = re.sub(r'\s+', ' ', nm).strip()
    nm = re.sub(r'^[!"\'|！,\-\.]+\s*', '', nm)
    nm = re.sub(r'\s*[!"\'|！,\-]+$', '', nm)
    return nm.strip()

def extract_items_cj(text: str) -> list:
    items  = []
    lines  = text.split('\n')

    start_idx = 0
    _DATE_RE2 = re.compile(r'\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}')
    for idx, line in enumerate(lines):
        if _DATE_RE2.search(line) or _DATE_RE2.search(_collapse(line)):
            start_idx = idx + 1; break

    stop_kw = ["ยอดรวม","รวมทั้งสิ้น","เงินสด","เงินทอน",
               "จำนวนสินค้า","จำนวนรายการ","จานวนสินค้า","จํานวนสินค้า"]
    skip_kw = ["BNO","8NO","POS","TAX","INCLUDED","ใบเสร็จ","โปรโม",
               "ส่วนลด","แต้ม","ขอบคุณ","สาขา","RECEIPT","INVOICE",
               "สมาชิก","ID:","QR","User"]

    # suffix ท้ายบรรทัด: รองรับขยะหลายตัวอักษร/ช่องว่างปนกัน (เช่น "Vt", "บ หู")
    # ใช้ Unicode range ของอักษรไทยทั้งหมด (\u0E00-\u0E7F) แทนการ list ทีละตัว
    # เพราะตัวอักษรไทยรวมสระ/วรรณยุกต์มีมากเกินจะ list หมดและพลาดง่าย (เช่น ู, ึ, ์)
    _SUFFIX = r'[\sA-Za-z\u0E00-\u0E7F"\u201c\u201d|!！]*'
    _full  = re.compile(r'^[\.]?\s*(\d+)\s+(.+?)\s+(\d+[.,]\d{2})\s+(\d+[.,]\d{2})' + _SUFFIX + r'$')
    _fb_a  = re.compile(r'^(.+?)\s+(\d+[.,]\d{2})\s+(\d+[.,]\d{2})' + _SUFFIX + r'$')
    _fb_b  = re.compile(r'^[\.]?\s*(\d+)\s+(.+?)\s+(\d+[.,]\d{2})' + _SUFFIX + r'$')
    _fb_c  = re.compile(r'^[\.]?\s*(\d+)\s+(.+?)\s+(\d+[.,]\d{2})\s*$')

    for line in lines[start_idx:]:
        line = line.strip()
        if not line: continue
        compact = _collapse(line)

        if any(k in line or k in compact for k in stop_kw): break
        if re.search(r'จ.{0,3}นวนส', compact): break
        # หยุดเมื่อเจอ "เลข+รายการ" (สรุปจำนวนรายการ) แม้คำว่า "จำนวนสินค้า"
        # ข้างหน้าจะเพี้ยนจน regex ข้างบนจับไม่ได้ก็ตาม — "รายการ" เป็นคำที่ OCR
        # มักอ่านถูกอยู่เสมอเพราะไม่มีสระ/วรรณยุกต์ซับซ้อนเท่าคำอื่น
        # รองรับ OCR สับสน ร↔ง ด้วย (เช่น "รายการ" อ่านเป็น "งายการ") และไม่
        # บังคับตัวเลขนำหน้าติดกันเป๊ะ (รองรับตัวอักษรแทรกระหว่าง ร/ง กับ ย)
        if re.search(r'[รง]\s*[ก-๙]{0,2}\s*ย\s*ก\s*า\s*ร', line): break
        if any(k.lower() in compact.lower() for k in skip_kw): continue
        if _RE_DATE.search(line): continue
        if re.search(r'-\s*\d+[.,]\d{2}', line): continue   # ส่วนลด (ติดลบ)
        if len(line) < 5: continue

        # ตัดอักขระขยะที่ติดหัว/ท้ายบรรทัดออกก่อน (เช่น ". [", "> ]", "es ", '"')
        # OCR มักใส่สัญลักษณ์แปลกปลอมไว้ต้นบรรทัดแทนเลขจำนวนที่อ่านไม่ออก
        line_stripped = re.sub(r'^[.\[\]>"\'`*es]+\s*', '', line.strip())
        if not line_stripped: line_stripped = line.strip()

        lf = _fix_spaced_price(re.sub(r'[\|！｜\[\]]+\s*$','',line_stripped).strip())

        m = _full.match(lf)
        if m:
            qty, nm, up, tp = m.groups()
            nm = _clean_item_name(nm)
            if len(re.sub(r'[^\u0E00-\u0E7FA-Za-z0-9]','',nm)) >= 2:
                u=parse_price(up); tt=parse_price(tp); q=int(qty)
                if tt < u*0.5: tt = round(u*q,2)
                items.append({"ชื่อสินค้า":nm,"จำนวน":q,"ราคาต่อหน่วย":u,"ยอดรวมสินค้า":tt})
            continue

        m = _fb_a.match(lf)
        if m:
            nm, up, tp = m.groups()
            nm = _clean_item_name(nm)
            if len(re.sub(r'[^\u0E00-\u0E7FA-Za-z0-9]','',nm)) >= 2:
                items.append({"ชื่อสินค้า":nm,"จำนวน":1,
                              "ราคาต่อหน่วย":parse_price(up),"ยอดรวมสินค้า":parse_price(tp)})
            continue

        m = _fb_b.match(lf)
        if m:
            qty, nm, price = m.groups()
            nm = _clean_item_name(nm)
            if len(re.sub(r'[^\u0E00-\u0E7FA-Za-z0-9]','',nm)) >= 2:
                p = parse_price(price)
                items.append({"ชื่อสินค้า":nm,"จำนวน":int(qty),"ราคาต่อหน่วย":p,"ยอดรวมสินค้า":p})
            continue

        m = _fb_c.match(lf)
        if m:
            qty, nm, price = m.groups()
            nm = re.sub(r'\s+[0-9]+[.,][0-9]{2}\s*',' ',nm)
            nm = _clean_item_name(nm)
            if len(re.sub(r'[^\u0E00-\u0E7FA-Za-z]','',nm)) >= 3:
                p = parse_price(price)
                items.append({"ชื่อสินค้า":nm,"จำนวน":int(qty),"ราคาต่อหน่วย":p,"ยอดรวมสินค้า":p})

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
                    "เลขที่ใบเสร็จ":d['rcpt_no']}
            for it in (b['items'] or []):
                row = base.copy()
                row.update({"ชื่อสินค้า":it.get('ชื่อสินค้า',''),
                            "จำนวน":it.get('จำนวน',1),
                            "ราคาต่อหน่วย":it.get('ราคาต่อหน่วย',0),
                            "ยอดรวมสินค้า":it.get('ยอดรวมสินค้า',0),
                            "ยอดรวม":"","เงินสด":"","เงินทอน":""})
                rows.append(row)
            summary = base.copy()
            summary.update({"ชื่อสินค้า":"","จำนวน":"","ราคาต่อหน่วย":"","ยอดรวมสินค้า":"",
                            "ยอดรวม":d['total_amount'],"เงินสด":d['cash'],
                            "เงินทอน":d['change']})
            rows.append(summary)
        pd.DataFrame(rows).to_excel(writer, index=False, sheet_name='ใบเสร็จ')
    return output.getvalue()

# ─────────────────────────────────────────────────────────────────────────────
# Interactive crop canvas — v4
# เส้นแดงแบ่งบิล (drag ปรับได้) + ส่งตำแหน่งสุดท้ายกลับมาผ่าน query param
# ใช้ st.query_params (หรือ session_state ผ่าน on_change ของ hidden widget)
# ─────────────────────────────────────────────────────────────────────────────
def _compute_split_positions(pil_img: Image.Image, n_bills: int) -> list:
    """
    คำนวณตำแหน่งแนวตั้งเริ่มต้นที่ควรตัดบิล (0.0–1.0 เป็น fraction ของความกว้าง)
    ใช้เป็นค่าเริ่มต้นให้ user ปรับต่อ — ไม่ใช่ค่าสุดท้ายที่ใช้ตัดจริง
    """
    if n_bills <= 1:
        return []
    img_cv = pil_to_cv(pil_img)
    h, w = img_cv.shape[:2]
    gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
    col_ratio = (gray > 180).astype(np.uint8).mean(axis=0)
    splits, smooth, _ = _find_bill_splits(col_ratio, w, min_bill_frac=0.12, n_expected=n_bills)
    if len(splits) >= n_bills - 1:
        return [s / w for s in splits[:n_bills - 1]]
    # fallback: แบ่งเท่าๆ กัน
    return [i / n_bills for i in range(1, n_bills)]

def crop_component_html(pil_img: Image.Image,
                        crop_mode: str = "free",
                        bill_count: int = 1,
                        component_key: str = "crop1") -> str:
    b64 = img_to_b64(pil_img)
    orig_w, orig_h = pil_img.size
    is_a4 = "true" if crop_mode == "a4" else "false"

    split_fracs = _compute_split_positions(pil_img, bill_count) if bill_count > 1 else []
    split_js = json.dumps(split_fracs)
    n_bills_js = bill_count

    label_hint = {
        1: "ลากบนรูปเพื่อเลือกพื้นที่ที่ต้องการ Crop",
        2: "เส้นแดง = จุดตัดบิล · ลากปรับตำแหน่งได้ · ลากบนรูปเพื่อ Crop",
        3: "เส้นแดง = จุดตัดบิล · ลากปรับได้ · ลากบนรูปเพื่อ Crop",
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
#copy-data{{background:#10B981;color:#fff;display:none}}
textarea#data-out{{width:100%;margin-top:8px;font-size:11px;font-family:monospace;
  height:60px;border-radius:6px;border:1px solid #ddd;padding:6px;display:none}}
</style></head><body>
<div id="wrap"><canvas id="c"></canvas></div>
<div id="info">{label_hint}</div>
<div id="result"></div>
<div id="split-info"></div><br>
<button id="btn-confirm" disabled>✅ ยืนยัน Crop</button>
<button id="btn-reset">🔄 วาดใหม่</button>
<br>
<div style="margin-top:10px;font-size:12px;color:#666">
👇 คัดลอกข้อความนี้ไปวางในช่อง "ข้อมูล Crop" ด้านล่างเพื่อยืนยัน
</div>
<textarea id="data-out" readonly onclick="this.select()"></textarea>
<script>
const IMG_W={orig_w}, IMG_H={orig_h}, A4={is_a4}, N_BILLS={n_bills_js};
const SPLIT_FRACS = {split_js};

const canvas = document.getElementById('c'), ctx = canvas.getContext('2d');
const MAX_W = 660, MAX_H = 460;
let scaleX, scaleY;
let sx, sy, ex, ey, drawing = false, hasCrop = false, cropRect = null;

let splitLines = SPLIT_FRACS.map(f => ({{frac: f, dragging: false}}));
let dragSplitIdx = -1, dragStartX = 0, dragStartFrac = 0;

const img = new Image();
img.onload = () => {{
  let dw = Math.min(IMG_W, MAX_W), dh = dw * IMG_H / IMG_W;
  if (dh > MAX_H) {{ dh = MAX_H; dw = dh * IMG_W / IMG_H; }}
  canvas.width = Math.round(dw); canvas.height = Math.round(dh);
  scaleX = IMG_W / canvas.width; scaleY = IMG_H / canvas.height;
  draw(); updateSplitInfo();
  // เริ่มต้น: ถือว่าทั้งภาพถูกเลือกแล้ว (ไม่บังคับให้ลาก crop ถ้าไม่ต้องการ)
  cropRect = {{x:0, y:0, w:canvas.width, h:canvas.height}};
  hasCrop = true; draw(); finalise();
}};
img.src = 'data:image/png;base64,{b64}';

function draw() {{
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
  if (hasCrop && cropRect) drawCropBox();
  if (N_BILLS > 1) drawSplitLines();
}}

function drawCropBox() {{
  const {{x, y, w, h}} = cropRect;
  ctx.fillStyle = 'rgba(0,0,0,0.30)';
  ctx.fillRect(0, 0, canvas.width, y);
  ctx.fillRect(0, y + h, canvas.width, canvas.height - y - h);
  ctx.fillRect(0, y, x, h);
  ctx.fillRect(x + w, y, canvas.width - x - w, h);
  ctx.strokeStyle = '#534AB7'; ctx.lineWidth = 2; ctx.setLineDash([]);
  ctx.strokeRect(x, y, w, h);
  [[x,y],[x+w,y],[x,y+h],[x+w,y+h]].forEach(([cx,cy]) => {{
    ctx.fillStyle = '#534AB7'; ctx.fillRect(cx-5, cy-5, 10, 10);
  }});
  ctx.strokeStyle = 'rgba(255,255,255,0.35)'; ctx.lineWidth = 0.5; ctx.setLineDash([4,4]);
  ctx.beginPath();
  ctx.moveTo(x+w/3,y); ctx.lineTo(x+w/3,y+h);
  ctx.moveTo(x+w*2/3,y); ctx.lineTo(x+w*2/3,y+h);
  ctx.moveTo(x,y+h/3); ctx.lineTo(x+w,y+h/3);
  ctx.moveTo(x,y+h*2/3); ctx.lineTo(x+w,y+h*2/3);
  ctx.stroke(); ctx.setLineDash([]);
}}

function drawSplitLines() {{
  splitLines.forEach((sl, i) => {{
    const px = Math.round(sl.frac * canvas.width);
    ctx.shadowColor = 'rgba(220,0,0,0.4)';
    ctx.shadowBlur = 6;
    ctx.strokeStyle = '#EF4444'; ctx.lineWidth = 2.5;
    ctx.setLineDash([8, 5]);
    ctx.beginPath(); ctx.moveTo(px, 0); ctx.lineTo(px, canvas.height); ctx.stroke();
    ctx.setLineDash([]); ctx.shadowBlur = 0;

    const mid = canvas.height / 2;
    ctx.fillStyle = '#EF4444';
    ctx.beginPath(); ctx.arc(px, mid, 10, 0, Math.PI*2); ctx.fill();
    ctx.fillStyle = '#fff';
    ctx.font = 'bold 9px sans-serif'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.fillText('↔', px, mid);

    const prevPx = i === 0 ? 0 : Math.round(splitLines[i-1].frac * canvas.width);
    const labelX = (prevPx + px) / 2;
    const lbl = `บิล ${{i+1}}`;
    const tw = ctx.measureText(lbl).width + 16;
    ctx.fillStyle = 'rgba(239,68,68,0.85)';
    ctx.beginPath();
    ctx.roundRect ? ctx.roundRect(labelX - tw/2, 6, tw, 20, 6)
                  : ctx.rect(labelX - tw/2, 6, tw, 20);
    ctx.fill();
    ctx.fillStyle = '#fff'; ctx.font = 'bold 11px sans-serif';
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.fillText(lbl, labelX, 16);
  }});

  const lastPx = splitLines.length > 0
    ? Math.round(splitLines[splitLines.length-1].frac * canvas.width) : 0;
  const lastLbl = `บิล ${{N_BILLS}}`;
  const tw2 = ctx.measureText(lastLbl).width + 16;
  const labelX2 = (lastPx + canvas.width) / 2;
  ctx.fillStyle = 'rgba(239,68,68,0.85)';
  ctx.beginPath();
  ctx.roundRect ? ctx.roundRect(labelX2 - tw2/2, 6, tw2, 20, 6)
                : ctx.rect(labelX2 - tw2/2, 6, tw2, 20);
  ctx.fill();
  ctx.fillStyle = '#fff'; ctx.font = 'bold 11px sans-serif';
  ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
  ctx.fillText(lastLbl, labelX2, 16);
}}

function getSplitHitIdx(px) {{
  for (let i = 0; i < splitLines.length; i++) {{
    const lx = splitLines[i].frac * canvas.width;
    if (Math.abs(px - lx) <= 14) return i;
  }}
  return -1;
}}

function updateSplitInfo() {{
  if (N_BILLS <= 1) return;
  const info = document.getElementById('split-info');
  info.style.display = 'block';
  let fracs = [0, ...splitLines.map(s => s.frac), 1.0];
  let parts = [];
  for (let i = 0; i < N_BILLS; i++) {{
    const pct = Math.round((fracs[i+1] - fracs[i]) * 100);
    const px = Math.round(fracs[i+1] * IMG_W) - Math.round(fracs[i] * IMG_W);
    parts.push(`บิล ${{i+1}}: ${{pct}}% (${{px}}px)`);
  }}
  info.innerHTML = '🔴 จุดตัด: ' + splitLines.map((s,i) =>
    `<b>${{Math.round(s.frac * IMG_W)}}px</b>`).join(' · ') +
    '<br>' + parts.join(' &nbsp;|&nbsp; ');
}}

function getPos(e) {{
  const r = canvas.getBoundingClientRect();
  const cx = (e.touches ? e.touches[0].clientX : e.clientX) - r.left;
  const cy = (e.touches ? e.touches[0].clientY : e.clientY) - r.top;
  return [Math.max(0, Math.min(cx, canvas.width)), Math.max(0, Math.min(cy, canvas.height))];
}}

canvas.addEventListener('mousedown', e => {{
  const [px, py] = getPos(e);
  dragSplitIdx = getSplitHitIdx(px);
  if (dragSplitIdx >= 0) {{
    dragStartX = px; dragStartFrac = splitLines[dragSplitIdx].frac;
    canvas.style.cursor = 'ew-resize'; return;
  }}
  drawing = true; sx = px; sy = py; hasCrop = false;
}});

canvas.addEventListener('mousemove', e => {{
  const [px, py] = getPos(e);
  if (dragSplitIdx >= 0) {{
    let newFrac = dragStartFrac + (px - dragStartX) / canvas.width;
    const minF = dragSplitIdx > 0 ? splitLines[dragSplitIdx-1].frac + 0.03 : 0.03;
    const maxF = dragSplitIdx < splitLines.length-1 ? splitLines[dragSplitIdx+1].frac - 0.03 : 0.97;
    newFrac = Math.max(minF, Math.min(maxF, newFrac));
    splitLines[dragSplitIdx].frac = newFrac;
    draw(); updateSplitInfo(); finalise(); return;
  }}
  if (drawing) {{
    ex = px; ey = py;
    let w = ex - sx, h = ey - sy;
    if (A4) h = Math.abs(w) * 1.4142 * Math.sign(h);
    cropRect = {{x: Math.min(sx,ex), y: Math.min(sy,ey), w: Math.abs(w), h: Math.abs(h)}};
    hasCrop = true; draw(); return;
  }}
  canvas.style.cursor = getSplitHitIdx(px) >= 0 ? 'ew-resize' : 'crosshair';
}});

canvas.addEventListener('mouseup', e => {{
  if (dragSplitIdx >= 0) {{ dragSplitIdx = -1; canvas.style.cursor = 'crosshair'; draw(); finalise(); return; }}
  drawing = false; finalise();
}});
canvas.addEventListener('mouseleave', () => {{
  if (dragSplitIdx >= 0) {{ dragSplitIdx = -1; draw(); finalise(); }}
  drawing = false;
}});
canvas.addEventListener('touchstart', e => {{
  e.preventDefault();
  const [px,py] = getPos(e);
  dragSplitIdx = getSplitHitIdx(px);
  if (dragSplitIdx >= 0) {{ dragStartX = px; dragStartFrac = splitLines[dragSplitIdx].frac; return; }}
  drawing = true; sx = px; sy = py; hasCrop = false;
}}, {{passive:false}});
canvas.addEventListener('touchmove', e => {{
  e.preventDefault();
  const [px,py] = getPos(e);
  if (dragSplitIdx >= 0) {{
    let newFrac = dragStartFrac + (px - dragStartX) / canvas.width;
    const minF = dragSplitIdx > 0 ? splitLines[dragSplitIdx-1].frac + 0.03 : 0.03;
    const maxF = dragSplitIdx < splitLines.length-1 ? splitLines[dragSplitIdx+1].frac - 0.03 : 0.97;
    splitLines[dragSplitIdx].frac = Math.max(minF, Math.min(maxF, newFrac));
    draw(); updateSplitInfo(); return;
  }}
  if (!drawing) return;
  ex = px; ey = py;
  let w = ex-sx, h = ey-sy;
  if (A4) h = Math.abs(w)*1.4142*Math.sign(h);
  cropRect = {{x:Math.min(sx,ex),y:Math.min(sy,ey),w:Math.abs(w),h:Math.abs(h)}};
  hasCrop = true; draw();
}}, {{passive:false}});
canvas.addEventListener('touchend', e => {{
  e.preventDefault();
  if (dragSplitIdx >= 0) {{ dragSplitIdx = -1; draw(); finalise(); return; }}
  drawing = false; finalise();
}}, {{passive:false}});

function finalise() {{
  if (!hasCrop || !cropRect || cropRect.w < 10 || cropRect.h < 10) return;
  const ox = Math.round(cropRect.x * scaleX), oy = Math.round(cropRect.y * scaleY);
  const ow = Math.round(cropRect.w * scaleX), oh = Math.round(cropRect.h * scaleY);
  const splitPx = splitLines.map(s => Math.round(s.frac * IMG_W));
  document.getElementById('result').style.display = 'block';
  document.getElementById('result').textContent =
    `เลือก: ${{ow}}×${{oh}} px (ตำแหน่ง ${{ox}},${{oy}})` +
    (splitPx.length ? `  |  จุดตัด: ${{splitPx.join(', ')}} px` : '');
  const btn = document.getElementById('btn-confirm');
  btn.disabled = false;
  const dataObj = {{x:ox, y:oy, w:ow, h:oh, splits:splitPx}};
  btn.dataset.crop = JSON.stringify(dataObj);

  const out = document.getElementById('data-out');
  out.style.display = 'block';
  out.value = JSON.stringify(dataObj);
}}

document.getElementById('btn-confirm').onclick = () => {{
  const out = document.getElementById('data-out');
  out.select();
  document.execCommand('copy');
  out.style.background = '#d1fae5';
}};

document.getElementById('btn-reset').onclick = () => {{
  hasCrop = false; cropRect = null;
  splitLines = SPLIT_FRACS.map(f => ({{frac: f, dragging: false}}));
  document.getElementById('result').style.display = 'none';
  document.getElementById('btn-confirm').disabled = true;
  document.getElementById('data-out').style.display = 'none';
  draw(); updateSplitInfo();
}};
</script></body></html>"""

# ─────────────────────────────────────────────────────────────────────────────
# Batch mode: ประมวลผลหลายไฟล์พร้อมกัน (1 ไฟล์ = 1 บิล)
# ─────────────────────────────────────────────────────────────────────────────
def run_batch_analysis(files: list, progress_cb=None, auto_detect_multi: bool = False,
                        ocr_engine: str = "tesseract") -> list:
    """
    files: list of (filename, bytes)
    progress_cb: callable(i, n, filename) เรียกก่อนประมวลผลแต่ละไฟล์ (สำหรับแสดง progress)
    auto_detect_multi: ถ้า True จะพยายามตรวจจับว่าภาพหนึ่งมีหลายใบเสร็จไหม
        (เช่น ถ่ายบนโต๊ะรวมหลายใบ) แล้วแยกประมวลผลเป็นบิลละรายการอัตโนมัติ
        พร้อมทำพื้นหลังขาวให้ก่อน OCR เพื่อความแม่นยำที่ดีขึ้น
    คืนค่า: list of bill dicts พร้อม OCR แล้ว (รูปแบบเดียวกับ S.all_bills)
    """
    results = []
    n = len(files)
    for i, (fname, fbytes) in enumerate(files, 1):
        if progress_cb: progress_cb(i, n, fname)
        try:
            pil = Image.open(io.BytesIO(fbytes)).convert("RGB")
            img_cv = pil_to_cv(pil)

            if auto_detect_multi:
                sub_crops = auto_crop_receipts(img_cv)
            else:
                sub_crops = [img_cv]

            if len(sub_crops) == 1:
                text  = run_ocr(sub_crops[0], engine=ocr_engine)
                bill  = extract_receipt(text)
                items = extract_items_cj(text)
                results.append({"filename": fname, "bill": bill, "items": items,
                                "raw_text": text, "image": img_to_bytes_png(sub_crops[0])})
            else:
                for ci, crop in enumerate(sub_crops, 1):
                    label = f"{fname} — บิล {ci}"
                    text  = run_ocr(crop, engine=ocr_engine)
                    bill  = extract_receipt(text)
                    items = extract_items_cj(text)
                    results.append({"filename": label, "bill": bill, "items": items,
                                    "raw_text": text, "image": img_to_bytes_png(crop)})
        except Exception as e:
            results.append({"filename": fname,
                            "bill": {"date":"ไม่พบ","time":"ไม่พบ","branch":"ไม่พบ","name":"ไม่พบ",
                                     "total_amount":0.0,"cash":0.0,"change":0.0,
                                     "pos_id":"ไม่พบ","rcpt_no":"ไม่พบ","tax_id":"ไม่พบ","user":"ไม่พบ"},
                            "items": [], "raw_text": f"[ERROR] {e}", "image": None})
    return results


def run_batch_mode_ui():
    """
    โหมด Batch: อัปโหลดได้ทั้งโฟลเดอร์ (multi-select ไฟล์ทั้งหมดในครั้งเดียว)
    สมมติฐาน: 1 ไฟล์ = 1 บิล (ไม่มีการ crop/แยกบิลในรูป)
    วิเคราะห์ทุกไฟล์พร้อมกันในคลิกเดียว
    """
    st.markdown(f'<p class="sec-header">{t("upload_label")}</p>', unsafe_allow_html=True)
    st.markdown(f'<div class="hint-box">{t("batch_upload_hint")}</div>', unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "batch_files", type=["png","jpg","jpeg","heic"],
        accept_multiple_files=True, label_visibility="collapsed", key="batch_uploader")

    if uploaded:
        new_files = [(f.name, f.read()) for f in uploaded]
        names_new = [x[0] for x in new_files]
        names_old = [x[0] for x in S.batch_files]
        if names_new != names_old:
            S.batch_files = new_files
            S.all_bills = []

    if not S.batch_files:
        st.info("👆 เลือกไฟล์ภาพทั้งหมดที่ต้องการวิเคราะห์ (เลือกได้ทีเดียวทั้งโฟลเดอร์)"
                if S.lang=="th" else
                "👆 Select all image files you want to analyze (can select an entire folder at once)")
        return

    st.success(t("batch_found")(len(S.batch_files)))

    # ── พรีวิวรูปย่อทั้งหมด ──
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

    auto_detect = st.checkbox(
        "🪄 ตรวจจับหลายใบเสร็จในภาพเดียวอัตโนมัติ + ทำพื้นหลังขาว"
        if S.lang=="th" else
        "🪄 Auto-detect multiple receipts per image + whiten background",
        value=False, key="batch_auto_detect",
        help="เปิดใช้เมื่อถ่ายภาพหลายใบเสร็จรวมกัน (เช่น วางบนโต๊ะ) "
             "ระบบจะครอปแยกแต่ละใบและลบพื้นหลัง/เงาออกให้อัตโนมัติ ช่วยให้ OCR แม่นขึ้น"
             if S.lang=="th" else
             "Enable when a photo contains multiple receipts together (e.g. laid out "
             "on a table). The app will auto-crop each one and remove background/shadow "
             "to improve OCR accuracy.")

    # ── พรีวิวผลการตัด/แก้ไขภาพ ก่อนกดวิเคราะห์ ──
    # ให้เห็นว่าระบบครอป/แก้มุมเอียง/ลบพื้นหลังถูกต้องไหม ก่อนเสียเวลา OCR จริง
    if auto_detect and S.batch_files:
        st.markdown("**👁️ พรีวิวผลการตัด/แก้ไขภาพ**" if S.lang=="th"
                     else "**👁️ Preview of cropped/corrected images**")
        preview_n = min(len(S.batch_files), 3)
        st.caption(f"แสดงตัวอย่าง {preview_n} จาก {len(S.batch_files)} ไฟล์แรก"
                   if S.lang=="th" else
                   f"Showing {preview_n} of the first {len(S.batch_files)} files")

        for fname, fbytes in S.batch_files[:preview_n]:
            try:
                pil = Image.open(io.BytesIO(fbytes)).convert("RGB")
                img_cv = pil_to_cv(pil)
                sub_crops = auto_crop_receipts(img_cv)
            except Exception as e:
                st.warning(f"⚠️ {fname}: {e}")
                continue

            st.markdown(f"📄 **{fname}**")
            pc1, pc2 = st.columns([1, 2])
            with pc1:
                st.image(pil, use_container_width=True,
                         caption="ต้นฉบับ" if S.lang=="th" else "Original")
            with pc2:
                if len(sub_crops) <= 1:
                    st.image(cv_to_pil(sub_crops[0]), use_container_width=True,
                             caption="หลังแก้ไข (พบ 1 บิล)" if S.lang=="th"
                                     else "After correction (1 receipt found)")
                else:
                    st.caption(f"พบ {len(sub_crops)} บิลในภาพนี้" if S.lang=="th"
                               else f"Found {len(sub_crops)} receipts in this image")
                    sub_cols = st.columns(min(len(sub_crops), 3))
                    for si, sc in enumerate(sub_crops):
                        with sub_cols[si % len(sub_cols)]:
                            st.image(cv_to_pil(sc), use_container_width=True,
                                     caption=f"บิล {si+1}" if S.lang=="th" else f"Receipt {si+1}")
            st.markdown("---")

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
        results = run_batch_analysis(S.batch_files, progress_cb=_cb, auto_detect_multi=auto_detect,
                                      ocr_engine=S.ocr_engine)
        progress_bar.progress(1.0, text=t("batch_done")(len(results)))
        S.all_bills = results
        st.rerun()

    # ── RESULTS (ใช้ layout เดียวกับโหมดปกติ) ──
    if S.all_bills:
        st.divider()
        st.success(t("found")(len(S.all_bills)))
        for idx, b in enumerate(S.all_bills):
            with st.expander(f"📄 {b['filename']}", expanded=False):
                ic, dc = st.columns([1,2])
                with ic:
                    if b.get('image'): st.image(b['image'], use_container_width=True)
                with dc:
                    d = b['bill']
                    st.markdown("**ตรวจสอบ / แก้ไข**")
                    r1 = st.columns(4)
                    d['date']    = r1[0].text_input("วันที่",          d['date'],    key=f"bdt{idx}")
                    d['time']    = r1[1].text_input("เวลา",            d['time'],    key=f"btm{idx}")
                    d['pos_id']  = r1[2].text_input("รหัสสาขา/POS",   d['pos_id'],  key=f"bps{idx}")
                    d['rcpt_no'] = r1[3].text_input("เลขที่ใบเสร็จ",  d['rcpt_no'], key=f"brc{idx}")
                    r2 = st.columns(3)
                    # ใช้ text_input แทน number_input สำหรับยอดเงิน — เพราะ <input type="number">
                    # ของ HTML บางครั้งค้างค่า min/max จาก render ก่อนหน้าไว้ใน browser
                    # ทำให้พิมพ์แก้ค่าใหม่ไม่ได้ (โดยเฉพาะเมื่อ OCR ค่าเดิมสูงผิดปกติ)
                    tot_str = r2[0].text_input("ยอดรวม",  f"{float(d['total_amount']):.2f}", key=f"btot{idx}")
                    csh_str = r2[1].text_input("เงินสด",  f"{float(d['cash']):.2f}",         key=f"bcsh{idx}")
                    chg_str = r2[2].text_input("เงินทอน", f"{float(d['change']):.2f}",        key=f"bchg{idx}")
                    d['total_amount'] = parse_price(tot_str)
                    d['cash']         = parse_price(csh_str)
                    d['change']       = parse_price(chg_str)

                mc = st.columns(3)
                mc[0].metric("💰 ยอดรวม",  f"{d['total_amount']:.2f} ฿")
                mc[1].metric("💵 เงินสด",  f"{d['cash']:.2f} ฿")
                mc[2].metric("🔄 เงินทอน", f"{d['change']:.2f} ฿")

                if b['items']:
                    st.markdown("**🛒 รายการสินค้า**")
                    st.dataframe(pd.DataFrame(b['items']), use_container_width=True, hide_index=True)
                else:
                    st.info(t("no_items"))

                with st.expander(f"🔬 {t('raw_text')} + debug"):
                    st.text_area("Raw OCR", b['raw_text'], height=160, key=f"braw{idx}", disabled=True)
                    st.json({k:v for k,v in b['bill'].items()})

        st.divider()
        st.download_button(t("download"), data=build_excel(S.all_bills),
                           file_name="receipts_batch.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# MAIN UI
# ─────────────────────────────────────────────────────────────────────────────
def main():
    col_t, col_l = st.columns([5,1])
    with col_t: st.markdown(f"## {t('title')}")
    with col_l:
        lc = st.radio("", ["ไทย","EN"], horizontal=True,
                      label_visibility="collapsed",
                      index=0 if S.lang=="th" else 1)
        nl = "th" if lc=="ไทย" else "en"
        if nl != S.lang: S.lang = nl; st.rerun()

    # ── เลือก OCR Engine ─────────────────────────────────────────────────
    with st.expander(
        "⚙️ ตัวเลือก OCR Engine" if S.lang=="th" else "⚙️ OCR Engine options",
        expanded=False
    ):
        vision_ready = is_vision_api_configured()
        engine_options = (
            ["🆓 Tesseract (ฟรี, รันในเครื่อง)", "🎯 Google Cloud Vision API (แม่นกว่ามาก)"]
            if S.lang == "th" else
            ["🆓 Tesseract (free, local)", "🎯 Google Cloud Vision API (much more accurate)"]
        )
        current_idx = 1 if S.ocr_engine == "vision" else 0
        choice = st.radio(
            "เลือก OCR Engine" if S.lang=="th" else "Choose OCR Engine",
            engine_options, index=current_idx, key="ocr_engine_radio",
            label_visibility="collapsed",
        )
        new_engine = "vision" if engine_options.index(choice) == 1 else "tesseract"
        if new_engine != S.ocr_engine:
            S.ocr_engine = new_engine

        if S.ocr_engine == "vision":
            if vision_ready:
                st.success(
                    "✅ ตั้งค่า Google Vision API key แล้ว — พร้อมใช้งาน "
                    "(ฟรี 1,000 ภาพแรกของทุกเดือน เกินจากนั้นมีค่าใช้จ่ายเล็กน้อย)"
                    if S.lang=="th" else
                    "✅ Google Vision API key configured — ready to use "
                    "(free for first 1,000 images/month, small cost after that)"
                )
            else:
                st.warning(
                    "⚠️ ยังไม่ได้ตั้งค่า API key — ระบบจะใช้ Tesseract แทนโดยอัตโนมัติ\n\n"
                    "วิธีตั้งค่า: เปิดใช้งาน Cloud Vision API ที่ Google Cloud Console "
                    "แล้วตั้งค่า environment variable `GOOGLE_VISION_API_KEY` "
                    "ก่อนรันแอป (ดูรายละเอียดในคอมเมนต์เหนือฟังก์ชัน "
                    "`run_ocr_google_vision` ในโค้ด)"
                    if S.lang=="th" else
                    "⚠️ API key not configured yet — will automatically fall back "
                    "to Tesseract\n\nSetup: enable Cloud Vision API in Google Cloud "
                    "Console, then set the `GOOGLE_VISION_API_KEY` environment "
                    "variable before running the app (see comment above the "
                    "`run_ocr_google_vision` function in the code)"
                )
        else:
            st.caption(
                "💡 Tesseract ฟรีและรันในเครื่องทั้งหมด แต่แม่นน้อยกว่า Google Vision "
                "โดยเฉพาะกับภาพที่เอียง/เงา/ยับมาก ลองสลับไปใช้ Google Vision "
                "ถ้าผลลัพธ์ยังไม่แม่นพอ"
                if S.lang=="th" else
                "💡 Tesseract is free and fully local, but less accurate than "
                "Google Vision, especially on tilted/shadowed/wrinkled photos. "
                "Try switching to Google Vision if results aren't accurate enough."
            )

    # ── เลือกโหมดการอัปโหลด ──────────────────────────────────────────────
    st.markdown(f'<p class="sec-header">{t("mode_label")}</p>', unsafe_allow_html=True)
    m1, m2 = st.columns(2)
    is_single = S.bill_count != -1   # -1 = batch mode (sentinel)
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

    # ════════════════════════════════════════════════════════════════════
    # โหมด BATCH: อัปโหลดทั้งโฟลเดอร์ วิเคราะห์ทุกไฟล์พร้อมกัน (1 ไฟล์ = 1 บิล)
    # ════════════════════════════════════════════════════════════════════
    if S.bill_count == -1:
        run_batch_mode_ui()
        return

    render_steps()

    # ── STEP 1: จำนวนบิล (CJ Express only) ──────────────────────────────
    st.markdown(f'<p class="sec-header">🏪 CJ Express</p>', unsafe_allow_html=True)

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
        st.info("👆 กรุณาเลือกจำนวนบิลก่อน" if S.lang=="th"
                else "👆 Please select the number of bills first")
        return

    st.divider()

    # ── STEP 2: อัปโหลด ─────────────────────────────────────────────────
    st.markdown(f'<p class="sec-header">{t("upload_label")}</p>', unsafe_allow_html=True)
    st.markdown(f'<div class="hint-box">{t("upload_hint")}</div>', unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "files", type=["png","jpg","jpeg","heic"],
        accept_multiple_files=True, label_visibility="collapsed", key="uploader")

    if uploaded:
        new_files = [(f.name, f.read()) for f in uploaded]
        names_new = [x[0] for x in new_files]
        names_old = [x[0] for x in S.gallery_files]
        if names_new != names_old:
            S.gallery_files=new_files; S.selected_idx=-1
            S.crop_result=None; S.step=2
            S.manual_splits_px=None; S.crop_applied=False

    # ── GALLERY ──────────────────────────────────────────────────────────
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

    # ── STEP 3: CROP ─────────────────────────────────────────────────────
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
            mode_choice = st.radio("โหมด Crop", [t("crop_free"), t("crop_a4")],
                                   key="crop_mode_r")
            mode_key = "a4" if t("crop_a4") in mode_choice else "free"
            st.info(t("crop_hint"))
            if S.bill_count > 1:
                st.caption("💡 ลากวงกลมแดงเพื่อขยับเส้นตัดบิลให้ตรงตำแหน่งจริง "
                           "แม้บิลจะวางห่างกันแค่ไม่กี่มิลก็ปรับได้")

        with crop_col:
            st_html(crop_component_html(pil_orig, mode_key, bill_count=S.bill_count),
                    height=620 if S.bill_count > 1 else 580, scrolling=False)

        # ── รับข้อมูล Crop จาก JS (วาง JSON ที่คัดลอกมา) ──
        st.markdown("**📋 วางข้อมูล Crop ที่คัดลอกจากด้านบน** "
                     "(หรือกรอกตัวเลขเอง / ข้ามเพื่อใช้รูปทั้งหมด):")
        paste_col, btn_col = st.columns([4,1])
        crop_json_str = paste_col.text_input(
            "JSON ข้อมูล Crop", value="", key="crop_json_input",
            label_visibility="collapsed",
            placeholder='วาง {"x":0,"y":0,"w":800,"h":600,"splits":[300,560]} ที่นี่')

        w0, h0 = pil_orig.size
        parsed_ok = False
        cx, cy, cw, ch, parsed_splits = 0, 0, w0, h0, []
        if crop_json_str.strip():
            try:
                data = json.loads(crop_json_str.strip())
                cx = int(data.get("x", 0)); cy = int(data.get("y", 0))
                cw = int(data.get("w", w0)); ch = int(data.get("h", h0))
                parsed_splits = [int(p) for p in data.get("splits", [])]
                parsed_ok = True
            except Exception:
                st.warning("⚠️ รูปแบบข้อมูลไม่ถูกต้อง กรุณาคัดลอกจากปุ่ม ✅ ยืนยัน Crop อีกครั้ง")

        st.caption("หรือระบุพื้นที่ Crop ด้วยตัวเลขเอง:")
        mc = st.columns(4)
        cx = mc[0].number_input("X (px)", 0, w0, cx, key="cx")
        cy = mc[1].number_input("Y (px)", 0, h0, cy, key="cy")
        cw = mc[2].number_input("W (px)", 1, w0, cw if cw>0 else w0, key="cw")
        ch = mc[3].number_input("H (px)", 1, h0, ch if ch>0 else h0, key="ch")

        if S.bill_count > 1:
            st.caption("จุดตัดบิล (พิกเซล X ในรูปต้นฉบับ — คั่นด้วยจุลภาค หากต้องการแก้ไขเอง):")
            default_splits_str = ", ".join(str(p) for p in parsed_splits) if parsed_splits else ""
            splits_str = st.text_input(
                "จุดตัดบิล", value=default_splits_str, key="splits_input",
                label_visibility="collapsed",
                placeholder="เช่น 300, 560  (เว้นว่าง = ให้ระบบเดาเอง)")
        else:
            splits_str = ""

        bc1, bc2 = st.columns(2)
        with bc1:
            if st.button(t("crop_confirm"), use_container_width=True, type="primary"):
                S.crop_result = pil_orig.crop((cx, cy, min(cx+cw,w0), min(cy+ch,h0)))
                S.crop_applied = True
                # ── เก็บตำแหน่งตัดบิลที่ user ระบุ (พิกเซลในรูปต้นฉบับ → แปลงเป็นพิกเซลในรูปที่ crop แล้ว) ──
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
                S.crop_applied = False   # ข้าม crop → ยังไม่ได้ตัดเอง ปล่อยให้ auto-split ทำงานทั้งภาพ
                S.manual_splits_px = None
                S.step=4; st.rerun()

    # ── STEP 4: OCR ──────────────────────────────────────────────────────
    working_pil = S.crop_result
    if working_pil is None and active_bytes and S.step == 4:
        working_pil = Image.open(io.BytesIO(active_bytes)).convert("RGB")

    if working_pil:
        st.divider()
        prev_c, ocr_c = st.columns([1,2])
        with prev_c:
            st.image(working_pil, caption="รูปที่จะวิเคราะห์", use_container_width=True)

            # ── แสดงตัวอย่างการแบ่งบิล + ผลแก้ไขภาพ (deskew/ลบพื้นหลัง/แก้แสงเงา) ก่อนกด OCR ──
            if S.bill_count > 1:
                img_cv_preview = pil_to_cv(working_pil)
                if S.manual_splits_px:
                    preview_crops = split_by_positions(img_cv_preview, S.manual_splits_px)
                    st.caption(f"✋ ใช้จุดตัดที่ปรับเอง → แยกได้ {len(preview_crops)} บิล")
                else:
                    preview_crops = split_receipts_image(img_cv_preview, n_expected=S.bill_count)
                    st.caption(f"🤖 ระบบตรวจจับอัตโนมัติ → แยกได้ {len(preview_crops)} บิล")
                st.caption("👁️ ภาพด้านล่างคือผลหลังตัด/ลบพื้นหลัง/แก้แสงเงา (ภาพจริงที่จะใช้วิเคราะห์)"
                           if S.lang=="th" else
                           "👁️ Images below show the result after crop/background removal/"
                           "lighting correction (the actual images used for analysis)")
                pc = st.columns(min(len(preview_crops), 3))
                for i, crop in enumerate(preview_crops[:3]):
                    with pc[i % len(pc)]:
                        corrected_preview = whiten_background(crop)
                        st.image(corrected_preview, caption=f"บิล {i+1}", use_container_width=True)
            else:
                # บิลเดี่ยว: แสดงผลหลังแก้ไขด้วย เผื่อภาพมีเงา/เอียง/พื้นหลังไม่ขาว
                img_cv_preview = pil_to_cv(working_pil)
                results_preview, _ = find_paper_contours(img_cv_preview)
                if results_preview:
                    st.caption("👁️ ภาพหลังแก้ไข (ตัด/แก้มุมเอียง/ลบพื้นหลัง/แก้แสงเงา)"
                               if S.lang=="th" else
                               "👁️ Corrected image (crop/deskew/background removal/lighting fix)")
                    bbox0, contour0 = results_preview[0]
                    corrected_single = whiten_background(img_cv_preview, contour=contour0, bbox=bbox0)
                    st.image(corrected_single, use_container_width=True)

        with ocr_c:
            if S.bill_count > 1:
                st.caption("ถ้าผลแบ่งบิลด้านซ้ายไม่ตรง ให้ย้อนกลับไปปรับเส้นแดงในขั้นตอน Crop ใหม่")
            if st.button(t("analyze"), use_container_width=True, type="primary"):
                with st.spinner("กำลัง OCR..." if S.lang=="th" else "Running OCR..."):
                    img_cv = pil_to_cv(working_pil)

                    if S.bill_count == 1:
                        crops = [img_cv]
                    elif S.manual_splits_px:
                        # ── ใช้ตำแหน่งที่ user ปรับเองตรงๆ ไม่เดาซ้ำ ──
                        crops = split_by_positions(img_cv, S.manual_splits_px)
                    else:
                        # ── ไม่มีข้อมูล manual → auto-detect โดยรู้จำนวนบิลที่คาดไว้ ──
                        crops = split_receipts_image(img_cv, n_expected=S.bill_count)

                    # เติม/ตัดให้ตรงจำนวนบิลที่เลือกไว้ตอนแรกเสมอ
                    if len(crops) < S.bill_count:
                        crops = crops + [img_cv] * (S.bill_count - len(crops))
                    crops = crops[:S.bill_count]

                    all_bills = []
                    fname = (S.gallery_files[S.selected_idx][0]
                             if 0 <= S.selected_idx < len(S.gallery_files) else "image")
                    for ci, crop in enumerate(crops):
                        label = fname if len(crops)==1 else f"{fname} — บิล {ci+1}"
                        text  = run_ocr(crop, engine=S.ocr_engine)
                        bill  = extract_receipt(text)
                        items = extract_items_cj(text)
                        all_bills.append({"filename":label,"bill":bill,"items":items,
                                          "raw_text":text,"image":img_to_bytes_png(crop)})
                    S.all_bills = all_bills; S.step=4; st.rerun()

    # ── RESULTS ───────────────────────────────────────────────────────────
    if S.all_bills:
        st.success(t("found")(len(S.all_bills)))
        for idx, b in enumerate(S.all_bills):
            with st.expander(f"📄 {b['filename']}", expanded=True):
                ic, dc = st.columns([1,2])
                with ic:
                    if b.get('image'): st.image(b['image'], use_container_width=True)
                with dc:
                    d = b['bill']
                    st.markdown("**ตรวจสอบ / แก้ไข**")
                    r1 = st.columns(4)
                    d['date']    = r1[0].text_input("วันที่",          d['date'],    key=f"dt{idx}")
                    d['time']    = r1[1].text_input("เวลา",            d['time'],    key=f"tm{idx}")
                    d['pos_id']  = r1[2].text_input("รหัสสาขา/POS",   d['pos_id'],  key=f"ps{idx}")
                    d['rcpt_no'] = r1[3].text_input("เลขที่ใบเสร็จ",  d['rcpt_no'], key=f"rc{idx}")
                    r2 = st.columns(3)
                    # ใช้ text_input แทน number_input — กัน browser ค้างค่า min/max
                    # จาก state ก่อนหน้า ทำให้พิมพ์แก้ยอดเงินใหม่ไม่ได้
                    tot_str = r2[0].text_input("ยอดรวม",  f"{float(d['total_amount']):.2f}", key=f"tot{idx}")
                    csh_str = r2[1].text_input("เงินสด",  f"{float(d['cash']):.2f}",         key=f"csh{idx}")
                    chg_str = r2[2].text_input("เงินทอน", f"{float(d['change']):.2f}",        key=f"chg{idx}")
                    d['total_amount'] = parse_price(tot_str)
                    d['cash']         = parse_price(csh_str)
                    d['change']       = parse_price(chg_str)

                mc = st.columns(3)
                mc[0].metric("💰 ยอดรวม",  f"{d['total_amount']:.2f} ฿")
                mc[1].metric("💵 เงินสด",  f"{d['cash']:.2f} ฿")
                mc[2].metric("🔄 เงินทอน", f"{d['change']:.2f} ฿")

                if b['items']:
                    st.markdown("**🛒 รายการสินค้า**")
                    st.dataframe(pd.DataFrame(b['items']), use_container_width=True, hide_index=True)
                else:
                    st.info(t("no_items"))

                with st.expander(f"🔬 {t('raw_text')} + debug"):
                    st.text_area("Raw OCR", b['raw_text'], height=160, key=f"raw{idx}", disabled=True)
                    st.json({k:v for k,v in b['bill'].items()})

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
