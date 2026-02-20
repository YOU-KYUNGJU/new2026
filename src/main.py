import os
import re
import sys
import time
import shutil
import threading
import configparser
import gc  # Garbage Collection explicit call
from datetime import datetime
from collections import deque
from threading import Lock

from pdf2image import convert_from_path
import pytesseract
from PIL import ImageOps, ImageFilter

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

# exe 파일만들기(권장)
# pyinstaller --onefile --windowed --name FITI_PDF_ReportSplitter main.py
# pyinstaller FITI_PDF_ReportSplitter.spec

__version__ = "3.6.0"
# ================== 설정 파일 로드 ==================
def load_config():
    config = configparser.ConfigParser()
    
    # 우선순위:
    # 1. 실행 파일(또는 스크립트)과 같은 폴더 (배포판/단일폴더)
    # 2. 상위 폴더의 config/config.ini (개발환경: src/main.py -> config/config.ini)
    
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 1. 같은 폴더
    path_same_dir = os.path.join(base_dir, 'config.ini')
    
    # 2. ../config/config.ini
    path_dev_dir = os.path.join(os.path.dirname(base_dir), 'config', 'config.ini')

    if os.path.exists(path_same_dir):
        config_path = path_same_dir
    elif os.path.exists(path_dev_dir):
        config_path = path_dev_dir
    else:
        # 없으면 기본값으로 같은 폴더에 생성
        config_path = path_same_dir

    # 기본값 설정
    defaults = {
        'PATHS': {
            'POPPLER_PATH': r"C:\poppler\Library\bin",
            'TESSERACT_EXE': r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        },
        'OCR': {
            'DPI': '500',
            'BATCH_SIZE': '10'
        },
        'WATCHER': {
            'SCAN_INTERVAL_SEC': '1.0',
            'STABLE_CHECK_SEC': '0.7',
            'STABLE_RETRY': '3'
        },
        'SERVER': {
            'BASE_ROOT_PH': r"\\192.168.1.7\유해물질시험팀\3. 폼알데히드,pH파트\pH\RDMS\H111분석일지 SCAN",
            'BASE_ROOT_FORMALDEHYDE': r"\\192.168.1.7\유해물질시험팀\3. 폼알데히드,pH파트\FORMALDEHYDE\RDMS\완료",
            'ALT_SERVER_NAME': r"\\fiti_fileserver"
        }
    }

    if not os.path.exists(config_path):
        # 설정 파일이 없으면 기본값으로 생성
        config.read_dict(defaults)
        # config 폴더가 없으면 생성 시도 (개발환경 고려)
        try:
            os.makedirs(os.path.dirname(config_path), exist_ok=True)
            with open(config_path, 'w', encoding='utf-8') as f:
                config.write(f)
        except Exception:
            # 실패하면 그냥 메모리 상의 기본값 사용
            pass
    else:
        config.read(config_path, encoding='utf-8')

    return config

CONFIG = load_config()

# ================== 환경 설정 (Config 적용) ==================
POPPLER_PATH = CONFIG['PATHS'].get('POPPLER_PATH', r"C:\poppler\Library\bin")
TESSERACT_EXE = CONFIG['PATHS'].get('TESSERACT_EXE', r"C:\Program Files\Tesseract-OCR\tesseract.exe")
DPI = CONFIG['OCR'].getint('DPI', 500)
BATCH_SIZE = CONFIG['OCR'].getint('BATCH_SIZE', 10)  # 메모리 최적화를 위한 배치 크기

SCAN_INTERVAL_SEC = CONFIG['WATCHER'].getfloat('SCAN_INTERVAL_SEC', 1.0)
STABLE_CHECK_SEC = CONFIG['WATCHER'].getfloat('STABLE_CHECK_SEC', 0.7)
STABLE_RETRY = CONFIG['WATCHER'].getint('STABLE_RETRY', 3)

BASE_ROOT_PH = CONFIG['SERVER'].get('BASE_ROOT_PH', r"\\192.168.1.7\유해물질시험팀\3. 폼알데히드,pH파트\pH\RDMS\H111분석일지 SCAN")
BASE_ROOT_FORMALDEHYDE = CONFIG['SERVER'].get('BASE_ROOT_FORMALDEHYDE', r"\\192.168.1.7\유해물질시험팀\3. 폼알데히드,pH파트\FORMALDEHYDE\RDMS\완료")
ALT_SERVER_NAME = CONFIG['SERVER'].get('ALT_SERVER_NAME', r"\\fiti_fileserver")
# =============================================

pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE

# ---------- pypdf (없으면 설치) ----------
try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pypdf"])
    from pypdf import PdfReader, PdfWriter

# ---------- 정규식 ----------
RE_RECEIPT_ANY = re.compile(r"H\d{3}\s*[-]?\s*\d{2}\s*[-]?\s*\d{5}", re.IGNORECASE)

REFS = [
    "FITI-T04-001-02(REV.0)",
    "FITI-IQP-011-13(REV.0)",
]

# ---------- 전역 상태 ----------
stop_requested = False
paused_by_checkbox = False

queue_lock = threading.Lock()
pdf_queue = deque()            # 처리 대기 큐: [(pdf_path, forced_code_or_None), ...]
seen_files = set()             # 이미 큐에 넣었거나 처리한 파일들
currently_processing = None    # 현재 처리 중인 pdf 경로
last_done_filename = None   # ✅ 마지막 완료 파일명(표시용)



# ================== 유틸 ==================
def now_ts():
    return datetime.now().strftime("%H:%M:%S")

def safe_filename(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", s)

def norm_keep(s: str) -> str:
    return re.sub(r"\s+", "", s).upper()

def norm_code(s: str) -> str:
    s = re.sub(r"\s+", "", s)
    s = s.replace("O", "0").replace("o", "0")
    s = s.replace("I", "1").replace("l", "1")
    return s.upper()

def preprocess_gray_sharp(img):
    g = img.convert("L")
    g = ImageOps.autocontrast(g)
    g = g.filter(ImageFilter.MedianFilter(size=3))
    g = g.filter(ImageFilter.SHARPEN)
    return g

def binarize(img_gray, thr):
    return img_gray.point(lambda x, t=thr: 255 if x > t else 0)

def file_is_stable(path: str) -> bool:
    """
    파일이 쓰기 중일 수 있으므로(복사/다운로드 중),
    size가 일정해질 때까지 짧게 확인.
    """
    try:
        prev = os.path.getsize(path)
        for _ in range(STABLE_RETRY):
            time.sleep(STABLE_CHECK_SEC)
            cur = os.path.getsize(path)
            if cur != prev:
                prev = cur
            else:
                return True
        return False
    except Exception:
        return False


# ================== 출력 폴더(로컬/서버) ==================
# exe로 빌드 시, sys.executable이 exe 경로
RUN_DIR = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
SPLIT_OUTPUT_DIR = os.path.join(RUN_DIR, "split_output")
os.makedirs(SPLIT_OUTPUT_DIR, exist_ok=True)

# ================== 텍스트 로그 파일 ==================
LOG_DIR = os.path.join(RUN_DIR, "logs")
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE_PATH = os.path.join(LOG_DIR, f"FITI_PDF_ReportSplitter_{datetime.now().strftime('%Y%m%d')}.txt")
log_file_lock = Lock()

def resolve_base_root(primary_root: str) -> str:
    """
    \\192.168.1.7 경로가 실제로 존재하지 않으면 \\fiti_fileserver 로 대체
    """
    try:
        if os.path.exists(primary_root):
            return primary_root
    except Exception:
        pass

    # 서버명 대체
    alt_root = primary_root.replace(r"\\192.168.1.7", ALT_SERVER_NAME, 1)
    return alt_root


def get_output_dir_by_rdms(rdms_code: str) -> str:
    """
    RDMS 코드에 따라 저장 폴더 결정
    - P: pH 서버 폴더 (연/월/일 자동 생성)
    - F: Formaldehyde 서버 폴더 (연/월/일 자동 생성)
    - 그 외: 로컬 split_output
    """
    code = (rdms_code or "").strip().upper()

    today = datetime.now()
    yyyy = today.strftime("%Y")
    mm = today.strftime("%m")
    yyyymmdd = today.strftime("%Y%m%d")

    if code == "P":
        base_root = resolve_base_root(BASE_ROOT_PH)  # ✅ Config 적용
        out_dir = os.path.join(base_root, yyyy, mm, yyyymmdd)

    elif code == "F":
        base_root = resolve_base_root(BASE_ROOT_FORMALDEHYDE)  # ✅ Config 적용
        out_dir = os.path.join(base_root, yyyy, mm, yyyymmdd)

    else:
        out_dir = SPLIT_OUTPUT_DIR

    # (서버/로컬 모두) 폴더 자동 생성
    os.makedirs(out_dir, exist_ok=True)
    return out_dir

def ensure_dir_or_fallback(preferred_dir: str, rdms_code: str, log):
    """
    ✅ 서버 폴더 생성/접근 실패 시 로컬로 자동 fallback
    - preferred_dir 생성/쓰기 테스트
    - 실패하면 split_output/_fallback_server_failed/YYYYMMDD 로 저장
    """
    # 로컬 fallback 폴더
    day = datetime.now().strftime("%Y%m%d")
    fallback_dir = os.path.join(SPLIT_OUTPUT_DIR, "_fallback_server_failed", day)
    os.makedirs(fallback_dir, exist_ok=True)

    try:
        os.makedirs(preferred_dir, exist_ok=True)

        # 쓰기 테스트(0바이트 파일 생성 후 삭제)
        test_path = os.path.join(preferred_dir, f"__write_test__{int(time.time())}.tmp")
        with open(test_path, "wb") as f:
            f.write(b"")
        os.remove(test_path)

        return preferred_dir, False
    except Exception as e:
        log(f"⚠️ 서버 저장 폴더 접근/생성 실패 → 로컬로 저장합니다. (RDMS={rdms_code})")
        log(f"   원인: {e}")
        log(f"   로컬 저장 폴더: {fallback_dir}")
        return fallback_dir, True


# ================== 하단 양식 검출 ==================
def make_substrings(ref: str, k=12):
    ref_n = norm_keep(ref)
    if len(ref_n) <= k:
        return {ref_n}
    return {ref_n[i:i+k] for i in range(0, len(ref_n) - k + 1)}

REF_SUBS = [(ref, make_substrings(ref, 12)) for ref in REFS]

def bottom_has_fiti_code(ocr_norm: str):
    for ref, subs in REF_SUBS:
        for sub in subs:
            if sub and sub in ocr_norm:
                return True, ref
    return False, ""

def ocr_bottom_band(img):
    """
    하단 띠 OCR → FITI-xxx(Rev.0) 검출
    """
    w, h = img.size
    bottom = img.crop((
        int(w * 0.00),
        int(h * 0.72),
        int(w * 1.00),
        int(h * 1.00),
    ))

    g = preprocess_gray_sharp(bottom)
    thr_list = [160, 175, 190]
    best_norm = ""

    config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-().'

    for thr in thr_list:
        bw = binarize(g, thr)
        raw = pytesseract.image_to_string(bw, lang="eng", config=config)
        n = norm_keep(raw)

        hit, which = bottom_has_fiti_code(n)
        if hit:
            return True, which, n

        if len(n) > len(best_norm):
            best_norm = n

    hit, which = bottom_has_fiti_code(best_norm)
    return hit, (which if hit else ""), best_norm


# ================== 접수번호 OCR ==================
def extract_receipt_no_anywhere_top(img):
    """
    접수번호 OCR
    - 우측(시험일/승인칸) 제외하고 좌측만 OCR
    - 위치 흔들림 대응 위해 ROI 여러개 시도
    """
    w, h = img.size

    top_band = img.crop((
        int(w * 0.00),
        int(h * 0.00),
        int(w * 0.66),   # 우측 제외
        int(h * 0.34),
    ))

    g = preprocess_gray_sharp(top_band)
    thr_list = [150, 165, 180, 195]
    psm_list = [6, 11, 12]
    cfg = r'--oem 3 --psm {psm} -c tessedit_char_whitelist=Hh0123456789-'

    best_candidate = None

    for thr in thr_list:
        bw = binarize(g, thr)
        for psm in psm_list:
            raw = pytesseract.image_to_string(bw, lang="eng", config=cfg.format(psm=psm))
            n = norm_code(raw)

            m = RE_RECEIPT_ANY.search(n)
            if m:
                v = re.sub(r"\s+", "", m.group(0)).upper()
                digits = re.sub(r"[^H0-9]", "", v)
                m2 = re.fullmatch(r"H(\d{3})(\d{2})(\d{5})", digits)
                if m2:
                    return f"H{m2.group(1)}-{m2.group(2)}-{m2.group(3)}"
                return v

            digits = re.sub(r"[^H0-9]", "", n)
            m2 = re.search(r"(H\d{3})(\d{2})(\d{5})", digits)
            if m2:
                best_candidate = f"{m2.group(1)}-{m2.group(2)}-{m2.group(3)}"

    if best_candidate:
        return best_candidate

    rois = [
        (0.00, 0.03, 0.66, 0.18),
        (0.00, 0.05, 0.66, 0.22),
        (0.02, 0.06, 0.62, 0.20),
        (0.00, 0.08, 0.66, 0.28),
        (0.00, 0.10, 0.66, 0.30),
    ]

    for (lx, ty, rx, by) in rois:
        roi = img.crop((int(w * lx), int(h * ty), int(w * rx), int(h * by)))
        g2 = preprocess_gray_sharp(roi)

        for thr in [150, 170, 190]:
            bw2 = binarize(g2, thr)
            raw2 = pytesseract.image_to_string(bw2, lang="eng", config=cfg.format(psm=6))
            n2 = norm_code(raw2)

            m = RE_RECEIPT_ANY.search(n2)
            if m:
                v = re.sub(r"\s+", "", m.group(0)).upper()
                digits = re.sub(r"[^H0-9]", "", v)
                m2 = re.fullmatch(r"H(\d{3})(\d{2})(\d{5})", digits)
                if m2:
                    return f"H{m2.group(1)}-{m2.group(2)}-{m2.group(3)}"
                return v

            digits = re.sub(r"[^H0-9]", "", n2)
            m2 = re.search(r"(H\d{3})(\d{2})(\d{5})", digits)
            if m2:
                return f"{m2.group(1)}-{m2.group(2)}-{m2.group(3)}"

    return None


# ================== 그룹/보정 ==================
def group_consecutive(pages):
    pages = sorted(set(pages))
    if not pages:
        return []
    groups = []
    start = prev = pages[0]
    for p in pages[1:]:
        if p == prev + 1:
            prev = p
        else:
            groups.append((start, prev))
            start = prev = p
    groups.append((start, prev))
    return groups

def parse_receipt(r: str):
    if not r:
        return None
    m = re.match(r"^(H\d{3})-(\d{2})-(\d{5})$", r.upper())
    if not m:
        return None
    return (m.group(1), int(m.group(2)), m.group(3))

def fix_group_receipt(group_pages, page_to_receipt):
    """
    ✅ 연속 페이지(그룹)에서는 '2번째 페이지 접수번호'를 우선
    """
    parsed = {}
    for p in group_pages:
        pr = parse_receipt(page_to_receipt.get(p))
        if pr:
            parsed[p] = pr

    if not parsed:
        return None

    if len(group_pages) >= 2:
        p2 = group_pages[1]
        if p2 in parsed:
            h, yy, tail = parsed[p2]
            return f"{h}-{yy:02d}-{tail}"

    p1 = group_pages[0]
    if p1 in parsed:
        h, yy, tail = parsed[p1]
        return f"{h}-{yy:02d}-{tail}"

    for p in group_pages:
        if p in parsed:
            h, yy, tail = parsed[p]
            return f"{h}-{yy:02d}-{tail}"

    return None


# ================== 저장/이동 ==================
def split_and_save(pdf_path, out_dir, groups, fixed_receipts, rdms_code, log):
    """
    분할 규칙:
    - 그룹 시작 페이지 ~ 다음 그룹 시작-1 까지 저장
    - 마지막은 PDF 끝까지
    파일명: 접수번호-입력값.pdf
    """
    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)

    used = set()

    for idx, ((gs, _ge), fixed) in enumerate(zip(groups, fixed_receipts)):
        if stop_requested:
            return

        start_page = gs
        if idx + 1 < len(groups):
            next_start = groups[idx + 1][0]
            end_page = next_start - 1
        else:
            end_page = total_pages

        if end_page < start_page:
            end_page = start_page

        if fixed:
            base = safe_filename(fixed)
        else:
            base = f"{idx + 1:04d}"

        fname = f"{base}-{rdms_code}.pdf"
        low = fname.lower()
        k = 2
        while low in used:
            fname = f"{base}-{rdms_code}_{k}.pdf"
            low = fname.lower()
            k += 1
        used.add(low)

        out_path = os.path.join(out_dir, fname)

        writer = PdfWriter()
        for p in range(start_page - 1, end_page):
            writer.add_page(reader.pages[p])

        with open(out_path, "wb") as f:
            writer.write(f)

        log(f"✅ 저장: {out_path}  (구간 {start_page}~{end_page}, FIXED={fixed})")

def move_original_to_done(pdf_path: str, split_output_dir: str, log):
    """
    원본 PDF는 항상 로컬 split_output/_inbox_done 으로 이동
    """
    done_dir = os.path.join(split_output_dir, "_inbox_done")
    os.makedirs(done_dir, exist_ok=True)

    base = os.path.basename(pdf_path)
    name, ext = os.path.splitext(base)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = os.path.join(done_dir, f"{safe_filename(name)}_{ts}{ext}")

    try:
        shutil.move(pdf_path, dst)
        log(f"📦 원본 이동: {dst}")
    except Exception as e:
        log(f"❌ 원본 이동 실패: {e}")


# ================== UI ==================
root = tk.Tk()
root.title("FITI PDF Report Splitter (Folder Watch)")

status_var = tk.StringVar(value="대기 중...")
info_var = tk.StringVar(value="")
progress_var = tk.DoubleVar(value=0)

rdms_var = tk.StringVar(value="")             # UI에서 수시 변경
auto_next_var = tk.BooleanVar(value=True)     # 체크박스(기본 ON): 로그숨김+자동진행

frm_top = tk.Frame(root)
frm_top.pack(fill="x", padx=10, pady=8)

lbl_status = tk.Label(frm_top, textvariable=status_var, font=("맑은 고딕", 11, "bold"))
lbl_status.grid(row=0, column=0, sticky="w")

lbl_info = tk.Label(frm_top, textvariable=info_var, font=("맑은 고딕", 9))
lbl_info.grid(row=1, column=0, sticky="w", pady=(2, 0))

pbar = ttk.Progressbar(frm_top, maximum=100, variable=progress_var)
pbar.grid(row=2, column=0, sticky="ew", pady=(6, 0))

frm_top.grid_columnconfigure(0, weight=1)

frm_mid = tk.Frame(root)
frm_mid.pack(fill="x", padx=10, pady=(0, 6))

tk.Label(frm_mid, text="RDMS 코드:", font=("맑은 고딕", 9, "bold")).pack(side="left")
rdms_entry = tk.Entry(frm_mid, textvariable=rdms_var, width=10)
rdms_entry.pack(side="left", padx=(6, 14))

auto_chk = tk.Checkbutton(
    frm_mid,
    text="자동 진행(로그 숨김)",
    variable=auto_next_var,
    onvalue=True,
    offvalue=False
)
auto_chk.pack(side="left")

txt = tk.Text(root, height=16, wrap="word")
txt.pack(fill="both", expand=True, padx=10, pady=(0, 10))

frm_btn = tk.Frame(root)
frm_btn.pack(fill="x", padx=10, pady=(0, 10))

exit_btn = tk.Button(frm_btn, text="종료", width=10)
exit_btn.pack(side="right")


def append_log(s: str):
    # 화면 로그와 별도로 txt 파일에도 누적 저장
    try:
        with log_file_lock:
            with open(LOG_FILE_PATH, "a", encoding="utf-8") as f:
                f.write(s + "\n")
    except Exception:
        pass

    def _append():
        txt.insert("end", s + "\n")
        txt.see("end")
    root.after(0, _append)

def set_status(s: str):
    root.after(0, lambda: status_var.set(s))

def set_info(s: str):
    root.after(0, lambda: info_var.set(s))

def set_progress(pct: float):
    root.after(0, lambda: progress_var.set(pct))

def refresh_log_visibility():
    # 체크 ON: 로그 숨김 / OFF: 로그 표시
    show = not auto_next_var.get()
    if show:
        txt.pack(fill="both", expand=True, padx=10, pady=(0, 10))
    else:
        txt.pack_forget()

def on_auto_chk_toggle():
    global paused_by_checkbox
    refresh_log_visibility()
    paused_by_checkbox = (not auto_next_var.get())
    if paused_by_checkbox:
        append_log(f"[{now_ts()}] ⏸️ 자동 진행 OFF: 다음 작업은 대기합니다. (체크를 다시 켜면 재개)")
    else:
        append_log(f"[{now_ts()}] ▶️ 자동 진행 ON: 대기 중이면 다음 작업을 진행합니다.")

auto_next_var.trace_add("write", lambda *_: on_auto_chk_toggle())

def request_exit():
    global stop_requested
    stop_requested = True
    set_status("종료 중...")
    set_info("프로그램을 종료합니다.")
    append_log(f"[{now_ts()}] 🛑 종료 요청됨")
    root.after(300, root.destroy)

exit_btn.config(command=request_exit)

def on_window_close():
    if messagebox.askyesno("종료", "프로그램을 종료할까요?"):
        request_exit()

root.protocol("WM_DELETE_WINDOW", on_window_close)


# ================== 시작 로그/폴더 ==================
append_log(f"[{now_ts()}] 감시 폴더: {RUN_DIR}")
append_log(f"[{now_ts()}] 로컬 출력 폴더: {SPLIT_OUTPUT_DIR}")
refresh_log_visibility()


# ================== 폴더 감시 스레드 ==================
def scan_loop():
    """
    RUN_DIR 최상위 + RUN_DIR 하위 1레벨 폴더들을 스캔하여 PDF 자동 처리
    - 최상위 PDF: forced_code=None (UI RDMS 사용)
    - 하위 폴더 PDF: forced_code=폴더명(예: AA, N, F, P ...)
    - split_output 및 시스템 폴더는 제외
    """
    global stop_requested

    # 스캔 제외할 폴더명(필요하면 추가)
    EXCLUDE_DIRS = {
        "split_output",
        "_inbox_done",
        "_fallback_server_failed",
        "__pycache__",
    }

    while not stop_requested:
        try:
            # 1) 최상위(RUN_DIR) PDF 스캔
            for name in os.listdir(RUN_DIR):
                if stop_requested:
                    break

                full = os.path.join(RUN_DIR, name)

                # 디렉토리면 여기서는 스킵(하위 폴더는 아래에서 처리)
                if os.path.isdir(full):
                    continue

                if not name.lower().endswith(".pdf"):
                    continue

                # split_output 내부 파일이면 제외
                try:
                    if os.path.commonpath([full, SPLIT_OUTPUT_DIR]) == SPLIT_OUTPUT_DIR:
                        continue
                except Exception:
                    pass

                if full in seen_files:
                    continue
                if not file_is_stable(full):
                    continue

                with queue_lock:
                    pdf_queue.append((full, None))  # ✅ 최상위는 UI RDMS 사용
                    seen_files.add(full)

                append_log(f"[{now_ts()}] 📥 새 PDF 감지(최상위): {name}")

            # 2) 하위 1레벨 폴더 스캔 (폴더명=강제코드)
            for dname in os.listdir(RUN_DIR):
                if stop_requested:
                    break

                dpath = os.path.join(RUN_DIR, dname)
                if not os.path.isdir(dpath):
                    continue

                # 제외 폴더
                if dname in EXCLUDE_DIRS:
                    continue

                forced_code = dname.strip().upper()
                if not forced_code:
                    continue

                for fname in os.listdir(dpath):
                    if stop_requested:
                        break

                    if not fname.lower().endswith(".pdf"):
                        continue

                    fpath = os.path.join(dpath, fname)

                    # split_output 내부면 제외(방어)
                    try:
                        if os.path.commonpath([fpath, SPLIT_OUTPUT_DIR]) == SPLIT_OUTPUT_DIR:
                            continue
                    except Exception:
                        pass

                    if fpath in seen_files:
                        continue
                    if not file_is_stable(fpath):
                        continue

                    with queue_lock:
                        pdf_queue.append((fpath, forced_code))  # ✅ 폴더명으로 강제 RDMS
                        seen_files.add(fpath)

                    append_log(f"[{now_ts()}] 📥 새 PDF 감지({forced_code}): {fname}")

        except Exception as e:
            append_log(f"[{now_ts()}] ❌ 스캔 오류: {e}")

        time.sleep(SCAN_INTERVAL_SEC)



# ================== 작업 처리 ==================
def process_one_pdf(pdf_path: str, rdms_code: str):
    """
    단일 PDF 처리 (메모리 최적화 Chunk 방식):
    - pypdf로 총 페이지 수 확인
    - BATCH_SIZE(예: 10장)씩 끊어서 이미지 변환 -> OCR -> 메모리 해제
    - 하단 양식 검출 + 접수번호 추출된 결과들을 모아둠
    - 모든 페이지 처리 후 분할 저장 로직 수행
    - 원본 이동
    """
    set_status("📄 PDF 처리 중...")
    set_info(os.path.basename(pdf_path))

    def log(msg): append_log(f"[{now_ts()}] {msg}")

    try:
        # 1. 총 페이지 수 먼저 확인
        reader = PdfReader(pdf_path)
        total_pages = len(reader.pages)
        if total_pages == 0:
            log(f"❌ 페이지 0: {pdf_path}")
            return

        matched_pages_raw = []
        page_receipt_raw = {}

        set_status("OCR 추출 중...")

        # 2. 배치 단위 처리 (메모리 절약)
        # BATCH_SIZE = 10 (Config 로드값)
        
        for start_idx in range(1, total_pages + 1, BATCH_SIZE):
            if stop_requested:
                return
            
            # 마지막 페이지 번호 계산
            end_idx = min(start_idx + BATCH_SIZE - 1, total_pages)
            
            # 해당 구간만 이미지 변환
            # pdf2image는 1-based index 사용 (first_page, last_page)
            # fmt='jpeg' 사용 시 메모리 사용량이 조금 더 줄어들 수 있음 (기본값은 ppm -> PIL)
            images = convert_from_path(
                pdf_path, 
                poppler_path=POPPLER_PATH, 
                dpi=DPI, 
                first_page=start_idx, 
                last_page=end_idx
            )
            
            # 변환된 이미지들에 대해 OCR 수행
            for i, img in enumerate(images):
                page_num = start_idx + i
                
                # UI 진행률 업데이트
                set_progress((page_num / total_pages) * 85.0)
                set_info(f"{os.path.basename(pdf_path)} | OCR Page {page_num}/{total_pages}")

                hit, which, _norm = ocr_bottom_band(img)
                if hit:
                    matched_pages_raw.append(page_num)
                    receipt = extract_receipt_no_anywhere_top(img)
                    page_receipt_raw[page_num] = receipt
                    log(f"Page {page_num}/{total_pages} | FORM ✅ ({which}) | OCR Receipt: {receipt if receipt else '❌'}")
            
            # 중요: 이미지 사용 후 즉시 메모리 해제
            del images
            gc.collect()

        # 3. OCR 완료 후 그룹핑 및 후처리 (기존 로직 동일)
        groups = group_consecutive(matched_pages_raw)
        if not groups:
            log(f"❌ 양식 페이지 미검출: {os.path.basename(pdf_path)}")
            return

        fixed_receipts = []
        log("✅ 양식 페이지 그룹(연속 묶음):")
        for (s, e) in groups:
            pages = list(range(s, e + 1))
            fixed = fix_group_receipt(pages, page_receipt_raw)
            fixed_receipts.append(fixed)
            log(f" - Pages {s}-{e}  =>  FIXED Receipt: {fixed if fixed else '❌'}")

        # ✅ RDMS 코드별 저장 폴더 결정 + 서버 실패 시 로컬 fallback
        preferred = get_output_dir_by_rdms(rdms_code)
        out_dir, fallback_used = ensure_dir_or_fallback(preferred, rdms_code, log)
        if fallback_used:
            set_status("저장 중(로컬 fallback)...")
        else:
            set_status("저장 중...")

        set_info(f"{os.path.basename(pdf_path)} | 저장 폴더: {out_dir}")
        set_progress(92.0)

        split_and_save(pdf_path, out_dir, groups, fixed_receipts, rdms_code, log)

        set_status("원본 이동 중...")
        set_info(f"{os.path.basename(pdf_path)} | 원본 이동")
        set_progress(98.0)

        move_original_to_done(pdf_path, SPLIT_OUTPUT_DIR, log)

        log(f"✅ 완료: {os.path.basename(pdf_path)} (RDMS={rdms_code})")
        set_progress(100.0)
        # ✅ 완료 표시용
        global last_done_filename
        last_done_filename = os.path.basename(pdf_path)
        set_status("✅ 처리 완료")
        set_info(f"{last_done_filename} 완료")


    except Exception as e:
        log(f"❌ 처리 오류: {e}")


def worker_loop():
    """
    큐에 들어온 PDF를 순차 처리.
    체크박스 OFF이면(로그 표시 모드) "다음 파일로 넘어가기 전 대기"
    """
    global currently_processing, paused_by_checkbox

    while not stop_requested:
        if paused_by_checkbox:
            set_status("대기 중(자동 진행 OFF)")
            set_info("체크를 다시 켜면 다음 작업 진행")
            time.sleep(0.2)
            continue

        pdf_item = None
        with queue_lock:
            if pdf_queue:
                pdf_item = pdf_queue.popleft()

        if not pdf_item:
            # ✅ 마지막 완료 파일명 표시
            if last_done_filename:
                set_status("대기 중...")
                set_info(f"감시 폴더에 PDF가 생성되면 자동 처리합니다.  |  {last_done_filename} 완료")
            else:
                set_status("대기 중...")
                set_info("감시 폴더에 PDF가 생성되면 자동 처리합니다.")
            set_progress(0)
            time.sleep(0.2)
            continue

        pdf_path, forced_code = pdf_item

        # ✅ RDMS 코드 결정:
        # - forced_code가 있으면 그걸 우선 (AA 폴더는 무조건 AA)
        # - 아니면 UI 입력값 사용
        if forced_code:
            rdms_code = forced_code
        else:
            rdms_code = safe_filename(rdms_var.get().strip()) or "X"
        if not rdms_code:
            rdms_code = "X"

        currently_processing = pdf_path
        append_log(f"[{now_ts()}] ▶ 처리 시작: {os.path.basename(pdf_path)} (RDMS={rdms_code})")
        process_one_pdf(pdf_path, rdms_code)
        currently_processing = None

        if not auto_next_var.get():
            paused_by_checkbox = True


# ================== 시작: RDMS 초기 입력 ==================
init = simpledialog.askstring("입력", "시험 항목 RDMS 코드 입력", parent=root)
if init is None:
    request_exit()
else:
    rdms_var.set(safe_filename(init.strip()))

paused_by_checkbox = (not auto_next_var.get())
refresh_log_visibility()

threading.Thread(target=scan_loop, daemon=True).start()
threading.Thread(target=worker_loop, daemon=True).start()

root.mainloop()
