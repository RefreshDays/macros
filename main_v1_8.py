# -*- coding: utf-8 -*-
"""
설비 수주 자동응답 매크로 v1.8
- 단일 실행 파일
- 후보 확정 2회
- 제한 지역 데이터 포함
- 카톡 시간 기반 로그
"""

import os
import re
import threading
import time
from typing import Dict, List, Optional, Set, Tuple

import pyautogui
import pyperclip
import pytesseract
import tkinter as tk
import win32gui
from PIL import Image, ImageEnhance, ImageOps
from pytesseract import Output
from tkinter import messagebox, scrolledtext, ttk

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.1

CONFIRM_THRESHOLD = 2
PENDING_TTL_SECONDS = 1.6
FAST_CONFIRM_THRESHOLD = 1
CHAT_TIME_STALE_MINUTES = 2
KAKAO_KEYWORDS = ["카카오톡", "KakaoTalk", "오픈채팅", "KAKAO"]
MIN_WINDOW_WIDTH = 250
MIN_WINDOW_HEIGHT = 300
TESSERACT_PATHS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
]

REGION_HIERARCHY = {
    "서울특별시": {
        "동작구": ["노량진동", "대방동", "동작동", "본동", "사당동", "상도1동", "상도동", "신대방동", "흑석동"],
        "관악구": ["남현동", "봉천동", "신림동"],
        "금천구": ["가산동", "독산동", "시흥동"],
        "영등포구": [
            "당산동", "당산동1가", "당산동2가", "당산동3가", "당산동4가", "당산동5가", "당산동6가",
            "대림동", "도림동", "문래동1가", "문래동2가", "문래동3가", "문래동4가", "문래동5가",
            "문래동6가", "신길동", "양평동", "양평동1가", "양평동2가", "양평동3가", "양평동4가",
            "양평동5가", "양평동6가", "양화동", "여의도동", "영등포동", "영등포동1가", "영등포동2가",
            "영등포동3가", "영등포동4가", "영등포동5가", "영등포동6가", "영등포동7가", "영등포동8가",
        ],
        "강서구": ["가양동", "개화동", "공항동", "과해동", "내발산동", "등촌동", "마곡동", "방화동", "염창동", "오곡동", "오쇠동", "외발산동", "화곡동"],
        "구로구": ["가리봉동", "개봉동", "고척동", "구로동", "궁동", "신도림동", "오류동", "온수동", "천왕동", "항동"],
        "마포구": [
            "공덕동", "구수동", "노고산동", "당인동", "대흥동", "도화동", "동교동", "마포동", "망원동",
            "상수동", "상암동", "서교동", "성산동", "신공덕동", "신수동", "신정동", "아현동", "연남동",
            "염리동", "용강동", "중동", "창전동", "토정동", "하중동", "합정동", "현석동",
        ],
        "양천구": ["목동", "신월동", "신정동"],
    },
    "경기도": {
        "광명시": ["가학동", "광명동", "노온사동", "소하동", "옥길동", "일직동", "철산동", "하안동"],
        "안산시": [
            "고잔동", "대부남동", "대부동동", "대부북동", "목내동", "선감동", "선부동", "성곡동", "신길동",
            "와동", "원곡동", "원시동", "초지동", "풍도동", "화정동", "건건동", "본오동", "부곡동", "사동",
            "사사동", "성포동", "수암동", "양상동", "월피동", "이동", "일동", "장상동", "장하동", "팔곡이동", "팔곡일동",
        ],
        "안양시": ["관양동", "비산동", "평촌동", "호계동", "박달동", "석수동", "안양동"],
        "군포시": ["금정동", "당동", "당정동", "대야미동", "도마교동", "둔대동", "부곡동", "산본동", "속달동"],
    },
    "인천광역시": {
        "강화군": ["강화읍", "교동면", "길상면", "내가면", "불은면", "삼산면", "서도면", "선원면", "송해면", "양도면", "양사면", "하점면", "화도면"],
        "계양구": ["갈현동", "계산동", "귤현동", "다남동", "동양동", "둑실동", "목상동", "박촌동", "방축동", "병방동", "상야동", "서운동", "선주지동", "오류동", "용종동", "이화동", "임학동", "작전동", "장기동", "평동", "하야동", "효성동"],
        "남동구": ["간석동", "고잔동", "구월동", "남촌동", "논현동", "도림동", "만수동", "서창동", "수산동", "운연동", "장수동"],
        "동구": ["금곡동", "만석동", "송림동", "송현동", "창영동", "화수동", "화평동"],
        "미추홀구": ["관교동", "도화동", "문학동", "숭의동", "용현동", "주안동", "학익동"],
        "부평구": ["갈산동", "구산동", "부개동", "부평동", "산곡동", "삼산동", "십정동", "일신동", "청천동"],
        "서구": ["가정동", "가좌동", "검암동", "경서동", "공촌동", "금곡동", "당하동", "대곡동", "마전동", "백석동", "불로동", "석남동", "시천동", "신현동", "심곡동", "연희동", "오류동", "왕길동", "원당동", "원창동", "청라동"],
        "연수구": ["동춘동", "선학동", "송도동", "연수동", "옥련동", "청학동"],
        "옹진군": ["대청면", "덕적면", "백령면", "북도면", "연평면", "영흥면", "자월면"],
        "중구": [
            "경동", "관동1가", "관동2가", "관동3가", "남북동", "내동", "답동", "덕교동", "도원동", "무의동",
            "북성동1가", "북성동2가", "북성동3가", "사동", "선린동", "선화동", "송월동1가", "송월동2가",
            "송월동3가", "송학동1가", "송학동2가", "송학동3가", "신생동", "신포동", "신흥동1가", "신흥동2가",
            "신흥동3가", "용동", "운남동", "운북동", "운서동", "유동", "율목동", "을왕동", "인현동", "전동",
            "중산동", "중앙동1가", "중앙동2가", "중앙동3가", "중앙동4가", "항동1가", "항동2가", "항동3가",
            "항동4가", "항동5가", "항동6가", "항동7가", "해안동1가", "해안동2가", "해안동3가", "해안동4가",
        ],
    },
}

JOB_NORMALIZE = [
    ("싱크대막힘", ["싱크대막힘", "싱크대 막힘"]),
    ("하수구막힘", ["하수구막힘", "하수구 막힘"]),
    ("싱막", ["싱막", "성막", "생막", "씽막"]),
    ("하막", ["하막", "하 막"]),
    ("역류", ["역류"]),
    ("누수", ["누수"]),
]
REGION_SUFFIXES = ["동", "구", "시", "군", "읍", "면", "리"]


def init_tesseract():
    for path in TESSERACT_PATHS:
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
            return True
    return False


def preprocess_image(img: Image.Image) -> Image.Image:
    gray = img.convert("L")
    gray = ImageEnhance.Contrast(gray).enhance(2.5)
    width, height = gray.size
    resized = gray.resize((max(width * 2, 1), max(height * 2, 1)), Image.Resampling.LANCZOS)
    binary = resized.point(lambda px: 255 if px > 185 else 0, mode="1").convert("L")
    return ImageOps.autocontrast(binary)


def _normalize_ocr_text(text: str) -> str:
    text = text.replace("|", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def extract_positioned_lines_from_image(img: Image.Image, lang: str = "kor+eng") -> List[Dict[str, int | str]]:
    if img is None:
        return []
    processed = preprocess_image(img)
    data = pytesseract.image_to_data(
        processed,
        lang=lang,
        config=r"--oem 3 --psm 11",
        output_type=Output.DICT,
    )
    grouped: Dict[tuple, Dict[str, List[int] | List[str]]] = {}
    total = len(data.get("text", []))
    for idx in range(total):
        raw_text = (data["text"][idx] or "").strip()
        confidence = data["conf"][idx]
        if not raw_text:
            continue
        try:
            if float(confidence) < 0:
                continue
        except (TypeError, ValueError):
            continue
        key = (data["block_num"][idx], data["par_num"][idx], data["line_num"][idx])
        entry = grouped.setdefault(key, {"words": [], "tops": []})
        entry["words"].append(raw_text)
        entry["tops"].append(int(data["top"][idx]))

    lines = []
    for entry in grouped.values():
        text = _normalize_ocr_text(" ".join(entry["words"]))
        if not text:
            continue
        avg_top = int(sum(entry["tops"]) / max(len(entry["tops"]), 1))
        lines.append({"text": text, "y": avg_top})
    lines.sort(key=lambda item: item["y"])
    return lines


def _enum_kakao_callback(hwnd, results):
    if win32gui.IsWindowVisible(hwnd):
        title = win32gui.GetWindowText(hwnd)
        if title and any(keyword in title for keyword in KAKAO_KEYWORDS):
            results.append((hwnd, title))
    return True


def _enum_all_visible_callback(hwnd, results):
    if win32gui.IsWindowVisible(hwnd):
        try:
            title = win32gui.GetWindowText(hwnd)
            rect = win32gui.GetWindowRect(hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            if title and width >= MIN_WINDOW_WIDTH and height >= MIN_WINDOW_HEIGHT:
                results.append((hwnd, title))
        except Exception:
            pass
    return True


def find_kakao_window(chat_room_name: str):
    room = (chat_room_name or "").strip()
    results = []
    win32gui.EnumWindows(_enum_kakao_callback, results)
    for hwnd, title in results:
        if room:
            if room in title:
                return hwnd, title
        else:
            return hwnd, title

    if room:
        all_visible = []
        win32gui.EnumWindows(_enum_all_visible_callback, all_visible)
        for hwnd, title in all_visible:
            if room in title:
                return hwnd, title
    return None, None


def get_window_rect(hwnd):
    try:
        left, top, right, bottom = win32gui.GetClientRect(hwnd)
        x, y = win32gui.ClientToScreen(hwnd, (left, top))
        width = right - left
        height = bottom - top
        return x, y, width, height
    except Exception:
        return None


def capture_chat_area(hwnd):
    rect = get_window_rect(hwnd)
    if not rect:
        return None
    wx, wy, width, height = rect
    x_margin = int(width * 0.03)
    chat_top = int(height * 0.18)
    chat_bottom = int(height * 0.88)
    chat_height = max(chat_bottom - chat_top, 1)
    return pyautogui.screenshot(
        region=(wx + x_margin, wy + chat_top, max(width - (x_margin * 2), 1), chat_height)
    )


def sanitize_ocr_line(line: str) -> str:
    cleaned = line.strip()
    cleaned = re.sub(r"^[^가-힣0-9]+", "", cleaned)
    cleaned = re.sub(r"[^가-힣0-9% ]+$", "", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip()


def normalize_region(region: str) -> str:
    for suffix in REGION_SUFFIXES:
        if region.endswith(suffix) and len(region) > len(suffix):
            return region[:-len(suffix)]
    return region


def build_region_variants(region: str) -> Set[str]:
    variants = {region, normalize_region(region)}
    base = normalize_region(region)
    for suffix in REGION_SUFFIXES:
        variants.add(base + suffix)
    return variants


def normalize_job(text: str) -> Optional[str]:
    compact_text = text.replace(" ", "")
    for standard, variants in JOB_NORMALIZE:
        for variant in variants:
            if variant in text or variant.replace(" ", "") in compact_text:
                return standard
    return None


def extract_commission(text: str) -> Optional[int]:
    """
    수수료는 반드시 '수수/수수료' 맥락이 있을 때만 인정한다.
    임의의 1~30 숫자(시간, 잡음)를 수수료로 쓰지 않는다.
    """
    compact_text = re.sub(r"\s+", "", text)
    patterns = [
        r"수수\s*료?\s*(\d{1,2})\s*%?",
        r"(\d{1,2})\s*%\s*수수",
        r"수수\s*(\d{1,2})",
        r"수수료\s*(\d{1,2})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            value = int(match.group(1))
            if 1 <= value <= 30:
                full = match.group(0)
                if value == 1 and "%" not in text and "료" not in full:
                    continue
                return value
    spaced = re.search(r"수\s+수\s*(?:료)?\s*(\d{1,2})\s*%?", text)
    if spaced:
        value = int(spaced.group(1))
        if 1 <= value <= 30:
            if not (value == 1 and "%" not in text and "료" not in spaced.group(0)):
                return value
    compact_patterns = [r"수수료?(\d{1,2})%?", r"(\d{1,2})%수수"]
    for pattern in compact_patterns:
        match = re.search(pattern, compact_text, re.IGNORECASE)
        if match:
            value = int(match.group(1))
            if 1 <= value <= 30:
                full = match.group(0)
                if value == 1 and "%" not in text and "료" not in full:
                    continue
                return value
    return None


def is_noise_line(line: str, response_text: str = "네") -> bool:
    line = sanitize_ocr_line(line)
    if not line or len(line) < 3:
        return True
    if response_text and response_text in line and len(line) <= len(response_text) + 2:
        return True
    if re.search(r"오[전후추]\s*\d{1,2}:?\d{2}", line):
        return True
    if re.search(r"\d{1,2}:\d{2}", line) and len(line) < 15:
        return True
    if any(keyword in line for keyword in ["공지", "시스템", "알림", "입장", "퇴장"]):
        return True
    hangul = len(re.findall(r"[가-힣]", line))
    if hangul < 2:
        return True
    return False


def fix_ocr_place_tokens(line: str) -> str:
    """자주 틀리는 OCR 패턴을 보정한다."""
    s = line
    s = re.sub(r"대\s*방\s*등(?=\s|$|[가-힣\s]*(?:싱|성|씽|하))", "대방동", s)
    s = re.sub(r"금\s*전(?=\s|$|[가-힣\s]*하)", "금천", s)
    s = re.sub(r"가\s*포(?=\s|$|[가-힣\s]*(?:싱|성|씽|하|쉬))", "마포", s)
    s = re.sub(r"과\s*악(?=\s|$|[가-힣\s]*하)", "관악", s)
    return s


def parse_message(line: str, selected_regions: List[str], response_text: str = "네") -> Optional[Tuple[str, str, int]]:
    line = fix_ocr_place_tokens(line)
    clean_line = sanitize_ocr_line(line)
    if is_noise_line(clean_line, response_text):
        return None

    compact_line = re.sub(r"\s+", "", clean_line)
    commission = extract_commission(clean_line)
    if commission is None:
        return None
    job = normalize_job(clean_line)
    if job is None:
        return None

    selected_normalized = {normalize_region(region): region for region in selected_regions}
    selected_variants = set()
    for region in selected_regions:
        selected_variants.update(build_region_variants(region))

    for variant in selected_variants:
        compact_variant = re.sub(r"\s+", "", variant)
        compact_variant_norm = normalize_region(compact_variant)
        if compact_variant and (compact_variant in compact_line or compact_variant_norm in compact_line):
            matched_region = selected_normalized.get(normalize_region(variant)) or variant
            return matched_region, job, commission

    region_candidates = re.findall(r"[가-힣]+(?:동|구|시|군|읍|면|리)?|[가-힣]{2,}", clean_line)
    job_keywords = {"싱막", "하막", "역류", "누수", "싱크대막힘", "하수구막힘", "성막", "생막", "씽막"}
    region_candidates = [
        candidate
        for candidate in region_candidates
        if candidate not in job_keywords and "수수" not in candidate and len(candidate) >= 2
    ]

    matched_region = None
    for candidate in region_candidates:
        candidate_norm = normalize_region(candidate)
        if candidate in selected_variants or candidate_norm in selected_normalized:
            matched_region = selected_normalized.get(candidate_norm) or candidate
            break
        for selected in selected_regions:
            selected_norm = normalize_region(selected)
            if len(candidate_norm) == len(selected_norm):
                diff = sum(1 for a, b in zip(candidate_norm, selected_norm) if a != b)
                if diff <= 1:
                    matched_region = selected
                    break
        if matched_region:
            break
    if matched_region is None:
        return None

    return matched_region, job, commission


def make_duplicate_key(region: str, job: str, commission: int) -> str:
    return f"{normalize_region(region)}|{job}|{commission}"


def is_strong_candidate(raw_line: str, region: str, commission: int) -> bool:
    compact_line = re.sub(r"\s+", "", sanitize_ocr_line(fix_ocr_place_tokens(raw_line)))
    region_base = normalize_region(region)
    commission_tokens = [
        f"수수료{commission}",
        f"수수{commission}",
        f"{commission}%수수",
        f"수수료{commission}%",
        f"수수{commission}%",
    ]
    has_region = region_base in compact_line or region in compact_line
    has_commission = any(token in compact_line for token in commission_tokens)
    return has_region and has_commission


class MacroApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("설비 수주 자동응답 매크로 v1.8")
        self._set_initial_window_position()
        self.root.resizable(True, True)

        self.running = False
        self.thread: Optional[threading.Thread] = None
        # {dup_key: {'count': int, 'line': str, 'parsed': tuple, 'first_seen': float, 'last_seen': float}}
        self.pending_candidates: Dict[str, dict] = {}
        self.processed_keys: Set[str] = set()
        self.processed_records: Dict[str, dict] = {}
        self.last_send_time = 0.0
        self.empty_ocr_streak = 0
        self.selected_region_cache: List[str] = []
        self.selected_region_signature = ""
        self.current_chat_time = "시간없음"

        self._build_ui()
        self._check_tesseract()

    def _set_initial_window_position(self):
        width = 800
        height = 900
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = max(int((screen_width - width) / 2) - 320, 8)
        y = max(int((screen_height - height) / 2) - 20, 20)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _check_tesseract(self):
        if init_tesseract():
            self._log("Tesseract OCR 준비 완료")
        else:
            self._log("[경고] Tesseract 경로를 찾을 수 없습니다. C:\\Program Files\\Tesseract-OCR 설치 필요", "warning")

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)
        cfg = ttk.LabelFrame(main, text="설정", padding=8)
        cfg.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(cfg, text="채팅방 이름:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.entry_room = ttk.Entry(cfg, width=40)
        self.entry_room.grid(row=0, column=1, padx=4, pady=2, sticky=tk.EW)

        ttk.Label(cfg, text="응답 문구:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.entry_response = ttk.Entry(cfg, width=40)
        self.entry_response.grid(row=1, column=1, padx=4, pady=2, sticky=tk.EW)
        self.entry_response.insert(0, "네")

        ttk.Label(cfg, text="스캔 주기(초):").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.entry_scan = ttk.Entry(cfg, width=10)
        self.entry_scan.grid(row=2, column=1, padx=4, pady=2, sticky=tk.W)
        self.entry_scan.insert(0, "0.3")

        ttk.Label(cfg, text="전송 간격(초):").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.entry_interval = ttk.Entry(cfg, width=10)
        self.entry_interval.grid(row=3, column=1, padx=4, pady=2, sticky=tk.W)
        self.entry_interval.insert(0, "2.0")
        cfg.columnconfigure(1, weight=1)

        region_frame = ttk.LabelFrame(main, text="지역 설정 (시/도 → 구/군/시 → 동)", padding=8)
        region_frame.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(region_frame, text="시/도:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.var_sido = tk.StringVar()
        self.combo_sido = ttk.Combobox(region_frame, textvariable=self.var_sido, width=18, state="readonly")
        self.combo_sido["values"] = ["(전체)"] + list(REGION_HIERARCHY.keys())
        self.combo_sido.grid(row=0, column=1, padx=4, pady=2)
        self.combo_sido.bind("<<ComboboxSelected>>", self._on_sido_select)

        ttk.Label(region_frame, text="구/군/시:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.var_gugun = tk.StringVar()
        self.combo_gugun = ttk.Combobox(region_frame, textvariable=self.var_gugun, width=18, state="readonly")
        self.combo_gugun.grid(row=1, column=1, padx=4, pady=2)
        self.combo_gugun.bind("<<ComboboxSelected>>", self._on_gugun_select)

        ttk.Label(region_frame, text="동:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.var_dong = tk.StringVar()
        self.combo_dong = ttk.Combobox(region_frame, textvariable=self.var_dong, width=18, state="readonly")
        self.combo_dong.grid(row=2, column=1, padx=4, pady=2)

        ttk.Button(region_frame, text="추가", command=self._add_region).grid(row=3, column=1, padx=4, pady=4, sticky=tk.W)
        ttk.Button(region_frame, text="삭제", command=self._del_region).grid(row=3, column=1, padx=4, pady=4, sticky=tk.E)
        ttk.Button(region_frame, text="전체삭제", command=self._clear_regions).grid(row=3, column=2, padx=4, pady=4)

        ttk.Label(region_frame, text="선택 지역:").grid(row=4, column=0, sticky=tk.NW, pady=2)
        self.list_regions = tk.Listbox(region_frame, height=4, width=48)
        self.list_regions.grid(row=4, column=1, columnspan=2, padx=4, pady=2, sticky=tk.EW)
        region_frame.columnconfigure(1, weight=1)

        cond_frame = ttk.LabelFrame(main, text="마지막 감지 조건", padding=8)
        cond_frame.pack(fill=tk.X, pady=(0, 8))
        self.label_cond = ttk.Label(cond_frame, text="지역: - | 작업: - | 수수료: -")
        self.label_cond.pack(anchor=tk.W)

        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=8)
        self.btn_start = ttk.Button(btn_frame, text="시작", command=self._on_start)
        self.btn_start.pack(side=tk.LEFT, padx=4)
        self.btn_stop = ttk.Button(btn_frame, text="중지", command=self._on_stop, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=4)
        self.btn_clear_log = ttk.Button(btn_frame, text="로그삭제", command=self._clear_log)
        self.btn_clear_log.pack(side=tk.LEFT, padx=4)

        log_frame = ttk.LabelFrame(main, text="로그", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=24, wrap=tk.WORD, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.tag_configure("default", foreground="#222222")
        self.log_text.tag_configure("match", foreground="#cc0000")
        self.log_text.tag_configure("warning", foreground="#a55b00")
        self.log_text.tag_configure("error", foreground="#b00020")
        self.log_text.tag_configure("debug", foreground="#1f4e79")

    def _format_regions_for_log(self) -> str:
        items = list(self.list_regions.get(0, tk.END))
        return ", ".join(items) if items else "없음"

    def _on_sido_select(self, event=None):
        sido = self.var_sido.get()
        self.var_gugun.set("")
        self.var_dong.set("")
        self.combo_dong["values"] = []
        if sido == "(전체)":
            self.combo_gugun["values"] = ["(전체)"]
            return
        if sido and sido in REGION_HIERARCHY:
            self.combo_gugun["values"] = ["(전체)"] + list(REGION_HIERARCHY[sido].keys())

    def _on_gugun_select(self, event=None):
        sido = self.var_sido.get()
        gugun = self.var_gugun.get()
        self.var_dong.set("")
        if gugun == "(전체)":
            self.combo_dong["values"] = ["(전체)"]
            return
        if sido and gugun and sido in REGION_HIERARCHY and gugun in REGION_HIERARCHY[sido]:
            self.combo_dong["values"] = ["(전체)"] + REGION_HIERARCHY[sido][gugun]

    def _expand_region_selection(self, sido: str, gugun: str, dong: str) -> List[str]:
        if sido == "(전체)":
            return ["(전체)"]
        if not sido:
            return []
        if gugun == "(전체)":
            return [f"{sido} (전체)"]
        if not gugun:
            return []
        if dong == "(전체)":
            return [f"{sido} {gugun} (전체)"]
        if dong:
            return [f"{sido} {gugun} {dong}"]
        return [f"{sido} {gugun}"]

    def _compute_selected_regions(self) -> List[str]:
        items = list(self.list_regions.get(0, tk.END))
        result: List[str] = []

        if "(전체)" in items:
            for sido_name, guguns in REGION_HIERARCHY.items():
                for gugun_name, dongs in guguns.items():
                    result.append(gugun_name)
                    result.extend(dongs)
            return list(dict.fromkeys(result))

        for item in items:
            parts = item.split()
            if len(parts) == 2 and parts[1] == "(전체)":
                sido = parts[0]
                for gugun_name, dongs in REGION_HIERARCHY.get(sido, {}).items():
                    result.append(gugun_name)
                    result.extend(dongs)
                continue
            if len(parts) == 3 and parts[2] == "(전체)":
                sido, gugun = parts[0], parts[1]
                result.append(gugun)
                result.extend(REGION_HIERARCHY.get(sido, {}).get(gugun, []))
                continue
            if len(parts) == 2:
                result.append(parts[1])
                result.extend(REGION_HIERARCHY.get(parts[0], {}).get(parts[1], []))
            elif len(parts) >= 3:
                result.append(parts[1])
                result.append(parts[2])
        return list(dict.fromkeys(result))

    def _refresh_region_cache(self):
        self.selected_region_cache = self._compute_selected_regions()
        self.selected_region_signature = "|".join(self.selected_region_cache)

    def _add_region(self):
        sido = self.var_sido.get().strip()
        gugun = self.var_gugun.get().strip()
        dong = self.var_dong.get().strip()
        if not sido:
            messagebox.showwarning("경고", "시/도를 선택하세요.")
            return
        if sido != "(전체)" and not gugun:
            messagebox.showwarning("경고", "구/군/시를 선택하세요.")
            return

        items = self.list_regions.get(0, tk.END)
        added_names = []
        for name in self._expand_region_selection(sido, gugun, dong):
            if name not in items:
                self.list_regions.insert(tk.END, name)
                added_names.append(name)
        if added_names:
            self._refresh_region_cache()
            self._log(f"[설정] 지역을 추가했습니다: {', '.join(added_names)}", "warning")
            self._log(f"[설정] 현재 감시 지역: {self._format_regions_for_log()}", "warning")

    def _del_region(self):
        selection = self.list_regions.curselection()
        if selection:
            removed = self.list_regions.get(selection[0])
            self.list_regions.delete(selection[0])
            self._refresh_region_cache()
            self._log(f"[설정] 지역을 삭제했습니다: {removed}", "warning")
            self._log(f"[설정] 현재 감시 지역: {self._format_regions_for_log()}", "warning")

    def _clear_regions(self):
        self.list_regions.delete(0, tk.END)
        self._refresh_region_cache()
        self._log("[설정] 감시 지역을 전체삭제했습니다.", "warning")

    def _clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.current_chat_time = self._system_time_label()
        self._log("로그를 비웠습니다.")

    def _extract_chat_time(self, positioned_lines) -> Optional[str]:
        if not positioned_lines:
            return None
        candidates = []
        for item in positioned_lines[-24:]:
            text = sanitize_ocr_line(item["text"])
            compact = re.sub(r"\s+", "", text).replace("오추", "오후").replace("오주", "오후")
            match = re.search(r"(오전|오후)(\d{1,2})(\d{2})", compact)
            if match:
                meridiem, hour, minute = match.groups()
                formatted = self._format_chat_time(meridiem, hour, minute)
                if formatted:
                    candidates.append(formatted)
                continue
            match = re.search(r"(오전|오후)\s*(\d{1,2}):(\d{2})", text)
            if match:
                meridiem, hour, minute = match.groups()
                formatted = self._format_chat_time(meridiem, hour, minute)
                if formatted:
                    candidates.append(formatted)
        if not candidates:
            return None
        return max(candidates, key=self._time_to_minutes)

    def _format_chat_time(self, meridiem: str, hour_text: str, minute_text: str) -> Optional[str]:
        try:
            hour = int(hour_text)
            minute = int(minute_text)
        except ValueError:
            return None
        if not (1 <= hour <= 12 and 0 <= minute <= 59):
            return None
        return f"{meridiem} {hour}:{minute:02d}"

    def _time_to_minutes(self, label: str) -> int:
        match = re.match(r"(오전|오후)\s*(\d{1,2}):(\d{2})", label or "")
        if not match:
            return -1
        meridiem, hour_text, minute_text = match.groups()
        hour_raw = int(hour_text)
        minute = int(minute_text)
        if not (1 <= hour_raw <= 12 and 0 <= minute <= 59):
            return -1
        hour = hour_raw % 12
        if meridiem == "오후":
            hour += 12
        return hour * 60 + minute

    def _system_time_label(self) -> str:
        now = time.localtime()
        hour = now.tm_hour
        minute = now.tm_min
        meridiem = "오전" if hour < 12 else "오후"
        display_hour = hour % 12
        if display_hour == 0:
            display_hour = 12
        return f"{meridiem} {display_hour}:{minute:02d}"

    def _refresh_log_time(self, detected_chat_time: Optional[str]):
        system_label = self._system_time_label()
        if not detected_chat_time:
            self.current_chat_time = system_label
            return
        chat_minutes = self._time_to_minutes(detected_chat_time)
        system_minutes = self._time_to_minutes(system_label)
        if chat_minutes >= 0 and system_minutes >= 0 and chat_minutes + CHAT_TIME_STALE_MINUTES < system_minutes:
            self.current_chat_time = system_label
            return
        self.current_chat_time = detected_chat_time

    def _log(self, msg: str, tag: str = "default"):
        def _do():
            prefix = self.current_chat_time if self.current_chat_time else "시간없음"
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, f"[{prefix}] {msg}\n", tag)
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)
        self.root.after(0, _do)

    def _update_cond(self, region: str, job: str, commission: int):
        def _do():
            self.label_cond.config(text=f"지역: {region} | 작업: {job} | 수수료: {commission}%")
        self.root.after(0, _do)

    def _extract_focus_lines(self, positioned_lines):
        if not positioned_lines:
            return []
        ordered = sorted(positioned_lines, key=lambda item: item["y"])
        tail_lines = ordered[-12:]
        max_y = tail_lines[-1]["y"]
        min_y = tail_lines[0]["y"]
        threshold = min_y + int((max_y - min_y) * 0.35)
        focus = [item for item in tail_lines if item["y"] >= threshold] or tail_lines
        return focus[-6:]

    def _extract_all_candidates(self, focus_lines, region_names: List[str], response_text: str) -> List[Tuple[float, str, str, tuple, bool]]:
        """focus_lines에서 파싱 가능한 모든 후보를 반환. (y, dup_key, display_line, parsed, strong)"""
        candidates = []
        seen_keys: Set[str] = set()
        trials: List[Tuple[float, str]] = [(item["y"], item["text"]) for item in focus_lines]
        ordered = sorted(focus_lines, key=lambda item: item["y"])
        for i in range(len(ordered) - 1):
            merged = f"{ordered[i]['text']} {ordered[i + 1]['text']}".strip()
            trials.append((ordered[i + 1]["y"], merged))
        for y, raw in trials:
            parsed = parse_message(raw, region_names, response_text)
            if not parsed:
                continue
            region, job, commission = parsed
            dup_key = make_duplicate_key(region, job, commission)
            if dup_key in seen_keys:
                continue
            seen_keys.add(dup_key)
            display_line = sanitize_ocr_line(fix_ocr_place_tokens(raw))
            strong = is_strong_candidate(raw, region, commission)
            candidates.append((y, dup_key, display_line, parsed, strong))
        candidates.sort(key=lambda x: x[0])
        return candidates

    def _reset_tracking(self):
        self.pending_candidates = {}
        self.processed_keys = set()
        self.processed_records = {}
        self.last_send_time = 0.0
        self.empty_ocr_streak = 0
        self.current_chat_time = "시간없음"

    def _on_start(self):
        self._refresh_region_cache()
        if not self.selected_region_cache:
            messagebox.showwarning("경고", "지역을 최소 1개 이상 추가하세요.")
            return
        try:
            scan = float(self.entry_scan.get())
            interval = float(self.entry_interval.get())
            if scan <= 0 or interval <= 0:
                raise ValueError("스캔 주기와 전송 간격은 0보다 커야 합니다.")
        except ValueError as exc:
            messagebox.showerror("오류", str(exc))
            return
        self.running = True
        self.btn_start.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        self._reset_tracking()
        self._log("=== 감시 시작 ===")
        self._log(f"[설정] 감시 지역: {self._format_regions_for_log()}", "warning")
        self.thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.thread.start()

    def _on_stop(self):
        self.running = False
        self.btn_start.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)
        self._log("=== 감시 중지 ===")

    def _send_response(self, text: str):
        try:
            pyperclip.copy(text)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(0.1)
            pyautogui.press("enter")
            return True
        except Exception as exc:
            self._log(f"[오류] 전송 실패: {exc}", "error")
            return False

    def _monitor_loop(self):
        room_name = self.entry_room.get().strip()
        response_text = self.entry_response.get().strip() or "네"
        scan_interval = float(self.entry_scan.get())
        send_interval = float(self.entry_interval.get())
        baseline_set = False
        last_region_signature = self.selected_region_signature

        while self.running:
            try:
                current_region_names = list(self.selected_region_cache)
                current_region_signature = self.selected_region_signature

                if current_region_signature != last_region_signature:
                    last_region_signature = current_region_signature
                    baseline_set = False
                    self.pending_candidates = {}
                    self._log(f"[설정] 감시 지역이 변경되었습니다: {self._format_regions_for_log()}", "warning")
                    self._log("[설정] 지역 변경 감지, 기준선을 다시 설정합니다.", "warning")

                hwnd, _ = find_kakao_window(room_name)
                if not hwnd:
                    self._log("[경고] 카카오톡 창을 찾을 수 없습니다.", "warning")
                    time.sleep(scan_interval)
                    continue

                img = capture_chat_area(hwnd)
                if img is None:
                    time.sleep(scan_interval)
                    continue

                positioned_lines = extract_positioned_lines_from_image(img)
                detected_chat_time = self._extract_chat_time(positioned_lines)
                self._refresh_log_time(detected_chat_time)

                if not positioned_lines:
                    self.empty_ocr_streak += 1
                    if self.empty_ocr_streak in (1, 10) or self.empty_ocr_streak % 30 == 0:
                        self._log("[대기] OCR 텍스트를 읽지 못했습니다. 카톡 창을 더 크게 띄우고 라이트 모드를 유지하세요.", "warning")
                    time.sleep(scan_interval)
                    continue

                self.empty_ocr_streak = 0
                focus_lines = self._extract_focus_lines(positioned_lines)

                # ── 기준선 설정 ──────────────────────────────────────────────
                if not baseline_set:
                    baseline_set = True
                    baseline_candidates = self._extract_all_candidates(
                        focus_lines, current_region_names, response_text
                    )
                    if baseline_candidates:
                        for _, dup_key, _, _, _ in baseline_candidates:
                            self.processed_keys.add(dup_key)
                        latest_line = baseline_candidates[-1][2]
                        n = len(baseline_candidates)
                        suffix = f"{n}개 메시지" if n > 1 else "마지막 메시지"
                        self._log(f"기준선 설정 완료 (기존 {suffix} 무시: {latest_line})")
                    else:
                        self._log("기준선 설정 완료 (현재 화면에 조건 일치 메시지 없음)")
                    time.sleep(scan_interval)
                    continue

                # ── 후보 탐색 ────────────────────────────────────────────────
                all_candidates = self._extract_all_candidates(
                    focus_lines, current_region_names, response_text
                )

                now = time.time()

                # 현재 화면에 보이는 미처리 후보 키셋
                current_keys: Set[str] = {
                    dup_key
                    for _, dup_key, _, _, _ in all_candidates
                    if dup_key not in self.processed_keys
                }

                # OCR이 한 번 튀어도 바로 버리지 않고 잠시 유지
                for stale in list(self.pending_candidates):
                    info = self.pending_candidates[stale]
                    if stale in current_keys:
                        continue
                    stale_limit = PENDING_TTL_SECONDS
                    if info["count"] >= info["required_count"]:
                        stale_limit = max(PENDING_TTL_SECONDS, send_interval + 1.0)
                    if now - info["last_seen"] > stale_limit:
                        del self.pending_candidates[stale]

                # 카운트 갱신 및 신규 후보 등록
                for _, dup_key, line, parsed, strong in all_candidates:
                    if dup_key in self.processed_keys:
                        continue
                    required_count = FAST_CONFIRM_THRESHOLD if strong else CONFIRM_THRESHOLD
                    if dup_key not in self.pending_candidates:
                        self.pending_candidates[dup_key] = {
                            "count": 1,
                            "line": line,
                            "parsed": parsed,
                            "required_count": required_count,
                            "first_seen": now,
                            "last_seen": now,
                        }
                        if required_count > 1:
                            self._log(f"[OCR후보] {line} (확정 대기 1/{required_count})", "debug")
                    else:
                        self.pending_candidates[dup_key]["count"] += 1
                        self.pending_candidates[dup_key]["line"] = line
                        self.pending_candidates[dup_key]["parsed"] = parsed
                        self.pending_candidates[dup_key]["required_count"] = min(
                            self.pending_candidates[dup_key]["required_count"],
                            required_count,
                        )
                        self.pending_candidates[dup_key]["last_seen"] = now

                # 확정 임계값 도달 후보 (가장 먼저 등장한 것 우선 처리)
                ready = [
                    (info["first_seen"], dup_key, info["line"], info["parsed"])
                    for dup_key, info in self.pending_candidates.items()
                    if info["count"] >= info["required_count"]
                ]

                if not ready:
                    time.sleep(scan_interval)
                    continue

                ready.sort(key=lambda x: x[0])
                _, dup_key, candidate_line, parsed = ready[0]
                region, job, commission = parsed

                # 중복 방지 (동일 조건 이미 처리)
                if dup_key in self.processed_keys:
                    del self.pending_candidates[dup_key]
                    time.sleep(scan_interval)
                    continue

                # 전송 간격 대기
                send_now = time.time()
                if send_now - self.last_send_time < send_interval:
                    wait = send_interval - (send_now - self.last_send_time)
                    self._log(f"[대기] 전송 간격 {wait:.1f}초 대기")
                    time.sleep(wait)

                self._log(f"[조건충족] {candidate_line} → 응답 전송", "match")
                self._update_cond(region, job, commission)

                try:
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.2)
                except Exception:
                    pass

                if self._send_response(response_text):
                    self.processed_keys.add(dup_key)
                    self.processed_records[dup_key] = {
                        "line": candidate_line,
                        "chat_time": self.current_chat_time,
                    }
                    self.last_send_time = time.time()
                    self._log(f"[전송완료] '{response_text}'")
                    self.pending_candidates.pop(dup_key, None)
                else:
                    self._log("[전송실패]", "error")

            except Exception as exc:
                self._log(f"[오류] {exc}", "error")

            time.sleep(scan_interval)


def main():
    app = MacroApp()
    app.root.mainloop()


if __name__ == "__main__":
    main()
