# -*- coding: utf-8 -*-
"""
OCR 엔진 - Tesseract 기반 텍스트 추출
"""

import re
from typing import Dict, List

import pytesseract
from PIL import Image, ImageEnhance, ImageOps
from pytesseract import Output

# Tesseract 기본 경로 (Windows)
TESSERACT_PATHS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
]


def init_tesseract():
    """Tesseract 경로 설정"""
    import os
    for path in TESSERACT_PATHS:
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
            return True
    return False


def preprocess_image(img: Image.Image) -> Image.Image:
    """채팅 OCR 인식률을 높이기 위한 전처리."""
    gray = img.convert("L")
    gray = ImageEnhance.Contrast(gray).enhance(2.5)

    width, height = gray.size
    resized = gray.resize((max(width * 2, 1), max(height * 2, 1)), Image.Resampling.LANCZOS)

    # 밝은 배경의 채팅 글자를 선명하게 분리한다.
    binary = resized.point(lambda px: 255 if px > 185 else 0, mode="1").convert("L")
    return ImageOps.autocontrast(binary)


def _normalize_ocr_text(text: str) -> str:
    """OCR 결과의 자주 나오는 줄바꿈/공백 노이즈를 줄인다."""
    text = text.replace("|", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def extract_text_from_image(img: Image.Image, lang: str = "kor+eng") -> str:
    """
    이미지에서 텍스트 추출
    kor: 한글, eng: 영문 (OCR 오인식 보정용)
    """
    if img is None:
        return ""
    try:
        processed = preprocess_image(img)
        configs = [
            r"--oem 3 --psm 6",
            r"--oem 3 --psm 11",
        ]
        results = []
        for config in configs:
            text = pytesseract.image_to_string(processed, lang=lang, config=config)
            normalized = _normalize_ocr_text(text or "")
            if normalized:
                results.append(normalized)

        if not results:
            return ""

        # 더 긴 결과를 우선 사용한다.
        return max(results, key=len)
    except Exception as e:
        raise RuntimeError(f"OCR 실패: {e}. Tesseract 및 kor.traineddata 설치 확인.")


def extract_positioned_lines_from_image(img: Image.Image, lang: str = "kor+eng") -> List[Dict[str, int | str]]:
    """이미지에서 줄 단위 텍스트와 세로 위치를 함께 추출."""
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

        key = (
            data["block_num"][idx],
            data["par_num"][idx],
            data["line_num"][idx],
        )
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


def extract_lines_from_image(img: Image.Image) -> List[str]:
    """이미지에서 텍스트를 줄 단위로 추출"""
    positioned = extract_positioned_lines_from_image(img)
    if positioned:
        return [item["text"] for item in positioned]

    text = extract_text_from_image(img)
    return [ln.strip() for ln in text.splitlines() if ln.strip()]
