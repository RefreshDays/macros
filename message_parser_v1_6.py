# -*- coding: utf-8 -*-
"""
메시지 파싱 및 조건 매칭 모듈 v1.6
- OCR 노이즈 정리
- 지역/작업명/수수료 추출
- 미리보기용 의미 있는 OCR 줄 판별
"""

import re
from typing import List, Optional, Set, Tuple

JOB_NORMALIZE = [
    ("싱크대막힘", ["싱크대막힘", "싱크대 막힘"]),
    ("하수구막힘", ["하수구막힘", "하수구 막힘"]),
    ("싱막", ["싱막", "성막", "생막", "씽막"]),
    ("하막", ["하막", "하 막"]),
    ("역류", ["역류"]),
    ("누수", ["누수"]),
]

REGION_SUFFIXES = ["동", "구", "시", "군", "읍", "면", "리"]


def sanitize_ocr_line(line: str) -> str:
    """OCR 줄에서 앞뒤 잡음을 제거하고 공백을 정리한다."""
    cleaned = line.strip()
    cleaned = re.sub(r"^[^가-힣0-9]+", "", cleaned)
    cleaned = re.sub(r"[^가-힣0-9% ]+$", "", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip()


def normalize_region(region: str) -> str:
    """지역명에서 접미사 제거 (송파구 → 송파, 거여동 → 거여)"""
    for suffix in REGION_SUFFIXES:
        if region.endswith(suffix) and len(region) > len(suffix):
            return region[:-len(suffix)]
    return region


def build_region_variants(region: str) -> Set[str]:
    """지역명의 다양한 표현 생성 (원형, 축약형, 접미사 변형)"""
    variants = {region, normalize_region(region)}
    base = normalize_region(region)
    for suffix in REGION_SUFFIXES:
        variants.add(base + suffix)
    return variants


def normalize_job(text: str) -> Optional[str]:
    """작업명 정규화 - 6개 핵심 작업 키워드 매칭"""
    compact_text = text.replace(" ", "")
    for standard, variants in JOB_NORMALIZE:
        for variant in variants:
            if variant in text or variant.replace(" ", "") in compact_text:
                return standard
    return None


def extract_commission(text: str) -> Optional[int]:
    """수수료 추출 (1~30만 유효)"""
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
                return value

    compact_patterns = [
        r"수수료?(\d{1,2})%?",
        r"(\d{1,2})%수수",
    ]
    for pattern in compact_patterns:
        match = re.search(pattern, compact_text, re.IGNORECASE)
        if match:
            value = int(match.group(1))
            if 1 <= value <= 30:
                return value

    numbers = re.findall(r"\b(\d{1,2})\b", text)
    for number in numbers:
        value = int(number)
        if 1 <= value <= 30:
            return value
    return None


def is_noise_line(line: str, response_text: str = "네") -> bool:
    """응답문구, 시간, 공지, 시스템 메시지 등 제외"""
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


def is_meaningful_ocr_line(line: str) -> bool:
    """로그 미리보기로 볼 가치가 있는 줄만 통과시킨다."""
    cleaned = sanitize_ocr_line(line)
    if len(cleaned) < 3:
        return False

    hangul_count = len(re.findall(r"[가-힣]", cleaned))
    if hangul_count < 2 and not re.search(r"\d{3,4}", cleaned):
        return False

    keywords = ["수수", "싱막", "하막", "역류", "누수", "동", "구", "시", "오전", "오후", "메시지"]
    return any(keyword in cleaned for keyword in keywords) or hangul_count >= 4


def parse_message(
    line: str,
    selected_regions: List[str],
    response_text: str = "네",
) -> Optional[Tuple[str, str, int]]:
    """
    메시지에서 (지역, 작업명, 수수료) 추출
    조건 충족 시 (region, job, commission) 반환, 아니면 None
    """
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
        if compact_variant and (
            compact_variant in compact_line
            or compact_variant_norm in compact_line
        ):
            matched_region = selected_normalized.get(normalize_region(variant)) or variant
            return matched_region, job, commission

    region_candidates = re.findall(
        r"[가-힣]+(?:동|구|시|군|읍|면|리)?|[가-힣]{2,}",
        clean_line,
    )
    job_keywords = {"싱막", "하막", "역류", "누수", "싱크대막힘", "하수구막힘", "성막", "생막", "씽막"}
    region_candidates = [candidate for candidate in region_candidates if candidate not in job_keywords and "수수" not in candidate]

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
    """중복 판단 키 생성"""
    return f"{normalize_region(region)}|{job}|{commission}"
