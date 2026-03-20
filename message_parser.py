# -*- coding: utf-8 -*-
"""
메시지 파싱 및 조건 매칭 모듈
- 지역/작업명/수수료 추출
- 정규화 및 매칭
"""

import re
from typing import Optional, Tuple, List, Set

# 작업명 정규화 매핑 (OCR 오인식 및 변형 → 표준형)
# 긴 키워드 먼저 매칭 (싱크대막힘, 하수구막힘 → 싱막, 하막 등)
JOB_NORMALIZE = [
    ("싱크대막힘", ["싱크대막힘", "싱크대 막힘"]),
    ("하수구막힘", ["하수구막힘", "하수구 막힘"]),
    ("싱막", ["싱막", "성막", "생막", "씽막"]),
    ("하막", ["하막", "하 막"]),
    ("역류", ["역류"]),
    ("누수", ["누수"]),
]

# 지역 접미사 제거용
REGION_SUFFIXES = ["동", "구", "시", "군", "읍", "면", "리"]


def normalize_region(region: str) -> str:
    """지역명에서 접미사 제거 (송파구 → 송파, 거여동 → 거여)"""
    for suffix in REGION_SUFFIXES:
        if region.endswith(suffix) and len(region) > len(suffix):
            return region[:-len(suffix)]
    return region


def normalize_region_reverse(base: str, suffix: str = "동") -> str:
    """축약형에 접미사 붙여 원형 생성 (거여 → 거여동)"""
    if not base:
        return base
    return base + suffix


def build_region_variants(region: str) -> Set[str]:
    """지역명의 다양한 표현 생성 (원형, 축약형, 접미사 변형)"""
    variants = {region, normalize_region(region)}
    base = normalize_region(region)
    for s in REGION_SUFFIXES:
        variants.add(base + s)
    return variants


def normalize_job(text: str) -> Optional[str]:
    """작업명 정규화 - 6개 핵심 작업 키워드 매칭"""
    text_no_space = text.replace(" ", "")
    for standard, variants in JOB_NORMALIZE:
        for v in variants:
            if v in text or v.replace(" ", "") in text_no_space:
                return standard
    return None


def extract_commission(text: str) -> Optional[int]:
    """
    수수료 추출 (1~30만 유효)
    수수10, 수수료10, 수수 10, 수수료 10%, 수수 20% 등
    """
    compact_text = re.sub(r"\s+", "", text)

    # 수수/수수료 패턴
    patterns = [
        r"수수\s*료?\s*(\d{1,2})\s*%?",  # 수수10, 수수료 10%, 수수 20%
        r"(\d{1,2})\s*%\s*수수",  # 10% 수수
        r"수수\s*(\d{1,2})",
        r"수수료\s*(\d{1,2})",
    ]
    for pattern in patterns:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            val = int(m.group(1))
            if 1 <= val <= 30:
                return val

    compact_patterns = [
        r"수수료?(\d{1,2})%?",
        r"(\d{1,2})%수수",
    ]
    for pattern in compact_patterns:
        m = re.search(pattern, compact_text, re.IGNORECASE)
        if m:
            val = int(m.group(1))
            if 1 <= val <= 30:
                return val

    # 숫자만 있는 경우 (1~30)
    numbers = re.findall(r"\b(\d{1,2})\b", text)
    for n in numbers:
        val = int(n)
        if 1 <= val <= 30:
            return val
    return None


def is_noise_line(line: str, response_text: str = "네") -> bool:
    """노이즈 제거: 응답문구, 시간, 공지, 시스템 메시지 등"""
    line = line.strip()
    if not line or len(line) < 3:
        return True
    # 응답 문구 자체
    if response_text and response_text in line:
        return True
    # 시간 패턴 (오전 11:42, 오후 3:00 등)
    if re.search(r"오[전후]\s*\d{1,2}:\d{2}", line):
        return True
    if re.search(r"\d{1,2}:\d{2}", line) and len(line) < 15:
        return True
    # 공지/시스템
    if any(kw in line for kw in ["공지", "시스템", "알림", "입장", "퇴장"]):
        return True
    # 한글 비중 체크 (너무 적으면 노이즈)
    hangul = len(re.findall(r"[가-힣]", line))
    if hangul < 2:
        return True
    return False


def parse_message(
    line: str,
    selected_regions: List[str],
    response_text: str = "네",
) -> Optional[Tuple[str, str, int]]:
    """
    메시지에서 (지역, 작업명, 수수료) 추출
    조건 충족 시 (region, job, commission) 반환, 아니면 None
    """
    if is_noise_line(line, response_text):
        return None

    compact_line = re.sub(r"\s+", "", line)

    # 수수료 먼저 추출
    commission = extract_commission(line)
    if commission is None:
        return None

    # 작업명 추출
    job = normalize_job(line)
    if job is None:
        return None

    # 지역 매칭
    selected_normalized = {normalize_region(r): r for r in selected_regions}
    selected_variants = set()
    for r in selected_regions:
        selected_variants.update(build_region_variants(r))

    for variant in selected_variants:
        variant_compact = re.sub(r"\s+", "", variant)
        variant_norm = normalize_region(variant_compact)
        if variant_compact and (
            variant_compact in compact_line
            or variant_norm in compact_line
        ):
            matched_region = selected_normalized.get(normalize_region(variant)) or variant
            return (matched_region, job, commission)

    # 라인에서 지역 후보 추출 (한글, 접미사 포함)
    region_candidates = re.findall(
        r"[가-힣]+(?:동|구|시|군|읍|면|리)?|[가-힣]{2,}",
        line,
    )
    # 작업 키워드 제외
    job_keywords = {"싱막", "하막", "역류", "누수", "싱크대막힘", "하수구막힘", "성막", "생막", "씽막"}
    region_candidates = [c for c in region_candidates if c not in job_keywords and "수수" not in c]

    matched_region = None
    for cand in region_candidates:
        cand_norm = normalize_region(cand)
        cand_variants = build_region_variants(cand)
        # 정확 매칭
        if cand in selected_variants or cand_norm in selected_normalized:
            matched_region = selected_normalized.get(cand_norm) or cand
            break
        # 1글자 오차 허용 (OCR 보정)
        for sel in selected_regions:
            sel_norm = normalize_region(sel)
            if len(cand_norm) == len(sel_norm):
                diff = sum(1 for a, b in zip(cand_norm, sel_norm) if a != b)
                if diff <= 1:
                    matched_region = sel
                    break
        if matched_region:
            break

    if matched_region is None:
        return None

    return (matched_region, job, commission)


def make_duplicate_key(region: str, job: str, commission: int) -> str:
    """중복 판단 키 생성"""
    return f"{normalize_region(region)}|{job}|{commission}"
