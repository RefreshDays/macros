# -*- coding: utf-8 -*-
"""
메시지 파싱 및 조건 매칭 모듈 v1.7
- v1.6 기반
- OCR 미리보기에서 시간/입력창 단독 노이즈 더 강하게 제외
"""

from message_parser_v1_6 import *  # noqa: F403


def is_meaningful_ocr_line(line: str) -> bool:
    cleaned = sanitize_ocr_line(line)  # noqa: F405
    if len(cleaned) < 3:
        return False

    # 시간/입력창 단독 줄은 로그 미리보기에서 제외
    if cleaned in {"메시지 입력", "오전", "오후"}:
        return False

    time_only = cleaned.replace(" ", "")
    if time_only.startswith(("오전", "오후", "오추")) and any(ch.isdigit() for ch in time_only):
        if "수수" not in cleaned and "싱막" not in cleaned and "하막" not in cleaned:
            return False

    hangul_count = sum(1 for ch in cleaned if "가" <= ch <= "힣")
    strong_keywords = ["수수", "싱막", "하막", "역류", "누수", "동", "구", "시"]
    if any(keyword in cleaned for keyword in strong_keywords):
        return True

    return hangul_count >= 4 and "메시지" not in cleaned
