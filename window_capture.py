# -*- coding: utf-8 -*-
"""
카카오톡 PC 창 탐색 및 채팅 영역 캡처 모듈
"""

import pyautogui
import win32gui

# 창 제목에 포함될 키워드 (카카오톡)
KAKAO_KEYWORDS = ["카카오톡", "KakaoTalk", "오픈채팅", "KAKAO"]
# 최소 창 크기 (채팅창으로 인정할 최소 크기)
MIN_WINDOW_WIDTH = 250
MIN_WINDOW_HEIGHT = 300


def _enum_kakao_callback(hwnd, results):
    """카카오톡 키워드 포함 창 열거"""
    if win32gui.IsWindowVisible(hwnd):
        title = win32gui.GetWindowText(hwnd)
        if title and any(kw in title for kw in KAKAO_KEYWORDS):
            results.append((hwnd, title))
    return True


def _enum_all_visible_callback(hwnd, results):
    """채팅방 이름 포함 창 열거 (fallback)"""
    if win32gui.IsWindowVisible(hwnd):
        try:
            title = win32gui.GetWindowText(hwnd)
            rect = win32gui.GetWindowRect(hwnd)
            w = rect[2] - rect[0]
            h = rect[3] - rect[1]
            if title and w >= MIN_WINDOW_WIDTH and h >= MIN_WINDOW_HEIGHT:
                results.append((hwnd, title))
        except Exception:
            pass
    return True


def find_kakao_window(chat_room_name: str):
    """
    지정된 채팅방 이름과 일치하는 카카오톡 창 핸들 반환
    chat_room_name: 감시할 오픈채팅방 제목
    """
    room = (chat_room_name or "").strip()

    # 1) 카카오톡 키워드 포함 창에서 채팅방 이름 매칭
    results = []
    win32gui.EnumWindows(_enum_kakao_callback, results)
    for hwnd, title in results:
        if room:
            if room in title:
                return hwnd, title
        else:
            return hwnd, title

    # 2) Fallback: 채팅방 이름이 제목에 포함된 창 (카톡이 제목만 표시하는 경우)
    if room:
        all_visible = []
        win32gui.EnumWindows(_enum_all_visible_callback, all_visible)
        for hwnd, title in all_visible:
            if room in title:
                return hwnd, title

    return None, None


def get_window_rect(hwnd):
    """창의 클라이언트 영역 화면 좌표 반환 (x, y, w, h)"""
    try:
        left, top, right, bottom = win32gui.GetClientRect(hwnd)
        x, y = win32gui.ClientToScreen(hwnd, (left, top))
        w = right - left
        h = bottom - top
        return (x, y, w, h)
    except Exception:
        return None


def capture_screen_region(x, y, w, h):
    """화면 지정 영역 캡처 (창이 화면에 보일 때 사용)"""
    return pyautogui.screenshot(region=(x, y, w, h))


def capture_chat_area(hwnd):
    """
    카카오톡 채팅 내용 영역 캡처
    상단 검색/공지 영역과 하단 입력 영역을 제외한 중앙 채팅 영역만 캡처
    """
    cr = get_window_rect(hwnd)
    if not cr:
        return None
    wx, wy, w, h = cr

    x_margin = int(w * 0.03)
    chat_top = int(h * 0.18)
    chat_bottom = int(h * 0.88)
    chat_height = max(chat_bottom - chat_top, 1)

    x = wx + x_margin
    y = wy + chat_top
    return capture_screen_region(x, y, max(w - (x_margin * 2), 1), chat_height)
