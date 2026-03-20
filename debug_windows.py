# -*- coding: utf-8 -*-
"""창 제목 디버깅 - 모든 보이는 창 목록 출력"""
import win32gui

def callback(hwnd, results):
    if win32gui.IsWindowVisible(hwnd):
        title = win32gui.GetWindowText(hwnd)
        if title:
            try:
                rect = win32gui.GetWindowRect(hwnd)
                w, h = rect[2] - rect[0], rect[3] - rect[1]
                results.append((title, w, h))
            except:
                results.append((title, 0, 0))
    return True

results = []
win32gui.EnumWindows(callback, results)
for t, w, h in sorted(results, key=lambda x: -x[1]*x[2]):
    if w > 100 and h > 100:
        print(f"[{w}x{h}] {t}")
