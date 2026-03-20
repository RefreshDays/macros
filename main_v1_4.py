# -*- coding: utf-8 -*-
"""
설비 수주 자동응답 매크로 - 메인 애플리케이션
카카오톡 PC 오픈채팅 OCR 자동응답
"""

import threading
import time
from typing import List, Optional, Set

import pyautogui
import pyperclip
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

from region_data import REGION_HIERARCHY
from window_capture import find_kakao_window, capture_chat_area
import win32gui
from ocr_engine import extract_positioned_lines_from_image, init_tesseract
from message_parser import parse_message, make_duplicate_key

# pyautogui 안전 설정
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.1


class MacroApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("설비 수주 자동응답 매크로 v1.4")
        self._set_initial_window_position()
        self.root.resizable(True, True)

        self.running = False
        self.thread: Optional[threading.Thread] = None
        self.last_seen_candidate: Optional[str] = None
        self.pending_candidate: Optional[str] = None
        self.pending_candidate_count = 0
        self.last_duplicate_logged: Optional[str] = None
        self.last_candidate_logged: Optional[str] = None
        self.last_preview_logged: Optional[str] = None
        self.processed_keys: Set[str] = set()
        self.last_send_time = 0.0
        self.empty_ocr_streak = 0

        self._build_ui()
        self._check_tesseract()

    def _set_initial_window_position(self):
        """화면 중앙보다 약간 왼쪽 위치에 앱을 띄운다."""
        width = 760
        height = 860
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = max(int((screen_width - width) / 2) - 140, 20)
        y = max(int((screen_height - height) / 2) - 20, 20)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _check_tesseract(self):
        if init_tesseract():
            self._log("Tesseract OCR 준비 완료")
        else:
            self._log("[경고] Tesseract 경로를 찾을 수 없습니다. C:\\Program Files\\Tesseract-OCR 설치 필요")

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        # === 설정 영역 ===
        cfg = ttk.LabelFrame(main, text="설정", padding=8)
        cfg.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(cfg, text="채팅방 이름:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.entry_room = ttk.Entry(cfg, width=35)
        self.entry_room.grid(row=0, column=1, padx=4, pady=2, sticky=tk.EW)
        self.entry_room.insert(0, "")

        ttk.Label(cfg, text="응답 문구:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.entry_response = ttk.Entry(cfg, width=35)
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

        # === 지역 선택 ===
        region_frame = ttk.LabelFrame(main, text="지역 설정 (시/도 → 구/군/시 → 동)", padding=8)
        region_frame.pack(fill=tk.X, pady=(0, 8))

        row_r = 0
        ttk.Label(region_frame, text="시/도:").grid(row=row_r, column=0, sticky=tk.W, pady=2)
        self.var_sido = tk.StringVar()
        self.combo_sido = ttk.Combobox(region_frame, textvariable=self.var_sido, width=18, state="readonly")
        self.combo_sido["values"] = ["(전체)"] + list(REGION_HIERARCHY.keys())
        self.combo_sido.grid(row=row_r, column=1, padx=4, pady=2)
        self.combo_sido.bind("<<ComboboxSelected>>", self._on_sido_select)

        row_r += 1
        ttk.Label(region_frame, text="구/군/시:").grid(row=row_r, column=0, sticky=tk.W, pady=2)
        self.var_gugun = tk.StringVar()
        self.combo_gugun = ttk.Combobox(region_frame, textvariable=self.var_gugun, width=18, state="readonly")
        self.combo_gugun.grid(row=row_r, column=1, padx=4, pady=2)
        self.combo_gugun.bind("<<ComboboxSelected>>", self._on_gugun_select)

        row_r += 1
        ttk.Label(region_frame, text="동:").grid(row=row_r, column=0, sticky=tk.W, pady=2)
        self.var_dong = tk.StringVar()
        self.combo_dong = ttk.Combobox(region_frame, textvariable=self.var_dong, width=18, state="readonly")
        self.combo_dong.grid(row=row_r, column=1, padx=4, pady=2)

        row_r += 1
        btn_add = ttk.Button(region_frame, text="추가", command=self._add_region)
        btn_add.grid(row=row_r, column=1, padx=4, pady=4, sticky=tk.W)
        btn_del = ttk.Button(region_frame, text="삭제", command=self._del_region)
        btn_del.grid(row=row_r, column=1, padx=4, pady=4, sticky=tk.E)
        btn_clear = ttk.Button(region_frame, text="전체삭제", command=self._clear_regions)
        btn_clear.grid(row=row_r, column=2, padx=4, pady=4)

        row_r += 1
        ttk.Label(region_frame, text="선택 지역:").grid(row=row_r, column=0, sticky=tk.NW, pady=2)
        self.list_regions = tk.Listbox(region_frame, height=4, width=40)
        self.list_regions.grid(row=row_r, column=1, columnspan=2, padx=4, pady=2, sticky=tk.EW)
        region_frame.columnconfigure(1, weight=1)

        # === 마지막 감지 조건 ===
        cond_frame = ttk.LabelFrame(main, text="마지막 감지 조건", padding=8)
        cond_frame.pack(fill=tk.X, pady=(0, 8))
        self.label_cond = ttk.Label(cond_frame, text="지역: - | 작업: - | 수수료: -")
        self.label_cond.pack(anchor=tk.W)

        # === 시작/중지 ===
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=8)
        self.btn_start = ttk.Button(btn_frame, text="시작", command=self._on_start)
        self.btn_start.pack(side=tk.LEFT, padx=4)
        self.btn_stop = ttk.Button(btn_frame, text="중지", command=self._on_stop, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=4)

        # === 로그 ===
        log_frame = ttk.LabelFrame(main, text="로그", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=22, wrap=tk.WORD, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.tag_configure("default", foreground="#222222")
        self.log_text.tag_configure("match", foreground="#cc0000")
        self.log_text.tag_configure("warning", foreground="#a55b00")
        self.log_text.tag_configure("error", foreground="#b00020")
        self.log_text.tag_configure("debug", foreground="#1f4e79")

    def _on_sido_select(self, event=None):
        sido = self.var_sido.get()
        self.var_gugun.set("")
        self.var_dong.set("")
        self.combo_dong["values"] = []

        if sido == "(전체)":
            self.combo_gugun["values"] = ["(전체)"]
            return

        if sido and sido in REGION_HIERARCHY:
            guguns = list(REGION_HIERARCHY[sido].keys())
            self.combo_gugun["values"] = ["(전체)"] + guguns

    def _on_gugun_select(self, event=None):
        sido = self.var_sido.get()
        gugun = self.var_gugun.get()
        self.var_dong.set("")

        if gugun == "(전체)":
            self.combo_dong["values"] = ["(전체)"]
            return

        if sido and gugun and sido in REGION_HIERARCHY and gugun in REGION_HIERARCHY[sido]:
            dongs = REGION_HIERARCHY[sido][gugun]
            self.combo_dong["values"] = ["(전체)"] + dongs

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

    def _add_region(self):
        dong = self.var_dong.get().strip()
        gugun = self.var_gugun.get().strip()
        sido = self.var_sido.get().strip()
        if not sido:
            messagebox.showwarning("경고", "시/도를 선택하세요.")
            return
        if sido != "(전체)" and not gugun:
            messagebox.showwarning("경고", "구/군/시를 선택하세요.")
            return

        items = self.list_regions.get(0, tk.END)
        for name in self._expand_region_selection(sido, gugun, dong):
            if name not in items:
                self.list_regions.insert(tk.END, name)

    def _del_region(self):
        sel = self.list_regions.curselection()
        if sel:
            self.list_regions.delete(sel[0])

    def _clear_regions(self):
        self.list_regions.delete(0, tk.END)

    def _get_selected_regions_simple(self) -> List[str]:
        """선택된 지역을 실제 매칭 가능한 구/동 목록으로 확장."""
        items = list(self.list_regions.get(0, tk.END))

        if "(전체)" in items:
            all_regions = []
            for sido_name, guguns in REGION_HIERARCHY.items():
                for gugun_name, dongs in guguns.items():
                    all_regions.append(gugun_name)
                    all_regions.extend(dongs)
            return list(dict.fromkeys(all_regions))

        result: List[str] = []
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
                continue

            if len(parts) >= 3:
                result.append(parts[1])
                result.append(parts[2])

        return list(dict.fromkeys(result))

    def _log(self, msg: str, tag: str = "default"):
        def _do():
            timestamp = time.strftime("%M:%S")
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, f"[{timestamp}] {msg}\n", tag)
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)

        self.root.after(0, _do)

    def _update_cond(self, region: str, job: str, commission: int):
        def _do():
            self.label_cond.config(text=f"지역: {region} | 작업: {job} | 수수료: {commission}%")

        self.root.after(0, _do)

    def _extract_latest_candidate(self, positioned_lines, region_names: List[str], response_text: str):
        """
        OCR 결과에서 화면상 가장 아래쪽 유효 메시지를 찾는다.
        반환값: (signature, raw_line, parsed_tuple) 또는 None
        """
        candidates = []
        for item in positioned_lines:
            line = item["text"].strip()
            parsed = parse_message(line, region_names, response_text)
            if not parsed:
                continue
            region, job, commission = parsed
            signature = f"{make_duplicate_key(region, job, commission)}|{line.strip()}"
            candidates.append((item["y"], signature, line, parsed))

        if not candidates:
            return None

        # 실제 최신 메시지는 가장 아래쪽에 위치한 경우가 많으므로 y가 가장 큰 후보를 사용한다.
        _, signature, line, parsed = max(candidates, key=lambda item: item[0])
        return signature, line, parsed

    def _log_ocr_preview(self, positioned_lines):
        """최근 OCR 결과 일부를 로그에 남겨 인식 과정을 확인한다."""
        if not positioned_lines:
            return
        preview = " | ".join(item["text"] for item in positioned_lines[-3:])
        if preview and preview != self.last_preview_logged:
            self.last_preview_logged = preview
            self._log(f"[OCR] 최근 인식: {preview}", "debug")

    def _on_start(self):
        regions = self._get_selected_regions_simple()
        if not regions:
            messagebox.showwarning("경고", "지역을 최소 1개 이상 추가하세요.")
            return

        try:
            scan = float(self.entry_scan.get())
            interval = float(self.entry_interval.get())
            if scan <= 0 or interval <= 0:
                raise ValueError("스캔 주기와 전송 간격은 0보다 커야 합니다.")
        except ValueError as e:
            messagebox.showerror("오류", str(e))
            return

        self.running = True
        self.btn_start.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        self.last_seen_candidate = None
        self.pending_candidate = None
        self.pending_candidate_count = 0
        self.last_duplicate_logged = None
        self.last_candidate_logged = None
        self.last_preview_logged = None
        self.processed_keys = set()
        self.last_send_time = 0.0
        self.empty_ocr_streak = 0

        self._log("=== 감시 시작 ===")
        self.thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.thread.start()

    def _on_stop(self):
        self.running = False
        self.btn_start.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)
        self._log("=== 감시 중지 ===")

    def _send_response(self, text: str):
        """응답 문구 전송 (클립보드 붙여넣기 후 엔터 - 한글 지원)"""
        try:
            pyperclip.copy(text)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(0.1)
            pyautogui.press("enter")
            return True
        except Exception as e:
            self._log(f"[오류] 전송 실패: {e}")
            return False

    def _monitor_loop(self):
        room_name = self.entry_room.get().strip()
        response_text = self.entry_response.get().strip() or "네"
        scan_interval = float(self.entry_scan.get())
        send_interval = float(self.entry_interval.get())
        region_names = self._get_selected_regions_simple()

        baseline_set = False

        while self.running:
            try:
                hwnd, title = find_kakao_window(room_name)
                if not hwnd:
                    self._log("[경고] 카카오톡 창을 찾을 수 없습니다.")
                    time.sleep(scan_interval)
                    continue

                img = capture_chat_area(hwnd)
                if img is None:
                    time.sleep(scan_interval)
                    continue

                positioned_lines = extract_positioned_lines_from_image(img)
                lines = [item["text"] for item in positioned_lines]

                if not lines:
                    self.empty_ocr_streak += 1
                    if self.empty_ocr_streak in (1, 10) or self.empty_ocr_streak % 30 == 0:
                        self._log("[대기] OCR 텍스트를 읽지 못했습니다. 카톡 창을 더 크게 띄우고 라이트 모드를 유지하세요.", "warning")
                    time.sleep(scan_interval)
                    continue

                self.empty_ocr_streak = 0
                self._log_ocr_preview(positioned_lines)
                latest_candidate = self._extract_latest_candidate(positioned_lines, region_names, response_text)

                if not baseline_set:
                    baseline_set = True
                    if latest_candidate:
                        self.last_seen_candidate = latest_candidate[0]
                        self.pending_candidate = None
                        self.pending_candidate_count = 0
                        self._log(f"기준선 설정 완료 (기존 마지막 메시지 무시: {latest_candidate[1]})")
                    else:
                        self.last_seen_candidate = None
                        self._log("기준선 설정 완료 (현재 화면에 조건 일치 메시지 없음)")
                    time.sleep(scan_interval)
                    continue

                if not latest_candidate:
                    self.pending_candidate = None
                    self.pending_candidate_count = 0
                    time.sleep(scan_interval)
                    continue

                candidate_signature, candidate_line, parsed = latest_candidate
                if self.last_seen_candidate == candidate_signature:
                    self.pending_candidate = None
                    self.pending_candidate_count = 0
                    time.sleep(scan_interval)
                    continue

                region, job, commission = parsed
                dup_key = make_duplicate_key(region, job, commission)

                if self.pending_candidate != candidate_signature:
                    self.pending_candidate = candidate_signature
                    self.pending_candidate_count = 1
                    if self.last_candidate_logged != candidate_signature:
                        self._log(f"[OCR후보] {candidate_line} (확정 대기 1/2)", "debug")
                        self.last_candidate_logged = candidate_signature
                    time.sleep(scan_interval)
                    continue

                self.pending_candidate_count += 1
                if self.pending_candidate_count < 2:
                    time.sleep(scan_interval)
                    continue

                self.last_seen_candidate = candidate_signature
                self.pending_candidate = None
                self.pending_candidate_count = 0

                if dup_key in self.processed_keys:
                    if self.last_duplicate_logged != candidate_signature:
                        self._log(f"[중복] 이미 처리함: {candidate_line}")
                        self.last_duplicate_logged = candidate_signature
                    time.sleep(scan_interval)
                    continue

                self.last_duplicate_logged = None

                now = time.time()
                if now - self.last_send_time < send_interval:
                    wait = send_interval - (now - self.last_send_time)
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
                    self.last_send_time = time.time()
                    self._log(f"[전송완료] '{response_text}'")
                else:
                    self._log("[전송실패]", "error")

            except Exception as e:
                self._log(f"[오류] {e}", "error")

            time.sleep(scan_interval)


def main():
    app = MacroApp()
    app.root.mainloop()


if __name__ == "__main__":
    main()
