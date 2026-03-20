# -*- coding: utf-8 -*-
"""
설비 수주 자동응답 매크로 v1.7
- 후보 확정 2회
- 제한 지역 데이터 적용
- 중복 로그에 이전 응답 이력 표시
"""

import time

import main_v1_6_전국1차 as base
from message_parser_v1_7 import (
    is_meaningful_ocr_line,
    make_duplicate_key,
    parse_message,
    sanitize_ocr_line,
)
from region_data_v1_7 import REGION_HIERARCHY

base.CONFIRM_THRESHOLD = 2
base.REGION_HIERARCHY = REGION_HIERARCHY
base.is_meaningful_ocr_line = is_meaningful_ocr_line
base.make_duplicate_key = make_duplicate_key
base.parse_message = parse_message
base.sanitize_ocr_line = sanitize_ocr_line


class MacroApp(base.MacroApp):
    def __init__(self):
        self.processed_records = {}
        super().__init__()
        self.root.title("설비 수주 자동응답 매크로 v1.7")

    def _reset_tracking(self):
        super()._reset_tracking()
        self.processed_records = {}

    def _log_ocr_preview(self, focus_lines):
        preview_lines = [sanitize_ocr_line(item["text"]) for item in focus_lines]
        preview_lines = [line for line in preview_lines if is_meaningful_ocr_line(line)]
        preview = " | ".join(preview_lines[-2:])
        if preview and preview != self.last_preview_logged:
            self.last_preview_logged = preview
            self._log(f"[OCR] 하단 인식: {preview}", "debug")

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
                    self.last_seen_candidate_key = None
                    self.pending_candidate_key = None
                    self.pending_candidate_line = None
                    self.pending_candidate_count = 0
                    self.last_candidate_logged = None
                    self.last_duplicate_logged = None
                    self._log(f"[설정] 감시 지역이 변경되었습니다: {self._format_regions_for_log()}", "warning")
                    self._log("[설정] 지역 변경 감지, 기준선을 다시 설정합니다.", "warning")

                hwnd, _ = base.find_kakao_window(room_name)
                if not hwnd:
                    self._log("[경고] 카카오톡 창을 찾을 수 없습니다.", "warning")
                    time.sleep(scan_interval)
                    continue

                img = base.capture_chat_area(hwnd)
                if img is None:
                    time.sleep(scan_interval)
                    continue

                positioned_lines = base.extract_positioned_lines_from_image(img)
                detected_chat_time = self._extract_chat_time(positioned_lines)
                if detected_chat_time:
                    self.current_chat_time = detected_chat_time

                if not positioned_lines:
                    self.empty_ocr_streak += 1
                    if self.empty_ocr_streak in (1, 10) or self.empty_ocr_streak % 30 == 0:
                        self._log("[대기] OCR 텍스트를 읽지 못했습니다. 카톡 창을 더 크게 띄우고 라이트 모드를 유지하세요.", "warning")
                    time.sleep(scan_interval)
                    continue

                self.empty_ocr_streak = 0
                focus_lines = self._extract_focus_lines(positioned_lines)
                self._log_ocr_preview(focus_lines)
                latest_candidate = self._extract_latest_candidate(focus_lines, current_region_names, response_text)

                if not baseline_set:
                    baseline_set = True
                    if latest_candidate:
                        self.last_seen_candidate_key = latest_candidate[0]
                        self._log(f"기준선 설정 완료 (기존 마지막 메시지 무시: {latest_candidate[1]})")
                    else:
                        self.last_seen_candidate_key = None
                        self._log("기준선 설정 완료 (현재 화면에 조건 일치 메시지 없음)")
                    time.sleep(scan_interval)
                    continue

                if not latest_candidate:
                    self.pending_candidate_key = None
                    self.pending_candidate_line = None
                    self.pending_candidate_count = 0
                    time.sleep(scan_interval)
                    continue

                candidate_key, candidate_line, parsed = latest_candidate
                if self.last_seen_candidate_key == candidate_key:
                    self.pending_candidate_key = None
                    self.pending_candidate_line = None
                    self.pending_candidate_count = 0
                    time.sleep(scan_interval)
                    continue

                region, job, commission = parsed
                dup_key = make_duplicate_key(region, job, commission)

                if self.pending_candidate_key != candidate_key:
                    self.pending_candidate_key = candidate_key
                    self.pending_candidate_line = candidate_line
                    self.pending_candidate_count = 1
                    if self.last_candidate_logged != candidate_key:
                        self._log(f"[OCR후보] {candidate_line} (확정 대기 1/{base.CONFIRM_THRESHOLD})", "debug")
                        self.last_candidate_logged = candidate_key
                    time.sleep(scan_interval)
                    continue

                self.pending_candidate_count += 1
                if self.pending_candidate_count < base.CONFIRM_THRESHOLD:
                    self._log(
                        f"[OCR후보] {self.pending_candidate_line} (확정 대기 {self.pending_candidate_count}/{base.CONFIRM_THRESHOLD})",
                        "debug",
                    )
                    time.sleep(scan_interval)
                    continue

                self.last_seen_candidate_key = candidate_key
                self.pending_candidate_key = None
                self.pending_candidate_line = None
                self.pending_candidate_count = 0

                if dup_key in self.processed_keys:
                    if self.last_duplicate_logged != candidate_key:
                        previous = self.processed_records.get(dup_key)
                        if previous:
                            self._log(
                                f"[중복] 이미 처리함: {candidate_line} (이전응답: {previous['line']} / {previous['chat_time']})"
                            )
                        else:
                            self._log(f"[중복] 이미 처리함: {candidate_line}")
                        self.last_duplicate_logged = candidate_key
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
                    base.win32gui.SetForegroundWindow(hwnd)
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
