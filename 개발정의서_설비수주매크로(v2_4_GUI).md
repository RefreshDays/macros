# 개발정의서 — 설비 수주 자동응답 매크로 (v2_4_GUI)

## 1. 개요

| 항목 | 내용 |
|------|------|
| 파일명 | `macro_v2_4_GUI.py` |
| 버전 | v2.4 |
| 기반 | `macro_v2_3_GUI.py` |
| 추가 의존성 | OpenCV (`opencv-python` 또는 `opencv-python-headless`) |

---

## 2. v2.4 수정 원인 및 해결

### 2.1 기준선인데 위쪽(이미 있던) 말풍선에 다시 “네”가 나가던 현상

**원인**

1. **`FOCUS_BLOCK_COUNT`(28) 절단**  
   `build_focus_blocks_from_incoming_bubbles` / `_focus_blocks_from_positioned`가 **세로 기준 하단 28개 말풍선만** 후보로 쓰고, 기준선 등록도 그 **절단된 목록**만 사용함.  
   화면 **상단에 남아 있는 오래된 흰 말풍선**은 기준선에 **한 번도 `processed_keys`에 안 올라감** → 이후 스크롤·OCR 변동 시 “새 글”처럼 처리될 수 있음.

2. **OCR 변동 vs `dup_key`**  
   같은 말풍선이라도 스캔마다 작업명·지역 인식이 달라 `dup_key`(지역|작업|수수료)가 바뀌면, `processed_keys`만으로는 기존 글을 막기 어려움.

**해결**

| 구분 | v2.3 | v2.4 |
|------|------|------|
| 기준선에 쓰는 블록 | 하단 28개만 | **화면에 검출된 수신 말풍선 전체** (`trim_to_focus=False`로 수집) |
| 일반 감시 루프 | 하단 28개 | 동일(성능 유지) |
| 추가 무시 규칙 | — | 말풍선 합친 원문의 **`bubble_raw_fingerprint`**(한글·숫자만 정규화 후 MD5)를 `baseline_fingerprints`에 넣고, 이후 동일 지문 블록은 후보에서 제외 |

기준선 시점에 **줄별·합친 텍스트**로 파싱 가능한 **`dup_key`는 모두** `processed_keys`에 넣음 (`collect_dup_keys_from_block`).

### 2.2 `[조건충족]` 로그가 지저분함

- v2.4: `[조건충족]`·`[OCR후보]`(확정 대기)는 **`format_match_summary(parsed)`** 한 줄만 출력  
  예: `지역:신도림동 | 작업:역류 | 수수료:10%`  
- 전송 완료 후 `processed_records`의 `line`도 동일 요약으로 저장.

### 2.3 GUI 기본값

| 항목 | v2.3 기본 | v2.4 기본 |
|------|-----------|-----------|
| 스캔 주기 | 0.3초 | **0.2초** |
| 전송 간격 | 2.0초 | **1.0초** |

입력창 값은 그대로 사용자가 바꾸면 런타임에 반영됨.

---

## 3. 상수·함수 (v2.4 추가)

| 이름 | 설명 |
|------|------|
| `bubble_raw_fingerprint` | 말풍선 OCR 합친 문자열 정규 지문 |
| `collect_dup_keys_from_block` | 한 블록에서 나올 수 있는 모든 `dup_key` 수집 |
| `format_match_summary` | 로그용 파싱 요약 문자열 |
| `build_focus_blocks_from_incoming_bubbles(..., trim_to_focus=)` | 기준선은 `False`, 일반 경로는 내부에서 슬라이스 |

---

## 4. 전체 흐름 요약 (v2.4 반영)

```
채팅창 캡처
    ↓
[OpenCV] 흰 말풍선 ROI → 블록 전체 수집 (trim 없음)
    ↓
폴백 시 전체 OCR → group → 블록 전체 (trim 없음)
    ↓
focus_blocks = 하단 FOCUS_BLOCK_COUNT개만 (감시용)
    ↓
[감시 시작 첫 프레임] 기준선:
  · 모든 블록 지문 등록 + 가능한 dup_key 전부 processed_keys
    ↓
_extract_block_candidates: baseline 지문 일치 블록 스킵
    ↓
(v2.3와 동일) sent_region_keys, pending, 큐, pre-mark, 전송
```

---

## 5. 파일 목록

| 파일 | 설명 |
|------|------|
| `macro_v2_4_GUI.py` | 본 버전 소스 |
| `requirements.txt` | 기존과 동일 |
| `개발정의서_설비수주매크로(v2_4_GUI).md` | 본 문서 |
