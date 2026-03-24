# 개발정의서 — 설비 수주 자동응답 매크로 (v2_1_GUI)

## 1. 개요

| 항목 | 내용 |
|------|------|
| 파일명 | `macro_v2_1_GUI.py` |
| 실행파일 | `설비수주매크로_v2_1.exe` |
| 버전 | v2.1 |
| 작성일 | 2026-03-24 |
| 기반 버전 | macro_v1_3_GUI.py |
| 실행 환경 | Windows 10 이상, Python 3.10+, Tesseract-OCR |

---

## 2. 핵심 업그레이드 요약 (v1.3 → v2.1)

| 구분 | v1.3 | v2.1 |
|------|------|------|
| 인식 단위 | 줄(Line) | 말풍선 블록(Message Block) |
| 후보 추출 함수 | `_extract_all_candidates` (줄 단위) | `_extract_block_candidates` (블록 단위) |
| 파싱 함수 | `parse_message(단일 줄)` | `parse_message(블록 합친 텍스트)` → 실패 시 줄별 재시도 |
| 기준선 방식 | 줄 기반 dup_key 등록 | 블록 기반 dup_key 등록 (Smart Baseline) |
| 중복 차단 | `sent_group_keys` 30초 grace | `seen_hashes` 300초 TTL (강화 대체) |
| 확인 카운터 | `pending_candidates` CONFIRM_THRESHOLD=2 | 동일하게 유지 |
| 큐 순서 | y 기준 최신 우선 (FIFO) | 동일하게 유지 (`_enqueue_ready_candidates`) |
| 작업 키워드 수 | 4개 | 8개 |
| 공백 처리 | 부분 적용 | 블록 결합 후 공백 제거(_norm) 전면 적용 |

---

## 3. 버전별 변경 이력

### v1.1
- 전송 직전 `processed_keys` 선행 등록(pre-mark)
- 수수료 OCR 오인식 계열(1↔10, 2↔20, 3↔30) 동시 차단
- `QUEUE_TTL` 2.0s → 6.0s 확장

### v1.2
- 중지 버튼 시 큐/대기 후보 즉시 비움
- fuzzy 매칭 최소 3글자 이상 제한
- 동일 (지역+작업) 쌍 30초 유예(grace) 적용

### v1.3
- 수수료 허용 단축 패턴 추가 (`수10`, `수 10`, `수10%`, `수 10%`)
- 인식 키워드 4개로 변경 (싱막, 하막, 역류, 고압)

### v2.1 ← 현재 버전
- **말풍선 블록 기반 인식 엔진** (줄 단위 → 블록 단위 교체)
- **스마트 기준선**: 시작 시 현재 블록 파싱 → dup_key 등록
- **`_extract_block_candidates`**: 블록 합친 텍스트에 `parse_message` 적용 (v1.3 파싱 로직 완전 재사용)
- **작업 키워드 8종 확장** (싱크대, 싱크대역류, 싱크대막힘, 하수구막힘 추가)
- **`seen_hashes` 300초 TTL** (v1.3의 `sent_group_keys` 30초 → 300초 강화 대체)
- **v1.3 모든 안정성 메커니즘 완전 유지**: pending_candidates, processed_keys, CONFIRM_THRESHOLD, `_enqueue_ready_candidates`, 수수료 오인식 차단 등

---

## 4. 수정 상세

### 4-1. 신규 상수

```python
BLOCK_Y_THRESHOLD = 40    # 인접 줄 y 좌표 차이 임계값(px) — 블록 분리 기준 (2x 이미지 기준)
SEEN_HASH_TTL     = 300.0 # 동일 (지역+작업) 재전송 차단 시간(초) — 5분
```

### 4-2. 작업 키워드 확장 (JOB_NORMALIZE)

| 키워드 | 상태 | OCR 변형 허용값 |
|--------|------|----------------|
| 싱막 | 유지 | 성막, 생막, 씽막, 싱 막, 성 막, 생 막 |
| 하막 | 유지 | 하 막 |
| 역류 | 유지 | 역 류, 역루, 여류, 역유 |
| 고압 | 유지 | 고 압, 고암, 코압 |
| 싱크대 | 신규 | 씽크대, 싱크 대, 씽크 대 |
| 싱크대역류 | 신규 | 씽크대역류, 싱크대 역류, 씽크대 역류, 싱크역류 |
| 싱크대막힘 | 신규 | 씽크대막힘, 싱크대 막힘, 씽크대 막힘 |
| 하수구막힘 | 신규 | 하수구 막힘 |

`normalize_job` OCR 보정 추가: `씽크` → `싱크`

### 4-3. 말풍선 블록 그룹화 (`group_lines_into_blocks`)

```
extract_positioned_lines_from_image 결과 (y 좌표 정렬)
    ↓
인접 두 줄 y 차이 > BLOCK_Y_THRESHOLD(40px) → 새 블록 시작
시간 정보 줄(오전/오후 H:MM[:SS]) 감지 → 현재 블록 종료 후 타임스탬프 줄 제외
    ↓
focus_blocks = 하단 15개 블록 사용
```

> OCR 전처리 이미지는 2배 확대이므로 y 좌표도 2배.
> 실제 라인 간격 ~20px → OCR y 차이 ~40px. 임계값 40으로 설정.

### 4-4. `_extract_block_candidates` (핵심 변경)

v1.3의 `_extract_all_candidates`를 블록 단위로 교체한 함수.
`parse_message` (v1.3과 동일)를 블록 합친 텍스트에 그대로 적용한다.

```
focus_blocks 순회
    ↓
block_combined_text(block) → block_text
fix_ocr_place_tokens(block_text) → raw
    ↓
parse_message(raw, regions, response_text) 시도
    ↓ 실패 시
각 줄 개별로 parse_message 재시도
    ↓
(block_y, dup_key, display_line[:80], parsed, strong) 반환
```

**왜 `parse_message`를 사용하는가:**
- v1.3의 강력한 노이즈 필터(`is_noise_line`), 지역 fuzzy 매칭, 수수료 패턴 파싱을 그대로 활용
- 블록 합친 텍스트 = 여러 줄 의뢰문구를 하나의 문자열로 취급 → 다중 줄 의뢰 인식 가능

### 4-5. 스마트 기준선 (Smart Baseline)

시작 버튼 클릭 시:

```
blocks = group_lines_into_blocks(현재 화면)
baseline_block_hash = compute_block_hash(blocks[-1])  ← 로그 출력용
for block in focus_blocks:
    result = _extract_block_candidates([block], ...)
    → dup_key를 processed_keys에 등록 (의미 기반 키 차단)
```

### 4-6. `_monitor_loop` 구조 (v1.3 완전 동일)

| 단계 | v1.3 | v2.1 |
|------|------|------|
| OCR 결과 처리 | `_extract_focus_lines` | `group_lines_into_blocks` |
| 후보 추출 | `_extract_all_candidates` | `_extract_block_candidates` |
| 카운터 갱신 | pending_candidates | 동일 |
| 큐 적재 | `_enqueue_ready_candidates` | 동일 |
| 순서 보장 | y 기준 최신 우선 | 동일 |
| 중복 차단 | `sent_group_keys` | `seen_hashes` (TTL만 300초로 강화) |

### 4-7. `seen_hashes` TTL 관리 (300초)

v1.3의 `sent_group_keys` 30초 grace를 대체하며 TTL을 300초로 강화한다.

```python
# 전송 완료 후
self.seen_hashes[group_key] = self.last_send_time  # group_key = "지역|작업"

# 매 스캔 만료 정리
self.seen_hashes = {k: t for k, t in self.seen_hashes.items() if now - t < SEEN_HASH_TTL}

# 재진입 차단
if group_key in self.seen_hashes:
    if now - self.seen_hashes[group_key] < SEEN_HASH_TTL:
        processed_keys.add(dup_key)  # 영구 차단 등록
        continue
```

---

## 5. v1.3 안정성 메커니즘 유지 목록

| 메커니즘 | v1.3 | v2.1 |
|----------|------|------|
| `pending_candidates` 확인 카운터 | O | O (유지) |
| `CONFIRM_THRESHOLD = 2` | O | O (유지) |
| `FAST_CONFIRM_THRESHOLD = 1` (강확인) | O | O (유지) |
| `JOB_CONFIRM_OVERRIDES` | O | O (유지) |
| `processed_keys` 세션 영구 차단 | O | O (유지) |
| `_mark_related_commissions_processed` | O | O (유지) |
| `_is_ambiguous_commission_candidate` | O | O (유지) |
| 수수료 오인식 쌍 차단 (1↔10 등) | O | O (유지) |
| `QUEUE_TTL_SECONDS` 큐 만료 | O | O (유지) |
| `_prune_send_queue` | O | O (유지) |
| `_enqueue_ready_candidates` (y 기준 정렬) | O | O (유지) |
| `_pop_next_queued_candidate` | O | O (유지) |
| `pending_candidates["y"]` 필드 | O | O (유지) |
| `sent_group_keys` 30초 grace | O | → `seen_hashes` 300초로 강화 대체 |

---

## 6. MacroApp 상태 변수

| 변수 | v1.3 | v2.1 |
|------|------|------|
| `pending_candidates` | O | O (유지) |
| `send_queue` | O | O (유지) |
| `queued_keys` | O | O (유지) |
| `processed_keys` | O | O (유지) |
| `sent_group_keys` | O | 제거 (`seen_hashes`로 대체) |
| `processed_records` | O | 제거 (불필요) |
| `seen_hashes` | — | 신규 — group_key → timestamp, 300초 TTL |
| `baseline_block_hash` | — | 신규 — 로그 출력용 기준 블록 해시 |

---

## 7. 주요 상수 및 설정값

| 상수 | 값 | 설명 |
|------|----|------|
| `BLOCK_Y_THRESHOLD` | 40px | 블록 분리 y 차이 임계값 (2x 이미지 기준) |
| `SEEN_HASH_TTL` | 300.0s | 동일 (지역+작업) 재전송 차단 시간 |
| `CONFIRM_THRESHOLD` | 2 | 일반 후보 확정 횟수 |
| `FAST_CONFIRM_THRESHOLD` | 1 | 강확인 후보 확정 횟수 |
| `PENDING_TTL_SECONDS` | 1.6s | 후보 유지 TTL |
| `QUEUE_TTL_SECONDS` | 6.0s | 큐 항목 만료 시간 |
| `MIN_OCR_WORD_CONF` | 20.0 | 최소 OCR 신뢰도 |

---

## 8. 전체 인식 흐름 (v2.1 최종)

```
카카오톡 창 캡처
    ↓
OCR (Tesseract kor+eng, PSM 6/11, threshold/contrast 유지 — v1.3 동일)
    ↓
extract_positioned_lines_from_image → positioned_lines
    ↓
group_lines_into_blocks (y_threshold=40px, 타임스탬프 분리)
    ↓
focus_blocks = 하단 15개 블록
    ↓
[시작 시] _extract_block_candidates → dup_key 등록 → processed_keys (Smart Baseline)
    ↓
[이후 스캔] _extract_block_candidates 호출
  ① block_combined_text → parse_message (v1.3 로직 그대로 적용)
  ② 실패 시 각 줄 개별 parse_message 재시도
    ↓
(y, dup_key, display_line, parsed, strong) → pending_candidates 카운터 갱신
    ↓
CONFIRM_THRESHOLD 도달 → _enqueue_ready_candidates (y 기준 최신 우선 정렬)
    ↓
processed_keys.add(dup_key) (pre-mark) → SetForegroundWindow → 응답 전송 ("네")
    ↓
seen_hashes[group_key] = now (300초 TTL) → 동일 (지역+작업) 재응답 차단
```

---

## 9. 실행 환경 요구사항

- Python 3.10+
- Tesseract-OCR (`C:\Program Files\Tesseract-OCR\tesseract.exe`)
- 언어팩: `kor`, `eng`
- 패키지: `pyautogui`, `pyperclip`, `pytesseract`, `Pillow`, `pywin32`
- 카카오톡 PC 버전 (라이트 모드 권장)

---

## 10. 파일 목록

| 파일 | 설명 |
|------|------|
| `macro_v2_1_GUI.py` | v2.1 소스 코드 |
| `dist/설비수주매크로_v2_1.exe` | 단독 실행 파일 (약 44MB) |
| `개발정의서_설비수주매크로(v2_1_GUI).md` | 본 문서 |
| `개발정의서_설비수주매크로(v2_1_GUI).docx` | 워드 문서 |
