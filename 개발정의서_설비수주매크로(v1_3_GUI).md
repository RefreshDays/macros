# 개발정의서 — 설비 수주 자동응답 매크로 (v1_3_GUI)

## 1. 개요

| 항목 | 내용 |
|------|------|
| 파일명 | `macro_v1_3_GUI.py` |
| 버전 | v1.3 |
| 작성일 | 2026-03-23 |
| 기반 버전 | macro_v1_2_GUI.py |
| 실행 환경 | Windows 10 이상, Python 3.10+, Tesseract-OCR |

---

## 2. 버전별 변경 이력

### v1.1
- 중복 응답 완전 차단
  - 전송 직전 `processed_keys` 선행 등록 (pre-mark)
  - 수수료 OCR 오인식 계열 (1↔10, 2↔20, 3↔30) 동시 차단
  - `QUEUE_TTL` 2.0s → 6.0s 확장

### v1.2
- 중지 버튼 시 큐/대기 후보 즉시 비움 + 인터럽트 sleep + `running` 재확인
- fuzzy 매칭 대상 최소 3글자 이상으로 제한 (2글자 지역 오인식 방지)
- 동일 (지역+작업) 쌍 전송 후 30초 유예(grace) 적용

### v1.3 ← 현재 버전
- **수수료 허용 단축 패턴 추가** (`수10`, `수 10`, `수10%`, `수 10%`)
- **인식 키워드 4개로 변경** (`싱막`, `하막`, `역류`, `고압`)

---

## 3. 수정 상세

### 3-1. 허용 수수료 패턴 추가

#### 배경
기존에는 수수료를 `수수료10`, `수수10`, `10%수수` 형태만 인식했다.  
채팅에서 `수10`, `수 10`, `수10%`, `수 10%` 와 같이 `수` 단독 접두 형태로 입력되는 케이스를 추가 인식해야 한다.

#### 변경 위치

**`has_exact_commission_token`** — 강확인(strong candidate) 판별 토큰 목록

```python
# 추가된 토큰
f"수{commission}",
f"수{commission}%",
```

**`extract_commission`** — 수수료 숫자 추출 로직 (compact_patterns 이후 최후 시도)

```python
# 허용 단축 패턴: 수10 / 수 10 / 수10% / 수 10%
short = re.search(r"수\s*(\d{1,2})\s*%?", text)
if short:
    value = int(short.group(1))
    if 1 <= value <= 30:
        full = short.group(0)
        if not (value == 1 and "%" not in full):
            return value
```

> **설계 원칙**: 기존 `수수료`, `수수` 패턴보다 후순위로 배치하여 오인식 최소화.  
> 값이 `1`이고 `%` 표기가 없을 경우 노이즈로 간주하여 건너뜀.

---

### 3-2. 인식 키워드 변경

#### 배경
기존 6개 작업 키워드(`싱크대막힘`, `하수구막힘`, `싱막`, `하막`, `역류`, `누수`)에서  
`싱크대막힘`, `하수구막힘`, `누수`를 제거하고 `고압`을 신규 추가하여 4개로 정리.

#### `JOB_NORMALIZE` 변경

| 구분 | v1.2 | v1.3 |
|------|------|------|
| 유지 | 싱막, 하막, 역류 | 싱막, 하막, 역류 |
| 제거 | 싱크대막힘, 하수구막힘, 누수 | — |
| 추가 | — | **고압** |

```python
# v1.3 JOB_NORMALIZE
JOB_NORMALIZE = [
    ("싱막", ["싱막", "성막", "생막", "씽막", "싱 막", "성 막", "생 막"]),
    ("하막", ["하막", "하 막"]),
    ("역류", ["역류", "역 류", "역루", "여류", "역유"]),
    ("고압", ["고압", "고 압", "고암", "코압"]),
]
```

#### 연관 코드 일괄 수정

| 위치 | v1.2 | v1.3 |
|------|------|------|
| `JOB_CONFIRM_OVERRIDES` | `{"역류": 1, "누수": 1}` | `{"역류": 1}` |
| `normalize_job` replacements | 누수 오인식 보정 포함 | 누수 제거, 고압(`고암`→`고압`, `코압`→`고압`) 추가 |
| `job_keywords` (parse_message) | `싱막, 하막, 역류, 누수, 싱크대막힘, 하수구막힘, 성막, 생막, 씽막` | `싱막, 하막, 역류, 고압, 성막, 생막, 씽막` |
| GUI 타이틀 | `macro_v1_2 GUI` | `macro_v1_3 GUI` |

---

## 4. 주요 상수 및 설정값

| 상수 | 값 | 설명 |
|------|----|------|
| `CONFIRM_THRESHOLD` | 2 | 일반 후보 확정 횟수 |
| `FAST_CONFIRM_THRESHOLD` | 1 | 강확인 후보 확정 횟수 |
| `PENDING_TTL_SECONDS` | 1.6s | 후보 유지 TTL |
| `QUEUE_TTL_SECONDS` | 6.0s | 큐 항목 만료 시간 |
| `SENT_GROUP_GRACE_SECONDS` | 30.0s | 동일 (지역+작업) 재전송 차단 유예 |
| `MAX_QUEUE_SIZE` | 10 | 최대 큐 크기 |
| `MIN_OCR_WORD_CONF` | 20.0 | 최소 OCR 신뢰도 |

---

## 5. 인식 흐름 요약

```
카카오톡 창 캡처
    ↓
OCR (Tesseract kor+eng, PSM 6/11)
    ↓
sanitize → fix_ocr_place_tokens → is_noise_line 필터
    ↓
extract_commission  ←─ 수수료 패턴 매칭
normalize_job       ←─ 작업 키워드 매칭 (싱막/하막/역류/고압)
    ↓
지역 매칭 (selected_regions 기반)
    ↓
pending_candidates 카운트 누적 → CONFIRM_THRESHOLD 도달
    ↓
send_queue 적재 → 응답 전송 ("네")
    ↓
processed_keys 등록 (중복 차단)
```

---

## 6. 실행 환경 요구사항

- Python 3.10+
- Tesseract-OCR (`C:\Program Files\Tesseract-OCR\tesseract.exe`)
- 언어팩: `kor`, `eng`
- 패키지: `pyautogui`, `pyperclip`, `pytesseract`, `Pillow`, `pywin32`
- 카카오톡 PC 버전 (라이트 모드 권장)
