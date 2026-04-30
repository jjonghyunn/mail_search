# mail_search_to_msg.py 가이드
<!-- 2026-04-30  Jonghyun Park w/ Claude -->

`team_name` 메일함에서 키워드 매칭되는 메일을 `.msg` 파일 **+ 첨부파일**로 다운로드하는 스크립트.

검색 결과는 `~/Downloads/mail_search_<YYMMDD>/` 폴더에 일괄 저장. 메일과 첨부가 모두 같은 날짜 prefix(`YYMMDD_HHMM_`)로 시작해서 어느 첨부가 어느 메일에 속하는지 한눈에 보임.

---

## 동작 흐름

```
Outlook.Application (win32com)
    ↓
GetNamespace("MAPI")
    ↓
Stores 중 STORE_NAME(부분 일치) 선택  → team_name
    ↓
SAVE_DIR / _processed_entry_ids.txt 로드 (있으면) → processed_ids set
    ↓
Inbox(받은편함) 또는 지정 폴더 진입
    ↓
(옵션) 하위 폴더 재귀 순회
    ↓
각 메일에 대해 — Subject + (옵션) Body lowercase 결합
    ↓
KEYWORDS 중 어느 하나라도 substring 매칭? (OR, 대소문자 무관)
    ↓ 매칭 시
EntryID 가 processed_ids 에 있으면 → skip (이미 저장된 메일)
    ↓ 신규
MailItem.SaveAs(path, 3)  # 3 = olMSG
    ↓ (옵션) 첨부 저장
Attachment.SaveAsFile(...) — 인라인 이미지는 SKIP_INLINE_IMAGES 시 제외
    ↓
EntryID 를 _processed_entry_ids.txt 에 append
    ↓
~/Downloads/mail_search_<YYMMDD>/
    ├─ <YYMMDD_HHMM>_<safe subject>.msg
    ├─ <YYMMDD_HHMM>_<원본 첨부 파일명>
    └─ _processed_entry_ids.txt    ← 재실행 시 중복 skip 마커
```

---

## 설정 (스크립트 상단)

| 변수 | 기본값 | 의미 |
|---|---|---|
| `KEYWORDS` | `["CAMPAIGN NAME"]` | 검색 키워드 리스트. 어느 하나라도 포함되면 매칭 (OR). 대소문자 무관 |
| `STORE_NAME` | `"team_name"` | Outlook 메일함 DisplayName 부분 일치로 검색 |
| `FOLDER_NAME` | `None` | `None`이면 받은편함(Inbox). 다른 폴더 이름(예: `"매칭 메일"`) 지정 가능 |
| `RECURSE_SUBFOLDERS` | `False` | True면 하위 폴더까지 재귀 검색 |
| `SEARCH_BODY` | `True` | False면 제목만 검색 (수천 개 메일 처리 시 훨씬 빠름) |
| `WHOLE_WORD` | `True` | True면 단어 경계(`\b`) 매칭 — `"ai"`가 `email`/`available` 안에서 매칭 안 됨. False면 단순 substring (예전 동작) |
| `SAVE_ATTACHMENTS` | `True` | 매칭 메일의 첨부파일도 같은 폴더에 저장 |
| `SKIP_INLINE_IMAGES` | `True` | 서명·인라인 이미지(`image001.png` 등 Outlook 자동 생성 이름) 제외 |

> ⚠️ **짧은 키워드 주의**: 2~3글자 영문 키워드(`ai`, `kv` 등)는 `WHOLE_WORD=True` 권장. False로 두면 `email`, `available`, `again`, `details`, `main` 등 흔한 단어 안에 substring으로 잡혀 과매칭 발생.

---

## 실행

```bash
python mail_search_to_msg.py
```

매번 키워드 바뀌면 스크립트 상단 `KEYWORDS` 리스트 수정 후 실행.

---

## 같은 날짜 폴더에서 키워드 바꿔가며 재실행

두 단계 dedup 으로 중복 저장 방지:

### 1. 메일 dedup — `_processed_entry_ids.txt` 마커

처리한 메일의 Outlook EntryID를 한 줄씩 append 기록.
재실행 시 같은 EntryID가 매칭되면 → 메일·첨부 모두 skip.

### 2. 첨부 dedup — 폴더 내 기존 첨부파일 스캔

SAVE_DIR 의 기존 파일들에서 `<YYMMDD_HHMM>_` prefix와 `(N)` counter 제거하여 원본명 set 구성.  
**다른 메일**이 매칭됐어도 첨부 원본명이 set에 있으면 skip.

→ "여러 메일에 같은 첨부파일이 들어있는" 경우에도 첨부는 한 번만 저장됨.

### 동작 예시

```
1차 실행 (KEYWORDS=["A"]):
  메일 X (첨부: report.xlsx) 저장 → 260415_0903_<X 제목>.msg + 260415_0903_report.xlsx
  EntryID(X) → _processed_entry_ids.txt
  saved_att_originals: {"report.xlsx"}

2차 실행 (KEYWORDS=["B"], 같은 날):
  메일 X (다시 매칭됨) → EntryID 매칭 → skip
  메일 Y (첨부: report.xlsx, 같은 파일명) → EntryID 신규지만 첨부 원본명 매칭 → 첨부만 skip
                                          → .msg는 저장됨
  메일 Z (첨부: notes.pdf) → 둘 다 신규 → 정상 저장
```

### 강제 재저장

`_processed_entry_ids.txt` 삭제 + 기존 첨부파일들 삭제 후 실행.

---

## 출력 파일명 규칙

```
메일:   <YYMMDD_HHMM>_<safe subject>.msg
첨부:   <YYMMDD_HHMM>_<원본 첨부 파일명>
```

- `YYMMDD_HHMM` — 메일 수신 시각 (`ReceivedTime`)
- `safe subject` — Windows 금지문자(`<>:"/\\|?*` + 제어문자) 모두 `_` 로 치환, 길이 150자 제한
- 첨부 원본명도 동일하게 sanitize (확장자 보존)
- 동일 파일명 중복 시 `(2)`, `(3)` ... 자동 부여 (.msg + 첨부 모두 같은 seen_names 셋 공유)
- 메일 + 그 메일의 첨부가 모두 같은 prefix로 시작 → 폴더에서 이름순 정렬 시 자연스럽게 묶임

---

## 성능 팁

- 받은편함 수천 개 메일에서 본문(Body)까지 검색하면 **수십 초~수 분** 소요
- 빠른 검색이 필요하면 `SEARCH_BODY = False` 로 제목만 검색
- 폴더 정렬: `Items.Sort("[ReceivedTime]", True)` 최신 순
- 진행 표시: 500개마다 `진행 N/Total (저장 X, 실패 Y)` 출력

---

## 주의사항

| 항목 | 내용 |
|---|---|
| Outlook 실행 필요 | `win32com.client.Dispatch("Outlook.Application")` — Outlook 데스크톱 앱 설치/로그인 상태여야 함 |
| 메일함 권한 | `team_name` 가 본인 Outlook 프로필에 등록돼 있어야 `Stores`에서 찾힘 |
| 본문 인코딩 | `mail.Body`는 plain text. HTMLBody가 필요하면 별도 처리 |
| `.msg` 포맷 | Outlook 기본 메일 저장 포맷. 다른 OS / 클라이언트에선 열기 제한적 |
| 매칭 0개 | 폴더만 만들고 종료 (저장 0개) — 키워드 / 폴더 / 메일함 명 재확인 |
| `.msg` 안의 첨부 | `.msg` 자체에 첨부가 임베드되어 있음. 추가로 별도 첨부파일도 폴더에 저장됨 (`SAVE_ATTACHMENTS=True` 기본) |
| 인라인 이미지 | `image001.png` 형식의 자동 생성 이름은 `SKIP_INLINE_IMAGES=True` 일 때 자동 skip — 일반 첨부만 저장 |
| MailItem 외 항목 | `Class != 43` (회의·작업·연락처 등)은 자동 skip |

---

## 변경 이력

- **2026-04-30** (Jonghyun Park) — 초기 작성
- **2026-04-30** (Jonghyun Park) — 매칭 메일의 첨부파일도 같은 폴더에 저장하도록 추가 (`SAVE_ATTACHMENTS`, `SKIP_INLINE_IMAGES` 옵션). 첨부 파일명도 `<YYMMDD_HHMM>_<원본명>` prefix로 통일.
- **2026-04-30** (Jonghyun Park) — EntryID 기반 중복 방지 마커 (`_processed_entry_ids.txt`) 추가. 같은 날짜 폴더에서 키워드 바꿔가며 재실행 시 같은 메일 두 번 저장되지 않음.
- **2026-04-30** (Jonghyun Park) — 첨부 원본명 기반 dedup 추가. 서로 다른 메일에 같은 이름의 첨부가 들어있어도 첨부는 한 번만 저장. SAVE_DIR 의 기존 첨부 파일들에서 prefix·counter 제거하여 원본명 set 구성.
- **2026-04-30** (Jonghyun Park) — `WHOLE_WORD` 옵션 추가 (기본 True). 짧은 키워드(`ai`, `kv` 등)가 다른 단어 안에 substring으로 잡혀 과매칭되던 문제 해결. `\b` 정규식 경계로 단어 단위 매칭.
