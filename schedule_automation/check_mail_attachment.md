# check_mail_attachment — 첨부파일 자동 감지·저장  
<sub>2026-08-19  Jonghyun Park w/ Claude</sub>  

Outlook 메일함을 주기적으로 훑어 **조건에 맞는 첨부파일만** 지정 폴더에 내려받는 스크립트 모음.
`mail_search_to_msg.py` 가 "찾아서 한 번 받아오는" 도구라면, 이쪽은 **작업 스케줄러에 걸어두고
계속 지켜보는** 쪽이다. 새 파일이 오면 조용히 저장하고, 이미 처리한 메일은 다시 건드리지 않는다.

## 3가지 변형

조건만 다른 **자립 스크립트**다. 공통 모듈로 묶지 않았다 — 한 곳의 조건을 바꿀 때 다른 감시까지
흔들리지 않게 하려는 것이고, 새 용도가 생기면 파일 하나 복사해 상단 상수만 고치면 된다.

| | `_byname` | `_status` | `_url` |
|---|---|---|---|
| 메일 제목 조건 | — | — | `SUBJECT_KEYS` |
| 첨부파일명 조건 | `ATTACHMENT_KEYS` | `ATTACHMENT_KEYS` | `ATTACHMENT_KEYS` |
| 수신일 필터 | `RECEIVED_FROM` 이후 | `RECEIVED_FROM` 이후 | `RECEIVED_FROM` ~ `RECEIVED_TO` |
| 처리 이력 마커 | `Documents/` 아래 | 저장 폴더 안 | 저장 폴더 안 |

- **`_byname`** — 첨부파일명만 보고 판단. 제목이 매번 다르게 오는 메일에 쓴다.
- **`_status`** — 같은 방식이되 마커를 저장 폴더 안에 둔다. 폴더를 통째로 복사·이동하면
  처리 이력도 따라가므로, 건마다 폴더를 새로 파는 운영에 편하다.
- **`_url`** — 제목과 첨부파일명을 **둘 다** 봐야 걸러지는 경우. 수신 기간도 양끝으로 막는다.

## 설정 (각 파일 상단)

| 상수 | 설명 |
|---|---|
| `SAVE_FOLDER` | 첨부를 내려받을 폴더 |
| `STORE_NAME` | Outlook 메일함 DisplayName (부분 일치). 공유 메일함이면 그 이름 |
| `ATTACHMENT_KEYS` | 첨부파일명 조건 |
| `SUBJECT_KEYS` (`_url`) | 메일 제목 조건 |
| `RECEIVED_FROM` / `RECEIVED_TO` | 수신일 범위 |
| `ALLOWED_EXT` (`_url`) | 저장할 확장자 |
| `MARKER_FILE` | 처리한 메일 EntryID 기록 파일 |

### 키워드 조건 읽는 법

리스트 원소는 **모두 포함(AND)**, 원소가 다시 리스트면 그 안은 **OR**.

```python
ATTACHMENT_KEYS = ["campaign name", ["Campaign", "캠페인"]]
```

→ 파일명에 `campaign name` 이 있고, **그리고** `Campaign` 또는 `캠페인` 중 하나가 있어야 통과.
영문/한글 표기가 섞여 오는 첨부에 쓴다. 매칭은 소문자 비교라 대소문자는 무관.

## 동작

1. `STORE_NAME` 과 이름이 맞는 메일함을 찾아 받은편지함을 연다 (기본 프로필이 아니라 **지정 메일함**).
   못 찾으면 그대로 종료한다.
2. `RECEIVED_FROM` 이전 메일은 건너뛴다 — 지난 건의 옛 메일이 딸려오는 걸 막는다.
3. 마커에 이미 있는 EntryID 는 건너뛴다.
4. 조건에 맞는 첨부를 `SAVE_FOLDER` 에 저장한다.
5. 처리한 메일의 EntryID 를 마커에 덧붙인다.

### 같은 파일명이 다시 왔을 때

수신일시를 붙여 따로 저장한다.

```
이미 있음 : ···20260409_v2.xlsx
새로 수신 : ···20260409_v2_260414_1432.xlsx     (260414 14:32 수신)
```

날짜만(`_yymmdd`) 붙이면 **같은 날 두 번 온 파일이 이름까지 같아져 두 번째가 유실**된다.
받는 쪽에서 "가장 최신 파일" 을 고를 때도 두 파일의 정렬 키가 동점이 되어 순서가 갈리지 않는다.
그래서 분 단위 시각까지 붙인다. 시각까지 같은 파일이 또 오면 그건 같은 메일이므로 저장 없이 넘긴다.

### 긴 경로(MAX_PATH)

저장 경로가 260자를 넘으면 `att.SaveAsFile()` 이 실패한다. 임시 폴더에 먼저 저장한 뒤
`shutil.move()` 로 옮겨 우회한다. 레지스트리 `LongPathsEnabled = 1` 도 함께 켜두면 안전하다.

## 실행

```bash
python check_mail_attachment_byname.py
```

무인 실행은 작업 스케줄러에 `pythonw.exe` 로 등록한다 (창이 뜨지 않는다).
등록 명령어 예시는 같은 폴더의 [`create_schtasks_v2.txt`](create_schtasks_v2.txt) 에
실행 시각 배치와 함께 정리돼 있다.

## 받은 파일을 쓰는 쪽

내려받은 일정 파일을 리포트 워크북에 반영하는 도구는 같은 폴더의
[`update_schedule_summary.py`](update_schedule_summary.py) 다.
이 스크립트가 붙이는 `_YYMMDD_HHMM` 접미사를 그쪽 최신 파일 판별이 그대로 읽는다.

요약 정제 단계가 없는 앞 세대 `update_schedule.py` 는 [`data-preprocessing`](https://github.com/jjonghyunn/data-preprocessing/tree/main/260324_schedule) 에 있다.

## 의존성

`win32com` (Outlook 데스크톱 앱 필요) — `summarize_msgs.py` 와 달리 Outlook 이 깔려 있어야 한다.

```bash
pip install pywin32
```
