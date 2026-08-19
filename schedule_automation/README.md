# schedule_automation — 메일 첨부 감시 → 워크북 자동 반영  
<sub>2026-08-19  Jonghyun Park w/ Claude</sub>  

메일함에 새 일정 파일이 오는지 지켜보다가, 오면 받아서 리포트 워크북에 반영하고
지난번과 달라진 칸에 음영을 칠하는 데까지가 한 묶음이다.
매일 사람이 메일함을 열어보고 손으로 붙여넣던 일을 대신한다.

상위 폴더의 `mail_search_to_msg.py` 가 "찾아서 한 번 받아오는" 도구라면,
이쪽은 **작업 스케줄러에 걸어두고 계속 도는** 쪽이다.

## 흐름

```
메일 도착
   ↓  check_mail_attachment_*.py      (조건 맞는 첨부만 저장, 이미 받은 건 건너뜀)
소스 폴더에 xlsx 쌓임
   ↓  update_schedule_summary.py      (최신본 선별 → 요약 시트 정제 → 워크북 반영 → 변경 셀 음영)
리포트 워크북 갱신
```

## 파일

| 파일 | 역할 |
|---|---|
| `check_mail_attachment_byname.py` | 첨부 감시 — **첨부파일명** 키워드, 수신일 이후만 |
| `check_mail_attachment_status.py` | 위와 같되 처리 이력 마커를 **저장 폴더 안**에 둠 |
| `check_mail_attachment_url.py` | **메일 제목 + 첨부파일명 2중 조건** + 수신 기간(From~To) |
| `check_mail_attachment.md` | 위 3종 가이드 |
| `update_schedule_summary.py` | 최신 소스 선별 → 요약 시트 정제 → 워크북 반영 → 변경 셀 음영 → 강제 재계산 |
| `update_schedule_summary.md` | 위 스크립트 가이드 |
| `26_Schedule(Auto)_example.xlsx` | 반영 대상 워크북이 어떻게 생겼는지 보여주는 예시 |
| `create_schtasks_v2.txt` | 작업 스케줄러 등록 명령어 + 실행 시각 배치 |

## 감시 스크립트 3종

조건만 다른 **자립 스크립트**다. 공통 모듈로 묶지 않았다 — 한 곳의 조건을 바꿀 때 다른 감시까지
흔들리지 않게 하려는 것이고, 새 용도가 생기면 파일 하나 복사해 상단 상수만 고치면 된다.
자세한 설정은 [`check_mail_attachment.md`](check_mail_attachment.md).

## 반영 스크립트

`update_schedule_summary.py` 는 받은 파일 중 **진짜 최신본**을 골라 워크북에 넣는다.
파일명에서 문서날짜·시각·버전·메일수신일시를 뽑아 튜플로 비교하므로,
`_v2` 와 `_v10`, 같은 날 두 번 온 파일이 뒤섞여도 순서가 갈린다.

앞에 **요약 시트 정제** 단계가 붙어 있는 것이 특징이다. 메일로 오는 파일이 사람이 읽으라고 만든
자유 서식이라, `6/24~9/13` 같은 기간 문자열을 시작일·종료일 두 칸으로 쪼개고 `TBU`·빈칸을
정리해야 기계가 쓸 수 있다. 그 손작업을 대신한다. 자세한 규칙은
[`update_schedule_summary.md`](update_schedule_summary.md).

## 예시 워크북

`26_Schedule(Auto)_example.xlsx` 는 반영 대상 워크북의 구조를 보여주는 예시다.

- 13개 시트, 컬럼 배치와 서식이 실제와 같고 **변경 셀 노란 음영도 그대로** 들어있다.
- **수식이 제거된 값 스냅샷**이다. 원본은 1만 개가 넘는 수식으로 시트끼리 얽혀 있는데,
  예시는 계산 결과만 남겨 열어보기만 해도 구조가 보이도록 했다.
- 회사·법인·국가·사이트코드·날짜·비고는 마스킹돼 있다. 토큰 단위 **일관 치환**이라
  시트 간 사이트코드 매칭은 예시에서도 그대로 성립하고, 날짜는 전체를 같은 폭으로 옮겨
  **기간 길이와 선후관계가 보존**된다.
- `update_schedule_summary.py` 의 `TARGET_SHEET` 가 가리키는 `고객법인일정파일` 시트가
  붙여넣기 대상이다.

## 실행

```bash
python check_mail_attachment_byname.py
python update_schedule_summary.py
```

무인 실행은 `pythonw.exe` 로 작업 스케줄러에 등록한다 (창이 뜨지 않는다).
등록 명령어와 분 단위 시각 배치는 [`create_schtasks_v2.txt`](create_schtasks_v2.txt).

## 의존성

```bash
pip install pywin32 openpyxl
```

`win32com` 을 쓰므로 **Outlook 데스크톱 앱과 Excel 이 설치된 Windows** 에서만 돈다.
(상위 폴더의 `summarize_msgs.py` 는 `extract-msg` 기반이라 Outlook 없이도 돈다.)

## 관련

일정 자동화의 앞 세대 도구 — 제목 기준 첨부 감지 `check_mail_attachment.py` 와
요약 정제 단계가 없는 `update_schedule.py` — 는
[`data-preprocessing/260324_schedule`](https://github.com/jjonghyunn/data-preprocessing/tree/main/260324_schedule)
에 있다. 워크북 시트 구조 설명(`26_Schedule_separate(Auto).md`)과 긴 경로 설정 안내도 그쪽에 있다.
