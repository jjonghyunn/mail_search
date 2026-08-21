# update_schedule_summary.py  
<sub>2026-08-21  Jonghyun Park w/ Claude</sub>  

`update_schedule.py` 의 사본 + **Summary 시트 자동 정제 단계**를 앞에 붙인 버전.
같은 폴더의 `README.md`(= `update_schedule.py` 문서)와 겹치는 부분(최신 파일 선택 규칙, Auto 파일 하류 체인,
작업 스케줄러 등록)은 그쪽을 보고, 여기서는 **달라진 부분만** 다룬다.

## 왜 만들었나

메일로 오는 고객 법인 일정 파일의 포맷이 바뀌어 이제 **`Summary` 시트 하나만** 온다.
예전엔 첨부의 첫 시트가 이미 `일정`(B~N 13열) 형태라 `update_schedule.py` 가 그대로 읽어 붙여넣을 수 있었는데,
지금은 사람이 손으로 `일정` 시트를 만들어야(`=Summary!B7` 수식 + 기간 `6/24~9/13` 을 날짜 2개로 쪼개기)
`update_schedule.py` 가 돌아간다. **그 수동 정제 단계를 스크립트가 대신한다.**

## 동작 순서

1. `1.고객 법인 일정 파일/` 폴더에서 최신 파일 자동 선택 — **`.xlsx`·`.xlsb` 모두 후보, 도착순 단일 축**
   - `.xlsb` 는 openpyxl 이 못 읽으므로 Excel COM 으로 `.xlsx` 변환본을 떠서 읽는다 (`ensure_openpyxl_readable()`).
     변환본은 **소스 폴더 직속에 원본과 같은 이름**(`<원본>.xlsx`)으로 남는다.
     변환본이 소스보다 새것이면 재변환하지 않는다 (매 실행 Excel 을 띄우지 않기 위함).
2. `Summary` 시트 정제 → 일정 13열(B~N) 데이터 생성 (**메모리 처리 — 소스 파일 미변경**)
3. **[2026-08-19 변경] `target_is_saved()` 로 Auto 파일에 그 결과가 실제로 저장돼 있는지 확인**
   → 이미 반영돼 있으면 SKIP (마커가 아니라 **파일 내용**으로 판정 — 아래 `## 재실행 판정` 참조)
4. `일정` 시트 생성 (`WRITE_SHEET_TO_SOURCE=True` 이고 **그 시트가 없을 때만**)
   - **xlsx 소스** → 소스 파일에 기록. 종전대로 `[SKIP]` 뒤 단계라 **Auto 갱신이 필요할 때만** 돈다.
   - **xlsb 소스** → 원본엔 못 쓰므로 **변환본에** 기록하고, 이 단계만 `[SKIP]` 판정 **앞**에 둔다
     → **Auto 갱신이 SKIP 돼도 변환본과 `일정` 시트는 항상 만들어진다.**
     고객이 xlsb 만 보내는 회차엔 폴더에 사람이 열어볼 xlsx 도 `일정` 시트도 없기 때문 (2026-08-20).
   - 2026-08-20: 이미 있으면 **지우지 않고 그대로 둔다** — 수기 편집·고객 원본 보존.
     다시 만들려면 소스에서 `일정` 시트를 지우고 재실행.
   - ⚠ 이 단계는 Auto 파일 **갱신이 필요할 때만** 실행된다. 이미 반영돼 있으면
     그 앞의 `[SKIP]` 에서 종료되므로 시트가 없어도 만들어지지 않는다.
5. 직전 소스 파일도 같은 정제 → 전후 비교(노란 음영)용
6. Auto 파일 `고객법인일정파일` 시트 **B2:N999 클리어 후 B5 부터** 붙여넣기
7. Excel COM 으로 `CalculateFull()` → 저장 → **저장 직후 `target_is_saved()` 재검증** → 통과했을 때만 마커 기록
   - 그 앞에 `is_locked()` 로 Auto 파일 잠김을 먼저 본다 — 잠겨 있으면 **Excel 을 아예 안 띄우고** 물러난다
     (아래 'Excel 팝업 차단' 참조). 이때 Auto 파일은 수식 캐시가 빈 상태라 상세 사유에 그 사실을 남긴다.
   - 재검증은 **읽기 실패(잠김) 사유일 때만** `VERIFY_RETRIES`(3회 × 2초) 재시도한다.
     Excel 이 막 놓은 파일을 곧바로 읽으면 `PermissionError` 가 나 '저장이 되돌려짐' 으로 오판된다
     (2026-08-20 실제 발생). 값 불일치·잔재는 재시도해도 그대로라 즉시 판정.
8. 실행 결과를 소스 폴더의 `_schedule_update_status.txt` 에 기록 — **파일 1개**,
   상단 `── 마지막 실행` 블록 + 하단 `── 실행 이력` **누적**
   - 스케줄러로 돌면 콘솔이 아무데도 안 남아, 갱신됐는지 SKIP 인지 실패인지 확인할 방법이
     Auto 파일을 직접 열어보는 것뿐이었다. 그 확인을 파일 하나로 대신한다.
   - 기록 항목: 실행 시각 / 결과(`갱신 완료`·`이미 반영됨 (SKIP)`·`실패`·`실패(예외)`·`경고`) / 상세 사유 /
     소스 파일 / 변환본(+`일정 시트 O|X`) / Auto 파일 / 데이터 행수
   - **이력 (2026-08-21 추가)**: 매 실행 `시각 / 결과 / 상세·행수` 한 줄을 하단에 append,
     `STATUS_LOG_MAX_LINES`(200줄) 초과분은 오래된 것부터 잘린다. 팝업 없이 조용히 물러난 회차가
     쌓이므로 **'언제부터 실패했는지'** 를 여기서 본다. 구분선 문자열(`STATUS_LOG_HEADER`)이
     바뀌면 그 시점에 이력이 끊긴다.
   - `.txt` 라 `SOURCE_EXTS` 에 안 걸려 소스 후보를 오염시키지 않는다.

## 정제 룰 (Summary → 일정)

열은 **헤더 문자열로 찾는다** (열 문자 하드코딩 X) — 고객이 열을 끼워넣어도 견디도록.

| 출력 열 | 소스 (Summary) | 규칙 |
|---|---|---|
| B Global | `Global` 헤더가 있는 열 (`Region`/`Global` 라벨 열) | 값 그대로. **그룹 시작행에만 값**, 나머지는 빈칸 |
| — | `No.` | **건너뜀** |
| C Subs | 헤더가 **정확히** `Subs` 인 열 | 라벨행의 `Subs.` 는 매칭 안 됨 |
| D Country | `Country` | 값 그대로 (옛 소스의 `PIC` 자리) |
| E Participation | 파생 | B2B/B2C 날짜가 **1개라도 파싱되면 `O`**, 아니면 빈칸 |
| F / H | `캠페인 기간(B2B)` | `~` 로 분리 → 각각 `M/D` 파싱 → `date` |
| G / I | — | **공백** (원래 WEEKNUM 자리) |
| J / L | `캠페인 기간(B2C)` | 위와 동일 |
| K / M | — | **공백** |
| N note | `Remark` | 값 그대로 |

**날짜 파싱**
- 구분자 `~` (전각 `～`/`∼`/`〜` 도 허용). **`~` 가 없으면 시작·종료 모두 빈칸** → `TBU`, `-`, `b2`(오타), 빈칸이 전부 여기로 흡수
- 토큰은 `M/D` (`.` `-` 구분자도 허용). 연도는 상수 `CAMPAIGN_YEAR`
- 종료일 < 시작일이면 **종료일만 다음 해로** (`12/20~1/15` → 2026-12-20 ~ 2027-01-15)

**데이터 행 판정**: 헤더행 다음부터 `Subs` 또는 `Country` 중 하나라도 값이 있으면 데이터 행.
(Korea 행처럼 기간이 전부 빈 행도 포함 — 수동 작업본과 동일)

## 생성되는 `일정` 시트

| 행 | 내용 |
|---|---|
| 6 | Region 라벨행 (B열에 `Region`) |
| 7 | 헤더행 — `Global / Subs / Country / (공백) / 캠페인 기간(B2B) / … / 캠페인 기간(B2C) / … / note` |
| 8~ | 데이터 |

- 시트 위치는 **index 0**(맨 앞). 이미 `일정` 시트가 있으면 삭제 후 재생성 (수동 작업본 덮어쓰기).
- 6/7행 헤더 + 8행~ 데이터 레이아웃은 **수동 작업본을 그대로 재현**한 것.

### → Auto `고객법인일정파일` 로의 행 매핑 (⚠ 2026-08-07 수정)

| 일정 시트 | → Auto | 내용 |
|---|---|---|
| 6행 | **5행** | Region 라벨행 |
| 7행 | **6행** | 헤더행 |
| 8행~ | **7행~** | 데이터 (55행 → 7~61행) |

`TGT_START_ROW = 5`. **B2 가 아니라 B5 부터** 들어가야 한다 (Auto 파일이 그렇게 짜여 있음).
클리어는 잔재 제거를 위해 한 칸 위인 `TGT_CLEAR_ROW = 2` 부터 (B2:N999).

## 재실행 판정 (⚠ 2026-08-19 변경)

**SKIP 여부를 마커가 아니라 `Auto 파일에 실제로 저장된 내용`으로 판정한다.**

### 왜 바꿨나

종전엔 마커(`campaign_schedule_last_source.txt`)가 최신이면 무조건 SKIP 했다. 그런데 마커는
`recalc_and_save()` 성공 **직후**에 쓰이므로, 그 뒤 Auto 파일이 외부 요인으로 되돌려지면
**마커만 최신이고 내용은 옛 소스인 상태**로 굳어 스케줄러가 영원히 SKIP 한다.

실제 발생 (2026-08-19):

| 대상 | 값 |
|---|---|
| 마커 | `…20260819_v1(…).xlsx` (최신) |
| Auto `고객법인일정파일!D1` | `…20260811_v2.xlsx` ← **한 세대 전** |
| Auto 데이터 | 7~61행 (55행, 옛 값) / 실제 필요 = 5~70행 (64행) |
| mtime | 소스 15:51 → 마커 15:54 → **Auto 15:55** (우리 저장 *이후* 외부가 덮어씀) |

### `target_is_saved(output_file, source_name, src_data)`

`(True, "")` = 이미 반영됨 → 실행 불필요 / `(False, 사유)` = 미반영·유실 → 실행 필요.
아래를 순서대로 보고 하나라도 걸리면 즉시 `False` + **사유를 로그에 찍는다**.

| # | 검사 | 잡아내는 것 |
|---|---|---|
| 1 | `load_workbook(..., data_only=True, read_only=True)` 성공 여부 | 파일 사용 중·손상 |
| 2 | `TARGET_SHEET` 시트 존재 | 시트 삭제 |
| 3 | **`D1` == 소스 파일명** (`STAMP_COL`) | 다른 소스가 들어가 있음 / 되돌려짐 |
| 4 | 붙여넣기 영역(B5~ × 13열) 값 1:1 대조 | 부분 저장·값 변조 (첫 불일치 셀 좌표를 사유에 표기) |
| 5 | 영역 밖 잔재 행 | 행 수가 줄었는데 옛 행이 남음 |

- 값 비교는 `_norm()` 을 거친다 — openpyxl 은 `date` 로 쓴 값을 **`datetime` 으로 되읽고**,
  빈 문자열과 빈칸도 구분하므로 그대로 비교하면 매번 불일치가 난다.
- 읽기는 `read_only=True` + `iter_rows(values_only=True)` **한 번**으로 끝낸다
  (실측 로드 0.13s + 순회 0.03s — 20분 주기 실행에 부담 없음. `read_only` 에서 `ws.cell()` 랜덤 접근은 느리다).

### 마커의 역할 (기록·경고 전용)

- **판정에는 안 쓴다.** 다만 `마커는 최신인데 target_is_saved() 가 False` 면
  `[경고] 마커는 '처리 완료'인데 Auto 파일 내용은 최신이 아닙니다` 를 찍어 **저장 유실을 드러낸다.**
- **마커는 저장 직후 재검증까지 통과했을 때만 기록**한다. 검증 실패 시 기록하지 않지만,
  어차피 다음 실행이 내용 기준으로 재시도하므로 무한 SKIP 은 발생하지 않는다.
- SKIP 인데 마커만 안 맞으면 마커를 **뒤늦게 동기화**한다 (자기치유).
- 이 정책 덕에 마커는 `update_schedule.py` 와 **같은 파일(`campaign_schedule_last_source.txt`)을 계속 공유**한다.

### mtime 복원 (종전과 동일)

소스 파일에 `일정` 시트를 쓰면 mtime 이 바뀌어 마커 값(`파일명|mtime`)이 흔들린다.
→ 저장 직후 `os.utime()` 으로 **원래 mtime 을 복원**해서 마커 의미(= 메일로 받은 버전)를 유지한다.

## 전후 비교(노란 음영)

`COMPARE` 대상 열(E Participation / F·H B2B / J·L B2C)은 종전과 같지만, **비교 소스가 바뀌었다.**

- 종전: 직전 파일의 `일정` 시트를 읽음 → 이제 옛 파일엔 그 시트가 없어 못 씀
- 현재: **직전 파일의 `Summary` 도 똑같이 정제**해서 비교 → 포맷 혼재에 안 흔들림
- 매칭 키: `(Subs, Country, 같은 조합의 몇 번째)`.
  `SUB_C`(북유럽 4국)처럼 같은 Subs 가 여러 행인 경우 종전 방식(Subs 단독 키)은 **마지막 행만 남아** 비교가 어긋났다.

## 상수 (파일 상단)

| 상수 | 기본값 | 설명 |
|---|---|---|
| `BASE` / `SOURCE_FOLDER` / `TARGET_SHEET` / `LAST_SOURCE_FILE` | CAMPAIGN NAME 캠페인 경로 | `update_schedule.py` 와 동일 |
| `SOURCE_EXTS` | `(".xlsx", ".xlsb")` | 소스로 인정할 확장자. 고객이 회차마다 오락가락 보낸다 |
| `SOURCE_NAME_KEYS` | `["schedule", "캠페인", "일정", "monitoring"]` | 파일명에 하나라도 있어야 소스 후보 (아래 참조) |
| `XLSB_WORK_DIR` | `%TEMP%\campaign_schedule_xlsb_work` | Excel 변환 작업용 **짧은 경로** 임시 폴더 (아래 `MAX_PATH` 참조). 매번 비운다 |
| `STATUS_FILE` | `<소스 폴더>\_schedule_update_status.txt` | 마지막 실행 블록 + 실행 이력(누적) — 파일 1개 |
| `STATUS_LOG_MAX_LINES` | `200` | 상태 txt 하단 '실행 이력' 보관 줄 수 |
| `VERIFY_RETRIES` / `VERIFY_WAIT_SEC` | `3` / `2` | 저장 직후 재검증 재시도 (**잠김 사유일 때만**) |
| `XL_SECURITY_FORCE_DISABLE` / `XL_FEATURE_INSTALL_NONE` | `3` / `0` | Excel 팝업 차단용 (`_new_excel()`) |
| `CAMPAIGN_YEAR` | `2026` | 기간 `M/D` 에 붙일 연도. **캠페인 해가 바뀌면 여기만 수정** |
| `SUMMARY_SHEET` | `"Summary"` | 없으면 첫 번째 시트로 fallback |
| `SCHEDULE_SHEET` | `"일정"` | 생성할 정제 시트명 |
| `WRITE_SHEET_TO_SOURCE` | `True` | `일정` 시트가 **없을 때만** 생성. 있으면 손대지 않음. `False` 면 메모리 처리만 (소스 파일 미변경) |
| `H_GLOBAL` / `H_SUBS` / `H_COUNTRY` / `H_B2B` / `H_B2C` / `H_REMARK` | Summary 헤더 문자열 | 고객이 헤더 문구를 바꾸면 여기 수정 |
| `SCHED_LABEL_ROW` / `SCHED_HEADER_ROW` | `6` / `7` | 생성 시트 레이아웃 |
| `SRC_MIN_COL` / `SRC_MAX_COL` | `2` / `14` | 읽기·붙여넣기 열 범위 (B~N) |
| `TGT_CLEAR_ROW` | `2` | 클리어 시작 행 (B2:N999) |
| `TGT_START_ROW` | **`5`** | 붙여넣기 시작 행 = Region 라벨행 |
| `TGT_MAX_ROW`, `COMPARE` | 종전과 동일 | 클리어 하단·음영 대상 |
| `STAMP_COL` | `4` | D열 — 소스 파일명을 기록·대조하는 열(`D1`). **재실행 판정의 1차 키** |

## ⚠ Auto 파일 하류 동작 변화 (수동 작업본 대비)

수동 작업본은 `=Summary!B9` 수식으로 Global 열을 채웠는데, **Summary 쪽이 빈 셀이면 수식 결과가 `0`** 이 된다.
이 스크립트는 값으로 쓰므로 그 자리가 **진짜 빈칸**이다. 이 차이가 하류에서 다음을 **고친다**:

- `태깅기획site_code!N4 = FILTER(고객법인일정파일!B4:B99, …<>"")` (Region 순서용 매핑)
  - 붙여넣기가 5행부터라 B5(`Region`)·B6(`Global`) 라벨이 FILTER 에 같이 잡힌다 →
    `Region(1), Global(2), N.America(3), EU(4), L.America(5), S.E.Asia(6), S.W.Asia(7), MENA(8), CIS(9), China(10), Korea(11)`
  - 앞 2개는 더미지만 **Region 간 상대 순서가 그대로**라 `Appendix` 정렬 결과는 정상 (수기 정렬본과 diff 0 확인)
  - 이전(수동 작업본)엔 `0` 이 섞여 `Region, Global, N.America, 0, 0, EU, …` 로 오염됐고,
    그 `0` 때문에 상당수 Region 이 `#N/A → ""` 가 되어 Appendix 그룹이 섞였다
- `MASTER!Z`,`AP` = `VLOOKUP(Region, 태깅기획site_code!$N$4:$O$18, 2, 0)` → **Appendix 정렬 인덱스**. 위 개선이 그대로 반영됨
- `MASTER!Q`(고객Note): 문자열 `'0'` → 빈칸

> **참고**: FILTER 결과 행 수가 줄면 `태깅기획site_code!N`/`O` 아래쪽에 `#N/A` 가 남을 수 있다.
> openpyxl 라운드트립이 배열 수식 `ref` 를 고정 범위로 저장해 Excel 이 남는 칸을 `#N/A` 로 채우는 것.
> 이 범위를 쓰는 건 `IFERROR(VLOOKUP(...))` 뿐이라 **결과에는 영향 없음**. 거슬리면 Excel 에서 `N4` 수식을
> 지웠다 다시 입력해 동적 배열로 되돌리면 사라진다.

## Excel 팝업 차단 (2026-08-21)

무인(`pythonw`) 실행인데 **실패 시 Excel 대화상자가 화면에 떴다.** 모달 창이 뜨면 사람이 닫아줄 때까지
그 Excel 인스턴스가 멈춰 서서 다음 회차까지 물리고, 고아 `EXCEL.EXE` 가 쌓인다.

`DisplayAlerts = False` 만으로는 부족하다 — 그건 '저장/덮어쓰기 확인' 류만 막고
**링크 업데이트 · 읽기전용 권장 · 매크로 보안 · 기능 설치 · '파일 사용 중'** 은 그대로 뜬다.
그래서 4겹으로 막는다:

| 층 | 무엇 | 막는 것 |
|---|---|---|
| ① `_new_excel()` | `AskToUpdateLinks` / `AlertBeforeOverwriting` / `EnableEvents` / `AutomationSecurity` / `FeatureInstall` 까지 전부 끈 전용 인스턴스 | 링크 갱신·매크로 보안·기능 설치·Open 이벤트 창 |
| ② `Workbooks.Open(..., Notify=False, IgnoreReadOnlyRecommended=True, UpdateLinks=0)` | 잠긴 파일에 **알림 창 대신 `com_error`** | '파일 사용 중' 알림, '읽기 전용으로 여시겠습니까' |
| ③ `is_locked()` 선확인 | 잠겨 있으면 **Excel 을 아예 안 띄운다** | 위 둘을 통과해도 남는 창·고아 프로세스 |
| ④ `sys.excepthook` | 남은 예외를 받아 상태 txt 에 `실패(예외)` 로 기록만 하고 종료 | traceback 이 어디에도 안 남는 문제 |

- ②·③은 겹치는 방어다 — Excel 은 파일을 **deny-write** 로 잠그므로 `Notify=False` 면 창 없이
  읽기 전용으로 열리고 `Save()` 단계에서 `com_error` 가 난다. ③이 그 앞에서 걸러 Excel 기동 자체를 없앤다.
- ④의 `_log_uncaught` 는 콘솔이 있으면(`sys.stderr is not None`) traceback 도 같이 찍는다 —
  기록만 남기되 사람이 직접 돌릴 땐 원인이 보여야 한다.
- **콘솔 인코딩 방어도 같이 넣었다** — 진행 로그의 `—`(em dash) 한 글자에 `cp949` 콘솔에서
  `UnicodeEncodeError` 가 나 실행 전체가 죽는 걸 실측했다(2026-08-21). 시작 시
  `sys.stdout.reconfigure(encoding="utf-8", errors="replace")` 로 막는다(`pythonw` 면 `stdout` 이 `None`).
- 스케줄러 `raise`(Excel 재계산 최종 실패)는 **그대로 뒀다** — 작업 스케줄러에 실패 코드를 남기는
  유일한 신호이고, 이제 ④가 받아 창 없이 로그로 마무리한다.

## 알려진 제약

- **제목·형식과 무관하게 '나중에 도착한 파일'이 이긴다** (2026-08-20, `latest_file_key` v2.0).
  정렬 키가 `(도착시각, 문서날짜, 문서시각, 버전float, 버전int, 끝번호)` 라 **도착시각이 1순위**다.
  도착시각은 파일명 끝의 수신일시 스탬프(`_YYMMDD_HHMM`)에서 읽고, 스탬프가 없는 옛 파일은
  파일명 안의 문서날짜를 그 날 00:00 도착으로 환산한다(그래야 스탬프 있는 옛 파일이 최신을
  이기는 회귀가 없다). 이름에 날짜가 아예 없으면 그때만 mtime 을 쓴다.
  `_shared` 처럼 버전 숫자가 없는 접미사, `.xlsb`/`.xlsx` 차이, SW형(8자리)/MD형(6자리) 날짜 —
  **어느 것도 순서에 영향을 주지 않는다.**
  ⚠ 바뀐 점: 문서날짜가 옛것이어도 나중에 도착하면 이긴다(고객이 옛 파일을 재전송하면 그게 소스).
  배경·검증은 `README.md` 의 '최신 파일 선택 기준' 참조.
- **소스 후보는 확장자 + 이름으로 거른다.** `SOURCE_EXTS` 확장자이면서 `SOURCE_NAME_KEYS` 중
  하나가 파일명에 있어야 후보가 된다. 그 안에서는 오직 도착순으로 최신을 고른다.
- **⚠ 2026-08-20 전제 변경 — Monitoring 파일이 이제 '일정 소스'다.**
  종전엔 `SOURCE_NAME_KEYS` 3개(`schedule`/`캠페인`/`일정`) + `glob("*.xlsx")` 조합이
  **Qualitative Monitoring 첨부를 소스에서 배제**하는 장치였다 (같은 폴더에 `check_mail_attachment_byname.py` 가
  Monitoring 도 저장하므로). 그런데 고객이 일정 내용을 Monitoring 파일로 보내기 시작하면서
  그 두 장치가 정반대로 작동했다 — `26 CAMPAIGN NAME Qualitative Monitoring_260819_shared.xlsb`(20:08, 최신)가
  **이름에서도 확장자에서도 탈락**해, 한 세대 전인
  `..._20260819_v1(Qualitative Monitoring기반제작).xlsx`(00:00)가 소스로 잡히고 `[SKIP]` 으로 끝났다.
  → `SOURCE_EXTS` 에 `.xlsb` 추가 + `SOURCE_NAME_KEYS` 에 `monitoring` 추가.
  이제 이 상수는 **'무관한 파일 배제' 용 화이트리스트**이지 Monitoring 배제 장치가 아니다.
  (Monitoring xlsb 는 시트가 `Summary` 하나뿐이고 — A1 제목 셀이
  `2026 CAMPAIGN NAME Campaign Summary` 라 시트명으로 오인하기 쉽다 —
  `Global`/`Subs`/`Country`/`캠페인 기간(B2B)`/`캠페인 기간(B2C)`/`Remark` 헤더가 그대로 있어
  정제 로직은 무수정으로 먹힌다. Monitoring 쪽 추가 열 — `B2B 모니터링`/`홈KV`/`L0메뉴` 등 — 은
  헤더 이름으로 열을 찾는 구조라 자동 무시된다.)
- **변환본은 소스 폴더 직속, 원본과 같은 이름**(`260819_2008_..._shared.xlsb` → `..._shared.xlsx`).
  같은 이름의 `.xlsb` 가 있는 `.xlsx` 는 `is_converted_twin()` 이 **소스 후보에서 제외**한다.
  이 룰이 없으면 다음 실행이 변환본을 소스로 집는다 — 도착시각이 같고 ext tiebreak 이 xlsx 우선이라
  확정적으로 뒤집힌다. (고객이 같은 회차에 진짜 xlsx·xlsb 를 둘 다 보내도 내용이 같아 무해.)
- **⚠ Excel 에는 짧은 임시 경로만 넘긴다 — 복사 → 변환 → 되옮기기 3단계인 이유.**
  소스도 변환본도 전체 경로가 **264자**로 `MAX_PATH`(260)를 넘어, 원본 경로를 그대로 넘기면 Excel 이
  `Workbooks 클래스 중 Open 메서드에 오류가 있습니다`(0x800A03EC)로 거부한다.
  Python 은 같은 경로를 **읽고 쓰는 데** 문제가 없어서(롱패스) 원인이 잘 안 보인다 — 인자 조합·파일 손상
  문제로 오인하기 쉽다 (2026-08-20 실측). Excel 공식 한도는 경로+파일명 **218자**라 Auto 파일(223자)도
  이미 아슬아슬하다 — 캠페인 폴더명이 한 번만 더 길어지면 `recalc_and_save()` 쪽도 같은 이유로 깨진다.
- 소스 파일을 openpyxl 로 열었다 저장하므로 **`Summary!J5`/`K5`(COUNTIF) 캐시값이 지워진다.**
  소스 파일을 Excel 로 열면 자동 재계산돼 원래 값이 돌아온다 (이 스크립트는 그 셀을 안 읽음).
  캐시가 꼭 필요하면 `WRITE_SHEET_TO_SOURCE=False` 로 두고 소스를 안 건드리면 된다.
- 소스 파일이 Excel 에서 열려 있으면 `일정` 시트 기록만 건너뛰고 **Auto 업데이트는 계속 진행**한다.
- Excel COM `CalculateFull()+Save()` 단계는 **제거 금지** — 빼면 수식 캐시가 비어 후속 도구가 전부 `None` 을 본다.
- **`wb_com.Close()` 의 `TypeError: 'bool' object is not callable`** (2026-08-11, Python 3.14 로 실행 시 발생) —
  win32com 은 `%TEMP%\gen_py\<파이썬버전>\` 에 타입라이브러리 캐시가 있어야 조기 바인딩을 한다.
  파이썬을 새로 깔면 그 캐시가 비어 **지연 바인딩**으로 떨어지고, 이때 `wb_com.Close` 는
  **속성 접근만으로 COM 메서드가 실행**되어 `True` 를 돌려준다 → 이어지는 `()` 가 TypeError.
  이미 저장·닫기는 끝난 뒤라 **출력 xlsx 는 정상**이지만, 그 다음 줄인 마커(`LAST_SOURCE_FILE`) 기록이
  안 돼서 다음 실행이 같은 소스를 재처리한다(작업 스케줄러엔 실패로 기록).
  → 현재 코드는 `Close(SaveChanges=False)` 를 `try/except TypeError` 로 감싸 양쪽 바인딩 모두에서 통과한다.
  (`%TEMP%` 캐시는 디스크 정리로 언제든 지워지므로 gen_py 재생성이 아니라 코드로 막는 게 맞다.)
- **⚠ Auto 파일이 외부에 의해 옛 내용으로 되돌려지는 사례 (2026-08-19)** —
  스크립트가 정상 저장하고 `[완료]` 까지 찍은 뒤에도, 몇 분 지나 Auto 파일이 **한 세대 전 소스 상태로 롤백**되는
  현상이 두 번 관측됐다 (저장 직후 재검증은 통과 → 이후 롤백). 파일은 클라우드 플레이스홀더가 아니라
  로컬 고정(`attrib` 의 `P`) 상태이고 Excel 도 안 떠 있었으므로, **OneDrive 동기화(공유 폴더라 서버 버전이
  내려오는 경우) 또는 다른 사람이 같은 파일을 열어 둔 co-authoring** 이 유력하다.
  → 이제 재실행 판정이 내용 기준이라 **다음 실행(최대 20분 뒤)이 자동으로 다시 병합**한다. 스크립트는 스스로 복구되지만,
    근본 원인은 OneDrive 쪽이므로 자주 반복되면 동기화 상태·공동 편집자를 확인할 것.
  → **스케줄러 실행 시간대에 Auto 파일을 Excel 로 열어두지 말 것** (열어둔 Excel 이 옛 메모리 내용으로 덮어쓸 수 있다).
- **OneDrive 동기화 중 파일 잠김 (2026-08-19 보강)** — 같은 날 실행 중 Auto 파일과 직전 소스 파일이 동시에
  `PermissionError` 로 잠긴 순간이 관측됐다. 20분 주기 스케줄러라 **잠김은 그냥 다음 실행에 넘기면 되는 상황**인데,
  종전 코드는 두 군데서 traceback 으로 죽었다:
  | 위치 | 종전 | 현재 |
  |---|---|---|
  | 소스 파일 정제 (`build_schedule_rows(source_file)`) | 가드 없음 → traceback | `OSError` 잡아 `[SKIP] 소스 파일을 읽을 수 없습니다` 후 `exit(0)` |
  | 직전 파일 정제 (전후 비교용) | `except (ValueError, KeyError)` 로 좁아 `PermissionError` 통과 → traceback | `except Exception` — **부가 기능이므로 비교만 생략** |
  Auto 파일 자체를 못 읽는 경우는 `target_is_saved()` 가 `(False, "Auto 파일을 읽을 수 없음 …")` 을 돌려
  '미반영' 으로 보고 그대로 진행한다 (뒤쪽 쓰기 단계에 이미 `PermissionError` → `[SKIP]` 처리가 있다).

## 실행 / 스케줄러

```powershell
python "C:\Users\user_name\OneDrive - company_name\user_id\work\campaign_schedule\update_schedule_summary.py"
```

**2026-08-07 — 작업 스케줄러 `campaign_schedule_update` 를 이 스크립트로 교체 완료.**
(20분 주기 / 10:04 시작 / `/ed 2026/10/15` / 배터리 조건 둘 다 `False` — 트리거·설정은 그대로 두고 action 만 교체)

```powershell
# action 만 바꾸는 방식 — 트리거/설정 보존 (재등록보다 안전)
$py     = '"C:\Python314\pythonw.exe"'
$script = '"C:\Users\user_name\OneDrive - company_name\user_id\work\campaign_schedule\update_schedule_summary.py"'
Set-ScheduledTask -TaskName 'campaign_schedule_update' -Action (New-ScheduledTaskAction -Execute $py -Argument $script)

# Set-ScheduledTask 후 배터리 설정이 되돌아갈 수 있으므로 다시 적용
$t = Get-ScheduledTask -TaskName 'campaign_schedule_update'
$t.Settings.DisallowStartIfOnBatteries = $false
$t.Settings.StopIfGoingOnBatteries     = $false
Set-ScheduledTask -InputObject $t
```

> 통째로 재등록하려면 `create_schtasks_v2.txt` 의 `campaign_schedule_update` 줄(이미 새 파일명으로 갱신됨)을 쓰고,
> 그 뒤 배터리 모드 허용 설정(README `### 배터리 모드 허용`)을 다시 적용할 것.
> 구 `update_schedule.py` 는 롤백용으로 폴더에 그대로 남겨둔다 (스케줄러에는 안 걸림).

## 검증 기록 (2026-08-07)

| 항목 | 결과 |
|---|---|
| 정제 결과 vs 수동 작업본 `일정` 시트 (55행 × 13열) | **diff 0** |
| `Summary` 만 있는 원본 파일 종단 테스트 | `일정` 시트 생성 OK, mtime 보존 OK |
| 연속 2회 실행 | 2회차 `[SKIP] 소스 파일 변경 없음` |
| **빈 Auto 사본 복원 테스트** (아래) | 레이아웃·하류 전부 정상 |
| Auto 파일 새 오류(`#REF!`/`#VALUE!` 등) | 없음 |

### 빈 Auto 사본 복원 테스트 (2026-08-07)

`고객법인일정파일` 시트를 **B1:N999 전부 비운** Auto 사본 + Downloads 원본(`Summary` 기반) 을
격리 폴더(`Downloads\campaign_schedule_검증\`)에 놓고 실행:

| 확인 | 결과 |
|---|---|
| 붙여넣기 위치 | `B5=Region` / `B6=헤더` / `B7~61=데이터 55행` — **수기 정렬본과 동일** |
| `Appendix_Date` / `Appendix_URL` / `태깅기획site_code` | 수기 정렬본과 **diff 0** |
| `MASTER` | `Q`열 `'0' → 빈칸` 등 48건 (전부 위 "하류 동작 변화" 항목) |
| `api용` / `RAW_*` / `날짜세그*` / `Last고객법인일정파일` | **diff 0** |

## 검증 기록 (2026-08-19 — 재실행 판정 전환)

| 항목 | 결과 |
|---|---|
| 마커가 최신인데 Auto 내용이 옛 소스인 상태에서 실행 | `[경고] 마커는 '처리 완료'…` + `[갱신 필요] D1 소스명 불일치 …` 출력 후 **정상 병합** |
| 붙여넣기 결과 | `D1` = 0819 소스, `B5=Region` / `B6=헤더` / `B7~70 = 데이터 64행`, 71행 이하 잔재 없음 |
| 연속 실행 (3회차) | `[SKIP] Auto 파일에 이미 반영돼 있습니다` — Excel COM 미기동 |
| Auto 파일 오류값 | **신규 발생 없음.** `MASTER!Q` 등의 `#N/A` 는 `MASTER!D` 의 14개 법인(`SUB_A`/`SUB_B`/`HQ` …)이 고객 일정 파일에 애초에 없어서 나는 것으로, **0811·0819 소스에서 동일**(사라진 Subs 0건, 신규 `SUB_D-BE_FR` 1건) |
| 저장 유실 재현 | 1회차 `[완료]` 후 롤백 발생 → **2회차가 스스로 감지해 재병합** (위 '알려진 제약' 참조) |
| OneDrive 잠김 중 실행 | 가드 추가 전에는 `PermissionError` traceback → 가드 후 `[SKIP]`/`[알림]` 로 조용히 물러남 (exit 0) |
| 최종 상태 | 연속 2회 `[SKIP] Auto 파일에 이미 반영돼 있습니다` (exit 0), Excel COM 미기동 |


## 검증 기록 (2026-08-20 — 도착순 반영, `latest_file_key` v1.2)

**계기**: `26 CAMPAIGN NAME Qualitative Monitoring_260819_v1` 다음에 `..._260819_shared` 가 왔을 때
`_shared` 가 최신으로 인식되는지 확인 → **안 됐다.** 두 파일의 정렬 키가 `(260819, 0, 0.0, 0, 0)` 로
완전히 같아 동점이었고, 같은 패턴이 일정 파일에 오면 `_shared`(vN=0) 가 `_v1`(vN=1) 에 밀려
**나중에 온 파일이 무시**됐다. (상세 경위는 `README.md` v1.2 절)

| 검증 항목 | 결과 |
|---|---|
| 실제 소스 폴더(19개) 재정렬 — 회귀 없음 | 최신 = `2026 CAMPAIGN NAME Campaign Schedule_20260819_v1(Qualitative Monitoring기반제작).xlsx` (변경 전과 동일) |
| 스탬프 있는 옛 파일 vs 없는 최신 파일 | `_20260806_260807_0949`(2608070949) < `_20260819_v1`(2608190000) — 스탬프 없는 쪽을 문서날짜 00:00 으로 환산해 정상 |
| 도착시각 최우선 + 스탬프 없는 파일을 `0` 으로 둘 때 (기각안) | `_20260806_260807_0949` 가 `_20260819_v1` 을 이김 ✗ → 문서날짜 환산(v2.0)으로 해결 |
| 도착순 반영 (모니터링) — **실측** | `_260819_v1_260819_1548`(15:48) < `_260819_shared_260819_2008`(20:08) → `_shared` 최신 ✓ <br>2026-08-20 재수집으로 확인한 실제 수신시각 |
| 도착순 반영 (일정) — 가상 | `_20260819_v1_260819_1551` < `_20260819_shared_260819_1802` → `_shared` 최신 ✓ <br>(일정 파일엔 아직 이 패턴이 안 왔음 — 시뮬레이션)|
| 같은 메일 v1·v2 동봉 (수신일시 동일) | `ver_int` tiebreak 살아있어 v2 최신 ✓ |
| MD형(6자리 날짜 + `vX.XX`) 회귀 | `_v0.44_260319` / `_v0.48_260420` / `_v0.49_260420` 순서 종전과 동일 ✓ |
| **v2.0** 실제 소스 폴더(19개) 재정렬 | 선택 결과 변경 전과 동일 (`..._20260819_v1(Qualitative Monitoring기반제작).xlsx`) ✓ |
| **v2.0** SW/MD 자릿수 문제 | `sw_20260415_1543`(2604151543) < `md_v1.0_260416`(2604160000) — 6자리 정규화로 실제 날짜 비교 ✓ |
| **v2.0** 이름에 날짜 없음 → mtime | mtime 1시간 차 두 파일에서 늦은 쪽 선택 ✓ |
| **v2.0** 옛 문서날짜 + 늦은 도착 | `_20260818_shared_260825_1030` 가 `_20260819_v1_260819_1551` 을 이김 — **의도된 신규 동작** |

**같이 고친 것 (3번째 대응)**: 수신일시 꼬리를 **파일명 끝에서** 떼도록 변경.
종전 정규식은 문서날짜 *바로 뒤*에 붙은 꼬리만 읽어서 `..._20260819_shared_260819_1802` 처럼
사이에 토큰이 끼면 `mail = 0` 이 나왔다 (1차 수정만으로는 일정 파일 케이스가 안 고쳐졌음).

## 검증 기록 (2026-08-21 — Excel 팝업 차단 + 상태 txt 이력 누적)

| 검증 항목 | 결과 |
|---|---|
| 정상 실행 (갱신 필요 상태) | `[완료] … 저장 완료` — **`경고 — 저장이 되돌려짐` 안 나옴** (재검증 재시도가 흡수) ✓ |
| 연속 실행 2회 | 둘 다 `[SKIP] Auto 파일에 이미 반영돼 있습니다`, Excel COM 미기동 ✓ |
| 상태 txt | 상단 '마지막 실행' 블록은 최신 1회, 하단 이력은 4줄 누적 — **파일 1개** ✓ |
| `is_locked()` — Excel 방식(deny-write) 잠금 | `True` → Excel 기동 없이 `[SKIP]` ✓ |
| `Open(Notify=False)` — deny-write 잠금 | **대화상자 없이** 읽기 전용으로 열림 (`Save()` 에서 `com_error` → 재시도 루프가 처리) ✓ |
| `Open(Notify=False)` — 배타(공유 없음) 잠금 | **대화상자 없이** 즉시 `com_error` ✓ |
| `sys.excepthook` | `UnicodeEncodeError` 실사고에서 창 없이 상태 txt 에 `실패(예외)` + 사유 기록 ✓ |
| 콘솔 인코딩 (`cp949`) | 방어 전엔 `—` 출력에서 실행 전체가 죽음 → `reconfigure` 후 정상 ✓ |
