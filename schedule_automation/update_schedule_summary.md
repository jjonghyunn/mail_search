# update_schedule_summary.py  
<sub>2026-08-19  Jonghyun Park w/ Claude</sub>  

`update_schedule.py` 의 사본 + **Summary 시트 자동 정제 단계**를 앞에 붙인 버전.
최신 파일 선택 규칙·Auto 파일 하류 체인·작업 스케줄러 등록처럼 원본과 겹치는 설명은
[`data-preprocessing/260324_schedule/README.md`](https://github.com/jjonghyunn/data-preprocessing/tree/main/260324_schedule) 를 보고,
여기서는 **달라진 부분만** 다룬다.

## 왜 만들었나

메일로 오는 고객 법인 일정 파일의 포맷이 바뀌어 이제 **`Summary` 시트 하나만** 온다.
예전엔 첨부의 첫 시트가 이미 `일정`(B~N 13열) 형태라 `update_schedule.py` 가 그대로 읽어 붙여넣을 수 있었는데,
지금은 사람이 손으로 `일정` 시트를 만들어야(`=Summary!B7` 수식 + 기간 `6/24~9/13` 을 날짜 2개로 쪼개기)
`update_schedule.py` 가 돌아간다. **그 수동 정제 단계를 스크립트가 대신한다.**

## 동작 순서

1. `1.고객 법인 일정 파일/` 폴더에서 최신 파일 자동 선택 (`update_schedule.py` 와 동일 정렬 규칙)
2. 마커(`campaign_schedule_last_source.txt`)와 비교 → 파일명·mtime 동일하면 즉시 종료(SKIP)
3. **[신규] `Summary` 시트 정제 → 일정 13열(B~N) 데이터 생성**
4. **[신규] 소스 xlsx 에 `일정` 시트로 기록** (`WRITE_SHEET_TO_SOURCE=True` 일 때)
5. **[변경] 직전 소스 파일도 같은 정제** → 전후 비교(노란 음영)용
6. Auto 파일 `고객법인일정파일` 시트 **B2:N999 클리어 후 B5 부터** 붙여넣기
7. Excel COM 으로 `CalculateFull()` → 저장 → 마커 기록

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

## 마커 / mtime

소스 파일에 `일정` 시트를 쓰면 mtime 이 바뀌어 **다음 실행이 "새 파일"로 오인**한다.
→ 저장 직후 `os.utime()` 으로 **원래 mtime 을 복원**해서 마커 의미(= 메일로 받은 버전)를 유지한다.
덕분에 마커는 `update_schedule.py` 와 **같은 파일(`campaign_schedule_last_source.txt`)을 그대로 공유**한다.

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
| `CAMPAIGN_YEAR` | `2026` | 기간 `M/D` 에 붙일 연도. **캠페인 해가 바뀌면 여기만 수정** |
| `SUMMARY_SHEET` | `"Summary"` | 없으면 첫 번째 시트로 fallback |
| `SCHEDULE_SHEET` | `"일정"` | 생성할 정제 시트명 |
| `WRITE_SHEET_TO_SOURCE` | `True` | `False` 면 메모리 처리만 (소스 파일 미변경) |
| `H_GLOBAL` / `H_SUBS` / `H_COUNTRY` / `H_EPP` / `H_B2C` / `H_REMARK` | Summary 헤더 문자열 | 고객이 헤더 문구를 바꾸면 여기 수정 |
| `SCHED_LABEL_ROW` / `SCHED_HEADER_ROW` | `6` / `7` | 생성 시트 레이아웃 |
| `SRC_MIN_COL` / `SRC_MAX_COL` | `2` / `14` | 읽기·붙여넣기 열 범위 (B~N) |
| `TGT_CLEAR_ROW` | `2` | 클리어 시작 행 (B2:N999) |
| `TGT_START_ROW` | **`5`** | 붙여넣기 시작 행 = Region 라벨행 |
| `TGT_MAX_ROW`, `COMPARE` | 종전과 동일 | 클리어 하단·음영 대상 |

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

## 알려진 제약

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

## 실행 / 스케줄러

```powershell
python "C:\Users\user_name\OneDrive - company_name\user_id\your_folder\your_workspace\260324_schedule\update_schedule_summary.py"
```

**2026-08-07 — 작업 스케줄러 `cmp_schedule_update` 를 이 스크립트로 교체 완료.**
(20분 주기 / 10:04 시작 / `/ed 2026/10/15` / 배터리 조건 둘 다 `False` — 트리거·설정은 그대로 두고 action 만 교체)

```powershell
# action 만 바꾸는 방식 — 트리거/설정 보존 (재등록보다 안전)
$py     = '"C:\Python3xx\pythonw.exe"'
$script = '"C:\Users\user_name\OneDrive - company_name\user_id\your_folder\your_workspace\260324_schedule\update_schedule_summary.py"'
Set-ScheduledTask -TaskName 'cmp_schedule_update' -Action (New-ScheduledTaskAction -Execute $py -Argument $script)

# Set-ScheduledTask 후 배터리 설정이 되돌아갈 수 있으므로 다시 적용
$t = Get-ScheduledTask -TaskName 'cmp_schedule_update'
$t.Settings.DisallowStartIfOnBatteries = $false
$t.Settings.StopIfGoingOnBatteries     = $false
Set-ScheduledTask -InputObject $t
```

> 통째로 재등록하려면 `create_schtasks_v2.txt` 의 `cmp_schedule_update` 줄(이미 새 파일명으로 갱신됨)을 쓰고,
> 그 뒤 배터리 모드 허용 설정(`create_schtasks_v2.txt` 아래쪽 PowerShell 한 줄)을 다시 적용할 것.
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
