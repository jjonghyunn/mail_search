"""
update_schedule_summary.py   [CAMPAIGN NAME 폴더 전용 — update_schedule.py 의 정제 통합판]
2026-08-11  Jonghyun Park w/ Claude
2026-08-19  Jonghyun Park w/ Claude  — SKIP 판정을 마커 대신 Auto 파일 실제 저장내용으로 전환
2026-08-20  Jonghyun Park w/ Claude  — xlsb 소스 허용(Excel COM 변환) + Monitoring 파일을 일정 소스로 인정
2026-08-20  Jonghyun Park w/ Claude  — 변환본(.xlsx + 일정 시트)을 소스 폴더에 상시 생성 + 실행 상태 txt
2026-08-21  Jonghyun Park w/ Claude  — Excel 팝업 전면 차단(Notify/AskToUpdateLinks 등) + 상태 txt 이력 누적

update_schedule.py 와의 차이 = **Summary 시트 자동 정제 단계가 앞에 붙었다**.

배경: 메일로 오는 고객 법인 일정 파일의 포맷이 바뀌어 이제 `Summary` 시트 하나만 온다.
      예전엔 첨부의 첫 시트가 이미 `일정`(B~N 13열) 형태라 그대로 붙여넣을 수 있었지만,
      지금은 사람이 손으로 `일정` 시트를 만들어야(기간 `6/24~9/13` 을 날짜 2개로 쪼개기 등)
      update_schedule.py 가 돌아간다. 그 수동 단계를 이 스크립트가 대신한다.

흐름:
 1. `1.고객 법인 일정 파일/` 폴더에서 최신 파일 자동 선택 (xlsx·xlsb 모두 후보, 도착순 단일 축)
    — xlsb 는 openpyxl 이 못 읽으므로 Excel COM 으로 **같은 폴더에 같은 이름의 xlsx 변환본**을 떠서 읽는다
 2. Summary 시트 정제 → 일정 13열 데이터 생성 (**메모리 처리 — 소스 파일 미변경**)
 3. **Auto 파일에 그 결과가 실제로 저장돼 있는지 확인** → 이미 반영돼 있으면 SKIP
    (마커/로그가 아니라 파일 내용으로 판정 — 아래 '재실행 판정' 참조)
 4. `일정` 시트 기록 (WRITE_SHEET_TO_SOURCE=True 일 때)
    — xlsx 소스: 소스 파일에 기록 (Auto 갱신이 필요할 때만)
    — xlsb 소스: **변환본에 기록하고, Auto 갱신이 SKIP 돼도 항상 만든다**
      (고객이 xlsb 만 보내는 회차에 사람이 열어볼 xlsx·일정 시트가 폴더에 없어서)
 5. 직전 소스 파일도 같은 정제 → 전후 비교(노란 음영)용
 6. Auto 파일 `고객법인일정파일` 시트 B2:N999 클리어 후 **B5 부터** 붙여넣기
    (B5 = Region 라벨행, B6 = 헤더행, B7~ = 데이터)
 7. Excel COM 으로 전체 재계산 후 저장 → **저장 직후 재검증** → 통과했을 때만 마커 기록
    (Auto 파일이 잠겨 있으면 Excel 을 아예 안 띄우고 물러난다 — 아래 '팝업 차단' 참조)
 8. 실행 결과를 소스 폴더의 `_schedule_update_status.txt` 에 기록
    (상단 = 마지막 실행 블록, 하단 = 실행 이력 누적. 파일은 계속 **1개**)

팝업 차단 (2026-08-21):
  무인(pythonw) 실행이라 Excel 모달 창이 뜨면 사람이 닫을 때까지 멈춰 서서 다음 회차까지 물린다.
  실제로 실패 시 Excel 대화상자가 화면에 떴다. `DisplayAlerts=False` 만으로는 '저장 확인' 류만
  막히고 링크 업데이트·읽기전용 권장·매크로 보안·'파일 사용 중' 은 그대로 뜬다.
  → ① `_new_excel()` 이 낼 수 있는 창을 전부 끄고, ② `Workbooks.Open(..., Notify=False)` 로
    잠긴 파일에 창 대신 `com_error` 가 나게 하고, ③ 그 전에 `is_locked()` 로 걸러 Excel 을
    아예 안 띄우고, ④ `sys.excepthook` 이 남은 예외까지 받아 **기록만** 남기고 끝낸다.

재실행 판정 (2026-08-19 변경):
  종전엔 마커(campaign_schedule_last_source.txt)가 최신이면 무조건 SKIP 했다. 그런데 저장이 끝난 뒤
  Auto 파일이 외부(열려 있던 Excel / OneDrive 옛 버전 복원)에 의해 되돌려지면 **마커만 남고 내용은
  옛 소스인 상태로 굳어** 스케줄러가 영원히 SKIP 한다.
  실제 발생: 2026-08-19 — 마커는 0819 소스인데 Auto 파일 D1 은 0811 소스, 데이터도 55행(옛 값)이었다.
  (마커 mtime 15:54 < Auto 파일 mtime 15:55 → 우리 저장 이후 외부가 덮어쓴 정황)
  → 이제 target_is_saved() 가 Auto 파일의 D1 스탬프 + 붙여넣기 영역 값 + 잔재 행을 **실제로 대조**해
    판정한다. 마커는 기록·경고용으로만 남는다 (같은 파일을 쓰는 구 update_schedule.py 호환 유지).

정제 룰 (Summary → 일정):
  B Global      ← Summary 의 'Region'/'Global' 텍스트가 있는 열 (그룹 시작행에만 값)
  (No. 열은 건너뜀)
  C Subs        ← 헤더가 정확히 'Subs' 인 열  (5행 라벨 'Subs.' 아님)
  D Country     ← 'Country'
  E Participation ← B2B/B2C 날짜가 하나라도 파싱되면 'O', 아니면 빈칸
  F/H B2B 시작·종료 ← '캠페인 기간(B2B)' 을 '~' 로 분리 → M/D 파싱 → date
  G/I           ← 공백 (원래 WEEKNUM 자리)
  J/L B2C 시작·종료 ← '캠페인 기간(B2C)' 동일 처리
  K/M           ← 공백
  N note        ← 'Remark'

  ※ 'TBU' / '-' / 빈칸 처럼 '~' 가 없는 값은 시작·종료 모두 빈칸 처리.
  ※ 생성되는 `일정` 시트는 6행=Region 라벨, 7행=헤더, 8행~=데이터 (수동 작업본과 동일 레이아웃).
     구 update_schedule.py(SRC_MIN_ROW=3)를 돌려도 같은 결과가 나오도록 맞춘 것.
"""

import os
import re
import sys
import time
import shutil
import tempfile
import datetime as dt
from pathlib import Path
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import win32com.client
import pywintypes

# 콘솔 인코딩(cp949)에서 '—' 같은 문자를 print 하다 UnicodeEncodeError 로 죽는 걸 막는다.
# 2026-08-21 실제 발생 — 진행 로그 한 줄 때문에 실행 전체가 죽었다. pythonw 면 stdout 이 None.
if sys.stdout is not None:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, OSError):
        pass


# ════════════════ 사용자가 바꿔야 하는 부분 ════════════════

# ─── 경로 ────────────────────────────────────────────────────
BASE = Path(
    r"C:\Users\user_name\OneDrive - company_name"
    r"\Project_team_name - 1 company_name - 02 part_name"
    r"\part_name\2026\# CAMPAIGN_PROJECTS\03. CAMPAIGN NAME\01. SCHEDULE"
)

SOURCE_FOLDER    = BASE / "1.고객 법인 일정 파일"
# 소스로 인정할 확장자. 고객이 회차마다 xlsx/xlsb 를 오락가락 보낸다.
# ⚠ 종전엔 glob("*.xlsx") 가 곧 xlsb 배제 장치였다 — 이제 확장자로 거르지 않는다.
SOURCE_EXTS = (".xlsx", ".xlsb")

# 일정 소스로 인정할 파일명 키워드 (소문자 매칭, 하나라도 포함되면 후보).
# 2026-08-20: "monitoring" 추가 — 고객이 일정 내용을 Qualitative Monitoring 파일로 보내기 시작했다
#             (`26 CAMPAIGN NAME Qualitative Monitoring_260819_shared.xlsb`). 종전 3개 키워드로는 이름에서도,
#             glob("*.xlsx") 에서도 탈락해 최신본을 통째로 놓쳤다.
# ⚠ 이제 이 상수는 '무관한 파일 배제' 용 화이트리스트이지 Monitoring 배제 장치가 아니다.
SOURCE_NAME_KEYS = ["schedule", "캠페인", "일정", "monitoring"]
TARGET_SHEET     = "고객법인일정파일"
LAST_SOURCE_FILE = BASE / "campaign_schedule_last_source.txt"  # 마커: Auto 파일과 같은 폴더 (프로젝트별 독립 관리)

# ─── Excel COM 재계산 ────────────────────────────────────────
COM_RETRIES        = 2   # Excel 인스턴스가 도중에 죽었을 때(RPC_E_DISCONNECTED) 새 인스턴스로 재시도할 횟수
COM_RETRY_WAIT_SEC = 5   # 재시도 전 대기 (죽은 프로세스가 정리될 시간)

# ─── xlsb 변환본 · 실행 상태 ─────────────────────────────────
# 변환본은 **소스 폴더 직속**에 원본과 같은 이름(.xlsx)으로 남긴다 — 고객이 xlsb 만 보내는 회차에
# 사람이 열어볼 xlsx 와 `일정` 시트가 폴더에 하나도 없기 때문. Auto 갱신이 SKIP 돼도 항상 만든다.
# ⚠ 같은 이름의 .xlsb 가 있는 .xlsx 는 '변환본'으로 간주해 **소스 후보에서 제외**한다
#    (is_converted_twin). 안 그러면 다음 실행이 변환본을 소스로 집는다 — 도착시각이 같고
#    ext tiebreak 이 xlsx 우선이라 확정적으로 뒤집힌다.
XLSB_WORK_DIR = Path(tempfile.gettempdir()) / "campaign_schedule_xlsb_work"   # Excel 작업용 짧은 경로 (아래 ⚠ 참조)
STATUS_FILE   = SOURCE_FOLDER / "_schedule_update_status.txt"            # 마지막 실행 블록 + 실행 이력(누적)
STATUS_LOG_MAX_LINES = 200   # 상태 txt 하단 '실행 이력' 보관 줄 수 (초과분은 오래된 것부터 잘린다)

# ─── 정제 ────────────────────────────────────────────────────
CAMPAIGN_YEAR         = 2026        # 기간 텍스트 'M/D' 에 붙일 연도
SUMMARY_SHEET         = "Summary"   # 소스 원본 시트명 (없으면 첫 번째 시트로 fallback)
SCHEDULE_SHEET        = "일정"      # 생성할 정제 시트명
WRITE_SHEET_TO_SOURCE = True        # False 면 메모리 처리만 (소스 파일 미변경)

# Summary 헤더 문자열 — 열 문자 대신 이 이름으로 열을 찾는다 (고객이 열을 끼워넣어도 견디도록)
H_GLOBAL  = "Global"            # 이 헤더가 있는 열 = Region/Global 열
H_SUBS    = "Subs"              # 정확일치. 'Subs.'(라벨행) 는 매칭 안 됨
H_COUNTRY = "Country"
H_B2B     = "캠페인 기간(B2B)"
H_B2C     = "캠페인 기간(B2C)"
H_REMARK  = "Remark"

# 생성 시트 레이아웃 (수동 작업본과 동일)
SCHED_LABEL_ROW  = 6   # Region 라벨행
SCHED_HEADER_ROW = 7   # 헤더행 (데이터는 그 다음 행부터)

# 생성 시트 헤더행에 쓸 값 (B~N 13칸)
SCHED_HEADER = ["Global", "Subs", "Country", None,
                H_B2B, None, None, None,
                H_B2C, None, None, None,
                "note"]

# ─── 읽기/붙여넣기 범위 ───────────────────────────────────────
SRC_MIN_COL   = 2    # B (Global)
SRC_MAX_COL   = 14   # N (note)
TGT_CLEAR_ROW = 2    # 클리어 시작 행 (붙여넣기 위쪽 잔재까지 지우도록 더 위에서 시작)
TGT_START_ROW = 5    # 붙여넣기 시작 행 = Region 라벨행 → 헤더 6행, 데이터 7행~
TGT_MAX_ROW   = 999  # 클리어 범위 하단
STAMP_COL     = 4    # D — 붙여넣은 소스 파일명을 기록·대조하는 열 (D1). 재실행 판정의 1차 키

# 전후 비교(노란 음영) 대상 — {row_data 인덱스: 타겟 열번호}
# row_data 인덱스 0=B … 12=N,  타겟 열번호 = 인덱스 + 2
COMPARE = {
    3: 5,    # E Participation
    4: 6,    # F B2B 시작
    6: 8,    # H B2B 종료
    8: 10,   # J B2C 시작
    10: 12,  # L B2C 종료
}


# ════════════════ 내부 사용 ════════════════

CHANGED_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
NO_FILL      = PatternFill(fill_type=None)

PERIOD_SEP   = "~"                                  # 전각 ～ / ∼ 도 이 문자로 정규화 후 분리
PERIOD_ALTS  = ("～", "∼", "〜")
MD_PATTERN   = re.compile(r"^\s*(\d{1,2})\s*[/.\-]\s*(\d{1,2})\s*$")   # 'M/D' (. - 구분자도 허용)
ROW_LEN      = SRC_MAX_COL - SRC_MIN_COL + 1        # 13
XL_OPENXML_WORKBOOK = 51                            # Excel SaveAs FileFormat (xlsx)

# Excel 팝업 차단용 상수 (아래 _new_excel 참조)
XL_SECURITY_FORCE_DISABLE = 3                       # msoAutomationSecurityForceDisable — 매크로 보안 경고 차단
XL_FEATURE_INSTALL_NONE   = 0                       # msoFeatureInstallNone — 기능 설치 프롬프트 대신 오류 반환

# 저장 직후 재검증 재시도 — Excel 이 막 저장하고 핸들을 놓기 전 찰나에 읽으면
# PermissionError 가 나 '저장이 되돌려짐' 으로 오판된다 (2026-08-20 실측).
VERIFY_RETRIES  = 3
VERIFY_WAIT_SEC = 2

# 상태 txt 의 이력 구분선 — 이 줄 **아래**를 이력으로 보고 이어쓴다. 문자열이 바뀌면 이력이 끊긴다.
STATUS_LOG_HEADER = "── 실행 이력 (최신이 아래) ────────────────"


# ── 최신 파일 정렬 키 ────────────────────────────────────────
# 2026-08-20 (v2.0): **도착 시각 최우선**으로 전환.
#   종전 키 `(doc_date, mail_stamp, hhmm, ver_float, ver_int)` 는 파일명 안의 문서날짜가
#   1순위였다. 그래서 제목·형식(`_vN` / `_shared` / 8자리 vs 6자리 날짜)이 승패를 갈랐고,
#   고객이 옛 문서날짜로 새 파일을 보내면(재공유·수정본) 나중에 온 파일이 밀렸다.
#   이제 '언제 도착했나' 하나를 축으로 삼는다 — 제목·형식이 뭐가 됐든 최신이 이긴다.

def name_stamp(date_token: str, hhmm: str | None = None) -> int:
    """날짜 토큰(YYYYMMDD 또는 YYMMDD) + 선택적 HHMM → 비교용 정수 YYMMDDHHMM.

    8자리는 앞 2자리(세기)를 떼어 6자리로 맞춘다 — 20260819 → 260819 → 2608190000.
    이렇게 해야 SW형(8자리)과 MD형(6자리)이 **같은 축**에서 비교된다.
    (종전엔 20260819 vs 260819 를 그대로 비교해 자릿수만으로 승패가 갈렸다.)
    """
    d = date_token[2:] if len(date_token) == 8 else date_token
    return int(d) * 10000 + int(hhmm or 0)


def file_mtime_stamp(f: Path) -> int:
    """최후 수단 — 파일명에 날짜가 전혀 없을 때만 mtime 을 YYMMDDHHMM 으로 환산.

    ⚠ mtime 은 OneDrive 동기화·열어서 저장 등으로 쉽게 바뀌어 신뢰도가 낮다
      (실측: `_20260811_v2` 의 mtime 이 `_20260819_v1` 보다 나중이었다).
      파일명에 날짜가 하나라도 있으면 여기까지 오지 않는다.
    """
    try:
        t = dt.datetime.fromtimestamp(f.stat().st_mtime)
    except OSError:
        return 0
    return int(t.strftime("%y%m%d")) * 10000 + int(t.strftime("%H%M"))


def latest_file_key(f: Path):
    """(도착시각, 문서날짜, 문서시각, 버전float, 버전int, 끝번호) — 큰 튜플이 최신.

    **1순위 '도착시각'** 을 아래 우선순위로 정한다. 제목·버전 표기와 무관하다:
      a) 파일명 **맨 앞**의 `YYMMDD_HHMM_`  — 2026-08-20~ check_mail_attachment_byname.py 가 붙이는 수신일시
      b) 파일명 **맨 뒤**의 `_YYMMDD_HHMM` / `_YYMMDD` — 그 이전 수집분 (구형, 날짜만이면 00:00)
      c) 파일명 안의 문서날짜         — 스탬프 없는 옛 파일. 그 날 00:00 에 온 것으로 간주
      d) 파일 mtime                  — 날짜가 아예 없는 파일 (최후 수단)

    c) 가 필요한 이유: 2026-08-20 이전 수집분엔 스탬프가 없다. 이들을 0 으로 두면
    스탬프 붙은 옛 파일이 최신 파일을 이겨버린다 (`_20260806_260807_0949` 가
    `_20260819_v1` 을 이기는 회귀). 문서날짜를 도착시각의 근사값으로 써서 막는다.

    2~5순위는 **도착시각이 같을 때만** 쓰인다 (같은 메일에 v1·v2 가 동봉된 경우 등).

    예)
      260819_1548_..._20260819_v1      → (2608191548, 260819, 0, 0.0, 1, 0)  신형(접두)
      260819_2008_..._20260819_shared  → (2608192008, 260819, 0, 0.0, 0, 0)  ← 최신
      ..._20260819_v1_260819_1548      → (2608191548, 260819, 0, 0.0, 1, 0)  구형(접미)
      ..._20260819_v1                  → (2608190000, 260819, 0, 0.0, 1, 0)  스탬프 없음 → 문서날짜
      ..._v0.49_260420                 → (2604200000, 0,      0, 0.49, 0, 0) MD형
    """
    name = f.stem

    # ① 수신일시 스탬프를 뗀다. 2026-08-20 부터 **맨 앞**(YYMMDD_HHMM_)에 붙이고,
    #    그 이전 수집분은 **맨 뒤**(_YYMMDD_HHMM / _YYMMDD)에 붙어 있어 둘 다 읽는다.
    m = re.match(r"^(\d{6})_(\d{4})_", name)            # 신형: 접두
    if m:
        arrive, core = name_stamp(m.group(1), m.group(2)), name[m.end():]
    else:
        m = re.search(r"_(\d{6})_(\d{4})$", name)        # 구형: 접미 (날짜+시각)
        if m:
            arrive, core = name_stamp(m.group(1), m.group(2)), name[:m.start()]
        else:
            m = re.search(r"_(\d{6})$", name)             # 구형: 접미 (날짜만)
            arrive, core = (name_stamp(m.group(1)), name[:m.start()]) if m else (None, name)

    # ② 남은 이름에서 문서날짜(8자리 우선, 없으면 6자리)와 문서시각을 읽는다
    m8 = re.search(r"(?<!\d)(\d{8})(?!\d)", core)
    m6 = None if m8 else re.search(r"(?<!\d)(\d{6})(?!\d)", core)
    doc_token = m8.group(1) if m8 else (m6.group(1) if m6 else None)

    mh = re.search(r"(?<!\d)\d{8}_(\d{4})(?!\d)", core)
    doc_hhmm = int(mh.group(1)) if mh else 0

    # ③ 도착시각 확정 — 스탬프 없으면 문서날짜, 그것도 없으면 mtime
    if arrive is None:
        arrive = name_stamp(doc_token, doc_hhmm or None) if doc_token else file_mtime_stamp(f)

    # ④ 동점 tiebreak — 문서날짜도 6자리로 정규화해 SW/MD 형이 같은 축에 오게 한다
    doc_date = int(doc_token[2:] if doc_token and len(doc_token) == 8 else (doc_token or 0))

    mv = re.search(r"_v(\d+\.\d+)", core)
    ver_float = float(mv.group(1)) if mv else 0.0
    mv = re.search(r"_v(\d+)(?!\.\d)", core)
    ver_int = int(mv.group(1)) if mv else 0

    # MD형의 수기 끝번호 (`..._v0.44_260319_2`) — 도착시각·버전까지 같을 때의 마지막 tiebreak.
    # SW형은 끝이 `_vN` 이거나 8자리 날짜라 여기에 걸리지 않는다.
    ms = re.search(r"_(\d{1,5})$", core)
    suffix = int(ms.group(1)) if ms else 0

    # 같은 메일에 같은 이름의 xlsx/xlsb 가 동봉되면 여기까지 전부 동점이라 순서가
    # 파일시스템 순서에 좌우된다 → 내용이 같다면 변환이 필요 없는 xlsx 를 택한다.
    ext_rank = 1 if f.suffix.lower() == ".xlsx" else 0

    return (arrive, doc_date, doc_hhmm, ver_float, ver_int, suffix, ext_rank)


# ── Summary 정제 ─────────────────────────────────────────────
def parse_md(token: str) -> dt.date | None:
    """'8/17' 같은 M/D 토큰 → date(CAMPAIGN_YEAR, M, D). 형식이 아니면 None."""
    m = MD_PATTERN.match(token or "")
    if not m:
        return None
    try:
        return dt.date(CAMPAIGN_YEAR, int(m.group(1)), int(m.group(2)))
    except ValueError:          # 13/45 같은 오타
        return None


def parse_period(value) -> tuple[dt.date | None, dt.date | None]:
    """'6/24~9/13' → (date, date). '~' 가 없는 값(TBU, -, 빈칸)은 (None, None)."""
    if not isinstance(value, str):
        return None, None
    s = value
    for alt in PERIOD_ALTS:
        s = s.replace(alt, PERIOD_SEP)
    if PERIOD_SEP not in s:
        return None, None
    head, tail = s.split(PERIOD_SEP, 1)
    start, end = parse_md(head), parse_md(tail)
    # 12/20~1/15 처럼 해를 넘기는 기간 → 종료일만 다음 해로
    if start and end and end < start:
        end = end.replace(year=end.year + 1)
    return start, end


def find_summary_layout(ws):
    """Summary 시트에서 헤더행 위치와 필요한 열 번호를 찾아 반환.

    반환: (header_row, {논리명: 열번호})
    """
    header_row = global_col = None
    for r in range(1, min(ws.max_row, 30) + 1):
        for c in range(1, ws.max_column + 1):
            if isinstance(ws.cell(r, c).value, str) and ws.cell(r, c).value.strip() == H_GLOBAL:
                header_row, global_col = r, c
                break
        if header_row:
            break
    if not header_row:
        raise ValueError(f"'{H_GLOBAL}' 헤더를 찾을 수 없습니다 (시트: {ws.title})")

    cols = {"global": global_col}
    for key, header in (("subs", H_SUBS), ("country", H_COUNTRY),
                        ("b2b", H_B2B), ("b2c", H_B2C), ("remark", H_REMARK)):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(header_row, c).value
            if isinstance(v, str) and v.strip() == header:
                cols[key] = c
                break
        else:
            raise ValueError(f"'{header}' 헤더를 찾을 수 없습니다 (시트: {ws.title}, {header_row}행)")
    return header_row, cols


def _new_excel():
    """팝업을 낼 수 있는 경로를 전부 끈 전용 Excel 인스턴스.

    ⚠ `DisplayAlerts=False` 만으로는 부족하다 — 그건 '저장/덮어쓰기 확인' 류만 막고
      **링크 업데이트·읽기전용 권장·매크로 보안·기능 설치** 프롬프트는 그대로 뜬다.
      무인(pythonw) 실행에서 모달 창이 뜨면 사람이 닫아줄 때까지 Excel 이 멈춰 서서
      다음 회차까지 물리고 고아 EXCEL.EXE 가 쌓인다. 그래서 낼 수 있는 창을 전부 끈다.

    ※ DispatchEx = 전용 인스턴스 (이미 떠 있는 Excel 에 붙으면 그쪽이 먼저 종료될 때
      Quit 단계에서 죽는다 — 아래 두 호출부의 주석 참조).
    """
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible                = False
    excel.DisplayAlerts          = False   # 저장·덮어쓰기 확인
    excel.AskToUpdateLinks       = False   # "이 통합 문서에 링크가 있습니다" — DisplayAlerts 로 안 막힌다
    excel.AlertBeforeOverwriting = False
    excel.EnableEvents           = False   # 통합 문서 Open 이벤트가 띄우는 창
    excel.ScreenUpdating         = False
    excel.AutomationSecurity     = XL_SECURITY_FORCE_DISABLE
    excel.FeatureInstall         = XL_FEATURE_INSTALL_NONE
    return excel


def _com_convert_to_xlsx(src: Path, dest: Path) -> None:
    """Excel COM 으로 src(xlsb) → dest(xlsx) 변환. 실패 시 예외를 올린다.

    ⚠ **Excel 에는 짧은 임시 경로만 넘긴다** — 복사 → 변환 → 되옮기기 3단계인 이유.
      캠페인 폴더가 깊어서 소스도 dest 도 전체 경로가 264자(`MAX_PATH` 260 초과)다. 원본 경로를
      그대로 넘기면 Excel 이 `Workbooks 클래스 중 Open 메서드에 오류가 있습니다`(0x800A03EC) 로 거부한다.
      Python 은 같은 경로를 읽고 쓰는 데 문제가 없어서(롱패스) 원인이 잘 안 보인다 — 인자 조합·파일 손상
      문제로 오인하기 쉽다 (2026-08-20 실측). Excel 공식 한도는 경로+파일명 218자라 Auto 파일(223자)도
      아슬아슬하다 — 캠페인 폴더명이 한 번만 더 길어지면 recalc_and_save 쪽도 같은 이유로 깨진다.

    ※ 인스턴스 설정(팝업 차단)은 _new_excel() 에 모아 뒀다.
    """
    XLSB_WORK_DIR.mkdir(parents=True, exist_ok=True)
    work_src  = XLSB_WORK_DIR / src.name
    work_dest = work_src.with_suffix(".xlsx")
    shutil.copy2(src, work_src)
    try:
        excel = _new_excel()
        try:
            wb_com = excel.Workbooks.Open(
                str(work_src),
                UpdateLinks=0,                   # 외부링크 갱신 팝업 방지
                ReadOnly=True,                   # 사본이라 원본은 어차피 안전
                IgnoreReadOnlyRecommended=True,  # "읽기 전용으로 여시겠습니까"
                Notify=False,                    # 잠겨 있으면 '파일 사용 중' 창 대신 com_error
            )
            try:
                wb_com.SaveAs(str(work_dest), FileFormat=XL_OPENXML_WORKBOOK)
            finally:
                # 지연 바인딩이면 `.Close` 속성 접근만으로 COM 메서드가 실행되고 bool 을 돌려준다
                # → 이어지는 () 가 TypeError (recalc_and_save 의 같은 주석 참조).
                try:
                    wb_com.Close(SaveChanges=False)
                except TypeError:
                    pass
        finally:
            try:
                excel.Quit()
            except Exception as e:
                print(f"[알림] Excel 종료 중 무시된 예외: {e}")

        shutil.copy2(work_dest, dest)   # 긴 경로로 되옮기기는 Python 이 한다 (copy2 = 덮어쓰기 OK)
    finally:
        # Dispatch 자체가 실패해도 작업본은 반드시 지운다
        for f in (work_src, work_dest):
            try:
                f.unlink(missing_ok=True)
            except OSError:
                pass


def converted_twin(path: Path) -> Path:
    """xlsb 소스에 대응하는 변환본 경로 — **같은 폴더, 같은 이름, .xlsx**."""
    return path.with_suffix(".xlsx")


def is_converted_twin(f: Path) -> bool:
    """이 .xlsx 가 같은 이름 xlsb 의 변환본인가 → 소스 후보에서 제외해야 한다.

    고객이 같은 회차에 진짜 xlsx·xlsb 를 둘 다 보내도 내용은 같으므로 xlsb 쪽만 써도 무해하다.
    """
    return f.suffix.lower() == ".xlsx" and f.with_suffix(".xlsb").exists()


def ensure_openpyxl_readable(path: Path) -> Path:
    """openpyxl 이 읽을 수 있는 경로를 돌려준다 (xlsx 는 그대로, xlsb 는 변환본 경로).

    openpyxl 은 xlsb 를 아예 못 읽는다. 새 의존성(pyxlsb)을 들이는 대신, 이 스크립트가 어차피
    쓰고 있는 Excel COM 으로 xlsx 변환본을 떠서 기존 정제 경로를 그대로 태운다.

    변환본이 소스보다 새것이면 재변환하지 않는다 — SKIP 판정에도 정제 결과가 필요해서,
    이 가드가 없으면 20분마다 도는 스케줄러가 '변경 없음' 회차에도 매번 Excel 을 띄운다.
    (`일정` 시트를 써 넣으면 변환본 mtime 이 더 새것이 되므로 판정은 계속 유효하다.)
    """
    if path.suffix.lower() != ".xlsb":
        return path

    dest = converted_twin(path)
    if dest.exists() and dest.stat().st_mtime >= path.stat().st_mtime:
        return dest

    for attempt in range(1, COM_RETRIES + 1):
        try:
            _com_convert_to_xlsx(path, dest)
            break
        except pywintypes.com_error as e:
            print(f"[재시도 {attempt}/{COM_RETRIES}] xlsb → xlsx 변환 실패: {e}")
            if attempt == COM_RETRIES:
                raise
            time.sleep(COM_RETRY_WAIT_SEC)

    print(f"[xlsb 변환] {path.name} → {dest.name}")
    return dest


def build_schedule_rows(xlsx_path: Path) -> tuple[list, list]:
    """소스 xlsx 의 Summary 를 정제해 (헤더 2행, 데이터 행들) 반환.

    각 행은 B~N 13칸 리스트. 헤더 2행 = [Region 라벨행, 컬럼 헤더행].
    소스가 xlsb 면 여기서 xlsx 변환본으로 바꿔 읽는다 (소스 파일 자체는 무변경).
    """
    xlsx_path = ensure_openpyxl_readable(xlsx_path)
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[SUMMARY_SHEET] if SUMMARY_SHEET in wb.sheetnames else wb.worksheets[0]
    header_row, cols = find_summary_layout(ws)

    # Region 라벨행 = 헤더행 바로 위 (B열의 'Region')
    label_value = ws.cell(header_row - 1, cols["global"]).value if header_row > 1 else None
    label_row   = [label_value] + [None] * (ROW_LEN - 1)

    data_rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        subs    = ws.cell(r, cols["subs"]).value
        country = ws.cell(r, cols["country"]).value
        if subs in (None, "") and country in (None, ""):
            continue                                    # 데이터 행 아님

        b2b_start, b2b_end = parse_period(ws.cell(r, cols["b2b"]).value)
        b2c_start, b2c_end = parse_period(ws.cell(r, cols["b2c"]).value)
        participation = "O" if any((b2b_start, b2b_end, b2c_start, b2c_end)) else None

        data_rows.append([
            ws.cell(r, cols["global"]).value,   # B Global
            subs,                               # C Subs
            country,                            # D Country
            participation,                      # E Participation
            b2b_start, None, b2b_end, None,     # F~I  (G/I = WEEKNUM 자리, 공백)
            b2c_start, None, b2c_end, None,     # J~M  (K/M = WEEKNUM 자리, 공백)
            ws.cell(r, cols["remark"]).value,   # N note
        ])

    wb.close()
    return [label_row, list(SCHED_HEADER)], data_rows


def write_schedule_sheet(xlsx_path: Path, label_rows: list, data_rows: list) -> bool:
    """소스 xlsx 에 정제 결과를 `일정` 시트로 기록 (**없을 때만** 새로 만든다).

    2026-08-20: 기존 시트가 있으면 지우고 다시 쓰던 것을 **그대로 두는** 방식으로 변경.
    고객이 보낸 `일정` 시트나 손으로 고친 내용이 매 실행마다 덮여 사라지는 걸 막는다.
    → 다시 만들고 싶으면 소스 파일에서 `일정` 시트를 지우고 재실행하면 된다.

    파일 수정 시각(mtime)은 원래대로 되돌린다 — 마커가 '메일로 받은 버전'을 가리키도록 유지하고,
    우리가 쓴 것 때문에 다음 실행이 재처리로 오인하지 않게.

    2026-08-20: xlsb 소스는 건너뛴다 — openpyxl 은 xlsb 저장이 불가하고, 변환 캐시에 써봐야
    다음 회차에 버려지는 임시파일이라 의미가 없다.
    """
    if xlsx_path.suffix.lower() != ".xlsx":
        # xlsb 원본에는 못 쓴다. 대신 호출부가 **변환본(.xlsx)** 을 넘겨 거기에 기록한다.
        print(f"[정제 시트] xlsb 원본에는 '{SCHEDULE_SHEET}' 시트를 쓸 수 없습니다 — {xlsx_path.name}")
        return False

    orig_stat = xlsx_path.stat()
    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=False)
    except PermissionError:
        print(f"[알림] 소스 파일이 사용 중이라 '{SCHEDULE_SHEET}' 시트 기록을 건너뜁니다: {xlsx_path.name}")
        return False

    # 이미 있으면 손대지 않는다 (수기 편집·고객 원본 보존). 저장도 안 하므로 파일 무변경.
    if SCHEDULE_SHEET in wb.sheetnames:
        wb.close()
        print(f"[정제 시트] '{SCHEDULE_SHEET}' 시트가 이미 있어 그대로 둡니다 — {xlsx_path.name}")
        return True

    ws = wb.create_sheet(SCHEDULE_SHEET, 0)

    for r_offset, row_data in enumerate(label_rows + data_rows):
        r_idx = SCHED_LABEL_ROW + r_offset
        for c_offset, value in enumerate(row_data):
            cell = ws.cell(row=r_idx, column=SRC_MIN_COL + c_offset, value=value)
            if isinstance(value, dt.date):
                cell.number_format = "YYYY-MM-DD"

    try:
        wb.save(xlsx_path)
    except PermissionError:
        print(f"[알림] 소스 파일 저장 실패(사용 중) — '{SCHEDULE_SHEET}' 시트 기록 생략: {xlsx_path.name}")
        return False
    finally:
        wb.close()

    os.utime(xlsx_path, (orig_stat.st_atime, orig_stat.st_mtime))
    print(f"[정제 시트] '{SCHEDULE_SHEET}' 시트 신규 생성 ({len(data_rows)}행) — {xlsx_path.name}")
    return True


# ── Auto 파일 저장상태 확인 ─────────────────────────────────
def _norm(v):
    """저장된 값 ↔ 정제 결과를 비교 가능한 형태로 정규화.

    openpyxl 은 date 로 쓴 값을 datetime 으로 되읽고, 빈 문자열과 빈칸도 구분되므로
    그대로 비교하면 매번 불일치가 난다.
    """
    if isinstance(v, dt.datetime):          # date 로 썼어도 datetime 으로 돌아온다
        return v.date()
    if isinstance(v, str):
        v = v.replace("\r\n", "\n").strip()
        return v or None                    # 빈 문자열 == 빈칸
    return v


def target_is_saved(output_file: Path, source_name: str, src_data: list) -> tuple[bool, str]:
    """Auto 파일에 현재 소스의 정제 결과가 **실제로 저장돼 있는지** 확인.

    (True, "")    = 이미 반영됨 → 실행 불필요
    (False, 사유) = 미반영·유실 → 실행 필요

    마커(로그)를 믿지 않고 파일 내용으로만 판정한다. 저장이 끝난 뒤 외부(Excel/OneDrive)가
    파일을 되돌려도 다음 실행이 스스로 알아채고 재처리하게 하는 게 목적.
    """
    try:
        wb = openpyxl.load_workbook(output_file, data_only=True, read_only=True)
    except Exception as e:                  # 사용 중·손상 등 — 못 읽으면 '미반영' 으로 보고 진행
        return False, f"Auto 파일을 읽을 수 없음 ({type(e).__name__}: {e})"

    try:
        if TARGET_SHEET not in wb.sheetnames:
            return False, f"'{TARGET_SHEET}' 시트 없음"
        ws = wb[TARGET_SHEET]
        # read_only 모드는 ws.cell() 랜덤 접근이 느리므로 한 번에 훑어 dict 로 받는다
        rows = {r_idx: row for r_idx, row in enumerate(
            ws.iter_rows(min_row=1, max_row=TGT_MAX_ROW,
                         min_col=SRC_MIN_COL, max_col=SRC_MAX_COL,
                         values_only=True), start=1)}
    finally:
        wb.close()

    def cell(r_idx: int, col: int):
        row = rows.get(r_idx)
        idx = col - SRC_MIN_COL
        return row[idx] if row and idx < len(row) else None

    # 1) D1 소스 파일명 스탬프
    stamp = _norm(cell(1, STAMP_COL))
    if stamp != source_name:
        return False, f"D1 소스명 불일치 (저장됨 {stamp!r} != 현재 {source_name!r})"

    # 2) 붙여넣기 영역 값 대조 (B5~ / 13열)
    for r_offset, want_row in enumerate(src_data):
        r_idx = TGT_START_ROW + r_offset
        for c_offset, want in enumerate(want_row):
            got = cell(r_idx, SRC_MIN_COL + c_offset)
            if _norm(got) != _norm(want):
                col = get_column_letter(SRC_MIN_COL + c_offset)
                return False, f"{col}{r_idx} 값 불일치 (저장됨 {got!r} != 소스 {want!r})"

    # 3) 영역 밖 잔재 — 행 수가 줄었는데 옛 행이 남은 경우 (1행은 스탬프 행이라 제외)
    paste_end = TGT_START_ROW + len(src_data) - 1
    for r_idx, row in rows.items():
        if r_idx < TGT_CLEAR_ROW or TGT_START_ROW <= r_idx <= paste_end:
            continue
        if any(_norm(v) is not None for v in row):
            return False, f"{r_idx}행에 옛 데이터 잔재"

    return True, ""


# ── 실행 상태 기록 ──────────────────────────────────────────
def write_status(result: str, detail: str = "", rows: int | None = None) -> None:
    """소스 폴더의 상태 txt **한 개**를 갱신 — 상단 '마지막 실행' 블록 + 하단 '실행 이력'(누적).

    스케줄러(pythonw)로 돌면 콘솔 출력이 아무데도 안 남아서, 갱신이 됐는지 SKIP 인지 실패인지 알 방법이
    Auto 파일을 직접 열어보는 것뿐이었다. 그 확인을 파일 하나로 대신한다.

    2026-08-21: 마지막 1회분만 덮어쓰던 것을 **이력 누적**으로 바꿨다 — 팝업 없이 조용히 물러난 회차가
    쌓이면 '언제부터 실패했는지' 가 보여야 추적이 된다. 파일은 계속 1개다 (이력은 같은 txt 하단).

    ※ `.txt` 라 SOURCE_EXTS 에 안 걸려 소스 후보를 오염시키지 않는다.
    ※ 전역(source_file 등)이 **아직 안 정해진 시점에도 불릴 수 있다** (sys.excepthook 이 이른 단계의
      예외에서 호출) → globals().get() 으로 방어적으로 읽고, 없는 항목은 줄을 생략한다.
    """
    now  = dt.datetime.now()
    src  = globals().get("source_file")
    work = globals().get("source_work")
    out  = globals().get("output_file")

    lines = [
        "이 파일은 update_schedule_summary.py 가 매 실행마다 갱신합니다.",
        "",
        "── 마지막 실행 ────────────────────────────",
        f"실행 시각 : {now:%Y-%m-%d %H:%M:%S}",
        f"결과      : {result}",
    ]
    if detail:
        lines.append(f"상세      : {detail}")
    if src is not None:
        lines.append(f"소스 파일 : {src.name}")
    if work is not None and work != src:
        mark = "일정 시트 O" if globals().get("schedule_sheet_ok") else "일정 시트 X"
        lines.append(f"변환본    : {work.name} ({mark})")
    if out is not None:
        lines.append(f"Auto 파일 : {out.name}")
    if rows is not None:
        lines.append(f"데이터    : {rows}행")

    # ── 이력: 기존 파일에서 구분선 아래를 회수 → 이번 줄 append → 오래된 것부터 잘라냄
    history = []
    try:
        old = STATUS_FILE.read_text(encoding="utf-8").splitlines()
        history = [ln for ln in old[old.index(STATUS_LOG_HEADER) + 1:] if ln.strip()]
    except (OSError, ValueError):
        pass                                    # 파일 없음 / 구분선 없음(구 포맷) → 이력 새로 시작
    tail = detail or (f"{rows}행" if rows is not None else "")
    history.append(f"{now:%Y-%m-%d %H:%M:%S}  {result:<20}  {tail}".rstrip())
    history = history[-STATUS_LOG_MAX_LINES:]

    lines += ["", STATUS_LOG_HEADER] + history

    try:
        STATUS_FILE.write_text("\n".join(lines) + "\n", encoding="utf-8")
    except OSError as e:
        print(f"[알림] 상태 파일 기록 실패 ({type(e).__name__}: {e})")


# ── 마지막 안전망 — 예상 못 한 예외도 창 없이 기록만 남기고 끝낸다 ──
# pythonw 로 돌면 traceback 은 어차피 어디에도 안 남는다. 상태 txt 이력에 한 줄이라도 남겨야
# "몇 시에 뭘로 죽었는지" 를 나중에 볼 수 있다.
def _log_uncaught(exc_type, exc, tb):
    try:
        write_status("실패(예외)", f"{exc_type.__name__}: {exc}")
    except Exception:
        pass                                    # 기록조차 못 해도 창은 띄우지 않는다
    if sys.stderr is not None:                  # 콘솔에서 돌릴 땐 traceback 도 보여준다
        import traceback
        traceback.print_exception(exc_type, exc, tb)


sys.excepthook = _log_uncaught


# ── Auto 파일 자동 탐색 ──────────────────────────────────────
auto_files = list(BASE.glob("*Auto*.xlsx"))
if not auto_files:
    raise FileNotFoundError(f"Auto 파일을 찾을 수 없습니다: {BASE}")
output_file = auto_files[0]
print(f"[업데이트 대상] {output_file.name}")

# ── 소스 폴더에서 최신 파일 선택 ────────────────────────────
# SOURCE_EXTS 확장자 + 이름에 SOURCE_NAME_KEYS 가 하나라도 있는 파일만 소스 후보.
# 확장자(.xlsx/.xlsb)도 제목 형태(_vN/_shared)도 순위에 관여하지 않는다 — 오직 도착순(latest_file_key).
source_files = sorted(
    (f for f in SOURCE_FOLDER.iterdir()
     if f.is_file()
     and f.suffix.lower() in SOURCE_EXTS
     and not f.name.startswith("~$")                     # Excel 잠금 임시파일
     and not is_converted_twin(f)                       # 우리가 만든 xlsb 변환본
     and any(k in f.name.lower() for k in SOURCE_NAME_KEYS)),
    key=latest_file_key,
)
if not source_files:
    raise FileNotFoundError(
        f"소스 폴더에 일정 파일이 없습니다 "
        f"(확장자 {SOURCE_EXTS} + 이름에 {SOURCE_NAME_KEYS} 중 하나 필요): {SOURCE_FOLDER}")

source_file = source_files[-1]
print(f"[소스 파일] {source_file.name}")

# ── Summary 정제 (메모리 — 소스 파일 미변경) ─────────────────
# ※ OneDrive 동기화 중이면 잠깐 잠길 수 있다. 작업 스케줄러에 '실패'로 남기지 말고
#   조용히 물러나 다음 실행(20분 뒤)이 재시도하게 한다.
source_work = source_file          # xlsb 면 아래에서 변환본 경로로 바뀐다 (write_status 가 참조)
schedule_sheet_ok = False

try:
    source_work = ensure_openpyxl_readable(source_file)
    label_rows, data_rows = build_schedule_rows(source_work)
except (OSError, pywintypes.com_error) as e:
    # com_error = xlsb → xlsx 변환 실패 (Excel 이 죽었거나 파일이 잠김). 이것도 재시도 대상이다.
    print(f"[SKIP] 소스 파일을 읽을 수 없습니다 ({type(e).__name__}). 다음 실행 시 재시도합니다: {source_file.name}")
    write_status("실패", f"소스 파일을 읽을 수 없음 ({type(e).__name__})")
    exit(0)
src_data = label_rows + data_rows
print(f"[정제 완료] 데이터 {len(data_rows)}행 (+ 헤더 {len(label_rows)}행)")

# ── 변환본에 `일정` 시트 — Auto 갱신이 SKIP 돼도 **항상** 만든다 ──
# 고객이 xlsb 만 보내는 회차엔 폴더에 사람이 열어볼 xlsx 도 `일정` 시트도 없다.
# 그래서 이 단계만 SKIP 판정 **앞**에 둔다 (xlsx 소스는 종전대로 갱신이 필요할 때만 쓴다).
if WRITE_SHEET_TO_SOURCE and source_work != source_file:
    schedule_sheet_ok = write_schedule_sheet(source_work, label_rows, data_rows)

# ── 재실행 판정 — 마커(로그)가 아니라 Auto 파일에 실제 저장된 내용으로 ──
# ※ 마커는 '처리 완료' 기록·경고용으로만 남긴다. 판정에 쓰면 저장이 유실됐을 때 영원히 SKIP 된다.
src_mtime = int(source_file.stat().st_mtime)
current_marker = f"{source_file.name}|{src_mtime}"
marker_says_done = (LAST_SOURCE_FILE.exists()
                    and LAST_SOURCE_FILE.read_text(encoding="utf-8").strip() == current_marker)

saved, reason = target_is_saved(output_file, source_file.name, src_data)
if saved:
    print(f"[SKIP] Auto 파일에 이미 반영돼 있습니다 ({source_file.name})")
    if not marker_says_done:
        LAST_SOURCE_FILE.write_text(current_marker, encoding="utf-8")
        print("[알림] 내용은 최신이라 마커만 뒤늦게 동기화했습니다.")
    write_status("이미 반영됨 (SKIP)", rows=len(data_rows))
    exit(0)

if marker_says_done:
    print("[경고] 마커는 '처리 완료'인데 Auto 파일 내용은 최신이 아닙니다 — 저장이 유실된 것으로 보고 재실행합니다.")
    print("        (Auto 파일이 Excel 에서 열려 있었거나 OneDrive 가 옛 버전으로 되돌렸을 수 있습니다)")
print(f"[갱신 필요] {reason}")

if WRITE_SHEET_TO_SOURCE:
    write_schedule_sheet(source_file, label_rows, data_rows)

# ── 이전 파일 정제 (전후 비교용) ─────────────────────────────
# 키가 Subs 만이면 SUB_C(북유럽 4국) 처럼 같은 Subs 가 여러 행인 경우 마지막 행만 남는다.
# → (Subs, Country, 같은 조합의 몇 번째) 로 키를 잡아 행 단위로 정확히 대응시킨다.
def compare_key(row_data, seen: dict):
    key = (row_data[1], row_data[2])
    seen[key] = seen.get(key, 0) + 1
    return key + (seen[key],)


prev_data = {}
if len(source_files) >= 2:
    prev_file = source_files[-2]
    try:
        _, prev_rows = build_schedule_rows(prev_file)
    except Exception as e:
        # 전후 비교는 **부가 기능**이라 여기서 죽으면 안 된다. 실패 원인 2종 모두 비교만 생략:
        #  - ValueError/KeyError : 옛 포맷(Summary 없음/헤더 다름)
        #  - OSError(PermissionError 등) : OneDrive 동기화 중 파일 잠김 (2026-08-19 실제 발생 —
        #    이 절이 좁아서 스크립트 전체가 traceback 으로 죽었다)
        prev_rows = []
        print(f"[알림] 이전 파일 정제 불가 — 전후 비교 생략 ({prev_file.name}): {type(e).__name__}: {e}")
    seen_prev = {}
    for row_data in prev_rows:
        prev_data[compare_key(row_data, seen_prev)] = row_data
    if prev_data:
        print(f"[이전 파일] {prev_file.name} ({len(prev_data)}행 로드)")

# ── 타겟 파일 업데이트 ───────────────────────────────────────
try:
    tgt_wb = openpyxl.load_workbook(output_file)
except PermissionError:
    print(f"[SKIP] 파일이 사용 중입니다. 다음 실행 시 재시도합니다: {output_file.name}")
    write_status("실패", "Auto 파일이 사용 중 (다음 실행 시 재시도)", rows=len(data_rows))
    exit(0)

if TARGET_SHEET not in tgt_wb.sheetnames:
    raise ValueError(f"'{TARGET_SHEET}' 시트를 찾을 수 없습니다. 시트 목록: {tgt_wb.sheetnames}")

tgt_ws = tgt_wb[TARGET_SHEET]

# D1에 소스 파일명 기록 (재실행 판정의 1차 키 — target_is_saved() 가 이 값을 대조한다)
tgt_ws.cell(row=1, column=STAMP_COL, value=source_file.name)

# 대상 영역(B2:N999)과 겹치는 병합셀 해제 — MergedCell 은 value 설정 불가(read-only)라
# 클리어/붙여넣기에서 충돌. 어차피 이 영역은 소스값으로 덮어쓰므로 해제해도 무방.
for rng in list(tgt_ws.merged_cells.ranges):
    if (rng.max_row >= TGT_CLEAR_ROW and rng.min_row <= TGT_MAX_ROW
            and rng.max_col >= SRC_MIN_COL and rng.min_col <= SRC_MAX_COL):
        tgt_ws.unmerge_cells(str(rng))

# B2:N999 값·음영 클리어 (서식 유지)
for row in tgt_ws.iter_rows(min_row=TGT_CLEAR_ROW, max_row=TGT_MAX_ROW,
                            min_col=SRC_MIN_COL, max_col=SRC_MAX_COL):
    for cell in row:
        cell.value = None
        cell.fill  = NO_FILL

# B5(라벨행)부터 값 붙여넣기 → 헤더 6행, 데이터 7행~
for r_idx, row_data in enumerate(src_data, start=TGT_START_ROW):
    for c_idx, value in enumerate(row_data, start=SRC_MIN_COL):  # B열=2
        cell = tgt_ws.cell(row=r_idx, column=c_idx, value=value)
        if isinstance(value, dt.date):
            cell.number_format = "YYYY-MM-DD"

# 변경 셀 음영 표시 (대상 = 상단 COMPARE). 헤더 2행은 건너뛰고 데이터 행만 비교.
if prev_data:
    changed_count = 0
    seen_cur = {}
    data_start_row = TGT_START_ROW + len(label_rows)
    for r_idx, row_data in enumerate(data_rows, start=data_start_row):
        key = compare_key(row_data, seen_cur)
        if key not in prev_data:
            continue
        prev_row = prev_data[key]
        for src_idx, tgt_col in COMPARE.items():
            cur_val = row_data[src_idx] if src_idx < len(row_data) else None
            prv_val = prev_row[src_idx] if src_idx < len(prev_row) else None
            if cur_val != prv_val:
                tgt_ws.cell(row=r_idx, column=tgt_col).fill = CHANGED_FILL
                changed_count += 1
    print(f"[변경 셀] {changed_count}개 음영 표시")

try:
    tgt_wb.save(output_file)
    tgt_wb.close()
except PermissionError:
    tgt_wb.close()
    print(f"[SKIP] 저장 중 파일이 잠겼습니다. 다음 실행 시 재시도합니다: {output_file.name}")
    write_status("실패", "저장 중 Auto 파일이 잠김 (다음 실행 시 재시도)", rows=len(data_rows))
    exit(0)

# Excel로 열어서 전체 재계산 후 저장 (FILTER/SORT 등 동적 배열 함수 반영)
# ※ openpyxl 저장은 수식의 캐시값을 지우므로 이 재계산 저장이 반드시 성공해야 한다.
#   (캐시값이 없으면 openpyxl data_only=True 로 읽는 후속 도구가 전부 None 을 본다)
# ※ DispatchEx = 전용 인스턴스. Dispatch 는 이미 떠 있는 Excel 에 붙어서, 그쪽이 먼저
#   종료되면 Quit 단계에서 AttributeError 로 죽는다(작업 스케줄러가 실패로 기록).
#   인스턴스 설정(팝업 차단)은 _new_excel() 참조.
def recalc_and_save(path: Path) -> None:
    """Excel COM 으로 전체 재계산 후 저장. 실패 시 예외를 올린다."""
    excel = _new_excel()
    try:
        wb_com = excel.Workbooks.Open(
            str(path.resolve()),
            UpdateLinks=0,                   # 링크 갱신 안 함
            IgnoreReadOnlyRecommended=True,  # "읽기 전용으로 여시겠습니까"
            Notify=False,                    # ★ 잠겨 있으면 '파일 사용 중' 창 대신 com_error 를 던진다
                                             #   (기본값 True 면 알림 대화상자를 띄우고 사람을 기다린다)
        )
        excel.CalculateFull()
        wb_com.Save()
        # ※ gen_py 캐시가 없으면(파이썬 새 버전 설치 직후 등) win32com 이 지연 바인딩으로 동작해
        #   `wb_com.Close` 는 속성 접근만으로 COM 메서드가 실행되고 bool(True) 을 돌려준다
        #   → 이어지는 `()` 가 TypeError: 'bool' object is not callable 로 죽는다.
        #   그 시점엔 이미 저장·닫기가 끝난 상태이므로 TypeError 만 무시한다.
        #   (조기 바인딩이면 아래 호출이 정상 동작 — 양쪽 다 안전)
        try:
            wb_com.Close(SaveChanges=False)
        except TypeError:
            pass
    finally:
        try:
            excel.Quit()
        except Exception as e:      # 이미 죽은 인스턴스 등 — 저장 성패는 위에서 판정
            print(f"[알림] Excel 종료 중 무시된 예외: {e}")


def is_locked(path: Path) -> bool:
    """다른 프로세스가 쓰기 잠금 중인가 — Excel 을 띄우기 **전에** 값싸게 확인.

    Notify=False 로 대화상자 대신 오류가 나게 만들어 놨어도, 이미 떠버린 Excel 은
    닫힐 때까지 메모리를 잡고 있다. 사람이 Auto 파일을 열어둔 흔한 상황에선
    아예 안 띄우는 게 가장 안전하다 — 20분 뒤 다음 실행이 재시도하면 된다.
    """
    try:
        with open(path, "r+b"):
            return False
    except OSError:
        return True


# 잠긴 파일이면 Excel 을 아예 띄우지 않는다 (팝업·고아 EXCEL.EXE 방지).
# ⚠ 이 시점엔 openpyxl 저장이 이미 끝나 수식 캐시가 빈 상태다 — 그 사실을 상태에 남긴다.
if is_locked(output_file):
    print(f"[SKIP] Auto 파일이 잠겨 있어 재계산을 건너뜁니다: {output_file.name}")
    write_status("실패", "Auto 파일 잠김 — 재계산 미실행(수식 캐시 빈 상태). 다음 실행 시 재시도",
                 rows=len(data_rows))
    exit(0)


# ※ 재시도 이유: Excel 인스턴스가 재계산 도중 죽으면 이후 호출이
#   com_error(-2147417848, RPC_E_DISCONNECTED '호출된 개체가 연결이 끊겼습니다') 로 실패한다.
#   프로세스 단위 사고라 같은 인스턴스로는 복구가 안 되고, 새 인스턴스로 다시 열면 대개 성공한다.
#   여기서 끝까지 실패하면 Auto 파일은 **수식 캐시가 빈 상태**로 남으므로 그 사실을 크게 알린다.
for attempt in range(1, COM_RETRIES + 1):
    try:
        recalc_and_save(output_file)
    except pywintypes.com_error as e:
        print(f"[재시도 {attempt}/{COM_RETRIES}] Excel COM 실패: {e}")
        if attempt == COM_RETRIES:
            print(f"[실패] Excel 재계산·저장이 {COM_RETRIES}회 모두 실패했습니다.")
            print(f"        {output_file.name} 은 지금 **수식 캐시가 빈 상태**입니다 — "
                  f"Excel 로 한 번 열었다가 저장하거나 이 스크립트를 다시 실행하세요.")
            print(f"        (마커를 기록하지 않았으므로 다음 실행이 같은 소스를 재처리합니다)")
            write_status("실패", "Excel 재계산·저장 실패 — Auto 파일의 수식 캐시가 빈 상태",
                         rows=len(data_rows))
            raise
        time.sleep(COM_RETRY_WAIT_SEC)
        continue

    # 저장 직후 재검증 — 마커는 여기를 통과했을 때만 기록한다.
    # ※ 저장이 외부(Excel/OneDrive)에 의해 되돌려진 경우를 즉시 드러내기 위한 단계.
    # ※ Excel 이 막 놓은 파일을 곧바로 읽으면 PermissionError 가 난다 — 그건 '되돌려짐' 이
    #   아니라 단순 타이밍이므로 **읽기 실패 사유일 때만** 잠시 기다렸다 다시 본다.
    #   값 불일치·잔재는 재시도해도 그대로라 즉시 판정한다.
    for _ in range(VERIFY_RETRIES):
        ok, why = target_is_saved(output_file, source_file.name, src_data)
        if ok or "읽을 수 없음" not in why:
            break
        time.sleep(VERIFY_WAIT_SEC)
    if ok:
        LAST_SOURCE_FILE.write_text(current_marker, encoding="utf-8")
        print(f"[완료] {output_file.name} 저장 완료")
        write_status("갱신 완료", rows=len(data_rows))
    else:
        print(f"[경고] 저장 직후 검증 실패 — {why}")
        write_status("경고 — 저장이 되돌려짐", why, rows=len(data_rows))
        print("        다른 프로그램(Excel/OneDrive)이 파일을 되돌렸을 수 있습니다.")
        print("        마커를 기록하지 않았으므로 다음 실행이 다시 처리합니다.")
    break
