"""
check_mail_attachment_byname.py
2026-07-22  Jonghyun Park w/ Claude  — 중복 suffix _yymmdd → _yymmdd_HHMM
2026-08-20  Jonghyun Park w/ Claude  — 수신일시 '항상' 부착 + ATTACHMENT_KEYS 원복
2026-08-20  Jonghyun Park w/ Claude  — RULES 리스트 구조로 전환, Monitoring 흡수
                                       (check_mail_attachment_monitoring.py 폐기)
2026-08-20  Jonghyun Park w/ Claude  — Monitoring 저장폴더를 `1.고객 법인 일정 파일` 로 변경
2026-08-20  Jonghyun Park w/ Claude  — 수신일시를 파일명 **맨 앞**으로 (이름순=도착순)

team_name mailbox 받은편함을 한 번만 순회하면서, RULES 에 정의된 조건별로
첨부파일을 각자의 폴더에 저장한다.
- RULES 항목 1개 = 저장폴더 1개 = 마커파일 1개 (룰끼리 완전히 독립)
- received_from ~ received_to 사이에 받은 메일만 처리 (이전 캠페인 옛 메일 제외)
- 이미 처리한 메일은 EntryID로 기록해 재처리 방지 (마커는 룰별로 따로)
- 저장 파일명 **맨 앞**에 수신일시(yymmdd_HHMM_)를 **항상** 붙임
  ※ 맨 앞이라 탐색기 이름순 정렬 = 도착순 정렬. 폴더만 열어봐도 최신본이 맨 아래에 온다
  ※ 일정 폴더는 update_schedule_summary.py 가 읽는 소스 폴더 — 최신 판정에 도착 순서가 필요
- SaveAsFile은 MAX_PATH 제한으로 임시폴더에 저장 후 shutil.move로 이동

── 왜 '항상' 붙이나 (2026-08-20) ────────────────────────────────
종전엔 같은 파일명이 이미 있을 때만 수신일시를 덧붙였다. 그래서 이름이 서로 다르면
파일명에 도착 순서 정보가 전혀 없었다. 실제 사례 — 2026-08-19 에 고객이

    26 CAMPAIGN NAME Monitoring_260819_v1.xlsb        (15:48 수신)
    26 CAMPAIGN NAME Monitoring_260819_shared.xlsb    (20:08 수신, 이게 최신본)

두 개를 순서대로 보냈는데, 정렬 키가 `(260819, 0, 0.0, 0, 0)` 로 **완전히 동일**해
어느 쪽이 최신인지 판정 자체가 불가능했다. 같은 패턴이 일정 파일에 오면 더 나쁘다 —
`_shared` 는 `_vN` 이 없어 `ver_int=0` 이라 `_v1`(=1) 에 밀려서 **나중에 온 파일이 통째로
무시**된다. 지금은 저장 시 `260819_1548_..._v1` / `260819_2008_..._shared` 처럼 **맨 앞에** 도착 시각이
박히므로 순서가 파일명만으로 확정되고, 탐색기 이름순 정렬만으로도 도착순이 보인다.
자세한 경위·검증은 README.md 의 '최신 파일 선택 기준 › v1.2' 참조.

── 왜 스크립트를 하나로 합쳤나 (2026-08-20) ──────────────────────
Monitoring 첨부를 받으려고 check_mail_attachment_monitoring.py 를 따로 뒀지만,
Outlook 받은편함을 스크립트 수만큼 중복 순회하고 스케줄러 작업도 그만큼 늘어난다.
같은 메일함을 보는 조건은 RULES 한 줄로 표현하는 게 맞아 통합했다.
(gro / url 변형은 아직 별도 유지 — 필요해지면 같은 방식으로 RULES 에 얹으면 된다.)

⚠ Monitoring 조건을 일정 룰의 attachment_keys 에 얹지 말 것 — 저장폴더가 같아진
   지금도 **반드시 별개 룰**이어야 한다. attachment_keys 는 AND 매칭이라 한 룰에 합치면
   `CAMPAIGN NAME + Campaign + Monitoring` 을 모두 만족해야 해서, 정작 일정 파일
   (`...Campaign Schedule...`)이 조건에서 탈락해 **수집이 통째로 멎는다**
   (2026-08-19 실제 발생). 확장자(.xlsx vs .xlsb)와 마커도 룰마다 달라야 한다.
"""

import win32com.client
import shutil
import tempfile
from pathlib import Path
from datetime import datetime, date

# ════════ 사용자가 바꿔야 하는 부분 ════════
STORE_NAME = "team_name mailbox"

# 캠페인 폴더 루트 — 룰들의 저장폴더가 여기(또는 하위)에 붙는다
CAMPAIGN_BASE = Path(
    r"C:\Users\user_name\OneDrive - company_name"
    r"\Project_team_name - 1 company_name - 02 part_name"
    r"\part_name\2026\# CAMPAIGN_PROJECTS\03. CAMPAIGN NAME"
    r"\01. SCHEDULE"
)

# 고객 첨부(일정 · Monitoring)를 받는 폴더. update_schedule_summary.py 가 이 폴더에서
# 최신 1개를 일정 소스로 집어간다 — 확장자(.xlsx/.xlsb)·제목 형태와 무관하게 **가장 늦게 도착한 것**.
# 2026-08-20: 고객이 일정 내용을 Qualitative Monitoring 파일로 보내기 시작해, 소비 쪽
#    (update_schedule_summary.py) 의 SOURCE_EXTS 에 .xlsb, SOURCE_NAME_KEYS 에 monitoring 을 추가했다.
#    즉 Monitoring 파일은 더 이상 '배제 대상' 이 아니라 정상적인 일정 소스 후보다.
#    ⚠ 두 룰의 마커·attachment_keys 분리는 그대로 유지할 것 (아래 RULES 주석 참조).
CUSTOMER_FILE_FOLDER = CAMPAIGN_BASE / "1.고객 법인 일정 파일"

# 한 첨부가 여러 룰에 걸릴 때 처리 방식
#   True  = 위에서부터 처음 걸린 룰 1개만 적용 (한 파일 = 한 번만 저장)
#   False = 걸리는 룰 전부에 저장
# ⚠ True 를 권장. 예) "2026 CAMPAIGN NAME Campaign Schedule_20260804_v1(Monitoring기반제작).xlsx"
#    는 일정 룰(CAMPAIGN NAME+Campaign)과 Monitoring 룰(CAMPAIGN NAME+Monitoring)에 **둘 다** 걸린다.
#    두 룰의 저장폴더가 같아진 지금 False 로 두면 같은 파일을 두 번 저장하려다
#    두 번째가 '이미 처리됨' 스킵으로 빠진다(무해하지만 로그가 지저분).
FIRST_MATCH_ONLY = True

# ── 수집 룰 ──────────────────────────────────────────────────
# name            : 로그 표기용 이름
# save_folder     : 저장 폴더 (미리 존재해야 함)
# subject_keys    : 메일제목에 '모두' 포함돼야 통과. [] = 제목 필터 없음
# attachment_keys : 첨부파일명에 '모두' 포함돼야 통과 (소문자 매칭)
#                   리스트 안의 list = OR 그룹 → ["Campaign","캠페인"] = 영/한 둘 다 허용
# allowed_exts    : 저장할 확장자 (소문자 튜플)
# received_from   : 이 날짜(받은날짜) 이상만 처리 — 캠페인 바뀌면 갱신
# received_to     : 이 날짜 이하만 처리. None = 상한 없음
# marker_file     : 처리한 EntryID 기록 파일 (룰별로 반드시 다른 파일)
RULES = [
    {
        "name":            "일정",
        "save_folder":     CUSTOMER_FILE_FOLDER,
        "subject_keys":    [],
        "attachment_keys": ["CAMPAIGN NAME", ["Campaign", "캠페인"]],
        "allowed_exts":    (".xlsx",),
        "received_from":   date(2026, 6, 18),
        "received_to":     None,
        "marker_file":     Path(r"C:\Users\user_name\Documents\campaign_mail_processed_ids.txt"),
    },
    {
        # 고객이 회차마다 xlsb/xlsx 를 오락가락 보내므로 확장자 2종 허용
        # (260804 은 xlsx, 260819 는 xlsb)
        "name":            "Monitoring",
        "save_folder":     CUSTOMER_FILE_FOLDER,
        "subject_keys":    [],
        "attachment_keys": ["CAMPAIGN NAME", "Monitoring"],
        "allowed_exts":    (".xlsb", ".xlsx"),
        "received_from":   date(2026, 6, 18),
        "received_to":     None,
        # 마커는 소스 폴더가 아니라 캠페인 루트에 둔다 (고객 파일 폴더를 깨끗하게 유지)
        "marker_file":     CAMPAIGN_BASE / "_monitoring_processed_ids.txt",
    },
]

# ════════ 내부 사용 ════════

# ── 긴 경로 존재 확인 (\\?\ 접두사 사용) ─────────────────────
def long_path_exists(p: Path) -> bool:
    lp = Path("\\\\?\\" + str(p.resolve()))
    return lp.exists()

# ── 키워드 매칭 (리스트 안의 list = OR 그룹) ──────────────────
def keys_match(text: str, keys) -> bool:
    t = text.lower()
    return all(
        any(alt.lower() in t for alt in (k if isinstance(k, list) else [k]))
        for k in keys
    )

# ── 룰 초기화: 저장폴더 확인 + 처리된 EntryID 로드 ────────────
for rule in RULES:
    folder = rule["save_folder"]
    if not folder.exists():
        raise RuntimeError(f"[{rule['name']}] 저장 폴더가 없습니다: {folder}")
    marker = rule["marker_file"]
    rule["processed_ids"] = (
        set(marker.read_text(encoding="utf-8").splitlines()) if marker.exists() else set()
    )
    rule["new_ids"] = []
    rule["saved"]   = 0
    rule["skipped"] = 0

# ── Outlook 연결 ──────────────────────────────────────────────
outlook = win32com.client.Dispatch("Outlook.Application")
ns      = outlook.GetNamespace("MAPI")

target_store = None
for store in ns.Stores:
    if STORE_NAME.lower() in store.DisplayName.lower():
        target_store = store
        break

if target_store is None:
    raise RuntimeError(f"메일함을 찾을 수 없습니다: {STORE_NAME}")

inbox = target_store.GetDefaultFolder(6)  # 6 = 받은편함
print(f"[메일함] {target_store.DisplayName} / 받은편함 ({inbox.Items.Count}개)")
print(f"[룰] {', '.join(r['name'] for r in RULES)}")

# ── 메일 순회 (한 번만 돌면서 모든 룰 적용) ───────────────────
for mail in inbox.Items:
    try:
        entry_id = mail.EntryID
        rt       = mail.ReceivedTime
    except Exception:
        continue

    rt_date = date(rt.year, rt.month, rt.day)
    subject = getattr(mail, "Subject", "") or ""

    # 이 메일에 적용할 룰 추리기 (받은날짜 · 메일제목 · 기처리 EntryID)
    active = []
    for rule in RULES:
        if rt_date < rule["received_from"]:
            continue
        if rule["received_to"] and rt_date > rule["received_to"]:
            continue
        if rule["subject_keys"] and not keys_match(subject, rule["subject_keys"]):
            continue
        if entry_id in rule["processed_ids"]:
            rule["skipped"] += 1
            continue
        active.append(rule)

    if not active:
        continue

    # 파일명 앞에 붙일 수신일시. 날짜만(yymmdd) 쓰면 같은 날 두 번 온 파일이 이름 충돌로
    # 스킵되어 유실 + 최신 판정이 동점 → 시각(_HHMM)까지 (2026-07-22)
    received = datetime.strftime(rt, "%y%m%d_%H%M")

    for att in mail.Attachments:
        name   = att.FileName
        suffix = Path(name).suffix

        for rule in active:
            if suffix.lower() not in rule["allowed_exts"]:
                continue
            if rule["attachment_keys"] and not keys_match(name, rule["attachment_keys"]):
                continue

            # 수신일시를 파일명 **맨 앞**에 붙인다 (2026-08-20)
            # → 탐색기에서 이름순 정렬 = 도착순 정렬이 된다
            dest = rule["save_folder"] / f"{received}_{Path(name).stem}{suffix}"

            # 같은 이름이 이미 있으면 = 같은 메일을 다시 도는 것 (같은 분 단위 시각) → 스킵
            if long_path_exists(dest):
                print(f"[스킵][{rule['name']}] 이미 처리됨: {dest.name}")
                rule["skipped"] += 1
                # 이미 저장된 것으로 간주 → EntryID 기록
                if entry_id not in rule["new_ids"]:
                    rule["new_ids"].append(entry_id)
            else:
                # 임시폴더 저장 후 이동
                with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
                    tmp_path = Path(tmp.name)
                att.SaveAsFile(str(tmp_path))
                shutil.move(str(tmp_path), str(dest))
                print(f"[저장][{rule['name']}] {name} → {dest.name} (수신 {received})")

                rule["saved"] += 1
                if entry_id not in rule["new_ids"]:
                    rule["new_ids"].append(entry_id)

            if FIRST_MATCH_ONLY:
                break

# ── 처리된 EntryID 저장 + 결과 요약 ───────────────────────────
print()
for rule in RULES:
    if rule["new_ids"]:
        with rule["marker_file"].open("a", encoding="utf-8") as f:
            f.write("\n".join(rule["new_ids"]) + "\n")
    print(f"완료 [{rule['name']}] - 저장 {rule['saved']}개 / 스킵 {rule['skipped']}개")
