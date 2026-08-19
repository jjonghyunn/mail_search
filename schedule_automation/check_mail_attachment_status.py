"""
check_mail_attachment_status.py
2026-07-22  Jonghyun Park w/ Claude  — 중복 suffix _yymmdd → _yymmdd_HHMM
2026-08-19  Jonghyun Park w/ Claude  — public repo 공개판 (설정값 placeholder 화)

team_name mailbox 받은편함에서 ATTACHMENT_KEYS 조건에 맞는 xlsx 첨부파일을 감지해 저장.
- 이미 처리한 메일은 EntryID로 기록해 재처리 방지
- 같은 파일명이 있으면 수신일시(_yymmdd_HHMM) 붙여 저장
- SaveAsFile은 MAX_PATH 제한으로 임시폴더에 저장 후 shutil.move로 이동
"""

import win32com.client
import shutil
import tempfile
from pathlib import Path
from datetime import datetime, date

SAVE_FOLDER  = Path(
    r"C:\Users\user_name\OneDrive - company_name"
    r"\Project_team_name - 1 company_name - 02 part_name"
    r"\part_name\2026\# CAMPAIGN_PROJECTS\03. CAMPAIGN NAME"
    r"\02. SCHEDULE\2.Status"
)
STORE_NAME   = "team_name mailbox"
ATTACHMENT_KEYS = ["Status", "campaign name"]  # 첨부파일명에 모두 포함되어야 통과 (소문자 매칭)
RECEIVED_FROM   = date(2026, 6, 18)         # 이 날짜(받은날짜) 이상 메일만 처리 — 캠페인 바뀌면 갱신
MARKER_FILE     = SAVE_FOLDER / "_processed_ids.txt"

# ── 긴 경로 존재 확인 (\\?\ 접두사 사용) ──────────────────────
def long_path_exists(p: Path) -> bool:
    lp = Path("\\\\?\\" + str(p.resolve()))
    return lp.exists()

# ── 처리된 EntryID 로드 ───────────────────────────────────────
if MARKER_FILE.exists():
    processed_ids = set(MARKER_FILE.read_text(encoding="utf-8").splitlines())
else:
    processed_ids = set()

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

# ── 메일 순회 ────────────────────────────────────────────────
saved = 0
skipped = 0
new_ids = []

for mail in inbox.Items:
    try:
        entry_id = mail.EntryID
        rt       = mail.ReceivedTime
    except Exception:
        continue

    if date(rt.year, rt.month, rt.day) < RECEIVED_FROM:
        continue

    if entry_id in processed_ids:
        skipped += 1
        continue

    mail_saved = False
    # 중복 파일명에 붙일 수신일시. 날짜만(_yymmdd) 쓰면 같은 날 두 번 온 파일이
    # 이름 충돌로 스킵되어 유실 → 시각(_HHMM)까지 붙여 분리 (2026-07-22)
    received = datetime.strftime(rt, "%y%m%d_%H%M")

    for att in mail.Attachments:
        name = att.FileName
        if not name.lower().endswith(".xlsx"):
            continue
        if ATTACHMENT_KEYS and not all(k.lower() in name.lower() for k in ATTACHMENT_KEYS):
            continue

        dest = SAVE_FOLDER / name

        # 이미 같은 파일명 존재 → 날짜 붙인 이름으로
        if long_path_exists(dest):
            stem     = Path(name).stem
            suffix   = Path(name).suffix
            new_name = f"{stem}_{received}{suffix}"
            dest     = SAVE_FOLDER / new_name

            # 날짜 붙인 파일도 이미 존재 → 완전 스킵
            if long_path_exists(dest):
                print(f"[스킵] 이미 처리됨: {new_name}")
                skipped += 1
                mail_saved = True   # 이미 저장된 것으로 간주 → EntryID 기록
                continue

            # 임시폴더 저장 후 이동
            with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
                tmp_path = Path(tmp.name)
            att.SaveAsFile(str(tmp_path))
            shutil.move(str(tmp_path), str(dest))
            print(f"[저장-중복] {name} → {new_name} (수신일: {received})")
        else:
            # 임시폴더 저장 후 이동
            suffix = Path(name).suffix
            with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
                tmp_path = Path(tmp.name)
            att.SaveAsFile(str(tmp_path))
            shutil.move(str(tmp_path), str(dest))
            print(f"[저장] {name}")

        saved += 1
        mail_saved = True

    if mail_saved:
        new_ids.append(entry_id)

# ── 처리된 EntryID 저장 ───────────────────────────────────────
if new_ids:
    with MARKER_FILE.open("a", encoding="utf-8") as f:
        f.write("\n".join(new_ids) + "\n")

print(f"\n완료 - 저장 {saved}개 / 스킵 {skipped}개")
