"""
mail_search_to_msg.py
2026-04-30  Jonghyun Park w/ Claude

team_name 메일함에서 키워드 매칭되는 메일을 .msg + 첨부파일로 다운로드.

검색 대상:
  - 제목(Subject) 또는 본문(Body) 부분일치 (대소문자 무관)
  - KEYWORDS 리스트 — 어느 하나라도 포함되면 매칭 (OR)

저장 위치 (.msg 와 첨부 모두 같은 폴더):
  C:\\Users\\<user>\\Downloads\\mail_search_<YYMMDD>\\

저장 파일명:
  - 메일:   <YYMMDD_HHMM>_<safe subject>.msg
  - 첨부:   <YYMMDD_HHMM>_<원본 첨부 파일명>
  → 같은 메일에서 나온 .msg + 첨부가 날짜 prefix로 묶여 정렬됨.

같은 날짜 폴더에서 키워드 바꿔가며 재실행 가능:
  - 메일 dedup: EntryID 를 _processed_entry_ids.txt 에 기록 → 재실행 시 같은 메일 skip
  - 첨부 dedup: SAVE_DIR 안의 기존 첨부파일명에서 원본명 추출하여 set 구성
                → 다른 메일이 매칭됐어도 같은 원본명 첨부는 skip
  - 강제 재저장: _processed_entry_ids.txt 삭제 + 기존 첨부파일들도 삭제 후 실행

사용:
  스크립트 상단 ── 설정 ── 섹션에서 KEYWORDS 등을 바꾼 뒤 실행.
    python mail_search_to_msg.py
"""

import re
import win32com.client
from pathlib import Path
from datetime import datetime

# ── 설정 ────────────────────────────────────────────────────────
# 검색할 키워드 — 제목 또는 본문에 어느 하나라도 포함되면 매칭 (OR, 대소문자 무관)
KEYWORDS = [
    "campaign_name",
    # "ai",
    # "추가 키워드 ...",
]

# Outlook 메일함 이름들 (DisplayName 부분 일치). 여러 개 박으면 다 검색.
# 같은 메일이 여러 메일함에 동시 수신된 경우 — InternetMessageID 로 dedup 되어 한 번만 저장 (아래 참고).
STORE_NAMES = [
    "team_name",
    # "your.email@example.com",
]

# 온라인 보관(아카이브) store 도 함께 검색? (디폴트 True — 개인 mailbox + 그 아카이브 둘 다)
#   개인 mailbox 의 오래된 메일은 '온라인 보관 - <이메일>' 아카이브로 이동돼 있어,
#   기본으로 아카이브까지 봐야 옛 메일이 누락되지 않는다. False 면 개인 mailbox 만.
INCLUDE_ARCHIVE = True
# 공용 폴더(Public Folders) store 는 검색에서 제외 (개인 mailbox 와 이름이 substring 으로 겹쳐 오매칭 방지)
SKIP_PUBLIC_FOLDERS = True

# 검색 대상 폴더 — None 이면 받은편함(Inbox)
FOLDER_NAME = None

# 하위 폴더까지 재귀 검색?
RECURSE_SUBFOLDERS = False

# 본문(Body)도 검색? False면 제목(Subject)만 검색 — 빠름
SEARCH_BODY = True

# 단어 단위(whole-word) 매칭? True 면 'ai'가 'email'/'available' 안에서 매칭 안 됨 (\b 경계 사용)
# False 면 단순 substring 매칭 (이전 동작 — 짧은 키워드 시 과매칭 위험)
WHOLE_WORD = True

# 매칭 메일의 첨부파일도 같은 폴더에 저장? (.msg 파일과 같은 위치)
SAVE_ATTACHMENTS = True

# 서명·인라인 이미지(image001.png 등) 자동 skip? — Outlook 자동 생성 이름 패턴 매칭
SKIP_INLINE_IMAGES = True

# ── 자동 결정 ───────────────────────────────────────────────────
TODAY = datetime.now().strftime("%y%m%d")
SAVE_DIR = Path.home() / "Downloads" / f"mail_search_{TODAY}"

# 같은 날짜 폴더에서 키워드 바꿔가며 재실행해도 같은 메일 중복 저장 방지용
# 처리한 메일의 EntryID를 한 줄씩 기록. 강제 재저장 원하면 이 파일 삭제 후 실행.
PROCESSED_MARKER = SAVE_DIR / "_processed_entry_ids.txt"

# Outlook 상수
OL_FOLDER_INBOX = 6   # GetDefaultFolder
OL_CLASS_MAIL   = 43  # MailItem
OL_SAVE_AS_MSG  = 3   # SaveAs Type

# Outlook이 자동 생성하는 인라인 이미지 이름 패턴 (서명 등)
_INLINE_IMG_PAT = re.compile(r'^image\d+\.(png|jpg|jpeg|gif|bmp)$', re.IGNORECASE)


def safe_filename(name: str) -> str:
    """Windows 파일명 허용 문자만 남기고 길이 제한."""
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name or "")
    cleaned = cleaned.strip(". ")
    return cleaned[:150] or "no_subject"


def unique_path(save_dir: Path, base_name: str, seen: set) -> Path:
    """저장 폴더 내 중복 회피. 같은 이름이면 ' (2)', ' (3)' ... 부여 (확장자 보존)."""
    name = base_name
    counter = 1
    while name.lower() in seen or (save_dir / name).exists():
        counter += 1
        stem, dot, ext = base_name.rpartition('.')
        if dot:
            name = f"{stem} ({counter}).{ext}"
        else:
            name = f"{base_name} ({counter})"
    seen.add(name.lower())
    return save_dir / name


def load_processed_ids(marker_path: Path) -> set:
    """마커 파일에서 처리한 EntryID 집합 로드. 파일 없으면 빈 집합."""
    if not marker_path.exists():
        return set()
    with open(marker_path, encoding="utf-8") as f:
        return {line.strip() for line in f if line.strip()}


def append_processed_id(marker_path: Path, entry_id: str) -> None:
    """EntryID를 마커 파일에 1줄 append (즉시 flush — 중간 종료에도 안전)."""
    with open(marker_path, "a", encoding="utf-8") as f:
        f.write(entry_id + "\n")


# 저장된 첨부파일에서 <YYMMDD_HHMM>_ prefix와 (N) counter suffix 제거하여 원본명 추출
_DATE_PREFIX_PAT = re.compile(r"^\d{6}_\d{4}_(.+)$")
_COUNTER_SUFFIX_PAT = re.compile(r" \(\d+\)$")


def extract_attachment_original(filename: str) -> str | None:
    """저장된 첨부 파일명에서 원본 첨부 파일명 복원.
       '260415_0903_report (2).xlsx' → 'report.xlsx'
       prefix 매칭 안 되면 None.
    """
    p = Path(filename)
    m = _DATE_PREFIX_PAT.match(p.stem)
    if not m:
        return None
    original_stem = _COUNTER_SUFFIX_PAT.sub("", m.group(1))
    return original_stem + p.suffix


def scan_saved_attachments(save_dir: Path) -> set:
    """저장 폴더 내 기존 첨부파일들의 원본명 집합 (소문자) 반환.
       .msg 파일과 marker(_*) 파일은 제외.
    """
    if not save_dir.exists():
        return set()
    found = set()
    for f in save_dir.iterdir():
        if not f.is_file():
            continue
        if f.name.startswith("_"):
            continue
        if f.suffix.lower() == ".msg":
            continue
        original = extract_attachment_original(f.name)
        if original:
            found.add(original.lower())
    return found


def _compile_keyword_patterns():
    """KEYWORDS를 매칭용 검사 함수 리스트로 변환.
       WHOLE_WORD=True 이면 \\b 경계 regex 사용 — 'ai'가 'email' 안에서 매칭 안 됨.
       WHOLE_WORD=False 이면 단순 substring (lowercase) 매칭.
    """
    checkers = []
    for kw in KEYWORDS:
        if WHOLE_WORD:
            pat = re.compile(rf"\b{re.escape(kw)}\b", re.IGNORECASE)
            checkers.append(pat.search)
        else:
            kw_lower = kw.lower()
            checkers.append(lambda text, k=kw_lower: k in text)
    return checkers


_KEYWORD_CHECKERS = _compile_keyword_patterns()


def matches_keywords(subject: str, body: str) -> bool:
    text = subject if WHOLE_WORD else subject.lower()
    if SEARCH_BODY:
        text = text + "\n" + (body if WHOLE_WORD else body.lower())
    return any(check(text) for check in _KEYWORD_CHECKERS)


def iter_folders(root, recurse: bool):
    """폴더 + (옵션) 하위 폴더 재귀 yield."""
    yield root
    if recurse:
        for sub in root.Folders:
            yield from iter_folders(sub, True)


def _resolve_stores(ns, name: str) -> list:
    """name(부분일치)에 매칭되는 Outlook store 들을 반환 (개인 mailbox 우선, 아카이브는 맨 뒤).
    - SKIP_PUBLIC_FOLDERS 면 공용 폴더 store 제외
    - INCLUDE_ARCHIVE=False 면 온라인 보관(아카이브) store 제외 → 개인 mailbox 만
    - INCLUDE_ARCHIVE=True 면 개인 mailbox + 그 아카이브 둘 다 (온라인 보관 DisplayName 이
      개인 mailbox 이메일을 통째로 포함하므로 같은 name 으로 둘 다 잡힌다)
    DisplayName 정확일치 store 를 맨 앞에 둔다."""
    key = (name or "").lower()
    exact, mains, archives = [], [], []
    for store in ns.Stores:
        try:
            dn = store.DisplayName or ""
        except Exception:
            continue
        low = dn.lower()
        is_archive = ("온라인 보관" in low or "archive" in low)
        is_public = ("공용 폴더" in low or "public folders" in low)
        if is_public and SKIP_PUBLIC_FOLDERS:
            continue
        if is_archive and not INCLUDE_ARCHIVE:
            continue
        if key and key not in low:
            continue
        if low == key:
            exact.append(store)
        elif is_archive:
            archives.append(store)
        else:
            mains.append(store)
    return exact + mains + archives


def find_stores(ns, store_names: list[str]):
    """STORE_NAMES 각 이름을 _resolve_stores 로 확장(개인 mailbox + 아카이브 포함)해
    (DisplayName, store) 튜플 list 반환. StoreID 로 dedup — 같은 store 중복 방지."""
    results: list[tuple[str, object]] = []
    seen_ids: set[str] = set()
    for nm in store_names:
        matched = _resolve_stores(ns, nm)
        if not matched:
            print(f"  ⚠️ 메일함 못 찾음 — {nm!r}  (skip)")
            continue
        for s in matched:
            sid = getattr(s, "StoreID", None) or s.DisplayName
            if sid in seen_ids:
                continue
            seen_ids.add(sid)
            results.append((s.DisplayName, s))
    return results


# Outlook 의 PR_INTERNET_MESSAGE_ID — RFC 5322 Message-ID (globally unique 메일 식별자)
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"


def get_message_id(mail) -> str | None:
    """메일의 InternetMessageID 반환. 없거나 빈 값이면 None.
    같은 메일이 여러 store/folder 에 있어도 동일 → store 교차 dedup 의 핵심 키."""
    try:
        v = mail.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID)
    except Exception:
        return None
    if not v:
        return None
    v = str(v).strip()
    return v or None


def find_folder(store, folder_name: str | None):
    """folder_name이 None이면 받은편함, 아니면 루트 하위에서 이름 일치 폴더."""
    if not folder_name:
        return store.GetDefaultFolder(OL_FOLDER_INBOX)
    root = store.GetRootFolder()
    for f in root.Folders:
        if f.Name == folder_name:
            return f
    raise RuntimeError(f"폴더를 찾을 수 없습니다: {folder_name}")


def main():
    print(f"[키워드] {KEYWORDS}  (OR 매칭, 대소문자 무관)")
    print(f"[매칭 단위] {'단어 경계(\\b)' if WHOLE_WORD else 'substring'}")
    print(f"[검색 범위] 제목{' + 본문' if SEARCH_BODY else ' (본문 미검색)'}")
    print(f"[메일함] {STORE_NAMES}")
    print(f"[저장]   {SAVE_DIR}")
    print()

    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")

    stores = find_stores(ns, STORE_NAMES)
    if not stores:
        raise RuntimeError(f"메일함 하나도 못 찾음: {STORE_NAMES}")
    print(f"[발견] {len(stores)} 메일함:")
    for nm, _ in stores:
        print(f"   - {nm}")
    print()

    SAVE_DIR.mkdir(parents=True, exist_ok=True)

    # dedup 마커 — InternetMessageID (msgid:...) 우선, EntryID (entry:...) fallback, legacy prefix-less entry 도 호환
    processed_ids = load_processed_ids(PROCESSED_MARKER)
    saved_att_originals = scan_saved_attachments(SAVE_DIR)
    print(f"[중복방지] 이미 처리한 메일 마커 {len(processed_ids):,}개 / 첨부 원본명 {len(saved_att_originals):,}개")
    print()

    saved = 0
    saved_atts = 0
    failed = 0
    failed_atts = 0
    skipped_inline = 0
    skipped_dup = 0
    skipped_att_dup = 0
    seen_names = set()

    for store_name, target_store in stores:
        try:
            target_folder = find_folder(target_store, FOLDER_NAME)
        except Exception as e:
            # RuntimeError (custom) + pywintypes.com_error (Public Folders 등 Inbox 없는 store) 둘 다 catch
            print(f"[메일함 '{store_name}'] 폴더 못 찾음 — {e}  (skip)")
            continue
        print(f"\n=== 메일함: {store_name} ===")

        for folder in iter_folders(target_folder, RECURSE_SUBFOLDERS):
            items = folder.Items
            try:
                items.Sort("[ReceivedTime]", True)  # 최신 순
            except Exception:
                pass

            total = items.Count
            print(f"▶ 폴더 '{folder.Name}' — {total}개 검색 중...")

            for idx, mail in enumerate(items, 1):
                if idx % 500 == 0:
                    print(f"    진행 {idx}/{total} (저장 {saved}, 실패 {failed})")
                try:
                    if mail.Class != OL_CLASS_MAIL:
                        continue
                    subject = mail.Subject or ""
                    body = mail.Body if SEARCH_BODY else ""
                except Exception:
                    continue

                if not matches_keywords(subject, body):
                    continue

                # dedup 키 추출 — Message-ID (전 store 공통) 우선, EntryID (store local) fallback
                msg_id = get_message_id(mail)
                try:
                    entry_id = mail.EntryID
                except Exception:
                    entry_id = None

                # 이미 처리한 메일? msgid / entry / legacy entry (prefix 없음) 어느 하나라도 매칭이면 skip
                already = False
                if msg_id and f"msgid:{msg_id}" in processed_ids:
                    already = True
                elif entry_id and f"entry:{entry_id}" in processed_ids:
                    already = True
                elif entry_id and entry_id in processed_ids:
                    already = True   # legacy 마커 (prefix 없는 EntryID)
                if already:
                    skipped_dup += 1
                    continue

                # 파일명: <YYMMDD_HHMM>_<safe subject>.msg
                try:
                    received = mail.ReceivedTime
                    date_prefix = received.strftime("%y%m%d_%H%M")
                except Exception:
                    date_prefix = "unknown"

                msg_base = f"{date_prefix}_{safe_filename(subject)}.msg"
                dest = unique_path(SAVE_DIR, msg_base, seen_names)
                try:
                    mail.SaveAs(str(dest), OL_SAVE_AS_MSG)
                    print(f"    [저장] {dest.name}")
                    saved += 1
                except Exception as e:
                    print(f"    [실패] {subject[:50]} → {e}")
                    failed += 1
                    continue  # .msg 저장 실패 시 첨부도 skip + 마커 미기록 (다음 실행에 재시도)

                # 첨부파일도 같은 폴더에 <YYMMDD_HHMM>_<원본명> 으로 저장
                if SAVE_ATTACHMENTS:
                    try:
                        atts = mail.Attachments
                    except Exception:
                        atts = None
                    if atts:
                        for att in atts:
                            try:
                                att_name = att.FileName or ""
                            except Exception:
                                continue
                            if not att_name:
                                continue
                            if SKIP_INLINE_IMAGES and _INLINE_IMG_PAT.match(att_name):
                                skipped_inline += 1
                                continue
                            # 첨부 원본명 dedup — 다른 메일에 같은 이름 첨부 있으면 skip
                            safe_att = safe_filename(att_name)
                            att_key = safe_att.lower()
                            if att_key in saved_att_originals:
                                skipped_att_dup += 1
                                continue
                            att_base = f"{date_prefix}_{safe_att}"
                            att_dest = unique_path(SAVE_DIR, att_base, seen_names)
                            try:
                                att.SaveAsFile(str(att_dest))
                                print(f"      [첨부] {att_dest.name}")
                                saved_atts += 1
                                saved_att_originals.add(att_key)
                            except Exception as e:
                                print(f"      [첨부실패] {att_name} → {e}")
                                failed_atts += 1

                # .msg 저장 성공 시 마커에 기록 — msgid (전 store 공통) + entry (store local fallback) 둘 다
                for key in (
                    f"msgid:{msg_id}" if msg_id else None,
                    f"entry:{entry_id}" if entry_id else None,
                ):
                    if key and key not in processed_ids:
                        processed_ids.add(key)
                        append_processed_id(PROCESSED_MARKER, key)

    print()
    print(f"완료 — .msg {saved}개 / 첨부 {saved_atts}개")
    print(f"  skip: 메일 중복 (msgid/entry) {skipped_dup} / 첨부 원본명 중복 {skipped_att_dup} / 인라인 이미지 {skipped_inline}")
    print(f"  실패: {failed + failed_atts}개")
    print(f"위치: {SAVE_DIR}")


if __name__ == "__main__":
    main()
