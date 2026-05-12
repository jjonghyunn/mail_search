# summarize_msgs.py
# 2026-05-11  Jonghyun Park w/ Claude
"""
지정된 폴더의 .msg 파일들을 읽어서 한 개의 마크다운 요약 리포트를 생성한다.

목적
  · 메일 스레드를 한 눈에 훑으며 요청/이슈/TODO/회신 필요 항목을 빠르게 캐치
  · 시간순으로 정렬해서 누가 언제 무엇을 요청·결정했는지 timeline 으로 본다

휴리스틱 방식 (API 호출 없음)
  · 각 메일에서 ACTION_KEYWORDS 가 포함된 줄을 "액션 아이템 후보" 로 추출
  · 본문은 forwarded thread 시작 지점에서 자르고 BODY_PREVIEW_LINES 까지만 보여줌
  · 인용(>) 라인 제거 + 빈 줄 정리

의존성
  pip install extract-msg

사용
  스크립트 상단 SOURCE_DIR / OUTPUT_DIR 만 본인 환경에 맞게 두고
    python summarize_msgs.py
"""
from __future__ import annotations

import re
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

try:
    import extract_msg
except ImportError:
    print("❌ extract_msg 가 설치되어 있지 않습니다.")
    print("   설치: pip install extract-msg")
    sys.exit(1)

# ════════════════════════════════════════════════════════════════════
# 사용자가 바꿔야 하는 부분
# ════════════════════════════════════════════════════════════════════

# .msg 파일이 있는 폴더 — 본인 환경에 맞게 수정.
# 기본값은 `mail_search_to_msg.py` 가 만드는 Downloads 폴더 패턴.
SOURCE_DIR = Path.home() / "Downloads" / "mail_search_YYMMDD"

# 요약 리포트가 저장될 폴더.
# None 으로 두면 SOURCE_DIR 안에 저장됨 (.msg 들과 같이 묶임 — 추천).
# 별도 경로를 박으면 거기에 저장.
OUTPUT_DIR: Path | None = None

# 출력 파일명 base (timestamp 자동 추가)
# Prefix '_' 는 Windows/일반 파일 매니저에서 영문보다 앞에 정렬 → 폴더 맨 위로 올라옴.
# 다른 prefix 가 좋으면 여기만 바꾸면 됨. (예: "0.summary" / "(00)summary")
OUTPUT_BASENAME = "_summary"

# 액션 아이템 키워드 — 줄 안에 포함되면 "후보" 로 추출됨
ACTION_KEYWORDS = [
    # 요청/부탁
    "요청", "부탁", "공유 부탁", "회신 부탁", "전달 부탁", "확인 부탁",
    "확인 후", "확인해", "확인 바랍", "확인이 필요", "검토 부탁",
    # 이슈/지연/오류
    "이슈", "문제", "오류", "지연", "딜레이", "지체", "어려움",
    # TODO/할 일
    "TODO", "To-do", "to do", "todo", "할 일", "해야",
    # 일정/마감
    "마감", "데드라인", "deadline", "기한", "까지",
    # 결정 필요
    "결정", "협의", "확정", "확정 후", "협의 후",
]

# 본문 미리보기에서 보여줄 최대 줄 수 (forwarded thread 시작점 이전까지)
BODY_PREVIEW_LINES = 25

# 본문 1줄의 최대 길이 (너무 긴 줄은 잘라서 표시)
MAX_LINE_LENGTH = 200

# 발신자 TOP 표시 개수
TOP_SENDERS = 10

# True 면 본문 전체를 리포트에 그대로 포함 (디버그/원문 확인용)
INCLUDE_FULL_BODY = False

# ════════════════════════════════════════════════════════════════════
# 내부 사용
# ════════════════════════════════════════════════════════════════════

def fix_mojibake(text: str) -> str:
    """CP949 로 인코딩된 바이트가 latin-1 으로 잘못 디코딩된 텍스트를 복원.

    extract_msg 가 일부 한글 .msg 파일을 잘못 디코딩하는 경우 대응.
    원래 텍스트보다 복원본의 한글 글자 수가 더 많으면 복원본 채택.
    """
    if not text:
        return text
    try:
        candidate = text.encode("latin-1", errors="strict").decode("cp949", errors="strict")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return text

    def korean_count(s: str) -> int:
        return sum(1 for c in s if "가" <= c <= "힣")

    return candidate if korean_count(candidate) > korean_count(text) else text


# Forwarded thread 시작 신호 — 이 줄을 만나면 본문 미리보기 cut
FORWARD_MARKERS = [
    re.compile(r"^보낸\s*사람\s*[:：]", re.IGNORECASE),
    re.compile(r"^From\s*[:：]", re.IGNORECASE),
    re.compile(r"^발신\s*[:：]", re.IGNORECASE),
    re.compile(r"^-{3,}\s*Original\s*Message\s*-{3,}", re.IGNORECASE),
    re.compile(r"^_{5,}"),                                  # ____________________
    re.compile(r"^={5,}"),
    re.compile(r"^-{20,}"),
]

# 외부 메일 경고 텍스트 — Outlook/Exchange 가 외부 메일 본문 앞에 자동 삽입.
# 본문 시작에 있으면 separator(_____) 와 같이 들어와 FORWARD_MARKERS 에 잘못 잡혀
# 본문 전체가 cut 되는 문제 발생 → clean_body_lines 가 이 블록 먼저 skip 함.
EXTERNAL_WARNING_RE = re.compile(
    r"please\s+be\s+cautious|external\s+email|이\s*메일은\s*외부",
    re.IGNORECASE,
)
# 본문 시작부터 N 줄 안에 외부 경고가 등장하면 그 블록 + 잇따르는 separator/빈줄을 skip
EXTERNAL_WARNING_SCAN_LINES = 12

# 파일명 prefix 에서 날짜 추출용: 260406_0933_subject.msg
FILENAME_DATE_RE = re.compile(r"^(\d{6})_(\d{4})_")

# 인용 줄(>로 시작) 제거용
QUOTE_LINE_RE = re.compile(r"^[\s>]*>")

# 한국어 사람 이름 패턴 — ACTION_KEYWORD 매칭 false positive 방지용.
# 예: "김지연" 안의 "지연" 이 액션 키워드로 잘못 매칭 → 매칭 검사 전 line 에서
# 이런 이름 패턴을 임시 제거함 (출력 시엔 원본 line 유지).
#   1) 흔한 한국어 성씨 한 글자 + 이름 1~2글자 (김지연/이지연/박지연/곽지은 등 흔한 케이스)
#   2) 한글 2-4자 + 직장 호칭 ('이형조 차장', '김권영 프로' 등)
KOREAN_NAME_RE = re.compile(
    r"[김이박최정조강윤장임한오서신권황안송류홍전고문양손배백허남심노유진곽우주구함변도천표명피하][가-힣]{1,2}"
    r"|"
    r"[가-힣]{2,4}\s*(?:차장|부장|과장|대리|매니저|프로|책임|상무|이사|팀장|선임|수석|연구원)"
)


def parse_msg(path: Path) -> dict:
    """단일 .msg 파일을 dict 로 파싱."""
    with extract_msg.openMsg(str(path)) as m:
        # 날짜: msg 메타데이터 우선, 실패하면 파일명 prefix 에서 추출
        sent_at = None
        try:
            if m.date:
                sent_at = m.date if isinstance(m.date, datetime) else None
        except Exception:
            pass
        if sent_at is None:
            mt = FILENAME_DATE_RE.match(path.name)
            if mt:
                try:
                    sent_at = datetime.strptime(
                        f"{mt.group(1)}_{mt.group(2)}", "%y%m%d_%H%M"
                    )
                except ValueError:
                    pass

        return {
            "path": path,
            "sent_at": sent_at,
            "sender": fix_mojibake((m.sender or "").strip()),
            "to": fix_mojibake((m.to or "").strip()),
            "cc": fix_mojibake((m.cc or "").strip()),
            "subject": fix_mojibake((m.subject or "").strip()),
            "body": fix_mojibake(m.body or ""),
            "attachments": [
                fix_mojibake(getattr(att, "longFilename", "") or getattr(att, "shortFilename", ""))
                for att in (m.attachments or [])
            ],
        }


def _skip_external_warning_block(lines: list[str]) -> int:
    """본문 첫 부분에 외부 메일 경고가 있으면 그 블록 (경고 + 잇따르는 separator/빈줄) 통째로 skip.
    진짜 본문 시작 line index 반환. 경고 없으면 0."""
    n = len(lines)
    if n == 0:
        return 0
    head_has_warning = any(
        EXTERNAL_WARNING_RE.search(lines[i] or "")
        for i in range(min(EXTERNAL_WARNING_SCAN_LINES, n))
    )
    if not head_has_warning:
        return 0
    i = 0
    while i < n:
        s = (lines[i] or "").strip()
        if (
            not s
            or EXTERNAL_WARNING_RE.search(s)
            or re.match(r"^_{5,}$|^={5,}$|^-{20,}$", s)
        ):
            i += 1
            continue
        break
    return i


def clean_body_lines(body: str) -> list[str]:
    """본문을 줄 단위로 정제 — 인용·과한 공백·forwarded marker 이후 라인 제거.
    본문 시작의 외부 메일 경고 + separator 블록은 forward marker 검사 전 먼저 skip."""
    lines = body.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    start = _skip_external_warning_block(lines)
    cleaned: list[str] = []
    for raw in lines[start:]:
        if any(rx.match(raw.strip()) for rx in FORWARD_MARKERS):
            break  # forwarded thread 시작 — 여기서 cut
        if QUOTE_LINE_RE.match(raw):
            continue
        stripped = raw.rstrip()
        if not stripped:
            # 연속 빈 줄은 1개로 압축
            if cleaned and cleaned[-1] == "":
                continue
            cleaned.append("")
        else:
            if len(stripped) > MAX_LINE_LENGTH:
                stripped = stripped[:MAX_LINE_LENGTH] + " …"
            cleaned.append(stripped)
    # 앞뒤 빈 줄 제거
    while cleaned and cleaned[0] == "":
        cleaned.pop(0)
    while cleaned and cleaned[-1] == "":
        cleaned.pop()
    return cleaned


def find_action_lines(body_lines: list[str]) -> list[tuple[str, str]]:
    """본문 줄 중 ACTION_KEYWORDS 가 포함된 줄을 (keyword, line) tuple 로 반환.
    한 줄에 여러 키워드 매칭되면 ACTION_KEYWORDS 리스트 순서상 첫 번째만 기록.

    매칭 검사 시 한국어 사람 이름 패턴(KOREAN_NAME_RE)은 line 에서 임시 제거 후 검사.
    '김지연' 안의 '지연' 이 액션 키워드로 잘못 매칭되는 false positive 방지."""
    matched: list[tuple[str, str]] = []
    for line in body_lines:
        if not line:
            continue
        # 사람 이름 제거한 sanitized line 으로 키워드 검사 (출력은 원본 line)
        scan_line = KOREAN_NAME_RE.sub("", line).lower()
        for kw in ACTION_KEYWORDS:
            if kw.lower() in scan_line:
                matched.append((kw, line))
                break
    return matched


def short_addr(addr: str) -> str:
    """긴 이메일 주소 라인을 디스플레이용으로 축약."""
    if not addr:
        return "(없음)"
    # 'Name <name@domain>' 형식이면 Name 만, ; 로 여러 명이면 처음 3명 + ...
    parts = [p.strip() for p in re.split(r"[;,]", addr) if p.strip()]
    short = []
    for p in parts[:3]:
        m = re.match(r"^\s*([^<]+?)\s*<", p)
        short.append(m.group(1) if m else p)
    if len(parts) > 3:
        short.append(f"… (+{len(parts) - 3}명)")
    return "; ".join(short)


def build_report(msgs: list[dict]) -> str:
    """파싱된 메일들로 마크다운 리포트 본문 생성."""
    now = datetime.now()
    msgs_sorted = sorted(msgs, key=lambda x: x["sent_at"] or datetime.min)

    # 통계 + 전체 메일에서 action 줄 미리 수집 (통합 섹션 + per-email 양쪽에서 사용)
    dates = [m["sent_at"] for m in msgs_sorted if m["sent_at"]]
    sender_counter = Counter(short_addr(m["sender"]) for m in msgs_sorted)
    kw_counter: Counter[str] = Counter()
    # all_actions: 통합 섹션용 — {idx, sent_at, sender, subject, kw, line}
    all_actions: list[dict] = []
    # actions_by_msg: per-email 섹션용 — msg_idx → list of (kw, line)
    actions_by_msg: dict[int, list[tuple[str, str]]] = {}
    for i, m in enumerate(msgs_sorted, 1):
        body_lines = clean_body_lines(m["body"])
        actions = find_action_lines(body_lines)
        actions_by_msg[i] = actions
        for kw, line in actions:
            kw_counter[kw] += 1
            all_actions.append({
                "idx": i,
                "sent_at": m["sent_at"],
                "sender": short_addr(m["sender"]),
                "subject": m["subject"],
                "kw": kw,
                "line": line,
            })

    out: list[str] = []
    out.append(f"# 메일 요약 — `{SOURCE_DIR.name}`")
    out.append("")
    out.append(f"생성: {now:%Y-%m-%d %H:%M:%S}")
    out.append(f"원본 폴더: `{SOURCE_DIR}`")
    out.append("")

    # 개요
    out.append("## 개요")
    out.append("")
    out.append(f"- 메일 수: **{len(msgs_sorted)}개**")
    if dates:
        out.append(
            f"- 기간: **{min(dates):%Y-%m-%d %H:%M}** ~ **{max(dates):%Y-%m-%d %H:%M}**"
        )
    out.append("")

    out.append(f"### 발신자 TOP {TOP_SENDERS}")
    out.append("")
    out.append("| 발신자 | 건수 |")
    out.append("|---|---|")
    for sender, cnt in sender_counter.most_common(TOP_SENDERS):
        out.append(f"| {sender} | {cnt} |")
    out.append("")

    out.append("### 액션 키워드 빈도 (전체 본문 합산)")
    out.append("")
    out.append("| 키워드 | 발견 횟수 |")
    out.append("|---|---|")
    for kw, cnt in kw_counter.most_common():
        if cnt == 0:
            continue
        out.append(f"| `{kw}` | {cnt} |")
    out.append("")

    # 🔑 액션 아이템 통합 — 키워드별 그룹핑
    if all_actions:
        out.append("## 🔑 액션 아이템 통합 — 키워드별")
        out.append("")
        out.append(
            f"전체 매칭 줄: **{len(all_actions)}건** "
            f"(메일 번호 `#NN` 으로 타임라인 섹션 참조)"
        )
        out.append("")
        grouped: dict[str, list[dict]] = defaultdict(list)
        for a in all_actions:
            grouped[a["kw"]].append(a)
        # ACTION_KEYWORDS 정의 순서로 출력 (빈도순이 아니라 의미 그룹 순서 유지)
        for kw in ACTION_KEYWORDS:
            items = grouped.get(kw)
            if not items:
                continue
            out.append(f"### `{kw}` ({len(items)}건)")
            out.append("")
            for a in items:
                date_str = (
                    a["sent_at"].strftime("%m-%d %H:%M") if a["sent_at"] else "?"
                )
                out.append(
                    f"- **#{a['idx']:02d}** `{date_str}` "
                    f"{a['sender']} — {a['line']}"
                )
            out.append("")

    # 타임라인
    out.append("## 타임라인")
    out.append("")

    for i, m in enumerate(msgs_sorted, 1):
        body_lines = clean_body_lines(m["body"])
        action_lines = actions_by_msg.get(i, [])

        date_str = (
            m["sent_at"].strftime("%Y-%m-%d %H:%M") if m["sent_at"] else "(날짜 미상)"
        )
        out.append(f"### {i:02d}. {date_str} — {m['subject'] or '(제목 없음)'}")
        out.append("")
        out.append(f"- **From**: {short_addr(m['sender'])}")
        out.append(f"- **To**: {short_addr(m['to'])}")
        if m["cc"]:
            out.append(f"- **Cc**: {short_addr(m['cc'])}")
        if m["attachments"]:
            atts = ", ".join(a for a in m["attachments"] if a)
            if atts:
                out.append(f"- **첨부**: {atts}")
        out.append(f"- **파일**: `{m['path'].name}`")
        out.append("")

        if action_lines:
            out.append("**🔑 액션 아이템 후보**")
            out.append("")
            for kw, line in action_lines:
                out.append(f"  - `{kw}` {line}")
            out.append("")

        preview = body_lines[:BODY_PREVIEW_LINES]
        if preview:
            out.append("**📝 본문 미리보기**")
            out.append("")
            out.append("```text")
            out.extend(preview)
            if len(body_lines) > BODY_PREVIEW_LINES:
                out.append(
                    f"… (이하 {len(body_lines) - BODY_PREVIEW_LINES}줄 생략 — "
                    f"forwarded thread 또는 본문 길이 초과)"
                )
            out.append("```")
            out.append("")

        if INCLUDE_FULL_BODY:
            out.append("<details><summary>본문 전체</summary>")
            out.append("")
            out.append("```text")
            out.append(m["body"])
            out.append("```")
            out.append("")
            out.append("</details>")
            out.append("")

        out.append("---")
        out.append("")

    return "\n".join(out)


def main() -> int:
    # Windows cp949 콘솔에서 unicode (↻, →, ✅) 출력 시 UnicodeEncodeError 방지
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    if not SOURCE_DIR.exists():
        print(f"❌ SOURCE_DIR 가 없습니다: {SOURCE_DIR}")
        return 1

    msg_paths = sorted(SOURCE_DIR.glob("*.msg"))
    if not msg_paths:
        print(f"❌ {SOURCE_DIR} 안에 .msg 파일이 없습니다.")
        return 1

    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {len(msg_paths)}개 .msg 파일 파싱 중 ...")

    msgs: list[dict] = []
    failed: list[tuple[Path, Exception]] = []
    for p in msg_paths:
        try:
            msgs.append(parse_msg(p))
        except Exception as e:
            failed.append((p, e))
            print(f"  ⚠️ 파싱 실패: {p.name} — {e}")

    if not msgs:
        print("❌ 파싱된 메일이 없습니다.")
        return 1

    print(f"  → 성공: {len(msgs)}, 실패: {len(failed)}")

    report = build_report(msgs)

    # OUTPUT_DIR 미지정시 SOURCE_DIR (분석한 .msg 폴더) 안에 저장
    out_dir = OUTPUT_DIR if OUTPUT_DIR is not None else SOURCE_DIR
    out_dir.mkdir(parents=True, exist_ok=True)

    # 기존 _summary_*.md 중 가장 최근 1개를 keeper 로 살리고 (파일 ID 보존),
    # 새 timestamp 이름으로 rename → 내용만 덮어쓰는 방식.
    # 이렇게 하면 OneDrive/M365 file ID 가 유지돼서 공유 링크가 살아있음.
    timestamp = datetime.now().strftime("%y%m%d_%H%M")
    out_path = out_dir / f"{OUTPUT_BASENAME}_{timestamp}.md"

    existing = sorted(out_dir.glob(f"{OUTPUT_BASENAME}_*.md"))
    keeper = existing[-1] if existing else None  # 가장 최근 것 = 유지 대상

    # keeper 외 나머지(중복/누적 잔재) 삭제
    for old in existing[:-1]:
        try:
            old.unlink()
            print(f"  - 정리: {old.name}")
        except OSError as e:
            print(f"  ⚠️ 정리 실패: {old.name} — {e}")

    # keeper 가 새 이름과 다르면 rename (file ID 유지 → 링크 보존)
    rename_log = None
    if keeper is not None and keeper.name != out_path.name:
        try:
            keeper.rename(out_path)
            rename_log = f"  ↻ rename: {keeper.name} → {out_path.name} (M365 파일 ID 유지)"
        except OSError as e:
            rename_log = f"  ⚠️ rename 실패 → 새 파일로 작성 (링크 깨질 수 있음): {e}"
            try:
                keeper.unlink()
            except OSError:
                pass

    # 내용 덮어쓰기 (file write 가 print 보다 우선 — print 에서 unicode 에러 나도 파일은 갱신됨)
    out_path.write_text(report, encoding="utf-8")

    if rename_log:
        try:
            print(rename_log)
        except UnicodeEncodeError:
            pass

    print(f"\n✅ 요약 리포트 생성: {out_path}")
    print(f"   메일 {len(msgs)}건 → {out_path.stat().st_size:,} bytes")
    return 0


if __name__ == "__main__":
    sys.exit(main())
