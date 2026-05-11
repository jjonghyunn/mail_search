# mail_search

Outlook 공유 메일함에서 키워드 매칭되는 메일을 `.msg` 로 다운로드하고, 받은 `.msg` 들을 한 개 마크다운 리포트로 요약하는 도구 모음.

## 파일 구성

| 파일 | 설명 |
|---|---|
| `mail_search_to_msg.py` | 메일 다운로드 — Outlook 공유 메일함에서 키워드 매칭 메일을 `.msg` + 첨부로 일괄 저장 (`win32com` 기반, Outlook 앱 필요) |
| `mail_search_to_msg.md` | 위 스크립트 사용 가이드 |
| `summarize_msgs.py` | 받은 `.msg` 폴더를 시간순 마크다운 리포트로 요약 — 발신자 TOP / 액션 키워드 통합 / 메일별 본문 미리보기 + 액션 아이템 후보. `extract-msg` 기반, Outlook 설치 불필요 |
| `summarize_msgs.md` | 위 스크립트 사용 가이드 |

## 빠른 시작

1. 스크립트 상단 `── 설정 ──` 섹션에서:
   - `KEYWORDS` — 검색할 단어 리스트 (제목/본문 OR 매칭, 대소문자 무관)
   - `STORE_NAME` — Outlook DisplayName 부분 일치 (공유 메일함 이름)
   - `FOLDER_NAME` — `None` 이면 받은편함, 다른 폴더 이름 입력 가능
   - `RECURSE_SUBFOLDERS` — 하위 폴더까지 재귀 검색 여부
   - `SEARCH_BODY` — 본문 검색 ON/OFF (대량 메일함이면 OFF 권장)
   - `SAVE_ATTACHMENTS` — 매칭 메일의 첨부파일도 저장
   - `SKIP_INLINE_IMAGES` — 서명·인라인 이미지(`image001.png` 등) 자동 skip

2. 실행:
   ```bash
   python mail_search_to_msg.py
   ```

3. 결과:
   ```
   ~/Downloads/mail_search_<YYMMDD>/
     ├─ <YYMMDD_HHMM>_<safe subject>.msg
     ├─ <YYMMDD_HHMM>_<원본 첨부 파일명>
     └─ _processed_entry_ids.txt    ← 재실행 시 중복 skip 마커
   ```

4. **같은 날짜에 키워드 바꿔서 재실행** → 같은 메일은 자동 skip (EntryID 마커 기반).
   강제 재저장은 `_processed_entry_ids.txt` 삭제 후 실행.

## 요구사항

`mail_search_to_msg.py` (메일 다운로드):
- Windows + Outlook 데스크톱 앱 (실행 중 + 본인 프로필에 대상 메일함 등록)
- Python 3.x
- `pywin32` 패키지 (`pip install pywin32`)

`summarize_msgs.py` (요약 리포트):
- Python 3.x
- `extract-msg` 패키지 (`pip install extract-msg`) — Outlook 설치 불필요

상세 가이드:
- 메일 다운로드: [`mail_search_to_msg.md`](mail_search_to_msg.md)
- 요약 리포트: [`summarize_msgs.md`](summarize_msgs.md)
