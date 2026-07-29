# mail_search  
<sub>2026-07-29  Jonghyun Park w/ Claude</sub>  

Outlook 공유 메일함에서 키워드 매칭되는 메일을 `.msg` 로 다운로드하고, 받은 `.msg` 들을 한 개 마크다운 리포트로 요약하는 도구 모음.

## 파일 구성

| 파일 | 설명 |
|---|---|
| `mail_search_to_msg.py` | 메일 다운로드 — Outlook 공유 메일함에서 키워드 매칭 메일을 `.msg` + 첨부로 일괄 저장 (`win32com` 기반, Outlook 앱 필요) |
| `mail_search_to_msg_onlymsg.py` | `.msg` 파일만 저장하는 경량 버전 — 첨부파일 제외, 메일 본문/메타데이터만 필요할 때 사용. ⚠️ 설정 스키마가 본편과 다름 — ① 메일함 지정이 `STORE_NAME`(문자열 1개)이라 여러 메일함 동시 검색 불가, ② `SAVE_DIR` 이 날짜 폴더 없는 **고정 경로**라 아래 "결과" 트리가 적용되지 않음 (실행 전 직접 교체할 것) |
| `mail_search_to_msg.md` | 위 스크립트 사용 가이드 |
| `summarize_msgs.py` | 받은 `.msg` 폴더를 시간순 마크다운 리포트로 요약 — 발신자 TOP / 액션 키워드 통합 / 메일별 본문 미리보기 + 액션 아이템 후보. `extract-msg` 기반, Outlook 설치 불필요 |
| `summarize_msgs.md` | 위 스크립트 사용 가이드 |

## 빠른 시작

1. 스크립트 상단 `── 설정 ──` 섹션에서:
   - `KEYWORDS` — 제목/본문에서 검색할 단어 리스트 (그룹 내부 OR 매칭, 대소문자 무관)
   - `SENDER_KEYWORDS` — 발신자(이름/이메일)에서 검색할 단어 리스트. "그 사람이 보낸 메일" 을 찾을 때 사용 (그룹 내부 OR)
   - `MATCH_LOGIC` — 위 두 그룹의 결합 방식. `"OR"`(기본) = 제목/본문 **또는** 발신자 매칭, `"AND"` = 둘 다 매칭. 한쪽 그룹이 비어 있으면 나머지 한쪽으로만 검색
   - `RECEIVED_FROM` — 받은 날짜 하한. `None`(기본) 이면 제한 없음, `date(2026,1,1)` 처럼 주면 그 이후 메일만
   - `WHOLE_WORD` — 단어 경계(`\b`) 매칭 여부. **기본 True 로 두는 것을 권장** — False 면 `ai`·`kv` 같은 2~3글자 키워드가 `email`/`available` 안에 substring 으로 잡혀 과매칭됨
   - `STORE_NAMES` — Outlook DisplayName 부분 일치 **리스트** (여러 메일함 동시 검색 가능)
   - `INCLUDE_ARCHIVE` — 온라인 보관(아카이브) store 도 검색 (기본 True — 옛 메일 누락 방지)
   - `SKIP_PUBLIC_FOLDERS` — 공용 폴더(Public Folders) store 제외 (기본 True — 개인 mailbox 와 이름이 겹쳐 생기는 오매칭 방지)
   - `SEARCH_WHOLE_STORE` — True 면 store 전 폴더 검색, 아래 두 항목은 무시됨 (기본 False)
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

4. **같은 날짜에 키워드 바꿔서 재실행** → 같은 메일은 자동 skip (Message-ID 우선 / EntryID fallback 마커 기반).
   여러 메일함·아카이브를 동시에 봐도 같은 메일은 한 번만 저장됨.
   강제 재저장은 `_processed_entry_ids.txt` 삭제 후 실행.

## 요구사항

전체 패키지 일괄 설치:
```bash
pip install -r requirements.txt
```

`mail_search_to_msg.py` (메일 다운로드):
- Windows + Outlook 데스크톱 앱 (실행 중 + 본인 프로필에 대상 메일함 등록)
- Python 3.x
- `pywin32` 패키지

`summarize_msgs.py` (요약 리포트):
- Python 3.x
- `extract-msg` 패키지 — Outlook 설치 불필요

상세 가이드:
- 메일 다운로드: [`mail_search_to_msg.md`](mail_search_to_msg.md)
- 요약 리포트: [`summarize_msgs.md`](summarize_msgs.md)

## License

MIT
