# mail_search

Outlook 공유 메일함에서 키워드 매칭되는 메일을 `.msg` 파일 + 첨부파일로 일괄 다운로드하는 단일 파일 스크립트.

## 파일 구성

| 파일 | 설명 |
|---|---|
| `mail_search_to_msg.py` | 메인 스크립트 (win32com 기반, Outlook 데스크톱 앱 필요) |
| `mail_search_to_msg.md` | 사용 가이드 (설정 변수, 동작 흐름, 출력 파일명 규칙, 트러블슈팅) |

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
     └─ <YYMMDD_HHMM>_<원본 첨부 파일명>
   ```

## 요구사항

- Windows + Outlook 데스크톱 앱 (실행 중 + 본인 프로필에 대상 메일함 등록)
- Python 3.x
- `pywin32` 패키지 (`pip install pywin32`)

상세 가이드: [`mail_search_to_msg.md`](mail_search_to_msg.md)
