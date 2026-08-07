# txt → docx 자동 변환 시스템

메모장(.txt)으로 작성한 속기록·회의록을 자동으로 정리된 Word 문서(.docx)로 변환하는 개인 업무 자동화 도구.
여러 실무 프로젝트의 인수인계 교육·회의록·주간보고를 정리하는 데 매일 사용 중.

## 하는 일

`doc_input/` 폴더에 txt 파일을 드롭하면:

1. 감시 프로그램(`doc_watcher.py`)이 자동 감지
2. Claude API가 내용 분석
   - 날짜 / 담당자 / 제목 자동 추출
   - 문서 타입 분류 (교육·온보딩 / 회의록)
   - 실제 WBS 폴더 트리를 실시간 스캔해서 저장 경로 자동 판단 (경로 하드코딩 없음)
   - 섹션 구성 및 내용 정제
3. Word 문서(.docx) 생성 및 배포
   - 판단된 WBS 경로에 저장 (메인)
   - `3_finished/`에 복사 (결과 확인용)
   - `4_archive/`에 복사 (시간순 보관, 타임스탬프 파일명)
4. 원본 txt는 `1_processed/`로 이동
5. `2_log/`에 변환 결과 로그 생성 (실패 시 재처리 가이드 자동 포함)
6. 완료 시 토스트 알림

재처리가 필요하면 원본 txt 상단에 `#경로:`, `#제목:`, `#타입:`, `#날짜:`, `#담당자:` 힌트 줄을 추가해서
다시 드롭하면 힌트를 우선 적용해 재변환한다.

## 구조

```
doc_input/                 (입력 폴더 — 여기 txt를 드롭)
├── 1_processed/            원본 txt 보관
├── 2_log/                  변환 결과 로그
├── 3_finished/             생성된 docx 복사본 (확인용)
└── 4_archive/               전체 결과 시간순 보관
```

## 저장 경로 판단 (WBS 라우팅)

저장 경로는 사람이 지정하지 않는다. Claude가 매 처리 시 실제 폴더 트리를 스캔해서:

- 적합한 기존 폴더가 있으면 → 거기 저장
- 없으면 WBS 번호 규칙에 맞게 새 폴더명을 직접 생성 후 저장

`wbs_routing.json`이 스캔 대상 최상위 폴더(`scan_roots`)와 스캔 깊이(`scan_depth`)만 관리하고,
실제 경로 테이블은 존재하지 않는다 — 폴더 구조가 바뀌어도 자동 반영된다.

## 파일

| 파일 | 역할 |
|---|---|
| `doc_watcher.py` | 핵심 스크립트. 폴더 감시 + Claude API 호출 + docx 생성 + 로그 기록 |
| `wbs_routing.json` | WBS 스캔 설정 (스캔 대상 루트 폴더, 깊이) |
| `run_watcher.bat` | 수동 실행용 (평소엔 Windows 작업 스케줄러로 자동 실행) |
| `requirements.txt` | 의존 패키지 목록 |
| `skills/doc-writer/SKILL.md` | 채팅창에 속기록을 직접 붙여넣을 때 쓰는 수동 버전 (Claude Code Skill) |

## 실행 조건

- Python 3.x
- 패키지: `pip install -r requirements.txt`
- `.env` 파일에 `ANTHROPIC_API_KEY` 설정 (`.env.example` 참고)
- 평소엔 Windows 작업 스케줄러 자동 실행, 꺼져있으면 `run_watcher.bat` 수동 실행

## `doc-writer` 스킬

`doc_watcher.py`가 자동 감시 파이프라인이라면, `skills/doc-writer/SKILL.md`는 같은 규칙을
채팅창(Claude Code)에서 속기록을 직접 붙여넣을 때 쓰는 수동 버전이다. 두 경로 모두 같은 규칙을
따르도록 관리한다 (파일명·시트명 원문 보존, 출처 추적, WBS 경로 자동 판단 등).
