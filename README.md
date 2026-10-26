# automation-modules

업무 자동화를 위해 직접 만든 모듈과, 그 과정에서 만든 Claude 스킬·규칙·기록을 모아둔 레포.
각 하위 폴더는 독립적이며, 모듈성 폴더는 자체 README/문서에 상세 내용이 있다.

## 자동화 모듈

| 모듈 | 설명 | 최초 생성일 | 최근 업데이트 |
|---|---|---|---|
| [`txt2docx/`](./txt2docx) | 메모장(.txt) 속기록·회의록을 감시 폴더에 넣으면 Claude API가 분석해 정리된 Word(.docx)로 자동 변환. 저장 경로도 실제 폴더 구조를 스캔해 자동 판단. | 2026-05-21 | 2026-10-20 |
| [`mail2xlsx/`](./mail2xlsx) | 로컬 Outlook(COM) 받은편지함에서 특정 발신자 알림 메일을 스캔·파싱해 관리대장 xlsx로 자동 정리. | 2026-07-29 | 2026-10-20 |

## Claude 스킬·규칙

| 폴더 | 설명 | 최근 업데이트 |
|---|---|---|
| [`skills_claude/`](./skills_claude) | Claude Code 스킬·규칙 백업본. folder-audit·gitgit·update-md·folder_structure 규칙 md와, 스크립트 동반 스킬 `weekly-plan/`(전주 to-do + 금주 주간보고 → 이번주 개인업무 계획 txt+pptx 생성). | 2026-10-26 |

## 기록 / 기타

| 폴더 | 설명 |
|---|---|
| [`troubleshooting_log/`](./troubleshooting_log) | PC 트러블슈팅 기록 (예: 탐색기 행(hang) — OneDrive 쉘확장 이슈 해결기록). |
| [`3.1_Agents/`](./3.1_Agents) | 에이전트 관련 (현재 비어 있음, placeholder). |

> 날짜는 모듈 최초 생성일 / git 커밋 기준 최근 업데이트. 새 작업물은 성격에 맞는 폴더에 쌓는다.
