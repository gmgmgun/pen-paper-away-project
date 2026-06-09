# AGENTS.md

이 파일은 Codex가 이 저장소에서 작업할 때 따르는 가이드다.

## 프로젝트

PPAP(Pen-Paper-Away-Project) — 법인 차량 운행일지 디지털화.
Google Apps Script + Google Sheets 백엔드(`Code.js`), 모바일 웹 폼(`ppap_form.html`) + 주차 보드(`parking_board.html`) 프론트.

데이터 흐름:
- READ는 차량 탭에서만 (주차현황은 RAW에서)
- WRITE는 RAW + 차량 탭 양쪽 — 한쪽만 건드리지 말 것
- RAW는 SSOT, 절대 직접 수정 금지 (이력 누적 원칙)
- `CONFIG.DATA_START_ROW = 15`, `CAR_TOTAL_COLS = 44` — 양식 변경 시 동시 갱신
- 차량 탭의 `=T{n-1}` 수식이 직전 행을 참조하므로 행 삽입/이동 시 수식 무결성 확인 필수

## 배포 워크플로

`.github/workflows/deploy.yml`의 GitHub Actions가 자동 배포를 처리한다. 로컬에서 `clasp push`를 직접 칠 필요 없음.

| 브랜치 | 환경 | scriptId | deployment |
|---|---|---|---|
| `staging` | **test** | `secrets.GAS_SCRIPT_ID_TEST` | `secrets.GAS_DEPLOYMENT_ID_TEST` |
| `main` | **prod** | `.clasp.json` 기본값 | `secrets.GAS_DEPLOYMENT_ID` |

배포 트리거 파일: `Code.js`, `*.html`, `appsscript.json`, `.clasp.json`, `.claspignore`, `.github/workflows/deploy.yml`. 다른 파일만 변경된 푸시는 워크플로가 안 돈다.

## 적용 지시 처리 규칙

사용자가 적용을 지시하면 변경만 만들고 끝내지 말고, **사용자가 해당 환경 웹앱 화면에서 즉시 검증 가능한 상태까지** 처리한다.

### "test 적용", "테스트 반영", "스테이징 푸시" 등 → test 환경 끝까지 처리

1. `git status`로 working tree 확인
2. 현재 브랜치가 `staging`이 아니면 staging으로 전환 (없으면 `git checkout -b staging origin/staging`)
3. 변경 파일을 명시적으로 stage 후 커밋
   - 메시지는 한국어 + conventional commit 접두어 (feat/fix/chore/refactor)
   - HEREDOC으로 작성, `Co-Authored-By: Codex` 라인 포함
4. `git push origin staging`
5. `gh run list --branch staging --limit 1`로 GitHub Actions 워크플로 완료까지 대기 (필요시 `gh run watch`)
6. 워크플로 성공 확인 후 사용자에게 "test 웹앱에서 확인하세요"를 안내. URL은 운영자가 보유.

### "운영 적용", "prod 반영", "main 머지" 등 → prod 환경 끝까지 처리

1. `origin/staging`이 `origin/main` 대비 앞서 있는지 확인. **test 검증이 안 된 변경의 prod 적용은 금지** — 그런 상황이면 사용자에게 "test 검증 먼저" 안내하고 중단.
2. 기존 패턴(`Merge pull request from gmgmgun/staging`)에 따라 PR 생성: `gh pr create --base main --head staging`
3. PR 머지: `gh pr merge --merge` (방식이 망설여지면 사용자에게 옵션 확인)
4. main 푸시 후 GitHub Actions가 prod 배포. `gh run watch`로 완료 대기.
5. 사용자에게 "운영 웹앱에서 확인하세요" 안내.

### 공통 주의

- prod 적용은 사용자가 **명시적으로 지시했을 때만**. 자동으로 staging→main 머지하지 않는다.
- 워크플로가 deploy 단계까지 성공했는지 확인 후 보고. push 단계만 성공하고 deploy 실패한 경우 그 사실을 명시.

## 수동 1회 작업 (배포로 해결되지 않는 것)

GAS 에디터에서 직접 실행해야 하는 setup 함수들. 코드 배포만으로는 트리거가 안 걸린다:

- `setupProperties` — `config.json` 변경 후 ScriptProperties 갱신
- `setupExportTrigger` — 매월 1일 자정 Excel 백업 트리거
- `setupWarmupTrigger` — 5분 워밍업 트리거
- `setupDailyResyncTrigger` — 매일 04:00 KST 차량 탭 재정렬 (과거 날짜 입력 보정)
- `setupNoticeSheet` — `공지사항` 시트 생성/헤더·체크박스·검증 세팅 (배너 기능 최초 1회)

새 setup 함수를 추가하거나 트리거 변경 시 사용자에게 GAS 에디터에서 어떤 함수를 어느 환경(test/prod)에서 실행해야 하는지 명확히 안내.

## 작업 스타일

- 사용자는 1인 개발자, GAS·Sheets에 익숙. 기초 설명 생략, 트레이드오프·파일/라인 단위 코멘트 중심.
- 한국어 응답. 코드 식별자는 한글 변수명(`차량번호`, `주행전` 등) 자연스럽게 사용.
- 커밋 메시지·주석·PR 본문 모두 한국어.
- `SPREADSHEET_ID`는 `Code.js`의 `setupProperties` 안에 하드코딩되어 있어 환경 분리 시 이 부분도 확인.
