# CLAUDE.md

이 파일은 Claude Code가 이 저장소에서 작업할 때 따르는 가이드다.

## 프로젝트

PPAP(Pen-Paper-Away-Project) — 법인 차량 운행일지 디지털화.
Google Apps Script + Google Sheets 백엔드(`Code.js`), 모바일 웹 폼(`ppap_form.html`) + 주차 보드(`parking_board.html`) 프론트.

데이터 흐름:
- READ는 차량 탭에서만 (주차현황은 차량 목록 = 차량_마스터, 위치 = RAW)
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
   - HEREDOC으로 작성, `Co-Authored-By: Claude` 라인 포함
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

## 차량 반납·추가 (리스 반납 / 렌트 신규)

`차량_마스터` F열 `반납일`과 G열 `위치기록여부`(체크박스)가 기준이다. 시트를 고친 뒤 GAS 에디터에서 `syncCarRegistry`를 실행한다(인자 없음, 여러 번 실행해도 결과 동일).

- **반납**: F열에 반납일 입력 → **반드시** `syncCarRegistry` 실행. `CAR_META_JSON`이 갱신되기 전에는 차단되지 않는다.
  (`setupProperties`도 `CAR_META_JSON`을 다시 만들어 차단은 반영하지만, 탭 숨김·생성은 `syncCarRegistry`만 한다.)
  - F열은 반납일 전용. F1에 `반납일`이 아닌 다른 값이 있으면 두 함수 모두 아무것도 바꾸기 전에 throw — 다른 용도의 F열 값을 반납일로 오인해 전 차량이 막히는 것을 방지.
  - 반납일 **당일까지** 폼 입력 허용, 다음 날부터 폼은 안내 페이지 · 저장은 throw · 주차 보드에서 제외.
  - 차량 탭은 숨김만 한다(삭제·이름 변경 금지 — 04:00 재정렬과 월간 백업에 계속 포함).
  - 날짜가 아닌 값(예: `반납`)은 즉시 반납으로 처리되고 실행 로그에 경고가 남는다.
- **추가**: 마스터에 행 추가(A 차량번호, B 차종, D 법인명, E 사업자번호, G 체크) → `syncCarRegistry` → 새 탭 머리말의 차량 정보 수정 →
  주차 보드에 `미기록` 카드 → 카드의 "운행일지 기록하기"(또는 QR)로 폼에서 "초기값 등록" → 보드에 위치 표시.
  - 새 탭은 기존 차량 탭(반납 안 된 탭 우선)을 복제하고 15행 이하를 비운 것이다. 머리말(1~14행)에는 원본 차량 정보가
    남아 있으므로 새 차량 값으로 고친다. 어느 탭을 복제했는지 실행 로그에 나온다.
  - sync 로그의 "생성" 목록에 예상하지 못한 차량이 있으면 탭 이름이 마스터 A열과 다른 것이다 — 중복 탭을 지우고 이름을 맞춘다. 전담(`fixedUser`)·출장(`businessTripCars`) 차량이면 `config.html` 수정·배포 후 `setupProperties`도 실행.
- **보드에서만 빼기**: G 체크 해제 → `syncCarRegistry`. 폼·저장은 그대로(주차위치 필수, RAW 기록 유지)라 다시 체크하면 최신 위치가 바로 보인다.
- 반납 취소: F열을 비우고 `syncCarRegistry` → 탭 숨김 해제.

### 주차 보드 표시 규칙

- 차량 목록은 `CAR_META_JSON`(마스터 행 순서), 위치는 RAW 최신 행. 표시 = G TRUE + 반납 아님.
  RAW 기록이 없으면 `미기록` 카드, 초기값등록만 있으면 위치만 표시(운전자 `—`).
- G 값은 체크박스 TRUE 또는 문자열 `TRUE`만 TRUE, 빈칸은 FALSE.
- G1 규칙: 비어 있으면 첫 `syncCarRegistry`가 `위치기록여부`를 쓰고 반납 아닌 차량을 체크한다(그 전까지는 전 차량 표시).
  다른 값이거나, G1은 비었는데 G2 아래에 값이 있으면(G1을 실수로 지운 경우 포함) `setupProperties`·`syncCarRegistry` 모두
  `CAR_META_JSON`을 쓰기 전에 throw — 체크 해제해 둔 값을 마이그레이션이 덮어쓰지 않게.
- 마스터 맨 아래에 추가한 행에는 체크박스가 없어 G가 FALSE로 읽힌다. sync가 새 탭을 만든 차량의 G가 FALSE면
  `경고: <차량번호>: G열 위치기록여부가 체크되지 않아 …`를 남긴다 — 체크박스를 복사해 넣고 다시 실행한다.
- sync 로그 `경고: <차량번호>: RAW 에 기록이 있지만 차량_마스터에 없어 …` → 그 차량은 보드에 나오지 않는다.
  운행 중인 차량이면 마스터에 행을 추가하고, 오타 차량번호면 무시한다.
- 반영 시점은 `syncCarRegistry`(또는 `setupProperties`) 실행 시 — 보드는 시트가 아니라 `CAR_META_JSON`을 읽는다.

## 작업 스타일

- 사용자는 1인 개발자, GAS·Sheets에 익숙. 기초 설명 생략, 트레이드오프·파일/라인 단위 코멘트 중심.
- 한국어 응답. 코드 식별자는 한글 변수명(`차량번호`, `주행전` 등) 자연스럽게 사용.
- 커밋 메시지·주석·PR 본문 모두 한국어.
- `SPREADSHEET_ID`는 `Code.js`의 `setupProperties` 안에 하드코딩되어 있어 환경 분리 시 이 부분도 확인.
