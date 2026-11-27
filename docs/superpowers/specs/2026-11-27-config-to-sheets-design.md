# config.html → 스프레드시트 이관 설계 (짧은 설계)

## 의도

- config.html(직원·전담차량·출장차량·거래처)을 바꿀 때마다 secret `CONFIG_HTML` 갱신 → push → 환경별 `setupProperties`를
  거쳐야 한다. GAS 에디터에서 고친 config.html은 다음 배포에 덮인다. 편집자는 운영자 1인.
- 시트를 고치고 `syncSettings` 한 번으로 반영되게 한다. 차량 교체 때 이미 하는 절차와 같다.

## 목표 / 비목표 / 안티골

- **목표**: 4종 데이터를 시트에서 관리. config.html·secret `CONFIG_HTML`·CI 복원 단계 제거. 폼 동작 불변.
- **비목표**: 자동 반영(onEdit·캐시), 폼 UI 변경, 관리자 화면, test↔prod 시트 동기화(각 환경이 자기 시트를 가진다).
- **안티골**: 이관 도중 props가 비거나 잘못된 값으로 덮여 운영 폼의 성명 목록·전담·출장 동작이 바뀌는 것.

## 접근

| 안 | 내용 | 결론 |
|---|---|---|
| **A** | 시트 → sync 실행 → Script Properties (`CAR_META_JSON`과 같은 방식) | **채택** (사용자 선택) |
| B | 폼 요청마다 시트를 읽고 CacheService | 캐시 miss 때 로딩 지연, 반영이 TTL·onEdit(바인딩 미확인)에 걸림, 차량 마스터와 반영 규칙이 갈림 |
| C | config.html 유지, secret 갱신만 자동화 | 시트로 관리하자는 취지와 반대 |

## 설계

### 시트 배치

- **`직원`**: A1 `이름`, B1 `부서`. 2행부터. 이름이 빈 행은 건너뛴다. 행 순서 = 폼 드롭다운 순서.
- **`거래처`**: A1 `거래처`. 2행부터, 빈 행은 건너뛴다. 행 순서 = 폼 선택지 순서.
- **`차량_마스터`** H `전담자`, I `전담자부서`, J `출장차량`(체크박스).
  - H가 비어 있지 않은 행 → `FIXED_USER_JSON[차량번호] = { name: H, dept: I }` (trim).
  - J 값 해석은 G와 같다(체크박스 true 또는 문자열 `TRUE`) → `BUSINESS_TRIP_CARS_JSON`(마스터 행 순서).
  - 반납 차량 행도 그대로 포함한다. 폼이 반납 차량을 먼저 막으므로 무해하다.
- 전담자 부서를 `직원`에서 찾지 않고 I열에 따로 둔다(내 판단 — 앞선 대화에서 말한 "직원에서 찾기"를 바꿈).
  지금 `FIXED_USER_JSON`은 `{name, dept}`를 따로 들고 있고, 전담 차량은 드롭다운을 쓰지 않으므로 전담자가 직원 목록에
  없을 수 있다. 직원에서 찾게 하면 이관 때 그런 전담자를 `직원`에 넣어야 하고, 그러면 다른 차량의 드롭다운에 나타난다.

### 반영 (`syncSettings` 하나로 통합, `setupProperties` 삭제)

- 공용 읽기 함수가 `직원`·`거래처`·마스터 H~J를 읽어 `{ staff, fixedUser, businessTripCars, clients }`를 만든다.
  마스터 값은 `_readCarMaster`가 이미 `getDataRange()`로 읽는 값을 쓴다.
- **아무것도 바꾸기 전에 throw** (메시지에 `migrateConfigToSheets` 먼저 실행 안내 포함):
  - `직원` 또는 `거래처` 시트가 없거나 1행 헤더가 위와 다름
  - 차량_마스터 H1·I1·J1이 정확히 `전담자`·`전담자부서`·`출장차량`이 아님
  - `직원`에 이름이 하나도 없음 (빈 드롭다운으로 운영 폼이 막히는 것 방지)
- 경고: `직원` 이름 중복 — 폼은 이름으로 첫 행 부서를 찾는다(`ppap_form.html:1466`, `1616`).
- `syncSettings`: 기존 탭 처리 + 5개 prop(`STAFF_JSON`, `FIXED_USER_JSON`, `BUSINESS_TRIP_CARS_JSON`, `CLIENTS_JSON`,
  `CAR_META_JSON`) 기록.
- **`syncCarRegistry`를 `syncSettings`로 이름을 바꾼다**(사용자 결정). 이제 차량만 다루지 않는다. `syncData`는 RAW→차량 탭
  동기화(`syncAllCarSheets`, `dailyResyncCarSheets`)와 헷갈려 쓰지 않았다. 시간 트리거로 호출하는 곳이 없어 이름 변경이 안전하다.
  옛 이름 별칭은 두지 않는다. 이전 설계·계획 문서(`2026-11-2x-*`)는 기록이라 고치지 않는다.
- **`setupProperties`는 삭제한다**(사용자 질문 "나눌 필요가 있음?"에서 결정). config.html을 읽지 않게 되면 남는 차이는 셋인데
  어느 것도 따로 둘 이유가 아니다:
  - `SPREADSHEET_ID`가 없으면 운영 ID로 폴백(`Code.js:887-889`) — 테스트 GAS가 운영 시트를 건드리는 원인(README 트러블슈팅).
    없애면 `getSpreadsheet()`가 실패해 바로 드러난다. 새 환경은 Script Properties에 `SPREADSHEET_ID`를 손으로 넣는다.
  - 탭 처리를 하지 않음 — 탭 처리는 마스터를 따라가는 멱등 동작이라 prop만 갱신하고 탭은 두고 싶은 경우가 없다.
  - G1이 빈칸이면 전 차량을 `위치기록 TRUE`로 담음 — sync는 같은 상황에서 G열을 초기화한다. 같은 시트에 두 함수가
    다른 결과를 내 혼란을 만든다.
- prop 키와 모양이 그대로이므로 `serveForm`(`Code.js:99-106`)과 `ppap_form.html`은 바꾸지 않는다.
  `serveForm`은 요청마다 props를 읽으므로(`getAllScriptProps`, `Code.js:66`) sync 직후 반영된다.

### 이관 `migrateConfigToSheets` (환경별 1회, GAS 에디터에서 수동)

- **입력은 현재 Script Properties**(`STAFF_JSON`·`FIXED_USER_JSON`·`BUSINESS_TRIP_CARS_JSON`·`CLIENTS_JSON`)다.
  config.html이 아니라서 배포로 GAS의 config.html이 지워진 뒤에도 실행할 수 있고, 각 환경이 자기 값을 옮긴다.
- 쓰기 전 검사 → throw: `STAFF_JSON`이 없거나 빈 배열 / `직원`·`거래처` 시트가 이미 있음 / 차량_마스터 H~J에
  값이 있음(헤더 포함).
- 쓰기: `직원`·`거래처` 시트 생성, H1:J1 헤더, 마스터 각 행의 전담자·전담자부서, J열 체크박스 + 출장 여부.
- `FIXED_USER_JSON`·`BUSINESS_TRIP_CARS_JSON`의 차량 중 마스터에 없는 차량은 건너뛰고 경고 로그를 남긴다
  (필요하면 마스터에 행을 추가하고 H~J를 손으로 채운다).
- props는 건드리지 않는다. 이어서 실행하는 sync가 시트에서 다시 만든다.

### 배포·정리

- `deploy.yml` "Restore config.html from secret" 단계(96-105) 삭제.
- `.claspignore`에 `config.html` 추가 — 로컬에 남은 config.html이 수동 `clasp push`로 올라가지 않게.
- README "config.html 처리" 절, 트러블슈팅의 `CONFIG_HTML`·`setupProperties` 항목, 초기 설정 절(`setupProperties` 실행
  안내)을 갱신. CLAUDE.md의 `setupProperties` 언급(수동 1회 작업, 반납·추가 절, 보드 규칙, 작업 스타일의 `SPREADSHEET_ID`
  하드코딩 문구)을 `syncSettings` 기준으로 고치고, `syncCarRegistry` 언급도 모두 `syncSettings`로 바꾼다.
  시트 관리 규칙과 이관 순서를 추가.
- `setupProperties`를 대상으로 한 테스트(`test/car-registry.test.js`)는 같은 검증을 `syncSettings`로 옮기거나,
  sync 테스트와 겹치면 지운다.
- secret `CONFIG_HTML` 삭제는 **prod 적용 후**. main의 `deploy.yml`에는 머지 전까지 복원 단계가 남아 있어,
  먼저 지우면 main 배포가 그 단계에서 중단된다.

### 운영 순서 (환경별)

1. 배포 (test: staging push / prod: main 머지). 이때 `clasp push -f`가 GAS의 config.html을 지운다. props는 그대로라 폼은 정상.
2. GAS 에디터에서 `migrateConfigToSheets` 실행 → 시트 확인.
3. `syncSettings` 실행 → props를 시트에서 다시 만든다. 폼에서 성명 목록·전담 차량·출장 차량 확인.
4. 1과 2 사이에 `syncSettings`를 실행하면 throw하고 props는 바뀌지 않는다.
5. prod까지 끝나면 secret `CONFIG_HTML` 삭제, 로컬 config.html 삭제(선택).

## 제약

- Script property 값당 9KB.
- 반영 시점은 `syncSettings` 실행 시.
- test·prod 시트는 따로 관리한다.

## 빌더에게 남긴 것

- 테스트 하네스에 `insertSheet` 스텁이 없다.
- 차량_마스터 열 수가 10보다 적을 때 열을 늘리는 방법(`insertColumnsAfter` 등).
- G열 해석 함수(`isLocationTracked`)를 J열에도 쓸지, 이름을 일반화할지.
- 이관 → sync 후 props가 이관 전 props와 같은지(round-trip) 테스트.
