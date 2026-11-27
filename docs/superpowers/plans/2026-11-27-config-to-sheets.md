# config.html → 스프레드시트 이관 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 직원·거래처·전담·출장 설정을 config.html 대신 시트에서 읽어 `syncSettings` 한 번으로 Script Properties에 반영하고, config.html·`setupProperties`·CI 복원 단계를 없앤다.

**Architecture:** `Code.js`에 시트 읽기 함수(`_readStaff`, `_readClients`, `_readCarMaster` 확장)를 두고, `syncCarRegistry`를 `syncSettings`로 바꿔 모든 검증 뒤 5개 prop을 쓴다. 기존 props를 시트로 옮기는 1회 함수 `migrateConfigToSheets`를 추가한다. 폼(`serveForm`, `ppap_form.html`)은 prop 키·모양이 그대로라 건드리지 않는다.

**Tech Stack:** Google Apps Script V8, Node 24 `node:test` + `node:vm` 하네스(`test/harness.js`).

**Spec:** `docs/superpowers/specs/2026-11-27-config-to-sheets-design.md`

## Global Constraints

- 작업 위치: `C:\Users\GIC\repos\pen-paper-away-project-staging`, 브랜치 `staging`. push·배포는 하지 않는다(사용자의 "test 적용" 지시가 따로 있어야 한다).
- 건드리지 않는 파일: `ppap_form.html`, `parking_board.html`, `serveForm`(`Code.js:99-106`), 미커밋 `docs/superpowers/specs/2026-11-20-record-correction-sidebar-design.md`, 이전 설계·계획 문서.
- prop 키와 JSON 모양 불변: `STAFF_JSON` = `[{name, dept}]`, `FIXED_USER_JSON` = `{차량번호: {name, dept}}`, `BUSINESS_TRIP_CARS_JSON` = `[차량번호]`, `CLIENTS_JSON` = `[문자열]`, `CAR_META_JSON` 기존과 같음.
- 시트·헤더 이름(정확히): 시트 `직원` A1 `이름` B1 `부서` / 시트 `거래처` A1 `거래처` / `차량_마스터` H1 `전담자` I1 `전담자부서` J1 `출장차량`.
- 모든 throw는 시트·props·탭을 바꾸기 전에 일어난다. 설정 시트가 없을 때의 throw 메시지에는 `migrateConfigToSheets`를 넣는다.
- 셀 값은 `String(v == null ? "" : v).trim()`으로 읽는다. 체크박스 해석은 `isChecked`(true 또는 문자열 `TRUE`, 대소문자·공백 무시).
- 주석·로그·커밋 메시지는 한국어. 커밋은 conventional 접두어 + HEREDOC + 마지막 줄 `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.
- 테스트 명령: `npm test` (= `node --test "test/**/*.test.js"`). vm에서 나온 배열·객체는 교차 realm이라 `[...arr]`, `{...obj}`, 또는 `JSON.parse(JSON.stringify(x))`로 바꿔 비교한다. 실패 원인은 `node --test --test-reporter=tap <file>`로 본다.

## Review Focus

1. 전담자(H)는 있는데 전담자부서(I)가 빈칸 → `{ name, dept: "" }`로 담긴다(throw 아님). → Task 1 테스트
2. 전담자부서(I)만 있고 전담자(H)가 빈칸 → 전담 차량이 아니다. → Task 1 테스트
3. `거래처` 시트에 헤더만 있음 → `CLIENTS_JSON`은 `[]`이고 throw하지 않는다(출장 폼은 "직접 입력"만 남음). → Task 1 테스트
4. 이관할 환경에 `FIXED_USER_JSON`·`BUSINESS_TRIP_CARS_JSON`·`CLIENTS_JSON` 키가 없음 → 빈 값으로 보고 직원만 옮긴다. → Task 3 테스트
5. 차량_마스터에 헤더만 있는(차량 0대) 상태에서 이관 → H1:J1만 쓰고 행 수 0짜리 `getRange`·`insertCheckboxes`를 호출하지 않는다(GAS는 행 수 0이면 throw). → Task 3 테스트

결정(스펙에 없음): 월간 Excel 백업(`EXPORT_EXCLUDE_SHEETS`, `Code.js:715`)에서 `직원`·`거래처`를 빼지 않는다 — `공지사항`도 포함되는 기존 방식과 같고, 이름은 차량 탭에 이미 있다.

---

### Task 1: 설정 시트 읽기 (`_readStaff`, `_readClients`, `_readCarMaster` H~J)

**Files:**
- Modify: `Code.js` — `CONFIG`(16-22), `isLocationTracked`(919-923), `_readCarMaster`(945-982)
- Modify: `test/car-registry.test.js` — `MASTER_HEADER`, `masterRow`, `isLocationTracked` 테스트(302)
- Create: `test/settings-sheets.test.js`

**Interfaces:**
- Produces:
  - `CONFIG.SHEET_STAFF = "직원"`, `CONFIG.SHEET_CLIENTS = "거래처"`
  - `isChecked(value) -> boolean` — `isLocationTracked`의 새 이름. G열 해석도 이것을 쓴다(옛 이름은 지운다).
  - `_readStaff(ss) -> { staff: Array<{name: string, dept: string}>, warnings: string[] }` — 이름 빈 행 건너뜀, 시트 행 순서.
  - `_readClients(ss) -> string[]` — 빈 행 건너뜀, 시트 행 순서.
  - `_readCarMaster(masterSh) -> { carMeta, warnings, hasLocationCol, fixedUser: {[차량번호]: {name, dept}}, businessTripCars: string[] }` — 기존 필드에 두 필드 추가. 차량번호 빈 행은 건너뛴다. `businessTripCars`는 마스터 행 순서.
- throw 조건과 메시지에 꼭 들어갈 말:
  - `직원` 시트 없음 → `직원 시트가 없습니다` + `migrateConfigToSheets`
  - `직원` 1행이 `이름 | 부서`가 아님 → `직원 시트 1행`
  - `직원`에 이름 0개 → `직원 시트에 이름이 없습니다`
  - `거래처` 시트 없음 → `거래처 시트가 없습니다` + `migrateConfigToSheets`
  - `거래처` A1이 `거래처`가 아님 → `거래처 시트 1행`
  - 차량_마스터 H1:J1이 `전담자 | 전담자부서 | 출장차량`이 아님(빈칸 포함) → `H1:J1` + `migrateConfigToSheets`. F1·G1 검사 뒤에 한다.
- 경고: 같은 이름이 두 번 이상 → `직원 "<이름>" 이 여러 번 있습니다 — 폼은 첫 행의 부서를 씁니다.` (이름당 한 번)

- [ ] **Step 1: 기존 마스터 fixture를 10열로 바꾼다**

`test/car-registry.test.js`:
```js
const MASTER_HEADER = ["차량번호", "차종", "메모", "법인명", "사업자번호", "", "", "전담자", "전담자부서", "출장차량"];
function masterRow(carNo, 반납일 = "", 위치기록 = "", 전담자 = "", 전담자부서 = "", 출장 = "") {
  return [carNo, "K5", "", "(주)샘플", "123-45-67890", 반납일, 위치기록, 전담자, 전담자부서, 출장];
}
```
`isLocationTracked` 테스트(302)는 이름을 `isChecked: true 와 문자열 TRUE 만 TRUE`로 바꾸고 `gas.isChecked`를 부른다.

- [ ] **Step 2: 새 테스트 파일을 쓴다**

`test/settings-sheets.test.js` — fixture는 `loadGasWithSheets({ sheets })`에 `makeSheet`로 만든 시트를 넘긴다. 테스트:

```js
test("_readStaff: 이름 빈 행은 건너뛰고 trim, 시트 순서를 지킨다", () => {
  // 직원: [["이름","부서"],[" 홍길동 ","영업팀"],["",""],["김철수",""]]
  assert.deepEqual(plain(gas._readStaff(ss).staff), [
    { name: "홍길동", dept: "영업팀" },
    { name: "김철수", dept: "" },
  ]);
});
test("_readStaff: 직원 시트가 없으면 migrateConfigToSheets 를 안내하며 throw", ...);   // /migrateConfigToSheets/
test("_readStaff: 1행이 이름|부서 가 아니면 throw", ...);                               // /직원 시트 1행/
test("_readStaff: 이름이 하나도 없으면 throw", ...);                                    // 헤더만 → /직원 시트에 이름이 없습니다/
test("_readStaff: 같은 이름이 있으면 경고 한 번", ...);                                 // 홍길동 3번 → warnings.length === 1, includes("홍길동")
test("_readClients: 빈 행을 건너뛰고 순서를 지킨다", ...);                              // [["거래처"],["A상사"],[""],["B물산"]] → ["A상사","B물산"]
test("_readClients: 헤더만 있으면 빈 배열", ...);                                       // Review Focus 3
test("_readClients: 거래처 시트가 없으면 migrateConfigToSheets 를 안내하며 throw", ...);
test("_readCarMaster: H~J 에서 전담·출장을 읽는다", () => {
  // rows: A 전담자 "김전담"/"임원", B 출장 true, C 출장 "TRUE"
  // → fixedUser { A: {name:"김전담", dept:"임원"} }, businessTripCars [B, C]
});
test("_readCarMaster: 전담자만 있고 부서가 비면 dept 는 빈 문자열", ...);               // Review Focus 1
test("_readCarMaster: 전담자부서만 있으면 전담 차량이 아니다", ...);                     // Review Focus 2 → fixedUser {}
test("_readCarMaster: H1:J1 이 비어 있으면 migrateConfigToSheets 를 안내하며 throw", ...);
test("_readCarMaster: H1:J1 이 다른 값이면 throw", ...);                                // ["전담","",""] → /H1:J1/
```
`plain = (x) => JSON.parse(JSON.stringify(x))`.

- [ ] **Step 3: 실패 확인**

Run: `node --test test/settings-sheets.test.js test/car-registry.test.js`
Expected: settings-sheets는 `_readStaff is not a function` 등으로 FAIL, car-registry의 `isChecked` 테스트 FAIL. 나머지 car-registry 테스트는 PASS(H~J 검사가 아직 없음).

- [ ] **Step 4: 구현**

`Code.js`에 위 Interfaces대로 추가한다. `isLocationTracked`를 `isChecked`로 이름만 바꾸고 `_buildCarMeta`의 G 해석도 `isChecked`를 쓴다. `_readCarMaster`는 이미 `getDataRange()`로 읽은 `values`에서 H~J(인덱스 7~9)를 읽는다. 열 주석(`Code.js:926`)에 `| H 전담자 | I 전담자부서 | J 출장차량`을 더한다.

- [ ] **Step 5: 통과 확인**

Run: `npm test`
Expected: 전부 PASS. `syncCarRegistry` 테스트도 Step 1의 10열 헤더 덕분에 통과한다.

- [ ] **Step 6: 커밋**

```bash
git add Code.js test/car-registry.test.js test/settings-sheets.test.js
git commit  # feat: 직원·거래처·전담·출장 설정을 시트에서 읽는 함수 추가
```

---

### Task 2: `syncSettings` — 이름 변경, 설정 prop 기록, `setupProperties` 삭제

**Files:**
- Modify: `Code.js` — `getSpreadsheet`(59-63), `setupProperties`(878-916 삭제), `_readCarMaster` 주석(945, 965-966), `syncCarRegistry`(1037-1145)
- Modify: `test/car-registry.test.js`

**Interfaces:**
- Consumes: Task 1의 `_readStaff`, `_readClients`, `_readCarMaster`(fixedUser·businessTripCars 포함)
- Produces: `syncSettings() -> { created, hidden, shown, warnings }` — 반환 모양은 `syncCarRegistry`와 같다. 맨 앞에서 `_readCarMaster`·`_readStaff`·`_readClients`를 모두 부른 뒤에야 시트를 바꾼다. 끝에서 `CAR_META_JSON`과 함께 `STAFF_JSON`, `FIXED_USER_JSON`, `BUSINESS_TRIP_CARS_JSON`, `CLIENTS_JSON`을 `JSON.stringify`로 쓴다. `_readStaff` 경고는 `warnings`에 합친다. 실행 로그 첫 줄 접두어는 `syncSettings:`이고 `직원 N명 · 거래처 N곳 · 전담 N대 · 출장 N대`를 덧붙인다.
- `getSpreadsheet()`: `SPREADSHEET_ID`가 없거나 빈 문자열이면 `SPREADSHEET_ID 스크립트 속성이 없습니다 — 프로젝트 설정 → 스크립트 속성에 이 환경의 스프레드시트 ID 를 넣으세요.`로 throw.
- `setupProperties`와 운영 ID 기본값(`DEFAULT_PROD_ID`)은 지운다. 옛 이름 `syncCarRegistry` 별칭은 두지 않는다.

- [ ] **Step 1: fixture와 기존 테스트를 바꾼다**

`test/car-registry.test.js`:
- 모든 `gas.syncCarRegistry()`를 `gas.syncSettings()`로, 테스트 제목의 `syncCarRegistry`도 `syncSettings`로 바꾼다.
- `CONFIG_FILE`과 `files: { config: CONFIG_FILE }`을 지운다.
- `registryFixture`에 기본 설정 시트를 넣는다: `직원` `[["이름","부서"],["홍길동","영업팀"]]`, `거래처` `[["거래처"],["A상사"]]`. 옵션 `staffGrid`, `clientsGrid`(null이면 그 시트를 만들지 않음), `scriptProps`(기본 `{ CAR_META_JSON: "{}" }`에 합침), `masterHeader`(기본 `MASTER_HEADER`, F1·G1 덮어쓰기는 지금처럼 `fHeader`·`gHeader`)를 받는다.
- `setupProperties` 테스트 6개(268, 279, 290, 325, 349, 459)를 지운다. 각 검증은 같은 조건의 sync 테스트(256, 174, 362, 336, 444)가 이미 덮는다.

- [ ] **Step 2: 새 테스트를 쓴다**

```js
test("syncSettings 는 직원·거래처·전담·출장을 props 에 쓴다", () => {
  // rows: masterRow(OLD_CAR, "", "", "김전담", "임원"), masterRow(ACTIVE_CAR, "", "", "", "", true) + ACTIVE_CAR 탭
  // → JSON.parse(props.STAFF_JSON) = [{name:"홍길동",dept:"영업팀"}]
  //   CLIENTS_JSON = ["A상사"], FIXED_USER_JSON = {[OLD_CAR]: {name:"김전담",dept:"임원"}}, BUSINESS_TRIP_CARS_JSON = [ACTIVE_CAR]
});
test("직원 시트가 없으면 syncSettings 는 아무것도 바꾸기 전에 throw", () => {
  // staffGrid: null, scriptProps: { STAFF_JSON: '[{"name":"옛직원","dept":""}]' }, rows: [masterRow(NEW_CAR)]
  // → throws /migrateConfigToSheets/; props.STAFF_JSON 그대로; 마스터 F1(_grid[0][5]) === ""; NEW_CAR 탭 없음
});
test("H1:J1 이 없으면 syncSettings 는 아무것도 바꾸기 전에 throw", ...);  // masterHeader: MASTER_HEADER.slice(0, 7) → /H1:J1/, F1 그대로 ""
test("직원 이름 중복 경고가 sync 로그에 남는다", ...);                    // Logger._logs 에 "경고:" 와 이름
test("setupProperties 는 더 이상 없다", () => {
  assert.equal(typeof gas.setupProperties, "undefined");
  assert.equal(typeof gas.syncCarRegistry, "undefined");
});
test("SPREADSHEET_ID 가 없으면 getSpreadsheet 는 원인을 말하며 throw", () => {
  // loadGasWithSheets({ sheets: {}, scriptProps: { SPREADSHEET_ID: null } }) → throws /SPREADSHEET_ID/
});
```

- [ ] **Step 3: 실패 확인**

Run: `node --test test/car-registry.test.js`
Expected: `syncSettings is not a function`으로 대부분 FAIL.

- [ ] **Step 4: 구현**

`Code.js`를 Interfaces대로 바꾼다. 머리 주석(1037-1045)에 "직원·거래처·전담·출장 설정을 props 에 반영"을 더하고, `_readCarMaster`의 `setupProperties` 언급 주석(945, 966)을 고친다.

- [ ] **Step 5: 통과 확인**

Run: `npm test` 그리고 `grep -n "setupProperties\|syncCarRegistry\|DEFAULT_PROD_ID\|createHtmlOutputFromFile(\"config\")" Code.js test/*.js`
Expected: 테스트 전부 PASS. grep은 `test/car-registry.test.js`의 "더 이상 없다" 테스트 줄만 나온다.

- [ ] **Step 6: 커밋**

```bash
git add Code.js test/car-registry.test.js
git commit  # feat: syncCarRegistry 를 syncSettings 로 바꾸고 설정 시트를 props 에 반영, setupProperties 삭제
```

---

### Task 3: `migrateConfigToSheets` (환경별 1회 이관)

**Files:**
- Modify: `test/harness.js` — `makeSheet`, `makeSpreadsheet`
- Modify: `Code.js` — `syncSettings` 위에 추가
- Modify: `test/settings-sheets.test.js`

**Interfaces:**
- Consumes: `CONFIG.SHEET_STAFF`, `CONFIG.SHEET_CLIENTS`, `isChecked` 해석 규칙과 같은 J 값(true/false), Task 2의 `syncSettings`(round-trip 테스트)
- Produces: `migrateConfigToSheets() -> { staff: number, clients: number, fixedUser: number, businessTripCars: number, skipped: string[] }` — 옮긴 개수와 마스터에 없어 건너뛴 차량번호(중복 없음).
  - 입력: props `STAFF_JSON`(필수), `FIXED_USER_JSON`(없으면 `{}`), `BUSINESS_TRIP_CARS_JSON`(없으면 `[]`), `CLIENTS_JSON`(없으면 `[]`). props는 쓰지 않는다.
  - 쓰기 전 throw: `STAFF_JSON` 없음·빈 배열 → `STAFF_JSON`; `직원` 또는 `거래처` 시트가 이미 있음 → `이미 있습니다`; 차량_마스터 H~J(1행부터 마지막 행까지)에 빈칸 아닌 값 → 그 셀 주소(예: `H3`).
  - 쓰기: 마스터 열이 10개보다 적으면 `insertColumnsAfter(getMaxColumns(), 10 - getMaxColumns())`. `ss.insertSheet("직원")`에 헤더 + `[name, dept]` 행, `ss.insertSheet("거래처")`에 헤더 + 거래처 행. 마스터 H1:J1 헤더, 2행부터 각 행의 차량번호(trim)로 H·I(전담자·부서)와 J(체크박스, 출장이면 true)를 쓴다. 마스터 데이터 행이 0이면 H1:J1만 쓴다. J는 `insertCheckboxes()` 뒤에 `setValues`(체크박스 삽입이 값을 false로 만든다).
  - 마스터에 없는 전담·출장 차량은 `경고: <차량번호>: 전담/출장 설정이 있지만 차량_마스터에 없어 옮기지 않았습니다.`로 로그하고 `skipped`에 넣는다. 끝에 `이관 완료 — 시트를 확인한 뒤 syncSettings 를 실행하세요.`를 로그한다.
- 하네스:
  - `makeSheet(name, grid, { maxColumns } = {})` — `getMaxColumns()`(기본값 `Math.max(26, 가장 넓은 행 길이)`), `insertColumnsAfter(after, n)`(최대 열 수를 n만큼 늘리고 sheet 반환). `getRange`는 지금처럼 최대 열을 강제하지 않는다.
  - `makeSpreadsheet`에 `insertSheet(name)` — 같은 이름이 있으면 throw, 없으면 빈 시트를 만들어 붙이고 반환.

- [ ] **Step 1: 실패하는 테스트를 쓴다**

`test/settings-sheets.test.js`에 이관 fixture(마스터 10열 아님 — 7열 헤더 `MASTER7`, 차량 A·B, 각 차량 탭, `scriptProps`로 4개 JSON)를 두고:

```js
test("이관: props 의 직원·거래처를 새 시트로 옮기고 props 는 건드리지 않는다", ...);
  // 직원 _grid = [["이름","부서"],["홍길동","영업팀"],["김철수","관리팀"]]; 거래처 = [["거래처"],["A상사"],["B물산"]]; props.STAFF_JSON 문자열 그대로
test("이관: 전담·출장을 차량_마스터 H~J 로 옮긴다", ...);
  // H1:J1 헤더; A 행 H,I = "김전담","임원", J false; B 행 H,I = "","", J true; _checkboxRanges 에 {row:2,col:10,numRows:2,numCols:1}
test("이관: 마스터에 없는 전담·출장 차량은 건너뛰고 경고", ...);           // FIXED_USER_JSON 에 "99하9999" → skipped ["99하9999"], 로그 "경고:"
test("이관: 직원 시트가 이미 있으면 아무것도 바꾸기 전에 throw", ...);      // /이미 있습니다/, 마스터 H1 === "" 그대로, 거래처 시트 없음
test("이관: 차량_마스터 H~J 에 값이 있으면 throw", ...);                   // B 행 H = "메모" → /H3/, 직원 시트 없음
test("이관: STAFF_JSON 이 없으면 throw", ...);                             // /STAFF_JSON/
test("이관: 마스터 열이 10개보다 적으면 늘린다", ...);                     // maxColumns: 7 → 이관 뒤 getMaxColumns() >= 10
test("이관: 전담·출장·거래처 prop 이 없는 환경은 직원만 옮긴다", ...);     // Review Focus 4 → 거래처 _grid = [["거래처"]], H2:I 빈칸, J 전부 false
test("이관: 차량이 0대면 H1:J1 만 쓰고 체크박스를 넣지 않는다", ...);       // Review Focus 5 → _checkboxRanges.length === 0
test("이관 뒤 syncSettings 는 이관 전과 같은 props 를 만든다", () => {
  // 원래 4개 prop 을 기억 → migrateConfigToSheets() → syncSettings()
  // → 4개 prop 을 JSON.parse 한 값이 원래 값과 deepEqual (BUSINESS_TRIP_CARS_JSON 은 마스터 행 순서와 같은 순서로 fixture 를 만든다)
});
```

- [ ] **Step 2: 실패 확인**

Run: `node --test test/settings-sheets.test.js`
Expected: 새 테스트가 `migrateConfigToSheets is not a function`으로 FAIL, Task 1 테스트는 PASS.

- [ ] **Step 3: 하네스 확장과 구현**

`test/harness.js`와 `Code.js`를 Interfaces대로 바꾼다. 검사를 모두 끝낸 뒤에만 `insertColumnsAfter`·`insertSheet`·`setValues`를 부른다.

- [ ] **Step 4: 통과 확인**

Run: `npm test`
Expected: 전부 PASS.

- [ ] **Step 5: 커밋**

```bash
git add Code.js test/harness.js test/settings-sheets.test.js
git commit  # feat: 기존 props 를 설정 시트로 옮기는 migrateConfigToSheets 추가
```

---

### Task 4: 배포 단계와 문서 정리

**Files:**
- Modify: `.github/workflows/deploy.yml:96-105` — "Restore config.html from secret" 단계 삭제
- Modify: `.claspignore` — `config.html` 줄 추가(`.gitignore`의 `config.html`은 그대로)
- Modify: `README.md` — Setup(66-70), 테스트 환경 5단계(109), "config.html 처리" 절(148-159), 트러블슈팅(165-169)
- Modify: `CLAUDE.md`

**Interfaces:**
- Consumes: Task 2·3의 함수 이름 `syncSettings`, `migrateConfigToSheets`, 시트·헤더 이름(Global Constraints)

- [ ] **Step 1: 워크플로·claspignore**

복원 단계를 통째로 지운다(앞뒤 단계 들여쓰기 유지). `.claspignore` 끝에 `config.html`.

- [ ] **Step 2: README**

- Setup: `config.json`·`setupProperties` 설명을 "직원·거래처는 `직원`·`거래처` 시트, 전담·출장은 `차량_마스터` H~J에서 고치고 GAS 에디터에서 `syncSettings` 실행"으로 바꾼다.
- 테스트 환경 5단계: `setupProperties()` → `syncSettings()`. 4단계의 "(운영 ID 폴백 방지)"는 "(없으면 모든 실행이 SPREADSHEET_ID 오류로 멈춤)"으로.
- "config.html 처리" 절 → "설정 시트" 절: 시트 위치, 반영은 `syncSettings`, 환경별 1회 이관(`migrateConfigToSheets` → 시트 확인 → `syncSettings`), secret `CONFIG_HTML`은 prod 적용 후 삭제.
- 트러블슈팅: `CONFIG_HTML secret 미설정` 줄 삭제, "테스트 GAS가 운영 Sheets를 건드림" 줄은 `SPREADSHEET_ID` 확인 후 `syncSettings` 재실행으로, "GAS에서 파일이 갑자기 사라짐" 줄에서 config.html 언급 삭제.

- [ ] **Step 3: CLAUDE.md**

- "수동 1회 작업": `setupProperties` 줄을 `migrateConfigToSheets` — 환경별 1회, 기존 props를 설정 시트로 이관(이후 `syncSettings`)으로 바꾼다.
- `syncCarRegistry`를 전부 `syncSettings`로. 73행의 `setupProperties` 괄호 문장 삭제. 82행 "전담·출장 차량이면 config.html 수정·배포 후 setupProperties도 실행" → "전담·출장 차량이면 H~J를 채우고 `syncSettings`". 92·98행의 `setupProperties` 언급 삭제. 105행 → "`SPREADSHEET_ID`는 각 GAS 프로젝트의 Script Properties에 직접 넣는다(코드 기본값 없음)."
- 새 절 "### 설정 시트 (직원·거래처·전담·출장)": 시트·열 위치, 반영 시점 = `syncSettings`, throw 조건(시트·헤더 없음, 직원 0명), 이관 순서(배포 → `migrateConfigToSheets` → 시트 확인 → `syncSettings`, 이관 전 sync는 throw하고 props 불변), secret `CONFIG_HTML`은 prod 적용 후 삭제.

- [ ] **Step 4: 확인**

Run: `grep -rn "setupProperties\|syncCarRegistry\|CONFIG_HTML\|config\.html" Code.js README.md CLAUDE.md .github .claspignore` 그리고 `npm test`
Expected: grep 결과는 `.claspignore`의 `config.html`, README·CLAUDE.md의 이관·secret 삭제 안내 줄뿐이다. 테스트 전부 PASS.

- [ ] **Step 5: 커밋**

```bash
git add .github/workflows/deploy.yml .claspignore README.md CLAUDE.md docs/superpowers/specs/2026-11-27-config-to-sheets-design.md docs/superpowers/plans/2026-11-27-config-to-sheets.md
git commit  # chore: config.html 복원 단계 제거, 설정 시트 운영 문서화
```
