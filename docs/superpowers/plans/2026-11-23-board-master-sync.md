# 주차 보드 ↔ 차량_마스터 연동 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 주차 보드 목록을 RAW 기준에서 `차량_마스터` 기준으로 바꾸고, 마스터 G열 `위치기록여부`가 TRUE이면서 반납 아닌 차량만 띄운다.

**Architecture:** 마스터 G열을 `_readCarMaster`가 읽어 `CAR_META_JSON`의 `위치기록: boolean`으로 담는다. `_buildParkingBoard`는 `CAR_META_JSON` 키 순서(=마스터 행 순서)로 차량을 고르고 RAW 최신 기록을 붙인다.
`syncCarRegistry`는 G열 첫 도입(마이그레이션)과 "RAW에만 있는 차량번호" 경고를 맡는다.

**Tech Stack:** Google Apps Script(V8), Node 24 `node:test` + `node:vm` 하네스(`test/harness.js`)

**Spec:** `docs/superpowers/specs/2026-11-23-board-master-sync-design.md`

## Global Constraints

- 작업 위치: `C:\Users\GIC\repos\pen-paper-away-project-staging` (`staging` 브랜치). push는 사용자 지시 후에만.
- `Code.js`에 `module.exports` 금지(GAS에서 깨짐). 테스트는 `test/harness.js`의 `loadGas` / `loadGasWithSheets`로 로드.
- G1 헤더 문자열은 정확히 `위치기록여부`.
- 값 해석: `true` 또는 문자열 `"TRUE"`(대소문자·앞뒤 공백 무시)만 TRUE, 나머지(빈 칸 포함)는 FALSE.
- G1이 비어 있으면 모든 차량 `위치기록: true`. G1이 다른 값이면 `setupProperties`·`syncCarRegistry` 모두 아무것도 바꾸기 전에 throw.
- `CAR_META_JSON`에 `위치기록` 키가 없으면 TRUE로 본다(`!== false`).
- 보드 표시 조건: `위치기록 !== false && !isRetiredCar(반납일, today)`.
- 폼(`ppap_form.html`)과 저장 경로는 바꾸지 않는다((a)안).
- 커밋 메시지는 한국어 + conventional 접두어 + `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.
- 테스트 명령: `npm test` (전체 GREEN 유지).

## Review Focus

1. G 셀에 체크박스 대신 손으로 `true`/` TRUE `를 친 경우 → TRUE로 읽혀야 한다 (Task 1 `isLocationTracked` 테스트).
2. 배포 직후 sync 전, 예전 `CAR_META_JSON`(위치기록 키 없음) → 보드가 비면 안 된다 (Task 2 테스트).
3. RAW 차량번호 뒤에 공백이 붙은 기록 → 마스터 차량 카드에 붙어야 한다 (Task 2 테스트).
4. 마스터 중간에 차량번호가 빈 행이 있을 때의 첫 sync 마이그레이션 → 깨지지 않고 그 행은 FALSE (Task 3 테스트).
5. 미래 반납일 차량의 첫 sync 마이그레이션 → 아직 운행 중이므로 체크(TRUE) (Task 3 테스트).

---

### Task 1: 마스터 G열 `위치기록여부` 해석

**Files:**
- Modify: `Code.js` — `_buildCarMeta`(917), `_readCarMaster`(935), 열 설명 주석(915). 새 함수 `isLocationTracked`는 `_buildCarMeta` 바로 위.
- Test: `test/car-registry.test.js`

**Interfaces:**
- Produces:
  - `isLocationTracked(value: any) -> boolean`
  - `_buildCarMeta(rows: any[][], hasLocationCol = false) -> { [carNo]: { 차종, 법인명, 사업자번호, 반납일, 위치기록: boolean } }` — `hasLocationCol`이 false면 모든 차량 `위치기록: true`, true면 `isLocationTracked(r[6])`.
  - `_readCarMaster(masterSh) -> { carMeta, warnings: string[], hasLocationCol: boolean }` — `hasLocationCol`은 G1(trim)이 `"위치기록여부"`인지. G1이 비어 있지 않고 `"위치기록여부"`도 아니면 메시지에 `G1`을 담아 throw(F1 검사와 같은 위치, 어떤 쓰기보다 먼저).

- [ ] **Step 1: 테스트 fixture 확장**

`MASTER_HEADER`를 7열로(G1 = `""`), `masterRow(carNo, 반납일 = "", 위치기록 = "")`가 7번째 값으로 `위치기록`을 넣게, `registryFixture`에 `gHeader = ""` 인자를 추가해 `header[6] = gHeader`.

- [ ] **Step 2: 실패하는 테스트 작성**

```js
test("isLocationTracked: true 와 문자열 TRUE 만 TRUE", () => {
  const gas = loadGas();
  for (const v of [true, "TRUE", " true "]) assert.equal(gas.isLocationTracked(v), true, String(v));
  for (const v of [false, "", "FALSE", null, undefined, "Y", 1]) assert.equal(gas.isLocationTracked(v), false, String(v));
});

test("G1 이 위치기록여부면 행별 G 값을 CAR_META_JSON 위치기록으로 담는다", () => {
  const { gas, props } = registryFixture({
    rows: [masterRow(OLD_CAR, "", true), masterRow(ACTIVE_CAR, "", false)],
    gHeader: "위치기록여부",
    extraTabs: { [ACTIVE_CAR]: carTabGrid("운행기록부", 2000) },
  });
  gas.syncCarRegistry();
  assert.equal(meta(props)[OLD_CAR].위치기록, true);
  assert.equal(meta(props)[ACTIVE_CAR].위치기록, false);
});

test("G1 이 비어 있으면 setupProperties 는 모든 차량을 위치기록 TRUE 로 담는다", () => {
  const { gas, props } = registryFixture({ rows: [masterRow(OLD_CAR, "", false)] });
  gas.setupProperties();
  assert.equal(meta(props)[OLD_CAR].위치기록, true);
});

test("G1 이 다른 헤더면 syncCarRegistry 는 아무것도 바꾸기 전에 throw", () => {
  const { gas, ss, props } = registryFixture({ rows: [masterRow(OLD_CAR, "2026-11-30")], gHeader: "비고" });
  assert.throws(() => gas.syncCarRegistry(), /G1/);
  assert.equal(props.CAR_META_JSON, "{}");
  assert.equal(ss.getSheetByName(OLD_CAR).isSheetHidden(), false);
  assert.equal(ss.getSheetByName(MASTER).getRange(1, 6).getValue(), "");
});

test("G1 이 다른 헤더면 setupProperties 도 throw", () => {
  const { gas, props } = registryFixture({ rows: [masterRow(OLD_CAR)], gHeader: "비고" });
  assert.throws(() => gas.setupProperties(), /G1/);
  assert.equal(props.CAR_META_JSON, "{}");
});
```

기존 `_buildCarMeta` 테스트(인자 1개 호출)의 기대 객체에 `위치기록: true`를 추가하고, 같은 테스트에 `hasLocationCol = true` 호출 한 건을 붙인다: 7번째 값 `true` → `위치기록: true`, `""` → `false`.

- [ ] **Step 3: RED 확인**

Run: `npm test`
Expected: 위 신규 테스트와 수정한 `_buildCarMeta` 테스트가 FAIL (`isLocationTracked is not a function`, `위치기록` undefined, G1 throw 없음). 기존 테스트는 PASS.

- [ ] **Step 4: 구현** — Interfaces 그대로. `_buildCarMeta`의 열 주석에 `G 위치기록여부` 추가.

- [ ] **Step 5: GREEN 확인**

Run: `npm test`
Expected: 전부 PASS.

- [ ] **Step 6: 커밋**

```bash
git add Code.js test/car-registry.test.js
git commit  # feat: 차량_마스터 G열 위치기록여부 → CAR_META_JSON 위치기록
```

---

### Task 2: 주차 보드 목록을 마스터 기준으로

**Files:**
- Modify: `Code.js` — `_buildParkingBoard`(186-248)와 그 위 주석(182 "초기값등록 행은 … 제외" 문구 갱신)
- Create: `test/parking-board.test.js`

**Interfaces:**
- Consumes: `CAR_META_JSON`의 `위치기록`(Task 1). 없으면 TRUE.
- Produces: `_buildParkingBoard() -> Array<{ carNo, carName, parking, name, time }>` — 형태는 그대로, 순서는 `Object.keys(carMeta)` 순서. RAW 기록이 없으면 `{ parking: "", name: "", time: "—" }`. 초기값등록 행도 최신 기록 후보에 넣는다(그 행은 성명·부서가 비어 `name: ""`).

- [ ] **Step 1: 실패하는 테스트 작성** (`test/parking-board.test.js`)

fixture: `boardFixture({ meta, raw })` — `meta`는 `CAR_META_JSON` 객체, `raw`는 `{ carNo, flag = "정상", ts, parking = "B1", name = "홍길동", dept = "영업팀" }` 배열. 각 RAW 행은 17열, `COL` 인덱스로 채운다(`test/car-retire.test.js`의 `boardFixture` 참고). `ts`는 `"yyyy-MM-dd HH:mm:ss"` KST 문자열.

```js
test("RAW 기록이 없는 마스터 차량도 미기록 카드로 나온다", () => {
  const g = boardFixture({ meta: { [A]: { 차종: "K5", 반납일: "", 위치기록: true } }, raw: [] });
  assert.deepEqual({ ...g._buildParkingBoard()[0] }, { carNo: A, carName: "K5", parking: "", name: "", time: "—" });
});

test("초기값등록 행만 있으면 위치는 보이고 운전자는 비어 있다", () => {
  const g = boardFixture({
    meta: { [A]: { 차종: "K5", 위치기록: true } },
    raw: [{ carNo: A, flag: "초기값등록", parking: "B2", name: "", dept: "" }],
  });
  const card = g._buildParkingBoard()[0];
  assert.equal(card.parking, "B2");
  assert.equal(card.name, "");
});

test("초기값등록 뒤 운행 기록이 있으면 최신 운행 기록을 쓴다", () => { /* 초기값 B2(이른 ts) → 정상 B3(늦은 ts) → parking B3, name "영업팀 홍길동" */ });

test("위치기록 FALSE 차량은 보드에서 빠진다", () => { /* meta A 위치기록 false, RAW 기록 있음 → board 길이 0 */ });

test("마스터에 없는 차량번호는 RAW 기록이 있어도 빠진다", () => { /* meta {A}, RAW에 B 기록 → carNo 목록 [A] */ });

test("카드 순서는 마스터 순서를 따른다", () => { /* meta 키 순서 [A, B], RAW는 B 먼저 기록 → [A, B] */ });

test("위치기록 키가 없는 예전 CAR_META_JSON 은 표시한다", () => { /* meta { [A]: { 차종: "K5" } } → 길이 1 */ });

test("RAW 차량번호 뒤 공백은 마스터 차량에 붙는다", () => { /* RAW carNo `${A} `, parking B1 → A 카드 parking B1 */ });
```

본문을 주석으로 적은 테스트는 주석의 입력·기대값 그대로 작성한다. 차량번호는 `A = "12가3456"`, `B = "34나5678"`.

- [ ] **Step 2: RED 확인**

Run: `npm test`
Expected: 미기록 카드·초기값등록·마스터 미등록·순서 테스트가 FAIL. `test/car-retire.test.js`의 보드 테스트 2건은 PASS 유지.

- [ ] **Step 3: 구현** — RAW를 한 번 돌아 차량번호(trim)별 최신 기록 맵을 만들고(초기값등록 포함), `Object.keys(carMeta)`를 돌며 표시 조건을 거른 뒤 카드를 만든다. 시간 라벨 계산은 지금 코드를 그대로 쓴다.

- [ ] **Step 4: GREEN 확인**

Run: `npm test`
Expected: 전부 PASS.

- [ ] **Step 5: 커밋**

```bash
git add Code.js test/parking-board.test.js
git commit  # feat: 주차 보드 목록을 차량_마스터 기준으로 (위치기록 TRUE + 반납 아님)
```

---

### Task 3: sync — G열 첫 도입, RAW 미등록 차량 경고, 운영 문서

**Files:**
- Modify: `Code.js` — `syncCarRegistry`(검증 뒤, F1 기록 옆에서 마이그레이션 호출 / 경고 합치기), 새 함수 `_initLocationColumn`, `_unregisteredRawCars`(둘 다 `syncCarRegistry` 바로 위)
- Modify: `test/harness.js` — range에 `insertCheckboxes()` 스텁
- Modify: `CLAUDE.md` — "차량 반납·추가" 절
- Test: `test/car-registry.test.js`

**Interfaces:**
- Consumes: `_readCarMaster(...).hasLocationCol`(Task 1), `isRetiredCar`.
- Produces:
  - `_initLocationColumn(masterSh, carMeta, todayStr) -> void` — G1에 `위치기록여부`를 쓰고, G2부터 마스터 마지막 행까지 `insertCheckboxes()` 후 값을 쓴다. 차량 행 = `!isRetiredCar(반납일, todayStr)`, 차량번호 빈 행 = `false`. 쓴 값으로 `carMeta[carNo].위치기록`을 맞춘다. `hasLocationCol`이 false일 때만, 모든 검증(F1·G1·복제 원본)이 끝난 뒤 호출한다.
  - `_unregisteredRawCars(ss, carMeta) -> string[]` — RAW `차량번호` 열(trim)에서 `carMeta`에 없는 값을 처음 나온 순서로 중복 없이. RAW 시트가 없거나 데이터 행이 없으면 `[]`.
  - `syncCarRegistry().warnings`에 `${carNo}: RAW 에 기록이 있지만 차량_마스터에 없어 주차 보드에 나오지 않습니다.`를 더한다(기존 `경고:` 로그로 함께 출력).
- 하네스: `getRange(...).insertCheckboxes()`는 시트의 `_checkboxRanges`에 `{ row, col, numRows, numCols }`를 push하고 range를 반환.

- [ ] **Step 1: 하네스 스텁 + fixture 확장**

`makeSheet`에 `_checkboxRanges: []`와 range의 `insertCheckboxes`를 추가한다. `registryFixture`에 `rawCars = []` 인자를 추가해, 값이 있으면 `RAW_운행일지` 시트(헤더 1행 + 차량번호만 채운 17열 행)를 만든다.

- [ ] **Step 2: 실패하는 테스트 작성**

```js
test("G1 이 비면 첫 sync 가 헤더와 체크박스를 쓴다 — 반납 아닌 차량만 체크", () => {
  const { gas, ss, props } = registryFixture({
    rows: [masterRow(OLD_CAR), masterRow(ACTIVE_CAR, "2020-01-01")],
    extraTabs: { [ACTIVE_CAR]: carTabGrid("운행기록부", 2000) },
  });
  gas.syncCarRegistry();
  const master = ss.getSheetByName(MASTER);
  assert.equal(master.getRange(1, 7).getValue(), "위치기록여부");
  assert.equal(master.getRange(2, 7).getValue(), true);
  assert.equal(master.getRange(3, 7).getValue(), false);
  assert.ok(master._checkboxRanges.some((r) => r.row === 2 && r.col === 7 && r.numRows === 2));
  assert.equal(meta(props)[OLD_CAR].위치기록, true);
  assert.equal(meta(props)[ACTIVE_CAR].위치기록, false);
});

test("첫 sync 마이그레이션: 미래 반납일 차량은 아직 운행 중이라 체크", () => {
  /* rows [masterRow(OLD_CAR, "2999-12-31")] → G2 true */
});

test("첫 sync 마이그레이션: 차량번호가 빈 행이 있어도 깨지지 않고 그 행은 FALSE", () => {
  /* rows [masterRow(OLD_CAR), ["", "", "", "", "", "", ""], masterRow(ACTIVE_CAR)] (ACTIVE_CAR 탭 있음) → G3 false, G4 true, meta 두 차량 true */
});

test("G1 이 이미 있으면 sync 는 G 값을 덮어쓰지 않는다", () => {
  /* gHeader "위치기록여부", rows [masterRow(OLD_CAR, "", false)] → sync 두 번 뒤 G2 false, _checkboxRanges 비어 있음 */
});

test("RAW 에만 있는 차량번호는 sync 경고로 남는다", () => {
  /* rows [masterRow(OLD_CAR)], rawCars [OLD_CAR, "99하9999", "99하9999 "] → warnings 중 "99하9999" 포함 항목이 정확히 1개, Logger 에 "경고:" + "99하9999" */
});
```

기존 "복제할 기존 차량 탭이 없으면 … throw" 테스트에 `assert.equal(ss.getSheetByName(MASTER).getRange(1, 7).getValue(), "")`를 추가한다.

- [ ] **Step 3: RED 확인**

Run: `npm test`
Expected: 신규 테스트 FAIL (G1 미기록, 경고 없음). 다른 테스트 PASS.

- [ ] **Step 4: 구현** — Interfaces 그대로. 실행 로그 요약 줄은 바꾸지 않는다.

- [ ] **Step 5: GREEN 확인**

Run: `npm test`
Expected: 전부 PASS.

- [ ] **Step 6: CLAUDE.md 갱신** — "차량 반납·추가" 절을 spec의 "운영 절차"대로 고친다:
  - 추가 절차에 "G열 `위치기록여부` 체크"를 넣는다.
  - "보드에서만 빼기: G 체크 해제 → `syncCarRegistry`" 항목을 추가한다.
  - 보드 표시 규칙 한 줄: G TRUE + 반납 아님, 마스터 순서, 기록 없으면 `미기록`.
  - G1 규칙: 비어 있으면 첫 sync가 채움, 다른 값이면 중단.
  - RAW 미등록 경고의 의미.

- [ ] **Step 7: 커밋** (spec·plan 문서 포함)

```bash
git add Code.js test/harness.js test/car-registry.test.js CLAUDE.md docs/superpowers/specs/2026-11-23-board-master-sync-design.md docs/superpowers/plans/2026-11-23-board-master-sync.md
git commit  # feat: syncCarRegistry 위치기록여부 첫 도입 + RAW 미등록 차량 경고
```
