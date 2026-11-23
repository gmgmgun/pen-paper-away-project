# 차량 반납·추가 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 리스 반납한 차량은 폼·저장·주차 보드에서 막고 탭을 숨긴다. 렌트 신규 차량은 `차량_마스터`에 행을 추가하고 `syncCarRegistry` 한 번을 실행하면 탭이 생기고 등록된다.

**Architecture:** `차량_마스터` F열 `반납일`을 기준으로 삼는다. 이 값을 `CAR_META_JSON`에 담고, 순수 함수 `isRetiredCar`로 판정해 `serveForm`·`saveRecord`·`_updateRecordInner`·`_buildParkingBoard`에 적용한다. 탭 숨김·복원·생성과 메타 갱신은 인자 없는 idempotent 함수 `syncCarRegistry`가 맡는다.

**Tech Stack:** Google Apps Script(V8), Google Sheets, Node 24 `node:test` + `node:vm` 하네스

**Spec:** `docs/superpowers/specs/2026-11-23-car-retire-add-design.md`

## Global Constraints

- 작업 위치: `C:\Users\GIC\repos\pen-paper-away-project-staging`(브랜치 `staging`). 2026-11-20의 미커밋 변경(`normalizeReturnDate`, `isRetiredCar`, 보드 제외, `setupProperties` 반납일)을 이어받는다.
- push 금지. 각 Task는 로컬 커밋까지만 한다. 원격 반영은 사용자의 "test 적용" 지시가 있을 때 `CLAUDE.md` 절차로 한다.
- `docs/superpowers/specs/2026-11-20-record-correction-sidebar-design.md`(다른 기능)는 stage하지 않는다. 파일은 항상 명시적으로 `git add`한다.
- 반납 판정: 반납일이 `yyyy-MM-dd`이면 `반납일 < todayStr`(KST)일 때 반납이다. 당일은 허용한다. 비어 있지 않고 날짜 형식이 아닌 값은 즉시 반납이다.
- RAW는 읽기만 한다(주차 보드). 이 기능은 RAW에 쓰지 않는다. 차단은 모든 쓰기보다 먼저 throw한다.
- 차량 탭 시트명 = 차량번호(trim). 반납 탭은 이름 변경·삭제 금지, `hideSheet()`만 한다.
- `CONFIG.DATA_START_ROW = 15`, `CAR_TOTAL_COLS = 44`. `차량_마스터` 열: A 차량번호 | B 차종 | C 메모 | D 법인명 | E 사업자번호 | F 반납일
- 차단 문구(정확히): `반납된 차량입니다 (반납일 ${반납일}).`
- `Code.js`에 `module.exports`를 달지 않는다(GAS 런타임에서 깨진다). 테스트는 `test/harness.js`의 `loadGas`/`loadGasWithSheets`로 로드한다.
- 커밋 메시지는 한국어 + conventional 접두어 + HEREDOC으로 쓰고, 마지막 줄은 `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`이다.

## Review Focus

1. **반납일 당일**: 폼·저장·주차 보드 모두 평소처럼 동작해야 한다. → Task 1 보드 테스트, Task 2 저장 테스트
2. **차량번호 앞뒤 공백**(`" 78라1234 "`): 탭명·메타 키는 trim한 값이어야 하며, 공백이 붙은 중복 탭이 생기면 안 된다. → Task 3 테스트
3. **날짜가 아닌 반납일 텍스트**(`"반납"`, `"2026.11.30"`): 즉시 반납으로 처리하되, 실행 로그에 경고가 남아야 한다. → Task 3 테스트
4. **15행 이하에 잔재가 남은 `차량_양식`**: 새 탭의 `getPrevOdoData().prevOdo`는 `null`이어야 한다(초기값 등록 모드). → Task 3 테스트
5. **F열만 채우고 `syncCarRegistry`를 실행하지 않은 경우**: `CAR_META_JSON`이 갱신되지 않아 차단도 안 된다. 테스트로 막을 수 없으므로 `CLAUDE.md` 운영 절차에 "입력 후 반드시 실행"을 명시한다. → Task 3 Step 7

---

### Task 1: 테스트 하네스 이식 + 반납 판정 규칙 확정

**Files:**
- Create: `test/harness.js`(`C:\Users\GIC\repos\pen-paper-away-project\test\harness.js`를 그대로 복사), `package.json`(같은 위치의 파일 복사)
- Modify: `.claspignore`(끝에 `test/`, `test/**`, `package.json`, `package-lock.json` 추가), `Code.js:139-145`(`isRetiredCar`)
- Test: `test/car-retire.test.js`

**Interfaces:**
- Produces: `normalizeReturnDate(value: any) -> string`(기존), `isRetiredCar(반납일: string, todayStr: string) -> boolean`(규칙 변경)

- [ ] **Step 1: 하네스·package.json 복사, `.claspignore` 갱신**

`sample` 작업 트리의 다른 테스트 파일(`trip`, `prev-odo`, `record-date`, `save-record`)은 그쪽의 미커밋 리팩터에 의존하므로 복사하지 않는다.

- [ ] **Step 2: 실패하는 테스트 작성 — `test/car-retire.test.js`**

```js
const { test } = require("node:test");
const assert = require("node:assert/strict");
const { loadGas, loadGasWithSheets, makeSheet, formatDate, parseDate } = require("./harness");

const gas = loadGas();
const TODAY = "2026-11-23";

test("반납일이 비어 있으면 반납 아님", () => {
  assert.equal(gas.isRetiredCar("", TODAY), false);
  assert.equal(gas.isRetiredCar(undefined, TODAY), false);
});
test("반납일 당일은 아직 반납 아님 — 반납하러 가는 운행 기록 허용", () => {
  assert.equal(gas.isRetiredCar("2026-11-23", TODAY), false);
});
test("반납일 다음 날부터 반납", () => {
  assert.equal(gas.isRetiredCar("2026-11-22", TODAY), true);
});
test("미래 반납일은 미리 입력해 둬도 아직 반납 아님", () => {
  assert.equal(gas.isRetiredCar("2026-11-30", TODAY), false);
});
test("날짜 형식이 아닌 값은 즉시 반납", () => {
  assert.equal(gas.isRetiredCar("반납", TODAY), true);
});
test("Date 셀은 KST yyyy-MM-dd 로 정규화", () => {
  assert.equal(gas.normalizeReturnDate(parseDate("2026-11-23", "Asia/Seoul", "yyyy-MM-dd")), "2026-11-23");
  assert.equal(gas.normalizeReturnDate(" 2026-11-23 "), "2026-11-23");
  assert.equal(gas.normalizeReturnDate(null), "");
  assert.equal(gas.normalizeReturnDate(new Date(NaN)), "");
});
```

주차 보드 테스트 2개도 같은 파일에 넣는다. `loadGasWithSheets`로 RAW 시트를 만든다(헤더 1행 + 차량 `78라1234` 행 1개, 17열, `COL.플래그="정상"`, `COL.타임스탬프`=현재 시각 문자열, `COL.주차위치="B1"`). `scriptProps.CAR_META_JSON`에 반납일을 넣고 `gas._buildParkingBoard()`를 호출한다.
- `"주차 보드: 반납일 당일 차량은 표시된다"`: 반납일 = 오늘 KST → 결과에 `carNo === "78라1234"` 포함
- `"주차 보드: 반납일 지난 차량은 빠진다"`: 반납일 = 어제 KST → 결과에 없음

- [ ] **Step 3: 실패 확인**

Run: `npm test`
Expected: FAIL — "반납일 당일은 아직 반납 아님", "주차 보드: 반납일 당일 차량은 표시된다" 2건(현재 규칙이 `<=`)

- [ ] **Step 4: `isRetiredCar` 비교를 `반납일 < todayStr`로 변경, 위 주석도 "당일까지 허용"으로 수정**

- [ ] **Step 5: 통과 확인**

Run: `npm test`
Expected: PASS 전부

- [ ] **Step 6: 커밋** — 이어받은 미커밋 변경과 spec·plan 문서를 함께 커밋한다.

```bash
git add Code.js .claspignore package.json test/harness.js test/car-retire.test.js \
  docs/superpowers/specs/2026-11-23-car-retire-add-design.md docs/superpowers/plans/2026-11-23-car-retire-add.md
git commit  # "feat: 반납 차량 판정(반납일 당일까지 허용) + 주차 보드 제외 + 단위테스트 하네스"
```

---

### Task 2: 반납 차량 입력 차단 — 폼·저장·수정

**Files:**
- Modify: `Code.js` — `serveForm`(:99), `saveRecord`(:317, `props` 로드 직후 :348), `_updateRecordInner`(:467, `props` 로드 직후 :481), `isRetiredCar` 아래에 helper 2개 추가
- Modify: `test/harness.js` — `loadGasWithSheets`에 `HtmlService` 스텁 추가
- Test: `test/car-retire.test.js`

**Interfaces:**
- Consumes: `isRetiredCar`, `normalizeReturnDate`(Task 1)
- Produces:
  - `retiredReturnDate(carNo: string, props: object, now: Date) -> string` — 반납이면 메타의 반납일, 아니면 `""`
  - `assertCarActive(carNo: string, props: object, now: Date) -> void` — 반납이면 `throw new Error(\`반납된 차량입니다 (반납일 ${반납일}).\`)`

- [ ] **Step 1: 하네스에 `HtmlService` 스텁 추가**

`createHtmlOutput(html)`은 `getContent()`가 html을 돌려주고 `setTitle`/`addMetaTag`/`setXFrameOptionsMode`가 자기 자신을 반환하는 객체를 만든다. `XFrameOptionsMode: { ALLOWALL: "ALLOWALL" }`도 넣는다.

- [ ] **Step 2: 실패하는 테스트 작성**

save-record 픽스처 방식과 같다. 차량 탭은 1~14행 헤더, 15행에 T열=1000. RAW는 헤더 1행. `CAR_META_JSON = { "12가3456": { 차종: "쏘나타", 반납일 } }`.
- `"반납일 지난 차량 saveRecord 는 throw 하고 RAW·차량 탭에 아무것도 쓰지 않는다"`: 반납일=어제 → `assert.throws(..., /반납된 차량입니다 \(반납일 \d{4}-\d{2}-\d{2}\)/)`, RAW `getLastRow() === 1`, 차량 탭 `getLastRow() === 15`
- `"반납일 당일 saveRecord 는 정상 저장된다"`: 반납일=오늘 → `res.success === true`, RAW `getLastRow() === 2`
- `"반납일 지난 차량 updateRecord 는 throw 한다"`: 반납일=어제, `originalId: "x"`, `carRowIndex: 15` → throws 같은 문구
- `"반납일 지난 차량 serveForm 은 폼 대신 안내 페이지를 돌려준다"`: `gas.serveForm({ parameter: { car: "12가3456" } }).getContent()`가 `"반납된 차량입니다"`와 반납일 문자열을 포함

payload는 `{ carNo, name: "홍길동", dept: "영업팀", useType: "일반업무용", prevOdo: 1000, currentOdo: 1100, note: "", parking: "B1" }`이다.

- [ ] **Step 3: 실패 확인** — Run: `npm test` / Expected: 위 4건 중 "당일" 제외 3건 FAIL

- [ ] **Step 4: 구현**
  - `retiredReturnDate`: `carMeta[carNo]?.반납일`과 `Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd")`로 `isRetiredCar`를 판정한다.
  - `assertCarActive`: `retiredReturnDate`를 써서 판정하고, 반납이면 위 문구로 throw한다.
  - `saveRecord`·`_updateRecordInner`: `props` 로드 직후, 모든 쓰기 전에 `assertCarActive(차량번호, props, now)`를 호출한다.
  - `serveForm`: `carNo`가 있으면 `prevOdoData` 조회 전에 `retiredReturnDate`를 확인한다. 반납이면 `HtmlService.createHtmlOutput(...)`으로 차량번호와 차단 문구를 담은 짧은 페이지를 반환하고, `.setTitle("반납된 차량")`, `.addMetaTag("viewport", "width=device-width, initial-scale=1")`, `.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)`를 적용한다.
  - carNo는 메타 키와 일치할 때만 이 경로에 오므로 HTML 이스케이프는 필요 없다.

- [ ] **Step 5: 통과 확인** — Run: `npm test` / Expected: PASS 전부

- [ ] **Step 6: 커밋** — `git add Code.js test/harness.js test/car-retire.test.js` → `feat: 반납 차량 폼·저장·수정 차단`

---

### Task 3: `syncCarRegistry` — 탭 숨김·복원·생성 + 메타 갱신

**Files:**
- Modify: `Code.js`
  - `CONFIG`(:16)에 `SHEET_CAR_TEMPLATE: "차량_양식"` 추가
  - `EXPORT_EXCLUDE_SHEETS`(:674)에 `CONFIG.SHEET_CAR_TEMPLATE` 추가
  - `setupProperties`(:838)의 carMeta 루프를 `_buildCarMeta`로 교체
  - `setupProperties` 아래에 `syncCarRegistry` 추가
- Modify: `test/harness.js` — 가짜 시트·스프레드시트 확장
- Modify: `CLAUDE.md` — "수동 1회 작업"과 운영 절차
- Test: `test/car-registry.test.js`

**Interfaces:**
- Consumes: `normalizeReturnDate`, `_prevOdoCacheKey`, `getLastDataRow`, `getPrevOdoData`(기존)
- Produces:
  - `_buildCarMeta(rows: any[][]) -> { [차량번호]: { 차종, 법인명, 사업자번호, 반납일 } }` — rows는 헤더를 제외한 마스터 값이다. 차량번호는 `String(r[0]).trim()`이고, 빈 값은 건너뛴다. 반납일은 `normalizeReturnDate(r[5])`이다.
  - `syncCarRegistry() -> { created: string[], hidden: string[], shown: string[], warnings: string[] }`

- [ ] **Step 1: 하네스 확장**
  - `makeSheet`: `hideSheet()`/`showSheet()`/`isSheetHidden()`, range `setValue(v)`, `setName(n)`(소속 스프레드시트의 이름 맵 키도 바꿈), `copyTo(ss)`(grid를 깊은 복사한 새 시트를 `"Copy of <name>"`으로 대상 ss에 등록하고 반환. 숨김 상태도 복사)
  - `makeSpreadsheet`: 시트가 자기 ss를 알도록 연결한다.

- [ ] **Step 2: 실패하는 테스트 작성 — `test/car-registry.test.js`**

픽스처:
- `차량_마스터`: 헤더 `["차량번호","차종","메모","법인명","사업자번호",""]`(F1 비움)
- `차량_양식`: 14행 머리말(A1=`"운행기록부"`) + **15행 T열에 999 잔재**, 숨김
- 기존 차량 탭 `12가3456`
- `CAR_META_JSON`은 `"{}"`에서 시작한다.

| 테스트 | 마스터 행 | 기대 |
|---|---|---|
| 반납일 있는 차량 탭은 숨겨진다 | `["12가3456","쏘나타","","(주)샘플","123-45-67890","2026-11-30"]` | 탭 `isSheetHidden()===true`, `hidden` 포함, `CAR_META_JSON["12가3456"].반납일 === "2026-11-30"` |
| 반납일을 지우면 숨김이 풀린다 | 같은 행, F="" + 탭을 미리 `hideSheet()` | `isSheetHidden()===false`, `shown` 포함 |
| 신규 차량은 차량_양식 복제로 탭이 생긴다 | `["78라1234","K5","","(주)샘플","123-45-67890",""]` | `getSheetByName("78라1234")` 존재, A1=`"운행기록부"`, 숨김 아님, `created` 포함, 메타 등록 |
| 새 탭은 양식의 15행 이하 잔재를 비운다 | 위와 같음 | `gas.getPrevOdoData("78라1234").prevOdo === null` |
| 차량번호 앞뒤 공백은 trim 한다 | `[" 78라1234 ", ...]` | `getSheetByName("78라1234")` 존재, `" 78라1234 "` 시트 없음, 메타 키 `"78라1234"` |
| 날짜가 아닌 반납일은 경고를 남긴다 | F=`"반납"` | `warnings.some(w => w.includes("12가3456"))`, `Logger._logs`에도 포함 |
| 두 번 실행해도 결과가 같다 | 신규 + 반납 행 | 두 번째 실행의 `created`·`hidden`·`shown` 모두 빈 배열, 시트 수 동일 |
| F1 헤더가 비면 반납일을 쓴다 | 아무 행 | 마스터 F1 === `"반납일"` |
| 차량_양식이 없는데 만들 탭이 있으면 변경 전에 throw | 신규 행, 양식 시트 제거 | `throws(/차량_양식/)`, `CAR_META_JSON === "{}"` 그대로 |
| 만들 탭이 없으면 양식이 없어도 동작 | 기존 차량 행만, 양식 시트 제거 | throw 없음 |

`_buildCarMeta` 단위 테스트 1건: 빈 차량번호 행은 건너뛰고, Date 반납일을 정규화한다.

- [ ] **Step 3: 실패 확인** — Run: `npm test` / Expected: `syncCarRegistry is not a function` 등으로 FAIL

- [ ] **Step 4: `_buildCarMeta` 추출, `setupProperties`를 이 함수로 교체**

- [ ] **Step 5: `syncCarRegistry` 구현**

순서는 다음과 같다.
1. 마스터를 읽고 `_buildCarMeta`를 만든다.
2. 탭이 없는 신규 차량이 있는데 `차량_양식`이 없으면 **아무것도 바꾸기 전에** throw한다. 문구: `차량_양식 시트가 없습니다. 기존 차량 탭을 복제해 15행 이하와 머리말의 차량 정보를 비우고 이름을 차량_양식으로 바꾼 뒤 다시 실행하세요.`
3. F1이 비었으면 `반납일`을 쓴다.
4. 차량별로 숨김·복원·생성을 처리한다. 생성은 `copyTo(ss)` → `setName(차량번호)` → `DATA_START_ROW`부터 마지막 행까지 `CAR_TOTAL_COLS`열 `clearContent` → `showSheet()` 순서다.
5. 반납일이 비어 있지 않고 `yyyy-MM-dd`가 아니면 warning을 남긴다.
6. `CAR_META_JSON`을 저장한다.
7. 캐시 `parking_board`와 마스터 전 차량의 `_prevOdoCacheKey`를 지운다.
8. 결과 요약과 warnings를 `Logger.log`로 남기고 결과 객체를 반환한다.

- [ ] **Step 6: 통과 확인** — Run: `npm test` / Expected: PASS 전부

- [ ] **Step 7: `CLAUDE.md` 갱신**
  - "수동 1회 작업"에 추가: `차량_양식` 시트 준비(test/prod 각각). 기존 차량 탭을 복제하고 15행 이하와 머리말 차량 정보를 비운 뒤 이름을 `차량_양식`으로 바꾸고 숨긴다.
  - 새 소절 "차량 반납·추가":
    - 반납: F열 `반납일` 입력 → **반드시** `syncCarRegistry` 실행. 실행 전에는 차단되지 않는다.
    - 추가: 마스터 행 추가 → `syncCarRegistry` → 새 탭 머리말 기입 → 폼에서 초기값 등록. 전담·출장 차량이면 `config.html` 수정·배포 후 `setupProperties`도 실행한다.
    - 반납일 당일까지는 입력을 허용한다.

- [ ] **Step 8: 커밋** — `git add Code.js test/harness.js test/car-registry.test.js CLAUDE.md` → `feat: syncCarRegistry — 차량 탭 숨김·복원·양식 복제 생성`

---

## 배포 후 수동 작업 (사용자 안내용, 구현 범위 밖)

"test 적용" 지시가 오면 `CLAUDE.md` 절차로 push하고 워크플로 성공을 확인한다. 그다음 test 환경에서 아래를 진행하도록 안내한다.
1. `차량_양식` 시트를 준비한다.
2. GAS 에디터에서 `syncCarRegistry`를 실행해 F1 헤더와 `CAR_META_JSON`을 확인한다.
3. 반납일을 넣은 차량으로 다음을 확인한다.
   - 폼 URL을 열면 안내 페이지가 나오는지
   - 주차 보드에서 빠지는지
   - 탭이 숨겨지는지
4. prod에서도 1~2를 반복한다.
