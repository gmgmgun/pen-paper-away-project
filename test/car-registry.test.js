const { test } = require("node:test");
const assert = require("node:assert/strict");
const { loadGas, loadGasWithSheets, makeSheet, parseDate } = require("./harness");

const MASTER = "차량_마스터";
const RAW = "RAW_운행일지";
const OLD_CAR = "12가3456";
const ACTIVE_CAR = "34나5678";
const NEW_CAR = "78라1234";
const MASTER_HEADER = [
  "차량번호",
  "차종",
  "메모",
  "법인명",
  "사업자번호",
  "",
  "",
  "전담자",
  "전담자부서",
  "출장차량",
];

function masterRow(
  carNo,
  반납일 = "",
  위치기록 = "",
  전담자 = "",
  전담자부서 = "",
  출장 = "",
) {
  return [
    carNo,
    "K5",
    "",
    "(주)샘플",
    "123-45-67890",
    반납일,
    위치기록,
    전담자,
    전담자부서,
    출장,
  ];
}

// 1~14행 머리말 + 15행 데이터 1건(T열 = 주행후)
function carTabGrid(header, odo) {
  const grid = [];
  for (let i = 0; i < 14; i++) grid.push(new Array(44).fill(""));
  grid[0][0] = header;
  const seed = new Array(44).fill("");
  seed[0] = "11/20(목)";
  seed[19] = odo;
  grid.push(seed);
  return grid;
}

const STAFF_GRID = [
  ["이름", "부서"],
  ["홍길동", "영업팀"],
];
const CLIENTS_GRID = [["거래처"], ["A상사"]];

// OLD_CAR 탭은 항상 있다(15행 주행후 1000). extraTabs 로 다른 차량 탭을 더 둘 수 있다.
// rawCars 가 있으면 차량번호만 채운 RAW_운행일지 시트를 만든다.
// staffGrid · clientsGrid 가 null 이면 그 설정 시트를 만들지 않는다.
function registryFixture({
  rows,
  fHeader = "",
  gHeader = "",
  masterHeader = MASTER_HEADER,
  extraTabs = {},
  rawCars = [],
  staffGrid = STAFF_GRID,
  clientsGrid = CLIENTS_GRID,
  scriptProps = {},
}) {
  const header = masterHeader.slice();
  header[5] = fHeader;
  header[6] = gHeader;
  const sheets = {
    [MASTER]: makeSheet(MASTER, [header, ...rows]),
    [OLD_CAR]: makeSheet(OLD_CAR, carTabGrid("운행기록부", 1000)),
  };
  if (staffGrid) sheets["직원"] = makeSheet("직원", staffGrid);
  if (clientsGrid) sheets["거래처"] = makeSheet("거래처", clientsGrid);
  Object.entries(extraTabs).forEach(([name, grid]) => {
    sheets[name] = makeSheet(name, grid);
  });
  if (rawCars.length > 0) {
    const rawRows = rawCars.map((carNo) => {
      const r = new Array(17).fill("");
      r[1] = carNo; // COL.차량번호
      return r;
    });
    sheets[RAW] = makeSheet(RAW, [new Array(17).fill("헤더"), ...rawRows]);
  }
  return loadGasWithSheets({
    sheets,
    scriptProps: Object.assign({ CAR_META_JSON: "{}" }, scriptProps),
  });
}

function meta(props) {
  return JSON.parse(props.CAR_META_JSON);
}

test("반납일 있는 차량 탭은 숨겨진다", () => {
  const { gas, ss, props } = registryFixture({
    rows: [masterRow(OLD_CAR, "2026-11-30")],
  });

  const res = gas.syncSettings();

  assert.equal(ss.getSheetByName(OLD_CAR).isSheetHidden(), true);
  assert.ok(res.hidden.includes(OLD_CAR));
  assert.equal(meta(props)[OLD_CAR].반납일, "2026-11-30");
});

test("반납일을 지우면 숨김이 풀린다", () => {
  const { gas, ss } = registryFixture({ rows: [masterRow(OLD_CAR, "")] });
  ss.getSheetByName(OLD_CAR).hideSheet();

  const res = gas.syncSettings();

  assert.equal(ss.getSheetByName(OLD_CAR).isSheetHidden(), false);
  assert.ok(res.shown.includes(OLD_CAR));
});

test("신규 차량은 기존 차량 탭 복제로 탭이 생긴다", () => {
  const { gas, ss, props } = registryFixture({
    rows: [masterRow(OLD_CAR), masterRow(NEW_CAR)],
  });

  const res = gas.syncSettings();

  const sh = ss.getSheetByName(NEW_CAR);
  assert.ok(sh, "새 탭이 있어야 함");
  assert.equal(sh.getRange(1, 1).getValue(), "운행기록부");
  assert.equal(sh.isSheetHidden(), false);
  assert.ok(res.created.includes(NEW_CAR));
  assert.equal(meta(props)[NEW_CAR].차종, "K5");
});

test("새 탭은 원본 탭의 15행 이하 기록을 비운다 — 초기값 등록 모드로 열려야 함", () => {
  const { gas } = registryFixture({
    rows: [masterRow(OLD_CAR), masterRow(NEW_CAR)],
  });

  gas.syncSettings();

  assert.equal(gas.getPrevOdoData(NEW_CAR).prevOdo, null);
  assert.equal(gas.getPrevOdoData(OLD_CAR).prevOdo, 1000, "원본 탭은 그대로");
});

test("리스 교체: 기존 차량 반납과 신규 추가를 한 번에 처리한다", () => {
  const { gas, ss } = registryFixture({
    rows: [masterRow(OLD_CAR, "2026-11-30"), masterRow(NEW_CAR)],
  });

  gas.syncSettings();

  assert.equal(ss.getSheetByName(OLD_CAR).isSheetHidden(), true);
  assert.equal(ss.getSheetByName(NEW_CAR).isSheetHidden(), false);
  assert.equal(gas.getPrevOdoData(NEW_CAR).prevOdo, null);
});

test("복제 원본은 반납 안 된 차량 탭을 우선한다", () => {
  const { gas, ss } = registryFixture({
    rows: [
      masterRow(OLD_CAR, "2026-11-30"),
      masterRow(ACTIVE_CAR),
      masterRow(NEW_CAR),
    ],
    extraTabs: { [ACTIVE_CAR]: carTabGrid("현행양식", 2000) },
  });

  gas.syncSettings();

  assert.equal(ss.getSheetByName(NEW_CAR).getRange(1, 1).getValue(), "현행양식");
});

test("새 탭을 만들면 머리말 수정 안내를 원본 차량과 함께 로그에 남긴다", () => {
  const { gas } = registryFixture({
    rows: [masterRow(OLD_CAR), masterRow(NEW_CAR)],
  });

  gas.syncSettings();

  assert.ok(
    gas.Logger._logs.some(
      (l) => l.includes("머리말") && l.includes(NEW_CAR) && l.includes(OLD_CAR),
    ),
  );
});

test("차량번호 앞뒤 공백은 trim 한다", () => {
  const { gas, ss, props } = registryFixture({
    rows: [masterRow(OLD_CAR), masterRow(` ${NEW_CAR} `)],
  });

  gas.syncSettings();

  assert.ok(ss.getSheetByName(NEW_CAR));
  assert.equal(ss.getSheetByName(` ${NEW_CAR} `), null);
  assert.ok(NEW_CAR in meta(props));
});

test("날짜가 아닌 반납일은 경고를 남긴다", () => {
  const { gas } = registryFixture({ rows: [masterRow(OLD_CAR, "반납")] });

  const res = gas.syncSettings();

  assert.ok(res.warnings.some((w) => w.includes(OLD_CAR)));
  assert.ok(gas.Logger._logs.some((l) => l.includes(OLD_CAR)));
});

test("두 번 실행해도 결과가 같다", () => {
  const { gas, ss } = registryFixture({
    rows: [masterRow(OLD_CAR, "2026-11-30"), masterRow(NEW_CAR)],
  });

  gas.syncSettings();
  const sheetCount = ss.getSheets().length;
  const res = gas.syncSettings();

  assert.deepEqual([...res.created], []);
  assert.deepEqual([...res.hidden], []);
  assert.deepEqual([...res.shown], []);
  assert.equal(ss.getSheets().length, sheetCount);
});

test("F1 헤더가 비면 반납일을 쓴다", () => {
  const { gas, ss } = registryFixture({ rows: [masterRow(OLD_CAR)] });

  gas.syncSettings();

  assert.equal(ss.getSheetByName(MASTER).getRange(1, 6).getValue(), "반납일");
});

test("복제할 기존 차량 탭이 없으면 아무것도 바꾸기 전에 throw", () => {
  // OLD_CAR 탭은 있지만 마스터에 없으므로 원본 후보가 아니다
  const { gas, ss, props } = registryFixture({ rows: [masterRow(NEW_CAR)] });

  assert.throws(
    () => gas.syncSettings(),
    (e) => /원본/.test(e.message) && e.message.includes(NEW_CAR),
  );

  assert.equal(props.CAR_META_JSON, "{}");
  assert.equal(ss.getSheetByName(MASTER).getRange(1, 6).getValue(), "");
  assert.equal(ss.getSheetByName(MASTER).getRange(1, 7).getValue(), "");
  assert.equal(ss.getSheetByName(NEW_CAR), null);
});

test("_buildCarMeta: 빈 차량번호 행은 건너뛰고 Date 반납일을 정규화한다", () => {
  const gas = loadGas();

  const m = gas._buildCarMeta([
    ["", "", "", "", "", ""],
    [
      ` ${NEW_CAR} `,
      "K5",
      "",
      "(주)샘플",
      "123-45-67890",
      parseDate("2026-11-30", "Asia/Seoul", "yyyy-MM-dd"),
    ],
  ]);

  assert.deepEqual(Object.keys(m), [NEW_CAR]);
  assert.deepEqual({ ...m[NEW_CAR] }, {
    차종: "K5",
    법인명: "(주)샘플",
    사업자번호: "123-45-67890",
    반납일: "2026-11-30",
    위치기록: true,
  });

  // G열이 도입된 마스터(hasLocationCol = true)는 행별 G 값을 따른다
  const withG = gas._buildCarMeta(
    [masterRow(OLD_CAR, "", true), masterRow(ACTIVE_CAR, "", "")],
    true,
  );
  assert.equal(withG[OLD_CAR].위치기록, true);
  assert.equal(withG[ACTIVE_CAR].위치기록, false);
});

// ── F열 보호: 다른 용도로 쓰이던 F열을 반납일로 오인하면 전 차량이 차단된다 ──

test("F1 이 반납일이 아닌 다른 헤더면 syncSettings 는 아무것도 바꾸기 전에 throw", () => {
  const { gas, ss, props } = registryFixture({
    rows: [masterRow(OLD_CAR, "리스 만기 2027-03")],
    fHeader: "비고",
  });

  assert.throws(() => gas.syncSettings(), /F1/);

  assert.equal(props.CAR_META_JSON, "{}");
  assert.equal(ss.getSheetByName(OLD_CAR).isSheetHidden(), false);
});

// ── G열 위치기록여부: 주차 보드 표시 여부 ──

test("isChecked: true 와 문자열 TRUE 만 TRUE", () => {
  const gas = loadGas();
  for (const v of [true, "TRUE", " true "]) {
    assert.equal(gas.isChecked(v), true, String(v));
  }
  for (const v of [false, "", "FALSE", null, undefined, "Y", 1]) {
    assert.equal(gas.isChecked(v), false, String(v));
  }
});

test("G1 이 위치기록여부면 행별 G 값을 CAR_META_JSON 위치기록으로 담는다", () => {
  const { gas, props } = registryFixture({
    rows: [masterRow(OLD_CAR, "", true), masterRow(ACTIVE_CAR, "", false)],
    gHeader: "위치기록여부",
    extraTabs: { [ACTIVE_CAR]: carTabGrid("운행기록부", 2000) },
  });

  gas.syncSettings();

  assert.equal(meta(props)[OLD_CAR].위치기록, true);
  assert.equal(meta(props)[ACTIVE_CAR].위치기록, false);
});

test("G1 이 다른 헤더면 syncSettings 는 아무것도 바꾸기 전에 throw", () => {
  const { gas, ss, props } = registryFixture({
    rows: [masterRow(OLD_CAR, "2026-11-30")],
    gHeader: "비고",
  });

  assert.throws(() => gas.syncSettings(), /G1/);

  assert.equal(props.CAR_META_JSON, "{}");
  assert.equal(ss.getSheetByName(OLD_CAR).isSheetHidden(), false);
  assert.equal(ss.getSheetByName(MASTER).getRange(1, 6).getValue(), "");
});

// ── G열 첫 도입(마이그레이션)과 RAW 미등록 차량 경고 ──

test("G1 이 비면 첫 sync 가 헤더와 체크박스를 쓴다 — 반납 아닌 차량만 체크", () => {
  const { gas, ss, props } = registryFixture({
    rows: [masterRow(OLD_CAR), masterRow(ACTIVE_CAR, "2020-01-01")],
    extraTabs: { [ACTIVE_CAR]: carTabGrid("운행기록부", 2000) },
  });

  gas.syncSettings();

  const master = ss.getSheetByName(MASTER);
  assert.equal(master.getRange(1, 7).getValue(), "위치기록여부");
  assert.equal(master.getRange(2, 7).getValue(), true);
  assert.equal(master.getRange(3, 7).getValue(), false);
  assert.ok(
    master._checkboxRanges.some(
      (r) => r.row === 2 && r.col === 7 && r.numRows === 2,
    ),
  );
  assert.equal(meta(props)[OLD_CAR].위치기록, true);
  assert.equal(meta(props)[ACTIVE_CAR].위치기록, false);
});

test("첫 sync 마이그레이션: 미래 반납일 차량은 아직 운행 중이라 체크", () => {
  const { gas, ss } = registryFixture({
    rows: [masterRow(OLD_CAR, "2999-12-31")],
  });

  gas.syncSettings();

  assert.equal(ss.getSheetByName(MASTER).getRange(2, 7).getValue(), true);
});

test("첫 sync 마이그레이션: 차량번호가 빈 행이 있어도 깨지지 않고 그 행은 FALSE", () => {
  const { gas, ss, props } = registryFixture({
    rows: [
      masterRow(OLD_CAR),
      ["", "", "", "", "", "", ""],
      masterRow(ACTIVE_CAR),
    ],
    extraTabs: { [ACTIVE_CAR]: carTabGrid("운행기록부", 2000) },
  });

  gas.syncSettings();

  const master = ss.getSheetByName(MASTER);
  assert.equal(master.getRange(2, 7).getValue(), true);
  assert.equal(master.getRange(3, 7).getValue(), false);
  assert.equal(master.getRange(4, 7).getValue(), true);
  assert.equal(meta(props)[OLD_CAR].위치기록, true);
  assert.equal(meta(props)[ACTIVE_CAR].위치기록, true);
});

test("G1 이 이미 있으면 sync 는 G 값을 덮어쓰지 않는다", () => {
  const { gas, ss } = registryFixture({
    rows: [masterRow(OLD_CAR, "", false)],
    gHeader: "위치기록여부",
  });

  gas.syncSettings();
  gas.syncSettings();

  const master = ss.getSheetByName(MASTER);
  assert.equal(master.getRange(2, 7).getValue(), false);
  assert.equal(master._checkboxRanges.length, 0);
});

test("RAW 에만 있는 차량번호는 sync 경고로 남는다", () => {
  const { gas } = registryFixture({
    rows: [masterRow(OLD_CAR)],
    rawCars: [OLD_CAR, "99하9999", "99하9999 "],
  });

  const res = gas.syncSettings();

  assert.equal(res.warnings.filter((w) => w.includes("99하9999")).length, 1);
  assert.ok(!res.warnings.some((w) => w.includes(OLD_CAR)));
  assert.ok(
    gas.Logger._logs.some((l) => l.includes("경고:") && l.includes("99하9999")),
  );
});

// ── 리뷰 반영: G1 빈칸인데 G열에 값이 있으면 덮어쓰지 않는다 / 신규 차량 G 미체크 경고 ──

test("G1 이 비었는데 G열에 값이 있으면 sync 는 아무것도 바꾸기 전에 throw", () => {
  // 도입 후 G1 을 실수로 지운 상황 — 마이그레이션이 체크 해제를 되돌리면 안 된다
  const { gas, ss, props } = registryFixture({
    rows: [masterRow(OLD_CAR, "", false)],
  });

  assert.throws(() => gas.syncSettings(), /G1/);

  const master = ss.getSheetByName(MASTER);
  assert.equal(master.getRange(2, 7).getValue(), false);
  assert.equal(master.getRange(1, 6).getValue(), "");
  assert.equal(master._checkboxRanges.length, 0);
  assert.equal(props.CAR_META_JSON, "{}");
});

test("새 탭을 만든 차량의 G 가 체크되지 않았으면 보드에 안 나온다고 경고", () => {
  const { gas } = registryFixture({
    rows: [masterRow(OLD_CAR, "", true), masterRow(NEW_CAR, "", "")],
    gHeader: "위치기록여부",
  });

  const res = gas.syncSettings();

  assert.ok(res.created.includes(NEW_CAR));
  assert.ok(
    res.warnings.some((w) => w.includes(NEW_CAR) && w.includes("위치기록여부")),
  );
  assert.ok(!res.warnings.some((w) => w.includes(OLD_CAR)));
});

test("새 탭을 만든 차량의 G 가 체크돼 있으면 경고 없음", () => {
  const { gas } = registryFixture({
    rows: [masterRow(OLD_CAR, "", true), masterRow(NEW_CAR, "", true)],
    gHeader: "위치기록여부",
  });

  const res = gas.syncSettings();

  assert.ok(res.created.includes(NEW_CAR));
  assert.equal(res.warnings.length, 0);
});

// ── 설정 시트 → props (직원 · 거래처 · 전담 · 출장) ──

test("syncSettings 는 직원·거래처·전담·출장을 props 에 쓴다", () => {
  const { gas, props } = registryFixture({
    rows: [
      masterRow(OLD_CAR, "", "", "김전담", "임원"),
      masterRow(ACTIVE_CAR, "", "", "", "", true),
    ],
    extraTabs: { [ACTIVE_CAR]: carTabGrid("운행기록부", 2000) },
  });

  gas.syncSettings();

  assert.deepEqual(JSON.parse(props.STAFF_JSON), [{ name: "홍길동", dept: "영업팀" }]);
  assert.deepEqual(JSON.parse(props.CLIENTS_JSON), ["A상사"]);
  assert.deepEqual(JSON.parse(props.FIXED_USER_JSON), {
    [OLD_CAR]: { name: "김전담", dept: "임원" },
  });
  assert.deepEqual(JSON.parse(props.BUSINESS_TRIP_CARS_JSON), [ACTIVE_CAR]);
});

test("직원 시트가 없으면 syncSettings 는 아무것도 바꾸기 전에 throw", () => {
  const OLD_STAFF = '[{"name":"옛직원","dept":""}]';
  const { gas, ss, props } = registryFixture({
    rows: [masterRow(NEW_CAR)],
    staffGrid: null,
    scriptProps: { STAFF_JSON: OLD_STAFF },
  });

  assert.throws(() => gas.syncSettings(), /migrateConfigToSheets/);
  assert.equal(props.STAFF_JSON, OLD_STAFF);
  assert.equal(ss.getSheetByName(MASTER).getRange(1, 6).getValue(), "");
  assert.equal(ss.getSheetByName(NEW_CAR), null);
});

test("H1:J1 이 없으면 syncSettings 는 아무것도 바꾸기 전에 throw", () => {
  const { gas, ss } = registryFixture({
    rows: [masterRow(NEW_CAR)],
    masterHeader: MASTER_HEADER.slice(0, 7),
  });

  assert.throws(() => gas.syncSettings(), /H1:J1/);
  assert.equal(ss.getSheetByName(MASTER).getRange(1, 6).getValue(), "");
  assert.equal(ss.getSheetByName(NEW_CAR), null);
});

test("직원 이름 중복 경고가 sync 로그에 남는다", () => {
  const { gas } = registryFixture({
    rows: [masterRow(OLD_CAR)],
    staffGrid: [
      ["이름", "부서"],
      ["홍길동", "영업팀"],
      ["홍길동", "관리팀"],
    ],
  });

  gas.syncSettings();

  assert.ok(
    gas.Logger._logs.some((l) => l.includes("경고:") && l.includes("홍길동")),
  );
});

test("setupProperties 는 더 이상 없다", () => {
  const { gas } = registryFixture({ rows: [masterRow(OLD_CAR)] });

  assert.equal(typeof gas.setupProperties, "undefined");
  assert.equal(typeof gas.syncCarRegistry, "undefined");
});

test("SPREADSHEET_ID 가 없으면 getSpreadsheet 는 원인을 말하며 throw", () => {
  const { gas } = loadGasWithSheets({
    sheets: {},
    scriptProps: { SPREADSHEET_ID: null },
  });

  assert.throws(() => gas.getSpreadsheet(), /SPREADSHEET_ID 스크립트 속성이 없습니다/);
});
