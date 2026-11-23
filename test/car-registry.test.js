const { test } = require("node:test");
const assert = require("node:assert/strict");
const { loadGas, loadGasWithSheets, makeSheet, parseDate } = require("./harness");

const MASTER = "차량_마스터";
const TEMPLATE = "차량_양식";
const OLD_CAR = "12가3456";
const NEW_CAR = "78라1234";
const MASTER_HEADER = ["차량번호", "차종", "메모", "법인명", "사업자번호", ""];

function masterRow(carNo, 반납일 = "") {
  return [carNo, "K5", "", "(주)샘플", "123-45-67890", 반납일];
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

const CONFIG_FILE = JSON.stringify({
  staff: [],
  fixedUser: {},
  businessTripCars: [],
  clients: [],
});

function registryFixture({ rows, withTemplate = true, fHeader = "" }) {
  const header = MASTER_HEADER.slice();
  header[5] = fHeader;
  const sheets = {
    [MASTER]: makeSheet(MASTER, [header, ...rows]),
    [OLD_CAR]: makeSheet(OLD_CAR, carTabGrid("운행기록부", 1000)),
  };
  if (withTemplate) {
    // 양식에 15행 잔재(999)가 남아 있는 상황을 일부러 만든다
    sheets[TEMPLATE] = makeSheet(TEMPLATE, carTabGrid("운행기록부", 999));
    sheets[TEMPLATE].hideSheet();
  }
  return loadGasWithSheets({
    sheets,
    scriptProps: { CAR_META_JSON: "{}" },
    files: { config: CONFIG_FILE },
  });
}

function meta(props) {
  return JSON.parse(props.CAR_META_JSON);
}

test("반납일 있는 차량 탭은 숨겨진다", () => {
  const { gas, ss, props } = registryFixture({
    rows: [masterRow(OLD_CAR, "2026-11-30")],
  });

  const res = gas.syncCarRegistry();

  assert.equal(ss.getSheetByName(OLD_CAR).isSheetHidden(), true);
  assert.ok(res.hidden.includes(OLD_CAR));
  assert.equal(meta(props)[OLD_CAR].반납일, "2026-11-30");
});

test("반납일을 지우면 숨김이 풀린다", () => {
  const { gas, ss } = registryFixture({ rows: [masterRow(OLD_CAR, "")] });
  ss.getSheetByName(OLD_CAR).hideSheet();

  const res = gas.syncCarRegistry();

  assert.equal(ss.getSheetByName(OLD_CAR).isSheetHidden(), false);
  assert.ok(res.shown.includes(OLD_CAR));
});

test("신규 차량은 차량_양식 복제로 탭이 생긴다", () => {
  const { gas, ss, props } = registryFixture({ rows: [masterRow(NEW_CAR)] });

  const res = gas.syncCarRegistry();

  const sh = ss.getSheetByName(NEW_CAR);
  assert.ok(sh, "새 탭이 있어야 함");
  assert.equal(sh.getRange(1, 1).getValue(), "운행기록부");
  assert.equal(sh.isSheetHidden(), false);
  assert.ok(res.created.includes(NEW_CAR));
  assert.equal(meta(props)[NEW_CAR].차종, "K5");
});

test("새 탭은 양식의 15행 이하 잔재를 비운다 — 초기값 등록 모드로 열려야 함", () => {
  const { gas } = registryFixture({ rows: [masterRow(NEW_CAR)] });

  gas.syncCarRegistry();

  assert.equal(gas.getPrevOdoData(NEW_CAR).prevOdo, null);
});

test("차량번호 앞뒤 공백은 trim 한다", () => {
  const { gas, ss, props } = registryFixture({
    rows: [masterRow(` ${NEW_CAR} `)],
  });

  gas.syncCarRegistry();

  assert.ok(ss.getSheetByName(NEW_CAR));
  assert.equal(ss.getSheetByName(` ${NEW_CAR} `), null);
  assert.ok(NEW_CAR in meta(props));
});

test("날짜가 아닌 반납일은 경고를 남긴다", () => {
  const { gas } = registryFixture({ rows: [masterRow(OLD_CAR, "반납")] });

  const res = gas.syncCarRegistry();

  assert.ok(res.warnings.some((w) => w.includes(OLD_CAR)));
  assert.ok(gas.Logger._logs.some((l) => l.includes(OLD_CAR)));
});

test("두 번 실행해도 결과가 같다", () => {
  const { gas, ss } = registryFixture({
    rows: [masterRow(OLD_CAR, "2026-11-30"), masterRow(NEW_CAR)],
  });

  gas.syncCarRegistry();
  const sheetCount = ss.getSheets().length;
  const res = gas.syncCarRegistry();

  assert.deepEqual([...res.created], []);
  assert.deepEqual([...res.hidden], []);
  assert.deepEqual([...res.shown], []);
  assert.equal(ss.getSheets().length, sheetCount);
});

test("F1 헤더가 비면 반납일을 쓴다", () => {
  const { gas, ss } = registryFixture({ rows: [masterRow(OLD_CAR)] });

  gas.syncCarRegistry();

  assert.equal(ss.getSheetByName(MASTER).getRange(1, 6).getValue(), "반납일");
});

test("차량_양식이 없는데 만들 탭이 있으면 아무것도 바꾸기 전에 throw", () => {
  const { gas, ss, props } = registryFixture({
    rows: [masterRow(OLD_CAR, "2026-11-30"), masterRow(NEW_CAR)],
    withTemplate: false,
  });

  assert.throws(() => gas.syncCarRegistry(), /차량_양식/);

  assert.equal(props.CAR_META_JSON, "{}");
  assert.equal(ss.getSheetByName(OLD_CAR).isSheetHidden(), false);
});

test("만들 탭이 없으면 차량_양식이 없어도 동작한다", () => {
  const { gas } = registryFixture({
    rows: [masterRow(OLD_CAR)],
    withTemplate: false,
  });

  assert.doesNotThrow(() => gas.syncCarRegistry());
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
  });
});

// ── F열 보호: 다른 용도로 쓰이던 F열을 반납일로 오인하면 전 차량이 차단된다 ──

test("F1 이 반납일이 아닌 다른 헤더면 syncCarRegistry 는 아무것도 바꾸기 전에 throw", () => {
  const { gas, ss, props } = registryFixture({
    rows: [masterRow(OLD_CAR, "리스 만기 2027-03")],
    fHeader: "비고",
  });

  assert.throws(() => gas.syncCarRegistry(), /F1/);

  assert.equal(props.CAR_META_JSON, "{}");
  assert.equal(ss.getSheetByName(OLD_CAR).isSheetHidden(), false);
});

test("F1 이 다른 헤더면 setupProperties 도 throw 해 전 차량 차단을 막는다", () => {
  const { gas, props } = registryFixture({
    rows: [masterRow(OLD_CAR, "리스 만기 2027-03")],
    fHeader: "비고",
  });

  assert.throws(() => gas.setupProperties(), /F1/);

  assert.equal(props.CAR_META_JSON, "{}");
});

test("F1 이 반납일이면 setupProperties 는 정상 동작한다", () => {
  const { gas, props } = registryFixture({
    rows: [masterRow(OLD_CAR, "2026-11-30")],
    fHeader: "반납일",
  });

  gas.setupProperties();

  assert.equal(meta(props)[OLD_CAR].반납일, "2026-11-30");
});

test("setupProperties 도 날짜가 아닌 반납일을 경고 로그로 남긴다", () => {
  const { gas } = registryFixture({ rows: [masterRow(OLD_CAR, "반납")] });

  gas.setupProperties();

  assert.ok(
    gas.Logger._logs.some((l) => l.includes("경고") && l.includes(OLD_CAR)),
  );
});
