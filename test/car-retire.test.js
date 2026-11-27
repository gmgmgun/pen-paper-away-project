const { test } = require("node:test");
const assert = require("node:assert/strict");
const {
  loadGas,
  loadGasWithSheets,
  makeSheet,
  formatDate,
  parseDate,
} = require("./harness");

const gas = loadGas();
const TODAY = "2026-11-23";

// ── 반납 판정 ─────────────────────────────────────────────────────────

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
  assert.equal(
    gas.normalizeReturnDate(parseDate("2026-11-23", "Asia/Seoul", "yyyy-MM-dd")),
    "2026-11-23",
  );
  assert.equal(gas.normalizeReturnDate(" 2026-11-23 "), "2026-11-23");
  assert.equal(gas.normalizeReturnDate(null), "");
  assert.equal(gas.normalizeReturnDate(new Date(NaN)), "");
});

// ── 주차 보드 ─────────────────────────────────────────────────────────

const RAW = "RAW_운행일지";
const BOARD_CAR = "78라1234";

function kstDateStr(offsetDays) {
  return formatDate(
    new Date(Date.now() + offsetDays * 86400000),
    "Asia/Seoul",
    "yyyy-MM-dd",
  );
}

function boardFixture(반납일) {
  const { COL } = loadGas();
  const row = new Array(17).fill("");
  row[COL.ID] = "id-1";
  row[COL.차량번호] = BOARD_CAR;
  row[COL.성명] = "홍길동";
  row[COL.부서] = "영업팀";
  row[COL.플래그] = "정상";
  row[COL.타임스탬프] = formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  row[COL.주차위치] = "B1";

  return loadGasWithSheets({
    sheets: {
      [RAW]: makeSheet(RAW, [new Array(17).fill("헤더"), row]),
    },
    scriptProps: {
      CAR_META_JSON: JSON.stringify({ [BOARD_CAR]: { 차종: "K5", 반납일 } }),
    },
  });
}

test("주차 보드: 반납일 당일 차량은 표시된다", () => {
  const { gas: g } = boardFixture(kstDateStr(0));

  const board = g._buildParkingBoard();

  assert.ok(board.some((c) => c.carNo === BOARD_CAR));
});

test("주차 보드: 반납일 지난 차량은 빠진다", () => {
  const { gas: g } = boardFixture(kstDateStr(-1));

  const board = g._buildParkingBoard();

  assert.ok(!board.some((c) => c.carNo === BOARD_CAR));
});

// ── 입력 차단: 폼·저장·수정 ───────────────────────────────────────────

const CAR = "12가3456";
const DATA_START_ROW = 15;
const BLOCK_MSG = /반납된 차량입니다 \(반납일 \d{4}-\d{2}-\d{2}\)\./;

// 차량 탭: 1~14행 헤더, 15행에 주행후 1000. RAW 는 헤더 1행.
function retireFixture(반납일) {
  const carGrid = [];
  for (let i = 0; i < 14; i++) carGrid.push(new Array(44).fill(""));
  carGrid[13][0] = "날짜";
  const seed = new Array(44).fill("");
  seed[0] = "11/20(목)";
  seed[19] = 1000; // T열 = 주행후
  carGrid.push(seed);

  return loadGasWithSheets({
    sheets: {
      [RAW]: makeSheet(RAW, [new Array(17).fill("헤더")]),
      [CAR]: makeSheet(CAR, carGrid),
    },
    scriptProps: {
      CAR_META_JSON: JSON.stringify({ [CAR]: { 차종: "쏘나타", 반납일 } }),
    },
  });
}

function savePayload(overrides) {
  return Object.assign(
    {
      carNo: CAR,
      name: "홍길동",
      dept: "영업팀",
      useType: "일반업무용",
      prevOdo: 1000,
      currentOdo: 1100,
      note: "",
      parking: "B1",
    },
    overrides,
  );
}

test("반납일 지난 차량 saveRecord 는 throw 하고 RAW·차량 탭에 아무것도 쓰지 않는다", () => {
  const { gas: g, ss } = retireFixture(kstDateStr(-1));

  assert.throws(() => g.saveRecord(savePayload()), BLOCK_MSG);

  assert.equal(ss.getSheetByName(RAW).getLastRow(), 1);
  assert.equal(ss.getSheetByName(CAR).getLastRow(), DATA_START_ROW);
});

test("반납일 당일 saveRecord 는 정상 저장된다", () => {
  const { gas: g, ss } = retireFixture(kstDateStr(0));

  const res = g.saveRecord(savePayload());

  assert.equal(res.success, true);
  assert.equal(ss.getSheetByName(RAW).getLastRow(), 2);
});

test("반납일 지난 차량 updateRecord 는 throw 한다", () => {
  const { gas: g, ss } = retireFixture(kstDateStr(-1));

  assert.throws(
    () =>
      g.updateRecord(
        savePayload({ originalId: "x", carRowIndex: DATA_START_ROW }),
      ),
    BLOCK_MSG,
  );
  assert.equal(ss.getSheetByName(RAW).getLastRow(), 1);
});

test("반납일 지난 차량 serveForm 은 폼 대신 안내 페이지를 돌려준다", () => {
  const 반납일 = kstDateStr(-1);
  const { gas: g } = retireFixture(반납일);

  const html = g.serveForm({ parameter: { car: CAR } }).getContent();

  assert.match(html, /반납된 차량입니다/);
  assert.ok(html.includes(반납일));
  assert.ok(html.includes(CAR));
});
