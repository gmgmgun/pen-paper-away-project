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
