const { test } = require("node:test");
const assert = require("node:assert/strict");
const { loadGas, loadGasWithSheets, makeSheet, formatDate } = require("./harness");

const RAW = "RAW_운행일지";
const A = "12가3456";
const B = "34나5678";

const NOW_TS = () => formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

// meta: CAR_META_JSON 객체. raw: RAW 행 정의 배열 (17열, COL 인덱스로 채움).
function boardFixture({ meta, raw }) {
  const { COL } = loadGas();
  const rows = raw.map((r, i) => {
    const row = new Array(17).fill("");
    row[COL.ID] = `id-${i + 1}`;
    row[COL.차량번호] = r.carNo;
    row[COL.성명] = r.name === undefined ? "홍길동" : r.name;
    row[COL.부서] = r.dept === undefined ? "영업팀" : r.dept;
    row[COL.플래그] = r.flag || "정상";
    row[COL.타임스탬프] = r.ts || NOW_TS();
    row[COL.주차위치] = r.parking || "B1";
    return row;
  });

  const { gas } = loadGasWithSheets({
    sheets: {
      [RAW]: makeSheet(RAW, [new Array(17).fill("헤더"), ...rows]),
    },
    scriptProps: { CAR_META_JSON: JSON.stringify(meta) },
  });
  return gas;
}

test("RAW 기록이 없는 마스터 차량도 미기록 카드로 나온다", () => {
  const g = boardFixture({
    meta: { [A]: { 차종: "K5", 반납일: "", 위치기록: true } },
    raw: [],
  });

  assert.deepEqual({ ...g._buildParkingBoard()[0] }, {
    carNo: A,
    carName: "K5",
    parking: "",
    name: "",
    time: "—",
  });
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

test("초기값등록 뒤 운행 기록이 있으면 최신 운행 기록을 쓴다", () => {
  const g = boardFixture({
    meta: { [A]: { 차종: "K5", 위치기록: true } },
    raw: [
      {
        carNo: A,
        flag: "초기값등록",
        parking: "B2",
        name: "",
        dept: "",
        ts: "2026-11-20 09:00:00",
      },
      { carNo: A, parking: "B3", ts: "2026-11-21 09:00:00" },
    ],
  });

  const card = g._buildParkingBoard()[0];

  assert.equal(card.parking, "B3");
  assert.equal(card.name, "영업팀 홍길동");
});

test("위치기록 FALSE 차량은 보드에서 빠진다", () => {
  const g = boardFixture({
    meta: { [A]: { 차종: "K5", 위치기록: false } },
    raw: [{ carNo: A }],
  });

  assert.equal(g._buildParkingBoard().length, 0);
});

test("마스터에 없는 차량번호는 RAW 기록이 있어도 빠진다", () => {
  const g = boardFixture({
    meta: { [A]: { 차종: "K5", 위치기록: true } },
    raw: [{ carNo: B }],
  });

  assert.deepEqual(
    [...g._buildParkingBoard().map((c) => c.carNo)],
    [A],
  );
});

test("카드 순서는 마스터 순서를 따른다", () => {
  const g = boardFixture({
    meta: {
      [A]: { 차종: "K5", 위치기록: true },
      [B]: { 차종: "쏘나타", 위치기록: true },
    },
    raw: [{ carNo: B }, { carNo: A }],
  });

  assert.deepEqual(
    [...g._buildParkingBoard().map((c) => c.carNo)],
    [A, B],
  );
});

test("위치기록 키가 없는 예전 CAR_META_JSON 은 표시한다", () => {
  const g = boardFixture({
    meta: { [A]: { 차종: "K5" } },
    raw: [],
  });

  assert.equal(g._buildParkingBoard().length, 1);
});

test("RAW 차량번호 뒤 공백은 마스터 차량에 붙는다", () => {
  const g = boardFixture({
    meta: { [A]: { 차종: "K5", 위치기록: true } },
    raw: [{ carNo: `${A} `, parking: "B1" }],
  });

  const card = g._buildParkingBoard()[0];

  assert.equal(card.carNo, A);
  assert.equal(card.parking, "B1");
});
