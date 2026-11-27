const { test } = require("node:test");
const assert = require("node:assert/strict");
const { loadGasWithSheets, makeSheet } = require("./harness");

const MASTER = "차량_마스터";
const STAFF = "직원";
const CLIENTS = "거래처";
const A = "12가3456";
const B = "34나5678";
const C = "56다7890";
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

const plain = (x) => JSON.parse(JSON.stringify(x));

function row(carNo, 전담자 = "", 전담자부서 = "", 출장 = "") {
  return [carNo, "K5", "", "(주)샘플", "123-45-67890", "", "", 전담자, 전담자부서, 출장];
}

// grids: 시트 이름 → 2차원 배열. null 이면 그 시트를 만들지 않는다.
function settingsFixture({
  staffGrid = [["이름", "부서"], ["홍길동", "영업팀"]],
  clientsGrid = [["거래처"], ["A상사"]],
  masterGrid = [MASTER_HEADER, row(A)],
} = {}) {
  const sheets = {};
  if (staffGrid) sheets[STAFF] = makeSheet(STAFF, staffGrid);
  if (clientsGrid) sheets[CLIENTS] = makeSheet(CLIENTS, clientsGrid);
  if (masterGrid) sheets[MASTER] = makeSheet(MASTER, masterGrid);
  return loadGasWithSheets({ sheets });
}

// ── 직원 ──

test("_readStaff: 이름 빈 행은 건너뛰고 trim, 시트 순서를 지킨다", () => {
  const { gas, ss } = settingsFixture({
    staffGrid: [
      ["이름", "부서"],
      [" 홍길동 ", "영업팀 "],
      ["", ""],
      ["김철수", ""],
    ],
  });

  assert.deepEqual(plain(gas._readStaff(ss).staff), [
    { name: "홍길동", dept: "영업팀" },
    { name: "김철수", dept: "" },
  ]);
});

test("_readStaff: 직원 시트가 없으면 migrateConfigToSheets 를 안내하며 throw", () => {
  const { gas, ss } = settingsFixture({ staffGrid: null });

  assert.throws(() => gas._readStaff(ss), /직원 시트가 없습니다[\s\S]*migrateConfigToSheets/);
});

test("_readStaff: 1행이 이름|부서 가 아니면 throw", () => {
  const { gas, ss } = settingsFixture({
    staffGrid: [["성명", "부서"], ["홍길동", "영업팀"]],
  });

  assert.throws(() => gas._readStaff(ss), /직원 시트 1행/);
});

test("_readStaff: 이름이 하나도 없으면 throw", () => {
  const { gas, ss } = settingsFixture({ staffGrid: [["이름", "부서"]] });

  assert.throws(() => gas._readStaff(ss), /직원 시트에 이름이 없습니다/);
});

test("_readStaff: 같은 이름이 있으면 경고 한 번", () => {
  const { gas, ss } = settingsFixture({
    staffGrid: [
      ["이름", "부서"],
      ["홍길동", "영업팀"],
      ["홍길동", "관리팀"],
      ["홍길동", "총무팀"],
      ["김철수", "영업팀"],
    ],
  });

  const { warnings } = gas._readStaff(ss);

  assert.equal(warnings.length, 1);
  assert.ok(warnings[0].includes("홍길동"));
});

// ── 거래처 ──

test("_readClients: 빈 행을 건너뛰고 순서를 지킨다", () => {
  const { gas, ss } = settingsFixture({
    clientsGrid: [["거래처"], ["A상사"], [""], [" B물산 "]],
  });

  assert.deepEqual([...gas._readClients(ss)], ["A상사", "B물산"]);
});

test("_readClients: 헤더만 있으면 빈 배열", () => {
  const { gas, ss } = settingsFixture({ clientsGrid: [["거래처"]] });

  assert.deepEqual([...gas._readClients(ss)], []);
});

test("_readClients: 거래처 시트가 없으면 migrateConfigToSheets 를 안내하며 throw", () => {
  const { gas, ss } = settingsFixture({ clientsGrid: null });

  assert.throws(() => gas._readClients(ss), /거래처 시트가 없습니다[\s\S]*migrateConfigToSheets/);
});

test("_readClients: 1행이 거래처 가 아니면 throw", () => {
  const { gas, ss } = settingsFixture({ clientsGrid: [["업체"], ["A상사"]] });

  assert.throws(() => gas._readClients(ss), /거래처 시트 1행/);
});

// ── 차량_마스터 H~J ──

test("_readCarMaster: H~J 에서 전담·출장을 읽는다", () => {
  const { gas, ss } = settingsFixture({
    masterGrid: [
      MASTER_HEADER,
      row(A, "김전담", "임원"),
      row(B, "", "", true),
      row(C, "", "", "TRUE"),
    ],
  });

  const { fixedUser, businessTripCars } = gas._readCarMaster(ss.getSheetByName(MASTER));

  assert.deepEqual(plain(fixedUser), { [A]: { name: "김전담", dept: "임원" } });
  assert.deepEqual([...businessTripCars], [B, C]);
});

test("_readCarMaster: 전담자만 있고 부서가 비면 dept 는 빈 문자열", () => {
  const { gas, ss } = settingsFixture({
    masterGrid: [MASTER_HEADER, row(A, "김전담", "")],
  });

  const { fixedUser } = gas._readCarMaster(ss.getSheetByName(MASTER));

  assert.deepEqual(plain(fixedUser), { [A]: { name: "김전담", dept: "" } });
});

test("_readCarMaster: 전담자부서만 있으면 전담 차량이 아니다", () => {
  const { gas, ss } = settingsFixture({
    masterGrid: [MASTER_HEADER, row(A, "", "임원")],
  });

  const { fixedUser } = gas._readCarMaster(ss.getSheetByName(MASTER));

  assert.deepEqual(plain(fixedUser), {});
});

test("_readCarMaster: H1:J1 이 비어 있으면 migrateConfigToSheets 를 안내하며 throw", () => {
  const { gas, ss } = settingsFixture({
    masterGrid: [MASTER_HEADER.slice(0, 7), row(A)],
  });

  assert.throws(
    () => gas._readCarMaster(ss.getSheetByName(MASTER)),
    /H1:J1[\s\S]*migrateConfigToSheets/,
  );
});

test("_readCarMaster: H1:J1 이 다른 값이면 throw", () => {
  const header = MASTER_HEADER.slice();
  header[7] = "전담";
  const { gas, ss } = settingsFixture({ masterGrid: [header, row(A)] });

  assert.throws(() => gas._readCarMaster(ss.getSheetByName(MASTER)), /H1:J1/);
});
