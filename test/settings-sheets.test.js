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

// ── migrateConfigToSheets: 기존 props → 설정 시트 (환경별 1회) ──

const GHOST = "99하9999"; // props 에만 있고 마스터에 없는 차량
const MASTER7 = MASTER_HEADER.slice(0, 7);
const row7 = (carNo) => row(carNo).slice(0, 7);

const DEFAULT_PROPS = {
  STAFF_JSON: JSON.stringify([
    { name: "홍길동", dept: "영업팀" },
    { name: "김철수", dept: "관리팀" },
  ]),
  CLIENTS_JSON: JSON.stringify(["A상사", "B물산"]),
  FIXED_USER_JSON: JSON.stringify({ [A]: { name: "김전담", dept: "임원" } }),
  BUSINESS_TRIP_CARS_JSON: JSON.stringify([B]),
};

// 이관 전 상태: 마스터 A~G(7열), 차량 A·B 탭, 직원·거래처 시트 없음.
function migrateFixture({
  props = DEFAULT_PROPS,
  masterGrid = [MASTER7, row7(A), row7(B)],
  maxColumns,
  extraSheets = {},
} = {}) {
  const sheets = {
    [MASTER]: makeSheet(MASTER, masterGrid, { maxColumns }),
    [A]: makeSheet(A, [["운행기록부"]]),
    [B]: makeSheet(B, [["운행기록부"]]),
  };
  Object.entries(extraSheets).forEach(([name, grid]) => {
    sheets[name] = makeSheet(name, grid);
  });
  return loadGasWithSheets({ sheets, scriptProps: props });
}

test("이관: props 의 직원·거래처를 새 시트로 옮기고 props 는 건드리지 않는다", () => {
  const { gas, ss, props } = migrateFixture();

  gas.migrateConfigToSheets();

  assert.deepEqual(plain(ss.getSheetByName(STAFF)._grid), [
    ["이름", "부서"],
    ["홍길동", "영업팀"],
    ["김철수", "관리팀"],
  ]);
  assert.deepEqual(plain(ss.getSheetByName(CLIENTS)._grid), [
    ["거래처"],
    ["A상사"],
    ["B물산"],
  ]);
  assert.equal(props.STAFF_JSON, DEFAULT_PROPS.STAFF_JSON);
  assert.equal(props.FIXED_USER_JSON, DEFAULT_PROPS.FIXED_USER_JSON);
});

test("이관: 전담·출장을 차량_마스터 H~J 로 옮긴다", () => {
  const { gas, ss } = migrateFixture();

  gas.migrateConfigToSheets();

  const master = ss.getSheetByName(MASTER);
  assert.deepEqual(plain(master.getRange(1, 8, 3, 3).getValues()), [
    ["전담자", "전담자부서", "출장차량"],
    ["김전담", "임원", false],
    ["", "", true],
  ]);
  assert.deepEqual(plain(master._checkboxRanges), [
    { row: 2, col: 10, numRows: 2, numCols: 1 },
  ]);
});

test("이관: 마스터에 없는 전담·출장 차량은 건너뛰고 경고", () => {
  const { gas } = migrateFixture({
    props: {
      ...DEFAULT_PROPS,
      FIXED_USER_JSON: JSON.stringify({
        [A]: { name: "김전담", dept: "임원" },
        [GHOST]: { name: "박유령", dept: "" },
      }),
      BUSINESS_TRIP_CARS_JSON: JSON.stringify([B, GHOST]),
    },
  });

  const res = gas.migrateConfigToSheets();

  assert.deepEqual([...res.skipped], [GHOST]);
  assert.ok(
    gas.Logger._logs.some((l) => l.includes("경고:") && l.includes(GHOST)),
  );
});

test("이관: 직원 시트가 이미 있으면 아무것도 바꾸기 전에 throw", () => {
  const { gas, ss } = migrateFixture({
    extraSheets: { [STAFF]: [["이름", "부서"]] },
  });

  assert.throws(() => gas.migrateConfigToSheets(), /이미 있습니다/);
  assert.equal(ss.getSheetByName(MASTER).getRange(1, 8).getValue(), "");
  assert.equal(ss.getSheetByName(CLIENTS), null);
});

test("이관: 차량_마스터 H~J 에 값이 있으면 throw", () => {
  const { gas, ss } = migrateFixture({
    masterGrid: [MASTER7, row7(A), [...row7(B), "메모"]],
  });

  assert.throws(() => gas.migrateConfigToSheets(), /H3/);
  assert.equal(ss.getSheetByName(STAFF), null);
});

test("이관: STAFF_JSON 이 없거나 비어 있으면 throw", () => {
  for (const staff of [null, "[]"]) {
    const { gas, ss } = migrateFixture({
      props: { ...DEFAULT_PROPS, STAFF_JSON: staff },
    });

    assert.throws(() => gas.migrateConfigToSheets(), /STAFF_JSON/, String(staff));
    assert.equal(ss.getSheetByName(STAFF), null);
  }
});

test("이관: 마스터 열이 10개보다 적으면 늘린다", () => {
  const { gas, ss } = migrateFixture({ maxColumns: 7 });

  gas.migrateConfigToSheets();

  assert.ok(ss.getSheetByName(MASTER).getMaxColumns() >= 10);
});

test("이관: 전담·출장·거래처 prop 이 없는 환경은 직원만 옮긴다", () => {
  const { gas, ss } = migrateFixture({
    props: { STAFF_JSON: DEFAULT_PROPS.STAFF_JSON },
  });

  const res = gas.migrateConfigToSheets();

  assert.deepEqual(plain(ss.getSheetByName(CLIENTS)._grid), [["거래처"]]);
  assert.deepEqual(plain(ss.getSheetByName(MASTER).getRange(2, 8, 2, 3).getValues()), [
    ["", "", false],
    ["", "", false],
  ]);
  assert.deepEqual([...res.skipped], []);
});

test("이관: 차량이 0대면 H1:J1 만 쓰고 체크박스를 넣지 않는다", () => {
  const { gas, ss } = migrateFixture({ masterGrid: [MASTER7] });

  const res = gas.migrateConfigToSheets();

  const master = ss.getSheetByName(MASTER);
  assert.deepEqual(plain(master.getRange(1, 8, 1, 3).getValues()), [
    ["전담자", "전담자부서", "출장차량"],
  ]);
  assert.equal(master._checkboxRanges.length, 0);
  assert.deepEqual([...res.skipped], [A, B]);
});

test("이관 뒤 syncSettings 는 이관 전과 같은 props 를 만든다", () => {
  const { gas, props } = migrateFixture();

  gas.migrateConfigToSheets();
  gas.syncSettings();

  for (const key of Object.keys(DEFAULT_PROPS)) {
    assert.deepEqual(JSON.parse(props[key]), JSON.parse(DEFAULT_PROPS[key]), key);
  }
});
