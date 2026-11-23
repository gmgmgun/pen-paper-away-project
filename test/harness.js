// GAS 런타임 없이 Code.js 를 Node 에서 로드하기 위한 하네스.
//
// Code.js 는 최상위에서 GAS 서비스를 호출하지 않으므로(전부 함수 선언 + 상수),
// vm 컨텍스트에 넣고 필요한 서비스만 스텁으로 주입하면 순수 로직을 그대로 쓸 수 있다.
// Code.js 에 module.exports 를 달지 않는 이유: GAS 런타임에서 깨진다.

// GAS 스크립트 타임존은 Asia/Seoul (appsscript.json). getFormattedDate 가
// Utilities.formatDate(KST) 와 date.getDay()/getMonth()(로컬) 를 섞어 쓰므로
// 테스트 프로세스도 KST 로 맞춰야 GAS 와 같은 결과가 나온다.
process.env.TZ = "Asia/Seoul";

const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");

const CODE_PATH = path.join(__dirname, "..", "Code.js");

// ── Utilities 스텁 ────────────────────────────────────────────────────
function tzParts(date, timeZone) {
  const dtf = new Intl.DateTimeFormat("en-US", {
    timeZone,
    hourCycle: "h23",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
  const out = {};
  for (const p of dtf.formatToParts(date)) out[p.type] = p.value;
  return out;
}

function formatDate(date, timeZone, pattern) {
  const p = tzParts(date, timeZone);
  return String(pattern)
    .replaceAll("yyyy", p.year)
    .replaceAll("MM", p.month)
    .replaceAll("dd", p.day)
    .replaceAll("HH", p.hour)
    .replaceAll("mm", p.minute)
    .replaceAll("ss", p.second);
}

// "yyyy-MM-dd" 를 해당 타임존 자정으로 파싱. Code.js 가 쓰는 유일한 포맷.
function parseDate(str, timeZone, pattern) {
  if (pattern !== "yyyy-MM-dd") {
    throw new Error(`parseDate 스텁 미지원 포맷: ${pattern}`);
  }
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(str).trim());
  if (!m) throw new Error(`parseDate 스텁: 잘못된 입력 ${str}`);
  const [, y, mo, d] = m;
  // 해당 타임존의 자정에 해당하는 UTC 시각을 역산
  const guess = Date.UTC(Number(y), Number(mo) - 1, Number(d), 0, 0, 0);
  const asTz = tzParts(new Date(guess), timeZone);
  const drift =
    guess -
    Date.UTC(
      Number(asTz.year),
      Number(asTz.month) - 1,
      Number(asTz.day),
      Number(asTz.hour),
      Number(asTz.minute),
      Number(asTz.second),
    );
  return new Date(guess + drift);
}

function makeUtilities() {
  return {
    formatDate,
    parseDate,
    getUuid: () => crypto.randomUUID(),
  };
}

// ── 기본 스텁 세트 ────────────────────────────────────────────────────
// 시트 I/O 가 필요한 테스트는 overrides 로 가짜 서비스를 넘긴다.
function defaultStubs() {
  const logs = [];
  return {
    Utilities: makeUtilities(),
    Logger: {
      log: (msg) => logs.push(String(msg)),
      _logs: logs,
    },
    console,
  };
}

/**
 * Code.js 를 vm 컨텍스트에 로드하고, 그 컨텍스트(전역 객체)를 반환한다.
 * 반환값에서 Code.js 의 모든 최상위 함수/상수에 바로 접근할 수 있다.
 *
 *   const gas = loadGas();
 *   gas.computeTrip({ ... });
 */
// 최상위 const 는 vm 전역 객체의 프로퍼티가 되지 않으므로(함수 선언만 된다),
// 같은 스크립트 끝에 에필로그를 붙여 렉시컬 스코프에서 끌어낸다.
const EXPORTED_CONSTS = [
  "CONFIG",
  "COL",
  "CAR_COL",
  "CAR_TOTAL_COLS",
  "NOTICE_SHEET",
  "NOTICE_HEADER",
  "NOTICE_CACHE_KEYS",
  "EXPORT_EXCLUDE_SHEETS",
];

function loadGas(overrides = {}) {
  const src = fs.readFileSync(CODE_PATH, "utf8");
  const epilogue = `\n;this.__consts = { ${EXPORTED_CONSTS.join(", ")} };\n`;
  const sandbox = Object.assign(defaultStubs(), overrides);
  vm.createContext(sandbox);
  vm.runInContext(src + epilogue, sandbox, { filename: "Code.js" });
  Object.assign(sandbox, sandbox.__consts);
  delete sandbox.__consts;
  return sandbox;
}

// ── 가짜 스프레드시트 ─────────────────────────────────────────────────
// saveRecord / _updateRecordInner / syncAllCarSheets 가 쓰는 범위만 구현.
// grid 는 0-indexed 2차원 배열, GAS 의 행/열은 1-based 라 변환해서 접근한다.

function makeSheet(name, grid = []) {
  const g = grid.map((r) => r.slice());

  const ensure = (rows, cols) => {
    while (g.length < rows) g.push([]);
    for (const row of g) while (row.length < cols) row.push("");
  };

  const isBlank = (v) => v === "" || v === null || v === undefined;

  const sheet = {
    getName: () => name,
    _grid: g,

    getLastRow() {
      for (let i = g.length - 1; i >= 0; i--) {
        if (g[i].some((v) => !isBlank(v))) return i + 1;
      }
      return 0;
    },

    getLastColumn() {
      return g.reduce((max, r) => Math.max(max, r.length), 0);
    },

    getRange(row, col, numRows = 1, numCols = 1) {
      ensure(row + numRows - 1, col + numCols - 1);
      return {
        getValues() {
          const out = [];
          for (let i = 0; i < numRows; i++) {
            const r = [];
            for (let j = 0; j < numCols; j++) r.push(g[row - 1 + i][col - 1 + j]);
            out.push(r);
          }
          return out;
        },
        getValue() {
          return g[row - 1][col - 1];
        },
        setValues(values) {
          for (let i = 0; i < numRows; i++) {
            for (let j = 0; j < numCols; j++) {
              g[row - 1 + i][col - 1 + j] = values[i][j];
            }
          }
          return this;
        },
        clearContent() {
          for (let i = 0; i < numRows; i++) {
            for (let j = 0; j < numCols; j++) g[row - 1 + i][col - 1 + j] = "";
          }
          return this;
        },
      };
    },

    getDataRange() {
      const rows = sheet.getLastRow();
      const cols = sheet.getLastColumn();
      return sheet.getRange(1, 1, Math.max(rows, 1), Math.max(cols, 1));
    },
  };
  return sheet;
}

function makeSpreadsheet(sheetsByName) {
  const sheets = sheetsByName;
  return {
    getSheetByName: (n) => sheets[n] || null,
    getSheets: () => Object.values(sheets),
    _sheets: sheets,
  };
}

// 시트 I/O 까지 포함해 Code.js 를 로드. scriptProps 는 스크립트 속성 초기값.
function loadGasWithSheets({ sheets, scriptProps = {} }) {
  const ss = makeSpreadsheet(sheets);
  const props = Object.assign({ SPREADSHEET_ID: "fake-ss-id" }, scriptProps);
  const cache = new Map();

  const gas = loadGas({
    SpreadsheetApp: {
      openById: () => ss,
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => (k in props ? props[k] : null),
        getProperties: () => Object.assign({}, props),
        setProperty: (k, v) => {
          props[k] = v;
        },
      }),
    },
    CacheService: {
      getScriptCache: () => ({
        get: (k) => (cache.has(k) ? cache.get(k) : null),
        put: (k, v) => cache.set(k, v),
        remove: (k) => cache.delete(k),
        removeAll: (keys) => keys.forEach((k) => cache.delete(k)),
      }),
    },
    LockService: {
      getScriptLock: () => ({
        waitLock: () => {},
        tryLock: () => true,
        releaseLock: () => {},
      }),
    },
  });

  return { gas, ss, cache, props };
}

module.exports = {
  loadGas,
  loadGasWithSheets,
  makeSheet,
  makeSpreadsheet,
  formatDate,
  parseDate,
};
