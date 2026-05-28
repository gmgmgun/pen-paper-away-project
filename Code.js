// ============================================================
// PPAP 운행기록 시스템 — Google Apps Script 백엔드
// ============================================================
//
// [데이터 흐름 원칙]
// - READ  : 차량 탭에서만 (RAW는 절대 읽지 않음)
//           단, 주차현황 조회(?mode=where)는 RAW에서 읽음
// - WRITE : 차량 탭 + RAW 동시 저장 (RAW는 복구용 백업)
//
// [변경 이력]
// - 주차위치(parking) 필드 추가: RAW 시트 17번째 열(인덱스 16)에 저장
// - ?mode=where: 주차현황 조회 페이지 추가
//
// ============================================================

const CONFIG = {
  SHEET_RAW: "RAW_운행일지",
  SHEET_MASTER: "차량_마스터",
  MAX_DAILY_KM: 500,
  GAP_ALERT_DAYS: 3,
  DATA_START_ROW: 15,
};

const COL = {
  ID: 0,
  차량번호: 1,
  차종: 2,
  사용일자: 3,
  요일: 4,
  부서: 5,
  성명: 6,
  주행전: 7,
  주행후: 8,
  주행거리: 9,
  사용구분: 10,
  출퇴근: 11,
  일반업무: 12,
  비고: 13,
  플래그: 14,
  타임스탬프: 15,
  주차위치: 16,
};

// 차량 탭 열 위치 (1-based)
const CAR_COL = {
  날짜: 1, // A
  부서: 6, // F
  성명: 10, // J
  주행전: 14, // N
  주행후: 20, // T
  주행거리: 26, // Z
  출퇴근: 32, // AF
  일반업무: 38, // AL
  비고: 44, // AR
};

const CAR_TOTAL_COLS = 44;

function getSpreadsheet() {
  const id =
    PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  return SpreadsheetApp.openById(id);
}

// ── 스크립트 속성 한 번에 가져오기 ──────────────────────────────────────
function getAllScriptProps() {
  return PropertiesService.getScriptProperties().getProperties();
}

// ── 날짜 포맷 헬퍼 ────────────────────────────────────────────────────
function getFormattedDate(date) {
  date = date || new Date();
  const DAYS = ["일", "월", "화", "수", "목", "금", "토"];
  const 사용일자 = Utilities.formatDate(date, "Asia/Seoul", "yyyy-MM-dd");
  const 요일 = DAYS[date.getDay()];
  const dateStr = `${date.getMonth() + 1}/${date.getDate()}(${요일})`;
  return { 사용일자, 요일, dateStr };
}

// ── GET 라우터 ────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const mode =
      e && e.parameter && e.parameter.mode ? e.parameter.mode : "form";

    if (mode === "where") {
      return serveParkingBoard();
    }

    return serveForm(e);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: err.message }),
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 운행기록 폼 서빙 ──────────────────────────────────────────────────
function serveForm(e) {
  const props = getAllScriptProps();
  const config = {
    staff: JSON.parse(props.STAFF_JSON || "[]"),
    fixedUser: JSON.parse(props.FIXED_USER_JSON || "{}"),
    businessTripCars: JSON.parse(props.BUSINESS_TRIP_CARS_JSON || "[]"),
    clients: JSON.parse(props.CLIENTS_JSON || "[]"),
  };

  const carNo = e && e.parameter && e.parameter.car ? e.parameter.car : "";
  const prevOdoData = carNo
    ? getPrevOdoData(carNo, props)
    : { prevOdo: null, prevDate: null, carName: "" };

  const tpl = HtmlService.createTemplateFromFile("ppap_form.html");
  tpl.configJson = JSON.stringify(config);
  tpl.carNo = carNo;
  tpl.carName = prevOdoData.carName || "";
  tpl.prevOdoJson = JSON.stringify(prevOdoData);
  tpl.noticesJson = JSON.stringify(getNotices("form"));

  return tpl
    .evaluate()
    .setTitle("운행 기록")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── READ: RAW 시트에서 차량별 최근 주차 현황 조회 ────────────────────
//
// 초기값등록 행은 운전자 정보가 없으므로 제외
// 차량번호별로 타임스탬프 기준 가장 최근 행만 추출
//
function _buildParkingBoard() {
  const ss = getSpreadsheet();
  const rawSh = ss.getSheetByName(CONFIG.SHEET_RAW);
  if (!rawSh) return [];

  const props = getAllScriptProps();
  const carMeta = JSON.parse(props.CAR_META_JSON || "{}");
  const allData = rawSh.getDataRange().getValues().slice(1); // 헤더 제외

  const latestMap = {};
  allData.forEach((r) => {
    const carNo = String(r[COL.차량번호]).trim();
    if (!carNo) return;
    if (String(r[COL.플래그]).includes("초기값등록")) return;

    const ts = r[COL.타임스탬프];
    // 문자열/Date 혼합 대비: getTime() 으로 정규화
    const tsMs = ts ? new Date(ts).getTime() : 0;
    if (!latestMap[carNo] || tsMs > latestMap[carNo].tsMs) {
      latestMap[carNo] = {
        ts,
        tsMs,
        carNo,
        parking: String(r[COL.주차위치] || "").trim(),
        name: String(r[COL.성명] || "").trim(),
        dept: String(r[COL.부서] || "").trim(),
      };
    }
  });

  // 시간 포맷: 오늘이면 "오늘 HH:mm", 어제면 "어제 HH:mm", 그 외 "MM/dd HH:mm"
  const now = new Date();
  const todayStr = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd");
  const yesterday = new Date(now);
  yesterday.setDate(yesterday.getDate() - 1);
  const yesterdayStr = Utilities.formatDate(
    yesterday,
    "Asia/Seoul",
    "yyyy-MM-dd",
  );

  return Object.values(latestMap).map((item) => {
    let timeLabel = "—";
    if (item.ts) {
      const tsDate = new Date(item.ts);
      const tsDay = Utilities.formatDate(tsDate, "Asia/Seoul", "yyyy-MM-dd");
      const tsTime = Utilities.formatDate(tsDate, "Asia/Seoul", "HH:mm");
      if (tsDay === todayStr) timeLabel = "오늘 " + tsTime;
      else if (tsDay === yesterdayStr) timeLabel = "어제 " + tsTime;
      else timeLabel = tsDay.slice(5).replace("-", "/") + " " + tsTime;
    }

    return {
      carNo: item.carNo,
      carName: carMeta[item.carNo]?.차종 || "",
      parking: item.parking,
      name: item.dept ? item.dept + " " + item.name : item.name,
      time: timeLabel,
    };
  });
}

function getParkingBoard() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("parking_board");
  if (cached) return JSON.parse(cached);

  const result = _buildParkingBoard();
  // writes 시 invalidate 되므로 TTL 길게 (최대 21600초). 30분 사용.
  cache.put("parking_board", JSON.stringify(result), 1800);
  return result;
}

// ── READ: RAW 시트 직접 조회 (캐시 우회, 테스트용) ──────────────────
function fetchParkingBoardFromRaw() {
  return _buildParkingBoard();
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.action === "submit") {
      return ContentService.createTextOutput(
        JSON.stringify(saveRecord(payload)),
      ).setMimeType(ContentService.MimeType.JSON);
    }
    if (payload.action === "update") {
      return ContentService.createTextOutput(
        JSON.stringify(updateRecord(payload)),
      ).setMimeType(ContentService.MimeType.JSON);
    }
    throw new Error("알 수 없는 action");
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, message: err.message }),
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 차량 탭에서 마지막 데이터 행 번호 반환 ────────────────────────────
// A열을 역방향으로 스캔 — 중간 공백 행이 있어도 안전.
function getLastDataRow(carSh) {
  const lastRow = carSh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return -1;

  const aCol = carSh
    .getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.DATA_START_ROW + 1, 1)
    .getValues();

  for (let i = aCol.length - 1; i >= 0; i--) {
    if (String(aCol[i][0]).trim() !== "") {
      return CONFIG.DATA_START_ROW + i;
    }
  }
  return -1;
}

// ── READ: 직전 계기판 조회 ────────────────────────────────────────────
// 차량별로 결과 캐싱(30분). saveRecord / _updateRecordInner / dailyResyncCarSheets
// 에서 해당 키를 명시적으로 invalidate 하므로 stale 위험 없음.
function _prevOdoCacheKey(carNo) {
  return `prev_odo:${carNo}`;
}

function getPrevOdoData(carNo, props) {
  const cache = CacheService.getScriptCache();
  const cacheKey = _prevOdoCacheKey(carNo);
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const ss = getSpreadsheet();
  const _props = props || getAllScriptProps();
  const carMeta = JSON.parse(_props.CAR_META_JSON || "{}");
  const carName = carMeta[carNo]?.차종 || "";

  const empty = { prevOdo: null, prevDate: null, carName };
  const carSh = ss.getSheetByName(carNo);
  if (!carSh) {
    cache.put(cacheKey, JSON.stringify(empty), 1800);
    return empty;
  }

  const lastDataRow = getLastDataRow(carSh);
  if (lastDataRow === -1) {
    cache.put(cacheKey, JSON.stringify(empty), 1800);
    return empty;
  }

  const rowData = carSh
    .getRange(lastDataRow, 1, 1, CAR_COL.주행후)
    .getValues()[0];
  const prevOdo = rowData[CAR_COL.주행후 - 1];
  const prevDate = rowData[CAR_COL.날짜 - 1];

  const result =
    !prevOdo || Number(prevOdo) === 0
      ? empty
      : { prevOdo: Number(prevOdo), prevDate: String(prevDate), carName };
  cache.put(cacheKey, JSON.stringify(result), 1800);
  return result;
}

// ── WRITE: 운행 기록 저장 ─────────────────────────────────────────────
function saveRecord(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = getSpreadsheet();
    const now = new Date();

    const 차량번호 = payload.carNo;
    const 주행후 = Number(payload.currentOdo);
    if (!Number.isFinite(주행후)) {
      throw new Error("계기판 값이 유효하지 않습니다.");
    }
    const 사용구분 = payload.useType;

    // payload.usageDate(yyyy-MM-dd) 가 오면 그 날짜로 운행기록을 작성.
    // 미래 날짜는 거부. 비어 있으면 현재 시각 사용 (이전 클라이언트 호환).
    let recordDate = now;
    if (payload.usageDate) {
      const todayStr = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd");
      if (String(payload.usageDate) > todayStr) {
        throw new Error("미래 날짜는 입력할 수 없습니다.");
      }
      recordDate = Utilities.parseDate(
        String(payload.usageDate),
        "Asia/Seoul",
        "yyyy-MM-dd",
      );
    }
    const { 사용일자, 요일, dateStr } = getFormattedDate(recordDate);
    const 주차위치 = payload.parking || "";

    const props = getAllScriptProps();
    const carMeta = JSON.parse(props.CAR_META_JSON || "{}");
    const 차종 = carMeta[차량번호]?.차종 || "";

    const isFirst = payload.prevOdo === null;
    let prevOdoNum = null;
    if (!isFirst) {
      prevOdoNum = Number(payload.prevOdo);
      if (!Number.isFinite(prevOdoNum)) {
        throw new Error("직전 계기판 값이 유효하지 않습니다.");
      }
    }
    const 주행전 = isFirst ? "" : prevOdoNum;
    const 주행거리 = isFirst ? "" : 주행후 - prevOdoNum;
    const 출퇴근 = isFirst ? "" : 사용구분 === "출퇴근용" ? 주행거리 : 0;
    const 일반업무 = isFirst ? "" : 사용구분 === "일반업무용" ? 주행거리 : 0;
    const flags = isFirst ? ["초기값등록"] : [];

    const id = Utilities.getUuid();
    const flagStr = flags.length > 0 ? flags.join(" | ") : "정상";
    const 타임스탬프 = Utilities.formatDate(
      now,
      "Asia/Seoul",
      "yyyy-MM-dd HH:mm:ss",
    );

    // ① RAW 시트 저장
    const rawSh = ss.getSheetByName(CONFIG.SHEET_RAW);
    if (rawSh) {
      const newRow = new Array(17).fill("");
      newRow[COL.ID] = id;
      newRow[COL.차량번호] = 차량번호;
      newRow[COL.차종] = 차종;
      newRow[COL.사용일자] = 사용일자;
      newRow[COL.요일] = 요일;
      newRow[COL.부서] = isFirst ? "" : payload.dept;
      newRow[COL.성명] = isFirst ? "" : payload.name;
      newRow[COL.주행전] = 주행전;
      newRow[COL.주행후] = 주행후;
      newRow[COL.주행거리] = 주행거리;
      newRow[COL.사용구분] = isFirst ? "" : 사용구분;
      newRow[COL.출퇴근] = 출퇴근;
      newRow[COL.일반업무] = 일반업무;
      newRow[COL.비고] = isFirst ? "" : payload.note || "";
      newRow[COL.플래그] = flagStr;
      newRow[COL.타임스탬프] = 타임스탬프;
      newRow[COL.주차위치] = 주차위치;
      const nextRawRow = rawSh.getLastRow() + 1;
      rawSh.getRange(nextRawRow, 1, 1, newRow.length).setValues([newRow]);
    }

    // ② 차량 탭 저장
    let carRowIndex = -1;
    const carSh = ss.getSheetByName(차량번호);
    if (carSh) {
      const lastDataRow = getLastDataRow(carSh);
      const insertRow =
        lastDataRow === -1 ? CONFIG.DATA_START_ROW : lastDataRow + 1;
      carRowIndex = insertRow;

      // 첫 행이거나, 차량탭이 비어 있는데 RAW엔 이력이 있는 복구 시나리오:
      // 직전 행 참조(=T${insertRow-1})가 헤더를 가리키지 않도록 prevOdo를
      // 리터럴로 기록.
      const isCarSheetFirstWrite = insertRow === CONFIG.DATA_START_ROW;
      if (isFirst || isCarSheetFirstWrite) {
        const row = new Array(CAR_TOTAL_COLS).fill("");
        row[CAR_COL.날짜 - 1] = dateStr;
        row[CAR_COL.주행후 - 1] = 주행후;
        if (!isFirst) {
          row[CAR_COL.부서 - 1] = payload.dept;
          row[CAR_COL.성명 - 1] = payload.name;
          row[CAR_COL.주행전 - 1] = prevOdoNum;
          row[CAR_COL.주행거리 - 1] = 주행거리;
          row[CAR_COL.출퇴근 - 1] = 출퇴근;
          row[CAR_COL.일반업무 - 1] = 일반업무;
          row[CAR_COL.비고 - 1] = payload.note || "";
        }
        carSh.getRange(insertRow, 1, 1, CAR_TOTAL_COLS).setValues([row]);
      } else {
        const row = new Array(CAR_TOTAL_COLS).fill("");
        row[CAR_COL.날짜 - 1] = dateStr;
        row[CAR_COL.부서 - 1] = payload.dept;
        row[CAR_COL.성명 - 1] = payload.name;
        row[CAR_COL.주행후 - 1] = 주행후;
        row[CAR_COL.주행전 - 1] = `=T${insertRow - 1}`;
        row[CAR_COL.주행거리 - 1] = `=T${insertRow}-N${insertRow}`;
        row[CAR_COL.출퇴근 - 1] = 출퇴근;
        row[CAR_COL.일반업무 - 1] = 일반업무;
        row[CAR_COL.비고 - 1] = payload.note || "";
        carSh.getRange(insertRow, 1, 1, CAR_TOTAL_COLS).setValues([row]);
      }
    }

    const cache = CacheService.getScriptCache();
    cache.remove("parking_board");
    cache.remove(_prevOdoCacheKey(차량번호));
    return {
      success: true,
      id,
      carRowIndex,
      mileage: isFirst ? 0 : 주행거리,
      flags,
    };
  } finally {
    lock.releaseLock();
  }
}

// ── WRITE: 기존 기록 수정 ─────────────────────────────────────────────
function updateRecord(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return _updateRecordInner(payload);
  } finally {
    lock.releaseLock();
  }
}

function _updateRecordInner(payload) {
  const ss = getSpreadsheet();
  const now = new Date();

  const 차량번호 = payload.carNo;
  const 주행후 = Number(payload.currentOdo);
  if (!Number.isFinite(주행후)) {
    throw new Error("계기판 값이 유효하지 않습니다.");
  }
  const 사용구분 = payload.useType;
  const { 사용일자, 요일 } = getFormattedDate(now);
  const carRowIndex = payload.carRowIndex;
  const 주차위치 = payload.parking || "";

  const props = getAllScriptProps();
  const carMeta = JSON.parse(props.CAR_META_JSON || "{}");
  const 차종 = carMeta[차량번호]?.차종 || "";

  const 주행전 = Number(payload.prevOdo);
  if (!Number.isFinite(주행전)) {
    throw new Error("직전 계기판 값이 유효하지 않습니다.");
  }
  const 주행거리 = 주행후 - 주행전;
  const 출퇴근 = 사용구분 === "출퇴근용" ? 주행거리 : 0;
  const 일반업무 = 사용구분 === "일반업무용" ? 주행거리 : 0;
  const flags = [];

  const newId = Utilities.getUuid();
  const flagStr = `수정됨(원본:${payload.originalId})`;
  const 타임스탬프 = Utilities.formatDate(
    now,
    "Asia/Seoul",
    "yyyy-MM-dd HH:mm:ss",
  );

  // ① RAW 시트: 수정 이력 추가
  const rawSh = ss.getSheetByName(CONFIG.SHEET_RAW);
  if (rawSh) {
    const newRow = new Array(17).fill("");
    newRow[COL.ID] = newId;
    newRow[COL.차량번호] = 차량번호;
    newRow[COL.차종] = 차종;
    newRow[COL.사용일자] = 사용일자;
    newRow[COL.요일] = 요일;
    newRow[COL.부서] = payload.dept;
    newRow[COL.성명] = payload.name;
    newRow[COL.주행전] = 주행전;
    newRow[COL.주행후] = 주행후;
    newRow[COL.주행거리] = 주행거리;
    newRow[COL.사용구분] = 사용구분;
    newRow[COL.출퇴근] = 출퇴근;
    newRow[COL.일반업무] = 일반업무;
    newRow[COL.비고] = payload.note || "";
    newRow[COL.플래그] = flagStr;
    newRow[COL.타임스탬프] = 타임스탬프;
    newRow[COL.주차위치] = 주차위치;
    const nextRawRow = rawSh.getLastRow() + 1;
    rawSh.getRange(nextRawRow, 1, 1, newRow.length).setValues([newRow]);
  }

  // ② 차량 탭: 해당 행 덮어쓰기
  const carSh = ss.getSheetByName(차량번호);
  if (carSh && carRowIndex > 0) {
    const existingDate = carSh.getRange(carRowIndex, CAR_COL.날짜).getValue();
    const row = new Array(CAR_TOTAL_COLS).fill("");
    row[CAR_COL.날짜 - 1] = existingDate;
    row[CAR_COL.부서 - 1] = payload.dept;
    row[CAR_COL.성명 - 1] = payload.name;
    row[CAR_COL.주행후 - 1] = 주행후;
    // 첫 데이터 행을 수정하는 경우 =T${row-1} 가 헤더를 가리키지 않도록 리터럴 사용
    if (carRowIndex === CONFIG.DATA_START_ROW) {
      row[CAR_COL.주행전 - 1] = 주행전;
      row[CAR_COL.주행거리 - 1] = 주행거리;
    } else {
      row[CAR_COL.주행전 - 1] = `=T${carRowIndex - 1}`;
      row[CAR_COL.주행거리 - 1] = `=T${carRowIndex}-N${carRowIndex}`;
    }
    row[CAR_COL.출퇴근 - 1] = 출퇴근;
    row[CAR_COL.일반업무 - 1] = 일반업무;
    row[CAR_COL.비고 - 1] = payload.note || "";
    carSh.getRange(carRowIndex, 1, 1, CAR_TOTAL_COLS).setValues([row]);
  }

  const cache = CacheService.getScriptCache();
  cache.remove("parking_board");
  cache.remove(_prevOdoCacheKey(차량번호));
  return { success: true, newId, mileage: 주행거리, flags };
}

// ── RAW → 차량 탭 전체 재동기화 (복구용) ─────────────────────────────
function syncAllCarSheets() {
  const ss = getSpreadsheet();
  const rawSh = ss.getSheetByName(CONFIG.SHEET_RAW);
  if (!rawSh) {
    Logger.log("RAW 시트 없음 — 동기화 불가");
    return;
  }

  const DAYS = ["일", "월", "화", "수", "목", "금", "토"];
  const rawValues = rawSh.getDataRange().getValues().slice(1);

  // 수정 이력에서 참조된 원본 id 는 무효화 — 수정본만 차량 탭에 그려지도록.
  // _updateRecordInner 가 RAW 에 새 row 추가하면서 원본은 그대로 남기는 구조라
  // 이걸 안 거르면 같은 운행이 차량 탭에 두 줄로 보임.
  const invalidIds = new Set();
  rawValues.forEach((r) => {
    const m = String(r[COL.플래그] || "").match(/원본:([\w-]+)/);
    if (m) invalidIds.add(m[1]);
  });

  const allData = rawValues
    .filter(
      (r) =>
        r[COL.차량번호] &&
        String(r[COL.차량번호]).trim() !== "" &&
        r[COL.주행후] > 0 &&
        !invalidIds.has(String(r[COL.ID])),
    )
    .sort((a, b) => new Date(a[COL.사용일자]) - new Date(b[COL.사용일자]));

  const carMap = {};
  allData.forEach((r) => {
    const car = String(r[COL.차량번호]).trim();
    if (!carMap[car]) carMap[car] = [];
    carMap[car].push(r);
  });

  // (차량+사용일자+주행후) 키로 RAW 중복 row 제거. 마지막(=최신 타임스탬프) 우선.
  // 같은 차량의 같은 날 동일 주행후 = 물리적으로 중복일 수밖에 없음.
  Object.keys(carMap).forEach((carNo) => {
    const seen = new Map();
    carMap[carNo].forEach((r) => {
      const dateKey = Utilities.formatDate(
        new Date(r[COL.사용일자]),
        "Asia/Seoul",
        "yyyy-MM-dd",
      );
      const key = `${dateKey}|${r[COL.주행후]}`;
      seen.set(key, r); // 같은 키면 뒤 row 가 덮어씀 (정렬상 더 늦은 행 우선)
    });
    carMap[carNo] = Array.from(seen.values());
  });

  Object.entries(carMap).forEach(([carNo, rows]) => {
    const carSh = ss.getSheetByName(carNo);
    if (!carSh) {
      Logger.log(`시트 없음: ${carNo}`);
      return;
    }

    const writeData = rows.map((r, idx) => {
      const d = new Date(r[COL.사용일자]);
      const 요일 = DAYS[d.getDay()];
      const dateStr = `${d.getMonth() + 1}/${d.getDate()}(${요일})`;
      const absRow = CONFIG.DATA_START_ROW + idx;
      const isFirst = String(r[COL.플래그]).includes("초기값등록");

      const row = new Array(CAR_TOTAL_COLS).fill("");
      row[CAR_COL.날짜 - 1] = dateStr;
      row[CAR_COL.주행후 - 1] = r[COL.주행후];
      if (!isFirst) {
        row[CAR_COL.부서 - 1] = r[COL.부서];
        row[CAR_COL.성명 - 1] = r[COL.성명];
        row[CAR_COL.주행전 - 1] = `=T${absRow - 1}`;
        row[CAR_COL.주행거리 - 1] = `=T${absRow}-N${absRow}`;
        row[CAR_COL.출퇴근 - 1] = r[COL.출퇴근];
        row[CAR_COL.일반업무 - 1] = r[COL.일반업무];
        row[CAR_COL.비고 - 1] = r[COL.비고] || "";
      }
      return row;
    });

    // 데이터 시작행 이하 전체 clearContent — 이전 sync 가 더 많은 row 를 그렸을
    // 경우의 잔재 0 row 제거. 원인 미상 중복 row 도 함께 정리됨.
    const lastRow = carSh.getLastRow();
    if (lastRow >= CONFIG.DATA_START_ROW) {
      carSh
        .getRange(
          CONFIG.DATA_START_ROW,
          1,
          lastRow - CONFIG.DATA_START_ROW + 1,
          CAR_TOTAL_COLS,
        )
        .clearContent();
    }

    if (writeData.length > 0) {
      carSh
        .getRange(CONFIG.DATA_START_ROW, 1, writeData.length, CAR_TOTAL_COLS)
        .setValues(writeData);
    }
    Logger.log(`${carNo}: ${rows.length}건 동기화 완료`);
  });
}

// ── Excel 내보내기 (Drive 저장 + 이메일 발송) ────────────────────────
//
// [스크립트 속성 설정 필요]
// - EXPORT_FOLDER_ID : 저장할 Drive 폴더 ID (없으면 My Drive 루트)
// - EXPORT_EMAIL     : 발송할 이메일 주소 (없으면 메일 발송 생략)
//
// [제외 시트]
// - RAW_운행일지, 차량_마스터
//
// [사용법]
// setupExportTrigger() 를 1회 수동 실행 → 이후 매월 1일 자정(KST) 자동 실행
//
const EXPORT_EXCLUDE_SHEETS = ["RAW_운행일지", "차량_마스터"];

function exportToExcel() {
  const ss = getSpreadsheet();
  const dateStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");
  const fileName = `운행일지_${dateStr}.xlsx`;

  // ① 제외 시트를 빼고 임시 스프레드시트 생성
  const tmpSs = SpreadsheetApp.create(`tmp_export_${dateStr}`);
  const defaultSheet = tmpSs.getSheets()[0]; // 자동 생성된 빈 시트 핸들

  try {
    ss.getSheets().forEach((sheet) => {
      if (EXPORT_EXCLUDE_SHEETS.includes(sheet.getName())) return;
      sheet.copyTo(tmpSs).setName(sheet.getName());
    });

    // 기본 빈 시트 삭제 (다른 시트가 하나라도 복사된 경우에만)
    if (tmpSs.getSheets().length > 1) {
      tmpSs.deleteSheet(defaultSheet);
    }

    // ② Excel blob 생성
    const exportUrl = `https://docs.google.com/spreadsheets/d/${tmpSs.getId()}/export?format=xlsx`;
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
      muteHttpExceptions: true,
    });

    if (response.getResponseCode() !== 200) {
      Logger.log("Excel 내보내기 실패: HTTP " + response.getResponseCode());
      return;
    }

    const blob = response.getBlob().setName(fileName);

    // ③ Drive 저장 (같은 날짜 파일은 덮어쓰기)
    const props = PropertiesService.getScriptProperties();
    const folderId = props.getProperty("EXPORT_FOLDER_ID");
    const folder = folderId
      ? DriveApp.getFolderById(folderId)
      : DriveApp.getRootFolder();

    const existing = folder.getFilesByName(fileName);
    if (existing.hasNext()) existing.next().setTrashed(true);

    folder.createFile(blob);
    Logger.log(`Drive 저장 완료: ${fileName} → ${folder.getName()}`);

    // ④ 이메일 발송
    const email = props.getProperty("EXPORT_EMAIL");
    if (email) {
      MailApp.sendEmail({
        to: email,
        subject: `[운행일지 백업] ${dateStr}`,
        body: `${dateStr} 운행일지 Excel 파일을 첨부합니다.`,
        attachments: [blob],
      });
      Logger.log(`이메일 발송 완료: ${email}`);
    }
  } finally {
    // 성공/실패 무관 임시 스프레드시트 정리
    try {
      DriveApp.getFileById(tmpSs.getId()).setTrashed(true);
    } catch (e) {
      Logger.log("임시 파일 삭제 실패: " + e.message);
    }
  }
}

// exportToExcel 트리거 등록 (1회 수동 실행)
function setupExportTrigger() {
  // 기존 트리거 삭제
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === "exportToExcel") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 매월 1일 자정(KST) 실행
  ScriptApp.newTrigger("exportToExcel")
    .timeBased()
    .onMonthDay(1)
    .atHour(0)
    .inTimezone("Asia/Seoul")
    .create();

  Logger.log("Excel 내보내기 트리거 등록 완료 (매월 1일 자정 KST)");
}

// ── 워밍업 트리거 (5분마다) ──────────────────────────────────────────
function warmup() {
  try {
    getSpreadsheet();
    Logger.log("warmup OK — " + new Date().toISOString());
  } catch (e) {
    Logger.log("warmup ERROR: " + e.message);
  }
}

// warmup 트리거 등록 (1회 수동 실행)
function setupWarmupTrigger() {
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === "warmup") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("warmup").timeBased().everyMinutes(5).create();

  Logger.log("warmup 트리거 등록 완료 (5분마다)");
}

// ── 일일 차량 탭 재정렬 (과거 날짜 입력으로 어긋난 순서 보정) ─────────
// 신규 입력은 append 정책이므로, 사용자가 과거 날짜로 등록하면 차량 탭의
// 행 순서가 시간순과 어긋날 수 있다. RAW를 SSOT로 두고 매일 새벽
// syncAllCarSheets 로 차량 탭을 날짜순으로 다시 작성한다.
function dailyResyncCarSheets() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log("dailyResyncCarSheets: lock 획득 실패, 건너뜀");
    return;
  }
  try {
    Logger.log("dailyResyncCarSheets 시작");
    syncAllCarSheets();

    // 차량 탭 재정렬로 prev_odo 결과가 바뀔 수 있으므로 차량 캐시 일괄 invalidate.
    const props = getAllScriptProps();
    const carMeta = JSON.parse(props.CAR_META_JSON || "{}");
    const keys = Object.keys(carMeta).map(_prevOdoCacheKey);
    if (keys.length > 0) {
      const cache = CacheService.getScriptCache();
      cache.removeAll(keys);
      cache.remove("parking_board");
    }

    Logger.log("dailyResyncCarSheets 완료");
  } catch (e) {
    Logger.log("dailyResyncCarSheets ERROR: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// dailyResyncCarSheets 트리거 등록 (1회 수동 실행)
function setupDailyResyncTrigger() {
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === "dailyResyncCarSheets") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("dailyResyncCarSheets")
    .timeBased()
    .atHour(4)
    .everyDays(1)
    .inTimezone("Asia/Seoul")
    .create();

  Logger.log("dailyResyncCarSheets 트리거 등록 완료 (매일 04:00 KST)");
}

// ── 초기 설정 (최초 1회 수동 실행) ──────────────────────────────────
function setupProperties() {
  const config = JSON.parse(
    HtmlService.createHtmlOutputFromFile("config").getContent(),
  );
  const props = PropertiesService.getScriptProperties();

  // 환경 분리: 테스트 GAS 에서는 사전에 SPREADSHEET_ID 를 세팅해두면
  // 그 값을 그대로 사용. 비어 있으면 운영 ID 로 폴백.
  const DEFAULT_PROD_ID = "1sgzKrRD47t8429NpSOiaRJHeRCIBPf98TsqIjlGYU9A";
  const ssId = props.getProperty("SPREADSHEET_ID") || DEFAULT_PROD_ID;
  props.setProperty("SPREADSHEET_ID", ssId);

  props.setProperty("STAFF_JSON", JSON.stringify(config.staff));
  props.setProperty("FIXED_USER_JSON", JSON.stringify(config.fixedUser));
  props.setProperty(
    "BUSINESS_TRIP_CARS_JSON",
    JSON.stringify(config.businessTripCars),
  );
  props.setProperty("CLIENTS_JSON", JSON.stringify(config.clients || []));

  const ss = SpreadsheetApp.openById(ssId);
  const masterSh = ss.getSheetByName(CONFIG.SHEET_MASTER);
  const masterData = masterSh.getDataRange().getValues().slice(1);
  const carMeta = {};
  masterData.forEach((r) => {
    if (r[0]) {
      carMeta[String(r[0])] = {
        차종: r[1] || "",
        법인명: r[3] || "",
        사업자번호: r[4] || "",
      };
    }
  });
  props.setProperty("CAR_META_JSON", JSON.stringify(carMeta));

  Logger.log(
    "설정 완료: 직원 " +
      config.staff.length +
      "명, 차량 " +
      (Object.keys(config.fixedUser).length + config.businessTripCars.length) +
      "대, 차량 메타 " +
      Object.keys(carMeta).length +
      "대 캐싱",
  );
}

// ── 주차현황 보드 서빙 ────────────────────────────────────────────────
function serveParkingBoard() {
  const boardData = getParkingBoard();
  const now = Utilities.formatDate(new Date(), "Asia/Seoul", "MM/dd HH:mm");

  // 현재 웹앱 배포 URL에서 mode 파라미터를 제거한 베이스 URL
  const baseUrl = ScriptApp.getService().getUrl();

  const tpl = HtmlService.createTemplateFromFile("parking_board.html");
  tpl.boardJson = JSON.stringify(boardData);
  tpl.updatedAt = now;
  tpl.baseUrl = baseUrl;
  tpl.noticesJson = JSON.stringify(getNotices("board"));

  return tpl
    .evaluate()
    .setTitle("주차 현황")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── 공지사항 ──────────────────────────────────────────────────────────
// 시트 `공지사항` (헤더: id | 제목 | 내용 | 시작일 | 종료일 | 활성 | 페이지)
// 페이지: form / board / both. 즉시 반영을 위해 onEdit 가 캐시 invalidate.
const NOTICE_SHEET = "공지사항";
const NOTICE_HEADER = [
  "id",
  "제목",
  "내용",
  "시작일",
  "종료일",
  "활성",
  "페이지",
];
const NOTICE_CACHE_KEYS = ["notices:form", "notices:board"];

// 셀 값(Date 객체 또는 문자열)을 ms 로 변환. 시작 시각용 — 시간 미입력 시 KST 자정 = 그 날 시작.
function _noticeStartMs(v) {
  if (v == null || v === "") return null;
  const d = v instanceof Date ? v : new Date(v);
  const t = d.getTime();
  return Number.isFinite(t) ? t : null;
}

// 종료 시각용. 시간 부분이 모두 0(KST 자정) 이면 "그 날 종일" 의미로 보정 — 다음 날 자정 직전까지 노출.
// 운영자가 "yyyy-MM-dd HH:mm" 처럼 시간까지 적으면 분 단위 만료, "yyyy-MM-dd" 만 적으면 종일.
function _noticeEndMs(v) {
  if (v == null || v === "") return null;
  const d = v instanceof Date ? v : new Date(v);
  const t = d.getTime();
  if (!Number.isFinite(t)) return null;
  if (d.getHours() === 0 && d.getMinutes() === 0 && d.getSeconds() === 0) {
    return t + 86400000 - 1;
  }
  return t;
}

function _buildNotices(page) {
  const ss = getSpreadsheet();
  const sh = ss.getSheetByName(NOTICE_SHEET);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const rows = sh
    .getRange(2, 1, lastRow - 1, NOTICE_HEADER.length)
    .getValues();
  const nowMs = Date.now();

  const result = [];
  rows.forEach((r) => {
    const id = String(r[0] || "").trim();
    if (!id) return;
    const title = String(r[1] || "").trim();
    const body = String(r[2] || "").trim();
    if (!title && !body) return;

    const active = r[5] === true || String(r[5]).toUpperCase() === "TRUE";
    if (!active) return;

    const startMs = _noticeStartMs(r[3]);
    const endMs = _noticeEndMs(r[4]);
    if (startMs !== null && nowMs < startMs) return;
    if (endMs !== null && nowMs > endMs) return;

    const target = String(r[6] || "both").trim().toLowerCase();
    if (target !== "both" && target !== page) return;

    result.push({ id, title, body });
  });
  return result;
}

function getNotices(page) {
  const cache = CacheService.getScriptCache();
  const key = `notices:${page}`;
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);

  const result = _buildNotices(page);
  // TTL 10분. 사용자 편집은 onEdit 가 즉시 invalidate.
  cache.put(key, JSON.stringify(result), 600);
  return result;
}

function invalidateNoticesCache() {
  CacheService.getScriptCache().removeAll(NOTICE_CACHE_KEYS);
}

// 공지사항 시트 편집 시 캐시 즉시 비움. simple onEdit 트리거(자동 동작).
// throw 하면 사용자 편집 UI에 에러 토스트가 뜨므로 조용히 무시.
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    if (e.range.getSheet().getName() !== NOTICE_SHEET) return;
    CacheService.getScriptCache().removeAll(NOTICE_CACHE_KEYS);
  } catch (_) {}
}

// 공지사항 시트 초기화 (1회 수동 실행). 시트 없으면 생성 + 헤더/검증/체크박스 세팅.
function setupNoticeSheet() {
  const ss = getSpreadsheet();
  let sh = ss.getSheetByName(NOTICE_SHEET);
  if (!sh) sh = ss.insertSheet(NOTICE_SHEET);

  sh.getRange(1, 1, 1, NOTICE_HEADER.length)
    .setValues([NOTICE_HEADER])
    .setFontWeight("bold")
    .setBackground("#eff6ff");
  sh.setFrozenRows(1);

  // 활성 컬럼(F) 체크박스
  sh.getRange("F2:F").insertCheckboxes();

  // 페이지 컬럼(G) 검증
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["form", "board", "both"], true)
    .build();
  sh.getRange("G2:G").setDataValidation(rule);

  sh.setColumnWidth(1, 60);
  sh.setColumnWidth(2, 220);
  sh.setColumnWidth(3, 420);
  sh.setColumnWidth(4, 110);
  sh.setColumnWidth(5, 110);
  sh.setColumnWidth(6, 70);
  sh.setColumnWidth(7, 90);

  invalidateNoticesCache();
  Logger.log("공지사항 시트 초기화 완료");
}
