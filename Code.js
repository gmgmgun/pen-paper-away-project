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
// 정상 케이스(=마지막 행이 그대로 데이터 행)는 1셀 read 로 끝낸다.
// 중간 공백 행이 있을 가능성에 대비해 비어 있으면 A열 전체 역스캔으로 폴백.
function getLastDataRow(carSh) {
  const lastRow = carSh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return -1;

  const lastCellA = carSh.getRange(lastRow, 1).getValue();
  if (String(lastCellA).trim() !== "") return lastRow;

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
  const allData = rawSh
    .getDataRange()
    .getValues()
    .slice(1)
    .filter(
      (r) =>
        r[COL.차량번호] &&
        String(r[COL.차량번호]).trim() !== "" &&
        r[COL.주행후] > 0,
    )
    .sort((a, b) => new Date(a[COL.사용일자]) - new Date(b[COL.사용일자]));

  const carMap = {};
  allData.forEach((r) => {
    const car = String(r[COL.차량번호]).trim();
    if (!carMap[car]) carMap[car] = [];
    carMap[car].push(r);
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

    carSh
      .getRange(CONFIG.DATA_START_ROW, 1, writeData.length, CAR_TOTAL_COLS)
      .setValues(writeData);
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

  return tpl
    .evaluate()
    .setTitle("주차 현황")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
