// ============================================================
// PPAP 운행기록 시스템 — Google Apps Script 백엔드
// ============================================================
//
// [데이터 흐름 원칙]
// - READ  : 차량 탭에서만 (RAW는 절대 읽지 않음)
// - WRITE : 차량 탭 + RAW 동시 저장 (RAW는 복구용 백업)
//
// ============================================================

const CONFIG = {
  SHEET_RAW: "RAW_운행일지",
  SHEET_MASTER: "차량_마스터",
  MAX_DAILY_KM: 500,
  GAP_ALERT_DAYS: 3,
  DATA_START_ROW: 15, // 차량 탭 데이터 시작 행
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
};

// 차량 탭 열 위치 (1-based)
const CAR_COL = {
  날짜: 1,    // A
  부서: 6,    // F
  성명: 10,   // J
  주행전: 14,  // N
  주행후: 20,  // T
  주행거리: 26, // Z
  출퇴근: 32,  // AF
  일반업무: 38, // AL
  비고: 44,   // AR
};

function getSpreadsheet() {
  const id =
    PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  return SpreadsheetApp.openById(id);
}

function doGet(e) {
  try {
    const props = PropertiesService.getScriptProperties();
    const config = {
      staff: JSON.parse(props.getProperty("STAFF_JSON") || "[]"),
      fixedUser: JSON.parse(props.getProperty("FIXED_USER_JSON") || "{}"),
      businessTripCars: JSON.parse(
        props.getProperty("BUSINESS_TRIP_CARS_JSON") || "[]",
      ),
      clients: JSON.parse(props.getProperty("CLIENTS_JSON") || "[]"),
    };

    const carNo = e && e.parameter && e.parameter.car ? e.parameter.car : "";
    const prevOdoData = carNo
      ? getPrevOdoData(carNo)
      : { prevOdo: null, prevDate: null, carName: "" };

    const tpl = HtmlService.createTemplateFromFile("ppap_form.html");
    tpl.configJson = JSON.stringify(config);
    tpl.carNo = carNo;
    tpl.prevOdoJson = JSON.stringify(prevOdoData);

    return tpl
      .evaluate()
      .setTitle("운행 기록")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: err.message }),
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.action === "submit") {
      const result = saveRecord(payload);
      return ContentService.createTextOutput(
        JSON.stringify(result),
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
// 없으면 -1 반환
function getLastDataRow(carSh) {
  const lastRow = carSh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return -1;

  const aCol = carSh
    .getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.DATA_START_ROW + 1, 1)
    .getValues();

  let lastDataRow = -1;
  for (let i = 0; i < aCol.length; i++) {
    if (String(aCol[i][0]).trim() !== "") {
      lastDataRow = CONFIG.DATA_START_ROW + i;
    } else {
      break;
    }
  }
  return lastDataRow;
}

// ── READ: 직전 계기판 조회 — 차량 탭 T열에서 직접 읽기 ───────────────
function getPrevOdoData(carNo) {
  const ss = getSpreadsheet();
  const masterSh = ss.getSheetByName(CONFIG.SHEET_MASTER);
  const carName = getMasterValue(masterSh, carNo, "차종");
  const carSh = ss.getSheetByName(carNo);

  if (!carSh) return { prevOdo: null, prevDate: null, carName };

  const lastDataRow = getLastDataRow(carSh);
  if (lastDataRow === -1) return { prevOdo: null, prevDate: null, carName };

  const prevOdo = carSh.getRange(lastDataRow, CAR_COL.주행후).getValue();
  const prevDate = carSh.getRange(lastDataRow, CAR_COL.날짜).getValue();

  if (!prevOdo || Number(prevOdo) === 0) return { prevOdo: null, prevDate: null, carName };

  return {
    prevOdo: Number(prevOdo),
    prevDate: String(prevDate),
    carName,
  };
}

// ── READ: 이상 감지 — 차량 탭에서 직접 읽기 ──────────────────────────
function detectAnomalies({ carSh, 주행거리,  prevOdo }) {
  const flags = [];

  if (주행거리 < 0) {
    flags.push(`역주행감지(${주행거리}km)`);
    return flags;
  }
  if (prevOdo !== null && prevOdo !== undefined) {
    // prevOdo와 주행전이 일치하는지는 프론트에서 이미 검증하므로 여기선 생략
  }

  return flags;
}

function saveRecord(payload) {
  const ss = getSpreadsheet();
  const masterSh = ss.getSheetByName(CONFIG.SHEET_MASTER);

  const now = new Date();
  const DAYS = ["일", "월", "화", "수", "목", "금", "토"];
  const 차량번호 = payload.carNo;
  const 주행후 = Number(payload.currentOdo);
  const 사용구분 = payload.useType;
  const 차종 = getMasterValue(masterSh, 차량번호, "차종") || "";
  const 사용일자 = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd");
  const 요일 = DAYS[now.getDay()];

  const isFirst = payload.prevOdo === null;

  const 주행전 = isFirst ? "" : Number(payload.prevOdo);
  const 주행거리 = isFirst ? "" : 주행후 - Number(payload.prevOdo);
  const 출퇴근 = isFirst ? "" : (사용구분 === "출퇴근용" ? 주행거리 : 0);
  const 일반업무 = isFirst ? "" : (사용구분 === "일반업무용" ? 주행거리 : 0);

  const carSh = ss.getSheetByName(차량번호);

  const flags = isFirst
    ? ["초기값등록"]
    : detectAnomalies({
        carSh,
        주행거리,
        사용일자: now,
        prevOdo: payload.prevOdo,
      });

  const id = Utilities.getUuid();
  const flagStr = flags.length > 0 ? flags.join(" | ") : "정상";
  const 타임스탬프 = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

  // ── ① WRITE: RAW 시트 백업 저장 ───────────────────────────
  const rawSh = ss.getSheetByName(CONFIG.SHEET_RAW);
  if (rawSh) {
    const newRow = new Array(16).fill("");
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
    newRow[COL.비고] = isFirst ? "" : (payload.note || "");
    newRow[COL.플래그] = flagStr;
    newRow[COL.타임스탬프] = 타임스탬프;
    rawSh.appendRow(newRow);
  }

  // ── ② WRITE: 차량 탭 저장 ─────────────────────────────────
  if (carSh) {
    const dateStr = `${now.getMonth() + 1}/${now.getDate()}(${요일})`;

    let lastDataRow = getLastDataRow(carSh);
    const insertRow = lastDataRow === -1
      ? CONFIG.DATA_START_ROW
      : lastDataRow + 1;

    if (isFirst) {
      // 최초 등록: 날짜 + T열(주행후)만 기록
      carSh.getRange(insertRow, CAR_COL.날짜).setValue(dateStr);
      carSh.getRange(insertRow, CAR_COL.주행후).setValue(주행후);
    } else {
      // 일반 기록
      carSh.getRange(insertRow, CAR_COL.날짜).setValue(dateStr);
      carSh.getRange(insertRow, CAR_COL.부서).setValue(payload.dept);
      carSh.getRange(insertRow, CAR_COL.성명).setValue(payload.name);
      carSh.getRange(insertRow, CAR_COL.주행후).setValue(주행후);
      carSh.getRange(insertRow, CAR_COL.주행전).setFormula(`=T${insertRow - 1}`);
      carSh.getRange(insertRow, CAR_COL.주행거리).setFormula(`=T${insertRow}-N${insertRow}`);
      carSh.getRange(insertRow, CAR_COL.출퇴근).setValue(출퇴근);
      carSh.getRange(insertRow, CAR_COL.일반업무).setValue(일반업무);
      carSh.getRange(insertRow, CAR_COL.비고).setValue(payload.note || "");
    }
  }

  return { success: true, id, mileage: isFirst ? 0 : 주행거리, flags };
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

    rows.forEach((r, idx) => {
      const d = new Date(r[COL.사용일자]);
      const 요일 = DAYS[d.getDay()];
      const dateStr = `${d.getMonth() + 1}/${d.getDate()}(${요일})`;
      const row = CONFIG.DATA_START_ROW + idx;
      const isFirst = String(r[COL.플래그]).includes("초기값등록");

      carSh.getRange(row, CAR_COL.날짜).setValue(dateStr);
      carSh.getRange(row, CAR_COL.주행후).setValue(r[COL.주행후]);

      if (!isFirst) {
        carSh.getRange(row, CAR_COL.부서).setValue(r[COL.부서]);
        carSh.getRange(row, CAR_COL.성명).setValue(r[COL.성명]);
        carSh.getRange(row, CAR_COL.주행전).setFormula(`=T${row - 1}`);
        carSh.getRange(row, CAR_COL.주행거리).setFormula(`=T${row}-N${row}`);
        carSh.getRange(row, CAR_COL.출퇴근).setValue(r[COL.출퇴근]);
        carSh.getRange(row, CAR_COL.일반업무).setValue(r[COL.일반업무]);
        carSh.getRange(row, CAR_COL.비고).setValue(r[COL.비고] || "");
      }
    });

    Logger.log(`${carNo}: ${rows.length}건 동기화 완료`);
  });
}

function getMasterValue(masterSh, carNo, field) {
  const FIELD_COL = { 차종: 1, 법인명: 3, 사업자번호: 4 };
  const row = masterSh
    .getDataRange()
    .getValues()
    .find((r) => r[0] === carNo);
  return row ? row[FIELD_COL[field]] : "";
}

function getMasterRow(masterSh, carNo) {
  return (
    masterSh
      .getDataRange()
      .getValues()
      .find((r) => r[0] === carNo) || null
  );
}

function setupProperties() {
  const config = JSON.parse(
    HtmlService.createHtmlOutputFromFile("config").getContent(),
  );
  const props = PropertiesService.getScriptProperties();

  props.setProperty(
    "SPREADSHEET_ID",
    "1sgzKrRD47t8429NpSOiaRJHeRCIBPf98TsqIjlGYU9A",
  );
  props.setProperty("STAFF_JSON", JSON.stringify(config.staff));
  props.setProperty("FIXED_USER_JSON", JSON.stringify(config.fixedUser));
  props.setProperty(
    "BUSINESS_TRIP_CARS_JSON",
    JSON.stringify(config.businessTripCars),
  );

  Logger.log(
    "설정 완료: 직원 " +
      config.staff.length +
      "명, 차량 " +
      (Object.keys(config.fixedUser).length + config.businessTripCars.length) +
      "대",
  );
}