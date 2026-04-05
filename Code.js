// ============================================================
// PPAP 운행기록 시스템 — Google Apps Script 백엔드
// ============================================================
//
// [데이터 흐름 원칙]
// - READ  : 차량 탭에서만 (RAW는 절대 읽지 않음)
// - WRITE : 차량 탭 + RAW 동시 저장 (RAW는 복구용 백업)
//
// [성능 최적화 내역]
// 1. saveRecord() — setValue() 다중 호출 → setValues() 1회로 통합
// 2. 차종 조회 — getMasterValue() 매번 시트 전체 읽기 → Script Properties 캐시로 대체
// 3. getLastDataRow() 중복 호출 제거 — saveRecord() 안에서 1회만 호출
// 4. SpreadsheetApp.flush() 명시적 호출 추가
//
// [버그 수정]
// - updateRecord 함수 누락 복구
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

// 차량 탭 마지막 열 (비고 열 = 44)
const CAR_TOTAL_COLS = 44;

function getSpreadsheet() {
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  return SpreadsheetApp.openById(id);
}

function doGet(e) {
  try {
    const props = PropertiesService.getScriptProperties();
    const config = {
      staff: JSON.parse(props.getProperty("STAFF_JSON") || "[]"),
      fixedUser: JSON.parse(props.getProperty("FIXED_USER_JSON") || "{}"),
      businessTripCars: JSON.parse(props.getProperty("BUSINESS_TRIP_CARS_JSON") || "[]"),
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
      JSON.stringify({ error: err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.action === "submit") {
      const result = saveRecord(payload);
      return ContentService.createTextOutput(
        JSON.stringify(result)
      ).setMimeType(ContentService.MimeType.JSON);
    }
    if (payload.action === "update") {
      const result = updateRecord(payload);
      return ContentService.createTextOutput(
        JSON.stringify(result)
      ).setMimeType(ContentService.MimeType.JSON);
    }
    throw new Error("알 수 없는 action");
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, message: err.message })
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
  // [최적화 2] 마스터 시트 읽기 대신 캐시된 Script Properties 사용
  const carMeta = JSON.parse(
    PropertiesService.getScriptProperties().getProperty("CAR_META_JSON") || "{}"
  );
  const carName = carMeta[carNo]?.차종 || "";

  const carSh = ss.getSheetByName(carNo);
  if (!carSh) return { prevOdo: null, prevDate: null, carName };

  const lastDataRow = getLastDataRow(carSh);
  if (lastDataRow === -1) return { prevOdo: null, prevDate: null, carName };

  // [최적화 3] 날짜(A)와 주행후(T) 두 셀을 getValues() 1회로 읽기
  // A열(1)~T열(20) 범위를 한 번에 읽어 인덱스로 접근
  const rowData = carSh.getRange(lastDataRow, 1, 1, CAR_COL.주행후).getValues()[0];
  const prevOdo = rowData[CAR_COL.주행후 - 1];
  const prevDate = rowData[CAR_COL.날짜 - 1];

  if (!prevOdo || Number(prevOdo) === 0) return { prevOdo: null, prevDate: null, carName };

  return {
    prevOdo: Number(prevOdo),
    prevDate: String(prevDate),
    carName,
  };
}

// ── READ: 이상 감지 ───────────────────────────────────────────────────
function detectAnomalies({ 주행거리 }) {
  const flags = [];
  if (주행거리 < 0) {
    flags.push(`역주행감지(${주행거리}km)`);
  }
  return flags;
}

// ── WRITE: 운행 기록 저장 ─────────────────────────────────────────────
function saveRecord(payload) {
  const ss = getSpreadsheet();

  const now = new Date();
  const DAYS = ["일", "월", "화", "수", "목", "금", "토"];
  const 차량번호 = payload.carNo;
  const 주행후 = Number(payload.currentOdo);
  const 사용구분 = payload.useType;
  const 사용일자 = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd");
  const 요일 = DAYS[now.getDay()];
  const dateStr = `${now.getMonth() + 1}/${now.getDate()}(${요일})`;

  // [최적화 2] 마스터 시트 전체 읽기 대신 캐시된 Script Properties 사용
  const carMeta = JSON.parse(
    PropertiesService.getScriptProperties().getProperty("CAR_META_JSON") || "{}"
  );
  const 차종 = carMeta[차량번호]?.차종 || "";

  const isFirst = payload.prevOdo === null;
  const 주행전 = isFirst ? "" : Number(payload.prevOdo);
  const 주행거리 = isFirst ? "" : 주행후 - Number(payload.prevOdo);
  const 출퇴근 = isFirst ? "" : (사용구분 === "출퇴근용" ? 주행거리 : 0);
  const 일반업무 = isFirst ? "" : (사용구분 === "일반업무용" ? 주행거리 : 0);

  const flags = isFirst
    ? ["초기값등록"]
    : detectAnomalies({ 주행거리 });

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
  const carSh = ss.getSheetByName(차량번호);
  if (carSh) {
    // [최적화 3] getLastDataRow는 여기서 1회만 호출 (getPrevOdoData와 중복 제거)
    const lastDataRow = getLastDataRow(carSh);
    const insertRow = lastDataRow === -1 ? CONFIG.DATA_START_ROW : lastDataRow + 1;

    if (isFirst) {
      // [최적화 1] 최초 등록: 날짜(A) + 주행후(T) — setValues() 2셀 배치
      // 두 셀이 연속하지 않아 개별 호출이 불가피하나, flush 전 버퍼에 묶임
      carSh.getRange(insertRow, CAR_COL.날짜).setValue(dateStr);
      carSh.getRange(insertRow, CAR_COL.주행후).setValue(주행후);
    } else {
      // [최적화 1] 일반 기록: CAR_TOTAL_COLS(44)열짜리 행 배열을 setValues() 1회로 저장
      // 열 번호가 띄엄띄엄(1,6,10,14,20,26,32,38,44)이므로
      // 빈 배열을 만들고 각 위치에 값/수식을 주입한 뒤 한 번에 씀
      const row = new Array(CAR_TOTAL_COLS).fill("");
      row[CAR_COL.날짜    - 1] = dateStr;
      row[CAR_COL.부서    - 1] = payload.dept;
      row[CAR_COL.성명    - 1] = payload.name;
      row[CAR_COL.주행후  - 1] = 주행후;
      // 수식: 주행전 = 바로 윗 행의 T열, 주행거리 = T - N
      row[CAR_COL.주행전  - 1] = `=T${insertRow - 1}`;
      row[CAR_COL.주행거리- 1] = `=T${insertRow}-N${insertRow}`;
      row[CAR_COL.출퇴근  - 1] = 출퇴근;
      row[CAR_COL.일반업무- 1] = 일반업무;
      row[CAR_COL.비고    - 1] = payload.note || "";

      carSh.getRange(insertRow, 1, 1, CAR_TOTAL_COLS).setValues([row]);
    }
  }

  // [최적화 4] 모든 쓰기가 끝난 후 명시적 flush — 응답 시간 안정화
  SpreadsheetApp.flush();

  return { success: true, id, mileage: isFirst ? 0 : 주행거리, flags };
}

// ── WRITE: 기존 기록 수정 ─────────────────────────────────────────────
//
// 처리 순서:
// 1. RAW 시트에 "수정" 플래그가 붙은 새 행 추가 (원본 ID 이력 보존)
// 2. 차량 탭의 해당 행(carRowIndex)을 새 데이터로 덮어씀
//    - 수식(주행전, 주행거리)은 행 번호 기준으로 재설정
//
function updateRecord(payload) {
  const ss = getSpreadsheet();

  const now = new Date();
  const DAYS = ["일", "월", "화", "수", "목", "금", "토"];
  const 차량번호 = payload.carNo;
  const 주행후 = Number(payload.currentOdo);
  const 사용구분 = payload.useType;
  const 요일 = DAYS[now.getDay()];
  const carRowIndex = payload.carRowIndex;

  // [최적화 2] 캐시에서 차종 조회
  const carMeta = JSON.parse(
    PropertiesService.getScriptProperties().getProperty("CAR_META_JSON") || "{}"
  );
  const 차종 = carMeta[차량번호]?.차종 || "";
  const 사용일자 = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd");

  const 주행전 = Number(payload.prevOdo);
  const 주행거리 = 주행후 - 주행전;
  const 출퇴근 = 사용구분 === "출퇴근용" ? 주행거리 : 0;
  const 일반업무 = 사용구분 === "일반업무용" ? 주행거리 : 0;

  const carSh = ss.getSheetByName(차량번호);
  const flags = detectAnomalies({ 주행거리 });

  const newId = Utilities.getUuid();
  const flagStr = (flags.length > 0 ? flags.join(" | ") + " | " : "") +
                  `수정됨(원본:${payload.originalId})`;
  const 타임스탬프 = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

  // ── ① RAW 시트: 수정 이력 행 추가 ────────────────────────
  const rawSh = ss.getSheetByName(CONFIG.SHEET_RAW);
  if (rawSh) {
    const newRow = new Array(16).fill("");
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
    rawSh.appendRow(newRow);
  }

  // ── ② 차량 탭: 해당 행 덮어쓰기 — setValues() 1회로 배치 처리 ──
  if (carSh && carRowIndex > 0) {
    // 날짜는 원본 유지 → 날짜 열은 건드리지 않고 나머지만 배열로 구성
    // 기존 행을 먼저 읽어 날짜값 보존
    const existingDate = carSh.getRange(carRowIndex, CAR_COL.날짜).getValue();

    const row = new Array(CAR_TOTAL_COLS).fill("");
    row[CAR_COL.날짜    - 1] = existingDate;
    row[CAR_COL.부서    - 1] = payload.dept;
    row[CAR_COL.성명    - 1] = payload.name;
    row[CAR_COL.주행후  - 1] = 주행후;
    row[CAR_COL.주행전  - 1] = `=T${carRowIndex - 1}`;
    row[CAR_COL.주행거리- 1] = `=T${carRowIndex}-N${carRowIndex}`;
    row[CAR_COL.출퇴근  - 1] = 출퇴근;
    row[CAR_COL.일반업무- 1] = 일반업무;
    row[CAR_COL.비고    - 1] = payload.note || "";

    carSh.getRange(carRowIndex, 1, 1, CAR_TOTAL_COLS).setValues([row]);
  }

  // [최적화 4] flush
  SpreadsheetApp.flush();

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
        r[COL.주행후] > 0
    )
    .sort((a, b) => new Date(a[COL.사용일자]) - new Date(b[COL.사용일자]));

  const carMap = {};
  allData.forEach((r) => {
    const car = String(r[COL.차량번호]).trim();
    if (!carMap[car]) carMap[car] = [];
    carMap[car].push(r);
  });

  // [최적화 1 — syncAllCarSheets] 차량별로 행 배열을 모아 setValues() 1회로 일괄 저장
  Object.entries(carMap).forEach(([carNo, rows]) => {
    const carSh = ss.getSheetByName(carNo);
    if (!carSh) {
      Logger.log(`시트 없음: ${carNo}`);
      return;
    }

    // 전체 데이터를 2D 배열로 구성
    const writeData = rows.map((r, idx) => {
      const d = new Date(r[COL.사용일자]);
      const 요일 = DAYS[d.getDay()];
      const dateStr = `${d.getMonth() + 1}/${d.getDate()}(${요일})`;
      const absRow = CONFIG.DATA_START_ROW + idx;
      const isFirst = String(r[COL.플래그]).includes("초기값등록");

      const row = new Array(CAR_TOTAL_COLS).fill("");
      row[CAR_COL.날짜   - 1] = dateStr;
      row[CAR_COL.주행후 - 1] = r[COL.주행후];

      if (!isFirst) {
        row[CAR_COL.부서    - 1] = r[COL.부서];
        row[CAR_COL.성명    - 1] = r[COL.성명];
        row[CAR_COL.주행전  - 1] = `=T${absRow - 1}`;
        row[CAR_COL.주행거리- 1] = `=T${absRow}-N${absRow}`;
        row[CAR_COL.출퇴근  - 1] = r[COL.출퇴근];
        row[CAR_COL.일반업무- 1] = r[COL.일반업무];
        row[CAR_COL.비고    - 1] = r[COL.비고] || "";
      }
      return row;
    });

    // 한 번의 setValues()로 전체 차량 데이터 저장
    carSh
      .getRange(CONFIG.DATA_START_ROW, 1, writeData.length, CAR_TOTAL_COLS)
      .setValues(writeData);

    Logger.log(`${carNo}: ${rows.length}건 동기화 완료`);
  });

  // [최적화 4] 복구 완료 후 flush
  SpreadsheetApp.flush();
}

// ── 마스터 시트 헬퍼 (setupProperties 전용 — 런타임 호출 금지) ──────────
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

// ── 워밍업 트리거 (5분마다 자동 실행) ───────────────────────────────────
//
// 설정 방법 (최초 1회 수동):
// GAS 편집기 → 트리거 → 추가 → warmup / 시간 기반 / 5분마다
//
// 역할: GAS 인스턴스를 웜 상태로 유지해 콜드 스타트(2~5초 지연)를 제거
// 비용: 스프레드시트를 열기만 하고 아무것도 쓰지 않으므로 할당량 소모 최소
//
function warmup() {
  try {
    getSpreadsheet(); // 인스턴스 초기화만 수행
    Logger.log("warmup OK — " + new Date().toISOString());
  } catch (e) {
    Logger.log("warmup ERROR: " + e.message);
  }
}

// ── 초기 설정 (최초 1회 수동 실행) ─────────────────────────────────────
function setupProperties() {
  const config = JSON.parse(
    HtmlService.createHtmlOutputFromFile("config").getContent()
  );
  const props = PropertiesService.getScriptProperties();

  props.setProperty("SPREADSHEET_ID", "1sgzKrRD47t8429NpSOiaRJHeRCIBPf98TsqIjlGYU9A");
  props.setProperty("STAFF_JSON", JSON.stringify(config.staff));
  props.setProperty("FIXED_USER_JSON", JSON.stringify(config.fixedUser));
  props.setProperty("BUSINESS_TRIP_CARS_JSON", JSON.stringify(config.businessTripCars));

  // [최적화 2] 차량-차종 매핑을 Script Properties에 캐싱
  // → 런타임에서 getMasterValue() 대신 이 캐시를 사용
  const ss = SpreadsheetApp.openById("1sgzKrRD47t8429NpSOiaRJHeRCIBPf98TsqIjlGYU9A");
  const masterSh = ss.getSheetByName(CONFIG.SHEET_MASTER);
  const masterData = masterSh.getDataRange().getValues().slice(1); // 헤더 제외
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
      "대 캐싱"
  );
}