// ============================================================
// PPAP 운행기록 시스템 — Google Apps Script 백엔드
// ============================================================

const CONFIG = {
  SHEET_RAW: "RAW_운행일지",
  SHEET_MASTER: "차량_마스터",
  MAX_DAILY_KM: 500,
  GAP_ALERT_DAYS: 3,
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

function getPrevOdoData(carNo) {
  const ss = getSpreadsheet();
  const rawSh = ss.getSheetByName(CONFIG.SHEET_RAW);
  const masterSh = ss.getSheetByName(CONFIG.SHEET_MASTER);
  const carName = getMasterValue(masterSh, carNo, "차종");

  const lastRow = rawSh.getLastRow();
  if (lastRow < 2) return { prevOdo: null, prevDate: null, carName };

  const data = rawSh.getRange(2, 1, lastRow - 1, 16).getValues();
  const carRows = data
    .filter((r) => r[COL.차량번호] === carNo && r[COL.주행후] > 0)
    .sort((a, b) => b[COL.주행후] - a[COL.주행후]);

  if (carRows.length === 0) return { prevOdo: null, prevDate: null, carName };

  const latest = carRows[0];
  const prevDate = Utilities.formatDate(
    new Date(latest[COL.사용일자]),
    "Asia/Seoul",
    "yyyy-MM-dd",
  );
  return { prevOdo: latest[COL.주행후], prevDate, carName };
}

function saveRecord(payload) {
  const ss = getSpreadsheet();
  const rawSh = ss.getSheetByName(CONFIG.SHEET_RAW);
  const masterSh = ss.getSheetByName(CONFIG.SHEET_MASTER);

  const now = new Date();
  const DAYS = ["일", "월", "화", "수", "목", "금", "토"];
  const 차량번호 = payload.carNo;
  const 주행후 = Number(payload.currentOdo);
  const 주행전 = payload.prevOdo !== null ? Number(payload.prevOdo) : 주행후;
  const 주행거리 = 주행후 - 주행전;
  const 사용구분 = payload.useType;
  const 출퇴근 = 사용구분 === "출퇴근용" ? 주행거리 : 0;
  const 일반업무 = 사용구분 === "일반업무용" ? 주행거리 : 0;
  const 차종 = getMasterValue(masterSh, 차량번호, "차종") || "";
  const 사용일자 = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd");
  const 요일 = DAYS[now.getDay()];

  const flags = detectAnomalies({
    rawSh,
    차량번호,
    주행전,
    주행후,
    주행거리,
    사용일자: now,
    prevOdo: payload.prevOdo,
  });

  const id = Utilities.getUuid();
  const flagStr = flags.length > 0 ? flags.join(" | ") : "정상";
  const 타임스탬프 = Utilities.formatDate(
    now,
    "Asia/Seoul",
    "yyyy-MM-dd HH:mm:ss",
  );

  // ① RAW 시트 저장
  const newRow = new Array(16).fill("");
  newRow[COL.ID] = id;
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

  // ② 차량별 탭 저장
  const carSh = ss.getSheetByName(차량번호);
  if (carSh) {
    const DATA_START_ROW = 15;
    const dateStr = `${now.getMonth() + 1}/${now.getDate()}(${요일})`;

    let lastDataRow = DATA_START_ROW - 1;
    const aCol = carSh.getRange(DATA_START_ROW, 1, 1000, 1).getValues();
    for (let i = 0; i < aCol.length; i++) {
      if (String(aCol[i][0]).trim() !== "") {
        lastDataRow = DATA_START_ROW + i;
      } else {
        break; // 빈값 만나면 즉시 중단
      }
    }

    const insertRow = lastDataRow + 1;
    carSh.getRange(insertRow, 1).setValue(dateStr);
    carSh.getRange(insertRow, 6).setValue(payload.dept);
    carSh.getRange(insertRow, 10).setValue(payload.name);

    if (insertRow === DATA_START_ROW) {
      carSh.getRange(insertRow, 14).setValue(주행전);
    } else {
      carSh.getRange(insertRow, 14).setFormula(`=T${insertRow - 1}`);
    }

    carSh.getRange(insertRow, 20).setValue(주행후);
    carSh.getRange(insertRow, 26).setFormula(`=T${insertRow}-N${insertRow}`);
    carSh.getRange(insertRow, 32).setValue(출퇴근);
    carSh.getRange(insertRow, 38).setValue(일반업무);
    carSh.getRange(insertRow, 44).setValue(payload.note || "");
  }

  return { success: true, id, mileage: 주행거리, flags };
}

function detectAnomalies({
  rawSh,
  차량번호,
  주행전,
  주행후,
  주행거리,
  사용일자,
  prevOdo,
}) {
  const flags = [];

  if (주행거리 < 0) {
    flags.push(`역주행감지(${주행거리}km)`);
    return flags;
  }
  if (주행거리 > CONFIG.MAX_DAILY_KM) flags.push(`과다주행(${주행거리}km)`);
  if (prevOdo !== null && prevOdo !== undefined) {
    const diff = Math.abs(주행전 - Number(prevOdo));
    if (diff > 0) flags.push(`계기판불일치(차이:${diff}km)`);
  }

  const lastRow = rawSh.getLastRow();
  if (lastRow >= 2) {
    const data = rawSh.getRange(2, 1, lastRow - 1, 16).getValues();
    const carRows = data
      .filter((r) => r[COL.차량번호] === 차량번호 && r[COL.주행후] > 0)
      .sort((a, b) => b[COL.주행후] - a[COL.주행후]);
    if (carRows.length > 0) {
      const dayGap = Math.floor(
        (사용일자 - new Date(carRows[0][COL.사용일자])) / (1000 * 60 * 60 * 24),
      );
      if (dayGap > CONFIG.GAP_ALERT_DAYS)
        flags.push(`${dayGap}일공백(누락의심)`);
    }
  }
  return flags;
}

function syncAllCarSheets() {
  const ss = getSpreadsheet();
  const rawSh = ss.getSheetByName(CONFIG.SHEET_RAW);
  const DAYS = ["일", "월", "화", "수", "목", "금", "토"];
  const DATA_START_ROW = 15;

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
      const row = DATA_START_ROW + idx;

      carSh.getRange(row, 1).setValue(dateStr);
      carSh.getRange(row, 6).setValue(r[COL.부서]);
      carSh.getRange(row, 10).setValue(r[COL.성명]);

      if (row === DATA_START_ROW) {
        carSh.getRange(row, 14).setValue(r[COL.주행전]);
      } else {
        carSh.getRange(row, 14).setFormula(`=T${row - 1}`);
      }

      carSh.getRange(row, 20).setValue(r[COL.주행후]);
      carSh.getRange(row, 26).setFormula(`=T${row}-N${row}`);
      carSh.getRange(row, 32).setValue(r[COL.출퇴근]);
      carSh.getRange(row, 38).setValue(r[COL.일반업무]);
      carSh.getRange(row, 44).setValue(r[COL.비고] || "");
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
