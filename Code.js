// ============================================================
// PPAP 운행기록 시스템 — Google Apps Script 백엔드
// ============================================================

const CONFIG = {
  SHEET_RAW: "RAW_운행일지",
  SHEET_MASTER: "차량_마스터",
  ADMIN_EMAIL: "dmlee@greenia.co.kr",
  MAX_DAILY_KM: 500,
  GAP_ALERT_DAYS: 3,
};

const COL = {
  ID: 0, // A
  차량번호: 1, // B
  차종: 2, // C
  사용일자: 3, // D
  요일: 4, // E
  부서: 5, // F
  성명: 6, // G
  주행전: 7, // H
  주행후: 8, // I
  주행거리: 9, // J
  사용구분: 10, // K
  출퇴근: 11, // L
  일반업무: 12, // M
  비고: 13, // N
  플래그: 14, // O
  타임스탬프: 15, // P
};

// ── GET: HTML 서빙 + JSON API ──────────────────────────────
function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : null;

    // JSON API — 이전 계기판 조회
    if (action === "getPrevOdo") {
      const carNo = e.parameter.car;
      if (!carNo) throw new Error("차량번호 없음");
      const result = getPrevOdoData(carNo);
      return ContentService.createTextOutput(
        JSON.stringify(result),
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // HTML 서빙 — 설정값을 템플릿으로 주입
    const props = PropertiesService.getScriptProperties();

    // 모든 데이터를 하나의 객체로 묶음
    const config = {
      staff: JSON.parse(props.getProperty("STAFF_JSON") || "[]"),
      fixedUser: JSON.parse(props.getProperty("FIXED_USER_JSON") || "{}"),
      businessTripCars: JSON.parse(
        props.getProperty("BUSINESS_TRIP_CARS_JSON") || "[]",
      ),
    };

    const tpl = HtmlService.createTemplateFromFile("ppap_form.html");
    // JSON 문자열로 변환해서 하나만 전달
    tpl.configJson = JSON.stringify(config);

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

// ── POST: 운행 기록 저장 ───────────────────────────────────
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

// ── 이전 계기판 조회 ──────────────────────────────────────
function getPrevOdoData(carNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSh = ss.getSheetByName(CONFIG.SHEET_RAW);
  const masterSh = ss.getSheetByName(CONFIG.SHEET_MASTER);
  const carName = getMasterValue(masterSh, carNo, "차종");

  const lastRow = rawSh.getLastRow();
  if (lastRow < 2) return { prevOdo: null, prevDate: null, carName };

  const data = rawSh.getRange(2, 1, lastRow - 1, 16).getValues();
  const carRows = data
    .filter((r) => r[COL.차량번호] === carNo && r[COL.주행후] > 0)
    .sort((a, b) => new Date(b[COL.사용일자]) - new Date(a[COL.사용일자]));

  if (carRows.length === 0) return { prevOdo: null, prevDate: null, carName };

  const latest = carRows[0];
  const prevDate = Utilities.formatDate(
    new Date(latest[COL.사용일자]),
    "Asia/Seoul",
    "yyyy-MM-dd",
  );
  return { prevOdo: latest[COL.주행후], prevDate, carName };
}

// ── 운행 기록 저장 ────────────────────────────────────────
function saveRecord(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
  const newRow = new Array(16).fill("");
  newRow[COL.ID] = id;
  newRow[COL.차량번호] = 차량번호;
  newRow[COL.차종] = 차종;
  newRow[COL.사용일자] = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd");
  newRow[COL.요일] = DAYS[now.getDay()];
  newRow[COL.부서] = payload.dept;
  newRow[COL.성명] = payload.name;
  newRow[COL.주행전] = 주행전;
  newRow[COL.주행후] = 주행후;
  newRow[COL.주행거리] = 주행거리;
  newRow[COL.사용구분] = 사용구분;
  newRow[COL.출퇴근] = 출퇴근;
  newRow[COL.일반업무] = 일반업무;
  newRow[COL.비고] = payload.note || "";
  newRow[COL.플래그] = flags.length > 0 ? flags.join(" | ") : "정상";
  newRow[COL.타임스탬프] = Utilities.formatDate(
    now,
    "Asia/Seoul",
    "yyyy-MM-dd HH:mm:ss",
  );

  rawSh.appendRow(newRow);

  if (flags.length > 0) {
    sendAlertEmail({
      차량번호,
      성명: payload.name,
      주행전,
      주행후,
      주행거리,
      flags,
      date: now,
    });
  }

  return { success: true, id, mileage: 주행거리, flags };
}

// ── 이상 감지 ─────────────────────────────────────────────
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
      .sort((a, b) => new Date(b[COL.사용일자]) - new Date(a[COL.사용일자]));
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

// ── 이메일 알림 ───────────────────────────────────────────
function sendAlertEmail({
  차량번호,
  성명,
  주행전,
  주행후,
  주행거리,
  flags,
  date,
}) {
  const dateStr = Utilities.formatDate(date, "Asia/Seoul", "yyyy-MM-dd HH:mm");
  GmailApp.sendEmail(
    CONFIG.ADMIN_EMAIL,
    `[PPAP 이상감지] ${차량번호} · ${성명} · ${dateStr}`,
    `이상 유형: ${flags.join(", ")}\n\n차량: ${차량번호}\n운전자: ${성명}\n` +
      `기록일시: ${dateStr}\n주행전: ${주행전}km\n주행후: ${주행후}km\n주행거리: ${주행거리}km`,
  );
}

// ── 월간 리포트 생성 ──────────────────────────────────────
function generateAllReports() {
  const now = new Date();
  const target = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const year = target.getFullYear();
  const month = target.getMonth() + 1;
  // 차량 목록도 스크립트 속성에서 읽어옴
  const props = PropertiesService.getScriptProperties();
  const allCars = [
    ...Object.keys(JSON.parse(props.getProperty("FIXED_USER_JSON") || "{}")),
    ...JSON.parse(props.getProperty("BUSINESS_TRIP_CARS_JSON") || "[]"),
  ];
  allCars.forEach((carNo) => generateMonthlyReport(carNo, year, month));
}

function generateMonthlyReport(targetCarNo, year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSh = ss.getSheetByName(CONFIG.SHEET_RAW);
  const masterSh = ss.getSheetByName(CONFIG.SHEET_MASTER);

  const monthData = rawSh
    .getDataRange()
    .getValues()
    .slice(1)
    .filter((row) => {
      const d = new Date(row[COL.사용일자]);
      return (
        row[COL.차량번호] === targetCarNo &&
        d.getFullYear() === year &&
        d.getMonth() + 1 === month
      );
    })
    .sort((a, b) => new Date(a[COL.사용일자]) - new Date(b[COL.사용일자]));

  if (monthData.length === 0) return;

  const sheetName = `${targetCarNo}_${year}${String(month).padStart(2, "0")}`;
  let reportSh = ss.getSheetByName(sheetName);
  if (reportSh) ss.deleteSheet(reportSh);
  reportSh = ss.insertSheet(sheetName);

  const masterRow = getMasterRow(masterSh, targetCarNo);
  const 차종 = masterRow ? masterRow[1] : "";
  const 법인명 = masterRow ? masterRow[3] : "";
  const 사업자번호 = masterRow ? masterRow[4] : "";

  reportSh
    .getRange("A1")
    .setValue("【업무용승용차 운행기록부】 별지 제25호 서식");
  reportSh.getRange("A2").setValue(`사업연도: ${year}년`);
  reportSh
    .getRange("A3")
    .setValue(`법인명: ${법인명}   사업자등록번호: ${사업자번호}`);
  reportSh
    .getRange("A4")
    .setValue(`①차종: ${차종}   ②자동차등록번호: ${targetCarNo}`);

  const headers = [
    "③사용일자(요일)",
    "④부서",
    "④성명",
    "⑤주행전(km)",
    "⑥주행후(km)",
    "⑦주행거리(km)",
    "⑧출퇴근용(km)",
    "⑨일반업무용(km)",
    "⑩비고",
  ];
  reportSh.getRange(6, 1, 1, headers.length).setValues([headers]);

  const START = 7;
  const DAYS = ["일", "월", "화", "수", "목", "금", "토"];
  monthData.forEach((row, idx) => {
    const r = START + idx;
    const d = new Date(row[COL.사용일자]);
    const 주행전 = idx === 0 ? row[COL.주행전] : monthData[idx - 1][COL.주행후];
    reportSh
      .getRange(r, 1)
      .setValue(`${d.getMonth() + 1}/${d.getDate()}(${DAYS[d.getDay()]})`);
    reportSh.getRange(r, 2).setValue(row[COL.부서]);
    reportSh.getRange(r, 3).setValue(row[COL.성명]);
    reportSh.getRange(r, 4).setValue(주행전);
    reportSh.getRange(r, 5).setValue(row[COL.주행후]);
    reportSh.getRange(r, 6).setValue(row[COL.주행거리]);
    reportSh.getRange(r, 7).setValue(row[COL.출퇴근]);
    reportSh.getRange(r, 8).setValue(row[COL.일반업무]);
    reportSh.getRange(r, 9).setValue(row[COL.비고]);
  });

  const sumRow = START + monthData.length;
  reportSh.getRange(sumRow, 1).setValue("합 계");
  ["F", "G", "H"].forEach((col, i) => {
    reportSh
      .getRange(sumRow, 6 + i)
      .setFormula(`=SUM(${col}${START}:${col}${sumRow - 1})`);
  });
  SpreadsheetApp.flush();
}

// ── 헬퍼 ──────────────────────────────────────────────────
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

// ── 스크립트 속성 초기 설정 ──────────────────────────────
// config.json 수정 후 clasp push → 이 함수 실행 → 재배포
function setupProperties() {
  const config = JSON.parse(
    HtmlService.createHtmlOutputFromFile("config").getContent(),
  );
  const props = PropertiesService.getScriptProperties();

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
