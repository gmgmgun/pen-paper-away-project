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

// ── 엑셀 양식 컬럼 위치 (1-based) ────────────────────────
// A=1, F=6, J=10, N=14, T=20, Z=26, AF=32, AL=38, AR=44
const XLSX_COL = {
  사용일자: 1, // A
  부서: 6, // F
  성명: 10, // J
  주행전: 14, // N
  주행후: 20, // T
  주행거리: 26, // Z
  출퇴근: 32, // AF
  일반업무: 38, // AL
  비고: 44, // AR
};

const XLSX_DATA_START_ROW = 15; // 데이터 시작 행

// 고정 사용자 차량 (1인 전담): T열 = =N+Z 수식, N열에 주행전 직접 입력
// 출장/공용 차량: N열 = 이전행 T 수식, T열에 주행후 직접 입력, Z=T-N 수식
const FIXED_CAR_PATTERN = ["240서7489", "07누8546"]; // 고정 차량 번호
const SHARED_CAR_PATTERN = ["200호7074", "208호1041"]; // 공용/출장 차량 번호

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

// ── GET: HTML 서빙 + JSON API ──────────────────────────────
function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : null;

    if (action === "getPrevOdo") {
      const carNo = e.parameter.car;
      if (!carNo) throw new Error("차량번호 없음");
      const result = getPrevOdoData(carNo);
      return ContentService.createTextOutput(
        JSON.stringify(result),
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const props = PropertiesService.getScriptProperties();
    const config = {
      staff: JSON.parse(props.getProperty("STAFF_JSON") || "[]"),
      fixedUser: JSON.parse(props.getProperty("FIXED_USER_JSON") || "{}"),
      businessTripCars: JSON.parse(
        props.getProperty("BUSINESS_TRIP_CARS_JSON") || "[]",
      ),
    };

    const tpl = HtmlService.createTemplateFromFile("ppap_form.html");
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

  // 1) RAW 시트에서 최근 기록 조회
  const rawSh = ss.getSheetByName(CONFIG.SHEET_RAW);
  const masterSh = ss.getSheetByName(CONFIG.SHEET_MASTER);
  const carName = getMasterValue(masterSh, carNo, "차종");

  const lastRow = rawSh.getLastRow();
  if (lastRow >= 2) {
    const data = rawSh.getRange(2, 1, lastRow - 1, 16).getValues();
    const carRows = data
      .filter((r) => r[COL.차량번호] === carNo && r[COL.주행후] > 0)
      .sort((a, b) => new Date(b[COL.사용일자]) - new Date(a[COL.사용일자]));

    if (carRows.length > 0) {
      const latest = carRows[0];
      const prevDate = Utilities.formatDate(
        new Date(latest[COL.사용일자]),
        "Asia/Seoul",
        "yyyy-MM-dd",
      );
      return { prevOdo: latest[COL.주행후], prevDate, carName };
    }
  }

  // 2) RAW에 없으면 해당 차량 탭 엑셀 시트에서 마지막 주행후 값 조회
  const xlsxOdo = getLastOdoFromXlsxSheet(ss, carNo);
  if (xlsxOdo !== null) {
    return { prevOdo: xlsxOdo.odo, prevDate: xlsxOdo.date, carName };
  }

  return { prevOdo: null, prevDate: null, carName };
}

// ── 엑셀 시트에서 마지막 계기판 값 조회 ─────────────────
function getLastOdoFromXlsxSheet(ss, carNo) {
  const sh = ss.getSheetByName(carNo);
  if (!sh) return null;

  const isFixed = FIXED_CAR_PATTERN.indexOf(carNo) !== -1;
  // 고정 차량: T열(주행후=수식), N열(주행전)에서 마지막 값
  // 공용 차량: T열(주행후)에서 마지막 값
  const odoCol = XLSX_COL.주행후; // T = 20

  let lastRow = XLSX_DATA_START_ROW - 1;
  const maxRow = sh.getLastRow();
  for (let r = XLSX_DATA_START_ROW; r <= maxRow; r++) {
    const dateVal = sh.getRange(r, XLSX_COL.사용일자).getValue();
    const odoVal = sh.getRange(r, odoCol).getValue();
    if (dateVal && odoVal && !isNaN(Number(odoVal))) {
      lastRow = r;
    }
  }
  if (lastRow < XLSX_DATA_START_ROW) return null;

  const odo = Number(sh.getRange(lastRow, odoCol).getValue());
  const rawDate = sh.getRange(lastRow, XLSX_COL.사용일자).getValue();
  const date = rawDate
    ? Utilities.formatDate(new Date(rawDate), "Asia/Seoul", "yyyy-MM-dd")
    : null;
  return { odo, date };
}

// ── 운행 기록 저장 (RAW + 엑셀 양식 탭) ──────────────────
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

  // 1) RAW 시트에 저장 (기존 방식 유지)
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

  // 2) 차량별 엑셀 양식 탭에 저장
  writeToXlsxSheet(ss, {
    차량번호,
    부서: payload.dept,
    성명: payload.name,
    주행전,
    주행후,
    주행거리,
    출퇴근,
    일반업무,
    비고: payload.note || "",
    사용일자: now,
    요일: DAYS[now.getDay()],
  });

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

// ── 엑셀 양식 탭에 행 추가 ───────────────────────────────
function writeToXlsxSheet(ss, data) {
  const sh = ss.getSheetByName(data.차량번호);
  if (!sh) {
    Logger.log("시트 없음: " + data.차량번호);
    return;
  }

  const isFixed = FIXED_CAR_PATTERN.indexOf(data.차량번호) !== -1;

  // 빈 행 찾기: A열(사용일자)이 비어 있는 첫 번째 행
  let targetRow = XLSX_DATA_START_ROW;
  const maxRow = sh.getLastRow();
  for (let r = XLSX_DATA_START_ROW; r <= maxRow + 1; r++) {
    const aVal = sh.getRange(r, XLSX_COL.사용일자).getValue();
    const nVal = sh.getRange(r, XLSX_COL.주행전).getValue();
    // 빈 행 = A열과 N열 모두 비어있는 경우
    if (!aVal && (nVal === "" || nVal === null)) {
      targetRow = r;
      break;
    }
    targetRow = r + 1; // 계속 채워져 있으면 다음 행
  }

  const r = targetRow;
  const dateStr =
    data.사용일자.getMonth() +
    1 +
    "/" +
    data.사용일자.getDate() +
    "(" +
    data.요일 +
    ")";

  // A열: 사용일자(요일)
  sh.getRange(r, XLSX_COL.사용일자).setValue(dateStr);
  // F열: 부서
  sh.getRange(r, XLSX_COL.부서).setValue(data.부서);
  // J열: 성명
  sh.getRange(r, XLSX_COL.성명).setValue(data.성명);

  if (isFixed) {
    // 고정 차량 패턴: N=주행전(값), T=수식(=N+Z), Z=주행거리(값)
    sh.getRange(r, XLSX_COL.주행전).setValue(data.주행전);
    sh.getRange(r, XLSX_COL.주행후).setFormula("=N" + r + "+Z" + r);
    sh.getRange(r, XLSX_COL.주행거리).setValue(data.주행거리);
    // 출퇴근/일반업무
    if (data.출퇴근 > 0) {
      sh.getRange(r, XLSX_COL.출퇴근).setValue(data.출퇴근);
      sh.getRange(r, XLSX_COL.일반업무).setValue(0);
    } else {
      sh.getRange(r, XLSX_COL.출퇴근).setValue(0);
      sh.getRange(r, XLSX_COL.일반업무).setValue(data.일반업무);
    }
  } else {
    // 공용/출장 차량 패턴: N=이전행T 수식(단 첫행은 값), T=주행후(값), Z=수식(T-N), AL=수식
    if (r === XLSX_DATA_START_ROW) {
      // 첫 데이터 행: N열은 초기 계기판 값
      sh.getRange(r, XLSX_COL.주행전).setValue(data.주행전);
    } else {
      // 이후 행: N열 = 이전 행 T열 수식
      sh.getRange(r, XLSX_COL.주행전).setFormula("=T" + (r - 1));
    }
    sh.getRange(r, XLSX_COL.주행후).setValue(data.주행후);
    sh.getRange(r, XLSX_COL.주행거리).setFormula("=T" + r + "-N" + r);

    // 출퇴근(AF열)과 일반업무(AL열)
    sh.getRange(r, XLSX_COL.출퇴근).setValue(data.출퇴근 > 0 ? data.출퇴근 : 0);
    sh.getRange(r, XLSX_COL.일반업무).setFormula("=Z" + r + "-AF" + r);
  }

  // AR열: 비고
  if (data.비고) {
    sh.getRange(r, XLSX_COL.비고).setValue(data.비고);
  }

  SpreadsheetApp.flush();
  Logger.log(
    "[XLSX 기록] " +
      data.차량번호 +
      " row" +
      r +
      " → " +
      data.성명 +
      " / " +
      data.주행후 +
      "km",
  );
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

// ── 월간 리포트 생성 (RAW 기준 — 별도 요약 시트 필요 시) ─
function generateAllReports() {
  const now = new Date();
  const target = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const year = target.getFullYear();
  const month = target.getMonth() + 1;
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
