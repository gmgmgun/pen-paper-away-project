// ============================================================

// PPAP 운행기록 시스템 — Google Apps Script 백엔드

// ============================================================

const CONFIG = {
  SHEET_RAW: "RAW_운행일지",

  SHEET_MASTER: "차량_마스터",

  ADMIN_EMAIL: "dmlee@greenia.co.kr",

  MAX_DAILY_KM: 500,

  GAP_ALERT_DAYS: 3,

  XLSX_DATA_START_ROW: 15, // 엑셀 양식 데이터 시작 행 (고정)
};

// RAW 시트 컬럼 인덱스 (0-based)

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

// 엑셀 양식 컬럼 번호 (1-based, openpyxl 기준과 동일)

// Google Sheets에서도 getRange(row, col)은 1-based

const XLSX_COL = {
  사용일자: 1, // A

  부서: 6, // F

  성명: 10, // J

  주행전: 14, // N  ← 공용차: =T(r-1) 수식 / 고정차: 숫자값

  주행후: 20, // T  ← 반드시 setValue (수식 구조의 기준점)

  주행거리: 26, // Z  ← 공용차: =T(r)-N(r) 수식 유지

  출퇴근: 32, // AF

  일반업무: 38, // AL ← 공용차: =Z(r)-AF(r) 수식 유지

  비고: 44, // AR
};

// ── GET: HTML 서빙 ────────────────────────────────────────

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

    const carNo =
      e && e.parameter && e.parameter.car ? e.parameter.car.trim() : "";
    let prevOdoJson = JSON.stringify({
      prevOdo: null,
      prevDate: null,
      carName: "",
    });

    if (carNo) {
      try {
        prevOdoJson = JSON.stringify(getPrevOdoData(carNo));
      } catch (err) {
        prevOdoJson = JSON.stringify({
          prevOdo: null,
          prevDate: null,
          carName: "",
        });
      }
    }

    const tpl = HtmlService.createTemplateFromFile("ppap_form");
    tpl.configJson = JSON.stringify(config);
    tpl.carNo = carNo;
    tpl.prevOdoJson = prevOdoJson;

    return tpl
      .evaluate()
      .setTitle("PPAP 운행기록")
      .addMetaTag(
        "viewport",
        "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no",
      ) // 모바일 확대 방지
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: err.message }),
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ── saveRecord: google.script.run으로 직접 호출 ──────────

function saveRecord(payload) {
  try {
    return _saveRecord(payload);
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ============================================================

// ── getPrevOdoData: 직전 계기판 조회 ─────────────────────

// 우선순위: 엑셀 양식 탭(T열 마지막값) > RAW 시트(주행후 최신값)

// 이유: 엑셀 탭에 수기 보정된 값이 있을 수 있고,

//       RAW보다 엑셀 탭이 실제 운행기록부의 원본이기 때문

// ============================================================

function getPrevOdoData(carNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const masterSh = ss.getSheetByName(CONFIG.SHEET_MASTER);

  const carName = getMasterValue(masterSh, carNo, "차종");

  // 1) 엑셀 양식 탭에서 조회 (차량번호와 동일한 이름의 시트)

  const xlsxSh = ss.getSheetByName(carNo);

  if (xlsxSh) {
    const xlsxResult = _getLastOdoFromXlsxSheet(xlsxSh);

    if (xlsxResult.prevOdo !== null) {
      return {
        prevOdo: xlsxResult.prevOdo,
        prevDate: xlsxResult.prevDate,
        carName,
      };
    }
  }

  // 2) 엑셀 탭에 값 없으면 RAW 시트 fallback

  const rawSh = ss.getSheetByName(CONFIG.SHEET_RAW);

  if (!rawSh) return { prevOdo: null, prevDate: null, carName };

  const lastRow = rawSh.getLastRow();

  if (lastRow < 2) return { prevOdo: null, prevDate: null, carName };

  const data = rawSh.getRange(2, 1, lastRow - 1, 16).getValues();

  const carRows = data

    .filter((r) => r[COL.차량번호] === carNo && Number(r[COL.주행후]) > 0)

    .sort((a, b) => new Date(b[COL.사용일자]) - new Date(a[COL.사용일자]));

  if (carRows.length === 0) return { prevOdo: null, prevDate: null, carName };

  const latest = carRows[0];

  const prevDate = Utilities.formatDate(
    new Date(latest[COL.사용일자]),
    "Asia/Seoul",
    "yyyy-MM-dd",
  );

  return { prevOdo: Number(latest[COL.주행후]), prevDate, carName };
}

// 엑셀 탭에서 T열(주행 후) 마지막 실제 값 조회

// ✅ 핵심: T열에 getValue()로 읽은 값이 있는 마지막 행을 찾음

//    N열은 수식으로 가득 차 있어서 기준이 될 수 없음

function _getLastOdoFromXlsxSheet(sheet) {
  const START = CONFIG.XLSX_DATA_START_ROW;

  const lastRow = sheet.getLastRow();

  if (lastRow < START) return { prevOdo: null, prevDate: null };

  // T열(주행후, 20번) 전체를 한번에 읽어 마지막 유효값 탐색

  const tColValues = sheet

    .getRange(START, XLSX_COL.주행후, lastRow - START + 1, 1)

    .getValues(); // [[val], [val], ...]

  let lastOdo = null;

  let lastOdoRowOffset = -1;

  for (let i = 0; i < tColValues.length; i++) {
    const val = tColValues[i][0];

    const num = Number(val);

    if (val !== "" && val !== null && !isNaN(num) && num > 0) {
      lastOdo = num;

      lastOdoRowOffset = i;
    }
  }

  if (lastOdo === null) return { prevOdo: null, prevDate: null };

  // 날짜는 A열에서 같은 행 조회

  const dateRow = START + lastOdoRowOffset;

  const dateVal = sheet.getRange(dateRow, XLSX_COL.사용일자).getValue();

  let prevDate = null;

  if (dateVal instanceof Date && !isNaN(dateVal)) {
    prevDate = Utilities.formatDate(dateVal, "Asia/Seoul", "yyyy-MM-dd");
  }

  return { prevOdo: lastOdo, prevDate };
}

// ============================================================

// ── _saveRecord: 운행 기록 저장 ──────────────────────────

// ① RAW 시트에 로그 적재

// ② 엑셀 양식 탭에 데이터 기입 (writeToXlsxSheet 호출)

// ============================================================

function _saveRecord(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const rawSh = ss.getSheetByName(CONFIG.SHEET_RAW);

  const masterSh = ss.getSheetByName(CONFIG.SHEET_MASTER);

  if (!rawSh) throw new Error("RAW_운행일지 시트를 찾을 수 없습니다.");

  const now = new Date();

  const DAYS = ["일", "월", "화", "수", "목", "금", "토"];

  const 차량번호 = payload.carNo;

  const 주행후 = Number(payload.currentOdo);

  // prevOdo가 null이면(최초 등록) 주행전 = 주행후로 처리 (주행거리 0)

  const 주행전 =
    payload.prevOdo !== null && payload.prevOdo !== undefined
      ? Number(payload.prevOdo)
      : 주행후;

  const 주행거리 = 주행후 - 주행전;

  // NaN 가드

  if (isNaN(주행후) || isNaN(주행전)) {
    throw new Error("계기판 값이 유효하지 않습니다. (NaN)");
  }

  const 사용구분 = payload.useType;

  const 출퇴근 = 사용구분 === "출퇴근용" ? 주행거리 : 0;

  const 일반업무 = 사용구분 === "일반업무용" ? 주행거리 : 0;

  const 차종 = getMasterValue(masterSh, 차량번호, "차종") || "";

  const 날짜문자열 = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd");

  const flags = detectAnomalies({
    rawSh,

    차량번호,

    주행전,

    주행후,

    주행거리,

    사용일자: now,

    prevOdo: payload.prevOdo,
  });

  // ① RAW 시트 기록

  const id = Utilities.getUuid();

  const newRow = new Array(16).fill("");

  newRow[COL.ID] = id;

  newRow[COL.차량번호] = 차량번호;

  newRow[COL.차종] = 차종;

  newRow[COL.사용일자] = 날짜문자열;

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

  SpreadsheetApp.flush(); // RAW 기록 먼저 확정

  // ② 엑셀 양식 탭 기록

  try {
    writeToXlsxSheet({
      ss,

      차량번호,

      날짜: now,

      부서: payload.dept,

      성명: payload.name,

      주행전,

      주행후,

      주행거리,

      출퇴근,

      일반업무,

      비고: payload.note || "",
    });
  } catch (xlsxErr) {
    // 엑셀 탭 기록 실패 시 RAW는 보존, 에러 플래그만 추가

    Logger.log("엑셀 탭 기록 실패: " + xlsxErr.message);

    // RAW 시트 마지막 행 플래그에 기록

    const lastRawRow = rawSh.getLastRow();

    const currentFlag = rawSh.getRange(lastRawRow, COL.플래그 + 1).getValue();

    rawSh
      .getRange(lastRawRow, COL.플래그 + 1)

      .setValue(currentFlag + " | XLSX기록실패");
  }

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

// ============================================================

// ── writeToXlsxSheet: 엑셀 양식 탭에 데이터 기입 ─────────

//

// 차량 유형별 시트 구조:

//

//   [공용차] 200호7074, 208호1041

//     N열(주행전) 행15  : 숫자값 (초기 계기판)

//     N열(주행전) 행16~ : =T(r-1) 수식 (전행 주행후 자동참조)

//     T열(주행후)       : setValue — 수식 체인의 기준점

//     Z열(주행거리)     : =T(r)-N(r) 수식 유지

//     AF열(출퇴근)      : setValue

//     AL열(일반업무)    : =Z(r)-AF(r) 수식 유지

//     SUM 범위          : Z15:Z188, AF15:AF188, AL15:AL188

//

//   [고정차] 240서7489, 07누8546

//     N열(주행전)       : 숫자값 (최초 계기판 고정, 이후 업데이트)

//     T열(주행후)       : =N(r)+Z(r) 수식 — 건드리지 않음

//     Z열(주행거리)     : setValue — 직접 입력값

//     AF열(출퇴근)      : setValue

//     AL열(일반업무)    : setValue

//     SUM 범위          : Z15:Z(SUM행-2)

//

// targetRow 결정 기준:

//   공용차 → T열 비어있는 첫 행  (N열은 수식으로 가득 차 있어 사용 불가)

//   고정차 → Z열 비어있는 첫 행  (T열은 =N+Z 수식이라 값 비교 불가)

//

// 다음 행 Pre-fill 없음: 현재 targetRow 행만 처리,

//   r+1 행에 수식/값을 미리 넣지 않음 → SUM 오염 방지

// ============================================================

function writeToXlsxSheet({
  ss,
  차량번호,
  날짜,
  부서,
  성명,
  주행전,
  주행후,
  주행거리,
  출퇴근,
  일반업무,
  비고,
}) {
  // ── 1. 시트 존재 확인 ──────────────────────────────────

  const sheet = ss.getSheetByName(차량번호);

  if (!sheet) {
    throw new Error("엑셀 양식 탭을 찾을 수 없습니다: " + 차량번호);
  }

  const START = CONFIG.XLSX_DATA_START_ROW; // 15

  // ── 2. 차량 유형 판별 ──────────────────────────────────

  // 고정차: config의 fixedUser 키 목록과 일치하는 차량

  // 공용차: businessTripCars 목록 (출장용)

  // 판별 기준: 행15 T열에 수식(=N+Z)이 있으면 고정차, 없으면 공용차

  const row15TFormula = sheet

    .getRange(START, XLSX_COL.주행후)

    .getFormula()

    .trim()

    .toUpperCase();

  const isFixedCar = row15TFormula.startsWith("=N");

  // ── 3. targetRow 결정 ──────────────────────────────────

  // ✅ Pre-fill 없이 실제 데이터가 없는 첫 번째 행을 탐색

  //    공용차: T열(주행후) 기준 — N열은 수식으로 가득 차 기준 불가

  //    고정차: Z열(주행거리) 기준 — T열은 =N+Z 수식이라 값 비교 불가

  const scanColIdx = isFixedCar ? XLSX_COL.주행거리 : XLSX_COL.주행후;

  const lastRow = sheet.getLastRow();

  const searchEnd = Math.max(lastRow, START);

  // 해당 열 전체를 한 번에 읽어 API 호출 최소화

  const scanColData = sheet

    .getRange(START, scanColIdx, searchEnd - START + 1, 1)

    .getValues();

  let targetRow = -1;

  for (let i = 0; i < scanColData.length; i++) {
    const val = scanColData[i][0];

    // 수식 문자열, 숫자 0을 제외한 빈 셀만 미기입 행으로 판단

    // getValues()는 수식의 계산 결과를 반환 — 미기입이면 빈 문자열 또는 0

    if (val === "" || val === null) {
      targetRow = START + i;

      break;
    }
  }

  // 모든 행이 차있으면 SUM 행 직전에 기록 (SUM 수식 오염 방지)

  // SUM 행은 마지막 데이터 행 + 2 위치에 있음 (빈 행 1개 간격)

  if (targetRow === -1) {
    targetRow = searchEnd; // 마지막으로 값 있는 행 다음

    Logger.log("경고: 모든 데이터 행이 소진됨 — 행 " + targetRow + "에 기록");
  }

  // ── 4. 날짜 포맷 ───────────────────────────────────────

  const DAYS = ["일", "월", "화", "수", "목", "금", "토"];

  const dateLabel = Utilities.formatDate(날짜, "Asia/Seoul", "yyyy-MM-dd");
  sheet.getRange(targetRow, XLSX_COL.사용일자).setValue(dateLabel);
  // ── 5. 공통 필드 기입 ──────────────────────────────────

  sheet.getRange(targetRow, XLSX_COL.사용일자).setValue(dateLabel);

  sheet.getRange(targetRow, XLSX_COL.부서).setValue(부서);

  sheet.getRange(targetRow, XLSX_COL.성명).setValue(성명);

  if (비고) sheet.getRange(targetRow, XLSX_COL.비고).setValue(비고);

  // ── 6. 차량 유형별 분기 기입 ───────────────────────────

  if (isFixedCar) {
    // ────────────────────────────────────────────────────

    // [고정차] N=숫자값, T=수식(건드리지않음), Z=직접입력

    // ────────────────────────────────────────────────────

    // N열(주행전): 최초 행이면 숫자값 갱신, 이후 행은 이전 T열 계산값을 직접 기입

    // 고정차는 N열에 수식 체인이 없으므로 모두 setValue

    if (targetRow === START) {
      // 행15: 이미 초기 계기판 값이 세팅된 경우 덮어쓰지 않음

      // (최초 등록이 아닌 경우 prevOdo가 주행전으로 전달됨)

      const existingN = sheet.getRange(START, XLSX_COL.주행전).getValue();

      if (!existingN || existingN === 0) {
        sheet.getRange(targetRow, XLSX_COL.주행전).setValue(주행전);
      }
    } else {
      // 행16~: 이전 행의 T열 계산 결과를 직접 읽어 숫자로 기입

      // T열 수식(=N+Z)이 이미 계산되어 있으므로 getValue()가 숫자 반환

      const prevTVal = sheet
        .getRange(targetRow - 1, XLSX_COL.주행후)
        .getValue();

      const prevTNum = Number(prevTVal);

      sheet
        .getRange(targetRow, XLSX_COL.주행전)
        .setValue(isNaN(prevTNum) || prevTNum === 0 ? 주행전 : prevTNum);
    }

    // Z열(주행거리): 직접 입력 — T열 수식(=N+Z)의 기준점

    sheet.getRange(targetRow, XLSX_COL.주행거리).setValue(주행거리);

    // T열(주행후): =N(r)+Z(r) 수식이 이미 존재 → 절대 건드리지 않음

    // 단, 수식 부재 시에만 복구

    const tFormula = sheet.getRange(targetRow, XLSX_COL.주행후).getFormula();

    if (!tFormula) {
      sheet
        .getRange(targetRow, XLSX_COL.주행후)

        .setFormula("=N" + targetRow + "+Z" + targetRow);
    }

    // AF열(출퇴근), AL열(일반업무): 직접 입력

    sheet.getRange(targetRow, XLSX_COL.출퇴근).setValue(출퇴근);

    sheet.getRange(targetRow, XLSX_COL.일반업무).setValue(일반업무);
  } else {
    // ────────────────────────────────────────────────────

    // [공용차] N=수식(=T(r-1)), T=실제값, Z/AL=수식

    // ────────────────────────────────────────────────────

    // N열(주행전) 처리:

    //   행15 (첫 행): 수식 대신 숫자 직접 입력 (이전 행 없음)

    //   행16~ : =T(r-1) 수식이 기존에 있으면 유지, 없거나 깨지면 복구

    //   ✅ r+1 행에 미리 수식 주입하는 Pre-fill 완전 제거

    if (targetRow === START) {
      sheet.getRange(targetRow, XLSX_COL.주행전).setValue(주행전);
    } else {
      const expectedNFormula = "=T" + (targetRow - 1);

      const existingNFormula = sheet

        .getRange(targetRow, XLSX_COL.주행전)

        .getFormula()

        .trim()

        .toUpperCase();

      if (existingNFormula !== expectedNFormula.toUpperCase()) {
        // 수식 없거나 잘못된 경우에만 복구 — 정상이면 절대 건드리지 않음

        sheet.getRange(targetRow, XLSX_COL.주행전).setFormula(expectedNFormula);
      }
    }

    // T열(주행후): 반드시 실제값 — Z열 수식(=T-N)의 기준점

    sheet.getRange(targetRow, XLSX_COL.주행후).setValue(주행후);

    // Z열(주행거리): =T(r)-N(r) 수식 유지, 없으면 복구

    const zFormula = sheet.getRange(targetRow, XLSX_COL.주행거리).getFormula();

    if (!zFormula) {
      sheet
        .getRange(targetRow, XLSX_COL.주행거리)

        .setFormula("=T" + targetRow + "-N" + targetRow);
    }

    // AF열(출퇴근): 직접 입력

    sheet.getRange(targetRow, XLSX_COL.출퇴근).setValue(출퇴근);

    // AL열(일반업무): =Z(r)-AF(r) 수식 유지, 없으면 복구

    const alFormula = sheet.getRange(targetRow, XLSX_COL.일반업무).getFormula();

    if (!alFormula) {
      sheet
        .getRange(targetRow, XLSX_COL.일반업무)

        .setFormula("=Z" + targetRow + "-AF" + targetRow);
    }
  }

  // ── 7. 반영 확정 ───────────────────────────────────────

  SpreadsheetApp.flush();

  Logger.log(
    "엑셀 탭 기록 완료: [" +
      (isFixedCar ? "고정차" : "공용차") +
      "] " +
      차량번호 +
      " 행" +
      targetRow +
      " | 주행전=" +
      주행전 +
      " 주행후=" +
      주행후 +
      " 거리=" +
      주행거리,
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

  if (isNaN(주행거리)) {
    flags.push("주행거리NaN");

    return flags;
  }

  if (주행거리 < 0) {
    flags.push("역주행감지(" + 주행거리 + "km)");

    return flags;
  }

  if (주행거리 > CONFIG.MAX_DAILY_KM) {
    flags.push("과다주행(" + 주행거리 + "km)");
  }

  if (prevOdo !== null && prevOdo !== undefined) {
    const diff = Math.abs(주행전 - Number(prevOdo));

    if (diff > 0) flags.push("계기판불일치(차이:" + diff + "km)");
  }

  const lastRow = rawSh.getLastRow();

  if (lastRow >= 2) {
    const data = rawSh.getRange(2, 1, lastRow - 1, 16).getValues();

    const carRows = data

      .filter((r) => r[COL.차량번호] === 차량번호 && Number(r[COL.주행후]) > 0)

      .sort((a, b) => new Date(b[COL.사용일자]) - new Date(a[COL.사용일자]));

    if (carRows.length > 0) {
      const dayGap = Math.floor(
        (사용일자 - new Date(carRows[0][COL.사용일자])) / (1000 * 60 * 60 * 24),
      );

      if (dayGap > CONFIG.GAP_ALERT_DAYS) {
        flags.push(dayGap + "일공백(누락의심)");
      }
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

    "[PPAP 이상감지] " + 차량번호 + " · " + 성명 + " · " + dateStr,

    "이상 유형: " +
      flags.join(", ") +
      "\n\n차량: " +
      차량번호 +
      "\n운전자: " +
      성명 +
      "\n기록일시: " +
      dateStr +
      "\n주행전: " +
      주행전 +
      "km\n주행후: " +
      주행후 +
      "km\n주행거리: " +
      주행거리 +
      "km",
  );
}

// ── 월간 리포트 생성 ──────────────────────────────────────

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

  const sheetName = targetCarNo + "_" + year + String(month).padStart(2, "0");

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

  reportSh.getRange("A2").setValue("사업연도: " + year + "년");

  reportSh
    .getRange("A3")
    .setValue("법인명: " + 법인명 + "   사업자등록번호: " + 사업자번호);

  reportSh
    .getRange("A4")
    .setValue("①차종: " + 차종 + "   ②자동차등록번호: " + targetCarNo);

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
      .setValue(
        d.getMonth() + 1 + "/" + d.getDate() + "(" + DAYS[d.getDay()] + ")",
      );

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

      .setFormula("=SUM(" + col + START + ":" + col + (sumRow - 1) + ")");
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
