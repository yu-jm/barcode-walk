// ========== API 라우터 ==========
// 화면(HTML)은 버셀에 있고, 이 GAS는 데이터 API로만 동작한다.
// 호출 형식: .../exec?action=함수명&payload=JSON문자열&callback=콜백명
// 응답: callback이 있으면 JSONP( callback({...}) )로 반환 → 브라우저 CORS 우회
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || "";
  var callback = (e && e.parameter && e.parameter.callback) || "";

  var payload = {};
  try {
    if (e && e.parameter && e.parameter.payload) {
      payload = JSON.parse(e.parameter.payload);
    }
  } catch (err) {
    payload = {};
  }

  var result;
  switch (action) {
    case "getManagerList":
      result = getManagerList();
      break;
    case "getDataForSearch":
      result = getDataForSearch();
      break;
    case "getRecentWalks":
      result = getRecentWalks();
      break;
    case "addStudent":
      result = addStudent(
        payload.id,
        payload.ban,
        payload.kbn,
        payload.name,
        payload.dateTime,
        payload.manager,
        payload.ksuk,
        payload.qtumAll
      );
      break;
    case "setEndWalkTime":
      result = setEndWalkTime(payload.id, payload.walkTime, payload.endTime);
      break;
    default:
      result = { error: "알 수 없는 action: " + action };
  }

  var json = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

var url =
  "https://docs.google.com/spreadsheets/d/1b55BZPhlX5A4LBru37BcPEpmx9d7sV2V_uBhJIskTpo/edit?gid=0";

function getDataForSearch() {
  const ss = SpreadsheetApp.openByUrl(url);
  const ws = ss.getSheetByName("master");
  const data = ws.getRange(2, 1, ws.getLastRow() - 1, 11).getValues();

  return data;
}

// 담당자 목록 조회 함수
function getManagerList() {
  const ss = SpreadsheetApp.openByUrl(url);
  let ws;
  try {
    ws = ss.getSheetByName("wakeup담당자");
  } catch (e) {
    // 시트가 없으면 빈 배열 반환
    return [];
  }

  if (!ws) {
    return [];
  }

  const lastRow = ws.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  // A열에서 담당자명 가져오기 (A1은 헤더이므로 A2부터)
  const data = ws.getRange(2, 1, lastRow - 1, 1).getValues();
  const managers = [];

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() !== "") {
      managers.push(data[i][0].toString().trim());
    }
  }

  return managers;
}

function addStudent(id, ban, kbn, name, h_dateTime, manager, ksuk, qtumAll) {
  const ss = SpreadsheetApp.openByUrl(url);
  const ws = ss.getSheetByName("walk");

  // 저장 순서: 학번, 반, 구분, 이름, walk(날짜시간), 담당자, 기숙사, 퀀텀정보
  ws.appendRow([id, ban, kbn, name, h_dateTime, manager, ksuk, qtumAll || ""]);

  // 저장 후 최신 산책 목록 반환
  return getRecentWalks();
}

// 새로운 함수: 최신 산책 목록 및 타임아웃 설정 조회
function getRecentWalks() {
  const ss = SpreadsheetApp.openByUrl(url);
  const ws = ss.getSheetByName("walk");
  const lastRow = ws.getLastRow();
  
  // J1에서 타임아웃(분) 설정 가져오기 (기본값 10)
  let timeoutMinutes = 10;
  try {
    const j1Value = ws.getRange("J1").getValue();
    if (j1Value && !isNaN(j1Value)) {
      timeoutMinutes = parseInt(j1Value, 10);
    }
  } catch(e) {}
  
  if (lastRow <= 1) return { data: [], timeout: timeoutMinutes };
  
  // 최근 100건 정도 가져오기
  const startRow = Math.max(2, lastRow - 99);
  const numRows = lastRow - startRow + 1;
  
  // 컬럼 1~9: 학번, 반, 구분, 이름, walk, 담당자, 기숙사, 퀀텀, 산책종료시간
  const data = ws.getRange(startRow, 1, numRows, 9).getDisplayValues();
  
  // 역순 정렬 (최신이 위로)
  data.reverse();
  
  const recentData = data.map(function(row) {
    return {
      id: row[0],
      ban: row[1],
      kbn: row[2],
      name: row[3],
      walkTime: row[4],
      manager: row[5],
      ksuk: row[6],
      qtumAll: row[7],
      endTime: row[8] || ""
    };
  });
  
  return { data: recentData, timeout: timeoutMinutes };
}

// 새로운 함수: 산책 종료 시간 기록
function setEndWalkTime(id, walkTime, endTime) {
  const ss = SpreadsheetApp.openByUrl(url);
  const ws = ss.getSheetByName("walk");
  const lastRow = ws.getLastRow();
  
  if (lastRow <= 1) return false;
  
  // 최근 500건 정도 안에서 찾기 (역순 검색)
  const startRow = Math.max(2, lastRow - 499);
  const numRows = lastRow - startRow + 1;
  const data = ws.getRange(startRow, 1, numRows, 5).getDisplayValues();
  
  for (let i = numRows - 1; i >= 0; i--) {
    if (data[i][0] == id && data[i][4] == walkTime) {
      // 찾았을 경우 I열(9번째)에 기록
      const targetRow = startRow + i;
      ws.getRange(targetRow, 9).setValue(endTime);
      return true;
    }
  }
  return false;
}

function jandi(id, ban, kbn, name, h_dateTime, manager, ksuk, qtumAll) {
  var message =
    "[산책 체크]" +
    "\n" +
    "- " +
    name +
    "(" +
    id +
    ")" +
    " / " +
    ban +
    kbn +
    "반" +
    "\n" +
    "- " +
    h_dateTime +
    "\n" +
    "  " +
    manager +
    "\n";
  var jandi_incoming_url =
    "https://wh.jandi.com/connect-api/webhook/30603710/90a34782951d3395bf4f5e25aef14647";
  var jandi_headers = {
    Accept: "application/vnd.tosslab.jandi-v2+json",
    "Content-type": "application/json",
  };
  var jandi_formData = {
    body: message,
  };
  var jandi_options = {
    method: "POST",
    payload: JSON.stringify(jandi_formData),
    headers: jandi_headers,
  };

  response = UrlFetchApp.fetch(jandi_incoming_url, jandi_options);
}

// 학생 통계 조회 함수
function getStudentStats(studentId) {
  const ss = SpreadsheetApp.openByUrl(url);
  const ws = ss.getSheetByName("walk");

  // 모든 데이터 가져오기 (헤더 제외)
  const lastRow = ws.getLastRow();
  if (lastRow <= 1) {
    return {
      total: 0,
      week: 0,
      month: 0,
      t1: 0, // 아침 06:30-07:59
      t2: 0, // 오전 08:00-11:59
      t3: 0, // 오후 13:00-17:00
      t4: 0, // 저녁 18:00-23:00
    };
  }

  const data = ws.getRange(2, 1, lastRow - 1, 5).getValues(); // id, ban, kbn, name, h_dateTime

  // 현재 날짜 기준 (시간 제거하여 날짜만 비교)
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth() + 1; // 1-12
  const currentWeekStart = getWeekStart(now);
  const currentWeekEnd = new Date(currentWeekStart);
  currentWeekEnd.setDate(currentWeekEnd.getDate() + 6);
  currentWeekEnd.setHours(23, 59, 59, 999);

  let total = 0;
  let week = 0;
  let month = 0;
  let t1 = 0; // 아침 06:30-07:59
  let t2 = 0; // 오전 08:00-11:59
  let t3 = 0; // 오후 13:00-17:00
  let t4 = 0; // 저녁 18:00-23:00

  for (let i = 0; i < data.length; i++) {
    // 첫 번째 컬럼이 학생 ID
    if (data[i][0] == studentId) {
      total++;

      // 날짜 파싱 (h_dateTime 형식: "YYYY-MM-DD HH:mm:ss" 또는 Date 객체)
      let recordDate;
      const dateTimeValue = data[i][4];

      if (!dateTimeValue) continue;

      // Date 객체인 경우
      if (dateTimeValue instanceof Date) {
        recordDate = new Date(dateTimeValue);
      }
      // 문자열인 경우
      else if (typeof dateTimeValue === "string") {
        // "YYYY-MM-DD HH:mm:ss" 형식 파싱
        const dateTimeStr = dateTimeValue.trim();
        // 공백을 'T'로 변경하여 ISO 형식으로 변환
        const isoStr = dateTimeStr.replace(" ", "T");
        recordDate = new Date(isoStr);
      } else {
        recordDate = new Date(dateTimeValue);
      }

      if (isNaN(recordDate.getTime())) continue;

      const recordYear = recordDate.getFullYear();
      const recordMonth = recordDate.getMonth() + 1;
      const recordHours = recordDate.getHours();
      const recordMinutes = recordDate.getMinutes();
      const recordTime = recordHours * 60 + recordMinutes; // 분 단위로 변환

      // 날짜만 비교하기 위한 Date 객체 (시간 제거)
      const recordDateOnly = new Date(
        recordYear,
        recordMonth - 1,
        recordDate.getDate(),
      );
      const currentWeekStartOnly = new Date(
        currentWeekStart.getFullYear(),
        currentWeekStart.getMonth(),
        currentWeekStart.getDate(),
      );
      const currentWeekEndOnly = new Date(
        currentWeekEnd.getFullYear(),
        currentWeekEnd.getMonth(),
        currentWeekEnd.getDate(),
      );

      // 월별 건수 (현재 년도와 월만 비교)
      if (recordYear == currentYear && recordMonth == currentMonth) {
        month++;
      }

      // 주별 건수 (날짜만 비교)
      if (
        recordDateOnly >= currentWeekStartOnly &&
        recordDateOnly <= currentWeekEndOnly
      ) {
        week++;
      }

      // 시간대별 건수
      // 아침 06:30-07:59 (06:30 포함, 07:59 포함)
      if (recordTime >= 390 && recordTime <= 479) {
        // 06:30 = 390분, 07:59 = 479분
        t1++;
      }
      // 오전 08:00-11:59 (08:00 포함, 11:59 포함)
      else if (recordTime >= 480 && recordTime <= 719) {
        // 08:00 = 480분, 11:59 = 719분
        t2++;
      }
      // 오후 13:00-17:00 (13:00 포함, 17:00 포함)
      else if (recordTime >= 780 && recordTime <= 1020) {
        // 13:00 = 780분, 17:00 = 1020분
        t3++;
      }
      // 저녁 18:00-23:00 (18:00 포함, 23:00 포함, 23:59까지 포함)
      else if (recordTime >= 1080 && recordTime <= 1439) {
        // 18:00 = 1080분, 23:59 = 1439분
        t4++;
      }
    }
  }

  return {
    total: total,
    week: week,
    month: month,
    t1: t1,
    t2: t2,
    t3: t3,
    t4: t4,
  };
}

// 주의 시작일(월요일) 구하기
function getWeekStart(date) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = d.getDate() - day + (day == 0 ? -6 : 1); // 월요일로 조정
  const weekStart = new Date(d.getFullYear(), d.getMonth(), diff);
  weekStart.setHours(0, 0, 0, 0);
  return weekStart;
}

// ========== 조회 관련 함수들 ==========

// 테스트 함수: 시트 정보 확인
function testGetSheetInfo() {
  try {
    const ss = SpreadsheetApp.openByUrl(url);
    const sheets = ss.getSheets();
    const sheetNames = sheets.map(function (sheet) {
      return sheet.getName();
    });

    return {
      success: true,
      sheetNames: sheetNames,
      url: url,
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
    };
  }
}

// 테스트 함수: 간단한 데이터 조회 테스트
function testGetData() {
  try {
    const data = getAllwalkData();
    return {
      success: true,
      count: data.length,
      sample: data.length > 0 ? data[0] : null,
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
    };
  }
}

// walk 데이터 전체 조회
function getAllwalkData() {
  try {
    Logger.log("getAllwalkData called");
    const ss = SpreadsheetApp.openByUrl(url);
    Logger.log("Spreadsheet opened successfully");

    let ws = null;

    // 시트 이름 시도 (대소문자 구분)
    // getSheetByName은 시트가 없으면 null을 반환 (예외를 던지지 않음)
    ws = ss.getSheetByName("walk");
    if (ws) {
      Logger.log("Found sheet: walk");
    } else {
      Logger.log("Sheet 'walk' not found, trying 'walk'");
      ws = ss.getSheetByName("walk");
      if (ws) {
        Logger.log("Found sheet: walk");
      } else {
        Logger.log("Sheet 'walk' also not found");
        // 모든 시트 이름 확인
        const allSheets = ss.getSheets();
        const sheetNames = allSheets.map(function (s) {
          return s.getName();
        });
        Logger.log("Available sheets: " + sheetNames.join(", "));

        // 대소문자 무시하고 찾기
        for (let i = 0; i < allSheets.length; i++) {
          const sheetName = allSheets[i].getName();
          if (sheetName.toLowerCase() === "walk") {
            ws = allSheets[i];
            Logger.log("Found sheet (case-insensitive): " + sheetName);
            break;
          }
        }

        if (!ws) {
          Logger.log("Could not find walk sheet with any case variation");
          return [];
        }
      }
    }

    const lastRow = ws.getLastRow();
    Logger.log("Last row: " + lastRow);

    if (lastRow <= 1) {
      Logger.log("No data rows (lastRow <= 1)");
      return [];
    }

    // 헤더 제외하고 모든 데이터 가져오기 (학번, 반, 구분, 이름, walk, 담당자, 기숙사, 퀀텀)
    // 날짜 필드를 텍스트로 읽기 위해 getDisplayValues 사용
    const data = ws.getRange(2, 1, lastRow - 1, 8).getDisplayValues();
    Logger.log("Data retrieved, rows: " + data.length);
    Logger.log("Last row in sheet: " + lastRow);

    // 빈 행 제거 (모든 셀이 비어있는 행)
    const nonEmptyRows = data.filter(function (row) {
      return row.some(function (cell) {
        return (
          cell !== null && cell !== undefined && String(cell).trim() !== ""
        );
      });
    });

    Logger.log("Non-empty rows: " + nonEmptyRows.length);

    if (nonEmptyRows.length > 0) {
      Logger.log("First row: " + JSON.stringify(nonEmptyRows[0]));
      Logger.log("All rows sample:");
      for (let i = 0; i < Math.min(nonEmptyRows.length, 5); i++) {
        Logger.log("Row " + (i + 2) + ": " + JSON.stringify(nonEmptyRows[i]));
      }
    }

    return nonEmptyRows;
  } catch (error) {
    Logger.log("getAllwalkData error: " + error.toString());
    Logger.log("Error stack: " + error.stack);
    return [];
  }
}

// 학생 목록 조회 (walk 시트에서)
function getStudentList() {
  const ss = SpreadsheetApp.openByUrl(url);
  let ws;

  // walk 시트 찾기 (대소문자 구분)
  ws = ss.getSheetByName("walk");
  if (!ws) {
    ws = ss.getSheetByName("walk");
    if (!ws) {
      // 대소문자 무시하고 찾기
      const allSheets = ss.getSheets();
      for (let i = 0; i < allSheets.length; i++) {
        const sheetName = allSheets[i].getName();
        if (sheetName.toLowerCase() === "walk") {
          ws = allSheets[i];
          break;
        }
      }
      if (!ws) {
        return [];
      }
    }
  }

  const lastRow = ws.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  // walk 시트 구조: 학번(컬럼1), 반(컬럼2), 구분(컬럼3), 이름(컬럼4), walk(컬럼5), 담당자(컬럼6)
  const data = ws.getRange(2, 1, lastRow - 1, 6).getValues();
  const studentMap = {}; // 학생 정보와 건수를 저장하는 맵 (학번 기준)

  for (let i = 0; i < data.length; i++) {
    const id = data[i][0] ? data[i][0].toString().trim() : "";
    const ban = data[i][1] ? data[i][1].toString().trim() : "";
    const name = data[i][3] ? data[i][3].toString().trim() : "";

    // 학번이 있는 경우
    if (id) {
      if (!studentMap[id]) {
        // 처음 등장하는 학생인 경우
        studentMap[id] = {
          id: id,
          name: name,
          ban: ban,
          count: 1, // 건수 초기화
        };
      } else {
        // 이미 등장한 학생인 경우 건수만 증가
        studentMap[id].count++;
      }
    }
  }

  // 맵을 배열로 변환
  const students = [];
  for (let id in studentMap) {
    students.push(studentMap[id]);
  }

  return students;
}

// 학생 검색 (코드 또는 이름으로)
function searchStudents(keyword) {
  if (!keyword || keyword.trim() === "") {
    return [];
  }

  const students = getStudentList();
  const searchKeyword = keyword.toLowerCase().trim();
  const results = [];

  for (let i = 0; i < students.length; i++) {
    const student = students[i];
    if (
      (student.id &&
        student.id.toString().toLowerCase().includes(searchKeyword)) ||
      (student.name && student.name.toLowerCase().includes(searchKeyword))
    ) {
      results.push(student);
    }
  }

  return results;
}

// walk 데이터 조회 (필터링)
function querywalkData(filters) {
  try {
    Logger.log(
      "querywalkData called with filters: " + JSON.stringify(filters),
    );

    const allData = getAllwalkData();

    // 디버깅용 로그
    Logger.log("Total data count: " + allData.length);
    Logger.log("Filters: " + JSON.stringify(filters));

    if (allData.length === 0) {
      Logger.log("No data found, returning empty result");
      return { count: 0, data: [] };
    }

    Logger.log("First row sample: " + JSON.stringify(allData[0]));
    Logger.log("All data rows: " + allData.length);

    let filteredData = allData;
    Logger.log("Before filtering: " + filteredData.length + " rows");

    // 학생 필터 (학번 또는 이름)
    if (filters.studentId && filters.studentId !== "") {
      const studentIdStr = String(filters.studentId).trim();
      Logger.log("Filtering by student ID: '" + studentIdStr + "'");
      Logger.log("Before student filter: " + filteredData.length + " rows");

      let matchCount = 0;
      filteredData = filteredData.filter(function (row) {
        const rowStudentId = row[0] ? String(row[0]).trim() : "";
        const match = rowStudentId === studentIdStr;
        if (match) {
          matchCount++;
          if (matchCount <= 3) {
            Logger.log(
              "Matched student row " + matchCount + ": " + JSON.stringify(row),
            );
          }
        }
        // 처음 몇 개만 로그
        if (!match && matchCount === 0) {
          Logger.log(
            "Sample non-matching row - studentId in data: '" +
              rowStudentId +
              "', looking for: '" +
              studentIdStr +
              "'",
          );
        }
        return match;
      });
      Logger.log(
        "After student filter: " +
          filteredData.length +
          " rows (matched: " +
          matchCount +
          ")",
      );
    }

    // 담당자 필터
    if (
      filters.manager &&
      filters.manager !== "" &&
      filters.manager !== "전체"
    ) {
      filteredData = filteredData.filter(function (row) {
        return row[5] && row[5].toString() === filters.manager; // 담당자 비교
      });
      Logger.log("After manager filter: " + filteredData.length + " rows");
    }

    // 날짜 필터
    if (filters.dateType && filters.dateType !== "전체") {
      Logger.log("Applying date filter: " + filters.dateType);
      const beforeDateFilter = filteredData.length;

      // 기간별인 경우 시작일과 종료일 미리 파싱
      let startDateOnly = null;
      let endDateOnly = null;

      if (
        filters.dateType === "기간별" &&
        filters.startDate &&
        filters.endDate
      ) {
        const startDateStr = String(filters.startDate).trim();
        const endDateStr = String(filters.endDate).trim();

        Logger.log(
          "Date filter - Start: " + startDateStr + ", End: " + endDateStr,
        );

        try {
          const startDateParts = startDateStr.split("-");
          const endDateParts = endDateStr.split("-");

          if (startDateParts.length === 3 && endDateParts.length === 3) {
            startDateOnly = new Date(
              parseInt(startDateParts[0], 10), // year
              parseInt(startDateParts[1], 10) - 1, // month (0-based)
              parseInt(startDateParts[2], 10), // day
            );
            startDateOnly.setHours(0, 0, 0, 0);

            endDateOnly = new Date(
              parseInt(endDateParts[0], 10), // year
              parseInt(endDateParts[1], 10) - 1, // month (0-based)
              parseInt(endDateParts[2], 10), // day
            );
            endDateOnly.setHours(23, 59, 59, 999);

            Logger.log(
              "Parsed dates - Start: " +
                startDateOnly.toISOString() +
                ", End: " +
                endDateOnly.toISOString(),
            );
          }
        } catch (e) {
          Logger.log("Error parsing date filter: " + e.toString());
        }
      }

      let processedCount = 0;
      filteredData = filteredData.filter(function (row) {
        processedCount++;
        const dateTimeValue = row[4];
        if (!dateTimeValue) {
          if (processedCount <= 5) {
            Logger.log(
              "Row " +
                processedCount +
                ": No date value, row=" +
                JSON.stringify(row),
            );
          }
          return false;
        }

        let recordDate;
        try {
          // 텍스트로 읽어오므로 항상 문자열로 처리
          let dateTimeStr = String(dateTimeValue || "").trim();

          if (!dateTimeStr || dateTimeStr === "") {
            if (processedCount <= 5) {
              Logger.log(
                "Row " +
                  processedCount +
                  ": Empty date string, row=" +
                  JSON.stringify(row),
              );
            }
            return false;
          }

          // 처음 몇 개만 로그
          if (processedCount <= 5) {
            Logger.log(
              "Row " +
                processedCount +
                ": Original date string: '" +
                dateTimeStr +
                "'",
            );
          }

          // 괄호나 특수문자 제거 (예: "2026-01-12()" -> "2026-01-12")
          dateTimeStr = dateTimeStr.replace(/[()]/g, "").trim();

          // 슬래시 형식도 처리 (예: "2026/01/12" -> "2026-01-12")
          if (dateTimeStr.includes("/")) {
            dateTimeStr = dateTimeStr.replace(/\//g, "-");
          }

          // 여러 형식 시도
          // "YYYY-MM-DD HH:mm:ss" 또는 "YYYY-MM-DD HH:mm" 형식
          if (dateTimeStr.includes(" ")) {
            const parts = dateTimeStr.split(" ");
            const datePart = parts[0];
            const timePart = parts[1] || "00:00:00";

            // 날짜 부분이 "YYYY-MM-DD" 형식인지 확인
            if (datePart.match(/^\d{4}-\d{2}-\d{2}$/)) {
              // 시간 부분이 "HH:mm:ss" 또는 "HH:mm" 형식인지 확인
              if (timePart.match(/^\d{2}:\d{2}(:\d{2})?$/)) {
                const isoStr =
                  datePart +
                  "T" +
                  (timePart.length === 5 ? timePart + ":00" : timePart);
                recordDate = new Date(isoStr);
              } else {
                recordDate = new Date(datePart + "T00:00:00");
              }
            } else {
              // 기타 형식 시도
              const isoStr = dateTimeStr.replace(" ", "T");
              recordDate = new Date(isoStr);
            }
          }
          // "YYYY-MM-DD" 형식
          else if (dateTimeStr.match(/^\d{4}-\d{2}-\d{2}$/)) {
            recordDate = new Date(dateTimeStr + "T00:00:00");
          }
          // 기타 형식
          else {
            recordDate = new Date(dateTimeStr);
          }

          if (isNaN(recordDate.getTime())) {
            // 날짜 파싱 실패 시 해당 행 제외
            if (processedCount <= 5) {
              Logger.log(
                "Row " +
                  processedCount +
                  ": Date parsing failed for: '" +
                  dateTimeStr +
                  "' (original: '" +
                  dateTimeValue +
                  "')",
              );
            }
            return false;
          }

          if (processedCount <= 5) {
            Logger.log(
              "Row " +
                processedCount +
                ": Successfully parsed date: '" +
                dateTimeStr +
                "' -> " +
                recordDate.toISOString(),
            );
          }
        } catch (e) {
          // 날짜 파싱 에러 시 해당 행 제외
          if (processedCount <= 5) {
            Logger.log(
              "Row " +
                processedCount +
                ": Date parsing error: " +
                e.toString() +
                " for value: '" +
                dateTimeValue +
                "'",
            );
          }
          return false;
        }

        // 기간별 필터 적용
        if (filters.dateType === "기간별") {
          if (startDateOnly && endDateOnly) {
            // 날짜 문자열에서 직접 년/월/일 추출 (타임존 문제 방지)
            let recordYear, recordMonth, recordDay;

            // 파싱된 날짜 문자열에서 년/월/일 추출 시도
            const dateTimeStr = String(dateTimeValue || "")
              .trim()
              .replace(/[()]/g, "")
              .trim();
            const dateMatch = dateTimeStr.match(
              /(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/,
            );

            if (dateMatch) {
              // 문자열에서 직접 추출
              recordYear = parseInt(dateMatch[1], 10);
              recordMonth = parseInt(dateMatch[2], 10);
              recordDay = parseInt(dateMatch[3], 10);
            } else {
              // 파싱 실패 시 Date 객체에서 추출 (fallback)
              recordYear = recordDate.getFullYear();
              recordMonth = recordDate.getMonth() + 1;
              recordDay = recordDate.getDate();
            }

            // 시작일과 종료일의 년/월/일 추출 (이미 파싱된 Date 객체에서)
            const startYear = startDateOnly.getFullYear();
            const startMonth = startDateOnly.getMonth() + 1;
            const startDay = startDateOnly.getDate();

            const endYear = endDateOnly.getFullYear();
            const endMonth = endDateOnly.getMonth() + 1;
            const endDay = endDateOnly.getDate();

            // 날짜 비교 (년/월/일을 숫자로 변환하여 비교)
            // 년월일을 YYYYMMDD 형식의 숫자로 변환
            const recordDateNum =
              recordYear * 10000 + recordMonth * 100 + recordDay;
            const startDateNum =
              startYear * 10000 + startMonth * 100 + startDay;
            const endDateNum = endYear * 10000 + endMonth * 100 + endDay;

            const isInRange =
              recordDateNum >= startDateNum && recordDateNum <= endDateNum;

            // 처음 몇 개만 상세 로그
            if (processedCount <= 5) {
              // padStart 함수 대신 직접 포맷팅
              const formatDate = function (year, month, day) {
                const m = month < 10 ? "0" + month : month;
                const d = day < 10 ? "0" + day : day;
                return year + "-" + m + "-" + d;
              };

              Logger.log(
                "Row " +
                  processedCount +
                  " Date comparison: " +
                  "record=" +
                  formatDate(recordYear, recordMonth, recordDay) +
                  " (num=" +
                  recordDateNum +
                  ", from string: '" +
                  dateTimeStr +
                  "'), " +
                  "range=[" +
                  formatDate(startYear, startMonth, startDay) +
                  " (num=" +
                  startDateNum +
                  ") ~ " +
                  formatDate(endYear, endMonth, endDay) +
                  " (num=" +
                  endDateNum +
                  ")], " +
                  "match=" +
                  isInRange +
                  ", studentId=" +
                  (row[0] || ""),
              );
            }

            return isInRange;
          } else {
            // 기간별인데 날짜가 없으면 필터링하지 않음 (전체 조회)
            if (processedCount === 1) {
              Logger.log(
                "기간별 선택되었지만 시작일 또는 종료일이 없음. 전체 조회로 처리.",
              );
            }
            return true;
          }
        }

        // 기간별이 아닌 경우 (현재는 "전체"만 있음)
        return true;
      });
      Logger.log(
        "After date filter: " +
          filteredData.length +
          " rows (was " +
          beforeDateFilter +
          ")",
      );
    } else {
      Logger.log("No date filter applied (dateType: " + filters.dateType + ")");
    }

    // 전체/기간별 조회 시 일자별 정렬 (최신순) — 타임스탬프로 비교해 정확히 적용
    try {
      filteredData.sort(function (a, b) {
        var dateA = parseDateTime(a[4]);
        var dateB = parseDateTime(b[4]);
        var tsA = dateA && !isNaN(dateA.getTime()) ? dateA.getTime() : 0;
        var tsB = dateB && !isNaN(dateB.getTime()) ? dateB.getTime() : 0;
        if (tsA === 0 && tsB === 0) return 0;
        if (tsA === 0) return 1;
        if (tsB === 0) return -1;
        return tsB - tsA;
      });
    } catch (e) {
      Logger.log("Sort function error: " + e.toString());
    }

    Logger.log("Filtered data count: " + filteredData.length);

    // 결과 객체 생성
    const result = {
      count: filteredData.length,
      data: filteredData,
    };

    Logger.log("Returning result with count: " + result.count);

    return result;
  } catch (error) {
    Logger.log("querywalkData error: " + error.toString());
    Logger.log("Error stack: " + (error.stack || "No stack"));
    return { count: 0, data: [], error: error.toString() };
  }
}

// ========== 순위별 조회 ==========
// rankCondition: 'all' | 'month_1'~'month_12' | 'week_2026_1_1' (년_월_주차)
function queryRankData(filters) {
  try {
    const allData = getAllwalkData();
    if (allData.length === 0) {
      return { isRank: true, rankData: [], rankLabel: "", count: 0 };
    }

    let filteredData = allData;

    // 담당자 필터
    if (
      filters.manager &&
      filters.manager !== "" &&
      filters.manager !== "전체"
    ) {
      filteredData = filteredData.filter(function (row) {
        return row[5] && row[5].toString() === filters.manager;
      });
    }

    const rankCondition = filters.rankCondition || "all";
    const currentYear = new Date().getFullYear();

    // 날짜 필터 적용 (rankCondition 기준)
    if (rankCondition !== "all") {
      if (rankCondition.indexOf("month_") === 0) {
        const month = parseInt(rankCondition.replace("month_", ""), 10);
        filteredData = filteredData.filter(function (row) {
          const dt = parseDateTime(row[4]);
          if (!dt || isNaN(dt.getTime())) return false;
          return (
            dt.getFullYear() === currentYear && dt.getMonth() + 1 === month
          );
        });
      } else if (rankCondition.indexOf("quarter_") === 0) {
        const quarter = parseInt(rankCondition.replace("quarter_", ""), 10);
        const startMonth = (quarter - 1) * 3 + 1;
        const endMonth = quarter * 3;
        filteredData = filteredData.filter(function (row) {
          const dt = parseDateTime(row[4]);
          if (!dt || isNaN(dt.getTime())) return false;
          const m = dt.getMonth() + 1;
          return (
            dt.getFullYear() === currentYear && m >= startMonth && m <= endMonth
          );
        });
      } else if (rankCondition.indexOf("week_") === 0) {
        const parts = rankCondition.replace("week_", "").split("_");
        const year = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10);
        const weekOfMonth = parseInt(parts[2], 10);
        const lastDayOfMonth = new Date(year, month, 0).getDate();
        const startDay = (weekOfMonth - 1) * 7 + 1;

        if (startDay > lastDayOfMonth) {
          filteredData = [];
        } else {
          const endDay = weekOfMonth * 7;
          const actualEndDay = Math.min(endDay, lastDayOfMonth);
          const startDate = new Date(year, month - 1, startDay);
          startDate.setHours(0, 0, 0, 0);
          const endDate = new Date(year, month - 1, actualEndDay);
          endDate.setHours(23, 59, 59, 999);

          filteredData = filteredData.filter(function (row) {
            const dt = parseDateTime(row[4]);
            if (!dt || isNaN(dt.getTime())) return false;
            return dt >= startDate && dt <= endDate;
          });
        }
      }
    }

    // 학생별 건수 집계
    const countMap = {};
    const infoMap = {};
    for (let i = 0; i < filteredData.length; i++) {
      const row = filteredData[i];
      const id = row[0] ? String(row[0]).trim() : "";
      if (!id) continue;

      countMap[id] = (countMap[id] || 0) + 1;
      // 학생별 기본 정보 + 기숙사·퀀텀 (여러 건이면 마지막 기록 값으로 표시용)
      infoMap[id] = {
        ban: row[1] || "",
        kbn: row[2] || "",
        name: row[3] || "",
        ksuk: row[6] || "",
        qtum: row[7] || "",
      };
    }

    // 건수 내림차순 정렬 후 상위 10명 (건수는 절대 생략하지 않음)
    const sorted = [];
    for (const id in countMap) {
      const info = infoMap[id];
      sorted.push({
        id: id,
        ban: info.ban,
        kbn: info.kbn,
        name: info.name,
        ksuk: info.ksuk,
        qtum: info.qtum,
        count: countMap[id],
      });
    }
    sorted.sort(function (a, b) {
      return b.count - a.count;
    });
    const top10 = sorted.slice(0, 10);

    // 순위 부여 (순위, 학번, 반, 구분, 이름, 기숙사, 퀀텀, 건수)
    const rankData = top10.map(function (item, idx) {
      return [
        idx + 1,
        item.id,
        item.ban,
        item.kbn,
        item.name,
        item.ksuk,
        item.qtum,
        item.count,
      ];
    });

    let rankLabel = "";
    if (rankCondition === "all") {
      rankLabel = "전체 누적 1~10순위";
    } else if (rankCondition.indexOf("month_") === 0) {
      const month = parseInt(rankCondition.replace("month_", ""), 10);
      rankLabel = currentYear + "년 " + month + "월 1~10순위";
    } else if (rankCondition.indexOf("quarter_") === 0) {
      const quarter = parseInt(rankCondition.replace("quarter_", ""), 10);
      const labels = ["", "1~3월", "4~6월", "7~9월", "10~12월"];
      rankLabel = currentYear + "년 " + labels[quarter] + " 분기 1~10순위";
    } else if (rankCondition.indexOf("week_") === 0) {
      const parts = rankCondition.replace("week_", "").split("_");
      rankLabel =
        parts[0] + "년 " + parts[1] + "월 " + parts[2] + "주차 1~10순위";
    }

    return {
      isRank: true,
      rankData: rankData,
      rankLabel: rankLabel,
      count: rankData.length,
    };
  } catch (error) {
    Logger.log("queryRankData error: " + error.toString());
    return {
      isRank: true,
      rankData: [],
      rankLabel: "",
      count: 0,
      error: error.toString(),
    };
  }
}

// 날짜/시간 파싱 (시트 표시 형식·로케일 다양 대응, 일자별 정렬용)
function parseDateTime(val) {
  if (!val) return null;
  if (val instanceof Date) return new Date(val.getTime());
  var str = String(val).trim().replace(/[()]/g, "").replace(/\//g, "-");
  str = str.replace(/\s*오전\s*/gi, " ").replace(/\s*오후\s*/gi, " ");
  var datePart = null;
  var timePart = "00:00:00";
  var match = str.match(/(\d{4})[-\s.]*(\d{1,2})[-\s.]*(\d{1,2})/);
  if (match) {
    var y = match[1];
    var m = ("0" + parseInt(match[2], 10)).slice(-2);
    var d = ("0" + parseInt(match[3], 10)).slice(-2);
    datePart = y + "-" + m + "-" + d;
  }
  var timeMatch = str.match(/(\d{1,2})\s*:\s*(\d{2})(?:\s*:\s*(\d{2}))?/);
  if (timeMatch) {
    var h = ("0" + parseInt(timeMatch[1], 10)).slice(-2);
    var min = timeMatch[2];
    var sec = timeMatch[3] !== undefined ? timeMatch[3] : "00";
    timePart = h + ":" + min + ":" + sec;
  }
  if (datePart) {
    var iso = datePart + "T" + timePart;
    var dt = new Date(iso);
    if (!isNaN(dt.getTime())) return dt;
  }
  if (str.indexOf("-") !== -1 && str.indexOf(" ") !== -1) {
    var fallback = new Date(str.replace(" ", "T"));
    if (!isNaN(fallback.getTime())) return fallback;
  }
  if (str.match(/^\d{4}-\d{2}-\d{2}$/)) {
    return new Date(str + "T00:00:00");
  }
  return new Date(str);
}

// 차트용 데이터: 동일 필터 적용 후 막대(상위 10명)·라인(일별 건수) 집계
function getChartData(filters) {
  try {
    let data = getAllwalkData();
    if (data.length === 0) {
      return { barData: [], lineData: [] };
    }

    if (filters.studentId && filters.studentId !== "") {
      const sid = String(filters.studentId).trim();
      data = data.filter(function (row) {
        return (row[0] ? String(row[0]).trim() : "") === sid;
      });
    }
    if (
      filters.manager &&
      filters.manager !== "" &&
      filters.manager !== "전체"
    ) {
      data = data.filter(function (row) {
        return row[5] && row[5].toString() === filters.manager;
      });
    }

    const currentYear = new Date().getFullYear();
    const rankCondition = filters.rankCondition || "";
    const dateType = filters.dateType || "전체";

    if (dateType === "기간별" && filters.startDate && filters.endDate) {
      const startParts = String(filters.startDate).trim().split("-");
      const endParts = String(filters.endDate).trim().split("-");
      if (startParts.length === 3 && endParts.length === 3) {
        const start = new Date(
          parseInt(startParts[0], 10),
          parseInt(startParts[1], 10) - 1,
          parseInt(startParts[2], 10),
        );
        const end = new Date(
          parseInt(endParts[0], 10),
          parseInt(endParts[1], 10) - 1,
          parseInt(endParts[2], 10),
        );
        end.setHours(23, 59, 59, 999);
        data = data.filter(function (row) {
          const dt = parseDateTime(row[4]);
          return dt && !isNaN(dt.getTime()) && dt >= start && dt <= end;
        });
      }
    } else if (
      dateType === "순위별" &&
      rankCondition &&
      rankCondition !== "all"
    ) {
      if (rankCondition.indexOf("month_") === 0) {
        const month = parseInt(rankCondition.replace("month_", ""), 10);
        data = data.filter(function (row) {
          const dt = parseDateTime(row[4]);
          return (
            dt &&
            !isNaN(dt.getTime()) &&
            dt.getFullYear() === currentYear &&
            dt.getMonth() + 1 === month
          );
        });
      } else if (rankCondition.indexOf("quarter_") === 0) {
        const q = parseInt(rankCondition.replace("quarter_", ""), 10);
        const startM = (q - 1) * 3 + 1;
        const endM = q * 3;
        data = data.filter(function (row) {
          const dt = parseDateTime(row[4]);
          if (!dt || isNaN(dt.getTime())) return false;
          const m = dt.getMonth() + 1;
          return dt.getFullYear() === currentYear && m >= startM && m <= endM;
        });
      } else if (rankCondition.indexOf("week_") === 0) {
        const parts = rankCondition.replace("week_", "").split("_");
        const y = parseInt(parts[0], 10);
        const m = parseInt(parts[1], 10);
        const w = parseInt(parts[2], 10);
        const lastDay = new Date(y, m, 0).getDate();
        const startDay = (w - 1) * 7 + 1;
        if (startDay <= lastDay) {
          const start = new Date(y, m - 1, startDay);
          start.setHours(0, 0, 0, 0);
          const end = new Date(y, m - 1, Math.min(w * 7, lastDay));
          end.setHours(23, 59, 59, 999);
          data = data.filter(function (row) {
            const dt = parseDateTime(row[4]);
            return dt && !isNaN(dt.getTime()) && dt >= start && dt <= end;
          });
        } else {
          data = [];
        }
      }
    }

    var barMap = {};
    var nameMap = {};
    var lineMap = {};
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var id = row[0] ? String(row[0]).trim() : "";
      var name = row[3] ? String(row[3]).trim() : "";
      if (id) {
        barMap[id] = (barMap[id] || 0) + 1;
        nameMap[id] = name || id;
      }
      var dt = parseDateTime(row[4]);
      if (dt && !isNaN(dt.getTime())) {
        var mo = dt.getMonth() + 1;
        var da = dt.getDate();
        var dateStr =
          dt.getFullYear() +
          "-" +
          (mo < 10 ? "0" + mo : mo) +
          "-" +
          (da < 10 ? "0" + da : da);
        lineMap[dateStr] = (lineMap[dateStr] || 0) + 1;
      }
    }

    var barList = [];
    for (var id in barMap) {
      barList.push({
        label: nameMap[id] || id,
        value: barMap[id],
      });
    }
    barList.sort(function (a, b) {
      return b.value - a.value;
    });
    var barData = barList.slice(0, 10);

    var lineList = [];
    for (var dateStr in lineMap) {
      lineList.push({ date: dateStr, count: lineMap[dateStr] });
    }
    lineList.sort(function (a, b) {
      return a.date.localeCompare(b.date);
    });

    return { barData: barData, lineData: lineList };
  } catch (e) {
    Logger.log("getChartData error: " + e.toString());
    return { barData: [], lineData: [] };
  }
}
