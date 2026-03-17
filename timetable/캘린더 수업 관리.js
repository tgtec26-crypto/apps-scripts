/**
 * 🎯 구글 스프레드시트 커스텀 메뉴 생성
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📚 수업 관리')
    .addItem('📅 날짜별 수업 정보 생성', 'generateDailyClassInfoUI')
    .addSeparator()
    .addItem('📝 구글 캘린더에 수업 등록', 'registerDailyClassEventsWithTime')
    .addSeparator()
    .addItem('🗑️ 구글 캘린더 학기 수업 전체 삭제', 'deleteSemesterCalendarEvents')
    .addToUi();
}

/**
 * 📌 1-A. 교사 선택 드롭다운 UI 띄우기
 */
function generateDailyClassInfoUI() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scheduleSheet = ss.getSheetByName("주간시간표");
  
  if (!scheduleSheet) {
    Browser.msgBox("오류", "'주간시간표' 시트를 찾을 수 없습니다.", Browser.Buttons.OK);
    return;
  }

  // 주간시간표 A열에서 교사 이름 추출
  var data = scheduleSheet.getRange("A:A").getValues();
  var teacherSet = new Set();
  
  for (var i = 0; i < data.length; i++) {
    var val = String(data[i][0]).trim();
    if (val && !val.match(/교사|이름|시간표/)) {
      var cleanName = val.split('(')[0].trim(); 
      if (cleanName) {
        teacherSet.add(cleanName);
      }
    }
  }
  
  var teacherList = Array.from(teacherSet);
  if (teacherList.length === 0) {
    Browser.msgBox("알림", "'주간시간표' A열에서 교사 명단을 찾을 수 없습니다.", Browser.Buttons.OK);
    return;
  }

  var userProperties = PropertiesService.getUserProperties();
  var lastTeacher = userProperties.getProperty('LAST_TEACHER') || "";

  var htmlString = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: 'Malgun Gothic', sans-serif; padding: 20px; color: #333; }
        label { font-size: 14px; font-weight: bold; margin-bottom: 10px; display: block; }
        select { width: 100%; padding: 10px; font-size: 15px; border: 1px solid #ccc; border-radius: 5px; margin-bottom: 20px; outline: none; }
        select:focus { border-color: #4CAF50; }
        button { width: 100%; padding: 12px; background-color: #1a73e8; color: white; border: none; font-size: 15px; font-weight: bold; border-radius: 5px; cursor: pointer; transition: background 0.3s; }
        button:hover { background-color: #1557b0; }
        button:disabled { background-color: #ccc; cursor: not-allowed; }
      </style>
    </head>
    <body>
      <label for="teacher">👨‍🏫 수업 정보를 생성할 교사를 선택하세요:</label>
      <select id="teacher"></select>
      <button id="btn" onclick="submit()">수업 정보 생성하기</button>

      <script>
        var teachers = ${JSON.stringify(teacherList)};
        var lastTeacher = "${lastTeacher}";

        window.onload = function() {
          var select = document.getElementById('teacher');
          teachers.forEach(function(t) {
            var opt = document.createElement('option');
            opt.value = t;
            opt.text = t;
            if (t === lastTeacher) opt.selected = true;
            select.appendChild(opt);
          });
        };

        function submit() {
          var teacher = document.getElementById('teacher').value;
          var btn = document.getElementById('btn');
          btn.disabled = true;
          btn.innerText = "데이터 처리 중...";

          google.script.run
            .withSuccessHandler(function() {
              google.script.host.close();
            })
            .withFailureHandler(function(err) {
              alert("오류 발생: " + err.message);
              btn.disabled = false;
              btn.innerText = "수업 정보 생성하기";
            })
            .processDailyClassInfo(teacher);
        }
      </script>
    </body>
    </html>
  `;

  var htmlOutput = HtmlService.createHtmlOutput(htmlString)
    .setWidth(350)
    .setHeight(220);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '교사 선택');
}

/**
 * 📌 1-B. 실제 '날짜별 수업 정보' 생성 
 */
function processDailyClassInfo(teacherName) {
  PropertiesService.getUserProperties().setProperty('LAST_TEACHER', teacherName);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast(teacherName + " 교사의 수업 정보를 생성 중입니다...", "⏳ 처리 중", -1);

  var settingsSheet = ss.getSheetByName("설정");
  var scheduleSheet = ss.getSheetByName("주간시간표");
  var timeSheet = ss.getSheetByName("교시별 시간");
  
  var dailySheet = ss.getSheetByName("날짜별 수업 정보");
  if (!dailySheet) dailySheet = ss.insertSheet("날짜별 수업 정보");
  
  if (!settingsSheet || !scheduleSheet || !timeSheet) {
    throw new Error("필수 시트(설정, 주간시간표, 교시별 시간)가 누락되었습니다.");
  }

  // ✅ 학년+과목별 주제 데이터 수집 ('{숫자}학년-{과목명}' 탭 자동 탐색)
  var topicMap = {};
  ss.getSheets().forEach(function(sheet) {
    var match = sheet.getName().match(/^(\d)학년-(.+)$/);
    if (match) {
      var g = parseInt(match[1], 10);
      var subj = match[2].trim();
      if (!topicMap[g]) topicMap[g] = {};
      if (!topicMap[g][subj]) topicMap[g][subj] = {};
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) { // 첫 줄(헤더) 제외
        var round = parseInt(data[i][1], 10); // B열: 회차
        var topic = data[i][2] ? String(data[i][2]).trim() : ""; // C열: 주제
        if (!isNaN(round) && topic) {
          topicMap[g][subj][round] = topic;
        }
      }
    }
  });

  dailySheet.clear();
  // ✅ B열 헤더를 '휴일 또는 학사일정'으로 변경
  dailySheet.appendRow(["날짜", "휴일 또는 학사일정", "학년-반", "시작 시각", "끝나는 시각", "교시", "회차", "주제"]);
  
  function getSafeDateStr(val) {
    if (!val) return null;
    if (val instanceof Date) return Utilities.formatDate(val, "Asia/Seoul", "yyyy-MM-dd");
    var s = String(val).trim();
    var match = s.match(/^(\d{4})[^0-9]+(\d{1,2})[^0-9]+(\d{1,2})/);
    if (match) return match[1] + "-" + ("0" + match[2]).slice(-2) + "-" + ("0" + match[3]).slice(-2);
    try { return Utilities.formatDate(new Date(val), "Asia/Seoul", "yyyy-MM-dd"); } catch(e) { return null; }
  }
  
  try {
    var startDate = new Date(settingsSheet.getRange("A2").getValue());
    var endDate = new Date(settingsSheet.getRange("B2").getValue());
    
    var holidayMap = {};
    var holidayData = settingsSheet.getRange("C2:D" + settingsSheet.getLastRow()).getValues();
    holidayData.forEach(row => {
      var d = getSafeDateStr(row[0]);
      if (d) holidayMap[d] = row[1] || "학교 휴업";
    });

    var closureMap = {};
    var closureData = settingsSheet.getRange("E2:H" + settingsSheet.getLastRow()).getValues();
    
    closureData.forEach(row => {
      var d = getSafeDateStr(row[0]); 
      if (d) {
        var gradeVal = row[1];
        var closureGrades = [];
        if (gradeVal !== "" && gradeVal !== null && gradeVal !== undefined) {
          var gStr = String(gradeVal).replace(/[^0-9,]/g, ""); 
          closureGrades = gStr ? gStr.split(',').map(g => parseInt(g, 10)).filter(g => !isNaN(g)) : [];
        }

        var pStr = String(row[2]).replace(/[^0-9,]/g, ""); 
        var periods = pStr ? pStr.split(',').map(p => parseInt(p, 10)).filter(p => !isNaN(p)) : [];
        var reason = row[3] || "일부 교시 휴업";

        if (!closureMap[d]) closureMap[d] = {}; 
        periods.forEach(p => {
          if (!closureMap[d][p]) closureMap[d][p] = [];
          closureMap[d][p].push({ reason: reason, grades: closureGrades });
        });
      }
    });

    var dayChangeMap = {};
    var changeData = settingsSheet.getRange("I2:J" + settingsSheet.getLastRow()).getValues();
    changeData.forEach(row => {
      var d = getSafeDateStr(row[0]);
      if (d && row[1]) dayChangeMap[d] = String(row[1]).trim();
    });

    var timeData = timeSheet.getDataRange().getValues();
    var scheduleData = scheduleSheet.getDataRange().getValues();
    
    var teacherRow;
    for (var i = 0; i < scheduleData.length; i++) {
      if (scheduleData[i][0]) {
        var cellName = String(scheduleData[i][0]).split('(')[0].trim();
        if (cellName === teacherName) {
          teacherRow = scheduleData[i];
          break;
        }
      }
    }
    
    if (!teacherRow) {
      throw new Error("해당 교사를 찾을 수 없습니다.");
    }

    var dayColumns = { "월": [1, 6], "화": [7, 13], "수": [14, 19], "목": [20, 26], "금": [27, 32] };
    var dayNameToNumber = { "월": 1, "화": 2, "수": 3, "목": 4, "금": 5 };
    var weekDays = ["일", "월", "화", "수", "목", "금", "토"];
    var classCounter = {};
    var dataToAppend = [];

    for (var date = new Date(startDate); date <= endDate; date.setDate(date.getDate() + 1)) {
      var dateStr = Utilities.formatDate(date, "Asia/Seoul", "yyyy-MM-dd");
      var dayOfWeek = date.getDay();
      
      var changedDay = dayChangeMap[dateStr];
      var actualDayName = weekDays[dayOfWeek];
      var scheduleDayName = changedDay || actualDayName;
      var displayDateStr = dateStr + `(${actualDayName}` + (changedDay ? `→${scheduleDayName}` : "") + `)`;

      if ((dayOfWeek === 0 || dayOfWeek === 6) && !changedDay) continue;

      if (holidayMap[dateStr]) {
        dataToAppend.push([displayDateStr, holidayMap[dateStr], "", "", "", "", "", ""]);
        continue;
      }

      var dayInfo = dayColumns[scheduleDayName];
      if (!dayInfo) continue;
      
      var [startCol, endCol] = dayInfo;
      var scheduleDayNum = dayNameToNumber[scheduleDayName];
      var maxPeriods = (scheduleDayNum === 2 || scheduleDayNum === 4) ? 7 : 6;
      var todayClosures = closureMap[dateStr]; 

      for (var period = 1; period <= maxPeriods; period++) {
        var colIndex = startCol + period - 1;
        if (colIndex >= teacherRow.length) continue;
        
        var classInfo = teacherRow[colIndex];
        if (!classInfo || String(classInfo).trim() === "") continue;

        var parts = String(classInfo).split(/\r?\n/);
        var gradeClass = parts.length > 0 ? "'" + parts[0].trim() : "";
        var subject = parts.length > 1 ? parts[1].trim() : "";

        if (subject.includes("적응 교육")) continue;

        var classGradeMatch = String(gradeClass).match(/[1-9]/);
        var classGrade = classGradeMatch ? parseInt(classGradeMatch[0], 10) : null;
        var specificReason = null;
        
        if (todayClosures && todayClosures[period]) {
          var closureEntries = todayClosures[period]; 
          for (var ci = 0; ci < closureEntries.length; ci++) {
            var entry = closureEntries[ci];
            if (entry.grades.length === 0) {
              specificReason = entry.reason;
              break;
            } else if (classGrade !== null && entry.grades.includes(classGrade)) {
              specificReason = entry.reason;
              break;
            }
          }
        }

        if (specificReason) {
          dataToAppend.push([
            displayDateStr, specificReason, gradeClass, 
            timeData[period] ? timeData[period][1] : "", 
            timeData[period] ? timeData[period][2] : "", 
            period, "", ""
          ]);
        } else {
          var classKey = `${gradeClass}-${subject}`;
          if (!classCounter[classKey]) classCounter[classKey] = 1;
          else classCounter[classKey]++;
          
          var currentRound = classCounter[classKey];
          var topic = (classGrade && subject && topicMap[classGrade] && topicMap[classGrade][subject] && topicMap[classGrade][subject][currentRound]) ? topicMap[classGrade][subject][currentRound] : "";
          
          var changedDayNote = changedDay ? `(${actualDayName}→${scheduleDayName} 변경)` : "";
          var finalTopic = topic;
          if (changedDayNote) {
            finalTopic = finalTopic ? finalTopic + " " + changedDayNote : changedDayNote;
          }

          dataToAppend.push([
            displayDateStr, subject, gradeClass, 
            timeData[period] ? timeData[period][1] : "", 
            timeData[period] ? timeData[period][2] : "", 
            period, currentRound, finalTopic
          ]);
        }
      }
    }

    if (dataToAppend.length > 0) {
      dailySheet.getRange(dailySheet.getLastRow() + 1, 1, dataToAppend.length, 8).setValues(dataToAppend);
    }
    
    ss.toast("✅ 주제가 포함된 수업 정보가 성공적으로 생성되었습니다.", "완료", 5);
    
  } catch (error) {
    throw new Error(error.message);
  }
}

/**
 * 📌 2. '날짜별 수업 정보' 시트를 기반으로 Google 캘린더에 이벤트 등록
 */
/**
 * 📌 2-A. 캘린더 업로드 대상 선택 팝업창 띄우기 (기존 함수 대체)
 */
function registerDailyClassEventsWithTime() {
  var htmlString = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: 'Malgun Gothic', sans-serif; padding: 20px; color: #333; }
        h3 { margin-top: 0; font-size: 16px; color: #1a73e8; }
        button { width: 100%; padding: 12px; font-size: 14px; font-weight: bold; border: none; border-radius: 5px; cursor: pointer; transition: background 0.3s; margin-bottom: 10px; }
        .btn-default { background-color: #1a73e8; color: white; }
        .btn-default:hover { background-color: #1557b0; }
        .btn-other { background-color: #34a853; color: white; }
        .btn-other:hover { background-color: #2b803e; }
        button:disabled { background-color: #ccc; cursor: not-allowed; }
        input { width: 92%; padding: 10px; font-size: 14px; border: 1px solid #ccc; border-radius: 5px; margin-bottom: 10px; outline: none; }
        input:focus { border-color: #34a853; }
        .divider { text-align: center; margin: 15px 0; color: #888; font-size: 12px; position: relative; }
        .divider::before, .divider::after { content: ''; position: absolute; top: 50%; width: 40%; height: 1px; background-color: #ddd; }
        .divider::before { left: 0; }
        .divider::after { right: 0; }
      </style>
    </head>
    <body>
      <h3>📅 일정을 등록할 캘린더 선택</h3>
      <button class="btn-default" id="btnDefault" onclick="useDefault()">지금 로그인한 계정에 업로드</button>
      
      <div class="divider">또는</div>
      
      <label style="font-size:13px; font-weight:bold; display:block; margin-bottom:5px;">다른 계정 캘린더 (이메일 입력):</label>
      <input type="text" id="otherEmail" placeholder="예: other_account@gmail.com">
      <button class="btn-other" id="btnOther" onclick="useOther()">입력한 계정에 업로드</button>

      <script>
        function useDefault() {
          disableButtons();
          google.script.run
            .withSuccessHandler(closeWindow)
            .withFailureHandler(showError)
            .processCalendarRegistration('default');
        }

        function useOther() {
          var email = document.getElementById('otherEmail').value.trim();
          if(!email) { 
            alert('업로드할 다른 계정의 이메일 주소를 입력해주세요!'); 
            return; 
          }
          disableButtons();
          google.script.run
            .withSuccessHandler(closeWindow)
            .withFailureHandler(showError)
            .processCalendarRegistration(email);
        }

        function disableButtons() {
          document.getElementById('btnDefault').disabled = true;
          document.getElementById('btnOther').disabled = true;
          document.getElementById('btnDefault').innerText = "처리 중...";
          document.getElementById('btnOther').innerText = "처리 중...";
        }

        function closeWindow() { 
          google.script.host.close(); 
        }

        function showError(err) { 
          alert("오류 발생: " + err.message + "\\n(접근 권한이 없거나 없는 이메일일 수 있습니다.)"); 
          google.script.host.close();
        }
      </script>
    </body>
    </html>
  `;

  var htmlOutput = HtmlService.createHtmlOutput(htmlString)
    .setWidth(350)
    .setHeight(300);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '캘린더 업로드');
}

/**
 * 📌 2-B. 전달받은 캘린더 정보로 실제 이벤트 등록 처리 (백그라운드 실행)
 */
function processCalendarRegistration(calendarId) {
  var calendar;
  
  // 전달받은 값이 'default'면 내 캘린더, 이메일이면 해당 캘린더를 불러옴
  if (calendarId === 'default') {
    calendar = CalendarApp.getDefaultCalendar();
  } else {
    calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      throw new Error("해당 캘린더를 찾을 수 없거나 공유 권한이 없습니다: " + calendarId);
    }
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dailySheet = ss.getSheetByName("날짜별 수업 정보");

  if (!dailySheet) {
    throw new Error("'날짜별 수업 정보' 시트를 찾을 수 없습니다.");
  }

  var data = dailySheet.getDataRange().getValues();
  var displayData = dailySheet.getDataRange().getDisplayValues(); 

  if (data.length < 2) {
    throw new Error("처리할 데이터가 없습니다.");
  }

  function parseDisplayTime(timeStr) {
    var str = String(timeStr).trim();
    var isPM = str.indexOf("오후") !== -1 || str.indexOf("PM") !== -1 || str.indexOf("pm") !== -1;
    var isAM = str.indexOf("오전") !== -1 || str.indexOf("AM") !== -1 || str.indexOf("am") !== -1;
    
    var cleanStr = str.replace(/오전|오후|AM|PM|am|pm/g, "").trim();
    var parts = cleanStr.split(":");
    var h = parseInt(parts[0], 10) || 0;
    var m = parseInt(parts[1], 10) || 0;
    
    if (isPM && h < 12) h += 12; 
    if (isAM && h === 12) h = 0; 
    
    return [h, m];
  }

  var minDate = new Date(data[1][0]);
  var maxDate = new Date(data[1][0]);
  
  for (var i = 1; i < data.length; i++) {
    var currentDate = new Date(data[i][0]);
    if (currentDate < minDate) minDate = currentDate;
    if (currentDate > maxDate) maxDate = currentDate;
  }
  
  minDate.setHours(0, 0, 0, 0);
  maxDate.setHours(23, 59, 59, 999);
  
  ss.toast("기존 일정을 삭제하고 새로운 일정을 등록하는 중입니다...", "⏳ 캘린더 처리 중", -1);
  var deletedCount = deleteExistingEvents(calendar, minDate, maxDate);
  Utilities.sleep(2000);
  
  var batchSize = 10;
  var waitTime = 3000;
  var processedCount = 0;
  var dateGroups = {};

  for (var i = 1; i < data.length; i++) {
    var dateStr = Utilities.formatDate(new Date(data[i][0]), "Asia/Seoul", "yyyy-MM-dd");
    if (!dateGroups[dateStr]) dateGroups[dateStr] = [];
    dateGroups[dateStr].push({ rawRow: data[i], displayRow: displayData[i] });
  }

  for (var dateStr in dateGroups) {
    var dateData = dateGroups[dateStr];
    for (var i = 0; i < dateData.length; i++) {
      var rowInfo = dateData[i];
      var row = rowInfo.rawRow;
      var displayRow = rowInfo.displayRow;

      var date = new Date(row[0]);
      var subject = row[1]; // B열
      var gradeClass = row[2] ? row[2].toString().trim() : "";
      var period = row[5];
      var round = row[6] ? row[6].toString().trim() : "";
      var topic = row[7] ? row[7].toString().trim() : "";

      var startTimeStr = displayRow[3];
      var endTimeStr = displayRow[4];

      if (!startTimeStr || !endTimeStr) continue;

      try {
        var parsedStart = parseDisplayTime(startTimeStr);
        var startHour = parsedStart[0];
        var startMinute = parsedStart[1];
        
        var parsedEnd = parseDisplayTime(endTimeStr);
        var endHour = parsedEnd[0];
        var endMinute = parsedEnd[1];

        var startDateTime = new Date(date);
        var endDateTime = new Date(date);

        startDateTime.setHours(startHour, startMinute, 0, 0);
        endDateTime.setHours(endHour, endMinute, 0, 0);

        if (startDateTime >= endDateTime) continue;

        var eventTitle;
        if (!round) { 
          eventTitle = `${gradeClass}-${period}교시-${subject}`;
        } else {
          eventTitle = `${gradeClass}-${period}교시-${round}차시`;
          if (topic) {
            eventTitle += `-${topic}`;
          }
        }

        calendar.createEvent(eventTitle, startDateTime, endDateTime, {description: '#수업관리스크립트'});
        processedCount++;
        Utilities.sleep(300);

        if (processedCount % batchSize === 0) {
          Utilities.sleep(waitTime);
        }
      } catch (rowError) {
        Logger.log(`⚠️ 행 처리 중 오류 발생: ${rowError.message}`);
      }
    }
  }

  ss.toast(`선택한 캘린더에서 기존 ${deletedCount}개 삭제 후\n새로 ${processedCount}개 일정 등록 완료!`, "✅ 완료", 5);
}

/**
 * 📌 3. 구글 캘린더 학기 수업 전체 삭제
 */
function deleteSemesterCalendarEvents() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('경고', '설정 시트의 [수업 생성 시작일 ~ 종료일] 기간 내의 모든 캘린더 일정을 삭제하시겠습니까?', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = ss.getSheetByName("설정");

    if (!settingsSheet) {
      ui.alert("오류", "설정 시트를 찾을 수 없습니다.", ui.ButtonSet.OK);
      return;
    }

    var startDate = new Date(settingsSheet.getRange("A2").getValue());
    var endDate = new Date(settingsSheet.getRange("B2").getValue());
    endDate.setHours(23, 59, 59, 999);

    var calendar = CalendarApp.getDefaultCalendar();

    ss.toast("일정 삭제를 시작합니다. 이벤트 양에 따라 1~2분 정도 소요될 수 있습니다.", "삭제 중...", 10);
    var deletedCount = deleteExistingEvents(calendar, startDate, endDate);

    ui.alert('완료', `총 ${deletedCount}개의 캘린더 일정이 성공적으로 삭제되었습니다.`, ui.ButtonSet.OK);
  }
}

/**
 * 🛠️ 공통 헬퍼 함수: 지정된 기간의 모든 캘린더 이벤트 삭제
 */
function deleteExistingEvents(calendar, startDate, endDate) {
  var deletedCount = 0;
  try {
    var currentDate = new Date(startDate);
    var intervalDays = 7; 
    
    while (currentDate <= endDate) {
      var nextDate = new Date(currentDate);
      nextDate.setDate(nextDate.getDate() + intervalDays);
      if (nextDate > endDate) nextDate = new Date(endDate);
      
      var events = calendar.getEvents(currentDate, nextDate);
      
      for (var i = 0; i < events.length; i++) {
        try {
          if (events[i].getDescription().indexOf('#수업관리스크립트') === -1) continue;
          events[i].deleteEvent();
          deletedCount++;
          if (deletedCount % 10 === 0) Utilities.sleep(500);
        } catch (deleteError) {
          Logger.log(`⚠️ 이벤트 삭제 실패: ${deleteError.message}`);
        }
      }
      currentDate.setTime(nextDate.getTime() + 1);
      if (currentDate <= endDate) Utilities.sleep(1000);
    }
  } catch (error) {
    Logger.log(`❌ 삭제 중 오류: ${error.message}`);
  }
  return deletedCount;
}