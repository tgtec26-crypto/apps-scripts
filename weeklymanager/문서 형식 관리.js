// ======================== [0] 전역 상수 ========================
const DOC_ID = "1FFPBrZb56rxKeNWw6z9fmnpjxu7ejy7oaC7XEUwGr1M";
// 개강일 설정 (매년 변경 필요, month는 0-indexed: 2=3월, 7=8월)
const SEM1_START_MONTH = 2;
const SEM1_START_DAY = 4;
const SEM2_START_MONTH = 7;
const SEM2_START_DAY = 18;

// ======================== [1] 메뉴 및 트리거 설정 ========================
function onOpen() {
  const ui = DocumentApp.getUi();
  ui.createMenu('추가 기능')
    .addItem('🚀 모든 섹션 일괄 업데이트', 'runAutomatedTask')
    .addSeparator()
    .addItem('숫자 매기기 (부서별 전달 사항)', 'runEmojiNumbering')
    .addItem('✿ 말머리 추가 (협의/논의 사항)', 'runFlowerPrefix')
    .addItem('일정 가져오기 (주간 일정)', 'runImportEvents')
    .addSeparator()
    .addItem('📅 새 주차 만들기', 'runCreateNewWeek')
    .addSeparator()
    .addItem('⏰ 자동 트리거 등록 (1회 실행)', 'setupTriggers')
    .addToUi();
}
function setupTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();
  allTriggers.forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('runAutomatedTask').timeBased().everyDays(1).atHour(7).create();
  ScriptApp.newTrigger('runAutomatedTask').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(12).create();
  DocumentApp.getUi().alert("트리거가 등록되었습니다.");
}
// ======================== [2] 상위 5개 탭 가져오기 유틸리티 ========================
function getTopFiveTabs() {
  const doc = DocumentApp.openById(DOC_ID);
  const tabs = doc.getTabs();
  const result = [];
  for (let i = 0; i < tabs.length; i++) {
    if (tabs[i].getTitle() === '주간계획 틀') continue;
    result.push({
      tab: tabs[i],
      body: tabs[i].asDocumentTab().getBody(),
      header: tabs[i].asDocumentTab().getHeader(),
      title: tabs[i].getTitle()
    });
    if (result.length >= 5) break;
  }
  return result;
}
// ======================== [3] 메뉴에서 호출되는 래퍼 함수들 ========================
function runAutomatedTask() {
  const tabsData = getTopFiveTabs();
  tabsData.forEach((tabData, i) => {
    importEventsWithBody(tabData.body);
    addEmojiNumberingWithBody(tabData.body);
    addFlowerPrefixWithBody(tabData.body);
    console.log(`탭 ${i + 1} (${tabData.title}) 업데이트 완료`);
  });
  console.log("전체 업데이트 완료");
}
function runEmojiNumbering() {
  const tabsData = getTopFiveTabs();
  tabsData.forEach((tabData, i) => {
    addEmojiNumberingWithBody(tabData.body);
    console.log(`탭 ${i + 1} (${tabData.title}) 숫자 매기기 완료`);
  });
}
function runFlowerPrefix() {
  const tabsData = getTopFiveTabs();
  tabsData.forEach((tabData, i) => {
    addFlowerPrefixWithBody(tabData.body);
    console.log(`탭 ${i + 1} (${tabData.title}) 말머리 추가 완료`);
  });
}
function runImportEvents() {
  const tabsData = getTopFiveTabs();
  tabsData.forEach((tabData, i) => {
    importEventsWithBody(tabData.body);
    console.log(`탭 ${i + 1} (${tabData.title}) 일정 가져오기 완료`);
  });
}
function runCreateNewWeek() {
  const doc = DocumentApp.openById(DOC_ID);
  const activeTab = doc.getActiveTab();
  const body = activeTab.asDocumentTab().getBody();
  const header = activeTab.asDocumentTab().getHeader();
  const title = activeTab.getTitle();

  // 오늘이 월요일이면 이번 주 월요일, 아니면 다음 주 월요일
  const nextMon = getMondayOfWeekFrom(new Date());
  const month = nextMon.getMonth();
  const year = nextMon.getFullYear();

  let sem, academicYear;
  if (month >= SEM1_START_MONTH && month < SEM2_START_MONTH) {
    sem = 1;
    academicYear = year;
  } else {
    sem = 2;
    academicYear = (month >= SEM2_START_MONTH) ? year : year - 1;
  }

  const semStart = (sem === 1)
    ? new Date(academicYear, SEM1_START_MONTH, SEM1_START_DAY)
    : new Date(academicYear, SEM2_START_MONTH, SEM2_START_DAY);
  const semStartMon = new Date(semStart);
  semStartMon.setDate(semStart.getDate() - (semStart.getDay() === 0 ? 6 : semStart.getDay() - 1));
  const week = Math.floor((nextMon - semStartMon) / 604800000) + 1;

  if (header) {
    header.replaceText("\\d*학기", sem + "학기");
    header.replaceText("\\d*주차", week + "주차");
  }

  const table = findTableByHeader(body, "주간일정");
  if (table) updateScheduleTable(table, nextMon);

  importEventsWithBody(body);

  console.log(`탭 (${title}) 새 주차 만들기 완료 - ${sem}학기 ${week}주차`);
}




// ======================== [4] 섹션별 서식 적용 함수 ========================
function addEmojiNumberingWithBody(body) {
  const emojiNumbers = ["1️⃣", "2️⃣", "3️⃣", "4️⃣", "5️⃣", "6️⃣", "7️⃣", "8️⃣", "9️⃣"];
  const fixedSpace = "\u2009";
  const tables = findTablesByKeyword(body, "부서별 전달사항");
  tables.forEach(table => {
    for (let r = 1; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      const numCells = row.getNumCells();
      for (let c = 0; c < numCells; c++) {
        if (numCells > 1 && c === 0) continue;
        processCellParagraphs(row.getCell(c), emojiNumbers, fixedSpace, 16, true);
      }
    }
  });
}
function addFlowerPrefixWithBody(body) {
  const flowerPrefix = ["✿"];
  const fixedSpace = "\u2009";
  const keywords = ["주요 협의 내용", "논의사항"];
  keywords.forEach(keyword => {
    const tables = findTablesByKeyword(body, keyword);
    tables.forEach(table => {
      for (let r = 1; r < table.getNumRows(); r++) {
        const row = table.getRow(r);
        const numCells = row.getNumCells();
        for (let c = 0; c < numCells; c++) {
          processCellParagraphs(row.getCell(c), flowerPrefix, fixedSpace, 12, false);
        }
      }
    });
  });
}
function findTablesByKeyword(body, searchText) {
  const tables = body.getTables();
  return tables.filter(table => table.getText().includes(searchText));
}
// ======================== [5] 일정 가져오기 로직 ========================
function importEventsWithBody(body) {
  const spreadsheetId = "1BsS_nRsdrBpyv0ebcuuQ3fKGG_6ZitHVSye97ceWaCw";
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName("주요 일정");
  if (!sheet) return;
  const targetTable = findTableByHeader(body, "주간일정");
  if (!targetTable) return;
  const rawValues = sheet.getDataRange().getValues();
  const richTextValues = sheet.getDataRange().getRichTextValues();
  // 실제 테이블 행 수를 확인하여 범위 초과 방지 (헤더 행 제외)
  const maxRow = Math.min(7, targetTable.getNumRows() - 1);
  for (let row = 1; row <= maxRow; row++) {
    const dateCell = targetTable.getCell(row, 0);
    const eventCell = targetTable.getCell(row, 1);
    const targetDate = getDateFromCell(dateCell);
    if (targetDate) {
      const events = findEventsFromDatabase(rawValues, richTextValues, targetDate);
      eventCell.clear();
      if (events.length > 0) {
        const textObj = eventCell.editAsText();
        let pos = 0;
        events.forEach((ev, i) => {
          const line = ev.fullText + (i < events.length - 1 ? "\n" : "");
          textObj.appendText(line);
          if (ev.deptLink) textObj.setLinkUrl(pos + 1, pos + ev.deptText.length, ev.deptLink);
          if (ev.contentLink) {
            const start = pos + ev.deptText.length + 3;
            textObj.setLinkUrl(start, start + ev.contentText.length - 1, ev.contentLink);
          }
          pos += line.length;
        });
        textObj.setFontFamily("Open Sans").setFontSize(10).setBold(false);
        applyBoldToBrackets(textObj);
      }
    } else {
      eventCell.clear();
    }
  }
}
// ======================== [6] 서식 및 링크 보존 유틸리티 ========================
function processCellParagraphs(cell, prefixArr, space, indent, isNum) {
  let count = 0;
  for (let i = 0; i < cell.getNumChildren(); i++) {
    const child = cell.getChild(i);
    const childType = child.getType();
    if (childType === DocumentApp.ElementType.PARAGRAPH || childType === DocumentApp.ElementType.LIST_ITEM) {
      const p = (childType === DocumentApp.ElementType.PARAGRAPH)
                ? child.asParagraph()
                : child.asListItem();
      const text = p.getText().trim();
      if (text !== "") {
        const isContent = (childType === DocumentApp.ElementType.LIST_ITEM) ||
                         text.match(/^([*\-•∙◦]|\d+\.)\s*/);
        if (isContent) {
          p.setLineSpacing(1.1).setSpacingBefore(0).setSpacingAfter(3);
          p.editAsText().setFontSize(10).setFontFamily("Open Sans").setBold(false);
        } else {
          applyFormatting(p, count, prefixArr, space, isNum, indent);
          count++;
        }
      }
    }
  }
}
function applyFormatting(paragraph, count, prefixArr, space, isNum, indent) {
  const textObj = paragraph.editAsText();
  const original = textObj.getText();
  const match = original.match(/^([1-9]️⃣|✿|[*]|[-]|•)\s*/);
  const removedLen = match ? match[0].length : 0;
  const core = original.substring(removedLen);
  const links = [];
  for (let i = 0; i < original.length; i++) {
    const url = textObj.getLinkUrl(i);
    if (url) {
      let s = i, e = i;
      while (e < original.length - 1 && textObj.getLinkUrl(e + 1) === url) e++;
      // prefix 영역 안에서 시작하는 링크는 제거된 텍스트에 속하므로 제외
      if (s >= removedLen) links.push({ s: s - removedLen, e: e - removedLen, url });
      i = e;
    }
  }
  let pfx = isNum ? (count < prefixArr.length ? prefixArr[count] + space : (count + 1) + "." + space) : prefixArr[0] + space;
  textObj.setText(pfx + core);
  links.forEach(l => {
    let nS = pfx.length + l.s, nE = pfx.length + l.e;
    if (nS < textObj.getText().length) textObj.setLinkUrl(nS, Math.min(nE, textObj.getText().length - 1), l.url);
  });
  paragraph.setIndentStart(indent).setIndentFirstLine(0).setLineSpacing(1.2).setSpacingBefore(8).setSpacingAfter(0);
  textObj.setFontSize(11).setFontFamily("Open Sans").setBold(false);
}
// ======================== [7] 헬퍼 및 날짜 관련 ========================
function findTableByHeader(body, searchText) {
  const tables = body.getTables();
  for (let i = 0; i < tables.length; i++) {
    if (tables[i].getText().includes(searchText)) return tables[i];
  }
  return null;
}
function applyBoldToBrackets(textObj) {
  const fullText = textObj.getText();
  const regex = /\[.*?\]/g;
  let m;
  while ((m = regex.exec(fullText)) !== null) textObj.setBold(m.index, m.index + m[0].length - 1, true);
}
function findEventsFromDatabase(raw, rich, targetDate) {
  const events = [];
  const ty = targetDate.getFullYear(), tm = targetDate.getMonth(), td = targetDate.getDate();
  for (let i = 1; i < raw.length; i++) {
    let d = raw[i][0];
    if (!(d instanceof Date)) d = new Date(d);
    if (!isNaN(d.getTime()) && d.getFullYear() === ty && d.getMonth() === tm && d.getDate() === td) {
      const deptText = rich[i][1].getText(), deptLink = rich[i][1].getLinkUrl();
      const contentText = rich[i][2].getText(), contentLink = rich[i][2].getLinkUrl();
      const startP = raw[i][3], endP = raw[i][4], startT = raw[i][5], endT = raw[i][6];
      let time = "";
      if (endP !== "") time = " (" + startP + "-" + endP + ")";
      else if (startT !== "") {
        const f = (startT instanceof Date)
          ? String(startT.getHours()).padStart(2, '0') + ":" + String(startT.getMinutes()).padStart(2, '0')
          : startT;
        const g = (endT instanceof Date)
          ? String(endT.getHours()).padStart(2, '0') + ":" + String(endT.getMinutes()).padStart(2, '0')
          : endT;
        time = " (" + f + "-" + g + ")";
      }
      events.push({ fullText: "[" + deptText + "] " + contentText + time, deptText, deptLink, contentText, contentLink });
    }
  }
  return events;
}
function getDateFromCell(cell) {
  const m = cell.getText().match(/(\d{1,2})월\s*(\d{1,2})일/);
  if (!m) return null;
  const month = parseInt(m[1]) - 1;
  const day = parseInt(m[2]);
  const today = new Date();
  let year = today.getFullYear();
  if (today.getMonth() >= 10 && month <= 1) year++;
  else if (today.getMonth() <= 1 && month >= 10) year--;
  return new Date(year, month, day);
}
// 오늘이 월요일이면 오늘을 반환, 아니면 다음 주 월요일을 반환
function getMondayOfWeekFrom(d) {
  const day = d.getDay();
  const diff = (day === 1 ? 0 : day === 0 ? 1 : 8 - day);
  const res = new Date(d);
  res.setDate(d.getDate() + diff);
  return res;
}
function updateScheduleTable(table, startDate) {
  let curr = new Date(startDate);
  for (let row = 1; row <= 7; row++) {
    while (curr.getDay() === 0) curr.setDate(curr.getDate() + 1);
    const month = curr.getMonth() + 1;
    const date = curr.getDate();
    const dayOfWeek = ["일", "월", "화", "수", "목", "금", "토"][curr.getDay()];
    table.getRow(row).getCell(0).setText(month + "월 " + date + "일 " + dayOfWeek + "요일");
    curr.setDate(curr.getDate() + 1);
  }
}
