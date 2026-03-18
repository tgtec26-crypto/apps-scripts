function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('데이터 관리')
    .addItem('교사별 대출 현황 자료 생성', 'createTeacherLoanStatus')
    .addItem('기존 관리번호 처리', 'processExistingAdminNumbers')
    .addItem('재고 현황 생성', 'createInventoryStatus')
    .addToUi();
  setupDependentDropdowns();
}

// 드롭다운 설정
function setupDependentDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('교사용');

  // 건물명 드롭다운 설정 - start from row 2
  const buildingSheet = ss.getSheetByName('건물별 실 및 망 현황');
  const buildingData = buildingSheet.getRange('A2:A' + buildingSheet.getLastRow()).getValues().flat();
  const buildingRule = SpreadsheetApp.newDataValidation().requireValueInList(buildingData).build();
  sheet.getRange('G2:G' + sheet.getLastRow()).setDataValidation(buildingRule);

  // 교과군 드롭다운 설정 - start from row 2
  const subjectSheet = ss.getSheetByName('과목직급별 성함');
  const subjectData = subjectSheet.getRange('A2:A' + subjectSheet.getLastRow()).getValues().flat();
  const subjectRule = SpreadsheetApp.newDataValidation().requireValueInList(subjectData).build();
  sheet.getRange('J2:J' + sheet.getLastRow()).setDataValidation(subjectRule);
}

// 기자재 정보 조회 함수
function getDeviceInfo(adminNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const deviceSheet = ss.getSheetByName('기자재 구매 시기와 수량');
  const data = deviceSheet.getRange('B2:E' + deviceSheet.getLastRow()).getValues();

  const prefix = adminNumber.split('-')[0];
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === prefix) {
      return {
        manufacturer: data[i][1], // C열 제조업체
        purchaseYear: data[i][2], // D열 취득년도
        modelName: data[i][3]     // E열 제품명
      };
    }
  }
  
  // 일치하는 정보가 없을 경우 기본값 반환
  return {
    manufacturer: '기타',
    purchaseYear: '',
    modelName: ''
  };
}

// 편집 이벤트 처리
function onEdit(e) {
  const sheet = e.range.getSheet();
  const range = e.range;

  // 교사별 대출 현황 시트 체크박스 처리
  if (sheet.getName() === '교사별 대출 현황') {
    const row = range.getRow();
    const col = range.getColumn();
    if (row > 1 && col === 3) { // C열 체크박스
      const value = range.getValue();
      if (value) {
        sheet.getRange(row, 4).setValue(new Date()); // D열 확인일자
        sheet.getRange(row, 5).setValue(Session.getActiveUser().getEmail()); // E열 확인자 이메일
      } else {
        sheet.getRange(row, 4).clearContent();
        sheet.getRange(row, 5).clearContent();
      }
    }
    return;
  }

  // 교사용 시트에서만 처리
  if (sheet.getName() === '교사용') {
    const startRow = range.getRow();
    const endRow = range.getLastRow();
    const startCol = range.getColumn();
    const endCol = range.getLastColumn();

    // 헤더 행(1행)은 처리하지 않음
    if (startRow === 1) return;

    // 드래그로 자동 채우기된 범위 처리
    if (startRow !== endRow || startCol !== endCol) {
      handleBatchUpdate(sheet, startRow, endRow, startCol, endCol);
    } else {
      // 단일 셀 편집 처리
      handleSingleCellUpdate(e);
    }
  }
}

// 드래그로 자동 채우기된 범위 처리
function handleBatchUpdate(sheet, startRow, endRow, startCol, endCol) {
  for (let row = startRow; row <= endRow; row++) {
    // 헤더 행(1행)은 처리하지 않음
    if (row === 1) continue;
    
    for (let col = startCol; col <= endCol; col++) {
      const cell = sheet.getRange(row, col);
      const value = cell.getValue();

      // B열 (관리번호) 처리
      if (col === 2) {
        if (value) {
          const deviceInfo = getDeviceInfo(value);
          sheet.getRange(row, 3).setValue(deviceInfo.manufacturer); // C열 (제조업체)
          sheet.getRange(row, 4).setValue(deviceInfo.purchaseYear); // D열 (구매년도)
          sheet.getRange(row, 5).setValue(deviceInfo.modelName);   // E열 (모델명)
        } else {
          // B열이 비어있으면 C, D, E열도 비우기
          sheet.getRange(row, 3).clearContent(); // C열 (제조업체) 비우기
          sheet.getRange(row, 4).clearContent(); // D열 (구매년도) 비우기
          sheet.getRange(row, 5).clearContent(); // E열 (모델명) 비우기
        }
      }

      // G열 (건물명) 처리
      if (col === 7 && value) {
        updateRoomDropdown(sheet, row, value);
      }

      // J열 (교과군) 처리
      if (col === 10 && value) {
        updateBorrowerDropdown(sheet, row, value);
      }

      // L열 (대출일) 처리
      if (col === 12 && value) {
        const returnDate = sheet.getRange(row, 13).getValue();
        if (returnDate && new Date(value) > new Date(returnDate)) {
          sheet.getRange(row, 13).clearContent();
        }
      }

      // N열 (체크박스) 처리
      if (col === 14) {
        handleCheckbox(sheet, row, value);
      }
    }
  }
}

// 단일 셀 편집 처리
function handleSingleCellUpdate(e) {
  const sheet = e.range.getSheet();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  const value = range.getValue();

  // 헤더 행(1행)은 처리하지 않음
  if (row === 1) return;

  // B열 (관리번호) 처리
  if (col === 2) {
    if (value) {
      const deviceInfo = getDeviceInfo(value);
      sheet.getRange(row, 3).setValue(deviceInfo.manufacturer); // C열 (제조업체)
      sheet.getRange(row, 4).setValue(deviceInfo.purchaseYear); // D열 (구매년도)
      sheet.getRange(row, 5).setValue(deviceInfo.modelName);   // E열 (모델명)
    } else {
      // B열이 비어있으면 C, D, E열도 비우기
      sheet.getRange(row, 3).clearContent(); // C열 (제조업체) 비우기
      sheet.getRange(row, 4).clearContent(); // D열 (구매년도) 비우기
      sheet.getRange(row, 5).clearContent(); // E열 (모델명) 비우기
    }
  }

  // G열 (건물명) 처리
  if (col === 7 && value) {
    updateRoomDropdown(sheet, row, value);
  }

  // J열 (교과군) 처리
  if (col === 10 && value) {
    updateBorrowerDropdown(sheet, row, value);
  }

  // L열 (대출일) 처리
  if (col === 12 && value) {
    const returnDate = sheet.getRange(row, 13).getValue();
    if (returnDate && new Date(value) > new Date(returnDate)) {
      sheet.getRange(row, 13).clearContent();
    }
  }

  // N열 (체크박스) 처리
  if (col === 14) {
    handleCheckbox(sheet, row, value);
  }
}

// 건물명 선택 시 실명 드롭다운 업데이트
function updateRoomDropdown(sheet, row, building) {
  // 헤더 행(1행)은 처리하지 않음
  if (row === 1) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const buildingSheet = ss.getSheetByName('건물별 실 및 망 현황');
  const data = buildingSheet.getRange('A2:B' + buildingSheet.getLastRow()).getValues();
  const rooms = data.filter(rowData => rowData[0] === building).map(rowData => rowData[1]);

  const roomRule = SpreadsheetApp.newDataValidation().requireValueInList(rooms).build();
  sheet.getRange(row, 8).setDataValidation(roomRule); // H열 (실명)
}

// 교과군 선택 시 대출자 드롭다운 업데이트
function updateBorrowerDropdown(sheet, row, subject) {
  // 헤더 행(1행)은 처리하지 않음
  if (row === 1) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const subjectSheet = ss.getSheetByName('과목직급별 성함');
  const data = subjectSheet.getRange('A2:B' + subjectSheet.getLastRow()).getValues();
  const borrowers = data.filter(rowData => rowData[0] === subject).map(rowData => rowData[1]);

  const borrowerRule = SpreadsheetApp.newDataValidation().requireValueInList(borrowers).build();
  sheet.getRange(row, 11).setDataValidation(borrowerRule); // K열 (대출자)
}

// 체크박스 처리
function handleCheckbox(sheet, row, checked) {
  // 헤더 행(1행)은 처리하지 않음
  if (row === 1) return;

  if (checked) {
    sheet.getRange(row, 15).setValue(new Date()); // O열 (확인일자)
    sheet.getRange(row, 16).setValue(Session.getActiveUser().getEmail()); // P열 (확인자 이메일)
  } else {
    sheet.getRange(row, 15).clearContent();
    sheet.getRange(row, 16).clearContent();
  }
}

// 교사별 대출 현황 자료 생성
function createTeacherLoanStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teacherSheet = ss.getSheetByName('교사용');
  const loanStatusSheet = ss.getSheetByName('교사별 대출 현황');
  
  // 시트가 없으면 생성
  let sheet = loanStatusSheet;
  if (!sheet) {
    sheet = ss.insertSheet('교사별 대출 현황');
  }
  
  // 기존 데이터 지우기
  if (sheet.getLastRow() > 1) {
    sheet.getRange('A2:B' + sheet.getLastRow()).clearContent();
  }

  // 헤더 설정
  sheet.getRange('A1').setValue('성함');
  sheet.getRange('B1').setValue('대출 기자재');
  sheet.getRange('A1:B1').setBackground('darkblue').setFontColor('white');

  // 데이터 가져오기
  const dataRange = teacherSheet.getDataRange();
  const values = dataRange.getValues();
  
  // 헤더 행에서 대출자(K열)와 관리번호(A열), 반납일(M열) 열 인덱스 찾기
  const headers = values[0];
  // K열은 일반적으로 인덱스 10
  const borrowerColIndex = headers.findIndex(header => header === '대출자') !== -1 ? 
                            headers.findIndex(header => header === '대출자') : 10;
  // A열은 일반적으로 인덱스 0
  const idColIndex = headers.findIndex(header => header === '관리번호') !== -1 ? 
                      headers.findIndex(header => header === '관리번호') : 0;
  // M열은 일반적으로 인덱스 12
  const returnDateColIndex = headers.findIndex(header => header === '반납일') !== -1 ? 
                              headers.findIndex(header => header === '반납일') : 12;
  
  // 대출자 목록 생성 (중복 제거)
  const borrowerMap = {};
  
  // 첫 번째 행(헤더)은 건너뛰고 처리
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const borrower = row[borrowerColIndex];
    const equipmentId = row[idColIndex];
    const returnDate = row[returnDateColIndex];
    
    // 대출자가 있고, 관리번호가 있으며, 반납일이 없는 경우만 처리
    if (borrower && equipmentId && !returnDate) {
      if (!borrowerMap[borrower]) {
        borrowerMap[borrower] = [];
      }
      borrowerMap[borrower].push(equipmentId);
    }
  }
  
  // 대출자별 대출 현황 시트에 데이터 입력
  let rowIndex = 2; // 데이터는 2행부터 시작
  
  for (const borrower in borrowerMap) {
    const equipmentIds = borrowerMap[borrower];
    
    // 대출자 이름 입력
    sheet.getRange(rowIndex, 1).setValue(borrower);

    // 관리번호를 쉼표로 구분해 B열에 입력
    sheet.getRange(rowIndex, 2).setValue(equipmentIds.join(', '));

    rowIndex++;
  }
  
  // 열 너비 자동 조정
  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);
  
  // 완료 메시지 표시
  SpreadsheetApp.getUi().alert('교사별 대출 현황이 업데이트되었습니다.');
}
// 기존 관리번호 처리
function processExistingAdminNumbers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('교사용');
  const data = sheet.getRange('A2:P' + sheet.getLastRow()).getValues();

  data.forEach((row, index) => {
    const adminNumber = row[1]; // B열 (관리번호)
    if (adminNumber && (!row[2] || !row[3] || !row[4])) { // C열, D열, E열 중 비어있는 경우
      const deviceInfo = getDeviceInfo(adminNumber);
      
      // 제조업체 추출 및 입력 (C열)
      if (!row[2]) {
        sheet.getRange(index + 2, 3).setValue(deviceInfo.manufacturer);
      }
      
      // 구매년도 추출 및 입력 (D열)
      if (!row[3]) {
        sheet.getRange(index + 2, 4).setValue(deviceInfo.purchaseYear);
      }
      
      // 모델명 추출 및 입력 (E열)
      if (!row[4]) {
        sheet.getRange(index + 2, 5).setValue(deviceInfo.modelName);
      }
    }
  });

  SpreadsheetApp.getUi().alert('기존 관리번호 처리 완료: 제조업체, 구매년도, 모델명이 업데이트되었습니다.');
}

// 재고 현황 생성
function createInventoryStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teacherSheet = ss.getSheetByName('교사용');
  const inventorySheet = ss.getSheetByName('교사용 재고 현황');

  // 기존 데이터 지우기
  if (inventorySheet.getLastRow() > 1) {
    inventorySheet.getRange('A2:C' + inventorySheet.getLastRow()).clearContent();
  }

  // 데이터 추출
  const data = teacherSheet.getRange('A2:M' + teacherSheet.getLastRow()).getValues();
  const inventoryMap = {};

  data.forEach(row => {
    const adminNumber = row[1];
    if (adminNumber) {
      const prefix = adminNumber.split('-')[0];
      if (!inventoryMap[prefix]) {
        inventoryMap[prefix] = { total: 0, available: 0 };
      }
      inventoryMap[prefix].total++;
      
      // 대출자가 없거나 반납일이 있는 경우 사용 가능으로 처리
      if (!row[10] || row[12]) {
        inventoryMap[prefix].available++;
      }
    }
  });

  // 데이터 입력
  let inventoryData = Object.keys(inventoryMap).map(prefix => [
    prefix,
    inventoryMap[prefix].available,
    inventoryMap[prefix].total
  ]);
  
  // A열(prefix) 기준으로 오름차순 정렬
  inventoryData.sort((a, b) => {
    if (a[0] < b[0]) return -1;
    if (a[0] > b[0]) return 1;
    return 0;
  });
  
  if (inventoryData.length > 0) {
    inventorySheet.getRange(2, 1, inventoryData.length, 3).setValues(inventoryData);
  }
}