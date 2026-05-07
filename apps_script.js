// ── 출석체크 Apps Script (사유 메모 지원) ──
// 스프레드시트 ID를 아래에 입력하세요
const SHEET_ID = 'YOUR_SPREADSHEET_ID';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'saveAttendance') return saveAttendance(data);
    return jsonResp({ error: 'unknown action' });
  } catch(err) {
    return jsonResp({ error: err.message });
  }
}

function doGet(e) {
  return jsonResp({ status: 'ok', message: '출석체크 API 정상 작동 중' });
}

function saveAttendance(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sessionLabel = data.session === 'morning' ? '아침자습' : '수업';
  const sheetName = `${sessionLabel}_${data.class}`;

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // 헤더 행: 학생명 | 날짜(출석) | 날짜(사유) | ...
    const hdrCell = sheet.getRange(1, 1);
    hdrCell.setValue('학생명');
    hdrCell.setBackground('#185FA5').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);
  }

  const date = data.date;       // e.g. "5/8"
  const records = data.records; // [{name, status, memo}, ...]

  // 1행에서 날짜 열 찾기 (출석 열)
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // 날짜 헤더: "5/8" (출석), "5/8 사유" (메모)
  const dateHeader = date;
  const memoHeader = date + ' 사유';

  let dateCol = headerRow.indexOf(dateHeader) + 1;
  let memoCol = headerRow.indexOf(memoHeader) + 1;

  // 새 날짜면 열 추가 (출석 + 사유 두 열)
  if (dateCol === 0) {
    dateCol = lastCol + 1;
    memoCol = lastCol + 2;
    const dateCell = sheet.getRange(1, dateCol);
    dateCell.setValue(dateHeader);
    dateCell.setBackground('#185FA5').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    const memoCell = sheet.getRange(1, memoCol);
    memoCell.setValue(memoHeader);
    memoCell.setBackground('#4472C4').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  }

  // 학생별 행 찾아서 출석 + 사유 기록
  records.forEach((rec) => {
    const lastRow = Math.max(sheet.getLastRow(), 1);
    const nameVals = sheet.getRange(2, 1, Math.max(lastRow - 1, 1), 1).getValues();
    let row = -1;
    for (let i = 0; i < nameVals.length; i++) {
      if (nameVals[i][0] === rec.name) { row = i + 2; break; }
    }
    if (row === -1) {
      row = lastRow + 1;
      sheet.getRange(row, 1).setValue(rec.name);
    }

    // 출석 상태
    const cell = sheet.getRange(row, dateCol);
    cell.setValue(rec.status).setHorizontalAlignment('center');
    if (rec.status === '출석') {
      cell.setBackground('#E1F5EE').setFontColor('#085041');
    } else if (rec.status === '결석') {
      cell.setBackground('#FCEBEB').setFontColor('#501313').setFontWeight('bold');
    } else if (rec.status === '지각') {
      cell.setBackground('#FAEEDA').setFontColor('#412402').setFontWeight('bold');
    } else {
      cell.setBackground(null).setFontColor(null).setFontWeight('normal');
    }

    // 사유 메모
    const memoCell = sheet.getRange(row, memoCol);
    memoCell.setValue(rec.memo || '').setHorizontalAlignment('left');
    if (rec.memo) {
      memoCell.setBackground('#FFF9E6').setFontColor('#5c4000');
    } else {
      memoCell.setBackground(null).setFontColor(null);
    }
  });

  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(memoCol);

  return jsonResp({ message: `${sessionLabel} ${data.class} ${date} 저장 완료 ✓` });
}

function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
