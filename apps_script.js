// ── 출석체크 Apps Script ──
const SHEET_ID = '1J0S09jZSGetnA8Bg75dqWc9sNUcdRVnN5_hqZc8GhQI';

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'getAttendance') return getAttendance(e.parameter);
  return jsonResp({ status: 'ok', message: '출석체크 API 정상 작동 중' });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'saveAttendance') return saveAttendance(data);
    return jsonResp({ error: 'unknown action' });
  } catch(err) {
    return jsonResp({ error: err.message });
  }
}

// ── 오늘 출석 데이터 불러오기 ──
function getAttendance(params) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const session = params.session;
  const cls = params.class;
  const date = params.date;

  const sessionLabel = session === 'morning' ? '아침자습' : '수업';
  const sheetName = `${sessionLabel}_${cls}`;
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) return jsonResp({ records: [] });

  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const lastRow = Math.max(sheet.getLastRow(), 1);
  if (lastRow < 2) return jsonResp({ records: [] });

  const headerRow = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const dateCol = headerRow.indexOf(date) + 1;
  const memoCol = headerRow.indexOf(date + ' 사유') + 1;

  if (dateCol === 0) return jsonResp({ records: [] });

  const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const statuses = sheet.getRange(2, dateCol, lastRow - 1, 1).getValues();
  const memos = memoCol > 0
    ? sheet.getRange(2, memoCol, lastRow - 1, 1).getValues()
    : [];

  const records = names.map((r, i) => ({
    name: r[0],
    status: statuses[i][0] || '',
    memo: memos[i] ? memos[i][0] : ''
  })).filter(r => r.name);

  return jsonResp({ records });
}

// ── 출석 저장 ──
function saveAttendance(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sessionLabel = data.session === 'morning' ? '아침자습' : '수업';
  const sheetName = `${sessionLabel}_${data.class}`;

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    const hdrCell = sheet.getRange(1, 1);
    hdrCell.setValue('학생명');
    hdrCell.setBackground('#185FA5').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);
  }

  const date = data.date;
  const records = data.records;

  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];

  let dateCol = headerRow.indexOf(date) + 1;
  let memoCol = headerRow.indexOf(date + ' 사유') + 1;

  if (dateCol === 0) {
    dateCol = lastCol + 1;
    memoCol = lastCol + 2;
    const dateCell = sheet.getRange(1, dateCol);
    dateCell.setNumberFormat('@');
    dateCell.setValue(date);
    dateCell.setBackground('#185FA5').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    const memoCell = sheet.getRange(1, memoCol);
    memoCell.setNumberFormat('@');
    memoCell.setValue(date + ' 사유');
    memoCell.setBackground('#4472C4').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  }

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

    const cell = sheet.getRange(row, dateCol);
    cell.setValue(rec.status).setHorizontalAlignment('center');
    if (rec.status === '출석') {
      cell.setBackground('#E1F5EE').setFontColor('#085041').setFontWeight('normal');
    } else if (rec.status === '결석') {
      cell.setBackground('#FCEBEB').setFontColor('#501313').setFontWeight('bold');
    } else if (rec.status === '지각') {
      cell.setBackground('#FAEEDA').setFontColor('#412402').setFontWeight('bold');
    } else {
      cell.setBackground(null).setFontColor(null).setFontWeight('normal');
    }

    const memoCellRange = sheet.getRange(row, memoCol);
    memoCellRange.setValue(rec.memo || '').setHorizontalAlignment('left');
    if (rec.memo) {
      memoCellRange.setBackground('#FFF9E6').setFontColor('#5c4000');
    } else {
      memoCellRange.setBackground(null).setFontColor(null);
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
