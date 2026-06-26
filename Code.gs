// =============================================
// 2026학년도 대토론회 학교 공통의제 - 3주체 의견 모음 웹앱
// =============================================

const SS_ID = '1FMsxZkM0wOVvY-CmBqwxYZ7vcfEp8ZxaFnPPzPQgYo0';
const SHEET_CANDIDATES = ['설문지 응답 시트1', '설문지응답시트1'];

const AGENDA_LABELS = ['소통과 존중', '자율과 조율', '배움과 평화'];

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (!action) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('대토론회 의견 모음')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  let data;
  try {
    if (action === 'getOpinions') data = { opinions: getOpinions() };
    else if (action === 'debug') data = debugSheetInfo([
      { id: SS_ID, label: '대토론회', infoSheetCandidates: SHEET_CANDIDATES }
    ]);
    else data = { error: 'unknown action: ' + action };
  } catch (err) {
    data = { error: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOpinions() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = findSheet(ss, SHEET_CANDIDATES);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  const result = [];

  rows.forEach(row => {
    const subjectRaw = norm(row[2] || ''); // C열
    let subject = '';
    if (subjectRaw.includes('보호자')) subject = '학부모';
    else if (subjectRaw.includes('학생')) subject = '학생';
    else if (subjectRaw.includes('교사')) subject = '교사';
    else return;

    let grade = '', name = '', attend = '';
    let contents = ['', '', ''];

    if (subject === '학생') {
      grade = row[3];           // D열
      contents = [row[4], row[5], row[6]]; // E,F,G
    } else if (subject === '교사') {
      contents = [row[7], row[8], row[9]]; // H,I,J
    } else if (subject === '학부모') {
      grade = row[10];          // K열 (자녀 학년)
      contents = [row[11], row[12], row[13]]; // L,M,N
      attend = row[14];         // O열
      name = row[15];           // P열
    }

    const timestamp = row[0] ? String(row[0]) : '';

    contents.forEach((c, idx) => {
      const text = String(c || '').trim();
      if (!text) return;
      result.push({
        subject: subject,
        grade: String(grade || '').trim(),
        name: String(name || '').trim(),
        attend: String(attend || '').trim(),
        agendaIndex: idx + 1,
        agendaLabel: AGENDA_LABELS[idx],
        content: text,
        timestamp: timestamp
      });
    });
  });

  return result;
}

// =============================================
// 유틸리티
// =============================================

function norm(s) {
  return String(s).replace(/[\s　​]/g, '').normalize('NFC');
}

function findSheet(ss, candidates) {
  for (const name of candidates) {
    const s = ss.getSheetByName(name);
    if (s) return s;
  }
  const all = ss.getSheets();
  for (const cand of candidates) {
    const found = all.find(s => s.getName().includes(cand.replace(/\s/g, '')));
    if (found) return found;
  }
  return null;
}

function debugSheetInfo(sheetConfigs) {
  const result = {};
  sheetConfigs.forEach(({ id, label, infoSheetCandidates }) => {
    try {
      const ss = SpreadsheetApp.openById(id);
      const sheets = ss.getSheets().map(s => s.getName());
      result[label] = { sheets };
      const infoSheet = findSheet(ss, infoSheetCandidates || []);
      if (infoSheet) {
        result[label].header = infoSheet.getRange(1, 1, 1, infoSheet.getLastColumn()).getValues()[0];
        result[label].row2 = infoSheet.getRange(2, 1, 1, infoSheet.getLastColumn()).getValues()[0];
      } else {
        result[label].infoSheetMissing = true;
      }
    } catch (e) {
      result[label] = { error: e.toString() };
    }
  });
  return result;
}
