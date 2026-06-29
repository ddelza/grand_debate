// =============================================
// 2026학년도 대토론회 학교 공통의제 - 3주체 의견 모음 웹앱
// =============================================

const SS_ID = '1FMsxZkM0wOVvY-CmBqwxYZ7vcfEp8ZxaFnPPzPQgYo0';
const SHEET_CANDIDATES = ['설문지 응답 시트1', '설문지응답시트1'];
const HIDDEN_SHEET_NAME = '숨김처리';

const AGENDA_LABELS = ['소통과 존중', '자율과 조율', '배움과 평화'];

// 교사용 숨기기/복원 비밀번호 — 운영 전에 원하는 값으로 바꿔서 재배포하세요.
const TEACHER_PASSWORD = '7683101';

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  const page = e && e.parameter && e.parameter.page;
  if (!action) {
    if (page === 'attendance') {
      return HtmlService.createHtmlOutputFromFile('attendance')
        .setTitle('대토론회 참석 가능여부 조회')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('대토론회 의견 모음')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  const p = (e && e.parameter) || {};
  let data;
  try {
    if (action === 'getOpinions') {
      const includeHidden = p.includeHidden === 'true';
      if (includeHidden) checkPassword_(p.password);
      data = { opinions: getOpinions(includeHidden) };
    } else if (action === 'getAttendance') {
      data = { records: getAttendanceList() };
    } else if (action === 'hideOpinion') {
      checkPassword_(p.password);
      setHidden_(p.id, true);
      data = { ok: true };
    } else if (action === 'unhideOpinion') {
      checkPassword_(p.password);
      setHidden_(p.id, false);
      data = { ok: true };
    } else if (action === 'resetHidden') {
      checkPassword_(p.password);
      resetHidden_();
      data = { ok: true };
    } else if (action === 'debug') {
      data = debugSheetInfo([
        { id: SS_ID, label: '대토론회', infoSheetCandidates: SHEET_CANDIDATES }
      ]);
    } else {
      data = { error: 'unknown action: ' + action };
    }
  } catch (err) {
    data = { error: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// google.script.run 호출용 래퍼 (Apps Script 자체 웹뷰에서 열렸을 때)
function hideOpinion_(id, password) {
  checkPassword_(password);
  setHidden_(id, true);
  return { ok: true };
}

function unhideOpinion_(id, password) {
  checkPassword_(password);
  setHidden_(id, false);
  return { ok: true };
}

function checkPassword_(pw) {
  if (String(pw || '') !== TEACHER_PASSWORD) {
    throw new Error('비밀번호가 일치하지 않습니다.');
  }
}

function getOpinions(includeHidden) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = findSheet(ss, SHEET_CANDIDATES);
  if (!sheet) return [];

  const hiddenSet = getHiddenSet_(ss);

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  const result = [];

  rows.forEach((row, i) => {
    const rowNumber = i + 2; // 시트 상의 실제 행 번호 (헤더가 1행)

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

    const tsDate = row[0] instanceof Date ? row[0] : null;
    const timestamp = tsDate ? Utilities.formatDate(tsDate, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : String(row[0] || '');
    const timestampMs = tsDate ? tsDate.getTime() : 0;

    contents.forEach((c, idx) => {
      const text = String(c || '').trim();
      if (!text) return;
      const id = rowNumber + '-' + (idx + 1);
      const isHidden = hiddenSet.has(id);
      if (isHidden && !includeHidden) return;

      result.push({
        id: id,
        subject: subject,
        grade: String(grade || '').trim(),
        name: String(name || '').trim(),
        attend: String(attend || '').trim(),
        agendaIndex: idx + 1,
        agendaLabel: AGENDA_LABELS[idx],
        content: text,
        timestamp: timestamp,
        timestampMs: timestampMs,
        hidden: isHidden
      });
    });
  });

  return result;
}

// 보호자(학부모) 응답자 단위로 한 행씩 — 대토론회 참석 가능여부 조회용
function getAttendanceList() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = findSheet(ss, SHEET_CANDIDATES);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  const result = [];

  rows.forEach(row => {
    const subjectRaw = norm(row[2] || ''); // C열
    if (!subjectRaw.includes('보호자')) return;

    const gradeRaw = String(row[10] || '').trim();   // K열 (자녀 학년, 중복 선택 가능: "2, 3" 등)
    const attend = String(row[14] || '').trim();      // O열
    const nameRaw = String(row[15] || '').trim();     // P열

    const grades = gradeRaw
      .split(/[,，\/]/)
      .map(g => g.trim())
      .filter(g => g);

    result.push({
      name: nameRaw ? nameRaw : '익명',
      attend: attend,
      grades: grades
    });
  });

  return result;
}

// =============================================
// 숨김 상태 저장 ('숨김처리' 탭, 설문지 응답 시트는 건드리지 않음)
// =============================================

function getHiddenSheet_(ss) {
  let sheet = ss.getSheetByName(HIDDEN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(HIDDEN_SHEET_NAME);
    sheet.getRange(1, 1, 1, 2).setValues([['id', 'hiddenAt']]);
  }
  // id 칼럼이 "6-1", "7-2" 같은 값을 날짜로 자동 변환해버리는 걸 막기 위해
  // 항상 일반 텍스트 서식으로 강제한다 (그렇지 않으면 문자열 비교가 깨진다).
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  return sheet;
}

function getHiddenSet_(ss) {
  const sheet = ss.getSheetByName(HIDDEN_SHEET_NAME);
  if (!sheet) return new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().map(r => String(r[0]));
  return new Set(ids);
}

function setHidden_(id, hide) {
  if (!id) throw new Error('id가 필요합니다.');

  // 다중 선택으로 여러 id를 동시에(병렬 요청으로) 처리할 때, 잠금 없이 같은
  // 시트의 행 번호를 동시에 읽고 쓰면 서로 덮어써서 엉뚱하게 토글되는 것처럼
  // 보이는 문제가 있었다. 시트 수정 구간을 잠궈서 한 번에 하나씩만 처리한다.
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = getHiddenSheet_(ss);

    // 과거 버그(브라우저 캐시)로 같은 id가 중복으로 쌓인 행이 있을 수 있으므로,
    // 매번 해당 id의 기존 행을 전부 지운 뒤 필요하면 한 줄만 새로 추가한다.
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
      for (let i = ids.length - 1; i >= 0; i--) {
        if (String(ids[i][0]) === String(id)) {
          sheet.deleteRow(i + 2);
        }
      }
    }

    if (hide) {
      const newRow = sheet.getLastRow() + 1;
      sheet.getRange(newRow, 1).setNumberFormat('@').setValue(id);
      sheet.getRange(newRow, 2).setValue(new Date());
    }
  } finally {
    lock.releaseLock();
  }
}

// 관리자용: 누적된 숨김 기록을 전부 비운다 (설문 응답 시트는 영향 없음).
function resetHidden_() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = getHiddenSheet_(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
  }
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
