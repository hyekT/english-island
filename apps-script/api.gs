/**
 * 🏝️ 영어 학습 섬 — Vercel 화면이 부르는 데이터 API
 *
 * 사용법 (내년용으로 새로 준비할 때)
 *  1) 지금 쓰는 Apps Script 프로젝트와 구글 시트를 '사본 만들기'로 복사해요. (올해 것은 그대로!)
 *  2) 복사한 프로젝트에 이 파일(api.gs)을 새 스크립트 파일로 추가해요.
 *  3) 배포 → 새 배포 → 유형 '웹 앱' → 실행: 나 / 액세스: 모든 사용자 → 배포
 *  4) 나온 웹 앱 주소(/exec로 끝나요)를 index.html의 APP_CONFIG.API_URL에 넣어요.
 *
 * 화면이 { fn: '함수이름', args: [...] } 를 보내면, 같은 프로젝트에 있는 그 함수를 실행해서 결과를 돌려줘요.
 * 기존 함수(getAllData, studentLogin, addEntry …)는 하나도 고칠 필요가 없어요.
 */

// 화면에서 부를 수 있는 함수 목록 (이 목록에 없는 함수는 실행되지 않아요)
const API_ALLOWED = [
  'getAllData', 'studentLogin', 'addClassName', 'addEntry',
  'addWord', 'deleteWord', 'lookupWord',
  'sendMail', 'submitMailReply', 'sendStudyLink',
  'startWordGame', 'submitWordGame',
  'verifyTeacherPassword', 'submitFeedback', 'updateFeedback', 'addBonus',
  'listStudentAccounts', 'resetStudentPassword',
  'getIslandLayout', 'saveIslandLayout'
];

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    if (API_ALLOWED.indexOf(req.fn) < 0) throw new Error('허용되지 않은 요청이에요: ' + req.fn);
    const fn = globalThis[req.fn];
    if (typeof fn !== 'function') throw new Error('서버에 ' + req.fn + ' 함수가 없어요.');
    out = { ok: true, data: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: (err && err.message) ? err.message : String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------
// 🛠️ 섬 꾸미기 배치 저장 (시트 'IslandLayouts'를 자동으로 만들어요)
// ---------------------------------------------------------------
// 이 스크립트가 구글 시트에 붙어 있지 않다면(독립 스크립트) 시트 ID를 넣어 주세요.
const ISLAND_SHEET_ID = '';

function islandLayoutSheet_() {
  const ss = ISLAND_SHEET_ID ? SpreadsheetApp.openById(ISLAND_SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('IslandLayouts');
  if (!sh) {
    sh = ss.insertSheet('IslandLayouts');
    sh.appendRow(['className', 'studentNumber', 'layout', 'updatedAt']);
  }
  return sh;
}

function getIslandLayout(className, studentNumber) {
  const rows = islandLayoutSheet_().getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === className && String(rows[i][1]) === String(studentNumber)) {
      return { layout: JSON.parse(rows[i][2] || 'null'), updatedAt: Number(rows[i][3]) || 0 };
    }
  }
  return null;
}

function saveIslandLayout(className, studentNumber, layoutJson) {
  if (String(layoutJson).length > 20000) throw new Error('배치 정보가 너무 커요.');
  const L = JSON.parse(layoutJson);
  if (!L || L.v !== 1) throw new Error('배치 정보가 올바르지 않아요.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = islandLayoutSheet_();
    const rows = sh.getDataRange().getValues();
    const now = Date.now();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === className && String(rows[i][1]) === String(studentNumber)) {
        sh.getRange(i + 1, 3, 1, 2).setValues([[layoutJson, now]]);
        return { updatedAt: now };
      }
    }
    sh.appendRow([className, Number(studentNumber), layoutJson, now]);
    return { updatedAt: now };
  } finally {
    lock.releaseLock();
  }
}
