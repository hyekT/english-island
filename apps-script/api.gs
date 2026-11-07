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
  'addWord', 'deleteWord',
  'sendMail', 'submitMailReply', 'sendStudyLink',
  'startWordGame', 'submitWordGame',
  'verifyTeacherPassword', 'submitFeedback', 'updateFeedback', 'addBonus',
  'listStudentAccounts', 'resetStudentPassword',
  'getIslandProfile', 'saveIslandLayout', 'saveIslandPet', 'getClassIslands', 'likeIsland',
  'getBossAttacks', 'recordBossAttack'
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
// 🏝️ 내 섬 정보: 친구(동물)·꾸미기 배치·도장 (시트 'Islands'를 자동으로 만들어요)
// ---------------------------------------------------------------
// 이 스크립트가 구글 시트에 붙어 있지 않다면(독립 스크립트) 시트 ID를 넣어 주세요.
const ISLAND_SHEET_ID = '';
const ISLAND_COLS = ['className', 'studentNumber', 'studentName', 'pet', 'layout', 'updatedAt', 'likes', 'likedBy'];

function islandSheet_() {
  const ss = ISLAND_SHEET_ID ? SpreadsheetApp.openById(ISLAND_SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Islands');
  if (!sh) { sh = ss.insertSheet('Islands'); sh.appendRow(ISLAND_COLS); sh.setFrozenRows(1); }
  return sh;
}
function islandRows_() { return islandSheet_().getDataRange().getValues(); }
function islandFind_(rows, className, studentNumber) {
  for (let i = 1; i < rows.length; i++) if (rows[i][0] === className && String(rows[i][1]) === String(studentNumber)) return i;
  return -1;
}
function islandParse_(v) { try { return v ? JSON.parse(v) : null; } catch (e) { return null; } }
// 행을 찾아서(없으면 만들어서) fn(row)로 고친 뒤 저장해요
function islandUpdate_(className, studentNumber, studentName, fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = islandSheet_(), rows = sh.getDataRange().getValues();
    let i = islandFind_(rows, className, studentNumber);
    let row = i >= 0 ? rows[i].slice() : [className, Number(studentNumber), studentName || '', '', '', 0, 0, '{}'];
    while (row.length < ISLAND_COLS.length) row.push('');
    if (studentName) row[2] = studentName;
    const out = fn(row);
    if (i >= 0) sh.getRange(i + 1, 1, 1, ISLAND_COLS.length).setValues([row]);
    else sh.appendRow(row);
    return out;
  } finally {
    lock.releaseLock();
  }
}

function getIslandProfile(className, studentNumber) {
  const rows = islandRows_(), i = islandFind_(rows, className, studentNumber);
  if (i < 0) return null;
  const r = rows[i];
  return { pet: islandParse_(r[3]), layout: islandParse_(r[4]), updatedAt: Number(r[5]) || 0, likes: Number(r[6]) || 0 };
}

function saveIslandLayout(className, studentNumber, studentName, layoutJson) {
  if (String(layoutJson).length > 20000) throw new Error('배치 정보가 너무 커요.');
  const L = JSON.parse(layoutJson);
  if (!L || L.v !== 1) throw new Error('배치 정보가 올바르지 않아요.');
  return islandUpdate_(className, studentNumber, studentName, function (row) {
    row[4] = layoutJson; row[5] = Date.now(); return { updatedAt: row[5] };
  });
}

function saveIslandPet(className, studentNumber, studentName, petJson) {
  const p = JSON.parse(petJson);
  if (!p || ['tiger', 'wolf', 'chick', 'bear', 'cat', 'quokka'].indexOf(p.species) < 0) throw new Error('친구 정보가 올바르지 않아요.');
  const clean = JSON.stringify({ species: p.species, name: String(p.name || '').slice(0, 8), adoptedAt: Number(p.adoptedAt) || Date.now() });
  return islandUpdate_(className, studentNumber, studentName, function (row) { row[3] = clean; return true; });
}

function getClassIslands(className) {
  const rows = islandRows_(), out = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] !== className) continue;
    out.push({ studentNumber: Number(rows[i][1]), studentName: rows[i][2], pet: islandParse_(rows[i][3]), layout: islandParse_(rows[i][4]), likes: Number(rows[i][6]) || 0 });
  }
  return out;
}

// 친구 섬에 👍 도장: 한 친구 섬에 하루 한 번
function likeIsland(className, targetNumber, visitorNumber) {
  if (String(targetNumber) === String(visitorNumber)) throw new Error('내 섬에는 도장을 찍을 수 없어요.');
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const rows = islandRows_();
  if (islandFind_(rows, className, targetNumber) < 0) throw new Error('친구 섬을 찾지 못했어요.');
  return islandUpdate_(className, targetNumber, '', function (row) {
    const by = islandParse_(row[7]) || {};
    if (by[visitorNumber] === today) return { likes: Number(row[6]) || 0, already: true };
    by[visitorNumber] = today; row[7] = JSON.stringify(by); row[6] = (Number(row[6]) || 0) + 1;
    return { likes: row[6], already: false };
  });
}

// ---------------------------------------------------------------
// ⚔️ 보스 단어 대결 기록 (시트 'BossAttacks'를 자동으로 만들어요)
//   학생 한 명당 하루 1번, 한 번에 최대 20 데미지
// ---------------------------------------------------------------
const BOSS_COLS = ['className', 'studentNumber', 'studentName', 'date', 'weekKey', 'dmg', 'correct', 'createdAt'];
function bossSheet_() {
  const ss = ISLAND_SHEET_ID ? SpreadsheetApp.openById(ISLAND_SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('BossAttacks');
  if (!sh) { sh = ss.insertSheet('BossAttacks'); sh.appendRow(BOSS_COLS); sh.setFrozenRows(1); }
  return sh;
}
function bossDateStr_(v) {
  return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v);
}

function getBossAttacks(weekKey) {
  const rows = bossSheet_().getDataRange().getValues(), out = [];
  for (let i = 1; i < rows.length; i++) {
    if (bossDateStr_(rows[i][4]) !== String(weekKey)) continue;
    out.push({ className: rows[i][0], studentNumber: Number(rows[i][1]), studentName: rows[i][2], date: bossDateStr_(rows[i][3]), weekKey: String(weekKey), dmg: Number(rows[i][5]) || 0, correct: Number(rows[i][6]) || 0 });
  }
  return out;
}

function recordBossAttack(className, studentNumber, studentName, weekKey, dmg, correct) {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = bossSheet_(), rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === className && String(rows[i][1]) === String(studentNumber) && bossDateStr_(rows[i][3]) === today) {
        throw new Error('오늘은 이미 보스와 대결했어요. 내일 다시 도전해요!');
      }
    }
    const row = [className, Number(studentNumber), studentName, today, String(weekKey), Math.max(0, Math.min(20, Math.round(Number(dmg) || 0))), Number(correct) || 0, Date.now()];
    // 날짜가 자동으로 Date로 바뀌지 않게 글자로 저장해요
    sh.appendRow(row);
    sh.getRange(sh.getLastRow(), 4, 1, 2).setNumberFormat('@').setValues([[today, String(weekKey)]]);
    return { className: className, studentNumber: Number(studentNumber), studentName: studentName, date: today, weekKey: String(weekKey), dmg: row[5], correct: row[6] };
  } finally {
    lock.releaseLock();
  }
}
