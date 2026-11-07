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
  'getIslandProfile', 'saveIslandPet', 'getClassIslands',
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
// 🏝️ 내 친구 정보: 동물 종류·이름·꾸미기 (시트 'Islands'를 자동으로 만들어요)
// ---------------------------------------------------------------
// 이 스크립트가 구글 시트에 붙어 있지 않다면(독립 스크립트) 시트 ID를 넣어 주세요.
const ISLAND_SHEET_ID = '';
const ISLAND_COLS = ['className', 'studentNumber', 'studentName', 'pet', 'updatedAt'];
// 꾸미기 아이템 (화면의 OUTFIT_CATALOG와 같아요)
const OUTFIT_IDS = {
  head: ['sprout', 'cap', 'headphones', 'flowerCrown', 'crown'],
  face: ['roundGlasses', 'sunglasses', 'starGlasses'],
  neck: ['bowtie', 'bell', 'scarf', 'medal'],
  back: ['backpack', 'babyWings', 'ballPack', 'swordShield', 'batWings', 'jetpack', 'dragonWings'],
  feet: ['sneakers', 'skateboard', 'cloudRide', 'fireShoes', 'hoverboard'],
  buddy: ['cloudBuddy', 'bee', 'ghost', 'drone', 'babyStar', 'babyDragon'],
  effect: ['sparkle', 'notes', 'hearts', 'lightning', 'aura', 'fire']
};

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
    let row = i >= 0 ? rows[i].slice(0, ISLAND_COLS.length) : [className, Number(studentNumber), studentName || '', '', 0];
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
  return { pet: islandParse_(rows[i][3]), updatedAt: Number(rows[i][4]) || 0 };
}

function saveIslandPet(className, studentNumber, studentName, petJson) {
  const p = JSON.parse(petJson);
  if (!p || ['tiger', 'wolf', 'chick', 'bear', 'cat', 'quokka'].indexOf(p.species) < 0) throw new Error('친구 정보가 올바르지 않아요.');
  const outfit = {};
  Object.keys(OUTFIT_IDS).forEach(function (slot) {
    const id = p.outfit && p.outfit[slot];
    if (id && OUTFIT_IDS[slot].indexOf(id) >= 0) outfit[slot] = id;
  });
  const clean = JSON.stringify({ species: p.species, name: String(p.name || '').slice(0, 8), adoptedAt: Number(p.adoptedAt) || Date.now(), outfit: outfit });
  return islandUpdate_(className, studentNumber, studentName, function (row) { row[3] = clean; row[4] = Date.now(); return true; });
}

// 같은 반 친구 목록 (보스 체력을 정할 때 반 인원 수로 써요)
function getClassIslands(className) {
  const rows = islandRows_(), out = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] !== className) continue;
    out.push({ studentNumber: Number(rows[i][1]), studentName: rows[i][2], pet: islandParse_(rows[i][3]) });
  }
  return out;
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
