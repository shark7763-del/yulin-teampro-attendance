/**
 * 育林國中體育班出席管理系統 2.0 ─ 教練戰情室版
 * 後端 Google Apps Script
 *
 * 設計重點：手機優先、4 狀態快速點名、今日戰情室、需要關心學生、三種 LINE 回報、今日/週/月報表。
 * 資料庫：Google Sheet（Students / Attendance / Users / 報表）
 *
 * 部署：
 *  1. 試算表 > 擴充功能 > Apps Script，貼上本檔與 index.html（檔名 index）。
 *  2. 第一次執行 setup() 建表與預設帳號；舊表升級執行 upgradeSchema()。
 *  3. 部署 > 網頁應用程式（執行身分：我；存取：任何人）。
 */

/* ===================== 設定 ===================== */
var SHEET_STUDENTS = 'Students';
var SHEET_ATT      = 'Attendance';
var SHEET_USERS    = 'Users';
var SHEET_REPORT   = '報表';

var SPECIALTIES = ['田徑', '跆拳道', '武術'];
var SESSIONS    = ['晨操', '專長訓練', '體育班課程', '暑期集訓', '晚自習', '道館訓練', '比賽', '移地訓練', '其他'];
var STATUSES    = ['出席', '遲到', '早退', '病假', '事假', '公假', '傷病觀察', '無故缺席', '賽事公假', '停訓', '移地訓練'];
var WEEKDAY     = ['日', '一', '二', '三', '四', '五', '六'];

// TeamPro Attendance 欄位。保留舊欄位「時段 / 姓名 / 專項 / 記錄者」以利舊版資料升級與相容。
var ATT_HEADERS = ['日期', '星期', '課程時段', '時段', '隊伍', '地點', '負責教練',
                   '學生ID', '學生姓名', '姓名', '年級', '班級', '項目', '專項',
                   '狀態', '備註', '請假原因', '遲到分鐘', '修改人', '記錄者',
                   '建立時間', '修改時間'];

function getSS() { return SpreadsheetApp.getActiveSpreadsheet(); }
function tz()    { return Session.getScriptTimeZone(); }

/* ===================== 網頁進入點 ===================== */
function doGet(e) {
  if (e && e.parameter && e.parameter.api === 'ping') {
    return json_({ ok: true, app: 'TeamPro 體育班出勤戰情室', time: new Date() });
  }
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('TeamPro 體育班出勤戰情室')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) {}
  var allow = {
    getUsers: getUsers,
    login: login,
    getStudents: getStudents,
    saveStudent: saveStudent,
    deleteStudent: deleteStudent,
    getAttendanceByDate: getAttendanceByDate,
    saveAttendance: saveAttendance,
    getDashboard: getDashboard,
    getCareList: getCareList,
    getReport: getReport,
    exportReport: exportReport,
    getStudentHistory: getStudentHistory,
    buildParentNotice: buildParentNotice
  };
  try {
    var fn = String(body.fn || '');
    if (!allow[fn]) return json_({ ok: false, error: 'API not allowed: ' + fn });
    var out = allow[fn].apply(null, body.args || []);
    return json_({ ok: true, data: out });
  } catch (err2) {
    return json_({ ok: false, error: String(err2 && err2.message ? err2.message : err2) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ===================== 初始化 ===================== */
function setup() {
  var ss = getSS();

  var st = getOrCreateSheet(ss, SHEET_STUDENTS);
  if (st.getLastRow() === 0) {
    st.appendRow(['學生ID', '姓名', '年級', '班級', '專項', '家長姓名', '家長聯絡方式', '家長查詢碼', '狀態', '晚自習報名']);
    formatHeader(st);
  }

  var at = getOrCreateSheet(ss, SHEET_ATT);
  if (at.getLastRow() === 0) { at.appendRow(ATT_HEADERS); formatHeader(at); }

  var us = getOrCreateSheet(ss, SHEET_USERS);
  if (us.getLastRow() === 0) {
    us.appendRow(['帳號', '密碼', '角色', '專項', '姓名']);
    us.appendRow(['admin', '9999', 'admin', '全部',   '系統管理員']);
    us.appendRow(['track', '1111', 'coach', '田徑',   '田徑教練']);
    us.appendRow(['tkd',   '2222', 'coach', '跆拳道', '跆拳道教練']);
    us.appendRow(['wushu', '3333', 'coach', '武術',   '武術教練']);
    formatHeader(us);
  }
  return '初始化完成！登入 PIN：admin=9999、track=1111、tkd=2222、wushu=3333';
}

/** 舊表升級（保留資料，只補欄位／工作表）。改版後執行一次。 */
function upgradeSchema() {
  var ss = getSS(), msg = [];
  var st = ss.getSheetByName(SHEET_STUDENTS);
  if (st) {
    var sh = st.getRange(1, 1, 1, Math.max(st.getLastColumn(), 10)).getValues()[0];
    if (sh[9] !== '晚自習報名') { st.getRange(1, 10).setValue('晚自習報名'); msg.push('Students 補「晚自習報名」'); }
    var filled = fillMissingParentCodes();
    if (filled.updated) msg.push('Students 補「家長查詢碼」' + filled.updated + ' 筆');
  }
  var at = ss.getSheetByName(SHEET_ATT);
  if (!at) { at = getOrCreateSheet(ss, SHEET_ATT); at.appendRow(ATT_HEADERS); formatHeader(at); msg.push('建立 Attendance'); }
  else {
    var hrow = at.getRange(1, 1, 1, at.getLastColumn()).getValues()[0];
    ATT_HEADERS.forEach(function (h) {
      if (hrow.indexOf(h) < 0) { at.getRange(1, at.getLastColumn() + 1).setValue(h); msg.push('Attendance 補「' + h + '」'); }
    });
    formatHeader(at);
  }
  return msg.length ? ('升級完成：' + msg.join('、')) : '結構已是最新，無需升級。';
}

function getOrCreateSheet(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}
function formatHeader(sh) {
  var r = sh.getRange(1, 1, 1, sh.getLastColumn());
  r.setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
  sh.setFrozenRows(1);
}
function genCode() {
  var s = '', chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  for (var i = 0; i < 6; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}
function makeStudentId(s, existing) {
  var cls = String(s.cls || '').replace(/\D/g, '') || String(s.grade || '').replace(/\D/g, '') || 'STU';
  var name = String(s.name || '').replace(/\s/g, '');
  var seed = 0;
  for (var i = 0; i < name.length; i++) seed += name.charCodeAt(i);
  var suffix = ('00' + (seed % 100)).slice(-2);
  var base = cls + '-' + suffix;
  var id = base, n = 1;
  while (existing[id]) {
    id = base + '-' + (++n);
  }
  return id;
}

/* ===================== 一鍵匯入名單 ===================== */
function importRoster() {
  var roster = [
    ['701','02','田沂昊','田徑'],['701','03','李程皓','田徑'],['701','04','林駿堯','田徑'],
    ['701','05','洪翌智','田徑'],['701','08','黃紹誌','田徑'],['701','10','董以杰','田徑'],
    ['701','13','鄭添鴻','田徑'],['701','14','盧宥熙','田徑'],['701','21','王怡方','田徑'],
    ['701','22','王翎伊','田徑'],['701','23','田凡鋗','田徑'],['701','24','朱璇晴','田徑'],
    ['701','28','陳薈欣','田徑'],['701','29','楊喬甯','田徑'],['701','30','盧品瑄','田徑'],
    ['701','31','簡宜慧','田徑'],
    ['801','02','王品睿','田徑'],['801','03','李明潔','田徑'],['801','04','張師豪','田徑'],
    ['801','05','張澤羲','田徑'],['801','06','陳炫希','田徑'],['801','08','蔡宸睿','田徑'],
    ['801','09','盧彥勳','田徑'],['801','11','羅承瀚','田徑'],['801','21','吳宜蓁','田徑'],
    ['801','25','陳霈真','田徑'],['801','26','鄭琇云','田徑'],['801','27','蕭蘿槿','田徑'],
    ['701','01','王柏鈞','跆拳道'],['701','06','許景皓','跆拳道'],['701','15','上官哲忻','跆拳道'],
    ['701','25','徐洧翎','跆拳道'],['701','26','張晏慈','跆拳道'],['701','27','曹絜綺','跆拳道'],
    ['801','01','王冠霖','跆拳道'],['801','07','葉承祐','跆拳道'],['801','10','謝昊恩','跆拳道'],
    ['801','12','蘇宥嘉','跆拳道'],['801','22','吳昀蓁','跆拳道'],['801','23','林子棠','跆拳道'],
    ['801','24','唐霈昕','跆拳道']
  ];
  var sh = getSS().getSheetByName(SHEET_STUDENTS);
  if (!sh) { setup(); sh = getSS().getSheetByName(SHEET_STUDENTS); }
  var data = sh.getDataRange().getValues(), existing = {};
  for (var i = 1; i < data.length; i++) existing[String(data[i][0])] = true;
  var added = 0, skipped = 0;
  roster.forEach(function (r) {
    var id = r[0] + '-' + r[1];
    if (existing[id]) { skipped++; return; }
    sh.appendRow([id, r[2], r[0].charAt(0) + '年級', r[0], r[3], '', '', genCode(), '在學', '']);
    added++;
  });
  return '匯入完成！新增 ' + added + ' 人，略過 ' + skipped + ' 人。';
}

/* ===================== 登入 ===================== */
function getUsers() {
  var sh = getSS().getSheetByName(SHEET_USERS), data = sh.getDataRange().getValues(), list = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    list.push({ account: data[i][0], role: data[i][2], specialty: data[i][3], name: data[i][4] });
  }
  return list;
}
function resetDefaultPins() {
  var pins = { admin: '9999', track: '1111', tkd: '2222', wushu: '3333' };
  var sh = getSS().getSheetByName(SHEET_USERS), data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var acc = String(data[i][0]).trim(); if (!acc) continue;
    sh.getRange(i + 1, 2).setValue(pins[acc] || '0000');
  }
  return '已重設 PIN：admin=9999、track=1111、tkd=2222、wushu=3333，其餘=0000。';
}
function login(account, password) {
  var sh = getSS().getSheetByName(SHEET_USERS), data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(account).trim() && String(data[i][1]) === String(password)) {
      return { ok: true, account: data[i][0], role: data[i][2], specialty: data[i][3], name: data[i][4] };
    }
  }
  return { ok: false, message: '帳號或 PIN 錯誤' };
}

/* ===================== 學生資料 ===================== */
function getStudents(specialty) {
  var sh = getSS().getSheetByName(SHEET_STUDENTS), data = sh.getDataRange().getValues(), list = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i]; if (!row[0]) continue;
    if (specialty && specialty !== '全部' && row[4] !== specialty) continue;
    list.push({
      id: row[0], name: row[1], grade: row[2], cls: row[3], specialty: row[4],
      parentName: row[5], parentContact: row[6], parentCode: row[7], status: row[8],
      studyHall: row[9] || ''
    });
  }
  return list;
}
function saveStudent(s) {
  var sh = getSS().getSheetByName(SHEET_STUDENTS), data = sh.getDataRange().getValues();
  var existing = {};
  for (var e = 1; e < data.length; e++) if (data[e][0]) existing[String(data[e][0])] = true;
  if (!String(s.id || '').trim()) s.id = makeStudentId(s, existing);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(s.id)) {
      var parentCode = String(s.parentCode || data[i][7] || '').trim() || genCode();
      sh.getRange(i + 1, 1, 1, 10).setValues([[
        s.id, s.name, s.grade, s.cls, s.specialty, s.parentName, s.parentContact,
        parentCode, s.status || '在學', s.studyHall || '']]);
      return { ok: true, message: '已更新' };
    }
  }
  var newParentCode = String(s.parentCode || '').trim() || genCode();
  sh.appendRow([s.id, s.name, s.grade, s.cls, s.specialty, s.parentName, s.parentContact,
                newParentCode, s.status || '在學', s.studyHall || '']);
  return { ok: true, message: '已新增' };
}

function fillMissingParentCodes() {
  var sh = getSS().getSheetByName(SHEET_STUDENTS), data = sh.getDataRange().getValues();
  if (data.length < 2) return { ok: true, updated: 0 };
  var used = {};
  for (var i = 1; i < data.length; i++) {
    var code = String(data[i][7] || '').trim();
    if (code) used[code] = true;
  }
  var updated = 0;
  for (var r = 1; r < data.length; r++) {
    var code2 = String(data[r][7] || '').trim();
    if (code2) continue;
    var next = genCode();
    while (used[next]) next = genCode();
    sh.getRange(r + 1, 8).setValue(next);
    used[next] = true;
    updated++;
  }
  return { ok: true, updated: updated };
}
function deleteStudent(id) {
  var sh = getSS().getSheetByName(SHEET_STUDENTS), data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) { sh.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false, message: '找不到學生' };
}

/* ===================== TeamPro 點名讀寫 / 統計 ===================== */
function hidx(headerRow) { var H = {}; headerRow.forEach(function (n, i) { H[String(n)] = i; }); return H; }
function hget(row, H, names) {
  for (var i = 0; i < names.length; i++) if (H[names[i]] != null) return row[H[names[i]]];
  return '';
}
function hset(row, H, names, value) {
  names.forEach(function (n) { if (H[n] != null) row[H[n]] = value; });
}
function normalizeStatus(st) {
  if (st === '請假') return '事假';
  if (st === '未到') return '無故缺席';
  if (st === '比賽') return '賽事公假';
  if (STATUSES.indexOf(st) >= 0) return st;
  return st || '';
}
function isActual(st) { return ['出席', '遲到', '早退'].indexOf(normalizeStatus(st)) >= 0; }
function isExcluded(st) { return ['公假', '賽事公假', '移地訓練', '停訓'].indexOf(normalizeStatus(st)) >= 0; }
function isLeave(st) { return ['病假', '事假', '公假', '賽事公假', '傷病觀察', '停訓', '移地訓練'].indexOf(normalizeStatus(st)) >= 0; }
function isAbsent(st) { return normalizeStatus(st) === '無故缺席'; }
function isAbnormal(st) { return ['遲到', '早退', '病假', '事假', '無故缺席'].indexOf(normalizeStatus(st)) >= 0; }
function statBucket(st) {
  st = normalizeStatus(st);
  return {
    actual: isActual(st) ? 1 : 0,
    late: st === '遲到' ? 1 : 0,
    early: st === '早退' ? 1 : 0,
    leave: isLeave(st) ? 1 : 0,
    absent: isAbsent(st) ? 1 : 0,
    excluded: isExcluded(st) ? 1 : 0,
    abnormal: isAbnormal(st) ? 1 : 0,
    severe: isAbsent(st) ? 1 : 0,
    attended: isActual(st) ? 1 : 0,
    expected: isExcluded(st) ? 0 : 1
  };
}

// 該時段「應到」學生（晚自習只算當天有報名者）
function eligibleStudents(specialty, dateStr, session) {
  var list = getStudents(specialty).filter(function (s) { return s.status === '在學'; });
  if (session === '晚自習') {
    var wc = weekdayOf(dateStr);
    list = list.filter(function (s) { return (s.studyHall || '').indexOf(wc) >= 0; });
  }
  return list;
}

// 回傳 { 學生ID: {status, note, reason, lateMin, location, coach, team, modifiedBy} }
function getAttendanceByDate(date, session, specialty) {
  var sh = getSS().getSheetByName(SHEET_ATT), data = sh.getDataRange().getValues();
  if (data.length < 2) return {};
  var H = hidx(data[0]), map = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i], rowSession = hget(row, H, ['課程時段', '時段']);
    if (formatDate(hget(row, H, ['日期'])) !== date || rowSession !== session) continue;
    if (specialty && specialty !== '全部' && hget(row, H, ['項目', '專項']) !== specialty) continue;
    map[String(hget(row, H, ['學生ID']))] = {
      status: normalizeStatus(hget(row, H, ['狀態'])),
      note: hget(row, H, ['備註']) || '',
      reason: hget(row, H, ['請假原因']) || '',
      lateMin: hget(row, H, ['遲到分鐘']) || '',
      location: hget(row, H, ['地點']) || '',
      coach: hget(row, H, ['負責教練', '記錄者']) || '',
      team: hget(row, H, ['隊伍']) || '',
      modifiedBy: hget(row, H, ['修改人']) || ''
    };
  }
  return map;
}

/**
 * 儲存點名（同日期 + 課程時段 + 學生ID 僅更新一列）
 * payload: { date, session, team, location, coach, recorder, records:[{id,name,grade,cls,specialty,status,note,reason,lateMin}] }
 */
function saveAttendance(payload) {
  var sh = getSS().getSheetByName(SHEET_ATT);
  if (!sh) { setup(); sh = getSS().getSheetByName(SHEET_ATT); }
  upgradeSchema();
  var data = sh.getDataRange().getValues(), H = hidx(data[0]), ncol = sh.getLastColumn();
  var date = payload.date, session = payload.session, now = new Date();
  var coach = payload.coach || payload.recorder || '', recorder = payload.recorder || coach;
  var team = payload.team || payload.specialty || '', location = payload.location || '';
  var index = {};
  for (var i = 1; i < data.length; i++) {
    if (formatDate(hget(data[i], H, ['日期'])) === date && hget(data[i], H, ['課程時段', '時段']) === session) {
      index[String(hget(data[i], H, ['學生ID']))] = i + 1;
    }
  }
  var added = 0, updated = 0;
  (payload.records || []).forEach(function (r) {
    var st = normalizeStatus(r.status);
    if (!r.id || !st) return;
    var rowNum = index[String(r.id)];
    var arr = rowNum ? data[rowNum - 1].slice() : new Array(ncol);
    for (var j = 0; j < ncol; j++) if (arr[j] == null) arr[j] = '';
    hset(arr, H, ['日期'], date);
    hset(arr, H, ['星期'], weekdayOf(date));
    hset(arr, H, ['課程時段', '時段'], session);
    hset(arr, H, ['隊伍'], team || r.specialty || '');
    hset(arr, H, ['地點'], location || r.location || '');
    hset(arr, H, ['負責教練'], coach);
    hset(arr, H, ['學生ID'], r.id);
    hset(arr, H, ['學生姓名', '姓名'], r.name || '');
    hset(arr, H, ['年級'], r.grade || '');
    hset(arr, H, ['班級'], r.cls || '');
    hset(arr, H, ['項目', '專項'], r.specialty || team || '');
    hset(arr, H, ['狀態'], st);
    hset(arr, H, ['請假原因'], r.reason || '');
    hset(arr, H, ['遲到分鐘'], st === '遲到' ? (r.lateMin || '') : '');
    hset(arr, H, ['備註'], r.note || '');
    hset(arr, H, ['修改人', '記錄者'], recorder);
    hset(arr, H, ['修改時間'], now);
    if (rowNum) {
      sh.getRange(rowNum, 1, 1, ncol).setValues([arr]);
      updated++;
    } else {
      hset(arr, H, ['建立時間'], now);
      sh.appendRow(arr);
      added++;
    }
  });
  return { ok: true, added: added, updated: updated, lineText: buildLineReport(date, session, team || payload.specialty || '全部') };
}

/* ===================== 今日戰情室 ===================== */
function getDashboard(specialty, dateStr, session) {
  dateStr = dateStr || todayServer();
  session = session || currentSession();
  var students = eligibleStudents(specialty, dateStr, session);
  var map = getAttendanceByDate(dateStr, session, specialty);
  var c = { actual: 0, late: 0, early: 0, leave: 0, absent: 0, excluded: 0, abnormal: 0, severe: 0, marked: 0 };
  students.forEach(function (s) {
    var st = map[s.id] && map[s.id].status;
    if (!st) return;
    c.marked++;
    var b = statBucket(st);
    c.actual += b.actual; c.late += b.late; c.early += b.early; c.leave += b.leave;
    c.absent += b.absent; c.excluded += b.excluded; c.abnormal += b.abnormal; c.severe += b.severe;
  });
  var denom = students.length - c.excluded;
  return {
    date: dateStr, weekday: weekdayOf(dateStr), session: session,
    expected: students.length, countedExpected: denom, actual: c.actual, present: c.actual,
    late: c.late, early: c.early, leave: c.leave, absent: c.absent, excluded: c.excluded,
    abnormal: c.abnormal, severe: c.severe, marked: c.marked, unmarked: students.length - c.marked,
    rate: denom > 0 ? Math.round(c.actual / denom * 100) : null
  };
}

/* ===================== 需要關心學生 ===================== */
function getCareList(specialty) {
  var sh = getSS().getSheetByName(SHEET_ATT), data = sh.getDataRange().getValues();
  var meta = {}; getStudents(specialty).forEach(function (s) { if (s.status === '在學') meta[s.id] = s; });
  if (data.length < 2) return [];
  var H = hidx(data[0]), today = todayServer(), monthPref = today.substring(0, 7), byStu = {};
  for (var i = 1; i < data.length; i++) {
    var id = String(hget(data[i], H, ['學生ID']));
    if (!meta[id]) continue;
    if (!byStu[id]) byStu[id] = [];
    byStu[id].push({
      date: formatDate(hget(data[i], H, ['日期'])),
      session: hget(data[i], H, ['課程時段', '時段']),
      status: normalizeStatus(hget(data[i], H, ['狀態']))
    });
  }
  var care = [];
  Object.keys(meta).forEach(function (id) {
    var recs = (byStu[id] || []).sort(function (a, b) { return (a.date + a.session) < (b.date + b.session) ? -1 : 1; });
    var consecAbs = 0, monthLate = 0, obsStart = null, monthActual = 0, monthExpected = 0, severeToday = false;
    for (var j = recs.length - 1; j >= 0; j--) { if (recs[j].status === '無故缺席') consecAbs++; else break; }
    recs.forEach(function (r) {
      if (r.date.indexOf(monthPref) === 0) {
        if (r.status === '遲到') monthLate++;
        var b = statBucket(r.status); monthActual += b.attended; monthExpected += b.expected;
      }
      if (r.status === '無故缺席' && r.date === today) severeToday = true;
      if (r.status === '傷病觀察') { if (!obsStart || r.date < obsStart) obsStart = r.date; }
    });
    var rate = monthExpected > 0 ? Math.round(monthActual / monthExpected * 100) : null;
    var reasons = [];
    if (consecAbs >= 2) reasons.push('連續缺席 ' + consecAbs + ' 次');
    if (monthLate >= 3) reasons.push('本月遲到 ' + monthLate + ' 次');
    if (rate !== null && monthExpected >= 3 && rate < 80) reasons.push('本月出席率 ' + rate + '%');
    if (obsStart && daysBetween(obsStart, today) > 7) reasons.push('傷病觀察超過 7 天');
    if (severeToday) reasons.push('今日無故缺席');
    if (reasons.length) {
      var s = meta[id];
      care.push({ id: id, name: s.name, grade: s.grade, cls: s.cls, specialty: s.specialty, parentName: s.parentName, reasons: reasons, rate: rate });
    }
  });
  care.sort(function (a, b) { return b.reasons.length - a.reasons.length; });
  return care;
}

/* ===================== 報表（今日 / 本週 / 本月） ===================== */
function rangeBounds(range, refDate) {
  refDate = refDate || todayServer();
  if (range === 'week')  { var mon = mondayOf(refDate); return { from: mon, to: shiftDate(mon, 6) }; }
  if (range === 'month') {
    var pref = refDate.substring(0, 7), p = pref.split('-');
    var last = new Date(Number(p[0]), Number(p[1]), 0).getDate();
    return { from: pref + '-01', to: pref + '-' + ('0' + last).slice(-2) };
  }
  return { from: refDate, to: refDate };
}
function emptyStudentStat(s) {
  return { id: s.id, name: s.name, grade: s.grade, cls: s.cls, specialty: s.specialty,
    total: 0, present: 0, actual: 0, late: 0, early: 0, leave: 0, absent: 0,
    excluded: 0, abnormal: 0, severe: 0, rate: null, note: '' };
}
function getReport(range, specialty, refDate) {
  var b = rangeBounds(range, refDate), meta = {}, stat = {};
  getStudents(specialty).forEach(function (s) { if (s.status === '在學') { meta[s.id] = s; stat[s.id] = emptyStudentStat(s); } });
  var sh = getSS().getSheetByName(SHEET_ATT), data = sh.getDataRange().getValues();
  if (data.length >= 2) {
    var H = hidx(data[0]);
    for (var i = 1; i < data.length; i++) {
      var id = String(hget(data[i], H, ['學生ID'])); if (!stat[id]) continue;
      var d = formatDate(hget(data[i], H, ['日期'])); if (d < b.from || d > b.to) continue;
      var st = normalizeStatus(hget(data[i], H, ['狀態'])), x = stat[id], bk = statBucket(st);
      x.total += bk.expected; x.present += st === '出席' ? 1 : 0; x.actual += bk.actual;
      x.late += bk.late; x.early += bk.early; x.leave += bk.leave; x.absent += bk.absent;
      x.excluded += bk.excluded; x.abnormal += bk.abnormal; x.severe += bk.severe;
    }
  }
  var sum = { total: 0, actual: 0, present: 0, late: 0, early: 0, leave: 0, absent: 0, excluded: 0, abnormal: 0, severe: 0 };
  var students = Object.keys(stat).map(function (id) {
    var x = stat[id];
    x.rate = x.total > 0 ? Math.round(x.actual / x.total * 100) : null;
    x.light = lightOf(x.rate);
    Object.keys(sum).forEach(function (k) { sum[k] += x[k] || 0; });
    return x;
  });
  var teamRate = sum.total > 0 ? Math.round(sum.actual / sum.total * 100) : null;
  var absentRank = students.filter(function (s) { return s.absent > 0; }).sort(function (a, b2) { return b2.absent - a.absent; }).slice(0, 10);
  var lateRank = students.filter(function (s) { return s.late > 0; }).sort(function (a, b2) { return b2.late - a.late; }).slice(0, 10);
  var abnormalStudents = students.filter(function (s) { return s.abnormal > 0 || (s.rate !== null && s.rate < 80); })
    .sort(function (a, b2) { return (b2.abnormal - a.abnormal) || (a.rate || 999) - (b2.rate || 999); });
  students.sort(function (a, b2) { return (a.rate === null ? 999 : a.rate) - (b2.rate === null ? 999 : b2.rate); });
  return { range: range, from: b.from, to: b.to, summary: sum, teamRate: teamRate, students: students,
           absentRank: absentRank, lateRank: lateRank, abnormalStudents: abnormalStudents };
}
function getTeamRates(specialty, refDate) {
  return {
    today: getReport('today', specialty, refDate).teamRate,
    week: getReport('week', specialty, refDate).teamRate,
    month: getReport('month', specialty, refDate).teamRate
  };
}
function lightOf(r) { if (r === null) return 'gray'; if (r >= 90) return 'green'; if (r >= 80) return 'yellow'; return 'red'; }

// 匯出報表到「報表」工作表
function exportReport(range, specialty, refDate, reportType) {
  var rep = getReport(range || 'month', specialty, refDate), ss = getSS(), sh = getOrCreateSheet(ss, SHEET_REPORT);
  sh.clear();
  reportType = reportType || '全隊出勤月報';
  sh.appendRow(['TeamPro 體育班出勤戰情室｜' + reportType + '｜' + rep.from + ' ~ ' + rep.to + '｜隊伍：' + specialty + '｜全隊出席率：' + (rep.teamRate === null ? '-' : rep.teamRate + '%')]);
  sh.appendRow(['學生姓名', '年級', '班級', '項目', '總課程數', '出席次數', '遲到次數', '請假次數', '無故缺席次數', '出席率', '備註']);
  formatHeader2(sh, 2, 11);
  rep.students.forEach(function (s) {
    var note = [];
    if (s.excluded) note.push('合理排除 ' + s.excluded);
    if (s.abnormal) note.push('異常 ' + s.abnormal);
    sh.appendRow([s.name, s.grade, s.cls, s.specialty, s.total, s.present, s.late, s.leave, s.absent,
                  s.rate === null ? '-' : s.rate + '%', note.join('；')]);
  });
  return { ok: true, count: rep.students.length, url: ss.getUrl() };
}
function formatHeader2(sh, row, cols) {
  var r = sh.getRange(row, 1, 1, cols || sh.getLastColumn());
  r.setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
}

// 個別學生出席紀錄（近 limit 筆）
function getStudentHistory(id, limit) {
  limit = limit || 30;
  var sh = getSS().getSheetByName(SHEET_ATT), data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  var H = hidx(data[0]), rows = [];
  for (var i = 1; i < data.length; i++) {
    if (String(hget(data[i], H, ['學生ID'])) !== String(id)) continue;
    rows.push({
      date: formatDate(hget(data[i], H, ['日期'])), weekday: hget(data[i], H, ['星期']),
      session: hget(data[i], H, ['課程時段', '時段']), status: normalizeStatus(hget(data[i], H, ['狀態'])),
      reason: hget(data[i], H, ['請假原因']) || '', lateMin: hget(data[i], H, ['遲到分鐘']) || '',
      note: hget(data[i], H, ['備註']) || '', modifiedBy: hget(data[i], H, ['修改人', '記錄者']) || ''
    });
  }
  rows.sort(function (a, b) { return (a.date + a.session) < (b.date + b.session) ? 1 : -1; });
  return rows.slice(0, limit);
}

function buildLineReport(date, session, specialty) {
  var students = eligibleStudents(specialty, date, session), map = getAttendanceByDate(date, session, specialty);
  var c = { expected: students.length, actual: 0, late: 0, leave: 0, absent: 0, severe: 0, lateNames: [], leaveNames: [], absentNames: [] };
  students.forEach(function (s) {
    var rec = map[s.id], st = rec && rec.status;
    if (!st) return;
    if (isActual(st)) c.actual++;
    if (st === '遲到') { c.late++; c.lateNames.push(s.name); }
    if (isLeave(st)) { c.leave++; c.leaveNames.push(s.name + (rec.note ? '（' + rec.note + '）' : '（' + st + '）')); }
    if (isAbsent(st)) { c.absent++; c.severe++; c.absentNames.push(s.name + (rec.note ? '（' + rec.note + '）' : '')); }
  });
  return '育林體育班｜' + date + ' ' + session + ' 點名\n\n' +
    '應到：' + c.expected + '人\n' +
    '實到：' + c.actual + '人\n' +
    '遲到：' + c.late + '人\n' +
    '請假：' + c.leave + '人\n' +
    '缺席：' + c.absent + '人\n' +
    '無故缺席：' + c.severe + '人\n\n' +
    '遲到：\n' + (c.lateNames.length ? c.lateNames.join('、') : '無') + '\n\n' +
    '請假：\n' + (c.leaveNames.length ? c.leaveNames.join('、') : '無') + '\n\n' +
    '缺席：\n' + (c.absentNames.length ? c.absentNames.join('、') : '無') + '\n\n' +
    '今日點名已完成。';
}

function buildParentNotice(studentId, status, date, session) {
  var s = null, list = getStudents('全部');
  list.forEach(function (x) { if (String(x.id) === String(studentId)) s = x; });
  var count = 0, pref = (date || todayServer()).substring(0, 7), sh = getSS().getSheetByName(SHEET_ATT), data = sh.getDataRange().getValues();
  if (sh && data.length >= 2) {
    var H = hidx(data[0]);
    for (var i = 1; i < data.length; i++) {
      if (String(hget(data[i], H, ['學生ID'])) === String(studentId) &&
          formatDate(hget(data[i], H, ['日期'])).indexOf(pref) === 0 &&
          normalizeStatus(hget(data[i], H, ['狀態'])) === normalizeStatus(status)) count++;
    }
  }
  return '家長您好，今日 ' + (date || todayServer()) + ' ' + (session || currentSession()) + '，孩子 ' +
    (s ? s.name : '') + ' 紀錄為 ' + normalizeStatus(status) + '。本月累計 ' + normalizeStatus(status) + ' ' + count +
    ' 次，教練會協助孩子調整，也請家長一起留意。';
}

/* ===================== 工具 ===================== */
function formatDate(d) {
  if (!d) return '';
  if (typeof d === 'string') return d.length >= 10 ? d.substring(0, 10) : d;
  return Utilities.formatDate(d, tz(), 'yyyy-MM-dd');
}
function todayServer() { return Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd'); }
function currentSession() {
  var h = Number(Utilities.formatDate(new Date(), tz(), 'H'));
  if (h < 12) return '晨操';
  if (h < 18) return '專長訓練';
  return '晚自習';
}
function weekdayOf(dateStr) { return WEEKDAY[new Date(dateStr + 'T12:00:00').getDay()]; }
function shiftDate(dateStr, days) {
  var d = new Date(dateStr + 'T12:00:00'); d.setDate(d.getDate() + days);
  return Utilities.formatDate(d, tz(), 'yyyy-MM-dd');
}
function mondayOf(dateStr) {
  var d = new Date(dateStr + 'T12:00:00'), day = d.getDay(), diff = (day === 0 ? -6 : 1 - day);
  return shiftDate(dateStr, diff);
}
