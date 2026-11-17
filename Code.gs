/**
 * 앱 사랑방 — 구글 시트 서버 (Google Apps Script)
 * 1) 새 구글 시트 → 확장 프로그램 → Apps Script → 이 코드를 통째로 붙여넣기
 * 2) 위쪽 함수 목록에서 setup 선택 → ▶실행 (권한 허용)
 * 3) 배포 → 새 배포 → 유형: 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → /exec 주소를 config.js에 넣기
 * 글 숨기기: 각 시트의 '숨김' 칸에 Y 를 적으면 앱에서 보이지 않습니다.
 */
var TZ = 'Asia/Seoul';
var T = {
  apps:      {name:'앱',       cols:[['id','번호'],['date','등록일'],['name','앱 이름'],['desc','한 줄 설명'],['cat','분야'],['status','상태'],['url','주소'],['maker','만든 분'],['howto','사용법'],['shot1','스크린샷1'],['shot2','스크린샷2'],['shot3','스크린샷3'],['story','이렇게 만들었어요'],['tool','AI 도구'],['hours','걸린 시간'],['prompt','사용한 프롬프트'],['likes','공감'],['pw','비밀번호(암호화)'],['hidden','숨김']]},
  updates:   {name:'업데이트', cols:[['id','번호'],['appId','앱 번호'],['date','날짜'],['ver','버전'],['text','내용'],['hidden','숨김']]},
  feedback:  {name:'피드백',   cols:[['id','번호'],['appId','앱 번호'],['date','날짜'],['type','유형'],['author','작성자'],['text','내용'],['likes','공감'],['done','반영'],['doneNote','반영 메모'],['hidden','숨김']]},
  ideas:     {name:'아이디어', cols:[['id','번호'],['date','날짜'],['title','제목'],['text','내용'],['author','작성자'],['likes','나도 필요해요'],['status','상태'],['builder','만드는 분'],['hidden','숨김']]},
  topics:    {name:'강의주제', cols:[['id','번호'],['date','날짜'],['title','주제'],['text','내용'],['author','신청자'],['likes','투표'],['hidden','숨김']]},
  demos:     {name:'시연신청', cols:[['id','번호'],['date','신청일'],['app','앱 이름'],['presenter','발표자'],['when','희망 일정'],['note','메모'],['hidden','숨김']]},
  questions: {name:'질문',     cols:[['id','번호'],['date','날짜'],['title','제목'],['text','내용'],['author','작성자'],['solved','해결'],['hidden','숨김']]},
  answers:   {name:'답변',     cols:[['id','번호'],['qId','질문 번호'],['date','날짜'],['author','작성자'],['text','내용'],['likes','도움됐어요'],['hidden','숨김']]}
};
var LIKEABLE = {apps:1, feedback:1, ideas:1, topics:1, answers:1};

/* ---------- 처음 한 번 실행 ---------- */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(T).forEach(function (k) {
    var sh = ss.getSheetByName(T[k].name) || ss.insertSheet(T[k].name);
    var heads = T[k].cols.map(function (c) { return c[1]; });
    sh.getRange(1, 1, 1, heads.length).setValues([heads]).setFontWeight('bold').setBackground('#E6F2EE');
    sh.setFrozenRows(1);
    sh.getRange('A:Z').setNumberFormat('@'); // 모두 글자로 저장 (날짜·숫자 자동변환 방지)
  });
  var cf = ss.getSheetByName('설정') || ss.insertSheet('설정');
  if (cf.getLastRow() < 2) {
    cf.getRange(1, 1, 5, 3).setValues([
      ['항목', '값', '설명'],
      ['공지', '앱 사랑방에 오신 것을 환영합니다! 서로의 앱을 구경하고, 따뜻한 의견을 나눠 주세요.', '첫 화면 위쪽에 나오는 한 줄 공지'],
      ['다음모임', '', '예: 12월 1일(월) 오후 2시 · 장유중앙교회 교육관'],
      ['문의메일', '', '아이디어·문의를 메일로도 받고 싶을 때'],
      ['관리자비밀번호', '', '모든 앱의 수정·반영 표시를 할 수 있는 비밀번호 (잊어버린 분 도울 때)']
    ]);
    cf.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#E6F2EE');
    cf.setColumnWidth(2, 420);
  }
  var s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('시트1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  folder_();
  SpreadsheetApp.getUi && Logger.log('준비 완료! 이제 [배포 → 새 배포 → 웹 앱]으로 배포하세요.');
}

/* ---------- 입구 ---------- */
function doGet() { return out_(all_()); }

function doPost(e) {
  var b;
  try { b = JSON.parse(e.postData.contents || '{}'); } catch (x) { return out_({ok:false, error:'잘못된 요청입니다'}); }
  try {
    if (b.action === 'all') return out_(all_());
    if (b.action === 'upload') return out_(upload_(b));
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try { return out_(write_(b)); } finally { lock.releaseLock(); }
  } catch (err) {
    return out_({ok:false, error: String(err && err.message || err)});
  }
}

function write_(b) {
  var now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  var me = clip_(b.author, 30) || '익명';
  switch (b.action) {
    case 'addApp': {
      var a = b.app || {};
      need_(a.name, '앱 이름'); need_(a.desc, '한 줄 설명'); need_(a.url, '주소');
      if (String(b.pw || '').length < 4) throw new Error('비밀번호는 4자리 이상으로 정해 주세요');
      var id = newId_('A');
      append_('apps', appRow_(id, now, a, hash_(b.pw), 0));
      return {ok:true, id:id};
    }
    case 'editApp': {
      var r = find_('apps', b.id); auth_(r, b.pw);
      var a2 = b.app || {};
      need_(a2.name, '앱 이름'); need_(a2.desc, '한 줄 설명'); need_(a2.url, '주소');
      var pw = b.newPw ? hash_(b.newPw) : r.obj.pw;
      setRow_('apps', r.row, appRow_(r.obj.id, r.obj.date, a2, pw, r.obj.likes));
      return {ok:true};
    }
    case 'checkPw': { auth_(find_('apps', b.id), b.pw); return {ok:true}; }
    case 'addUpdate': {
      auth_(find_('apps', b.appId), b.pw); need_(b.text, '내용');
      append_('updates', [newId_('U'), b.appId, now, clip_(b.ver, 20), clip_(b.text, 1500), '']);
      return {ok:true};
    }
    case 'addFeedback': {
      find_('apps', b.appId); need_(b.text, '내용');
      var type = ['idea','fix','bug','thanks'].indexOf(b.type) >= 0 ? b.type : 'idea';
      append_('feedback', [newId_('F'), b.appId, now, type, me, clip_(b.text, 1500), 0, '', '', '']);
      return {ok:true};
    }
    case 'markFeedback': {
      var f = find_('feedback', b.id);
      auth_(find_('apps', f.obj.appId), b.pw);
      setCell_('feedback', f.row, 'done', b.done ? 'Y' : '');
      setCell_('feedback', f.row, 'doneNote', clip_(b.note, 300));
      return {ok:true};
    }
    case 'addIdea': {
      need_(b.title, '제목');
      append_('ideas', [newId_('I'), now, clip_(b.title, 80), clip_(b.text, 1500), me, 0, '제안', '', '']);
      return {ok:true};
    }
    case 'claimIdea': {
      var i = find_('ideas', b.id);
      var who = clip_(b.builder, 30); need_(who, '이름');
      var cur = String(i.obj.builder || '');
      var list = cur ? cur.split(', ') : [];
      if (list.indexOf(who) < 0) list.push(who);
      setCell_('ideas', i.row, 'builder', list.join(', '));
      if (i.obj.status !== '완성') setCell_('ideas', i.row, 'status', '만드는 중');
      return {ok:true};
    }
    case 'addTopic': {
      need_(b.title, '주제');
      append_('topics', [newId_('T'), now, clip_(b.title, 80), clip_(b.text, 1000), me, 0, '']);
      return {ok:true};
    }
    case 'addDemo': {
      need_(b.app, '앱 이름');
      append_('demos', [newId_('D'), now, clip_(b.app, 60), me, clip_(b.when, 60), clip_(b.note, 500), '']);
      return {ok:true};
    }
    case 'addQuestion': {
      need_(b.title, '제목');
      append_('questions', [newId_('Q'), now, clip_(b.title, 100), clip_(b.text, 2000), me, '', '']);
      return {ok:true};
    }
    case 'addAnswer': {
      find_('questions', b.qId); need_(b.text, '내용');
      append_('answers', [newId_('R'), b.qId, now, me, clip_(b.text, 2000), 0, '']);
      return {ok:true};
    }
    case 'solveQuestion': {
      var q = find_('questions', b.id);
      setCell_('questions', q.row, 'solved', b.solved ? 'Y' : '');
      return {ok:true};
    }
    case 'like': {
      if (!LIKEABLE[b.kind]) throw new Error('잘못된 요청입니다');
      var l = find_(b.kind, b.id);
      var n = (Number(l.obj.likes) || 0) + 1;
      setCell_(b.kind, l.row, 'likes', String(n));
      return {ok:true, likes:n};
    }
  }
  throw new Error('알 수 없는 요청입니다');
}

/* ---------- 읽기 ---------- */
function all_() {
  var d = {ok:true};
  Object.keys(T).forEach(function (k) {
    d[k] = rows_(k).filter(function (o) { return String(o.hidden).toUpperCase() !== 'Y'; })
      .map(function (o) { delete o.pw; delete o.hidden; delete o._row; return o; });
  });
  var c = cfg_();
  d.config = {notice: c['공지'] || '', next: c['다음모임'] || '', email: c['문의메일'] || ''};
  return d;
}
function rows_(k) {
  var sh = sheet_(k), n = sh.getLastRow();
  if (n < 2) return [];
  var cols = T[k].cols, vals = sh.getRange(2, 1, n - 1, cols.length).getValues();
  var outv = [];
  vals.forEach(function (v, i) {
    if (!v[0]) return;
    var o = {_row: i + 2};
    cols.forEach(function (c, j) {
      var x = v[j];
      if (x instanceof Date) x = Utilities.formatDate(x, TZ, 'yyyy-MM-dd HH:mm');
      o[c[0]] = (x === null || x === undefined) ? '' : x;
    });
    if (o.likes !== undefined) o.likes = Number(o.likes) || 0;
    outv.push(o);
  });
  return outv;
}
function cfg_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('설정'), o = {};
  if (!sh || sh.getLastRow() < 2) return o;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { o[String(r[0]).trim()] = String(r[1]).trim(); });
  return o;
}

/* ---------- 사진 올리기 ---------- */
function upload_(b) {
  var m = String(b.data || '').match(/^data:(image\/(?:jpeg|png|webp|gif));base64,(.+)$/);
  if (!m) throw new Error('사진 파일만 올릴 수 있어요');
  var bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > 4 * 1024 * 1024) throw new Error('사진이 너무 큽니다 (4MB 이하)');
  var name = 'shot_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmmss') + '_' + Math.floor(Math.random() * 1e4) + '.jpg';
  var file = folder_().createFile(Utilities.newBlob(bytes, m[1], name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {ok:true, url:'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1600'};
}
function folder_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (x) {} }
  var f = DriveApp.createFolder('앱사랑방_스크린샷');
  p.setProperty('FOLDER', f.getId());
  return f;
}

/* ---------- 작은 도구 ---------- */
function sheet_(k) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(T[k].name);
  if (!sh) { setup(); sh = ss.getSheetByName(T[k].name); }
  return sh;
}
function appRow_(id, date, a, pw, likes) {
  var shots = (a.shots || []).filter(okUrl_).slice(0, 3);
  var url = String(a.url || '').trim();
  if (!okUrl_(url)) throw new Error('주소는 https:// 로 시작해야 해요');
  return [id, date, clip_(a.name, 40), clip_(a.desc, 100), clip_(a.cat, 20), clip_(a.status, 20), url,
    clip_(a.maker, 30) || '익명', clip_(a.howto, 1500), shots[0] || '', shots[1] || '', shots[2] || '',
    clip_(a.story, 2000), clip_(a.tool, 80), clip_(a.hours, 30), clip_(a.prompt, 6000), String(likes || 0), pw, ''];
}
function okUrl_(u) { return /^https?:\/\/[^\s"'<>]+$/i.test(String(u || '').trim()); }
function append_(k, row) { sheet_(k).appendRow(row.map(safe_)); }
function setRow_(k, r, row) { sheet_(k).getRange(r, 1, 1, row.length).setValues([row.map(safe_)]); }
function setCell_(k, r, key, v) {
  var j = T[k].cols.map(function (c) { return c[0]; }).indexOf(key);
  sheet_(k).getRange(r, j + 1).setValue(safe_(v));
}
function find_(k, id) {
  var hit = null;
  if (id) rows_(k).some(function (o) { if (String(o.id) === String(id)) { hit = o; return true; } });
  if (!hit || String(hit.hidden).toUpperCase() === 'Y') throw new Error('해당 글을 찾을 수 없어요');
  return {row: hit._row, obj: hit};
}
function auth_(r, pw) {
  pw = String(pw || '');
  if (pw && hash_(pw) === r.obj.pw) return;
  var admin = cfg_()['관리자비밀번호'];
  if (pw && admin && pw === admin) return;
  throw new Error('비밀번호가 맞지 않아요');
}
function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'appbang|' + s, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function newId_(p) { return p + Date.now().toString(36) + Math.floor(Math.random() * 1296).toString(36); }
function need_(v, label) { if (!String(v || '').trim()) throw new Error(label + '을(를) 적어 주세요'); }
function clip_(v, n) { return String(v == null ? '' : v).trim().slice(0, n); }
function safe_(v) { v = (v == null) ? '' : String(v); return /^[=+\-@]/.test(v) ? "'" + v : v; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
