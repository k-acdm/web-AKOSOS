/**
 * web-AKOSOS — 講習日程希望調査アプリ バックエンド（GAS）
 * 春日部アカデミー
 *
 * 【スクリプトプロパティ（プロジェクトの設定）】
 *   MYKATSU_SS_ID … マイ活スプレッドシート（Students 読み取り元）
 *   SURVEY_SS_ID  … 専用スプレッドシート（Terms / resp_* 書き込み先）
 *
 * 【シート構成】
 *   マイ活SS : Students（生徒ID / 氏名 / GRADE_LEVEL を使用）
 *   専用SS   : Terms（講習レジストリ）, resp_<term_id>（講習別の入力結果）
 *
 * 【○×の意味】 ○=授業OK=1 ／ ×=入れないで=0 ／ 既定は全コマ○
 */

// ====== 定数 ======
var PROP          = PropertiesService.getScriptProperties();
var MYKATSU_SS_ID = PROP.getProperty('MYKATSU_SS_ID');
var SURVEY_SS_ID  = PROP.getProperty('SURVEY_SS_ID');

var STUDENTS_SHEET = 'Students';     // ※タブ名がこれで正しいか要確認
var TERMS_SHEET    = 'Terms';
var RESP_PREFIX    = 'resp_';
var EXPORT_PREFIX  = 'export_';

var SLOTS        = ['A', 'B', 'C', 'D', 'E', 'F'];          // 6コマ固定
var SLOT_HEADERS = SLOTS.map(function (s) { return s + 'コマ'; }); // Aコマ..Fコマ
var RESP_HEADERS = ['生徒ID', '生徒名', '日付'].concat(SLOT_HEADERS);
var TERM_COLS    = ['term_id', 'title', 'start_date', 'num_blocks', 'active',
                    'time_A', 'time_B', 'time_C', 'time_D', 'time_E', 'time_F'];
var DOW          = ['日', '月', '火', '水', '木', '金', '土'];
var TZ           = 'Asia/Tokyo';
var DATE_FMT     = 'yyyy/M/d';

// ====== Web エントリポイント ======
function doGet(e)  { return handle_(e); }
function doPost(e) { return handle_(e); }

function handle_(e) {
  var out;
  try {
    var params = parseParams_(e);
    var action = params.action || 'ping';
    switch (action) {
      case 'ping':  out = { ok: true, msg: 'web-AKOSOS backend alive' }; break;
      case 'login': out = apiLogin_(params); break;
      case 'save':  out = apiSave_(params);  break;
      default:      out = { ok: false, error: 'unknown action: ' + action };
    }
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/** POST(JSON / text/plain) と GET(query) の両対応でパラメータを取得 */
function parseParams_(e) {
  if (e && e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch (x) { /* fallthrough */ }
  }
  return (e && e.parameter) ? e.parameter : {};
}

// ====== API：ログイン ======
function apiLogin_(p) {
  var sid = String(p.studentId || '').trim();
  if (!sid) return { ok: false, error: '生徒IDを入力してください' };

  var stu = getStudent_(sid);
  if (!stu) return { ok: false, error: '生徒IDが見つかりません。番号をご確認ください。' };

  var term = getActiveTerm_();
  if (!term) {
    return { ok: true, student: stu, term: null, dates: [], grid: {},
             message: '現在受付中の調査はありません。' };
  }

  var dates = buildDates_(term);
  var grid  = readStudentGrid_(term.term_id, sid, dates);
  return {
    ok: true,
    student: stu,
    term: { id: term.term_id, title: term.title, slots: slotLabels_(term) },
    dates: dates,
    grid: grid
  };
}

// ====== API：保存（upsert）======
function apiSave_(p) {
  var sid = String(p.studentId || '').trim();
  if (!sid) return { ok: false, error: '生徒IDがありません' };

  var stu = getStudent_(sid);
  if (!stu) return { ok: false, error: '生徒IDが見つかりません' };

  var term = getActiveTerm_();
  if (!term) return { ok: false, error: '受付中の調査がありません' };
  if (p.termId && String(p.termId) !== String(term.term_id)) {
    return { ok: false, error: '調査が切り替わりました。画面を再読み込みしてください。' };
  }

  var dates = buildDates_(term);
  writeStudentGrid_(term.term_id, stu, dates, p.grid || {});
  return { ok: true, saved: dates.length };
}

// ====== Students 参照 ======
function getStudent_(sid) {
  var sh = openMykatsu_().getSheetByName(STUDENTS_SHEET);
  if (!sh) throw new Error('Studentsシートが見つかりません');
  var data = sh.getDataRange().getValues();
  var ci = colIndex_(data[0], ['生徒ID', '氏名', 'GRADE_LEVEL']);
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][ci['生徒ID']]).trim() === sid) {
      return {
        id: sid,
        name:  String(data[r][ci['氏名']] || ''),
        grade: String(data[r][ci['GRADE_LEVEL']] || '')
      };
    }
  }
  return null;
}

function getAllStudents_() {
  var sh = openMykatsu_().getSheetByName(STUDENTS_SHEET);
  if (!sh) throw new Error('Studentsシートが見つかりません');
  var data = sh.getDataRange().getValues();
  var ci = colIndex_(data[0], ['生徒ID', '氏名']);
  var out = [];
  for (var r = 1; r < data.length; r++) {
    var id = String(data[r][ci['生徒ID']]).trim();
    if (id) out.push({ id: id, name: String(data[r][ci['氏名']] || '') });
  }
  return out;
}

// ====== Terms 参照 ======
function getActiveTerm_() {
  var sh = openSurvey_().getSheetByName(TERMS_SHEET);
  if (!sh) throw new Error('Termsシートが見つかりません');
  var data = sh.getDataRange().getValues();
  var idx = colIndex_(data[0], TERM_COLS);
  for (var r = 1; r < data.length; r++) {
    var a = data[r][idx['active']];
    if (a === true || String(a).toUpperCase() === 'TRUE') {
      return rowToTerm_(data[r], idx);
    }
  }
  return null;
}

function getTermById_(termId) {
  var sh = openSurvey_().getSheetByName(TERMS_SHEET);
  if (!sh) throw new Error('Termsシートが見つかりません');
  var data = sh.getDataRange().getValues();
  var idx = colIndex_(data[0], TERM_COLS);
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idx['term_id']]).trim() === String(termId).trim()) {
      return rowToTerm_(data[r], idx);
    }
  }
  return null;
}

function rowToTerm_(row, idx) {
  var t = {};
  Object.keys(idx).forEach(function (k) { t[k] = row[idx[k]]; });
  return t;
}

function slotLabels_(term) {
  return SLOTS.map(function (s) {
    return { slot: s, time: String(term['time_' + s] || '') };
  });
}

// ====== 日付生成（num_blocks × 14日）======
function buildDates_(term) {
  var start = toDate_(term.start_date);
  var n = Number(term.num_blocks) * 14;
  if (!(n > 0)) throw new Error('num_blocks が不正です: ' + term.num_blocks);
  var arr = [];
  for (var i = 0; i < n; i++) {
    var d = new Date(start.getTime());
    d.setDate(d.getDate() + i);
    arr.push({
      key:   fmt_(d, DATE_FMT),
      label: fmt_(d, 'M/d'),
      dow:   DOW[d.getDay()],
      block: Math.floor(i / 14) + 1
    });
  }
  return arr;
}

// ====== resp_<term> 読み取り（既定は全○）======
function readStudentGrid_(termId, sid, dates) {
  var sh = ensureRespSheet_(termId);
  var last = sh.getLastRow();
  var map = {};
  if (last > 1) {
    var data = sh.getRange(2, 1, last - 1, 9).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === sid) {
        map[dkeyOf_(data[i][2])] = [3, 4, 5, 6, 7, 8].map(function (c) {
          return Number(data[i][c]) ? 1 : 0;
        });
      }
    }
  }
  var grid = {};
  dates.forEach(function (d) {
    grid[d.key] = map[d.key] || [1, 1, 1, 1, 1, 1]; // 行が無ければ全○
  });
  return grid;
}

// ====== resp_<term> 書き込み（upsert）======
function writeStudentGrid_(termId, stu, dates, grid) {
  var sh = ensureRespSheet_(termId);
  var last = sh.getLastRow();
  var values = (last > 1) ? sh.getRange(2, 1, last - 1, 9).getValues() : [];

  var rowByDate = {};
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === stu.id) {
      rowByDate[dkeyOf_(values[i][2])] = i;
    }
  }

  var appends = [];
  dates.forEach(function (d) {
    var g = grid[d.key];
    var v = (g && g.length === 6)
      ? g.map(function (x) { return Number(x) ? 1 : 0; })
      : [1, 1, 1, 1, 1, 1];
    if (rowByDate[d.key] !== undefined) {
      var r = rowByDate[d.key];
      values[r][1] = stu.name;
      values[r][3] = v[0]; values[r][4] = v[1]; values[r][5] = v[2];
      values[r][6] = v[3]; values[r][7] = v[4]; values[r][8] = v[5];
    } else {
      appends.push([stu.id, stu.name, d.key, v[0], v[1], v[2], v[3], v[4], v[5]]);
    }
  });

  if (values.length) sh.getRange(2, 1, values.length, 9).setValues(values);
  if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, 9).setValues(appends);
}

// ====== resp_<term> シート用意 ======
function ensureRespSheet_(termId) {
  var ss = openSurvey_();
  var name = RESP_PREFIX + termId;
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, RESP_HEADERS.length).setValues([RESP_HEADERS]).setFontWeight('bold');
    sh.getRange(1, 3, sh.getMaxRows(), 1).setNumberFormat('@'); // 日付は文字列扱い
    sh.setFrozenRows(1);
  }
  return sh;
}

// ====== 全○一括初期化（論点14・(b)）======
// active講習について、全生徒×全日付を「全○(=1)」で投入。既存の回答行は保持。
function initTermResponses_(termId) {
  var term = getTermById_(termId);
  if (!term) throw new Error('term_id が見つかりません: ' + termId);
  var dates = buildDates_(term);
  var sh = ensureRespSheet_(termId);

  var last = sh.getLastRow();
  var existing = {};
  if (last > 1) {
    var head = sh.getRange(2, 1, last - 1, 3).getValues();
    head.forEach(function (row) {
      existing[String(row[0]).trim() + '|' + dkeyOf_(row[2])] = true;
    });
  }

  var students = getAllStudents_();
  var appends = [];
  students.forEach(function (st) {
    dates.forEach(function (d) {
      if (!existing[st.id + '|' + d.key]) {
        appends.push([st.id, st.name, d.key, 1, 1, 1, 1, 1, 1]);
      }
    });
  });
  if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, 9).setValues(appends);
  return appends.length;
}

// ====== 業者出力（生徒ID無し版）======
// resp_<term> から先頭の生徒ID列を落とした export_<term> タブを作成/更新。
// 業者システムが生徒ID列を取り込めない場合に使用（取り込める場合は resp_ タブをそのままCSV化）。
function makeExportNoId_(termId) {
  var ss = openSurvey_();
  var src = ensureRespSheet_(termId);
  var data = src.getDataRange().getValues();
  var outName = EXPORT_PREFIX + termId;
  var out = ss.getSheetByName(outName);
  if (out) ss.deleteSheet(out);
  out = ss.insertSheet(outName);
  var rows = data.map(function (r) { return r.slice(1); }); // 生徒ID列を除外
  out.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  out.setFrozenRows(1);
  return outName;
}

// ====== スプレッドシートメニュー（塾長用）======
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('講習調査')
    .addItem('active講習を初期化（全○投入）', 'menuInitActive_')
    .addItem('active講習：生徒ID無しCSV用シート作成', 'menuExportNoId_')
    .addToUi();
}

function menuInitActive_() {
  var ui = SpreadsheetApp.getUi();
  var term = getActiveTerm_();
  if (!term) { ui.alert('active=TRUE の講習がありません'); return; }
  var n = initTermResponses_(term.term_id);
  ui.alert('初期化完了\nterm: ' + term.term_id + '\n追加した行: ' + n + '（既存の回答は保持）');
}

function menuExportNoId_() {
  var ui = SpreadsheetApp.getUi();
  var term = getActiveTerm_();
  if (!term) { ui.alert('active講習がありません'); return; }
  var name = makeExportNoId_(term.term_id);
  ui.alert('「' + name + '」を作成/更新しました。\nこのタブを ファイル→ダウンロード→CSV で書き出してください。');
}

// ====== 小物 ======
function openMykatsu_() {
  if (!MYKATSU_SS_ID) throw new Error('スクリプトプロパティ MYKATSU_SS_ID が未設定です');
  return SpreadsheetApp.openById(MYKATSU_SS_ID);
}
function openSurvey_() {
  if (!SURVEY_SS_ID) throw new Error('スクリプトプロパティ SURVEY_SS_ID が未設定です');
  return SpreadsheetApp.openById(SURVEY_SS_ID);
}
function colIndex_(headers, names) {
  var map = {};
  names.forEach(function (n) {
    var i = headers.indexOf(n);
    if (i < 0) throw new Error('列が見つかりません: ' + n);
    map[n] = i;
  });
  return map;
}
function fmt_(d, f) { return Utilities.formatDate(d, TZ, f); }
function dkeyOf_(v) { return (v instanceof Date) ? fmt_(v, DATE_FMT) : String(v).trim(); }
function toDate_(v) {
  if (v instanceof Date) return v;
  var s = String(v).trim().replace(/-/g, '/');
  var m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  throw new Error('開始日(start_date)の形式が不正です: ' + v);
}

// ====== 動作確認用（エディタから手動実行）======
function test_login_() {
  var res = apiLogin_({ studentId: '1001' }); // テスト用ID
  Logger.log(JSON.stringify(res, null, 2));
}
