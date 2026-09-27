// ============================================================
// gas_takamura ／ [62] 設定シート（コードを触らずに運用を変える）
// ------------------------------------------------------------
//  五十嵐商会の「設定はすべてスプレッドシートに置く」方式を持ってきた。
//  加藤さん・石川さんがシートを書き換えるだけで、次の3つが変わる。
//
//   GC設定       … 部署（またはソースID）ごとの GC曜日。
//                  毎日17:30の tkGC日次トリガー が「明日GCの部署」の下書きだけ作る。
//                  対象期間は「GC前日までの7日間」。
//   推進メンバー  … 作業フォルダ「GC下書き」全体を編集権で共有する人。
//                  親フォルダに付くので、中の全員分のフォルダ・下書きが見える。
//   取込ソース   … 新しい振り返りフォームの回答SS。列は**見出し名**で自動判定する
//                  （名前・メール・タイムスタンプ以外の見出し付き列を全部本文にする）。
//                  config.gs の TK.SOURCES と併用。同じIDならシートが勝つ。
//
//  初回: tk設定シート_作成() を1回実行（既にあるシートは触らない）
// ============================================================

var TK_SET = {
  HEADER_GC:  ['部署またはソースID', 'GC曜日', '有効', 'メモ'],
  HEADER_MGR: ['メール', '氏名', '役割', '共有'],
  HEADER_SRC: ['ID', '名称', '部署', 'スプレッドシートID', 'シートgid（空＝フォーム連携シート全部）',
               '名前の列見出し（空＝名前/氏名）', 'メールの列見出し（空＝自動）', '有効', '開始日（YYYY-MM-DD）', 'メモ'],
  WEEKDAYS: ['日', '月', '火', '水', '木', '金', '土'],
  DEFAULT_GC_DAY: '土'    // 高村社長のグループが土曜。前日（金曜）の11時に作る
};

// GC設定シートの既定の行＝GC下書きフォルダの事業所と同じ並び。
//   ここを部署（薬局・ネットワークなど）にすると薬局とネットワークが分かれないため、
//   フォルダと同じ「事業所」で持つ。
function tkSetDefaultGroups_() {
  var out = [];
  (TK.SOURCES || []).forEach(function (s) {
    if (s.ENABLED === false) return;
    var o = String(s.OFFICE || s.DEPT || '').trim();
    if (o && out.indexOf(o) < 0) out.push(o);
  });
  var other = (TK.GC_SLIDE && TK.GC_SLIDE.OFFICE_OTHER) || 'その他';
  if (out.indexOf(other) < 0) out.push(other);
  return out;
}

// ---------- 曜日 ----------
function tkSetWeekdayIndex_(s) {
  s = String(s || '').trim().replace(/曜日?$/, '');
  var i = TK_SET.WEEKDAYS.indexOf(s);
  if (i >= 0) return i;
  return ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'].indexOf(s.slice(0, 3).toLowerCase());
}
function tkSetIsOn_(v) {
  var s = String(v == null ? '' : v).trim();
  return !(s === 'しない' || s === '無効' || s === 'いいえ' || s === 'false' || s === 'FALSE' || s === '0' || s === '×');
}

// ---------- GC設定を読む ----------
//  戻り: [{ group:'薬局・ネットワーク', day:6, enabled:true }, ...]
//  シートが無い／空のときは TK.DEPTS 全部を既定曜日（土）で返す＝従来の金曜夕方の動きと同じ
function tkSetGcGroups_() {
  var out = [];
  try {
    var sh = tkSS_().getSheetByName(TK.SHEET_GCSET);
    if (sh && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (r) {
        var g = String(r[0] || '').trim();
        if (!g || g.indexOf('（例）') === 0) return;
        var d = tkSetWeekdayIndex_(r[1]);
        if (d < 0) d = tkSetWeekdayIndex_(TK_SET.DEFAULT_GC_DAY);
        out.push({ group: g, day: d, enabled: tkSetIsOn_(r[2]) });
      });
    }
  } catch (e) {}
  if (!out.length) {
    tkSetDefaultGroups_().forEach(function (d) {
      out.push({ group: d, day: tkSetWeekdayIndex_(TK_SET.DEFAULT_GC_DAY), enabled: true });
    });
  }
  return out;
}

// ---------- 推進メンバー ----------
//  シート + config の TK.GC_SLIDE.MANAGERS。小文字・重複なし
function tkSetManagers_() {
  var seen = {}, out = [];
  function add(m) {
    m = String(m || '').trim().toLowerCase();
    if (m && m.indexOf('@') > 0 && m.indexOf('（') < 0 && !seen[m]) { seen[m] = 1; out.push(m); }
  }
  try {
    var sh = tkSS_().getSheetByName(TK.SHEET_MANAGERS);
    if (sh && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(function (r) {
        if (tkSetIsOn_(r[3])) add(r[0]);
      });
    }
  } catch (e) {}
  ((TK.GC_SLIDE && TK.GC_SLIDE.MANAGERS) || []).forEach(add);
  return out;
}

// ---------- 取込ソース（シート定義） ----------
//  TK.SOURCES と同じ形にして返す。AUTO:true が付くので 15_ingest が見出し名で列を決める
function tkSetSources_() {
  var out = [];
  try {
    var sh = tkSS_().getSheetByName(TK.SHEET_SOURCES);
    if (!sh || sh.getLastRow() < 2) return out;
    sh.getRange(2, 1, sh.getLastRow() - 1, TK_SET.HEADER_SRC.length).getValues().forEach(function (r) {
      var id = String(r[0] || '').trim();
      var ssId = String(r[3] || '').trim();
      if (!id || !ssId || id.indexOf('（例）') === 0 || ssId.indexOf('（') === 0) return;
      var gid = String(r[4] || '').trim();
      var since = r[8] instanceof Date ? Utilities.formatDate(r[8], TK.TZ, 'yyyy-MM-dd') : String(r[8] || '').trim();
      out.push({
        ID: id, NAME: String(r[1] || id).trim(), DEPT: String(r[2] || '').trim(),
        SS_ID: ssId, SHEET_NAME: '', GIDS: gid ? gid.split(/[,\s、]+/).filter(String) : null,
        SINCE: since, ENABLED: tkSetIsOn_(r[7]),
        AUTO: true,
        HEAD_NAME: String(r[5] || '').trim(), HEAD_EMAIL: String(r[6] || '').trim(),
        BODY_COLS: [], NOTE_COLS: [], COL_TS: 0, COL_NAME: null, COL_EMAIL: null,
        COL_WEEKGOAL: null, COL_RESOURCE: null
      });
    });
  } catch (e) {}
  return out;
}

// config と シートを合わせた「いま有効なソース」
function tkSetAllSources_() {
  var byId = {}, order = [];
  (TK.SOURCES || []).forEach(function (s) { byId[s.ID] = s; order.push(s.ID); });
  tkSetSources_().forEach(function (s) { if (!byId[s.ID]) order.push(s.ID); byId[s.ID] = s; });
  return order.map(function (id) { return byId[id]; }).filter(function (s) { return s && s.ENABLED !== false; });
}

// ---------- ソースID → 事業所フォルダ名（OFFICE。無ければ部署、それも無ければ名称） ----------
function tkSetOfficeBySource_() {
  var map = {};
  (TK.SOURCES || []).concat(tkSetSources_()).forEach(function (s) {
    if (!s || !s.ID) return;
    map[s.ID] = String(s.OFFICE || s.DEPT || s.NAME || s.ID).trim();
  });
  return map;
}

// ---------- 人 → どのソースに書いているか（GC設定でソースIDを使えるようにする） ----------
//  日報統合の（氏名/メール → ソース）。最後に書いたソースを採用。
function tkSetPersonSources_() {
  var out = { byMid: {}, byName: {} };
  try {
    var sh = tkSS_().getSheetByName(TK.SHEET_RAW);
    if (!sh || sh.getLastRow() < 2) return out;
    var iSrc = TK_UNI_HEADER.indexOf('ソース');
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, iSrc + 1).getValues();
    vals.forEach(function (r) {
      var src = String(r[iSrc] || '').trim(); if (!src) return;
      var mid = tkMemberId_(r[TK.RAW_EMAIL]), nm = String(r[TK.RAW_NAME] || '').trim();
      if (mid) out.byMid[mid] = src;
      if (nm) out.byName[nm] = src;
    });
  } catch (e) {}
  return out;
}

// グループ（事業所 or 部署 or ソースID or ソース名）に、この人が入るか
//  ★事業所（OFFICE）で絞れることが要点。フォルダの分け方と同じ言葉で指定できる。
//    どのフォームにも書いていない人は「その他」に入る。
function tkSetGroupMatch_(group, a, ps) {
  if (!group) return true;
  if (String(a.office || '').trim() === group) return true;
  if (String(a.dept || '').trim() === group) return true;
  ps = ps || { byMid: {}, byName: {} };
  var src = (a.mid && ps.byMid[a.mid]) || (a.name && ps.byName[a.name]) || '';
  var other = (TK.GC_SLIDE && TK.GC_SLIDE.OFFICE_OTHER) || 'その他';
  if (!src) return group === other;          // どこにも書いていない人＝その他
  if (src === group) return true;
  var def = null;
  (TK.SOURCES || []).concat(tkSetSources_()).forEach(function (s) { if (s.ID === src) def = s; });
  if (!def) return false;
  return def.NAME === group || String(def.OFFICE || '').trim() === group;
}

// ============================================================
//  初回：設定シートを作る（既にあるものは触らない）
// ============================================================
function tk設定シート_作成() {
  var ss = tkSS_();
  var lines = ['=== 設定シート ==='];

  // GC設定（既定：部署ごとに土曜＝前日の金曜17:30に作る。従来の動きと同じ）
  var gc = ss.getSheetByName(TK.SHEET_GCSET);
  if (!gc) {
    gc = tkEnsureSheet_(TK.SHEET_GCSET, TK_SET.HEADER_GC);
    var defs = tkSetDefaultGroups_();
    var rows = defs.map(function (d) { return [d, TK_SET.DEFAULT_GC_DAY, 'する', 'GC下書きフォルダの事業所と同じ名前です']; });
    rows.push(['（例）薬局・ネットワーク', '金', 'しない', '部署名やソースID（yakkyoku / network / hananoaru）でも指定できます']);
    gc.getRange(2, 1, rows.length, TK_SET.HEADER_GC.length).setValues(rows);
    gc.setColumnWidth(1, 220); gc.setColumnWidth(4, 520);
    lines.push('　「' + TK.SHEET_GCSET + '」を作成（事業所' + defs.length + '件・既定 ' + TK_SET.DEFAULT_GC_DAY +
      '曜＝前日の' + tkGCSRunHour_() + '時に作成）');
  } else lines.push('　「' + TK.SHEET_GCSET + '」は既にあります（触っていません）');

  // 推進メンバー
  var mg = ss.getSheetByName(TK.SHEET_MANAGERS);
  if (!mg) {
    mg = tkEnsureSheet_(TK.SHEET_MANAGERS, TK_SET.HEADER_MGR);
    mg.getRange(2, 1, 1, 4).setValues([['（ここにメール）', '（氏名）', '推進メンバー', 'する']]);
    mg.setColumnWidth(1, 260);
    lines.push('　「' + TK.SHEET_MANAGERS + '」を作成（メールを入れて ⑬- を実行すると共有されます）');
  } else lines.push('　「' + TK.SHEET_MANAGERS + '」は既にあります（触っていません）');

  // 取込ソース
  var sc = ss.getSheetByName(TK.SHEET_SOURCES);
  if (!sc) {
    sc = tkEnsureSheet_(TK.SHEET_SOURCES, TK_SET.HEADER_SRC);
    sc.getRange(2, 1, 1, TK_SET.HEADER_SRC.length).setValues([[
      '（例）honbu', '本部 振り返り', '本部', '（回答スプレッドシートのID）', '', '', '', 'しない', '',
      '新しいフォームはここに1行足すだけ。列は見出し名で自動判定（名前・メール・タイムスタンプ以外は全部本文）'
    ]]);
    sc.setColumnWidth(4, 320); sc.setColumnWidth(10, 480);
    lines.push('　「' + TK.SHEET_SOURCES + '」を作成（config.gs の3ソースはそのまま有効）');
  } else lines.push('　「' + TK.SHEET_SOURCES + '」は既にあります（触っていません）');

  lines.push('');
  lines.push('次に: ⑬- で推進メンバーに共有 → ⑭ で毎日17:30の自動作成に切り替え');
  var msg = lines.join('\n');
  tkLog_('設定', msg); Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return msg;
}
