// ============================================================
// gas_takamura ／ [15] 既存の振り返りフォームを取り込む（複数ソース対応）
// ------------------------------------------------------------
//  ★方針転換（2026-08-18）
//   新しいフォームに書かせるのをやめた。SUNグループでは部署ごとに
//   **すでに別々のGoogleフォームで毎日振り返りを書いている**。
//   項目構成も部署ごとに違う。統一フォームに寄せると現場の入力が変わる＝
//   「別々に書く工数をなくす」という契約の本質に反する。
//   → **既存フォームをそのまま残し、システム側が吸い込む。**
//
//  つくり:
//   各ソースSS → 見出し付きで1本のテキストに正規化 → シート「日報統合」へ追記
//   以降は既存の tkBulkRun_ / 70_gc / 30_okr が**一切変更なし**で動く
//   （TK.SHEET_RAW = '日報統合' なので tkRawSheet_ がこれを返す）
//
//  日報統合の列: [タイムスタンプ, メール, 氏名, 部署, 本文, ソース, 元行, 今週の目標, リソース, 備考]
//                 ↑0            ↑1      ↑2    ↑3     ↑4     ↑5      ↑6    ↑7          ↑8      ↑9
//   0〜4 は TK.RAW_TS/EMAIL/NAME/DEPT/BODY とぴったり同じ並び（下流が無改造で動く理由）
//   7〜9 は分類には使わず、GC下書き等で参照する
//
//  使い方:
//    tk取込_全ソース()        … 未取込分を取り込む（1時間ごとトリガー推奨）
//    tk取込_トリガー設定()     … 30分ごとに自動取込
//    tk取込_状況()            … ソースごとの件数と未取込数
//    tk取込_ソース確認()       … 各ソースの列見出しを実物から表示（設定合わせに使う）
//
//  ⚠ 取り込んだだけでは分類されない。分類は tk日次キャッチアップ() が行う。
// ============================================================

var TK_UNI_HEADER = ['タイムスタンプ', 'メール', '氏名', '部署', '本文',
  'ソース', '元行', '今週の目標', 'リソース', '備考'];

// ============================================================
//  取り込み本体
// ============================================================
function tk取込_全ソース() {
  var srcs = (TK.SOURCES || []).filter(function (s) { return s.ENABLED !== false; });
  if (!srcs.length) throw new Error('TK.SOURCES が未設定です（config.gs）');

  var out = tkEnsureSheet_(TK.SHEET_RAW, TK_UNI_HEADER);
  var seen = tkUniSeen_(out);
  var roster = tkUniRoster_();
  var rows = [], lines = ['=== 日報の取り込み ==='];

  srcs.forEach(function (src) {
    var got = 0, skip = 0, err = '';
    try {
      var sheets = tkUniSourceSheets_(src);
      if (!sheets || !sheets.length) throw new Error('シートが見つかりません');

      sheets.forEach(function (sh) {
      var sid = sh.getSheetId();
      var last = sh.getLastRow();
      if (last < 2) return;

      var width = sh.getLastColumn();
      var vals = sh.getRange(2, 1, last - 1, width).getValues();

      for (var i = 0; i < vals.length; i++) {
        var rowNo = i + 2;
        var key = src.ID + '#' + sid + '#' + rowNo;
        if (seen[key]) { skip++; continue; }

        var v = vals[i];
        var ts = v[src.COL_TS != null ? src.COL_TS : 0];
        if (!(ts instanceof Date)) { var d = new Date(ts); ts = isNaN(d.getTime()) ? '' : d; }
        if (!ts) { skip++; continue; }
        if (src.SINCE) {
          var since = new Date(src.SINCE + 'T00:00:00');
          if (ts < since) { skip++; continue; }
        }

        var rawName = src.COL_NAME != null ? String(v[src.COL_NAME] || '') : '';
        var name = tkUniPersonName_(rawName, roster);
        var email = src.COL_EMAIL != null ? String(v[src.COL_EMAIL] || '').trim() : '';
        if (!email && name && roster.byName[name]) email = roster.byName[name];

        var body = tkUniBody_(v, src);
        if (!body) { skip++; continue; }   // 本文が全部空の行は取り込まない

        rows.push([
          ts, email, name, src.DEPT || '',
          tkSafeCell_(body),
          src.ID, sid + '#' + rowNo,
          tkSafeCell_(tkUniPick_(v, src.COL_WEEKGOAL)),
          tkSafeCell_(tkUniPick_(v, src.COL_RESOURCE)),
          tkSafeCell_(tkUniNote_(v, src))
        ]);
        seen[key] = 1;
        got++;
      }
      });
    } catch (e) {
      err = String(e && e.message || e);
    }
    lines.push('　' + src.NAME + '：取込 ' + got + '件／スキップ ' + skip + '件'
      + (err ? '　⚠ ' + err : ''));
  });

  if (rows.length) {
    // 時系列に並べてから追記（活動記録が古い順に積まれる）
    rows.sort(function (a, b) { return new Date(a[0]) - new Date(b[0]); });
    out.getRange(out.getLastRow() + 1, 1, rows.length, TK_UNI_HEADER.length).setValues(rows);
  }

  lines.push('');
  lines.push('合計 ' + rows.length + '件を「' + TK.SHEET_RAW + '」に追加しました。');
  lines.push('※ 分類は tk日次キャッチアップ()（1時間ごと）が行います。');
  var msg = lines.join('\n');
  tkLog_('取込', '追加 ' + rows.length + '件');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return { added: rows.length };
}

// ---------- 本文を見出し付きで1本にまとめる ----------
function tkUniBody_(v, src) {
  var parts = [];
  (src.BODY_COLS || []).forEach(function (c) {
    var t = String(v[c.col] == null ? '' : v[c.col]).trim();
    // 全角スペースだけ・「なし」「特になし」等は中身なしとして落とす
    if (!t || /^[\s　。．・]*$/.test(t)) return;
    if (/^(なし|特になし|特にありません|無し)[。．]?$/.test(t)) return;
    parts.push('■' + c.label + '\n' + t);
  });
  return parts.join('\n\n');
}

function tkUniPick_(v, col) {
  if (col == null) return '';
  var t = String(v[col] == null ? '' : v[col]).trim();
  return /^[\s　。．・]*$/.test(t) ? '' : t;
}

// 分類には使わないが残しておきたい情報（コンディション等）
function tkUniNote_(v, src) {
  var out = [];
  (src.NOTE_COLS || []).forEach(function (c) {
    var t = tkUniPick_(v, c.col);
    if (t) out.push(c.label + '：' + t);
  });
  return out.join(' ／ ');
}

// ---------- 氏名の取り出し ----------
//  「名前」列に氏名＋フィロソフィ目標が連結されている。区切り方がフォームで違う。
//   ・ネットワーク：コロン区切り「大串徹:燃える闘魂」（コロン無しの期間も混在）
//   ・薬局：スペース区切り
//   例）「大島 恵美子 常に明るく前向きに夢と希望をもって素直な心で」
//      「神田 愛  常に創造的な仕事をする」
//      「光野直子」（フィロソフィなし）
//  → メンバー名簿の氏名と前方一致（最長一致）で確定する。
//    名簿に無ければ、姓+名の2トークンまでを氏名とみなす。
function tkUniPersonName_(raw, roster) {
  var s = String(raw || '').replace(/[　]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!s) return '';

  // ⓪ コロン区切りは、そこから後ろが全部フィロソフィ（ネットワークフォームの形式）
  //    例）「大串徹:燃える闘魂」「上岡美之:常に謙虚でなければならない」
  //    同じフォームでもコロン無し（「大串徹」）の期間があるため、両方を1人に寄せる
  var ci = s.search(/[:：]/);
  if (ci > 0) s = s.slice(0, ci).trim();

  // ① 名簿と前方一致（スペース有無を無視して比較）
  var flat = s.replace(/\s/g, '');
  var best = '';
  roster.names.forEach(function (n) {
    var nf = n.replace(/\s/g, '');
    if (nf && flat.indexOf(nf) === 0 && nf.length > best.replace(/\s/g, '').length) best = n;
  });
  if (best) return best;

  // ② 名簿に無い → 先頭2トークンまで（3トークン目以降はフィロソフィとみなす）
  var t = s.split(' ');
  if (t.length === 1) return t[0];
  return t[0] + ' ' + t[1];
}

// ---------- 既に取り込んだ行 ----------
function tkUniSeen_(out) {
  var seen = {};
  var last = out.getLastRow();
  if (last < 2) return seen;
  var iSrc = TK_UNI_HEADER.indexOf('ソース') + 1;
  var iRow = TK_UNI_HEADER.indexOf('元行') + 1;
  var w = Math.max(iSrc, iRow);
  var vals = out.getRange(2, 1, last - 1, w).getValues();
  vals.forEach(function (r) {
    var s = String(r[iSrc - 1] || '').trim();
    var n = String(r[iRow - 1] || '').trim();
    if (s && n) seen[s + '#' + n] = 1;   // n は 'sid#row' 形式
  });
  return seen;
}

// ---------- メンバー名簿（氏名リストとメール対応） ----------
function tkUniRoster_() {
  var names = [], byName = {};
  try {
    var sh = tkSS_().getSheetByName(TK.SHEET_MEMBER);
    if (sh && sh.getLastRow() > 1) {
      var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
      vals.forEach(function (r) {
        var mail = String(r[0] || '').trim();
        var nm = String(r[1] || '').trim();
        if (!nm) return;
        names.push(nm);
        if (mail && mail.indexOf('@') > 0) byName[nm] = mail;
      });
    }
  } catch (e) {}
  // 長い名前を先に見るため降順
  names.sort(function (a, b) { return b.length - a.length; });
  return { names: names, byName: byName };
}

// ---------- ソースのシートを開く（1ソースに複数シートがありうる） ----------
//  薬局フォームは年度ごとに回答シートが分かれている（gid違い・ヘッダは同じ）。
//  GIDS に複数指定すれば全部まとめて取り込む。重複キーはシートIDを含めるので衝突しない。
function tkUniSourceSheets_(src) {
  var ss = src.SS_ID ? SpreadsheetApp.openById(src.SS_ID) : tkSS_();
  var all = ss.getSheets();
  var out = [];

  var gids = src.GIDS || (src.GID != null ? [src.GID] : null);
  if (gids && gids.length) {
    gids.forEach(function (g) {
      all.forEach(function (sh) { if (String(sh.getSheetId()) === String(g)) out.push(sh); });
    });
    if (out.length) return out;
  }
  if (src.SHEET_NAME) {
    var byName = ss.getSheetByName(src.SHEET_NAME);
    if (byName) return [byName];
  }
  // 指定なし → フォーム連携シートを全部（無ければ先頭）
  all.forEach(function (sh) {
    try { if (sh.getFormUrl && sh.getFormUrl()) out.push(sh); } catch (e) {}
  });
  return out.length ? out : [all[0]];
}

// 後方互換（tk取込_ソース確認 が使う）
function tkUniSourceSheet_(src) {
  var a = tkUniSourceSheets_(src);
  return a && a.length ? a[0] : null;
}

// ============================================================
//  設定合わせ用：各ソースの列見出しを実物から表示
// ============================================================
function tk取込_ソース確認() {
  var lines = ['=== ソースの列見出し（実物） ==='];
  (TK.SOURCES || []).forEach(function (src) {
    lines.push('');
    lines.push('▼ ' + src.NAME + '（' + src.ID + '）　部署: ' + (src.DEPT || '（未設定）'));
    if (!src.SS_ID) { lines.push('　⚠ SS_ID が未設定です'); return; }
    try {
      var ss = SpreadsheetApp.openById(src.SS_ID);
      lines.push('　SS: ' + ss.getName());
      ss.getSheets().forEach(function (sh) {
        lines.push('　― シート「' + sh.getName() + '」 gid=' + sh.getSheetId()
          + '　回答 ' + Math.max(0, sh.getLastRow() - 1) + '件');
      });
      var sh2 = tkUniSourceSheet_(src);
      if (sh2 && sh2.getLastColumn() > 0) {
        var head = sh2.getRange(1, 1, 1, sh2.getLastColumn()).getValues()[0];
        lines.push('　使用シート「' + sh2.getName() + '」の列:');
        head.forEach(function (h, i) {
          lines.push('　　[' + i + '] ' + String(h || '（空）'));
        });
      }
    } catch (e) {
      lines.push('　⚠ 開けません: ' + (e && e.message));
    }
  });
  lines.push('');
  lines.push('※ 上の [番号] を config の COL_* / BODY_COLS に設定してください。');
  var msg = lines.join('\n');
  Logger.log(msg);
  return msg;
}

// ============================================================
//  状況
// ============================================================
function tk取込_状況() {
  var out = tkEnsureSheet_(TK.SHEET_RAW, TK_UNI_HEADER);
  var last = out.getLastRow();
  var bySrc = {}, total = Math.max(0, last - 1);
  if (last > 1) {
    var iSrc = TK_UNI_HEADER.indexOf('ソース') + 1;
    out.getRange(2, iSrc, last - 1, 1).getValues().forEach(function (r) {
      var s = String(r[0] || '(不明)');
      bySrc[s] = (bySrc[s] || 0) + 1;
    });
  }
  var lines = ['=== 取り込み状況 ===', '「' + TK.SHEET_RAW + '」合計: ' + total + '件', ''];
  Object.keys(bySrc).forEach(function (k) { lines.push('　' + k + ': ' + bySrc[k] + '件'); });

  // 未分類数
  try {
    var cls = tkSS_().getSheetByName(TK.SHEET_CLASS);
    var done = cls ? Math.max(0, cls.getLastRow() - 1) : 0;
    lines.push('');
    lines.push('分類済み: ' + done + '件　未分類: ' + Math.max(0, total - done) + '件');
  } catch (e) {}

  var msg = lines.join('\n');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return msg;
}

// ============================================================
//  トリガー（30分ごと）
// ============================================================
function tk取込トリガー() { try { tk取込_全ソース(); } catch (e) { tkLog_('取込', '失敗: ' + (e && e.message)); } }

function tk取込_トリガー設定() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'tk取込トリガー') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('tk取込トリガー').timeBased().everyMinutes(30).create();
  var msg = '30分ごとに既存フォームから日報を取り込む設定にしました。\n'
          + '分類は1時間ごとの tk日次キャッチアップ() が行います。';
  tkLog_('取込', msg); Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}
