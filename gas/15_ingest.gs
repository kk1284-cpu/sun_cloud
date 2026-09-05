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
//  取り込んだ直後に tk日次キャッチアップ() を続けて呼ぶ（tk取込トリガー）。
//  → 日報が書かれてから最長30分で判定まで終わる（以前は取込30分＋判定1時間＝最長90分）。
//
//  ★2026-09-06 変更
//   ・身分証を「ソース#シートID#行番号」から「ソース#送信日時#氏名」に変えた。
//     元シートが並べ替え・行削除されても二重取込／取りこぼしが起きない（五十嵐で実際に起きた事故の予防）。
//     旧キーも併用するので既存2,500件の再取込は起きない。移行作業なし。
//   ・列を**見出し名**で解けるようにした（BODY_COLS の head、シート「取込ソース」の AUTO）。
//     設問を編集・並べ替えしても壊れない。
// ============================================================

var TK_UNI_HEADER = ['タイムスタンプ', 'メール', '氏名', '部署', '本文',
  'ソース', '元行', '今週の目標', 'リソース', '備考'];

// ============================================================
//  取り込み本体
// ============================================================
function tk取込_全ソース() {
  var srcs = tkSetAllSources_();   // config.gs の TK.SOURCES ＋ シート「取込ソース」
  if (!srcs.length) throw new Error('取込ソースがありません（config.gs の TK.SOURCES か、シート「取込ソース」）');

  var out = tkEnsureSheet_(TK.SHEET_RAW, TK_UNI_HEADER);
  var seenAll = tkUniSeen_(out);
  var seen = seenAll.rows, seen2 = seenAll.ids;
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
      var head = sh.getRange(1, 1, 1, width).getValues()[0];
      var cols = tkUniResolve_(src, head);          // 見出し名 → 列番号
      cols.warn.forEach(function (w) { lines.push('　⚠ ' + src.NAME + '「' + sh.getName() + '」: ' + w); });
      if (cols.name == null && src.AUTO) return;     // 誰の日報か分からないシートは取り込まない
      var vals = sh.getRange(2, 1, last - 1, width).getValues();

      for (var i = 0; i < vals.length; i++) {
        var rowNo = i + 2;
        var key = src.ID + '#' + sid + '#' + rowNo;          // 旧キー（行番号）
        if (seen[key]) { skip++; continue; }

        var v = vals[i];
        var ts = v[cols.ts];
        if (!(ts instanceof Date)) { var d = new Date(ts); ts = isNaN(d.getTime()) ? '' : d; }
        if (!ts) { skip++; continue; }
        if (src.SINCE) {
          var since = new Date(src.SINCE + 'T00:00:00');
          if (ts < since) { skip++; continue; }
        }

        var rawName = cols.name != null ? String(v[cols.name] || '') : '';
        var name = tkUniPersonName_(rawName, roster);
        var email = cols.email != null ? String(v[cols.email] || '').trim() : '';
        if (!email && name && roster.byName[name]) email = roster.byName[name];

        var key2 = tkUniIdKey_(src.ID, ts, name);            // 新キー（日時＋氏名）
        if (seen2[key2]) { skip++; continue; }

        var body = tkUniBody_(v, cols.body);
        if (!body) { skip++; continue; }   // 本文が全部空の行は取り込まない

        rows.push([
          ts, email, name, src.DEPT || '',
          tkSafeCell_(body),
          src.ID, sid + '#' + rowNo,
          tkSafeCell_(tkUniPick_(v, cols.weekgoal)),
          tkSafeCell_(tkUniPick_(v, cols.resource)),
          tkSafeCell_(tkUniNote_(v, cols.notes))
        ]);
        seen[key] = 1; seen2[key2] = 1;
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
  lines.push('※ 自動取込（30分ごと）のときは、続けて判定まで行います。手で実行したときは ② で判定してください。');
  var msg = lines.join('\n');
  tkLog_('取込', '追加 ' + rows.length + '件');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return { added: rows.length };
}

// ---------- 列の解決（見出し名 → 列番号） ----------
function tkUniNorm_(s) { return String(s == null ? '' : s).replace(/[\s　]+/g, '').trim(); }
// 完全一致 → 前方一致 → 部分一致
function tkUniFindHead_(head, want) {
  var w = tkUniNorm_(want); if (!w) return -1;
  var hs = head.map(tkUniNorm_);
  var i = hs.indexOf(w); if (i >= 0) return i;
  for (i = 0; i < hs.length; i++) if (hs[i] && hs[i].indexOf(w) === 0) return i;
  for (i = 0; i < hs.length; i++) if (hs[i] && hs[i].indexOf(w) >= 0) return i;
  return -1;
}
//  config 定義（従来）: BODY_COLS の col を使う。head があれば見出し名で解き、無ければ col。
//  シート定義（AUTO）  : 見出しから全部決める。名前・メール・タイムスタンプ以外の見出し付き列を本文にする。
function tkUniResolve_(src, head) {
  var r = { ts: src.COL_TS != null ? src.COL_TS : 0, name: src.COL_NAME, email: src.COL_EMAIL,
            body: [], weekgoal: src.COL_WEEKGOAL, resource: src.COL_RESOURCE, notes: [], warn: [] };
  if (!src.AUTO) {
    (src.BODY_COLS || []).forEach(function (c) {
      var col = c.col;
      if (c.head) {
        var i = tkUniFindHead_(head, c.head);
        if (i >= 0) col = i; else r.warn.push('見出し「' + c.head + '」が見つからず、列' + c.col + 'を使いました');
      }
      if (col != null) r.body.push({ col: col, label: c.label || String(head[col] || '') });
    });
    (src.NOTE_COLS || []).forEach(function (c) { r.notes.push(c); });
    if (src.HEAD_NAME)  { var n = tkUniFindHead_(head, src.HEAD_NAME);  if (n >= 0) r.name = n; }
    if (src.HEAD_EMAIL) { var m = tkUniFindHead_(head, src.HEAD_EMAIL); if (m >= 0) r.email = m; }
    return r;
  }
  var iTs = tkUniFindHead_(head, 'タイムスタンプ'); r.ts = iTs >= 0 ? iTs : 0;
  var iName = src.HEAD_NAME ? tkUniFindHead_(head, src.HEAD_NAME) : -1;
  ['名前', '氏名', 'お名前'].forEach(function (k) { if (iName < 0) iName = tkUniFindHead_(head, k); });
  r.name = iName >= 0 ? iName : null;
  var iMail = src.HEAD_EMAIL ? tkUniFindHead_(head, src.HEAD_EMAIL) : -1;
  if (iMail < 0) iMail = tkUniFindHead_(head, 'メール');
  r.email = iMail >= 0 ? iMail : null;
  r.weekgoal = null; r.resource = null;
  head.forEach(function (h, i) {
    var t = String(h || '').trim(); if (!t) return;
    if (i === r.ts || i === r.name || i === r.email) return;
    if (/^列\s*\d+$/.test(t)) return;                      // Googleフォームが付ける名無しの列
    if (r.weekgoal == null && /目標/.test(t)) { r.weekgoal = i; return; }
    if (r.resource == null && /リソース/.test(t)) { r.resource = i; return; }
    if (/コンディション|体調/.test(t)) { r.notes.push({ col: i, label: t.slice(0, 20) }); return; }
    r.body.push({ col: i, label: t.slice(0, 30) });
  });
  if (r.name == null) r.warn.push('「名前」の列が見つかりません（誰の日報か判定できないため、このシートは取り込みません）');
  if (!r.body.length) r.warn.push('本文に使える列がありません');
  return r;
}

// 身分証：ソース＋送信日時＋氏名（空白を除く）。行番号に依存しない
function tkUniIdKey_(srcId, ts, name) {
  var t = ts instanceof Date ? ts.getTime() : new Date(ts).getTime();
  return srcId + '#' + (isNaN(t) ? String(ts) : t) + '#' + String(name || '').replace(/\s/g, '');
}

// ---------- 本文を見出し付きで1本にまとめる ----------
function tkUniBody_(v, bodyCols) {
  var parts = [];
  (bodyCols || []).forEach(function (c) {
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
function tkUniNote_(v, noteCols) {
  var out = [];
  (noteCols || []).forEach(function (c) {
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
//  rows: 旧キー 'ソース#sid#行'（元行 列から） ／ ids: 新キー 'ソース#日時#氏名'（日時・氏名・ソース列から）
//  両方を見るので、切り替え時に既存行が二重に入ることはない
function tkUniSeen_(out) {
  var seen = {}, ids = {};
  var last = out.getLastRow();
  if (last < 2) return { rows: seen, ids: ids };
  var iSrc = TK_UNI_HEADER.indexOf('ソース') + 1;
  var iRow = TK_UNI_HEADER.indexOf('元行') + 1;
  var w = Math.max(iSrc, iRow);
  var vals = out.getRange(2, 1, last - 1, w).getValues();
  vals.forEach(function (r) {
    var s = String(r[iSrc - 1] || '').trim();
    var n = String(r[iRow - 1] || '').trim();
    if (s && n) seen[s + '#' + n] = 1;   // n は 'sid#row' 形式
    if (s && r[TK.RAW_TS]) ids[tkUniIdKey_(s, r[TK.RAW_TS], r[TK.RAW_NAME])] = 1;
  });
  return { rows: seen, ids: ids };
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
  (TK.SOURCES || []).concat(tkSetSources_()).forEach(function (src) {
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
        var cols = tkUniResolve_(src, head);
        lines.push('　→ 解決: 日時[' + cols.ts + '] 名前[' + cols.name + '] メール[' + cols.email + ']' +
          ' 目標[' + cols.weekgoal + '] リソース[' + cols.resource + ']');
        lines.push('　→ 本文: ' + cols.body.map(function (c) { return '[' + c.col + ']' + c.label; }).join(' / '));
        cols.warn.forEach(function (w) { lines.push('　⚠ ' + w); });
      }
    } catch (e) {
      lines.push('　⚠ 開けません: ' + (e && e.message));
    }
  });
  lines.push('');
  lines.push('※ config 定義のソースは [番号] を COL_* / BODY_COLS に。シート「取込ソース」のものは見出し名で自動です。');
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
function tk取込トリガー() {
  var r = null;
  try { r = tk取込_全ソース(); } catch (e) { tkLog_('取込', '失敗: ' + (e && e.message)); }
  // 取り込んだら続けて判定（日報が書かれてから最長30分で活動記録まで届く）
  if (r && r.added) {
    try { tk日次キャッチアップ(); } catch (e) { tkLog_('取込', '判定を続けて実行できず: ' + (e && e.message)); }
  }
}

function tk取込_トリガー設定() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'tk取込トリガー') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('tk取込トリガー').timeBased().everyMinutes(30).create();
  var msg = '30分ごとに既存フォームから日報を取り込む設定にしました。\n'
          + '取り込んだ直後に判定も続けて行います（1時間ごとの tk日次キャッチアップ() は取りこぼしの保険として残します）。';
  tkLog_('取込', msg); Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}
