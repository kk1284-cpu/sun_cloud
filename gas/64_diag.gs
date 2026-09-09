// ============================================================
// gas_takamura ／ [64] なぜ下書きが出ないのかを、人ごとに出す
// ------------------------------------------------------------
//  「37名いるのに19名しか下書きが出ない」を、推測ではなく数で答えるための1本。
//  メニュー ⑯ から実行。期間は既定で「今日までの7日間」。
//
//  下書きが出るまでに人が落ちる場所は4つある。どこで落ちたかを1人1行で出す。
//
//    ① 日報を書いていない            → 日報統合に行が無い
//    ② 取り込めていない              → 元フォームにはあるが日報統合に無い（氏名が名簿と違う等）
//    ③ まだ判定していない            → 日報統合にあるが分類済みに無い
//    ④ 該当なしで終わった            → 判定はしたが96項目のどれにも当たらず、活動記録が0件
//    ⑤ メールが無くて落ちた          → 活動記録はあるが メンバーID（メール）が空
//                                       ★2026-09-09 まではここで静かに消えていた
//
//  ⑤は 70_gc / 90_gcslide を直したので、いまは氏名で拾う（共有だけができない）。
// ============================================================

function tkGC_なぜ出ないか() {
  var days = 7;
  try {
    var ui = SpreadsheetApp.getUi();
    var r = ui.prompt('下書きが出ない理由を調べる',
      '何日ぶんを見ますか（既定 7 ＝今日までの7日間）。\n数字だけ入れて OK。',
      ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    var v = parseInt(String(r.getResponseText() || '').trim(), 10);
    if (v > 0 && v <= 90) days = v;
  } catch (e) {}
  return tkGCDiag_(days);
}

function tkGCDiag_(days) {
  var ss = tkSS_();
  var to = new Date(); to.setHours(23, 59, 59, 999);
  var from = new Date(to.getFullYear(), to.getMonth(), to.getDate() - (days - 1));
  var inRange = function (t) { if (!t) return false; var d = new Date(t); return d >= from && d <= to; };
  var label = Utilities.formatDate(from, TK.TZ, 'M/d') + '〜' + Utilities.formatDate(to, TK.TZ, 'M/d');

  // ---- 名簿 ----
  var people = {}, order = [];           // key = メール or '名前:氏名'
  function keyOf(mid, name) { return mid ? mid : (name ? '名前:' + String(name).replace(/\s/g, '') : ''); }
  function touch(mid, name, dept) {
    var k = keyOf(mid, name); if (!k) return null;
    if (!people[k]) {
      people[k] = { key: k, mid: mid || '', name: name || '', dept: dept || '',
                    roster: false, raw: 0, cls: 0, none: 0, act: 0, draft: 0 };
      order.push(k);
    }
    var p = people[k];
    if (mid && !p.mid) p.mid = mid;
    if (name && !p.name) p.name = name;
    if (dept && !p.dept) p.dept = dept;
    return p;
  }

  var mem = ss.getSheetByName(TK.SHEET_MEMBER);
  if (mem && mem.getLastRow() > 1) {
    mem.getRange(2, 1, mem.getLastRow() - 1, 3).getValues().forEach(function (r) {
      var p = touch(tkMemberId_(r[0]), String(r[1] || '').trim(), String(r[2] || '').trim());
      if (p) p.roster = true;
    });
  }

  // ---- 日報統合 ----
  var raw = ss.getSheetByName(TK.SHEET_RAW);
  var rawTotal = 0;
  if (raw && raw.getLastRow() > 1) {
    raw.getRange(2, 1, raw.getLastRow() - 1, TK_UNI_HEADER.length).getValues().forEach(function (r) {
      if (!inRange(r[TK.RAW_TS])) return;
      rawTotal++;
      var p = touch(tkMemberId_(r[TK.RAW_EMAIL]), String(r[TK.RAW_NAME] || '').trim(), String(r[TK.RAW_DEPT] || '').trim());
      if (p) p.raw++;
    });
  }

  // ---- 分類済み（判定コードが空＝該当なし） ----
  var cls = ss.getSheetByName(TK.SHEET_CLASS);
  var clsTotal = 0, noneTotal = 0;
  if (cls && cls.getLastRow() > 1) {
    cls.getRange(2, 1, cls.getLastRow() - 1, TK_CLASS_HEADER.length).getValues().forEach(function (r) {
      if (!inRange(r[1])) return;
      clsTotal++;
      var p = touch(tkMemberId_(r[2]), String(r[3] || '').trim(), String(r[4] || '').trim());
      var hit = String(r[6] || '').trim();
      if (!hit) noneTotal++;
      if (p) { p.cls++; if (!hit) p.none++; }
    });
  }

  // ---- 活動記録 ----
  var act = ss.getSheetByName(TK.SHEET_ACT);
  var actTotal = 0, actNoMid = 0;
  if (act && act.getLastRow() > 1) {
    act.getRange(2, 1, act.getLastRow() - 1, TK_ACT_HEADER.length).getValues().forEach(function (r) {
      if (!inRange(r[0])) return;
      actTotal++;
      var mid = tkMemberId_(r[2]);
      if (!mid) actNoMid++;
      var p = touch(mid, String(r[3] || '').trim(), String(r[4] || '').trim());
      if (p) p.act++;
    });
  }

  // ---- できている下書き（フォルダを見る） ----
  var draftTotal = 0;
  try {
    var work = tkGCSWorkFolder_();
    var folders = work.getFolders();
    while (folders.hasNext()) {
      var f = folders.next();
      var files = f.getFiles(), n = 0;
      while (files.hasNext()) {
        var file = files.next();
        if (file.getName().indexOf('GCシート ') !== 0) continue;
        if (file.getDateCreated() < from) continue;
        n++;
      }
      if (!n) continue;
      draftTotal += n;
      var nm = f.getName().replace(/\s/g, '');
      var hit = null;
      order.forEach(function (k) { if (!hit && String(people[k].name).replace(/\s/g, '') === nm) hit = people[k]; });
      if (hit) hit.draft += n; else { var p2 = touch('', f.getName(), ''); if (p2) p2.draft += n; }
    }
  } catch (e) {}

  // ---- 理由を決める ----
  function reasonOf(p) {
    if (p.draft) return '出ている';
    if (!p.raw && !p.act) return '① 日報なし';
    if (p.raw && !p.cls) return '③ 未判定';
    if (p.cls && !p.act) return '④ 該当なしのみ';
    if (p.act && !p.mid) return '⑤ メール無し（要名簿）';
    if (p.act) return '⑥ 材料はあるが未作成（実行待ち）';
    return '② 取り込めていない';
  }
  var tally = {};
  order.forEach(function (k) {
    var p = people[k]; p.reason = reasonOf(p);
    tally[p.reason] = (tally[p.reason] || 0) + 1;
  });

  // ---- 出力 ----
  var lines = ['=== 下書きが出ない理由（' + label + '／' + days + '日間） ===', ''];
  lines.push('日報統合 ' + rawTotal + '件　判定 ' + clsTotal + '件（うち該当なし ' + noneTotal + '件）　'
    + '活動記録 ' + actTotal + '件（うちメール無し ' + actNoMid + '件）　下書き ' + draftTotal + '件');
  lines.push('名簿 ' + order.filter(function (k) { return people[k].roster; }).length + '名　'
    + '登場した人 ' + order.length + '名');
  lines.push('');
  lines.push('■ 内訳');
  Object.keys(tally).sort().forEach(function (k) { lines.push('　' + k + '：' + tally[k] + '名'); });
  lines.push('');
  lines.push('■ 人ごと（日報／判定／該当なし／活動／下書き）');
  order.sort(function (a, b) {
    var A = people[a], B = people[b];
    if (A.reason !== B.reason) return A.reason < B.reason ? -1 : 1;
    return (B.raw - A.raw);
  });
  order.forEach(function (k) {
    var p = people[k];
    lines.push('　' + p.reason + '　' + (p.name || p.mid) + '（' + (p.dept || '部署未設定') + '）'
      + '　日報' + p.raw + '／判定' + p.cls + '／該当なし' + p.none + '／活動' + p.act + '／下書き' + p.draft
      + (p.mid ? '' : '　⚠メール無し') + (p.roster ? '' : '　⚠名簿に無し'));
  });
  lines.push('');
  lines.push('※ ①が多いのは、その期間に日報を書いた人がそれだけだったということです（仕組みの不具合ではありません）。');
  lines.push('※ ⑤は名簿にメールを入れると解消します。いまは氏名で拾って下書きは作りますが、本人への共有ができません。');
  lines.push('※ ⑥は ⑫ GC下書きを今すぐ作る で出ます。');

  var msg = lines.join('\n');
  tkLog_('診断', '下書き診断（' + label + '）' + Object.keys(tally).map(function (k) { return k + tally[k]; }).join(' '));
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg.slice(0, 1800) + (msg.length > 1800 ? '\n\n…続きは実行ログ（表示→ログ）で見てください。' : '')); } catch (e) {}
  return msg;
}
