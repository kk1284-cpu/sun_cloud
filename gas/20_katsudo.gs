// ============================================================
// gas_takamura ／ [20] 分類済み → 活動記録（マンダラOKR反映の土台）
// ------------------------------------------------------------
//  「活動記録」は OKR反映ビューと 週次応援（賞賛）の共通データ源。
//  分類済みの1行 = 活動1件として、メンバーID・部署・観点・小見出し・EXP・実践内容を貯める。
//  raw行キーで冪等（同じ日報を二重計上しない）。
// ============================================================

var TK_ACT_HEADER = [
  '日時', 'raw行', 'メンバーID', '氏名', '部署', '観点', '小見出し', '対象',
  'EXP', '確信度', '実践内容', '判定コード'
];
var TK_MEMBER_HEADER = ['メンバーID', '氏名', '部署', '累計EXP', '報告数', '最終報告日', '配信停止', '社員番号'];

function tkExpFor_(confidence) {
  var c = String(confidence || '').trim();
  if (c === '高') return TK.EXP.高;
  if (c === '中') return TK.EXP.中;
  if (c === '低') return TK.EXP.低;
  return TK.EXP.none;
}

// 既に活動記録に入っている raw行 の集合
function tkActivityDoneRows_() {
  var sh = tkEnsureSheet_(TK.SHEET_ACT, TK_ACT_HEADER);
  var set = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().forEach(function (r) {
      var v = String(r[0] || '').trim(); if (v) set[v] = true;
    });
  }
  return set;
}

// 分類済みレコード（rec）から活動記録へ1行追記＋メンバー累計更新
function tkAppendActivityFromRec_(rec) {
  var q = (rec.result && rec.result.quest) || {};
  if (!q.code) return false;                     // 判定不能はOKRに反映しない
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sh = tkEnsureSheet_(TK.SHEET_ACT, TK_ACT_HEADER);
    // 冪等チェック
    if (rec.row) {
      var done = tkActivityDoneRows_();
      if (done[String(rec.row)]) return false;
    }
    var exp = tkExpFor_(rec.result.confidence);
    var mid = tkMemberId_(rec.email);
    sh.appendRow([
      rec.ts || new Date(), rec.row || '', mid, rec.name || '', rec.dept || '',
      q.axis || '', q.title || '', q.level || '', exp, rec.result.confidence || '',
      String(rec.text || '').slice(0, 500), q.code || ''
    ]);
    tkUpsertMember_(mid, rec.name, rec.dept, exp, rec.ts);
    return true;
  } finally {
    lock.releaseLock();
  }
}

// メンバー名簿の累計EXP・報告数・最終報告日を更新（無ければ追加）
function tkUpsertMember_(mid, name, dept, addExp, ts) {
  if (!mid) return;
  var sh = tkEnsureSheet_(TK.SHEET_MEMBER, TK_MEMBER_HEADER);
  var last = sh.getLastRow();
  var rowIdx = -1;
  if (last > 1) {
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim().toLowerCase() === mid) { rowIdx = i + 2; break; }
    }
  }
  var when = ts ? new Date(ts) : new Date();
  if (rowIdx < 0) {
    sh.appendRow([mid, name || '', dept || '', addExp || 0, 1, when, '', '']);
  } else {
    var cur = sh.getRange(rowIdx, 4, 1, 2).getValues()[0];
    var totalExp = (Number(cur[0]) || 0) + (addExp || 0);
    var reports = (Number(cur[1]) || 0) + 1;
    sh.getRange(rowIdx, 4, 1, 3).setValues([[totalExp, reports, when]]);
    if (name) sh.getRange(rowIdx, 2).setValue(name);
    if (dept) sh.getRange(rowIdx, 3).setValue(dept);
  }
}

// 分類済みで「活動記録済み」空欄の行をまとめて活動記録へ流す（手動修正後の再取り込み用）
function tkSyncClassifiedToActivity_() {
  var sh = tkSS_().getSheetByName(TK.SHEET_CLASS);
  if (!sh || sh.getLastRow() < 2) return { ok: true, added: 0 };
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, TK_CLASS_HEADER.length).getValues();
  var done = tkActivityDoneRows_();
  var quests = tkLoadQuests_();
  var byCode = {}; quests.forEach(function (q) { byCode[q.code] = q; });
  var added = 0;
  vals.forEach(function (r) {
    var rawRow = String(r[0] || '').trim();
    // 手動修正コードがあれば優先
    var code = String(r[15] || '').trim() || String(r[6] || '').trim();
    if (!code || !rawRow || done[rawRow]) return;
    var q = byCode[code];
    if (!q) return;
    var rec = {
      row: rawRow, ts: r[1], email: r[2], name: r[3], dept: r[4], text: r[5],
      result: { ok: true, confidence: r[10], quest: { code: q.code, axis: q.axis, level: q.level, title: q.title, dept: q.dept } }
    };
    if (tkAppendActivityFromRec_(rec)) added++;
  });
  return { ok: true, added: added };
}
