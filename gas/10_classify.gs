// ============================================================
// gas_takamura ／ [10] 日報 → Gemini分類（サンクスUP評価基準）
// ------------------------------------------------------------
//  日報フォーム送信時に発火し、その人の「部署」に対応するサンクスUP評価基準
//  （成果・能力・姿勢 × 一般/ベテラン）のどれに最も当てはまるかを1つ判定して
//  「分類済み」シートに書く。判定は人が後から修正できる。
// ============================================================

var TK_CLASS_HEADER = [
  'raw行', '日時', 'メール', '氏名', '部署', '本文',
  '判定コード', '観点', '小見出し', '対象',
  '確信度', '判定理由', '候補2', '判定日時', '活動記録済み', '修正コード（手動）'
];

// クエスト表（サンクスUP基準）を読む。10分キャッシュ。
//   クエスト表の列: [コード, 部署, 対象, 観点, No, 小見出し, 内容, 出典]
function tkLoadQuests_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('tk_quests');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  var sh = tkSS_().getSheetByName(TK.SHEET_QUEST);
  if (!sh || sh.getLastRow() < 2) return [];
  // 列: コード/部署/対象/観点/GC分類/No/小見出し/内容/出典（GC分類は後から追加した列）
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
  var quests = [];
  vals.forEach(function (r) {
    var code = String(r[0] || '').trim();
    if (!code) return;
    var axis = String(r[3] || '').trim();
    var gc = String(r[4] || '').trim();
    // GC分類が空（旧レイアウトのシート）なら観点から自動補完
    if (!gc) gc = (axis === '成果') ? '直接の成果' : (axis === '姿勢') ? '価値への取り組み' : '人材育成';
    quests.push({
      code: code, dept: String(r[1] || '').trim(), level: String(r[2] || '').trim(),
      axis: axis, gc: gc, no: r[5], title: String(r[6] || '').trim(),
      desc: String(r[7] || '').trim()
    });
  });
  try { cache.put('tk_quests', JSON.stringify(quests), 600); } catch (e) {}
  return quests;
}

// 分類本体。dept 指定で候補を絞る（未指定なら全部署）。
function tkClassify_(text, dept) {
  if (!text) return { ok: false, error: 'empty text' };
  var quests = tkLoadQuests_();
  if (!quests.length) return { ok: false, error: 'クエスト表が空です（先にSETUPを実行）' };

  var cand = quests.filter(function (q) { return !dept || q.dept === dept; });
  if (!cand.length) cand = quests;   // 部署不明時は全件から

  var list = cand.map(function (q) {
    return q.code + '｜' + q.axis + '｜' + q.level + '｜' + q.title + '｜例:' + q.desc;
  }).join('\n');

  var prompt =
    'あなたはSUNグループ（介護・障がい・薬局・本部）の日報を、人事評価制度「サンクスUP！」の\n' +
    '評価基準（成果・能力・姿勢）に分類する係です。\n' +
    (dept ? ('報告者の部署は「' + dept + '」です。\n') : '') +
    '以下は評価基準の一覧です（形式: コード｜観点｜対象｜小見出し｜記入例）:\n' + list + '\n\n' +
    '次の日報が最もよく当てはまる基準を1つ選び、JSONで答えてください。\n' +
    '- code: 最有力の基準コード。どうしても該当がなければ ""\n' +
    '- alt: 2番目に近いコード（なければ ""）\n' +
    '- confidence: "高"|"中"|"低"\n' +
    '- reason: 判定理由をひとこと（30字以内）\n\n' +
    '日報:\n' + String(text).slice(0, 1200);

  var schema = {
    type: 'OBJECT',
    properties: {
      code: { type: 'STRING' }, alt: { type: 'STRING' },
      confidence: { type: 'STRING' }, reason: { type: 'STRING' }
    },
    required: ['code', 'confidence', 'reason']
  };
  var g = tkGeminiJson_(prompt, schema);
  if (!g.ok) return { ok: false, error: g.error };

  var byCode = {};
  quests.forEach(function (q) { byCode[q.code] = q; });
  var q = byCode[String(g.data.code || '').trim()] || null;
  return {
    ok: true,
    quest: q ? { code: q.code, axis: q.axis, level: q.level, title: q.title, dept: q.dept } : null,
    alt: String(g.data.alt || '').trim(),
    confidence: String(g.data.confidence || ''),
    reason: String(g.data.reason || '')
  };
}

// フォーム送信トリガー本体
function tkOnFormSubmit(e) {
  // SS_ID未設定でも動くよう、トリガー発火元のSSを使う
  if (e && e.source) TK_SS_OVERRIDE = e.source;
  var vals = (e && e.values) || [];
  var row = (e && e.range) ? e.range.getRow() : 0;
  var ts    = vals[TK.RAW_TS]   || new Date();
  var email = String(vals[TK.RAW_EMAIL] || '').trim();
  var name  = String(vals[TK.RAW_NAME]  || '').trim();
  var dept  = String(vals[TK.RAW_DEPT]  || '').trim();
  var text  = String(vals[TK.RAW_BODY]  || '').trim();
  if (!text) return;

  // 送信時は判定しない設定（TK.CLASSIFY_ON_SUBMIT=false）のときは何も書かない。
  // → 1時間ごとの「キャッチアップ」がまとめて判定する（Geminiの1日上限を節約）
  if (TK.CLASSIFY_ON_SUBMIT === false) return;

  var r = tkClassify_(text, dept);

  // AI側の失敗（利用上限・通信エラー等）は記録しない。
  //   記録すると「処理済み」扱いになって二度と判定されなくなるため、
  //   未処理のまま残して後続のキャッチアップ／一括分類に拾わせる。
  if (!r || !r.ok) {
    tkLog_('分類', 'row ' + row + ' 判定できず（' + ((r && r.error) || '不明') + '）→ 未処理のまま残します');
    return;
  }

  var rec = { row: row, ts: ts, email: email, name: name, dept: dept, text: text, result: r };
  tkWriteClass_(rec);

  // 分類できたら、そのまま活動記録へ追記（マンダラOKR反映の土台）
  try {
    if (r.ok && r.quest) tkAppendActivityFromRec_(rec);
  } catch (err) { tkLog_('OKR反映', 'activity追記失敗: ' + (err && err.message)); }
}

// 「分類済み」へ1行書く
function tkWriteClass_(rec) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sh = tkEnsureSheet_(TK.SHEET_CLASS, TK_CLASS_HEADER);
    var r = rec.result || {};
    var q = (r.ok && r.quest) || {};
    sh.appendRow([
      rec.row, rec.ts, rec.email, rec.name, rec.dept, rec.text,
      q.code || '', q.axis || '', q.title || '', q.level || '',
      r.confidence || (r.ok ? '' : 'エラー'), r.reason || r.error || '', r.alt || '',
      new Date(), q.code ? '' : '', ''
    ]);
  } finally {
    lock.releaseLock();
  }
}

// 過去分の一括分類（バックログ）。新しい順に未分類n件。
function tkProcessBacklog_(n) {
  n = n || 10;
  var raw = tkRawSheet_();
  if (!raw || raw.getLastRow() < 2) return { ok: true, processed: 0 };
  var out = tkEnsureSheet_(TK.SHEET_CLASS, TK_CLASS_HEADER);
  var done = {};
  if (out.getLastRow() > 1) {
    out.getRange(2, 1, out.getLastRow() - 1, 1).getValues().forEach(function (r) { done[String(r[0])] = 1; });
  }
  var last = raw.getLastRow();
  var width = Math.max(TK.RAW_BODY + 1, 5);
  var vals = raw.getRange(2, 1, last - 1, width).getValues();
  var processed = 0, results = [];
  for (var i = vals.length - 1; i >= 0 && processed < n; i--) {
    var rowNo = i + 2;
    if (done[String(rowNo)]) continue;
    var text = String(vals[i][TK.RAW_BODY] || '').trim();
    if (!text) continue;
    var dept = String(vals[i][TK.RAW_DEPT] || '').trim();
    var rec = {
      row: rowNo, ts: vals[i][TK.RAW_TS], email: String(vals[i][TK.RAW_EMAIL] || ''),
      name: String(vals[i][TK.RAW_NAME] || ''), dept: dept, text: text,
      result: tkClassify_(text, dept)
    };
    tkWriteClass_(rec);
    try { if (rec.result.ok && rec.result.quest) tkAppendActivityFromRec_(rec); } catch (e) {}
    results.push({ row: rowNo, code: rec.result.quest && rec.result.quest.code, conf: rec.result.confidence });
    processed++;
    Utilities.sleep(400);
  }
  return { ok: true, processed: processed, results: results };
}
