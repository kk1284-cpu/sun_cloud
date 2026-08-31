// ============================================================
// gas_takamura ／ [80] まとめて分類（キャッチアップ／過去分の一括分類）
// ------------------------------------------------------------
//  ★なぜ必要か：Geminiの無料枠は「1日20回」（RPMではなくRPDで詰まる）。
//    日報が届くたびに1件ずつ判定すると、日報15件で使い切って
//    その日の21件目から分類が止まる。3部署でトライアルを始めた瞬間に踏む。
//    → config で TK.CLASSIFY_ON_SUBMIT: false にして送信時判定をやめ、
//      ここが 15件まとめて1回のAPI呼び出しで処理する（1日1〜2回に圧縮）。
//
//  日々の運用（これだけ設定すればよい）:
//    tkキャッチアップ_設定()  … 1時間ごとの自動判定トリガーを作る（最初に1回だけ）
//    tk日次キャッチアップ()    … 実処理（トリガーが自動で呼ぶ）
//
//  過去分をまとめて片付けたいとき:
//    tk一括分類_開始()          … 開始（以後5分おきに自動で続く・レジューム式）
//    tk一括分類_状況()          … 進捗（何件済み／残り何件）
//    tk一括分類_停止()          … 中断（分類済みのデータは残る）
//    tk一括分類_エラー行を削除() … 失敗行を消して再判定させる
//    tkAPIキー確認()            … いまのキーが無料枠か有料かを判定（25回連打）
//
//  設計上ゆずれない3点（過去に事故った）:
//    ① 429/5xx はリトライ→それでもダメなら null を返して「記録せず中断」。
//       記録すると処理済み扱いになり、二度と判定されなくなる。
//    ② 1件も進まない実行が3回続いたら自分でトリガーを止める（空回り防止）。
//    ③ 読みは実行ごと1回だけ・書きは setValues で一括（1件ごと全読みはO(n²)で激遅）。
// ============================================================

var TK_BULK = {
  RUN_SEC: 280,        // 1回の実行で使う秒数（6分上限に対する余裕）
  HEADROOM_SEC: 40,    // 次のバッチを始めるのに必要な残り時間
  EVERY_MIN: 5,        // 継続トリガーの間隔（分）
  BATCH: 15,           // Gemini 1回でまとめて判定する件数
                       //   判定がぼやける感じがしたら 10 に下げる。
  PACE_MS: 1500,       // バッチ間の待ち
  MAX_BATCHES: 0,      // 1回の実行で呼ぶGeminiの最大回数（0=無制限）
  CATCHUP_BATCHES: 8,  // キャッチアップ1回で呼ぶ回数。BATCH×これ＝1時間あたりの処理件数。
                       //   案内メッセージもここを見るので、数字がずれない。
                       //   2026-08-31 有料枠にしたので 2 → 8（30件/時 → 120件/時）
  RETRY: 2,            // 429/5xx のときの再試行回数
  RETRY_WAIT_MS: 8000, // 再試行までの待ち（1日上限なら待っても無駄なので短く）
  TRIGGER_FN: 'tk一括分類_継続',
  SINCE: ''            // 対象開始日（''なら全期間）。例 '2026-08-01'（トライアル開始日）
};

// ============================================================
//  日々の運用：1時間ごとに「まだ判定していない日報」をまとめて判定
// ============================================================
function tkキャッチアップ_設定() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'tk日次キャッチアップ') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('tk日次キャッチアップ').timeBased().everyHours(1).create();
  var msg = '1時間ごとの自動判定（キャッチアップ）を設定しました。\n'
          + '未判定の日報を、1回あたり最大' + (TK_BULK.BATCH * TK_BULK.CATCHUP_BATCHES)
            + '件（API ' + TK_BULK.CATCHUP_BATCHES + '回）まで処理します。\n\n'
          + '※ あわせて config の TK.CLASSIFY_ON_SUBMIT が false になっていることを確認してください。';
  tkLog_('キャッチアップ', msg);
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}

function tk日次キャッチアップ() {
  var saveMax = TK_BULK.MAX_BATCHES, savePace = TK_BULK.PACE_MS;
  TK_BULK.MAX_BATCHES = TK_BULK.CATCHUP_BATCHES;   // 上の CATCHUP_BATCHES に従う
                                //   2026-08-31 有料枠に切り替えたので 2 から上げた。
                                //   無料枠(5回/分)に戻す場合は 2 に戻す。
  TK_BULK.PACE_MS = 3000;
  try {
    var r = tkBulkRun_();
    if (r.processed) tkLog_('キャッチアップ', '判定 ' + r.processed + '件 / 未判定の残り ' + r.remaining + '件');
    else if (r.quota) tkLog_('キャッチアップ', 'APIの利用上限のため今回は判定できませんでした（次の時間に再挑戦）');
  } catch (e) {
    tkLog_('キャッチアップ', '失敗: ' + (e && e.message));
  } finally {
    TK_BULK.MAX_BATCHES = saveMax; TK_BULK.PACE_MS = savePace;
  }
}

// ============================================================
//  期間を絞って判定する（今週分を先に押さえるため）
// ============================================================
//  ⚠ tkBulkRun_ は「古い順」に回す。過去1年分(2458件)が入っている状態で
//     素直に流すと、今週の日報がいちばん最後になり、金曜のGC下書きが空になる。
//     そこで SINCE を一時的に効かせて、その日以降だけを先に判定する。
//     終わったら SINCE は元に戻すので、あとから過去分を流せる。
function tk分類_期間を絞って(since, maxBatches) {
  var saveSince = TK_BULK.SINCE, saveMax = TK_BULK.MAX_BATCHES;
  TK_BULK.SINCE = since;
  TK_BULK.MAX_BATCHES = maxBatches || 0;   // 0=無制限（6分の実行上限まで回す）
  var r;
  try {
    r = tkBulkRun_();
  } finally {
    TK_BULK.SINCE = saveSince; TK_BULK.MAX_BATCHES = saveMax;
  }
  var msg = '期間を絞って判定（' + since + ' 以降）: 今回 ' + r.processed + '件'
          + ' / この期間の残り ' + r.remaining + '件';
  if (r.remaining > 0) msg += '\nもう一度実行すると続きを判定します。';
  else msg += '\nこの期間は完了しました。';
  tkLog_('期間指定分類', msg);
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return r;
}

// トライアル開始日(9/1)以降だけを判定する。金曜のGC下書きに間に合わせるため
function tk分類_今週分を先に() {
  return tk分類_期間を絞って('2026-09-01', 0);
}

// ============================================================
//  過去分の一括分類
// ============================================================
function tk一括分類_開始() {
  tkBulkClearTriggers_();
  var r = tkBulkRun_();
  var msg = '一括分類 開始: 今回 ' + r.processed + '件 / 残り ' + r.remaining + '件'
          + (r.processed ? ('（' + r.secPerItem + '秒/件）') : '');
  if (r.remaining > 0) {
    tkBulkInstallTrigger_();
    msg += '\n\n' + TK_BULK.EVERY_MIN + '分おきに自動で続きます。この画面は閉じて構いません。'
        + '\n残り時間の目安: 約' + Math.ceil(r.remaining * (r.secPerItem || 2) / 60 / 60 * 10) / 10 + '時間'
        + '\n進捗は tk一括分類_状況() で確認できます。';
  } else {
    msg += '\n\nすべて完了しました。';
  }
  tkLog_('一括分類', msg);
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return r;
}

function tk一括分類_継続() {
  var r = tkBulkRun_();
  tkLog_('一括分類', '継続: 今回 ' + r.processed + '件 / 残り ' + r.remaining + '件');

  if (r.remaining <= 0) {
    tkBulkClearTriggers_();
    tkBulkResetStall_();
    var done = '=== 一括分類 すべて完了 ===\n分類済み合計: ' + r.totalDone + '件';
    tkLog_('一括分類', done);
    try {
      if (TK.ADMIN_EMAIL) MailApp.sendEmail(TK.ADMIN_EMAIL, '[' + TK.COMPANY + ' 日報解析] 過去分の一括分類が完了しました', done);
    } catch (e) {}
    return;
  }

  // 1件も進まない実行が続いたら自分で止める（APIの上限などで延々と空回りしないように）
  if (r.processed > 0) { tkBulkResetStall_(); return; }
  var props = PropertiesService.getScriptProperties();
  var n = (Number(props.getProperty('tkBulkStall')) || 0) + 1;
  props.setProperty('tkBulkStall', String(n));
  if (n >= 3) {
    tkBulkClearTriggers_();
    tkBulkResetStall_();
    var msg = '⚠ 一括分類を自動停止しました（3回連続で1件も処理できませんでした）\n'
      + '残り ' + r.remaining + '件\n\n'
      + 'ほぼ確実にAPIキーの利用上限です（無料枠は1日20回）。\n'
      + 'tkAPIキー確認() を実行して、有料のキーかどうか確認してください。\n'
      + '有料キーに差し替えたら tk一括分類_開始() で再開できます（処理済みは再処理しません）。';
    tkLog_('一括分類', msg);
    Logger.log(msg);
    try { if (TK.ADMIN_EMAIL) MailApp.sendEmail(TK.ADMIN_EMAIL, '[' + TK.COMPANY + ' 日報解析] 一括分類が停止しました（APIの上限）', msg); } catch (e) {}
  }
}

function tkBulkResetStall_() {
  try { PropertiesService.getScriptProperties().deleteProperty('tkBulkStall'); } catch (e) {}
}

function tk一括分類_状況() {
  var s = tkBulkScan_();
  var msg = [
    '=== 分類の進捗 ===',
    '日報（回答）: ' + s.totalRows + '件',
    '分類済み: ' + s.doneCount + '件',
    '未判定の残り: ' + s.pending.length + '件',
    '継続トリガー: ' + (tkBulkHasTrigger_() ? 'あり（自動で進行中）' : 'なし（停止中）')
  ].join('\n');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return msg;
}

function tk一括分類_停止() {
  tkBulkClearTriggers_();
  var msg = '一括分類の自動継続を停止しました。（分類済みのデータはそのまま残ります）\n'
          + '再開したいときは tk一括分類_開始() を実行してください。';
  tkLog_('一括分類', '停止');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}

// ============================================================
//  失敗（HTTP 429 等）で「エラー」になった行を削除して、やり直せるようにする
//  → 削除後に tk一括分類_開始() を実行すると、その行だけ再判定される
// ============================================================
function tk一括分類_エラー行を削除() {
  var sh = tkEnsureSheet_(TK.SHEET_CLASS, TK_CLASS_HEADER);
  var last = sh.getLastRow();
  if (last < 2) { Logger.log('分類済みは空です'); return { deleted: 0 }; }

  var confCol = TK_CLASS_HEADER.indexOf('確信度') + 1;
  var codeCol = TK_CLASS_HEADER.indexOf('判定コード') + 1;
  var vals = sh.getRange(2, 1, last - 1, TK_CLASS_HEADER.length).getValues();

  var keep = [], deleted = 0, byConf = {};
  vals.forEach(function (r) {
    var conf = String(r[confCol - 1] || '').trim();
    var code = String(r[codeCol - 1] || '').trim();
    byConf[conf || '(空)'] = (byConf[conf || '(空)'] || 0) + 1;
    // 「エラー」または（コードが無く確信度も空）の行は捨てる＝再判定させる
    if (conf === 'エラー' || (!code && !conf)) { deleted++; return; }
    keep.push(r);
  });

  sh.getRange(2, 1, last - 1, TK_CLASS_HEADER.length).clearContent();
  if (keep.length) sh.getRange(2, 1, keep.length, TK_CLASS_HEADER.length).setValues(keep);

  var msg = ['=== エラー行の削除 ===',
    '削除（再判定の対象に戻した）: ' + deleted + '件',
    '残した（判定できていた）: ' + keep.length + '件',
    '',
    '内訳: ' + JSON.stringify(byConf),
    '',
    'このあと tk一括分類_開始() を実行すると、削除した分だけ再判定されます。'].join('\n');
  tkLog_('一括分類', msg);
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return { deleted: deleted, kept: keep.length, byConf: byConf };
}

// ---------- 未処理の洗い出し（古い順） ----------
function tkBulkScan_() {
  var raw = tkRawSheet_();
  if (!raw) throw new Error('回答シート（フォーム連携）が見つかりません');
  var out = tkEnsureSheet_(TK.SHEET_CLASS, TK_CLASS_HEADER);

  var done = {}, doneCount = 0;
  if (out.getLastRow() > 1) {
    out.getRange(2, 1, out.getLastRow() - 1, 1).getValues().forEach(function (r) {
      var v = String(r[0] || '').trim();
      if (v) { done[v] = 1; doneCount++; }
    });
  }

  var last = raw.getLastRow();
  var pending = [];
  if (last >= 2) {
    // 必要な列数だけ読む（実際の列数を超えないようにクランプ＝狭いシートでも落ちない）
    var need = Math.max(TK.RAW_TS, TK.RAW_EMAIL, TK.RAW_BODY, TK.RAW_NAME, TK.RAW_DEPT) + 1;
    var width = Math.max(1, Math.min(need, raw.getLastColumn()));
    var vals = raw.getRange(2, 1, last - 1, width).getValues();
    var since = TK_BULK.SINCE ? new Date(TK_BULK.SINCE + 'T00:00:00') : null;
    for (var i = 0; i < vals.length; i++) {          // 古い順
      var rowNo = i + 2;
      if (done[String(rowNo)]) continue;
      var text = String(vals[i][TK.RAW_BODY] || '').trim();
      if (!text) continue;
      if (since) {
        var d = vals[i][TK.RAW_TS];
        if (d instanceof Date && d < since) continue;
      }
      pending.push({ row: rowNo, vals: vals[i] });
    }
  }
  return { totalRows: Math.max(0, last - 1), doneCount: doneCount, pending: pending, out: out };
}

// ---------- 本体（バッチ処理） ----------
function tkBulkRun_() {
  var started = Date.now();
  var s = tkBulkScan_();
  if (!s.pending.length) return { processed: 0, remaining: 0, totalDone: s.doneCount, secPerItem: 0 };

  var quests = tkLoadQuests_();
  if (!quests.length) throw new Error('クエスト表が空です（先にSETUP_初期構築を実行）');
  var byCode = {};
  quests.forEach(function (q) { byCode[q.code] = q; });
  // 96項目すべてを提示する（部署は日報1件ごとに添えて、AI側で絞らせる）
  var questList = quests.map(function (q) {
    return q.code + '｜' + q.dept + '｜' + q.axis + '｜' + q.level + '｜' + q.title + '｜例:' + q.desc;
  }).join('\n');

  // ---- 1回だけ読む：活動記録の済みrow / メンバー表 ----
  var actSh = tkEnsureSheet_(TK.SHEET_ACT, TK_ACT_HEADER);
  var actDone = {};
  if (actSh.getLastRow() > 1) {
    actSh.getRange(2, 2, actSh.getLastRow() - 1, 1).getValues().forEach(function (r) {
      var v = String(r[0] || '').trim(); if (v) actDone[v] = 1;
    });
  }
  var memSh = tkEnsureSheet_(TK.SHEET_MEMBER, TK_MEMBER_HEADER);
  var mem = {}, memOrder = [];
  if (memSh.getLastRow() > 1) {
    var mv = memSh.getRange(2, 1, memSh.getLastRow() - 1, TK_MEMBER_HEADER.length).getValues();
    mv.forEach(function (r, i) {
      var id = String(r[0] || '').trim().toLowerCase();
      if (!id) return;
      mem[id] = { row: i + 2, name: String(r[1] || ''), dept: String(r[2] || ''),
        exp: Number(r[3]) || 0, reports: Number(r[4]) || 0, last: r[5] || '', stop: r[6] || '',
        emp: String(r[7] || '') };
      memOrder.push(id);
    });
  }

  var classRows = [], actRows = [], processed = 0, quota = false, calls = 0;

  for (var b = 0; b < s.pending.length; b += TK_BULK.BATCH) {
    var elapsed = (Date.now() - started) / 1000;
    if (elapsed > TK_BULK.RUN_SEC - TK_BULK.HEADROOM_SEC) break;    // 時間切れ→次回へ
    if (TK_BULK.MAX_BATCHES && calls >= TK_BULK.MAX_BATCHES) break; // API呼び出し数の上限
    calls++;

    var chunk = s.pending.slice(b, b + TK_BULK.BATCH);
    if (b > 0) Utilities.sleep(TK_BULK.PACE_MS);
    var judged = tkBulkClassifyChunk_(chunk, questList, mem);

    // APIの利用上限（429）に当たった → 記録せずに中断し、次のトリガーで再挑戦
    if (judged === null) { quota = true; break; }

    for (var j = 0; j < chunk.length; j++) {
      var p = chunk[j], v = p.vals;
      var text  = String(v[TK.RAW_BODY]  || '').trim();
      var email = String(v[TK.RAW_EMAIL] || '');
      var mid   = tkMemberId_(email);
      var res   = judged[j] || { code: '', alt: '', confidence: '低', reason: '判定できませんでした' };

      // 氏名・部署（メモリ上のメンバー表から補完）
      var name = String(v[TK.RAW_NAME] || '').trim();
      var dept = String(v[TK.RAW_DEPT] || '').trim();
      if ((!name || !dept) && mid && mem[mid]) {
        if (!name) name = mem[mid].name;
        if (!dept) dept = mem[mid].dept;
      }
      if (!name && mid) name = mid.split('@')[0];

      var q = byCode[String(res.code || '').trim()] || null;
      var ts = v[TK.RAW_TS] || new Date();

      // 「分類済み」16列: raw行/日時/メール/氏名/部署/本文/判定コード/観点/小見出し/対象/
      //                   確信度/判定理由/候補2/判定日時/活動記録済み/修正コード（手動）
      classRows.push([
        p.row, ts, email, name, dept, tkSafeCell_(text),
        q ? q.code : '', q ? q.axis : '', q ? q.title : '', q ? q.level : '',
        res.confidence || '', tkSafeCell_(res.reason || ''), res.alt || '',
        new Date(), q ? '済' : '', ''
      ]);

      if (q && !actDone[String(p.row)]) {
        var exp = tkExpFor_(res.confidence);
        // 「活動記録」12列: 日時/raw行/メンバーID/氏名/部署/観点/小見出し/対象/EXP/確信度/実践内容/判定コード
        actRows.push([ts, p.row, mid, name, dept, q.axis, q.title, q.level, exp,
          res.confidence || '', tkSafeCell_(String(text).slice(0, 500)), q.code]);
        actDone[String(p.row)] = 1;
        // メンバー集計をメモリ上で更新（シートへは最後に一括反映）
        if (mid) {
          if (!mem[mid]) { mem[mid] = { row: 0, name: name, dept: dept, exp: 0, reports: 0, last: '', stop: '', emp: '' }; memOrder.push(mid); }
          var m = mem[mid];
          if (name) m.name = name;
          if (dept) m.dept = dept;
          m.exp += exp; m.reports += 1;
          var t = ts ? new Date(ts) : null;
          if (t && (!m.last || t > new Date(m.last))) m.last = t;
        }
      }
      processed++;
    }
  }

  // ---- まとめて書き込み ----
  if (classRows.length) {
    s.out.getRange(s.out.getLastRow() + 1, 1, classRows.length, TK_CLASS_HEADER.length).setValues(classRows);
  }
  if (actRows.length) {
    actSh.getRange(actSh.getLastRow() + 1, 1, actRows.length, TK_ACT_HEADER.length).setValues(actRows);
  }
  // メンバー表：既存行は更新、新規はまとめて追記
  var newMem = [];
  memOrder.forEach(function (id) {
    var m = mem[id];
    var row = [id, m.name, m.dept, m.exp, m.reports, m.last || '', m.stop || '', m.emp || ''];
    if (m.row > 0) memSh.getRange(m.row, 1, 1, TK_MEMBER_HEADER.length).setValues([row]);
    else newMem.push(row);
  });
  if (newMem.length) {
    memSh.getRange(memSh.getLastRow() + 1, 1, newMem.length, TK_MEMBER_HEADER.length).setValues(newMem);
  }

  var sec = (Date.now() - started) / 1000;
  if (quota) tkLog_('一括分類', 'APIの利用上限に当たったため中断しました（今回 ' + processed + '件）。次回のトリガーで自動的に再挑戦します。');
  return {
    processed: processed,
    remaining: Math.max(0, s.pending.length - processed),
    totalDone: s.doneCount + processed,
    secPerItem: processed ? Math.round(sec / processed * 10) / 10 : 0,
    quota: quota
  };
}

// 15件まとめてGeminiに判定させる → [{code,alt,confidence,reason}, ...]（入力順）
function tkBulkClassifyChunk_(chunk, questList, mem) {
  var items = chunk.map(function (p, i) {
    var v = p.vals;
    var t = String(v[TK.RAW_BODY] || '').replace(/\s+/g, ' ').slice(0, 500);
    // 部署は基準の絞り込みに効くので、日報1件ごとに添える
    var dept = String(v[TK.RAW_DEPT] || '').trim();
    if (!dept) {
      var mid = tkMemberId_(String(v[TK.RAW_EMAIL] || ''));
      if (mid && mem && mem[mid]) dept = mem[mid].dept || '';
    }
    return '[' + (i + 1) + ']（部署: ' + (dept || '不明') + '） ' + t;
  }).join('\n');

  var prompt =
    'あなたはSUNグループ（介護・障がい・薬局・本部）の日報を、人事評価制度「サンクスUP！」の\n' +
    '評価基準（成果・能力・姿勢）に分類する係です。\n' +
    '以下は評価基準の一覧です（形式: コード｜部署｜観点｜対象｜小見出し｜記入例）:\n' + questList + '\n\n' +
    '次に日報が ' + chunk.length + '件あります。**1件ずつ**、最もよく当てはまる基準を1つ選び、\n' +
    'results 配列に入力と同じ順番・同じ件数で返してください。\n' +
    '各日報には報告者の部署が付いています。**その部署の基準の中から選んでください**\n' +
    '（部署が「不明」のものだけ、全部署の基準から選んでよい）。\n' +
    '- n: 日報の番号（1から）\n' +
    '- code: 最有力の基準コード。どうしても該当がなければ ""\n' +
    '- alt: 2番目に近いコード（なければ ""）\n' +
    '- confidence: "高"|"中"|"低"\n' +
    '- reason: 判定理由をひとこと（30字以内）\n\n' +
    '日報:\n' + items;

  var schema = {
    type: 'OBJECT',
    properties: {
      results: {
        type: 'ARRAY',
        items: {
          type: 'OBJECT',
          properties: {
            n: { type: 'INTEGER' }, code: { type: 'STRING' }, alt: { type: 'STRING' },
            confidence: { type: 'STRING' }, reason: { type: 'STRING' }
          },
          required: ['n', 'code', 'confidence']
        }
      }
    },
    required: ['results']
  };

  // 429（利用上限）/5xx は少し待って再試行。それでもダメなら null を返して呼び出し側に中断させる
  var g = null;
  for (var attempt = 0; attempt <= TK_BULK.RETRY; attempt++) {
    g = tkGeminiJson_(prompt, schema);
    if (g.ok) break;
    var err = String(g.error || '');
    var busy = /HTTP 429|HTTP 5\d\d|quota|rate/i.test(err);
    if (!busy) break;                                  // 上限以外の失敗は再試行しない
    if (attempt < TK_BULK.RETRY) Utilities.sleep(TK_BULK.RETRY_WAIT_MS);
    else return null;                                  // ← 上限のまま：記録せず中断
  }

  var out = [];
  if (g.ok && g.data && g.data.results) {
    // n をたよりに入力順へ並べ直す（欠けは空判定）
    var byN = {};
    g.data.results.forEach(function (r) {
      var n = Number(r.n); if (n >= 1) byN[n] = r;
    });
    for (var i = 0; i < chunk.length; i++) {
      var r = byN[i + 1];
      out.push(r ? { code: String(r.code || ''), alt: String(r.alt || ''),
                     confidence: String(r.confidence || '中'), reason: String(r.reason || '') }
                 : { code: '', alt: '', confidence: '低', reason: 'AI応答が不足' });
    }
  } else {
    // 失敗時：エラーとして記録（tk一括分類_エラー行を削除() で再判定できる）
    for (var k = 0; k < chunk.length; k++) {
      out.push({ code: '', alt: '', confidence: 'エラー', reason: String((g && g.error) || 'Gemini失敗') });
    }
  }
  return out;
}

// ============================================================
//  APIキーの診断：いま設定されているキーが「有料(Tier1)」か「無料枠」かを判定
//  ------------------------------------------------------------
//  キーの値をどこにも表示せず、上限の種類だけを確認する。
//  ※ 3回テストでは判別できない（無料枠でも20回までは通る）ので25回連打する。
// ============================================================
function tkAPIキー確認() {
  var keys = String(SECRET_CONFIG.GEMINI_API_KEYS || '').split(',')
    .map(function (s) { return s.trim(); }).filter(String);
  if (!keys.length) { Logger.log('APIキーが設定されていません'); return; }

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
    SECRET_CONFIG.GEMINI_MODEL + ':generateContent';
  var payload = JSON.stringify({
    contents: [{ parts: [{ text: '1+1' }] }],
    generationConfig: { maxOutputTokens: 8 }
  });

  var PROBE = 25;   // 無料枠(20/日)を確実に見抜くには20回超の連打が必要
  var lines = ['=== APIキー診断（' + SECRET_CONFIG.GEMINI_MODEL + '・最大' + PROBE + '回連打） ==='];
  keys.forEach(function (k, idx) {
    var label = 'キー' + (idx + 1) + '（長さ' + k.length + '・末尾 ' + k.slice(-4) + '）';
    var ok = 0, quotaId = '', quotaVal = '', err = '', msg = '', httpCode = 0;
    for (var i = 0; i < PROBE; i++) {
      try {
        var res = UrlFetchApp.fetch(url + '?key=' + encodeURIComponent(k), {
          method: 'post', contentType: 'application/json',
          payload: payload, muteHttpExceptions: true
        });
        var code = res.getResponseCode();
        if (code === 200) { ok++; continue; }
        httpCode = code;
        var body = {};
        try { body = JSON.parse(res.getContentText()); } catch (e) {}
        var e2 = body.error || {};
        err = 'HTTP ' + code + ' ' + (e2.status || '');
        msg = String(e2.message || '').slice(0, 400);
        (e2.details || []).forEach(function (d) {
          if (String(d['@type'] || '').indexOf('QuotaFailure') >= 0) {
            (d.violations || []).forEach(function (v) {
              quotaId = String(v.quotaId || ''); quotaVal = String(v.quotaValue || '');
            });
          }
        });
        break;
      } catch (e) { err = String(e && e.message || e); break; }
    }
    var verdict;
    if (ok >= PROBE) verdict = '✅ 有料です（' + ok + '回連続OK）→ 大量処理に使えます';
    else if (/FreeTier/i.test(quotaId)) verdict = '❌ 無料枠です（' + ok + '回で停止・上限 ' + quotaVal +
      '）→ このキーでは大量処理できません';
    else if (quotaId) verdict = '⚠ ' + ok + '回で上限（' + quotaId + ' = ' + quotaVal + '）';
    else verdict = '❌ ' + ok + '回でエラー: ' + err;
    lines.push(label + ': ' + verdict);
    if (msg) lines.push('　理由: ' + msg);

    // 404/403 のときは「このキーで使えるモデル一覧」を出す（原因の切り分け用）
    if (httpCode === 404 || httpCode === 403) {
      try {
        var lm = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?key=' +
          encodeURIComponent(k) + '&pageSize=100', { muteHttpExceptions: true });
        if (lm.getResponseCode() === 200) {
          var ms = (JSON.parse(lm.getContentText()).models || [])
            .filter(function (m) { return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0; })
            .map(function (m) { return String(m.name || '').replace('models/', ''); });
          var flash = ms.filter(function (n) { return /flash/i.test(n) && !/lite|thinking|image|tts|live/i.test(n); });
          lines.push('　使えるモデル数: ' + ms.length);
          lines.push('　flash系: ' + (flash.slice(0, 8).join(', ') || '（なし）'));
          lines.push('　→ 上の名前のどれかを SECRET_CONFIG.GEMINI_MODEL に設定してください');
        } else {
          var lb = {}; try { lb = JSON.parse(lm.getContentText()); } catch (e) {}
          lines.push('　モデル一覧も取得できません: HTTP ' + lm.getResponseCode() +
            ' ' + String((lb.error && lb.error.message) || '').slice(0, 200));
          lines.push('　→ そのキーのプロジェクトで「Generative Language API」が有効か確認してください');
        }
      } catch (e) { lines.push('　モデル一覧の取得に失敗: ' + (e && e.message)); }
    }
  });
  lines.push('');
  lines.push('※「無料枠(20/日)」なら有料プロジェクトのキーに差し替え。');
  lines.push('※ 404なら、上の「flash系」に出たモデル名に GEMINI_MODEL を合わせるか、APIの有効化を確認。');
  var out = lines.join('\n');
  tkLog_('APIキー診断', out);
  Logger.log(out);
  try { SpreadsheetApp.getUi().alert(out); } catch (e) {}
  return out;
}

// ---------- トリガー ----------
function tkBulkInstallTrigger_() {
  tkBulkClearTriggers_();
  ScriptApp.newTrigger(TK_BULK.TRIGGER_FN).timeBased().everyMinutes(TK_BULK.EVERY_MIN).create();
}
function tkBulkClearTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === TK_BULK.TRIGGER_FN) ScriptApp.deleteTrigger(t);
  });
}
function tkBulkHasTrigger_() {
  return ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === TK_BULK.TRIGGER_FN;
  });
}
