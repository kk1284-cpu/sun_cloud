// ============================================================
// gas_takamura ／ [70] 週次GCシートの下書き自動生成
// ------------------------------------------------------------
//  日報（活動記録）から、既存の週次GCシートの記入内容を下書きする。
//  ねらい：GCで最も負担の大きい「今週何をやったか思い出して書く」をなくし、
//         会議の時間を"報告"から"対話"に使えるようにする。
//
//  埋める欄（既存GCシートの項目に対応）:
//    ・今週の中心課題・取り組むテーマ
//    ・何をどのように実践したか
//    ・外に生み出した成果 →「直接の成果／価値への取り組み／人材育成」の3分類
//        ※ 3分類はAIの推測ではなく、クエスト表の「GC分類」列による事実の振り分け
//    ・鍵となる発見・気づき／予期せぬ成功・失敗
//    ・用いたフィロソフィ
//    ・すぐに取り組むべきこと
//    ・今週のKey Takeaway
//
//  使い方（エディタで関数を選んで▷実行）:
//    tkGC下書き_今週()     … 今週（月〜日）ぶんを生成
//    tkGC下書き_先週()     … 先週ぶんを生成
//    tkGC下書き_トリガー設定() … 毎週金曜17時に自動生成
//
//  出力先: シート「GC下書き」＋（TK.GC_MAIL=true なら本人へメール）
// ============================================================

var TK_GC = {
  SHEET: 'GC下書き',
  HEADER: ['週', 'メンバーID', '氏名', '部署', '日報件数',
    '今週の中心課題・取り組むテーマ', '何をどのように実践したか',
    '直接の成果', '価値への取り組み', '人材育成',
    '鍵となる発見・気づき／予期せぬ成功・失敗', '用いたフィロソフィ',
    'すぐに取り組むべきこと', '今週のKey Takeaway', '生成日時'],
  ORDER: ['直接の成果', '価値への取り組み', '人材育成'],
  MAIL: false,          // true にすると本人にも下書きをメール送信
  MAX_MEMBERS: 40       // 1回の実行で処理する上限（実行時間対策）
};

// ---------- 週の範囲 ----------
//  opt.endDate を渡すと「その日までの7日間」（GC前日までの7日間＝部署別の曜日運用）。
//  渡さなければ従来どおり 月〜日。
function tkGCWeekRange_(offsetWeeks, opt) {
  opt = opt || {};
  if (opt.endDate) {
    var e = new Date(opt.endDate);
    var to = new Date(e.getFullYear(), e.getMonth(), e.getDate(), 23, 59, 59);
    var from = new Date(e.getFullYear(), e.getMonth(), e.getDate() - 6);
    return { from: from, to: to };
  }
  var now = new Date();
  var day = now.getDay();                      // 0=日
  var mondayDiff = (day === 0 ? -6 : 1 - day);  // 月曜まで戻す
  var mon = new Date(now.getFullYear(), now.getMonth(), now.getDate() + mondayDiff);
  mon.setDate(mon.getDate() + (offsetWeeks || 0) * 7);
  var sun = new Date(mon.getFullYear(), mon.getMonth(), mon.getDate() + 6, 23, 59, 59);
  return { from: mon, to: sun };
}
function tkGCWeekLabel_(r) {
  var f = Utilities.formatDate(r.from, TK.TZ, 'M/d');
  var t = Utilities.formatDate(r.to, TK.TZ, 'M/d');
  return Utilities.formatDate(r.from, TK.TZ, 'yyyy') + ' ' + f + '〜' + t;
}

function tkGC下書き_今週() { return tkGCBuild_(0); }
function tkGC下書き_先週() { return tkGCBuild_(-1); }
function tkGC週次トリガー() { tkGCBuild_(0); }   // 旧：金曜17時固定。⑭で日次（90_gcslide の tkGC日次トリガー）に置き換わる

function tkGC下書き_トリガー設定() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'tkGC週次トリガー') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('tkGC週次トリガー').timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(17).inTimezone(TK.TZ).create();
  var msg = '毎週金曜17時に「GC下書き」を自動生成する設定にしました。';
  tkLog_('GC下書き', msg); Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}

// ---------- 本体 ----------
//  opt.range  … 期間の上書き（日次トリガーが渡す）
//  opt.group  … 部署またはソースIDで絞る（GC設定の1行ぶん）
//  opt.silent … ダイアログを出さない（トリガー用）
function tkGCBuild_(offsetWeeks, opt) {
  opt = opt || {};
  var range = opt.range || tkGCWeekRange_(offsetWeeks);
  var label = tkGCWeekLabel_(range);
  var ps = opt.group ? tkSetPersonSources_() : null;

  // クエスト表（GC分類つき）
  var quests = tkLoadQuests_();
  var gcByCode = {};
  quests.forEach(function (q) { gcByCode[q.code] = q.gc; });

  // 今週の活動記録をメンバー別に集める
  var acts = tkReadActivity_().filter(function (a) {
    if (!a.ts) return false;
    var t = new Date(a.ts);
    if (t < range.from || t > range.to) return false;
    return opt.group ? tkSetGroupMatch_(opt.group, a, ps) : true;
  });
  if (!acts.length) {
    var none = '【' + label + (opt.group ? '／' + opt.group : '') + '】該当期間の活動記録がありません（日報が入っていないか、まだ分類されていません）。';
    tkLog_('GC下書き', none); Logger.log(none);
    if (!opt.silent) { try { SpreadsheetApp.getUi().alert(none); } catch (e) {} }
    return { members: 0 };
  }

  var per = {};
  acts.forEach(function (a) {
    if (!a.mid) return;
    if (!per[a.mid]) per[a.mid] = { name: a.name, dept: a.dept, items: [] };
    if (a.name) per[a.mid].name = a.name;
    if (a.dept) per[a.mid].dept = a.dept;
    per[a.mid].items.push(a);
  });

  var sh = tkEnsureSheet_(TK_GC.SHEET, TK_GC.HEADER);
  var rows = [], done = 0, dup = 0;
  var ids = Object.keys(per).slice(0, TK_GC.MAX_MEMBERS);

  // 同じ週・同じ人の行が既にあれば作らない（作り直しで行が二重に増えないように）
  var have = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      have[String(r[0]) + '#' + tkMemberId_(r[1])] = 1;
    });
  }

  ids.forEach(function (mid) {
    var p = per[mid];
    if (have[label + '#' + mid]) { dup++; return; }

    // ① 3分類は「事実の振り分け」（AIに推測させない）
    var buckets = { '直接の成果': [], '価値への取り組み': [], '人材育成': [] };
    p.items.forEach(function (a) {
      var gc = gcByCode[a.code] || (a.axis === '成果' ? '直接の成果'
        : a.axis === '姿勢' ? '価値への取り組み' : '人材育成');
      if (!buckets[gc]) buckets[gc] = [];
      buckets[gc].push('・' + (a.title || a.skill || '') + '：' + String(a.practice || '').slice(0, 120));
    });

    // ② 文章欄はAIが下書き
    var d = tkGCDraftText_(p, range);

    rows.push([label, mid, p.name, p.dept, p.items.length,
      d.theme, d.practice,
      buckets['直接の成果'].join('\n'),
      buckets['価値への取り組み'].join('\n'),
      buckets['人材育成'].join('\n'),
      d.insight, d.philosophy, d.next, d.takeaway, new Date()]);
    done++;

    if (TK_GC.MAIL && mid) {
      try { tkGCMail_(mid, p, label, buckets, d); } catch (e) {}
    }
  });

  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, TK_GC.HEADER.length).setValues(rows);
  }

  var msg = ['=== 週次GCシート 下書き生成 ===',
    '対象週: ' + label + (opt.group ? '　対象: ' + opt.group : ''),
    '生成: ' + done + '名（活動記録 ' + acts.length + '件）' + (dup ? '　既にあり: ' + dup + '名' : ''),
    '出力先: シート「' + TK_GC.SHEET + '」',
    '',
    '※「直接の成果／価値への取り組み／人材育成」はクエスト表の【GC分類】列に沿った',
    '　事実の振り分けです（AIの推測ではありません）。',
    '※ 文章欄はAIの下書きです。ご本人が確認・微修正してGCシートに転記してください。'].join('\n');
  tkLog_('GC下書き', 'GC下書き生成 ' + done + '名' + (dup ? '・既存 ' + dup : '') + '（' + label + (opt.group ? '／' + opt.group : '') + '）');
  Logger.log(msg);
  if (!opt.silent) { try { SpreadsheetApp.getUi().alert(msg); } catch (e) {} }
  return { members: done, week: label, dup: dup };
}

// AIで文章欄を下書き
function tkGCDraftText_(p, range) {
  var list = p.items.map(function (a) {
    return '・[' + (a.axis || '') + '/' + (a.title || '') + '] ' + String(a.practice || '').slice(0, 150);
  }).join('\n').slice(0, 6000);

  var prompt =
    'あなたは有限会社高村（SUNグループ）の週次グループコーチング（GC）を支援する role です。\n' +
    'ドラッカーの考え方（成果は「直接の成果」「価値への取り組み」「人材育成」の3つ）と、\n' +
    '同社の評価制度「サンクスUP！」を前提にしてください。\n\n' +
    '以下は ' + (p.name || '本人') + ' さん（' + (p.dept || '') + '）の今週の日報から抽出した実践の記録です。\n' +
    'これをもとに、週次GCシートの各欄の**下書き**を作ってください。\n' +
    '本人が読んで「そうそう」と思える、事実に基づいた素直な文章にしてください。\n' +
    '推測で話を大きくしないこと。記録に無いことは書かないこと。\n\n' +
    '記録:\n' + list + '\n\n' +
    '出力する項目:\n' +
    '- theme: 今週の中心課題・取り組むテーマ（40字以内）\n' +
    '- practice: 何をどのように実践したか（150字程度・箇条書きでなく文章）\n' +
    '- insight: 鍵となる発見・気づき／予期せぬ成功・失敗（100字以内。記録から読み取れる範囲で）\n' +
    '- philosophy: 用いたフィロソフィ（同社の価値観に照らして。40字以内）\n' +
    '- next: すぐに取り組むべきこと（60字以内）\n' +
    '- takeaway: 今週のKey Takeaway（40字以内・一言）';

  var schema = {
    type: 'OBJECT',
    properties: {
      theme: { type: 'STRING' }, practice: { type: 'STRING' }, insight: { type: 'STRING' },
      philosophy: { type: 'STRING' }, next: { type: 'STRING' }, takeaway: { type: 'STRING' }
    },
    required: ['theme', 'practice', 'takeaway']
  };

  var g = tkGeminiJson_(prompt, schema);
  if (!g.ok) {
    return { theme: '', practice: '（AIの下書きを取得できませんでした：' + (g.error || '') + '）',
      insight: '', philosophy: '', next: '', takeaway: '' };
  }
  var d = g.data || {};
  return {
    theme: tkSafeCell_(d.theme || ''), practice: tkSafeCell_(d.practice || ''),
    insight: tkSafeCell_(d.insight || ''), philosophy: tkSafeCell_(d.philosophy || ''),
    next: tkSafeCell_(d.next || ''), takeaway: tkSafeCell_(d.takeaway || '')
  };
}

// 本人へメール（TK_GC.MAIL = true のとき）
function tkGCMail_(mid, p, label, buckets, d) {
  var b = function (k) { return buckets[k] && buckets[k].length ? buckets[k].join('\n') : '（今週は該当なし）'; };
  var body = [
    (p.name || '') + ' さん',
    '',
    '【' + label + '】週次GCシートの下書きです（日報から自動作成）。',
    'そのまま使える部分はそのまま、違うところは直してお使いください。',
    '',
    '■ 今週の中心課題・取り組むテーマ',
    d.theme,
    '',
    '■ 何をどのように実践したか',
    d.practice,
    '',
    '■ 外に生み出した成果',
    '［直接の成果］', b('直接の成果'),
    '［価値への取り組み］', b('価値への取り組み'),
    '［人材育成］', b('人材育成'),
    '',
    '■ 鍵となる発見・気づき／予期せぬ成功・失敗',
    d.insight,
    '',
    '■ 用いたフィロソフィ',
    d.philosophy,
    '',
    '■ すぐに取り組むべきこと',
    d.next,
    '',
    '■ 今週のKey Takeaway',
    d.takeaway,
    '',
    '---',
    '今週の日報 ' + p.items.length + '件から作成しました。',
    '（' + TK.PRAISE_CHARA + '）'
  ].join('\n');
  MailApp.sendEmail({ to: mid, subject: '【GC下書き】' + label + '　' + (p.name || ''), body: body });
}
