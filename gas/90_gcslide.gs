// ============================================================
// gas_takamura ／ [90] 週次GCシート（Googleスライド）へ下書きを直接書き込む
// ------------------------------------------------------------
//  ねらい：シート「GC下書き」を見て転記する手間もなくす。
//         **顧客が実際に使っているスライドの形のまま**、個人フォルダに
//         「下書きが入った今週のGCシート」を毎週金曜に置く。
//
//  ★絶対に守ること★
//   1) 顧客原本（TK.GC_SLIDE.MASTER_ID）には **一切書き込まない**。読むだけ。
//      原本を1回コピーして作った「下書き用テンプレ」を毎週複製して使う。
//   2) 本人が書く欄にはトークンを置かない（空のまま渡す）。
//      ＝外的な変化／内的な変化／目指していること／私のフィロソフィ目標／
//        実践に用いたドラッカーのコンセプト／効果的な計画／フィードバックコメント
//   3) **EXPは書かない。** 内省の場に点数を持ち込まない（既定の設計判断）。
//
//  なぜトークン方式か：
//   スライドの回答欄は「空のテキストボックス」なので、文字列置換では狙えない。
//   objectId を固定する方法もあるが、テンプレを編集した瞬間に壊れる。
//   テンプレに {{THEME}} 等を1回書いておけば、以後レイアウトを自由に変えても
//   コードは壊れない。運用でいちばん壊れにくい。
//
//  手順（初回だけ）:
//    ① tkGCスライド_構造を調べる()     … 原本のスライド構成とテキストを一覧（どこに何があるか把握）
//    ② tkGCスライド_下書きテンプレを用意() … 原本をコピーして作業フォルダに置く
//    ③ ②で作ったコピーを開き、各回答欄に下のトークンを1つずつ貼る（人の作業・1回だけ）
//    ④ config の TK.GC_SLIDE.TEMPLATE_ID に②のIDを設定
//    ⑤ tkGCスライド_個人フォルダ準備()  … メンバーごとのフォルダを作って本人に共有
//  毎週:
//    tkGCスライド_今週分を作る() / tkGCスライド_先週分を作る()
//    tkGCスライド_トリガー設定()  … 毎週金曜17時30分（シートのGC下書きの30分後）
// ============================================================

// テンプレに貼るトークン（← ③でこれを回答欄に貼る）
var TK_GCS_TOKENS = {
  '{{NAME}}':          '氏名',
  '{{DEPT}}':          '部署',
  '{{WEEK}}':          '対象週',
  '{{COUNT}}':         '今週の日報件数',
  '{{THEME}}':         '今週の中心課題・取り組むテーマは何か',
  '{{PRACTICE}}':      '何をどのように実践したか',
  '{{RESULT_DIRECT}}': '外に生み出した成果 ／ 直接の成果',
  '{{RESULT_VALUE}}':  '外に生み出した成果 ／ 価値への取り組み',
  '{{RESULT_PEOPLE}}': '外に生み出した成果 ／ 人材育成',
  '{{INSIGHT}}':       '鍵となる発見・気づきは何か／予期せぬ成功・失敗は？',
  '{{PHILOSOPHY}}':    '用いたフィロソフィは',
  '{{NEXT}}':          'すぐに取り組むべきことは何か',
  '{{TAKEAWAY}}':      '今週のＫey Takeawayは何か'
};

// ============================================================
//  ① 原本の構造を調べる（読むだけ・書き込まない）
// ============================================================
function tkGCスライド_構造を調べる() {
  var id = TK.GC_SLIDE && TK.GC_SLIDE.MASTER_ID;
  if (!id) throw new Error('TK.GC_SLIDE.MASTER_ID が未設定です');
  var pres = SlidesApp.openById(id);          // ← 読み取りのみ。save/変更はしない
  var lines = ['=== ' + pres.getName() + ' ===', 'スライド数: ' + pres.getSlides().length, ''];

  pres.getSlides().forEach(function (sl, i) {
    lines.push('── スライド ' + (i + 1) + '（id: ' + sl.getObjectId() + '）');
    sl.getShapes().forEach(function (sh) {
      var t = '';
      try { t = String(sh.getText().asString() || '').replace(/\s+/g, ' ').trim(); } catch (e) {}
      lines.push('   [' + sh.getObjectId() + '] ' + (t ? t.slice(0, 60) : '（空欄）'));
    });
    if (i >= 8) { lines.push('   …以降は省略'); return; }
  });

  lines.push('');
  lines.push('※「（空欄）」が回答欄です。ここにトークンを貼ります。');
  Object.keys(TK_GCS_TOKENS).forEach(function (k) {
    lines.push('   ' + k + ' → ' + TK_GCS_TOKENS[k]);
  });

  var msg = lines.join('\n');
  tkLog_('GCスライド', 'テンプレ構造を確認（スライド' + pres.getSlides().length + '枚）');
  Logger.log(msg);
  return msg;
}

// ============================================================
//  ② 原本をコピーして「下書き用テンプレ」を作る（原本は無傷）
// ============================================================
function tkGCスライド_下書きテンプレを用意() {
  var c = TK.GC_SLIDE;
  if (!c || !c.MASTER_ID) throw new Error('TK.GC_SLIDE.MASTER_ID が未設定です');
  var work = tkGCSWorkFolder_();
  var name = '_GC下書きテンプレ（トークン入り・自動生成用）';

  // 既にあれば作り直さない（人がトークンを入れた作業を消さないため）
  var ex = work.getFilesByName(name);
  if (ex.hasNext()) {
    var f0 = ex.next();
    var m0 = 'すでにテンプレがあります。作り直しません。\n\nID: ' + f0.getId() + '\n' + f0.getUrl();
    Logger.log(m0);
    try { SpreadsheetApp.getUi().alert(m0); } catch (e) {}
    return { id: f0.getId(), created: false };
  }

  var copy = DriveApp.getFileById(c.MASTER_ID).makeCopy(name, work);
  var lines = ['=== 下書き用テンプレを作成しました ===',
    'ID: ' + copy.getId(), copy.getUrl(), '',
    '次にやること（人の作業・1回だけ）:',
    ' 1. 上のURLを開く',
    ' 2. 週次GCシートのスライドの各回答欄に、次のトークンを1つずつ貼る',
    ''];
  Object.keys(TK_GCS_TOKENS).forEach(function (k) {
    lines.push('    ' + k + '  →  ' + TK_GCS_TOKENS[k]);
  });
  lines.push('');
  lines.push(' 3. 本人が書く欄には **何も貼らない**（空のまま）:');
  lines.push('    外的な変化／内的な変化／目指していること・探求していること／');
  lines.push('    私のフィロソフィ目標／実践に用いたドラッカーのコンセプト／');
  lines.push('    効果的な計画として、どのようなことが考えられるか／フィードバックコメント');
  lines.push('');
  lines.push(' 4. config の TK.GC_SLIDE.TEMPLATE_ID に上のIDを設定する');
  lines.push('');
  lines.push('※ 週次スライドが複数枚ある場合は、1枚だけトークンを入れて他は削除してください');
  lines.push('　（毎週このテンプレを1枚コピーして使います）');

  var msg = lines.join('\n');
  tkLog_('GCスライド', '下書きテンプレ作成 ' + copy.getId());
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return { id: copy.getId(), created: true };
}

// ============================================================
//  ⑤ 個人フォルダを作って本人に共有する
// ============================================================
function tkGCスライド_個人フォルダ準備() {
  var work = tkGCSWorkFolder_();
  var sh = tkEnsureSheet_(TK.SHEET_MEMBER, TK_MEMBER_HEADER);
  if (sh.getLastRow() < 2) {
    var m = 'メンバーシートが空です。先に名簿を登録してください。';
    Logger.log(m); try { SpreadsheetApp.getUi().alert(m); } catch (e) {}
    return { made: 0 };
  }
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  var made = 0, shared = 0, lines = ['=== 個人フォルダの準備 ==='];

  vals.forEach(function (r) {
    var mid = String(r[0] || '').trim().toLowerCase();
    var name = String(r[1] || '').trim() || (mid ? mid.split('@')[0] : '');
    if (!mid) return;
    var folder = tkGCSPersonFolder_(work, name, mid);
    made++;
    if (TK.GC_SLIDE.SHARE_WITH_MEMBER) {
      try { folder.addEditor(mid); shared++; }
      catch (e) { lines.push('　⚠ 共有できませんでした: ' + mid + '（' + (e && e.message) + '）'); }
    }
    lines.push('　' + name + '（' + mid + '） → ' + folder.getUrl());
  });

  lines.push('');
  lines.push('作成/確認: ' + made + '名　共有: ' + shared + '名');
  var msg = lines.join('\n');
  tkLog_('GCスライド', '個人フォルダ ' + made + '名');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return { made: made, shared: shared };
}

// ============================================================
//  毎週：個人フォルダに「下書き入りの今週のGCシート」を置く
// ============================================================
function tkGCスライド_今週分を作る() { return tkGCSBuild_(0); }
function tkGCスライド_先週分を作る() { return tkGCSBuild_(-1); }
function tkGCスライド週次トリガー() { tkGCSBuild_(0); }

function tkGCスライド_トリガー設定() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'tkGCスライド週次トリガー') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('tkGCスライド週次トリガー').timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(17).nearMinute(30).inTimezone(TK.TZ).create();
  var msg = '毎週金曜17時30分に、個人フォルダへ下書き入りGCシートを作る設定にしました。\n'
          + '（シート「GC下書き」の生成は17時。その30分後に走ります）';
  tkLog_('GCスライド', msg); Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}

function tkGCSBuild_(offsetWeeks) {
  var c = TK.GC_SLIDE;
  if (!c || !c.TEMPLATE_ID) {
    throw new Error('TK.GC_SLIDE.TEMPLATE_ID が未設定です。先に tkGCスライド_下書きテンプレを用意() を実行し、'
      + 'トークンを貼ってIDを設定してください。');
  }
  if (c.TEMPLATE_ID === c.MASTER_ID) {
    throw new Error('TEMPLATE_ID が顧客原本と同じです。原本を書き換える事故になるため中止しました。');
  }

  var range = tkGCWeekRange_(offsetWeeks);       // 70_gc.gs と同じ週の切り方
  var label = tkGCWeekLabel_(range);
  var per = tkGCSCollect_(range);                // メンバー別の材料
  var ids = Object.keys(per);
  if (!ids.length) {
    var none = '【' + label + '】該当期間の活動記録がありません。';
    tkLog_('GCスライド', none); Logger.log(none);
    try { SpreadsheetApp.getUi().alert(none); } catch (e) {}
    return { members: 0 };
  }

  var work = tkGCSWorkFolder_();
  var done = 0, skipped = 0, lines = ['=== 下書き入りGCシートの作成（' + label + '） ==='];

  ids.slice(0, TK_GC.MAX_MEMBERS).forEach(function (mid) {
    var p = per[mid];
    var folder = tkGCSPersonFolder_(work, p.name, mid);
    var fname = 'GCシート ' + label.replace(/\//g, '-') + '　' + (p.name || mid);

    // 同じ週のものが既にあれば作らない（本人が手を入れたものを潰さない）
    if (folder.getFilesByName(fname).hasNext()) {
      skipped++; lines.push('　スキップ（既存）: ' + fname);
      return;
    }

    var copy = DriveApp.getFileById(c.TEMPLATE_ID).makeCopy(fname, folder);
    var pres = SlidesApp.openById(copy.getId());

    var d = tkGCDraftText_(p, range);            // 70_gc.gs のAI下書きをそのまま使う
    var b = function (k) { return p.buckets[k] && p.buckets[k].length ? p.buckets[k].join('\n') : '（今週は該当なし）'; };

    var map = {
      '{{NAME}}':          p.name || mid,
      '{{DEPT}}':          p.dept || '',
      '{{WEEK}}':          label,
      '{{COUNT}}':         String(p.items.length) + '件',
      '{{THEME}}':         d.theme || '',
      '{{PRACTICE}}':      d.practice || '',
      '{{RESULT_DIRECT}}': b('直接の成果'),
      '{{RESULT_VALUE}}':  b('価値への取り組み'),
      '{{RESULT_PEOPLE}}': b('人材育成'),
      '{{INSIGHT}}':       d.insight || '',
      '{{PHILOSOPHY}}':    d.philosophy || '',
      '{{NEXT}}':          d.next || '',
      '{{TAKEAWAY}}':      d.takeaway || ''
    };
    Object.keys(map).forEach(function (k) {
      try { pres.replaceAllText(k, map[k], false); } catch (e) {}
    });
    pres.saveAndClose();

    if (c.SHARE_WITH_MEMBER) { try { copy.addEditor(mid); } catch (e) {} }
    if (c.MAIL) { try { tkGCSMail_(mid, p, label, copy); } catch (e) {} }

    done++;
    lines.push('　作成: ' + fname + '　' + copy.getUrl());
  });

  lines.push('');
  lines.push('作成 ' + done + '名／スキップ ' + skipped + '名');
  lines.push('※ 3分類は事実の振り分け、文章欄はAIの下書きです。ご本人が確認・微修正してお使いください。');
  lines.push('※ 外的/内的な変化・ドラッカーのコンセプト・フィロソフィ目標は空欄のままです（ご本人の記入欄）。');
  var msg = lines.join('\n');
  tkLog_('GCスライド', '作成 ' + done + '名・スキップ ' + skipped + '名（' + label + '）');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return { members: done, skipped: skipped, week: label };
}

// ---------- 材料集め（70_gc.gs の集計と同じ考え方） ----------
function tkGCSCollect_(range) {
  var quests = tkLoadQuests_();
  var gcByCode = {};
  quests.forEach(function (q) { gcByCode[q.code] = q.gc; });

  var acts = tkReadActivity_().filter(function (a) {
    if (!a.ts) return false;
    var t = new Date(a.ts);
    return t >= range.from && t <= range.to;
  });

  var per = {};
  acts.forEach(function (a) {
    if (!a.mid) return;
    if (!per[a.mid]) {
      per[a.mid] = { name: a.name, dept: a.dept, items: [],
        buckets: { '直接の成果': [], '価値への取り組み': [], '人材育成': [] } };
    }
    var p = per[a.mid];
    if (a.name) p.name = a.name;
    if (a.dept) p.dept = a.dept;
    p.items.push(a);
    var gc = gcByCode[a.code] || (a.axis === '成果' ? '直接の成果'
      : a.axis === '姿勢' ? '価値への取り組み' : '人材育成');
    if (!p.buckets[gc]) p.buckets[gc] = [];
    p.buckets[gc].push('・' + (a.title || '') + '：' + String(a.practice || '').slice(0, 120));
  });
  return per;
}

// ---------- フォルダ ----------
function tkGCSWorkFolder_() {
  var c = TK.GC_SLIDE;
  var root = c.ROOT_FOLDER_ID ? DriveApp.getFolderById(c.ROOT_FOLDER_ID) : DriveApp.getRootFolder();
  var nm = c.WORK_FOLDER_NAME || 'GC下書き';
  var it = root.getFoldersByName(nm);
  return it.hasNext() ? it.next() : root.createFolder(nm);
}

function tkGCSPersonFolder_(work, name, mid) {
  var nm = (name || mid.split('@')[0]);
  var it = work.getFoldersByName(nm);
  return it.hasNext() ? it.next() : work.createFolder(nm);
}

// ---------- 本人への案内メール（TK.GC_SLIDE.MAIL = true のとき） ----------
function tkGCSMail_(mid, p, label, file) {
  var body = [
    (p.name || '') + ' さん',
    '',
    '【' + label + '】今週のGCシートの下書きができました。',
    file.getUrl(),
    '',
    '今週の日報 ' + p.items.length + '件から作成しています。',
    'そのまま使える部分はそのまま、違うところは直してお使いください。',
    '',
    '・「外に生み出した成果」の3分類は、日報の事実をそのまま振り分けたものです。',
    '・文章の欄はAIの下書きです。ご本人の言葉に直してください。',
    '・「外的な変化／内的な変化」「ドラッカーのコンセプト」「フィロソフィ目標」は',
    '　空欄にしてあります。ここはご本人にしか書けないところです。',
    '',
    '（' + TK.PRAISE_CHARA + '）'
  ].join('\n');
  MailApp.sendEmail({ to: mid, subject: '【GCシート下書き】' + label + '　' + (p.name || ''), body: body });
}
