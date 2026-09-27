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
//    tkGCスライド_トリガー設定()  … 旧：毎週金曜17時30分（全員一斉）
//  ★2026-09-06 から（石川さん「グループでGCの曜日が違う」に対応）:
//    tkGC_トリガー設定()        … 毎日17:30 に tkGC日次トリガー。シート「GC設定」で
//                                  「明日がGC」の部署だけ、GC前日までの7日間で作る
//    tkGC下書き_今すぐ作る()     … メニューから部署を選んで即時作成（待たせない）
//    tkGCスライド_推進メンバーに共有() … 作業フォルダ全体を推進メンバーに編集権で
//    ?view=gc&key=…            … 推進メンバー向けの下書き一覧（tkGCSListHtml_）
//
//  ★2026-09-09 から
//   ・フォルダを「GC下書き ／ 事業所 ／ 氏名」の2階層にした（TK.GC_SLIDE.BY_OFFICE）。
//     すでに直下にある人フォルダは、実行時に事業所フォルダへ移す（1回で済む・冪等）。
//     事業所は メンバー表の「事業所」列（無ければ「部署」列）。シートを直せば分け方が変わる。
//   ・名簿にいる全員分を作る（TK.GC_SLIDE.ALL_MEMBERS）。日報が無い方は文章欄が空のまま、
//     「日報0件」と入ったGCシートが置かれる。GCは全員参加なので白紙でも要る。
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
  var roster0 = tkGCSRoster_();
  var made = 0, shared = 0, lines = ['=== 個人フォルダの準備 ==='];

  roster0.list.forEach(function (m) {
    var mid = m.mid, name = m.name || (mid ? mid.split('@')[0] : '');
    if (!mid) return;
    var folder = tkGCSPersonFolder_(work, name, mid, m.office);
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
    .onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(tkGCSRunHour_()).nearMinute(tkGCSRunMinute_()).inTimezone(TK.TZ).create();
  var msg = '毎週金曜 ' + tkGCSRunHour_() + '時に、個人フォルダへ下書き入りGCシートを作る設定にしました。\n'
          + '★事業所ごとに曜日を変えるなら ⑭ を使ってください（こちらは金曜固定の古い方式です）。';
  tkLog_('GCスライド', msg); Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}

// ============================================================
//  混み合ったときの再挑戦
// ============================================================
//  Drive/Slides は連続で叩くと「不明なエラーが発生しました」を返すことがある。
//  中身の問題ではなく一時的なものなので、少し待ってやり直せば通る。
//  待ち時間は 2秒 → 5秒 → 12秒 と延ばす。
var TK_GCS_WAIT = [2000, 5000, 12000];

// 下書きを作る時刻（GCの前日）。config で変えられる
function tkGCSRunHour_() {
  var h = TK.GC_SLIDE && TK.GC_SLIDE.RUN_HOUR;
  return (h === 0 || h) ? h : 11;
}
function tkGCSRunMinute_() {
  var m = TK.GC_SLIDE && TK.GC_SLIDE.RUN_MINUTE;
  return (m === 0 || m) ? m : 0;
}

function tkGCSRetry_(fn) {
  var last = null;
  for (var i = 0; i <= TK_GCS_WAIT.length; i++) {
    try { return fn(); }
    catch (e) {
      last = e;
      if (i < TK_GCS_WAIT.length) Utilities.sleep(TK_GCS_WAIT[i]);
    }
  }
  throw last;
}

//  opt.range / opt.group / opt.silent は 70_gc.gs の tkGCBuild_ と同じ意味
function tkGCSBuild_(offsetWeeks, opt) {
  opt = opt || {};
  var c = TK.GC_SLIDE;
  var tplId = tkGCSTemplateId_();          // 新フレーム（スクリプトプロパティ）が優先
  if (!tplId) {
    throw new Error('下書き用テンプレが未設定です。tkGCスライド_新フレームを作る() を実行してください。');
  }
  if (tplId === c.MASTER_ID) {
    throw new Error('テンプレが顧客原本と同じです。原本を書き換える事故になるため中止しました。');
  }

  var range = opt.range || tkGCWeekRange_(offsetWeeks);   // 70_gc.gs と同じ週の切り方
  var label = tkGCWeekLabel_(range);
  var per = tkGCSCollect_(range, opt.group);               // メンバー別の材料
  var roster = tkGCSRoster_();
  var work = tkGCSWorkFolder_();
  var lines = ['=== 下書き入りGCシートの作成（' + label + (opt.group ? '／' + opt.group : '') + '） ==='];
  var madeForChat = [];                                    // Chatへのお知らせ用（94_chat.gs）
  if (!opt.noShare) tkGCSShareManagers_(work, lines);      // 個別再生成では既存の共有設定を変えない

  // 作る相手＝「材料がある人」＋「名簿にいる全員」（ALL_MEMBERS のとき）。
  //   GCは全員参加なので、日報が無い方にも白紙のシートが要る。
  var ids = Object.keys(per);
  if (c.ALL_MEMBERS) {
    var ps = opt.group ? tkSetPersonSources_() : null;
    roster.list.forEach(function (m) {
      if (per[m.key] || m.stop) return;
      if (opt.group && !tkSetGroupMatch_(opt.group, { mid: m.mid, name: m.name, dept: m.dept, office: m.office }, ps)) return;
      per[m.key] = { name: m.name, dept: m.dept, mid: m.mid, items: [], blank: true,
        buckets: { '直接の成果': [], '価値への取り組み': [], '人材育成': [] } };
      ids.push(m.key);
    });
  }
  if (opt.onlyKeys) ids = ids.filter(function (k) { return !!opt.onlyKeys[k]; });
  if (!ids.length) {
    var none = '【' + label + (opt.group ? '／' + opt.group : '') + '】対象になる方がいません（名簿が空か、絞り込みに誰も当たりません）。';
    lines.push(none);
    tkLog_('GCスライド', none); Logger.log(lines.join('\n'));
    if (!opt.silent) { try { SpreadsheetApp.getUi().alert(lines.join('\n')); } catch (e) {} }
    return { members: 0 };
  }

  var done = 0, skipped = 0, failed = 0, blanks = 0, left = 0;
  // 複数の事業所をまとめて回すときは、6分の上限を全体で分け合う
  var deadline = opt.deadline || (Date.now() + TK_GCS_RUN_MS);

  ids.slice(0, TK_GC.MAX_MEMBERS).forEach(function (mid) {
    // 1人あたり数秒かかる。6分の上限に当たる前に切り上げ、続きは5分後の自動実行に回す
    if (Date.now() > deadline) { left++; return; }
    var p = per[mid];
    var rm = roster.byKey[mid];
    var office = (rm && rm.office) || p.dept;
    var folder = tkGCSPersonFolder_(work, p.name, mid, office);
    var fname = 'GCシート ' + label.replace(/\//g, '-') + '　' + (p.name || mid) +
      (opt.versionLabel ? '（' + opt.versionLabel + '）' : '');

    // 同じ週のものが既にあれば作らない（本人が手を入れたものを潰さない）
    if (folder.getFilesByName(fname).hasNext()) {
      skipped++; lines.push('　スキップ（既存）: ' + fname);
      return;
    }

    // ⚠ Drive/Slides は混み合うと「不明なエラー」を返す（2026-09-01 実際に発生）。
    //    1人分の失敗で全員分を落とさないよう、ここから1人ずつ包む。
    //    金曜の自動実行では作り直せないので、取りこぼしを残さないことが大事。
    var copy, pres;
    try {
      copy = tkGCSRetry_(function () {
        return DriveApp.getFileById(tplId).makeCopy(fname, folder);
      });
      pres = SlidesApp.openById(copy.getId());
    } catch (e) {
      failed++;
      lines.push('　🔴 失敗: ' + (p.name || mid) + '（' + (e && e.message) + '）');
      return;
    }

    // 日報が無い方はAIを呼ばない（材料が無いので推測になる）。文章欄は空のまま渡す
    var d = p.blank ? { theme: '', practice: '', insight: '', philosophy: '', next: '', takeaway: '' }
                    : tkGCDraftText_(p, range);   // 70_gc.gs のAI下書きをそのまま使う
    // 箱に収まる分だけ。3件を超えたら件数で示す
    var b = function (k) {
      var v = p.buckets[k] || [];
      if (!v.length) return p.blank ? '' : '（今週は該当なし）';
      var head = v.slice(0, 3).map(function (x) { return '・' + x; }).join('\n');
      return v.length > 3 ? head + '\n・ほか' + (v.length - 3) + '件' : head;
    };

    var map = {
      '{{NAME}}':          p.name || mid,
      '{{DEPT}}':          p.dept || '',
      '{{WEEK}}':          label,
      '{{COUNT}}':         String(p.items.length) + '件',   // 0件なら「日報0件」と入る
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
    try { pres.saveAndClose(); }
    catch (e) {
      failed++;
      lines.push('　🔴 保存できず: ' + (p.name || mid) + '（' + (e && e.message) + '）');
      return;
    }

    // メールが分かる人だけ本人に共有できる。氏名だけの人は推進メンバー経由で渡す
    if (!opt.noShare && c.SHARE_WITH_MEMBER && tkKeyIsMail_(mid)) { try { copy.addEditor(mid); } catch (e) {} }
    if (!opt.noMail && c.MAIL && tkKeyIsMail_(mid)) { try { tkGCSMail_(mid, p, label, copy); } catch (e) {} }
    if (!tkKeyIsMail_(mid)) lines.push('　（メール未登録のため本人共有なし: ' + (p.name || mid) + '）');

    done++;
    if (p.blank) blanks++;
    madeForChat.push({ office: office, name: p.name || mid, id: copy.getId(),
      url: copy.getUrl(), blank: !!p.blank });
    lines.push('　作成: ' + fname + (p.blank ? '（日報なし・白紙）' : '') + '　' + copy.getUrl());
  });

  // 名簿に無い方（フォームの氏名が名簿と違う等）は、ここで気づけるようにする
  var offRoster = ids.filter(function (k) { return !roster.byKey[k]; })
    .map(function (k) { return per[k].name || k; });
  if (offRoster.length) lines.push('　⚠ 名簿に無い氏名 ' + offRoster.length + '名: ' + offRoster.join('、'));

  if (left && !opt.onlyKeys) {
    tkGCSScheduleResume_(range, opt.group);
    lines.push('　⏳ 時間の都合でここまで。残り ' + left + '名は5分後に自動で続きます。');
  }

  lines.push('');
  lines.push('作成 ' + done + '名（うち日報なしの白紙 ' + blanks + '名）／既存 ' + skipped + '名' +
    (failed ? '／🔴失敗 ' + failed + '名' : '') + (left ? '／残り ' + left + '名（自動で続行）' : ''));
  if (failed) lines.push('※ 失敗した方は、もう一度この関数を実行すれば作られます（できている分はスキップされます）。');
  lines.push('※ 3分類は事実の振り分け、文章欄はAIの下書きです。ご本人が確認・微修正してお使いください。');
  lines.push('※ 外的/内的な変化・ドラッカーのコンセプト・フィロソフィ目標は空欄のままです（ご本人の記入欄）。');
  // ⚠ 途中で切れた回は投げない。残りが片付いた回にまとめて1本だけ出す。
  if (!opt.noChat && !left && madeForChat.length) {
    try { tkChatNotifyGC_(label, madeForChat, work.getUrl()); } catch (e) {
      lines.push('　（Chatへのお知らせは送れませんでした: ' + (e && e.message) + '）');
    }
  }

  var msg = lines.join('\n');
  tkLog_('GCスライド', '作成 ' + done + '名（白紙' + blanks + '）・既存 ' + skipped + '名' +
    (failed ? '・失敗 ' + failed + '名' : '') +
    '（' + label + (opt.group ? '／' + opt.group : '') + '）');
  Logger.log(msg);
  if (!opt.silent) { try { SpreadsheetApp.getUi().alert(msg.slice(0, 1800)); } catch (e) {} }
  return { members: done, skipped: skipped, failed: failed, blanks: blanks, left: left,
    week: label, files: madeForChat };
}

// ============================================================
//  途中で止まったら、5分後に自分で続きを作る
// ============================================================
//  ⚠ 36名分になったので1回の実行では収まらないことがある。
//    できている分はスキップされるので、同じ引数でもう一度走らせれば足りる。
//    一括分類（80_bulk）と同じ「レジューム式」の考え方。
var TK_GCS_RUN_MS = 260000;                  // ここまで来たら切り上げる（上限6分に対する余裕）
var TK_GCS_PROP = 'TK_GCS_RESUME';

function tkGCSScheduleResume_(range, group) {
  try {
    PropertiesService.getScriptProperties().setProperty(TK_GCS_PROP, JSON.stringify({
      from: range.from.getTime(), to: range.to.getTime(), group: group || ''
    }));
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (t.getHandlerFunction() === 'tkGCスライド_続きを作る') ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger('tkGCスライド_続きを作る').timeBased().after(5 * 60 * 1000).create();
  } catch (e) { tkLog_('GCスライド', '続行の予約に失敗: ' + (e && e.message)); }
}

function tkGCスライド_続きを作る() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(TK_GCS_PROP);
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'tkGCスライド_続きを作る') ScriptApp.deleteTrigger(t);
  });
  if (!raw) { tkLog_('GCスライド', '続きの予約がありません'); return { members: 0 }; }
  var s = JSON.parse(raw);
  var r = tkGCSBuild_(0, { range: { from: new Date(s.from), to: new Date(s.to) },
                           group: s.group || '', silent: true });
  if (!r || !r.left) props.deleteProperty(TK_GCS_PROP);   // 全部できたら予約を消す
  return r;
}

// ============================================================
//  推進メンバーへの共有（作業フォルダ「GC下書き」全体・編集権）
// ============================================================
//  親フォルダに付けるので、中の全員分のフォルダと下書きにそのまま効く。
//  既に付いている人は触らない（毎回の実行で差分だけ足す）。
function tkGCSShareManagers_(work, lines) {
  var want = tkSetManagers_();
  if (!want.length) return 0;
  var have = {};
  try { work.getEditors().forEach(function (u) { have[String(u.getEmail() || '').toLowerCase()] = 1; }); } catch (e) {}
  try { have[String(work.getOwner().getEmail() || '').toLowerCase()] = 1; } catch (e) {}
  var added = 0;
  want.forEach(function (m) {
    if (have[m]) return;
    try { work.addEditor(m); added++; if (lines) lines.push('　推進メンバーに共有: ' + m); }
    catch (e) { if (lines) lines.push('　⚠ 共有できず: ' + m + '（' + (e && e.message) + '）'); }
  });
  if (added) tkLog_('GCスライド', '推進メンバーに共有 ' + added + '名');
  return added;
}
function tkGCスライド_推進メンバーに共有() {
  var work = tkGCSWorkFolder_();
  var want = tkSetManagers_();
  var lines = ['=== 推進メンバーへの共有 ===', 'フォルダ: ' + work.getUrl(), ''];
  if (!want.length) {
    lines.push('シート「' + TK.SHEET_MANAGERS + '」にメールが入っていません。');
    lines.push('（無い場合は ⑬ 設定シートを作る → メールを入れる → もう一度これを実行）');
  } else {
    var n = tkGCSShareManagers_(work, lines);
    lines.push('');
    lines.push('対象 ' + want.length + '名／今回追加 ' + n + '名（残りは既に共有済み）');
  }
  var msg = lines.join('\n'); Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return msg;
}

// ============================================================
//  部署ごとの曜日で作る（毎日17:30）／今すぐ作る
// ============================================================
//  シート「GC設定」の各行について、**明日がGC曜日**なら、その部署の下書きを
//  「今日までの7日間」で作る。GC当日の朝には置かれている。
function tkGC日次トリガー() {
  var today = new Date();
  var tomorrow = (today.getDay() + 1) % 7;
  var groups = tkSetGcGroups_().filter(function (g) { return g.enabled && g.day === tomorrow; });
  if (!groups.length) {
    tkLog_('GC', '本日は対象なし（明日がGC曜日の部署がない）');
    return { groups: 0 };
  }
  var range = tkGCWeekRange_(0, { endDate: today });
  var out = [];
  groups.forEach(function (g) {
    try { tkGCBuild_(0, { range: range, group: g.group, silent: true }); }
    catch (e) { tkLog_('GC下書き', g.group + ': ' + (e && e.message)); }
    try { out.push(tkGCSBuild_(0, { range: range, group: g.group, silent: true })); }
    catch (e) { tkLog_('GCスライド', g.group + ': ' + (e && e.message)); }
  });
  return { groups: groups.length, results: out };
}

// 旧トリガー2本（金曜17:00 シート／17:30 スライド）を外して、日次に置き換える
function tkGC_トリガー設定() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (['tkGC週次トリガー', 'tkGCスライド週次トリガー', 'tkGC日次トリガー'].indexOf(fn) >= 0) ScriptApp.deleteTrigger(t);
  });
  var H = tkGCSRunHour_(), M = tkGCSRunMinute_();
  ScriptApp.newTrigger('tkGC日次トリガー').timeBased().everyDays(1).atHour(H).nearMinute(M).inTimezone(TK.TZ).create();
  var hhmm = H + ':' + (M < 10 ? '0' + M : M);
  var gs = tkSetGcGroups_();
  var lines = ['毎日 ' + hhmm + ' に、シート「' + TK.SHEET_GCSET + '」を見て「明日がGC」の事業所の下書きを作る設定にしました。',
    '（金曜17:00／17:30 の2本は外しました）', '', 'いまの設定:'];
  gs.forEach(function (g) {
    var prev = TK_SET.WEEKDAYS[(g.day + 6) % 7];
    lines.push('　' + g.group + '：GCは' + TK_SET.WEEKDAYS[g.day] + '曜 → ' + prev + '曜 ' + hhmm + ' に作成' +
      (g.enabled ? '' : '（無効）'));
  });
  lines.push('', '曜日を変えるときはシートを直すだけです。');
  lines.push('時刻を変えるときは config.gs の GC_SLIDE.RUN_HOUR です（松山が直します）。');
  var msg = lines.join('\n');
  tkLog_('GC', msg); Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}

// メニュー ⑫：今日までの7日間で今すぐ作る
//  対象は空＝全員。事業所やソースIDを「/」「、」区切りで複数書いてもよい。
function tkGC下書き_今すぐ作る() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('GC下書きを今すぐ作る',
    '★何も入れずに OK を押すと、名簿の全員分を作ります。\n' +
    '　（すでにある下書きは作り直しません）\n\n' +
    '一部だけ作りたいときは、事業所名を入れてください。\n' +
    '　例）薬局　　例）ネットワーク　　例）薬局 / ネットワーク\n\n' +
    '期間は今日までの7日間です。',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;

  var groups = String(r.getResponseText() || '').split(/[\/,、\s]+/)
    .map(function (s) { return s.trim(); }).filter(String);
  if (!groups.length) groups = [''];          // 空＝全員

  var range = tkGCWeekRange_(0, { endDate: new Date() });
  var deadline = Date.now() + TK_GCS_RUN_MS;
  var out = [], total = 0, left = 0;
  groups.forEach(function (g) {
    try { tkGCBuild_(0, { range: range, group: g, silent: true }); } catch (e) {}
    var r2 = tkGCSBuild_(0, { range: range, group: g, silent: true, deadline: deadline });
    out.push((g || '全員') + '：作成 ' + (r2.members || 0) + '名' +
      (r2.blanks ? '（白紙 ' + r2.blanks + '）' : '') +
      (r2.skipped ? '／既存 ' + r2.skipped : '') +
      (r2.left ? '／残り ' + r2.left + '（自動で続行）' : ''));
    total += (r2.members || 0);
    left += (r2.left || 0);
  });

  var msg = ['=== GC下書きを今すぐ作りました ===',
    '期間: ' + tkGCWeekLabel_(range), ''].concat(out);
  if (!total && !left) {
    msg.push('');
    msg.push('⚠ 1件も作られませんでした。入れた対象名が事業所と一致していない可能性があります。');
    msg.push('　もう一度 ⑫ を押して、**何も入れずに OK**（＝全員）でお試しください。');
  }
  var text = msg.join('\n');
  Logger.log(text);
  try { ui.alert(text.slice(0, 1800)); } catch (e) {}
  return { members: total, left: left };
}

// ============================================================
//  推進メンバー向け：下書きの一覧（?view=gc&key=…）
// ============================================================
//  作業フォルダの中を歩いて、直近 days 日に作られた「GCシート …」を表にする。
//  更新日時が作成から2分以上あとなら、ご本人が手を入れたとみなす。
function tkGCSListHtml_(key, opt) {
  opt = opt || {};
  var days = Math.max(1, Math.min(60, +opt.days || 10));
  var group = String(opt.group || '');
  var since = new Date(Date.now() - days * 86400000);
  var work = tkGCSWorkFolder_();

  var deptBy = {};
  try {
    var sh = tkSS_().getSheetByName(TK.SHEET_MEMBER);
    if (sh && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (r) {
        var n = String(r[1] || '').trim(); if (n) deptBy[n] = String(r[2] || '').trim();
      });
    }
  } catch (e) {}

  var rows = [];
  // 「事業所 → 人」と、旧レイアウトの「人」直置きの両方を歩く
  var scan = function (folder, office) {
    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      var nm = file.getName();
      if (nm.indexOf('GCシート ') !== 0) continue;
      var cr = file.getDateCreated();
      if (cr < since) continue;
      var person = folder.getName();
      var dept = office || deptBy[person] || '';
      if (group && dept !== group) continue;
      var up = file.getLastUpdated();
      var m = nm.match(/^GCシート (\S+)/);
      rows.push({ person: person, dept: dept, week: m ? m[1] : '', created: cr, updated: up,
                  edited: (up.getTime() - cr.getTime()) > 120000, url: file.getUrl() });
    }
  };
  var top = work.getFolders();
  while (top.hasNext()) {
    var f = top.next();
    scan(f, deptBy[f.getName()] || '');          // 旧レイアウト（人フォルダが直下）
    var sub = f.getFolders();                     // 新レイアウト（事業所フォルダの中）
    while (sub.hasNext()) scan(sub.next(), f.getName());
  }
  rows.sort(function (a, b) {
    if (a.week !== b.week) return a.week < b.week ? 1 : -1;
    if (a.dept !== b.dept) return a.dept < b.dept ? -1 : 1;
    return a.person < b.person ? -1 : 1;
  });

  var fmt = function (d) { return Utilities.formatDate(d, TK.TZ, 'M/d HH:mm'); };
  var css = 'body{font-family:"Noto Sans JP","Hiragino Sans",Meiryo,sans-serif;background:#FFFDF5;color:#1B2E1B;margin:0;padding:20px}' +
    '.wrap{max-width:1040px;margin:0 auto}h1{font-size:20px;margin:0 0 4px}.s{color:#4E6B4E;font-size:13px;margin-bottom:14px}' +
    'table{border-collapse:collapse;width:100%;font-size:13.5px;background:#fff}th,td{border:1px solid #CFE3CF;padding:7px 9px;text-align:left;vertical-align:top}' +
    'th{background:#E8F5E9;font-weight:600}tr:nth-child(even) td{background:#F7FBF7}' +
    '.wk td{background:#DCEDC8;font-weight:600}.ok{color:#1B5E20;font-weight:600}.no{color:#7A7A7A}' +
    'a{color:#1B5E20}.note{font-size:12.5px;color:#4E6B4E;margin-top:14px;line-height:1.6}';
  var h = ['<style>' + css + '</style><div class="wrap">',
    '<h1>GC下書きの一覧</h1>',
    '<div class="s">直近 ' + days + ' 日に作られた下書き　' + rows.length + '件' + (group ? '　対象: ' + tkEsc_(group) : '') +
      '　／ <a href="' + work.getUrl() + '" target="_blank">フォルダを開く</a></div>',
    '<table><tr><th>部署</th><th>氏名</th><th>下書き</th><th>作成</th><th>本人の編集</th></tr>'];
  var lastWeek = null;
  rows.forEach(function (r) {
    if (r.week !== lastWeek) { h.push('<tr class="wk"><td colspan="5">' + tkEsc_(r.week) + '</td></tr>'); lastWeek = r.week; }
    h.push('<tr><td>' + tkEsc_(r.dept) + '</td><td>' + tkEsc_(r.person) + '</td>' +
      '<td><a href="' + r.url + '" target="_blank">開く</a></td><td>' + fmt(r.created) + '</td>' +
      '<td>' + (r.edited ? '<span class="ok">編集あり（' + fmt(r.updated) + '）</span>' : '<span class="no">未編集</span>') + '</td></tr>');
  });
  if (!rows.length) h.push('<tr><td colspan="5">この期間の下書きはありません。</td></tr>');
  h.push('</table>',
    '<p class="note">下書きは「GC設定」シートの曜日にしたがって、GC前日の17:30に置かれます。' +
    '「本人の編集」は、作成から2分以上あとに更新があったかで見ています。<br>' +
    '&amp;days=30 で遡る日数、&amp;group=部署名 で絞り込めます。</p></div>');
  return HtmlService.createHtmlOutput(h.join('')).setTitle('GC下書きの一覧');
}

// ---------- 材料集め（70_gc.gs の集計と同じ考え方） ----------
function tkGCSCollect_(range, group) {
  var quests = tkLoadQuests_();
  var gcByCode = {};
  quests.forEach(function (q) { gcByCode[q.code] = q.gc; });
  var ps = group ? tkSetPersonSources_() : null;
  var rs = group ? tkGCSRoster_() : null;          // 事業所で絞るために名簿を引く

  var acts = tkReadActivity_().filter(function (a) {
    if (!a.ts) return false;
    var t = new Date(a.ts);
    if (t < range.from || t > range.to) return false;
    if (!group) return true;
    var rm = rs && rs.byKey[tkPersonKey_(a.mid, a.name)];
    return tkSetGroupMatch_(group, { mid: a.mid, name: a.name, dept: a.dept,
      office: rm && rm.office }, ps);
  });

  var per = {};
  acts.forEach(function (a) {
    var k = tkPersonKey_(a.mid, a.name);       // メールが無い人は氏名で拾う
    if (!k) return;
    if (!per[k]) {
      per[k] = { name: a.name, dept: a.dept, mid: a.mid || '', items: [],
        buckets: { '直接の成果': [], '価値への取り組み': [], '人材育成': [] } };
    }
    var p = per[k];
    if (a.name) p.name = a.name;
    if (a.dept) p.dept = a.dept;
    if (a.mid && !p.mid) p.mid = a.mid;
    p.items.push(a);
    var gc = gcByCode[a.code] || (a.axis === '成果' ? '直接の成果'
      : a.axis === '姿勢' ? '価値への取り組み' : '人材育成');
    if (!p.buckets[gc]) p.buckets[gc] = [];
    // ⚠ ここに日報の生文を入れると箱からあふれる（2026-08-31 実物で確認）。
    //    220x106pt の箱に3分類ぶん入るので、小見出しだけにして重複も除く。
    //    詳しい中身は日報そのものを見ればよい。
    var ti = String(a.title || '').trim();
    if (ti && p.buckets[gc].indexOf(ti) < 0) p.buckets[gc].push(ti);
  });
  return per;
}

// ---------- 名簿（事業所つき） ----------
//  事業所の決め方（上が優先）:
//    ① メンバー表の「事業所」列（人が決めたものが最優先）
//    ② その人が実際に日報を書いているフォーム ＝ トライアルのチーム
//       （薬局／ネットワーク／花のある家（中吉田））
//    ③ どのフォームにも書いていない → 「その他」
//  ★②にしている理由：部署（薬局・ネットワークなど）ではチームが分かれない。
//    GCはチーム単位で行うので、実際に書いている場所で分けるのが実態に合う。
//  ★列の増減に強くするため、見出し名で探す（列番号を決め打ちしない）
var TK_GCS_PSRC = null;      // 日報統合の読み直しを1実行で1回にする
function tkGCSRoster_() {
  var out = { list: [], byKey: {} };
  var sh = tkSS_().getSheetByName(TK.SHEET_MEMBER);
  if (!sh || sh.getLastRow() < 2) return out;
  var other = (TK.GC_SLIDE && TK.GC_SLIDE.OFFICE_OTHER) || 'その他';
  if (!TK_GCS_PSRC) TK_GCS_PSRC = tkSetPersonSources_();
  var bySrc = tkSetOfficeBySource_();

  var w = Math.max(3, sh.getLastColumn());
  var head = sh.getRange(1, 1, 1, w).getValues()[0].map(function (v) { return String(v || '').trim(); });
  var iOffice = head.indexOf('事業所');
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, w).getValues();
  vals.forEach(function (r) {
    var mid = tkMemberId_(r[0]), name = String(r[1] || '').trim(), dept = String(r[2] || '').trim();
    var key = tkPersonKey_(mid, name);
    if (!key) return;
    var src = (mid && TK_GCS_PSRC.byMid[mid]) || (name && TK_GCS_PSRC.byName[name]) || '';
    var office = (iOffice >= 0 ? String(r[iOffice] || '').trim() : '')
              || (src && bySrc[src]) || other;
    var m = { key: key, mid: mid, name: name, dept: dept, office: office, source: src,
              stop: String(r[6] || '').trim() === '停止' };
    out.list.push(m);
    out.byKey[key] = m;
  });
  return out;
}

// ---------- フォルダ ----------
function tkGCSWorkFolder_() {
  var c = TK.GC_SLIDE;
  var root = c.ROOT_FOLDER_ID ? DriveApp.getFolderById(c.ROOT_FOLDER_ID) : DriveApp.getRootFolder();
  var nm = c.WORK_FOLDER_NAME || 'GC下書き';
  var it = root.getFoldersByName(nm);
  return it.hasNext() ? it.next() : root.createFolder(nm);
}

//  事業所フォルダ（BY_OFFICE が false なら作業フォルダそのまま）
function tkGCSOfficeFolder_(work, office) {
  if (!(TK.GC_SLIDE && TK.GC_SLIDE.BY_OFFICE)) return work;
  var nm = String(office || '').trim() || '部署未設定';
  var it = work.getFoldersByName(nm);
  return it.hasNext() ? it.next() : work.createFolder(nm);
}

//  人フォルダ。事業所フォルダの中に置く。
//  ⚠ 以前は作業フォルダの直下に作っていた。直下に見つかったら**移す**（中身・共有はそのまま）。
function tkGCSPersonFolder_(work, name, mid, office) {
  var nm = name || (tkKeyIsMail_(mid) ? String(mid).split('@')[0] : String(mid).replace(/^名前:/, ''));
  var parent = tkGCSOfficeFolder_(work, office);

  var here = parent.getFoldersByName(nm);
  if (here.hasNext()) return here.next();

  // 旧レイアウト（作業フォルダ直下）にあれば移設する
  if (parent.getId() !== work.getId()) {
    var old = work.getFoldersByName(nm);
    if (old.hasNext()) {
      var f = old.next();
      try { f.moveTo(parent); tkLog_('GCスライド', 'フォルダ移設 ' + nm + ' → ' + parent.getName()); return f; }
      catch (e) { tkLog_('GCスライド', 'フォルダ移設できず ' + nm + '（' + (e && e.message) + '）'); return f; }
    }
  }
  return parent.createFolder(nm);
}

// ============================================================
//  フォルダを「事業所 → 人」に並べ替える（作り直しはしない・何度でも実行可）
// ============================================================
function tkGCスライド_事業所フォルダに並べ替え() {
  var work = tkGCSWorkFolder_();
  var roster = tkGCSRoster_();
  var lines = ['=== 事業所ごとのフォルダに並べ替え ===', 'フォルダ: ' + work.getUrl(), ''];
  var moved = 0, made = 0, kept = 0;

  var byOffice = {};
  roster.list.forEach(function (m) { (byOffice[m.office] = byOffice[m.office] || []).push(m); });

  Object.keys(byOffice).sort().forEach(function (office) {
    var parent = tkGCSOfficeFolder_(work, office);
    lines.push('▼ ' + office + '（' + byOffice[office].length + '名）');
    byOffice[office].forEach(function (m) {
      var nm = m.name || m.mid;
      if (parent.getFoldersByName(nm).hasNext()) { kept++; return; }
      var old = work.getFoldersByName(nm);
      if (old.hasNext()) {
        var f = old.next();
        try { f.moveTo(parent); moved++; lines.push('　移した: ' + nm); }
        catch (e) { lines.push('　⚠ 移せず: ' + nm + '（' + (e && e.message) + '）'); }
      } else {
        var nf = parent.createFolder(nm);
        if (TK.GC_SLIDE.SHARE_WITH_MEMBER && tkKeyIsMail_(m.mid)) {
          try { nf.addEditor(m.mid); } catch (e) {}
        }
        made++; lines.push('　作った: ' + nm);
      }
    });
  });

  tkGCSShareManagers_(work, lines);
  lines.push('');
  lines.push('移した ' + moved + '名／新しく作った ' + made + '名／もう入っていた ' + kept + '名');
  lines.push('※ グループは「実際に日報を書いているフォーム」＝トライアルのチームです。');
  lines.push('　 どのフォームにも書いていない方は「' + ((TK.GC_SLIDE && TK.GC_SLIDE.OFFICE_OTHER) || 'その他') + '」に入ります。');
  lines.push('　 メンバー表に「事業所」列を作って書けば、その値が最優先になります。');
  var msg = lines.join('\n');
  tkLog_('GCスライド', '事業所フォルダ 移設' + moved + '・新規' + made);
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg.slice(0, 1800)); } catch (e) {}
  return msg;
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

// ============================================================
//  ③の代わり：目印を座標から自動で貼る
// ============================================================
//  週次GCシートは「見出しの図形」と「回答欄の図形」が別々にある。
//  IDだけでは対応が取れないが、見出しの真下（または右）にある空欄は
//  その見出しの回答欄なので、位置から機械的に対応づけられる。
//
//  ⚠ 触るのは TEMPLATE_ID（原本のコピー）だけ。MASTER_ID には一切書かない。
//
//  使い方:
//    tkGCスライド_目印を下見()    … どの空欄に何を貼るかを一覧で出す（書き込まない）
//    tkGCスライド_目印を貼る()    … 下見の内容どおりに実際に貼る
//
//  ★必ず「下見」で確認してから「貼る」を実行する。

// 見出しから回答欄までの許容距離（pt）。これを超える箱は拾わない。
//   飾りの図形を掴んでしまう事故を防ぐ。実測では見出しの直下 20〜60pt に回答欄がある。
var TK_GCS_MAX_DIST = 130;

// 回答欄の最小の大きさ（pt）。これより小さい図形は回答欄とみなさない。
//   実測: 本物は 100x32 〜 220x151。チェックボックスや飾りは 32x19 で混じってくる。
var TK_GCS_MIN_W = 60, TK_GCS_MIN_H = 25;

// 見出しの文字 → 貼る目印。空文字は「本人が書く欄なので貼らない」
//  ⚠ 並び順が結果を変える。先に処理した見出しが箱を取る。
//    成果は1見出しに3欄なので、単独の見出しより**先**に置く。
//    （後ろに置くと、隣のフィロソフィが3欄の1つ目を取ってしまう）
var TK_GCS_LABELS = [
  // ── 複数欄のものを先に
  //  ⚠ 週次シートの「外に生み出した成果」は**回答欄が1つ**。
  //     見出しの文字に「直接の成果／価値への取り組み／人材育成」と3つ書いてあるだけで、
  //     欄は分かれていない（3欄に分かれているのは3ヶ月シートのほう）。
  //     実測: 見出し 428,22 (220x109) / 空欄 428,28 (220x106) ＝ 1対1。
  //     なので1つの欄に、小見出しを付けて3つまとめて入れる。
  { label: '外に生み出した成果は何か直接の成果価値への取り組み人材育成',
    text: '【直接の成果】\n{{RESULT_DIRECT}}\n\n【価値への取り組み】\n{{RESULT_VALUE}}\n\n【人材育成】\n{{RESULT_PEOPLE}}' },
  // ── 単独の欄
  { label: '今週の中心課題・取り組むテーマは何か',       token: '{{THEME}}' },
  { label: '何をどのように実践したか',                 token: '{{PRACTICE}}' },
  { label: '鍵となる発見・気づきは何か予期せぬ成功・失敗は？', token: '{{INSIGHT}}' },
  { label: 'すぐに取り組むべきことは何か',             token: '{{NEXT}}' },
  { label: '用いたフィロソフィは',                    token: '{{PHILOSOPHY}}' },
  { label: '今週のＫeyTakeawayは何か',                token: '{{TAKEAWAY}}' },
  // ── 本人が書く欄。貼らない（他の見出しに取られないよう確保だけする）
  { label: '目指していること・探求していること・大切にしていることなど', token: '' },
  { label: '効果的な計画として、どのようなことが考えられるか', token: '' },
  { label: '外的な変化は何か',                        token: '' },
  { label: '内的な変化は何か',                        token: '' },
  { label: '実践に用いたドラッカーのコンセプトは何か',   token: '' },
  { label: '私のフィロソフィ目標：',                   token: '' },
  { label: 'フィードバックコメント',                   token: '' }
];

function tkGCSFlat_(s) {
  return String(s || '').replace(/[\s　]/g, '');
}

// 週次スライドを探す（「今週の振り返り」を含む最初のスライド）
function tkGCSWeeklySlide_(pres) {
  var hit = null;
  pres.getSlides().forEach(function (sl) {
    if (hit) return;
    sl.getShapes().forEach(function (sh) {
      if (hit) return;
      var t = '';
      try { t = tkGCSFlat_(sh.getText().asString()); } catch (e) {}
      if (t.indexOf('今週の振り返り') === 0) hit = sl;
    });
  });
  return hit;
}

function tkGCSBox_(sh) {
  try {
    return { l: sh.getLeft(), t: sh.getTop(), w: sh.getWidth(), h: sh.getHeight() };
  } catch (e) { return null; }
}

// 見出しに対応する空欄を選ぶ
//   ① 横に重なっていて、見出しより下にあるもののうち近い順
//   ② 無ければ、縦に重なっていて右にあるもののうち近い順
//   いずれも TK_GCS_MAX_DIST を超えるものは候補にしない
function tkGCSPick_(labelBox, blanks, want, stack) {
  var cand = [];
  blanks.forEach(function (b) {
    if (b.used) return;
    var ov = Math.min(labelBox.l + labelBox.w, b.box.l + b.box.w) - Math.max(labelBox.l, b.box.l);
    if (ov > Math.min(labelBox.w, b.box.w) * 0.35 && b.box.t >= labelBox.t - 2) {
      var d = b.box.t - labelBox.t;
      if (d <= TK_GCS_MAX_DIST) cand.push({ b: b, d: d, dir: '下' });
    }
  });
  if (!cand.length) {
    blanks.forEach(function (b) {
      if (b.used) return;
      var ov = Math.min(labelBox.t + labelBox.h, b.box.t + b.box.h) - Math.max(labelBox.t, b.box.t);
      if (ov > 0 && b.box.l >= labelBox.l) {
        var d = b.box.l - labelBox.l;
        if (d <= TK_GCS_MAX_DIST) cand.push({ b: b, d: d, dir: '右' });
      }
    });
  }
  cand.sort(function (x, y) { return x.d - y.d; });
  var take = cand.slice(0, want || 1);
  // 複数欄のときの並び。縦積み(stack)なら上→下、そうでなければ左→右
  if (stack) take.sort(function (x, y) { return x.b.box.t - y.b.box.t; });
  else       take.sort(function (x, y) { return x.b.box.l - y.b.box.l; });
  return take;
}

function tkGCSAutoToken_(apply) {
  var c = TK.GC_SLIDE;
  if (!c || !tkGCSTemplateId_()) {
    throw new Error('下書き用テンプレが未設定です。' +
      'まず tkGCスライド_下書きテンプレを用意() を実行し、そのIDを config に入れてください。');
  }
  if (tkGCSTemplateId_() === c.MASTER_ID) {
    throw new Error('テンプレが MASTER_ID と同じです。顧客原本には書き込めません。');
  }

  var pres = SlidesApp.openById(tkGCSTemplateId_());
  var sl = tkGCSWeeklySlide_(pres);
  if (!sl) throw new Error('週次スライド（「今週の振り返り」を含むもの）が見つかりません');

  var blanks = [], labels = {}, labelList = [], tiny = [];
  sl.getShapes().forEach(function (sh) {
    var t = '';
    try { t = tkGCSFlat_(sh.getText().asString()); } catch (e) {}
    var box = tkGCSBox_(sh);
    if (!box) return;
    if (!t) {
      if (box.w < TK_GCS_MIN_W || box.h < TK_GCS_MIN_H) { tiny.push({ sh: sh, box: box }); return; }
      blanks.push({ sh: sh, box: box, used: false });
    }
    else { labels[t] = { sh: sh, box: box }; labelList.push({ t: t, box: box }); }
  });

  var lines = ['=== 目印の貼り付け ' + (apply ? '【実行】' : '【下見・書き込みません】') + ' ===',
    'テンプレ: ' + pres.getName(),
    'スライド: ' + sl.getObjectId() + '（週次シート）',
    '空欄 ' + blanks.length + '個 / 見出し ' + labelList.length + '個'
      + (tiny.length ? '（小さすぎて除外 ' + tiny.length + '個）' : ''), ''];

  var pasted = 0, missLabel = 0, missBox = 0;

  TK_GCS_LABELS.forEach(function (d) {
    var key = tkGCSFlat_(d.label);
    var lb = labels[key];
    if (!lb) {
      Object.keys(labels).forEach(function (k) {
        if (!lb && (k.indexOf(key) === 0 || key.indexOf(k) === 0) && k.length > 6) lb = labels[k];
      });
    }
    if (!lb) {
      lines.push('🔴 見出しが見つからない: ' + d.label);
      missLabel++;
      return;
    }
    // text 指定：1つの欄に複数の目印をまとめて入れる
    if (d.text) {
      var one = tkGCSPick_(lb.box, blanks, 1);
      if (!one.length) {
        lines.push('🔴 回答欄が見つからない: ' + d.label +
          ' 見出し位置 ' + Math.round(lb.box.l) + ',' + Math.round(lb.box.t));
        missBox++;
        return;
      }
      one[0].b.used = true;
      lines.push('　【1欄にまとめて】' + d.label +
        '（' + one[0].dir + Math.round(one[0].d) + 'pt ' + one[0].b.sh.getObjectId() +
        ' 位置 ' + Math.round(one[0].b.box.l) + ',' + Math.round(one[0].b.box.t) + '）');
      lines.push('　　' + d.text.replace(/\n/g, ' ／ '));
      if (apply) { one[0].b.sh.getText().setText(d.text); tkGCSFit_(one[0].b.sh); pasted += 3; }
      return;
    }

    var toks = d.tokens || [d.token];
    if (toks.length === 1 && toks[0] === '') {
      var keep = tkGCSPick_(lb.box, blanks, 1);
      if (keep.length) keep[0].b.used = true;
      lines.push('　（貼らない・本人記入）' + d.label);
      return;
    }
    var picks = tkGCSPick_(lb.box, blanks, toks.length, d.stack);
    if (picks.length < toks.length) {
      lines.push('🔴 回答欄が足りない: ' + d.label +
        '（必要' + toks.length + '・見つかった' + picks.length + '）' +
        ' 見出し位置 ' + Math.round(lb.box.l) + ',' + Math.round(lb.box.t));
      missBox++;
      return;
    }
    picks.forEach(function (p, i) {
      p.b.used = true;
      lines.push('　' + toks[i] + '  ←  ' + d.label +
        '（' + p.dir + Math.round(p.d) + 'pt ' + p.b.sh.getObjectId() +
        ' 位置 ' + Math.round(p.b.box.l) + ',' + Math.round(p.b.box.t) + '）');
      if (apply) { p.b.sh.getText().setText(toks[i]); tkGCSFit_(p.b.sh); pasted++; }
    });
  });

  var HEAD = '{{NAME}}さん（{{DEPT}}）　{{WEEK}}　日報{{COUNT}}';
  var already = false;
  sl.getShapes().forEach(function (sh) {
    var t = '';
    try { t = sh.getText().asString(); } catch (e) {}
    if (t.indexOf('{{NAME}}') >= 0) already = true;
  });
  if (already) {
    lines.push('', '見出し行はすでにあります（作り直しません）');
  } else if (apply) {
    var tb = sl.insertTextBox(HEAD, 12, 6, 420, 22);
    tkGCSFit_(tb);
    try { tb.getText().getTextStyle().setFontSize(11).setBold(true); } catch (e) {}
    lines.push('', '見出し行を左上に作りました: ' + HEAD);
    pasted += 4;
  } else {
    lines.push('', '見出し行を左上に作ります（いまは作っていません）: ' + HEAD);
  }

  // 下見では全体の座標も出す。対応づけを人が検算できるようにするため
  if (!apply) {
    lines.push('', '── 見出しの位置 ──');
    labelList.sort(function (a, b) { return a.box.t - b.box.t || a.box.l - b.box.l; });
    labelList.forEach(function (x) {
      lines.push('  ' + Math.round(x.box.l) + ',' + Math.round(x.box.t) +
        ' (' + Math.round(x.box.w) + 'x' + Math.round(x.box.h) + ') ' + x.t.slice(0, 34));
    });
    lines.push('', '── 空欄の位置（★は未割り当て）──');
    var bs = blanks.slice().sort(function (a, b) { return a.box.t - b.box.t || a.box.l - b.box.l; });
    bs.forEach(function (b) {
      lines.push('  ' + (b.used ? '  ' : '★') + ' ' +
        Math.round(b.box.l) + ',' + Math.round(b.box.t) +
        ' (' + Math.round(b.box.w) + 'x' + Math.round(b.box.h) + ') ' + b.sh.getObjectId());
    });
  }

  var rest = blanks.filter(function (b) { return !b.used; }).length;
  lines.push('', '割り当てなかった空欄: ' + rest + '個');
  if (apply) {
    lines.push('貼った目印: ' + pasted + '個');
    lines.push('', '次にやること: tkGCスライド_個人フォルダ準備() → tkGCスライド_先週分を作る()');
  } else {
    lines.push('', '★この内容でよければ tkGCスライド_目印を貼る() を実行してください');
  }
  if (missLabel || missBox) {
    lines.push('', '⚠ 見出し ' + missLabel + '件・欄不足 ' + missBox + '件。上の🔴を確認してください');
  }

  var msg = lines.join('\n');
  tkLog_('GCスライド', (apply ? '目印を貼った ' : '目印の下見 ') + pasted + '個');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg.slice(0, 1400)); } catch (e) {}
  return msg;
}

// 下見（書き込まない）
function tkGCスライド_目印を下見() { return tkGCSAutoToken_(false); }

// 実際に貼る
function tkGCスライド_目印を貼る() { return tkGCSAutoToken_(true); }

// ============================================================
//  文字が箱からあふれないように整える
// ============================================================
//  ⚠ 2026-08-31 の初回生成で、文字が箱を突き抜けて重なり読めなくなった。
//     原因は2つ。
//      ① 成果の欄に日報の生文を120字×件数ぶん入れていた（下の tkGCSCollect_ で修正）
//      ② スライドの文字が既定で縮まない設定だった（この関数で修正）
//
//  テンプレの目印（{{...}}）が入っている箱に、
//    ・小さめの文字
//    ・はみ出したら自動で縮む設定
//  を入れておく。差し替え後の文字にもこの設定が引き継がれる。

var TK_GCS_FONT = 8;      // 目印の箱の文字の大きさ（pt）

// 1つの箱に、小さめの文字とはみ出し時の自動縮小を入れる
function tkGCSFit_(sh) {
  try { sh.getText().getTextStyle().setFontSize(TK_GCS_FONT); } catch (e) {}
  try { sh.getAutofit().setAutofitType(SlidesApp.AutofitType.SHRINK_ON_OVERFLOW); } catch (e) {}
}

function tkGCスライド_文字を整える() {
  var c = TK.GC_SLIDE;
  if (!c || !tkGCSTemplateId_()) throw new Error('下書き用テンプレが未設定です');
  if (tkGCSTemplateId_() === c.MASTER_ID) throw new Error('テンプレが MASTER_ID と同じです');

  var pres = SlidesApp.openById(tkGCSTemplateId_());
  var n = 0, lines = ['=== 目印の箱の文字を整える ==='];

  pres.getSlides().forEach(function (sl) {
    sl.getShapes().forEach(function (sh) {
      var t = '';
      try { t = String(sh.getText().asString() || ''); } catch (e) { return; }
      if (t.indexOf('{{') < 0) return;
      try { sh.getText().getTextStyle().setFontSize(TK_GCS_FONT); } catch (e) {}
      try { sh.getAutofit().setAutofitType(SlidesApp.AutofitType.SHRINK_ON_OVERFLOW); } catch (e) {}
      n++;
      lines.push('　' + sh.getObjectId() + '  ' + t.replace(/\s+/g, ' ').slice(0, 40));
    });
  });
  pres.saveAndClose();

  lines.push('');
  lines.push('整えた箱: ' + n + '個（文字' + TK_GCS_FONT + 'pt・はみ出したら自動で縮む）');
  var msg = lines.join('\n');
  tkLog_('GCスライド', '文字を整えた ' + n + '個');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg.slice(0, 1400)); } catch (e) {}
  return msg;
}

// ============================================================
//  作り直しのために、ある週の下書きを消す
// ============================================================
//  ⚠ 消すのは個人フォルダの中の「GCシート <週> <氏名>」だけ。
//     ご本人が手を入れたものも消えるので、作り直す前だけに使う。
//
//    tkGCスライド_先週分を消す()  … 直前の週のぶんを消す
function tkGCスライド_先週分を消す() { return tkGCSDelete_(-1); }
function tkGCスライド_今週分を消す() { return tkGCSDelete_(0); }

// ★「今すぐ作る」で作った分を消す。
//   ⚠ 上の「今週分」は月〜日なので、今すぐ作った（今日まで7日間の）分とは
//     期間の名前が違い、1件も消えない。2026-09-17 に踏んだので専用を用意した。
function tkGCスライド_今日作った分を消す() {
  var ui = null; try { ui = SpreadsheetApp.getUi(); } catch (e) {}
  var group = '';
  if (ui) {
    var msg0 = '★何も入れずに OK を押すと、今日の期間の下書きを全員分ゴミ箱に入れます。\n'
             + '（ゴミ箱なので元に戻せます）\n\n'
             + '一部だけ消すときは、事業所名を入れてください。\n'
             + '　例）ネットワーク　　例）薬局\n\n'
             + '期間は今日までの7日間です。';
    var r = ui.prompt('今日作った下書きを消す', msg0, ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return { deleted: 0 };
    group = String(r.getResponseText() || '').trim();
  }
  return tkGCSDelete_(0, { range: tkGCWeekRange_(0, { endDate: new Date() }), group: group });
}

function tkGCSDelete_(offsetWeeks, opt) {
  opt = opt || {};
  var range = opt.range || tkGCWeekRange_(offsetWeeks);
  var label = tkGCWeekLabel_(range);
  var work = tkGCSWorkFolder_();
  var pat = 'GCシート ' + label.replace(/\//g, '-');
  var group = String(opt.group || '').trim();
  var n = 0, lines = ['=== 下書きを消す（' + label + (group ? '／' + group : '') + '） ==='];

  var kill = function (folder) {
    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      if (file.getName().indexOf(pat) === 0) {
        lines.push('　消す: ' + file.getName());
        file.setTrashed(true);
        n++;
      }
    }
  };
  var folders = work.getFolders();
  while (folders.hasNext()) {
    var f = folders.next();
    var isOffice = f.getFolders().hasNext();      // 中にフォルダがあれば事業所フォルダ
    if (group && isOffice && f.getName().indexOf(group) < 0) continue;
    if (!group || !isOffice) kill(f);             // 旧レイアウト（直下に個人フォルダ）
    var sub = f.getFolders();                     // 事業所フォルダの中
    while (sub.hasNext()) kill(sub.next());
  }
  lines.push('');
  lines.push('消した下書き: ' + n + '件（ゴミ箱に入りました。元に戻せます）');
  var msg = lines.join('\n');
  tkLog_('GCスライド', '下書き削除 ' + n + '件（' + label + '）');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg.slice(0, 1400)); } catch (e) {}
  return { deleted: n };
}
