// ============================================================
// gas_takamura ／ [60] セットアップ／プロビジョニング
// ------------------------------------------------------------
//  高村Workspace上で最初に1回だけ実行する初期構築。
//    SETUP_初期構築()  … SS・各シート・クエスト表・日報フォーム・トリガーを一括作成
//  実行後、ログに出るSS_ID / フォームURL を config.gs と案内に反映する。
// ============================================================

var TK_QUEST_HEADER = ['コード', '部署', '対象', '観点', 'GC分類', 'No', '小見出し', '内容', '出典'];
//  GC分類 = 週次GCシートの「外に生み出した成果」3分類（直接の成果／価値への取り組み／人材育成）
//  既定は 成果→直接の成果／姿勢→価値への取り組み／能力→人材育成。内容により一部例外あり。
//  ★このシート上で書き換えれば、GC下書きの振り分けも変わる（criteria.gs の修正は不要）

// クエスト表（サンクスUP基準）をシートへ展開
function tkWriteQuestSheet_() {
  var sh = tkEnsureSheet_(TK.SHEET_QUEST, TK_QUEST_HEADER);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, TK_QUEST_HEADER.length).clearContent();
  if (typeof TK_CRITERIA === 'undefined' || !TK_CRITERIA.length) throw new Error('criteria.gs（TK_CRITERIA）が読み込まれていません');
  var rows = TK_CRITERIA.map(function (c) {
    return [c.code, c.dept, c.level, c.axis, c.gc || '', c.no, c.title, c.desc, ''];
  });
  sh.getRange(2, 1, rows.length, TK_QUEST_HEADER.length).setValues(rows);
  CacheService.getScriptCache().remove('tk_quests');
  return rows.length;
}

// ============================================================
//  既存クエスト表に「GC分類」列を足す（7/24導入版 → 7/31以降のGC連携版への移行）
//  ------------------------------------------------------------
//  ★これが必要な理由：tkEnsureSheet_ はシートを新規作成するときだけヘッダを書く。
//    既にあるクエスト表（旧8列）は、コードを貼り替えてもヘッダが変わらない。
//    そのまま tkWriteQuestSheet_ を走らせると、データは9列・ヘッダは8列のまま＝列ズレになる。
//  ★SETUP_初期構築 を丸ごと再実行しなくて済むように、この1本だけを実行する。
//    旧シートは「クエスト表_旧_YYYYMMDD」にコピーして残す（消さない）。
// ============================================================
// ============================================================
//  メンバー表に「社員番号」列を足す（既存シートの移行用・1回だけ）
//  評価表／人事評価集計表と突き合わせるキー。
//  ★末尾に足すので、既存の列番号はひとつもずれない
// ============================================================
function tkメンバー表_社員番号列を追加() {
  var sh = tkEnsureSheet_(TK.SHEET_MEMBER, TK_MEMBER_HEADER);
  var head = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0]
    .map(function (v) { return String(v || '').trim(); });
  if (head.indexOf('社員番号') >= 0) {
    var m1 = 'すでに「社員番号」列があります。何もしませんでした。';
    Logger.log(m1);
    try { SpreadsheetApp.getUi().alert(m1); } catch (e) {}
    return { already: true };
  }
  var col = TK_MEMBER_HEADER.indexOf('社員番号') + 1;
  sh.getRange(1, col).setValue('社員番号');
  sh.getRange(1, 1, 1, TK_MEMBER_HEADER.length).setFontWeight('bold');
  try { sh.setFrozenRows(1); } catch (e) {}
  var NL = String.fromCharCode(10);
  var m = [
    'メンバー表に「社員番号」列を追加しました（' + col + '列目）。',
    '',
    '評価表・人事評価集計表と同じキーで人を突き合わせるための列です。',
    '名簿を頂いたら、この列に社員番号を入れてください。空欄のままでも動作します。'
  ].join(NL);
  tkLog_('移行', m); Logger.log(m);
  try { SpreadsheetApp.getUi().alert(m); } catch (e) {}
  return { added: true, col: col };
}

function tkクエスト表_GC分類列を追加() {
  var ss = tkSS_();
  var sh = ss.getSheetByName(TK.SHEET_QUEST);
  if (!sh) {
    var made = tkWriteQuestSheet_();
    var m0 = 'クエスト表が無かったので新規作成しました（' + made + '項目・GC分類列あり）。';
    tkLog_('移行', m0); Logger.log(m0);
    try { SpreadsheetApp.getUi().alert(m0); } catch (e) {}
    return { created: true, rows: made };
  }

  var head = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0]
    .map(function (v) { return String(v || '').trim(); });
  if (head.indexOf('GC分類') >= 0) {
    var m1 = 'すでに「GC分類」列があります。何もしませんでした。';
    Logger.log(m1);
    try { SpreadsheetApp.getUi().alert(m1); } catch (e) {}
    return { already: true };
  }

  // ---- 旧シートを退避（コピーを残す。元シートは作り直す） ----
  var stamp = Utilities.formatDate(new Date(), TK.TZ, 'yyyyMMdd');
  var bakName = TK.SHEET_QUEST + '_旧_' + stamp;
  if (ss.getSheetByName(bakName)) ss.deleteSheet(ss.getSheetByName(bakName));
  var bak = sh.copyTo(ss).setName(bakName);
  bak.hideSheet();

  // ---- 列数を確保してヘッダを9列に書き替え ----
  var needCols = TK_QUEST_HEADER.length;
  if (sh.getMaxColumns() < needCols) sh.insertColumnsAfter(sh.getMaxColumns(), needCols - sh.getMaxColumns());
  sh.getRange(1, 1, 1, needCols).setValues([TK_QUEST_HEADER]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, needCols).setFontWeight('bold').setBackground('#1B5E20').setFontColor('#FFFFFF');

  // ---- 96項目を criteria.gs から入れ直す（GC分類つき） ----
  var rows = tkWriteQuestSheet_();

  var msg = ['=== クエスト表に「GC分類」列を追加しました ===',
    '項目数: ' + rows + '件',
    '退避した旧シート: ' + bakName + '（非表示。中身を見たいときはシートの再表示から）',
    '',
    '★GC分類はこのシート上で自由に書き換えられます（コードの修正は不要）。',
    '　書き換えると、週次GC下書きの振り分けもそのとおりに変わります。',
    '　既定: 成果→直接の成果／姿勢→価値への取り組み／能力→人材育成（内容により一部例外あり）'
  ].join('\n');
  tkLog_('移行', msg);
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return { rows: rows, backup: bakName };
}

// 日報フォームを作成し、指定SSへ回答を出力。回答シートを「回答」にリネーム。
function tkCreateForm_(ss) {
  var form = FormApp.create('SUNグループ 日報（今日の実践・気づき）');
  form.setDescription('今日のAI実践・仕事の気づきを一言で。サンクスUP！評価（成果・能力・姿勢）に自動で反映されます。');
  form.setCollectEmail(true);
  form.addTextItem().setTitle('氏名').setRequired(true);
  form.addListItem().setTitle('部署').setChoiceValues(TK.DEPTS).setRequired(true);
  form.addParagraphTextItem().setTitle('今日の実践・気づき').setHelpText('例：担当ファミリィさんの体調変化に気づき、すぐ看護に共有した／Claudeで議事録を5分で作成した').setRequired(true);

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();
  // 追加された回答シートを探して「回答」にリネーム
  var sheets = ss.getSheets();
  var respSh = null;
  sheets.forEach(function (s) {
    try { if (!respSh && s.getFormUrl && s.getFormUrl()) respSh = s; } catch (e) {}
  });
  if (!respSh) sheets.forEach(function (s) { if (/フォームの回答|Form[ _]?Responses/i.test(s.getName())) respSh = s; });
  if (respSh) respSh.setName(TK.SHEET_RAW);
  return { editUrl: form.getEditUrl(), publishedUrl: form.getPublishedUrl(), formId: form.getId() };
}

// トリガー整備（onFormSubmit / 金曜応援下書き / 月初レポート）
function tkInstallTriggers_(ss) {
  // 既存の同名トリガーを掃除
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (['tkOnFormSubmit', 'tkWeeklyPraiseDraftTrigger', 'tkMonthlyReportTrigger'].indexOf(fn) >= 0) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('tkOnFormSubmit').forSpreadsheet(ss.getId()).onFormSubmit().create();
  ScriptApp.newTrigger('tkWeeklyPraiseDraftTrigger').timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(16).inTimezone(TK.TZ).create();
  ScriptApp.newTrigger('tkMonthlyReportTrigger').timeBased()
    .onMonthDay(1).atHour(7).inTimezone(TK.TZ).create();
}

// 金曜16時：応援下書きを自動準備（送信はしない）
function tkWeeklyPraiseDraftTrigger() { tkPraiseDraft(); }
// 毎月1日7時：先月分の月次レポート生成＋配信
function tkMonthlyReportTrigger() { tkGenerateMonthlyReport(); }

// ============ 初期構築（1回だけ実行） ============
function SETUP_初期構築() {
  var ss;
  if (TK.SS_ID) {
    ss = SpreadsheetApp.openById(TK.SS_ID);
  } else {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    ss = active || SpreadsheetApp.create('SUN日報解析');
  }
  var ssId = ss.getId();
  TK_SS_OVERRIDE = ss;   // この実行中は helper が必ずこのSSを使う

  // シート作成
  tkEnsureSheet_(TK.SHEET_QUEST, TK_QUEST_HEADER);
  tkEnsureSheet_(TK.SHEET_CLASS, TK_CLASS_HEADER);
  tkEnsureSheet_(TK.SHEET_ACT, TK_ACT_HEADER);
  tkEnsureSheet_(TK.SHEET_MEMBER, TK_MEMBER_HEADER);
  tkEnsureSheet_(TK.SHEET_OKR, TK_OKR_HEADER);
  tkEnsureSheet_(TK.SHEET_LOG, ['日時', '種別', '内容']);
  var qn = tkWriteQuestSheet_();

  // フォーム作成（既に回答シートがあればスキップ）
  var formInfo = null;
  if (!ss.getSheetByName(TK.SHEET_RAW)) {
    formInfo = tkCreateForm_(ss);
  }

  // トリガー
  tkInstallTriggers_(ss);

  // デフォルトシート「シート1」があれば削除（安全に）
  try {
    var def = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
    if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  } catch (e) {}

  var msg = [
    '=== gas_takamura 初期構築 完了 ===',
    'SS_ID: ' + ssId + '  ← config.gs の TK.SS_ID に貼ってください',
    'SS_URL: ' + ss.getUrl(),
    'クエスト表: ' + qn + '項目（サンクスUP基準）',
    formInfo ? ('日報フォーム（記入用URL）: ' + formInfo.publishedUrl) : '日報フォーム: 既存の「回答」シートを使用',
    formInfo ? ('日報フォーム（編集用URL）: ' + formInfo.editUrl) : '',
    'トリガー: onFormSubmit / 金曜16時 応援下書き / 毎月1日7時 月次レポート',
    '',
    '次の作業:',
    '1) config.gs の TK.SS_ID・GEMINI_API_KEYS・ADMIN_EMAIL を設定',
    '2) メンバー名簿（メンバーシート）に職員のメール・氏名・部署を登録',
    '3) 日報フォームURLを職員へ周知'
  ].filter(String).join('\n');
  Logger.log(msg);
  tkLog_('SETUP', msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return { ok: true, ssId: ssId, ssUrl: ss.getUrl(), quests: qn, form: formInfo };
}
