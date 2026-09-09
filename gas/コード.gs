// ============================================================
// gas_takamura ／ Webアプリ入口 ＆ スプレッドシートメニュー
// ------------------------------------------------------------
//  Web API（?key=APP_TOKEN 必須）:
//    ?action=status                        稼働状況
//    ?action=backlog&n=10                   未分類の日報をまとめて分類
//    ?action=test&text=...&dept=輝らら       分類のお試し
//    ?action=syncActivity                   分類済み→活動記録 再取り込み（手動修正後）
//    ?action=rebuildOkr                     OKR集計 再計算
//    ?action=praiseDraft / praiseSend       週次応援 下書き/送信
//    ?action=monthlyReport&y=2026&m=8       月次レポート生成
//  Web表示（HTML）:
//    ?view=okr&member=<メール>&key=APP_TOKEN 到達マンダラ
//    ?view=gc&key=APP_TOKEN                 GC下書きの一覧（推進メンバー向け・直近の週）
//      &days=10 で遡る日数、&group=薬局・ネットワーク で部署を絞る
// ============================================================

function doGet(e) {
  var p = (e && e.parameter) || {};

  // ---- HTMLビュー（GC下書き一覧：推進メンバーが「どこで見るか」の答え）----
  if (p.view === 'gc') {
    if (p.key !== tkSecret_('APP_TOKEN')) {
      return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:24px">アクセスキーが必要です。</p>');
    }
    return tkGCSListHtml_(p.key, { days: +p.days || 10, group: String(p.group || '').trim() });
  }

  // ---- HTMLビュー（到達マンダラ）----
  if (p.view === 'okr') {
    if (p.key !== tkSecret_('APP_TOKEN')) {
      return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:24px">アクセスキーが必要です。</p>');
    }
    var mem = String(p.member || '').trim();
    // member 指定なし → 全員の一覧（社内に配るのはこのURL1本）
    return mem ? tkOkrMandalaHtml_(mem) : tkOkrIndexHtml_(p.key);
  }

  // ---- JSON API ----
  var out;
  try {
    if (p.key !== tkSecret_('APP_TOKEN')) out = { ok: false, error: 'bad key' };
    else if (p.action === 'status')        out = tkStatus_();
    // 未分類のまとめ判定はバッチ版を使う（1件1API の旧 tkProcessBacklog_ は無料枠を食い潰す）
    else if (p.action === 'backlog')       out = tkBulkRun_();
    else if (p.action === 'test')          out = tkClassify_(String(p.text || ''), String(p.dept || ''));
    else if (p.action === 'syncActivity')  out = tkSyncClassifiedToActivity_();
    else if (p.action === 'rebuildOkr')    out = tkRebuildOkr_();
    else if (p.action === 'praiseDraft')   out = tkPraiseDraft();
    else if (p.action === 'praiseSend')    out = tkPraiseSend();
    else if (p.action === 'monthlyReport') out = tkGenerateMonthlyReport(+p.y || 0, +p.m || 0);
    else out = { ok: false, error: 'unknown action' };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }

  var json = JSON.stringify(out);
  if (p.callback && /^[\w.]+$/.test(p.callback)) {
    return ContentService.createTextOutput(p.callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function tkStatus_() {
  var ss;
  try { ss = tkSS_(); } catch (e) { return { ok: false, error: e.message }; }
  function n(name) { var s = ss.getSheetByName(name); return s ? Math.max(0, s.getLastRow() - 1) : '(なし)'; }
  var rawSh = tkRawSheet_();
  var triggers = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  var gk = tkSecret_('GEMINI_API_KEYS');
  return {
    ok: true, ssId: ss.getId(),
    sheets: {
      クエスト表: n(TK.SHEET_QUEST), 回答: (rawSh ? Math.max(0, rawSh.getLastRow() - 1) : '(なし)'), 分類済み: n(TK.SHEET_CLASS),
      活動記録: n(TK.SHEET_ACT), メンバー: n(TK.SHEET_MEMBER), OKR集計: n(TK.SHEET_OKR)
    },
    triggers: triggers,
    geminiKeySet: (gk.indexOf('＜') < 0 && gk.trim().length > 10)
  };
}

// ---- スプレッドシートを開いたときのメニュー ----
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('🌞 SUN日報解析')
      .addItem('① 初期構築（最初に1回）', 'SETUP_初期構築')
      .addSeparator()
      .addItem('② 未分類の日報をまとめて分類', 'menu_backlog')
      .addItem('③ 分類済み→活動記録 再取り込み', 'menu_syncActivity')
      .addItem('④ OKR集計を再計算', 'menu_rebuildOkr')
      .addSeparator()
      .addItem('⑤ 週次応援 下書き（管理者へプレビュー）', 'tkPraiseDraft')
      .addItem('⑥ 週次応援 送信', 'tkPraiseSend')
      .addSeparator()
      .addItem('⑦ 月次レポート生成（先月分）', 'tkMonthlyReportTrigger')
      .addSeparator()
      .addItem('⑧ クエスト表にGC分類列を追加（移行時に1回）', 'tkクエスト表_GC分類列を追加')
      .addItem('⑧- メンバー表に社員番号列を追加（移行時に1回）', 'tkメンバー表_社員番号列を追加')
      .addItem('⑨ 自動判定を設定（1時間ごと・最初に1回）', 'tkキャッチアップ_設定')
      .addItem('⑩ 分類の進捗を見る', 'tk一括分類_状況')
      .addItem('⑪ APIキーの診断', 'tkAPIキー確認')
      .addSeparator()
      .addItem('⑫ GC下書きを今すぐ作る（部署を選ぶ）', 'tkGC下書き_今すぐ作る')
      .addItem('⑬ 設定シートを作る（GC設定・推進メンバー・取込ソース／1回）', 'tk設定シート_作成')
      .addItem('⑬- 推進メンバーに下書きフォルダを共有', 'tkGCスライド_推進メンバーに共有')
      .addItem('⑭ GCの自動作成を毎日17:30に設定（金曜固定をやめる／1回）', 'tkGC_トリガー設定')
      .addItem('⑮ AIの鍵をスクリプトプロパティに登録', 'tk鍵を登録')
      .addItem('⑯ 下書きが出ない理由を人ごとに出す', 'tkGC_なぜ出ないか')
      .addToUi();
  } catch (e) {}
}

function menu_backlog() {
  // バッチ版（Gemini 15件まとめて1回）。1件1APIの旧 tkProcessBacklog_ は無料枠を食い潰すため使わない。
  var r = tkBulkRun_();
  try {
    SpreadsheetApp.getUi().alert('分類しました: ' + r.processed + '件 / 未判定の残り ' + r.remaining + '件'
      + (r.quota ? '\n\n※APIの利用上限に当たって中断しました。次の自動判定で再挑戦します。' : ''));
  } catch (e) {}
}
function menu_syncActivity() {
  var r = tkSyncClassifiedToActivity_();
  try { SpreadsheetApp.getUi().alert('活動記録に追加: ' + r.added + '件'); } catch (e) {}
}
function menu_rebuildOkr() {
  var r = tkRebuildOkr_();
  try { SpreadsheetApp.getUi().alert('OKR集計 再計算: ' + r.members + '名'); } catch (e) {}
}
