// ============================================================
// gas_takamura ／ 共通ユーティリティ
// ============================================================

// 実行中だけ有効なSS上書き（SETUP直後・onFormSubmit時に e.source を使うため）
var TK_SS_OVERRIDE = null;

function tkSS_() {
  if (TK_SS_OVERRIDE) return TK_SS_OVERRIDE;
  if (TK.SS_ID) return SpreadsheetApp.openById(TK.SS_ID);
  // SS_ID未設定時は、このスクリプトにバインドされたSSを使う（コンテナバインド運用も可）
  var act = SpreadsheetApp.getActiveSpreadsheet();
  if (act) return act;
  throw new Error('TK.SS_ID が未設定です。config.gs に「SUN日報解析」SSのIDを入れてください。');
}

// シート取得（無ければヘッダ付きで作成）
function tkEnsureSheet_(name, header) {
  var ss = tkSS_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (header && header.length) {
      sh.getRange(1, 1, 1, header.length).setValues([header]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#1B5E20').setFontColor('#FFFFFF');
    }
  }
  return sh;
}

function tkEsc_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// メンバーID＝メールアドレス（小文字）を安定キーとして使う
function tkMemberId_(email) {
  return String(email || '').trim().toLowerCase();
}

// 集計のキー。メールが無い人は氏名で代用する。
//  ⚠ ネットワークのフォームにはメール列が無く、名簿にメールが入っていない方が多い。
//    メールだけをキーにしていたため、活動記録はあるのにGC下書きが作られない人がいた（2026-09-09 修正）。
function tkPersonKey_(mid, name) {
  var m = String(mid || '').trim().toLowerCase();
  if (m) return m;
  var n = String(name || '').replace(/\s/g, '').trim();
  return n ? '名前:' + n : '';
}
// キーがメールなら本人に共有できる。氏名キーは共有できない
function tkKeyIsMail_(k) { return String(k || '').indexOf('@') > 0; }

// 管理者メール（config優先／無ければデプロイ実行者）
function tkAdminEmail_() {
  if (TK.ADMIN_EMAIL) return TK.ADMIN_EMAIL;
  try { return Session.getEffectiveUser().getEmail(); } catch (e) { return ''; }
}

// 今週（月〜日）の範囲。refDate省略時は今日基準。
function tkWeekRange_(refDate) {
  var base = refDate ? new Date(refDate) : new Date();
  var day = base.getDay();                 // 0=日
  var diffToMon = (day === 0 ? -6 : 1 - day);
  var start = new Date(base); start.setDate(base.getDate() + diffToMon); start.setHours(0, 0, 0, 0);
  var end = new Date(start); end.setDate(start.getDate() + 6); end.setHours(23, 59, 59, 999);
  var label = Utilities.formatDate(start, TK.TZ, 'M/d') + '〜' + Utilities.formatDate(end, TK.TZ, 'M/d');
  return { start: start, end: end, label: label };
}

// 先頭が = + - @ だとSheetsが数式と誤認するので ' でエスケープ
function tkSafeCell_(s) {
  s = String(s == null ? '' : s);
  return /^[=+\-@]/.test(s) ? ("'" + s) : s;
}

// 実行ログ追記
function tkLog_(tag, msg) {
  try {
    var sh = tkEnsureSheet_(TK.SHEET_LOG, ['日時', '種別', '内容']);
    sh.appendRow([new Date(), tkSafeCell_(tag), tkSafeCell_(String(msg).slice(0, 4000))]);
  } catch (e) { Logger.log(tag + ': ' + msg); }
}

// 回答シート（フォームのリンク先）を名前ゆらぎに強く取得
//  「回答」/「Form Responses」/「Form_Responses」/「フォームの回答 1」等に対応
function tkRawSheet_() {
  var ss = tkSS_();
  var sh = ss.getSheetByName(TK.SHEET_RAW);
  if (sh) return sh;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    try { if (sheets[i].getFormUrl && sheets[i].getFormUrl()) return sheets[i]; } catch (e) {}
  }
  for (var j = 0; j < sheets.length; j++) {
    if (/回答|フォーム|Form[ _]?Responses/i.test(sheets[j].getName())) return sheets[j];
  }
  return null;
}

// ---------- 秘密の値（スクリプトプロパティ優先 → config の SECRET_CONFIG） ----------
//  五十嵐と同じ方式。プロパティに入れておけばコードに鍵を書かずに済み、Git管理できる。
//  名前: GEMINI_API_KEYS / GEMINI_MODEL / APP_TOKEN
function tkSecret_(name) {
  var v = '';
  try { v = PropertiesService.getScriptProperties().getProperty(name) || ''; } catch (e) {}
  if (!v && typeof SECRET_CONFIG !== 'undefined') v = SECRET_CONFIG[name] || '';
  return String(v || '').trim();
}
function tkGeminiKeys_() {
  return tkSecret_('GEMINI_API_KEYS').split(',').map(function (s) { return s.trim(); }).filter(String);
}
//  モデル名は「英小文字で始まり、英数字・ハイフン・ドット」だけ。
//  ⚠ ここに鍵や空白が入ると Google が「unexpected model name format」(HTTP 400) を返し、
//    判定が全部止まる。2026-09-17 に実際に踏んだので、形が違う値は既定値に読み替える。
var TK_MODEL_DEFAULT = 'gemini-flash-latest';
function tkModelLooksValid_(v) {
  return /^[a-z][a-z0-9.\-]{2,60}$/.test(String(v || ''));
}
function tkGeminiModel_() {
  var v = tkSecret_('GEMINI_MODEL').replace(/^models\//, '').trim();   // 'models/xxx' でも通す
  if (!v) return TK_MODEL_DEFAULT;
  if (!tkModelLooksValid_(v)) {
    try { tkLog_('鍵', 'モデル名の形が違うため既定値（' + TK_MODEL_DEFAULT + '）を使います'); } catch (e) {}
    return TK_MODEL_DEFAULT;
  }
  return v;
}

// 値そのものを出さずに、長さと末尾だけ見せる
function tkMaskKey_(k) {
  k = String(k || '');
  return k ? ('長さ' + k.length + '・末尾 ' + k.slice(-4)) : '（未設定）';
}

// エディタから実行：鍵をスクリプトプロパティに登録する（値はログに出さない）
function tk鍵を登録() {
  var ui = SpreadsheetApp.getUi();
  var names = ['GEMINI_API_KEYS', 'GEMINI_MODEL', 'APP_TOKEN'];
  var props = PropertiesService.getScriptProperties();
  var done = [];
  names.forEach(function (n) {
    var cur = props.getProperty(n) || '';
    var state = cur ? '登録済み' : (SECRET_CONFIG[n] ? 'config.gs の値を使用中' : '未設定');
    var r = ui.prompt('AIの鍵を登録', n + '\n（いま: ' + state + '）\n新しい値を入れて OK。空のまま OK で変更なし。', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    var v = String(r.getResponseText() || '').trim();
    if (v) { props.setProperty(n, v); done.push(n); }
  });
  var msg = done.length ? ('登録しました: ' + done.join(', ') + '\n\nconfig.gs の同じ項目は空にして構いません。')
                        : '変更はありませんでした。';
  tkLog_('鍵', done.length ? '登録: ' + done.join(', ') : '変更なし');
  ui.alert(msg);
}
// どこから読んでいるかを見る（値は出さない）
function tk鍵の状態() {
  var props = PropertiesService.getScriptProperties();
  var lines = ['=== 秘密の値の置き場 ==='];
  ['GEMINI_API_KEYS', 'GEMINI_MODEL', 'APP_TOKEN'].forEach(function (n) {
    var p = props.getProperty(n), c = SECRET_CONFIG[n];
    var v = String(p || c || '');
    lines.push('　' + n + ': ' + (p ? 'スクリプトプロパティ' : (c ? 'config.gs' : '⚠ 未設定')) +
      (n === 'GEMINI_API_KEYS' && v ? '（' + v.split(',').length + '本・' + tkMaskKey_(v) + '）' : ''));
    // ⚠ モデル名やトークンの欄に鍵を貼ってしまう取り違えが起きやすい
    if (n === 'GEMINI_MODEL' && v && !tkModelLooksValid_(v.replace(/^models\//, '').trim())) {
      lines.push('　　🔴 モデル名の形ではありません。既定値（' + TK_MODEL_DEFAULT + '）で動いています。');
      lines.push('　　　 この欄は空にするか gemini-flash-latest と入れてください。');
    }
    if (n === 'APP_TOKEN' && v && (v.indexOf('AQ.') === 0 || v.length > 40)) {
      lines.push('　　🔴 APIキーが入っている可能性があります。到達マンダラのURLの key= と一致しません。');
    }
  });
  lines.push('');
  lines.push('　いま使うモデル: ' + tkGeminiModel_());
  var msg = lines.join('\n'); Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return msg;
}

// ---------- Gemini（gas_nippo と同じ方式・JSON強制） ----------
function tkGeminiJson_(prompt, schema) {
  var keys = tkGeminiKeys_();
  if (!keys.length) return { ok: false, error: 'Gemini APIキー未設定' };
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
    tkGeminiModel_() + ':generateContent';
  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { responseMimeType: 'application/json', responseSchema: schema, temperature: 0.1 }
  };
  var lastErr = '';
  for (var i = 0; i < keys.length; i++) {
    try {
      var res = UrlFetchApp.fetch(url + '?key=' + encodeURIComponent(keys[i]), {
        method: 'post', contentType: 'application/json',
        payload: JSON.stringify(payload), muteHttpExceptions: true
      });
      var code = res.getResponseCode();
      if (code === 429 || code >= 500) { lastErr = 'HTTP ' + code; continue; }
      var body = JSON.parse(res.getContentText());
      if (code !== 200) {
        // ⚠ 400/403 は設定の誤り。上限（429）と同じ扱いにすると、ログで見分けがつかない
        lastErr = 'HTTP ' + code + ' ' + ((body.error && body.error.message) || '');
        if (code === 400 || code === 403) {
          try { tkLog_('鍵', '設定の誤りで呼び出せません: ' + lastErr.slice(0, 200)); } catch (e) {}
          return { ok: false, error: lastErr, config: true };
        }
        continue;
      }
      var textOut = body.candidates && body.candidates[0] && body.candidates[0].content &&
        body.candidates[0].content.parts && body.candidates[0].content.parts[0] &&
        body.candidates[0].content.parts[0].text;
      if (!textOut) { lastErr = 'empty response'; continue; }
      return { ok: true, data: JSON.parse(textOut) };
    } catch (e) { lastErr = String(e && e.message || e); }
  }
  return { ok: false, error: lastErr || 'gemini failed' };
}

// Gemini（自由文・JSONなし）。応援文/レポート生成用。
function tkGeminiText_(prompt, opt) {
  opt = opt || {};
  var keys = tkGeminiKeys_();
  if (!keys.length) return '';
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
    tkGeminiModel_() + ':generateContent';
  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: opt.temperature || 0.7, maxOutputTokens: opt.maxTokens || 2048 }
  };
  for (var i = 0; i < keys.length; i++) {
    try {
      var res = UrlFetchApp.fetch(url + '?key=' + encodeURIComponent(keys[i]), {
        method: 'post', contentType: 'application/json',
        payload: JSON.stringify(payload), muteHttpExceptions: true
      });
      var code = res.getResponseCode();
      if (code === 429 || code >= 500) continue;
      var body = JSON.parse(res.getContentText());
      if (code !== 200) continue;
      var t = body.candidates && body.candidates[0] && body.candidates[0].content &&
        body.candidates[0].content.parts && body.candidates[0].content.parts[0] &&
        body.candidates[0].content.parts[0].text;
      if (t) return String(t).trim();
    } catch (e) {}
  }
  return '';
}
