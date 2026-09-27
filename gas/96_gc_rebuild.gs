// Slack等の日報が週次GC作成後に届いた場合の、対象者だけの安全な再生成。
// 既存スライド（本人の追記を含む）は編集・削除しない。新しい版を同じ個人フォルダに作る。
function tkGC_個人を再生成() {
  var ui = SpreadsheetApp.getUi();
  var namesPrompt = ui.prompt('個人GCを再生成',
    '対象者の氏名を名簿どおりに入力してください。複数名は「、」で区切ります。\n例：青木佳子、長峯妙子\n※同名の方がいる場合は中止します。',
    ui.ButtonSet.OK_CANCEL);
  if (namesPrompt.getSelectedButton() !== ui.Button.OK) return;
  var names = String(namesPrompt.getResponseText() || '').split(/[、,，\n]+/)
    .map(function (s) { return s.trim(); }).filter(String);
  names = names.filter(function (n, i) { return names.indexOf(n) === i; });
  if (!names.length || names.length > 10) {
    ui.alert('氏名を1〜10名入力してください。'); return;
  }

  var datePrompt = ui.prompt('対象週の終了日',
    '対象のGCシートに書かれた週の最終日を YYYY-MM-DD で入力してください。\n例：2026-09-25 → 2026 9/19〜9/25 の7日間',
    ui.ButtonSet.OK_CANCEL);
  if (datePrompt.getSelectedButton() !== ui.Button.OK) return;
  var raw = String(datePrompt.getResponseText() || '').trim();
  var m = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) { ui.alert('日付は YYYY-MM-DD で入力してください。'); return; }
  var end = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  if (end.getFullYear() !== Number(m[1]) || end.getMonth() !== Number(m[2]) - 1 ||
      end.getDate() !== Number(m[3]) || end > new Date()) {
    ui.alert('存在する過去または今日の日付を入力してください。'); return;
  }
  var range = tkGCWeekRange_(0, { endDate: end });
  var roster = tkGCSRoster_(), only = {}, selected = [];
  names.forEach(function (name) {
    var matches = roster.list.filter(function (p) { return p.name === name && !p.stop; });
    if (matches.length !== 1) throw new Error(name + '：名簿に1名と特定できません。名簿を確認してください。');
    only[matches[0].key] = true;
    selected.push(matches[0]);
  });
  var materials = tkGCSCollect_(range, '');
  var missing = selected.filter(function (p) {
    return !materials[p.key] || !materials[p.key].items.length;
  }).map(function (p) { return p.name; });
  if (missing.length) {
    ui.alert('対象期間に分類済みの日報がないため、中止しました：' + missing.join('、') +
      '\n日報統合・活動記録への反映を先に確認してください。');
    return;
  }
  var summary = selected.map(function (p) {
    return p.name + '：日報 ' + materials[p.key].items.length + '件';
  }).join('\n');
  var label = tkGCWeekLabel_(range);
  var confirmation = ui.alert('再生成の確認', '対象週：' + label + '\n' + summary +
    '\n\n旧版や本人の追記は残し、新しいGCシートを別ファイルとして作ります。' +
    '\nGoogle Chatへの自動通知はしません。実行しますか？', ui.ButtonSet.YES_NO);
  if (confirmation !== ui.Button.YES) return;

  var draft = tkGCBuild_(0, { range: range, onlyKeys: only, silent: true });
  var stamp = Utilities.formatDate(new Date(), TK.TZ, 'yyyyMMdd-HHmmss');
  var result = tkGCSBuild_(0, { range: range, onlyKeys: only,
    versionLabel: '再生成 ' + stamp, noMail: true, noChat: true, silent: true });
  tkGC再生成_名簿リンク更新_(result.files || []);
  var links = (result.files || []).map(function (f) { return f.name + '：' + f.url; });
  var message = '対象週：' + label + '\n作成：' + result.members + '名' +
    '／失敗：' + result.failed + '名\nGC下書きの新規行：' + (draft.members || 0) + '名' +
    '\n\n' + links.join('\n') + '\n\n旧版は残しています。新しい版の内容を確認してください。';
  tkLog_('GC再生成', message);
  ui.alert(message.slice(0, 1800));
  return result;
}

// 既存のメンバー表にリンク列がある場合だけ、新しく作った版のリンクへ差し替える。
function tkGC再生成_名簿リンク更新_(files) {
  if (!files.length) return;
  var sh = tkSS_().getSheetByName(TK.SHEET_MEMBER);
  if (!sh || sh.getLastRow() < 2) return;
  var width = sh.getLastColumn();
  var header = sh.getRange(1, 1, 1, width).getValues()[0];
  var iUrl = header.indexOf('最新GCシートURL') + 1;
  var iTime = header.indexOf('GC最終更新（日本時間）') + 1;
  if (!iUrl || !iTime) return;
  var names = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
  files.forEach(function (file) {
    var rows = [];
    names.forEach(function (v, i) { if (String(v[0] || '').trim() === file.name) rows.push(i + 2); });
    if (rows.length !== 1) return;
    try {
      var updated = DriveApp.getFileById(file.id).getLastUpdated();
      sh.getRange(rows[0], iUrl).setValue(file.url);
      sh.getRange(rows[0], iTime).setValue(Utilities.formatDate(updated, TK.TZ, 'yyyy/MM/dd HH:mm'));
    } catch (e) { tkLog_('GC再生成', file.name + ' の名簿リンク更新に失敗: ' + e.message); }
  });
}
