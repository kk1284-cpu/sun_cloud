var TK_SLACK_CHANNEL = 'C0946CH2EGJ';
var TK_SLACK_SOURCE = 'slack_care';
var TK_SLACK_CURSOR_KEY = 'TK_SLACK_CARE_CURSOR';
var TK_SLACK_OLDEST_KEY = 'TK_SLACK_CARE_OLDEST';
var TK_SLACK_MAX_KEY = 'TK_SLACK_CARE_MAX_TS';

function tkSlack接続確認() {
  var token = PropertiesService.getScriptProperties().getProperty('TK_SLACK_BOT_TOKEN');
  if (!token) return { ready: false, reason: 'TK_SLACK_BOT_TOKEN が未設定です' };
  var info = tkSlackApi_('auth.test', {}, token);
  var channel = tkSlackApi_('conversations.info', { channel: TK_SLACK_CHANNEL }, token);
  return { ready: true, workspace: info.team, channel: channel.channel && channel.channel.name };
}

// Slack履歴は1回1ページ。続きは次回の定時実行で取得する。
function tkSlack日報取込() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { added: 0, status: 'running' };
  try {
    var props = PropertiesService.getScriptProperties();
    var token = props.getProperty('TK_SLACK_BOT_TOKEN');
    if (!token) throw new Error('TK_SLACK_BOT_TOKEN が未設定です。Slackからの自動取得は未開始です。');
    var oldest = props.getProperty(TK_SLACK_OLDEST_KEY);
    if (!oldest) {
      oldest = String(Math.floor(Date.now() / 1000) - 7 * 86400);
      props.setProperty(TK_SLACK_OLDEST_KEY, oldest);
    }
    var cursor = props.getProperty(TK_SLACK_CURSOR_KEY);
    var params = { channel: TK_SLACK_CHANNEL, oldest: oldest, inclusive: 'false', limit: 100 };
    if (cursor) params.cursor = cursor;
    var page = tkSlackApi_('conversations.history', params, token);
    var out = tkEnsureSheet_(TK.SHEET_RAW, TK_UNI_HEADER);
    var seen = tkUniSeen_(out);
    var roster = tkUniRoster_();
    var rows = [];
    var maxTs = props.getProperty(TK_SLACK_MAX_KEY) || oldest;
    var messages = page.messages || [];
    for (var i = 0; i < messages.length; i++) {
      var m = messages[i];
      if (!m || !m.ts) continue;
      if (Number(m.ts) > Number(maxTs)) maxTs = m.ts;
      if (m.subtype || m.bot_id || (m.thread_ts && m.thread_ts !== m.ts)) continue;
      var body = String(m.text || '').trim();
      if (!body || /^(test|テスト|接続テスト)$/i.test(body)) continue;
      var sourceRow = TK_SLACK_CHANNEL + '#' + m.ts;
      var key = TK_SLACK_SOURCE + '#' + sourceRow;
      if (seen[key]) continue;
      var person = tkSlackPerson_(m.user, token, roster);
      rows.push([new Date(Number(m.ts) * 1000), person.email, person.name, '介護事業部',
        tkSafeCell_(body), TK_SLACK_SOURCE, sourceRow, '', '', '']);
      seen[key] = 1;
    }
    rows.sort(function (a, b) { return a[0] - b[0]; });
    if (rows.length) out.getRange(out.getLastRow() + 1, 1, rows.length, TK_UNI_HEADER.length).setValues(rows);
    var next = page.response_metadata && page.response_metadata.next_cursor || '';
    if (next) {
      props.setProperty(TK_SLACK_CURSOR_KEY, next);
      props.setProperty(TK_SLACK_MAX_KEY, maxTs);
    } else {
      props.deleteProperty(TK_SLACK_CURSOR_KEY);
      props.deleteProperty(TK_SLACK_MAX_KEY);
      props.setProperty(TK_SLACK_OLDEST_KEY, maxTs);
    }
    tkLog_('Slack取込', '追加 ' + rows.length + '件／続き ' + (next ? 'あり' : 'なし'));
    return { added: rows.length, scanned: messages.length, more: !!next };
  } finally {
    lock.releaseLock();
  }
}

function tkSlackPerson_(userId, token, roster) {
  if (!userId) return { name: '', email: '' };
  var cache = CacheService.getScriptCache();
  var cached = cache.get('slack_user_' + userId);
  var person;
  if (cached) person = JSON.parse(cached);
  else {
    var user = tkSlackApi_('users.info', { user: userId }, token).user || {};
    var profile = user.profile || {};
    person = { name: String(profile.real_name || user.real_name || profile.display_name || user.name || userId).trim(),
      email: String(profile.email || '').trim() };
    cache.put('slack_user_' + userId, JSON.stringify(person), 21600);
  }
  var flat = person.name.replace(/[\s　]/g, '');
  for (var i = 0; i < roster.names.length; i++) {
    var name = roster.names[i];
    if (name.replace(/[\s　]/g, '') === flat)
      return { name: name, email: person.email || roster.byName[name] || '' };
  }
  return person;
}

function tkSlackApi_(method, params, token) {
  var query = Object.keys(params).map(function (k) {
    return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
  }).join('&');
  var response = UrlFetchApp.fetch('https://slack.com/api/' + method + (query ? '?' + query : ''), {
    method: 'get', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
  });
  var data = JSON.parse(response.getContentText());
  if (response.getResponseCode() !== 200 || !data.ok)
    throw new Error('Slack API ' + method + ': ' + (data.error || response.getResponseCode()));
  return data;
}

function tkSlack取込トリガー設定() {
  var ready = tkSlack接続確認();
  if (!ready.ready) throw new Error(ready.reason);
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++)
    if (triggers[i].getHandlerFunction() === 'tkSlack日報取込') return { existing: true };
  ScriptApp.newTrigger('tkSlack日報取込').timeBased().everyHours(1).create();
  return { created: true };
}
