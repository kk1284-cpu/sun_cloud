// ============================================================
// gas_takamura ／ [40] 週次「応援（がんばり賞賛）」メール
// ------------------------------------------------------------
//  活動記録から今週分を集計し、トップランナー＆素敵な実践を称えるメールを作る。
//  ・承認制：下書き生成 → 管理者へプレビュー → OKなら送信（誤送信防止）。
//  ・配信停止（メンバー名簿G列='停止'）は除外。
//  gas_main の MCQ110_cohort_praise.gs を自己完結・部署横断に簡素化した版。
// ============================================================

var TK_PRAISE_DRAFT_KEY = 'TK_PRAISE_DRAFT';

// 今週分ハイライト抽出
function tkPraiseCollect_(refDate) {
  var range = tkWeekRange_(refDate);
  var acts = tkReadActivity_();
  var per = {}, highlights = [], totalReports = 0, totalExp = 0;
  acts.forEach(function (a) {
    var ts = a.ts ? new Date(a.ts) : null;
    if (!ts || ts < range.start || ts > range.end) return;
    if (!a.mid) return;
    if (!per[a.mid]) per[a.mid] = { name: a.name || a.mid, dept: a.dept, exp: 0, reports: 0 };
    per[a.mid].exp += a.exp; per[a.mid].reports += 1;
    totalReports++; totalExp += a.exp;
    if (a.practice && a.practice.length >= 10) {
      highlights.push({ mid: a.mid, name: a.name || a.mid, dept: a.dept, title: a.title, exp: a.exp, practice: a.practice });
    }
  });
  var arr = Object.keys(per).map(function (id) { return per[id]; });
  arr.sort(function (a, b) { return (b.exp - a.exp) || (b.reports - a.reports); });
  highlights.sort(function (a, b) { return b.exp - a.exp; });
  var seen = {}, picked = [];
  highlights.forEach(function (h) {
    if (seen[h.mid]) return; seen[h.mid] = true;
    var p = h.practice.length > 120 ? h.practice.substring(0, 117) + '…' : h.practice;
    picked.push({ name: h.name, dept: h.dept, title: h.title, practice: p });
  });
  return {
    weekLabel: range.label, top3: arr.slice(0, 3), highlights: picked.slice(0, 5),
    totalReports: totalReports, totalExp: totalExp, activeMembers: arr.length
  };
}

// 賞賛文（Gemini＋キャラ／失敗時テンプレ）
function tkPraiseText_(data) {
  var top3 = data.top3.map(function (t, i) {
    return (i + 1) + '位 ' + t.name + 'さん（' + (t.dept || '') + '／報告' + t.reports + '回・EXP' + t.exp + '）';
  }).join('\n');
  var hl = data.highlights.map(function (h) {
    return '・' + h.name + 'さん' + (h.title ? '（' + h.title + '）' : '') + '：' + h.practice;
  }).join('\n');

  var prompt = [
    'あなたは「' + TK.PRAISE_CHARA + '」という、SUNグループのAI実践と評価制度「サンクスUP！」を応援する温かいナビゲーターです。',
    'SUNグループ全職員向けに、週1回の「がんばり賞賛」メール本文を書いてください。',
    '目的は、今週がんばった人を称え、全員のモチベーションを上げること。',
    '',
    '【厳守】入会勧誘・他社宣伝は書かない。温かく前向きに、具体的な実践を引用して褒める。',
    '本文は400〜600字。見出しに絵文字OK。末尾の署名や配信停止案内は書かない（システムが付与）。',
    '',
    '【今週のデータ（' + data.weekLabel + '）】',
    '■ トップ3', top3 || '（該当者少なめ）',
    '', '■ 素敵な実践', hl || '（実践記載少なめ）',
    '', '■ 全体：日報 ' + data.totalReports + '件／獲得EXP ' + data.totalExp + '／活動者 ' + data.activeMembers + '名',
    '',
    '出力はJSONのみ（件名に絵文字は使わない）:',
    '{"subject":"件名(30字以内・絵文字なし)","body":"本文"}'
  ].join('\n');

  var raw = tkGeminiText_(prompt, { temperature: 0.7, maxTokens: 2048 });
  if (raw) {
    try {
      var m = raw.match(/\{[\s\S]*\}/);
      if (m) { var o = JSON.parse(m[0]); if (o.subject && o.body) return { subject: String(o.subject).trim(), body: String(o.body).trim() }; }
    } catch (e) {}
  }
  return tkPraiseFallback_(data);
}

function tkPraiseFallback_(data) {
  var L = '━━━━━━━━━━━━━━━\n', b = [];
  b.push('SUNグループのみなさん、今週もお疲れさまです！\n');
  b.push(L + '🏆 今週のトップランナー\n' + L);
  if (data.top3.length) data.top3.forEach(function (t, i) { b.push((i + 1) + '位　' + t.name + 'さん（' + (t.dept || '') + '／報告' + t.reports + '回・EXP' + t.exp + '）'); });
  else b.push('今週は活動が少なめでした。来週はぜひ「今日やってみた」を1件！');
  if (data.highlights.length) {
    b.push('\n' + L + '✨ 今週の素敵な実践\n' + L);
    data.highlights.forEach(function (h) { b.push('▶ ' + h.name + 'さん' + (h.title ? '（' + h.title + '）' : '') + '\n　「' + h.practice + '」'); });
  }
  b.push('\n' + L + '📊 今週のチーム全体\n' + L);
  b.push('・日報：' + data.totalReports + '件\n・獲得EXP：' + data.totalExp + '\n・活動メンバー：' + data.activeMembers + '名');
  b.push('\n小さな一歩の積み重ねが、確かな力になります。来週も一緒にコツコツいきましょう！');
  return { subject: '今週のがんばり（' + data.weekLabel + '）｜SUNグループ', body: b.join('\n') };
}

// 送信対象（メンバー名簿・配信停止除外）
function tkPraiseTargets_() {
  var sh = tkSS_().getSheetByName(TK.SHEET_MEMBER);
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, TK_MEMBER_HEADER.length).getValues();
  var out = [];
  vals.forEach(function (r) {
    var email = String(r[0] || '').trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return;
    if (String(r[6] || '').trim() === '停止') return;
    out.push({ email: email, name: String(r[1] || '').trim(), dept: String(r[2] || '').trim() });
  });
  return out;
}

function tkBodyToHtml_(body) {
  return '<div style="font-family:\'Hiragino Sans\',\'Yu Gothic\',sans-serif;font-size:14px;line-height:1.8;color:#333;white-space:pre-wrap">' +
    tkEsc_(body) + '</div>';
}

// [OU1] 下書き生成＋管理者へプレビュー
function tkPraiseDraft() {
  var data = tkPraiseCollect_();
  var draft = tkPraiseText_(data);
  PropertiesService.getScriptProperties().setProperty(TK_PRAISE_DRAFT_KEY, JSON.stringify({
    subject: draft.subject, body: draft.body, weekLabel: data.weekLabel,
    savedAt: Utilities.formatDate(new Date(), TK.TZ, 'yyyy/MM/dd HH:mm')
  }));
  var admin = tkAdminEmail_();
  var targets = tkPraiseTargets_();
  if (admin) {
    MailApp.sendEmail({
      to: admin, subject: '【プレビュー／要確認】' + draft.subject,
      htmlBody: '<div style="background:#FFF3E0;padding:8px;border-radius:6px;margin-bottom:10px;font-size:13px">' +
        '⚠ これは確認用プレビューです（職員には未送信）。<br>対象週: ' + tkEsc_(data.weekLabel) +
        '／送信対象: ' + targets.length + '名。<br>内容OKなら tkPraiseSend() を実行してください。</div>' +
        tkBodyToHtml_(draft.body)
    });
  }
  tkLog_('応援', '下書き生成 ' + data.weekLabel + ' 対象' + targets.length + '名');
  return { ok: true, weekLabel: data.weekLabel, subject: draft.subject, targets: targets.length, top3: data.top3.length, highlights: data.highlights.length };
}

// [OU2] 送信実行（下書きを職員へ一斉）
function tkPraiseSend() {
  var raw = PropertiesService.getScriptProperties().getProperty(TK_PRAISE_DRAFT_KEY);
  if (!raw) return { ok: false, error: '下書きがありません。先に tkPraiseDraft() を実行してください' };
  var draft = JSON.parse(raw);
  var targets = tkPraiseTargets_();
  if (!targets.length) return { ok: false, error: '送信対象が0名です' };
  var sent = 0, fail = 0;
  targets.forEach(function (t) {
    try {
      MailApp.sendEmail({ to: t.email, subject: draft.subject, htmlBody: tkBodyToHtml_(draft.body) });
      sent++; Utilities.sleep(150);
    } catch (e) { fail++; tkLog_('応援', '送信失敗 ' + t.email + ': ' + e.message); }
  });
  tkLog_('応援', '送信完了 ' + draft.weekLabel + ' ' + sent + '名（失敗' + fail + '）');
  return { ok: true, sent: sent, fail: fail, weekLabel: draft.weekLabel };
}
