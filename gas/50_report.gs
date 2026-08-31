// ============================================================
// gas_takamura ／ [50] 月次 日報解析レポート自動生成（契約第5条）
// ------------------------------------------------------------
//  当月の日報（活動記録＋本文）をGeminiで解析し、契約第5条3項の4要素を含む
//  月次レポートを生成 → Googleドキュメント化 → 管理者へ自動配信。
//    (1) 各部署から提出された日報の要約
//    (2) 気づき・課題・好事例の抽出
//    (3) 評価制度（成果・能力・姿勢）観点からのコメント
//    (4) 翌月以降の重点フォロー候補
// ============================================================

// 対象月（yyyy, m）の範囲。省略時は先月。
function tkMonthRange_(y, m) {
  var now = new Date();
  if (!y || !m) {
    var first = new Date(now.getFullYear(), now.getMonth(), 1);
    var prev = new Date(first.getTime() - 1);   // 先月末
    y = prev.getFullYear(); m = prev.getMonth() + 1;
  }
  var start = new Date(y, m - 1, 1, 0, 0, 0);
  var end = new Date(y, m, 0, 23, 59, 59);
  return { start: start, end: end, label: y + '年' + m + '月', y: y, m: m };
}

// 月次レポート本文（テキスト）を生成
function tkBuildMonthlyReportText_(y, m) {
  var range = tkMonthRange_(y, m);
  var acts = tkReadActivity_().filter(function (a) {
    var ts = a.ts ? new Date(a.ts) : null; return ts && ts >= range.start && ts <= range.end;
  });
  // 部署別に集計
  var byDept = {};
  acts.forEach(function (a) {
    var d = a.dept || '（部署不明）';
    if (!byDept[d]) byDept[d] = { reports: 0, exp: 0, axis: { '成果': 0, '能力': 0, '姿勢': 0 }, samples: [] };
    byDept[d].reports++; byDept[d].exp += a.exp;
    if (byDept[d].axis[a.axis] != null) byDept[d].axis[a.axis]++;
    if (a.practice && a.practice.length >= 8 && byDept[d].samples.length < 12) {
      byDept[d].samples.push((a.name || '') + '｜' + a.title + '：' + a.practice);
    }
  });

  var deptBlocks = Object.keys(byDept).map(function (d) {
    var x = byDept[d];
    return '【' + d + '】報告' + x.reports + '件 / EXP' + x.exp +
      ' / 成果' + x.axis['成果'] + '・能力' + x.axis['能力'] + '・姿勢' + x.axis['姿勢'] + '\n' +
      x.samples.map(function (s) { return '  - ' + s; }).join('\n');
  }).join('\n\n');

  if (!acts.length) {
    return { label: range.label, text: range.label + 'は対象となる日報がありませんでした。', reports: 0 };
  }

  var prompt = [
    'あなたはSUNグループ（介護・障がい・薬局・本部）の人事・組織開発コンサルタントです。',
    '評価制度は「サンクスUP！」（成果・能力・姿勢の3観点）です。',
    '以下は' + range.label + 'に各部署から提出された日報の集計と実践内容です。',
    'これをもとに、経営層向けの「月次 日報解析レポート」を作成してください。',
    '',
    '=== 集計データ ===',
    deptBlocks,
    '',
    '=== 出力フォーマット（この見出し構成で・Markdownの##見出し使用） ===',
    '## 1. 今月の日報サマリー（部署別）',
    '## 2. 気づき・課題・好事例',
    '## 3. 評価観点コメント（成果・能力・姿勢）',
    '## 4. 翌月の重点フォロー候補',
    '',
    '各セクションは具体的に、実データを引用しつつ簡潔に。全体で1200〜1800字程度。',
    '個人を責める表現は避け、前向きな示唆として書く。'
  ].join('\n');

  var text = tkGeminiText_(prompt, { temperature: 0.4, maxTokens: 4096 });
  if (!text) {
    text = range.label + ' 日報解析レポート（自動集計のみ・AI生成失敗）\n\n' + deptBlocks;
  }
  return { label: range.label, text: text, reports: acts.length, deptBlocks: deptBlocks };
}

// Googleドキュメント化して管理者へ配信
function tkGenerateMonthlyReport(y, m) {
  var r = tkBuildMonthlyReportText_(y, m);
  var title = 'SUNグループ 月次日報解析レポート ' + r.label;
  var doc = DocumentApp.create(title);
  var body = doc.getBody();
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('生成日時：' + Utilities.formatDate(new Date(), TK.TZ, 'yyyy/MM/dd HH:mm') +
    '／対象日報 ' + r.reports + '件').setForegroundColor('#777777');
  // Markdown見出しを段落として流し込む（簡易）
  String(r.text).split('\n').forEach(function (line) {
    var l = line.replace(/\s+$/, '');
    if (/^##\s+/.test(l)) body.appendParagraph(l.replace(/^##\s+/, '')).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    else if (/^#\s+/.test(l)) body.appendParagraph(l.replace(/^#\s+/, '')).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    else body.appendParagraph(l);
  });
  doc.saveAndClose();

  var url = doc.getUrl();
  var admin = tkAdminEmail_();
  if (admin) {
    MailApp.sendEmail({
      to: admin, subject: '【' + r.label + '】SUNグループ 月次日報解析レポート',
      htmlBody: '<div style="font-family:sans-serif;line-height:1.8">' +
        r.label + 'の月次日報解析レポートを自動生成しました（対象日報 ' + r.reports + '件）。<br><br>' +
        '▶ <a href="' + url + '">' + tkEsc_(title) + '</a><br><br>' +
        '<span style="color:#777;font-size:12px">※契約第5条：日報要約／気づき・課題・好事例／評価観点コメント／翌月フォロー候補を含みます。</span></div>'
    });
  }
  tkLog_('月次レポート', r.label + ' 生成 ' + url);
  return { ok: true, label: r.label, url: url, reports: r.reports };
}
