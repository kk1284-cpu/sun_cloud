// ============================================================
// gas_takamura ／ [30] マンダラOKR反映（活動記録 → 到達マンダラ）
// ------------------------------------------------------------
//  活動記録を集計して、部署×観点（成果/能力/姿勢）の到達度を出す。
//  ・OKR集計シート：メンバー別サマリ（自動再計算）
//  ・Web表示 ?view=okr&member=<メール>：本人の到達マンダラ（3観点×4項目=12マス）
//  マンダラOKRの「成果1/能力1/姿勢1…」の枠に、日報から積み上がった実績が反映される。
// ============================================================

var TK_OKR_HEADER = [
  'メンバーID', '氏名', '部署', '累計EXP', '報告数',
  '成果_件数', '能力_件数', '姿勢_件数', 'カバー項目数', '最終報告日'
];

// 活動記録を読み込む（配列）
function tkReadActivity_() {
  var sh = tkSS_().getSheetByName(TK.SHEET_ACT);
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, TK_ACT_HEADER.length).getValues();
  return vals.map(function (r) {
    return {
      ts: r[0], mid: tkMemberId_(r[2]), name: String(r[3] || ''), dept: String(r[4] || ''),
      axis: String(r[5] || ''), title: String(r[6] || ''), level: String(r[7] || ''),
      exp: Number(r[8]) || 0, conf: String(r[9] || ''), practice: String(r[10] || ''), code: String(r[11] || '')
    };
  });
}

// OKR集計シートを再構築
function tkRebuildOkr_() {
  var acts = tkReadActivity_();
  var per = {};
  acts.forEach(function (a) {
    if (!a.mid) return;
    if (!per[a.mid]) per[a.mid] = {
      name: a.name, dept: a.dept, exp: 0, reports: 0,
      axis: { '成果': 0, '能力': 0, '姿勢': 0 }, codes: {}, last: null
    };
    var p = per[a.mid];
    if (a.name) p.name = a.name;
    if (a.dept) p.dept = a.dept;
    p.exp += a.exp; p.reports += 1;
    if (p.axis[a.axis] != null) p.axis[a.axis]++;
    if (a.code) p.codes[a.code] = true;
    var t = a.ts ? new Date(a.ts) : null;
    if (t && (!p.last || t > p.last)) p.last = t;
  });

  var sh = tkEnsureSheet_(TK.SHEET_OKR, TK_OKR_HEADER);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, TK_OKR_HEADER.length).clearContent();
  var rows = Object.keys(per).map(function (mid) {
    var p = per[mid];
    return [mid, p.name, p.dept, p.exp, p.reports,
      p.axis['成果'], p.axis['能力'], p.axis['姿勢'],
      Object.keys(p.codes).length, p.last || ''];
  });
  rows.sort(function (a, b) { return b[3] - a[3]; });   // EXP降順
  if (rows.length) sh.getRange(2, 1, rows.length, TK_OKR_HEADER.length).setValues(rows);
  return { ok: true, members: rows.length };
}

// 本人の到達マンダラ用データ（観点×項目の件数）を作る
function tkOkrMandalaData_(mid) {
  mid = tkMemberId_(mid);
  var acts = tkReadActivity_().filter(function (a) { return a.mid === mid; });
  var name = '', dept = '', totalExp = 0;
  var cnt = {};   // code -> 件数
  acts.forEach(function (a) {
    if (a.name) name = a.name; if (a.dept) dept = a.dept; totalExp += a.exp;
    if (a.code) cnt[a.code] = (cnt[a.code] || 0) + 1;
  });
  // 対象部署の基準（一般＋ベテラン）を観点別に整理
  var crit = (typeof TK_CRITERIA !== 'undefined') ? TK_CRITERIA : [];
  var byAxis = { '成果': [], '能力': [], '姿勢': [] };
  crit.forEach(function (c) {
    if (dept && c.dept !== dept) return;
    if (byAxis[c.axis]) byAxis[c.axis].push({ code: c.code, level: c.level, title: c.title, n: cnt[c.code] || 0 });
  });
  return { mid: mid, name: name, dept: dept, totalExp: totalExp, reports: acts.length, byAxis: byAxis };
}

// 到達マンダラのHTMLを返す（Web表示 ?view=okr&member=）
// 全員の一覧（Web表示 ?view=okr だけで開いたとき）＝これ1つのURLを社内に配れる
function tkOkrIndexHtml_(key) {
  var acts = tkReadActivity_();
  var TOTAL = 96;   // サンクスUP評価基準の項目数
  var per = {}, teamCodes = {};
  acts.forEach(function (a) {
    if (!a.mid) return;
    if (!per[a.mid]) per[a.mid] = { name: a.name, dept: a.dept, exp: 0, reports: 0, codes: {}, last: null };
    var p = per[a.mid];
    if (a.name) p.name = a.name;
    if (a.dept) p.dept = a.dept;
    p.exp += a.exp; p.reports += 1;
    if (a.code) { p.codes[a.code] = true; teamCodes[a.code] = true; }
    var t = a.ts ? new Date(a.ts) : null;
    if (t && (!p.last || t > p.last)) p.last = t;
  });

  var list = Object.keys(per).map(function (mid) {
    var p = per[mid];
    return { mid: mid, name: p.name || mid.split('@')[0], dept: p.dept || '',
      exp: p.exp, reports: p.reports, covered: Object.keys(p.codes).length, last: p.last };
  }).sort(function (a, b) { return b.exp - a.exp; });

  var fmt = function (d) {
    if (!d) return '';
    try { return Utilities.formatDate(d, TK.TZ, 'M/d'); } catch (e) { return ''; }
  };
  var base = ScriptApp.getService().getUrl();
  var rows = list.map(function (m, i) {
    var url = base + '?view=okr&member=' + encodeURIComponent(m.mid) + (key ? ('&key=' + encodeURIComponent(key)) : '');
    var pct = Math.round(m.covered / TOTAL * 100);
    return '<tr>' +
      '<td style="padding:8px 6px;color:#888;font-size:12px">' + (i + 1) + '</td>' +
      '<td style="padding:8px 6px"><a href="' + url + '" target="_blank" style="color:#1a73e8;font-weight:700;text-decoration:none">' + tkEsc_(m.name) + '</a>' +
        (m.dept ? '<div style="font-size:11px;color:#888">' + tkEsc_(m.dept) + '</div>' : '') + '</td>' +
      '<td style="padding:8px 6px;text-align:right;font-weight:700">' + m.exp + '</td>' +
      '<td style="padding:8px 6px;text-align:right">' + m.reports + '</td>' +
      '<td style="padding:8px 6px">' +
        '<div style="display:flex;align-items:center;gap:6px">' +
          '<div style="flex:1;min-width:70px;background:#eceff1;border-radius:6px;height:8px;overflow:hidden">' +
            '<div style="width:' + pct + '%;height:100%;background:linear-gradient(90deg,#26A69A,#43A047)"></div></div>' +
          '<span style="font-size:11px;color:#555;white-space:nowrap">' + m.covered + '/' + TOTAL + '</span>' +
        '</div></td>' +
      '<td style="padding:8px 6px;font-size:12px;color:#777;white-space:nowrap">' + fmt(m.last) + '</td>' +
      '</tr>';
  }).join('');

  var html =
    '<div style="font-family:\'Hiragino Sans\',\'Yu Gothic\',sans-serif;max-width:820px;margin:0 auto;padding:20px">' +
    '<div style="background:linear-gradient(135deg,#37474F,#546E7A);color:#fff;border-radius:12px;padding:16px 20px">' +
      '<div style="font-size:20px;font-weight:bold">' + tkEsc_(TK.COMPANY) + '　サンクスUP！到達マンダラ</div>' +
      '<div style="font-size:13px;opacity:.9">メンバー ' + list.length + '名｜日報 ' + acts.length + '件｜' +
        'チーム到達 ' + Object.keys(teamCodes).length + '/' + TOTAL + '</div>' +
    '</div>' +
    '<p style="font-size:12.5px;color:#666;margin:10px 0 4px">日報が自動でサンクスUP！評価基準（96項目）に反映されています。' +
      '<b>お名前をクリック</b>すると、その方の到達マンダラが開きます。</p>' +
    (list.length ?
      ('<table style="width:100%;border-collapse:collapse;font-size:13.5px">' +
        '<tr style="border-bottom:2px solid #ddd;font-size:11.5px;color:#666;text-align:left">' +
          '<th style="padding:6px"></th><th style="padding:6px">お名前</th>' +
          '<th style="padding:6px;text-align:right">EXP</th><th style="padding:6px;text-align:right">日報</th>' +
          '<th style="padding:6px">到達</th><th style="padding:6px">最終</th></tr>' +
        rows + '</table>')
      : '<p style="padding:20px;color:#888">まだ活動記録がありません。日報が投稿されるとここに並びます。</p>') +
    '</div>';
  return HtmlService.createHtmlOutput(html).setTitle(TK.COMPANY + ' 到達マンダラ');
}

function tkOkrMandalaHtml_(mid) {
  var d = tkOkrMandalaData_(mid);
  if (!d.name && !d.reports) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:24px">該当メンバーの活動記録がありません（' + tkEsc_(mid) + '）</p>');
  }

  // クエスト表から GC分類（直接の成果／価値への取り組み／人材育成）を引く
  var gcByCode = {};
  try {
    tkLoadQuests_().forEach(function (q) { gcByCode[q.code] = q.gc || ''; });
  } catch (e) {}
  function gcOf(it) {
    return gcByCode[it.code] ||
      (it.axis === '成果' ? '直接の成果' : it.axis === '姿勢' ? '価値への取り組み' : '人材育成');
  }

  var all = [], maxN = 1, covered = 0, total = 0;
  TK.AXES.forEach(function (ax) {
    (d.byAxis[ax] || []).forEach(function (it) {
      it.axis = ax; all.push(it); total++;
      if (it.n > 0) { covered++; if (it.n > maxN) maxN = it.n; }
    });
  });

  // 実践件数 → 緑の薄い背景（濃くしすぎない＝文字を必ず読めるように）
  function bg(n) {
    if (!n) return '';
    var a = 0.10 + 0.30 * Math.min(1, n / Math.max(3, maxN * 0.5));
    return ' style="background:rgba(52,168,83,' + (Math.round(a * 100) / 100) + ')"';
  }
  function cell(it) {
    var n = it.n || 0;
    return '<div class="c"' + bg(n) + '>' +
      '<span class="n">' + tkEsc_(it.code || '') + '　' + tkEsc_(it.level || '') + '</span>' +
      '<span class="tx">' + tkEsc_(it.title) + '</span>' +
      (n ? '<span class="met">⭐×' + n + '</span>' : '') + '</div>';
  }

  // ① GCシートの3分類で見る（週次GCにそのまま使える並び）
  var GC3 = ['直接の成果', '価値への取り組み', '人材育成'];
  var gcSections = GC3.map(function (g) {
    var items = all.filter(function (it) { return gcOf(it) === g; });
    var hit = items.filter(function (it) { return it.n > 0; }).length;
    return '<div class="sec"><h3>' + g + ' <small>' + hit + '/' + items.length + '</small></h3>' +
      '<div class="grid">' + items.map(cell).join('') + '</div></div>';
  }).join('');

  // ② サンクスUP！の観点で見る
  var axSections = TK.AXES.map(function (ax) {
    var items = (d.byAxis[ax] || []);
    var hit = items.filter(function (it) { return it.n > 0; }).length;
    return '<div class="sec"><h3>' + ax + ' <small>' + hit + '/' + items.length + '</small></h3>' +
      '<div class="grid">' + items.map(cell).join('') + '</div></div>';
  }).join('');

  var css = [
    'body{margin:0;background:#f6f7f8}',
    '*{box-sizing:border-box}',
    '.wrap{font-family:"Hiragino Sans","Yu Gothic","Noto Sans JP",sans-serif;max-width:1180px;margin:0 auto;padding:16px}',
    '.hd{background:#1a1a2e;color:#fff;border-radius:12px;padding:16px 20px}',
    '.hd .t{font-size:21px;font-weight:bold}',
    '.hd .s{font-size:13.5px;opacity:.92;margin-top:3px}.hd .s b{color:#e8c96b}',
    '.note{font-size:13px;color:#5f6368;line-height:1.8;margin:10px 0}',
    '.tabs{display:flex;gap:8px;margin:14px 0 6px}',
    '.tabs button{border:1.5px solid #dadce0;background:#fff;color:#3c4043;border-radius:100px;padding:7px 20px;font-size:13.5px;font-weight:bold;cursor:pointer}',
    '.tabs button.on{background:#1a73e8;border-color:#1a73e8;color:#fff}',
    '.sec{background:#fff;border-radius:10px;padding:12px 14px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,.08)}',
    '.sec h3{margin:0 0 10px;font-size:15px;color:#5f4300;background:#fef7e0;display:inline-block;padding:4px 12px;border-radius:6px}',
    '.sec h3 small{color:#b06000;font-weight:normal;margin-left:6px}',
    '.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:8px}',
    '@media(max-width:820px){.grid{grid-template-columns:repeat(2,1fr)}}',
    '.c{border:1px solid #e8eaed;border-radius:8px;padding:8px 9px;min-height:92px;display:flex;flex-direction:column;gap:3px;background:#fff}',
    '.c .n{font-size:10px;font-weight:bold;color:#9aa0a6}',
    '.c .tx{font-size:13px;line-height:1.45;color:#202124;font-weight:500;flex:1}',
    '.c .met{font-size:11.5px;font-weight:bold;color:#0b8043}',
    '.lg{font-size:12.5px;color:#5f6368;margin-top:6px;line-height:2}',
    '.sw{display:inline-block;width:13px;height:13px;border:1px solid #dadce0;vertical-align:-2px;margin:0 3px 0 10px;border-radius:3px}'
  ].join('');

  var js = "function tkTab(w){document.getElementById('vGc').style.display=(w==='gc')?'block':'none';" +
    "document.getElementById('vAx').style.display=(w==='ax')?'block':'none';" +
    "document.getElementById('bGc').className=(w==='gc')?'on':'';" +
    "document.getElementById('bAx').className=(w==='ax')?'on':'';}";

  var html =
    '<style>' + css + '</style>' +
    '<div class="wrap">' +
    '<div class="hd"><div class="t">' + tkEsc_(d.name || d.mid) + ' さんの到達マンダラ</div>' +
    '<div class="s">' + (d.dept ? tkEsc_(d.dept) + '　' : '') +
      '累計EXP <b>' + d.totalExp + '</b>　日報 <b>' + d.reports + '</b>件　到達 <b>' + covered + '/' + total + '</b></div></div>' +
    '<p class="note">日報から自動で「サンクスUP！評価基準」に反映されています。<br>' +
      '<b>「GCシートの3分類」タブ</b>は、週次GCシートの「外に生み出した成果」にそのまま転記できる並びです。</p>' +
    '<div class="tabs">' +
      '<button id="bGc" class="on" onclick="tkTab(\'gc\')">GCシートの3分類</button>' +
      '<button id="bAx" onclick="tkTab(\'ax\')">サンクスUP！の観点</button>' +
    '</div>' +
    '<div id="vGc">' + gcSections + '</div>' +
    '<div id="vAx" style="display:none">' + axSections + '</div>' +
    '<p class="lg">⭐＝その基準に該当した日報の件数。' +
      '<span class="sw" style="background:rgba(52,168,83,.12)"></span>薄い緑＝実践あり' +
      '<span class="sw" style="background:rgba(52,168,83,.40)"></span>濃い緑＝何度も実践' +
      '<span class="sw" style="background:#fff"></span>白＝まだ記録がない項目（伸びしろ）</p>' +
    '</div><script>' + js + '</script>';
  return HtmlService.createHtmlOutput(html)
    .setTitle('到達マンダラ｜' + (d.name || ''))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
