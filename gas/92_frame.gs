// ============================================================
// gas_takamura ／ [92] 新しいGCシートのフレームを作る
// ------------------------------------------------------------
//  9/9の打合せで決まったこと（加藤さま・高村社長）に合わせて、
//  下書き用テンプレを**一から作り直す**。顧客原本には一切触らない。
//
//   ・発表の順番どおりに、左から右へ読める並びにする
//     ① 今週の中心課題 → ② 何をどのように実践したか → ③ 外に生み出した成果
//     → ④ 鍵となる発見・気づき → ⑤ すぐに取り組むべきこと
//     （気づきと次の一手を右の列へ移した。以前は左下だった）
//   ・発表の順番を丸数字と ▶ で示す
//   ・文字を大きくする（本文11pt。以前は8pt）
//   ・箱の大きさを先に決めてから文字を入れるので、文字が重ならない
//   ・A4用紙1枚（横）で印刷できる
//   ・フィードバックコメントとKey Takeawayは小さく、下の帯に寄せる
//   ・問いかけの見出しは青文字
//
//  ★寸法・色・文字の大きさは下の TK_FRAME だけを直せば変わる。
//    スライドは 720 x 405 pt（16:9）。A4横に印刷するとちょうど収まる。
//
//  使い方（松山）:
//    1. tkGCスライド_新フレームを作る()   … 新しいテンプレを作り、そのIDを使う設定にする
//    2. tkGCスライド_今すぐ作る（⑫）で1名分を出して目で見る
//    3. 直したいところがあれば TK_FRAME を変えて 1 をやり直す（古いテンプレは残る）
//
//  ⚠ 元のテンプレに戻すときは tkGCスライド_フレームを元に戻す() を実行する。
// ============================================================

var TK_FRAME = {
  W: 720, H: 405,                 // スライドの大きさ（pt）
  PAD: 14,                        // 外側の余白
  COL_GAP: 10,
  FONT_BODY: 11,                  // 回答欄の文字
  FONT_LABEL: 9.5,                // 問いかけの見出し
  FONT_NAME: 20,                  // 氏名
  FONT_META: 10,                  // 対象週・日報件数
  FONT_SMALL: 8.5,                // 下の帯（フィードバック等）
  COLOR_LABEL: '#1A56A8',         // 見出しの青
  COLOR_BODY: '#1B2E1B',
  COLOR_BOX_BG: '#FFFFFF',
  COLOR_BOX_LINE: '#BFCEDD',
  COLOR_HEAD_BG: '#E8F2E8',       // ヘッダー帯（うすい緑）
  COLOR_OWN_BG: '#FBF7EC',        // ご本人の記入欄（うすいクリーム）
  COLOR_ARROW: '#7A96B5',
  NAME: '_GC下書きテンプレ（新フレーム）'
};

// テンプレのID。スクリプトプロパティが優先（新フレームに切り替えてもconfigを触らずに戻せる）
var TK_FRAME_PROP = 'GC_TEMPLATE_ID';
function tkGCSTemplateId_() {
  var v = '';
  try { v = PropertiesService.getScriptProperties().getProperty(TK_FRAME_PROP) || ''; } catch (e) {}
  return String(v || (TK.GC_SLIDE && TK.GC_SLIDE.TEMPLATE_ID) || '').trim();
}

// ---------- 部品 ----------
//  問いかけの見出し（青文字・太字）
function tkFrLabel_(slide, no, text, x, y, w, h) {
  var t = (no ? no + ' ' : '') + text;
  var sh = slide.insertTextBox(t || ' ', x, y, w, h);
  try {
    var st = sh.getText().getTextStyle();
    st.setFontSize(TK_FRAME.FONT_LABEL).setBold(true).setForegroundColor(TK_FRAME.COLOR_LABEL);
  } catch (e) {}
  try { sh.getText().getParagraphStyle().setLineSpacing(100); } catch (e) {}
  try { sh.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM); } catch (e) {}
  return sh;
}

//  回答欄（枠つきの箱）。中身がはみ出したら自動で縮む
function tkFrBox_(slide, text, x, y, w, h, opt) {
  opt = opt || {};
  // ⚠ 空文字の箱は getText().getTextStyle() が
  //    "The object has no text." で落ちる（2026-09-17 に踏んだ）。
  //    ご記入欄は空のまま使うので、半角空白を 1 つ入れて逆に見せない。
  var body = String(text == null ? '' : text);
  var sh = slide.insertTextBox(body || ' ', x, y, w, h);
  try { sh.getFill().setSolidFill(opt.bg || TK_FRAME.COLOR_BOX_BG); } catch (e) {}
  try {
    sh.getBorder().setWeight(1).getLineFill().setSolidFill(opt.line || TK_FRAME.COLOR_BOX_LINE);
  } catch (e) {}
  try {
    var st = sh.getText().getTextStyle();
    st.setFontSize(opt.font || TK_FRAME.FONT_BODY).setForegroundColor(TK_FRAME.COLOR_BODY);
    if (opt.bold) st.setBold(true);
  } catch (e) {}
  try { sh.getText().getParagraphStyle().setLineSpacing(105); } catch (e) {}
  try { sh.setContentAlignment(SlidesApp.ContentAlignment.TOP); } catch (e) {}
  // ★これが「文字が重ならない」の要。箱の大きさを先に決め、あふれたら文字が縮む
  try { sh.getAutofit().setAutofitType(SlidesApp.AutofitType.SHRINK_ON_OVERFLOW); } catch (e) {}
  return sh;
}

//  発表の順番を示す ▶
function tkFrArrow_(slide, x, y) {
  var sh = slide.insertTextBox('▶', x, y, 14, 18);
  sh.getText().getTextStyle().setFontSize(12).setBold(true).setForegroundColor(TK_FRAME.COLOR_ARROW);
  try { sh.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) {}
  return sh;
}

// ============================================================
//  新しいフレームを作る
// ============================================================
function tkGCスライド_新フレームを作る() {
  var F = TK_FRAME;
  var work = tkGCSWorkFolder_();

  var pres = SlidesApp.create(F.NAME + ' ' + Utilities.formatDate(new Date(), TK.TZ, 'yyyyMMdd-HHmm'));
  var slide = pres.getSlides()[0];
  // 置いてあるひな型の枠を消して白紙にする
  slide.getShapes().forEach(function (s) { try { s.remove(); } catch (e) {} });

  var pad = F.PAD;
  var colW = Math.floor((F.W - pad * 2 - F.COL_GAP * 2) / 3);       // 224
  var x1 = pad, x2 = pad + colW + F.COL_GAP, x3 = pad + (colW + F.COL_GAP) * 2;

  // ---------- ヘッダー ----------
  var head = slide.insertTextBox('', pad, 10, F.W - pad * 2, 30);
  try { head.getFill().setSolidFill(F.COLOR_HEAD_BG); } catch (e) {}
  try { head.getBorder().setTransparent(); } catch (e) {}

  var nameBox = slide.insertTextBox('{{NAME}}さん（{{DEPT}}）', pad + 8, 12, 360, 26);
  nameBox.getText().getTextStyle().setFontSize(F.FONT_NAME).setBold(true)
    .setForegroundColor(F.COLOR_BODY);
  try { nameBox.getBorder().setTransparent(); } catch (e) {}
  try { nameBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) {}

  var metaBox = slide.insertTextBox('週次グループコーチングシート　{{WEEK}}　日報{{COUNT}}',
    pad + 372, 12, F.W - pad * 2 - 380, 26);
  metaBox.getText().getTextStyle().setFontSize(F.FONT_META).setForegroundColor('#4E6B4E');
  try { metaBox.getText().getParagraphStyle().setTextEndIndent(0); } catch (e) {}
  try { metaBox.getBorder().setTransparent(); } catch (e) {}
  try { metaBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) {}
  try { metaBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END); } catch (e) {}

  // ---------- 本体（発表の順番に左から右） ----------
  var top = 48;
  //  ⚠ 下に「AIが入れる2つ」＋「ご本人の記入欄6つ（2段）」が入るので 268 で止める。
  //    ここを下げると、いちばん下の枠が用紙の端に当たって崩れる（2026-09-17）。
  var bodyBottom = 268;

  // ① 今週の中心課題（左上）
  tkFrLabel_(slide, '①', '今週の中心課題・取り組むテーマは何か', x1, top, colW, 14);
  tkFrBox_(slide, '{{THEME}}', x1, top + 15, colW, 56);

  // ② 何をどのように実践したか（左下・いちばん量が多い）
  tkFrLabel_(slide, '②', '何をどのように実践したか', x1, top + 79, colW, 14);
  tkFrBox_(slide, '{{PRACTICE}}', x1, top + 94, colW, bodyBottom - (top + 94));

  // ③ 外に生み出した成果（中央・3分類を1箱に小見出しつきで）
  tkFrLabel_(slide, '③', '外に生み出した成果は何か', x2, top, colW, 14);
  tkFrBox_(slide,
    '【直接の成果】\n{{RESULT_DIRECT}}\n\n【価値への取り組み】\n{{RESULT_VALUE}}\n\n【人材育成】\n{{RESULT_PEOPLE}}',
    x2, top + 15, colW, bodyBottom - (top + 15));

  // ④ 鍵となる発見・気づき（右上）★以前は左下にあった
  tkFrLabel_(slide, '④', '鍵となる発見・気づきは何か／予期せぬ成功・失敗は？', x3, top, colW, 22);
  tkFrBox_(slide, '{{INSIGHT}}', x3, top + 23, colW, 98);

  // ⑤ すぐに取り組むべきこと（右下）
  tkFrLabel_(slide, '⑤', 'すぐに取り組むべきことは何か', x3, top + 129, colW, 14);
  tkFrBox_(slide, '{{NEXT}}', x3, top + 144, colW, bodyBottom - (top + 144));

  // 発表の順番の ▶（列のあいだ）
  tkFrArrow_(slide, x1 + colW - 2, top + 130);
  tkFrArrow_(slide, x2 + colW - 2, top + 130);

  // ---------- 下：AIが入れる2つ（うすい白）----------
  //  ⚠ 1〜2行入る。本文11ptだと2行目が箱からあふれて下の見出しに重なるので 8.5pt。
  var s1 = bodyBottom + 8;                     // 276
  var halfW = Math.floor((F.W - pad * 2 - F.COL_GAP) / 2);
  tkFrLabel_(slide, '⑥', '用いたフィロソフィは', x1, s1, halfW, 11);
  tkFrBox_(slide, '{{PHILOSOPHY}}', x1, s1 + 12, halfW, 20, { font: F.FONT_SMALL });

  tkFrLabel_(slide, '', '今週のKey Takeaway', x1 + halfW + F.COL_GAP, s1, halfW, 11);
  tkFrBox_(slide, '{{TAKEAWAY}}', x1 + halfW + F.COL_GAP, s1 + 12, halfW, 20, { font: F.FONT_SMALL });

  // ---------- 下：ご本人の記入欄（クリーム色・意図して空けている）----------
  //  ★元のGCシートにあった6つの欄を、独立した枠のまま残す。
  //    1本の帯にまとめると、何を書く欄なのかが分からなくなる。
  var own = { bg: F.COLOR_OWN_BG, font: F.FONT_SMALL };
  var w4 = Math.floor((F.W - pad * 2 - 6 * 3) / 4);          // 168
  var s2 = s1 + 37;                            // 313
  ['目指していること', '外的な変化は何か', '内的な変化は何か', '効果的な計画は'].forEach(function (t, i) {
    var x = pad + (w4 + 6) * i;
    tkFrLabel_(slide, '', t, x, s2, w4, 12);
    tkFrBox_(slide, '', x, s2 + 13, w4, 26, own);
  });

  var s3 = s2 + 44;                            // 357
  ['私のフィロソフィ目標', '実践に用いたドラッカーのコンセプト', 'フィードバックコメント']
    .forEach(function (t, i) {
      var x = pad + (colW + F.COL_GAP) * i;
      tkFrLabel_(slide, '', t, x, s3, colW, 12);
      tkFrBox_(slide, '', x, s3 + 13, colW, 26, own);
    });

  pres.saveAndClose();

  // 作業フォルダへ移す
  var file = DriveApp.getFileById(pres.getId());
  try { file.moveTo(work); } catch (e) {}

  // これからはこのテンプレを使う（configは触らない＝戻せる）
  var before = tkGCSTemplateId_();
  PropertiesService.getScriptProperties().setProperty(TK_FRAME_PROP, pres.getId());

  var lines = ['=== 新しいフレームを作りました ===',
    '', pres.getUrl(), '',
    '発表の順番：① 中心課題 ▶ ② 実践 ▶ ③ 成果 ▶ ④ 気づき ▶ ⑤ すぐに取り組むこと',
    '本文の文字：' + F.FONT_BODY + 'pt（以前は8pt）。箱からあふれたら自動で縮みます。',
    'A4横1枚で印刷できます。',
    '',
    'これからの下書きは、このフレームで作られます。',
    '（前のテンプレ ' + (before || '（なし）') + ' はそのまま残っています）',
    '',
    '次にすること:',
    ' 1. ⑫ GC下書きを今すぐ作る で1名分だけ出して、目で見る',
    '    （すでに今日の分がある方はスキップされるので、先に ⑫- で今日の分を消すか、',
    '      対象に「ネットワーク」など1つだけ入れてください）',
    ' 2. 直したいところがあれば松山へ。TK_FRAME の数字を変えて作り直します',
    '',
    '元のテンプレに戻すときは tkGCスライド_フレームを元に戻す() を実行してください。'];
  var msg = lines.join('\n');
  tkLog_('GCスライド', '新フレーム作成 ' + pres.getId());
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg.slice(0, 1800)); } catch (e) {}
  return { id: pres.getId(), url: pres.getUrl() };
}

// ============================================================
//  元のテンプレに戻す
// ============================================================
function tkGCスライド_フレームを元に戻す() {
  PropertiesService.getScriptProperties().deleteProperty(TK_FRAME_PROP);
  var msg = ['元のテンプレに戻しました。',
    '', 'これから使うテンプレ: ' + tkGCSTemplateId_(),
    '', '新フレームは消していません。もう一度使うときは tkGCスライド_新フレームを作る() を実行してください。'].join('\n');
  tkLog_('GCスライド', 'テンプレを元に戻した');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return msg;
}

// ============================================================
//  いま使っているテンプレを確かめる
// ============================================================
function tkGCスライド_いまのテンプレ() {
  var id = tkGCSTemplateId_();
  var lines = ['=== いま使っているテンプレ ==='];
  if (!id) lines.push('未設定です。');
  else {
    lines.push('ID: ' + id);
    try {
      var f = DriveApp.getFileById(id);
      lines.push('名前: ' + f.getName());
      lines.push(f.getUrl());
    } catch (e) { lines.push('⚠ 開けません: ' + (e && e.message)); }
    var p = '';
    try { p = PropertiesService.getScriptProperties().getProperty(TK_FRAME_PROP) || ''; } catch (e) {}
    lines.push('出どころ: ' + (p ? 'スクリプトプロパティ（新フレーム）' : 'config.gs'));
  }
  var msg = lines.join('\n'); Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return msg;
}
