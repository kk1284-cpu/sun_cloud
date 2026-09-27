// ============================================================
// gas_takamura ／ [94] Google Chat へのお知らせ
// ------------------------------------------------------------
//  ねらい：GCシートができたことを、その事業所のスペースへ自動で投稿する。
//         「作られても気づかない」状態をなくす。2026-09-18 石川さまのご要望。
//
//  ★Webhook方式にした理由（Chatアプリにしなかった理由）：
//    ・スペースで「アプリと連携 → Webhook」を1本発行するだけで済む
//    ・Google Workspace の管理者権限が要らない
//    ・弊社アカウントからの承認・ログインが一切要らない（URLにPOSTするだけ）
//    個人へのDMが要るようになったら、そのときにChatアプリ（Chat API）を検討する。
//    ⚠ DMはドメイン全体の委任が必要で、髙村さまの管理者判断になる。
//
//  使い方（初回だけ）：
//    1. tkChat設定シート_作成()      … シート「Chat通知」を作る
//    2. Google Chat の各スペースで
//         スペース名 ▾ → アプリと連携 → Webhookを追加 → 名前を付けてURLをコピー
//    3. シート「Chat通知」に、事業所名とそのURLを貼る
//    4. tkChat_テスト送信()          … 1行ずつ「テストです」を流して届くか見る
//
//  毎週：
//    GCシートの作成が終わったところで、自動的にその事業所のスペースへ投稿される。
//
//  ⚠ 投稿にはお一人ずつのシートのリンクが並ぶ。スペースの参加者全員に見える。
//    グループコーチングは全員でなさるものなので、そういう前提で作っている。
//    見せたくない場合は「一覧を出さない」列を「はい」にすると、
//    フォルダのリンクと人数だけを投稿する。
// ============================================================

var TK_CHAT = {
  SHEET: 'Chat通知',
  HEADER: ['事業所', 'WebhookのURL', '使う', '一覧を出さない', '備考'],
  MAX_PER_POST: 30,        // 1回の投稿に並べる人数の上限（長すぎると読まれない）
  MAX_TEXT: 3800           // Chatの1メッセージの上限は4096文字。余裕をみる
};

// ------------------------------------------------------------
//  設定シート
// ------------------------------------------------------------
function tkChat設定シート_作成() {
  var sh = tkEnsureSheet_(TK_CHAT.SHEET, TK_CHAT.HEADER);
  if (sh.getLastRow() <= 1) {
    // 見えている事業所の分だけ、空の行を用意しておく（URLは松山が貼る）
    var rows = [
      ['薬局', '', 'はい', '', 'スペース「薬局振り返り」'],
      ['ネットワーク', '', 'はい', '', 'スペース「NW事業課_振返り」'],
      ['花のある家（中吉田）', '', 'はい', '', 'スペース「花のある家 振り返り」'],
      ['その他', '', 'はい', '', 'スペース「日報：ボードメンバー」'],
      ['＊', '', '', '', '事業所が見つからないときの送り先（空でよい）']
    ];
    sh.getRange(2, 1, rows.length, TK_CHAT.HEADER.length).setValues(rows);
    sh.setColumnWidth(2, 420);
  }
  var msg = 'シート「' + TK_CHAT.SHEET + '」を用意しました。\n\n'
    + '【WebhookのURLの取り方】\n'
    + '　1. Google Chat で対象のスペースを開く\n'
    + '　2. 画面上のスペース名 ▾ → アプリと連携\n'
    + '　3. 「Webhook を追加」→ 名前（例：GCシートのお知らせ）→ 保存\n'
    + '　4. 出てきたURLをコピーして、この表の「WebhookのURL」に貼る\n\n'
    + '貼り終えたら tkChat_テスト送信() で届くかご確認ください。\n\n'
    + '⚠ URLを知っていれば誰でもそのスペースに投稿できます。外に出さないでください。';
  tkLog_('Chat', '設定シートを用意');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}

//  設定を読む（事業所 → {url, hideList}）
function tkChatHooks_() {
  var out = {}, sh = null;
  try { sh = tkSS_().getSheetByName(TK_CHAT.SHEET); } catch (e) {}
  if (!sh || sh.getLastRow() < 2) return out;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, TK_CHAT.HEADER.length).getValues();
  v.forEach(function (r) {
    var office = String(r[0] || '').trim();
    var url = String(r[1] || '').trim();
    var use = String(r[2] || '').trim();
    if (!office || !url) return;
    if (use && ['いいえ', 'no', 'off', '×'].indexOf(use.toLowerCase()) >= 0) return;
    out[office] = { url: url, hideList: String(r[3] || '').trim() !== '' };
  });
  return out;
}

// ------------------------------------------------------------
//  投稿する
// ------------------------------------------------------------
//  ⚠ 失敗しても例外を投げない。Chatが止まってもGCシートの作成は続けたい。
function tkChatPost_(url, text) {
  if (!url || !text) return false;
  try {
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json; charset=UTF-8',
      payload: JSON.stringify({ text: String(text).slice(0, TK_CHAT.MAX_TEXT) }),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code !== 200) {
      tkLog_('Chat', '投稿できませんでした（HTTP ' + code + '）: ' + res.getContentText().slice(0, 200));
      return false;
    }
    return true;
  } catch (e) {
    tkLog_('Chat', '投稿できませんでした: ' + (e && e.message));
    return false;
  }
}

function tkChat_テスト送信() {
  var hooks = tkChatHooks_();
  var names = Object.keys(hooks);
  var lines = ['=== Chatへのテスト送信 ==='];
  if (!names.length) {
    lines.push('シート「' + TK_CHAT.SHEET + '」にURLが入っていません。');
    lines.push('先に tkChat設定シート_作成() を実行し、Webhookのアドレスを貼ってください。');
  } else {
    names.forEach(function (office) {
      var ok = tkChatPost_(hooks[office].url,
        '【テスト】日報解析システムからの接続確認です。この投稿は消していただいて構いません。');
      lines.push('　' + office + '：' + (ok ? '✅ 届きました' : '🔴 届きませんでした（URLをご確認ください）'));
    });
  }
  var msg = lines.join('\n');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}

// ------------------------------------------------------------
//  GCシートができたことを知らせる
// ------------------------------------------------------------
//  made: [{office, name, url, blank}] ／ label: '2026 9/11〜9/17'
function tkChatNotifyGC_(label, made, folderUrl) {
  var hooks = tkChatHooks_();
  if (!Object.keys(hooks).length || !made || !made.length) return { posted: 0 };

  // 事業所ごとにまとめる
  var byOffice = {};
  made.forEach(function (m) {
    var o = m.office || 'その他';
    (byOffice[o] = byOffice[o] || []).push(m);
  });

  var posted = 0;
  Object.keys(byOffice).forEach(function (office) {
    var hook = hooks[office] || hooks['＊'] || hooks['*'];
    if (!hook) return;
    var list = byOffice[office];
    var withText = list.filter(function (m) { return !m.blank; });

    var t = ['📄 *今週のグループコーチングシートができました*',
             '対象期間：' + label,
             office + '：' + list.length + '名（うち日報から下書きが入った方 ' + withText.length + '名）',
             ''];
    if (hook.hideList) {
      t.push('こちらのフォルダに、お一人ずつのシートが入っています。');
      if (folderUrl) t.push(folderUrl);
    } else {
      list.slice(0, TK_CHAT.MAX_PER_POST).forEach(function (m) {
        t.push('・' + m.name + (m.blank ? '（今週の日報なし・枠のみ）' : '') + '　' + m.url);
      });
      if (list.length > TK_CHAT.MAX_PER_POST) {
        t.push('・ほか ' + (list.length - TK_CHAT.MAX_PER_POST) + '名');
        if (folderUrl) t.push(folderUrl);
      }
    }
    t.push('');
    t.push('文章の欄はAIの下書きです。ご確認のうえ、手を入れてお使いください。');
    if (list.length !== withText.length) {
      t.push('日報が無い週は枠だけが入ります。日報を書いていただくと、次回から下書きが入ります。');
    }
    if (tkChatPost_(hook.url, t.join('\n'))) posted++;
  });
  if (posted) tkLog_('Chat', 'GCシートのお知らせを ' + posted + '件のスペースへ投稿しました（' + label + '）');
  return { posted: posted };
}
