// Base code Develoer by kijtra (https://kijtra.com/article/gmail-delay-send-by-google-apps-script/)
// custimized by Kimiya Kitani (@kimipooh) 

function bookingSendDraftMail() {
  var drafts = GmailApp.getDraftMessages();
  var len = drafts.length;

  //下書きがなければ終了
  if (!len) {
    return false;
  }

  //現在時刻
  var now = (new Date()).getTime();

  for (var i = 0, l = len; i < l; i++) {
    //メールデータをセット
    var mes = drafts[i];
    if ('object' !== typeof mes) {
      continue;
    }
    //スターがついていないと無視 by @kimipooh
    if(!mes.isStarred()) {
      continue;
    }
    //件名を取得
    var str = mes.getSubject();
    //件名から日時を抽出
    var match = str.match(/^(\{(\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2} \d{1,2}:\d{1,2})\}) ?(.*)?/);
    //日時が抽出できないなら無視
    if (!match || !match[1]) {
      continue;
    }

    //時間を取得
    var time = (new Date(match[2].replace(/\-/g,'/')+' +09:00')).getTime();
    //時間を取得できない、または未来の時間なら無視
    if(!time || (time && time>now)){
      continue;
    }

    //各情報をセット
    var to = mes.getTo();
    var subject = match[3] || '';
    var body = mes.getPlainBody();
    var options = {}, val;

    //必要な情報がなければ無視
    if (!to || !body) {
      continue;
    }

    // 2014.11.11 From を変更した場合でも対応(ただしFrom名はつけられない模様)
    var from = mes.getFrom();
    var aliases = GmailApp.getAliases();
    for (var i = 0, l = aliases.length; i < l; i++) {
      var val = aliases[i];
      // From エイリアス一覧にマッチすれば From として使用
      if (-1 !== from.indexOf(val)) {
        options['from'] = val;
        break;
      }
    }

    if (val = mes.getCc()) {//Cc
      options['cc'] = val;
    }

    if (val = mes.getBcc()) {//Bcc
      options['bcc'] = val;
    }

    if (val = mes.getBody()) {//HTML本文
      //bodyにdivタグがあればHTMLとみなす
      if ( val.indexOf('<div')!==-1 ) {
        options['htmlBody'] = val;
      }
    }

    //添付ファイル
    if (val = mes.getAttachments()) {
      options['attachments'] = val;
    }

    // 送信！
    var status = GmailApp.sendEmail(to, subject, body, options);

    //送信したら下書きをゴミ箱へ
    if (status) {
      mes.moveToTrash();
    }
  }
}