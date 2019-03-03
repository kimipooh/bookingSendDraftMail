// Base code Develoer by kijtra (https://kijtra.com/article/gmail-delay-send-by-google-apps-script/)
// custimized by Kimiya Kitani (@kimipooh) 

// 環境設定
var SEARCH_MAX = 10;  // 予約最大件数 = 一度に処理する件数（少ないほうが速くて制限を受けにくい、最大500）
var SEARCH_TERM = "in:draft is:starred";  // 検索条件：下書きかつスター付き

function bookingSendDraftMail2() {
  // GmailApp.search はスレッドを取得するため、下書きメール（下書きに入った下書きを持つスレッドがヒットする）以外もヒットしてしまう。
  // そのため、GmailApp.getMessagesForThreads で各メールを取得する必要がある。
  var myThreads = GmailApp.search(SEARCH_TERM, 0, SEARCH_MAX); //条件にマッチしたスレッドを検索して取得 / 最大500 
  var drafts = GmailApp.getMessagesForThreads(myThreads); //スレッドからメールを取得し二次元配列で格納
  
  //下書きがなければ終了
  if (!drafts.length) {
    return false;
  }

  //現在時刻
  var now = (new Date()).getTime();

  for(var i in drafts){
    for(var j in drafts[i]){
    //メールデータをセット
      var mes = drafts[i][j];
      
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
      for (var k = 0, l = aliases.length; k < l; k++) {
        var val = aliases[k];
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
}