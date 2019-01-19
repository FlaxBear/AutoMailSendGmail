// グローバル変数
// day     : 本機能の何日置きに起動するか
// time    : 本機能を何時に起動するか
// minutes : 本機能を何分に起動するか
var program_start_time = {
    day: 7,
    hour: 23,
    minutes: 0,
  }
  
  // メール送信時間割り振りリスト
  // title        : 割り振るメールの件名(正規表現していないので完全一致じゃないと起動しない)
  // day          : 何日後に送信するか
  // hour         : 何時に送信するか
  // minutes      : 何分に送信するか 
  // change_title : テンプレートなどの件名を変更する際に使用（使用しない場合''でok)
  var match_title = Array(
    {title: 'テンプレ',       day: '1', hour: '8', minutes: '0', change_title: ''},
    {title: 'テンプレ',       day: '1', hour: '8', minutes: '0', change_title: ''},
    {title: 'テンプレ',   day: '1',  hour: '8', minutes: '0', change_title: ''},
    {title: 'テンプレ', day: '1', hour: '8', minutes: '0', change_title: ''}
  );
  
  // メール送信リストに書かれている以外のメールの設定
  var template_send_data = {
    day: 1,
    hour: 8,
    minutes: 0
  };
  
  // 以下、システム
  // ==========================================================================================================
  // 再利用可能みたいなところ
  var sheet = SpreadsheetApp.getActiveSheet();
  var trigger_function_list = Array(
    'my_auto_send_mail',
    'sendMails'
  );
  
  // シートの初期化
  function clear_spreadsheet() {
    sheet.getRange(2, 1, sheet.getLastRow() + 1, 5).clearContent();
  }
  
    
  // この機能で使用している使用しているトリガーの削除
  function clear_trigger() {
    // トリガーの削除
    var triggers = ScriptApp.getProjectTriggers();
    for(var i = 0; i < triggers.length; i++) {
      if(trigger_function_list.indexOf(triggers[i].getHandlerFunction()) != -1){
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
  }
  
  // 機能の再トリガー
  function set_trigger()
  {
    var date = new Date();
    date.setDate(date.getDate() + program_start_time.day);
    date.setHours(program_start_time.hour);
    date.setMinutes(program_start_time.minutes);
    ScriptApp.newTrigger(trigger_function_list[0]).timeBased().at(date)   
    .inTimezone(SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone())
    .create();
  }
  // ==========================================================================================================
  // 本機能
  function my_auto_send_mail() {
    // シートの初期化
    clear_spreadsheet();
    
    // トリガーの削除
    clear_trigger();
    
    // 件名の変換
    change_mail_title();
  
    // 下書きメール検索＆自動送信割り当て
    setting_drafts_mail();
    
    // 土日以外のメール送信トリガーの作成
    setting_mail_trigger();
    
    // 本機能のトリガーの再登録
    set_trigger();
  }
  
  // 下書きメール検索＆自動送信割り当て
  function setting_drafts_mail() {
  
    var drafts = GmailApp.getDraftMessages();
    
    if(drafts.length > 0) {
      var rows = [];
      
       for(var i = 0; i < drafts.length; i++) {
         if(drafts[i].getTo() != "") {
           // タイトルを元に日付時間割り当て
           var list_check = false;
           var date = new Date();
           for(var j = 0; j < match_title.length; j++) {
             if( drafts[i].getSubject() == match_title[j].title) {
               var date = new Date();
               // リストにある場合のみその時間に設定
               list_check = true;
               date.setDate(date.getDate() + match_title[j].day);
               date.setHours(match_title[j].hour);
               date.setMinutes(match_title[j].minutes);
               date.setSeconds(0);
               rows.push([drafts[i].getId(), drafts[i].getTo(), drafts[i].getSubject(), date, ""]);
             }
           }
           if(list_check == false) {
             // それ以外であればテンプレ時間を設定
             date.setDate(date.getDate() + template_send_data.day);
             date.setHours(template_send_data.hour);
             date.setMinutes(template_send_data.minutes);
             date.setSeconds(0);
             rows.push([drafts[i].getId(), drafts[i].getTo(), drafts[i].getSubject(), date, ""]);
           }
         }
       }
      if(rows.length != 0) {
        // スプレットシートに記述
        sheet.getRange(2, 1, rows.length, 5).setValues(rows);
      }
    }
  }
  
  // 土日以外のメール送信トリガーの作成
  function setting_mail_trigger() {
    date = new Date();
    if(date.getDay() <= 4)
    {
      setSchedule();
    }
  }
  
  function change_mail_title() {
    var drafts = GmailApp.getDraftMessages();
    if(drafts.length > 0) {
      for(var i = 0; i < drafts.length; i++) {
        if(drafts[i].getTo() != "") {
          for(var j = 0; j < match_title.length; j++) {
            if( drafts[i].getSubject() == match_title[j].title && match_title[j].change_title != '') {
                var title = match_title[j].change_title;
                // 特殊文字が含まれているか確認
                if(title.indexOf('{date}') != -1) {
                  // 日付(とりあえず翌日にしているがverupで変更)
                  var date = new Date();
                  var year = String(date.getFullYear());
                  var month = ("0"+(date.getMonth() + 1)).slice(-2);
                  var day = date.getDate() + 1;
                  day = ("0"+day).slice(-2);
                  var title_date = year + month + day;
                  title = replaceAll(title, "{date}", title_date);
                }
                // 元の下書きメールを取得し、新たに件名を変更した下書きメール作成し、元のデータは削除
                var mail = GmailApp.getMessageById(drafts[i].getId());
                GmailApp.createDraft(mail.getTo(), title, mail.getBody());
                mail.moveToTrash();
            }
          }
        }
      }
    }
  }
  // 引用プログラム
  // ==========================================================================================================
  // http://ctrlq.org/code/19716-schedule-gmail-emails
  // setSchedule(),sendMails()を使用
  function initialize() {
  
      /* Clear the current sheet */
      var sheet = SpreadsheetApp.getActiveSheet();
      sheet.getRange(2, 1, sheet.getLastRow() + 1, 5).clearContent();
  
      /* Delete all existing triggers */
      var triggers = ScriptApp.getProjectTriggers();
      for (var i = 0; i < triggers.length; i++) {
          if (triggers[i].getHandlerFunction() === "sendMails") {
              ScriptApp.deleteTrigger(triggers[i]);
          }
      }
  
      /* Import Gmail Draft Messages into the Spreadsheet */
      var drafts = GmailApp.getDraftMessages();
      if (drafts.length > 0) {
          var rows = [];
          for (var i = 0; i < drafts.length; i++) {
              if (drafts[i].getTo() !== "") {
                  rows.push([drafts[i].getId(), drafts[i].getTo(), drafts[i].getSubject(), "", ""]);
              }
          }
          sheet.getRange(2, 1, rows.length, 5).setValues(rows);
      }
  }
  function setSchedule() {
      var sheet = SpreadsheetApp.getActiveSheet();
      var data = sheet.getDataRange().getValues();
      var time = new Date().getTime();
      var code = [];
      for (var row in data) {
          if (row != 0) {
              var schedule = data[row][3];
              if (schedule !== "") {
                  if (schedule.getTime() > time) {
                      ScriptApp.newTrigger("sendMails")
                          .timeBased()
                          .at(schedule)
                          .inTimezone(SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone())
                          .create();
                      code.push("Scheduled");
                  } else {
                      code.push("Date is in the past");
                  }
              } else {
                  code.push("Not Scheduled");
              }
          }
      }
      for (var i = 0; i < code.length; i++) {
          sheet.getRange("E" + (i + 2)).setValue(code[i]);
      }
  }
  
  function sendMails() {
      var sheet = SpreadsheetApp.getActiveSheet();
      var data = sheet.getDataRange().getValues();
      var time = new Date().getTime();
      for (var row = 1; row < data.length; row++) {
          if (data[row][4] == "Scheduled") {
              var schedule = data[row][3];
              if ((schedule != "") && (schedule.getTime() <= time)) {
                  var message = GmailApp.getMessageById(data[row][0]);
                  var body = message.getBody();
                  var options = {
                      cc: message.getCc(),
                      bcc: message.getBcc(),
                      //htmlBody: body,
                      htmlBody:replaceAll(body, "\n", "<br>"),
                      replyTo: message.getReplyTo(),
                      attachments: message.getAttachments()
                  }
  
                  /* Send a copy of the draft message and move it to Gmail trash */
                  GmailApp.sendEmail(message.getTo(), message.getSubject(), body, options);
                  message.moveToTrash();
                  sheet.getRange("E" + (row + 1)).setValue("Delivered");
              }
          }
      }
  }
  // ==========================================================================================================
  //  https://javascript.programmer-reference.com/js-function-replaceall/
  function replaceAll(str, beforeStr, afterStr){
    var reg = new RegExp(beforeStr, "g");
    return str.replace(reg, afterStr);
  }