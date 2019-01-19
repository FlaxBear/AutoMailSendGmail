// グローバル変数
// day     : 本機能の何日置きに起動するか
// time    : 本機能を何時に起動するか
// minutes : 本機能を何分に起動するか
var program_start_time = {
		// D5:曜日指定にすればよさそう
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
	// D5
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
var sheet = SpreadsheetApp.getActiveSheet();	//	スプレッドシートの設定
var trigger_function_list = Array(				//	使用しているトリガー関数の情報
	'my_auto_send_mail',
	'sendMails'
);

// 関数:clear_spreadsheet
// 説明:シートの初期化
// 入力: なし
// 出力: なし
function clear_spreadsheet() {
	sheet.getRange(2, 1, sheet.getLastRow() + 1, 5).clearContent();
}

// 関数:clear_trigger
// 説明:この機能で使用している使用しているトリガーの削除
// 入力:なし
// 出力:なし
function clear_trigger() {
	// トリガーの削除
	var triggers = ScriptApp.getProjectTriggers();
	for(var i = 0; i < triggers.length; i++) {
		// FE1:indexOfを使用しているので部分一致で削除されないか？
		if(trigger_function_list.indexOf(triggers[i].getHandlerFunction()) != -1){
			ScriptApp.deleteTrigger(triggers[i]);
		}
	}
}

// 関数:set_trigger
// 説明:本機能の再トリガー設定
// 入力:なし
// 出力:なし
function set_trigger()
{
	// D1:共通関数で日付の設定が出来れば
	var date = new Date();
	date.setDate(date.getDate() + program_start_time.day);
	date.setHours(program_start_time.hour);
	date.setMinutes(program_start_time.minutes);
	ScriptApp.newTrigger(trigger_function_list[0]).timeBased().at(date)
	.inTimezone(SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone())
	.create();
}
// ==========================================================================================================

// 関数:my_auto_send_mail
// 説明:本機能
// 入力:なし
// 出力:なし
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

// 関数:setting_drafts_mail
// 説明:下書きメール検索＆自動送信割り当て
// 入力:なし
// 出力:なし
function setting_drafts_mail() {

	// 下書きメールリストの取得
	var drafts = GmailApp.getDraftMessages();
	if(drafts.length > 0) {
		var rows = [];
			for(var i = 0; i < drafts.length; i++) {
				// FE2:テンプレに宛先指定出来たら終わりじゃない？
				if(drafts[i].getTo() != "") {
					// タイトルを元に日付時間割り当て
					var list_check = false;
					for(var j = 0; j < match_title.length; j++) {
						if( drafts[i].getSubject() == match_title[j].title) {
							list_check = true;
							// リストにある場合のみその時間に設定
							// D1
							var date = new Date();
							date.setDate(date.getDate() + match_title[j].day);
							date.setHours(match_title[j].hour);
							date.setMinutes(match_title[j].minutes);
							date.setSeconds(0);
							rows.push([drafts[i].getId(), drafts[i].getTo(), drafts[i].getSubject(), date, ""]);
						}
					}
					if(list_check == false) {
						// それ以外であればテンプレ時間を設定
						// D1
						var date = new Date();
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

// 関数:setting_mail_trigger
// 説明:土日以外のメール送信トリガーの作成
// 入力:なし
// 出力:なし
function setting_mail_trigger() {
	// D2:土日でも設定が行われていないだけで本体は起動しているので金曜日の実行時に次に実行を日曜日に出来ないか
	date = new Date();
	if(date.getDay() <= 4)
	{
		setSchedule();
	}
}

// 関数:change_mail_title
// 説明:下書きのメールの件名を別の件名に変換する関数
// match_title配列のchange_titleに変換する
// ・特殊文字
//  {date} 年月日を入れる(例:20190101)
// 入力:なし
// 出力:なし
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
								// D1
								// D3:例えば{date:20190101}記述し、日時を指定出来ればいいんじゃないかな？
								// D4:{date:EW:1}を記述すると毎週月曜日の日時を指定とか
								var date = new Date();
								var year = String(date.getFullYear());
								var month = ("0"+(date.getMonth() + 1)).slice(-2);
								var day = date.getDate() + 1;
								day = ("0"+day).slice(-2);
								var title_date = year + month + day;
								title = replaceAll(title, "{date}", title_date);
							}
							// D5:特殊文字氏名とかは？

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
										// プレーンテキストの場合、改行が\nになり、改行されないので変換
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