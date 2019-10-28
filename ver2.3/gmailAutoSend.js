// Global

// createDate Function keyword list
var createDateKeywordsList = Array(
	// Not Edit
	// =============================================================================
	{"keyword": "gmailAutoSend", "type": "add", "day": "0/0/1", "date": "23:00"},
	{"keyword": "MailTemplate", "type": "add", "day": "0/0/1", "date": "09:00"},
	// =============================================================================
	{"keyword": "", "type": "add", "day": "0/0/0", "date": "18:00"}
);

// changeMailTitle Function keyword list
var changeMailTitleKeywordList = Array(
	{"Title": "{tomorrow}テスト", "changeTitle": "{tomorrow}"}
);

// ==========================================================================================================
// System
// Trigger function info
var triggerFunctionList = Array(
	'gmailAutoSend',
	'sendMails'
);

// Name: clearTrigger
// Description: Clear userd trigger function
// Input: None
// Output: None
function clearTrigger() {
	var triggers = ScriptApp.getProjectTriggers();
	for(var i = 0; i < triggers.length; i++) {
		if(triggerFunctionList.indexOf(triggers[i].getHandlerFunction()) != -1){
			ScriptApp.deleteTrigger(triggers[i]);
		}
	}
	return;
}


// Name: resetTrigger
// Description: Reset trigger function
// Input: None
// Output: None
function resetTrigger(keyword) {
	var date = createDate({"keyword": keyword});
	ScriptApp.newTrigger(triggerFunctionList[0]).timeBased().at(date).create();
	return;
}

// Name: createDate
// Description: Create date data
// Input: dict Dictionary
// Output: date Date
function createDate(dict) 
{
	var outputDate = new Date();
	if(dict.keyword) 
	{
		for(var listCount = 0; listCount < createDateKeywordsList.length; listCount++) 
		{
			if(dict.keyword == createDateKeywordsList[listCount].keyword) 
			{
				var dayData = createDateKeywordsList[listCount].day.split('/');
				var timeData = createDateKeywordsList[listCount].date.split(':');
				if(createDateKeywordsList[listCount].type == "set") 
				{
					// type: set
					outputDate.setFullYear(dayData[0]);
					outputDate.setMonth(dayData[1]);
					outputDate.setDate(dayData[2]);
				} 
				else if (createDateKeywordsList[listCount].type == "add") 
				{
					// type: add
					outputDate.setFullYear(outputDate.getFullYear() + Number(dayData[0]));
					outputDate.setMonth(outputDate.getMonth() + Number(dayData[1]));
					outputDate.setDate(outputDate.getDate() + Number(dayData[2]));
				}
			}
		}
	} 
	else 
	{
		var dayData = dict.day.split('/');
		var timeData = dict.date.split(':');
		if(dict.type == "set")
		{
		// type: set
		outputDate.setFullYear(dayData[0]);
		outputDate.setMonth(dayData[1]);
		outputDate.setDate(dayData[2]);
		} 
		else if (dict.type == "add") 
		{
		// type: add
		outputDate.setFullYear(outputDate.getFullYear() + Number(dayData[0]));
		outputDate.setMonth(outputDate.getMonth() + Number(dayData[1]));
		outputDate.setDate(outputDate.getDate() + Number(dayData[2]));
		}
	}

	outputDate.setHours(timeData[0]);
	outputDate.setMinutes(timeData[1]);
	return outputDate;
}

// Spread Sheet
var sheet = SpreadsheetApp.getActiveSheet();

// Name: clearSpreadSheet
// Description: Clear spread sheet
// Input: None
// Output: None
function clearSpreadSheet() {
	sheet.getRange(2, 1, sheet.getLastRow() + 1, 5).clearContent();
	return;
}

// ==========================================================================================================

// Name: GmailAutoSend
// Description: Main
// Input: None
// Output: None
function gmailAutoSend()
{
	// Clear spreadsheet
	clearSpreadSheet();

	// Clear trigger
	clearTrigger();

	// Change Title
	changeMailTitle();

	// Create Send Mail List
	settingDraftsMail();

	// Create Send Mail Trigger
	settingMailTrigger();

	// Recreate Main Trigger
	resetTrigger("gmailAutoSend");
}

// Name: changeMailTitle
// Description: Main
// Input: None
// Output: None
function changeMailTitle() 
{
	var drafts = GmailApp.getDraftMessages();
	if(drafts.length > 0) 
	{
		for(var i = 0; i < drafts.length; i++) 
		{
			if(drafts[i].getTo() != "") 
			{
				for(var j = 0; j < changeMailTitleKeywordList.length; j++) 
				{
					if( drafts[i].getSubject() == changeMailTitleKeywordList[j].Title && changeMailTitleKeywordList[j].changeTitle != '') 
					{
						var title = changeMailTitleKeywordList[j].changeTitle;
						// check special character
						if(title.indexOf('{tomorrow}') != -1) {
							// {tomorrow}
							var date = new Date();
							var year = String(date.getFullYear());
							var month = ("0"+(date.getMonth() + 1)).slice(-2);
							var day = date.getDate() + 1;
							day = ("0"+day).slice(-2);
							var title_date = "[" + year + "/" + month + "/" + day + "]";
							title = replaceAll(title, "{tomorrow}", title_date);
						}
						var mail = GmailApp.getMessageById(drafts[i].getId());
						GmailApp.createDraft(mail.getTo(), title, mail.getBody());
						mail.moveToTrash();
					}
				}
			}
		}
	}
	return;
}

// Name: settingDraftsMail
// Description: Create Send Mail List
// Input: None
// Output: None
function settingDraftsMail() 
{
	var drafts = GmailApp.getDraftMessages();

	if(drafts.length > 0) {
		var rows = [];

		for(var i = 0; i < drafts.length; i++) 
		{
			if(drafts[i].getTo() != "") 
			{
				var list_check = false;
				var date = new Date();
				for(var j = 0; j < createDateKeywordsList.length; j++) 
				{
					if( drafts[i].getSubject() == createDateKeywordsList[j].keyword) {
						list_check = true;
						var date = createDate({"keyword": drafts[i].getSubject()});
						rows.push([drafts[i].getId(), drafts[i].getTo(), drafts[i].getSubject(), date, ""]);
					}
				}
				if(list_check == false)
				{
					var date = createDate({"keyword": "MailTemplate"});
					rows.push([drafts[i].getId(), drafts[i].getTo(), drafts[i].getSubject(), date, ""]);
				}
			}
		}
		if(rows.length != 0) {
			// Write SpreadSheet
			sheet.getRange(2, 1, rows.length, 5).setValues(rows);
		}
	}
	return;
}

// Name: settingMailTrigger
// Description: Create Send Mail List
// Input: None
// Output: None
function settingMailTrigger() 
{
	date = new Date();
	if(date.getDay() <= 4)
	{
		setSchedule();
	}
}

// ==========================================================================================================
// http://ctrlq.org/code/19716-schedule-gmail-emails
// setSchedule(),sendMails()
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