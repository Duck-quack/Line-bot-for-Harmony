var CHANNEL_ACCESS_TOKEN = 'your access token'; 
var line_endpoint = 'https://api.line.me/v2/bot/message/reply';

var manage_sps = SpreadsheetApp.openById('ID of manage spreadsheet');
var manage_sheet = manage_sps.getSheets()[0];
  
var using = manage_sheet.getRange(2, 1);
var title = manage_sheet.getRange(2, 2);
var article = manage_sheet.getRange(2, 3);
var inputing_row = manage_sheet.getRange(2, 5);
var entering_user = manage_sheet.getRange(2, 6);
var version = manage_sheet.getRange(2, 7);
var spreadsheet_id = manage_sheet.getRange(2, 8);
var initial_id = manage_sheet.getRange(2, 9);
var document_ver = manage_sheet.getRange(2, 11);
var input_data;


function create_spreadsheet(version){
  var spreadsheet = SpreadsheetApp.create(version);
  var article_sheet = spreadsheet.getSheets()[0];
  spreadsheet.insertSheet("letter", 1);
  spreadsheet.insertSheet("space", 2);
  var letter_sheet = spreadsheet.getSheetByName("letter");
  var space_sheet = spreadsheet.getSheetByName("space");
  //create article_sheet
  article_sheet.setName("article");
  article_sheet.appendRow(['title', 'article','writer_id','number']);
  article_sheet.setColumnWidth(1, 200);
  article_sheet.setColumnWidth(2, 500);
  article_sheet.setColumnWidth(3, 250);
  article_sheet.setColumnWidth(4, 100);
  
  //create letter_sheet
  letter_sheet.appendRow(["", 'size','font']);
  var letter_cells = letter_sheet.getRange(2,1,4,1);
  var L_words = [["Harmony"],["version"],["title"],["article"]];
  letter_cells.setValues(L_words);
  letter_sheet.setColumnWidth(1, 100);
  letter_sheet.setColumnWidth(2, 50);
  letter_sheet.setColumnWidth(3, 50);
  var initial_letter_a = letter_sheet.getRange(2,2,4,1);
  initial_letter_a.setValues([[60],[15],[20],[10]]);
  var initial_letter_b = letter_sheet.getRange(2,3,4,1);
  initial_letter_b.setValues([["Times New Roman"],["Arial"],["Arial"],["Arial"]]);
  
  //create space_sheet
  space_sheet.appendRow(["", 'size']);
  var space_cells = space_sheet.getRange(2,1,4,1);
  var S_words = [["H-V"],["V-T"],["T-A"],["A-T"]];
  space_cells.setValues(S_words);
  space_sheet.setColumnWidth(1, 100);
  space_sheet.setColumnWidth(2, 50);
  space_sheet.setColumnWidth(3, 50);
  var initial_space = space_sheet.getRange(2,2,4,1);
  initial_space.setValues([[0],[0],[10],[20]]);
  
  var textId = spreadsheet.getId();
  spreadsheet_id.setValue(textId);
  return spreadsheet;
}

function get_spreadsheet(sheet_id) {
  if (sheet_id == null) {
    return create_spreadsheet("initial");
  } else {
    try {
      return SpreadsheetApp.openById(sheet_id);
    } catch(e) {
      try{
        return SpreadsheetApp.openById(initial_id);
      } catch(e){
        return create_spreadsheet("initial");
      }
    }
  }
}

function cancelling(inputing_row,sheet){
  sheet.getRange(inputing_row,1,1,4).clearContent();//記事のセル内容を消去
  manage_sheet.getRange(2,1,1,3).clearContent();
  manage_sheet.getRange(2,6).clearContent();
  return;
}

function reply(bot_message,e){
  var json = JSON.parse(e.postData.contents);
  var reply_token= json.events[0].replyToken;
  if (typeof reply_token === 'undefined') {
    return;
  }
  return UrlFetchApp.fetch(line_endpoint, {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
    },
    'method': 'post',
    'payload': JSON.stringify({
      'replyToken': reply_token,
      'messages': [{
        'type': 'text',
        'text': bot_message ,
      }],
    }),
  });
}




function doPost(e) {
  var json = JSON.parse(e.postData.contents);
  
  var user_id = json.events[0].source.userId;
  var user_message = json.events[0].message.text;
  
  if(spreadsheet_id.getValue() == 0){
    create_spreadsheet("initial");
  }
  
  var sheet_id = spreadsheet_id.getValue();
  var spreadsheet = get_spreadsheet(sheet_id);
  var sheet = spreadsheet.getSheets()[0];
  var cell;
  if(inputing_row.getValue() == 0){
    inputing_row.setValue(sheet.lastRow()+1);
  }
  
  //新しいハーモニー用のスプレッドシート作成
  if(user_message == "ハーモニー作成"){
    if(user_id == "Line User ID of admin"){
      version.setValue(1);
      return reply("バージョン名を入力してください",e);
    }else {
      return reply(user_id,e);
    }
  }
  if(version.getValue() == 1){
    if(user_id == "Line User ID of admin"){
      var sps_url = create_spreadsheet(user_message).getUrl();
      version.setValue(0);
      document_ver.setValue(1);
      inputing_row.setValue(2);
      return reply("作成完了！\n" + sps_url,e);
    }
  }
  
  //寄稿
  if(user_message == "寄稿する" && using.getValue() == 0){
    using.setValue(1);
    entering_user.setValue(user_id);
    
    cell = sheet.getRange(inputing_row.getValue(), 3)
    cell.setValue(user_id);
      
    return reply("タイトルは何ですか？",e);
  }else if(user_message == "寄稿する"){
    return reply("現在使用中です！また後でお試しください！",e);
  }
  
  if(using.getValue() == 1 && user_id == entering_user.getValue()){
    if(user_message == "中止"){
      cancelling(inputing_row.getValue(),sheet);
      return reply("中止しました。\n再度寄稿する場合はもう一度「寄稿する」と言ってください♪",e);
    }
    if(title.getValue() == 0){
      title.setValue(1);
      
      cell = sheet.getRange(inputing_row.getValue(),1);
      cell.setValue(user_message);
      
      reply("記事の内容を書いてください",e);
      return reply(e);
    }else if(article.getValue() == 0){
      cell = sheet.getRange(inputing_row.getValue(),2);
      cell.setValue(user_message);
      
      manage_sheet.getRange(2,1,1,3).clearContent();
      manage_sheet.getRange(2,6).clearContent();
      inputing_row.setValue(inputing_row.getValue() + 1);
      return reply("寄稿完了しました！ありがとうございますー！",e);
    }
  }
  
  if(user_message == "ヘルプ"){
    return reply("寄稿したいときは「寄稿する」\n書いてる途中でやめたくなったら「中止」\nと言ってください♪\nその他バグ、質問、要望などありましたら畑中まで～",e);
  }
  
  else{
    return reply("分からないことがあったら「ヘルプ」って言ってね！",e);
  }
}