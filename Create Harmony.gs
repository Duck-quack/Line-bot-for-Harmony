var CHANNEL_ACCESS_TOKEN = 'your access token'; 
var line_endpoint = 'https://api.line.me/v2/bot/message/reply';

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

function doPost(e){
  var json = JSON.parse(e.postData.contents);
  
  var user_id = json.events[0].source.userId;
  var user_message = json.events[0].message.text;
  
  //open manege spread sheet
  var manage_sps = SpreadsheetApp.openById('ID of manage spreadsheet');
  var manage_sheet = manage_sps.getSheets()[0];
  //get ID of up-to-date version "Harmony" spreadsheet
  var spreadsheet_id = manage_sheet.getRange(2, 8);
  var sheet_id = spreadsheet_id.getValue();
  //open "Harmony" spread sheet 
  var spreadsheet = SpreadsheetApp.openById(sheet_id);
  var article_sheet = spreadsheet.getSheets()[0];
  var letter_sheet = spreadsheet.getSheets()[1];
  var space_sheet = spreadsheet.getSheets()[2];
  var version_name = spreadsheet.getName();
  
  if(user_id == "U82fa048b16511463e19e9b6e9436ad02" && user_message == "作成中"){
    return reply("作成中のスプレッドシートだよ！" + spreadsheet.getUrl(),e);
  }
  
  if(user_id == "U82fa048b16511463e19e9b6e9436ad02" && user_message == "実行"){
    //get document
    var document_ver = manage_sheet.getRange(2, 11);
    var document_id = manage_sheet.getRange(2, 10);
    var document = DocumentApp.create(version_name + " Ver " + document_ver.getValue() + ".0");
  
    /*function make_space(size){
    var space = document.appendParagraph(" ").setFontSize(size);
    }*/
  
    /*parameters
    var harmony_size = letter_sheet.getRange(2, 2);
    var harmony_font = letter_sheet.getRange(2, 3);
    var version_size = letter_sheet.getRange(3, 2);
    var version_font = letter_sheet.getRange(3, 3);
    var title_size = letter_sheet.getRange(4, 2);
    var title_font = letter_sheet.getRange(4, 3);
    var article_size = letter_sheet.getRange(5, 2);
    var article_font = letter_sheet.getRange(5, 3);

    var H-V_space_size = space_sheet.getRange(, );
    var H-V_space_amount = space_sheet.getRange(, );
    var V-T_space_size = space_sheet.getRange(, );
    var V-T_space_amount = space_sheet.getRange(, );
    var T-A_space_size = space_sheet.getRange(, );
    var T-A_space_amount = space_sheet.getRange(, );
    var A-T_space_size = space_sheet.getRange(, );
    var A-T_space_amount = space_sheet.getRange(, );  */
  
    var hs = letter_sheet.getRange(2, 2).getValue();
    var hf = letter_sheet.getRange(2, 3).getValue();
    var vs = letter_sheet.getRange(3, 2).getValue();
    var vf = letter_sheet.getRange(3, 3).getValue();
    var ts = letter_sheet.getRange(4, 2).getValue();
    var tf = letter_sheet.getRange(4, 3).getValue();
    var as = letter_sheet.getRange(5, 2).getValue();
    var af = letter_sheet.getRange(5, 3).getValue();
  
    var HVs = space_sheet.getRange(2, 2).getValue();
    var VTs = space_sheet.getRange(3, 2).getValue();
    var TAs = space_sheet.getRange(4, 2).getValue();
    var ATs = space_sheet.getRange(5, 2).getValue();
  
    var space;
  
    //sort
    var lastrow = article_sheet.getLastRow();
    var range = article_sheet.getRange(2,1,lastrow-1,4);
    range.sort([{column: 4,ascending: true}]);
  
    //write "Harmony"
    var harmony = document.insertParagraph(0,"Harmony");
    harmony.setFontSize(hs);
    harmony.setFontFamily(hf);
    harmony.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    space = document.appendParagraph(" ").setFontSize(HVs);
  
    //write version
    var version = document.appendParagraph(version_name);
    version.setFontSize(vs);
    version.setFontFamily(vf);
    version.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    space = document.appendParagraph(" ").setFontSize(VTs);
  
    //write articles
    for(var article_row = 2; article_row <= lastrow; article_row=article_row +1){
      //title
      var title = document.appendParagraph(article_sheet.getRange(article_row, 1).getValue());
      title.setFontSize(ts);
      title.setFontFamily(tf);
      title.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      space = document.appendParagraph(" ").setFontSize(TAs);
      //article
      var article = document.appendParagraph(article_sheet.getRange(article_row, 2).getValue());
      article.setFontSize(as);
      article.setFontFamily(af);
      article.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      space = document.appendParagraph(" ").setFontSize(ATs);
    }
    
    document_ver.setValue(document_ver.getValue() + 1);
    var document_url = document.getUrl();
    return reply("実行しました！" + document_url,e);
  }else{
    return reply("実行する権限がない、あるいはメッセージが間違っています。\n「実行」か「作成中」と言ってみてください！",e);
  }
}
