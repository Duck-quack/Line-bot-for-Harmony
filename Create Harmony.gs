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
  sheet.getRange(inputing_row,1,1,4).clearContent();//�L���̃Z�����e������
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
  
  //�V�����n�[���j�[�p�̃X�v���b�h�V�[�g�쐬
  if(user_message == "�n�[���j�[�쐬"){
    if(user_id == "Line User ID of admin"){
      version.setValue(1);
      return reply("�o�[�W����������͂��Ă�������",e);
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
      return reply("�쐬�����I\n" + sps_url,e);
    }
  }
  
  //��e
  if(user_message == "��e����" && using.getValue() == 0){
    using.setValue(1);
    entering_user.setValue(user_id);
    
    cell = sheet.getRange(inputing_row.getValue(), 3)
    cell.setValue(user_id);
      
    return reply("�^�C�g���͉��ł����H",e);
  }else if(user_message == "��e����"){
    return reply("���ݎg�p���ł��I�܂���ł��������������I",e);
  }
  
  if(using.getValue() == 1 && user_id == entering_user.getValue()){
    if(user_message == "���~"){
      cancelling(inputing_row.getValue(),sheet);
      return reply("���~���܂����B\n�ēx��e����ꍇ�͂�����x�u��e����v�ƌ����Ă���������",e);
    }
    if(title.getValue() == 0){
      title.setValue(1);
      
      cell = sheet.getRange(inputing_row.getValue(),1);
      cell.setValue(user_message);
      
      reply("�L���̓��e�������Ă�������",e);
      return reply(e);
    }else if(article.getValue() == 0){
      cell = sheet.getRange(inputing_row.getValue(),2);
      cell.setValue(user_message);
      
      manage_sheet.getRange(2,1,1,3).clearContent();
      manage_sheet.getRange(2,6).clearContent();
      inputing_row.setValue(inputing_row.getValue() + 1);
      return reply("��e�������܂����I���肪�Ƃ��������܂��[�I",e);
    }
  }
  
  if(user_message == "�w���v"){
    return reply("��e�������Ƃ��́u��e����v\n�����Ă�r���ł�߂����Ȃ�����u���~�v\n�ƌ����Ă���������\n���̑��o�O�A����A�v�]�Ȃǂ���܂����用���܂Ł`",e);
  }
  
  else{
    return reply("������Ȃ����Ƃ���������u�w���v�v���Č����ĂˁI",e);
  }
}