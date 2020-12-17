const SHEET_ID = "17vR1PKzyyaqF4RALmI9MuCw0fUlAnjetmcQXCKXsPl0";
const SHEET_NAME = "宛先一覧";
const template ='${管理番号} \n${バージョン}をリリースします。\n以下の管理番号を含みます。${管理番号}'; 
const placeHolders = ["${バージョン}", "${管理番号}"];
let paramMap = {};
const titleRegExp = new RegExp('\\$\\{(.*)\\}');

function onGmailCompose(e) {
  console.log(e);
  
  const header = CardService.newCardHeader()
      .setTitle('パラメータ入力')
      .setSubtitle('メール本文のパラメータ部分を入力してください');
  
  const dropdown = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName("dropdown_field");
  
//  var item = getValuesFromSpreadsheet();
//  
//  for(var i = 0; i < item.length; i++){
//    dropdown.addItem(item[i], item[i], false);
//  }

  const action = CardService.newAction()
      .setFunctionName('onGmailComposeBody');
  const footer = CardService.newFixedFooter()
      .setPrimaryButton(CardService.newTextButton()
      .setText('本文作成！')
      .setOnClickAction(action));

  const section = CardService.newCardSection();
  
  const checkbox = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName('is_plain_text');
  checkbox.addItem('テキストメール', true, false);
  
  section.addWidget(checkbox);
  
  for(let i = 0; i < placeHolders.length; i++) {
    paramMap["params_" + i] = placeHolders[i];
    let title = placeHolders[i].match(titleRegExp)[1];
    
//    console.log(placeHolders[i]);
//    console.log(title);
    
    let textInput = CardService.newTextInput().setFieldName("params_" + i).setTitle(title);
    section.addWidget(textInput);
  }

  const card = CardService.newCardBuilder()
      .setHeader(header)
      .addSection(section)
      .setFixedFooter(footer);
  return card.build();
}

function onGmailComposeBody(e) {
  console.log(e);
  
  let text = template;
  for(let i = 0; i < placeHolders.length; i++) {
    if(e.formInput[('params_' + i)]) {
      var textRegExp = new RegExp(placeHolders[i].replace('$', '\\$').replace('{', '\\{').replace('}', '\\}'), 'g');
      
//      console.log('textRegExp:' + textRegExp);
      
      text = text.replace(textRegExp, e.formInput[('params_' + i)]);
    }
  }
  
//  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
//  
//  var dropdownValue = e.formInput.dropdown_field;
//  
//  //指定セルを起点に最終列のカラム番号を取得する。
//  var lastColumn = sheet.getRange(6, 9).getNextDataCell(SpreadsheetApp.Direction.NEXT).getColumn();
//  //指定セルを起点に最終行の行番号を取得する。
//  var lastRow = sheet.getRange(7, 2).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
//  
//  var recipientColumn = "";
//  //引数の宛先（recipient）が情報共有宛先項目と一致するカラム番号を取得する。
//  for(var i = 9; i <= lastColumn; i++) {
//    if(dropdownValue == sheet.getRange(6, i).getValue()){
//      recipientColumn = i;
//      break;
//    }
//  }
//  if(recipientColumn == "") {
//    throw new Error("宛先の選択が不正です。");
//  }
//  
//  var recipients = [" "];
//  var cc = [" "];
//  for (var i = 7; i <= lastRow; i++) {
//    if(sheet.getRange(i, recipientColumn).getValue() == "○") {
//      if(!sheet.getRange(i, 4).isBlank()) {
//        recipients.push(sheet.getRange(i, 3).getValue() + " <" + sheet.getRange(i, 4).getValue() + ">");
//      }
//    } else if(sheet.getRange(i, recipientColumn).getValue() == "△") {
//      if(!sheet.getRange(i, 4).isBlank()) {
//        cc.push(sheet.getRange(i, 3).getValue() + " <" + sheet.getRange(i, 4).getValue() + ">");
//      }
//    }
//  }
//  console.log(recipients);

//console.log(text);

  const isPlainText = e.formInput['is_plain_text'];
  console.log('is_plain_text:' + isPlainText);

  const body = isPlainText ? text : text.replace(/\r?\n/g, '<br/>');
  const contentType = isPlainText ? CardService.ContentType.TEXT : CardService.ContentType.MUTABLE_HTML;
  
//  console.log('{body, contentType}' + body + ',' + contentType);

  const response =  CardService.newUpdateDraftActionResponseBuilder()
        .setUpdateDraftBodyAction(CardService.newUpdateDraftBodyAction()
        .addUpdateContent(body, contentType)
        .setUpdateType(CardService.UpdateDraftBodyType.IN_PLACE_INSERT)).build();
  
  return response;
}

function getValuesFromSpreadsheet(){
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  
  const lastColumn = sheet.getRange(6, 9).getNextDataCell(SpreadsheetApp.Direction.NEXT).getColumn();
  const range = sheet.getRange(6, 9, 1, lastColumn - 8).getValues();
  return range[0];
}