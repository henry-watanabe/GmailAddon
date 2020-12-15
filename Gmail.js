var SHEET_ID = "1PB7lYZAoERaT9vED0onNCCC3dVcEbCfd1VMxyEMLwvE";
var SHEET_NAME = "宛先一覧";

function onGmailCompose(e) {
  console.log(e);

  var header = CardService.newCardHeader()
      .setTitle('宛先選択')
      .setSubtitle('以下のプルダウンから宛先グループを選択してください。');
  
  var dropdown = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName("dropdown_field");
  var item = getValuesFromSpreadsheet();
  for(var i = 0; i < item.length; i++){
    dropdown.addItem(item[i], item[i], false);
  }

  var action = CardService.newAction()
      .setFunctionName('onGmailInsertRecipient');
  var footer = CardService.newFixedFooter()
    .setPrimaryButton(CardService.newTextButton()
      .setText('ＰＩＣＫ！')
      .setOnClickAction(action));

  var section = CardService.newCardSection()
      .addWidget(dropdown)
  var card = CardService.newCardBuilder()
      .setHeader(header)
      .addSection(section)
      .setFixedFooter(footer);
  return card.build();
}

function onGmailInsertRecipient(e) {
  console.log(e);
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  
  var dropdownValue = e.formInput.dropdown_field;
  
  //指定セルを起点に最終列のカラム番号を取得する。
  var lastColumn = sheet.getRange(6, 9).getNextDataCell(SpreadsheetApp.Direction.NEXT).getColumn();
  //指定セルを起点に最終行の行番号を取得する。
  var lastRow = sheet.getRange(7, 2).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  
  var recipientColumn = "";
  //引数の宛先（recipient）が情報共有宛先項目と一致するカラム番号を取得する。
  for(var i = 9; i <= lastColumn; i++) {
    if(dropdownValue == sheet.getRange(6, i).getValue()){
      recipientColumn = i;
      break;
    }
  }
  if(recipientColumn == "") {
    throw new Error("宛先の選択が不正です。");
  }
  
  var recipients = [" "];
  var cc = [" "];
  for (var i = 7; i <= lastRow; i++) {
    if(sheet.getRange(i, recipientColumn).getValue() == "○") {
      if(!sheet.getRange(i, 4).isBlank()) {
        recipients.push(sheet.getRange(i, 3).getValue() + " <" + sheet.getRange(i, 4).getValue() + ">");
      }
    } else if(sheet.getRange(i, recipientColumn).getValue() == "△") {
      if(!sheet.getRange(i, 4).isBlank()) {
        cc.push(sheet.getRange(i, 3).getValue() + " <" + sheet.getRange(i, 4).getValue() + ">");
      }
    }
  }
  console.log(recipients);
  
  var response = CardService.newUpdateDraftActionResponseBuilder()
    .setUpdateDraftToRecipientsAction(CardService.newUpdateDraftToRecipientsAction()
       .addUpdateToRecipients(recipients))
    .setUpdateDraftCcRecipientsAction(CardService.newUpdateDraftCcRecipientsAction()
       .addUpdateCcRecipients(cc))
    .build()
  
  return response;
}

function getValuesFromSpreadsheet(){
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  
  const lastColumn = sheet.getRange(6, 9).getNextDataCell(SpreadsheetApp.Direction.NEXT).getColumn();
  const range = sheet.getRange(6, 9, 1, lastColumn - 8).getValues();
  return range[0];
}