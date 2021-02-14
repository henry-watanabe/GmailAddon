/**
 * 設定画面（カード）の生成
 * 
 * @return {CardService.Card} 生成したカード
 */
function createSettingCard() {

  let builder = CardService.newCardBuilder()
    .setDisplayStyle(CardService.DisplayStyle.REPLACE)
    .setHeader(CardService.newCardHeader()
      .setTitle('宛先一覧設定')
    )
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextInput()
        .setTitle('宛先一覧ファイルのスプレッドシートID')
        .setFieldName('spreadSheetId')
        .setHint('上記入力後、追加ボタンを押してください。')
      )
      .addWidget(CardService.newTextButton()
        .setText('追加')
        .setBackgroundColor('#0000FF')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction()
          .setFunctionName('addSpreadSheet')
        )
      )
    )

  // 登録済みがある場合だけリストセクションを追加する。
  let sheets = RecipientListRepository.getSheetsFromProperty();
  if (sheets && sheets.length > 0) {
    builder.addSection(createListSection(sheets));
  }

  return builder.setFixedFooter(CardService.newFixedFooter()
    .setPrimaryButton(CardService.newTextButton()
      .setText('戻る')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onReturn')
      )
    )
    .setSecondaryButton(CardService.newTextButton()
      .setText('設定をクリア')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onClearSetting')
      )
    )
  )
    .build();
}

/**
 * 設定カードのリストセクションを生成する
 * @module Settings.gs
 * @param {Sheets} sheets 宛先一覧シート
 * @return {CardService.CardSection} 生成したセクション
 */
function createListSection(sheets) {
  let section = CardService.newCardSection()
    .setCollapsible(true)
    .setHeader('登録済みリスト');
  for (let i = 0; i < sheets.length; i++) {
    section.addWidget(CardService.newDecoratedText()
      .setWrapText(true)
      .setText(sheets[i].fileName)
      .setOpenLink(CardService.newOpenLink()
        .setUrl('https://docs.google.com/spreadsheets/d/' + sheets[i].fileId + '/')
        // .setOnClose(CardService.OnClose.RELOAD_ADD_ON)
      )
      .setButton(CardService.newTextButton()
        .setTextButtonStyle(CardService.TextButtonStyle.TEXT)
        .setText('<font color="#0000FF">除外</font>')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('removeSpreadSheet')
          .setParameters({ 'sheetId': sheets[i].fileId })
        )
      )
    )
  }
  return section;
}

/**
 * 追加ボタンのイベントハンドラ。
 * 指定されたスプレッドシートIDをリストに追加する。
 * 
 * @param {Event} e イベント
 */
function addSpreadSheet(e) {
  Logger.log('addSheet');
  Logger.log('selected file:' + e.formInput.spreadSheetId);

  // RecipientListSheet型で保持し続けるのは無駄なのでIDだけで扱う
  let fileIds = RecipientListRepository.getSheetsFromProperty().map(s => s.fileId);
  fileIds.push(e.formInput.spreadSheetId);
  RecipientListRepository.saveSheetsToProperty(fileIds);

  return CardService.newNavigation().updateCard(createSettingCard());
}

/**
 * 除外ボタンのイベントハンドラ。
 * 宛先一覧設定の中から指定のスプレッドシートを除外する。
 * 
 * @param {Event} e イベント
 */
function removeSpreadSheet(e) {
  Logger.log('removeSpreadSheet');
  Logger.log('sheetId:' + e.parameters.sheetId);

  let sheetId = e.parameters.sheetId;
  // RecipientListSheet型で保持し続けるのは無駄なのでIDだけで扱う
  let sheets = RecipientListRepository.getSheetsFromProperty().map(s => s.fileId);
  sheets.splice(sheets.indexOf(sheetId), 1);
  RecipientListRepository.saveSheetsToProperty(sheets);

  return CardService.newNavigation().updateCard(createSettingCard());
}

/**
 * 設定情報をクリアする。
 */
function onClearSetting(e) {
  PropertiesService.getUserProperties().deleteProperty('SPREAD_SHEET_IDS');
}

function onReturn(e) {
  return CardService.newNavigation().popCard();
}
