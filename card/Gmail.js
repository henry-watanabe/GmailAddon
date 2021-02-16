/**
 * メール作成画面下部のアイコンからアドオン画面を開いたときに呼ばれる処理。
 * アドオン画面の初期化を行う。
 */
function onGmailCompose(e) {
  Logger.log('user:' + Session.getActiveUser().getUserLoginId());
  // 配列を返さないと Android の Gmail アプリでコケる
  return [createMainCard(e)];
}

function createMainCard(e) {
  Logger.log('createMainCard');

  let selectedFileId;
  if (e.formInputs) {
    selectedFileId = e.formInputs.fileId;
  }

  Logger.log('selectedFileId:' + selectedFileId);

  let fileDropDown = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle('宛先一覧')
    .setOnChangeAction(CardService.newAction().setFunctionName('onFileSelected'))
    .setFieldName('fileId');

  let sheets = RecipientListRepository.getSheetsFromProperty();
  if (sheets && sheets.length > 0) {
    fileDropDown.addItem('', '', false);
    for (let i = 0; i < sheets.length; i++) {
      let isSelected = (sheets[i].fileId == selectedFileId);
      fileDropDown.addItem(sheets[i].fileName, sheets[i].fileId, isSelected);
    }
  } else {
    return createErrorCard('宛先一覧のスプレッドシートが設定されていません。\r設定画面より設定してください。');
  }

  let mainSection = CardService.newCardSection()
    .addWidget(CardService.newDecoratedText().setText('宛先一覧とグループ名を選択してください'))
    .addWidget(fileDropDown)
    ;

  // 宛先一覧ドロップダウンでファイルが選択されているときはそのファイルにある宛先グループと選択ボタンをリスト表示する
  if (selectedFileId) {
    let groups = new RecipientListSheet(e.formInputs.fileId).getGroups();
    for (let i = 0; i < groups.length; i++) {
      if (groups[i]) {
        mainSection.addWidget(CardService.newDecoratedText()
          .setText(groups[i])
          .setButton(CardService.newTextButton()
            .setText('選択')
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setBackgroundColor('#0000FF')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('onGroupSelected')
              .setParameters({ 'groupName': groups[i] })
            )
          )
        );
      }
    }
  }

  return CardService.newCardBuilder()
    .setDisplayStyle(CardService.DisplayStyle.REPLACE)
    .setHeader(CardService.newCardHeader().setTitle('下書き'))
    .addCardAction(CardService.newCardAction()
      .setText('宛先一覧設定')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('createSettingCard')
      )
    )
    .addSection(mainSection)
    .setFixedFooter(CardService.newFixedFooter()
      .setPrimaryButton(CardService.newTextButton()
        .setText('下書きに反映する')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('onGmailInsertRecipient')
        )
      )
    )
    .build();
}

function onFileSelected(e) {
  Logger.log('user:' + Session.getActiveUser().getUserLoginId());

  return CardService.newNavigation().updateCard(createMainCard(e));
}

function onGroupSelected(e) {
  Logger.log('user:' + Session.getActiveUser().getUserLoginId());

  e.formInputs.groupNameText = e.parameters.groupName;
  return onGmailInsertRecipient(e);
}

/**
 * グループ名入力時の入力補助。
 * 入力された文字列を含むグループ名を返す。
 */
function suggestGroupName(e) {
  Logger.log('user:' + Session.getActiveUser().getUserLoginId());
  Logger.log(e.formInputs.fileId);

  let spreadSheetId = e.formInputs.fileId;
  let groups = new RecipientListSheet(spreadSheetId).getGroups();
  const pattern = new RegExp(e.formInputs['groupNameText'], 'gi');

  let suggestions = CardService.newSuggestions();
  for (let i = 0; i < groups.length; i++) {
    if (groups[i].toString().match(pattern)) {
      suggestions.addSuggestion(groups[i].toString())
    }
  }
  return CardService.newSuggestionsResponseBuilder().setSuggestions(suggestions).build();
}

/**
 * ドロップダウンで選択した値をもとに、宛先一覧スプレッドシートからTOまたはCCの宛先を抽出し、メールの宛先に出力。
 * @param e - 作成ボタン押下時のイベント。
 * @return レスポンス。
 */
function onGmailInsertRecipient(e) {
  Logger.log('user:' + Session.getActiveUser().getUserLoginId());

  let dropdownValue = e.formInputs['groupNameText'];

  let recipientsList = new RecipientListSheet(e.formInputs.fileId).getRecipients(dropdownValue);

  Logger.log(recipientsList);

  let toRecipients = recipientsList.toRecipients;
  let ccRecipients = recipientsList.ccRecipients;

  if (!toRecipients.length && !ccRecipients.length) {
    throw new Error('宛先が1件もありません。');
  }

  let responseBuilder = CardService.newUpdateDraftActionResponseBuilder();
  if (toRecipients.length > 0) {
    responseBuilder.setUpdateDraftToRecipientsAction(CardService.newUpdateDraftToRecipientsAction()
      .addUpdateToRecipients(toRecipients.map(to => to.format())))
  }
  if (ccRecipients.length > 0) {
    responseBuilder.setUpdateDraftCcRecipientsAction(CardService.newUpdateDraftCcRecipientsAction()
      .addUpdateCcRecipients(ccRecipients.map(cc => cc.format())))
  }
  return responseBuilder.build();
}

/**
 * エラー画面（カード）を作成する。
 * build() は呼び出しもとで実施すること。
 *
 * @module Gmail.gs
 * @param {string} message エラー画面に表示するメッセージ
 * @return エラー画面
 */
function createErrorCard(message) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('エラー画面'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText("<font color='#FF0000'>" + message + "</font>")
      )
    )
    .build();
}
