/**
 * ホームページトリガーに対するハンドラ。
 * ホーム画面を生成し表示する。
 */
function onDefaultHomePageOpen(e) {
  Logger.log('user:' + Session.getActiveUser().getUserLoginId());

  const iconBytes = DriveApp.getFileById('1gRvKjR7ZTFHCnkP8KfWfq9DDjRhw6jmg').getBlob().getBytes();
  const encodedIconUrl = "data:image/jpeg;base64," + Utilities.base64Encode(iconBytes);

  const imageBytes = DriveApp.getFileById('10qdT_KqRwDd48IsjW5JVmZk2MT7X4pc-').getBlob().getBytes();
  const encodedImageUrl = "data:image/jpeg;base64," + Utilities.base64Encode(imageBytes);

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('ホーム')
      .setImageUrl(encodedIconUrl)
      .setImageStyle(CardService.ImageStyle.CIRCLE)
    )
    .addCardAction(CardService.newCardAction()
      .setText('宛先一覧設定')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('createSettingCard')
      )
    )
    .addSection(CardService.newCardSection()
      .setHeader('メッセージ')
      .addWidget(CardService.newDecoratedText().setText('<font color="#FF0000">初めての方は右上の三点リーダーから宛先一覧の設定をしてください。</font>').setWrapText(true))
      .addWidget(CardService.newImage().setImageUrl(encodedImageUrl))
    )
    .setFixedFooter(CardService.newFixedFooter()
      .setPrimaryButton(CardService.newTextButton()
        .setText('宛先一覧の設定をクリアする')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('onClearSetting')
        )
      )
    )
    .build();
}

/**
 * ユニバーサルアクション
 * @deprecated
 */
function createSettingsResponse(e) {
  return CardService.newUniversalActionResponseBuilder()
    .displayAddOnCards([createSettingCard()])
    .build();
}

