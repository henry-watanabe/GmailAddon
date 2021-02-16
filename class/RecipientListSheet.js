/** シート名 */
const SHEET_NAME = '宛先一覧';
/** 宛先グループ行 */
const GROUP_START_ROW = 1;
/** 宛先グループ開始列 */
const GROUP_START_COL = 4;
/** 宛先開始行 */
const NUMBER_START_ROW = 2;
/** 行番号（#）列 */
const NUMBER_START_COL = 1;
/** メールアドレス列 */
const MAIL_ADDRESS_COL = 3;
/** 氏名列 */
const MEMBER_COL = 2;

/**
 * 宛先一覧のスプレッドシートを体現したクラス
 */
class RecipientListSheet {

  /**
   * 渡されたファイルIDに対応する{@class SpreadSheet}から「宛先一覧」という名のシートを{@class Sheet}型で返す。
   */
  constructor(fileId) {
    try {
      this.fileId = fileId;
      const file = SpreadsheetApp.openById(fileId);
      this.fileName = file.getName();
      this._sheet = file.getSheetByName(SHEET_NAME);
    } catch (e) {
      Logger.log(e);
      throw Error('引数の fileId に対応するスプレッドシートが存在しません。[fileId: ' + fileId + '; fileName: ' + this.fileName + ']');
    }
  }

  /**
   * グループ名のリストを返す。
   * @return {Array} グループ名リスト
   */
  getGroups() {
    let lastCol = this._sheet.getRange(1, 4).getNextDataCell(SpreadsheetApp.Direction.NEXT).getColumn();
    return this._sheet.getRange(1, 4, 1, lastCol).getValues()[0];
  }

  /**
   * 指定されたグループ名に対応する宛先の一覧を返す。
   * [toRecipients : (to宛先), ccRecipients : (cc宛先)]
   * 
   * @param {String} groupName 宛先グループ名
   */
  getRecipients(groupName) {
    //指定セルを起点に最終列のカラム番号を取得する。
    let lastColumn = this._sheet.getRange(GROUP_START_ROW, GROUP_START_COL).getNextDataCell(SpreadsheetApp.Direction.NEXT).getColumn();
    //指定セルを起点に最終行の行番号を取得する。
    let lastRow = this._sheet.getRange(NUMBER_START_ROW, NUMBER_START_COL).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();

    let recipientColumn = 0;
    //引数の宛先グループ名が情報共有宛先項目と一致するカラム番号を取得する。
    for (let i = GROUP_START_COL; i <= lastColumn; i++) {
      if (groupName == this._sheet.getRange(GROUP_START_ROW, i).getValue()) {
        recipientColumn = i;
        break;
      }
    }

    if (recipientColumn == 0) {
      throw new Error('宛先の選択が不正です。');
    }

    let toRecipients = [];
    let ccRecipients = [];
    let identifier;
    for (let i = NUMBER_START_ROW; i <= lastRow; i++) {
      identifier = this._sheet.getRange(i, recipientColumn).getValue();
      if (identifier == '○') {
        this._createRecipients(toRecipients, i, this._sheet);
      } else if (identifier == '△') {
        this._createRecipients(ccRecipients, i, this._sheet);
      }
    }
    let recipients = {};
    recipients['toRecipients'] = toRecipients;
    recipients['ccRecipients'] = ccRecipients;

    return recipients;
  }

  /**
   * 宛先一覧スプレッドシートから氏名とメールアドレスを取得し、宛先の配列に追加する。
   * @param recipientsKinds - 宛先の配列
   * @param rowIndex - 行番号
   * @param sheet - シート
   */
  _createRecipients(recipientsKinds, rowIndex, sheet) {
    if (!sheet.getRange(rowIndex, MAIL_ADDRESS_COL).isBlank()) {
      recipientsKinds.push(new Recipient(sheet.getRange(rowIndex, MEMBER_COL).getValue(), sheet.getRange(rowIndex, MAIL_ADDRESS_COL).getValue()));
      // recipientsKinds.push(sheet.getRange(rowIndex, MEMBER_COL).getValue() + ' <' + sheet.getRange(rowIndex, MAIL_ADDRESS_COL).getValue() + '>');
    }
  }
}