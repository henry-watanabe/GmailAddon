/**
 * PropertyServiceで管理している宛先一覧スプレッドシートの設定保存や取得に関する関数群
 */
class RecipientListRepository {

  /**
   * PropertyService にJSON形式で保管されている「宛先一覧」スプレッドシートのIDから生成した{@class RecipientListSheet}を配列にして返す。
   * PropertyServiceのキー名は SPREAD_SHEET_IDS
   *
   * @return {Array} RecipientListSheet型の宛先一覧スプレッドシートの配列。1件もない時には空配列。
   */
  static getSheetsFromProperty() {
    const sheetIds = JSON.parse(PropertiesService.getUserProperties().getProperty('SPREAD_SHEET_IDS'));
    return sheetIds ? sheetIds.map(sid => new RecipientListSheet(sid)) : [];
  }

  /**
   * 宛先一覧のスプレッドシートIDリストをJSON形式で PropertyService に保管する。
   * PropertyServiceのキー名は SPREAD_SHEET_IDS
   *
   * @module Homepage.gs
   * @param {Array} sheets 宛先一覧スプレッドシートIDの配列。1件もない時には空配列。
   */
  static saveSheetsToProperty(sheets) {
    PropertiesService.getUserProperties().setProperty('SPREAD_SHEET_IDS', JSON.stringify(sheets));
  }

}
