/**
 * メール受信者を示すクラス。
 */
class Recipient {

  /**
   * @param {String} name 名前
   * @param {String} address メールアドレス
   */
  constructor(name, address) {
    this.name = name;
    this.address = address;
  }

  /**
   * 受信者名とメールアドレスを「受信者名 <メールアドレス>」形式に整形して文字列として返す。
   * @return {String} 整形後文字列
   */
  format() {
    return this.name + ' <' + this.address + '>'
  }
}
