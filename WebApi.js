/**
 * @file 宛先リストをJSON形式で返す WebAPI の RESTリソース
 */

function doGet(e) {
  const fileId = e.parameter.fileId;
  const group = e.parameter.group;

  const recipients = new RecipientListSheet(fileId).getRecipients(group);

  Logger.log(recipients);

  const output = ContentService.createTextOutput(JSON.stringify(recipients));
  output.setMimeType(ContentService.MimeType.JSON);

  return output;
}

function debug() {
  const response = doGet({'parameter': {'fileId':'1gvof0ZZGfSCSQ9FLOwjI_Dlor4ccqTGJ1Ioer6V_8Ik', 'group':'２つ目'}});
  Logger.log(response.toString());
}
