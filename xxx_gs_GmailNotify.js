// Gmail通知

// GASの仕様上、末尾に _ がついた関数は「プライベート関数」扱いになり、他の関数からの誤呼び出しを防げるほか、エディタ上の実行ボタンリストからも非表示になります。
function sendApprovalNotification_() {
  // 1. アクティブなスプレッドシートとシートを取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // 2. データの行数・列数を取得（例：2行目のデータを最新の申請データとする）
  var lastRow = sheet.getLastRow();
  
  // データがヘッダー行（1行目）しかない場合は何もしない
  if (lastRow < 2) {
    Logger.log("送信対象のデータがありません。");
    return;
  }
  
  // 最新行（一番下の行）のデータを取得
  var applicantName = sheet.getRange(lastRow, 3).getHeading ? sheet.getRange(lastRow, 3).getValue() : sheet.getRange(lastRow, 3).getValue(); // 申請者名
  var referenceNumber = sheet.getRange(lastRow, 1).getValue(); // 起案番号
  var companyName   = sheet.getRange(lastRow, 2).getValue(); // 起案会社
  
  // 3. 送信先メールアドレス（テストとして自分のメールアドレスを指定）
  var recipient = Session.getActiveUser().getEmail();
  var subject   = "【自動通知】新しい押印申請が登録されました (" + referenceNumber + ")";
  var body      = "承認者様\n\n" +
                  "新しい押印申請が起案されました。\n\n" +
                  "・起案番号: " + referenceNumber + "\n" +
                  "・起案会社: " + companyName + "\n" +
                  "・申請者: " + applicantName + "\n\n" +
                  "内容を確認し、社内ポータルより承認作業を行ってください。";
  
  // 4. GmailAppを使ってメールを送信（外部サービス連携の第一歩）
  GmailApp.sendEmail(recipient, subject, body);
  
  Logger.log("通知メールを送信しました: " + recipient);
}