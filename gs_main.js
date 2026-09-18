// ==========================================
// 押印申請システム メイン処理 (gs_main.gs)
// ==========================================

// システム設定（宛先メールアドレス・WebアプリURL） ⭐️⭐️ ポートフォリオ用にdummy値 ADMIN_EMAIL、SHOMU_EMAIL、WEB_APP_URL ⭐️⭐️
const ADMIN_EMAIL  = 'admin_dummy@gmail.com';
const SHOMU_EMAIL  = 'shomu_dummy@gmail.com';
const WEB_APP_URL  = 'https://script.google.com/macros/s/AKfycbzfv__dummy_Rc/exec';

/**
 * 💡 WebアプリのURLにアクセスがあったときに呼び出されるメイン関数です。
 * URLパラメータ（?page=list や ?id=OIN... 等）を見て、以下のどの画面を表示するか判定・描画します。
 *  - 一覧画面 (?page=list)
 *  - 承認画面 (?id=xxx&mode=approval)
 *  - 再申請画面 (?id=xxx&mode=resubmit)
 *  - 新規申請画面 (パラメータなし)
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('html_form');
  var params = (e && e.parameter) ? e.parameter : {};

  // 1. 一覧画面へのアクセス (?page=list)
  if (params.page === 'list') {
    return HtmlService.createTemplateFromFile('html_list')
      .evaluate()
      .setTitle('押印申請 一覧画面')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var id = params.id;              // 起案番号 (例: OIN2608-0090)
  var mode = params.mode || 'new'; // モード (approval / resubmit / 指定なしなら new)

  // 2. IDが指定されている場合（承認 または 再申請画面を表示）
  if (id) {
    var data = getApplicationDataByRefNum_(id);
    
    if (!data) {
      var errorHtml = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="font-family:sans-serif;padding:40px;text-align:center;">
        <h2>⚠️ エラー</h2><p>指定された起案番号 (<strong>${id}</strong>) の申請データが見つかりませんでした。</p>
      </body></html>`;
      return HtmlService.createHtmlOutput(errorHtml);
    }

    template.formData = data;
    template.mode = mode; // 'approval' または 'resubmit' をセット

    var pageTitle = (mode === 'approval') ? '押印申請 承認確認' : '押印申請 再申請';

    return template.evaluate()
      .setTitle(pageTitle)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 3. IDが無い場合 ➔ 新規申請フォームを表示
  template.formData = createDefaultFormData_(getNextReferenceNumber());
  template.mode = 'new';

  return template.evaluate()
    .setTitle('押印申請 フォーム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 💡 HTMLファイルから別ファイル（css_formやjs_formなど）を読み込むための共通パーツ関数です。
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 💡 ログイン中のユーザー（Googleアカウント）のメールアドレスを取得します。
 * 取得できない場合はシステム管理者のメールアドレスを返します。
 */
function getApplicantEmail_() {
  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail || userEmail.trim() === "") {
    return ADMIN_EMAIL;
  }
  return userEmail;
}

/**
 * 💡 申請フォームから送信されたデータをスプレッドシートに保存（新規登録または上書き更新）し、
 * 承認者への依頼メールや申請者への控え通知（Gmail / LINE）を発行する関数です。
 */
function submitData(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請データ');

  // 申請者メールアドレスの空欄防止対策
  var applicantEmail = data.applicantEmail;
  if (!applicantEmail || applicantEmail.trim() === "") {
    applicantEmail = SHOMU_EMAIL;
  }

  // 1. B列（起案番号）を検索し、既存データがあればその行に上書き、無ければ末尾に追加
  var rows = sheet.getDataRange().getValues();
  var targetRow = sheet.getLastRow() + 1;
  var existingHistory = [];

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === String(data.refNum).trim()) {
      targetRow = i + 1;
      try { existingHistory = JSON.parse(rows[i][18] || '[]'); } catch(e) {}
      break;
    }
  }

  // 2. 処理履歴（ログ）の追加
  existingHistory.push({
    action: "申請",
    user: data.employeeName,
    email: applicantEmail,
    time: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss"),
    comment: data.remarks || ""
  });

  // 3. シートへの書き込み用データ配列作成
  var rowData = [
    data.applicationDate, 
    data.refNum, 
    data.companyName,     
    data.deptName,        
    data.employeeId,      
    data.employeeName,    
    data.deptType,  
    data.sealCompany,   
    data.clientCompany,   
    data.documentName,    
    data.copies,          
    data.sealType,        
    data.approvalNum,     
    data.remarks,         
    applicantEmail, 
    data.approverEmail,   
    '申請中', // ステータスは「申請中」に設定
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss"),            
    JSON.stringify(existingHistory)
  ];

  sheet.getRange(targetRow, 1, 1, rowData.length).setNumberFormat("@");
  sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

  // 4. 各種メール・LINE通知の送信
  if (data.approverEmail) sendApprovalRequestMail_(data.approverEmail, data);
  if (data.notifyType === 'line') {
    sendLineNotification(data.refNum, data.companyName, data.employeeName, data.documentName);
  } else if (data.notifyType === 'gmail') {
    sendGmailNotification(data.refNum, data.companyName, data.employeeName, data.documentName);
  }

  return "成功";
}

/**
 * 💡 承認者に「承認依頼メール」を自動送信する関数です。
 * メール本文には承認操作用URL（?id=xxx&mode=approval）が含まれます。
 */
function sendApprovalRequestMail_(approverEmail, data) {
  var targetUrl = WEB_APP_URL + "?id=" + encodeURIComponent(data.refNum) + "&mode=approval";

  var subject = "【押印申請】承認依頼通知（" + data.refNum + "）";
  var body = "[承認者]" + " 様\n\n"
           + "承認依頼メールです。承認または差戻しの操作を行ってください。\n\n"
           + "----------------------------------------\n"
           + "■ 操作URL:\n"
           + targetUrl + "\n"
           + "----------------------------------------\n"
           + "※このメールはシステムから自動送信されています。";
  MailApp.sendEmail(approverEmail, subject, body);
}

/**
 * 💡 画面（HTML）の「承認」or「差戻し」ボタンが押されたときに呼び出される受付関数です。
 */
function processApprovalAction(refNum, action, comment) {
  return updateStatusAndNotify_(refNum, action, comment);
}

/**
 * 💡 承認・差戻し処理の本体です。スプレッドシートのステータスを更新し、
 * 申請者（承認/差戻し通知）および庶務課（最終承認通知）へ完了メールを送信します。
 */
function updateStatusAndNotify_(refNum, action, comment) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請データ');
  var rows = sheet.getDataRange().getValues();
  
  for (var i = 1; i < rows.length; i++) {
    var currentRowRefNum = String(rows[i][1]);

    if (currentRowRefNum === String(refNum)) {
      var rowIndex = i + 1;

      var actionName = (action === 'approve') ? '承認' : '差戻し';
      var newStatus  = (action === 'approve') ? '承認完了' : '差戻し';
      
      // Q列（ステータス）を更新
      sheet.getRange(rowIndex, 17).setValue(newStatus);

      var company        = rows[i][2];  
      var applicantName  = rows[i][5];  
      var docName        = rows[i][9];  
      var applicantEmail = rows[i][14]; 
      var approverEmail  = rows[i][15]; 

      // 履歴ログを追加
      addHistory(sheet, rowIndex, actionName, "承認者", approverEmail || ADMIN_EMAIL, comment || "");
      
      // 申請者への結果通知メール送信
      if (applicantEmail && String(applicantEmail).trim() !== "") {
        var userSubject = (action === 'approve')
          ? "【押印申請】最終承認通知（申請者宛）（" + refNum + "）"
          : "【押印申請】差戻し通知（申請者宛）（" + refNum + "）";

        var targetUrl = WEB_APP_URL + "?id=" + encodeURIComponent(refNum) + "&mode=resubmit";

        var resultMessage = (action === 'approve') 
          ? "申請が最終承認されました。" 
          : "申請が差戻されました。\n申請内容を確認・再申請してください。\n" + targetUrl;
          
        var userBody = applicantName + " 様\n\n"
                     + resultMessage + "\n\n"
                     + "----------------------------------------\n"
                     + "・起案番号: " + refNum + "\n"
                     + "・起案会社: " + company + "\n"
                     + "・申請者: " + applicantName + "\n"
                     + "・書類名: " + docName + "\n"
                     + "・処理結果: " + newStatus + "\n"
                     + (comment ? "・コメント(承認者): " + comment + "\n" : "")
                     + "----------------------------------------\n"
                     + "※このメールはシステムから自動送信されています。";
        MailApp.sendEmail(applicantEmail, userSubject, userBody);
      }

      // 承認完了時のみ庶務課へ通知メール送信
      if (action === 'approve') {
        var shomuSubject = "【押印申請】最終承認通知（庶務宛）（" + refNum + "）";
        var shomuBody = "庶務課各位\n\n"
                      + "以下の申請が最終承認されました。\n"
                      + "ご確認をお願いいたします。\n\n"
                      + "----------------------------------------\n"
                      + "・起案番号: " + refNum + "\n"
                      + "・起案会社: " + company + "\n"
                      + "・申請者: " + applicantName + "\n"
                      + "・書類名: " + docName + "\n"
                      + (comment ? "・コメント(承認者): " + comment + "\n" : "")
                      + "----------------------------------------\n"
                      + "※このメールはシステムから自動送信されています。";
        MailApp.sendEmail(SHOMU_EMAIL, shomuSubject, shomuBody);
      }
      return { success: true, newStatus: newStatus };
    }
  }

  return { success: false, message: "対象データが見つかりませんでした。" };
}

/**
 * 💡 申請完了時に、申請者本人宛に確認用（控え）メールをGmail送信する関数です。
 */
function sendGmailNotification(refNum, company, employee, docName) {
  var recipient = Session.getActiveUser().getEmail();
  if (!recipient || recipient.trim() === "") {
    Logger.log("申請者のメールアドレスが取得できなかったため、控えメール送信をスキップしました。");
    return;
  }

  var subject = "【押印申請】申請登録通知（申請者宛） （" + refNum + "）";
  var body = employee + " 様\n\n" +
             "以下の内容で申請が登録されています。\n承認者による確認まで今しばらくお待ちください。\n\n" +
             "----------------------------------------\n" +
             "・起案番号: " + refNum + "\n" +
             "・起案会社: " + company + "\n" +
             "・申請者: " + employee + "\n" +
             "・書類名: " + docName + "\n" +
             "----------------------------------------\n\n" +
             "※このメールはシステムから自動送信されています。";

  MailApp.sendEmail(recipient, subject, body);
}

/**
 * 💡 申請完了時に、LINE Messaging APIを使ってLINEに通知を送信する関数です。
 */
function sendLineNotification(refNum, company, employee, docName) {
　/* ⭐️⭐️ ポートフォリオ用にダミー値（CHANNEL_ACCESS_TOKEN、USER_ID） ⭐️⭐️ */
  var CHANNEL_ACCESS_TOKEN = "QKJdM_dummy";
  var USER_ID = "Ue634_dummy";
  var url = "https://api.line.me/v2/bot/message/push";
  
  var message = "【押印申請】申請登録通知（申請者宛）（" + refNum + "）\n\n" +
             "以下の内容で申請が登録されています。\n承認者による確認まで今しばらくお待ちください。\n\n" +
             "----------------------------------------\n" +
             "・起案番号: " + refNum + "\n" +
             "・起案会社: " + company + "\n" +
             "・申請者: " + employee + "\n" +
             "・書類名: " + docName + "\n" +
             "----------------------------------------\n\n" +
             "※このメールは申請者本人へ自動送信されています。";

  var payload = {
    "to": USER_ID,
    "messages": [{"type": "text", "text": message}]
  };
  
  var options = {
    "method": "post",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + CHANNEL_ACCESS_TOKEN
    },
    "payload": JSON.stringify(payload)
  };
  
  UrlFetchApp.fetch(url, options);
}

/**
 * スプレッドシートのS列（履歴情報JSON）に、操作内容（申請・承認・差戻し等）のログを追加記録する関数です。
 */
function addHistory(sheet, rowIndex, actionName, userName, userEmail, comment) {
  var historyCol = 19;
  var currentHistoryJson = sheet.getRange(rowIndex, historyCol).getValue();
  
  var historyList = [];
  try {
    if (currentHistoryJson) {
      historyList = (typeof currentHistoryJson === 'string') ? JSON.parse(currentHistoryJson) : currentHistoryJson;
    }
  } catch(e) {
    historyList = [];
  }
  
  historyList.push({
    action: actionName,
    user: userName,
    email: userEmail,
    time: Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
    comment: comment || ""
  });
  
  sheet.getRange(rowIndex, historyCol).setValue(JSON.stringify(historyList));
}

/**
 * 現在年月（例: 2608）を元に、最新の連番を取得して新しい起案番号（例: OIN2608-0001）を自動採番する関数です。
 */
function getNextReferenceNumber() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請データ');
  var rows = sheet.getDataRange().getValues();
  
  var currentYearMonth = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMM");
  var prefix = "OIN" + currentYearMonth + "-";
  
  var maxSeq = 0;
  
  for (var i = 1; i < rows.length; i++) {
    var refNum = String(rows[i][1]);
    if (refNum.indexOf(prefix) === 0) {
      var seqStr = refNum.replace(prefix, "");
      var seq = parseInt(seqStr, 10);
      if (!isNaN(seq) && seq > maxSeq) {
        maxSeq = seq;
      }
    }
  }
  
  var nextSeq = maxSeq + 1;
  return prefix + String(nextSeq).padStart(4, '0');
}

/**
 * 新規申請フォームを開いたときに表示するデフォルト初期データを生成する関数です。
 */
function createDefaultFormData_(newRefNum) {
  var today = new Date();
  var yyyy = today.getFullYear();
  var mm = ('0' + (today.getMonth() + 1)).slice(-2);
  var dd = ('0' + today.getDate()).slice(-2);

  return {
    applicationDate: yyyy + '-' + mm + '-' + dd,
    refNum: newRefNum,
    companyName: 'MGT株式会社',
    deptName: '情報システム部',
    employeeId: '000001',
    employeeName: '山田 太郎',
    deptType: 'オフィス',
    sealCompany: [],
    clientCompany: '',
    documentName: '',
    copies: 1,
    sealType: '社印(角印)',
    approvalNum: '',
    remarks: '',
    approverEmail: '',
    history: []
  };
}

/**
 * スプレッドシートの1行（データ配列）を、JavaScriptで扱いやすいオブジェクト（連想配列）へ変換する関数です。
 * フォーム入力用・一覧表示用の両方で利用できるよう項目名を共通化しています。
 */
function convertRowToObj_(row) {
  if (!row || row.length === 0) return null;

  // 押印会社（カンマ区切り文字列を配列化）
  var sealCompanyArr = [];
  if (row[7]) {
    sealCompanyArr = String(row[7]).split(',').map(function(s) { return s.trim(); });
  }

  // 履歴（JSON文字列を配列オブジェクト化）
  var historyArr = [];
  try {
    if (row[18]) {
      historyArr = (typeof row[18] === 'string') ? JSON.parse(row[18]) : row[18];
    }
  } catch(e) {
    historyArr = [];
  }

  return {
    applicationDate: formatDate_(row[0]),                 // A列: 申請日
    refNum:          String(row[1] || ''),               // B列: 起案番号
    companyName:     String(row[2] || ''),               // C列: 起案会社
    deptName:        String(row[3] || ''),               // D列: 所属部署
    employeeId:      String(row[4] || ''),               // E列: 申請者CD
    employeeName:    String(row[5] || ''),               // F列: 申請者名
    deptType:        String(row[6] || 'オフィス'),        // G列: 所属部署種別
    sealCompany:     sealCompanyArr,                     // H列: 押印会社(配列)
    sealCompanyStr:  String(row[7] || ''),              // H列: 押印会社(表示用文字列)
    clientCompany:   String(row[8] || ''),               // I列: 相手先会社名
    documentName:    String(row[9] || ''),               // J列: 書類名
    copies:          row[10] || 1,                       // K列: 部数
    sealType:        String(row[11] || '社印(角印)'),     // L列: 押印種別
    approvalNum:     String(row[12] || ''),              // M列: 稟決番号
    remarks:         String(row[13] || ''),              // N列: 備考
    applicantEmail:  String(row[14] || ''),              // O列: 申請者Mail
    approverEmail:   String(row[15] || ''),              // P列: 承認者Mail
    status:          row[16] ? String(row[16]) : '申請中',// Q列: ステータス
    timestamp:       row[17] ? String(new Date(row[17]).toLocaleString('ja-JP')) : '', // R列: タイムスタンプ
    history:         historyArr                          // S列: 履歴(配列)
  };
}

/**
 * 日付データを「YYYY-MM-DD」形式の文字列に変換するヘルパー関数です。
 */
function formatDate_(dateVal) {
  if (!dateVal) return '';
  var d = new Date(dateVal);
  if (isNaN(d.getTime())) return String(dateVal);
  var yyyy = d.getFullYear();
  var mm = ('0' + (d.getMonth() + 1)).slice(-2);
  var dd = ('0' + d.getDate()).slice(-2);
  return yyyy + '-' + mm + '-' + dd;
}

/**
 * 起案番号（例: OIN2608-0001）を指定して、該当する申請データをスプレッドシートから検索・取得する関数です。
 */
function getApplicationDataByRefNum_(refNum) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請データ');
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return null;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(refNum).trim()) {
      return convertRowToObj_(data[i]);
    }
  }
  return null;
}