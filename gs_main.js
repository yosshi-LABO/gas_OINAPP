// ==========================================
// 押印申請システム メイン処理 (gs_main.gs)
// ==========================================

// 庶務課のグループメールアドレス
const ADMIN_EMAIL = 'multistlipe.re@gmail.com'
const SHOMU_EMAIL = 'orangepop.0x0.hawks66@gmail.com'
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbzfv__eoVBunxXsdEuhUU2PEpPvDfTgCo3TTyZp07EgrKVAALyc2V6FwYkm3WDuxbRc/exec'
// ※画面のコピーボタンで取得したURL

/**
 * 1. WebアプリのURLにアクセスがあったときに呼び出される関数 (doGet)
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

  // 2. IDが指定されている場合（承認 または 再申請）
  if (id) {
    var data = getApplicationDataByRefNum_(id);
    
    if (!data) {
      var errorHtml = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="font-family:sans-serif;padding:40px;text-align:center;">
        <h2>⚠️ エラー</h2><p>指定された起案番号 (<strong>${id}</strong>) の申請データが見つかりませんでした。</p>
      </body></html>`;
      return HtmlService.createHtmlOutput(errorHtml);
    }

    template.formData = data;
    template.mode = mode; // 'approval' または 'resubmit' がセットされる

    var pageTitle = (mode === 'approval') ? '押印申請 承認確認' : '押印申請 再申請';

    return template.evaluate()
      .setTitle(pageTitle)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 3. IDが無い場合 ➔ 新規申請画面
  template.formData = createDefaultFormData_(getNextReferenceNumber());
  template.mode = 'new';

  return template.evaluate()
    .setTitle('押印申請 フォーム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getApplicantEmail_() {
  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail || userEmail.trim() === "") {
    return ADMIN_EMAIL;
  }
  return userEmail;
}

function submitData(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請データ');

  //申請者のメールアドレスが取得できない（undefined や ""）場合に、エラー落ちや空欄保存を防ぐ予備処理
  var applicantEmail = data.applicantEmail;
  if (!applicantEmail || applicantEmail.trim() === "") {
    applicantEmail = SHOMU_EMAIL;
  }

  // 1. B列（起案番号）を検索し、対象行を決定（無ければ末尾）
  var rows = sheet.getDataRange().getValues();
  var targetRow = sheet.getLastRow() + 1;
  var existingHistory = [];

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === String(data.referenceNumber).trim()) {
      targetRow = i + 1;
      try { existingHistory = JSON.parse(rows[i][18] || '[]'); } catch(e) {}
      break;
    }
  }

  // 2. 履歴の追加
  existingHistory.push({
    action: "申請",
    user: data.employeeName,
    email: applicantEmail,
    time: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss"),
    comment: data.remarks || ""
  });

  // 3. データの書き込み（既存行なら上書きUPDATE、無ければ追加）
  var rowData = [
    data.applicationDate, 
    data.referenceNumber, 
    data.companyName,     
    data.deptName,        
    data.employeeId,      
    data.employeeName,    
    data.departmentType,  
    data.sealCompanies,   
    data.clientCompany,   
    data.documentName,    
    data.copies,          
    data.sealType,        
    data.approvalNum,     
    data.remarks,         
    applicantEmail, 
    data.approverEmail,   
    '申請中', // ステータスは常に「申請中」で上書き
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss"),           
    JSON.stringify(existingHistory)
  ];

  sheet.getRange(targetRow, 1, 1, rowData.length).setNumberFormat("@");
  sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

  // 4. 各種通知
  if (data.approverEmail) sendApprovalRequestMail_(data.approverEmail, data);
  if (data.notifyType === 'line') {
    sendLineNotification(data.referenceNumber, data.companyName, data.employeeName, data.documentName);
  } else if (data.notifyType === 'gmail') {
    sendGmailNotification(data.referenceNumber, data.companyName, data.employeeName, data.documentName);
  }

  return "成功";
}



function sendApprovalRequestMail_(approverEmail, data) {
// var xxxUrl = ScriptApp.getService().getUrl()+"?id=" + encodeURIComponent(data.referenceNumber);
  var targetUrl = WEB_APP_URL + "?id=" + encodeURIComponent(data.referenceNumber)+ "&mode=approval";

  var subject = "【押印申請】承認依頼通知（" + data.referenceNumber + "）";
  var body = data.employeeName + " 様\n\n"
           + "承認依頼メールです。承認または差戻しの操作を行ってください。\n\n"
           + "----------------------------------------\n"
           + "■ 操作URL:\n"
           + targetUrl + "\n"
           + "----------------------------------------\n"
           + "※このメールはシステムから自動送信されています。";
  MailApp.sendEmail(approverEmail, subject, body);
}

function updateStatusAndNotify_(refNum, action, comment) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請データ');
  var rows = sheet.getDataRange().getValues();
  
  for (var i = 1; i < rows.length; i++) {
    var currentRowRefNum = String(rows[i][1]);

    if (currentRowRefNum === String(refNum)) {
      var rowIndex = i + 1;

      var actionName = (action === 'approve') ? '承認' : '差戻し';
      var newStatus  = (action === 'approve') ? '承認完了' : '差戻し';
      
      sheet.getRange(rowIndex, 17).setValue(newStatus);

      var company        = rows[i][2];  
      var applicantName  = rows[i][5];  
      var docName        = rows[i][9];  
      var applicantEmail = rows[i][14]; 
      var approverEmail  = rows[i][15]; 

      addHistory(sheet, rowIndex, actionName, "承認者", approverEmail || ADMIN_EMAIL, comment || "");
      
      if (applicantEmail && String(applicantEmail).trim() !== "") {
        //var userSubject = "【押印申請】処理完了通知（申請者宛）（" + refNum + "）";
        var userSubject = (action === 'approve')
          ? "【押印申請】最終承認通知（申請者宛）（" + refNum + "）"
          : "【押印申請】差戻し通知（申請者宛）（" + refNum + "）"

        // ★差し戻し用URL（id=起案番号&mode=resubmit）
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

function sendGmailNotification(refNum, company, employee, docName) {
  var recipient = Session.getActiveUser().getEmail(); // GWSログインユーザー
  // ★ 宛先が取得できない（未ログイン・シークレットモード等）場合は送信せず終了
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

function sendLineNotification(refNum, company, employee, docName) {
  var CHANNEL_ACCESS_TOKEN = "QKJdMeKep3UrncPRm2qdjPxHHXgomSl3ybMMSuu75QejvxSCw6KjBz8zFDEN3LUrDXZNZngUkZ2bTr2aroZSUhnFPOuxjXDzuiPI0G+csn2tu+AALtH7rz2WKOy5kRsnSpVYyNpu6upB5TJS46mayAdB04t89/1O/w1cDnyilFU=";
  var USER_ID = "Ue63486ad4d781f909c9a6ccfd4ed8c93";
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

function processApprovalAction(refNum, action, comment) {
  return updateStatusAndNotify_(refNum, action, comment);
}

function addHistory(sheet, rowIndex, actionName, userName, userEmail, comment) {
  var historyCol = 19;
  var currentHistoryJson = sheet.getRange(rowIndex, historyCol).getValue();
  
  var historyList = [];
  try {
    if (currentHistoryJson) {
      historyList = JSON.parse(currentHistoryJson);
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

function getNextReferenceNumber() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請データ');
  var rows = sheet.getDataRange().getValues();
  
  var currentYearMonth = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMM");
  var prefix = "OIN" + currentYearMonth +"-" ;
  
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

function getApplicationDataByRefNum_(refNum) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請データ');
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return null;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var currentRefNum = String(row[1]).trim();

    if (currentRefNum === String(refNum).trim()) {
      var appDateStr = '';
      if (row[0]) {
        var d = new Date(row[0]);
        appDateStr = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
      } else {
        var today = new Date();
        appDateStr = today.getFullYear() + '-' + ('0' + (today.getMonth() + 1)).slice(-2) + '-' + ('0' + today.getDate()).slice(-2);
      }

      var sealCompaniesArr = [];
      if (row[7]) {
        sealCompaniesArr = String(row[7]).split(',').map(function(s) { return s.trim(); });
      }

      // ★ 19列目(S列 / index:18)の履歴JSON文字列をオブジェクトに変換
      var historyArr = [];
      try {
        if (row[18]) {
          historyArr = JSON.parse(row[18]);
        }
      } catch(e) {
        historyArr = [];
      }

      return {
        applicationDate: appDateStr,            // A列
        referenceNumber: row[1] || refNum,       // B列
        companyName:     row[2] || '',           // C列
        deptName:        row[3] || '',           // D列
        employeeId:      row[4] || '',           // E列
        employeeName:    row[5] || '',           // F列
        departmentType:  row[6] || 'オフィス',    // G列
        sealCompanies:   sealCompaniesArr,       // H列
        clientCompany:   row[8] || '',           // I列
        documentName:    row[9] || '',           // J列
        copies:          row[10] || 1,           // K列
        sealType:        row[11] || '社印(角印)',// L列
        approvalNum:     row[12] || '',          // M列
        remarks:         row[13] || '',          // N列
        approverEmail:   row[15] || '',           // P列
        status:          row[16] || '',           // Q列
        history:         historyArr              // ★ 履歴データを追加
      };
    }
  }
  return null;
}

function createDefaultFormData_(newRefNum) {
  var today = new Date();
  var yyyy = today.getFullYear();
  var mm = ('0' + (today.getMonth() + 1)).slice(-2);
  var dd = ('0' + today.getDate()).slice(-2);

  return {
    applicationDate: yyyy + '-' + mm + '-' + dd,
    referenceNumber: newRefNum,
    companyName: 'MGT株式会社',
    deptName: '情報システム部',
    employeeId: '000001',
    employeeName: '山田 太郎',
    departmentType: 'オフィス',
    sealCompanies: [],
    clientCompany: '',
    documentName: '',
    copies: 1,
    sealType: '社印(角印)',
    approvalNum: '',
    remarks: '',
    approverEmail: '',
    history: [] // ★ 追記: 初期状態の空配列を持たせる
  };
}