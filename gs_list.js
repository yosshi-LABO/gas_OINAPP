// ==========================================
// 一覧・検索画面用のスクリプト
// ==========================================

// 検索キーワードに基づいてスプレッドシートからデータを取得する関数
function getFilteredData(keyword) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請データ');
  var rows = sheet.getDataRange().getValues();
  
  if (rows.length <= 1) {
    return []; // ヘッダー行のみの場合は空を返す
  }
  
  var data = [];
  
  // 2行目からループ（1行目はヘッダー）
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    

    // スプレッドシートの全列を明示的に文字列（String）に変換して格納
    var item = {
      applicationDate: formatDate_(row[0]),                              // A列: 申請日 (※formatDate_内で文字列に変換されている前提)
      referenceNumber: String(row[1] || ''),                             // B列: 起案番号
      companyName:     String(row[2] || ''),                             // C列: 起案会社
      deptName:        String(row[3] || ''),                             // D列: 所属部署
      employeeId:    String(row[4] || ''),                             // E列: 申請者CD
      employeeName:    String(row[5] || ''),                             // F列: 申請者名
      deptType:        String(row[6] || ''),                             // G列: 所属部署種別
      sealCompanies:   String(row[7] || ''),                             // H列: 押印が必要な会社
      partnerCompany:  String(row[8] || ''),                             // I列: 相手先会社名
      documentName:    String(row[9] || ''),                             // J列: 書類名
      copies:          String(row[10] || ''),                            // K列: 必要な部数
      sealType:        String(row[11] || ''),                            // L列: 必要な押印
      approvalNumber:  String(row[12] || ''),                            // M列: 稟決番号
      remarks:         String(row[13] || ''),                            // N列: 備考
      applicantEmail:  String(row[14] || ''),                            // O列: 申請者メールアドレス
      approverEmail:   String(row[15] || ''),                            // P列: 承認者メールアドレス
      status:          row[16] ? String(row[16]) : '申請中',             // Q列: ステータス
      timestamp:       row[17] ? String(new Date(row[17]).toLocaleString('ja-JP')) : '', // R列: タイムスタンプ
      history:         String(row[18] || '')                             // S列: 履歴情報
    };

    /*  スプレッドシートの全16列（A〜P列）をオブジェクトに格納
    var item = {
      applicationDate: formatDate_(row[0]), // A列: 申請日
      referenceNumber: row[1] || '',        // B列: 起案番号
      companyName:     row[2] || '',        // C列: 起案会社
      deptName:        row[3] || '',        // D列: 所属部署
      employeeCode:    row[4] || '',        // E列: 申請者CD
      employeeName:    row[5] || '',        // F列: 申請者名
      deptType:        row[6] || '',        // G列: 所属部署種別
      sealCompanies:   row[7] || '',        // H列: 押印が必要な会社
      partnerCompany:  row[8] || '',        // I列: 相手先会社名
      documentName:    row[9] || '',        // J列: 書類名
      copies:          row[10] || '',       // K列: 必要な部数
      sealType:        row[11] || '',       // L列: 必要な押印
      approvalNumber:  row[12] || '',       // M列: 稟決番号
      remarks:         row[13] || '',       // N列: 備考
      applicantEmail:  row[14],              // O列: 申請者メールアドレス ★追加
      approverEmail:   row[15],              // P列: 承認者メールアドレス ★追加
      status:          row[16] || '申請中',  // Q列: ステータス ★[14]から[16]へ変更！
      //timestamp:       formatDate_(row[17])  // R列: タイムスタンプ ★[15]から[17]へ変更！
      timestamp:       row[17] ? new Date(row[17]).toLocaleString('ja-JP') : '',
      history:         row[18] || ''        // S列: 履歴情報（JSON文字列） ★ここを追加！
    };
    */

    // キーワード検索の判定（キーワードが空、または各項目に含まれている場合）
    if (!keyword || keyword.trim() === "") {
      data.push(item);
    } else {
      var match = Object.values(item).some(function(val) {
        return String(val).toLowerCase().includes(keyword.toLowerCase());
      });
      if (match) {
        data.push(item);
      }
    }
  }
  
  // 新しい申請が上に来るように逆順にして返す
  return data.reverse();
}

// 日付オブジェクトをきれいに「YYYY/MM/DD」形式の文字列にするヘルパー関数
function formatDate_(date) {
  if (!date) return "";
  if (Object.prototype.toString.call(date) === '[object Date]') {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd");
  }
  return String(date);
}