/**
 * ==========================================
 * 一覧・検索画面用のバックエンド処理 (gs_list.gs)
 * ==========================================
 */

/**
 * 💡 検索キーワードに基づいてスプレッドシートから申請データを一覧取得する関数
 * 
 * @param {string} keyword - 検索ボックスに入力されたキーワード（空文字の場合は全件取得）
 * @return {Array<Object>} 統一プロパティ名（refNum, applicationDate 等）で整形された申請オブジェクトの配列（降順）
 */
function getFilteredData(keyword) {
  // 「申請データ」シートから全セルデータを取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請データ');
  var rows = sheet.getDataRange().getValues();
  
  // データが存在しない、またはヘッダー行（1行目）のみの場合は空配列を返却
  if (rows.length <= 1) {
    return [];
  }
  
  var data = [];
  
  // 2行目（データ行）から順にループ処理
  for (var i = 1; i < rows.length; i++) {
    // 共通関数（gs_main.gs）の convertRowToObj_ を呼び出してオブジェクト化
    // ※内部プロパティ名は統一後の名称（refNum, companyName, deptType 等）に変換されます
    var item = convertRowToObj_(rows[i]);

    // 検索キーワードの有無に応じてフィルタリング
    if (!keyword || keyword.trim() === "") {
      // キーワード指定がない場合は全件対象
      data.push(item);
    } else {
      // キーワードがある場合、オブジェクトの全値から部分一致検索（大文字・小文字を区別しない）
      var match = Object.values(item).some(function(val) {
        return String(val).toLowerCase().includes(keyword.toLowerCase());
      });
      if (match) {
        data.push(item);
      }
    }
  }

  // 新しい申請（スプレッドシートの下の行）が一覧の上に来るように逆順にして返却
  return data.reverse();
}


/**
 * 💡 日付オブジェクトを「YYYY/MM/DD」形式の文字列にフォーマットする内部ヘルパー関数
 * 
 * @param {Date|string} date - スプレッドシートから取得した日付データ
 * @return {string} フォーマット後の日付文字列（空値の場合は空文字）
 */
function formatDate_(date) {
  if (!date) return "";
  
  // Dateオブジェクト判定の上、タイムゾーンに合わせて「yyyy/MM/dd」形式に変換
  if (Object.prototype.toString.call(date) === '[object Date]') {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd");
  }
  
  return String(date);
}