/**
 * システム全体の共通設定とユーティリティ関数
 */

// データベースIDの定義
const DATABASE_ID = "1pvfd-mwkUFyjPGjMRFtZov_2PE6nR_2hUwrV5JqEV3Y";

// 各シート名の定数定義
const SHEETS = {
  CONFIG: "Config",
  RESERVATIONS: "Reservations",
  RESERVATION_HISTORY: "ReservationHistory",
  AUDIT_LOG: "AuditLog",
  USERS: "Users",
  SYSTEM_LOG: "SystemLog",
  MAIL_QUEUE: "MailQueue",
  CALENDAR_SYNC: "CalendarSync",
  HOLIDAYS: "Holidays"
};

/**
 * 指定されたシート名に対応するSheetオブジェクトを取得する共通関数
 * @param {string} sheetName - 使用するシートの名前（定数SHEETSの値）
 * @return {SpreadsheetApp.Sheet} 対象のシートオブジェクト
 */
function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(DATABASE_ID);
  
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error(`シート「${sheetName}」が見つかりません。データベースの初期化状態を確認してください。`);
  }
  
  return sheet;
}