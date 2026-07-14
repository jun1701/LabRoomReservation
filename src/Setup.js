/**
 * データベース（スプレッドシート）の初期化スクリプト
 * 初回構築時に1度だけ実行します。
 */
function initializeDatabase() {
  // 取得したデータベースIDを指定してください
  const DATABASE_ID = "1pvfd-mwkUFyjPGjMRFtZov_2PE6nR_2hUwrV5JqEV3Y";
  
  const ss = SpreadsheetApp.openById(DATABASE_ID);

  // 9シートの定義とそれぞれのヘッダー列
  const sheetsConfig = [
    {
      name: "Config",
      headers: ["Key", "Value", "Description"]
    },
    {
      name: "Reservations",
      headers: ["ReservationID", "UserEmail", "UserName", "Date", "StartTime", "EndTime", "Purpose", "Status", "CreatedAt", "UpdatedAt"]
    },
    {
      name: "ReservationHistory",
      headers: ["ReservationID", "ActionType", "BeforeData", "AfterData", "OperatorEmail", "Timestamp"]
    },
    {
      name: "AuditLog",
      headers: ["Timestamp", "UserEmail", "Action", "TargetID", "Details"]
    },
    {
      name: "Users",
      headers: ["UserEmail", "UserName", "Role", "CreatedAt", "UpdatedAt"]
    },
    {
      name: "SystemLog",
      headers: ["Timestamp", "LogLevel", "Message"]
    },
    {
      name: "MailQueue",
      headers: ["QueueID", "ReservationID", "MailType", "Recipient", "Subject", "Body", "ScheduledAt", "Status", "SentAt", "ErrorMessage"]
    },
    {
      name: "CalendarSync",
      headers: ["ReservationID", "CalendarEventID", "SyncStatus", "LastSyncedAt"]
    },
    {
      name: "Holidays",
      headers: ["Date", "Description"]
    }
  ];

  // 各シートの作成とヘッダー設定
  for (let i = 0; i < sheetsConfig.length; i++) {
    const config = sheetsConfig[i];
    
    let sheet = ss.getSheetByName(config.name);

    // シートが存在しない場合は新規作成
    if (!sheet) {
      sheet = ss.insertSheet(config.name);
    }

    // ヘッダー行（1行目）を書き込み
    sheet.getRange(1, 1, 1, config.headers.length).setValues([config.headers]);
    
    // ヘッダー行を固定
    sheet.setFrozenRows(1);
  }

  // デフォルトの初期シートが存在する場合は削除
  const defaultSheet = ss.getSheetByName("シート1");
  
  if (defaultSheet) {
    ss.deleteSheet(defaultSheet);
  }
  
  const defaultSheetEn = ss.getSheetByName("Sheet1");
  
  if (defaultSheetEn) {
    ss.deleteSheet(defaultSheetEn);
  }

  Logger.log("9シート構成でのデータベース初期化が完了しました。");
}