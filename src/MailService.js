/**
 * メールの配信管理とキュー処理を行うサービス
 */

/**
 * 送信待ちのメールをキュー（MailQueueシート）に登録する関数
 * @param {string} reservationId - 関連する予約ID
 * @param {string} mailType - メールの種類 (CREATE / UPDATE / CANCEL など)
 * @param {string} recipient - 送信先メールアドレス
 * @param {string} subject - 件名
 * @param {string} body - 本文
 */
function enqueueMail(reservationId, mailType, recipient, subject, body) {
  const sheet = getSheet(SHEETS.MAIL_QUEUE);
  
  const timestamp = new Date();
  
  const queueId = "Q_" + timestamp.getTime() + Math.floor(Math.random() * 1000);
  
  const status = "PENDING";
  
  sheet.appendRow([
    queueId,
    reservationId,
    mailType,
    recipient,
    subject,
    body,
    timestamp,
    status,
    "", // SentAt (送信完了時は空)
    ""  // ErrorMessage (初期状態は空)
  ]);
}

/**
 * キューに溜まっている未送信メール（PENDING）を抽出し、実際に送信する関数
 * ※この関数はトリガーで定期実行（例: 1分〜5分おき）させます
 */
function processMailQueue() {
  const sheet = getSheet(SHEETS.MAIL_QUEUE);
  
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return;
  }
  
  const systemLogSheet = getSheet(SHEETS.SYSTEM_LOG);
  
  for (let i = 1; i < data.length; i++) {
    const status = data[i][7];
    
    if (status !== "PENDING") {
      continue;
    }
    
    const targetRowIndex = i + 1;
    
    const queueId = data[i][0];
    
    const recipient = data[i][3];
    
    const subject = data[i][4];
    
    const body = data[i][5];
    
    const timestamp = new Date();
    
    try {
      // 実際のメール送信処理
      GmailApp.sendEmail(recipient, subject, body);
      
      // 送信成功時の更新
      sheet.getRange(targetRowIndex, 8).setValue("SENT");
      
      sheet.getRange(targetRowIndex, 9).setValue(timestamp);
      
    } catch (error) {
      // 送信失敗時の更新とログ記録
      sheet.getRange(targetRowIndex, 8).setValue("ERROR");
      
      sheet.getRange(targetRowIndex, 10).setValue(error.message);
      
      systemLogSheet.appendRow([
        timestamp,
        "ERROR",
        `メール送信失敗 (QueueID: ${queueId}): ${error.message}`
      ]);
    }
  }
}