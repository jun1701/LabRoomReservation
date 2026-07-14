/**
 * 予約管理のコアロジックを処理するサービス
 */

/**
 * 指定された日時に予約の重複がないかチェックする関数
 */
function checkOverlap(date, startTime, endTime, excludeReservationId) {
  const sheet = getSheet(SHEETS.RESERVATIONS);
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return false;
  }
  
  // 比較用にHH:mm形式の文字列として扱う
  const newStartStr = startTime;
  const newEndStr = endTime;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const currentId = row[0];
    const currentDate = Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    // シート上の時間を確実に HH:mm 形式の文字列に変換
    let currentStartStr = row[4];
    if (currentStartStr instanceof Date) {
      currentStartStr = Utilities.formatDate(currentStartStr, Session.getScriptTimeZone(), "HH:mm");
    } else if (typeof currentStartStr === "string" && currentStartStr.length < 5) {
      currentStartStr = ("0" + currentStartStr).slice(-5); // 9:00 -> 09:00
    }
    
    let currentEndStr = row[5];
    if (currentEndStr instanceof Date) {
      currentEndStr = Utilities.formatDate(currentEndStr, Session.getScriptTimeZone(), "HH:mm");
    } else if (typeof currentEndStr === "string" && currentEndStr.length < 5) {
      currentEndStr = ("0" + currentEndStr).slice(-5);
    }
    
    const currentStatus = row[7];
    
    if (excludeReservationId && currentId === excludeReservationId) {
      continue;
    }
    
    if (currentStatus === "Cancelled") {
      continue;
    }
    
    if (currentDate === date) {
      // 文字列の大小比較で重複判定 (例: "09:00" < "10:30")
      if (newStartStr < currentEndStr && newEndStr > currentStartStr) {
        return true;
      }
    }
  }
  
  return false;
}

function createReservation(userEmail, userName, date, startTime, endTime, purpose) {
  if (checkOverlap(date, startTime, endTime)) {
    throw new Error("指定された時間帯は既に他の予約が入っています。");
  }
  
  const sheet = getSheet(SHEETS.RESERVATIONS);
  const timestamp = new Date();
  const reservationId = "RES_" + timestamp.getTime() + Math.floor(Math.random() * 1000);
  const status = "Active";
  
  sheet.appendRow([
    reservationId, userEmail, userName, date, startTime, endTime, purpose, status, timestamp, timestamp
  ]);
  
  const historySheet = getSheet(SHEETS.RESERVATION_HISTORY);
  historySheet.appendRow([
    reservationId, "CREATE", "", JSON.stringify({ userEmail, userName, date, startTime, endTime, purpose, status }), userEmail, timestamp
  ]);
  
  const auditSheet = getSheet(SHEETS.AUDIT_LOG);
  auditSheet.appendRow([
    timestamp, userEmail, "CREATE_RESERVATION", reservationId, `新規予約登録: ${date} ${startTime}-${endTime}`
  ]);
  
  return reservationId;
}

function updateReservation(reservationId, operatorEmail, updatedFields) {
  const sheet = getSheet(SHEETS.RESERVATIONS);
  const data = sheet.getDataRange().getValues();
  let targetRowIndex = -1;
  let beforeData = null;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reservationId) {
      targetRowIndex = i + 1;
      beforeData = {
        reservationId: data[i][0],
        userEmail: data[i][1],
        userName: data[i][2],
        date: Utilities.formatDate(new Date(data[i][3]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
        startTime: data[i][4],
        endTime: data[i][5],
        purpose: data[i][6],
        status: data[i][7]
      };
      break;
    }
  }
  
  if (targetRowIndex === -1) throw new Error("指定された予約が見つかりません。");
  if (beforeData.status === "Cancelled") throw new Error("キャンセル済みの予約は変更できません。");
  
  const checkDate = updatedFields.Date || beforeData.date;
  let checkStart = updatedFields.StartTime || beforeData.startTime;
  if (checkStart instanceof Date) checkStart = Utilities.formatDate(checkStart, Session.getScriptTimeZone(), "HH:mm");
  let checkEnd = updatedFields.EndTime || beforeData.endTime;
  if (checkEnd instanceof Date) checkEnd = Utilities.formatDate(checkEnd, Session.getScriptTimeZone(), "HH:mm");
  
  if (checkOverlap(checkDate, checkStart, checkEnd, reservationId)) {
    throw new Error("変更先の時間帯は既に他の予約が入っています。");
  }
  
  const timestamp = new Date();
  
  if (updatedFields.Date) sheet.getRange(targetRowIndex, 4).setValue(updatedFields.Date);
  if (updatedFields.StartTime) sheet.getRange(targetRowIndex, 5).setValue(updatedFields.StartTime);
  if (updatedFields.EndTime) sheet.getRange(targetRowIndex, 6).setValue(updatedFields.EndTime);
  if (updatedFields.Purpose) sheet.getRange(targetRowIndex, 7).setValue(updatedFields.Purpose);
  sheet.getRange(targetRowIndex, 10).setValue(timestamp);
  
  const afterData = { ...beforeData, ...updatedFields, status: beforeData.status };
  
  const historySheet = getSheet(SHEETS.RESERVATION_HISTORY);
  historySheet.appendRow([
    reservationId, "UPDATE", JSON.stringify(beforeData), JSON.stringify(afterData), operatorEmail, timestamp
  ]);
  
  const auditSheet = getSheet(SHEETS.AUDIT_LOG);
  auditSheet.appendRow([
    timestamp, operatorEmail, "UPDATE_RESERVATION", reservationId, `予約変更: ${checkDate} ${checkStart}-${checkEnd}`
  ]);
}

function cancelReservation(reservationId, operatorEmail) {
  const sheet = getSheet(SHEETS.RESERVATIONS);
  const data = sheet.getDataRange().getValues();
  let targetRowIndex = -1;
  let beforeData = null;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reservationId) {
      targetRowIndex = i + 1;
      beforeData = {
        reservationId: data[i][0],
        userEmail: data[i][1],
        userName: data[i][2],
        date: Utilities.formatDate(new Date(data[i][3]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
        startTime: data[i][4],
        endTime: data[i][5],
        purpose: data[i][6],
        status: data[i][7]
      };
      break;
    }
  }
  
  if (targetRowIndex === -1) throw new Error("指定された予約が見つかりません。");
  if (beforeData.status === "Cancelled") throw new Error("この予約は既にキャンセルされています。");
  
  const timestamp = new Date();
  sheet.getRange(targetRowIndex, 8).setValue("Cancelled");
  sheet.getRange(targetRowIndex, 10).setValue(timestamp);
  
  const afterData = { ...beforeData, status: "Cancelled" };
  
  const historySheet = getSheet(SHEETS.RESERVATION_HISTORY);
  historySheet.appendRow([
    reservationId, "CANCEL", JSON.stringify(beforeData), JSON.stringify(afterData), operatorEmail, timestamp
  ]);
  
  const auditSheet = getSheet(SHEETS.AUDIT_LOG);
  auditSheet.appendRow([
    timestamp, operatorEmail, "CANCEL_RESERVATION", reservationId, "予約キャンセル"
  ]);
}