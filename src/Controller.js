/**
 * フロントエンド（HTML画面）からのリクエストを受け付けるコントローラー
 */

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('研究室予約システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function apiGetUserNameByEmail(fullEmail) {
  try {
    const sheet = getSheet(SHEETS.RESERVATIONS);
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][1] === fullEmail && data[i][2]) {
        return data[i][2];
      }
    }
  } catch (error) {
    // エラー時は空文字を返す
  }
  return "";
}

function timeToMinutes(timeStr) {
  const parts = timeStr.split(':');
  return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
}

function isOverlapping(dateStr, startTime, endTime, excludeResId) {
  const sheet = getSheet(SHEETS.RESERVATIONS);
  const data = sheet.getDataRange().getValues();
  const newStartMin = timeToMinutes(startTime);
  const newEndMin = timeToMinutes(endTime);
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][7] === "Cancelled") continue;
    if (excludeResId && data[i][0] === excludeResId) continue;
    
    const rDate = Utilities.formatDate(new Date(data[i][3]), Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (rDate === dateStr) {
      let rStart = data[i][4];
      if (rStart instanceof Date) rStart = Utilities.formatDate(rStart, Session.getScriptTimeZone(), "HH:mm");
      let rEnd = data[i][5];
      if (rEnd instanceof Date) rEnd = Utilities.formatDate(rEnd, Session.getScriptTimeZone(), "HH:mm");
      
      const rStartMin = timeToMinutes(rStart);
      const rEndMin = timeToMinutes(rEnd);
      
      if (newStartMin < rEndMin && newEndMin > rStartMin) {
        return true;
      }
    }
  }
  return false;
}

function checkReservationRules(reqDateStr, reqStartTime, reqEndTime) {
  try {
    const sheet = getSheet("Rules");
    if (!sheet) return null; // Rulesシートがない場合は制限なしとして通過
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return null; // ヘッダーのみの場合は通過
    
    const reqDate = new Date(reqDateStr);
    const reqTimeStart = timeToMinutes(reqStartTime);
    const reqTimeEnd = timeToMinutes(reqEndTime);
    const daysArr = ['日', '月', '火', '水', '木', '金', '土'];
    const reqDoW = daysArr[reqDate.getDay()];
    
    for (let i = 1; i < data.length; i++) {
      const [startDate, endDate, dow, startTime, endTime, reason] = data[i];
      if (!startDate && !endDate && !dow && !startTime && !endTime) continue;
      
      if (startDate) {
        const ruleStart = new Date(startDate);
        ruleStart.setHours(0, 0, 0, 0);
        if (reqDate < ruleStart) continue;
      }
      if (endDate) {
        const ruleEnd = new Date(endDate);
        ruleEnd.setHours(23, 59, 59, 999);
        if (reqDate > ruleEnd) continue;
      }
      
      if (dow && typeof dow === 'string' && dow.trim() !== '') {
        if (!dow.includes(reqDoW)) continue;
      }
      
      let ruleTimeStart = 0;
      let ruleTimeEnd = 1440;
      
      if (startTime) {
        const sTimeStr = startTime instanceof Date ? Utilities.formatDate(startTime, Session.getScriptTimeZone(), "HH:mm") : startTime.toString();
        ruleTimeStart = timeToMinutes(sTimeStr);
      }
      if (endTime) {
        const eTimeStr = endTime instanceof Date ? Utilities.formatDate(endTime, Session.getScriptTimeZone(), "HH:mm") : endTime.toString();
        ruleTimeEnd = timeToMinutes(eTimeStr);
      }
      
      if (reqTimeStart < ruleTimeEnd && reqTimeEnd > ruleTimeStart) {
        return reason || "予約制限ルールに該当するため、この日時は予約できません。";
      }
    }
  } catch (error) {
    const systemLogSheet = getSheet(SHEETS.SYSTEM_LOG);
    systemLogSheet.appendRow([new Date(), "ERROR", `checkReservationRules失敗: ${error.message}`]);
  }
  return null;
}

function apiGetReservations(dateStr) {
  try {
    const sheet = getSheet(SHEETS.RESERVATIONS);
    const data = sheet.getDataRange().getValues();
    const reservations = [];
    if (data.length <= 1) return reservations;
    
    for (let i = 1; i < data.length; i++) {
      const status = data[i][7];
      if (status === "Cancelled") continue;
      
      const resDate = Utilities.formatDate(new Date(data[i][3]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      if (resDate === dateStr) {
        let startTime = data[i][4];
        if (startTime instanceof Date) startTime = Utilities.formatDate(startTime, Session.getScriptTimeZone(), "HH:mm");
        let endTime = data[i][5];
        if (endTime instanceof Date) endTime = Utilities.formatDate(endTime, Session.getScriptTimeZone(), "HH:mm");
        
        reservations.push({
          reservationId: data[i][0],
          userName: data[i][2],
          startTime: startTime,
          endTime: endTime,
          purpose: data[i][6]
        });
      }
    }
    return reservations;
  } catch (error) {
    const systemLogSheet = getSheet(SHEETS.SYSTEM_LOG);
    systemLogSheet.appendRow([new Date(), "ERROR", `apiGetReservations失敗: ${error.message}`]);
    throw new Error("予約データの取得に失敗しました。");
  }
}

function apiCreateReservation(requestData) {
  const timestamp = new Date();
  try {
    const reqStartDateTime = new Date(`${requestData.date}T${requestData.startTime}:00`);
    if (reqStartDateTime < timestamp) {
      return { success: false, message: "過去の日時は予約できません。" };
    }
    
    const ruleReason = checkReservationRules(requestData.date, requestData.startTime, requestData.endTime);
    if (ruleReason) {
      return { success: false, message: ruleReason };
    }
    
    if (isOverlapping(requestData.date, requestData.startTime, requestData.endTime, null)) {
      return { success: false, message: "指定された時間は既に他の予約が入っています。" };
    }
    
    const reservationId = createReservation(
      requestData.userEmail,
      requestData.userName,
      requestData.date,
      requestData.startTime,
      requestData.endTime,
      requestData.purpose
    );
    
    syncCalendarEvent(reservationId);
    
    const historySheet = getSheet(SHEETS.RESERVATION_HISTORY);
    historySheet.appendRow([reservationId, "CREATE", "", `Date: ${requestData.date}, Time: ${requestData.startTime}-${requestData.endTime}, Purpose: ${requestData.purpose}`, requestData.userEmail, timestamp]);
    
    const appUrl = ScriptApp.getService().getUrl();
    const mailSubject = `【予約完了】研究室利用予約 (${requestData.date})`;
    const mailBody = `${requestData.userName} 様\n\n研究室の利用予約が完了しました。\n\n・予約ID: ${reservationId}\n・日付: ${requestData.date}\n・時間: ${requestData.startTime} ～ ${requestData.endTime}\n・目的: ${requestData.purpose}\n\n予約の確認、変更、キャンセルは以下のシステム画面より行ってください。\n${appUrl}`;
    
    enqueueMail(reservationId, "CREATE", requestData.userEmail, mailSubject, mailBody);
    processMailQueue();
    
    return {
      success: true,
      message: "予約が正常に完了しました。確認メールを送信しました。",
      reservationId: reservationId
    };
  } catch (error) {
    const systemLogSheet = getSheet(SHEETS.SYSTEM_LOG);
    systemLogSheet.appendRow([timestamp, "ERROR", `apiCreateReservation失敗: ${error.message}`]);
    return {
      success: false,
      message: error.message || "予約処理中に予期せぬエラーが発生しました。"
    };
  }
}

function apiGetReservationById(reservationId) {
  try {
    const sheet = getSheet(SHEETS.RESERVATIONS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === reservationId) {
        const status = data[i][7];
        if (status === "Cancelled") {
          throw new Error("この予約は既にキャンセルされています。");
        }
        let startTime = data[i][4];
        if (startTime instanceof Date) startTime = Utilities.formatDate(startTime, Session.getScriptTimeZone(), "HH:mm");
        let endTime = data[i][5];
        if (endTime instanceof Date) endTime = Utilities.formatDate(endTime, Session.getScriptTimeZone(), "HH:mm");
        
        return {
          reservationId: data[i][0],
          userEmail: data[i][1],
          userName: data[i][2],
          date: Utilities.formatDate(new Date(data[i][3]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
          startTime: startTime,
          endTime: endTime,
          purpose: data[i][6]
        };
      }
    }
    throw new Error("指定された予約IDが見つかりません。");
  } catch (error) {
    throw new Error(error.message);
  }
}

function apiGetReservationsByUser(userEmail) {
  try {
    const sheet = getSheet(SHEETS.RESERVATIONS);
    const data = sheet.getDataRange().getValues();
    const userReservations = [];
    const now = new Date();
    const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const nowTimeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm");
    
    if (data.length <= 1) return userReservations;
    
    for (let i = 1; i < data.length; i++) {
      const email = data[i][1];
      const status = data[i][7];
      if (email === userEmail && status !== "Cancelled") {
        const resDate = Utilities.formatDate(new Date(data[i][3]), Session.getScriptTimeZone(), "yyyy-MM-dd");
        let startTime = data[i][4];
        if (startTime instanceof Date) startTime = Utilities.formatDate(startTime, Session.getScriptTimeZone(), "HH:mm");
        let endTime = data[i][5];
        if (endTime instanceof Date) endTime = Utilities.formatDate(endTime, Session.getScriptTimeZone(), "HH:mm");
        
        if (resDate < todayStr) continue;
        if (resDate === todayStr && endTime <= nowTimeStr) continue;
        
        userReservations.push({
          reservationId: data[i][0],
          userName: data[i][2],
          date: resDate,
          startTime: startTime,
          endTime: endTime,
          purpose: data[i][6]
        });
      }
    }
    userReservations.sort(function(a, b) {
      const dateA = new Date(a.date + "T" + a.startTime);
      const dateB = new Date(b.date + "T" + b.startTime);
      return dateA - dateB;
    });
    return userReservations;
  } catch (error) {
    throw new Error("予約データの取得に失敗しました: " + error.message);
  }
}

function apiUpdateReservation(requestData) {
  const timestamp = new Date();
  try {
    const reqStartDateTime = new Date(`${requestData.date}T${requestData.startTime}:00`);
    if (reqStartDateTime < timestamp) {
      return { success: false, message: "過去の日時に変更することはできません。" };
    }
    
    const ruleReason = checkReservationRules(requestData.date, requestData.startTime, requestData.endTime);
    if (ruleReason) {
      return { success: false, message: ruleReason };
    }
    
    if (isOverlapping(requestData.date, requestData.startTime, requestData.endTime, requestData.reservationId)) {
      return { success: false, message: "変更先の時間は既に他の予約が入っています。" };
    }
    
    const updatedFields = {
      Date: requestData.date,
      StartTime: requestData.startTime,
      EndTime: requestData.endTime,
      Purpose: requestData.purpose
    };
    
    updateReservation(requestData.reservationId, requestData.userEmail, updatedFields);
    syncCalendarEvent(requestData.reservationId);
    
    const historySheet = getSheet(SHEETS.RESERVATION_HISTORY);
    historySheet.appendRow([requestData.reservationId, "UPDATE", "", `NewDate: ${requestData.date}, Time: ${requestData.startTime}-${requestData.endTime}, Purpose: ${requestData.purpose}`, requestData.userEmail, timestamp]);
    
    const appUrl = ScriptApp.getService().getUrl();
    const mailSubject = `【予約変更】研究室利用予約 (${requestData.date})`;
    const mailBody = `${requestData.userName} 様\n\n研究室の利用予約を変更しました。\n\n・予約ID: ${requestData.reservationId}\n・日付: ${requestData.date}\n・時間: ${requestData.startTime} ～ ${requestData.endTime}\n・目的: ${requestData.purpose}\n\n予約の確認、変更、キャンセルは以下のシステム画面より行ってください。\n${appUrl}`;
    
    enqueueMail(requestData.reservationId, "UPDATE", requestData.userEmail, mailSubject, mailBody);
    processMailQueue();
    
    return { success: true, message: "予約の変更が完了しました。" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function apiCancelReservation(reservationId, userEmail) {
  const timestamp = new Date();
  try {
    cancelReservation(reservationId, userEmail);
    deleteCalendarEvent(reservationId);
    
    const historySheet = getSheet(SHEETS.RESERVATION_HISTORY);
    historySheet.appendRow([reservationId, "CANCEL", "", "", userEmail, timestamp]);
    
    const appUrl = ScriptApp.getService().getUrl();
    const mailSubject = `【予約キャンセル】研究室利用予約`;
    const mailBody = `研究室の利用予約をキャンセルしました。\n\n・予約ID: ${reservationId}\n\nシステム画面:\n${appUrl}`;
    
    enqueueMail(reservationId, "CANCEL", userEmail, mailSubject, mailBody);
    processMailQueue();
    
    return { success: true, message: "予約をキャンセルしました。" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * ユーザーからの問い合わせ・要望を受け付けて処理するAPI
 */
function apiSendInquiry(requestData) {
  const timestamp = new Date();
  
  try {
    // 1. 予約シートを基準にして、スプレッドシート本体を確実に取得する
    const ss = getSheet(SHEETS.RESERVATIONS).getParent();
    let sheet = ss.getSheetByName("Inquiries");
    
    // シートが存在しない場合は自動作成する
    if (!sheet) {
      sheet = ss.insertSheet("Inquiries");
      sheet.appendRow(["タイムスタンプ", "メールアドレス", "氏名", "カテゴリ", "対象日時", "内容", "対応ステータス"]);
      sheet.getRange("A1:G1").setFontWeight("bold").setBackground("#f3f3f3");
    }
    
    const targetDateTimeStr = requestData.targetDate ? `${requestData.targetDate} ${requestData.targetTime}` : "指定なし";
    
    sheet.appendRow([
      timestamp, 
      requestData.userEmail, 
      requestData.userName, 
      requestData.category, 
      targetDateTimeStr, 
      requestData.message, 
      "未対応" // 初期ステータス
    ]);
    
    // 2. メールの送信
    const adminEmail = "seigyo.reserve.system@gmail.com"; // 管理者（システム）アドレス
    const subject = `【システム申請・問い合わせ】${requestData.category} (${requestData.userName}様より)`;
    const body = `
システムから新しい問い合わせ・申請を受信しました。

■ 申請者情報
・氏名: ${requestData.userName}
・メールアドレス: ${requestData.userEmail}

■ 問い合わせ内容
・カテゴリ: ${requestData.category}
・対象日時: ${targetDateTimeStr}
・メッセージ:
${requestData.message}

※ スプレッドシートの「Inquiries」シートにも記録されています。
    `.trim();

    // 管理者宛てに送信
    MailApp.sendEmail({
      to: adminEmail,
      subject: subject,
      body: body
    });
    
    // 申請者本人にも控えを送信
    MailApp.sendEmail({
      to: requestData.userEmail,
      subject: `【控え】${subject}`,
      body: `以下の内容で送信を完了しました。管理者の確認をお待ちください。\n\n---\n${body}`
    });
    
    return { success: true, message: "送信が完了しました。ご意見・ご要望ありがとうございます。" };
    
  } catch (error) {
    const systemLogSheet = getSheet(SHEETS.SYSTEM_LOG);
    if (systemLogSheet) {
      systemLogSheet.appendRow([timestamp, "ERROR", `apiSendInquiry失敗: ${error.message}`]);
    }
    return { success: false, message: "送信に失敗しました。" + error.message };
  }
}