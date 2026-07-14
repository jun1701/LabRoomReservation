/**
 * 管理者向けのシステム処理（トリガー設定・リマインド送信など）を管理するサービス
 */

function handleAdminEdit(e) {
  if (!e || !e.range) return;
  
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEETS.RESERVATIONS) return;
  
  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  if (row <= 1) return;
  
  if (col >= 4 && col <= 8) {
    const resId = sheet.getRange(row, 1).getValue();
    const status = sheet.getRange(row, 8).getValue();
    const timestamp = new Date();
    
    try {
      if (status === "Cancelled") {
        deleteCalendarEvent(resId);
      } else if (status === "Active") {
        syncCalendarEvent(resId);
      }
      sheet.getRange(row, 10).setValue(timestamp);
    } catch (error) {
      const logSheet = getSheet(SHEETS.SYSTEM_LOG);
      logSheet.appendRow([timestamp, "ERROR", `管理者手動編集の同期失敗 (ReservationID: ${resId}): ${error.message}`]);
    }
  }
}

function setupEditTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'handleAdminEdit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  const ss = getSheet(SHEETS.RESERVATIONS).getParent();
  
  ScriptApp.newTrigger('handleAdminEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
    
  console.log("スプレッドシート編集時のトリガー設定が完了しました。");
}

function executeReminders() {
  const now = new Date();
  const hour = now.getHours();
  const appUrl = ScriptApp.getService().getUrl();
  
  if (hour >= 7 && hour <= 9) {
    const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    sendReminders(todayStr, "【本日】", "本日", appUrl);
  } else if (hour >= 19 && hour <= 21) {
    const tomorrow = new Date(now.getTime() + 24 * 60 * 60 * 1000);
    const tomorrowStr = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), "yyyy-MM-dd");
    sendReminders(tomorrowStr, "【明日】", "明日", appUrl);
  }
  
  scheduleExactReminders();
}

function sendReminders(targetDateStr, subjectPrefix, dayWord, appUrl) {
  const sheet = getSheet(SHEETS.RESERVATIONS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const status = data[i][7];
    const resDate = Utilities.formatDate(new Date(data[i][3]), Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    if (status === "Active" && resDate === targetDateStr) {
      const resId = data[i][0];
      const email = data[i][1];
      const name = data[i][2];
      let startTime = data[i][4];
      if (startTime instanceof Date) startTime = Utilities.formatDate(startTime, Session.getScriptTimeZone(), "HH:mm");
      let endTime = data[i][5];
      if (endTime instanceof Date) endTime = Utilities.formatDate(endTime, Session.getScriptTimeZone(), "HH:mm");
      const purpose = data[i][6];
      
      const mailSubject = `${subjectPrefix} 研究室利用のリマインド`;
      const mailBody = `${name} 様\n\n${dayWord}、研究室の利用予約が入っています。\n\n・予約ID: ${resId}\n・日付: ${resDate}\n・時間: ${startTime} ～ ${endTime}\n・目的: ${purpose}\n\n予約の確認、変更、キャンセルは以下のシステム画面より行ってください。\n${appUrl}`;
      
      enqueueMail(resId, "REMINDER", email, mailSubject, mailBody);
    }
  }
  processMailQueue();
}

function scheduleExactReminders() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'executeReminders') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  const now = new Date();
  
  const next8 = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 8, 0, 0);
  if (now.getTime() >= next8.getTime()) {
    next8.setDate(next8.getDate() + 1);
  }
  
  const next20 = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 20, 0, 0);
  if (now.getTime() >= next20.getTime()) {
    next20.setDate(next20.getDate() + 1);
  }
  
  ScriptApp.newTrigger('executeReminders').timeBased().at(next8).create();
  ScriptApp.newTrigger('executeReminders').timeBased().at(next20).create();
  
  console.log(`次回の送信を ${Utilities.formatDate(next8, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm")} と ${Utilities.formatDate(next20, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm")} にセットしました。`);
}

/**
 * Rulesシートの内容を読み取り、カレンダーに「予約不可」予定を書き込む関数
 */
function syncRulesToCalendar() {
  console.log("=== 制限ルールのカレンダー反映処理を開始します ===");
  
  const sheet = getSheet("Rules");
  if (!sheet) {
    console.error("Rulesシートが見つかりません。");
    return;
  }

  const calendarId = "seigyo.reserve.system@gmail.com";
  const cal = CalendarApp.getCalendarById(calendarId);
  
  if (!cal) {
    console.error("カレンダーが取得できません。IDと権限を確認してください。");
    return;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const maxDate = new Date(today);
  maxDate.setDate(maxDate.getDate() + 60);
  maxDate.setHours(23, 59, 59, 999);

  // 1. カレンダー上の古い「予約不可」予定をクリーンアップ
  const existingEvents = cal.getEvents(today, maxDate);
  for (let i = 0; i < existingEvents.length; i++) {
    if (existingEvents[i].getTitle().indexOf("【予約不可】") === 0) {
      existingEvents[i].deleteEvent();
    }
  }

  // 2. Rulesシートから新しい制限予定を生成
  const data = sheet.getDataRange().getValues();
  const daysArr = ['日', '月', '火', '水', '木', '金', '土'];
  let addedCount = 0;

  for (let i = 1; i < data.length; i++) {
    const [startDate, endDate, dow, startTime, endTime, reason] = data[i];

    if (!startDate && !endDate && !dow && !startTime && !endTime) continue;

    let loopStart = startDate ? new Date(startDate) : new Date(today);
    if (loopStart < today) loopStart = new Date(today);

    let loopEnd = endDate ? new Date(endDate) : new Date(maxDate);
    if (loopEnd > maxDate) loopEnd = new Date(maxDate);

    loopStart.setHours(0, 0, 0, 0);
    loopEnd.setHours(23, 59, 59, 999);

    let currentDate = new Date(loopStart);
    while (currentDate <= loopEnd) {
      const currentDow = daysArr[currentDate.getDay()];

      let isMatch = true;
      if (dow && typeof dow === 'string' && dow.trim() !== '') {
        if (!dow.includes(currentDow)) isMatch = false;
      }

      if (isMatch) {
        const eventTitle = "【予約不可】" + (reason ? reason : "");

        if (!startTime && !endTime) {
          cal.createAllDayEvent(eventTitle, currentDate);
          addedCount++;
        } else {
          let sHour = 0, sMin = 0, eHour = 23, eMin = 59;

          if (startTime instanceof Date) {
            sHour = startTime.getHours(); sMin = startTime.getMinutes();
          } else if (typeof startTime === 'string') {
            const parts = startTime.split(':');
            if (parts.length >= 2) { sHour = parseInt(parts[0], 10); sMin = parseInt(parts[1], 10); }
          }

          if (endTime instanceof Date) {
            eHour = endTime.getHours(); eMin = endTime.getMinutes();
          } else if (typeof endTime === 'string') {
            const parts = endTime.split(':');
            if (parts.length >= 2) { eHour = parseInt(parts[0], 10); eMin = parseInt(parts[1], 10); }
          }

          const evStart = new Date(currentDate);
          evStart.setHours(sHour, sMin, 0, 0);
          const evEnd = new Date(currentDate);
          evEnd.setHours(eHour, eMin, 0, 0);

          cal.createEvent(eventTitle, evStart, evEnd);
          addedCount++;
        }
      }
      currentDate.setDate(currentDate.getDate() + 1);
    }
  }

  // 実行結果をログおよびシステムログに残す
  const logMessage = `制限ルールのカレンダー反映が完了しました（作成件数: ${addedCount}件）`;
  console.log(logMessage);
  const logSheet = getSheet(SHEETS.SYSTEM_LOG);
  logSheet.appendRow([new Date(), "INFO", logMessage]);
}

/**
 * 毎日深夜にカレンダー反映関数を自動実行するトリガーをセットアップする関数
 */
function setupDailySyncTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncRulesToCalendar') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 毎日深夜（午前2時〜3時の間）に自動実行
  ScriptApp.newTrigger('syncRulesToCalendar')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
    
  console.log("毎日深夜の自動カレンダー同期トリガーを設定しました。");
}