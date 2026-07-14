/**
 * Googleカレンダーとの同期処理を管理するサービス
 */

const TARGET_CALENDAR_ID = "seigyo.reserve.system@gmail.com";

function syncCalendarEvent(reservationId) {
  const resSheet = getSheet(SHEETS.RESERVATIONS);
  const resData = resSheet.getDataRange().getValues();
  let resRow = null;
  
  for (let i = 1; i < resData.length; i++) {
    if (resData[i][0] === reservationId) {
      resRow = resData[i];
      break;
    }
  }
  
  if (!resRow || resRow[7] === "Cancelled") return;
  
  const userName = resRow[2];
  const date = Utilities.formatDate(new Date(resRow[3]), Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  let startTime = resRow[4];
  if (startTime instanceof Date) startTime = Utilities.formatDate(startTime, Session.getScriptTimeZone(), "HH:mm");
  
  let endTime = resRow[5];
  if (endTime instanceof Date) endTime = Utilities.formatDate(endTime, Session.getScriptTimeZone(), "HH:mm");
  
  const purpose = resRow[6];
  const startDateTime = new Date(`${date}T${startTime}:00`);
  const endDateTime = new Date(`${date}T${endTime}:00`);
  
  const syncSheet = getSheet(SHEETS.CALENDAR_SYNC);
  const syncData = syncSheet.getDataRange().getValues();
  let syncRowIndex = -1;
  let calendarEventId = "";
  
  for (let i = 1; i < syncData.length; i++) {
    if (syncData[i][0] === reservationId) {
      syncRowIndex = i + 1;
      calendarEventId = syncData[i][1];
      break;
    }
  }
  
  const calendar = CalendarApp.getCalendarById(TARGET_CALENDAR_ID);
  if (!calendar) {
    throw new Error(`カレンダーID (${TARGET_CALENDAR_ID}) が見つからないか、アクセス権限がありません。`);
  }
  
  const timestamp = new Date();
  const title = `【予約】${userName} (${purpose})`;
  
  try {
    if (calendarEventId) {
      const event = calendar.getEventById(calendarEventId);
      if (event) {
        event.setTitle(title);
        event.setTime(startDateTime, endDateTime);
        syncSheet.getRange(syncRowIndex, 3).setValue("SUCCESS");
        syncSheet.getRange(syncRowIndex, 4).setValue(timestamp);
      } else {
        const newEvent = calendar.createEvent(title, startDateTime, endDateTime);
        syncSheet.getRange(syncRowIndex, 2).setValue(newEvent.getId());
        syncSheet.getRange(syncRowIndex, 3).setValue("SUCCESS");
        syncSheet.getRange(syncRowIndex, 4).setValue(timestamp);
      }
    } else {
      const newEvent = calendar.createEvent(title, startDateTime, endDateTime);
      syncSheet.appendRow([reservationId, newEvent.getId(), "SUCCESS", timestamp]);
    }
  } catch (error) {
    const systemLogSheet = getSheet(SHEETS.SYSTEM_LOG);
    systemLogSheet.appendRow([timestamp, "ERROR", `カレンダー同期失敗 (ReservationID: ${reservationId}): ${error.message}`]);
    if (calendarEventId && syncRowIndex !== -1) {
      syncSheet.getRange(syncRowIndex, 3).setValue("ERROR");
      syncSheet.getRange(syncRowIndex, 4).setValue(timestamp);
    } else if (!calendarEventId) {
      syncSheet.appendRow([reservationId, "", "ERROR", timestamp]);
    }
  }
}

function deleteCalendarEvent(reservationId) {
  const syncSheet = getSheet(SHEETS.CALENDAR_SYNC);
  const syncData = syncSheet.getDataRange().getValues();
  let syncRowIndex = -1;
  let calendarEventId = "";
  
  for (let i = 1; i < syncData.length; i++) {
    if (syncData[i][0] === reservationId) {
      syncRowIndex = i + 1;
      calendarEventId = syncData[i][1];
      break;
    }
  }
  
  if (!calendarEventId || syncRowIndex === -1) return;
  
  const calendar = CalendarApp.getCalendarById(TARGET_CALENDAR_ID);
  if (!calendar) return;
  
  const timestamp = new Date();
  
  try {
    const event = calendar.getEventById(calendarEventId);
    if (event) {
      event.deleteEvent();
    }
    syncSheet.getRange(syncRowIndex, 3).setValue("CANCELLED");
    syncSheet.getRange(syncRowIndex, 4).setValue(timestamp);
  } catch (error) {
    const systemLogSheet = getSheet(SHEETS.SYSTEM_LOG);
    systemLogSheet.appendRow([timestamp, "ERROR", `カレンダー削除失敗 (ReservationID: ${reservationId}): ${error.message}`]);
    syncSheet.getRange(syncRowIndex, 3).setValue("ERROR");
    syncSheet.getRange(syncRowIndex, 4).setValue(timestamp);
  }
}