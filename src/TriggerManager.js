/**
 * トリガー（定期実行）の管理を行うスクリプト
 */

/**
 * 指定した関数名に紐づくトリガーをすべて削除する関数
 * （トリガーの重複登録を防ぐために使用します）
 * @param {string} functionName - 削除対象の関数名
 */
function deleteTriggerByName(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * メールキュー処理（processMailQueue）を5分おきに実行するトリガーを設定する関数
 * ※この関数をGASエディタから手動で1度だけ実行してください。
 */
function setMailTrigger() {
  const functionName = "processMailQueue";
  
  deleteTriggerByName(functionName);
  
  ScriptApp.newTrigger(functionName).timeBased().everyMinutes(5).create();
  
  Logger.log("メール自動配信のトリガー（5分おき）をセットしました。");
}