/**
 * Web Performance(core web vital)を測定
 */
function getWebVital() {
  // 実行している関数名を取得
  var thisFunctionName =  arguments.callee.name;
  // スプレッドシート、スプレッドシート内の全シートを取得
  var spredSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spredSheet.getSheets();
  // 書き込みを行うシートを設定
  var sheetIndex = getScriptProperty('sheetIndex') ? 
      parseInt(getScriptProperty('sheetIndex')) : 0;

  // 再起動用に開始時間を取得
  var start = dayjs.dayjs();

  for (sheetIndex; sheetIndex < sheets.length; sheetIndex++) {
    var sheet = sheets[sheetIndex];  
    // シートの最終行を取得
    var row = sheet.getLastRow() + 1;
    // デバイス設定取得
    var device = sheet.getRange('B1').getValue();
    var strategy = (device === 'PC') ? 'desktop' : 'mobile';
    // URL取得
    var url = sheet.getRange('D1').getValue();
    // 計測するURLを出力
    Logger.log('シート：' + (sheetIndex + 1) + '/' + sheets.length + ' 計測対象：' + url + '; ' + strategy);
    
    // 1列目に書き込み日時を書き込む
    var today = dayjs.dayjs().format('MM-DD HH:mm');
    sheet.getRange(row, 1).setValue(today);
    // API呼び出し
    var values = callPageSpeedInsightsApi(url, strategy);
    // valuesが空だった場合、再度APIを呼び出す
    if (values.length == 0) {
      values = callPageSpeedInsightsApi(url, strategy);
    }
    
    // 取得したスコアを書き込む
    sheet.getRange(row, 2, 1, values.length).setValues([values]);
  
    // 現在時間を取得して、開始から4分経過していたらforループ処理を中断して再起動
    var now = dayjs.dayjs();
    if (now.diff(start, 'minutes') >= 4 && (sheetIndex + 1) < sheets.length) {
      Logger.log('4分経過しました。タイムアウト回避のため処理を中断して再起動します');
      break;
    }
  }
  
  if (sheetIndex < sheets.length) {
    // 最終シートまで処理していない場合は再起動し、途中から処理実行
    setScriptProperty('sheetIndex', sheetIndex + 1);
    setTriggerAfterMinitues(thisFunctionName, 2);
  } else {
    // 最終シートまで処理した場合、スクリプトプロパティ削除
    deleteScriptProperty('sheetIndex');

    // トリガーを起動する時刻(h)を取得
    var startHour = getScriptProperty('START_HOUR') ?
      parseInt(getScriptProperty('START_HOUR')) : 21;
    setTriggerDaily(thisFunctionName, startHour);

    Logger.log('処理を終了します');
  }
}

/*
 * PageSpeed Insight API呼び出し
 * @param {string} url 測定する対象URL
 * @aram {string} strategy デバイスタイプ
 * @return {array} Web Vital data
*/
function callPageSpeedInsightsApi(url, strategy) {
  // API KEY取得
  var API_TOKEN_PAGESPEED = getScriptProperty('API_TOKEN_PAGESPEED');
  // 言語設定
  var locale = 'ja_JP';
  // リクエストURLを作成
  var request = 'https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url=' + 
    url + '&key=' + API_TOKEN_PAGESPEED + '&local=' + locale + '&strategy=' + strategy;

  // URLをAPIに投げてみてエラーが返ってくる場合はログに残す
  try {
    var response = UrlFetchApp.fetch(request, { muteHttpExceptions: true })
  } catch (err) {
    Logger.log(err);
    return err;
  }
  // 返ってきたjsonをパース
  var parsedResult = Utilities.jsonParse(response.getContentText());
  var values = [];
  // フィールドデータ
  // First Contentful Paint
  values.push(parsedResult['loadingExperience'] ?
    parsedResult['loadingExperience']['metrics']['FIRST_CONTENTFUL_PAINT_MS']['percentile'] : '');
  // Largest Contentful Paint
  values.push(parsedResult['loadingExperience'] ?
    parsedResult['loadingExperience']['metrics']['LARGEST_CONTENTFUL_PAINT_MS']['percentile'] : '');
  // First Input Delay
  values.push(parsedResult['loadingExperience'] ?
    parsedResult['loadingExperience']['metrics']['FIRST_INPUT_DELAY_MS']['percentile'] : '');
  // Cumulative Layout Shift
  values.push(parsedResult['loadingExperience'] ?
    parsedResult['loadingExperience']['metrics']['CUMULATIVE_LAYOUT_SHIFT_SCORE']['percentile'] : '');

  // ラボデータ
  // スコア
  values.push(parsedResult['lighthouseResult'] ?
    parsedResult['lighthouseResult']['categories']['performance']['score'] : '');
  // First Contentful Paint(Lab)
  values.push(parsedResult['lighthouseResult'] ? 
    parsedResult['lighthouseResult']['audits']['first-contentful-paint']['displayValue'] : '');
  // Speed Index(Lab)
  values.push(parsedResult['lighthouseResult'] ? 
    parsedResult['lighthouseResult']['audits']['speed-index']['displayValue'] : '');
  // Largest Contentful Paint
  values.push(parsedResult['lighthouseResult'] ? 
    parsedResult['lighthouseResult']['audits']['largest-contentful-paint']['displayValue'] : '');
  // Total Blocking Time
  values.push(parsedResult['lighthouseResult'] ? 
    parsedResult['lighthouseResult']['audits']['total-blocking-time']['displayValue'] : '');
  // Cumulative Layout Shift
  values.push(parsedResult['lighthouseResult'] ? 
    parsedResult['lighthouseResult']['audits']['cumulative-layout-shift']['displayValue'] : '');

  return values;
}

/*
 * トリガー設定（設定した分後）
 * @param {string} functionName 対象関数名
 * @param {int} minitues 何分後にトリガーを実行するか
*/
function setTriggerAfterMinitues(functionName, minitues) {
   // 同名のトリガーを削除 
   deleteTrigger(functionName);
   var setTime = new Date();
   setTime.setMinutes(setTime.getMinutes() + minitues);
   ScriptApp.newTrigger(functionName).timeBased().at(setTime).create();

   Logger.log('トリガーの設定日時：' + setTime);
}

/*
 * トリガー設定（定期実行用）
 * @param {string} functionName 対象関数名
 * @param {int} hour トリガーを起動する時刻(h)を設定
*/
function setTriggerDaily(functionName, hour) {
  // 同名のトリガーを削除 
  deleteTrigger(functionName);
  var setTime = new Date();
  setTime.setDate(setTime.getDate() + 1)
  setTime.setHours(hour);
  setTime.setMinutes(00); 
  ScriptApp.newTrigger(functionName).timeBased().at(setTime).create();

  Logger.log('トリガーの設定日時：' + setTime);
}

/*
 * トリガー削除
 * @param {string} functionName 対象関数名
*/
function deleteTrigger(functionName) {
  // gets all installable triggers associated with the current project
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    // トリガー取得
    var trigger = triggers[i];
    
    // トリガーが起動したときに呼び出される関数が指定した関数と一致する場合、対象のトリガー削除
    if (trigger.getHandlerFunction() == functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

/*
 * スクリプトプロパティ取得
 * @param {string} key スクリプトプロパティ キー
 * @return {string} スクリプトプロパティ 値
*/
function getScriptProperty(key){
  return PropertiesService.getScriptProperties().getProperty(key);
}

/*
 * スクリプトプロパティ設定
 * @param {string} key スクリプトプロパティ キー
 * @param {string} value スクリプトプロパティ 値
*/
function setScriptProperty(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
}

/*
 * スクリプトプロパティ削除
 * @param {string} key スクリプトプロパティ キー
*/
function deleteScriptProperty(key){
  PropertiesService.getScriptProperties().deleteProperty(key);
}
