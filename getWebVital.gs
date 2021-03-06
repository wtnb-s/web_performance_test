/**
 * Web Performance(core web vital)を測定
 * ラボデータは５回取得し、外れ値考慮のため、配列中の最小値と最大値を取り除いた上で平均を計算
 */
function getWebVital() {
  // スプレッドシート、スプレッドシート内の全シートを取得
  var spredSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spredSheet.getSheets();
  // 書き込みを行うシートを設定
  var sheetIndex = getScriptProperty('sheetIndex') ? parseInt(getScriptProperty('sheetIndex')) : 0;

  // 再起動用に開始時間を取得
  var start = dayjs.dayjs();

  for (sheetIndex; sheetIndex < sheets.length; sheetIndex++) {
    // 強制終了してもスクリプトが最後まで実行されるように再起動用のタイマーをセット
    setTriggerAfterMinutes('getWebVital', 10);

    var sheet = sheets[sheetIndex];
    // シートの最終行を取得
    var row = sheet.getLastRow() + 1;
    // デバイス設定取得
    var device = sheet.getRange('B1').getValue();
    var strategy = device === 'PC' ? 'desktop' : 'mobile';
    // URL取得
    var url = sheet.getRange('D1').getValue();
    // 計測するURLを出力
    Logger.log('シート：' + (sheetIndex + 1) + '/' + sheets.length + ' 計測対象：' + url + '; ' + strategy);

    // Web Vitalデータ取得
    var values = [];
    for (var getCount = 0; getCount < 5; getCount++) {
      // API呼び出し
      values.push(callPageSpeedInsightsApi(url, strategy));
    }
    // データ変換（values['getCount']['key'] → items['key']['getCount']）
    var items = convertArrayFormat(values);
    // 各要素毎に平均値計算（外れ値考慮のため、配列中の最小値と最大値を取り除いた上で計算する）
    var aveList = getAverage(items);

    // 1列目に書き込み日時を書き込む
    var today = dayjs.dayjs().format('MM-DD HH:mm');
    sheet.getRange(row, 1).setValue(today);
    // 取得したスコアを書き込む
    sheet.getRange(row, 2, 1, aveList.length).setValues([aveList]);

    // 書き込みが完了したら、インクリメントしたsheetIndexを保存
    setScriptProperty('sheetIndex', sheetIndex + 1);

    // 現在時間を取得して、開始から3分経過していたらforループ処理を中断して再起動
    var now = dayjs.dayjs();
    if (now.diff(start, 'minutes') >= 2 && sheetIndex + 1 < sheets.length) {
      Logger.log('2分経過しました。タイムアウト回避のため処理を中断して再起動します');
      break;
    }
  }

  if (sheetIndex < sheets.length) {
    // 最終シートまで処理していない場合は再起動し、途中から処理実行
    setTriggerAfterMinutes('getWebVital', 2);
  } else {
    // 最終シートまで処理した場合、スクリプトプロパティ削除
    deleteScriptProperty('sheetIndex');

    // トリガーを起動する時刻(h)を取得
    var startHour = getScriptProperty('START_HOUR') ? parseInt(getScriptProperty('START_HOUR')) : 21;
    setTriggerDaily('getWebVital', startHour);

    Logger.log('処理を終了します');
  }
}

/*
 * PageSpeed Insight API呼び出し
 * @param {string} url 測定する対象URL
 * @param {string} strategy デバイスタイプ
 * @return {array} Web Vital data
 */
function callPageSpeedInsightsApi(url, strategy) {
  // API KEY取得
  var API_TOKEN_PAGESPEED = getScriptProperty('API_TOKEN_PAGESPEED');
  // 言語設定
  var locale = 'ja_JP';
  // リクエストURLを作成
  var request =
    'https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url=' + url + '&key=' + API_TOKEN_PAGESPEED + '&local=' + locale + '&strategy=' + strategy;

  // URLをAPIに投げ、エラーが返ってくる場合ログに残す
  try {
    var response = UrlFetchApp.fetch(request, { muteHttpExceptions: true });
  } catch (err) {
    Logger.log(err);
    return [];
  }
  // 返ってきたjsonをパース
  var parsedResult = Utilities.jsonParse(response.getContentText());
  var values = [];
  try {
    // フィールドデータ
    // First Contentful Paint(ms)
    values.push(parsedResult['loadingExperience'] ? parsedResult['loadingExperience']['metrics']['FIRST_CONTENTFUL_PAINT_MS']['percentile'] : '');
    // Largest Contentful Paint(ms)
    values.push(parsedResult['loadingExperience'] ? parsedResult['loadingExperience']['metrics']['LARGEST_CONTENTFUL_PAINT_MS']['percentile'] : '');
    // First Input Delay(ms)
    values.push(parsedResult['loadingExperience'] ? parsedResult['loadingExperience']['metrics']['FIRST_INPUT_DELAY_MS']['percentile'] : '');
    // Cumulative Layout Shift(単位なし)
    values.push(parsedResult['loadingExperience'] ? parsedResult['loadingExperience']['metrics']['CUMULATIVE_LAYOUT_SHIFT_SCORE']['percentile'] : '');

    // ラボデータ
    // スコア(単位なし)
    var score = parsedResult['lighthouseResult'] ? parsedResult['lighthouseResult']['categories']['performance']['score'] : '';
    values.push(score * 100);
    // First Contentful Paint(Lab)(ms)
    values.push(parsedResult['lighthouseResult'] ? parsedResult['lighthouseResult']['audits']['first-contentful-paint']['numericValue'] : '');
    // Speed Index(Lab)(ms)
    values.push(parsedResult['lighthouseResult'] ? parsedResult['lighthouseResult']['audits']['speed-index']['numericValue'] : '');
    // Largest Contentful Paint(ms)
    values.push(parsedResult['lighthouseResult'] ? parsedResult['lighthouseResult']['audits']['largest-contentful-paint']['numericValue'] : '');
    // Total Blocking Time(ms)
    values.push(parsedResult['lighthouseResult'] ? parsedResult['lighthouseResult']['audits']['total-blocking-time']['numericValue'] : '');
    // Cumulative Layout Shift(単位なし)
    var cls = parsedResult['lighthouseResult'] ? parsedResult['lighthouseResult']['audits']['cumulative-layout-shift']['numericValue'] : '';
    values.push(cls * 100);
  } catch (err) {
    Logger.log(err);
  } finally {
    return values;
  }
}

/*
 * 配列中の要素ごとに変換する
 * @param {array} values 変換配列
 * @return {array} items 変換後配列
 */
function convertArrayFormat(values) {
  var items = [];
  for (var key = 0; key < values[0].length; key++) {
    var item = [];
    for (var getCount = 0; getCount < values.length; getCount++) {
      // APIの取得に失敗している可能性を考慮して値が入っているか判定する
      if (values[getCount][key]) {
        item.push(values[getCount][key]);
      }
    }
    items.push(item);
  }
  return items;
}

/*
 * 各要素毎に平均値計算（外れ値考慮のため、配列中の最小値と最大値を取り除いた上で計算する）
 * @param {array} items 要素データの格納された配列データ
 * @return {array} averageList 要素毎の平均値データの配列
 */
function getAverage(items) {
  aveList = [];
  for (var key = 0; key < items.length; key++) {
    Logger.log(items[key].join(','));
    var maxIndex = items[key].indexOf(Math.max.apply(null, items[key]));
    var minIndex = items[key].indexOf(Math.min.apply(null, items[key]));
    items[key].splice(maxIndex, 1);
    items[key].splice(minIndex, 1);
    // 平均値を算出し、四捨五入する
    ave =
      items[key].reduce(function (pre, curr, i) {
        return pre + curr;
      }, 0) / items[key].length;

    aveList.push(Math.round(ave));
  }
  return aveList;
}

/*
 * トリガー設定（設定した分後）
 * @param {string} functionName 対象関数名
 * @param {int} minutes 何分後にトリガーを実行するか
 */
function setTriggerAfterMinutes(functionName, minutes) {
  // 同名のトリガーを削除
  deleteTrigger(functionName);
  var setTime = new Date();
  setTime.setMinutes(setTime.getMinutes() + minutes);
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
  setTime.setDate(setTime.getDate() + 1);
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
function getScriptProperty(key) {
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
function deleteScriptProperty(key) {
  PropertiesService.getScriptProperties().deleteProperty(key);
}
