/**
 * Web Performanceを測定
 * 測定する要素はメソッド(callPageSpeedInsightsApi)の変数内で指定すること
 */
function testWebPerformance() {
    // トリガーを起動する時刻(h)を取得
    var startHour = getScriptProperty('START_HOUR') ?
        parseInt(getScriptProperty('START_HOUR')) : 20;
    // 実行している関数名を取得
    var thisFunctionName =  arguments.callee.name;
    // スプレッドシート、スプレッドシート内の全シートを取得
    var spredSheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = spredSheet.getSheets();
    // 何番目のシートか設定
    var sheetIndex = getScriptProperty('sheetIndex') ? 
        parseInt(getScriptProperty('sheetIndex')) : 0;
    var sheet = sheets[sheetIndex];  
    // シートの最終行・列を取得
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    // 書き込み開始列の設定（再起動時は処理を中断したURLの列から処理を実行）
    var column = getScriptProperty('writtenColumn') ? 
        parseInt(getScriptProperty('writtenColumn')) : 2;
    // 書き込み開始行設定（対象シートへの書き込みが初回の場合、最終行＋１の行に書き込みを行う）
    var row = (column == 2) ? lastRow + 1 : lastRow;
    
    // シート名の後ろ2文字を切り出してデバイスを取得
    var device = sheet.getName().substr(-2); 
    // デバイス別にクエリの値を取得
    var strategy = (device === 'PC') ? 'desktop' : 'mobile';
   
    // 再起動用に開始時間を取得
    var start = dayjs.dayjs();
    // 対象シートへの書き込みが初回の場合のみ、日時を1列目に書き込む
    if (column == 2) {
      var today = dayjs.dayjs().format('MM-DD HH:mm');
      sheet.getRange(row, 1).setValue(today);
    }

    // URL配列を現在の行から最後まで取得
    var urls = sheet.getRange(2, column, 1, lastColumn - column + 1).getValues();
    var aveScoreList = [];
    for (var urlCount = 0; urlCount < urls[0].length; urlCount++) {
      // URLが空の場合はスキップ
      var url = urls[0][urlCount];
      if (!url) continue;
      // 計測するURLを出力
      Logger.log('計測対象：' + url + '; ' + strategy);

      // 一つのURLに対して複数回スコアを取得し、平均値を算出する
      var sumScore = 0, num = 0;
      for (var count = 0; count < 3; count++) {
        // API呼び出し
        var score = callPageSpeedInsightsApi(url, strategy);
        if (score) {
          Logger.log('計測' + (count + 1) + '回目のスコア：' + score);
          sumScore = sumScore + score;
          num++;
        }
      }

      // 小数点だ２位で四捨五入した上で平均値を算出
      var aveScore = (num > 0) ? Math.round(sumScore / num * 10) / 10 : '-';
      aveScoreList.push(aveScore);
      Logger.log('##平均値##');
      Logger.log(aveScore);
  
      // 現在時間を取得して、開始から4分経過していたらforループ処理を中断して再起動
      var now = dayjs.dayjs();
      if (now.diff(start, 'minutes') >= 4) {
        Logger.log('4分経過しました。タイムアウト回避のため処理を中断して再起動します');
        break;
      }
    }
    
    // 取得したスコアを一度に書き込む
    sheet.getRange(row, column, 1, aveScoreList.length).setValues([aveScoreList]);
  
    // columnを次の再起動用に設定
    var writtenColumn = column + aveScoreList.length;
    setScriptProperty('writtenColumn', writtenColumn);
    
    if (writtenColumn <= lastColumn) {
      // 最終行まで処理していない場合は再起動し、途中から処理実行
      setTriggerAfterMinitues(thisFunctionName, 2);
    } else {
      // 最終行まで処理している場合は保存していた行を削除
      deleteScriptProperty('writtenColumn');

      // 次のシート記載のURLについて計測するため、インデックスをインクリメント
      sheetIndex++;
      if (sheetIndex < sheets.length) {
        // 処理するシートが残っている場合、再起動し次のシートを処理する
        setScriptProperty('sheetIndex', sheetIndex);

        Logger.log('次のシートの処理を実施するため、再起動します');
        setTriggerAfterMinitues(thisFunctionName, 2);
      } else {
        // 最終シートまで処理した場合、スクリプトプロパティ削除
        deleteScriptProperty('sheetIndex');
        setTriggerDaily(thisFunctionName, startHour);
      }
    }
  }
  
  /*
   * PageSpeed Insight API呼び出し
   * @param {string} url 測定する対象URL
   * @aram {string} strategy デバイスタイプ
   * @return {int | string} speed index score
  */
  function callPageSpeedInsightsApi(url, strategy) {
      // API KEY取得
      var API_TOKEN_PAGESPEED = getScriptProperty('API_TOKEN_PAGESPEED');
      // 言語設定
      var locale = 'ja_JP';
      // 測定する要素
      var testValue = 'speed-index';

      // リクエストURLを作成
      var request = 'https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url=' + 
        url + '&key=' + API_TOKEN_PAGESPEED + '&local=' + locale + '&strategy=' + strategy;
  
      // URLをAPIに投げてみてエラーが返ってくる場合はログに残す
      try {
        var response = UrlFetchApp.fetch(request, { muteHttpExceptions: true })
      } catch (err) {
        Logger.log(err)
        return err
      }
      // 返ってきたjsonをパース
      var parsedResult = Utilities.jsonParse(response.getContentText());
      // speedIndexのスコアを取得
      var score = parsedResult['lighthouseResult'] ? 
        parsedResult['lighthouseResult']['audits'][testValue]['displayValue'] : '';
      // 単位の「s」を削除して返却する
      return parseFloat(score.replace('s', ''));
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
  