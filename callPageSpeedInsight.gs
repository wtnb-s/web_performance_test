/**
 * ページの表示速度を測定する
 */
function insightPagespeed() {
    // API KEY取得
    var API_TOKEN_PAGESPEED = getScriptProperty("API_TOKEN_PAGESPEED");
    // 言語設定
    var locale = 'ja_JP';
    
    // シートに関する変数を設定
    // スプレッドシート、スプレッドシート内の全シートを取得
    var spredSheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = spredSheet.getSheets();
    // 何番目のシートか設定
    var sheetIndex = getScriptProperty("sheetIndex") ? 
        parseInt(getScriptProperty("sheetIndex")) : 0;
    var sheet = sheets[sheetIndex];  
    // 再起動時は処理を中断したURLの列から処理を実行
    var column = getScriptProperty("writtenColumn") ? 
        parseInt(getScriptProperty("writtenColumn")) : 2;
    // シートの最終行・列を取得
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    
    // PageSpeedInsightsAPIのリクエストに関わる変数設定
    // シート名の後ろ2文字を切り出してデバイスを取得
    var device = sheet.getName().substr(-2); 
    // デバイス別にクエリの値を取得
    var strategy = device === "PC" ?
      "desktop" : "mobile";
   
    // URL配列を現在の行から最後まで取得
    var urls = sheet.getRange(2, column, 1, lastColumn - column + 1).getValues();
  
    // 再起動用に開始時間を取得
    var start = dayjs.dayjs();
    // 今日の日時を1列目最終行に書き込む
    var today = dayjs.dayjs().format('MM-DD HH:mm');
    sheet.getRange(lastRow + 1, 1).setValue(today);
  
    var scores = [];
    // 取得した全URLに対して処理
    for (var i = 0; i < urls[0].length; i++) {
      // URLが空の場合はスキップ
      var url = urls[0][i];
      if (!url) {
        continue;
      } 
  
      // リクエストURLを作成
      var request = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url=" + 
        url + "&key=" + API_TOKEN_PAGESPEED + '&local=' + locale + "&strategy=" + strategy;
  
      // URLをAPIに投げてみてエラーが返ってくる場合はログに残す
      try {
        var response = UrlFetchApp.fetch(request, { muteHttpExceptions: true })
      } catch (err) {
        Logger.log(err)
        return err
      }
  
      // 返ってきたjsonをパース
      var parsedResult = Utilities.jsonParse(response.getContentText());
      var score = parsedResult['lighthouseResult']['audits']['speed-index'] ?
        parsedResult['lighthouseResult']['audits']['speed-index']['displayValue'] : "-";
      // ページスピードスコアをscores配列に追加
      scores.push(score);
  
      // 現在時間を取得して、開始から5分経過していたらforループ処理を中断して再起動
      var now = dayjs.dayjs();
      if (now.diff(start, "minutes") >= 5) {
        Logger.log("5分経過しました。タイムアウト回避のため処理を中断して再起動します。");
        break;
      }
      // 翌日のトリガー設定
      // setTrigger();
    }
    
    // 取得したスコアを一度に書き込む
    sheet.getRange(lastRow + 1, column, 1, scores.length).setValues([scores]);
  
    // columnを次の再起動用に設定
    var writtenColumn = column + scores.length;
    setScriptProperty("writtenColumn", writtenColumn);
    
    if (writtenColumn < lastColumn) {
      // 最終行まで処理していない場合は再起動し、途中から処理実行
      insightPagespeed();
    } else {
      // 最終行まで処理している場合は保存していた行を削除
      deleteScriptProperty("writtenColumn");
      // 次のシート記載のURLについて計測するため、インデックスをインクリメント
      sheetIndex++;
      if (sheetIndex < sheets.length) {
        // sheetIndexをスクリプトプロパティにセット後、再起動
        setScriptProperty("sheetIndex", sheetIndex);
        insightPagespeed();
      } else {
        // 最終シートまで処理した場合、スクリプトプロパティ削除
        deleteScriptProperty("sheetIndex");
      }
    }
  }
  
  // トリガー設定（定期実行用）
  function setTrigger() {
   var setTime = new Date();
    setTime.setDate(setTime.getDate() + 1)
    setTime.setHours(20);
    setTime.setMinutes(00); 
    ScriptApp.newTrigger('dateExecution').timeBased().at(setTime).create();  
  }
  
  // スクリプトプロパティゲッター
  function getScriptProperty(key){
    return PropertiesService.getScriptProperties().getProperty(key);
  }
  
  // スクリプトプロパティセッター
  function setScriptProperty(key, value) {
    return PropertiesService.getScriptProperties().setProperty(key, value);
  }
  
  // スクリプトプロパティデリート
  function deleteScriptProperty(key){
    return PropertiesService.getScriptProperties().deleteProperty(key);
  }
  