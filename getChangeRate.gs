/**
 * Web Performanceの各値の変化率を計算
 * Web Performanceの値は別ファイルから読み込む
 */
function getChangeRate() {
  // スプレッドシートキー取得
  var sheetKey = getScriptProperty('sheetKey');
  // スプレッドシートキーがスクリプトプロパティに設定されていない場合、処理を終了する
  if (!sheetKey) {
    return;
  }
  //　読み取り用スプレッドシートの読み取り行・列取得
  var startRow = getScriptProperty('startRow');
  var startColumn = getScriptProperty('startColumn');

  // 読み取り用スプレッドシート、スプレッドシート内の全シートを取得
  var inputSpredSheet = SpreadsheetApp.openById(sheetKey);
  var inputSheets = inputSpredSheet.getSheets();
  // 書き込み用スプレッドシート、スプレッドシート内の全シートを取得
  var outputSpredSheet = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheets = outputSpredSheet.getSheets();

  for (var sheetIndex = 0; sheetIndex < inputSheets.length; sheetIndex++) {
    // 読み込み・書き込みを行うシートを設定
    var inputSheet = inputSheets[sheetIndex];
    var outputSheet = outputSheets[sheetIndex];
    // シートの最終行・列を取得
    var lastRow = inputSheet.getLastRow();
    var lastColumn = inputSheet.getLastColumn();

    // 日付配列データ、各項目データ配列の初期化
    var dateList = [];
    var itemsList = [];
    for (var column = startColumn; column < lastColumn; column++) {
      var valueList = [];
      var changeRateList = [];
      for (var row = startRow; row < lastRow; row++) {
        // 該当セルの生データ取得
        var value = inputSheet.getRange(row, column).getValue();
        // 変化率初期化
        var changeRate = '';
        // 一列目のセルには日付が格納されているため、日付格納用配列に格納する
        if (column == startColumn) {
          dateList.push([value]);
          continue;
        }

        // 前日x日間平均値から該当セルの変化率算出
        var beforeDay = 7;
        if (value && row > startRow + (beforeDay - 1)) {
          // 該当セルの前日〜指定したx日前までの値を取得
          var pastDaysValueList = inputSheet.getRange(row - beforeDay, column, beforeDay, 1).getValues();
          // 該当セルの前日〜指定したx日前までの値を算出
          var pastDaysSumValue = num = 0;
          for (var count = 0; count < pastDaysValueList.length; count++) {
            var pastDaysValue = pastDaysValueList[count][0];
            if (pastDaysValue) {
              pastDaysSumValue += pastDaysValue;
              num++;
            }
          }
          // 平均値算出
          var pastDaysAveValue = num > 0 ? pastDaysSumValue / num : '';
          // 前日x日間平均値から該当セルの変化率算出
          if (pastDaysAveValue) {
            var diffValue = (pastDaysAveValue - value) / pastDaysAveValue;
            // 小数からパーセントに直し、小数点第２で四捨五入する
            changeRate = Math.round(diffValue * 1000) / 10;
          }
        }

        // 生データを配列へ格納
        valueList.push(value);
        // 変化率を配列へ格納
        changeRateList.push(changeRate);
      }
      // 一列目の日付データ以外を各項目データ配列へ格納する
      if (column != startColumn) {
        itemsList.push(valueList);
        itemsList.push(changeRateList);
      }
    }
    // 日付データを書き込む
    outputSheet.getRange(2, 1, dateList.length, 1).setValues(dateList);
    // 各項目データを書き込む（書き込み前に配列形式を変換する）
    itemsList = convertArrayFormat(itemsList);
    outputSheet.getRange(2, 2, itemsList.length, itemsList[0].length).setValues(itemsList);
  }
}

/* 配列の形式を変換する
 * @param {array} values 変換配列
 * @return {array} items 変換後配列
 */
function convertArrayFormat(values) {
  var items = [];
  for (var key = 0; key < values[0].length; key++) {
    var item = [];
    for (var count = 0; count < values.length; count++) {
      item.push(values[count][key]);
    }
    items.push(item);
  }
  return items;
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
