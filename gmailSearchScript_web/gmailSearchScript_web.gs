function clickOutputStartButton() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var mainSheet = SpreadsheetApp.getActiveSheet();
  //スプレッドシートの名前付き定義で取得
  var searchRange = mainSheet.getRange('searchQuery');

  var result = ui.alert(
    'gmailから以下の検索条件で出力します',
    searchRange.getValue(),
    ui.ButtonSet.OK_CANCEL);

  if (result == ui.Button.OK) {
    var startTime = new Date();
    //メール出力処理実行
    var ret = outputMail(searchRange.getValue());
    var endTime = new Date();
    mainSheet.getRange('outputTime').setValue((endTime - startTime) / 1000);
    /*
    //値が返ってきたかで判定
    if (ret != null) {
      var result = ui.alert(
        '完了',
        'gmailからの出力に成功しました',
        ui.ButtonSet.OK);
      //URLをセルにセット
      //mainSheet.getRange('outputUrl').setValue(ret);
      
    } else {
      var result = ui.alert(
        '警告',
        '検索条件に一致するメールはありませんでした',
        ui.ButtonSet.OK);
    }
    */
  }
}

function outputMail(searchQuery) {
  /* 自身のGmailからクエリ条件と一致するメール（スレッド）を全て取得する */
  var allThreads = [];
  var max = 500;
  var loopCount = SpreadsheetApp.getActiveSheet().getRange('loopCount').getValue();
  //スレッドが取得できなくなるまでループ
  for(var i = 0; i < loopCount; i++){
    //0,500で0~499のスレッドが取得されるので、次に取得するのは500から
    var threads = GmailApp.search(searchQuery, i*max, max);
    if (threads.length == 0) break;
    //取得したスレッドをオールに結合
    Array.prototype.push.apply(allThreads, threads);
  }
  
  //データ格納用変数
  var valMsgs = [];
  //見出し行セット
  var frozenRow = 2;
  valMsgs.push([searchQuery,'','','']);
  valMsgs.push(['Date','From','Subject','PlainBody']);
  
  for(var n in allThreads){
    var thread = allThreads[n];
    var msgs = thread.getMessages();
    
    for(m in msgs){
      var msg = msgs[m];
      
      var date = msg.getDate();
      var from = msg.getFrom();
      var subj = msg.getSubject();
      var body = msg.getPlainBody();

      valMsgs.push([date,from,subj,body]);
    }
  }
  
  /* スプレッドシートに出力 */
  if(allThreads.length > 0){
    var sheetName = SpreadsheetApp.getActiveSpreadsheet().getName();
    var dateFormat = "yyyy/MM/dd HH:mm:ss"
    var nowTime = Utilities.formatDate(new Date(), "JST", dateFormat);
    
    //自身のgoogleDrive内に新規スプレッドシートを作成
    //var newSheet = SpreadsheetApp.create(sheetName + " 出力結果 (" + nowTime + ")");
    //var activeSheet = newSheet.getActiveSheet();
    
    //mainスプレッドシートに新規シートを作成
    var activeSheet = SpreadsheetApp.getActive().insertSheet("出力結果 (" + nowTime + ")");
    
    activeSheet.setFrozenRows(frozenRow);  //見出し行設定
    activeSheet.getRange((1 + frozenRow), 1, (valMsgs.length - frozenRow)).setNumberFormat(dateFormat);  //日付列にフォーマット設定
    activeSheet.getRange(1, 1, valMsgs.length, 4).setValues(valMsgs);  //メールデータセット
    //PlainBody分割処理
    activeSheet.getRange(2, 4, valMsgs.length).splitTextToColumns(String.fromCharCode(10));
    //空白列削除
    for (var i = activeSheet.getLastColumn(); i > 0; i--) {
      if (activeSheet.getRange(1, i, valMsgs.length).isBlank()) {
        activeSheet.deleteColumn(i);
      }
    }

    //return newSheet.getUrl();
  }
}

//グローバル変数
var gallThreads = [];
var gvalMsgs = [];
var gfrozenRow = 2;

function outputMail2(searchQuery) {
  /* 自身のGmailからクエリ条件と一致するメール（スレッド）を全て取得する */
  var allThreads = [];
  var max = 500;
  var loopCount = 20;
  //スレッドが取得できなくなるまでループ
  for(var i = 0; i < loopCount; i++){
    //0,500で0~499のスレッドが取得されるので、次に取得するのは500から
    var threads = GmailApp.search(searchQuery, i*max, max);
    if (threads.length == 0) break;
    //取得したスレッドをオールに結合
    Array.prototype.push.apply(allThreads, threads);
  }
  gallThreads = allThreads;
  
  //見出し行セット
  gvalMsgs.push([searchQuery,'','','']);
  gvalMsgs.push(['Date','From','Subject','PlainBody']);
  
  return gallThreads.length;
}

function dataSet(count) {
  var valMsgs = [];
  
  var thread = gallThreads[count];
  var msgs = thread.getMessages();
  for(m in msgs){
    var msg = msgs[m];
    
    var date = msg.getDate();
    var from = msg.getFrom();
    var subj = msg.getSubject();
    var body = msg.getPlainBody();
      
    valMsgs.push([date,from,subj,body]);
  }
  
  Array.prototype.push.apply(gvalMsgs, valMsgs);
}

function outputSheet2(valMsgs) {
  /* スプレッドシートに出力 */
  if(valMsgs.length > 0){
    var sheetName = SpreadsheetApp.getActiveSpreadsheet().getName();
    var dateFormat = "yyyy/MM/dd HH:mm:ss"
    var nowTime = Utilities.formatDate(new Date(), "JST", dateFormat);
    
    //自身のgoogleDrive内に新規スプレッドシートを作成
    //var newSheet = SpreadsheetApp.create(sheetName + " 出力結果 (" + nowTime + ")");
    //var activeSheet = newSheet.getActiveSheet();
    
    //mainスプレッドシートに新規シートを作成
    var activeSheet = SpreadsheetApp.getActive().insertSheet("出力結果 (" + nowTime + ")");
    
    activeSheet.setFrozenRows(frozenRow);  //見出し行設定
    activeSheet.getRange((1 + frozenRow), 1, (valMsgs.length - frozenRow)).setNumberFormat(dateFormat);  //日付列にフォーマット設定
    activeSheet.getRange(1, 1, valMsgs.length, 4).setValues(valMsgs);  //メールデータセット
    //PlainBody分割処理
    activeSheet.getRange(2, 4, valMsgs.length).splitTextToColumns(String.fromCharCode(10));
    //空白列削除
    for (var i = activeSheet.getLastColumn(); i > 0; i--) {
      if (activeSheet.getRange(1, i, valMsgs.length).isBlank()) {
        activeSheet.deleteColumn(i);
      }
    }

    //return newSheet.getUrl();
  }
}
