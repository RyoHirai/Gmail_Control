function onOpen(e){
    var arr = [
        {name: "データクリア", functionName: "MailErase"},
        {name: "メール情報取得", functionName: "MailSearchContact"},
        {name: "【開発中】メール情報取得", functionName: "MailSearchContact_2"}
    ];
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    spreadsheet.addMenu("追加機能一覧", arr);
}


function MailErase(){
  try{
    var sheet = SpreadsheetApp.getActiveSheet();
    sheet.getRange('A3:F50').clear();   
  }catch(e){
    Browser.msgBox(e);
  }
}


function MailSearchContact() {
  
  /* Gmailから特定条件のスレッドを検索しメールを取り出す */
  var strTerms = Browser.inputBox("件名検索：キーワードを入力してください");
      strTerms = "'subject:'"+ strTerms
  var myThreads = GmailApp.search(strTerms, 0, 30); //条件にマッチしたスレッドを取得
  var myMsgs = GmailApp.getMessagesForThreads(myThreads); //スレッドからメールを取得する　→二次元配列で格納
 
  var valMsgs = [];
 
  /* 各メールから日時、送信元、件名、内容を取り出す*/
  for(var i = 0;i < myMsgs.length;i++){
 
    valMsgs[i] = [];
    valMsgs[i][1] = myMsgs[i][0].getDate();
    valMsgs[i][2] = myMsgs[i][0].getFrom();
    valMsgs[i][3] = myMsgs[i][0].getSubject();
    valMsgs[i][4] = myMsgs[i][0].getPlainBody();
 
  }
 
  /* スプレッドシートに出力 */
  if(myMsgs.length>0){
    
    SpreadsheetApp.getActiveSheet().getRange(3, 1, i, 5).setValues(valMsgs); //シートに貼り付け
    
  }
  
  Browser.msgBox("終了しました");
}




function MailSearchContact_2() {
  
  /* Gmailから特定条件のスレッドを検索しメールを取り出す */
  var strTerms = Browser.inputBox("件名検索：キーワードを入力してください");
      strTerms = "'subject:'"+ strTerms
  var myThreads = GmailApp.search(strTerms, 0, 30); //条件にマッチしたスレッドを取得
  var myMsgs = GmailApp.getMessagesForThreads(myThreads); //スレッドからメールを取得する　→二次元配列で格納
 
  var valMsgs = [];
 
  /* 各メールから日時、送信元、件名、内容を取り出す*/
  for(var i = 0;i < myMsgs.length;i++){
 
    valMsgs[i] = [];
    valMsgs[i][1] = myMsgs[i][0].getDate();
    valMsgs[i][2] = myMsgs[i][0].getFrom();
    valMsgs[i][3] = myMsgs[i][0].getSubject();
    valMsgs[i][4] = myMsgs[i][0].getPlainBody();
 
  }
 
  /* スプレッドシートに出力 */
  if(myMsgs.length>0){
    
    SpreadsheetApp.getActiveSheet().getRange(3, 1, i, 5).setValues(valMsgs); //シートに貼り付け
    
  }
  
  Browser.msgBox("終了しました");
}