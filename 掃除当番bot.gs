// グローバル変数として過去の選出結果を管理する配列を定義




//1
//月の終わりで通知する月のペア
function sendMonthlyPairsToSlackIfFirstDayOfMonth() {
  //月曜日の回数分のペアを作成
  let pairs = pickPairsForMonth();   //2
  notifyPairsToSlack(pairs);
  backupPairsToSpreadsheet(pairs);   //10
  
}

//2
//月曜日の回数分のペアを作成
function pickPairsForMonth() {
  let roster = getRosterFromSpreadsheet();  //12
  
  //すべてまとめるための関数　例:｛ [Aさん,Bさん], [Cさん,Dさん] ,・・｝
  let pairs = pickTwoRandomPeople(roster);  //3


  return pairs;
}

//3
//ペアを選出するための関数
function pickTwoRandomPeople(roster) {

  //ペアを入れるための配列　例: [Aさん,Bさん]
  let selectedPeople = [];
  let pastSelections = [];


  //すべてまとめるための関数　例:｛ [Aさん,Bさん], [Cさん,Dさん] ,・・｝
  let pairs = []; // ペア配列

  let monthMondaysCount = CountMonday();
  console.log(monthMondaysCount);

  //二人組のペアを作って
  for (let i = 0; i < monthMondaysCount; i++) {
    while (selectedPeople.length < 2) {
      // 過去の選出結果がある場合、選出されていない人だけを roster から抽出
      let availableRoster = roster.filter(person => !pastSelections.includes(person));

      // roster が空の場合、選出履歴をリセット
      if (availableRoster.length === 0) {
        pastSelections = [];
        availableRoster = getRosterFromSpreadsheet(); // roster を元の状態に戻す
      }

      let randomIndex = Math.floor(Math.random() * availableRoster.length);
      let selectedPerson = availableRoster[randomIndex];
      selectedPeople.push(selectedPerson);
      pastSelections.push(selectedPerson); // 選出された人を過去の選出結果に追加
    }
    pairs.push(selectedPeople);
    selectedPeople = [];
  }
  console.log(pairs);
  return pairs;
}

//4
//祝日の配列を返す関数
function getHolidays() {
  // 現在の日付を取得
  const now = new Date();
  const today =  new Date(now.getFullYear() - 1, now.getMonth(), now.getDate()); // 1年後の日付
  // Google Calendar APIの呼び出し
  const calendarId = 'ja.japanese#holiday@group.v.calendar.google.com'; // 使用するカレンダーのID
  const calendar = CalendarApp.getCalendarById(calendarId);
  
  const oneYearLater = new Date(now.getFullYear() + 1, now.getMonth(), now.getDate()); // 1年後の日付
  const events = calendar.getEvents(today, oneYearLater); // 現在から1か月後までのイベントを取得

  // 祝日の日付を格納する配列
  const holidays = [];

  // 各イベント（祝日）について繰り返し処理
  for (const event of events) {
    // 祝日の日付を取得し、日付オブジェクトとして配列に追加
    holidays.push(event.getStartTime());
  }
  
  return holidays;
}

//5
// 日付が祝日かどうかを判定する関数
function isHoliday(date) {
  const holidays = getHolidays();

  for (let i = 0; i < holidays.length; i++) {
    if (isSameDate(date, holidays[i])) {
      return true;
    }
  }
  return false;
}

//6
// 2つの日付が同じ日付かどうかを判定する関数
function isSameDate(date1, date2) {
  return date1.getFullYear() === date2.getFullYear() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getDate() === date2.getDate();
}

//7
//来月に月曜日が何回あるか返す関数
function CountMonday(){
  // 来月の月曜日の回数を計算
  let today = new Date();
  let nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1);

  let firstMondayDate = new Date(nextMonth.getFullYear(), nextMonth.getMonth(), 1);

  while (firstMondayDate.getDay() !== 1) {
    firstMondayDate.setDate(firstMondayDate.getDate() + 1);
  }


  let monthMondaysCount = Math.floor((new Date(nextMonth.getFullYear(), nextMonth.getMonth() + 1, 0).getDate() - firstMondayDate.getDate()) / 7) + 1;

  return monthMondaysCount

}

//8
//次の月曜日が何日か返してくれる関数
function NextMonday(){
  let today = new Date();
  let nextMonday = new Date(today);

  // 今日が月曜日でない場合、次の月曜日まで日付を進める
  while (nextMonday.getDay() !== 1) {
    nextMonday.setDate(nextMonday.getDate() + 1);
  }
  console.log(nextMonday);
  return nextMonday;
}


//9
//来月の第一月曜日を返す関数
function NextMonthFirstMonday(){
  let today = new Date();
  let nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1);
  let firstMondayDate = new Date(nextMonth.getFullYear(), nextMonth.getMonth(), 1);

  while (firstMondayDate.getDay() !== 1) {
    firstMondayDate.setDate(firstMondayDate.getDate() + 1);
  }
  console.log(firstMondayDate);
  return firstMondayDate;
}

//10
//
function notifyPairsToSlack(pairs) {
  //来月の月の番号を取得(0スタート＋来月なので+2にした)
  let month = new Date().getMonth() + 2 ;
  let month_English = ["January","February","March","April","May","June","July","August","September","October","November","December" ];

  //掃除当番の日付リスト(基本月曜日祝日はずらす)を日本語verと英語verで取得
  const [messageHeaders,messageHeaders_English] = MondayList();

  //let url = 'https://hooks.slack.com/services/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';     //旧鯖
  let url = 'https://hooks.slack.com/services/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';


  
  console.log(messageHeaders);
  postMessageToSlack(url,"----------------Japanese---------------------");
  postMessageToSlack(url,month +"月のごみ当番のお知らせです。");

  let i=0;
  pairs.forEach(pair => {
    let text = messageHeaders[i]+": " + pair.join(", ");
    postMessageToSlack(url, text);
    i++
  });

  postMessageToSlack(url,"----------------English---------------------");

   postMessageToSlack(url,"Here is the notice for the  "+month_English[month-1]+" garbage duty");

  i=0;
  pairs.forEach(pair => {
    let text = messageHeaders_English[i]+": " + pair.join(", ");
    postMessageToSlack(url, text);
    i++
  });



}


//11
//当番表にコピーする
function backupPairsToSpreadsheet(pairs) {
  let messageHeaders = [];
  
  //次の月の日付を取得
  let today = new Date();
  let nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1);

  //今月の当番シートに決めた当番表を反映させる
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("今月の当番");
  if (!sheet) {
    Logger.log("指定されたシートが見つかりません。");
    return;
  }

  // ヘッダーを設定
  sheet.getRange("A1").setValue("日付");
  sheet.getRange("B1").setValue("ペア");

  // ペアと日付の情報を配列に追加
  let pairsWithDates = [];

  //来月の第一月曜日を登録
  let firstMondayDate = NextMonthFirstMonday();

  while (firstMondayDate.getMonth() === nextMonth.getMonth()) {
    let month = firstMondayDate.getMonth() + 1;
    let date = firstMondayDate.getDate();
    let monthString = month < 10 ? '0' + month : '' + month;
    let dateString = date < 10 ? '0' + date : '' + date;
    messageHeaders.push(monthString + "月" + dateString + "日"); // 〇月〇日のペアを配列に追加
    firstMondayDate.setDate(firstMondayDate.getDate() + 7); // 次の月曜日に移動
  }

  pairs.forEach((pair, index) => {
    pairsWithDates.push({
      date: messageHeaders[index],
      pair: pair.join(", ")
    });

  });

  // スプレッドシートに出力
  pairsWithDates.forEach((pair, index) => {
    sheet.getRange(index + 2, 1).setValue(pair.date);
    sheet.getRange(index + 2, 2).setValue(pair.pair);
  });


  //その月の当番シートにも決めた当番表を反映させる
  let ms = SpreadsheetApp.getActiveSpreadsheet();
  let monthly_sheet = ms.getSheetByName(firstMondayDate.getMonth()+"月");
  if (!monthly_sheet) {
    Logger.log("指定されたシートが見つかりません。");
    return;
  }
    // スプレッドシートに出力
  pairsWithDates.forEach((pair, index) => {
    monthly_sheet.getRange(index + 2, 1).setValue(pair.date);
    monthly_sheet.getRange(index + 2, 2).setValue(pair.pair);
  });
}

//12
//来月の月曜日配列作成用関数(ただし祝日はずらす)
function MondayList(){
  let today = new Date();
  let nextMonth = new Date(today.getFullYear(), today.getMonth() + 1 , 1);
  let firstMondayDate = new Date(nextMonth.getFullYear(), nextMonth.getMonth(), 1);
  let messageHeaders = [];
  let messageHeaders_English = [];
  let shift = 0;


  while (firstMondayDate.getDay() !== 1) {
    firstMondayDate.setDate(firstMondayDate.getDate() + 1);
  }

  // 月曜日の日付を配列に追加
  while (firstMondayDate.getMonth() === nextMonth.getMonth()) {
   
    shift = 0;
    while(isHoliday(firstMondayDate) == true){
      firstMondayDate.setDate(firstMondayDate.getDate() + 1);
      console.log(firstMondayDate.getDate());
      shift+=1;
    }
   


    let month = firstMondayDate.getMonth() + 1;
    let date = firstMondayDate.getDate();
    let monthString = month < 10 ? '0' + month : '' + month;
    let dateString = date < 10 ? '0' + date : '' + date;
    messageHeaders.push(monthString + "月" + dateString + "日"); // 〇月〇日のペアを配列に追加
    messageHeaders_English.push(monthString+"/"+dateString );
    firstMondayDate.setDate(firstMondayDate.getDate() + 7 - shift);
  }
  console.log("M:"+messageHeaders);
  

  return [messageHeaders,messageHeaders_English]
}


///////////////////////////////////////////////////////////////////////////////////
//slackやスプレッドシート、カレンダーを利用するための関数群

//13
//slack送信用関数
function postMessageToSlack(url, message) {
  let jsonData = {
    'text' : message,
  }
  let payload = JSON.stringify(jsonData);
  let params = {
    'method' : 'post',
    'contentType' : 'application/json',
    'payload' : payload,
  }
  UrlFetchApp.fetch(url, params);
}

//14
//スプレッドシートをヘッダーを除いて読み込むための関数
function getRosterFromSpreadsheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("名簿");
  if (!sheet) {
    Logger.log("指定されたシートが見つかりません。");
    return [];
  }
  let lastRow = sheet.getLastRow();
  let roster = sheet.getRange("A2:A" + lastRow).getValues().flat().filter(Boolean);
  return roster;
}




///////////////////////////////////////////////////////////////////////////////////




//週の終わりで通知する今週のペア
function sendWeeklyPairsToSlackIfFirstDayOfMonth() {
  //let url = 'https://hooks.slack.com/services/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';
  let url = 'https://hooks.slack.com/services/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';


  let nextMonday = NextMonday()


  let month = nextMonday.getMonth() + 1;
  let date = nextMonday.getDate();
  let monthString = month < 10 ? '0' + month : '' + month;
  let dateString = date < 10 ? '0' + date : '' + date;
  
  

  // スプレッドシートから該当する日付のペアを取得
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(month+"月"); // 月ごとのシートを取得
  if (!sheet) {
    Logger.log("指定されたシートが見つかりません。");
    return;
  }

  // 日付とペアの情報を取得
  let pairs = [];
  let lastRow = sheet.getLastRow();
  console.log("lastRow"+lastRow);
  for (let i = 2; i <= lastRow; i++) {
    let cellValue = sheet.getRange(i, 1).getValue();
    let cellDate = new Date(cellValue).getDate(); // 日付のみ抽出
    console.log("cellDate"+cellDate);
      
    if (cellDate == date) { // 日付の比較
      let pair = [
        sheet.getRange(i, 2).getValue(),
        sheet.getRange(i, 3).getValue()
      ];
      console.log("pair"+pair);
      pairs.push(pair);
    }
  }
  console.log(pairs);
  // 取得したペアをSlackに通知
  pairs.forEach(pair => {
    let text =pair.join(", ");
    postMessageToSlack(url,"----------------Japanese---------------------");
    let message = "お疲れ様です。\n来週"+monthString + "月" + dateString + "日"+"の当番は" +text+"です。\nよろしくお願いします。" ;
    postMessageToSlack(url, message);

    
    postMessageToSlack(url,"----------------English---------------------");
    let message_English = "Thank you for your hard work. \nThe duty for " + monthString + "/" + dateString + " next week is " + text + ".\nWe appreciate your cooperation.";
    postMessageToSlack(url, message_English);
  });
}


