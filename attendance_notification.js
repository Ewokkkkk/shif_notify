function sendMessage() {
  const postUrl = '';
  const sendMessage = createMessage();
  const jsonData = {
    "text": sendMessage
  };
  const payload = JSON.stringify(jsonData);
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": payload
  };
  UrlFetchApp.fetch(postUrl, options);
}

function createMessage() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // const sheet = spreadsheet.getActiveSheet();
  const sheet = spreadsheet.getSheetByName("シート1");
  const lastRow = sheet.getLastRow();
  const values = sheet.getDataRange().getValues();

  // 今日の日付取得
  const today = new Date();
  // 比較用に時間は削除
  today.setHours(0);
  today.setMinutes(0);
  today.setSeconds(0);
  today.setMilliseconds(0);

  const month = today.getMonth() + 1;
  const date = today.getDate();
  const day_arr = ['日', '月', '火', '水', '木', '金', '土'];
  const day = day_arr[today.getDay()];
  const today_str = month + "月" + date + "日" + "(" + day + ")";

  // 今日の日付の行用の配列
  var todays_data = [];

  // 出力用配列
  var line = [today_str + " の研修生の出勤予定", ""];

  // 今日の日付の行を取得してtodays_dataに入れる
  for (let i = 1; i < lastRow - 1; i++) {
    if (values[i][0].getTime() === today.getTime()) {
      todays_data = values[i];
    }
  }

  // todays_dataに出勤あれば
  for (let i = 1; i < todays_data.length; i++) {
    if (todays_data[i] != "" && todays_data[i] != "欠勤") {
      line.push(values[0][i] + "さん : " + todays_data[i]);
    }
  }

  // 出勤がいなければ
  if (line.length == 2) {
    line.push("研修生の出勤予定はありません");
  }

  line.push("勤務表のURL");

  var output = line.join('\n');
  return output;
}
