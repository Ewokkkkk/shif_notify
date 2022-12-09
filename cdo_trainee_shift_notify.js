// 翌日に実行するトリガを設定する処理
function setTrigger() {
  var time = new Date();
  var this_year = time.getFullYear();
  var day = time.getDay(); // 今日の曜日

  time.setDate(time.getDate() + 1);
  time.setHours(9); //時
  time.setMinutes(0); //分
  time.setSeconds(0); //秒

  // 次の日が年を越してればトリガを設定しない
  if (this_year == time.getFullYear()) {
    ScriptApp.newTrigger('sendMessage').timeBased().at(time).create();
  }
}

function sendMessage() {
  setTrigger();
  // slack appのWebhook URL
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
  // 今日の日付取得
  const today = new Date();
  // const today = new Date('December 9, 2023 03:24:00');
  // 比較用に時間は削除
  today.setHours(0);
  today.setMinutes(0);
  today.setSeconds(0);
  today.setMilliseconds(0);

  // ex:人材開発室勤務表_202212
  const fileName = "人材開発室勤務表_" + today.getFullYear() + (today.getMonth() + 1);

  var target = DriveApp.getFolderById(""); // 人材開発室勤務表のフォルダIDを指定
  var folders = target.getFolders();

  // 今年のフォルダのフォルダIDを取得
  while (folders.hasNext()) {
    var folder = folders.next();
    var folderName = folder.getName();

    if (folderName == today.getFullYear()) {
      var this_folder = folder.getId();
    }
  }

  // 勤務表があるフォルダーID・ファイル名からファイルを取得
  var files = DriveApp.getFolderById(this_folder).getFilesByName(fileName);
  var file = files.next();

  // 開いているファイルのURL
  var fileURL = file.getUrl();

  const spreadsheet = SpreadsheetApp.openById(file.getId());
  const sheet = spreadsheet.getSheetByName("シフト");
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const values = sheet.getDataRange().getValues();
  const firstColumnVals = sheet.getRange(1, 1, 1, lastCol).getValues();

  const month = today.getMonth() + 1;
  const date = today.getDate();
  const day_arr = ['日', '月', '火', '水', '木', '金', '土'];
  const day = day_arr[today.getDay()];
  const today_str = month + "月" + date + "日" + "(" + day + ")";

  var trainee_column_start;  // 研修生の列の開始列
  var todays_data = [];  // 今日の日付の行用の配列
  var name_row = values[1];  // 名前の列用配列
  var line = [today_str + " の研修生の出勤予定", ""];  // 出力用配列

  // 研修生の列の開始位置を取得
  for (let i = 0; i < firstColumnVals[0].length; i++) {
    // console.log(firstColumnVals[i]);
    if (firstColumnVals[0][i] == "研修生") {
      trainee_column_start = i;
    }
  }

  // 今日の日付の行を取得してtodays_dataに入れる
  for (let i = 2; i < lastRow - 1; i++) {
    if (values[i][0].getTime() === today.getTime()) {
      todays_data = values[i];
    }
  }

  // 研修生の列から名前が続く列までループ
  // todays_dataに出勤あれば、lineにpush
  for (let i = trainee_column_start; i < todays_data.length; i++) {
    // 名前のある列で終わり
    if (name_row[i] == "") {
      break;
    }
    // 空白か欠勤以外であれば、配列lineに値を追加
    if (todays_data[i] != "" && todays_data[i] != "欠勤") {
      line.push(name_row[i] + "さん : " + todays_data[i]);
    }
  }

  // 出勤がいなければ
  if (line.length == 2) {
    line.push("研修生の出勤予定はありません");
  }

  line.push(fileURL);

  var output = line.join('\n');
  return output;
}