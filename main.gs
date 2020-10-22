// 管理に使いたいフォルダの名前を入れる。この名前でドライブ直下にフォルダが作られる。（被らない名前にする）
var FOLDER_NAME = "勤怠管理シート";
var CONFIG_FILE_NAME = "設定";

// Outgoing WebhookのTokenを保存する（一応認証する）
var SLACK_OUTGOING_TOKEN = "";

// Incoming WebhookのURLを入力する
var SLACK_WEBHOOK_URL = "";

// アイコンにしたい画像のURLを入力する
var ICON_URL = "";

// シート打刻名とアルファベットの対応表とか(要改善)
var CELL_NAME_AL = {休暇: "C", 出勤: "D", 休憩開始: "E", 休憩終了: "F", 退勤: "G", 今日やること:"I", やったこと: "J", やり残したこと: "K", 明日やること: "L"};

// テキストの文字区切り設定 (休暇などは「"有給" + SPLIT_TEXT + "2020/02/22"」のように下記文字の区切りで入力判断させる)
var SPLIT_TEXT = ",";

// 最初に手動で実行する（直接フォルダを作ってもいい）
function setUp() {
  var folder = getFolder_(FOLDER_NAME);
  if (!folder) {
    DriveApp.createFolder(FOLDER_NAME);
  }
  var spreadsheet = openSpreadsheetByName(CONFIG_FILE_NAME);
  // メッセージ用のシートを作成
  getMessageSheet(spreadsheet);
}

// メッセージ用のシートを開く
function getMessageSheet(spreadsheet) {
  // メッセージテンプレート設定
  var sheet = spreadsheet.getSheetByName('_メッセージ');
  if(!sheet) {
    sheet = spreadsheet.insertSheet('_メッセージ');
    if(!sheet) {
      throw "エラー: メッセージシートを作れませんでした";
    }
    else {
      sheet.getRange("A1:O2").setValues([
        [
          "出勤", "出勤更新", "退勤", "退勤更新", "休暇", "休暇取消",
          "出勤中", "出勤なし", "休暇中", "休暇なし", "出勤確認", "退勤確認", "休憩開始", "休憩終了", "エラー"
        ],
        [
          "<@#1> おはようございます (#2)", "<@#1> 出勤時間を#2へ変更しました",
          "<@#1> お疲れ様でした (#2)", "<@#1> 退勤時間を#2へ変更しました",
          "<@#1> #2を休暇として登録しました", "<@#1> #2の休暇を取り消しました",
          "#1が出勤しています", "全員退勤しています",
          "#1は#2が休暇です", "#1に休暇の人はいません",
          "今日は休暇ですか？ #1", "退勤しましたか？ #1", "<@#1> 休憩時間を打刻しました（#2）", "<@#1> 休憩終了を打刻しました（#2）", "<@#1> さんの#3の打刻でエラーが発生しました。スプレッドシートを確認してみてください。"
        ]
      ]);
    }
  }

  return sheet;
};

// 設定シートからテンプレを取得
function getMessage(type) {
  var spreadsheet = openSpreadsheetByName(CONFIG_FILE_NAME);
  var sheet = getMessageSheet(spreadsheet);
  var types = sheet.getRange("A1:Q1").getValues()[0];
  var alp = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q'];
  var template = null;
  for(var i = 0; i < types.length; ++i) {
    if(types[i] == type) {
      template = sheet.getRange(alp[i]+'2').getValues();
    }
  }
  return String(template[0]);
}

// メイン処理
function doPost(e) {
  // return messageSend(JSON.stringify(e));

  var verificationToken = e.parameters.token;
  // 一応トークンチェックをする
  if (verificationToken != SLACK_OUTGOING_TOKEN) {
    var response = '不正なアクセスです。';
    return messageSend(response);
  }

  var text = distributeProcessing(String(e.parameters.user_name),String(e.parameters.text));
  return messageSend(text);
}

// 打刻の文字に合わせて処理を分ける感じにする
function distributeProcessing(user, text) {
  var currentTime = get_dateTime();

  var type = "";
  if (text.match(/^出勤/)) {
    type = "出勤";
  } else if (text.match(/^退勤/)) {
    type = "退勤";
  } else if (text.match(/^休憩/)) {
    type = "休憩開始";
  } else if (text.match(/^戻り/)) {
    type = "休憩終了";
  } else if (text.match(/^有給/) || text.match(/^半休/) || text.match(/^代休/)) {
    message = text.split(SPLIT_TEXT);
    if (message.length == 2 && message[1].match(/^\d{4}\/\d{1,2}\/\d{1,2}/)) {
      var date = message[1];
      var type = message[0];
      return vacation(user, date, type);
    } else {
      return "入力されたフォーマットがおかしいようです。「休みの種類,年/月/日」の形式で、コロン、スラッシュ、数字を半角で入力してください。\n例：「有給,2020/02/22」";
    }

  } else {
    // webhookのトリガー上ないとは思うが…
    return "不正な形式のようです。文頭にスペースが入っていないかなど、形式を確認してみてください。";
  }

  return stamp(user, text, currentTime, type);
}

// 送信パラメータ設定
function messageSend(text) {
  var options = {
                "method" : "post",
                "contentType" : "application/json",
                "payload" : JSON.stringify({
                  "username":"勤怠bot",
                  "icon_url": ICON_URL,
                  "text": text,
                  link_names: 1
                })
              };
  //投稿先
  return UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
}

// 現在時刻の取得と日本時間への加工
function get_dateTime() {
  var date = new Date();
  var format = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd/HH:mm');
  return format;
}

function get_time() {
  var date = new Date();
  var format = Utilities.formatDate(date, 'Asia/Tokyo', 'HH:mm');
  return format;
}

function get_hour() {
  var date = new Date();
  var format = Utilities.formatDate(date, 'Asia/Tokyo', 'HH');
  return format;
}

function get_date_int() {
  var date = new Date();
  var format = Utilities.formatDate(date, 'Asia/Tokyo', 'dd');
  return Number(format);
}

function get_month_int() {
  var date = new Date();
  var format = Utilities.formatDate(date, 'Asia/Tokyo', 'MM');
  return Number(format);
}

function get_year_int() {
  var date = new Date();
  var format = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy');
  return Number(format);
}

function stamp(user, text, currentTime, type) {
  var spreadsheet = openSpreadsheetByName(user);

  // 日付周りの単品取得
  var day = get_date_int();
  var time = get_time();
  var hour = get_hour();

  // 打刻するセル番号を設定（今日の日付+1のところになる）
  // 一応0~1時の打刻&&退勤のばあいは前日のセルに打刻するように
  if ((hour == '1' || hour == '0') && type == "退勤") {
    var todaysCellNum = day;
  } else {
    var todaysCellNum = day + 1;
  }
  var stampCell = CELL_NAME_AL[type] + todaysCellNum;

  var sheet = getStampSheet(spreadsheet, get_year_int(), get_month_int());
  if (sheet.getRange(stampCell).setValues([[time]])) {
    var template = getMessage(type);
  } else {
    var template = getMessage("エラー");
  }
  return processMessage(user, template, currentTime, type);
}


// 休日の設定
function vacation(user, date, vacationType) {
  var spreadsheet = openSpreadsheetByName(user);
  var type = "休暇";

  var dateArr = date.split('/');
  // 日付周りの単品取得
  var year = dateArr[0];
  var month = dateArr[1];
  var day = dateArr[2];

  // 打刻するセル番号を設定（今日の日付+1のところになる）
  var stampCellNum = CELL_NAME_AL[type] + String(Number(day) + 1);

  var sheet = getStampSheet(spreadsheet, Number(year), Number(month));
  if (sheet.getRange(stampCellNum).setValues([[vacationType]])) {
    var template = getMessage(type);
  } else {
    var template = getMessage("エラー");
  }
  return processMessage(user, template, date, type);
}

// メッセージテンプレの送信用加工
function processMessage (user, template, currentTime, type) {
  var template = template.replace(/#1/g, user);
  template = template.replace(/#2/g, currentTime);
  template = template.replace(/#3/g, type);
  return template;
}

// 打刻するシートを取得する
function getStampSheet(spreadsheet, year, month) {
  var sheetName = String(year) + "/" + String(month);
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = setUpCurrentMonthStampSheet(spreadsheet, year, month, sheetName);
    if(!sheet) {
      throw "エラー: "+sheetName+"のシートが作れませんでした";
    }
  }
  return sheet;
}

// 打刻するシートを錬成する
function setUpCurrentMonthStampSheet(spreadsheet, year, month, sheetName) {
  var newSheet = spreadsheet.insertSheet(sheetName);
  newSheet.getRange("A1:L1").setValues([["日付", "曜日・祝日", "休暇", "出勤", "休憩開始", "休憩終了", "退勤", "   ", "今日やること", "やったこと", "やり残したこと", "明日やること"]]);
  var dateArr = getMonthWeekDayAndHoliday(year, month);
  var rangeLength = String(dateArr.length + 1);
  newSheet.getRange("A2:C" + rangeLength).setValues(dateArr);
  return newSheet;
}

// フォーマットに沿った月の日付・曜日・祝日のデータを返す
function getMonthWeekDayAndHoliday(year, month) {
  var endDate = new Date(year, month, 0);
  var date = endDate.getDate();
  var dateArr = [...Array(date).keys()];
  dateArr = dateArr.map(d => buildDateData(year, month, d + 1));
  return dateArr;
}

function buildDateData(year, month, date) {
  var vacation = "";
  var processDate = new Date(year, month - 1, date);

  // 曜日の判定
  var dayArr = ['日', '月', '火', '水', '木', '金', '土'];
  var num = processDate.getDay();
  if (num === 0 || num === 6) {
    vacation = "定休日";
  }
  var day = dayArr[num];

  // 祝日の判定
  var id = 'ja.japanese#holiday@group.v.calendar.google.com'
  var cal = CalendarApp.getCalendarById(id);
  var events = cal.getEventsForDay(processDate);
  if (events.length) {
    vacation = "定休日";
    day += "・祝";
  }

  return [date, day, vacation]
}

// スプレッドシートの取得
function openSpreadsheetByName(fileName) {
  if (fileName == CONFIG_FILE_NAME) {
    if (fileExists_(fileName)) {
      var folder = getFolder_(FOLDER_NAME);
      var files = folder.getFilesByName(fileName);
      var file = files.next();
    } else {
      var file = createSpreadsheetInfolder(fileName);
    }
  } else {
    if (fileExists_(fileName + "_勤怠管理表")) {
      var folder = getFolder_(FOLDER_NAME);
      var files = folder.getFilesByName(fileName + "_勤怠管理表");
      var file = files.next();
    } else {
      var file = createSpreadsheetInfolder(fileName + "_勤怠管理表");
    }
  }

  return SpreadsheetApp.openById(file.getId());
}

// 新しいスプレッドシートを作成する
function createSpreadsheetInfolder(fileName) {
  var folder = getFolder_(FOLDER_NAME);
  var newSS = SpreadsheetApp.create(fileName);
  var originalFile = DriveApp.getFileById(newSS.getId());
  var copiedFile = originalFile.makeCopy(fileName, folder);
  DriveApp.getRootFolder().removeFile(originalFile);
  return copiedFile;
}

// ファイルの存在チェック
function fileExists_(fileName) {
  var folder = getFolder_(FOLDER_NAME);
  var files = folder.getFilesByName(fileName);

  while (files.hasNext()) {
    var file = files.next();
    if(fileName == file.getName()) {
      return true;
    }
  }

  return false;
}

// フォルダ取得
function getFolder_(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  while(folders.hasNext()) {
    var folder = folders.next();
    if(folder.getName() == folderName){
      break;
    }
  }

  return folder;
}
