const CHANNEL_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty('CHANNEL_ACCESS_TOKEN'); // LINEチャネルアクセストークン
const SHEET_ID = PropertiesService.getScriptProperties().getProperty('SHEET_ID'); // スプレッドシートのID

function doPost(e) {
  const json = JSON.parse(e.postData.contents);

  // LINEから送信されたメッセージの処理
  const events = json.events;
  for (let i = 0; i < events.length; i++) {
    const event = events[i];
    const userId = event.source.userId;
    const groupId = event.source.groupId;
    const messageText = event.message.text.trim();

    // スプレッドシートに記録
    handleUserAction(userId, groupId, messageText);
  }

  // LINEサーバーへの応答
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(ContentService.MimeType.JSON);
}

function handleUserAction(userId, groupId, messageText) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  const now = new Date();
  const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  const formattedTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm');

  // ユーザーのLINEプロフィールを取得
  const userName = getUserName(userId);

  // 「参加」アクションの場合
  if (messageText === '参加') {
    sheet.appendRow([userName, formattedDate, formattedTime, '', '']);
    sendReply(groupId, `${userName} さんが参加しました！`);
  }

  // 「退室」アクションの場合
  else if (messageText === '退室') {
    const rows = sheet.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 0; i--) {
      if (rows[i][0] === userName && rows[i][3] === '') {
        sheet.getRange(i + 1, 4).setValue(formattedTime);

        const startDateTime = new Date(rows[i][2]);
        startDateTime.setFullYear(rows[i][1].getFullYear()); 
        startDateTime.setMonth(rows[i][1].getMonth()); 
        startDateTime.setDate(rows[i][1].getDate()); 

        // 勉強時間の計算（時間単位）
        const elapsedTime = now.getTime() - startDateTime.getTime();
        const studyTime = formatElapsedTime(elapsedTime); // 結果を時間単位に変換

        // 勉強時間をスプレッドシートに記録
        sheet.getRange(i + 1, 5).setValue(studyTime);

        // ボットのレスポンス
        sendReply(groupId, `${userName} さん、お疲れ様でした！\n勉強時間: ${studyTime}`);
        return;
      }
    }
    sendReply(groupId, 'エラー: 参加記録が見つかりませんでした。');
  }
}

function getUserName(userId) {
  const url = 'https://api.line.me/v2/bot/profile/' + userId;
  const options = {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + CHANNEL_ACCESS_TOKEN,
    },
  };
  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());
  return json.displayName || '名前不明';
}

function sendReply(userId, message) {
  const url = 'https://api.line.me/v2/bot/message/push';
  const payload = {
    to: userId,
    messages: [
      {
        type: 'text',
        text: message,
      },
    ],
  };
  const options = {
    method: 'post',
    headers: {
      Authorization: 'Bearer ' + CHANNEL_ACCESS_TOKEN,
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify(payload),
  };
  UrlFetchApp.fetch(url, options);
}

// 毎週日曜日の24:00に1週間分の集計結果を送信
function sendWeeklySummary() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  const rows = sheet.getDataRange().getValues();
  const now = new Date();
  const startOfWeek = new Date(now);
  startOfWeek.setDate(now.getDate() - 7); // 一週間前
  startOfWeek.setHours(0, 0, 0, 0); // その日の開始時刻

  // 現在日時を終了日時に設定 
  const endOfWeek = new Date(now);
  endOfWeek.setHours(23, 59, 59, 999); // 現在の終了時刻

  const weeklyData = {};
  rows.forEach((row, index) => {
    if (index === 0 || row[1] === '' || row[2] === '' || row[3] === '') return; // ヘッダーまたは空行スキップ
    const date = new Date(row[1]);
    if (date >= startOfWeek && date <= endOfWeek) {
      weeklyData[row[0]] = (weeklyData[row[0]] || 0) + (row[3].getTime() - row[2].getTime());
    }
  });

  const summary = Object.entries(weeklyData)
    .sort((a, b) => b[1] - a[1])
    .map((entry, index) => `${index + 1}. ${entry[0]} 合計勉強時間 ${formatElapsedTime(entry[1])}`)
    .join('\n');

  const message = `${Utilities.formatDate(startOfWeek, Session.getScriptTimeZone(), 'MM/dd')}～${Utilities.formatDate(endOfWeek, Session.getScriptTimeZone(), 'MM/dd')} 集計結果\n${summary}`;

  // Botに送信
  const groupId = PropertiesService.getScriptProperties().getProperty('GROUP_ID'); // 対象のグループIDやユーザーIDを指定
  sendReply(groupId, message);
}

// 勉強時間を適切なフォーマットに変換する関数
function formatElapsedTime(milliseconds) {
  const totalSeconds = Math.floor(milliseconds / 1000);
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;

  let result = '';
  if (hours > 0) {
    result += `${hours}時間`;
  }
  if (minutes > 0 || hours > 0) { // 時間がある場合、分が0でも表示
    result += `${minutes}分`;
  }
  result += `${seconds}秒`;

  return result;
}