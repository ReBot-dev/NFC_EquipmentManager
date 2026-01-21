function doPost(e) {
  const userId = e.parameter.user_id; //eパラメータに含まれるSlackUserIDを取得
  const scriptProperties = PropertiesService.getScriptProperties(); 
  scriptProperties.setProperty('TARGET_USER_ID', userId);
  const triggerId = ScriptApp.newTrigger('delayedResponse')
    .timeBased() //時間駆動型
    .after(2 * 1000)
    .create() //処理実行の予約
    .getUniqueId(); // トリガーのIDを直接取得
  scriptProperties.setProperty('CURRENT_TRIGGER_ID', triggerId);

  // レスポンスメッセージの処理
  const responseData = {
    text: "処理を受け付けました。操作説明をDMで送信します。", 
  };
  const output = ContentService.createTextOutput(JSON.stringify(responseData));
  return output.setMimeType(ContentService.MimeType.JSON); 
}

function delayedResponse() {
  //IDの読み込みと削除
  const scriptProperties = PropertiesService.getScriptProperties();
  const userId = scriptProperties.getProperty('TARGET_USER_ID');
  const triggerId = scriptProperties.getProperty('CURRENT_TRIGGER_ID'); 
  scriptProperties.deleteProperty('TARGET_USER_ID');
  scriptProperties.deleteProperty('CURRENT_TRIGGER_ID');

  const message = `*基本的な操作方法*\nTabキーで選択、スペースキーで決定です。\n初めてスキャンする社員証や物品に貼ったNFCタグの場合、登録を求められます。\n貸出登録などが完了すると、Slackの個人DMに通知が届きます。これが届かない場合は、社員証の初期登録の入力が間違っている場合があります。その場合は以下のURLから修正してください。\n過去の貸出履歴や、今誰が何を借りているのかなども確認できます\nバグや改善案などありましたら記入のご協力をお願いします。\nhttps://docs.google.com/spreadsheets/d/1xEhbHo_Vy_hd8T6WuoRqQm3AaK5BIw9US3RpLze3e1E/edit?usp=sharing \n*もし電源が切れていたら*\n1. ログインパスワード : khadas\n2. Windowsキーを押して"cmd"と入力、Enter\n3. 黒い画面が出てきたら、"source /myenv/bin/activate" と入力してEnter\n4. "python3 EM.py" と入力、Enter\n5. しばらくすると初期画面が起動します！`;
  sendSlackDM(userId,message);
}

function sendSlackDM(userId, message) {

  const SLACK_TOKEN = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');

  if (!SLACK_TOKEN) {
    console.error("エラー: スクリプトプロパティに 'SLACK_BOT_TOKEN' が設定されていません。");
    return; //トークンがないため中断
  }
  const url = 'https://slack.com/api/chat.postMessage';
  const payload = {
    'channel': userId,
    'text': message
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json; charset=utf-8',
    'headers': {
      'Authorization': 'Bearer ' + SLACK_TOKEN 
    },
    'payload': JSON.stringify(payload),
   'muteHttpExceptions': true
  };
  try {
   const response = UrlFetchApp.fetch(url, options);
   const jsonResponse = JSON.parse(response.getContentText());
   if (!jsonResponse.ok) {
     console.error("Slack DM送信エラー: " + jsonResponse.error);
   } else {
     console.log("Slack DMを送信しました。");
   }
  } catch (error) {
    console.error("通信エラー: " + error.toString());
  }
}