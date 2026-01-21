function onChangeHandler(e) {
  // スクリプトの多重実行を防ぐためにロック
  const lock = LockService.getScriptLock();
  
  // 10秒間ロック
  if (lock.tryLock(10000)) { 
    try {
      console.log("ロックを取得。処理を開始。");
      
      // 社員マスタの差分検出
      checkEmployeeMasterDiff();

      // 貸出中一覧の差分検出
      checkLendListDiff();

    } catch (e) {
      console.error("エラーが発生: " + e.toString());
    } finally {
      // 処理が完了したらロック解放
      lock.releaseLock();
      console.log("ロックを解放");
    }
  } else {
    // 他のプロセスが実行中の場合は何もしない
    console.log("別の処理が実行中のため、今回のトリガーはスキップ");
  }
}

// 社員マスタの差分検出とSlackID自動書き込み
function checkEmployeeMasterDiff() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("社員マスタ");
  const data = sheet.getDataRange().getValues();
  const props = PropertiesService.getScriptProperties();
  const prevDataJson = props.getProperty('employee_master');
  let prevData = [];
  if (prevDataJson) prevData = JSON.parse(prevDataJson);

  for (let i = 1; i < data.length; i++) {
    const prevRow = prevData[i] || [];
    // メールアドレスが新規入力 or 変更された場合
    if (data[i][2] && data[i][2] !== (prevRow[2] || "")) {
      const slackId = getSlackIdByEmail(data[i][2]);
      sheet.getRange(i + 1, 4).setValue(slackId || "Not Found");
      const message = `【登録完了(Registration completed)】\n社員証の登録が完了しました。以後備品の貸出情報や返却のお願いなどがこのDMに送信されます。\n詳しい使い方を表示するには /help 、現在貸し出されている物のリストを表示するには /list と送信してください。\nYour employee ID card has now been registered. From now on, information about equipment rentals and requests for returns will be sent via this DM.\nSend /help for detailed instructions, or /list for a list of currently checked out items.`;
      sendSlackDM(slackId, message);
    }
  }
  // 現在の内容を保存
  props.setProperty('employee_master', JSON.stringify(data));
}

// 貸出中一覧の差分検出とSlack DM送信
function checkLendListDiff() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("貸出中一覧");
  const data = sheet.getDataRange().getValues();
  const props = PropertiesService.getScriptProperties();
  const prevDataJson = props.getProperty('lend_list');
  let prevData = [];
  if (prevDataJson) prevData = JSON.parse(prevDataJson);

  for (let i = 1; i < data.length; i++) {
    const prevRow = prevData[i] || [];
    // 申請者名が新規入力 or 変更
    if (data[i][1] && data[i][1] !== (prevRow[1] || "")) {
      const applicantName = data[i][1];
      const email = getEmailByName(applicantName);
      if (!email) continue;
      const slackId = getSlackIdByEmail(email);
      if (!slackId) continue;
      const itemName = data[i][2];
      const returnDate = data[i][3];
      const message = `【貸出完了(Registration completed)】\n・物品名: ${itemName}\n・返却予定日: ${returnDate}`;
      sendSlackDM(slackId, message);
    }
  }
  // 現在の内容を保存
  props.setProperty('lend_list', JSON.stringify(data));
}

/**
社員マスタから名前を検索し、対応するメールアドレスを取得する関数
@param {string} name - 検索する氏名
@return {string|null} - メールアドレス or null
*/
function getEmailByName(name) {
  try {
    const masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
    const data = masterSheet.getDataRange().getValues(); // 社員マスタの全データを取得

    // 名前を検索
        for (let i = 1; i < data.length; i++) {
      if (data[i][0] === name) { // A列の名前が一致したら
        return data[i][2];       // C列のメールアドレスを返す
      }
    }
    return null; // 見つからなかった場合
  } catch (error) {
    console.error("社員マスタの検索エラー: " + error.toString());
    return null;
  }
}


/**
指定したユーザーにSlackのDMを送信する関数
@param {string} userId - 送信先のSlackユーザーID
@param {string} message - 送信するメッセージ本文
*/
function sendSlackDM(userId, message) {

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

/**
メールアドレスからSlackのユーザーIDを取得する関数（これは既存のものを流用）
@param {string} email - 検索するメールアドレス
@return {string|null} - SlackのユーザーID or null
*/
function getSlackIdByEmail(email) {
  if (!email || !SLACK_TOKEN) return null;
  const url = `https://slack.com/api/users.lookupByEmail?email=${encodeURIComponent(email)}`;
  const options = {
    'method': 'get',
    'headers': { 'Authorization': 'Bearer ' + SLACK_TOKEN },
    'muteHttpExceptions': true
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    const jsonResponse = JSON.parse(response.getContentText());
    if (jsonResponse.ok) return jsonResponse.user.id;
    else {
      console.error("Slack API Error: " + jsonResponse.error);
      return null;
    }
  } catch (error) {
    console.error("Fetch Error: " + error.toString());
    return null;
  }
}

function remindUnreturnedItems() {
  const lock = LockService.getScriptLock();

  if (lock.tryLock(10000)) {try {
    console.log("ロックを取得。処理を開始。");
    remindUnreturnedItems();
  } catch (e) {
    console.error("エラーが発生: " + e.toString());
  } finally {
    lock.releaseLock();
    console.log("ロックを開放");
    }
  } else {
    console.log("別の処理が実行中のため，今回のトリガーはスキップされました。")
  }
}

function remindUnreturnedItems_unlocked() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("貸出中一覧");
  const data = sheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  for (let i = 1; i < data.length; i++) {
    const applicantName = data[i][1];
    const itemName = data[i][2];
    const returnDate = data[i][3];

    // 返却予定日を過ぎて返却されていない
    if (returnDate && returnDate <= today) {
      const email = getEmailByName(applicantName);
      const slackId = getSlackIdByEmail(email);
      if (slackId) {
        const message = `【返却催促(Return reminder)】\n「${itemName}」の返却予定日（${returnDate}）を過ぎています。継続して借りる場合は再度貸出登録をしてください。\nThe expected return date (${returnDate}) for "${itemName}" has passed. If you wish to continue borrowing the item, please register for loan again.`;
        sendSlackDM(slackId, message);
      }
    }
  }
}

function remindUnreturnedItems_2() {
  const lock = LockService.getScriptLock();

  if (lock.tryLock(10000)) {try {
    console.log("ロックを取得。処理を開始。");
    remindUnreturnedItems_2_unlocked();
  } catch (e) {
    console.error("エラーが発生: " + e.toString());
  } finally {
    lock.releaseLock();
    console.log("ロックを開放");
    }
  } else {
    console.log("別の処理が実行中のため，今回のトリガーはスキップされました。")
  }
}

function remindUnreturnedItems_2_unlocked() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("貸出中一覧");
  const data = sheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  for (let i = 1; i < data.length; i++) {
    const applicantName = data[i][1];
    const itemName = data[i][2];
    const returnDate = data[i][3];

    // 返却予定日を過ぎて返却されていない
    if (returnDate && returnDate <= today) {
      const email = getEmailByName(applicantName);
      const slackId = getSlackIdByEmail(email);
      if (slackId) {
        const message = `【返却催促(Return reminder)】\n「${itemName}」の返却予定日（${returnDate}）です。返却してからの帰宅をお願いします。継続して借りる場合は再度貸出登録をしてください。\nThe due date for "${itemName}" is (${returnDate}). Please return it before going home. If you wish to continue borrowing, please register for loan again.`;
        sendSlackDM(slackId, message);
      }
    }
  }
}