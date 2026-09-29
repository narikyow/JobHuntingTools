function generateAndSendToDo() {
  const props = PropertiesService.getScriptProperties();
  const GEMINI_API_KEY = props.getProperty('GEMINI_API_KEY');
  const EXTERNAL_DB_ID = props.getProperty('EXTERNAL_DB_ID'); // 外部データベースのID
  const HISTORY_DB_ID = props.getProperty('HISTORY_DB'); // ToDo_HistoryシートのID

  const emailAddress = Session.getActiveUser().getEmail(); 

  // 1. 履歴データの取得（前回のToDo結果を引き継ぐ）
  const historySheet = SpreadsheetApp.openById(HISTORY_DB_ID).getActiveSheet();
  const historyData = historySheet.getDataRange().getValues();
  let lastToDoResult = "前回の履歴はありません。";
  if (historyData.length > 0) {
    // 最終行のB列(インデックス1)を前回の結果として取得
    lastToDoResult = historyData[historyData.length - 1][1];
    if (historyData.length-1>0){
      lastlastToDoResult = historyData[historyData.length - 2][1]; 
    }
  }

  // 2. 就活メールDBの未処理データ取得
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeDataRaw = activeSheet.getDataRange().getValues();
  
  const newActiveData = [];
  const rowsToUpdate = []; 

  for (let i = 0; i < activeDataRaw.length; i++) {
    if (i === 0) {
      newActiveData.push(activeDataRaw[i]); // ヘッダーは保持
      continue;
    }
    const row = activeDataRaw[i];
    // 処理状態（E列、インデックス4）が「済」でないものを新規として扱う
    if (row[4] !== "済") {
      newActiveData.push(row);
      rowsToUpdate.push(i + 1); // getRangeの行指定用（1始まり）
    }
  }

  let externalData = [];
  try {
    const externalSheet = SpreadsheetApp.openById(EXTERNAL_DB_ID).getActiveSheet();
    externalData = externalSheet.getDataRange().getValues();
  } catch (e) {
    console.error("外部データベースの読み込みに失敗しました:", e);
  }
  
  const targetData = JSON.stringify({
    newActiveData: newActiveData,
    externalDb: externalData
  });

  // 3. プロンプトの構築
  const prompt = `
以下の【前回、前々回のToDo結果（引き継ぎ事項）】と【新規追加されたデータ（未処理）】を統合・分析し、本日の私に対する最新のToDoリストとリマインドを作成せよ。

【前提条件とルール】
・タスクの捏造（ハルシネーション）を固く禁ずる。必ず提供されたデータに基づき作成すること。
・前回からの引き継ぎタスクと、新規データから発生したタスクを適切に統合すること。
・ToDoリストは、現在の日時から「1週間以内」のタスクを対象とし、優先順位付けを行うこと。
・リマインドは緊急性が低く、現在の日時から「3週間以内」のタスクを対象とすること。
・選考に直結するもののみをToDoとリマインドに入れること。選考に直結しないセミナーや説明会などは「Event」に分類すること。
・完了メールが確認できている、あるいは既に終了したタスクについては「Complete」に移動すること。
・人事とのやりとりやメッセージ関連の内容（メールでのやりとり、Labbase、マイナビのようなスカウト等。ただしLabbase以外のスカウトは無視）は「Message」に含めること。
・「Message」については、どのような理由で、どの方向でメッセージを返すべきか、簡潔な方針を提示すること。
・Labbaseの「興味あり」は、MessageタスクではなくEvent欄に加え、届いた日のみメールに載せるものとする。

【出力形式】
## ToDo
* **企業名** : タスク名称 期限
* * タスク概要

## Message
* **企業名** : タスク名称 期限
* * タスク概要

## リマインド
* **企業名** : タスク名称 期限
* * タスク概要

## Event
* **企業名** : タスク名称 期限
* * タスク概要

## Complete
* **企業名** : タスク名称 期限
* * タスク概要

【前回のToDo結果（引き継ぎ事項）】
${lastToDoResult}

【前々回のToDo結果（引き継ぎ事項）】
${lastlastToDoResult}


【新規追加されたデータ（未処理）】
${targetData}
  `;

  // 4. Gemini APIの呼び出し
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  const payload = {
    "contents": [{ "parts": [{ "text": prompt }] }]
  };
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  let generatedText = "";
  const maxRetries = 3;
  let attempt = 0;
  let success = false;

  while (attempt < maxRetries && !success) {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      const json = JSON.parse(response.getContentText());
      generatedText = json.candidates[0].content.parts[0].text;
      success = true;
    } else if (responseCode === 503 || responseCode === 429) {
      attempt++;
      if (attempt >= maxRetries) {
        throw new Error(`Gemini APIエラー: 一時的な高負荷が解消されません。レスポンスコード:${responseCode} 詳細:${response.getContentText()}`);
      }
      Utilities.sleep(3000 * attempt);
    } else {
      throw new Error(`Gemini APIエラー: 予期せぬエラーが発生しました。レスポンスコード:${responseCode} 詳細:${response.getContentText()}`);
    }
  }

  // 5. HTML形式にマークアップしてメール送信
  let htmlBody = generatedText
    .replace(/## (.*)/g, '<h2>$1</h2>')
    .replace(/\* \*\*(.*?)\*\* : (.*)/g, '<ul><li><strong>$1</strong> : $2</li>')
    .replace(/\* \* (.*)/g, '<ul><li>$1</li></ul></ul>')
    .replace(/\n/g, '<br>');

  MailApp.sendEmail({
    to: emailAddress,
    subject: "本日のToDoリストとリマインド",
    htmlBody: htmlBody
  });

  // 6. ToDo_Historyシートへの記録
  updateHistorySheet(HISTORY_DB_ID, generatedText);
  
  // 7. 処理対象データの「済」フラグ書き込み
  if (rowsToUpdate.length > 0) {
    markAsProcessed(activeSheet, rowsToUpdate);
  }
}

function updateHistorySheet(historySheetId, content) {
  const sheet = SpreadsheetApp.openById(historySheetId).getActiveSheet();
  const now = new Date();

  sheet.appendRow([now, content]);

  const data = sheet.getDataRange().getValues();
  const oneWeekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

  for (let i = data.length - 1; i >= 0; i--) {
    const rowDate = new Date(data[i][0]);
    if (!isNaN(rowDate) && rowDate < oneWeekAgo) {
      sheet.deleteRow(i + 1);
    }
  }
}

function markAsProcessed(sheet, rowsToUpdate) {
  // 就活メールDBの5列目（E列）を「済」に更新する
  rowsToUpdate.forEach(rowIndex => {
    sheet.getRange(rowIndex, 5).setValue("済");
  });
}