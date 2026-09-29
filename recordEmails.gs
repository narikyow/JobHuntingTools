/**
 * 就活メール収集・管理スクリプト（件名・本文の個人情報マスキング機能付き）
 */

function fetchJobHuntingEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const props = PropertiesService.getScriptProperties();
  
  const now = new Date();
  const twoWeeksAgo = new Date(now.getTime() - 14 * 24 * 60 * 60 * 1000);
  
  let lastUpdateStr = props.getProperty('LAST_UPDATE_TIME');
  let startTime;
  
  if (!lastUpdateStr) {
    startTime = twoWeeksAgo;
    console.log("初回実行として認識しました。直近2週間のメールを取得対象とします。");
  } else {
    startTime = new Date(Number(lastUpdateStr));
    if (startTime < twoWeeksAgo) {
      startTime = twoWeeksAgo;
      console.log("最終更新日時から2週間以上経過しているため、直近2週間のメールを取得対象とします。");
    } else {
      console.log(`前回の更新日時（${startTime}）以降のメールを取得対象とします。`);
    }
  }

  const nextUpdateTime = now.getTime().toString();
  const startEpoch = Math.floor(startTime.getTime() / 1000);
  const query = `after:${startEpoch} in:anywhere`;
  console.log(`Gmail検索クエリ: ${query}`);

  let threads = [];
  const maxThreadsToFetch = 1000;
  const batchSize = 500; 

  for (let start = 0; start < maxThreadsToFetch; start += batchSize) {
    const batch = GmailApp.search(query, start, batchSize);
    threads = threads.concat(batch);
    if (batch.length < batchSize) {
      break; 
    }
  }
  
  let newEmails = [];
  let processedCount = 0;
  
  for (let i = 0; i < threads.length; i++) {
    const messages = threads[i].getMessages();
    
    for (let j = 0; j < messages.length; j++) {
      const msg = messages[j];
      const date = msg.getDate();
      
      if (date > startTime && date >= twoWeeksAgo) {
        processedCount++;
        
        if (processedCount % 50 === 0 || processedCount === 1) {
           console.log(`メール読み込み中... 現在 ${processedCount} 通目を処理しています。`);
        }
        
        let body = msg.getPlainBody();
        if (body.length > 30000) {
            body = body.substring(0, 30000) + "\n...[文字数制限により以降省略されました]";
        }

        // --- マスキング処理の適用（件名と本文に実行） ---
        let maskedSubject = maskPersonalData(msg.getSubject());
        let maskedBody = maskPersonalData(body);

        newEmails.push({
          date: date,
          from: msg.getFrom(),       // 送信元はそのまま保持
          subject: maskedSubject,    // マスキング済みの件名
          body: maskedBody,          // マスキング済みの本文
          status: "未"
        });
        
        if (newEmails.length >= 1000) {
          break; 
        }
      }
    }
    if (newEmails.length >= 1000) {
      break;
    }
  }
  
  console.log(`新規取得メール数: ${newEmails.length} 通`);

  let existingData = [];
  const lastRow = sheet.getLastRow();
  
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, 5).setValues([["日時", "送信元", "件名", "本文", "処理状態"]]);
  }
  
  if (lastRow > 1) {
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 5);
    const values = dataRange.getValues();
    
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const rowDate = new Date(row[0]);
      
      if (rowDate >= twoWeeksAgo) {
        existingData.push({
          date: rowDate,
          from: row[1],
          subject: row[2],
          body: row[3],
          status: row[4]
        });
      }
    }
  }
  
  let allEmails = existingData.concat(newEmails);
  allEmails.sort((a, b) => b.date.getTime() - a.date.getTime());
  
  if (allEmails.length > 1000) {
    allEmails = allEmails.slice(0, 1000);
  }
  
  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, 5).clearContent();
  }
  
  if (allEmails.length > 0) {
    const outputValues = allEmails.map(email => [
      email.date,
      email.from,
      email.subject,
      email.body,
      email.status
    ]);
    
    sheet.getRange(2, 1, outputValues.length, 5).setValues(outputValues);
  }
  
  props.setProperty('LAST_UPDATE_TIME', nextUpdateTime);
  
  console.log(`処理完了: スプレッドシートに ${allEmails.length} 件のデータを反映しました。`);
}

/**
 * テキスト内の個人情報（氏名、住所、メールアドレス、電話番号、URL、ID）をマスキングする関数
 * @param {string} text - マスキング対象のテキスト
 * @return {string} マスキング後のテキスト
 */
function maskPersonalData(text) {
  if (!text) return "";
  
  let maskedText = text;

  // 1. 氏名のマスキング
  // フルネーム、スペース有無、姓のみ、名のみに対応
  maskedText = maskedText.replace(/成田\s*享市/g, '[NAME_HIDDEN]');
  maskedText = maskedText.replace(/成田/g, '[NAME_HIDDEN]');
  maskedText = maskedText.replace(/享市/g, '[NAME_HIDDEN]');

  // 2. 住所を含む行のマスキング
  // テキストを行ごとに分割し、判定を行う
  const lines = maskedText.split('\n');
  // 都道府県名 + 市区町村名に合致する正規表現
  const addressRegex = /(東京都|北海道|(?:京都|大阪)府|.{2,3}県)\s*.{1,10}(市|区|町|村)/;
  
  for (let i = 0; i < lines.length; i++) {
    if (addressRegex.test(lines[i])) {
      lines[i] = '[住所関連情報が含まれるため、この部分はマスキングされました]';
    }
  }
  maskedText = lines.join('\n');

  // 3. メールアドレスのマスキング
  maskedText = maskedText.replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, '[EMAIL_HIDDEN]');
  
  // 4. 電話番号のマスキング
  maskedText = maskedText.replace(/0\d{1,4}[-(]?\d{1,4}[-)]?\d{3,4}/g, '[PHONE_HIDDEN]');
  
  // 5. URLのマスキング
  maskedText = maskedText.replace(/https?:\/\/[^\s]+/g, '[URL_HIDDEN]');
  
  // 6. IDのマスキング
  maskedText = maskedText.replace(/(ID|ＩＤ|id|Id)[\s:：]*[a-zA-Z0-9_-]+/g, 'ID:[ID_HIDDEN]');

  return maskedText;
}