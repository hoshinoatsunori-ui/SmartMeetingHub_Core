// SmartMeetingHub_Core
// V3.7 2026/02/08
// Update: 高精度プロンプト実装 + 不要ファイル掃除機能

// ==========================================
// 1. 設定エリア
// ==========================================
const PROPS = PropertiesService.getScriptProperties();

const GEMINI_API_KEY = PROPS.getProperty('GEMINI_API_KEY');
const NOTION_API_KEY = PROPS.getProperty('NOTION_API_KEY');
const DB_ID_LOGS     = PROPS.getProperty('DB_ID_LOGS');
const DB_ID_ACTIONS  = PROPS.getProperty('DB_ID_ACTIONS');
const ADMIN_EMAIL    = PROPS.getProperty('ADMIN_EMAIL') || Session.getActiveUser().getEmail();

const DICTIONARY_SS_ID = PROPS.getProperty('DICTIONARY_SS_ID'); 
const LOG_SS_ID        = PROPS.getProperty('LOG_SS_ID');        

const INPUT_FOLDER_ID     = PROPS.getProperty('INPUT_FOLDER_ID');
const TARGET_FOLDER_ID    = PROPS.getProperty('TARGET_FOLDER_ID');
const LARGE_FILE_FOLDER_ID = PROPS.getProperty('LARGE_FILE_FOLDER_ID'); 

const DEBUG_MODE = (PROPS.getProperty('DEBUG_MODE') === 'true');
const MODEL_NAME = 'models/gemini-2.5-flash'; 

const PROPS_MAP = {
  logs: { id: '会議ID', title: '会議名', category: 'カテゴリ', date: '開催日', attendees: '参加者', summary: '要約' },
  actions: { id: '会議ID', task: 'タスク名', status: 'ステータス', assignee: '担当者', dueDate: '期限', category: 'カテゴリ', relation: 'Relation' }
};

let EMAIL_LOGS = [];

// ==========================================
// 2. メイン実行関数
// ==========================================

function main() {
  EMAIL_LOGS = [];
  Logger.log('[開始] 議事録生成プロセスを実行します。');
  
  if (!GEMINI_API_KEY || !DICTIONARY_SS_ID || !LOG_SS_ID) {
    Logger.log('[エラー] スクリプトプロパティ(APIキーまたはSS_ID)が設定されていません。');
    return;
  }

  const folder = DriveApp.getFolderById(INPUT_FOLDER_ID);
  const targetFolder = DriveApp.getFolderById(TARGET_FOLDER_ID);
  const largeFileFolder = DriveApp.getFolderById(LARGE_FILE_FOLDER_ID);

  const files = folder.getFiles();
  let currentIdCounter = parseInt(PROPS.getProperty('LAST_MEETING_ID') || '0', 10);

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    const mimeType = file.getMimeType();

    // 不要なテキストファイルの削除ロジック
    if (fileName.endsWith('.txt')) {
      Logger.log(`[掃除] 不要なテキストファイルを削除: ${fileName}`);
      file.setTrashed(true);
      continue;
    }

    if (mimeType === 'application/vnd.google-apps.script' || fileName.includes('【処理済】')) continue;
    if (!mimeType.startsWith('audio/') && !mimeType.startsWith('video/')) continue;

    // サイズチェック
    if (file.getSize() > 50 * 1024 * 1024) {
      file.moveTo(largeFileFolder);
      EMAIL_LOGS.push(`■ [退避] ${fileName} (50MB超過)`);
      break;
    }

    currentIdCounter++; 
    const currentMeetingId = currentIdCounter.toString().padStart(4, '0');
    let logInfo = { file: fileName, id: currentMeetingId, category: '-', title: '-', result: '実行中' };

    try {
      const description = file.getDescription() || "";
      if (description.includes("、")) {
        const parts = description.split("、");
        logInfo.category = parts[0].trim();
        logInfo.title = parts[1].trim(); 
      }

      let jsonString;
      if (DEBUG_MODE) {
        jsonString = JSON.stringify({ "title": "テスト", "date": "2026-02-08", "attendees": ["テスト"], "summary": "デバッグ", "actions": [] });
      } else {
        const fileUri = uploadToGeminiLargeFile(file.getId(), fileName, mimeType); 
        waitForFileActive(fileUri);
        jsonString = generateMeetingLogWithRetry(fileUri, mimeType, logInfo);
      }
      
      const data = JSON.parse(jsonString);
      createMeetingNotes(data, (logInfo.category !== '-') ? logInfo.category : null, logInfo.title !== '-' ? logInfo.title : data.title, currentMeetingId);
      
      PROPS.setProperty('LAST_MEETING_ID', currentIdCounter.toString());
      
      if (!DEBUG_MODE) {
        const safeTitle = (logInfo.title !== '-' ? logInfo.title : data.title).replace(/[\\/:*?"<>|]/g, '-');
        moveFileToNewFolder(file, targetFolder, `${currentMeetingId}_${safeTitle}`);
        file.setName(`【処理済】${fileName}`); 
      }
      logInfo.result = '✅ 成功';
    } catch (e) {
      logInfo.result = `❌ 失敗: ${e.toString()}`;
    } finally {
      EMAIL_LOGS.push(`■ ${logInfo.file}\n・成否: ${logInfo.result}`);
    }
    break; 
  }
  if (EMAIL_LOGS.length > 0) sendEmailLog();
}

// ==========================================
// 3. Gemini プロンプト & ログ保存
// ==========================================

function generateMeetingLog(fileUri, mimeType, logInfo) {
  const now = new Date();
  const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy年M月d日");
  const currentYear = now.getFullYear();
  const dictionary = getDictionaryTextFromSheet(logInfo.category);

  // 精度向上のための強化プロンプト
  const promptText = `
  # 役割
  あなたはプロジェクト管理に精通した、極めて正確な書記です。提供された録音データを一言一句逃さず分析し、構造化された議事録を作成してください。

  # 前提条件
  - 本日の日付: ${todayStr}
  - 基準年: 文中で年が明示されない場合は、原則として${currentYear}年として処理してください。

  # 参照用・正解辞書
  ${dictionary}
  ※「★推奨」とある人物は、今回の会議に最も深く関与している参加者候補です。
  ※ 音声で名前が呼ばれた際、上記辞書の「読み」に類似する場合は、必ず「正解の表記」を適用してください。

  # 思考ステップ（Step-by-Step）
  1. 会議の冒頭から参加者の自己紹介や呼びかけをすべて書き出し、辞書と照らし合わせます。
  2. 会議中に行われたすべての「提案」「合意」「依頼」を抽出し、誰が担当するかを文脈から判断します。
  3. 日付に関する発言（明日、来週、30日など）を、本日(${todayStr})を基準とした具体的な日付(YYYY-MM-DD)に計算し直します。

  # 抽出ガイドライン
  - **参加者 (attendees)**: 辞書の表記を優先。辞書にない人物は、聞こえた通りに漢字やカタカナで推測してください。
  - **アクションアイテム (actions)**: 
    - 「〜をやる」「〜をお願い」という明示的な発言だけでなく、議論の結果「必要となった作業」も担当者と共に抽出してください。
    - 担当者が曖昧な場合は、その発言の主導者を割り当ててください。
  - **要約 (summary)**: 何が決まり、次に何をする必要があるかを300文字以内で簡潔にまとめてください。

  # 出力形式 (JSONのみ)
  {
    "title": "具体的で分かりやすい会議タイトル",
    "date": "YYYY-MM-DD",
    "attendees": ["名前1", "名前2"],
    "summary": "要約内容",
    "actions": [
      {
        "task": "タスク内容（何をすべきか）",
        "assignee": "担当者名（辞書の正解表記を使用）",
        "due_date": "YYYY-MM-DD"
      }
    ]
  }`;

  const url = `https://generativelanguage.googleapis.com/v1beta/${MODEL_NAME}:generateContent?key=${GEMINI_API_KEY}`;
  const payload = { 
    "contents": [{ "parts": [{ "text": promptText }, { "file_data": { "mime_type": mimeType, "file_uri": fileUri } }] }],
    "generationConfig": { "response_mime_type": "application/json", "temperature": 0.1 }
  };

  const res = UrlFetchApp.fetch(url, { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true });
  const resText = res.getContentText();
  const resultJson = JSON.parse(resText);
  const finalAnswer = (resultJson.candidates && resultJson.candidates[0].content) ? resultJson.candidates[0].content.parts[0].text.trim() : "ERROR";
  
  // 専用スプレッドシートへログ保存
  saveExecutionLog(logInfo.id, logInfo.file, logInfo.category, promptText, finalAnswer);

  if (res.getResponseCode() !== 200) throw new Error(`Gemini Error: ${resText}`);
  return finalAnswer;
}

// ---------------------------------------------------------
// 以下の関数は前回のロジックを継承
// ---------------------------------------------------------

function getDictionaryTextFromSheet(currentCategory) {
  try {
    const ss = SpreadsheetApp.openById(DICTIONARY_SS_ID);
    const sheet = ss.getSheetByName('辞書');
    const values = sheet.getDataRange().getValues();
    const header = values[0];
    const catIdx = header.indexOf(currentCategory);
    let dictText = "";
    for (let i = 1; i < values.length; i++) {
      if (!values[i][1]) continue;
      const isTarget = (catIdx !== -1 && values[i][catIdx] === '○');
      dictText += `${isTarget ? '★推奨' : '-'} 読み:${values[i][0]} → 表記:${values[i][1]}\n`;
    }
    return dictText;
  } catch (e) { return "辞書取得エラー"; }
}

function saveExecutionLog(meetingId, fileName, category, prompt, response) {
  try {
    const ss = SpreadsheetApp.openById(LOG_SS_ID);
    let sheet = ss.getSheetByName('実行ログ');
    if (!sheet) {
      // シートがない場合は自動作成を試みる
      sheet = ss.insertSheet('実行ログ');
      sheet.appendRow(['日時', 'ID', 'ファイル', 'カテゴリ', 'プロンプト', '応答']);
    }
    sheet.appendRow([new Date(), meetingId, fileName, category, prompt, response]);
    Logger.log(`[成功] 実行ログをスプレッドシートに保存しました。`);
  } catch (e) { 
    // ここで具体的なエラー内容をログに出す
    Logger.log(`[重大エラー] ログ保存に失敗しました: ${e.toString()}`); 
  }
}

function generateMeetingLogWithRetry(fileUri, mimeType, logInfo) {
  const maxRetries = 2;
  for (let i = 0; i <= maxRetries; i++) {
    try { return generateMeetingLog(fileUri, mimeType, logInfo); }
    catch (e) { if (i === maxRetries) throw e; Utilities.sleep(30000); }
  }
}

function uploadToGeminiLargeFile(fileId, fileName, mimeType) {
  const file = DriveApp.getFileById(fileId);
  const initUrl = `https://generativelanguage.googleapis.com/upload/v1beta/files?key=${GEMINI_API_KEY}`;
  const metadata = { file: { display_name: fileName } };
  const initRes = UrlFetchApp.fetch(initUrl, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(metadata),
    headers: { 'X-Goog-Upload-Protocol': 'resumable', 'X-Goog-Upload-Command': 'start', 'X-Goog-Upload-Header-Content-Length': file.getSize().toString(), 'X-Goog-Upload-Header-Content-Type': mimeType }
  });
  const res = UrlFetchApp.fetch(initRes.getAllHeaders()['x-goog-upload-url'], { method: 'post', payload: file.getBlob(), headers: { 'X-Goog-Upload-Protocol': 'resumable', 'X-Goog-Upload-Command': 'upload, finalize', 'X-Goog-Upload-Offset': '0' } });
  return JSON.parse(res.getContentText()).file.uri;
}

function waitForFileActive(fileUri) {
  const name = fileUri.split('/files/')[1];
  let state = 'PROCESSING';
  while (state === 'PROCESSING') {
    Utilities.sleep(5000);
    state = JSON.parse(UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/files/${name}?key=${GEMINI_API_KEY}`).getContentText()).state;
  }
}

function createMeetingNotes(data, category, title, meetingId) {
  const payload = {
    parent: { database_id: DB_ID_LOGS },
    properties: {
      [PROPS_MAP.logs.title]: { title: [{ text: { content: title } }] },
      [PROPS_MAP.logs.attendees]: { multi_select: toMultiSelectOptions(data.attendees) },
      [PROPS_MAP.logs.summary]: { rich_text: [{ text: { content: data.summary || "" } }] },
      [PROPS_MAP.logs.id]: { rich_text: [{ text: { content: meetingId } }] }
    }
  };
  if (data.date) payload.properties[PROPS_MAP.logs.date] = { date: { start: data.date } };
  if (category) payload.properties[PROPS_MAP.logs.category] = { select: { name: category } };
  const res = callNotionApi(payload);
  if (res && data.actions) data.actions.forEach(act => {
    const actPayload = {
      parent: { database_id: DB_ID_ACTIONS },
      properties: {
        [PROPS_MAP.actions.task]: { title: [{ text: { content: act.task } }] },
        [PROPS_MAP.actions.assignee]: { multi_select: toMultiSelectOptions(act.assignee) },
        [PROPS_MAP.actions.status]: { status: { name: '未着手' } },
        [PROPS_MAP.actions.relation]: { relation: [{ id: res.id }] },
        [PROPS_MAP.actions.id]: { rich_text: [{ text: { content: meetingId } }] }
      }
    };
    if (act.due_date) actPayload.properties[PROPS_MAP.actions.dueDate] = { date: { start: act.due_date } };
    if (category) actPayload.properties[PROPS_MAP.actions.category] = { select: { name: category } };
    callNotionApi(actPayload);
  });
}

function callNotionApi(payload) {
  const options = { method: 'post', headers: { 'Authorization': `Bearer ${NOTION_API_KEY}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' }, payload: JSON.stringify(payload), muteHttpExceptions: true };
  const res = UrlFetchApp.fetch('https://api.notion.com/v1/pages', options);
  return res.getResponseCode() === 200 ? JSON.parse(res.getContentText()) : null;
}

function toMultiSelectOptions(input) {
  const arr = Array.isArray(input) ? input : (input ? input.toString().split(/,|、/) : []);
  return arr.map(s => s.trim()).filter(s => s).map(s => ({ name: s }));
}

function sendEmailLog() { GmailApp.sendEmail(ADMIN_EMAIL, `【議事録bot】実行完了報告`, EMAIL_LOGS.join('\n---\n')); }

function moveFileToNewFolder(file, parent, name) { parent.createFolder(name).addFile(file); parent.removeFile(file); }
