// SmartMeetingHub_Core
// V3.2 2026/01/29
// Update: 1回の実行で1ファイルのみ処理するよう変更（タイムアウト対策）

// ==========================================
// 1. 設定エリア
// ==========================================
const PROPS = PropertiesService.getScriptProperties();

const GEMINI_API_KEY = PROPS.getProperty('GEMINI_API_KEY');
const NOTION_API_KEY = PROPS.getProperty('NOTION_API_KEY');
const DB_ID_LOGS     = PROPS.getProperty('DB_ID_LOGS');
const DB_ID_ACTIONS  = PROPS.getProperty('DB_ID_ACTIONS');
const ADMIN_EMAIL    = PROPS.getProperty('ADMIN_EMAIL') || Session.getActiveUser().getEmail();

const INPUT_FOLDER_ID    = PROPS.getProperty('INPUT_FOLDER_ID');
const TARGET_FOLDER_ID   = PROPS.getProperty('TARGET_FOLDER_ID');
const LARGE_FILE_FOLDER_ID = PROPS.getProperty('LARGE_FILE_FOLDER_ID'); 

const debugVal = PROPS.getProperty('DEBUG_MODE');
const DEBUG_MODE = (debugVal && debugVal.trim().toLowerCase() === 'true');

const MODEL_NAME = 'models/gemini-2.5-flash'; // 2026年時点の最新推奨モデルへ修正

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
  
  const modeMsg = DEBUG_MODE 
    ? '【🚧 デバッグモード有効】Gemini解析スキップ / ファイル移動なし' 
    : '【▶︎ 通常モード】1件の処理を実行します';
  Logger.log(`[開始] ${modeMsg}`);

  if (!GEMINI_API_KEY || !NOTION_API_KEY || !DB_ID_LOGS || !DB_ID_ACTIONS || !INPUT_FOLDER_ID || !TARGET_FOLDER_ID || !LARGE_FILE_FOLDER_ID) {
    Logger.log('[エラー] スクリプトプロパティの設定を確認してください。');
    return;
  }

  const folder = DriveApp.getFolderById(INPUT_FOLDER_ID);
  const targetFolder = DriveApp.getFolderById(TARGET_FOLDER_ID);
  const largeFileFolder = DriveApp.getFolderById(LARGE_FILE_FOLDER_ID);

  const files = folder.getFiles();
  let processedCount = 0;
  let currentIdCounter = parseInt(PROPS.getProperty('LAST_MEETING_ID') || '0', 10);

  // --------------------------------------------------
  // ループ内で1件見つけたら処理して break する
  // --------------------------------------------------
  while (files.hasNext()) {
    const file = files.next();
    const mimeType = file.getMimeType();

    // 対象外スキップ（これらは「1件」にカウントしない）
    if (mimeType === 'application/vnd.google-apps.script') continue;
    if (!mimeType.startsWith('audio/') && !mimeType.startsWith('video/')) continue;
    if (file.getName().includes('【処理済】') || file.getName().includes('【サイズ超過】')) continue;

    // --- A. サイズチェック & 退避 ---
    const fileSize = file.getSize();
    if (fileSize > 50 * 1024 * 1024) {
      const sizeMB = Math.round(fileSize / 1024 / 1024);
      Logger.log(`[サイズ超過] ${file.getName()} (${sizeMB}MB) -> 退避`);
      
      if (!DEBUG_MODE) {
        file.setName(`【サイズ超過】${file.getName()}`);
        file.moveTo(largeFileFolder);
        EMAIL_LOGS.push(`■ [退避] ${file.getName()} (50MB超過)`);
      }
      // 1件「処理（退避）」したのでループを抜ける
      break; 
    }

    // --- B. 通常の解析処理 ---
    Logger.log(`[処理開始] ${file.getName()}`);
    currentIdCounter++; 
    const currentMeetingId = currentIdCounter.toString().padStart(4, '0');

    let logInfo = { file: file.getName(), id: currentMeetingId, category: '-', title: '-', result: '処理中' };

    try {
      const description = file.getDescription() || "";
      if (description.includes("、")) {
        const parts = description.split("、");
        if (parts.length >= 2) {
          logInfo.category = parts[0].trim();
          logInfo.title = parts[1].trim(); 
        }
      }

      let jsonString;
      if (DEBUG_MODE) {
        jsonString = JSON.stringify({
          "title": "【デバッグ】テスト", "date": "2026-01-01", "attendees": ["テスト"], 
          "summary": "デバッグ中...", "actions": []
        });
      } else {
        const fileUri = uploadToGeminiLargeFile(file.getId(), file.getName(), mimeType); 
        waitForFileActive(fileUri);
        jsonString = generateMeetingLogWithRetry(fileUri, mimeType);
      }
      
      if (!jsonString) throw new Error("Geminiからの回答が空でした");

      const data = JSON.parse(jsonString);
      const dateMatch = file.getName().match(/^(\d{4})(\d{2})(\d{2})/);
      if (dateMatch) {
        const fileDate = `${dateMatch[1]}-${dateMatch[2]}-${dateMatch[3]}`;
        if (isValidDate(fileDate)) data.date = fileDate;
      }

      const finalTitle = (logInfo.title !== '-') ? logInfo.title : (data.title || file.getName());
      logInfo.title = finalTitle; 

      // Notion登録
      createMeetingNotes(data, (logInfo.category !== '-') ? logInfo.category : null, finalTitle, currentMeetingId);
      PROPS.setProperty('LAST_MEETING_ID', currentIdCounter.toString());

      if (!DEBUG_MODE) {
        const folderName = `${currentMeetingId}_${finalTitle}`.replace(/[\\/:*?"<>|]/g, '-'); 
        moveFileToNewFolder(file, targetFolder, folderName);
        file.setName(`【処理済】${file.getName()}`); 
        logInfo.result = '✅ 成功';
        processedCount++;
      } else {
        logInfo.result = '✅ 成功 (DEBUG)';
      }

    } catch (e) {
      logInfo.result = `❌ 失敗: ${e.toString()}`;
      Logger.log(`[エラー詳細] ${e.stack}`); 
    } finally {
      EMAIL_LOGS.push(`■ 処理ファイル: ${logInfo.file}\n・会議ID: ${logInfo.id}\n・成否: ${logInfo.result}`);
    }

    // 1件処理が終わったのでループを終了
    break; 
  }

  if (EMAIL_LOGS.length > 0) {
    sendEmailLog(processedCount);
  } else {
    Logger.log('[情報] 処理対象ファイルはありませんでした。');
  }
}

// ==========================================
// 3. ユーティリティ
// ==========================================

function isValidDate(dateString) {
  if (!dateString) return false;
  const regex = /^\d{4}-\d{2}-\d{2}$/;
  if (!regex.test(dateString)) return false;
  
  const date = new Date(dateString);
  const timestamp = date.getTime();
  if (typeof timestamp !== 'number' || Number.isNaN(timestamp)) return false;
  
  return date.toISOString().startsWith(dateString);
}

function sendEmailLog(processedCount) {
  const subject = `【議事録bot】処理レポート (${processedCount}件成功)`;
  const body = EMAIL_LOGS.join('\n----------------------------------\n');
  try {
    GmailApp.sendEmail(ADMIN_EMAIL, subject, body);
  } catch (e) {
    Logger.log(`[メール送信失敗] ${e.toString()}`);
  }
}

function moveFileToNewFolder(file, parentFolder, newFolderName) {
  try {
    const newFolder = parentFolder.createFolder(newFolderName);
    file.moveTo(newFolder);
  } catch (e) {
    throw new Error(`フォルダ移動失敗: ${e.toString()}`);
  }
}

function toMultiSelectOptions(input) {
  if (!input) return [];
  let candidates = [];
  if (Array.isArray(input)) {
    candidates = input;
  } else {
    candidates = input.toString().split(/,|、/);
  }
  return candidates
    .map(s => s.trim())
    .filter(s => s.length > 0)
    .map(s => ({ name: s }));
}

// ==========================================
// 4. Gemini 関連関数
// ==========================================

function uploadToGeminiLargeFile(fileId, fileName, mimeType) {
  const fileForSize = DriveApp.getFileById(fileId);
  const fileSize = fileForSize.getSize();
  Logger.log(`[アップロード] ${fileName} (${Math.round(fileSize / 1024 / 1024 * 10) / 10}MB)`);

  const initUrl = `https://generativelanguage.googleapis.com/upload/v1beta/files?key=${GEMINI_API_KEY}`;
  const metadata = { file: { display_name: fileName } };
  
  const initRes = UrlFetchApp.fetch(initUrl, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(metadata),
    headers: {
      'X-Goog-Upload-Protocol': 'resumable', 'X-Goog-Upload-Command': 'start',
      'X-Goog-Upload-Header-Content-Length': fileSize.toString(), 'X-Goog-Upload-Header-Content-Type': mimeType
    }
  });

  const uploadUrl = initRes.getAllHeaders()['x-goog-upload-url'];
  
  const CHUNK_SIZE = 8 * 1024 * 1024; 
  let offset = 0;
  let fileUri = null;
  const token = ScriptApp.getOAuthToken();

  while (offset < fileSize) {
    const end = Math.min(offset + CHUNK_SIZE, fileSize);
    const isFinal = (end === fileSize);
    
    const downloadUrl = `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`;
    let chunkBlob;
    
    try {
      const chunkRes = UrlFetchApp.fetch(downloadUrl, {
        headers: { 'Authorization': 'Bearer ' + token, 'Range': `bytes=${offset}-${end - 1}` },
        muteHttpExceptions: true
      });
      if (chunkRes.getResponseCode() !== 206 && chunkRes.getResponseCode() !== 200) {
        throw new Error(`Driveダウンロード失敗 Code:${chunkRes.getResponseCode()}`);
      }
      chunkBlob = chunkRes.getBlob();
    } catch (e) {
      throw new Error(`Driveデータ取得失敗: ${e.toString()}`);
    }

    const command = isFinal ? 'upload, finalize' : 'upload';
    let uploadSuccess = false;
    let retryCount = 0;
    
    while (!uploadSuccess && retryCount < 3) {
      try {
        const response = UrlFetchApp.fetch(uploadUrl, {
          method: 'post', payload: chunkBlob, 
          headers: { 
            'X-Goog-Upload-Protocol': 'resumable', 'X-Goog-Upload-Command': command, 'X-Goog-Upload-Offset': offset.toString()
          },
          muteHttpExceptions: true
        });

        const code = response.getResponseCode();
        if (code === 308 || code === 200 || code === 201) {
          uploadSuccess = true;
          if (isFinal) {
            const json = JSON.parse(response.getContentText());
            if (json.file && json.file.uri) {
              fileUri = json.file.uri;
              Logger.log(`[アップロード完了] URI: ${fileUri}`);
            }
          }
        } else {
          Logger.log(`[通信リトライ] Offset:${offset} Code:${code}`);
          retryCount++;
          Utilities.sleep(2000);
        }
      } catch (e) {
        Logger.log(`[通信例外] ${e.toString()}`);
        retryCount++;
        Utilities.sleep(2000);
      }
    }

    if (!uploadSuccess) throw new Error(`アップロード失敗: Offset ${offset}`);
    offset = end;
  }
  
  if (!fileUri) {
    Logger.log('[警告] URI取得失敗。一覧検索を試行します。');
    Utilities.sleep(3000);
    return getLatestFileUri(fileName);
  }
  return fileUri;
}

function getLatestFileUri(displayName) {
  const url = `https://generativelanguage.googleapis.com/v1beta/files?key=${GEMINI_API_KEY}`;
  const res = UrlFetchApp.fetch(url);
  const json = JSON.parse(res.getContentText());
  if (json.files && json.files.length > 0) {
    const target = json.files.find(f => f.displayName === displayName);
    if (target) return target.uri;
    return json.files[0].uri;
  }
  throw new Error("URI取得失敗");
}

function waitForFileActive(fileUri) {
  let state = 'PROCESSING';
  let attempts = 0; 
  const name = fileUri.split('/files/')[1];
  while (state === 'PROCESSING' && attempts < 60) {
    Utilities.sleep(5000);
    const res = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/files/${name}?key=${GEMINI_API_KEY}`);
    state = JSON.parse(res.getContentText()).state;
    attempts++;
  }
  if (state !== 'ACTIVE') throw new Error('解析準備タイムアウト');
}

function generateMeetingLogWithRetry(fileUri, mimeType) {
  const maxRetries = 3;
  let attempt = 0;
  while (attempt < maxRetries) {
    try {
      return generateMeetingLog(fileUri, mimeType);
    } catch (e) {
      if (e.toString().includes("429")) {
        attempt++;
        Logger.log(`[警告] API制限 (429)。60秒待機... (${attempt}/${maxRetries})`);
        Utilities.sleep(60000); 
      } else {
        throw e;
      }
    }
  }
  throw new Error("リトライ上限到達");
}

function generateMeetingLog(fileUri, mimeType) {
  const now = new Date();
  const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy年M月d日");
  const currentYear = now.getFullYear();
  const url = `https://generativelanguage.googleapis.com/v1beta/${MODEL_NAME}:generateContent?key=${GEMINI_API_KEY}`;
  
  // 精度向上のための詳細なシステム指示
  const promptText = `
  # 役割
  あなたは非常に優秀なエグゼクティブアシスタントです。提供された録音データから、正確な議事録を作成してください。

  # 前提条件
  - 本日の日付: ${todayStr}
  - 年が不明な日付は、原則として${currentYear}年として扱ってください。

  # 抽出のガイドライン
  1. **参加者 (attendees)**: 
     - 挨拶や自己紹介、発言内容から、会議に参加している全員の名前を抽出してください。
     - 名字だけでなくフルネームがわかる場合はフルネームで記載してください。
  2. **アクションアイテム (actions)**:
     - 誰かが「やります」「お願いします」と言ったタスクを漏らさず抽出してください。
     - **重要**: 担当者が明言されていないが、文脈から判断できる場合はその人を記載してください。
     - **重要**: 期限が「来週中」「今月末」などの相対的な表現の場合、本日(${todayStr})を基準に具体的な日付(YYYY-MM-DD)へ変換してください。
  3. **要約 (summary)**:
     - 決定事項を中心に、議論の経緯がわかるようにまとめてください。

  # 出力形式 (JSONのみ)
  {
    "title": "会議の目的がわかる具体的なタイトル",
    "date": "YYYY-MM-DD",
    "attendees": ["名前1", "名前2"],
    "summary": "要約テキスト（300文字以内）",
    "actions": [
      {
        "task": "具体的なタスク内容（〜を作成する、〜に連絡するなど）",
        "assignee": "担当者名",
        "due_date": "YYYY-MM-DD（不明な場合は空文字）"
      }
    ]
  }`;

  // response_mime_type を指定して JSON 出力を強制する設定を追加
  const payload = { 
    "contents": [{ 
      "parts": [
        { "text": promptText }, 
        { "file_data": { "mime_type": mimeType, "file_uri": fileUri } }
      ] 
    }],
    "generationConfig": {
      "response_mime_type": "application/json"
    }
  };

  const response = UrlFetchApp.fetch(url, { 
    "method": "post", 
    "contentType": "application/json", 
    "payload": JSON.stringify(payload), 
    "muteHttpExceptions": true 
  });
  
  if (response.getResponseCode() !== 200) throw new Error(`Gemini API Error: ${response.getContentText()}`);
  
  const json = JSON.parse(response.getContentText());
  if (json.candidates && json.candidates[0].content) {
    // generationConfigでJSON指定しているため、バッククォート除去の必要性が低くなります
    return json.candidates[0].content.parts[0].text.trim();
  }
  return null;
}

// ==========================================
// 5. Notion 関連関数 (日付エラー対策済)
// ==========================================

function createMeetingNotes(data, category, fixedTitle, meetingId) {
  if (fixedTitle) data.title = fixedTitle;
  if (category) data.category = category;
  data.meetingId = meetingId;

  const logPageId = createLogPage(data);
  if (logPageId) {
    if (data.actions && data.actions.length > 0) {
      createActionPages(data.actions, logPageId, category, meetingId);
    }
  }
}

function createLogPage(data) {
  const payload = { parent: { database_id: DB_ID_LOGS }, properties: {} };
  
  payload.properties[PROPS_MAP.logs.title] = { title: [{ text: { content: data.title } }] };
  
  if (data.date && isValidDate(data.date)) {
    payload.properties[PROPS_MAP.logs.date] = { date: { start: data.date } };
  } else {
    Logger.log(`[警告] 議事録の日付が無効なため空欄にします: ${data.date}`);
  }

  payload.properties[PROPS_MAP.logs.attendees] = { multi_select: toMultiSelectOptions(data.attendees) };
  payload.properties[PROPS_MAP.logs.summary] = { rich_text: [{ text: { content: data.summary } }] };
  
  if (data.category) {
    payload.properties[PROPS_MAP.logs.category] = { select: { name: data.category } };
  }
  if (data.meetingId) {
    payload.properties[PROPS_MAP.logs.id] = { rich_text: [{ text: { content: data.meetingId } }] };
  }

  const res = callNotionApi(payload);
  return res ? res.id : null;
}

function createActionPages(actions, logPageId, category, meetingId) {
  actions.forEach(action => {
    const payload = { parent: { database_id: DB_ID_ACTIONS }, properties: {} };
    
    payload.properties[PROPS_MAP.actions.task] = { title: [{ text: { content: action.task } }] };
    payload.properties[PROPS_MAP.actions.status] = { status: { name: '未着手' } };
    payload.properties[PROPS_MAP.actions.assignee] = { multi_select: toMultiSelectOptions(action.assignee) };
    
    if (action.due_date && isValidDate(action.due_date)) {
      payload.properties[PROPS_MAP.actions.dueDate] = { date: { start: action.due_date } };
    } else {
      Logger.log(`[警告] タスクの期限が無効なため空欄にします: ${action.due_date}`);
    }

    payload.properties[PROPS_MAP.actions.relation] = { relation: [{ id: logPageId }] };
    
    if (category) {
      payload.properties[PROPS_MAP.actions.category] = { select: { name: category } };
    }
    if (meetingId) {
      payload.properties[PROPS_MAP.actions.id] = { rich_text: [{ text: { content: meetingId } }] };
    }

    callNotionApi(payload);
  });
}

function callNotionApi(payload) {
  const url = 'https://api.notion.com/v1/pages';
  const options = {
    method: 'post',
    headers: { 'Authorization': `Bearer ${NOTION_API_KEY}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    const errText = response.getContentText();
    throw new Error(`Notion API Error: ${errText}`);
  }
  return JSON.parse(response.getContentText());
}

/**
 * 現在のAPIキーで利用可能なGeminiモデルの一覧をログに出力します
 */
function listGeminiModels() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  
  if (!apiKey) {
    Logger.log("エラー: スクリプトプロパティ 'GEMINI_API_KEY' が設定されていません。");
    return;
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;

  try {
    const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());

    if (json.error) {
      Logger.log("APIエラー: " + json.error.message);
      return;
    }

    Logger.log("=== 利用可能なモデル一覧 ===");
    json.models.forEach(model => {
      // テキスト生成(generateContent)に対応しているモデルのみ抽出
      if (model.supportedGenerationMethods.includes("generateContent")) {
        Logger.log(`名称: ${model.name}`);
        Logger.log(`説明: ${model.description}`);
        Logger.log("-----------------------------------");
      }
    });
    Logger.log("============================");

  } catch (e) {
    Logger.log("通信エラー: " + e.toString());
  }
}


