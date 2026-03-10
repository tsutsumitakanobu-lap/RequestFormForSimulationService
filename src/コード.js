/**
 * Main function executed on form submission.
 * Handles data logging, folder creation, file management, AI diagnosis, and URL generation.
 */
function createFolderAndMoveFiles(e) {
  // --- [1. CONFIGURATION & CONSTANTS] ---
  const CONFIG = {
    PARENT_FOLDER_ID: '1ZyCfHWZw35Bpx-a8zwlgBLLfQtQBwcjmGqSc3dEvclnAYrZD8j54hTZ7JNopfunDV_IH576s',
    ADDRESS_BOOK_ID: '1MEtptfvwTvmC6YhRoQIAyU5-2VO3DvFjTtffdUTi35c',
    ADDRESS_BOOK_SHEET_NAME: 'シート1',
    LOG_SS_ID: '1i9OesRVOtN5_fY8ntCEp_fHsMh_pefRawlaYZAZMbVo',
    LOG_SHEET_NAME: 'RequestLog',
    MITUMORI_PROMPT_FILE_ID: '1liF9c709fvFYvLGzg2_QHC8UWLpJ6OrW',
    SEISHIKI_PROMPT_FILE_ID: '1Eu8XlG-0IKIft6hDb3WAXm4DMG6V4brj',
    GEMINI_API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'),
    ADMIN_EMAIL: PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL')
  };

  // Critical Item IDs based on the provided IDTable
  const IDS = {
    FORM_TYPE: '2015761906',       // 依頼種別
    SALES_REP: '1973732404',        // ラプラス・システム担当営業
    PLANT_NAME: '1222779165',      // 発電所名
    FOLDER_ID_HIDDEN: '77031512',  // フォルダID (Hidden field)
    MITUMORI_CLIENT: '1732382986', // 見積もり依頼者氏名
    SEISHIKI_CLIENT: '323456789'   // 依頼者氏名 (Example ID, update if needed)
  };

  /**
   * MAPPING: Mitumori Item ID -> Seishiki Entry ID
   * This ensures high maintainability for pre-filling formal requests.
   */
  const URL_PARAM_MAP = {
    '393957065': 'entry.393957065',   // Affiliation
    '1732382986': 'entry.1732382986', // Client Name
    '292916935': 'entry.292916935',   // Sales Rep
    '1222779165': 'entry.1222779165', // Plant Name
    '625473123': 'entry.625473123',   // Address
    '880911839': 'entry.880911839',   // LatLon
    '659329608': 'entry.659329608',   // Sim Type
    '485460435': 'entry.485460435'    // Notes
  };

  try {
    if (!e) return;
    const form = e.source;
    const response = e.response;
    const itemResponses = response.getItemResponses();
    
    let answersMap = {};       // Key: Item ID
    let titleAnswersMap = {};  // Key: Question Title
    let filesToMove = [];
    let fileNames = [];

    // --- [2. PARSE RESPONSES & CLEAN FILE NAMES] ---
    itemResponses.forEach(itemRes => {
      const item = itemRes.getItem();
      const id = item.getId().toString();
      const title = item.getTitle();
      let val = itemRes.getResponse();

      // Handle File Uploads: Clean names and collect IDs
      if (item.getType() === FormApp.ItemType.FILE_UPLOAD && val) {
        const fileIds = Array.isArray(val) ? val : [val];
        let cleaned = fileIds.map(fid => {
          let file = DriveApp.getFileById(fid);
          let newName = file.getName().replace(/(.+) - [^-]+(\.[a-zA-Z0-9]+)$/, '$1$2');
          file.setName(newName);
          fileNames.push(newName);
          filesToMove.push(fid);
          return newName;
        });
        val = cleaned.join('\n');
      }
      answersMap[id] = val;
      titleAnswersMap[title] = val;
    });

    const isMitumori = (answersMap[IDS.FORM_TYPE] === "見積もり依頼");

    // --- [3. DATA MANAGEMENT (REQ 1)] ---
    // Log all data to spreadsheet using fixed columns based on form items
    logToMasterDatabase(CONFIG, form, answersMap);

    // --- [4. FOLDER RESOLUTION] ---
    const parentFolder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
    let targetFolder;
    const existingFolderId = answersMap[IDS.FOLDER_ID_HIDDEN];

    if (!isMitumori && existingFolderId) {
      try { targetFolder = DriveApp.getFolderById(existingFolderId); } catch(err) {}
    }
    if (!targetFolder) {
      const timeStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd_HHmm');
      const folderName = `${answersMap[IDS.PLANT_NAME] || 'UnknownPlant'}_${timeStr}`;
      targetFolder = parentFolder.createFolder(folderName);
    }

    // --- [5. GENERATE PRE-FILLED URL (REQ 2)] ---
    let prefilledUrl = "";
    if (isMitumori) {
      const baseUrl = form.getPublishedUrl().replace("viewform", "viewform?");
      let queryParts = [`entry.1739104995=${encodeURIComponent("正式依頼")}`]; // Set to Formal
      queryParts.push(`entry.77031512=${encodeURIComponent(targetFolder.getId())}`); // Link folder

      // Dynamically add mapped parameters
      for (let itemId in URL_PARAM_MAP) {
        if (answersMap[itemId]) {
          queryParts.push(`${URL_PARAM_MAP[itemId]}=${encodeURIComponent(answersMap[itemId])}`);
        }
      }
      prefilledUrl = baseUrl + queryParts.join("&");
    }

    // --- [6. PDF GENERATION & AI ANALYSIS] ---
    let pdfContent = "【シミュレーション依頼 回答内容まとめ】\n\n";
    for (let title in titleAnswersMap) {
      if (title === "フォルダID") continue;
      let val = titleAnswersMap[title];
      // Keep "Sama" logic
      if (title.includes("氏名") && val && !val.endsWith("様")) val += " 様";
      pdfContent += `■ ${title}\n${val}\n\n`;
    }

    const aiAnalysis = callGeminiAI(pdfContent, filesToMove, isMitumori, CONFIG);
    
    if (isMitumori) {
      pdfContent += `\n--------------------------------------------\n【正式依頼用URL】\n${prefilledUrl}\n--------------------------------------------\n`;
    }
    pdfContent += `\n【AI事前診断】\n${aiAnalysis}`;

    // Create and save PDF
    const pdfBlob = createPDF(pdfContent, answersMap[IDS.PLANT_NAME] || 'Request', targetFolder);

    // Move uploaded files to the target folder
    filesToMove.forEach(fid => { try { DriveApp.getFileById(fid).moveTo(targetFolder); } catch(e) {} });

    // --- [7. NOTIFICATIONS] ---
    sendNotification(CONFIG, titleAnswersMap, targetFolder.getUrl(), aiAnalysis, pdfBlob, isMitumori, prefilledUrl);

  } catch (err) {
    console.error("Critical Error: " + err.stack);
    GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, "GAS Error Alert", err.stack);
  }
}

/**
 * Logs response to a Master Database sheet with fixed columns for every question.
 */
function logToMasterDatabase(config, form, answersMap) {
  const ss = SpreadsheetApp.openById(config.LOG_SS_ID);
  let sheet = ss.getSheetByName(config.LOG_SHEET_NAME);
  const items = form.getItems();
  
  if (sheet.getLastRow() === 0) {
    const headers = ["Timestamp"].concat(items.map(i => i.getTitle()));
    sheet.appendRow(headers);
  }

  // Row structure: [Timestamp, Ques1, Ques2, ..., Ques57]
  let rowData = [new Date()];
  items.forEach(item => {
    const id = item.getId().toString();
    rowData.push(answersMap[id] || ""); // Fills blank if question was skipped
  });

  sheet.appendRow(rowData);
}

/**
 * Creates a PDF in the specified folder and trashes the temporary Doc.
 */
function createPDF(text, plantName, folder) {
  const name = `Details_${plantName}_${Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd_HHmm")}`;
  const doc = DocumentApp.create(name);
  doc.getBody().setText(text);
  doc.saveAndClose();
  const blob = DriveApp.getFileById(doc.getId()).getAs('application/pdf').setName(name + ".pdf");
  const file = folder.createFile(blob);
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  return file;
}

/**
 * Call Gemini API (Refactored to use CONFIG object)
 * @param {string} content - Text content to analyze
 * @param {Array} fileIds - IDs of uploaded PDF files
 * @param {boolean} isMitumori - Flag for request type
 * @param {Object} config - Configuration object
 */
function callGeminiAI(content, fileIds, isMitumori, config) {
  if (!config.GEMINI_API_KEY) return "API key is not set.";
  
  const modelName = "gemini-3-flash-preview"; 
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${config.GEMINI_API_KEY}`;
  
  let roleInstruction = "";
  try {
    const targetFileId = isMitumori ? config.MITUMORI_PROMPT_FILE_ID : config.SEISHIKI_PROMPT_FILE_ID;
    roleInstruction = DriveApp.getFileById(targetFileId).getBlob().getDataAsString("utf-8");
  } catch (e) {
    console.error("Failed to load prompt file: " + e.message);
    return "AI diagnosis error: Failed to load prompt settings.";
  }

  const prompt = `You are a solar power simulation expert.\n${roleInstruction}\n\n# Request Details:\n${content}`;

  let parts = [{ "text": prompt }];
  
  // Attach up to 2 PDFs for analysis
  if (fileIds && fileIds.length > 0) {
    let pdfCount = 0;
    for (let id of fileIds) {
      if (pdfCount >= 2) break; 
      try {
        const file = DriveApp.getFileById(id);
        if (file.getMimeType() === "application/pdf") {
          const base64Data = Utilities.base64Encode(file.getBlob().getBytes());
          parts.push({ "inline_data": { "mime_type": "application/pdf", "data": base64Data } });
          pdfCount++;
        }
      } catch (fileErr) {}
    }
  }

  const payload = { "contents": [{ "parts": parts }] };
  const options = { 
    "method": "post", 
    "contentType": "application/json", 
    "payload": JSON.stringify(payload), 
    "muteHttpExceptions": true 
  };
  
  // Retry Logic
  const maxRetries = 3;
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const res = UrlFetchApp.fetch(url, options);
      const code = res.getResponseCode();
      if (code === 200) {
        return JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
      } else if (attempt === maxRetries) {
        return `AI Error (${code})`;
      }
      Utilities.sleep(3000); 
    } catch (e) { 
      if (attempt === maxRetries) return "AI Connection Error: " + e.message; 
      Utilities.sleep(3000); 
    }
  }
}

/**
 * Sends notification email (Extracted from the original Phase 7 logic)
 */
function sendNotification(config, titleMap, folderUrl, aiText, pdfBlob, isMitumori, preUrl) {
  let targetEmail = "";
  const salesRepName = titleMap['ラプラス・システム担当営業'];
  const companyName = isMitumori ? titleMap['見積もり依頼者所属'] : titleMap['依頼者所属'];
  const clientName = isMitumori ? titleMap['見積もり依頼者氏名'] : titleMap['依頼者氏名'];
  const simType = isMitumori ? titleMap['シミュレーション種別（見積もり）'] : titleMap['シミュレーション種別'];

  // Address book lookup
  if (salesRepName) {
    try {
      const bookSs = SpreadsheetApp.openById(config.ADDRESS_BOOK_ID);
      const data = bookSs.getSheetByName(config.ADDRESS_BOOK_SHEET_NAME).getDataRange().getValues();
      for (let r = 0; r < data.length; r++) {
        if (data[r][0] == salesRepName) { targetEmail = data[r][1]; break; }
      }
    } catch (e) { console.error("Address book lookup failed"); }
  }

  if (!targetEmail) targetEmail = config.ADMIN_EMAIL;

  const typeLabel = isMitumori ? "【見積依頼】" : "【正式依頼】";
  const subject = `${typeLabel} シミュレーション代行: ${companyName || "No Name"}`;
  
  // Add "Sama" for the email body display if not present
  const clientDisplayName = (clientName && !clientName.endsWith("様")) ? clientName + " 様" : clientName;

  let htmlBody = `<p>${salesRepName || '管理者'} さん</p><p>新しい依頼が届きました。</p>`
               + `<ul><li><b>依頼者：</b>${companyName} ${clientDisplayName}</li>`
               + `<li><b>種別：</b>${simType}</li></ul>`
               + `<p>■ <b>格納フォルダ:</b><br><a href="${folderUrl}">${folderUrl}</a></p>`
               + `<div style="background:#f1f3f4; padding:10px; border-radius:5px;"><b>✨ AI事前診断:</b><br>${aiText.replace(/\n/g, '<br>')}</div>`;

  if (isMitumori && preUrl) {
    htmlBody += `<hr><p style="color:red; font-weight:bold;">★正式依頼用リンク（依頼者へ送付用）</p>`
              + `<a href="${preUrl}">[正式依頼用フォームを開く]</a>`;
  }

  GmailApp.sendEmail(targetEmail, subject, "", { 
    htmlBody: htmlBody, 
    attachments: [pdfBlob], 
    name: '自動通知システム' 
  });
}