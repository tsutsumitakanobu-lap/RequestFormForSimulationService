/**
 * Main function executed on form submission.
 * Handles dual-path logic (Direct Formal vs Estimate-to-Formal).
 */
function createFolderAndMoveFiles(e) {
  // --- [1. CONFIGURATION & CONSTANTS] ---
  const CONFIG = {
    PARENT_FOLDER_ID: '1ZyCfHWZw35Bpx-a8zwlgBLLfQtQBwcjmGqSc3dEvclnAYrZD8j54hTZ7JNopfunDV_IH576s',
    ADDRESS_BOOK_ID: '1MEtptfvwTvmC6YhRoQIAyU5-2VO3DvFjTtffdUTi35c',
    ADDRESS_BOOK_SHEET_NAME: 'シート1',
    PARAM_MAP_SHEET_NAME: 'フォーム項目一覧',
    LOG_SS_ID: '1i9OesRVOtN5_fY8ntCEp_fHsMh_pefRawlaYZAZMbVo',
    LOG_SHEET_NAME: 'RequestLog',
    MITUMORI_PROMPT_FILE_ID: '1liF9c709fvFYvLGzg2_QHC8UWLpJ6OrW',
    SEISHIKI_PROMPT_FILE_ID: '1Eu8XlG-0IKIft6hDb3WAXm4DMG6V4brj',
    GEMINI_API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'),
    ADMIN_EMAIL: PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL')
  };

  /** * CRITICAL ITEM IDs (GAS Item IDs)
   * Note: Some items exist in both Estimate and Formal sections with different IDs.
   */
  const IDS = {
    FORM_TYPE: '2015761906',
    SALES_REP: '1973732404',
    FOLDER_ID_HIDDEN: '605773322',
    
    // Plant Name (Check both sections)
    PLANT_NAME_M: '1618123851', 
    PLANT_NAME_S: '700137931', 
    
    // Client Name (Check both sections)
    CLIENT_M: '1200762231',
    CLIENT_S: '1940403378', 
    
    // Affiliation (Check both sections)
    COMPANY_M: '715482028',
    COMPANY_S: '356449795'
  };


  try {
    if (!e) return;
    const form = e.source;
    const itemResponses = e.response.getItemResponses();
    
    let answersMap = {};
    let titleAnswersMap = {};
    let filesToMove = [];

    // --- [2. PARSE RESPONSES] ---
    itemResponses.forEach(itemRes => {
      const item = itemRes.getItem();
      const id = item.getId().toString();
      const title = item.getTitle();
      let val = itemRes.getResponse();

      if (item.getType() === FormApp.ItemType.FILE_UPLOAD && val) {
        const fileIds = Array.isArray(val) ? val : [val];
        val = fileIds.map(fid => {
          let file = DriveApp.getFileById(fid);
          let newName = file.getName().replace(/(.+) - [^-]+(\.[a-zA-Z0-9]+)$/, '$1$2');
          file.setName(newName);
          filesToMove.push(fid);
          return newName;
        }).join('\n');
      }
      answersMap[id] = val;
      titleAnswersMap[title] = val;
    });

    const isMitumori = (answersMap[IDS.FORM_TYPE] === "見積もり依頼");

    // --- [3. RESOLVE SHARED FIELDS (COALESCE LOGIC)] ---
    // This logic picks whichever ID has data, supporting both direct and estimate paths.
    const activePlantName = answersMap[IDS.PLANT_NAME_M] || answersMap[IDS.PLANT_NAME_S] || 'UnknownPlant';
    const activeClientName = answersMap[IDS.CLIENT_M] || answersMap[IDS.CLIENT_S] || '';
    const activeCompanyName = answersMap[IDS.COMPANY_M] || answersMap[IDS.COMPANY_S] || '';

    // --- [4. DATA MANAGEMENT (Log all columns)] ---
    logToMasterDatabase(CONFIG, form, answersMap);

    // --- [5. FOLDER RESOLUTION] ---
    const parentFolder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
    let targetFolder;
    const existingFolderId = answersMap[IDS.FOLDER_ID_HIDDEN];

    // Reuse folder if Folder ID is passed (Formal request after estimate)
    if (!isMitumori && existingFolderId) {
      try { targetFolder = DriveApp.getFolderById(existingFolderId); } catch(err) {}
    }
    
    if (!targetFolder) {
      const timeStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd_HHmm');
      targetFolder = parentFolder.createFolder(`${activePlantName}_${timeStr}`);
    }

    // --- [6. GENERATE PRE-FILLED URL (For Estimate Path)] ---
    let prefilledUrl = "";
    if (isMitumori) {
      const baseUrl = form.getPublishedUrl().replace("viewform", "viewform?");
      let queryParts = [`entry.1739104995=${encodeURIComponent("正式依頼")}`];
      queryParts.push(`entry.77031512=${encodeURIComponent(targetFolder.getId())}`);

      const URL_PARAM_MAP = loadUrlParamMap(CONFIG);
      for (let itemId in URL_PARAM_MAP) {
        if (answersMap[itemId]) {
          queryParts.push(`${URL_PARAM_MAP[itemId]}=${encodeURIComponent(answersMap[itemId])}`);
        }
      }
      prefilledUrl = baseUrl + queryParts.join("&");
    }

    // --- [7. PDF, AI & NOTIFICATION] ---
    let pdfBody = "【シミュレーション依頼 回答内容まとめ】\n\n";
    for (let title in titleAnswersMap) {
      if (title === "フォルダID") continue;
      let val = titleAnswersMap[title];
      if (title.includes("氏名") && val && !val.endsWith("様")) val += " 様";
      pdfBody += `■ ${title}\n${val}\n\n`;
    }

    const aiAnalysis = callGeminiAI(pdfBody, filesToMove, isMitumori, CONFIG);
    
    if (isMitumori) {
      pdfBody += `\n----------------\n【正式依頼URL】\n${prefilledUrl}\n----------------\n`;
    }
    pdfBody += `\n【AI事前診断】\n${aiAnalysis}`;

    const pdfBlob = createPDF(pdfBody, activePlantName, targetFolder);
    filesToMove.forEach(fid => { try { DriveApp.getFileById(fid).moveTo(targetFolder); } catch(e) {} });

    // Use unified "Active" variables for notification
    sendNotification(CONFIG, titleAnswersMap, targetFolder.getUrl(), aiAnalysis, pdfBlob, isMitumori, prefilledUrl, activeCompanyName, activeClientName);

  } catch (err) {
    console.error(err.stack);
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

/**
 * Loads URL pre-fill parameter map from Spreadsheet, with Script Properties cache.
 * Cache key: 'URL_PARAM_MAP_CACHE'
 * To invalidate: run clearUrlParamMapCache()
 */
function loadUrlParamMap(config) {
  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperty('URL_PARAM_MAP_CACHE');
  if (cached) return JSON.parse(cached);

  const sheet = SpreadsheetApp.openById(config.ADDRESS_BOOK_ID)
    .getSheetByName(config.PARAM_MAP_SHEET_NAME);
  if (!sheet) throw new Error(`シート "${config.PARAM_MAP_SHEET_NAME}" が見つかりません`);

  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) { // Skip header row
    const entryId      = data[i][3]; // Column D: entry.ID (URL用)
    const srcGasItemId = data[i][4]; // Column E: 事前入力用対応Item ID
    if (srcGasItemId && srcGasItemId !== '—' && entryId && entryId !== '—') {
      map[srcGasItemId.toString()] = entryId;
    }
  }

  props.setProperty('URL_PARAM_MAP_CACHE', JSON.stringify(map));
  return map;
}

/**
 * Clears the URL_PARAM_MAP cache stored in Script Properties.
 * Run this manually after editing the mapping sheet.
 */
function clearUrlParamMapCache() {
  PropertiesService.getScriptProperties().deleteProperty('URL_PARAM_MAP_CACHE');
  console.log('URL_PARAM_MAP_CACHE を削除しました。次回実行時にSpreadsheetから再読み込みします。');
}