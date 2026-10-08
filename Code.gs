/**
 * 補習班講義印刷、對帳與發放管理系統 (PrintHub) - Google Apps Script 後端
 * 帳號：winners.educenter@gmail.com
 */

const CONFIG = {
  // 【試算表綁定方式說明】：
  // 1. 最推薦做法：在 Google 雲端硬碟建立一張 Google 試算表，點選上方「擴充功能」>「Apps Script」貼入程式碼。
  //    此時 SPREADSHEET_ID 請留空 ""，系統會 100% 自動綁定該試算表！
  // 2. 指定試算表：若您在 script.google.com 獨立建專案，可把您 Google 試算表網址中的那串 ID 貼在下方引號中。
  // 3. 全自動模式：若完全留空且非擴充功能建立，系統首次執行會在 Google Drive 自動生成一張「補習班講義印刷與發放數據庫」試算表。
  SPREADSHEET_ID: "", 

  DEFAULT_VENDOR_EMAIL: "888@print.com.tw", // 可在設定頁自訂
  SHEET_NAME_ORDERS: "講義工單紀錄",
  SHEET_NAME_ROSTERS: "班級名單暫存",
  SHEET_NAME_SETTINGS: "系統全域設定"
};

/**
 * Web App 入口
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("補習班講義印刷與發放雲端管理系統")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 確保試算表資料表與欄位初始化 (精準定位試算表)
 */
function getOrCreateDb() {
  let ss = null;

  // 1. 優先檢查是否指定 SPREADSHEET_ID
  if (CONFIG.SPREADSHEET_ID && CONFIG.SPREADSHEET_ID.trim() !== "") {
    try {
      ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID.trim());
    } catch (e) {
      Logger.log("指定 SPREADSHEET_ID 開啟失敗: " + e.message);
    }
  }

  // 2. 若無指定 ID，檢查是否為試算表容器綁定 (擴充功能 > Apps Script)
  if (!ss) {
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch (e) {}
  }

  // 3. 若皆無，則在 Google Drive 自動建立一張專用試算表
  if (!ss) {
    ss = SpreadsheetApp.create("補習班講義印刷與發放數據庫");
  }

  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME_ORDERS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME_ORDERS);
    sheet.appendRow([
      "工單編號", "母講義名稱", "是否加印", "色彩規格", "學用本數", 
      "教用本數", "PDF頁數", "交件日期", "封面規格", "特殊備註", 
      "郵件完整主旨", "Gmail訊息ID", "信件狀態", "到貨狀態", "到貨時間", 
      "已發放本數", "單頁單價", "裝訂費用", "特殊加價", "預估總金額", 
      "是否完成對帳", "建立時間", "附件檔案IDs", "學生名單JSON", "是否封存"
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 25).setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold");
  } else {
    // 自動相容既有資料庫，若未達 25 欄自動補齊表頭
    if (sheet.getLastColumn() < 25) {
      sheet.getRange(1, 25).setValue("是否封存").setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold");
    }
  }
  return { ss, sheet };
}

function formatDateSafe(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "GMT+8", "yyyy-MM-dd HH:mm");
  }
  return String(val);
}

/**
 * 讀取所有講義工單
 */
function getOrders() {
  SpreadsheetApp.flush();
  const { sheet } = getOrCreateDb();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const rows = data.slice(1);
  return rows.map((r, index) => {
    return {
      rowIndex: index + 2,
      orderId: String(r[0] || ""),
      materialName: String(r[1] || "未命名講義"),
      isReprint: r[2] === true || r[2] === "TRUE",
      colorType: String(r[3] || "黑白"),
      studentCopies: Number(r[4]) || 0,
      teacherCopies: Number(r[5]) || 0,
      pdfPages: Number(r[6]) || 0,
      dueDate: String(r[7] || ""),
      coverType: String(r[8] || "公版封面"),
      note: String(r[9] || ""),
      emailSubject: String(r[10] || ""),
      gmailMessageId: String(r[11] || ""),
      mailStatus: String(r[12] || "SENT"), // SENT, TRASHED
      deliveryStatus: String(r[13] || "PENDING"), // PENDING, RECEIVED
      receivedAt: formatDateSafe(r[14]),
      distributedCopies: Number(r[15]) || 0,
      unitPrice: Number(r[16]) || 0.38,
      bindingCost: Number(r[17]) || 20,
      extraCost: Number(r[18]) || 0,
      estimatedTotal: Number(r[19]) || 0,
      reconciled: r[20] === true || r[20] === "TRUE",
      createdAt: formatDateSafe(r[21]),
      attachmentIds: r[22] ? String(r[22]).split(",") : [],
      rosterJson: String(r[23] || "[]"),
      isArchived: r[24] === true || r[24] === "TRUE"
    };
  });
}

/**
 * 掃描 Google Drive 指定資料夾內的所有檔案 (支援搜尋講義與封面)
 */
function scanDriveFolder(folderId) {
  try {
    const folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
    const files = folder.getFiles();
    const result = [];
    while (files.hasNext()) {
      const f = files.next();
      const mime = f.getMimeType();
      const name = f.getName();
      const sizeMb = (f.getSize() / (1024 * 1024)).toFixed(2);
      
      // 支援 PDF, AI, ZIP 等
      const isPdf = name.toLowerCase().endsWith(".pdf") || mime === "application/pdf";
      const isAi = name.toLowerCase().endsWith(".ai") || name.toLowerCase().endsWith(".eps");

      result.push({
        id: f.getId(),
        name: name,
        sizeMb: sizeMb,
        isPdf: isPdf,
        isAi: isAi,
        downloadUrl: f.getDownloadUrl(),
        updatedAt: Utilities.formatDate(f.getLastUpdated(), "GMT+8", "yyyy-MM-dd HH:mm")
      });
    }
    return { success: true, files: result, folderName: folder.getName() };
  } catch (err) {
    return { success: false, message: "讀取 Drive 資料夾失敗：" + err.message };
  }
}

/**
 * 自動讀取 Google Drive 中 PDF 檔案的總頁數
 */
function getPdfPageCount(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const bytes = blob.getBytes();
    
    // 將前/後位元組轉為字串搜尋 /Count 或 /Type /Pages
    // PDF 內部 Pages 物件通常包含 /Count N
    let text = "";
    const len = Math.min(bytes.length, 100000); // 讀取部分資料或全部
    for (let i = 0; i < len; i++) {
      text += String.fromCharCode(bytes[i]);
    }
    
    // 正則匹配 /Count \d+
    const matches = text.match(/\/Count\s+(\d+)/g);
    let maxPages = 0;
    if (matches) {
      matches.forEach(m => {
        const num = parseInt(m.replace(/\/Count\s+/, ""), 10);
        if (num > maxPages) maxPages = num;
      });
    }

    if (maxPages > 0) {
      return { success: true, pages: maxPages };
    }
    return { success: true, pages: 0, note: "未能直接自動讀取頁數，請手動確認" };
  } catch (err) {
    return { success: false, message: "解析 PDF 失敗: " + err.message };
  }
}

/**
 * 正式送印發信：組裝主旨、掛載附件、發送 Gmail、記錄到 Sheet
 */
function sendPrintEmail(form) {
  try {
    const {
      materialName,
      isReprint,
      colorType,
      studentCopies,
      teacherCopies,
      pdfPages,
      dueDateFormatted, // 例如 "08/27(四)"
      coverType, // 例如 "公版封面"
      specialTag, // 例如 "要分隔(藍色)"
      note, // 內文備註
      vendorEmail,
      attachmentFileIds,
      rawAttachments, // 支援本機同步資料夾上傳之 Base64 檔案 [{name, mimeType, base64Data}]
      unitPrice,
      bindingCost,
      extraCost
    } = form;

    // 1. 組裝標準化主旨
    // 範例：[加印][高一下數學講義B2][黑白][學4][07/09(四)][公版封面]
    // 範例：[國三上社會講義][黑白][學21][08/27(四)][公版封面]
    // 範例：[高二物理講義][黑白][學8][08/20(四)][要分隔(藍色)][公版封面]
    let subjectParts = [];
    if (isReprint) subjectParts.push("[加印]");
    subjectParts.push(`[${materialName}]`);
    subjectParts.push(`[${colorType}]`);

    // 本數表示：學X / 教Y 或 學X
    let copyStr = "";
    if (studentCopies > 0 && teacherCopies > 0) {
      copyStr = `學${studentCopies}/教${teacherCopies}`;
    } else if (studentCopies > 0) {
      copyStr = `學${studentCopies}`;
    } else if (teacherCopies > 0) {
      copyStr = `教${teacherCopies}`;
    } else {
      copyStr = `學1`;
    }
    subjectParts.push(`[${copyStr}]`);
    subjectParts.push(`[${dueDateFormatted}]`);
    if (specialTag && specialTag.trim() !== "") {
      subjectParts.push(`[${specialTag.trim()}]`);
    }
    subjectParts.push(`[${coverType || "公版封面"}]`);

    const fullSubject = subjectParts.join("");

    // 2. 獲取附件 (優先支援本機同步資料夾之附件，亦相容 Drive 檔案)
    const attachments = [];
    if (rawAttachments && rawAttachments.length > 0) {
      rawAttachments.forEach(att => {
        try {
          const blob = Utilities.newBlob(Utilities.base64Decode(att.base64Data), att.mimeType || "application/octet-stream", att.name);
          attachments.push(blob);
        } catch (e) {
          Logger.log("解析本機附件失敗: " + att.name + " - " + e.message);
        }
      });
    } else if (attachmentFileIds && attachmentFileIds.length > 0) {
      attachmentFileIds.forEach(id => {
        try {
          const file = DriveApp.getFileById(id);
          attachments.push(file.getAs(file.getMimeType()).setName(file.getName()));
        } catch (e) {
          Logger.log("找不到 Drive 檔案 ID: " + id);
        }
      });
    }

    // 3. 信件內文
    let emailBody = `您好，麻煩請協助印製以下講義：\n\n`;
    emailBody += `【講義名稱】：${materialName}\n`;
    emailBody += `【色彩規格】：${colorType}\n`;
    emailBody += `【預計本數】：${copyStr}\n`;
    emailBody += `【交件日期】：${dueDateFormatted}\n`;
    emailBody += `【內頁頁數】：${pdfPages} 頁\n`;
    if (note && note.trim() !== "") {
      emailBody += `【附註說明】：${note.trim()}\n`;
    }
    emailBody += `\n謝謝您！\n贏家升大教學團隊 敬上`;

    // 4. 發送 Gmail
    const recipient = vendorEmail || CONFIG.DEFAULT_VENDOR_EMAIL;
    const sentMsg = GmailApp.sendEmail(recipient, fullSubject, emailBody, {
      attachments: attachments,
      name: "贏家升大"
    });

    // 取得剛寄出信件的 messageId (查詢最新寄件備份)
    const sentThreads = GmailApp.search(`subject:"${fullSubject}" to:${recipient}`, 0, 1);
    let messageId = "";
    if (sentThreads.length > 0) {
      const msgs = sentThreads[0].getMessages();
      messageId = msgs[msgs.length - 1].getId();
    }

    // 5. 計算預估金額
    const totalCopies = (Number(studentCopies) || 0) + (Number(teacherCopies) || 0);
    const singleBookCost = Math.round((Number(pdfPages) * Number(unitPrice)) + Number(bindingCost));
    const estimatedTotal = (singleBookCost * totalCopies) + Number(extraCost || 0);

    // 6. 寫入 Google Sheet
    const orderId = "ORD-" + Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd-HHmmss");
    const { sheet } = getOrCreateDb();
    sheet.appendRow([
      orderId,
      materialName,
      isReprint ? true : false,
      colorType,
      studentCopies,
      teacherCopies,
      pdfPages,
      dueDateFormatted,
      coverType || "公版封面",
      note || "",
      fullSubject,
      messageId,
      "SENT", // 信件狀態
      "PENDING", // 到貨狀態
      "", // 到貨時間
      0, // 已發放本數
      unitPrice,
      bindingCost,
      extraCost || 0,
      estimatedTotal,
      false, // reconciled
      Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss"),
      (attachmentFileIds || []).join(","),
      "[]"
    ]);

    return {
      success: true,
      orderId: orderId,
      subject: fullSubject,
      estimatedTotal: estimatedTotal,
      message: "送印信件已成功寄出並建立追蹤工單！"
    };
  } catch (err) {
    return { success: false, message: "寄送失敗：" + err.message };
  }
}

/**
 * 更新講義收書到貨狀態 (入庫)
 */
function markOrderReceived(orderId) {
  const { sheet } = getOrCreateDb();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) {
      // 欄位 14 (N欄) 是到貨狀態, 15 (O欄) 是到貨時間
      sheet.getRange(i + 1, 14).setValue("RECEIVED");
      sheet.getRange(i + 1, 15).setValue(Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss"));
      return { success: true, message: "已確認收書入庫！" };
    }
  }
  return { success: false, message: "找不到該筆工單" };
}

/**
 * 空間清理助手：獲取「已收到但 Gmail 尚未刪除」的信件清單與檔案大小
 */
function getCleanupCandidates() {
  const { sheet } = getOrCreateDb();
  const data = sheet.getDataRange().getValues();
  const list = [];

  for (let i = 1; i < data.length; i++) {
    const orderId = data[i][0];
    const materialName = data[i][1];
    const subject = data[i][10];
    const messageId = data[i][11];
    const mailStatus = data[i][12];
    const deliveryStatus = data[i][13];
    const sentTime = data[i][21];

    if (deliveryStatus === "RECEIVED" && mailStatus === "SENT" && messageId) {
      let attachmentSizeMb = 0;
      let attachmentCount = 0;
      let emailExists = true;

      try {
        const msg = GmailApp.getMessageById(messageId);
        if (msg) {
          const attachments = msg.getAttachments();
          attachmentCount = attachments.length;
          let totalBytes = 0;
          attachments.forEach(att => totalBytes += att.getSize());
          attachmentSizeMb = (totalBytes / (1024 * 1024)).toFixed(1);
        } else {
          emailExists = false;
        }
      } catch (e) {
        emailExists = false;
      }

      if (emailExists) {
        list.push({
          orderId: orderId,
          materialName: materialName,
          subject: subject,
          messageId: messageId,
          sentTime: sentTime,
          attachmentCount: attachmentCount,
          attachmentSizeMb: attachmentSizeMb
        });
      }
    }
  }
  return { success: true, candidates: list };
}

/**
 * 空間清理：將指定 Gmail 移至垃圾桶 (安全省空間，保留 Sheets 紀錄)
 */
function trashPrintEmail(orderId, messageId) {
  try {
    let freedMb = 0;
    if (messageId) {
      const msg = GmailApp.getMessageById(messageId);
      if (msg) {
        let totalBytes = 0;
        msg.getAttachments().forEach(att => totalBytes += att.getSize());
        freedMb = (totalBytes / (1024 * 1024)).toFixed(1);
        
        // 移至垃圾桶
        msg.moveToTrash();
      }
    }

    // 更新 Sheet 狀態為 TRASHED
    const { sheet } = getOrCreateDb();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === orderId) {
        sheet.getRange(i + 1, 13).setValue("TRASHED");
        break;
      }
    }

    return {
      success: true,
      orderId: orderId,
      freedMb: freedMb,
      message: `已成功將郵件移入垃圾桶，為 Gmail 釋放了約 ${freedMb} MB 空間！`
    };
  } catch (err) {
    return { success: false, message: "清理失敗：" + err.message };
  }
}

/**
 * 更新學生名單與發放數量
 */
function updateDistributionRoster(orderId, distributedCopies, rosterList) {
  const { sheet } = getOrCreateDb();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) {
      sheet.getRange(i + 1, 16).setValue(distributedCopies); // 已發放本數
      sheet.getRange(i + 1, 24).setValue(JSON.stringify(rosterList)); // 名單 JSON
      return { success: true, message: "發放紀錄更新成功！" };
    }
  }
  return { success: false, message: "找不到該工單" };
}

/**
 * 標記工單已完成對帳
 */
function markOrdersReconciled(orderIds) {
  const { sheet } = getOrCreateDb();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (orderIds.includes(data[i][0])) {
      sheet.getRange(i + 1, 21).setValue(true); // reconciled = true
    }
  }
  return { success: true, message: "已更新對帳完成標記！" };
}

/**
 * 深度掃描 Gmail 寄件備份中的歷史講義印刷信件，自動辨識標籤並批次記錄至資料庫
 * @param {Object|string} options 查詢條件或設定 { query, forceReload }
 */
function scanAndImportGmailHistory(options) {
  try {
    let forceReload = false;
    let query = "";
    if (typeof options === "boolean") {
      forceReload = options;
    } else if (typeof options === "string") {
      query = options;
    } else if (typeof options === "object" && options !== null) {
      forceReload = options.forceReload === true;
      query = options.query || "";
    }
    const { sheet } = getOrCreateDb();

    // 若要求強制重新掃描，清空標題行以下的舊資料
    if (forceReload === true) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, 24).clearContent();
        SpreadsheetApp.flush();
      }
    }

    const existingData = sheet.getDataRange().getValues();
    const existingRowMap = new Map(); // key -> { rowIndex, pages }

    for (let i = 1; i < existingData.length; i++) {
      const msgId = String(existingData[i][11] || '').trim();
      const subj = String(existingData[i][10] || '').trim();
      const pages = Number(existingData[i][6]) || 0;
      const info = { rowIndex: i + 1, pages: pages };
      if (msgId) existingRowMap.set(msgId, info);
      if (subj) existingRowMap.set(subj, info);
    }

    // 搜尋語法：優先依據講義標籤或影印社信件搜尋
    let searchQuery = query || 'label:"888講義DM印製" OR label:"888講義DM印製/泗商影印社" OR subject:"[公版封面]" OR subject:"[黑白]" OR subject:"[彩色]"';
    
    // 獲取信件執行緒 (最多抓取前 150 筆)
    let threads = [];
    try {
      threads = GmailApp.search(searchQuery, 0, 150);
    } catch (e) {
      // 若標籤不存在則寬鬆搜尋
      threads = GmailApp.search('subject:"[黑白]" OR subject:"[彩色]" OR subject:"[加印]"', 0, 150);
    }

    const discoveredLabels = new Set();
    const scanLogs = []; // 保存每一封信件的完整日誌
    let importedCount = 0;
    let skippedCount = 0;
    let scanIndex = 0;

    threads.forEach(thread => {
      // 取得該討論串所有標籤
      const labels = thread.getLabels().map(l => l.getName());
      labels.forEach(lbl => discoveredLabels.add(lbl));

      // 檢查標籤中是否含有「已收到」
      const isReceivedByLabel = labels.some(l => l.includes("已收到") || l.includes("講義已收到"));

      const msgs = thread.getMessages();
      msgs.forEach(msg => {
        scanIndex++;
        const msgId = msg.getId();
        const subject = msg.getSubject().trim();
        const sentDate = Utilities.formatDate(msg.getDate(), "GMT+8", "yyyy-MM-dd HH:mm");

        // 取得附件清單與大小
        const attachments = msg.getAttachments();
        const attDetails = attachments.map(a => {
          const mb = (a.getSize() / (1024 * 1024)).toFixed(1);
          return `${a.getName()} (${mb} MB)`;
        });

        // 解析主旨格式
        const parsed = parsePrintSubject(subject);

        const logEntry = {
          index: scanIndex,
          date: sentDate,
          subject: subject,
          materialName: parsed.isValid ? parsed.materialName : "非規範主旨",
          isReprint: parsed.isValid ? parsed.isReprint : false,
          copies: parsed.isValid ? (parsed.teacherCopies > 0 ? `學${parsed.studentCopies}/教${parsed.teacherCopies}` : `學${parsed.studentCopies}`) : "-",
          colorType: parsed.isValid ? parsed.colorType : "-",
          dueDate: parsed.isValid ? parsed.dueDate : "-",
          labels: labels,
          deliveryStatus: isReceivedByLabel ? "已收到" : "待收書",
          attachments: attDetails,
          isValid: parsed.isValid,
          importStatus: ""
        };

        // 檢查是否已存在於試算表中
        const existingInfo = existingRowMap.get(msgId) || existingRowMap.get(subject);
        if (existingInfo) {
          // 若已存在但試算表內頁數為 0，且本次能解析出頁數，立即自動就地補齊！
          if (existingInfo.pages === 0) {
            let detectedPages = extractPdfPagesFromMessage(msg, attachments, subject);
            if (detectedPages === 0 && parsed.pages) detectedPages = parsed.pages;

            if (detectedPages > 0) {
              const copies = (parsed.studentCopies || 1) + (parsed.teacherCopies || 0);
              const unitP = parsed.colorType === '彩色' ? 0.9 : 0.38;
              const estTotal = (Math.round((detectedPages * unitP) + 20) * copies);

              // G 欄為欄位 7 (PDF 頁數), T 欄為欄位 20 (預估總金額)
              sheet.getRange(existingInfo.rowIndex, 7).setValue(detectedPages);
              sheet.getRange(existingInfo.rowIndex, 20).setValue(estTotal);
              existingInfo.pages = detectedPages;

              logEntry.detectedPages = detectedPages;
              logEntry.importStatus = `🔄 已自動補齊頁數 (${detectedPages} 頁)`;
              scanLogs.push(logEntry);
              return;
            }
          }

          skippedCount++;
          logEntry.importStatus = "已存在資料庫 (已跳過)";
          scanLogs.push(logEntry);
          return;
        }

        if (!parsed.isValid) {
          logEntry.importStatus = "非印刷格式 (略過)";
          scanLogs.push(logEntry);
          return;
        }

        // 深度解析 PDF 附件二進位結構與內文，精準取得真實講義頁數
        let detectedPages = extractPdfPagesFromMessage(msg, attachments, subject);
        if (detectedPages === 0 && parsed.pages) {
          detectedPages = parsed.pages;
        }
        logEntry.detectedPages = detectedPages;

        const orderId = "HIST-" + Utilities.formatDate(msg.getDate(), "GMT+8", "yyyyMMdd-HHmmss");
        const copies = parsed.studentCopies + parsed.teacherCopies;
        const unitP = parsed.colorType === '彩色' ? 0.9 : 0.38;
        const estTotal = (Math.round((detectedPages * unitP) + 20) * copies);

        const row = [
          orderId,
          parsed.materialName,
          parsed.isReprint,
          parsed.colorType,
          parsed.studentCopies,
          parsed.teacherCopies,
          detectedPages,
          parsed.dueDate,
          parsed.coverType,
          parsed.note || "",
          subject,
          msgId,
          "SENT",
          isReceivedByLabel ? "RECEIVED" : "PENDING",
          isReceivedByLabel ? sentDate : "",
          0, // 已發放本數
          unitP,
          20, // 裝訂費
          0,
          estTotal,
          false,
          sentDate,
          attDetails.join(", "),
          "[]"
        ];

        sheet.appendRow(row);
        existingRowMap.set(msgId, { rowIndex: sheet.getLastRow(), pages: detectedPages });
        existingRowMap.set(subject, { rowIndex: sheet.getLastRow(), pages: detectedPages });
        importedCount++;

        logEntry.importStatus = "✅ 成功匯入資料庫";
        scanLogs.push(logEntry);
      });
    });

    // 強制將寫入的資料同步落盤，確保即時讀取最新狀態
    SpreadsheetApp.flush();

    // 重新載入最新所有工單
    const allLatestOrders = getOrders();

    return {
      success: true,
      importedCount: importedCount,
      skippedCount: skippedCount,
      totalScanned: threads.length,
      discoveredLabels: Array.from(discoveredLabels),
      scanLogs: scanLogs, // 包含每封信的序號、標題、附件與標籤詳細日誌
      allOrders: allLatestOrders, // 包含所有最新工單
      message: `掃描完成！共分析 ${threads.length} 串郵件，成功匯入 ${importedCount} 筆，已存在 ${skippedCount} 筆。`
    };
  } catch (err) {
    return { success: false, message: "掃描失敗: " + err.message };
  }
}

/**
 * 智慧容錯主旨解析器：解析正印、加印、教/學本數、色彩規格、日期與備註
 */
function parsePrintSubject(rawSubject) {
  if (!rawSubject || typeof rawSubject !== 'string') {
    return { isValid: false };
  }

  // 檢查是否為印刷格式 (至少包含中括號且包含黑白/彩色/講義/印製/加印等特徵)
  const isPrintRelated = rawSubject.includes("[") && (
    rawSubject.includes("黑白") || 
    rawSubject.includes("彩色") || 
    rawSubject.includes("加印") || 
    rawSubject.includes("講義") || 
    rawSubject.includes("學") || 
    rawSubject.includes("教")
  );

  if (!isPrintRelated) {
    return { isValid: false };
  }

  // 提取所有中括號內容 [xxx]
  const brackets = [];
  const regex = /\[(.*?)\]/g;
  let match;
  while ((match = regex.exec(rawSubject)) !== null) {
    brackets.push(match[1].trim());
  }

  // 檢查括號之後是否有破折號附註 (例如: - 115上國一自然_【生物B1學期講義】共126頁...)
  let extraNote = "";
  const dashIndex = rawSubject.indexOf("]");
  const lastBracketEnd = rawSubject.lastIndexOf("]");
  if (lastBracketEnd >= 0 && lastBracketEnd < rawSubject.length - 1) {
    extraNote = rawSubject.substring(lastBracketEnd + 1).replace(/^[\s\-—]+/, "").trim();
  }

  let isReprint = false;
  let colorType = "黑白";
  let studentCopies = 1;
  let teacherCopies = 0;
  let dueDate = "";
  let coverType = "公版封面";
  let specialTag = "";
  let materialName = "";
  let pages = 0;

  // 逐一分析括號 token
  brackets.forEach(token => {
    if (token === "加印") {
      isReprint = true;
    } else if (token === "黑白" || token === "彩色" || token === "黑白/彩色" || token === "彩色/黑白") {
      colorType = token;
    } else if (/^學\d+/.test(token) || /^教\d+/.test(token)) {
      // 本數解析：如 學21, 學26/教1, 學19/教1, 教2
      const sMatch = token.match(/學(\d+)/);
      const tMatch = token.match(/教(\d+)/);
      if (sMatch) studentCopies = parseInt(sMatch[1], 10);
      if (tMatch) teacherCopies = parseInt(tMatch[1], 10);
    } else if (/\d{1,2}\/\d{1,2}/.test(token)) {
      // 日期解析：如 08/27(四), 07/09(四), 09/2(三)
      dueDate = token;
    } else if (token.includes("封面")) {
      coverType = token;
    } else if (token.includes("分隔") || token.includes("急件")) {
      specialTag = token;
    } else {
      // 剩餘最長或具備課程/年級特徵的視為講義名稱
      if (!materialName || token.includes("講義") || token.length > materialName.length) {
        materialName = token;
      }
    }
  });

  if (!materialName && brackets.length > 0) {
    materialName = brackets[0] === "加印" && brackets.length > 1 ? brackets[1] : brackets[0];
  }

  // 嘗試從 extraNote 提取頁數
  if (extraNote) {
    const pageM = extraNote.match(/共\s*(\d+)\s*頁/);
    if (pageM) pages = parseInt(pageM[1], 10);
  }

  return {
    isValid: true,
    materialName: materialName || "未命名講義",
    isReprint: isReprint,
    colorType: colorType,
    studentCopies: studentCopies,
    teacherCopies: teacherCopies,
    dueDate: dueDate,
    coverType: coverType,
    specialTag: specialTag,
    note: extraNote,
    pages: pages
  };
}

/**
 * 從郵件物件及其附件中深度解析真實講義頁數 (支援 PDF 二進位結構秒讀、檔名與內文提取)
 * @param {GmailMessage} msg Gmail 郵件物件
 * @param {Array<GmailAttachment>} attachments 附件陣列
 * @param {string} subject 信件主旨
 * @return {number} 解析出的頁數
 */
function extractPdfPagesFromMessage(msg, attachments, subject) {
  let detectedPages = 0;

  // 1. 優先從 PDF 附件的二進位結構直接秒讀 /Count 或 /N
  if (attachments && attachments.length > 0) {
    const pdfAttachments = attachments.filter(a => {
      const name = a.getName().toLowerCase();
      return name.endsWith('.pdf') || a.getContentType() === 'application/pdf';
    });

    // 若有多個 PDF，且有包含「封面」字樣，優先區分內頁講義與封面
    const contentPdfs = pdfAttachments.filter(a => !a.getName().includes('封面'));
    const targetPdfs = contentPdfs.length > 0 ? contentPdfs : pdfAttachments;

    let totalPdfPages = 0;
    targetPdfs.forEach(att => {
      const pages = getPagesFromSinglePdfAttachment(att);
      if (pages > 0) {
        totalPdfPages += pages;
      }
    });

    if (totalPdfPages > 0) {
      return totalPdfPages;
    }
  }

  // 2. 次選：從附件檔名匹配頁數 (如 xxx_126p.pdf, xxx_126頁.pdf)
  if (attachments && attachments.length > 0) {
    for (const att of attachments) {
      const name = att.getName();
      const fnMatch = name.match(/[\(\[_\-\s](\d{1,4})\s*(?:p|頁|pages)[\)\]_\-\s\.]/i) || name.match(/共\s*(\d{1,4})\s*頁/);
      if (fnMatch) {
        const p = parseInt(fnMatch[1], 10);
        if (p > 0 && p < 3000) return p;
      }
    }
  }

  // 3. 第三層備援：從信件主旨與內文匹配
  try {
    const bodyText = msg ? msg.getPlainBody() : "";
    const combinedText = (subject || "") + "\n" + bodyText;
    
    // 匹配 "共 126 頁", "126 頁", "內頁: 126", "頁數: 126", "126p"
    const patterns = [
      /共\s*(\d{1,4})\s*頁/,
      /(?:內頁|頁數|總頁數)\s*[:：]?\s*(\d{1,4})/,
      /\[\s*(\d{1,4})\s*頁\s*\]/,
      /【\s*(\d{1,4})\s*頁\s*】/,
      /(?:P|p)\s*[:：]?\s*(\d{1,4})/
    ];

    for (const pat of patterns) {
      const m = combinedText.match(pat);
      if (m) {
        const p = parseInt(m[1], 10);
        if (p > 0 && p < 3000) return p;
      }
    }
  } catch (e) {}

  return 0;
}

/**
 * 從單一 PDF 附件解析二進位位元組取得頁數 (高速頭尾抽取 + 解除長度限制正則)
 */
function getPagesFromSinglePdfAttachment(att) {
  try {
    const bytes = att.getBytes();
    const len = bytes.length;
    if (len === 0) return 0;

    // 1. 高速讀取頭部 64KB (匹配 Linearized Web 快顯頁數)
    const headLimit = Math.min(len, 65536);
    let headChars = [];
    for (let i = 0; i < headLimit; i++) {
      const b = bytes[i];
      if ((b >= 32 && b <= 126) || b === 10 || b === 13) {
        headChars.push(String.fromCharCode(b));
      } else {
        headChars.push(" ");
      }
    }
    const headStr = headChars.join("");

    // 2. 高速讀取尾部 384KB (PDF 核心型錄 Catalog、Pages 節點與 /Count 必定在檔案尾端)
    const tailStart = Math.max(0, len - 393216);
    let tailChars = [];
    for (let i = tailStart; i < len; i++) {
      const b = bytes[i];
      if ((b >= 32 && b <= 126) || b === 10 || b === 13) {
        tailChars.push(String.fromCharCode(b));
      } else {
        tailChars.push(" ");
      }
    }
    const tailStr = tailChars.join("");

    // 3. 組合頭尾字串進行大跨度型錄解析
    const combined = headStr + "\n" + tailStr;
    const pages = extractPagesFromPdfText(combined);
    if (pages > 0) return pages;

    // 4. 備援：若檔案小於 3MB，直接全檔搜尋 /Type /Page 獨立頁面物件
    if (len <= 3145728) {
      let allChars = [];
      for (let i = 0; i < len; i++) {
        const b = bytes[i];
        if ((b >= 32 && b <= 126) || b === 10 || b === 13) {
          allChars.push(String.fromCharCode(b));
        } else {
          allChars.push(" ");
        }
      }
      const allStr = allChars.join("");
      const p = extractPagesFromPdfText(allStr);
      if (p > 0) return p;

      // 統計 /Type /Page (兼容 /Type/Page 與 /Type /Page)
      const pageMatches = allStr.match(/\/Type\s*\/Page[^s\w]/g);
      if (pageMatches && pageMatches.length > 0 && pageMatches.length < 3000) {
        return pageMatches.length;
      }
    }

    return 0;
  } catch (e) {
    Logger.log("讀取單一 PDF 頁數失敗: " + att.getName() + " - " + e.message);
    return 0;
  }
}

/**
 * 大跨度正規表示式匹配 PDF 中的 /Count 或 /N (徹底支援破百頁大講義之長 Kids 陣列)
 */
function extractPagesFromPdfText(text) {
  if (!text) return 0;
  let maxPages = 0;

  // 1. 匹配 /Linearized ... /N <頁數>
  const linMatches = text.match(/\/Linearized[\s\S]{0,300}?\/N\s+(\d+)/);
  if (linMatches) {
    const val = parseInt(linMatches[1], 10);
    if (val > 0 && val < 3000) return val;
  }

  // 2. 匹配 /Type /Pages ... /Count N (跨度放寬至 5000 字元，徹底容納長 Kids 陣列)
  const typePagesMatches = text.match(/\/Type\s*\/Pages[\s\S]{0,5000}?\/Count\s+(\d+)/g) || 
                           text.match(/\/Count\s+(\d+)[\s\S]{0,5000}?\/Type\s*\/Pages/g);
  if (typePagesMatches) {
    typePagesMatches.forEach(m => {
      const numMatch = m.match(/\/Count\s+(\d+)/);
      if (numMatch) {
        const val = parseInt(numMatch[1], 10);
        if (val > maxPages && val < 4000) maxPages = val;
      }
    });
  }
  if (maxPages > 0) return maxPages;

  // 3. 匹配所有獨立的 /Count N (如 /Count 126)
  const countMatches = text.match(/\/Count\s+(\d+)/g);
  if (countMatches) {
    countMatches.forEach(m => {
      const val = parseInt(m.replace(/\/Count\s+/, ''), 10);
      if (val > maxPages && val < 4000) maxPages = val;
    });
  }

  return maxPages;
}

/**
 * 一鍵修復補齊現有試算表中頁數為 0 的工單 (依據 Gmail messageId 或主旨讀取原信件附件)
 */
function repairZeroPageOrders() {
  try {
    const { sheet } = getOrCreateDb();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: false, message: "目前尚無任何工單資料" };

    let updatedCount = 0;
    Logger.log("=== 開始執行 0 頁工單自動修復作業 ===");

    for (let i = 1; i < data.length; i++) {
      const currentPages = Number(data[i][6]) || 0;
      const messageId = String(data[i][11] || '').trim();
      const subject = String(data[i][10] || '').trim();
      const materialName = data[i][1] || '講義';
      const studentCopies = Number(data[i][4]) || 0;
      const teacherCopies = Number(data[i][5]) || 0;
      const colorType = data[i][3] || "黑白";
      const unitP = colorType === "彩色" ? 0.9 : 0.38;
      const bindingCost = Number(data[i][17]) || 20;

      // 若頁數為 0
      if (currentPages === 0) {
        let msg = null;

        // 1. 優先依據 messageId 讀取信件
        if (messageId) {
          try {
            msg = GmailApp.getMessageById(messageId);
          } catch (e) {}
        }

        // 2. 備援：若找不到則依主旨精確搜尋
        if (!msg && subject) {
          try {
            const threads = GmailApp.search(`subject:"${subject}"`, 0, 1);
            if (threads.length > 0) {
              const msgs = threads[0].getMessages();
              msg = msgs[msgs.length - 1];
            }
          } catch (e) {}
        }

        if (msg) {
          try {
            const attachments = msg.getAttachments();
            const realPages = extractPdfPagesFromMessage(msg, attachments, subject);
            if (realPages > 0) {
              const copies = (studentCopies || 1) + (teacherCopies || 0);
              const estTotal = (Math.round((realPages * unitP) + bindingCost) * copies);

              // G 欄為欄位 7 (PDF 頁數)
              sheet.getRange(i + 1, 7).setValue(realPages);
              // T 欄為欄位 20 (預估總金額)
              sheet.getRange(i + 1, 20).setValue(estTotal);
              updatedCount++;
              Logger.log(`[修復成功] 第 ${i + 1} 列 【${materialName}】: 解析成功 ${realPages} 頁，更新金額 $${estTotal}`);
            } else {
              Logger.log(`[略過] 第 ${i + 1} 列 【${materialName}】: 附件中未找到可解析之 PDF 頁數`);
            }
          } catch (e) {
            Logger.log(`[異常] 第 ${i + 1} 列 【${materialName}】: ${e.message}`);
          }
        } else {
          Logger.log(`[未找到郵件] 第 ${i + 1} 列 【${materialName}】: ID ${messageId}`);
        }
      }
    }

    SpreadsheetApp.flush();
    const allLatestOrders = getOrders();

    const msg = `修復完成！已成功為 ${updatedCount} 筆工單讀取並補齊真實 PDF 頁數與金額！`;
    Logger.log(msg);
    return {
      success: true,
      updatedCount: updatedCount,
      allOrders: allLatestOrders,
      message: msg
    };
  } catch (err) {
    Logger.log("修復失敗: " + err.message);
    return { success: false, message: "修復失敗：" + err.message };
  }
}

/**
 * 直接更新指定講義的所有工單頁數與預估金額 (支援前端畫面上即時編輯落盤)
 */
function updateMaterialPages(materialName, newPages) {
  try {
    const { sheet } = getOrCreateDb();
    const data = sheet.getDataRange().getValues();
    const p = Number(newPages) || 0;
    let updatedRows = 0;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(materialName).trim()) {
        const studentCopies = Number(data[i][4]) || 0;
        const teacherCopies = Number(data[i][5]) || 0;
        const colorType = data[i][3] || "黑白";
        const unitP = colorType === "彩色" ? 0.9 : 0.38;
        const bindingCost = Number(data[i][17]) || 20;
        const copies = (studentCopies || 1) + (teacherCopies || 0);
        const estTotal = (Math.round((p * unitP) + bindingCost) * copies);

        // G 欄為欄位 7 (PDF 頁數), T 欄為欄位 20 (預估總金額)
        sheet.getRange(i + 1, 7).setValue(p);
        sheet.getRange(i + 1, 20).setValue(estTotal);
        updatedRows++;
      }
    }

    SpreadsheetApp.flush();
    const allLatestOrders = getOrders();
    return {
      success: true,
      allOrders: allLatestOrders,
      message: `已將【${materialName}】的頁數更新為 ${p} 頁，共同步 ${updatedRows} 筆工單！`
    };
  } catch (err) {
    return { success: false, message: "更新頁數失敗：" + err.message };
  }
}

/**
 * 切換講義的封存狀態 (支援封存與解除封存，同步落盤寫入試算表第 25 欄)
 */
function toggleArchiveMaterial(materialName, isArchived) {
  try {
    const { sheet } = getOrCreateDb();
    const data = sheet.getDataRange().getValues();
    const archiveVal = (isArchived === true);
    let updatedRows = 0;

    // 確保第 25 欄表頭存在
    if (sheet.getLastColumn() < 25) {
      sheet.getRange(1, 25).setValue("是否封存").setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold");
    }

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(materialName).trim()) {
        sheet.getRange(i + 1, 25).setValue(archiveVal); // Y 欄為欄位 25 (是否封存)
        updatedRows++;
      }
    }

    SpreadsheetApp.flush();
    const allLatestOrders = getOrders();
    return {
      success: true,
      isArchived: archiveVal,
      allOrders: allLatestOrders,
      message: archiveVal ? `已將【${materialName}】封存！` : `已將【${materialName}】解除封存並移回進行中！`
    };
  } catch (err) {
    return { success: false, message: "更新封存狀態失敗：" + err.message };
  }
}

/**
 * 儲存指定母講義的學生發放名單與扣減庫存 (落盤寫入試算表第 16 欄已發放數與第 24 欄名單JSON)
 */
function saveMaterialDistributionRoster(materialName, rosterList) {
  try {
    const { sheet } = getOrCreateDb();
    const data = sheet.getDataRange().getValues();
    const list = Array.isArray(rosterList) ? rosterList : [];
    const distributedCopies = list.filter(s => s && s.received === true).length;
    const rosterJson = JSON.stringify(list);
    let updatedRows = 0;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(materialName).trim()) {
        sheet.getRange(i + 1, 16).setValue(distributedCopies); // P 欄：已發放本數
        sheet.getRange(i + 1, 24).setValue(rosterJson); // X 欄：學生名單 JSON
        updatedRows++;
      }
    }

    SpreadsheetApp.flush();
    const allLatestOrders = getOrders();
    return {
      success: true,
      distributedCopies: distributedCopies,
      allOrders: allLatestOrders,
      message: `已成功保存【${materialName}】學生發放紀錄！共 ${distributedCopies} 位學生領取，庫存已即時扣減！`
    };
  } catch (err) {
    return { success: false, message: "儲存學生名單失敗：" + err.message };
  }
}

