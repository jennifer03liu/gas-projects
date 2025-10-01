/**
 * @fileoverview 每月壽星生日禮金發放名單產生器
 * 1. 每月月底自動執行，抓取次月生日的員工，並發送審核 Email。
 * 2. 待相關人員點擊 Email 中的「確認」按鈕後，系統會重新從 Sheet 抓取最新名單。
 * 3. 最終根據最新名單，產生兩份 Google Doc 名單 (集邦、拓墣) 並歸檔。
 */

// ===============================================================
// 主要觸發函式 (預計每月月底執行)
// ===============================================================
function createMonthlyBirthdayReport() {
  try {
    Logger.log('開始執行每月壽星報告流程...');
    // 1. 取得壽星名單，用於 Email 預覽
    const { trendforce, topology } = getEligibleBirthdayEmployees();

    if (trendforce.length === 0 && topology.length === 0) {
      Logger.log('下個月沒有符合資格的壽星，流程終止。');
      return;
    }

    // 2. 建立一個有時效性的批准ID (不含員工資料)
    const cacheId = Utilities.getUuid();
    CacheService.getScriptCache().put(cacheId, 'pending', 3600); // 存活1小時
    Logger.log(`產生批准ID: ${cacheId}`);

    // 3. 發送審核郵件
    sendApprovalEmail(trendforce, topology, cacheId);

  } catch (e) {
    Logger.log(`!!!!!!!!!! 建立壽星報告過程中發生嚴重錯誤 !!!!!!!!!!\n${e.stack}`);
    GmailApp.sendEmail(getSetting('PAYMENT_NOTICE_RECIPIENT'), '【錯誤】每月壽星報告產生失敗', `執行過程中發生錯誤：\n${e.stack}`);
  }
}

// ===============================================================
// Web App 介面 (由 Code.js 的 doGet 呼叫)
// ===============================================================
function handleBirthdayApproval(e) {
  try {
    const params = e.parameter;
    const action = params.action;
    const cacheId = params.cacheId;
    let message = '';

    Logger.log(`handleBirthdayApproval running: action=${action}, cacheId=${cacheId}`);

    if (action === 'approve') {
      if (!cacheId) {
        return ContentService.createTextOutput('無效的請求：缺少 cacheId。');
      }
      
      const token = CacheService.getScriptCache().get(cacheId);
      if (token !== 'pending') {
        return ContentService.createTextOutput('操作失敗或連結已過期/已被使用。請重新執行每月壽星報告流程。');
      }
      
      CacheService.getScriptCache().remove(cacheId);
      Logger.log(`Cache ID ${cacheId} 已驗證並移除。`);

      // --- 核心邏輯：重新從Sheet抓取最新資料並產生報告 ---
      message = generateFinalReports();
      
    } else if (action === 'reject') {
      if (cacheId) {
        CacheService.getScriptCache().remove(cacheId);
      }
      message = '操作已記錄。系統將不會產生文件。如有需要，請聯繫IT人員或重新執行流程。';
    } else {
      message = '無效的操作。';
    }
    return ContentService.createTextOutput(message);

  } catch (err) {
    Logger.log(`!!!!!!!!!! handleBirthdayApproval 處理過程中發生嚴重錯誤 !!!!!!!!!!\n${err.stack}`);
    return ContentService.createTextOutput(`處理您的請求時發生錯誤: ${err.message}。請聯繫IT人員。`);
  }
}


/**
 * 執行最終的報告產生流程
 * (此函式由 handleBirthdayApproval 觸發)
 */
function generateFinalReports() {
    Logger.log('點擊批准後，重新從 Google Sheet 抓取最新的壽星名單...');
    const { trendforce, topology } = getEligibleBirthdayEmployees();

    if (trendforce.length === 0 && topology.length === 0) {
      const msg = '操作完成，但根據最新資料，下個月已無符合資格的壽星。';
      Logger.log(msg);
      return msg;
    }

    let successMessages = '壽星名單產生成功！\n\n';
    Logger.log(`找到 ${trendforce.length} 位集邦壽星, ${topology.length} 位拓墣壽星。`);

    // 產生並歸檔「集邦」壽星名單
    if (trendforce.length > 0) {
      const trendforceDoc = generateBirthdayDoc(trendforce, '集邦');
      const trendforceFolderId = getSetting('TRENDFORCE_BIRTHDAY_FOLDER_ID');
      moveFileToFolder(trendforceDoc.getId(), trendforceFolderId);
      const successMsg = `檔案 "${trendforceDoc.getName()}" 已產生並歸檔至集邦資料夾。`;
      Logger.log(successMsg);
      successMessages += `✓ ${successMsg}\n(網址: ${trendforceDoc.getUrl()})\n`;
    }

    // 產生並歸檔「拓墣」壽星名單
    if (topology.length > 0) {
      const topologyDoc = generateBirthdayDoc(topology, '拓墣');
      const topologyFolderId = getSetting('TOPOLOGY_BIRTHDAY_FOLDER_ID');
      moveFileToFolder(topologyDoc.getId(), topologyFolderId);
      const successMsg = `檔案 "${topologyDoc.getName()}" 已產生並歸檔至拓墣資料夾。`;
      Logger.log(successMsg);
      successMessages += `✓ ${successMsg}\n(網址: ${topologyDoc.getUrl()})\n`;
    }

    Logger.log('---------- 報告產生完畢 ----------');
    return successMessages;
}


// ===============================================================
// 1. 取得並篩選符合資格的壽星 (*** 加入詳細偵錯紀錄 ***)
// ===============================================================
function getEligibleBirthdayEmployees() {
  const COLUMN_NAMES = getSetting('BIRTHDAY_COLUMN_NAMES');
  if (!COLUMN_NAMES) {
    throw new Error("在設定中找不到 'BIRTHDAY_COLUMN_NAMES'。");
  }

  const sheet = SpreadsheetApp.openById(getSetting('SPREADSHEET_ID')).getSheetByName(getSetting('EMPLOYEE_SHEET_NAME'));
  if (!sheet) {
    throw new Error(`找不到名為 "${getSetting('EMPLOYEE_SHEET_NAME')}" 的工作表`);
  }

  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // 移除並取得標頭列
  Logger.log(`在 "${getSetting('EMPLOYEE_SHEET_NAME')}" 工作表中找到 ${data.length} 筆員工資料。開始進行篩選...`);


  // 建立欄位名稱與其索引位置的對應
  const indices = {};
  const missingColumns = [];
  for (const key in COLUMN_NAMES) {
    const index = headers.indexOf(COLUMN_NAMES[key]);
    if (index === -1) {
      missingColumns.push(COLUMN_NAMES[key]);
    } else {
      indices[key] = index;
    }
  }

  if (missingColumns.length > 0) {
    throw new Error(`在員工總控制表中找不到以下欄位: "${missingColumns.join(', ')}"。請檢查 CONFIG 設定是否與 Sheet 標頭完全相符。`);
  }

  // --- 篩選邏輯 ---
  const today = new Date();
  const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1).getMonth();
  Logger.log(`目標月份: ${nextMonth + 1}月 (JavaScript 月份索引為: ${nextMonth})`);

  let processedCount = 0;
  const eligibleEmployees = data.filter(row => {
    const employeeId = row[indices.employeeId];
    const shouldLogDetails = processedCount < 5;
    processedCount++;

    if (!employeeId) return false;

    if (shouldLogDetails) Logger.log(`--- 正在處理員工 ID: ${employeeId} ---`);
    
    const companyName = row[indices.company];
    const dobValue = row[indices.dob];
    const hireDateValue = row[indices.hireDate];
    const resignationDateValue = row[indices.resignationDate];

    if (!companyName || !dobValue || !hireDateValue) {
      if (shouldLogDetails) Logger.log(` -> [篩選排除] 原因: 基本資料不齊全 (投保單位/生日/到職日)。`);
      return false;
    }
    if (resignationDateValue) {
      if (shouldLogDetails) Logger.log(` -> [篩選排除] 原因: 該員工已有離職日期 (${resignationDateValue})。`);
      return false;
    }
    if (employeeId.toString().includes('_') || ['新報', '荃富'].includes(companyName)) {
      if (shouldLogDetails) Logger.log(` -> [篩選排除] 原因: 不符合資格的工號或投保單位 (${companyName})。`);
      return false;
    }
    const dob = new Date(dobValue);
    if (isNaN(dob.getTime())) {
      if (shouldLogDetails) Logger.log(` -> [篩選排除] 原因: 出生日期格式無效: "${dobValue}"`);
      return false;
    }
    if (dob.getMonth() !== nextMonth) {
      if (shouldLogDetails) Logger.log(` -> [篩選排除] 原因: 生日月份不符 (員工生日: ${dob.getMonth() + 1}月, 目標月份: ${nextMonth + 1}月)。`);
      return false;
    }
    const hireDate = new Date(hireDateValue);
    if (isNaN(hireDate.getTime())) {
      if (shouldLogDetails) Logger.log(` -> [篩選排除] 原因: 到職日期格式無效: "${hireDateValue}"`);
      return false;
    }
    const monthsOfService = calculateSeniority(hireDate);
    if (monthsOfService < 3) {
      if (shouldLogDetails) Logger.log(` -> [篩選排除] 原因: 年資未滿三個月 (目前年資: ${monthsOfService} 個月)。`);
      return false;
    }

    if (shouldLogDetails) Logger.log(` -> ✓ [通過] 符合所有資格!`);
    return true;
  });

  Logger.log(`篩選完畢。共有 ${eligibleEmployees.length} 位符合資格的員工。`);

  // --- 整理並分類資料 ---
  const trendforce = [];
  const topology = [];
  eligibleEmployees.forEach(row => {
    const employeeData = {
      departmentCode: row[indices.departmentCode],
      departmentName: row[indices.departmentName],
      employeeId: row[indices.employeeId],
      employeeName: row[indices.employeeName],
      dob: new Date(row[indices.dob]),
      hireDate: new Date(row[indices.hireDate]),
      seniority: calculateSeniority(new Date(row[indices.hireDate]))
    };
    const company = row[indices.company];
    if (company.includes('集邦')) {
      trendforce.push(employeeData);
    } else if (company.includes('拓墣')) {
      topology.push(employeeData);
    } else {
       Logger.log(`[警告]：員工 ${row[indices.employeeId]} 的投保單位名稱 "${company}" 無法分類至集邦或拓墣。`);
    }
  });
  
  Logger.log(`分類完畢。集邦壽星: ${trendforce.length} 位, 拓墣壽星: ${topology.length} 位。`);
  return { trendforce, topology };
}

// ===============================================================
// 2. 產生 Google Doc 生日名單 (*** 已修改為表格格式 ***)
// ===============================================================
function generateBirthdayDoc(employees, companyName) {
  const fullCompanyName = companyName.includes('集邦') ? '集邦科技股份有限公司' : '拓墣科技股份有限公司';

  const nextMonthDate = new Date();
  nextMonthDate.setMonth(nextMonthDate.getMonth() + 1);
  const reportYear = nextMonthDate.getFullYear();
  const reportMonth = nextMonthDate.getMonth() + 1;

  const yy = (reportYear - 2000).toString();
  const mm = ('0' + reportMonth).slice(-2);
  const fileName = `${companyName}${yy}${mm}壽星`;

  const doc = DocumentApp.create(fileName);
  const body = doc.getBody();
  const printDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');

  const header = doc.addHeader();
  const headerTable = header.appendTable([
    [`${fullCompanyName}`],
    [`${reportYear}年${reportMonth}月 每月壽星一覽表`]
  ]);
  headerTable.setBorderWidth(0);
  headerTable.getCell(0, 0).getChild(0).asParagraph().setBold(true);
  headerTable.getCell(1, 0).getChild(0).asParagraph().setBold(true);

  body.appendParagraph('');
  const table = body.appendTable();
  
  const tableHeader = ['部門代號', '部門名稱', '員工代號', '員工姓名', '出生日期', '到職日期','年資'];
  const headerRow = table.appendTableRow();
  tableHeader.forEach(headerText => {
    headerRow.appendTableCell(headerText).getChild(0).asParagraph().setBold(true);
  });

  const currentYear = new Date().getFullYear();
  employees.forEach(emp => {
    let birthDate;
    if (emp.hireDate.getFullYear() === currentYear) {
      birthDate = Utilities.formatDate(emp.dob, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    } else {
      birthDate = Utilities.formatDate(emp.dob, Session.getScriptTimeZone(), 'MM/dd');
    }

    const seniorityText = formatSeniorityForDisplay(emp.seniority);
    const hireDateText = Utilities.formatDate(emp.hireDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');

    const dataRow = table.appendTableRow();
    dataRow.appendTableCell(emp.departmentCode || '');
    dataRow.appendTableCell(emp.departmentName || '');
    dataRow.appendTableCell(emp.employeeId || '');
    dataRow.appendTableCell(emp.employeeName || '');
    dataRow.appendTableCell(birthDate);
    dataRow.appendTableCell(hireDateText);
    dataRow.appendTableCell(seniorityText);
  });

  body.appendParagraph('');
  const bottomInfoTable = body.appendTable([
    [`製表日期： ${printDate}`, `合  計： ${employees.length} 人`]
  ]);
  bottomInfoTable.setBorderWidth(0);
  bottomInfoTable.getCell(0, 0).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  bottomInfoTable.getCell(0, 1).getChild(0).asParagraph().setBold(true).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  const footer = doc.addFooter();
  const footerTable = footer.appendTable([
    [`頁     次： 1 / 1`]
  ]);
  footerTable.setBorderWidth(0);
  footerTable.getCell(0, 0).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  doc.saveAndClose();
  Logger.log(`文件 "${fileName}" 已成功建立。`);
  return doc;
}

// ===============================================================
// 3. 發送包含按鈕的審核 Email
// ===============================================================
function sendApprovalEmail(trendforce, topology, cacheId, recipientEmail) {
  const recipient = recipientEmail || getSetting('PAYMENT_NOTICE_RECIPIENT');
  if (!recipient || recipient === 'your-email@example.com') {
    Logger.log("錯誤：找不到 'PAYMENT_NOTICE_RECIPIENT' 設定，或尚未修改預設值。無法發送審核郵件。");
    throw new Error("請在 CONFIG 中設定有效的 'PAYMENT_NOTICE_RECIPIENT'。");
  }
  const subject = '【審核】每月壽星生日禮金名單';
  
    const webAppUrl = getSetting('WEB_APP_URL');
  if (!webAppUrl) {
    throw new Error("找不到 WEB_APP_URL 設定。請在 Config.js 中設定，或重新部署 Web App。");
  }
  
  const approvalUrl = `${webAppUrl}?action=approve&cacheId=${cacheId}`;
  const rejectUrl = `${webAppUrl}?action=reject&cacheId=${cacheId}`;
  
  const nextMonthDate = new Date();
  nextMonthDate.setMonth(nextMonthDate.getMonth() + 1);
  const reportMonth = nextMonthDate.getMonth() + 1;
  
  let listHtml = '';
  if (trendforce && trendforce.length > 0) {
    listHtml += '<h3>集邦壽星</h3><ul style="padding-left: 20px;">';
    trendforce.forEach(e => {
      const birthDate = Utilities.formatDate(e.dob, Session.getScriptTimeZone(), 'MM/dd');
      listHtml += `<li>${e.employeeName} (${birthDate})</li>`;
    });
    listHtml += '</ul>';
  } else {
    listHtml += '<h3>集邦壽星</h3><p>(無)</p>';
  }
  if (topology && topology.length > 0) {
    listHtml += '<h3>拓墣壽星</h3><ul style="padding-left: 20px;">';
    topology.forEach(e => {
      const birthDate = Utilities.formatDate(e.dob, Session.getScriptTimeZone(), 'MM/dd');
      listHtml += `<li>${e.employeeName} (${birthDate})</li>`;
    });
    listHtml += '</ul>';
  } else {
    listHtml += '<h3>拓墣壽星</h3><p>(無)</p>';
  }
  
  const htmlBody = `
    <div style="font-family: Arial, 'Microsoft JhengHei', sans-serif; line-height: 1.6;">
      <p>您好：</p>
      <p>以下為 ${reportMonth} 月的壽星生日禮金發放建議名單，請您審核。</p>
      ${listHtml}
      <p>確認無誤後，請點擊以下按鈕，系統將自動抓取最新資料以產生名單文件並歸檔。</p>
      <p style="margin: 25px 0;">
        <a href="${approvalUrl}" style="background-color:#4CAF50;color:white;padding:12px 25px;text-decoration:none;border-radius:5px;font-size:16px;font-weight:bold;">✓ 確認無誤，產生並歸檔</a>
        <a href="${rejectUrl}" style="background-color:#f44336;color:white;padding:12px 25px;text-decoration:none;border-radius:5px;font-size:16px;font-weight:bold;margin-left: 10px;">✗ 名單有問題</a>
      </p>
      <p style="color: #888888; font-size: 14px;"><i>(此為系統自動發送郵件)</i></p>
    </div>
  `;
  
  GmailApp.sendEmail(recipient, subject, '', { 
    htmlBody: htmlBody,
    name: getSetting('SENDER_NAME') || 'HR-System'
  });
  Logger.log(`審核郵件已發送至 ${recipient}`);
}


// ===============================================================
// 4. 將檔案移動至指定資料夾
// ===============================================================
function moveFileToFolder(fileId, folderId) {
  if (!fileId || !folderId) {
    throw new Error(`移動檔案失敗：檔案 ID 或資料夾 ID 為空。 File ID: ${fileId}, Folder ID: ${folderId}`);
  }
  try {
    const file = DriveApp.getFileById(fileId);
    const folder = DriveApp.getFolderById(folderId);
    const parents = file.getParents();
    while (parents.hasNext()) {
      const parent = parents.next();
      parent.removeFile(file);
    }
    folder.addFile(file);
    Logger.log(`檔案 ${file.getName()} 已成功移動至資料夾 ${folder.getName()}`);
  } catch (e) {
    Logger.log(`移動檔案 ID ${fileId} 至資料夾 ID ${folderId} 時失敗: ${e.message}`);
    throw new Error(`移動檔案 (ID: ${fileId}) 失敗。請檢查腳本是否有權限存取該檔案以及目標資料夾 (ID: ${folderId})。`);
  }
}

// ===============================================================
// 輔助函式 (計算年資等)
// ===============================================================

/**
 * 計算到下個月底的完整年資(月)
 */
function calculateSeniority(hireDate) {
  const today = new Date();
  const seniorityBaseline = new Date(today.getFullYear(), today.getMonth() + 2, 0);

  let months = (seniorityBaseline.getFullYear() - hireDate.getFullYear()) * 12;
  months -= hireDate.getMonth();
  months += seniorityBaseline.getMonth();

  if (seniorityBaseline.getDate() < hireDate.getDate()) {
    months--;
  }
  
  return months <= 0 ? 0 : months;
}

/**
 * 格式化年資為 "X年Y個月" 的格式
 */
function formatSeniorityForDisplay(totalMonths) {
  if (totalMonths === null || totalMonths === undefined || totalMonths < 0) {
    return '';
  }
  if (totalMonths < 12) {
    return `${totalMonths}個月`;
  }
  const years = Math.floor(totalMonths / 12);
  const months = totalMonths % 12;

  if (months === 0) {
    return `${years}年`;
  } else {
    return `${years}年${months}個月`;
  }
}

// ===============================================================
// 測試用函式 (可由選單觸發)
// ===============================================================
function test_runBirthdayReport() {
  try {
    Logger.log('【測試】開始執行每月壽星報告流程...');
    
    const { trendforce, topology } = getEligibleBirthdayEmployees();

    if (trendforce.length === 0 && topology.length === 0) {
      Logger.log('【測試】下個月沒有符合資格的壽星，流程終止。');
      SpreadsheetApp.getUi().alert('【測試】下個月沒有符合資格的壽星。');
      return;
    }

    const cacheId = Utilities.getUuid();
    CacheService.getScriptCache().put(cacheId, 'pending', 3600);
    Logger.log(`【測試】產生批准ID: ${cacheId}`);

    // 將測試郵件發送給當前使用者
    const testRecipient = Session.getActiveUser().getEmail();
    sendApprovalEmail(trendforce, topology, cacheId, testRecipient);
    
    Logger.log(`【測試】審核郵件已發送至 ${testRecipient}`);
    SpreadsheetApp.getUi().alert(`【測試】審核郵件已成功發送至您的信箱 (${testRecipient})，請前往信箱點擊按鈕以完成後續步驟。`);

  } catch (e) {
    Logger.log(`!!!!!!!!!! 【測試】執行過程中發生嚴重錯誤 !!!!!!!!!!\n${e.stack}`);
    SpreadsheetApp.getUi().alert(`【測試】執行過程中發生錯誤：\n${e.message}`);
  }
}