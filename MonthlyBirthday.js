/**
 * @fileoverview 每月壽星生日禮金發放名單產生器
 * 1. 每月月底自動執行，抓取次月生日的員工。
 * 2. 根據條件篩選資格。
 * 3. 產生兩份 Google Doc 名單 (集邦、拓墣)。
 * 4. 發送一封帶有操作按鈕的 Email 給指定人員。
 * 5. 待點擊「確認」後，將文件歸檔至指定 Google Drive 資料夾。
 */

// ===============================================================
// 主要觸發函式 (預計每月月底執行)
// ===============================================================
function createMonthlyBirthdayReport() {
  try {
    const { trendforce, topology } = getEligibleBirthdayEmployees();

    if (trendforce.length === 0 && topology.length === 0) {
      Logger.log('下個月沒有符合資格的壽星。');
      // 或者可以發一封通知信
      // GmailApp.sendEmail(getSetting('PAYMENT_NOTICE_RECIPIENT'), '每月壽星報告 - 無符合資格人員', '系統已於今日執行，下個月沒有符合生日禮金資格的壽星。');
      return;
    }

    const trendforceDoc = trendforce.length > 0 ? generateBirthdayDoc(trendforce, '集邦') : null;
    const topologyDoc = topology.length > 0 ? generateBirthdayDoc(topology, '拓墣') : null;

    sendApprovalEmail(trendforceDoc, topologyDoc);

  } catch (e) {
    Logger.log(`建立壽星報告時發生錯誤: ${e.message}`);
    // 可加入錯誤通知信
    GmailApp.sendEmail(getSetting('PAYMENT_NOTICE_RECIPIENT'), '【錯誤】每月壽星報告產生失敗', `執行過程中發生錯誤：\n${e.stack}`);
  }
}

// ===============================================================
// 1. 取得並篩選符合資格的壽星
// ===============================================================
function getEligibleBirthdayEmployees() {
  // --- 設定 ---
  // 從 Config 讀取欄位名稱設定，並處理可能的字串格式
  const columnNamesSetting = getSetting('BIRTHDAY_COLUMN_NAMES');
  const COLUMN_NAMES = typeof columnNamesSetting === 'string'
    ? JSON.parse(columnNamesSetting)
    : columnNamesSetting;

  if (!COLUMN_NAMES) {
    throw new Error("在設定中找不到 'BIRTHDAY_COLUMN_NAMES'。");
  }

  // --- 取得資料 ---
  const sheet = SpreadsheetApp.openById(getSetting('SPREADSHEET_ID')).getSheetByName(getSetting('EMPLOYEE_SHEET_NAME'));
  if (!sheet) {
    throw new Error(`找不到名為 "${getSetting('EMPLOYEE_SHEET_NAME')}" 的工作表`);
  }
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  // --- 建立欄位索引 ---
  const indices = {};
  for (const key in COLUMN_NAMES) {
    const index = headers.indexOf(COLUMN_NAMES[key]);
    if (index === -1) {
      throw new Error(`在員工總控制表中找不到欄位: "${COLUMN_NAMES[key]}"`);
    }
    indices[key] = index;
  }

  // --- 篩選邏輯 ---
  const today = new Date();
  const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1).getMonth();
  const seniorityBaseline = new Date(today.getFullYear(), today.getMonth() + 2, 0); // 次月月底

  const eligibleEmployees = data.filter(row => {
    // 1. 排除員工代號包含 '_'
    const employeeId = row[indices.employeeId];
    if (employeeId.toString().includes('_')) return false;

    // 2. 排除特定投保單位
    const insuranceUnit = row[indices.insuranceUnit];
    if (['新報', '荃富'].includes(insuranceUnit)) return false;

    // 3. 判斷生日是否在下個月
    const dob = new Date(row[indices.dob]);
    if (dob.getMonth() !== nextMonth) return false;

    // 4. 判斷年資是否滿三個月
    const hireDate = new Date(row[indices.hireDate]);
    const monthsOfService = (seniorityBaseline.getFullYear() - hireDate.getFullYear()) * 12 + (seniorityBaseline.getMonth() - hireDate.getMonth());
    if (monthsOfService < 3) return false;
    
    return true;
  });

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
      age: calculateAge(new Date(row[indices.dob])),
      seniority: calculateSeniority(new Date(row[indices.hireDate]))
    };

    const company = row[indices.company];
    if (company.includes('集邦')) {
      trendforce.push(employeeData);
    } else if (company.includes('拓墣')) {
      topology.push(employeeData);
    }
  });

  return { trendforce, topology };
}


// ===============================================================
// 2. 產生 Google Doc 生日名單
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

  // --- 設定文件頁首 ---
  const header = doc.addHeader();
  const headerTable = header.appendTable([
    [`${fullCompanyName}`, `列印日期： ${printDate}`],
    [`${reportYear}年${reportMonth}月 每月壽星一覽表`, `頁    次： 1 / 1`]
  ]);
  headerTable.setBorderWidth(0);
  headerTable.getCell(0, 0).getChild(0).asParagraph().setBold(true);
  headerTable.getCell(1, 0).getChild(0).asParagraph().setBold(true);
  
  // --- 設定主內容 ---
  body.appendParagraph(''); // 確保內容從頁首底下開始

  const table = body.appendTable([
    ['部門代號', '部門名稱', '員工代號', '員工姓名', '出生日期', '年齡', '年資(月)']
  ]);

  employees.forEach(emp => {
    const birthDate = Utilities.formatDate(emp.dob, Session.getScriptTimeZone(), 'MM/dd');
    table.appendRow([
      emp.departmentCode,
      emp.departmentName,
      emp.employeeId,
      emp.employeeName,
      birthDate,
      emp.age,
      emp.seniority
    ]);
  });
  
  // 美化表格樣式
  table.getRow(0).editAsText().setBold(true);
  
  // --- 設定文件頁尾 ---
  const footer = doc.addFooter();
  const footerTable = footer.appendTable([
    [`合  計： ${employees.length} 人`, `NO：0050-A4-2`]
  ]);
  footerTable.setBorderWidth(0);
  footerTable.getCell(0, 0).getChild(0).asParagraph().setBold(true);
  footerTable.getCell(0, 1).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  doc.saveAndClose();
  return doc;
}


// ===============================================================
// 3. 發送包含按鈕的審核 Email
// ===============================================================
function sendApprovalEmail(trendforceDoc, topologyDoc, recipientEmail) {
  const recipient = recipientEmail || getSetting('PAYMENT_NOTICE_RECIPIENT');
  const subject = '【審核】每月壽星生日禮金名單';
  
  const webAppUrl = ScriptApp.getService().getUrl();
  
  let trendforceParams = trendforceDoc ? `&docId1=${trendforceDoc.getId()}` : '';
  let topologyParams = topologyDoc ? `&docId2=${topologyDoc.getId()}` : '';

  const approvalUrl = `${webAppUrl}?action=approve${trendforceParams}${topologyParams}`;
  const rejectUrl = `${webAppUrl}?action=reject`;

  let htmlBody = `
    <html><body>
      <p>您好：</p>
      <p>附件為下個月的壽星生日禮金發放建議名單，請您審核。</p>
      <ul>
  `;
  if (trendforceDoc) {
    htmlBody += `<li><b>集邦壽星名單:</b> <a href="${trendforceDoc.getUrl()}">${trendforceDoc.getName()}</a></li>`;
  }
  if (topologyDoc) {
    htmlBody += `<li><b>拓墣壽星名單:</b> <a href="${topologyDoc.getUrl()}">${topologyDoc.getName()}</a></li>`;
  }
  htmlBody += `
      </ul>
      <p>確認無誤後，請點擊以下按鈕，系統將自動將檔案歸檔至對應資料夾。</p>
      <a href="${approvalUrl}" style="text-decoration: none;">
        <button style="background-color: #4CAF50; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer;">
          ✓ 確認無誤，進行歸檔
        </button>
      </a>
      <a href="${rejectUrl}" style="text-decoration: none;">
        <button style="background-color: #f44336; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; margin-left: 10px;">
          ✗ 名單有問題
        </button>
      </a>
      <br><br>
      <p><i>(此為系統自動發送郵件)</i></p>
    </body></html>
  `;

  GmailApp.sendEmail(recipient, subject, '', { htmlBody: htmlBody });
}


// ===============================================================
// 4. Web App 處理 Email 按鈕點擊
// ===============================================================
function doGet(e) {
  try {
    const action = e.parameter.action;

    if (action === 'approve') {
      const docId1 = e.parameter.docId1;
      const docId2 = e.parameter.docId2;
      
      // 從 Config 讀取資料夾 ID
      const trendforceFolderId = getSetting('TRENDFORCE_BIRTHDAY_FOLDER_ID');
      const topologyFolderId = getSetting('TOPOLOGY_BIRTHDAY_FOLDER_ID');

      if (docId1) {
        moveFileToFolder(docId1, trendforceFolderId);
      }
      if (docId2) {
        moveFileToFolder(docId2, topologyFolderId);
      }
      
      return ContentService.createTextOutput('操作成功！壽星名單已確認並歸檔至指定資料夾。');

    } else if (action === 'reject') {
      // 可在此加入通知，例如發信給IT人員或記錄在某個試算表中
      return ContentService.createTextOutput('操作已記錄。系統將不會移動檔案。如有需要，請手動修正問題。');
    } else {
      return ContentService.createTextOutput('無效的操作。');
    }
  } catch (err) {
    return ContentService.createTextOutput(`處理您的請求時發生錯誤: ${err.message}`);
  }
}

function moveFileToFolder(fileId, folderId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const folder = DriveApp.getFolderById(folderId);
    file.moveTo(folder);
    Logger.log(`檔案 ${file.getName()} 已成功移動至資料夾 ${folder.getName()}`);
  } catch (e) {
    Logger.log(`移動檔案 ${fileId} 至資料夾 ${folderId} 時失敗: ${e.message}`);
    // 可以考慮拋出錯誤或發送失敗通知
    throw new Error(`移動檔案 ${fileId} 失敗`);
  }
}


// ===============================================================
// 輔助函式
// ===============================================================
function calculateAge(birthDate) {
  const today = new Date();
  let age = today.getFullYear() - birthDate.getFullYear();
  const m = today.getMonth() - birthDate.getMonth();
  if (m < 0 || (m === 0 && today.getDate() < birthDate.getDate())) {
    age--;
  }
  return age;
}

function calculateSeniority(hireDate) {
    const today = new Date();
    const seniorityBaseline = new Date(today.getFullYear(), today.getMonth() + 2, 0); // 次月月底
    let months = (seniorityBaseline.getFullYear() - hireDate.getFullYear()) * 12;
    months -= hireDate.getMonth();
    months += seniorityBaseline.getMonth();
    return months <= 0 ? 0 : months;
}

// ===============================================================
// 測試用函式
// ===============================================================
/**
 * 手動執行此函式以進行測試。
 * 它會執行完整流程，但會將審核 Email 發送給當前登入的使用者。
 */
function test_runBirthdayReport() {
  try {
    const { trendforce, topology } = getEligibleBirthdayEmployees();

    if (trendforce.length === 0 && topology.length === 0) {
      Logger.log('【測試】下個月沒有符合資格的壽星。');
      Browser.msgBox('測試執行完成', '下個月沒有符合資格的壽星。', Browser.Buttons.OK);
      return;
    }

    const trendforceDoc = trendforce.length > 0 ? generateBirthdayDoc(trendforce, '集邦') : null;
    const topologyDoc = topology.length > 0 ? generateBirthdayDoc(topology, '拓墣') : null;

    const testRecipient = Session.getActiveUser().getEmail();
    sendApprovalEmail(trendforceDoc, topologyDoc, testRecipient);
    
    Logger.log(`【測試】測試郵件已發送至 ${testRecipient}`);
    Browser.msgBox('測試執行完成', `審核郵件已發送至您的信箱: ${testRecipient}`, Browser.Buttons.OK);

  } catch (e) {
    Logger.log(`【測試】建立壽星報告時發生錯誤: ${e.message}`);
    Browser.msgBox('測試執行失敗', `執行過程中發生錯誤：\n${e.stack}`, Browser.Buttons.OK);
  }
}