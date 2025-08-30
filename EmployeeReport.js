/**
 * @fileoverview 每月定時發送新進與離職員工報告。
 * @version 11.2
 */

/**
 * 【主要執行函式】
 * 由時間觸發器或手動選單呼叫，產生並寄送上一個月的員工異動報告。
 */
function sendMonthlyEmployeeReports() {
  try {
    // 1. 計算報告應屬的月份（上一個月）
    const reportDate = new Date(); // e.g., 8月1日
    reportDate.setDate(0); // 回溯到上個月最後一天, e.g., 7月31日
    
    const reportYear = reportDate.getFullYear().toString(); // "2025"
    const reportMonth = reportDate.getMonth() + 1; // 7
    
    // 2. 準備檔案名稱
    const sourceContactListName = `${reportMonth}月員工通訊錄_Private`;
    const destinationContactListName = `${reportMonth}月員工通訊錄`;

    // 3. 找到來源檔案並產生 Excel 附件
    const sourceFile = findFileInYearFolder(getSetting('SOURCE_BASE_FOLDER_ID'), reportYear, sourceContactListName);
    if (!sourceFile) { 
      throw new Error(`在來源資料夾中找不到檔案: ${sourceContactListName}`);
    }

    const tabsToCopy = JSON.parse(getSetting('TABS_TO_COPY') || '[]');
    const excelBlob = createNewSheetAndExportAsExcel(sourceFile, getSetting('DESTINATION_BASE_FOLDER_ID'), reportYear, destinationContactListName, tabsToCopy);
    if (!excelBlob) {
      throw new Error(`建立並轉換 "${destinationContactListName}" 失敗。`);
    }
    
    // 4. 從員工總控制表篩選異動資料
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSetting('EMPLOYEE_SHEET_NAME'));
    if (!sheet) {
      throw new Error(`在試算表中找不到工作表: '${getSetting('EMPLOYEE_SHEET_NAME')}'`);
    }
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    // 建立一個物件來儲存各欄位的索引值，增加程式碼可讀性
    const colIndex = {
      department: headers.indexOf("部門"), chineseName: headers.indexOf("員工姓名"), englishName: headers.indexOf("匿稱"),
      jobTitle: headers.indexOf("職稱"), extension: headers.indexOf("分機"), mail: headers.indexOf("員工Email"),
      telegram: headers.indexOf("Telegram"), mobile: headers.indexOf("手機"), insuranceUnit: headers.indexOf("投保單位名稱"),
      employeeId: headers.indexOf("員工代號"), startDate: headers.indexOf("到職日期"), endDate: headers.indexOf("離職日期"),
      idNumber: headers.indexOf("身份證字號"), salary: headers.indexOf("薪資"), insurancePlan: headers.indexOf("意外險計畫"),
      birthDate: headers.indexOf("出生日期")
    };
    
    const reportJsMonth = reportMonth-1 ; // JavaScript 的月份是 0-11
    const newHires = data.filter(row => { const d = row[colIndex.startDate]; return d instanceof Date && d.getFullYear().toString() === reportYear && d.getMonth() === reportJsMonth; });
    const departingEmployees = data.filter(row => { const d = row[colIndex.endDate]; return d instanceof Date && d.getFullYear().toString() === reportYear && d.getMonth() === reportJsMonth; });

    // 4a. 【v11.2 修改】篩選出當月到職又離職的員工，在老闆的信件中僅顯示為離職
    // 建立一個包含所有離職員工 ID 的 Set，方便快速查找
    const departingEmployeeIds = new Set(departingEmployees.map(row => row[colIndex.employeeId]));
    // 產生一個給老闆用的新進員工名單，此名單會排除掉也出現在離職名單中的員工
    const newHiresForBoss = newHires.filter(row => !departingEmployeeIds.has(row[colIndex.employeeId]));

    // 5. 寄送 Email
    if (newHires.length > 0 || departingEmployees.length > 0) {
      // 【v11.2 修改】老闆的信件使用過濾後的新進名單 (newHiresForBoss)
      sendBossEmail(newHiresForBoss, departingEmployees, colIndex, excelBlob, reportYear, reportMonth);
      // 保險聯絡人的信件仍使用完整的新進名單 (newHires)，因為加退保都需要通知
      sendInsuranceEmail(newHires, departingEmployees, colIndex, reportYear, reportMonth);
    } else {
      Logger.log(`在 ${reportYear} 年 ${reportMonth} 月沒有偵測到員工異動，但仍會寄送該月通訊錄。`);
      sendBossEmail([], [], colIndex, excelBlob, reportYear, reportMonth);
    }
    // SpreadsheetApp.getUi().alert('員工異動報告已成功寄出！'); // 舊的 UI 通知
    console.log('員工異動報告已成功寄出！'); // 改為背景記錄，不在介面跳出通知
  } catch (e) {
    const errorMessage = `執行員工報告時發生錯誤: ${e.message}\n錯誤堆疊: ${e.stack}`;
    Logger.log(errorMessage); // 在日誌中記錄詳細錯誤
    // SpreadsheetApp.getUi().alert(`執行失敗: ${e.message}`); // 舊的 UI 通知
  }
}

/**
 * 輔助函式：建立新的 Google Sheet，複製分頁，並將其匯出為 Excel Blob。
 */
function createNewSheetAndExportAsExcel(sourceFile, destBaseFolderId, yearStr, fileName, tabsToCopy) {
  let newFileId = null;
  try {
    const destBaseFolder = DriveApp.getFolderById(destBaseFolderId);
    let yearFolders = destBaseFolder.getFoldersByName(yearStr);
    let destinationYearFolder = yearFolders.hasNext() ? yearFolders.next() : destBaseFolder.createFolder(yearStr);
    const newSpreadsheet = SpreadsheetApp.create(fileName);
    newFileId = newSpreadsheet.getId();
    DriveApp.getFileById(newFileId).moveTo(destinationYearFolder);
    const sourceSpreadsheet = SpreadsheetApp.openById(sourceFile.getId());
    tabsToCopy.forEach(tabName => {
      const sourceSheet = sourceSpreadsheet.getSheetByName(tabName);
      if (sourceSheet) { sourceSheet.copyTo(newSpreadsheet).setName(tabName); } 
      else { console.error(`錯誤：在來源檔案中找不到名為 "${tabName}" 的工作表。`); }
    });
    const defaultSheet = newSpreadsheet.getSheets()[0];
    if (newSpreadsheet.getSheets().length > tabsToCopy.length && defaultSheet.getName().match(/^(工作表1|Sheet1)$/)) {
        newSpreadsheet.deleteSheet(defaultSheet);
    }
    SpreadsheetApp.flush();
    const url = `https://docs.google.com/spreadsheets/d/${newFileId}/export?format=xlsx`;
    const options = { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() == 200) {
        return response.getBlob().setName(`${fileName}.xlsx`);
    } else {
        throw new Error(`匯出 Excel 失敗: ${response.getContentText()}`);
    }
  } catch (e) {
    console.error("建立與轉換新檔案時發生錯誤: " + e.toString());
    if (newFileId) {
        try { DriveApp.getFileById(newFileId).setTrashed(true); }
        catch (err) { console.error("刪除不完整檔案時失敗: " + err.toString()); }
    }
    return null;
  }
}


/**
 * 輔助函式：在指定的年份資料夾中尋找檔案。
 */
function findFileInYearFolder(baseFolderId, yearStr, fileName) {
  try {
    const baseFolder = DriveApp.getFolderById(baseFolderId);
    const yearFolders = baseFolder.getFoldersByName(yearStr);
    if (!yearFolders.hasNext()) { console.error(`錯誤：在來源資料夾(${baseFolder.getName()})中找不到年份資料夾：${yearStr}`); return null; }
    const yearFolder = yearFolders.next();
    const files = yearFolder.getFilesByName(fileName);
    if (!files.hasNext()) { console.error(`錯誤：在 ${yearStr} 資料夾中找不到檔案：${fileName}`); return null; }
    return files.next();
  } catch (e) {
    console.error("尋找來源檔案時發生錯誤: " + e.toString());
    return null;
  }
}

/**
 * 輔助函式：寄送 Email 給老闆。
 */
function sendBossEmail(newHires, departingEmployees, colIndex, attachmentBlob, reportYear, reportMonth) {
  const subject = `${reportYear}.${reportMonth}月員工通訊錄`;
  let htmlBody = `<p>Dear ${getSetting('BOSS_NAME')},</p><p>${reportYear}年${reportMonth}月員工異動名單整理如下，該月通訊錄已夾帶於附件中，請查收，謝謝。</p>`;
  if (newHires.length > 0) {
    htmlBody += `<p><b>${reportMonth}月新進員工名單:</b></p><table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse;"><tr style="background-color:#f2f2f2;"><th>部門</th><th>中文名</th><th>匿稱</th><th>職稱</th><th>分機</th><th>Mail</th><th>Telegram</th><th>手機</th></tr>`;
    newHires.forEach(row => { htmlBody += `<tr><td>${row[colIndex.department]}</td><td>${row[colIndex.chineseName]}</td><td>${row[colIndex.englishName]}</td><td>${row[colIndex.jobTitle]}</td><td>${row[colIndex.extension]}</td><td>${row[colIndex.mail]}</td><td>${row[colIndex.telegram]}</td><td>${row[colIndex.mobile]}</td></tr>`; });
    htmlBody += `</table>`;
  }
  if (departingEmployees.length > 0) {
    htmlBody += `<p><b>${reportMonth}月離職員工名單:</b></p><table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse;"><tr style="background-color:#f2f2f2;"><th>部門</th><th>員工姓名</th><th>匿稱</th><th>職稱</th></tr>`;
    departingEmployees.forEach(row => { htmlBody += `<tr><td>${row[colIndex.department]}</td><td>${row[colIndex.chineseName]}</td><td>${row[colIndex.englishName]}</td><td>${row[colIndex.jobTitle]}</td></tr>`; });
    htmlBody += `</table>`;
  }
  if (newHires.length === 0 && departingEmployees.length === 0) { htmlBody += `<p>p.s. 上個月無人員異動。</p>`; }
  
  const signature = getGmailSignature();
  const mailOptions = { htmlBody: htmlBody + signature, attachments: [attachmentBlob] };
  const cc = getSetting('BOSS_CC_EMAIL');
  if (cc) { mailOptions.cc = cc; }
  
  GmailApp.sendEmail(getSetting('BOSS_EMAIL'), subject, "", mailOptions);
  console.log(`已寄送 ${subject} 給老闆。`);
}

/**
 * 輔助函式：寄送 Email 給保險聯絡人。
 */
function sendInsuranceEmail(newHires, departingEmployees, colIndex, reportYear, reportMonth) {
    const subject = `${reportYear}年度${reportMonth}月之三家公司團保加退保名單`;
    let htmlBody = `<p>Dear ${getSetting('INSURANCE_NAME')},</p><p>以下通知您${reportYear}年${reportMonth}月新到職人員與離職人員名單，<br>如有問題請隨時回信告知，謝謝。</p>`;
    if (newHires.length > 0) {
      htmlBody += `<p><b>${reportYear}/${reportMonth} 新進人員名單如下:</b></p><table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse; background-color: #FFFFE0;"><tr style="background-color:#f2f2f2;"><th>投保單位名稱</th><th>員工代號</th><th>員工姓名</th><th>到職日期</th><th>身分證字號</th><th>實際薪資</th><th>意外險計畫</th><th>出生日期</th></tr>`;
      newHires.forEach(row => { htmlBody += `<tr><td>${row[colIndex.insuranceUnit]}</td><td>${row[colIndex.employeeId]}</td><td>${row[colIndex.chineseName]}</td><td>${Utilities.formatDate(new Date(row[colIndex.startDate]), "GMT+8", "yyyy/MM/dd")}</td><td>'${row[colIndex.idNumber]}</td><td>${row[colIndex.salary]}</td><td>${row[colIndex.insurancePlan]}</td><td>${Utilities.formatDate(new Date(row[colIndex.birthDate]), "GMT+8", "yyyy/MM/dd")}</td></tr>`; });
      htmlBody += `</table>`;
    }
    if (departingEmployees.length > 0) {
      htmlBody += `<p><b>${reportYear}/${reportMonth} 離職員工名單如下:</b></p><table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse; background-color: #FFFFE0;"><tr style="background-color:#f2f2f2;"><th>投保單位名稱</th><th>員工代號</th><th>員工姓名</th><th>離職日期</th></tr>`;
      departingEmployees.forEach(row => { htmlBody += `<tr><td>${row[colIndex.insuranceUnit]}</td><td>${row[colIndex.employeeId]}</td><td>${row[colIndex.chineseName]}</td><td>${Utilities.formatDate(new Date(row[colIndex.endDate]), "GMT+8", "yyyy/MM/dd")}</td></tr>`; });
      htmlBody += `</table>`;
    }
    
    const signature = getGmailSignature();
    const mailOptions = { htmlBody: htmlBody + signature };
    const cc = getSetting('INSURANCE_CC_EMAILS');
    if (cc) { mailOptions.cc = cc; }
    
    GmailApp.sendEmail(getSetting('INSURANCE_EMAIL'), subject, "", mailOptions);
    console.log(`已寄送 ${subject} 給保險聯絡人。`);
}
