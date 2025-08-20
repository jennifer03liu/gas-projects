/**
 * @fileoverview 此檔案包含專案中所有共用的輔助函式 (Helper Functions)。
 */
function formatDateSimple(date) {
  if (!date) return '';
  try {
    const d = new Date(date);
    return isNaN(d.getTime()) ? date.toString() : `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
  } catch (e) { return date.toString(); }
}

function validateEmail(email) {
  if (!email || typeof email !== 'string') return false;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function getGmailSignature() {
  try {
    const currentUserEmail = Session.getActiveUser().getEmail();
    const sendAs = Gmail.Users.Settings.SendAs.get('me', currentUserEmail);
    return sendAs?.signature || '';
  } catch (e) { 
    console.error(`取得 Gmail 簽名檔時發生錯誤: ${e.message}`);
    return ''; 
  }
}

function generateFileName(rowData) {
  const today = new Date();
  const datePrefix = `${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}`;
  return `${datePrefix}_${rowData['員工代號'] || '未知'}_${rowData['員工姓名'] || '未知'}_試用期考核表`;
}

function generateSheetTabName(rowData) {
  return `${rowData['員工代號'] || '未知'}_${rowData['員工姓名'] || '未知'}`;
}

function waitFileReady(fileId, timeoutMs) {
  const t0 = Date.now();
  while (Date.now() - t0 < timeoutMs) {
    try { DriveApp.getFileById(fileId).getName(); return; } catch (e) { Utilities.sleep(3000); }
  }
  throw new Error("等待超時，檔案仍然無法使用: " + fileId);
}

function fillBasicData(targetSheet, rowData) {
  // 【修正】處理來自 getSetting 的值，可能是字串或物件
  const dataMappingSetting = getSetting('DATA_MAPPING');
  const dataMapping = typeof dataMappingSetting === 'string' 
    ? JSON.parse(dataMappingSetting || '{}') 
    : (dataMappingSetting || {});

  Object.keys(dataMapping).forEach(header => {
    if (rowData[header] !== undefined && rowData[header] !== null) {
        targetSheet.getRange(dataMapping[header]).setValue(rowData[header]);
    }
  });

  if (rowData['試用起始日'] && rowData['試用截止日']) {
    targetSheet.getRange('G3').setValue(`${formatDateSimple(rowData['試用起始日'])} - ${formatDateSimple(rowData['試用截止日'])}`);
  }
  targetSheet.getRange('B4').setValue(`病假 ${rowData['病假時數'] || 0} 小時 / 事假 ${rowData['事假時數'] || 0} 小時`);
  
  let totalSalary = 0;
  // 【修正】處理來自 getSetting 的值，可能是字串或陣列
  const salaryColsSetting = getSetting('SALARY_COLUMNS');
  const salaryCols = typeof salaryColsSetting === 'string'
    ? JSON.parse(salaryColsSetting || '[]')
    : (salaryColsSetting || []);

  salaryCols.forEach(col => totalSalary += parseFloat(rowData[col] || 0));
  targetSheet.getRange(getSetting('TOTAL_SALARY_CELL')).setValue(totalSalary);
}

function setupFilePermissions(fileId, targetSheet, managerEmail, employeeEmail) {
  const file = DriveApp.getFileById(fileId);
  file.addEditor(managerEmail);
  if (employeeEmail && managerEmail.toLowerCase() !== employeeEmail.toLowerCase()) file.addEditor(employeeEmail);
  
  const protection = targetSheet.protect().setDescription('保護員工基本資料');
  
  // 【修正】處理來自 getSetting 的值，可能是字串或陣列
  const editableRangesSetting = getSetting('EDITABLE_RANGES');
  const editableRanges = typeof editableRangesSetting === 'string'
    ? JSON.parse(editableRangesSetting || '[]')
    : (editableRangesSetting || []);

  if(editableRanges.length > 0) {
    protection.setUnprotectedRanges(targetSheet.getRangeList(editableRanges).getRanges());
  }
  
  const me = Session.getEffectiveUser();
  protection.removeEditors(protection.getEditors()).addEditor(me);
  if (protection.canDomainEdit()) protection.setDomainEdit(false);
}

function searchSupervisors(searchTerm) {
  if (!searchTerm || searchTerm.length < 1) return [];
  const sheet = SpreadsheetApp.openById(getSetting('SPREADSHEET_ID')).getSheetByName('員工總控制表');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const nameIndex = headers.indexOf('員工姓名'), deptIndex = headers.indexOf('部門'), statusIndex = headers.indexOf('離職日期');
  if (nameIndex === -1 || deptIndex === -1 || statusIndex === -1) return [];
  return data.filter(r => r[statusIndex] === '' && r[nameIndex].toString().includes(searchTerm))
             .map(r => `${r[nameIndex]} (${r[deptIndex]})`).slice(0, 10);
}

/**
 * 【全新】[前端呼叫] 根據關鍵字搜尋副本收件人 (搜尋姓名與匿稱)
 * @param {string} searchTerm - 使用者輸入的搜尋關鍵字
 * @returns {Object[]} - 符合條件的人員物件陣列 [{display: "姓名 (匿稱)", email: "..."}]
 */
function searchCcRecipients(searchTerm) {
  try {
    if (!searchTerm || searchTerm.length < 1) return [];
    
    const mainSheet = SpreadsheetApp.openById(getSetting('SPREADSHEET_ID')).getSheetByName('員工總控制表');
    if (!mainSheet) {
      throw new Error("找不到名為 '員工總控制表' 的工作表。");
    }
    
    const data = mainSheet.getDataRange().getValues();
    const headers = data.shift();
    
    const nameIndex = headers.indexOf('員工姓名');
    const nicknameIndex = headers.indexOf('匿稱');
    const emailIndex = headers.indexOf('員工Email');
    const statusIndex = headers.indexOf('離職日期');
    
    // 【修正】加入欄位檢查
    if (nameIndex === -1 || emailIndex === -1 || statusIndex === -1) {
      throw new Error("工作表中缺少必要的欄位：員工姓名、員工Email 或 離職日期");
    }
    
    const recipients = data
      .filter(row => {
        const name = row[nameIndex] || '';
        const nickname = row[nicknameIndex] || '';
        const email = row[emailIndex] || '';
        const isActive = row[statusIndex] === '';
        
        return isActive && email && (name.includes(searchTerm) || nickname.includes(searchTerm));
      })
      .map(row => {
        const name = row[nameIndex];
        const nickname = row[nicknameIndex];
        const email = row[emailIndex];
        const display = nickname ? `${name} (${nickname})` : name;
        return { display: display, email: email };
      });
      
    return recipients.slice(0, 10); // 最多返回10筆結果
  } catch (e) {
    throw new Error(`搜尋副本收件人時發生錯誤: ${e.message}`);
  }
}

function generateEmployeeId(company, employeeType, onboardingDateStr) {
  const serialSheet = SpreadsheetApp.openById(getSetting('SPREADSHEET_ID')).getSheetByName('編碼紀錄');
  if (!serialSheet) throw new Error("找不到 '編碼紀錄' 工作表。");
  const rocYear = new Date(onboardingDateStr).getFullYear() - 1911;
  let ruleName = '', prefix = '', isDescending = false;

  switch (company) {
    case '集邦科技': case '荃富科技':
      ruleName = `${company}_${employeeType}_${rocYear}`;
      prefix = rocYear.toString();
      if (employeeType === '非正職') isDescending = true;
      break;
    case '拓墣科技': ruleName = `${company}_${employeeType}_${rocYear}`;
      prefix = `EM`.toString();
      if (employeeType === '非正職') isDescending = true; // 修正: 這段邏輯似乎重複且位置不對，但暫時保留原樣
      break;
    case '新報科技': ruleName = `${company}_${employeeType}_${rocYear}`;
      prefix = `TN${rocYear}`.toString();
      if (employeeType === '非正職') isDescending = true; // 修正: 這段邏輯似乎重複且位置不對，但暫時保留原樣
      break;
    default: throw new Error('無效的公司別');
  }

  const serialData = serialSheet.getDataRange().getValues();
  const ruleIndex = serialData.findIndex(row => row[0] === ruleName);
  if (ruleIndex === -1) throw new Error(`在[編碼紀錄]中找不到規則: ${ruleName}`);
  
  const currentSerial = parseInt(serialData[ruleIndex][1]);
  const newSerial = isDescending ? currentSerial - 1 : currentSerial + 1;
  
  serialSheet.getRange(ruleIndex + 1, 2).setValue(newSerial);
  
  if (company === '拓墣科技') return prefix + ('0000' + newSerial).slice(-4);
  if (company === '新報科技') return prefix + ('000' + newSerial).slice(-3);
  return prefix + newSerial.toString();
}

function getSupervisorEmail(supervisorString) {
  if (!supervisorString) return '';
  const sheet = SpreadsheetApp.openById(getSetting('SPREADSHEET_ID')).getSheetByName('員工總控制表');
  if (!sheet) return '';
  const supervisorName = supervisorString.split(' (')[0];
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const nameIndex = headers.indexOf('員工姓名'), emailIndex = headers.indexOf('員工Email'), statusIndex = headers.indexOf('離職日期');
  if (nameIndex === -1 || emailIndex === -1 || statusIndex === -1) {
    throw new Error("工作表中缺少必要的欄位：員工姓名、員工Email 或 離職日期");
  }
  const row = data.find(r => r[nameIndex] === supervisorName && r[statusIndex] === ''); // 增加在職判斷
  return row ? row[emailIndex] : '';
}
