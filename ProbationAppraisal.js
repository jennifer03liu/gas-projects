/**
 * @fileoverview 這是「試用期考核」功能的核心檔案，負責處理主要的自動化流程。
 */
function processAndEmailEditableSheets(sendTo) {
  try {
    const masterSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const ui = SpreadsheetApp.getUi();
    if (masterSheet.getLastRow() < 2) { ui.alert('工作表中沒有資料可處理'); return; }
    
    const data = masterSheet.getDataRange().getValues();
    const headers = data.shift();
    const colIndices = {
      status: headers.indexOf('通知信狀態'),
      managerEmail: headers.indexOf('主管Email'),
      employeeEmail: headers.indexOf('員工Email')
    };

    if (Object.values(colIndices).some(i => i === -1)) {
      ui.alert('錯誤：找不到「通知信狀態」、「主管Email」或「員工Email」等必要欄位。');
      return;
    }
    
    data.forEach((row, index) => {
      if (row[colIndices.status] !== '待處理') return;
      
      const currentRow = index + 2;
      const rowData = Object.fromEntries(headers.map((header, i) => [header, row[i]]));
      
      try {
        const managerEmail = String(rowData['主管Email'] || '').trim();
        const employeeEmail = String(rowData['員工Email'] || '').trim();
        if (!validateEmail(managerEmail)) throw new Error(`主管Email格式不正確: "${managerEmail}"`);
        if (!validateEmail(employeeEmail)) throw new Error(`員工Email格式不正確: "${employeeEmail}"`);
        
        const newSheetName = generateFileName(rowData);
        const templateFile = DriveApp.getFileById(getSetting('TEMPLATE_SHEET_ID'));
        const destinationFolder = DriveApp.getFolderById(getSetting('DESTINATION_FOLDER_ID'));
        const newFile = templateFile.makeCopy(newSheetName, destinationFolder);
        const fileId = newFile.getId();
        
        waitFileReady(fileId, 30000);
        
        const newSheet = SpreadsheetApp.openById(fileId);
        const targetSheet = newSheet.getSheets()[0];
        
        targetSheet.setName(generateSheetTabName(rowData));
        fillBasicData(targetSheet, rowData); 
        setupFilePermissions(fileId, targetSheet, managerEmail, employeeEmail);
        SpreadsheetApp.flush();
        
        if (sendTo === 'manager') {
          sendManagerNotificationEmail(newFile, rowData, managerEmail);
        } else {
          sendEmployeeNotificationEmail(newFile, rowData, employeeEmail, managerEmail);
        }
              
        masterSheet.getRange(currentRow, colIndices.status + 1).setValue(`已於 ${new Date().toLocaleDateString('zh-TW')} 寄出`);
      } catch (e) {
        masterSheet.getRange(currentRow, colIndices.status + 1).setValue(`處理失敗: ${e.message}`);
      }
    });
    
    ui.alert('所有「待處理」的線上考核表皆已處理完成！');
  } catch (error) {
    SpreadsheetApp.getUi().alert('系統發生致命錯誤：' + error.message);
  }
}
