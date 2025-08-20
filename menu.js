/**
 * @fileoverview 這份檔案負責在 Google Sheet 檔案開啟時，建立自訂的操作選單。
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('流程整合')
    .addItem('產生考核表 (寄給主管)', 'runProcessForManager')
    .addItem('產生考核表 (寄給員工)', 'runProcessForEmployee')
    .addSeparator()
    .addItem('開啟新進人員任用登錄', 'showRecruitmentWebApp')
    .addItem('僅產生任用單', 'generateOfferLetterOnly')
    .addItem('寄送每月款項申請通知', 'sendMonthlyEmail')
    .addSeparator()
    .addItem('手動同步部門群組', 'syncAllDepartmentGroups')
    .addSeparator()
    .addItem('手動觸發每月員工報告', 'sendMonthlyEmployeeReports')
    .addSeparator()
    .addItem('修改信件設定', 'showEmailSettingsUI')
    .addToUi();
}

function runProcessForManager() {
  processAndEmailEditableSheets('manager');
}

function runProcessForEmployee() {
  processAndEmailEditableSheets('employee');
}
