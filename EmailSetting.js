/**
 * @fileoverview 此檔案負責處理信件設定介面的後端邏輯，
 * 包括顯示 UI、讀取和儲存設定。
 */

function showEmailSettingsUI() {
  const html = HtmlService.createHtmlOutputFromFile('EmailSettings')
      .setWidth(800)
      .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, '信件設定管理');
}

function getEmailSettings() {
  const keys = [
    'HR_MANAGER_CC_EMAIL', 'SENDER_NAME', 'BOSS_EMAIL', 'BOSS_CC_EMAIL', 
    'INSURANCE_EMAIL', 'INSURANCE_CC_EMAILS'
  ];
  const settings = {};
  keys.forEach(key => {
    settings[key] = getSetting(key);
  });
  return settings;
}

function saveEmailSettings(settings) {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.setProperties(settings, false);
    return '設定已成功儲存！';
  } catch (e) {
    throw new Error(`儲存設定時發生錯誤: ${e.message}`);
  }
}
