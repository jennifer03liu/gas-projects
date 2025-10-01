/**
 * @fileoverview 集中管理專案的預設設定值，並提供動態讀取客製化設定的函式。
 */

/**
 * 【核心函式】取得設定值。
 * 優先從 PropertiesService (使用者自訂的儲存區) 讀取，
 * 如果沒有，則回傳 CONFIG 物件中的預設值。
 * @param {string} key - 要取得的設定鍵名。
 * @returns {string | null} 設定值。
 */
function getSetting(key) {
  const userProperties = PropertiesService.getScriptProperties();
  const userValue = userProperties.getProperty(key);
  // 如果使用者有儲存過自訂值，就使用它；否則，使用下面的預設值。
  return userValue !== null ? userValue : (CONFIG[key] || null);
}


/**
 * @description 專案的預設設定值。
 * 當您從未在 UI 介面儲存過設定時，程式會使用這些值。
 */
const CONFIG = {
  // ===============================================================
  // 考核表相關設定
  // ===============================================================
  TEMPLATE_SHEET_ID: '12VcBawYb8C6PIF8UuhoRgnDjtn7XhPL8jh1ekqkmHRc',
  DESTINATION_FOLDER_ID: '1NdnJ-p9PwFEJhh5HzfUKz6o0tgRWkjpV',
  HR_MANAGER_CC_EMAIL: 'rolinghsu@trendforce.com',
  EDITABLE_RANGES: ['B8:I12', 'I16:I22', 'G31', 'H27:I27', 'B26:B32', 'B24:I24', 'F34:I34'],
  DATA_MAPPING: { '部門': 'B2', '員工姓名': 'B3', '員工代號': 'G2', '職稱': 'G4' },
  SALARY_COLUMNS: ['薪資', '職務加給', '伙食津貼', '全勤獎金'],
  TOTAL_SALARY_CELL: 'E27',

  // ===============================================================
  // 新進人員任用流程設定
  // ===============================================================
  SENDER_NAME: 'Jennifer Liu',
  SPREADSHEET_ID: '1UuLepzRnekct5x4DGcmgFkn7FDNoSgPjR8CG1L1w1gM',
  PDF_FOLDER_ID: '18fIQE0sjV9HkiNes9g6nX9T-2wA0Y6tq',
  NEW_HIRE_FORM_ID: '1tHWGNwMzlhZ0WJvPC7Y0nSGSKQX_rLyA-suoetF_ILc',

  // ===============================================================
  // 每月款項申請通知設定
  // ===============================================================
  PAYMENT_NOTICE_RECIPIENT: 'jennifer03liu@gmail.com',
  PAYMENT_SENDER_NAME: '管理中心會計處',

  // ===============================================================
  // Google Groups 同步設定
  // ===============================================================
  GROUP_SYNC_SHEET_NAME: "test",
  DEPARTMENT_GROUP_MAPPING: { 'test123': 'test123@trendforce.com' },
  GROUP_SYNC_HEADER_NAMES: { EMAIL: '員工Email', START_DATE: '到職日期', END_DATE: '離職日期', DEPARTMENT: '部門' },

  // ===============================================================
  // 每月壽星生日報告設定
  // ===============================================================
  // 員工總控制表中對應的欄位名稱
  BIRTHDAY_COLUMN_NAMES: {
    company: '投保單位名稱',
    departmentCode: '部門代號',
    departmentName: '部門名稱',
    employeeId: '員工代號',
    employeeName: '員工姓名',
    dob: '出生日期',
    hireDate: '到職日期',
    insuranceUnit: '投保單位名稱'
  },
  // 產生後的壽星名單要歸檔到的資料夾 ID
  TRENDFORCE_BIRTHDAY_FOLDER_ID: '1fyV1ljHXoZP8LAY5Um4y4_Ik7JUMoqtK', // 集邦
  TOPOLOGY_BIRTHDAY_FOLDER_ID: '1Hm925rsc7KWuY9Bqm7VlZJnc4uznGjIq',   // 拓墣,

  // ===============================================================
  // 每月員工異動報告設定
  // ===============================================================
  BOSS_NAME: "Kevin",
  BOSS_EMAIL: "KevinLin@trendforce.com",
  BOSS_CC_EMAIL: "rolinghsu@trendforce.com,jenniferliu@trendforce.com",
  INSURANCE_NAME: "Elsie",
  INSURANCE_EMAIL: "Elsie-LY.Huang@nanshan.com.tw",
  INSURANCE_CC_EMAILS: "charlenelyn1118@gmail.com,rolinghsu@trendforce.com,jenniferliu@trendforce.com",
  SOURCE_BASE_FOLDER_ID: "1ISfrjj6i1Ci_8YN_CUD7roYcwuwNO8WD",
  DESTINATION_BASE_FOLDER_ID: "14M6p6oXXIAiedBgjSLAl8JUktYoEmENf",
  EMPLOYEE_SHEET_NAME: "員工總控制表",
  TABS_TO_COPY: ["集邦、拓墣通訊錄", "新報"]
};
