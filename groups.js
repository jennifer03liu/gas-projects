/**
 * @fileoverview 此檔案負責根據 Google Sheet 中的員工資料，每日自動同步成員至對應的 Google Groups。
 */
function syncAllDepartmentGroups() {
  Logger.log("========= 開始執行每日群組同步任務 =========");
  try {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const requiredMembersByDept = getRequiredMembersForToday(today);

    if (Object.keys(requiredMembersByDept).length === 0) {
      Logger.log("從工作表中未篩選出任何應在職員工，任務結束。");
      return;
    }
    
    const departmentMapping = JSON.parse(getSetting('DEPARTMENT_GROUP_MAPPING') || '{}');

    for (const department in departmentMapping) {
      const groupEmail = departmentMapping[department];
      Logger.log(`\n--- 正在處理部門 [${department}], 群組 [${groupEmail}] ---`);
      const requiredEmails = requiredMembersByDept[department] || new Set();
      syncSingleGroup(groupEmail, requiredEmails);
    }
    Logger.log("========= 所有群組同步任務執行完畢 =========");
    SpreadsheetApp.getUi().alert("所有部門群組皆已同步完成！");
  } catch (e) {
    SpreadsheetApp.getUi().alert(`同步群組時發生錯誤: ${e.message}`);
  }
}

function getRequiredMembersForToday(today) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSetting('GROUP_SYNC_SHEET_NAME'));
  if (!sheet) throw new Error(`找不到名為 "${getSetting('GROUP_SYNC_SHEET_NAME')}" 的工作表。`);

  const data = sheet.getDataRange().getValues();
  const headers = data.shift().map(h => h.trim());
  const headerNames = JSON.parse(getSetting('GROUP_SYNC_HEADER_NAMES') || '{}');
  
  const membersByDept = {};
  data.forEach(row => {
    const rowData = Object.fromEntries(headers.map((h, i) => [h, row[i]]));
    const email = rowData[headerNames.EMAIL];
    const startDate = rowData[headerNames.START_DATE] ? new Date(rowData[headerNames.START_DATE]) : null;
    const endDate = rowData[headerNames.END_DATE] ? new Date(rowData[headerNames.END_DATE]) : null;
    const department = rowData[headerNames.DEPARTMENT];

    if (!email || !startDate || !department) return;
    
    startDate.setHours(0,0,0,0);
    if(endDate) endDate.setHours(0,0,0,0);

    if ((startDate <= today) && (!endDate || endDate >= today)) {
      if (!membersByDept[department]) membersByDept[department] = new Set();
      membersByDept[department].add(email.toLowerCase().trim());
    }
  });
  return membersByDept;
}

function syncSingleGroup(groupEmail, requiredEmails) {
  try {
    const currentMemberEmails = new Set();
    let pageToken;
    do {
      const result = AdminDirectory.Members.list(groupEmail, { pageToken, maxResults: 200, fields: 'members(email),nextPageToken' });
      if (result.members) result.members.forEach(m => currentMemberEmails.add(m.email.toLowerCase().trim()));
      pageToken = result.nextPageToken;
    } while (pageToken);
    
    const toAdd = [...requiredEmails].filter(e => !currentMemberEmails.has(e));
    toAdd.forEach(email => {
      try { AdminDirectory.Members.insert({ email, role: 'MEMBER' }, groupEmail); Logger.log(`  (+) 加入: ${email}`); } 
      catch (e) { Logger.log(`  (x) 加入失敗: ${email}, ${e.message}`); }
    });

    const toRemove = [...currentMemberEmails].filter(e => !requiredEmails.has(e));
    toRemove.forEach(email => {
      try { AdminDirectory.Members.remove(groupEmail, email); Logger.log(`  (-) 移除: ${email}`); } 
      catch (e) { Logger.log(`  (x) 移除失敗: ${email}, ${e.message}`); }
    });

    if(!toAdd.length && !toRemove.length) Logger.log("  成員名單已是最新。");
  } catch (err) {
    Logger.log(`處理群組 ${groupEmail} 時發生錯誤: ${err.message}`);
  }
}
