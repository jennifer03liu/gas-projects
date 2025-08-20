/**
 * @fileoverview 此檔案包含所有與「試用期考核」Email 處理相關的輔助函式。
 */
function calculateDueDate(probationEndDateStr) {
  if (!probationEndDateStr) return '請確認試用截止日';
  try {
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const probationEndDate = new Date(probationEndDateStr); probationEndDate.setHours(0, 0, 0, 0);
    const isOverdue = today > probationEndDate;
    const baseDate = isOverdue ? today : probationEndDate;
    const daysToAdd = isOverdue ? 14 : 7;
    const dueDate = new Date(baseDate.getTime() + (daysToAdd * 24 * 60 * 60 * 1000));
    return formatDateSimple(dueDate); 
  } catch (e) {
    return '計算截止日時發生錯誤';
  }
}

function sendEmployeeNotificationEmail(newFile, rowData, employeeEmail, managerEmail) {
  const subject = `【試用期考核通知】${rowData['員工姓名']} 您好，請完成您的線上考核表`;
  const body = `
    <div style="font-family: Arial, 'Microsoft JhengHei', sans-serif; line-height: 1.6;">
      <p>Dear ${rowData['員工姓名']} 同仁,</p>
      <p>此信件通知您，您的試用期將於 <strong style="color: red;">${formatDateSimple(rowData['試用截止日'])}</strong> 屆滿。</p>
      <p>為完成後續的考核流程，請您點擊下方連結，開啟您的線上考核表並開始進行自我評估。</p>
      <h3 style="color: #0056b3; border-bottom: 2px solid #0056b3; padding-bottom: 5px;">考核流程與時程</h3>
      <ol style="padding-left: 20px;">
          <li><strong>員工自評：</strong>請您開啟線上考核表，完成您的自我評估。</li>
          <li><strong>通知主管並約定面談：</strong><strong style="color: blue;">完成自評後，請務必主動通知您的直屬主管 (${rowData['直屬主管'] || '主管'}) 並與他約定面談時間</strong>。</li>
          <li><strong>面談與評核溝通：</strong>面談後，您的主管將完成評核並與您溝通考核結果。</li>
          <li><strong>列印與簽核：</strong>面談完成後，請您將最終的考核表列印出來，並交由主管簽核。</li>
      </ol>
      <p style="text-align: center; margin: 25px 0;"><a href="${newFile.getUrl()}" style="background-color:#007bff;color:white;padding:12px 25px;text-decoration:none;border-radius:5px;font-size:16px;font-weight:bold;">點此開啟線上考核表</a></p>
      <p>敬請於 <strong style="color: red;">${calculateDueDate(rowData['試用截止日'])}</strong> 前完成所有考核流程並繳交已簽核的考核表至幕僚室，謝謝。</p>
      <p style="color: #888888; font-size: 14px;">(提醒：若繳回日適逢假日，則順延至次一工作日。)</p>
    </div>
  `;
  GmailApp.sendEmail(employeeEmail, subject, '', {
    htmlBody: body + getGmailSignature(),
    name: getSetting('SENDER_NAME'),
    cc: `${managerEmail},${getSetting('HR_MANAGER_CC_EMAIL')}`
  });
}

function sendManagerNotificationEmail(newFile, rowData, managerEmail) {
  const subject = `【試用期屆滿考核通知】${rowData['部門']}同仁 ${rowData['員工姓名']} (${rowData['員工代號']})`;
  const body = `
    <div style="font-family: Arial, 'Microsoft JhengHei', sans-serif; line-height: 1.6;">
      <p>Dear ${rowData['直屬主管'] || '主管'},</p>
      <p>此信件通知您，貴部門同仁 <strong>${rowData['員工姓名']}</strong> (員工代號: ${rowData['員工代號']}) 試用期將於 <strong style="color: red;">${formatDateSimple(rowData['試用截止日'])}</strong> 屆滿。</p>
      <p>請您與同仁依照下列流程，一同完成此次考核。</p>
      <h3 style="color: #0056b3; border-bottom: 2px solid #0056b3; padding-bottom: 5px;">考核流程說明</h3>
      <ol style="padding-left: 20px;">
          <li><strong>員工自評：</strong>此檔案已同步提供予 <strong>${rowData['員工姓名']}</strong>，請其於線上考核表中完成第一部分「工作項目與熟練度自評」。</li>
          <li><strong>主管評核：</strong>待員工完成後，請您開啟同一個線上考核表，接續完成「主管綜合評核」。</li>
          <li><strong>面談溝通：</strong>完成評核後，請您與員工安排面談，溝通考核結果。</li>
          <li><strong>簽核繳交：</strong>面談完成後，請將最終的考核表列印並簽核繳交。</li>
      </ol>
      <p style="text-align: center; margin: 25px 0;"><a href="${newFile.getUrl()}" style="background-color:#007bff;color:white;padding:12px 25px;text-decoration:none;border-radius:5px;font-size:16px;font-weight:bold;">點此開啟線上考核表</a></p>
      <p>敬請於 <strong style="color: red;">${calculateDueDate(rowData['試用截止日'])}</strong> 前將表單填寫完畢，並完成所有考核流程，謝謝。</p>
      <p style="color: #888888; font-size: 14px;">(提醒：若繳回日適逢假日，則順延至次一工作日。)</p>
    </div>
  `;
  GmailApp.sendEmail(managerEmail, subject, '', {
    htmlBody: body + getGmailSignature(),
    name: getSetting('SENDER_NAME'),
    cc: getSetting('HR_MANAGER_CC_EMAIL')
  });
}
