/**
 * @fileoverview 此檔案負責處理每月自動寄送款項申請通知的 Email。
 */
function sendMonthlyEmail() {
  try {
    const now = new Date();
    const rocYear = now.getFullYear() - 1911;
    const currentMonth = now.getMonth() + 1;

    const deadlineDate = (currentMonth === 12) ? `${rocYear}年12月31日` : `${rocYear}年${currentMonth + 1}月5日`;
    const dynamicSubject = `【通知】${rocYear}年${currentMonth}月款項申請(至${deadlineDate}前截止)`;
    const signature = getGmailSignature(); 

    const mainContent = `
      <p>Dear All,</p>
      <p>${rocYear}年${currentMonth}月份尚未請款的個人費用(代墊款、差旅費等)及廠商款項申請,<br>
      請於${deadlineDate}前送至會計處, 若有來不及請款的同仁請先與我們聯絡,</p>
      <p><span style="background-color:yellow; color:red; font-weight:bold;">即日起不受理逾期款項延後請款, 還請各位幫忙配合, 感謝!!</span></p>
      <br>
      <p><span style="background-color:yellow; font-weight:bold;">請款注意事項：</span></p>
      <p>l&nbsp;&nbsp;&nbsp;公司相關請款表單，請至雲端公檔查詢 <a href="https://drive.google.com/drive/folders/1mErE9a4yBYffjIMtOpqZg6x_axGFZF3Y">雲端公檔連結</a></p>
      <p>l&nbsp;&nbsp;&nbsp;發票抬頭：各請款公司別的公司名稱及統一編號請注意不要打錯</p>
      <p>l&nbsp;&nbsp;&nbsp;請款憑證金額及內容請先計算核對</p>
    `;

    const htmlBody = `<div style="font-family: 'Microsoft JhengHei', sans-serif; font-weight: bold;">${mainContent}${signature}</div>`;

    GmailApp.sendEmail(getSetting('PAYMENT_NOTICE_RECIPIENT'), dynamicSubject, '', {
      htmlBody: htmlBody,
      name: getSetting('PAYMENT_SENDER_NAME')
    });

    SpreadsheetApp.getUi().alert("款項申請通知郵件已成功寄送！");
  } catch (e) {
    SpreadsheetApp.getUi().alert("郵件寄送失敗: " + e.message);
  }
}
