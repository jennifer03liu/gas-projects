/**
 * @fileoverview 此檔案負責處理新進人員任用流程的後端邏輯。
 */
function showRecruitmentForm() {
  return HtmlService.createHtmlOutputFromFile('RecruitmentForm').setTitle('新進人員任用登錄').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function showRecruitmentWebApp() {
  const html = HtmlService.createHtmlOutputFromFile('RecruitmentForm').setWidth(800).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, '新進人員任用登錄');
}

function processRecruitmentForm(formData) {
  try {
    if (!formData || !formData.supervisor || !formData.company || !formData.employeeName) ;
    // ... 驗證和準備工作 ...
    const employeeId = generateEmployeeId(formData.company, formData.employeeType, formData.onboardingDate);
    const supervisorEmail = getSupervisorEmail(formData.supervisor);
    const verificationCode = Utilities.getUuid();
    const pdfFile = createOfferPdfFromTemplate(formData, employeeId);
    
    // 【修正】先寄送Email
    sendOfferEmail(formData, pdfFile, verificationCode);
    
    // 【修正】Email寄送成功後再寫入資料庫
    const mainSheet = SpreadsheetApp.openById(getSetting('SPREADSHEET_ID')).getSheetByName('員工總控制表');
    if (!mainSheet) throw new Error("找不到名為 '員工總控制表' 的工作表。");
    
    const dateParts = formData.onboardingDate.split('-');
    const onboardingDateObject = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]);
    const supervisorName = formData.supervisor.split(' (')[0];
    
    mainSheet.appendRow([
      employeeId, formData.employeeName, formData.department, supervisorName, '', onboardingDateObject,
      '', '', '', '', '', '', '', '', formData.salary, '', '', supervisorEmail, '', '', formData.company, 
      '已寄送Offer', formData.otherSalaryInfo, verificationCode, ''
    ]);
    
    return `成功提交！新進人員 ${formData.employeeName} (員工代號: ${employeeId}) 的錄取通知信已寄出。`;
  } catch (e) {
    throw new Error(`後端處理失敗: ${e.message}`);
  }
}


// function processRecruitmentForm(formData) {
//   try {
//     if (!formData || !formData.supervisor || !formData.company || !formData.employeeName) {
//       throw new Error("提交失敗：必填欄位（如主管、公司、姓名）不完整。");
//     }

//     const employeeId = generateEmployeeId(formData.company, formData.employeeType, formData.onboardingDate);
//     const supervisorEmail = getSupervisorEmail(formData.supervisor);
//     const verificationCode = Utilities.getUuid();
//     const pdfFile = createOfferPdfFromTemplate(formData, employeeId);
    
//     const mainSheet = SpreadsheetApp.openById(getSetting('SPREADSHEET_ID')).getSheetByName('員工總控制表');
//     if (!mainSheet) throw new Error("找不到名為 '員工總控制表' 的工作表。");
//     // 【更新】將副本收件人Email傳遞下去
//     // 使用更穩定的方式建立日期物件，避免時區問題
//     const dateParts = formData.onboardingDate.split('-'); // e.g., "2025-08-24" -> ["2025", "08", "24"]
//     const onboardingDateObject = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]); // 建立一個乾淨的日期物件 (當天零點)

//     const supervisorName = formData.supervisor.split(' (')[0];
//     mainSheet.appendRow([
//       employeeId, formData.employeeName, formData.department, supervisorName, '', onboardingDateObject, // 使用修正後的日期物件寫入
//       '', '', '', '', '', '', '', '', formData.salary, '', '', supervisorEmail, '', '', formData.company, 
//       '已寄送Offer', formData.otherSalaryInfo, verificationCode, ''
//     ]);

//     sendOfferEmail(formData, pdfFile, verificationCode);
//     return `成功提交！新進人員 ${formData.employeeName} (員工代號: ${employeeId}) 的錄取通知信已寄出。`;
//   } catch (e) {
//     throw new Error(`後端處理失敗: ${e.message}`);
//   }
// }

/**
 * 從工作表讀取 HTML 範本，填入動態資料後生成 PDF 檔案。
 * @param {Object} formData 表單資料。
 * @param {string} employeeId 員工代號。
 * @returns {GoogleAppsScript.Drive.File} 生成的 PDF 檔案物件。

 function createOfferPdfFromTemplate(formData, employeeId) {
  const ss = SpreadsheetApp.openById(getSetting('SPREADSHEET_ID'));
  const templateContentSheet = ss.getSheetByName('信件範本');
  if (!templateContentSheet) throw new Error("找不到名為 '信件範本' 的工作表。");

  // 使用動態方式從信件範本工作表中尋找 PDF 範本的內容
  const pdfTemplateRow = templateContentSheet.getRange("A2:B").getValues().find(row => row[0] === 'PDF Offer 範本');
  if (!pdfTemplateRow) {
    throw new Error("在 '信件範本' 中找不到 'PDF Offer 範本'。");
  }
  let templateHtml = pdfTemplateRow[1];

  // 使用更穩定的方式處理日期，無論傳入的是字串還是日期物件
  let onboardingDateObject;
  if (typeof formData.onboardingDate === 'string') {
    // 如果是從表單來的 'YYYY-MM-DD' 字串
    const dateParts = formData.onboardingDate.split('-');
    onboardingDateObject = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]);
  } else {
    // 如果是從工作表讀取的 Date 物件
    onboardingDateObject = new Date(formData.onboardingDate);
  }

  const replacements = {
    '{{員工姓名}}': formData.employeeName,
    '{{部門}}': formData.department,
    '{{職稱}}': formData.jobTitle,
    '{{薪資制度}}': `NT$${Number(formData.salary).toLocaleString()}/月，到職當月薪資依實際到職天數比例計算。`,
    '{{報到日期}}': `${onboardingDateObject.toLocaleDateString('zh-TW-u-ca-roc', { year: 'numeric', month: 'long', day: 'numeric' })} 上午 09 時 30 分`
  };
  
  // --- 修正開始 (1/3) ---
  // 修正：處理獎金制度，使用您指定的格式
  if (formData.otherSalaryInfo && formData.otherSalaryInfo.trim() !== '') {
      replacements['{{獎金制度}}'] = `<b>獎 金 制 度 ： </b>${formData.otherSalaryInfo}`;
  } else {
      replacements['{{獎金制度}}'] = '';
  }
  // --- 修正結束 (1/3) ---

  // --- 新增開始 (2/3) ---
  // 新增：產生頁尾的中華民國日期
  const today = new Date();
  const minguoYear = today.getFullYear() - 1911;
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const day = String(today.getDate()).padStart(2, '0');
  const dateString = `中 華 民 國 ${minguoYear} 年 ${month} 月 ${day} 日`;
  // 將帶有樣式的日期 HTML 加入替換物件中，您需要在範本中加入 {{文件日期}} 佔位符
  replacements['{{文件日期}}'] = `<div style="text-align: right; margin-top: 80px; letter-spacing: 4px;">${dateString}</div>`;
  // --- 新增結束 (2/3) ---


  Object.keys(replacements).forEach(key => {
    templateHtml = templateHtml.replace(new RegExp(key, 'g'), replacements[key]);
  });

  // 額外處理：為了讓版面更乾淨，移除替換後可能產生的空段落標籤 (例如 <p></p>)
  templateHtml = templateHtml.replace(/<p>\s*<\/p>/gi, '');

  // --- 修正開始 (3/3) ---
  // 修正：修改 CSS，讓段落文字左右對齊，並增加行高
  const finalHtml = `<html><head><style>body { font-family: 'Arial Unicode MS', sans-serif; font-size: 12pt; } .content { margin: 0 80px; } p { line-height: 1.8; text-align: justify; }</style></head><body><div class="content">${templateHtml}</div></body></html>`;
  // --- 修正結束 (3/3) ---
  
  const pdfBlob = Utilities.newBlob(finalHtml, 'text/html').getAs('application/pdf');
  
  const companyPrefix = (formData.company === '集邦科技') ? '集邦科技' : `集邦_${formData.company}`;
  const fileName = `${companyPrefix}_符合資格通知書-${formData.employeeName}.pdf`;

  const folder = DriveApp.getFolderById(getSetting('PDF_FOLDER_ID'));
  return folder.createFile(pdfBlob).setName(fileName);
}
*/

function createOfferPdfFromTemplate(formData, employeeId) {
  const ss = SpreadsheetApp.openById(getSetting('SPREADSHEET_ID'));
  const templateContentSheet = ss.getSheetByName('信件範本');
  const configSheet = ss.getSheetByName('設定與範本');
  
  if (!templateContentSheet) throw new Error("找不到名為 '信件範本' 的工作表。");
  if (!configSheet) throw new Error("找不到名為 '設定與範本' 的工作表。");

  // 1. 取得內文範本
  const pdfTemplateRow = templateContentSheet.getRange("A2:B").getValues().find(row => row[0] === 'PDF Offer 範本');
  if (!pdfTemplateRow) {
    throw new Error("在 '信件範本' 中找不到 'PDF Offer 範本'。");
  }
  let templateHtml = pdfTemplateRow[1];



  // 2. 【更新】取得所有 Logo 圖片ID
  const configData = configSheet.getDataRange().getValues();
  const getConfigValue = (key) => {
    const row = configData.find(r => r[0] === key);
    return row ? row[1] : '';
  };

  // 【新增】輔助函式：將Drive檔案ID轉換為Base64 Data URL
  const getImageAsBase64Url = (fileId) => {
    if (!fileId) return '';
    try {
      const file = DriveApp.getFileById(fileId);
      const blob = file.getBlob();
      const contentType = blob.getContentType();
      const base64Data = Utilities.base64Encode(blob.getBytes());
      return `data:${contentType};base64,${base64Data}`;
    } catch (e) {
      Logger.log(`無法讀取圖片檔案ID: ${fileId}, 錯誤: ${e.message}`);
      return ''; // 如果檔案找不到或無權限，返回空字串
    }
  };

  const trendforceLogoId = getConfigValue('集邦科技_Logo網址');
  const topologyLogoId = getConfigValue('拓墣科技_Logo網址');
  const topcoLogoId = getConfigValue('新報科技_Logo網址');
  const footerLogosId = getConfigValue('Bottom_Logo網址');

  // 3. 【更新】根據公司別，決定右上角的 Logo
  let headerRightLogoId;
  if (formData.company === '新報科技') {
    headerRightLogoId = topcoLogoId;
  } else {
    headerRightLogoId = topologyLogoId;
  }
  
  // 將圖片ID轉換為可直接使用的URL
  // const trendforceLogoUrl = trendforceLogoId ? `https://drive.google.com/uc?id=${trendforceLogoId}` : '';
  // const headerRightLogoUrl = headerRightLogoId ? `https://drive.google.com/uc?id=${headerRightLogoId}` : '';
  // const footerLogosUrl = footerLogosId ? `https://drive.google.com/uc?id=${footerLogosId}` : '';

     const trendforceLogoUrl = getImageAsBase64Url(trendforceLogoId);
     const headerRightLogoUrl = getImageAsBase64Url(headerRightLogoId);
     const footerLogosUrl = getImageAsBase64Url(footerLogosId);

  // 4. 準備動態資料並替換內文
  let onboardingDateObject;
  if (typeof formData.onboardingDate === 'string') {
    // 如果是從表單來的 'YYYY-MM-DD' 字串
    const dateParts = formData.onboardingDate.split('-');
    onboardingDateObject = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]);
  } else {
    // 如果是從工作表讀取的 Date 物件
    onboardingDateObject = new Date(formData.onboardingDate);
  }
  

  const replacements = {
    '{{員工姓名}}': formData.employeeName,
    '{{部門}}': formData.department,
    '{{職稱}}': formData.jobTitle,
    '{{薪資制度}}': `NT$${Number(formData.salary).toLocaleString()}/月，到職當月薪資依實際到職天數比例計算。`,
    '{{報到日期}}': `${onboardingDateObject.toLocaleDateString('zh-TW-u-ca-roc', { year: 'numeric', month: 'long', day: 'numeric' })} 上午 09 時 30 分`
  };

  if (formData.otherSalaryInfo && formData.otherSalaryInfo.trim() !== '') {
      replacements['{{獎金制度}}'] = `<b>獎 金 制 度 ： </b>${formData.otherSalaryInfo}`;
  } else {
      replacements['{{獎金制度}}'] = '';
  }

  // let bonusText = '';
  // const bonuses = formData.salaryStructure.split('+').slice(1);
  // if (bonuses.length > 0) {
  //   bonusText = `<p><b>獎 金 制 度 ：</b> &nbsp;${bonuses.join('，')}。</p>`;
  // }
  // replacements['{{獎金制度}}'] = bonusText;

  Object.keys(replacements).forEach(key => {
    templateHtml = templateHtml.replace(new RegExp(key, 'g'), replacements[key]);
  });
  
  // 5. 產生中華民國日期
  const today = new Date();
  const minguoYear = today.getFullYear() - 1911;
  const month = today.getMonth() + 1;
  const day = today.getDate();
  const dateString = `中 &nbsp; &nbsp; &nbsp; &nbsp; 華 &nbsp; &nbsp; &nbsp; &nbsp; 民 &nbsp; &nbsp; &nbsp; &nbsp; 國 &nbsp; &nbsp; ${minguoYear} &nbsp; 年 &nbsp; &nbsp; ${month} &nbsp; 月 &nbsp; &nbsp; ${day} &nbsp; &nbsp; 日`;

  // 6. 【更新】組合最終的HTML，包含複雜的頁首排版
  // const finalHtml = `
  //   <html>
  //     <head>
  //       <style>
  //         body { font-family: 'Arial Unicode MS', sans-serif; font-size: 12pt; }
  //         .page-container { width: 210mm; min-height: 297mm; margin: auto; display: flex; flex-direction: column; }
  //         .main-frame { border: 2px solid black; padding: 5mm; flex-grow: 1; display: flex; flex-direction: column; }
  //         .header { display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #ccc; padding-bottom: 10px; }
  //         .header-left img { max-width: 150px; }
  //         .header-right img { max-width: 120px; }
  //         .content { flex-grow: 1; }
  //         .footer { text-align: center; border-top: 1px solid #ccc; padding-top: 8px; margin-top: auto; }
  //         .footer img { max-width: 100%; height: auto; }
  //         p { line-height: 1.8; }
  //         .date-footer { text-align: center; margin-top: 3px; }
  //       </style>
  //     </head>
  //     <body>
  //       <div class="page-container">
  //         <div class="main-frame">
  //           <div class="header">
  //             <div class="header-left">
  //               ${trendforceLogoUrl ? `<img src="${trendforceLogoUrl}">` : ''}
  //             </div>
  //             <div class="header-right">
  //               ${headerRightLogoUrl ? `<img src="${headerRightLogoUrl}">` : ''}
  //             </div>
  //           </div>
  //           <div class="content">
  //             <h1 style="text-align: center; font-size: 20pt; font-weight: bold; margin-top: 20px; margin-bottom: 20px;">符 合 資 格 通 知 書</h1>
  //             ${templateHtml}
  //           </div>
  //           <div class="footer">
  //             ${footerLogosUrl ? `<img src="${footerLogosUrl}">` : ''}
  //           </div>
  //         </div>
  //         <div class="date-footer">
  //           <p>${dateString}</p>
  //         </div>
  //       </div>
  //     </body>
  //   </html>`;
  const finalHtml = `
  <html>
    <head>
      <style>
        body { font-family: 'Arial Unicode MS', sans-serif; font-size: 12pt; }
        .page-container { width: 210mm; min-height: 297mm; margin: 2mm; display: flex; flex-direction: column; }
        .main-frame { border: 2px solid black; padding: 5mm; flex-grow: 1; display: flex; flex-direction: column; }
        .header { display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #ccc; padding-bottom: 10px; }
        .header-left img { max-width: 150px; }
        .header-right img { max-width: 120px; }
        .content { flex-grow: 1; }
        .footer { text-align: center; border-top: 1px solid #ccc; padding-top: 8px; margin-top: auto; }
        .footer img { max-width: 100%; height: auto; }
        p { line-height: 1.8; }
        .date-footer { text-align: center; margin-top: 5px; margin-bottom: 3mm; }
      </style>
    </head>
    <body>
      <div class="page-container">
        <div class="main-frame">
          <div class="header">
            <div class="header-left">
              ${trendforceLogoUrl ? `<img src="${trendforceLogoUrl}">` : ''}
            </div>
            <div class="header-right">
              ${headerRightLogoUrl ? `<img src="${headerRightLogoUrl}">` : ''}
            </div>
          </div>
          <div class="content">
            <h1 style="text-align: center; font-size: 20pt; font-weight: bold; margin-top: 20px; margin-bottom: 20px;">符 合 資 格 通 知 書</h1>
            ${templateHtml}
          </div>
          <div class="footer">
            ${footerLogosUrl ? `<img src="${footerLogosUrl}">` : ''}
            <div class="date-footer">
              <p>${dateString}</p>
            </div>
          </div>
        </div>
      </div>
    </body>
  </html>`;


  const pdfBlob = Utilities.newBlob(finalHtml, 'text/html').getAs('application/pdf');
  const fileName = `符合資格通知書(${formData.company})-${formData.employeeName}.pdf`;
  const folder = DriveApp.getFolderById(getSetting('PDF_FOLDER_ID'));
  return folder.createFile(pdfBlob).setName(fileName);
}

function sendOfferEmail(formData, pdfFile, verificationCode) {
  const ss = SpreadsheetApp.openById(getSetting('SPREADSHEET_ID'));
  const templateSheet = ss.getSheetByName('信件範本');
  if (!templateSheet) throw new Error("找不到名為 '信件範本' 的工作表。");
  
  const bodyTemplateRow = templateSheet.getRange("A2:B").getValues().find(row => row[0] === '錄取通知Email內文');
  if (!bodyTemplateRow) throw new Error("在 '信件範本' 中找不到 '錄取通知Email內文'。");
  
  let bodyTemplate = bodyTemplateRow[1];
  const formUrl = FormApp.openById(getSetting('NEW_HIRE_FORM_ID')).getPublishedUrl();
  const prefilledUrl = `${formUrl}?usp=pp_url&entry.12345=${formData.employeeName}&entry.67890=${verificationCode}`; // 範例 entry ID

  // 替換範本中的動態連結
  bodyTemplate = bodyTemplate.replace('{{個人資料表單連結}}', prefilledUrl);
  
  // 取得簽名檔並加到信件內容後方
  const signature = getGmailSignature(); // 假設此函式存在於您的專案中
  const htmlBodyWithSignature = bodyTemplate + signature;
  
  const subject = (formData.company === '集邦科技') ? '集邦科技_符合資格通知書' : `集邦/${formData.company}_符合資格通知書`;

  // 【更新】準備寄件選項，加入 cc 欄位
  const options = {
    htmlBody: htmlBodyWithSignature,
    attachments: [pdfFile.getAs(MimeType.PDF)],
    name: getSetting('SENDER_NAME')
  };

  // 如果 formData 中有 ccEmails 且長度大於 0，就加入 cc 屬性
  if (formData.ccEmails && formData.ccEmails.length > 0) {
    options.cc = formData.ccEmails.join(',');
    Logger.log(`副本收件人: ${options.cc}`);
  }

  GmailApp.sendEmail(formData.candidateEmail, subject, '', options);
  Logger.log(`已成功寄送錄取通知信至: ${formData.candidateEmail}`);
}

//   // 修正後的 GmailApp.sendEmail 函式呼叫
//   GmailApp.sendEmail(formData.candidateEmail, subject, '', {
//     htmlBody: htmlBodyWithSignature,
//     attachments: [pdfFile.getAs(MimeType.PDF)],
//     name: getSetting('SENDER_NAME')
//   });
// }

// // 請在此處貼上您專案中的 getGmailSignature() 函式，以確保簽名檔功能正常運作。
// // 範例：
// // function getGmailSignature() {
// //   return '您的 HTML 簽名檔內容';
// // }
