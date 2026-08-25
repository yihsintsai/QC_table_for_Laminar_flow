// spreadsheetId 填 google 表單; photoFolderId 填 google 雲端資料夾
const CONFIG = {
  spreadsheetId: '', 
  mappingSheetGid: 0,
  responseSheetName: '表單回應',
  photoFolderId: ''
};

const HEADERS = [
  '填寫工程師', '時間', '中心名稱', '機台', 'QC數值1', 'QC數值2', 'QC數值3', 'QC數值4',
  'QC數值5', '機台照片1', '機台照片2', '機台照片3'
];


// --------------------- 開啟網頁 ------------------------- 
function doGet(e) {
  if (e && e.parameter && e.parameter.page === 'redirect') {
    return HtmlService
      .createHtmlOutputFromFile('Redirect')
      .setTitle('開啟機台QC填寫表')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService
    .createHtmlOutputFromFile('機台QC')
    .setTitle('機台QC填寫表')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/* --------------------- 讀取中心及機台 ------------------------- */
function getFormOptions() {
  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const mappingSheet = getSheetByGid_(ss, CONFIG.mappingSheetGid);

  if (!mappingSheet) {
    throw new Error(
      '找不到 gid=' + CONFIG.mappingSheetGid + ' 的機台對照分頁'
    );
  }

  const values = mappingSheet.getDataRange().getValues();
  const machinesByCenter = {};

  for (let i = 1; i < values.length; i++) {
    const centerName = String(values[i][0] || '').trim();
    const machine = String(values[i][1] || '').trim();

    if (!centerName || !machine) {
      continue;
    }

    if (!machinesByCenter[centerName]) {
      machinesByCenter[centerName] = [];
    }

    if (!machinesByCenter[centerName].includes(machine)) {
      machinesByCenter[centerName].push(machine);
    }
  }

  return {
    centers: Object.keys(machinesByCenter),
    machinesByCenter: machinesByCenter
  };
}


// --------------------- 上傳照片 ------------------------- 
function uploadPhoto( base64Data, originalFileName, mimeType,
                      centerName, machine, photoNumber){
  
  if (!base64Data) {
    return '';
  }

  const folder = DriveApp.getFolderById(CONFIG.photoFolderId);
  const bytes = Utilities.base64Decode(base64Data);
  const extension = getFileExtension_(
    originalFileName,
    mimeType
  );
  const now = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyyMMdd_HHmmss'
  );
  const safeCenter = cleanFileName_(
    centerName || '未填中心'
  );
  const safeMachine = cleanFileName_(
    machine || '未填機台'
  );
  const fileName =
    now + '_' + safeCenter + '_' + safeMachine + '_照片' + photoNumber + extension;
  const blob = Utilities.newBlob(
    bytes,
    mimeType || 'image/jpeg',
    fileName
  );

  const file = folder.createFile(blob);
  
  return file.getUrl();
}


// --------------------- 檢查機台是否填寫過 ------------------------- 
function checkMachineSubmitted(centerName, machine) {
  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const responseSheet =
    ss.getSheetByName(CONFIG.responseSheetName);

  if (
    !centerName ||
    !machine ||
    !responseSheet ||
    responseSheet.getLastRow() < 2
  ) {
    return {
      exists: false
    };
  }

  const values =
    responseSheet.getDataRange().getValues();

  for (let i = values.length - 1; i >= 1; i--) {
    const rowEngineer =
      String(values[i][0] || '').trim();
    const rowTime =
      formatValue_(values[i][1]);
    const rowCenterName =
      String(values[i][2] || '').trim();
    const rowMachine =
      String(values[i][3] || '').trim();
    
    if (
      rowCenterName === centerName &&
      rowMachine === machine
    ) {
      return {
        exists: true,
        engineer: rowEngineer || '未記錄',
        time: rowTime || '未記錄'
      };
    }
  }

  return {
    exists: false
  };
}


// --------------------- 寫入表單 ------------------------- 
function submitForm(formData) {
  const ss =
    SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const engineer =
    String(formData.engineer || '').trim();
  const time =
    String(formData.time || '').trim();
  const centerName =
    String(formData.centerName || '').trim();
  const machine =
    String(formData.machine || '').trim();
  const qc1 =
    String(formData.qc1 || '').trim();
  const qc2 =
    String(formData.qc2 || '').trim();
  const qc3 =
    String(formData.qc3 || '').trim();
  const qc4 =
    String(formData.qc4 || '').trim();
  const qc5 =
    String(formData.qc5 || '').trim();
  const photoUrl1 =
    String(formData.photoUrl1 || '').trim();
  const photoUrl2 =
    String(formData.photoUrl2 || '').trim();
  const photoUrl3 =
    String(formData.photoUrl3 || '').trim()
  const row = [
    engineer, time, centerName, machine, qc1, qc2, qc3, qc4, qc5, photoUrl1, photoUrl2,photoUrl3
  ];


  //---------------------- 寫入總表 ----------------------------

  const responseSheet =
    getOrCreateSheet_(
      ss,
      CONFIG.responseSheetName
    );

  ensureHeaders_(responseSheet);

  responseSheet.appendRow(row);


  // 寫入中心分頁

  const centerSheetName = centerName
    ? cleanSheetName_('中心_' + centerName)
    : '中心_未填中心';

  const centerSheet =
    getOrCreateSheet_(
      ss,
      centerSheetName
    );

  ensureHeaders_(centerSheet);

  centerSheet.appendRow(row);


  SpreadsheetApp.flush();

  return {
    ok: true,
    message:
      '填寫完成！\n' +
      '資料已寫入「' +
      CONFIG.responseSheetName +
      '」和「' +
      centerSheetName +
      '」。'
  };
}


// --------------------- 工具 function ------------------------- 

function getSheetByGid_(ss, gid) {
  const sheets = ss.getSheets();

  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gid) {
      return sheets[i];
    }
  }

  return null;
}


function getOrCreateSheet_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  return sheet;
}


function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet
      .getRange(
        1,
        1,
        1,
        HEADERS.length
      )
      .setValues([HEADERS]);

    sheet.setFrozenRows(1);

    return;
  }

  const firstRow = sheet
    .getRange(
      1,
      1,
      1,
      HEADERS.length
    )
    .getValues()[0];

  const hasHeaders =
    HEADERS.every(
      function(header, index) {
        return (
          String(firstRow[index] || '').trim()
          === header
        );
      }
    );

  if (!hasHeaders) {
    sheet
      .getRange(1, 1, 1, HEADERS.length)
      .setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
}


function cleanSheetName_(name) {
  const cleanedName =
    String(name)
      .replace(/[\\/?*\[\]:]/g, '_')
      .trim()
      .substring(0, 99);

  return cleanedName || '未命名';
}


function cleanFileName_(name) {
  return String(name)
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, '_')
    .trim();
}


function getFileExtension_(fileName, mimeType) {
  const name = String(fileName || '');

  const match =
    name.match(/\.[A-Za-z0-9]+$/);

  if (match) {
    return match[0];
  }

  if (mimeType === 'image/png') {
    return '.png';
  }

  if (mimeType === 'image/heic') {
    return '.heic';
  }

  if (mimeType === 'image/webp') {
    return '.webp';
  }

  return '.jpg';
}


function formatValue_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy/MM/dd HH:mm'
    );
  }

  return String(value || '').trim();
}
