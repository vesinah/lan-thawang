/**
 * ===================================================
 * ลานท่าวัง - Google Apps Script Backend API
 * ===================================================
 * ระบบลงทะเบียนคัมภีร์ใบลาน
 * 
 * วิธีใช้:
 * 1. เปิด Google Apps Script Editor (script.google.com)
 * 2. สร้างโปรเจกต์ใหม่
 * 3. Copy โค้ดนี้ไปวางทั้งหมด
 * 4. Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy URL ที่ได้ไปใส่ในไฟล์ app.js (ตัวแปร API_URL)
 */

// ===== CONFIG =====
const SHEET1_ID = '1RyEZqlL7u0oJ3UuT4ncc4VsIgMsx6o4eflhZbN28h_g'; // ทะเบียนใบลานวัดวังตะวันตก
const SHEET2_ID = '141QC6tHvIDnDk8fAp8A9qPRDBsa8_3TVURsUV6KZ0vA'; // ฐานข้อมูลคัมภีร์ใบลานเมืองนคร
const SHEET1_NAME = 'ชีต1';
const SHEET2_NAME = 'เรียงตามรหัส';

// ===== COLUMN MAPPING สำหรับ Sheet 1 =====
const SHEET1_COLUMNS = [
  'ที่', 'เลขมัด', 'รหัสคัมภีร์', 'ชื่อเรื่องบนคัมภีร์', 'หมวดคัมภีร์',
  'เรื่อง', 'จำนวนผูก', 'ผูกที่มี', 'ยุคสมัย', 'อักษร', 'ภาษา',
  'เส้น', 'ฉบับ', 'ชนิดไม้ประกับ', 'ฉลากคัมภีร์', 'ผ้าห่อคัมภีร์',
  'ขนาดคัมภีร์', 'ขนาดไม้ประกับ', 'ตู้ที่', 'ชั้นที่',
  'ปีที่สร้าง', 'ชื่อผู้สร้าง', 'ประวัติ/บริบท', 'สภาพคัมภีร์',
  'Digitize', 'หมายเหตุสภาพ', 'หมายเหตุ', 'ประวัติการจัดการ', 'ผู้บันทึก'
];

// ===== COLUMN MAPPING สำหรับ Sheet 2 =====
const SHEET2_COLUMNS = [
  'ที่', 'ที่เก็บรักษา', 'เลขมัด', 'รหัสคัมภีร์', 'ชื่อเรื่อง',
  'หมวด', 'จำนวนเรื่อง', 'จำนวนผูก', 'ผูกที่มี',
  'ยุคสมัย', 'อักษร', 'ภาษา', 'เส้น', 'ฉบับ',
  'ชนิดไม้ประกับ', 'หมายเหตุ'
];

// ===== DROPDOWN VALUES =====
const DROPDOWN_OPTIONS = {
  หมวดคัมภีร์: [
    'พระวินัยปิฎก', 'พระสุตตันตปิฎก', 'พระอภิธรรมปิฎก',
    'ปกรณ์วิเสส', 'คัมภีร์ประเพณี', 'วรรณกรรมท้องถิ่น',
    'คัมภีร์บรรณรักษ์', 'ยังไม่ระบุหมวด', 'อื่นๆ'
  ],
  ยุคสมัย: [
    'อยุธยา', 'รัตนโกสินทร์', 'ยังไม่ระบุสมัย'
  ],
  อักษร: [
    'ขอม', 'ไทย', 'ขอม-ไทย', 'อื่นๆ'
  ],
  ภาษา: [
    'บาลี', 'ไทย', 'บาลี-ไทย', 'อื่นๆ'
  ],
  เส้น: [
    'จาร', 'เขียน', 'จาร-เขียน', 'อื่นๆ'
  ],
  ฉบับ: [
    'ลานยาว', 'ลานสั้น', 'อื่นๆ'
  ],
  ชนิดไม้ประกับ: [
    'ไม้ธรรมดา', 'ไม้แกะสลัก', 'ไม้ลงรักปิดทอง', 'ไม้ลงรัก',
    'ไม้ประดับกระจก', 'ไม่มีไม้ประกับ', 'อื่นๆ'
  ],
  สภาพคัมภีร์: [
    'ดี', 'พอใช้', 'ชำรุดเล็กน้อย', 'ชำรุดมาก', 'เสียหายหนัก'
  ],
  Digitize: [
    'แล้ว', 'ยังไม่ได้', 'อยู่ระหว่างดำเนินการ'
  ]
};

// ===== CORS HEADERS =====
function createCorsOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== GET HANDLER =====
function doGet(e) {
  try {
    const action = e.parameter.action;

    switch (action) {
      case 'getSheet1Data':
        return createCorsOutput({ success: true, data: getSheetData(SHEET1_ID, SHEET1_NAME, SHEET1_COLUMNS) });
      
      case 'getSheet2Data':
        return createCorsOutput({ success: true, data: getSheetData(SHEET2_ID, SHEET2_NAME, SHEET2_COLUMNS) });
      
      case 'getAllData':
        return createCorsOutput({
          success: true,
          data: {
            sheet1: getSheetData(SHEET1_ID, SHEET1_NAME, SHEET1_COLUMNS),
            sheet2: getSheetData(SHEET2_ID, SHEET2_NAME, SHEET2_COLUMNS)
          }
        });

      case 'getDashboard':
        return createCorsOutput({ success: true, data: getDashboardData() });

      case 'getDropdowns':
        return createCorsOutput({ success: true, data: DROPDOWN_OPTIONS });

      default:
        return createCorsOutput({ success: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return createCorsOutput({ success: false, error: err.toString() });
  }
}

// ===== POST HANDLER =====
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    switch (action) {
      case 'addRecord':
        return createCorsOutput({ success: true, data: addRecord(body.record) });
      
      case 'updateRecord':
        return createCorsOutput({ success: true, data: updateRecord(body.rowIndex, body.record) });

      case 'deleteRecord':
        return createCorsOutput({ success: true, data: deleteRecord(body.rowIndex) });

      default:
        return createCorsOutput({ success: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return createCorsOutput({ success: false, error: err.toString() });
  }
}

// ===== DATA FUNCTIONS =====

/**
 * อ่านข้อมูลจาก Sheet แล้วแปลงเป็น array ของ objects
 */
function getSheetData(sheetId, sheetName, columns) {
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return []; // มีแค่ header

  const numCols = columns.length;
  const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  
  return data.map((row, index) => {
    const obj = { _rowIndex: index + 2 }; // เก็บ row index จริงใน sheet (1-indexed, +1 สำหรับ header)
    columns.forEach((col, i) => {
      obj[col] = row[i] !== undefined && row[i] !== null ? row[i].toString() : '';
    });
    return obj;
  }).filter(obj => {
    // กรองแถวที่ว่างทั้งหมดออก
    return columns.some(col => obj[col] && obj[col].toString().trim() !== '');
  });
}

/**
 * สรุปข้อมูลสำหรับแดชบอร์ด
 */
function getDashboardData() {
  const sheet1Data = getSheetData(SHEET1_ID, SHEET1_NAME, SHEET1_COLUMNS);
  const sheet2Data = getSheetData(SHEET2_ID, SHEET2_NAME, SHEET2_COLUMNS);

  // สรุปจำนวน
  const totalSheet1 = sheet1Data.length;
  const totalSheet2 = sheet2Data.length;
  const totalAll = totalSheet1 + totalSheet2;

  // นับจำนวนที่ digitize แล้ว (Sheet 1)
  const digitized = sheet1Data.filter(r => r['Digitize'] === 'แล้ว').length;

  // สรุปตามหมวด
  const categoryCount = {};
  sheet1Data.forEach(r => {
    const cat = r['หมวดคัมภีร์'] || 'ไม่ระบุ';
    categoryCount[cat] = (categoryCount[cat] || 0) + 1;
  });
  sheet2Data.forEach(r => {
    const cat = r['หมวด'] || 'ไม่ระบุ';
    categoryCount[cat] = (categoryCount[cat] || 0) + 1;
  });

  // สรุปตามยุคสมัย
  const eraCount = {};
  [...sheet1Data, ...sheet2Data].forEach(r => {
    const era = r['ยุคสมัย'] || 'ไม่ระบุ';
    eraCount[era] = (eraCount[era] || 0) + 1;
  });

  // สรุปตามอักษร
  const scriptCount = {};
  [...sheet1Data, ...sheet2Data].forEach(r => {
    const script = r['อักษร'] || 'ไม่ระบุ';
    scriptCount[script] = (scriptCount[script] || 0) + 1;
  });

  // สรุปตามภาษา
  const languageCount = {};
  [...sheet1Data, ...sheet2Data].forEach(r => {
    const lang = r['ภาษา'] || 'ไม่ระบุ';
    languageCount[lang] = (languageCount[lang] || 0) + 1;
  });

  // สรุปตามที่เก็บรักษา (Sheet 2)
  const locationCount = {};
  sheet2Data.forEach(r => {
    const loc = r['ที่เก็บรักษา'] || 'ไม่ระบุ';
    locationCount[loc] = (locationCount[loc] || 0) + 1;
  });
  // เพิ่มวัดวังตะวันตกจาก Sheet 1
  if (totalSheet1 > 0) {
    locationCount['วัดวังตะวันตก'] = totalSheet1;
  }

  // สรุปตามสภาพ (Sheet 1)
  const conditionCount = {};
  sheet1Data.forEach(r => {
    const cond = r['สภาพคัมภีร์'] || 'ไม่ระบุ';
    conditionCount[cond] = (conditionCount[cond] || 0) + 1;
  });

  return {
    summary: {
      totalAll,
      totalSheet1,
      totalSheet2,
      digitized
    },
    categoryCount,
    eraCount,
    scriptCount,
    languageCount,
    locationCount,
    conditionCount
  };
}

/**
 * เพิ่มรายการใหม่ลง Sheet 1
 */
function addRecord(record) {
  const ss = SpreadsheetApp.openById(SHEET1_ID);
  const sheet = ss.getSheetByName(SHEET1_NAME);
  const lastRow = sheet.getLastRow();
  
  // auto-number: ลำดับถัดไป
  const nextNum = lastRow <= 1 ? 1 : lastRow; // row 1 = header, row 2 = data start

  const rowData = SHEET1_COLUMNS.map((col, i) => {
    if (col === 'ที่') return nextNum;
    return record[col] || '';
  });

  sheet.appendRow(rowData);
  
  return { 
    message: 'เพิ่มรายการสำเร็จ', 
    rowIndex: lastRow + 1,
    recordNumber: nextNum 
  };
}

/**
 * แก้ไขรายการใน Sheet 1
 */
function updateRecord(rowIndex, record) {
  const ss = SpreadsheetApp.openById(SHEET1_ID);
  const sheet = ss.getSheetByName(SHEET1_NAME);
  
  const rowData = SHEET1_COLUMNS.map(col => record[col] || '');
  
  sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  
  return { message: 'แก้ไขรายการสำเร็จ', rowIndex };
}

/**
 * ลบรายการจาก Sheet 1
 */
function deleteRecord(rowIndex) {
  const ss = SpreadsheetApp.openById(SHEET1_ID);
  const sheet = ss.getSheetByName(SHEET1_NAME);
  
  sheet.deleteRow(rowIndex);
  
  return { message: 'ลบรายการสำเร็จ' };
}
