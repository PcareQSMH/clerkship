/**
 * Code.gs
 * สคริปต์ฝั่ง Server สำหรับจัดการข้อมูล Google Sheets และ Google Drive
 */

const FOLDER_ID = "1zLkYfVVvj_-a_z7tV08bRUjq2uPeq5vS"; // โฟลเดอร์สำหรับเก็บรูป
const SPREADSHEET_ID = "1cHGYDJkyIZlKZLJNg13RqPyzmvonog3hkBqRlNhFl7I"; // *** กรุณาใส่ ID ของ Google Sheet ที่คุณสร้างที่นี่ ***

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('ระบบรายงานตัวนิสิตฝึกงาน - รพ.สมเด็จฯ ณ ศรีราชา')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ฟังก์ชันดึงข้อมูลจาก Sheet 'data' และ 'submission'
function getSheetData() {
  const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  
  // ดึงข้อมูล Data หลัก
  const dataSheet = ss.getSheetByName("data");
  let dataValues = [];
  if (dataSheet) {
    // อ่านข้อมูลเริ่มจากแถวที่ 2 (ข้าม Header)
    const lastRow = dataSheet.getLastRow();
    if (lastRow > 1) {
      dataValues = dataSheet.getRange(2, 1, lastRow - 1, 6).getValues(); 
      // Column: 0:ปี, 1:ชั้นปี, 2:มหาลัย, 3:สาขา, 4:ชื่อ-สกุล, 5:ช่วงเวลา
    }
  }

  // ดึงข้อมูล Submission เพื่อเช็คว่าใครส่งแล้วบ้าง
  const subSheet = ss.getSheetByName("submission");
  let submissionList = [];
  if (subSheet) {
    const lastRow = subSheet.getLastRow();
    if (lastRow > 1) {
      // ดึงเลขบัตรประชาชน (col 0) และ ชื่อ-สกุล (col 5) เพื่อใช้เทียบ
      const subValues = subSheet.getRange(2, 1, lastRow - 1, 6).getValues();
      submissionList = subValues.map(row => ({ id: row[0], name: row[5] }));
    }
  }

  return {
    masterData: dataValues,
    submittedData: submissionList
  };
}

// ฟังก์ชันบันทึกข้อมูล
function submitForm(formData) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000); // ป้องกันการบันทึกซ้อนกัน

  try {
    const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("submission");
    
    // ตรวจสอบว่ามี Sheet หรือไม่ ถ้าไม่มีให้สร้างและใส่ Header
    if (!sheet) {
      throw new Error("ไม่พบ Sheet ชื่อ 'submission'");
    }

    // 1. อัปโหลดไฟล์ X-ray
    let xrayFileUrl = "";
    if (formData.xrayFile && formData.xrayFile.data) {
      xrayFileUrl = uploadToDrive(formData.xrayFile.data, formData.xrayFile.name, formData.nationalId + "_XRAY");
    }

    // 2. อัปโหลดไฟล์ Vaccine
    let vaccineFileUrl = "";
    if (formData.vaccineFile && formData.vaccineFile.data) {
      vaccineFileUrl = uploadToDrive(formData.vaccineFile.data, formData.vaccineFile.name, formData.nationalId + "_VAC");
    }

    // 3. บันทึกลง Sheet
    // Order: เลขบัตร, ปี, ชั้นปี, มหาลัย, สาขา, ชื่อ, ช่วงเวลา, วันที่ X-ray, รูป X-ray, วันที่ Vac, รูป Vac, Timestamp
    sheet.appendRow([
      "'" + formData.nationalId, // ใส่ ' เพื่อบังคับเป็น text
      formData.year,
      formData.level,
      formData.uni,
      formData.major,
      formData.name,
      formData.period,
      formData.xrayDate,
      xrayFileUrl,
      formData.vaccineDate,
      vaccineFileUrl,
      new Date()
    ]);

    return { success: true, message: "บันทึกข้อมูลสำเร็จ" };

  } catch (e) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ฟังก์ชันช่วยอัปโหลดไฟล์ (Helper)
function uploadToDrive(base64Data, fileName, prefix) {
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const contentType = base64Data.substring(5, base64Data.indexOf(';'));
    const bytes = Utilities.base64Decode(base64Data.substr(base64Data.indexOf('base64,') + 7));
    const blob = Utilities.newBlob(bytes, contentType, prefix + "_" + fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    return "Upload Failed: " + e.toString();
  }
}
