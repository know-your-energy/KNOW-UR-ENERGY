function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('KNOW YOUR ENERGY');
}

// รับข้อมูลจาก GitHub Pages
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const result = saveCalculatorData(data);

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: error.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// บันทึกข้อมูลลง Google Sheets เดิม
function saveCalculatorData(data) {

  if (
    !data ||
    !Array.isArray(data.devices) ||
    data.devices.length === 0
  ) {
    throw new Error('Invalid device data');
  }

  const days = Number(data.days);
  const carbon = Number(data.carbon);
  const cost = Number(data.cost);

  if (
    !Number.isInteger(days) || days <= 0 ||
    !Number.isFinite(carbon) || carbon < 0 ||
    !Number.isFinite(cost) || cost < 0
  ) {
    throw new Error('Invalid calculation data');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error('Spreadsheet not found');
  }

  let sheet = ss.getSheetByName('Calculator Data');

  if (!sheet) {
    sheet = ss.insertSheet('Calculator Data');
  }

  // สร้างหัวตารางเฉพาะเมื่อชีตยังว่าง
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp',
      'Device',
      'Power (W)',
      'Quantity',
      'Hours Saved/Day',
      'Days',
      'Energy Saved (kWh)',
      'CO2 Reduction (kgCO2e)',
      'Cost Saving (THB)'
    ]);
  }

  const timestamp = new Date();

  const rows = data.devices.map(device => {

    const watt = Number(device.watt);
    const quantity = Number(device.quantity);
    const hours = Number(device.hours);
    const energy = Number(device.energy);

    if (
      !device.name ||
      !Number.isFinite(watt) || watt <= 0 ||
      !Number.isInteger(quantity) || quantity <= 0 ||
      !Number.isFinite(hours) || hours <= 0 ||
      !Number.isFinite(energy) || energy < 0
    ) {
      throw new Error('Invalid device information');
    }

    return [
      timestamp,
      String(device.name),
      watt,
      quantity,
      hours,
      days,
      energy,
      carbon,
      cost
    ];
  });

  // ป้องกันการเขียนข้อมูลชนกัน
  const lock = LockService.getScriptLock();

  lock.waitLock(10000);

  try {
    sheet.getRange(
      sheet.getLastRow() + 1,
      1,
      rows.length,
      rows[0].length
    ).setValues(rows);

    SpreadsheetApp.flush();

  } finally {
    lock.releaseLock();
  }

  return {
    success: true,
    message: 'บันทึกข้อมูลเรียบร้อยแล้ว'
  };
}
