const RMA_DATA_SHEET_NAME = "RMA Data";
const LAST_UPDATED_DATE_CELL = "D2";
const NOW = new Date();
const TODAY = new Date(NOW.getFullYear(), NOW.getMonth(), NOW.getDate());
const FIRST_DATA_ROW = 5;
const NUMBER_OF_RECORDS = 100;
const VENDOR_ARRAY = ["VENDOR 1", "VENDOR 2", "VENDOR 3", "VENDOR 4"];

function getRMADataSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RMA_DATA_SHEET_NAME);
}

function generateRMAID(rmaIDset) {
  let rmaID = "";
  
  do {
    rmaID = "RMA-"
    for (let i = 0; i < 5; i++) {
      const digit = Math.floor(Math.random() * 10).toString();
      rmaID += digit;
    }
  } while (rmaIDset.includes(rmaID));

  return rmaID;
}

function generatePartID(totalPartIDPool) {
  let partID = "";
  
  do {
    partID = "PN-";
    for (let i = 0; i < 5; i++) {
      const digit = Math.floor(Math.random() * 10).toString();
      partID += digit;
    }
  } while (totalPartIDPool.includes(partID));

  return partID;
}

function generatePartIDPools(vendorArray) {
  const partIDPools = {};
  const totalPartIDPool = [];

  for (let i = 0; i < vendorArray.length; i++) {
    const partIDPool = [];
    const partIDsAvailable = 5 + Math.floor(Math.random() * 6);

    for (let j = 0; j < partIDsAvailable; j++) {
      const partID = generatePartID(totalPartIDPool);
      partIDPool.push(partID);
      totalPartIDPool.push(partID);
    }

    partIDPools[vendorArray[i]] = partIDPool;
  }

  return partIDPools;
}

function selectPartIDFromPool(vendor, partIDPools) {
  const pool = partIDPools[vendor];
  return pool[Math.floor(Math.random() * pool.length)];
}

function generateSN(serialSet) {
  let serial = "";
  
  do {
    serial = "SN-";
    for (let i = 0; i < 5; i++) {
      const digit = Math.floor(Math.random() * 10).toString();
      serial += digit;
    }
  } while (serialSet.includes(serial));

  return serial;
}

function selectStatus(closedThreshold, inProgressThreshold) {
  const randomPercentage = Math.random();
  let status = "";

  if (randomPercentage < closedThreshold) {
    status = "CLOSED";
  }
  else if (randomPercentage < inProgressThreshold) {
    status = "IN PROGRESS";
  }
  else {
    status = "OPEN";
  }

  return status;
}

function calculateRMAStatus(rmaAge) {
  let rmaStatus = "";
  
  if (rmaAge < 7) {
    rmaStatus = selectStatus(0.10, 0.40);
  }
  else if (rmaAge < 14) {
    rmaStatus = selectStatus(0.40, 0.80);
  }
  else if (rmaAge < 21) {
    rmaStatus = selectStatus(0.70, 0.80);
  }
  else if (rmaAge < 30) {
    rmaStatus = selectStatus(0.80, 0.90);
  }
  else if (rmaAge < 60) {
    rmaStatus = selectStatus(0.90, 0.97);
  }
  else {
    rmaStatus = selectStatus(0.97, 0.98);
  }

  return rmaStatus;
}

function updateLastUpdateData() {
  getRMADataSheet().getRange(LAST_UPDATED_DATE_CELL).setValue(TODAY);
  SpreadsheetApp.flush();
}

function generateData() {
  let dataRows = [];
  let rmaIDSet = [];
  let serialSet = [];
  const partIDPools = generatePartIDPools(VENDOR_ARRAY);

  const lastRow = getRMADataSheet().getLastRow();
  if (lastRow >= FIRST_DATA_ROW) {
    getRMADataSheet().getRange(FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, 8).clearContent();
  }

  for (let i = 0; i < NUMBER_OF_RECORDS; i++) {
    const rmaID = generateRMAID(rmaIDSet);
    rmaIDSet.push(rmaID);

    const vendor = VENDOR_ARRAY[Math.floor(Math.random() * VENDOR_ARRAY.length)];

    const partID = selectPartIDFromPool(vendor, partIDPools);

    const serialID = generateSN(serialSet);
    serialSet.push(serialID);

    const randomDayOffset = Math.floor(Math.random() * 91);
    const creationDate = new Date(TODAY.getFullYear(), TODAY.getMonth(), TODAY.getDate() - randomDayOffset);

    let rmaAge = Math.floor((TODAY.getTime() - creationDate.getTime()) / (1000 * 3600 * 24));
    
    const rmaStatus = calculateRMAStatus(rmaAge);

    let closureDate = "";
    if (rmaStatus == "CLOSED") {
      closureDate = new Date(creationDate.getTime() + Math.floor(Math.random() * (TODAY.getTime() - creationDate.getTime())));
      // Used to remove time of day for aga calculations
      closureDate = new Date(closureDate.getFullYear(), closureDate.getMonth(), closureDate.getDate());
      rmaAge = Math.floor((closureDate.getTime() - creationDate.getTime()) / (1000 * 3600 * 24));
    }

    dataRows.push([rmaID, vendor, partID, serialID, creationDate, rmaStatus, closureDate, rmaAge]);
  }

  if (dataRows.length > 0 && dataRows[0].length > 0) {
    getRMADataSheet().getRange(FIRST_DATA_ROW, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
  }
  updateLastUpdateData();
}

function importData() {
  const httpResponse = UrlFetchApp.fetch("https://raw.githubusercontent.com/avelez61/RMA-Tracker/refs/heads/main/data/rma_data.csv");
  const csvData = httpResponse.getContentText();
  const rmaData = Utilities.parseCsv(csvData);

  const lastRow = getRMADataSheet().getLastRow();
  if (lastRow >= FIRST_DATA_ROW) {
    getRMADataSheet().getRange(FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, 8).clearContent();
  }
  getRMADataSheet().getRange(FIRST_DATA_ROW, 1, rmaData.length - 1, 8).setValues(rmaData.slice(1).map(row => row.slice(0, 8)));

  updateLastUpdateData();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('RMA Tracker')
    .addItem('Generate Data', 'generateData')
    .addItem('Import Data', 'importData')
    .addToUi();
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();

  if (sheet.getName() != "RMA Data") return;
  if (e.range.getColumn() !== 6) return;

  if (e.value == "CLOSED") {
    sheet.getRange(e.range.getRow(), 7).setValue(TODAY);
  }
  else {
    sheet.getRange(e.range.getRow(), 7).clearContent();
  }
}
