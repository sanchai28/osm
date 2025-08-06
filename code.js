const VHV_SHEET_NAME = "อสม. Data";
const USER_SHEET_NAME = "Users";
const REPORT_SHEET_NAME = "osm1";
const UNIT_SHEET_NAME = "unit";
const PDF_SHEET_NAME = "pdf";
const DRIVE_FOLDER_ID = "1QpGglkAIcneqwb4DqnjmU_949JL56ROi";
const MEETING_SHEET_NAME = "Meetings";
const MEETING_HEADERS = ['id', 'date', 'topic', 'summary', 'attendeeIds', 'facilityId', 'createdBy'];

// --- Column Header Constants ---
const VHV_HEADERS = [
  'หมายเลขบัตรประชาชน', 'ชื่อ-นามสกุล', 'วันเกิด', 'ที่อยู่', 'หมู่ที่',
  'ตำบล', 'อำเภอ', 'จังหวัด', 'เลขสถานบริการ',
  'วันที่ขึ้นทะเบียนเป็น อสม.', 'เบอร์โทรศัพท์', 'บัญชีปฎิบัติงาน อสม.', 'สถานะ อสม.', 'วันที่สถานะ'
];

const USER_HEADERS = [
  'หมายเลขบัตรประชาชน', 'ชื่อ-นามสกุล', 'ตำแหน่ง', 'เลขสถานบริการ', 'สถานะ'
];

const REPORT_HEADERS = [
    'ID', 'หมายเลขบัตร อสม', 'วันที่ส่ง', 'รูปแบบการส่งรายงาน', 'การเบิกจ่ายค่าตอบแทน', 'หมายเลขบัตรเจ้าหน้าที่', 'เลขสถานบริการ'
];


// --- Main POST Handler ---
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ status: 'error', message: 'No post data received from client.' });
    }
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    if (!action) {
      return createJsonResponse({ status: 'error', message: 'Action not specified in payload.' });
    }
    switch (action) {
      case 'login': return handleLogin(payload);
      case 'create':
      case 'update':
      case 'delete': return handleCrud(payload);
      case 'saveMonthlyReport': return handleSaveMonthlyReport(payload);
      case 'uploadPdf': return handleUploadPdf(payload);
      default:
        return createJsonResponse({ status: 'error', message: `Invalid action specified: '${action}'` });
    }
  } catch (error) {
    Logger.log(`doPost Error: ${error.stack}`);
    return createJsonResponse({ status: 'error', message: error.message });
  }
}

// --- Main GET Handler ---
function doGet(e) {
  try {
    const action = e.parameter.action;
    switch(action) {
        case 'getVhvList': return getVhvList(e);
        case 'getReportSummary': return getReportSummary(e);
        case 'getMonthlyReport': return getMonthlyReport(e);
        case 'getDashboardData': return getDashboardData(e);
        
        default: return getVhvData(e);
    }
  } catch (error) {
    Logger.log(`doGet Error: ${error.stack}`);
    return createJsonResponse({ error: error.message });
  }
}

// --- Action Handlers ---
function handleUploadPdf(payload) {
    const { facilityId, month, year, fileData } = payload;
    const pdfSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PDF_SHEET_NAME);
    if (!pdfSheet) return createJsonResponse({ status: 'error', message: `Sheet "${PDF_SHEET_NAME}" not found.` });

    const uniqueId = `${facilityId}-${year}-${month}`;
    const newFileName = `${uniqueId}.pdf`;

    const allData = pdfSheet.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][0]).trim() == uniqueId) {
            foundRow = i + 1;
            const oldFileId = allData[i][1];
            if (oldFileId) {
                try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (err) {}
            }
            break;
        }
    }

    const decodedData = Utilities.base64Decode(fileData);
    const blob = Utilities.newBlob(decodedData, MimeType.PDF, newFileName);
    
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const file = folder.createFile(blob);
    const fileId = file.getId();
    const fileUrl = file.getUrl();

    const newRowData = [uniqueId, fileId, newFileName, fileUrl];
    if (foundRow > -1) {
        pdfSheet.getRange(foundRow, 1, 1, newRowData.length).setValues([newRowData]);
    } else {
        pdfSheet.appendRow(newRowData);
    }

    return createJsonResponse({ status: 'success', fileUrl: fileUrl });
}

function getReportSummary(e) {
    const { facilityId, fiscalYear } = e.parameter;
    const reportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REPORT_SHEET_NAME);
    const pdfSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PDF_SHEET_NAME);
    if (!reportSheet || !pdfSheet) return createJsonResponse({});

    const pdfData = pdfSheet.getDataRange().getValues();
    const pdfMap = {};
    for (let i = 1; i < pdfData.length; i++) {
        const id = String(pdfData[i][0]).trim();
        const fileUrl = String(pdfData[i][3]).trim();
        const parts = id.split('-');
        if (parts.length < 3) continue;
        const pdfFacilityId = parts[0];
        if (pdfFacilityId == facilityId) {
            pdfMap[id] = fileUrl;
        }
    }

    const data = reportSheet.getDataRange().getValues();
    if (data.length <= 1) return createJsonResponse({});

    const summary = {};
    const headers = data[0].map(h => h.trim());
    const idIndex = headers.indexOf('ID');
    const paymentIndex = headers.indexOf('การเบิกจ่ายค่าตอบแทน');
    const facilityIndex = headers.indexOf('เลขสถานบริการ');

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowFacilityId = String(row[facilityIndex]).trim();

        if (rowFacilityId == facilityId) {
            const id = String(row[idIndex]).trim();
            const parts = id.split('-');
            if (parts.length < 4) continue;

            const month = parseInt(parts[parts.length - 1], 10);
            const year = parseInt(parts[parts.length - 2], 10);

            if (isNaN(year) || isNaN(month)) continue;

            let rowFiscalYear = (month >= 10) ? year + 1 : year;

            if (rowFiscalYear == fiscalYear) {
                const monthKey = month;
                if (!summary[monthKey]) {
                    summary[monthKey] = { paid: 0, suspended: 0, hasPdf: false, fileUrl: null };
                }
                const paymentStatus = String(row[paymentIndex]).trim();
                if (paymentStatus === 'เบิกจ่าย') summary[monthKey].paid++;
                if (paymentStatus === 'ระงับจ่าย') summary[monthKey].suspended++;
                
                const pdfCheckId = `${facilityId}-${year}-${month}`;
                if (pdfMap[pdfCheckId]) {
                    summary[monthKey].hasPdf = true;
                    summary[monthKey].fileUrl = pdfMap[pdfCheckId];
                }
            }
        }
    }
    return createJsonResponse(summary);
}

function getDashboardData(e) {
    const { fiscalYear } = e.parameter;
    const unitSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNIT_SHEET_NAME);
    const reportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REPORT_SHEET_NAME);
    const pdfSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PDF_SHEET_NAME);

    if (!unitSheet || !reportSheet || !pdfSheet) {
        return createJsonResponse({ units: [], summary: {} });
    }

    const unitData = unitSheet.getDataRange().getValues();
    const units = [];
    for (let i = 1; i < unitData.length; i++) {
        const unitId = String(unitData[i][0]).trim();
        if (unitId && unitId !== '00126') {
            units.push({ id: unitId, name: unitData[i][1] });
        }
    }

    const pdfData = pdfSheet.getDataRange().getValues();
    const pdfMap = {};
    for (let i = 1; i < pdfData.length; i++) {
        const id = String(pdfData[i][0]).trim();
        const fileUrl = String(pdfData[i][3]).trim();
        pdfMap[id] = fileUrl;
    }

    const reportData = reportSheet.getDataRange().getValues();
    const summary = {};

    const headers = reportData[0].map(h => h.trim());
    const idIndex = headers.indexOf('ID');
    const paymentIndex = headers.indexOf('การเบิกจ่ายค่าตอบแทน');

    for (let i = 1; i < reportData.length; i++) {
        const row = reportData[i];
        const id = String(row[idIndex]).trim();
        const parts = id.split('-');
        if (parts.length < 4) continue;

        const facilityId = parts[0];
        const year = parseInt(parts[parts.length - 2], 10);
        const month = parseInt(parts[parts.length - 1], 10);

        if (isNaN(year) || isNaN(month)) continue;

        let rowFiscalYear = (month >= 10) ? year + 1 : year;

        if (rowFiscalYear == fiscalYear) {
            const monthKey = month;
            const unitIdKey = facilityId;

            if (!summary[monthKey]) summary[monthKey] = {};
            if (!summary[monthKey][unitIdKey]) {
                summary[monthKey][unitIdKey] = { paid: 0, suspended: 0, fileUrl: null };
            }

            const paymentStatus = String(row[paymentIndex]).trim();
            if (paymentStatus === 'เบิกจ่าย') summary[monthKey][unitIdKey].paid++;
            if (paymentStatus === 'ระงับจ่าย') summary[monthKey][unitIdKey].suspended++;
            
            const pdfCheckId = `${facilityId}-${year}-${month}`;
            if (pdfMap[pdfCheckId]) {
                summary[monthKey][unitIdKey].fileUrl = pdfMap[pdfCheckId];
            }
        }
    }

    return createJsonResponse({ units: units, summary: summary });
}

function handleLogin(payload) {
  const { idCard, password } = payload;
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET_NAME);
  const unitSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNIT_SHEET_NAME);
  
  if (!userSheet) return createJsonResponse({ status: 'error', message: 'User sheet not found.' });
  if (!unitSheet) return createJsonResponse({ status: 'error', message: 'Unit sheet not found.' });

  const userData = userSheet.getDataRange().getValues();
  const unitData = unitSheet.getDataRange().getValues();
  
  const userHeaders = userData[0].map(h => h.trim());
  const idCardIndex = userHeaders.indexOf('หมายเลขบัตรประชาชน');
  const passwordIndex = userHeaders.indexOf('เลขสถานบริการ');
  const statusIndex = userHeaders.indexOf('สถานะ');

  for (let i = 1; i < userData.length; i++) {
    const row = userData[i];
    if (String(row[idCardIndex]).trim() == idCard) {
      if (String(row[passwordIndex]).trim() == password) {
        if (String(row[statusIndex]).trim() === 'ใช้งาน') {
          
          const facilityId = String(row[passwordIndex]).trim();
          const facilityName = getFacilityNameById(unitData, facilityId);
          
          if (!facilityName) {
            return createJsonResponse({ status: 'error', message: 'ไม่พบข้อมูลสถานบริการ' });
          }

          const user = {
            idCard: String(row[idCardIndex]).trim(),
            fullName: row[userHeaders.indexOf('ชื่อ-นามสกุล')],
            position: row[userHeaders.indexOf('ตำแหน่ง')],
            facilityId: facilityId,
            facilityName: facilityName
          };
          return createJsonResponse({ status: 'success', user });
        } else {
          return createJsonResponse({ status: 'error', message: 'บัญชีผู้ใช้นี้ถูกระงับ' });
        }
      } else {
        return createJsonResponse({ status: 'error', message: 'รหัสผ่านไม่ถูกต้อง' });
      }
    }
  }
  return createJsonResponse({ status: 'error', message: 'ไม่พบผู้ใช้งานนี้ในระบบ' });
}

function handleCrud(payload) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VHV_SHEET_NAME);
    if (!sheet) return createJsonResponse({ status: 'error', message: `Sheet "${VHV_SHEET_NAME}" not found.` });
    
    switch (payload.action) {
        case 'create':
            sheet.appendRow(mapDataToRowArray(payload.data, VHV_HEADERS));
            return createJsonResponse({ status: 'success' });
        case 'update':
            const updateRange = sheet.getRange(payload.rowIndex, 1, 1, VHV_HEADERS.length);
            updateRange.setValues([mapDataToRowArray(payload.data, VHV_HEADERS)]);
            return createJsonResponse({ status: 'success' });
        case 'delete':
            sheet.deleteRow(payload.rowIndex);
            return createJsonResponse({ status: 'success' });
    }
}

function handleSaveMonthlyReport(payload) {
    const { reports, officerIdCard, facilityId, month, year } = payload;
    const reportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REPORT_SHEET_NAME);
    if (!reportSheet) return createJsonResponse({ status: 'error', message: `Sheet "${REPORT_SHEET_NAME}" not found.` });

    const today = new Date();
    const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd/MM/") + (today.getFullYear() + 543);
    
    const allData = reportSheet.getDataRange().getValues();
    const headers = allData[0].map(h => h.trim());
    const idIndex = headers.indexOf('ID');

    const existingIdMap = {};
    for (let i = 1; i < allData.length; i++) {
        const id = String(allData[i][idIndex]).trim();
        if (id) {
            existingIdMap[id] = i;
        }
    }

    const rowsToAdd = [];

    reports.forEach(report => {
        const vhvIdCard = report.vhvIdCard;
        const uniqueId = `${facilityId}-${vhvIdCard}-${year}-${month}`;
        const newRowData = [
            uniqueId, vhvIdCard, formattedDate, report.submissionMethod,
            report.paymentStatus, officerIdCard, facilityId
        ];

        if (existingIdMap[uniqueId] !== undefined) {
            const arrayIndex = existingIdMap[uniqueId];
            allData[arrayIndex] = newRowData;
        } else {
            rowsToAdd.push(newRowData);
        }
    });

    if (allData.length > 1) {
        reportSheet.getRange(1, 1, allData.length, headers.length).setValues(allData);
    }
    
    if (rowsToAdd.length > 0) {
        reportSheet.getRange(reportSheet.getLastRow() + 1, 1, rowsToAdd.length, REPORT_HEADERS.length).setValues(rowsToAdd);
    }
    
    return createJsonResponse({ status: 'success' });
}


function getVhvData(e) {
    const facilityId = e.parameter.facilityId;
    if (!facilityId) return createJsonResponse([]);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VHV_SHEET_NAME);
    if (!sheet) return createJsonResponse([]);
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    if (values.length <= 1) return createJsonResponse([]);
    
    const sheetHeaders = values[0].map(h => h.trim());
    const facilityColIndex = sheetHeaders.indexOf('เลขสถานบริการ');
    if (facilityColIndex === -1) return createJsonResponse([]);

    const dateHeaders = ['วันเกิด', 'วันที่ขึ้นทะเบียนเป็น อสม.', 'วันที่ขึ้นทะเบียน', 'วันที่สถานะ'];

    const data = values.slice(1)
      .map((row, index) => ({ row, rowIndex: index + 2 }))
      .filter(({ row }) => String(row[facilityColIndex]).trim() == facilityId)
      .map(({ row, rowIndex }) => {
        const rowData = {};
        sheetHeaders.forEach((header, i) => {
          const key = toCamelCase(header);
          if (key) {
            let value = row[i];
            if (dateHeaders.includes(header) && value instanceof Date) {
              rowData[key] = formatToBuddhistDate(value);
            } else {
              rowData[key] = value;
            }
          }
        });
        return { rowIndex, data: rowData };
      });
    
    return createJsonResponse(data);
}

function getVhvList(e) {
    const facilityId = e.parameter.facilityId;
    if (!facilityId) return createJsonResponse([]);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VHV_SHEET_NAME);
    if (!sheet) return createJsonResponse([]);
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    if (values.length <= 1) return createJsonResponse([]);
    
    const sheetHeaders = values[0].map(h => h.trim());
    const facilityColIndex = sheetHeaders.indexOf('เลขสถานบริการ');
    const idCardIndex = sheetHeaders.indexOf('หมายเลขบัตรประชาชน');
    const fullNameIndex = sheetHeaders.indexOf('ชื่อ-นามสกุล');
    const statusIndex = sheetHeaders.indexOf('สถานะ อสม.');
    const workAccountIndex = sheetHeaders.indexOf('บัญชีปฎิบัติงาน อสม.');
    const mooIndex = sheetHeaders.indexOf('หมู่ที่');

    if ([facilityColIndex, idCardIndex, fullNameIndex, statusIndex, workAccountIndex, mooIndex].includes(-1)) {
       return createJsonResponse([]);
    }

    const data = values.slice(1)
      .filter(row => 
          String(row[facilityColIndex]).trim() == facilityId && 
          String(row[statusIndex]).trim() === 'เป็น อสม.' &&
          row[workAccountIndex] == 1
      )
      .map(row => ({
          data: {
              idCard: row[idCardIndex],
              fullName: row[fullNameIndex],
              moo: row[mooIndex]
          }
      }));
    
    return createJsonResponse(data);
}

function getMonthlyReport(e, returnMap = false) {
    const { facilityId, month, year } = e.parameter;
    const reportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REPORT_SHEET_NAME);
    if (!reportSheet) return returnMap ? {} : createJsonResponse([]);

    const data = reportSheet.getDataRange().getValues();
    if (data.length <= 1) return returnMap ? {} : createJsonResponse([]);

    const headers = data[0].map(h => h.trim());
    const idIndex = headers.indexOf('ID');
    const results = [];
    const resultMap = {};

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row && row[idIndex]) {
            const id = String(row[idIndex]);
            const parts = id.split('-');
            if (parts.length < 4) continue;

            const rowFacilityId = parts[0];
            const rowYear = parseInt(parts[parts.length - 2], 10);
            const rowMonth = parseInt(parts[parts.length - 1], 10);

            if (rowFacilityId == facilityId && rowMonth == month && rowYear == year) {
                const entry = {};
                headers.forEach((header, index) => {
                    entry[header] = row[index];
                });
                results.push(entry);
                resultMap[id] = { rowIndex: i + 1, data: entry };
            }
        }
    }
    return returnMap ? resultMap : createJsonResponse(results);
}


// --- Helper Functions ---
function getFacilityNameById(unitData, facilityId) {
    const headers = unitData[0].map(h => h.trim());
    const idIndex = headers.indexOf('เลขสถานบริการ');
    const nameIndex = headers.indexOf('สถานบริการ');
    
    for (let i = 1; i < unitData.length; i++) {
        if (String(unitData[i][idIndex]).trim() == facilityId) {
            return unitData[i][nameIndex];
        }
    }
    return null; // Not found
}

function mapDataToRowArray(dataObject, headers) {
  const camelCaseMap = getCamelCaseMap();
  const headerToCamelCase = {};
  for (const key in camelCaseMap) {
    headerToCamelCase[camelCaseMap[key]] = key;
  }

  return headers.map(header => {
    const key = headerToCamelCase[header];
    return dataObject[key] || '';
  });
}

function formatToBuddhistDate(jsDate) {
  if (!jsDate || !(jsDate instanceof Date)) return '';
  let day = jsDate.getDate();
  let month = jsDate.getMonth() + 1;
  let year = jsDate.getFullYear();
  if (year < 2500) year += 543;
  return `${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}/${year}`;
}

function createJsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

function getCamelCaseMap() {
    return {
        idCard: 'หมายเลขบัตรประชาชน', fullName: 'ชื่อ-นามสกุล', dob: 'วันเกิด',
        address: 'ที่อยู่', moo: 'หมู่ที่', subDistrict: 'ตำบล', district: 'อำเภอ',
        province: 'จังหวัด', 
        facilityId: 'เลขสถานบริการ',
        regDate: 'วันที่ขึ้นทะเบียนเป็น อสม.', phone: 'เบอร์โทรศัพท์',
        workAccount: 'บัญชีปฎิบัติงาน อสม.', status: 'สถานะ อสม.', statusDate: 'วันที่สถานะ',
        position: 'ตำแหน่ง', facilityName: 'สถานบริการ'
    };
}

function toCamelCase(header) {
  const map = getCamelCaseMap();
  const trimmedHeader = header.trim();
  for (const key in map) {
    if (map[key] === trimmedHeader) {
      return key;
    }
  }
  return null;
}

// --- Meeting Management Functions ---
// แทนที่ฟังก์ชันเดิมทั้งหมดด้วยชุดนี้


function handleGetMeetings(data) {
  try {
    const meetingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEETING_SHEET_NAME);
    if (!meetingSheet) throw new Error(`Sheet "${MEETING_SHEET_NAME}" not found.`);
    
    const meetingData = getSheetData(meetingSheet, MEETING_HEADERS);
    const vhvData = getSheetData(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VHV_SHEET_NAME));

    // Create a map for quick VHV lookup by ID card
    const vhvMap = new Map(vhvData.map(vhv => [vhv.idCard, vhv.fullName]));

    const facilityMeetings = meetingData
      .filter(m => m.facilityId === data.facilityId)
      .map(meeting => {
        let attendees = [];
        if (meeting.attendeeIds) {
          try {
            const attendeeIdArray = JSON.parse(meeting.attendeeIds);
            attendees = attendeeIdArray.map(id => ({
              idCard: id,
              fullName: vhvMap.get(id) || 'ไม่พบชื่อ (อสม. อาจถูกลบ)'
            }));
          } catch (e) {
            // Handle cases where parsing fails
            console.error(`Could not parse attendeeIds for meeting ${meeting.id}: ${meeting.attendeeIds}`);
          }
        }
        return {
          ...meeting,
          attendees: attendees // Return full attendee objects
        };
      });

    return createJsonResponse({ status: 'success', data: facilityMeetings });
  } catch (error) {
    return createJsonResponse({ status: 'error', message: 'Error fetching meetings: ' + error.message });
  }
}

function handleAddMeeting(data) {
  try {
    const meetingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEETING_SHEET_NAME);
    if (!meetingSheet) throw new Error(`Sheet "${MEETING_SHEET_NAME}" not found.`);
    
    const newId = 'M' + new Date().getTime();
    const newRow = [
      newId,
      new Date(data.date),
      data.topic,
      data.summary,
      JSON.stringify(data.attendees || []), // Store attendee IDs as a JSON string
      data.facilityId,
      data.userId // Assuming userId is passed from frontend
    ];
    
    meetingSheet.appendRow(newRow);
    return createJsonResponse({ status: 'success', message: 'Meeting added successfully.', id: newId });
  } catch (error) {
    return createJsonResponse({ status: 'error', message: 'Error adding meeting: ' + error.message });
  }
}

function handleUpdateMeeting(data) {
  try {
    const meetingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEETING_SHEET_NAME);
    if (!meetingSheet) throw new Error(`Sheet "${MEETING_SHEET_NAME}" not found.`);

    const dataRange = meetingSheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];
    const idColIndex = headers.indexOf('id');
    
    if (idColIndex === -1) throw new Error("Column 'id' not found in Meetings sheet.");

    const rowIndex = values.findIndex(row => row[idColIndex] == data.id);

    if (rowIndex === -1) {
      return createJsonResponse({ status: 'error', message: 'Meeting not found.' });
    }

    const rowData = {
      id: data.id,
      date: new Date(data.date),
      topic: data.topic,
      summary: data.summary,
      attendeeIds: JSON.stringify(data.attendees || []),
      facilityId: data.facilityId,
      createdBy: values[rowIndex][headers.indexOf('createdBy')] // Keep original creator
    };

    const newRowValues = mapObjectToRow(rowData, headers, MEETING_HEADERS);
    meetingSheet.getRange(rowIndex + 1, 1, 1, newRowValues.length).setValues([newRowValues]);

    return createJsonResponse({ status: 'success', message: 'Meeting updated successfully.' });
  } catch (error) {
    return createJsonResponse({ status: 'error', message: 'Error updating meeting: ' + error.message });
  }
}

function handleDeleteMeeting(data) {
  try {
    const meetingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEETING_SHEET_NAME);
    if (!meetingSheet) throw new Error(`Sheet "${MEETING_SHEET_NAME}" not found.`);

    const dataRange = meetingSheet.getDataRange();
    const values = dataRange.getValues();
    const idColIndex = values[0].indexOf('id');

    if (idColIndex === -1) throw new Error("Column 'id' not found in Meetings sheet.");

    const rowIndex = values.findIndex(row => row[idColIndex] == data.id);

    if (rowIndex > 0) { // rowIndex > 0 to avoid deleting header
      meetingSheet.deleteRow(rowIndex + 1);
      return createJsonResponse({ status: 'success', message: 'Meeting deleted successfully.' });
    } else {
      return createJsonResponse({ status: 'error', message: 'Meeting not found.' });
    }
  } catch (error) {
    return createJsonResponse({ status: 'error', message: 'Error deleting meeting: ' + error.message });
  }
}