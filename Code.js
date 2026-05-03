// Cozy Stays PG — Google Apps Script Backend (Full Version)
//  Supports: Tenant, Staff, Manager portals + Photo upload
// ═══════════════════════════════════════════════════════════
// SETUP:
// 1. Extensions → Apps Script → paste this
// 2. Set SHEET_ID below
// 3. Run setupSheets() once
// 4. Deploy → New Deployment → Web App
//    Execute as: Me | Who has access: Anyone
// 5. Copy URL → paste in HTML as APPS_SCRIPT_URL
// ═══════════════════════════════════════════════════════════

const SHEET_ID = "1T9ZCGqgH2X9M0rEtMnMdamtZsyBRIDt-FaQhAOAmfpA"; // ← REPLACE THIS

// Sheet column layouts:
// User_Master:  [Property Name, Room Number, Tenant Name, Mobile, PIN]
// Staff_Master: [Staff ID, PIN, Name, Property]
// Requests:     [Property Name, Room Number, Tenant Name, Mobile, Timestamp,
//                Request Type, Comments, Status, Assigned To, Staff Name,
//                Completion Photo (base64), Manager Verified, CSAT Score, CSAT Comment]
// CSAT:         [Building, Room, Tenant, CSAT Score, Comment, Timestamp]

function doGet(e) {
  return handle(e.parameter);
}
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    return handle(body);
  } catch(_) {
    return handle(e.parameter);
  }
}

function handle(p) {
  let result;
  try {
    switch(p.action) {
      case "tenantLogin":      result = tenantLogin(p.mobile, p.pin);                    break;
      case "staffLogin":       result = staffLogin(p.staffId, p.pin);                    break;
      case "submitRequest":    result = submitRequest(p);                                break;
      case "getRequests":      result = getRequests();                                   break;
      case "getTenantRequests":result = getTenantRequests(p.mobile);                     break;
      case "getStaffRequests": result = getStaffRequests(p.staffId);                     break;
      case "assignStaff":      result = assignStaff(p.rowIndex, p.staffId, p.staffName);break;
      case "uploadCompletion": result = uploadCompletion(p.rowIndex, p.staffId, p.photo);break;
      case "verifyAndClose":   result = verifyAndClose(p.rowIndex);                      break;
      case "submitCSAT":       result = submitCSAT(p);                                   break;
      case "getStaffList":     result = getStaffList();                                  break;
      default:                 result = { success: false, message: "Unknown action" };
    }
  } catch(err) {
    result = { success: false, message: err.toString() };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── TENANT LOGIN ─────────────────────────────────────────────
function tenantLogin(mobile, pin) {
  const data = getSheet("User_Master").getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]).trim() === String(mobile).trim() &&
        String(data[i][4]).trim() === String(pin).trim()) {
      return { success: true, tenant: {
        propertyName: data[i][0], roomNumber: String(data[i][1]),
        tenantName: data[i][2],   mobile: String(data[i][3])
      }};
    }
  }
  return { success: false, message: "Invalid mobile or PIN" };
}

// ── STAFF LOGIN ──────────────────────────────────────────────
function staffLogin(staffId, pin) {
  const data = getSheet("Staff_Master").getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === String(staffId).trim().toUpperCase() &&
        String(data[i][1]).trim() === String(pin).trim()) {
      return { success: true, staff: {
        staffId: data[i][0], name: data[i][2], property: data[i][3]
      }};
    }
  }
  return { success: false, message: "Invalid Staff ID or PIN" };
}

// ── SUBMIT REQUEST ───────────────────────────────────────────
function submitRequest(p) {
  getSheet("Requests").appendRow([
    p.propertyName, p.roomNumber, p.tenantName, p.mobile,
    new Date().toISOString(), p.requestType, p.comments,
    "New", "", "", "", false, "", ""
  ]);
  return { success: true };
}

// ── GET ALL REQUESTS (manager) ───────────────────────────────
function getRequests() {
  return { success: true, requests: readRequests() };
}

// ── GET TENANT'S OWN REQUESTS ────────────────────────────────
function getTenantRequests(mobile) {
  const all = readRequests();
  return { success: true, requests: all.filter(r => String(r.mobile) === String(mobile)) };
}

// ── GET STAFF REQUESTS ───────────────────────────────────────
function getStaffRequests(staffId) {
  const all = readRequests();
  return {
    success: true,
    requests: all.filter(r => r.status === "New" || r.assignedTo === staffId)
  };
}

// ── ASSIGN STAFF ─────────────────────────────────────────────
function assignStaff(rowIndex, staffId, staffName) {
  const sheet = getSheet("Requests");
  const row = Number(rowIndex);
  sheet.getRange(row, 8).setValue("Assigned");   // Status
  sheet.getRange(row, 9).setValue(staffId);       // Assigned To
  sheet.getRange(row, 10).setValue(staffName);    // Staff Name
  return { success: true };
}

// ── UPLOAD COMPLETION PHOTO ──────────────────────────────────
function uploadCompletion(rowIndex, staffId, photo) {
  const sheet = getSheet("Requests");
  const row = Number(rowIndex);
  sheet.getRange(row, 8).setValue("Work Done");  // Status
  sheet.getRange(row, 11).setValue(photo);       // Completion Photo (base64)
  return { success: true };
}

// ── MANAGER VERIFY & CLOSE ───────────────────────────────────
function verifyAndClose(rowIndex) {
  const sheet = getSheet("Requests");
  const row = Number(rowIndex);
  sheet.getRange(row, 8).setValue("Closed");    // Status
  sheet.getRange(row, 12).setValue(true);       // Manager Verified
  return { success: true };
}

// ── SUBMIT CSAT ──────────────────────────────────────────────
function submitCSAT(p) {
  // Update Requests sheet
  const sheet = getSheet("Requests");
  const all = sheet.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (i + 1 == p.rowIndex) {
      sheet.getRange(i+1, 13).setValue(p.score);    // CSAT Score
      sheet.getRange(i+1, 14).setValue(p.comment);  // CSAT Comment
      break;
    }
  }
  // Also log in CSAT tab
  const ss2 = SpreadsheetApp.openById(SHEET_ID);
  let csatSheet = ss2.getSheetByName("CSAT");
  if (!csatSheet) {
    csatSheet = ss2.insertSheet("CSAT");
    csatSheet.appendRow(["Building Name","Room Number","Tenant Name","CSAT Score","Comment","Timestamp"]);
    csatSheet.getRange(1,1,1,6).setFontWeight("bold").setBackground("#1A3BAD").setFontColor("#FFFFFF");
  }
  csatSheet.appendRow([p.propertyName, p.roomNumber, p.tenantName, Number(p.score), p.comment||"", new Date().toISOString()]);
  return { success: true };
}

// ── GET STAFF LIST ───────────────────────────────────────────
function getStaffList() {
  const data = getSheet("Staff_Master").getDataRange().getValues();
  const staff = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) staff.push({ staffId: data[i][0], name: data[i][2], property: data[i][3] });
  }
  return { success: true, staff };
}

// ── READ REQUESTS HELPER ─────────────────────────────────────
function readRequests() {
  const data = getSheet("Requests").getDataRange().getValues();
  const reqs = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    reqs.push({
      rowIndex:        i + 1,
      propertyName:    data[i][0],
      roomNumber:      String(data[i][1]),
      tenantName:      data[i][2],
      mobile:          String(data[i][3]),
      timestamp:       data[i][4],
      requestType:     data[i][5],
      comments:        data[i][6],
      status:          data[i][7] || "New",
      assignedTo:      data[i][8] || "",
      staffName:       data[i][9] || "",
      completionPhoto: data[i][10] || "",
      managerVerified: data[i][11] || false,
      csatScore:       data[i][12] || "",
      csatComment:     data[i][13] || "",
    });
  }
  return reqs;
}

// ── GET SHEET ────────────────────────────────────────────────
function getSheet(name) {
  const s = SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
  if (!s) throw new Error(`Sheet "${name}" not found. Run setupSheets() first.`);
  return s;
}

// ── ONE-TIME SETUP ───────────────────────────────────────────
// Run this manually ONCE to create all sheets with headers + sample data
function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const BG = "#1A3BAD", FG = "#FFFFFF";

  function makeSheet(name, headers, sampleRows) {
    let s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    if (s.getLastRow() === 0) {
      s.appendRow(headers);
      s.getRange(1,1,1,headers.length).setFontWeight("bold").setBackground(BG).setFontColor(FG).setFrozenRows(1);
      sampleRows.forEach(r => s.appendRow(r));
    }
    return s;
  }

  makeSheet("User_Master",
    ["Property Name","Room Number","Tenant Name","Mobile Number","Login PIN"],
    [
      ["Sunshine PG","101","Arjun Sharma","9876543210","1234"],
      ["Sunshine PG","203","Priya Mehta","9123456780","2345"],
      ["Green Nest PG","305","Rahul Verma","9988776655","3456"],
    ]
  );

  makeSheet("Staff_Master",
    ["Staff ID","PIN","Name","Property"],
    [
      ["S001","1111","Ramu Kumar","Sunshine PG"],
      ["S002","2222","Sita Devi","Green Nest PG"],
    ]
  );

  makeSheet("Requests",
    ["Property Name","Room Number","Tenant Name","Mobile","Timestamp",
     "Request Type","Comments","Status","Assigned To","Staff Name",
     "Completion Photo","Manager Verified","CSAT Score","CSAT Comment"],
    []
  );

  makeSheet("CSAT",
    ["Building Name","Room Number","Tenant Name","CSAT Score","Comment","Timestamp"],
    []
  );

  SpreadsheetApp.getUi().alert("✅ All sheets created!\n\nNext: Deploy as Web App and copy the URL.");
}
