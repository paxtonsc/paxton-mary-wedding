// ── SETUP ──────────────────────────────────────────────────────
// 1. Create a Google Sheet with these tabs: "GuestList", "RSVPs", "Registry"
// 2. GuestList columns (row 1 = headers):
//      A: GuestID  B: GroupID  C: FirstName  D: LastName  E: Type
//    Type values: primary | plus-one | child
// 3. Registry sheet columns (row 1 = headers):
//      A: ID  B: Name  C: Description  D: Price  E: Store  F: URL
//      G: Purchased  H: PurchasedBy
//    Purchased values: TRUE / FALSE (or blank = not purchased)
// 4. Replace SHEET_ID below with your sheet's ID (from its URL)
// 5. Deploy: Extensions → Apps Script → Deploy → New deployment
//      Type: Web app · Execute as: Me · Who has access: Anyone
//    Copy the deployment URL and paste into index.html as RSVP_SCRIPT_URL
//
// RSVPs sheet columns:
//   0=Timestamp  1=RSVPDate  2=Email  3=GroupID  4=GuestID
//   5=FirstName  6=LastName  7=WelcomeDinner  8=Ceremony  9=SundayBrunch

const SHEET_ID = '1nsmis0D-yjUk8wR1plM1ZVQIU6Ug5tZXwKSG7cyTsJk';

function doGet(e) {
  const params = e.parameter;
  let result;

  try {
    if (params.action === 'lookup') {
      result = lookup(params.lastName, params.firstInitial);
    } else if (params.action === 'rsvp') {
      result = submitRSVP(params);
    } else if (params.action === 'registryList') {
      result = getRegistryList();
    } else if (params.action === 'markPurchased') {
      result = markPurchased(params.itemId, params.purchasedBy);
    } else {
      result = { error: 'Unknown action' };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function lookup(lastName, firstInitial) {
  if (!lastName || !firstInitial) {
    return { found: false, error: 'Missing parameters' };
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('GuestList');
  if (!sheet) return { found: false, error: 'GuestList sheet not found' };

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { found: false };

  // Columns: 0=GuestID  1=GroupID  2=FirstName  3=LastName  4=Type
  const lnLower = lastName.trim().toLowerCase();
  const fi = firstInitial.trim().toLowerCase().charAt(0);

  const matchingGroupIds = new Set();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowFirstName = String(row[2]).trim();
    if (!rowFirstName || rowFirstName.toLowerCase() === 'tbd') continue;
    const rowLN = String(row[3]).trim().toLowerCase();
    const rowFI = rowFirstName.toLowerCase().charAt(0);
    if (rowLN === lnLower && rowFI === fi) {
      matchingGroupIds.add(String(row[1]));
    }
  }

  if (matchingGroupIds.size === 0) return { found: false };

  const groups = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const gid = String(row[1]);
    if (matchingGroupIds.has(gid)) {
      if (!groups[gid]) groups[gid] = { groupId: gid, members: [] };
      const firstName = String(row[2]).trim();
      groups[gid].members.push({
        guestId: String(row[0]),
        firstName: firstName,
        lastName: String(row[3]),
        unnamed: firstName === '' || firstName.toLowerCase() === 'tbd'
      });
    }
  }

  // Fetch latest RSVP status per guest from RSVPs sheet
  // Columns: 0=Timestamp 1=RSVPDate 2=Email 3=GroupID 4=GuestID
  //          5=FirstName 6=LastName 7=WelcomeDinner 8=Ceremony 9=SundayBrunch
  const existingRsvps = {};
  const rsvpSheet = ss.getSheetByName('RSVPs');
  if (rsvpSheet && rsvpSheet.getLastRow() > 1) {
    const rsvpData = rsvpSheet.getDataRange().getValues();
    for (let i = 1; i < rsvpData.length; i++) {
      const r = rsvpData[i];
      const gid = String(r[4]);
      existingRsvps[gid] = {
        email: String(r[2]),
        rsvpDate: String(r[1]),
        firstName: String(r[5]),
        lastName: String(r[6]),
        welcomeDinner: String(r[7]).toLowerCase() === 'yes',
        ceremony: String(r[8]).toLowerCase() === 'yes',
        sundayBrunch: String(r[9]).toLowerCase() === 'yes'
      };
    }
  }

  return { found: true, groups: Object.values(groups), existingRsvps };
}

function submitRSVP(params) {
  const email = params.email;
  const groupId = params.groupId;
  const rsvpDate = params.rsvpDate;
  const rsvpsJson = params.rsvps;

  if (!email || !groupId || !rsvpsJson) {
    return { success: false, error: 'Missing required fields' };
  }

  let rsvps;
  try {
    rsvps = JSON.parse(rsvpsJson);
  } catch (e) {
    return { success: false, error: 'Invalid RSVP data' };
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('RSVPs');
  if (!sheet) {
    sheet = ss.insertSheet('RSVPs');
    sheet.appendRow([
      'Timestamp', 'RSVPDate', 'Email', 'GroupID', 'GuestID',
      'FirstName', 'LastName', 'WelcomeDinner', 'Ceremony', 'SundayBrunch'
    ]);
    sheet.setFrozenRows(1);
  }

  const timestamp = new Date().toISOString();
  const today = rsvpDate || timestamp.split('T')[0];

  // Build a map of guestId -> row index (1-based) from existing data
  const existingRows = {};
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const guestIdCol = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
    guestIdCol.forEach((r, i) => {
      existingRows[String(r[0])] = i + 2; // +2: 1-based + header row
    });
  }

  rsvps.forEach(r => {
    const rowData = [
      timestamp,
      today,
      email,
      groupId,
      r.guestId,
      r.firstName,
      r.lastName,
      r.welcomeDinner ? 'Yes' : 'No',
      r.ceremony     ? 'Yes' : 'No',
      r.sundayBrunch ? 'Yes' : 'No'
    ];
    if (existingRows[r.guestId]) {
      sheet.getRange(existingRows[r.guestId], 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
  });

  return { success: true };
}

function getRegistryList() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Registry');
  if (!sheet) return { items: [] };

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { items: [] };

  // Columns: 0=ID  1=Name  2=Description  3=Price  4=Store  5=URL  6=Purchased  7=PurchasedBy
  const items = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0] || !r[1]) continue;
    items.push({
      id:          String(r[0]),
      name:        String(r[1]),
      description: String(r[2] || ''),
      price:       String(r[3] || ''),
      store:       String(r[4] || ''),
      url:         String(r[5] || ''),
      purchased:   r[6] === true || String(r[6]).toLowerCase() === 'true',
      purchasedBy: String(r[7] || '')
    });
  }
  return { items };
}

function markPurchased(itemId, purchasedBy) {
  if (!itemId || !purchasedBy) {
    return { success: false, error: 'Missing required fields' };
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Registry');
  if (!sheet) return { success: false, error: 'Registry sheet not found' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(itemId)) {
      if (data[i][6] === true || String(data[i][6]).toLowerCase() === 'true') {
        return { success: false, error: 'This item has already been claimed.' };
      }
      sheet.getRange(i + 1, 7).setValue(true);
      sheet.getRange(i + 1, 8).setValue(purchasedBy);
      return { success: true };
    }
  }
  return { success: false, error: 'Item not found' };
}
