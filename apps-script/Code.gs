// ── SETUP ──────────────────────────────────────────────────────
// 1. Create a Google Sheet with these two tabs: "GuestList", "RSVPs"
// 2. GuestList columns (row 1 = headers):
//      A: GuestID  B: GroupID  C: FirstName  D: LastName  E: Type
//    Type values: primary | plus-one | child
// 3. Replace SHEET_ID below with your sheet's ID (from its URL)
// 4. Deploy: Extensions → Apps Script → Deploy → New deployment
//      Type: Web app · Execute as: Me · Who has access: Anyone
//    Copy the deployment URL and paste into index.html as RSVP_SCRIPT_URL

const SHEET_ID = '1nsmis0D-yjUk8wR1plM1ZVQIU6Ug5tZXwKSG7cyTsJk';

function doGet(e) {
  const params = e.parameter;
  let result;

  try {
    if (params.action === 'lookup') {
      result = lookup(params.lastName, params.firstInitial);
    } else if (params.action === 'rsvp') {
      result = submitRSVP(params);
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
    const rowLN = String(row[3]).trim().toLowerCase();
    const rowFI = String(row[2]).trim().toLowerCase().charAt(0);
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
        type: String(row[4]),
        unnamed: firstName === '' || firstName.toLowerCase() === 'tbd'
      });
    }
  }

  // Fetch latest RSVP status per guest from RSVPs sheet
  const existingRsvps = {};
  const rsvpSheet = ss.getSheetByName('RSVPs');
  if (rsvpSheet && rsvpSheet.getLastRow() > 1) {
    const rsvpData = rsvpSheet.getDataRange().getValues();
    // Columns: 0=Timestamp 1=RSVPDate 2=Email 3=GroupID 4=GuestID 5=FirstName 6=LastName 7=Attending
    // Iterating forward means the last write for each guestId wins (most recent)
    for (let i = 1; i < rsvpData.length; i++) {
      const r = rsvpData[i];
      const gid = String(r[4]);
      existingRsvps[gid] = {
        attending: String(r[7]).toLowerCase() === 'yes',
        email: String(r[2]),
        rsvpDate: String(r[1]),
        firstName: String(r[5]),
        lastName: String(r[6])
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
    sheet.appendRow(['Timestamp', 'RSVPDate', 'Email', 'GroupID', 'GuestID', 'FirstName', 'LastName', 'Attending']);
    sheet.setFrozenRows(1);
  }

  const timestamp = new Date().toISOString();

  rsvps.forEach(r => {
    sheet.appendRow([
      timestamp,
      rsvpDate || timestamp.split('T')[0],
      email,
      groupId,
      r.guestId,
      r.firstName,
      r.lastName,
      r.attending ? 'Yes' : 'No'
    ]);
  });

  return { success: true };
}
