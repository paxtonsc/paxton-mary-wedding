// ── SETUP ──────────────────────────────────────────────────────
// 1. Create a Google Sheet with these tabs: "GuestList", "RSVPs",
//      "RegistryItems", "RegistryFunds"  ("RegistryGifts" is created for you)
// 2. GuestList columns (row 1 = headers):
//      A: GuestID  B: GroupID  C: FirstName  D: LastName  E: Type
//    Type values: primary | plus-one | child
// 3. Replace SHEET_ID below with your sheet's ID (from its URL)
// 4. Deploy: Extensions → Apps Script → Deploy → New deployment
//      Type: Web app · Execute as: Me · Who has access: Anyone
//    Copy the deployment URL and paste into index.html as RSVP_SCRIPT_URL
//    (the registry re-uses this same deployment URL)
//
// RSVPs sheet columns:
//   0=Timestamp  1=RSVPDate  2=Email  3=GroupID  4=GuestID
//   5=FirstName  6=LastName  7=WelcomeDinner  8=Ceremony  9=SundayBrunch
//
// ── REGISTRY SETUP ───────────────────────────────────────────────
// RegistryItems columns (row 1 = headers, you fill in the rows by hand).
// There's no ID column — each item is identified by its spreadsheet row
// number automatically, so just add rows and leave Claimed as FALSE:
//   A=Name  B=Description  C=Link  D=Price
//   E=Claimed(TRUE/FALSE)  F=ClaimedBy  G=ClaimedEmail  H=ClaimedDate
//   I=PhotoUrl (optional — a direct image URL, e.g. a public Google Drive
//     or Photos link, or an image hosted in this repo's images/ folder)
//
// RegistryFunds columns (row 1 = headers, you fill in the rows by hand):
//   A=Name  B=Description  C=Goal (optional, numeric)
//
// RegistryGifts columns (auto-created, like RSVPs):
//   0=Timestamp 1=Method(stripe|venmo) 2=Name 3=Email 4=Amount 5=Fund 6=Message 7=StripeSessionId
//
// Stripe setup — Apps Script web apps can't read the Stripe-Signature header
// that real webhooks need, so instead of a push webhook we confirm each gift
// by pulling it from Stripe's API right after the guest is redirected back:
//   1. Stripe Dashboard → Payment Links → New. Enable "Customer chooses price".
//      Add custom fields with these exact Keys: "fund" (dropdown — set each
//      option's Value to match a RegistryFunds name exactly), "your_name" (text),
//      "message" (text, optional).
//   2. Under "After payment", set a custom confirmation URL back to your site:
//        https://yourdomain.com/?registry_session={CHECKOUT_SESSION_ID}
//   3. In the Apps Script editor: Project Settings → Script properties → add
//      STRIPE_SECRET_KEY = your Stripe secret key (sk_live_... or sk_test_...).
//      Never hardcode this key in this file.
//   4. Paste the Payment Link URL into index.html as STRIPE_PAYMENT_LINK_URL.

const SHEET_ID = '1nsmis0D-yjUk8wR1plM1ZVQIU6Ug5tZXwKSG7cyTsJk';

function doGet(e) {
  const params = e.parameter;
  let result;

  try {
    if (params.action === 'lookup') {
      result = lookup(params.lastName, params.firstInitial);
    } else if (params.action === 'rsvp') {
      result = submitRSVP(params);
    } else if (params.action === 'registry') {
      result = getRegistry();
    } else if (params.action === 'claimItem') {
      result = claimItem(params.itemId, params.name, params.email);
    } else if (params.action === 'logGift') {
      result = logGift(params);
    } else if (params.action === 'confirmStripeGift') {
      result = confirmStripeGift(params.sessionId);
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

// ── REGISTRY ─────────────────────────────────────────────────────

function getRegistry() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // Columns: 0=Name 1=Description 2=Link 3=Price
  //          4=Claimed 5=ClaimedBy 6=ClaimedEmail 7=ClaimedDate 8=PhotoUrl
  // No ID column — each item's spreadsheet row number is its ID.
  const items = [];
  const itemsSheet = ss.getSheetByName('RegistryItems');
  if (itemsSheet && itemsSheet.getLastRow() > 1) {
    const data = itemsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      items.push({
        itemId: String(i + 1), // spreadsheet row number
        name: String(row[0]),
        description: String(row[1]),
        link: String(row[2]),
        price: row[3] === '' ? null : Number(row[3]),
        claimed: String(row[4]).toLowerCase() === 'true',
        photoUrl: String(row[8] || '')
      });
    }
  }

  // Columns: 0=Name 1=Description 2=Goal
  const funds = [];
  const fundsSheet = ss.getSheetByName('RegistryFunds');
  if (fundsSheet && fundsSheet.getLastRow() > 1) {
    const data = fundsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      funds.push({
        name: String(row[0]),
        description: String(row[1]),
        goal: row[2] === '' ? null : Number(row[2]),
        raised: 0
      });
    }
  }

  // Columns: 0=Timestamp 1=Method 2=Name 3=Email 4=Amount 5=Fund 6=Message 7=StripeSessionId
  const giftsSheet = ss.getSheetByName('RegistryGifts');
  if (giftsSheet && giftsSheet.getLastRow() > 1) {
    const data = giftsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const fund = funds.find(f => f.name === String(row[5]));
      if (fund) fund.raised += Number(row[4]) || 0;
    }
  }

  return { items, funds };
}

function claimItem(itemId, name, email) {
  if (!itemId || !name) return { success: false, error: 'Please enter your name.' };

  // itemId is the item's spreadsheet row number (see getRegistry)
  const row = Number(itemId);
  if (!Number.isInteger(row) || row < 2) return { success: false, error: 'Item not found' };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('RegistryItems');
  if (!sheet) return { success: false, error: 'RegistryItems sheet not found' };
  if (row > sheet.getLastRow()) return { success: false, error: 'Item not found' };

  const rowValues = sheet.getRange(row, 1, 1, 5).getValues()[0]; // A..E
  if (!rowValues[0]) return { success: false, error: 'Item not found' };
  if (String(rowValues[4]).toLowerCase() === 'true') {
    return { success: false, error: 'Someone already claimed this gift.' };
  }

  sheet.getRange(row, 5, 1, 4).setValues([[true, name, email || '', new Date().toISOString()]]);
  return { success: true };
}

function appendGift(method, name, email, amount, fund, message, stripeSessionId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('RegistryGifts');
  if (!sheet) {
    sheet = ss.insertSheet('RegistryGifts');
    sheet.appendRow(['Timestamp', 'Method', 'Name', 'Email', 'Amount', 'Fund', 'Message', 'StripeSessionId']);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([new Date().toISOString(), method, name, email || '', amount, fund, message || '', stripeSessionId || '']);
}

function logGift(params) {
  const name = params.name;
  const amount = Number(params.amount);
  const fund = params.fund;

  if (!name || !amount || amount <= 0 || !fund) {
    return { success: false, error: 'Please fill in your name, an amount, and a fund.' };
  }

  appendGift('venmo', name, params.email, amount, fund, params.message);
  return { success: true };
}

function confirmStripeGift(sessionId) {
  if (!sessionId) return { success: false, error: 'Missing session' };

  const ss = SpreadsheetApp.openById(SHEET_ID);

  // Idempotency: don't log the same Stripe session twice (e.g. page refresh)
  const giftsSheet = ss.getSheetByName('RegistryGifts');
  if (giftsSheet && giftsSheet.getLastRow() > 1) {
    const ids = giftsSheet.getRange(2, 8, giftsSheet.getLastRow() - 1, 1).getValues();
    if (ids.some(r => String(r[0]) === sessionId)) {
      return { success: true, alreadyLogged: true };
    }
  }

  const secretKey = PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY');
  if (!secretKey) return { success: false, error: 'Stripe is not configured yet.' };

  const resp = UrlFetchApp.fetch(
    `https://api.stripe.com/v1/checkout/sessions/${encodeURIComponent(sessionId)}`,
    { headers: { Authorization: 'Bearer ' + secretKey }, muteHttpExceptions: true }
  );
  const session = JSON.parse(resp.getContentText());

  if (session.payment_status !== 'paid') {
    return { success: false, error: 'Payment not completed.' };
  }

  const customFields = {};
  (session.custom_fields || []).forEach(f => {
    customFields[f.key] = f.text ? f.text.value : (f.dropdown ? f.dropdown.value : '');
  });

  const amount = (session.amount_total || 0) / 100;
  const email = session.customer_details ? session.customer_details.email : '';
  const name = customFields['your_name'] || (session.customer_details ? session.customer_details.name : '') || 'Guest';
  const fund = customFields['fund'] || 'General Fund';
  const message = customFields['message'] || '';

  appendGift('stripe', name, email, amount, fund, message, sessionId);

  return { success: true, amount, fund, name };
}
