/**
 * Advanced Backend for Ecommerce, Retail, & Handy Services
 * Optimized for Google Apps Script with robust auth and error handling.
 */

const SHEET_NAMES = {
  USERS: 'Users',
  SESSIONS: 'Sessions',
  ORDERS: 'Orders',
  SERVICES: 'Services',
  PAYMENTS: 'Payments',
  ADMINS: 'Admins',
  LOCATIONS: 'Locations',
  PROVIDERS: 'Providers',
  AVAILABILITY: 'Availability'
};

const TOKEN_EXPIRY_HOURS = 24;

const ROUTES = {
  PUBLIC: ['login', 'signup', 'forgot_password', 'verify_otp', 'reset_password', 'get_public_catalog', 'check_location', 'get_available_slots'],
  PRIVATE: ['check_auth', 'create_order', 'get_my_orders', 'book_service', 'get_my_services', 'record_payment', 'get_booking_details'],
  ADMIN: ['admin_update_status', 'admin_get_all_records', 'admin_manage_inventory', 'admin_add_location', 'admin_add_provider', 'admin_set_slots']
};


/**
 * Mock pricing logic based on service niche
 */
function getPriceForNiche(niche) {
  const rates = {
    'Electrician': 75.00,
    'Plumber': 65.00,
    'Cleaning': 45.00,
    'General': 50.00
  };
  return rates[niche] || rates['General'];
}

function bookService(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.SERVICES, ['BookingID', 'User', 'Provider', 'Niche', 'Date', 'Slot', 'Amount', 'Status']);
  
  const bookingId = 'BK-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  const amount = getPriceForNiche(data.niche);
  
  // Update availability to 'Booked'
  const availSheet = ss.getSheetByName(SHEET_NAMES.AVAILABILITY);
  const availData = availSheet.getDataRange().getValues();
  for (let i = 1; i < availData.length; i++) {
    if (availData[i][0] === data.providerEmail && availData[i][2] === data.slot) {
      availSheet.getRange(i + 1, 4).setValue('Booked');
      break;
    }
  }

  sheet.appendRow([
    bookingId, 
    data.userEmail, 
    data.providerEmail, 
    data.niche, 
    data.date, 
    data.slot, 
    amount, 
    'Pending Payment'
  ]);

  return { 
    success: true, 
    bookingId: bookingId, 
    amount: amount,
    message: `Service booked! Amount due: $${amount}` 
  };
}


function recordPayment(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paySheet = getOrInsertSheet(ss, SHEET_NAMES.PAYMENTS, [
    'Timestamp', 'BookingID', 'User', 'Amount', 'Method', 'TxID', 'Card/UPI Info', 'RawDetails'
  ]);
  const serviceSheet = ss.getSheetByName(SHEET_NAMES.SERVICES);
  
  // Extracting data from the Google Pay response structure provided by the user
  const gPayDetails = data.details || {};
  const payerInfo = data.payer || {};
  const methodData = gPayDetails.paymentMethodData || {};
  
  const bookingId = data.bookingId;
  const amount = data.amount;
  const userEmail = data.userEmail || payerInfo.email || "Unknown";
  
  // Create a readable description for the spreadsheet
  const infoSummary = methodData.description || "N/A";
  const txId = gPayDetails.tokenizationData?.token || 'TX-' + Utilities.getUuid().substring(0, 8);

  // Record the rich payment data
  paySheet.appendRow([
    new Date(),
    bookingId,
    userEmail,
    amount,
    data.method || 'Google Pay',
    txId,
    infoSummary,
    JSON.stringify(gPayDetails) // Store raw JSON for auditing
  ]);

  const services = serviceSheet.getDataRange().getValues();
  for (let i = 1; i < services.length; i++) {
    if (services[i][0] === bookingId) {
      serviceSheet.getRange(i + 1, 9).setValue('Paid'); // Status Column
      break;
    }
  }

  return { 
    success: true, 
    transactionId: txId, 
    message: `Payment of $${amount} verified via Google Pay.` 
  };
}

  return { success: true, transactionId: txId, message: 'Payment recorded and booking confirmed.' };
}

    if (isAdmin && !data.isAdmin) {
      if (!checkIfEmailIsAdmin(data.userEmail)) {
        throw new Error('Unauthorized: Admin database verification failed');
      }
    }

    let result;
    switch (action) {
      case 'login': result = secureHandleLogin(data.username, data.password); break;
      case 'signup': result = handleSignup(data.username, data.password); break;
      case 'check_location': result = { success: isZipAllowed(data.zipcode) }; break;
      case 'get_available_slots': result = getAvailableSlots(data.niche, data.date); break;
      case 'check_auth': result = { success: true, email: data.userEmail, isAdmin: data.isAdmin }; break;
      case 'create_order': result = createOrder(data); break;
      case 'get_my_orders': result = getRecords(SHEET_NAMES.ORDERS, data.userEmail); break;
      case 'book_service': result = bookService(data); break;
      case 'get_my_services': result = getRecords(SHEET_NAMES.SERVICES, data.userEmail); break;
      case 'record_payment': result = recordPayment(data); break;
      case 'admin_update_status': result = updateStatus(data.type, data.id, data.newStatus); break;
      case 'admin_get_all_records': result = getAllRecordsForAdmin(data.type); break;
      case 'admin_add_provider': result = addProvider(data.name, data.email, data.niche); break;
      case 'admin_set_slots': result = setProviderSlots(data.providerEmail, data.date, data.slots); break;
      default: result = { success: false, message: 'Action not found' };
    }

    return createCORSResponse(JSON.stringify(result));
  } catch (error) {
    return createCORSResponse(JSON.stringify({ success: false, message: error.message }));
  }
}

function handleSignup(username, password) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!username || !password) return { success: false, message: 'Missing username or password' };
    
    const cleanUser = username.toString().trim().toLowerCase();
    const cleanPass = password.toString().trim();
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrInsertSheet(ss, SHEET_NAMES.USERS, ['Email', 'Password', 'Created']);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === cleanUser) {
        return { success: false, message: 'User already exists' };
      }
    }
    
    sheet.appendRow([cleanUser, cleanPass, new Date()]);
    return { success: true, message: 'User created successfully' };
  } finally {
    lock.releaseLock();
  }
}

function verifyUser(username, password) {
  if (!username || !password) return null;
  
  const cleanUser = username.toString().trim().toLowerCase();
  const cleanPass = password.toString().trim();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.USERS, ['Email', 'Password', 'Created']);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const sheetEmail = data[i][0].toString().trim().toLowerCase();
    const sheetPass = data[i][1].toString().trim();
    
    if (sheetEmail === cleanUser && sheetPass === cleanPass) {
      return { email: sheetEmail };
    }
  }
  return null;
}

function secureHandleLogin(username, password) {
  const user = verifyUser(username, password);
  if (!user) {
    return { success: false, message: 'Invalid credentials. Check email casing or password.' };
  }
  const isAdmin = checkIfEmailIsAdmin(user.email);
  const token = createSession(user.email, isAdmin);
  return { success: true, token: token, email: user.email, isAdmin: isAdmin };
}

function getAvailableSlots(niche, date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const providerSheet = getOrInsertSheet(ss, SHEET_NAMES.PROVIDERS, ['Name', 'Email', 'Niche']);
  const availSheet = getOrInsertSheet(ss, SHEET_NAMES.AVAILABILITY, ['Email', 'Date', 'TimeSlot', 'Status']);
  
  const providers = providerSheet.getDataRange().getValues()
    .filter(row => row[2].toString().toLowerCase() === niche.toLowerCase())
    .map(row => row[1].toString().toLowerCase());

  if (providers.length === 0) return { success: false, message: 'No providers found for this niche' };

  const availData = availSheet.getDataRange().getValues();
  const availableSlots = [];

  for (let i = 1; i < availData.length; i++) {
    const email = availData[i][0].toString().toLowerCase();
    const rowDate = availData[i][1];
    const slot = availData[i][2];
    const status = availData[i][3];

    const formattedRowDate = Utilities.formatDate(new Date(rowDate), "GMT", "yyyy-MM-dd");
    
    if (providers.includes(email) && formattedRowDate === date && status === 'Free') {
      availableSlots.push({ providerEmail: email, slot: slot });
    }
  }
  return { success: true, data: availableSlots };
}

function createSession(email, isAdmin) {
  const sheet = getOrInsertSheet(SpreadsheetApp.getActiveSpreadsheet(), SHEET_NAMES.SESSIONS, ['Token', 'Email', 'Expiry', 'IsAdmin']);
  const token = Utilities.getUuid();
  const expiry = new Date();
  expiry.setHours(expiry.getHours() + TOKEN_EXPIRY_HOURS);
  sheet.appendRow([token, email, expiry, isAdmin]);
  return token;
}

function validateSession(token) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.SESSIONS, ['Token', 'Email', 'Expiry', 'IsAdmin']);
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      if (now > new Date(data[i][2])) return { success: false, message: 'Session expired' };
      return { success: true, email: data[i][1], isAdmin: String(data[i][3]).toLowerCase() === 'true' };
    }
  }
  return { success: false, message: 'Invalid Session' };
}

function checkIfEmailIsAdmin(email) {
  const cleanEmail = email.toString().trim().toLowerCase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.ADMINS, ['Admin Email']);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim().toLowerCase() === cleanEmail) return true;
  }
  return false;
}

function addProvider(name, email, niche) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.PROVIDERS, ['Name', 'Email', 'Niche']);
  sheet.appendRow([name, email, niche]);
  return { success: true, message: `Provider ${name} added.` };
}

function setProviderSlots(providerEmail, date, slots) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.AVAILABILITY, ['Email', 'Date', 'TimeSlot', 'Status']);
  slots.forEach(slot => {
    sheet.appendRow([providerEmail, date, slot, 'Free']);
  });
  return { success: true, message: `Slots created for ${date}` };
}

function getOrInsertSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) { sheet = ss.insertSheet(name); sheet.appendRow(headers); }
  return sheet;
}

function createCORSResponse(payload) {
  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

function isZipAllowed(zipcode) {
  if (!zipcode) return false;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.LOCATIONS, ['ZipCode', 'City/Area', 'Active']);
  const data = sheet.getDataRange().getValues();
  const cleanZip = zipcode.toString().trim();
  for (let i = 1; i < data.length; i++) {
    const isActive = String(data[i][2]).toLowerCase() === 'true';
    if (data[i][0].toString().trim() === cleanZip && isActive) return true;
  }
  return false;
}
