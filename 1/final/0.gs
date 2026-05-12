/**
 * Advanced Backend for Ecommerce, Retail, & Handy Services
 * Features: Slot-based Scheduling, Geo-Fencing, Payment Tracking, Admin Notifications.
 * Optimized for Google Apps Script.
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
  PRIVATE: ['check_auth', 'create_order', 'get_my_orders', 'book_service', 'get_my_services', 'record_payment'],
  ADMIN: ['admin_update_status', 'admin_get_all_records', 'admin_manage_inventory', 'admin_add_location', 'admin_add_provider', 'admin_set_slots']
};

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    const isPublic = ROUTES.PUBLIC.includes(action);
    const isPrivate = ROUTES.PRIVATE.includes(action);
    const isAdmin = ROUTES.ADMIN.includes(action);

    let authContext = { success: true, email: null, isAdmin: false };

    if (isPrivate || isAdmin) {
      authContext = validateSession(data.token);
      if (!authContext.success) return createCORSResponse(JSON.stringify(authContext));
      data.userEmail = authContext.email; 
      data.isAdmin = authContext.isAdmin;
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
      case 'get_available_slots': 
        result = getAvailableSlots(data.niche, data.date); 
        break;
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrInsertSheet(ss, SHEET_NAMES.USERS, ['Email', 'Password', 'Created']);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username) return { success: false, message: 'User already exists' };
    }
    
    sheet.appendRow([username, password, new Date()]);
    return { success: true, message: 'User created' };
  } finally {
    lock.releaseLock();
  }
}

function verifyUser(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.USERS, ['Email', 'Password', 'Created']);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username && data[i][1] === password) {
      return { email: data[i][0] };
    }
  }
  return null;
}

function secureHandleLogin(username, password) {
  const user = verifyUser(username, password);
  if (!user) return { success: false, message: 'Invalid credentials' };
  const isAdmin = checkIfEmailIsAdmin(user.email);
  return { success: true, token: createSession(user.email, isAdmin), email: user.email, isAdmin: isAdmin };
}

function getAvailableSlots(niche, date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const providerSheet = getOrInsertSheet(ss, SHEET_NAMES.PROVIDERS, ['Name', 'Email', 'Niche']);
  const availSheet = getOrInsertSheet(ss, SHEET_NAMES.AVAILABILITY, ['Email', 'Date', 'TimeSlot', 'Status']);
  
  const providers = providerSheet.getDataRange().getValues()
    .filter(row => row[2].toString().toLowerCase() === niche.toLowerCase())
    .map(row => row[1]);

  if (providers.length === 0) return { success: false, message: 'No providers found' };

  const availData = availSheet.getDataRange().getValues();
  const availableSlots = [];

  for (let i = 1; i < availData.length; i++) {
    const [email, rowDate, slot, status] = availData[i];
    const formattedRowDate = Utilities.formatDate(new Date(rowDate), "GMT", "yyyy-MM-dd");
    
    if (providers.includes(email) && formattedRowDate === date && status === 'Free') {
      availableSlots.push({ providerEmail: email, slot: slot });
    }
  }
  return { success: true, data: availableSlots };
}

function bookService(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // 15s wait to prevent race conditions

    if (!isZipAllowed(data.zipcode)) return { success: false, message: 'Unsupported Zip' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const availSheet = ss.getSheetByName(SHEET_NAMES.AVAILABILITY);
    const availData = availSheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < availData.length; i++) {
      const formattedDate = Utilities.formatDate(new Date(availData[i][1]), "GMT", "yyyy-MM-dd");
      if (availData[i][0] === data.providerEmail && formattedDate === data.date && availData[i][2] === data.slot) {
        if (availData[i][3] === 'Free') {
          rowIndex = i + 1;
          break;
        }
      }
    }

    if (rowIndex === -1) return { success: false, message: 'Slot taken or not found' };

    availSheet.getRange(rowIndex, 4).setValue('Busy');

    const serviceSheet = getOrInsertSheet(ss, SHEET_NAMES.SERVICES, ['ServiceID', 'UserEmail', 'Category', 'ProviderEmail', 'Date', 'Slot', 'Address', 'ZipCode', 'Status']);
    const serviceId = 'SRV-' + Math.floor(Math.random() * 10000);
    
    serviceSheet.appendRow([serviceId, data.userEmail, data.category, data.providerEmail, data.date, data.slot, data.address, data.zipcode, 'Confirmed']);
    
    sendEmailNotification(data.userEmail, 'Service Booked', `Booked for ${data.date} @ ${data.slot}`);
    return { success: true, serviceId: serviceId };
  } finally {
    lock.releaseLock();
  }
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
      if (now > new Date(data[i][2])) return { success: false, message: 'Expired' };
      return { success: true, email: data[i][1], isAdmin: String(data[i][3]).toLowerCase() === 'true' };
    }
  }
  return { success: false, message: 'Invalid Session' };
}

function getOrInsertSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) { sheet = ss.insertSheet(name); sheet.appendRow(headers); }
  return sheet;
}

function createCORSResponse(payload) {
  // .setHeader() is removed as it is not a valid function of ContentService in GAS
  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

function sendEmailNotification(to, subject, body) {
  const html = `<div style="font-family:sans-serif;padding:20px;border:1px solid #eee;"><h2>Antinna</h2><p>${body}</p></div>`;
  try {
    MailApp.sendEmail({ to: to, subject: `[Antinna] ${subject}`, htmlBody: html, name: "Antinna" });
  } catch(e) { console.error("Email failed", e); }
}

function checkIfEmailIsAdmin(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.ADMINS, ['Admin Email']);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === email.toLowerCase()) return true;
  }
  return false;
}
