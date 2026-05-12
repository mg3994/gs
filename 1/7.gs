/**
 * Advanced Backend for Ecommerce, Retail, & Handy Services
 * Features: Slot-based Handy Service Scheduling, Provider Management, 
 * Sheet-based Geo-Fencing (Zip codes), Payment Tracking, Admin Notifications.
 */

const SHEET_NAMES = {
  USERS: 'Users',
  SESSIONS: 'Sessions',
  ORDERS: 'Orders',
  SERVICES: 'Services',
  PAYMENTS: 'Payments',
  ADMINS: 'Admins',
  LOCATIONS: 'Locations',
  PROVIDERS: 'Providers',      // Provider Name, Email, Niche (e.g., Electrician)
  AVAILABILITY: 'Availability' // ProviderEmail, Date, TimeSlot, Status (Free/Busy)
};

const TOKEN_EXPIRY_HOURS = 24;

/**
 * Route Configuration
 */
const ROUTES = {
  PUBLIC: ['login', 'signup', 'forgot_password', 'verify_otp', 'reset_password', 'get_public_catalog', 'check_location', 'get_available_slots'],
  PRIVATE: ['check_auth', 'create_order', 'get_my_orders', 'book_service', 'get_my_services', 'record_payment'],
  ADMIN: ['admin_update_status', 'admin_get_all_records', 'admin_manage_inventory', 'admin_add_location', 'admin_add_provider', 'admin_set_slots']
};

/**
 * API Entry Point
 */
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
      // --- PUBLIC ---
      case 'login': result = secureHandleLogin(data.username, data.password); break;
      case 'signup': result = handleSignup(data.username, data.password); break;
      case 'check_location': result = { success: isZipAllowed(data.zipcode) }; break;
      case 'get_available_slots': 
        result = getAvailableSlots(data.niche, data.date); // date format: YYYY-MM-DD
        break;

      // --- PRIVATE ---
      case 'check_auth': result = { success: true, email: data.userEmail, isAdmin: data.isAdmin }; break;
      case 'create_order': result = createOrder(data); break;
      case 'get_my_orders': result = getRecords(SHEET_NAMES.ORDERS, data.userEmail); break;
      case 'book_service': result = bookService(data); break;
      case 'get_my_services': result = getRecords(SHEET_NAMES.SERVICES, data.userEmail); break;
      case 'record_payment': result = recordPayment(data); break;

      // --- ADMIN ---
      case 'admin_update_status': result = updateStatus(data.type, data.id, data.newStatus); break;
      case 'admin_get_all_records': result = getAllRecordsForAdmin(data.type); break;
      case 'admin_add_provider':
        result = addProvider(data.name, data.email, data.niche);
        break;
      case 'admin_set_slots':
        result = setProviderSlots(data.providerEmail, data.date, data.slots); // slots: array of strings
        break;

      default:
        result = { success: false, message: 'Action not found' };
    }

    return createCORSResponse(JSON.stringify(result));
  } catch (error) {
    return createCORSResponse(JSON.stringify({ success: false, message: error.message }));
  }
}

/**
 * Fetch available slots for a specific niche and date
 */
function getAvailableSlots(niche, date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const providerSheet = getOrInsertSheet(ss, SHEET_NAMES.PROVIDERS, ['Name', 'Email', 'Niche']);
  const availSheet = getOrInsertSheet(ss, SHEET_NAMES.AVAILABILITY, ['Email', 'Date', 'TimeSlot', 'Status']);
  
  // 1. Get providers in this niche
  const providers = providerSheet.getDataRange().getValues()
    .filter(row => row[2].toString().toLowerCase() === niche.toLowerCase())
    .map(row => row[1]); // List of emails

  if (providers.length === 0) return { success: false, message: 'No providers found for this service.' };

  // 2. Get free slots for these providers on the requested date
  const availData = availSheet.getDataRange().getValues();
  const availableSlots = [];

  for (let i = 1; i < availData.length; i++) {
    const [email, rowDate, slot, status] = availData[i];
    // Compare date string
    const formattedDate = Utilities.formatDate(new Date(rowDate), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    
    if (providers.includes(email) && formattedDate === date && status === 'Free') {
      availableSlots.push({
        providerEmail: email,
        slot: slot
      });
    }
  }

  return { success: true, data: availableSlots };
}

/**
 * Private Action: Book Service with Slot Verification
 */
function bookService(data) {
  // 1. Geo-fence Check
  if (!isZipAllowed(data.zipcode)) {
    return { success: false, message: `Zip code ${data.zipcode} is not supported.` };
  }

  // 2. Verify Slot Availability (Prevent Double Booking)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const availSheet = ss.getSheetByName(SHEET_NAMES.AVAILABILITY);
  const availData = availSheet.getDataRange().getValues();
  let slotFound = false;
  let rowIndex = -1;

  for (let i = 1; i < availData.length; i++) {
    const formattedDate = Utilities.formatDate(new Date(availData[i][1]), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    if (availData[i][0] === data.providerEmail && formattedDate === data.date && availData[i][2] === data.slot) {
      if (availData[i][3] === 'Free') {
        slotFound = true;
        rowIndex = i + 1;
        break;
      }
    }
  }

  if (!slotFound) return { success: false, message: 'This time slot is no longer available.' };

  // 3. Mark Slot as Busy
  availSheet.getRange(rowIndex, 4).setValue('Busy');

  // 4. Create Service Record
  const serviceSheet = getOrInsertSheet(ss, SHEET_NAMES.SERVICES, ['ServiceID', 'UserEmail', 'Category', 'ProviderEmail', 'Date', 'Slot', 'Address', 'ZipCode', 'Status']);
  const serviceId = 'SRV-' + Math.floor(1000 + Math.random() * 9000);
  
  serviceSheet.appendRow([
    serviceId, 
    data.userEmail, 
    data.category, 
    data.providerEmail, 
    data.date, 
    data.slot, 
    data.address, 
    data.zipcode, 
    'Confirmed'
  ]);
  
  sendEmailNotification(data.userEmail, 'Service Booked', `Your ${data.category} is scheduled for ${data.date} at ${data.slot}. Provider: ${data.providerEmail}`);
  
  notifyAdmins('New Service Booking', `
    <strong>ID:</strong> ${serviceId}<br>
    <strong>Service:</strong> ${data.category}<br>
    <strong>Provider:</strong> ${data.providerEmail}<br>
    <strong>Time:</strong> ${data.date} @ ${data.slot}
  `);

  return { success: true, serviceId: serviceId };
}

/**
 * Admin: Setup Slots for a Provider
 */
function setProviderSlots(email, date, slots) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.AVAILABILITY, ['Email', 'Date', 'TimeSlot', 'Status']);
  slots.forEach(slot => {
    sheet.appendRow([email, date, slot, 'Free']);
  });
  return { success: true };
}

/**
 * Admin: Add Provider
 */
function addProvider(name, email, niche) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.PROVIDERS, ['Name', 'Email', 'Niche']);
  sheet.appendRow([name, email, niche]);
  return { success: true };
}

/**
 * Utility: Validate Zip Code
 */
function isZipAllowed(zipcode) {
  if (!zipcode) return false;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.LOCATIONS, ['ZipCode', 'City/Area', 'Active']);
  const data = sheet.getDataRange().getValues();
  const cleanZip = zipcode.toString().trim();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === cleanZip && (data[i][2] === true || data[i][2] === 'true' || data[i][2] === 'TRUE')) return true;
  }
  return false;
}

/**
 * Order Creation
 */
function createOrder(data) {
  if (!isZipAllowed(data.zipcode)) return { success: false, message: 'Zip code not supported.' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.ORDERS, ['OrderID', 'Email', 'Items', 'Total', 'Status', 'PaymentStatus', 'Address', 'ZipCode', 'TransactionID', 'Date']);
  const orderId = 'ORD-' + Math.floor(1000 + Math.random() * 9000);
  sheet.appendRow([orderId, data.userEmail, JSON.stringify(data.items), data.total, 'Processing', 'Unpaid', data.address, data.zipcode, '', new Date()]);
  sendEmailNotification(data.userEmail, 'Order Confirmed', `Order ${orderId} received.`);
  notifyAdmins('New Order', `ID: ${orderId}, Total: ${data.total}`);
  return { success: true, orderId: orderId };
}

/**
 * Payment Recording
 */
function recordPayment(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
  if (!sheet) return { success: false };
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === data.orderId) {
      sheet.getRange(i + 1, 6).setValue('Paid');
      sheet.getRange(i + 1, 9).setValue(data.transactionId || 'N/A');
      notifyAdmins('Payment Received', `Order: ${data.orderId}`);
      return { success: true };
    }
  }
  return { success: false };
}

/**
 * Auth & Notifications Helpers
 */
function getAdminEmailList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adminSheet = getOrInsertSheet(ss, SHEET_NAMES.ADMINS, ['Admin Emails']);
  return adminSheet.getDataRange().getValues().slice(1).map(row => row[0].toString().trim()).filter(e => e !== "");
}

function checkIfEmailIsAdmin(email) {
  return getAdminEmailList().some(e => e.toLowerCase() === email.toLowerCase());
}

function secureHandleLogin(username, password) {
  const user = verifyUser(username, password); // Assumes handleSignup-like storage exists
  if (!user) return { success: false, message: 'Invalid credentials' };
  const isAdmin = checkIfEmailIsAdmin(user.email);
  return { success: true, token: createSession(user.email, isAdmin), email: user.email, isAdmin: isAdmin };
}

function validateSession(token) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.SESSIONS, ['Token', 'Email', 'Expiry', 'IsAdmin']);
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      if (now > new Date(data[i][2])) return { success: false, message: 'Expired' };
      return { success: true, email: data[i][1], isAdmin: data[i][3] === true || data[i][3] === 'true' };
    }
  }
  return { success: false };
}

function createSession(email, isAdmin) {
  const sheet = getOrInsertSheet(SpreadsheetApp.getActiveSpreadsheet(), SHEET_NAMES.SESSIONS, ['Token', 'Email', 'Expiry', 'IsAdmin']);
  const token = Utilities.getUuid();
  const expiry = new Date();
  expiry.setHours(expiry.getHours() + TOKEN_EXPIRY_HOURS);
  sheet.appendRow([token, email, expiry, isAdmin]);
  return token;
}

function notifyAdmins(subject, details) {
  const admins = getAdminEmailList();
  if (admins.length === 0) return;
  MailApp.sendEmail({ to: admins.join(','), subject: `[ADMIN] ${subject}`, htmlBody: details, name: "Antinna System", replyTo: "contact@antinna.in" });
}

function sendEmailNotification(to, subject, body) {
  const html = `<div style="font-family:sans-serif;padding:20px;"><h2>Antinna</h2><p>${body}</p></div>`;
  MailApp.sendEmail({ to: to, subject: `[Antinna] ${subject}`, htmlBody: html, name: "Antinna Support", replyTo: "contact@antinna.in" });
}

function getOrInsertSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) { sheet = ss.insertSheet(name); sheet.appendRow(headers); }
  return sheet;
}

function createCORSResponse(payload) {
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON).setHeader('Access-Control-Allow-Origin', '*').setHeader('Access-Control-Allow-Methods', 'POST').setHeader('Access-Control-Allow-Headers', 'Content-Type');
}
