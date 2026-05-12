/**
 * Advanced Backend for Ecommerce, Retail, & Handy Services
 * Features: Sheet-based Geo-Fencing (Zip codes), Payment Tracking, Admin Notifications, Token-based Auth
 */

const SHEET_NAMES = {
  USERS: 'Users',
  SESSIONS: 'Sessions',
  ORDERS: 'Orders',
  SERVICES: 'Services',
  PAYMENTS: 'Payments',
  ADMINS: 'Admins',
  LOCATIONS: 'Locations' // New sheet for Zip code based geo-fencing
};

const TOKEN_EXPIRY_HOURS = 24;

/**
 * Route Configuration
 */
const ROUTES = {
  PUBLIC: ['login', 'signup', 'forgot_password', 'verify_otp', 'reset_password', 'get_public_catalog', 'check_location'],
  PRIVATE: ['check_auth', 'create_order', 'get_my_orders', 'book_service', 'get_my_services', 'record_payment'],
  ADMIN: ['admin_update_status', 'admin_get_all_records', 'admin_manage_inventory', 'admin_add_location']
};

/**
 * Enhanced API Entry Point
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
      case 'login': result = secureHandleLogin(data.username, data.password); break;
      case 'signup': result = handleSignup(data.username, data.password); break;
      case 'check_auth': result = { success: true, email: data.userEmail, isAdmin: data.isAdmin }; break;
      
      case 'check_location':
        result = { success: isZipAllowed(data.zipcode) };
        break;

      case 'create_order':
        result = createOrder(data);
        break;
      case 'get_my_orders':
        result = getRecords(SHEET_NAMES.ORDERS, data.userEmail);
        break;
      case 'book_service':
        result = bookService(data);
        break;
      case 'get_my_services':
        result = getRecords(SHEET_NAMES.SERVICES, data.userEmail);
        break;
      case 'record_payment':
        result = recordPayment(data);
        break;

      case 'admin_update_status':
        result = updateStatus(data.type, data.id, data.newStatus);
        break;
      case 'admin_get_all_records':
        result = getAllRecordsForAdmin(data.type);
        break;
      case 'admin_add_location':
        result = addAllowedLocation(data.zipcode, data.city);
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
 * Utility: Validate Zip Code against Locations Database
 */
function isZipAllowed(zipcode) {
  if (!zipcode) return false;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.LOCATIONS, ['ZipCode', 'City/Area', 'Active']);
  const data = sheet.getDataRange().getValues();
  
  const cleanZip = zipcode.toString().trim();
  
  // Look for zip in the sheet where 'Active' column (3) is true
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === cleanZip && (data[i][2] === true || data[i][2] === 'true' || data[i][2] === 'TRUE')) {
      return true;
    }
  }
  return false;
}

/**
 * Admin: Add new serviceable location
 */
function addAllowedLocation(zip, city) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.LOCATIONS, ['ZipCode', 'City/Area', 'Active']);
  sheet.appendRow([zip, city, true]);
  return { success: true, message: 'Location added to geo-fence' };
}

/**
 * Private Action: Create Order with Zip Verification
 */
function createOrder(data) {
  // Guard: Geo-Fence check using Zip Code
  if (!isZipAllowed(data.zipcode)) {
    return { success: false, message: `Sorry, we currently do not provide delivery to Zip Code: ${data.zipcode}.` };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.ORDERS, ['OrderID', 'Email', 'Items', 'Total', 'Status', 'PaymentStatus', 'Address', 'ZipCode', 'TransactionID', 'Date']);
  const orderId = 'ORD-' + Math.floor(1000 + Math.random() * 9000);
  
  sheet.appendRow([
    orderId, 
    data.userEmail, 
    JSON.stringify(data.items), 
    data.total, 
    'Processing', 
    'Unpaid', 
    data.address, 
    data.zipcode,
    '', 
    new Date()
  ]);
  
  sendEmailNotification(data.userEmail, 'Order Confirmed', `Order ${orderId} has been placed. We will deliver to ${data.address} (${data.zipcode}).`);
  
  notifyAdmins('New Order Received', `
    <strong>Order ID:</strong> ${orderId}<br>
    <strong>Customer:</strong> ${data.userEmail}<br>
    <strong>Zip Code:</strong> ${data.zipcode}<br>
    <strong>Address:</strong> ${data.address}<br>
    <strong>Total:</strong> ${data.total}
  `);

  return { success: true, orderId: orderId };
}

/**
 * Private Action: Book Service with Zip Verification
 */
function bookService(data) {
  if (!isZipAllowed(data.zipcode)) {
    return { success: false, message: `Service area restricted. Zip code ${data.zipcode} is not currently supported.` };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.SERVICES, ['ServiceID', 'Email', 'Category', 'Address', 'ZipCode', 'Status', 'Date']);
  const serviceId = 'SRV-' + Math.floor(1000 + Math.random() * 9000);
  
  sheet.appendRow([serviceId, data.userEmail, data.category, data.address, data.zipcode, 'Requested', new Date()]);
  
  sendEmailNotification(data.userEmail, 'Service Request Received', `We received your request for ${data.category} at ${data.address} (${data.zipcode}).`);
  
  notifyAdmins('New Service Booking', `
    <strong>Service ID:</strong> ${serviceId}<br>
    <strong>Category:</strong> ${data.category}<br>
    <strong>Customer:</strong> ${data.userEmail}<br>
    <strong>Zip Code:</strong> ${data.zipcode}
  `);

  return { success: true, serviceId: serviceId };
}

/**
 * Private Action: Record Payment Details
 */
function recordPayment(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
  if (!sheet) return { success: false, message: 'Order database not found' };

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === data.orderId) {
      sheet.getRange(i + 1, 6).setValue('Paid');
      sheet.getRange(i + 1, 9).setValue(data.transactionId || 'N/A'); // TransactionID moved to Col 9
      
      notifyAdmins('Payment Received', `Payment for ${data.orderId} recorded. Transaction ID: ${data.transactionId}`);
      return { success: true, message: 'Payment recorded successfully' };
    }
  }
  return { success: false, message: 'Order ID not found' };
}

/**
 * Admin Action: Get all orders or services
 */
function getAllRecordsForAdmin(type) {
  const sheetName = type === 'orders' ? SHEET_NAMES.ORDERS : SHEET_NAMES.SERVICES;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return { success: true, data: [] };
  return { success: true, data: sheet.getDataRange().getValues() };
}

/**
 * Admin Action: Status Update Logic
 */
function updateStatus(type, id, newStatus) {
  const sheetName = type === 'order' ? SHEET_NAMES.ORDERS : SHEET_NAMES.SERVICES;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, 5).setValue(newStatus);
      sendEmailNotification(data[i][1], 'Status Update', `Your ${type} ${id} is now: ${newStatus}`);
      return { success: true };
    }
  }
  return { success: false, message: 'ID not found' };
}

/**
 * Auth & Session Management
 */
function getAdminEmailList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adminSheet = getOrInsertSheet(ss, SHEET_NAMES.ADMINS, ['Admin Emails']);
  const values = adminSheet.getDataRange().getValues();
  return values.slice(1).map(row => row[0].toString().trim()).filter(email => email !== "");
}

function checkIfEmailIsAdmin(email) {
  if (!email) return false;
  const adminEmails = getAdminEmailList();
  return adminEmails.some(rowEmail => rowEmail.toLowerCase() === email.toLowerCase());
}

function secureHandleLogin(username, password) {
  const user = verifyUser(username, password); 
  if (!user) return { success: false, message: 'Invalid credentials' };
  const isAdmin = checkIfEmailIsAdmin(user.email);
  const token = createSession(user.email, isAdmin);
  return { success: true, token: token, email: user.email, isAdmin: isAdmin };
}

function validateSession(token) {
  if (!token) return { success: false, message: 'Auth token required' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sessionSheet = getOrInsertSheet(ss, SHEET_NAMES.SESSIONS, ['Token', 'Email', 'Expiry', 'IsAdmin']);
  const data = sessionSheet.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      const expiry = new Date(data[i][2]);
      if (now > expiry) return { success: false, message: 'Session expired' };
      return { success: true, email: data[i][1], isAdmin: data[i][3] === true || data[i][3] === 'true' };
    }
  }
  return { success: false, message: 'Invalid session' };
}

function createSession(email, isAdmin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.SESSIONS, ['Token', 'Email', 'Expiry', 'IsAdmin']);
  const token = Utilities.getUuid();
  const expiry = new Date();
  expiry.setHours(expiry.getHours() + TOKEN_EXPIRY_HOURS);
  sheet.appendRow([token, email, expiry, isAdmin]);
  return token;
}

function getRecords(sheetName, email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return { success: true, data: [] };
  return { success: true, data: sheet.getDataRange().getValues().filter((row, idx) => idx > 0 && row[1] === email) };
}

function notifyAdmins(subject, details) {
  const admins = getAdminEmailList();
  if (admins.length === 0) return;
  const html = `<div style="font-family:sans-serif;padding:20px;border:2px solid #00b894;border-radius:8px;"><h2>Admin Alert</h2><p>${subject}</p><div>${details}</div></div>`;
  MailApp.sendEmail({ to: admins.join(','), subject: `[ADMIN ALERT] ${subject}`, htmlBody: html, name: "Antinna System", replyTo: "contact@antinna.in" });
}

function sendEmailNotification(to, subject, body) {
  const html = `<div style="font-family:sans-serif;padding:25px;border:1px solid #eee;border-radius:8px;"><h2>Antinna</h2><h3>${subject}</h3><p>${body}</p></div>`;
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
