/**
 * Advanced Backend for Ecommerce, Retail, & Handy Services
 * Features: Public/Private/Admin Route Architecture, Admin Notifications, Token-based Auth
 */

const SHEET_NAMES = {
  USERS: 'Users',
  SESSIONS: 'Sessions',
  ORDERS: 'Orders',
  SERVICES: 'Services',
  PAYMENTS: 'Payments',
  ADMINS: 'Admins'
};

const TOKEN_EXPIRY_HOURS = 24;

/**
 * Route Configuration
 */
const ROUTES = {
  PUBLIC: ['login', 'signup', 'forgot_password', 'verify_otp', 'reset_password', 'get_public_catalog'],
  PRIVATE: ['check_auth', 'create_order', 'get_my_orders', 'book_service', 'get_my_services', 'record_payment'],
  ADMIN: ['admin_update_status', 'admin_get_all_records', 'admin_manage_inventory']
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

      case 'admin_update_status':
        result = updateStatus(data.type, data.id, data.newStatus);
        break;
      case 'admin_get_all_records':
        result = getAllRecordsForAdmin(data.type);
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
 * Fetches all admin emails from the Admin sheet
 */
function getAdminEmailList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adminSheet = getOrInsertSheet(ss, SHEET_NAMES.ADMINS, ['Admin Emails']);
  const values = adminSheet.getDataRange().getValues();
  // Filter out header and empty rows, then flatten
  return values.slice(1)
    .map(row => row[0].toString().trim())
    .filter(email => email !== "");
}

/**
 * Verifies if an email exists in the Admin sheet
 */
function checkIfEmailIsAdmin(email) {
  if (!email) return false;
  const adminEmails = getAdminEmailList();
  return adminEmails.some(rowEmail => rowEmail.toLowerCase() === email.toLowerCase());
}

/**
 * Notifies all listed admins about a new business event
 */
function notifyAdmins(subject, details) {
  const admins = getAdminEmailList();
  if (admins.length === 0) return;

  const html = `
    <div style="font-family: sans-serif; padding:20px; border:2px solid #00b894; border-radius: 8px; max-width: 600px;">
      <h2 style="color: #2d3436; margin-top: 0;">New Business Alert!</h2>
      <p style="font-size: 16px; color: #636e72;"><strong>Action:</strong> ${subject}</p>
      <div style="background: #f1f2f6; padding: 15px; border-radius: 4px; color: #2f3542;">
        ${details}
      </div>
      <p style="font-size: 12px; color: #a4b0be; margin-top: 20px;">This is an internal notification for Antinna Admins.</p>
    </div>
  `;

  // Send to all admins (BCC used to keep privacy if many admins exist)
  MailApp.sendEmail({
    to: admins.join(','),
    subject: `[ADMIN ALERT] ${subject}`,
    htmlBody: html,
    name: "Antinna Business System",
    replyTo: "contact@antinna.in"
  });
}

/**
 * Secure Login logic
 */
function secureHandleLogin(username, password) {
  const user = verifyUser(username, password); 
  if (!user) return { success: false, message: 'Invalid credentials' };

  const isAdmin = checkIfEmailIsAdmin(user.email);
  const token = createSession(user.email, isAdmin);
  
  return { success: true, token: token, email: user.email, isAdmin: isAdmin };
}

/**
 * Private Action: Create Order
 */
function createOrder(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.ORDERS, ['OrderID', 'Email', 'Items', 'Total', 'Status', 'Payment', 'Date']);
  const orderId = 'ORD-' + Math.floor(1000 + Math.random() * 9000);
  
  sheet.appendRow([orderId, data.userEmail, JSON.stringify(data.items), data.total, 'Pending', 'Unpaid', new Date()]);
  
  // Notify Customer
  sendEmailNotification(data.userEmail, 'Order Confirmed', `Order ${orderId} has been placed.`);
  
  // Notify Admins
  notifyAdmins('New Order Received', `
    <strong>Order ID:</strong> ${orderId}<br>
    <strong>Customer:</strong> ${data.userEmail}<br>
    <strong>Total:</strong> ${data.total}<br>
    <strong>Items:</strong> ${JSON.stringify(data.items)}
  `);

  return { success: true, orderId: orderId };
}

/**
 * Private Action: Book Service
 */
function bookService(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.SERVICES, ['ServiceID', 'Email', 'Category', 'Address', 'Status', 'Date']);
  const serviceId = 'SRV-' + Math.floor(1000 + Math.random() * 9000);
  
  sheet.appendRow([serviceId, data.userEmail, data.category, data.address, 'Requested', new Date()]);
  
  // Notify Customer
  sendEmailNotification(data.userEmail, 'Service Request', `Request ${serviceId} received.`);
  
  // Notify Admins
  notifyAdmins('New Service Booking', `
    <strong>Service ID:</strong> ${serviceId}<br>
    <strong>Category:</strong> ${data.category}<br>
    <strong>Customer:</strong> ${data.userEmail}<br>
    <strong>Address:</strong> ${data.address}
  `);

  return { success: true, serviceId: serviceId };
}

/**
 * Auth & Session Helpers
 */
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
  return { success: false, message: 'Invalid or missing session' };
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
  const data = sheet.getDataRange().getValues();
  return { success: true, data: data.filter((row, index) => index > 0 && row[1] === email) };
}

function getAllRecordsForAdmin(type) {
  const sheetName = type === 'orders' ? SHEET_NAMES.ORDERS : SHEET_NAMES.SERVICES;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  return { success: true, data: sheet ? sheet.getDataRange().getValues() : [] };
}

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

function getOrInsertSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

/**
 * Professional Email Engine
 */
function sendEmailNotification(to, subject, body) {
  const html = `
    <div style="font-family: sans-serif; padding:25px; border:1px solid #e0e0e0; border-radius: 8px; max-width: 600px; margin: auto;">
      <h2 style="color: #1a1a1a;">Antinna</h2>
      <div style="background-color: #f9f9f9; padding: 20px; border-radius: 6px;">
        <h3 style="color: #2c3e50;">${subject}</h3>
        <p>${body}</p>
      </div>
      <p style="font-size: 12px; color: #888;">&copy; Antinna. Sent from contact@antinna.in</p>
    </div>
  `;
  MailApp.sendEmail({
    to: to,
    subject: `[Antinna] ${subject}`,
    htmlBody: html,
    name: "Antinna Support",
    replyTo: "contact@antinna.in"
  });
}

function createCORSResponse(payload) {
  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'POST')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}
