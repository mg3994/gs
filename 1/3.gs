/**
 * Advanced Backend for Ecommerce, Retail, & Handy Services
 * Features: Public/Private/Admin Route Architecture, Secure Admin Allowlist, Token-based Auth
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
 * Defines which actions require what level of authorization
 */
const ROUTES = {
  PUBLIC: ['login', 'signup', 'forgot_password', 'verify_otp', 'reset_password', 'get_public_catalog'],
  PRIVATE: ['check_auth', 'create_order', 'get_my_orders', 'book_service', 'get_my_services', 'record_payment'],
  ADMIN: ['admin_update_status', 'admin_get_all_records', 'admin_manage_inventory']
};

/**
 * Enhanced API Entry Point with Route Guarding
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    // 1. Identify Route Level
    const isPublic = ROUTES.PUBLIC.includes(action);
    const isPrivate = ROUTES.PRIVATE.includes(action);
    const isAdmin = ROUTES.ADMIN.includes(action);

    let authContext = { success: true, email: null, isAdmin: false };

    // 2. Auth Guard for Private & Admin Routes
    if (isPrivate || isAdmin) {
      authContext = validateSession(data.token);
      if (!authContext.success) return createCORSResponse(JSON.stringify(authContext));
      
      // Upgrade context: Attach validated identity
      data.userEmail = authContext.email; 
      data.isAdmin = authContext.isAdmin;
    }

    // 3. Specific Admin Guard for Admin Routes
    if (isAdmin && !data.isAdmin) {
      // Security: Re-verify against sheet for Admin actions just in case session is stale
      if (!checkIfEmailIsAdmin(data.userEmail)) {
        throw new Error('Unauthorized: Admin database verification failed');
      }
    }

    let result;
    switch (action) {
      // --- PUBLIC ROUTES ---
      case 'login': 
        result = secureHandleLogin(data.username, data.password); 
        break;
      case 'signup': 
        result = handleSignup(data.username, data.password); 
        break;
      
      // --- PRIVATE ROUTES (Requires Token) ---
      case 'check_auth':
        result = { success: true, email: data.userEmail, isAdmin: data.isAdmin };
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

      // --- ADMIN ROUTES (Requires Admin Token + Sheet Verification) ---
      case 'admin_update_status':
        result = updateStatus(data.type, data.id, data.newStatus);
        break;
      case 'admin_get_all_records':
        result = getAllRecordsForAdmin(data.type); // type: 'orders' or 'services'
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
 * Verifies if an email exists in the Admin sheet
 */
function checkIfEmailIsAdmin(email) {
  if (!email) return false;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adminSheet = getOrInsertSheet(ss, SHEET_NAMES.ADMINS, ['Admin Emails']);
  const adminEmails = adminSheet.getDataRange().getValues().flat();
  return adminEmails.some(rowEmail => rowEmail.toString().toLowerCase().trim() === email.toLowerCase().trim());
}

/**
 * Secure Login: Checks Admin List during session creation
 */
function secureHandleLogin(username, password) {
  // Logic to verify user in 'Users' sheet
  const user = verifyUser(username, password); 
  
  if (!user) {
    return { success: false, message: 'Invalid credentials' };
  }

  const isAdmin = checkIfEmailIsAdmin(user.email);
  const token = createSession(user.email, isAdmin);
  
  return {
    success: true,
    token: token,
    email: user.email,
    isAdmin: isAdmin,
    message: 'Login successful'
  };
}

/**
 * Validates a session token and returns user context
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
      
      return { 
        success: true, 
        email: data[i][1], 
        isAdmin: data[i][3] === true || data[i][3] === 'true' 
      };
    }
  }
  return { success: false, message: 'Invalid or missing session' };
}

/**
 * Creates a unique session
 */
function createSession(email, isAdmin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.SESSIONS, ['Token', 'Email', 'Expiry', 'IsAdmin']);
  const token = Utilities.getUuid();
  const expiry = new Date();
  expiry.setHours(expiry.getHours() + TOKEN_EXPIRY_HOURS);
  
  sheet.appendRow([token, email, expiry, isAdmin]);
  return token;
}

/**
 * Private Action: Create Order
 */
function createOrder(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.ORDERS, ['OrderID', 'Email', 'Items', 'Total', 'Status', 'Payment', 'Date']);
  const orderId = 'ORD-' + Math.floor(1000 + Math.random() * 9000);
  sheet.appendRow([orderId, data.userEmail, JSON.stringify(data.items), data.total, 'Pending', 'Unpaid', new Date()]);
  sendEmailNotification(data.userEmail, 'Order Confirmed', `Order ${orderId} has been placed.`);
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
  sendEmailNotification(data.userEmail, 'Service Request', `Request ${serviceId} received.`);
  return { success: true, serviceId: serviceId };
}

/**
 * Private Action: Get User-Specific Records
 */
function getRecords(sheetName, email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return { success: true, data: [] };
  const data = sheet.getDataRange().getValues();
  const userRecords = data.filter((row, index) => index > 0 && row[1] === email);
  return { success: true, data: userRecords };
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
 * Utility: Manage Sheet Initialization
 */
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
    name: "Antinna",
    replyTo: "contact@antinna.in"
  });
}

/**
 * CORS Wrapper
 */
function createCORSResponse(payload) {
  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'POST')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}
