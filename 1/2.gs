/**
 * Advanced Backend for Ecommerce, Retail, & Handy Services
 * Features: Secure Admin Allowlist, Token-based Auth, Status Management
 */

const SHEET_NAMES = {
  USERS: 'Users',
  SESSIONS: 'Sessions',
  ORDERS: 'Orders',
  SERVICES: 'Services',
  PAYMENTS: 'Payments',
  ADMINS: 'Admins' // New sheet for secure admin email list
};

const TOKEN_EXPIRY_HOURS = 24;

/**
 * Enhanced API Entry Point
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    // Actions that DO NOT require a session token
    const publicActions = ['login', 'signup', 'forgot_password', 'verify_otp', 'reset_password'];
    
    if (!publicActions.includes(action)) {
      const auth = validateSession(data.token);
      if (!auth.success) return createCORSResponse(JSON.stringify(auth));
      data.userEmail = auth.email; 
      data.isAdmin = auth.isAdmin;
    }

    let result;
    switch (action) {
      case 'login': 
        result = secureHandleLogin(data.username, data.password); 
        break;
      case 'signup': 
        result = handleSignup(data.username, data.password); 
        break;
      
      // --- SECURE ECOMMERCE ---
      case 'create_order':
        result = createOrder(data);
        break;
      case 'get_my_orders':
        result = getRecords(SHEET_NAMES.ORDERS, data.userEmail);
        break;
      
      // --- SECURE HANDY SERVICES ---
      case 'book_service':
        result = bookService(data);
        break;
      case 'get_my_services':
        result = getRecords(SHEET_NAMES.SERVICES, data.userEmail);
        break;

      // --- ADMIN ONLY ACTIONS ---
      case 'admin_update_status':
        // Double Check: Consult the Admin sheet directly for critical operations
        if (!checkIfEmailIsAdmin(data.userEmail)) {
          throw new Error('Unauthorized: You are not in the admin database');
        }
        result = updateStatus(data.type, data.id, data.newStatus);
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
  
  // Clean strings and check for match
  return adminEmails.some(rowEmail => rowEmail.toString().toLowerCase().trim() === email.toLowerCase().trim());
}

/**
 * Secure Login: Checks Admin List during session creation
 */
function secureHandleLogin(username, password) {
  // 1. Existing password verification logic here...
  // (Assuming verifyUser(username, password) is a function in your existing code)
  const user = verifyUser(username, password); // This should return the user object or null
  
  if (!user) {
    return { success: false, message: 'Invalid credentials' };
  }

  // 2. Determine admin status FROM THE DATABASE, not the client
  const isAdmin = checkIfEmailIsAdmin(user.email);
  
  // 3. Create session with verified status
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
 * Validates a session token
 */
function validateSession(token) {
  if (!token) return { success: false, message: 'Token missing' };
  
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
  return { success: false, message: 'Invalid session' };
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
 * Create Order (Customer Action)
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
 * Book Service (Customer Action)
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
 * Get User-Specific Records
 */
function getRecords(sheetName, email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return { success: true, data: [] };
  const data = sheet.getDataRange().getValues();
  const userRecords = data.filter((row, index) => index > 0 && row[1] === email);
  return { success: true, data: userRecords };
}

/**
 * Admin Status Update Logic
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
 * Email Engine
 * Configured to send from Antinna Business Email
 */
function sendEmailNotification(to, subject, body) {
  const html = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding:25px; border:1px solid #e0e0e0; border-radius: 8px; max-width: 600px; margin: auto;">
      <div style="text-align: center; margin-bottom: 20px;">
        <h1 style="color: #1a1a1a; margin: 0;">Antinna</h1>
        <p style="color: #666; font-size: 14px;">Ecommerce & Handy Services</p>
      </div>
      <div style="background-color: #f9f9f9; padding: 20px; border-radius: 6px;">
        <h3 style="color: #2c3e50; margin-top: 0; border-bottom: 1px solid #eee; padding-bottom: 10px;">${subject}</h3>
        <p style="color: #444; line-height: 1.6; font-size: 16px;">${body}</p>
      </div>
      <div style="margin-top: 25px; text-align: center; font-size: 12px; color: #888;">
        <p>You received this email because of an activity on your account at antinna.in</p>
        <p>&copy; ${new Date().getFullYear()} Antinna. All rights reserved.</p>
      </div>
    </div>
  `;

  // Note: To use contact@antinna.in, you MUST add it as an Alias in your Gmail Settings.
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
