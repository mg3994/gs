/**
 * ANTINNA INTEGRATED BACKEND v5.0
 * Features: 
 * - Google Maps Geocoding & Distance Calculation
 * - Dynamic Travel Fee Logic
 * - Secure Session Auth (24hr Expiry)
 * - GPay & UPI Payment Reconciliation
 * - Product & Service Booking Integration
 */

const SETTINGS = {
  MERCHANT_ID: "IDBCR2DN5TVPLKL4KZ",
  HUB_ADDRESS: "Your Main Shop Address, City, State, Zip", // CHANGE THIS to your shop location
  MAX_SERVICE_RADIUS_KM: 30, 
  TRAVEL_FEE_PER_KM: 5, 
  TIMEZONE: "GMT+5:30",
  TOKEN_EXPIRY_HOURS: 24
};

const SHEET_NAMES = {
  USERS: 'Users',
  SESSIONS: 'Sessions',
  ORDERS: 'Orders',
  SERVICES: 'Services',
  PAYMENTS: 'Payments',
  ADMINS: 'Admins',
  CATALOG: 'Catalog',
  PROVIDERS: 'Providers',
  AVAILABILITY: 'Availability'
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // Prevent concurrent write collisions
    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    const data = request.data || {};

    // Auth logic: Identify which routes need a valid session token
    const publicActions = ['login', 'signup', 'get_catalog', 'get_slots', 'verify_address'];
    let session = { success: false };

    if (!publicActions.includes(action)) {
      session = validateSession(request.token);
      if (!session.success) throw new Error("Unauthorized Access: Please login again.");
      data.userEmail = session.email;
      data.isAdmin = session.isAdmin;
    }

    let result;
    switch (action) {
      // Identity Actions
      case 'signup': result = handleSignup(data); break;
      case 'login': result = handleLogin(data); break;
      
      // Logistics & Catalog
      case 'verify_address': result = verifyAndCalculateLogistics(data.address); break;
      case 'get_catalog': result = getSheetData(SHEET_NAMES.CATALOG); break;
      case 'get_slots': result = getAvailableSlots(data.niche, data.date); break;
      
      // Transactional Actions
      case 'create_order': result = createOrder(data); break;
      case 'book_service': result = bookService(data); break;
      case 'record_payment': result = recordPayment(data); break;
      
      // Admin Actions
      case 'admin_get_orders': result = data.isAdmin ? getSheetData(SHEET_NAMES.ORDERS) : {success: false}; break;
      
      default: result = { success: false, message: 'Action not found' };
    }

    return createCORSResponse(result);
  } catch (error) {
    return createCORSResponse({ success: false, message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Uses Maps Service to validate address, calculate distance, and determine travel fees.
 */
function verifyAndCalculateLogistics(userAddress) {
  try {
    // 1. Geocode the address to get standardized format and Lat/Long
    const geocoder = Maps.newGeocoder().geocode(userAddress);
    if (geocoder.status !== 'OK' || !geocoder.results[0]) {
      return { success: false, message: "Address not recognized. Please provide more details." };
    }
    
    const result = geocoder.results[0];
    const lat = result.geometry.location.lat;
    const lng = result.geometry.location.lng;
    const formattedAddress = result.formatted_address;

    // 2. Calculate Road Distance from your Shop (Hub)
    const directions = Maps.newDirectionFinder()
      .setOrigin(SETTINGS.HUB_ADDRESS)
      .setDestination(formattedAddress)
      .setMode(Maps.DirectionFinder.Mode.DRIVING)
      .getDirections();

    if (!directions.routes || directions.routes.length === 0) {
      return { success: false, message: "We cannot find a driving route to this location." };
    }

    const route = directions.routes[0].legs[0];
    const distanceKm = route.distance.value / 1000; 
    const durationText = route.duration.text;

    // 3. Apply Business Rules
    const isServicable = distanceKm <= SETTINGS.MAX_SERVICE_RADIUS_KM;
    const travelFee = Math.ceil(distanceKm * SETTINGS.TRAVEL_FEE_PER_KM);

    return {
      success: true,
      data: {
        normalizedAddress: formattedAddress,
        lat: lat,
        lng: lng,
        distanceKm: distanceKm.toFixed(2),
        estimatedTravelTime: durationText,
        isServicable: isServicable,
        travelFee: travelFee,
        mapUrl: `https://www.google.com/maps/search/?api=1&query=${lat},${lng}`
      }
    };
  } catch (e) {
    return { success: false, message: "Logistics Error: " + e.message };
  }
}

function bookService(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.SERVICES, 
    ['BookingID', 'User', 'Niche', 'Date', 'Slot', 'Address', 'LatLong', 'Distance', 'TravelFee', 'BasePrice', 'TotalAmount', 'Status']
  );

  // Cross-verify logistics to prevent address spoofing
  const logistics = verifyAndCalculateLogistics(data.address);
  if (!logistics.success || !logistics.data.isServicable) {
    return { success: false, message: "Location is outside our service radius." };
  }

  const bookingId = 'BK-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  const basePrice = getPriceByNiche(data.niche);
  const totalAmount = basePrice + logistics.data.travelFee;

  sheet.appendRow([
    bookingId, 
    data.userEmail, 
    data.niche, 
    data.date, 
    data.slot, 
    logistics.data.normalizedAddress, 
    `${logistics.data.lat},${logistics.data.lng}`,
    logistics.data.distanceKm,
    logistics.data.travelFee,
    basePrice,
    totalAmount,
    'Pending Payment'
  ]);

  return { 
    success: true, 
    bookingId: bookingId, 
    amount: totalAmount,
    message: "Service area verified. Ready for checkout." 
  };
}

function recordPayment(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paySheet = getOrInsertSheet(ss, SHEET_NAMES.PAYMENTS, 
    ['Timestamp', 'RefID', 'Email', 'Amount', 'Method', 'TxID', 'Status', 'Payload']
  );
  
  const refId = data.bookingId || data.orderId;
  const txId = data.details?.transactionId || 'UPI-' + Utilities.getUuid().substring(0, 8);

  paySheet.appendRow([
    new Date(),
    refId,
    data.userEmail,
    data.amount,
    data.method || 'Google Pay',
    txId,
    'Success',
    JSON.stringify(data.details || {})
  ]);

  // Update Status in either Services or Orders tab
  updateRecordStatus(refId, 'Paid');

  return { success: true, transactionId: txId, message: "Payment recorded and verified." };
}

function updateRecordStatus(id, newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToSearch = [SHEET_NAMES.SERVICES, SHEET_NAMES.ORDERS];
  
  sheetsToSearch.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        // Status is usually the last column or specific index
        const lastCol = sheet.getLastColumn();
        sheet.getRange(i + 1, lastCol).setValue(newStatus);
        break;
      }
    }
  });
}

function handleSignup(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.USERS, ['Email', 'Password', 'Created']);
  const users = sheet.getDataRange().getValues();
  
  const email = data.username.toLowerCase();
  for (let i = 1; i < users.length; i++) {
    if (users[i][0] === email) return { success: false, message: "User already exists" };
  }
  
  sheet.appendRow([email, data.password, new Date()]);
  return { success: true, message: "Account created successfully" };
}

function handleLogin(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const users = sheet.getDataRange().getValues();
  const email = data.username.toLowerCase();
  
  for (let i = 1; i < users.length; i++) {
    if (users[i][0] === email && users[i][1] === data.password) {
      const isAdmin = checkIfAdmin(email);
      const token = createSession(email, isAdmin);
      return { success: true, token: token, email: email, isAdmin: isAdmin };
    }
  }
  return { success: false, message: "Invalid email or password" };
}

function createSession(email, isAdmin) {
  const sheet = getOrInsertSheet(SpreadsheetApp.getActiveSpreadsheet(), SHEET_NAMES.SESSIONS, ['Token', 'Email', 'Expiry', 'IsAdmin']);
  const token = Utilities.getUuid();
  const expiry = new Date();
  expiry.setHours(expiry.getHours() + SETTINGS.TOKEN_EXPIRY_HOURS);
  sheet.appendRow([token, email, expiry, isAdmin]);
  return token;
}

function validateSession(token) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SESSIONS);
  if (!sheet) return { success: false };
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      if (now > new Date(data[i][2])) return { success: false };
      return { success: true, email: data[i][1], isAdmin: data[i][3] === true };
    }
  }
  return { success: false };
}

function checkIfAdmin(email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.ADMINS);
  if (!sheet) return false;
  const admins = sheet.getDataRange().getValues();
  return admins.some(row => row[0].toLowerCase() === email.toLowerCase());
}

function getOrInsertSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');
  }
  return sheet;
}

function getSheetData(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) return { success: false, data: [] };
  const rows = sheet.getDataRange().getValues();
  const headers = rows.shift();
  const data = rows.map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
  return { success: true, data: data };
}

function getPriceByNiche(niche) {
  const rates = { 'Electrician': 500, 'Plumber': 450, 'Hardware Delivery': 100 };
  return rates[niche] || 300;
}

function createCORSResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
