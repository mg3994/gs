/**
 * ANTINNA LOGISTICS-INTEGRATED BACKEND v4.0
 * Features: Google Maps Geocoding, Distance Calculation, and Address Normalization.
 */

const SETTINGS = {
  MERCHANT_ID: "IDBCR2DN5TVPLKL4KZ",
  HUB_ADDRESS: "Your Main Shop Address, City, State, Zip", // CHANGE THIS to your shop location
  MAX_SERVICE_RADIUS_KM: 30, // Limit service to 30km driving distance
  TRAVEL_FEE_PER_KM: 5, // Additional charge per KM
  TIMEZONE: "GMT+5:30"
};

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

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    const data = request.data || {};

    // Auth logic
    const isPublic = ['login', 'signup', 'get_catalog', 'get_slots', 'verify_address'].includes(action);
    let session = { success: false };

    if (!isPublic) {
      session = validateSession(request.token);
      if (!session.success) throw new Error("Unauthorized");
      data.userEmail = session.email;
    }

    let result;
    switch (action) {
      case 'signup': result = handleSignup(data); break;
      case 'login': result = handleLogin(data); break;
      case 'verify_address': result = verifyAndCalculateLogistics(data.address); break;
      case 'book_service': result = bookService(data); break;
      case 'create_order': result = createOrder(data); break;
      case 'record_payment': result = recordPayment(data); break;
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
 * Uses Google Maps Service to validate address and calculate distance from Hub.
 */
function verifyAndCalculateLogistics(userAddress) {
  try {
    // 1. Geocode the address
    const geocoder = Maps.newGeocoder().geocode(userAddress);
    if (geocoder.status !== 'OK') return { success: false, message: "Address not found. Please be more specific." };
    
    const result = geocoder.results[0];
    const lat = result.geometry.location.lat;
    const lng = result.geometry.location.lng;
    const formattedAddress = result.formatted_address;

    // 2. Calculate Distance from Shop (Hub)
    const directions = Maps.newDirectionFinder()
      .setOrigin(SETTINGS.HUB_ADDRESS)
      .setDestination(formattedAddress)
      .setMode(Maps.DirectionFinder.Mode.DRIVING)
      .getDirections();

    if (!directions.routes || directions.routes.length === 0) {
      return { success: false, message: "We cannot calculate a driving route to this location." };
    }

    const route = directions.routes[0].legs[0];
    const distanceKm = route.distance.value / 1000; // convert meters to km
    const durationText = route.duration.text;

    // 3. Validation Logic
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
    return { success: false, message: "Map Service Error: " + e.message };
  }
}

function bookService(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrInsertSheet(ss, SHEET_NAMES.SERVICES, 
    ['BookingID', 'User', 'Niche', 'Date', 'Slot', 'Address', 'LatLong', 'Distance', 'TravelFee', 'Status']
  );

  // Re-verify logistics to ensure accuracy
  const logistics = verifyAndCalculateLogistics(data.address);
  if (!logistics.success || !logistics.data.isServicable) {
    return { success: false, message: "Address is outside our service area." };
  }

  const bookingId = 'BK-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  const coords = `${logistics.data.lat},${logistics.data.lng}`;

  sheet.appendRow([
    bookingId, 
    data.userEmail, 
    data.niche, 
    data.date, 
    data.slot, 
    logistics.data.normalizedAddress, 
    coords,
    logistics.data.distanceKm,
    logistics.data.travelFee,
    'Pending Payment'
  ]);

  return { 
    success: true, 
    bookingId: bookingId, 
    travelFee: logistics.data.travelFee,
    message: "Service area verified. Travel fee calculated." 
  };
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

function validateSession(token) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SESSIONS);
  if (!sheet) return { success: false };
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      if (now > new Date(data[i][2])) return { success: false };
      return { success: true, email: data[i][1] };
    }
  }
  return { success: false };
}

function createCORSResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// ... Additional standard functions (login/signup) remain consistent with previous versions ...
