/****************************************************************
 * EVENT ATTENDANCE – Code.gs
 *
 * Sheets used:
 *   Registered : A ID | B Full Name | C First Name | D Last Name |
 *                E Contact | F Address | G Invited By | H Date Registered
 *   Attendance : A Timestamp | B ID | C Full Name | D Contact |
 *                E Address | F Invited By | G Type
 *   Settings   : created automatically (see SETTINGS_FIELDS below).
 *                You can edit column B directly in the sheet, or use
 *                the "Event settings" button on the app's home screen.
 ****************************************************************/

const REGISTERED_SHEET = 'Registered';
const ATTENDANCE_SHEET = 'Attendance';
const SETTINGS_SHEET = 'Settings';
const BACKGROUND_FOLDER = 'Event Attendance Backgrounds';

/*
 * One setting per row in the Settings sheet (row 1 = first item).
 * Row 1 is still the event name, so older Settings sheets keep working.
 */
const SETTINGS_FIELDS = [
  { key: 'eventName',       label: 'Event Name',                     def: 'Event Attendance' },
  { key: 'welcomeMessage',  label: 'Welcome Message',                def: 'Welcome! Please select an option below.' },
  { key: 'background',      label: 'Background (color or gradient)', def: '#f4f6f8' },
  { key: 'backgroundImage', label: 'Background Image URL',           def: '' },
  { key: 'overlay',         label: 'Image Darkness (0-80)',          def: 35 },
  { key: 'accentColor',     label: 'Button Color',                   def: '#222222' },
  { key: 'adminPin',        label: 'Admin PIN',                      def: '1234' }
];


/* ============================================================
 * WEB APP
 * ============================================================ */

function doGet() {
  const settings = getPublicSettings();

  const template = HtmlService.createTemplateFromFile('Index');

  // Settings are injected into the page so the background shows
  // immediately (no flash of the default look).
  template.settingsJson = JSON.stringify(settings).replace(/</g, '\\u003c');

  return template
    .evaluate()
    .setTitle(settings.eventName)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/* ============================================================
 * SETTINGS
 * ============================================================ */

function getSettingsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SETTINGS_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET);
  }

  // Write labels into column A where they are missing.
  const labelRange = sheet.getRange(1, 1, SETTINGS_FIELDS.length, 1);
  const labels = labelRange.getValues();
  let changed = false;

  SETTINGS_FIELDS.forEach(function(field, i) {
    if (!String(labels[i][0] || '').trim()) {
      labels[i][0] = field.label;
      changed = true;
    }
  });

  if (changed) {
    labelRange.setValues(labels);
  }

  return sheet;
}


function readSettings_() {
  const sheet = getSettingsSheet_();
  const values = sheet
    .getRange(1, 2, SETTINGS_FIELDS.length, 1)
    .getValues();

  const settings = {};

  SETTINGS_FIELDS.forEach(function(field, i) {
    const raw = values[i][0];
    settings[field.key] =
      raw === '' || raw === null ? field.def : raw;
  });

  settings.eventName =
    String(settings.eventName).trim() || 'Event Attendance';

  settings.welcomeMessage = String(settings.welcomeMessage);

  settings.background =
    sanitizeBackground_(settings.background) || '#f4f6f8';

  settings.backgroundImage =
    sanitizeImageUrl_(settings.backgroundImage);

  settings.overlay = clampOverlay_(settings.overlay);

  settings.accentColor =
    sanitizeHexColor_(settings.accentColor) || '#222222';

  settings.adminPin =
    String(settings.adminPin).trim() || '1234';

  return settings;
}


/**
 * Settings that are safe to send to the browser (no PIN).
 */
function getPublicSettings() {
  const settings = readSettings_();
  delete settings.adminPin;
  return settings;
}


/**
 * Kept for backward compatibility.
 */
function getEventName() {
  return readSettings_().eventName;
}


function checkPin_(pin) {
  const settings = readSettings_();

  if (String(pin || '').trim() !== settings.adminPin) {
    throw new Error('Incorrect PIN.');
  }
}


function verifyAdminPin(pin) {
  try {
    checkPin_(pin);
    return { ok: true };
  } catch (e) {
    return { ok: false };
  }
}


/**
 * Save event settings from the in-app settings screen.
 */
function saveSettings(pin, data) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(15000)) {
    throw new Error('The system is busy. Please try again.');
  }

  try {
    checkPin_(pin);

    data = data || {};
    const current = readSettings_();

    const eventName = String(data.eventName || '').trim();
    if (!eventName) {
      throw new Error('Please enter the event name.');
    }

    const background = sanitizeBackground_(data.background);
    if (!background) {
      throw new Error('That background color is not valid.');
    }

    const imageInput = String(data.backgroundImage || '').trim();
    const backgroundImage = sanitizeImageUrl_(imageInput);
    if (imageInput && !backgroundImage) {
      throw new Error('The image link must start with https:// and contain no spaces or quotes.');
    }

    const accentColor = sanitizeHexColor_(data.accentColor);
    if (!accentColor) {
      throw new Error('That button color is not valid.');
    }

    let newPin = current.adminPin;
    const requestedPin = String(data.newPin || '').trim();
    if (requestedPin) {
      if (!/^[0-9A-Za-z]{4,12}$/.test(requestedPin)) {
        throw new Error('The new PIN must be 4 to 12 letters or numbers.');
      }
      newPin = requestedPin;
    }

    const values = {
      eventName: eventName,
      welcomeMessage: String(data.welcomeMessage || '').trim(),
      background: background,
      backgroundImage: backgroundImage,
      overlay: clampOverlay_(data.overlay),
      accentColor: accentColor,
      adminPin: newPin
    };

    const sheet = getSettingsSheet_();
    const range = sheet.getRange(1, 2, SETTINGS_FIELDS.length, 1);

    // Keep the PIN as text so a PIN like 0123 keeps its leading zero.
    sheet.getRange(SETTINGS_FIELDS.length, 2).setNumberFormat('@');

    range.setValues(
      SETTINGS_FIELDS.map(function(field) {
        return [values[field.key]];
      })
    );

    SpreadsheetApp.flush();

    return getPublicSettings();

  } finally {
    lock.releaseLock();
  }
}


/**
 * Upload a background image (sent from the browser as a data URL)
 * to Google Drive and return a link the page can display.
 */
function uploadBackgroundImage(pin, dataUrl, fileName) {
  checkPin_(pin);

  const match = String(dataUrl || '')
    .match(/^data:(image\/(?:png|jpeg|webp|gif));base64,(.+)$/);

  if (!match) {
    throw new Error('Please choose a PNG, JPG, WEBP or GIF image.');
  }

  const mimeType = match[1];
  const bytes = Utilities.base64Decode(match[2]);

  if (bytes.length > 8 * 1024 * 1024) {
    throw new Error('The image is too large. Please use one under 8 MB.');
  }

  const safeName =
    String(fileName || 'background')
      .replace(/[^\w.\- ]+/g, '')
      .slice(0, 80) || 'background';

  const blob = Utilities.newBlob(
    bytes,
    mimeType,
    getEventName() + ' - ' + safeName
  );

  const folders = DriveApp.getFoldersByName(BACKGROUND_FOLDER);
  const folder = folders.hasNext()
    ? folders.next()
    : DriveApp.createFolder(BACKGROUND_FOLDER);

  const file = folder.createFile(blob);

  try {
    file.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );
  } catch (e) {
    file.setTrashed(true);
    throw new Error(
      'The image was uploaded, but Google Drive would not allow it to be shared ' +
      'publicly (your organization may block this). Paste an image link instead.'
    );
  }

  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w2000';
}


function sanitizeBackground_(value) {
  const v = String(value || '').trim();
  if (!v || v.length > 300) return '';
  // Colors and gradients only: letters, digits, # ( ) , . % - and spaces.
  return /^[#a-zA-Z0-9(),.%\s-]+$/.test(v) ? v : '';
}


function sanitizeImageUrl_(value) {
  const v = String(value || '').trim();
  if (!v) return '';
  return /^https:\/\/[^\s"'()<>\\]+$/.test(v) && v.length <= 1000 ? v : '';
}


function sanitizeHexColor_(value) {
  const v = String(value || '').trim();
  return /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(v) ? v : '';
}


function clampOverlay_(value) {
  const n = Number(value);
  if (isNaN(n)) return 35;
  return Math.max(0, Math.min(80, Math.round(n)));
}


/* ============================================================
 * NAME DIRECTORY (for live suggestions)
 * ============================================================ */

/**
 * Return every registered name with its ID and whether the
 * person has already checked in today. Only names and IDs are
 * sent to the browser – no contact numbers or addresses.
 */
function getAttendeeDirectory() {
  const sheet = getSheet_(REGISTERED_SHEET);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const presentIds = getTodayPresentIds_();
  const list = [];

  rows.forEach(function(row) {
    const id = normalizeId(row[0]);
    const fullName = String(row[1] || '').trim();

    if (id && fullName) {
      list.push({
        id: id,
        fullName: fullName,
        present: presentIds.has(id)
      });
    }
  });

  return list;
}


/* ============================================================
 * SEARCH
 * ============================================================ */

/**
 * Search registered attendees. Every word typed must appear
 * somewhere in the name, so "Juan Cruz" finds "Juan Dela Cruz".
 */
function searchAttendees(searchTerm) {
  const term = normalizeName(searchTerm);

  if (!term) {
    return [];
  }

  const tokens = term.split(' ');

  return getAttendeeDirectory().filter(function(person) {
    const name = normalizeName(person.fullName);
    return tokens.every(function(t) {
      return name.indexOf(t) !== -1;
    });
  });
}


/**
 * Used by the NEW ATTENDEE screen to catch people who are already
 * registered. Matches:
 *   1. The exact same full name, or
 *   2. Names that contain every typed word as a whole word
 *      (only when at least two words were typed).
 */
function findNameMatches(fullName) {
  const name = normalizeName(fullName);

  if (!name) {
    return [];
  }

  const tokens = name.split(' ');
  const directory = getAttendeeDirectory();

  const exact = directory.filter(function(p) {
    return normalizeName(p.fullName) === name;
  });

  if (exact.length || tokens.length < 2) {
    return exact;
  }

  return directory.filter(function(p) {
    const words = normalizeName(p.fullName).split(' ');
    return tokens.every(function(t) {
      return words.indexOf(t) !== -1;
    });
  });
}


/* ============================================================
 * NORMALIZERS
 * ============================================================ */

function normalizeName(name) {
  return String(name || '')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '') // ñ -> n, é -> e
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}


/**
 * "0001", "1" and 1 are all treated as the same ID.
 */
function normalizeId(id) {
  return String(id === null || id === undefined ? '' : id)
    .trim()
    .replace(/^0+(?=\d)/, '');
}


/**
 * ID as stored in the sheet: numeric IDs padded to 4 digits.
 */
function displayId_(id) {
  const n = normalizeId(id);
  return /^\d+$/.test(n) ? n.padStart(4, '0') : n;
}


/**
 * Contact numbers are compared by digits only, so these all match:
 * 0917 123 4567, 09171234567, +63 917 123 4567, 9171234567
 */
function normalizeContact(contact) {
  let digits = String(contact || '').replace(/\D/g, '');

  if (digits.length === 12 && digits.indexOf('63') === 0) {
    digits = '0' + digits.slice(2);
  } else if (digits.length === 10 && digits.charAt(0) === '9') {
    digits = '0' + digits;
  }

  return digits;
}


/* ============================================================
 * DUPLICATE CHECK
 * ============================================================ */

/**
 * Duplicate is based on:
 *   1. Exact full name, or
 *   2. Same contact number.
 * First name alone or last name alone is NOT a duplicate.
 */
function checkExistingAttendee(fullName, contactNumber) {
  const sheet = getSheet_(REGISTERED_SHEET);
  const data = sheet.getDataRange().getValues();

  const name = normalizeName(fullName);
  const contact = normalizeContact(contactNumber);

  for (let i = 1; i < data.length; i++) {
    const existingName = normalizeName(data[i][1]);
    const existingContact = normalizeContact(data[i][4]);

    if (name && existingName === name) {
      return {
        exists: true,
        reason: 'name',
        id: normalizeId(data[i][0]),
        fullName: String(data[i][1]).trim()
      };
    }

    if (contact && existingContact === contact) {
      return {
        exists: true,
        reason: 'contact',
        id: normalizeId(data[i][0]),
        fullName: String(data[i][1]).trim()
      };
    }
  }

  return { exists: false };
}


/* ============================================================
 * CHECK-IN
 * ============================================================ */

/**
 * Mark an existing attendee as present (once per day).
 */
function markAttendance(attendeeId) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(15000)) {
    throw new Error('The system is busy. Please try again.');
  }

  try {
    const registeredSheet = getSheet_(REGISTERED_SHEET);
    const attendanceSheet = getSheet_(ATTENDANCE_SHEET);

    const targetId = normalizeId(attendeeId);

    if (!targetId) {
      throw new Error('No attendee ID was provided.');
    }

    // Find the attendee by ID.
    const data = registeredSheet.getDataRange().getValues();
    let attendee = null;

    for (let i = 1; i < data.length; i++) {
      if (normalizeId(data[i][0]) === targetId) {
        attendee = {
          id: displayId_(data[i][0]),
          fullName: String(data[i][1] || '').trim(),
          contact: String(data[i][4] || '').trim(),
          address: String(data[i][5] || '').trim(),
          invitedBy: String(data[i][6] || '').trim()
        };
        break;
      }
    }

    if (!attendee) {
      throw new Error('Attendee could not be found.');
    }

    // Already checked in today?
    const existing = findTodayAttendance_(attendanceSheet, targetId);

    if (existing) {
      return {
        success: false,
        alreadyPresent: true,
        id: targetId,
        fullName: attendee.fullName,
        timestamp: formatTimestamp(existing)
      };
    }

    const now = new Date();

    appendRowWithTextColumns_(
      attendanceSheet,
      [
        now,
        attendee.id,
        attendee.fullName,
        attendee.contact,
        attendee.address,
        attendee.invitedBy,
        'Registered'
      ],
      [2, 4] // ID and contact stay as text (keeps leading zeros)
    );

    return {
      success: true,
      alreadyPresent: false,
      id: targetId,
      fullName: attendee.fullName,
      timestamp: formatTimestamp(now)
    };

  } finally {
    lock.releaseLock();
  }
}


/* ============================================================
 * REGISTRATION
 * ============================================================ */

/**
 * Register a new attendee and automatically check them in.
 */
function registerNewAttendee(formData) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(15000)) {
    throw new Error('The system is busy. Please try again.');
  }

  try {
    const registeredSheet = getSheet_(REGISTERED_SHEET);
    const attendanceSheet = getSheet_(ATTENDANCE_SHEET);

    formData = formData || {};

    const fullName = String(formData.fullName || '').trim().replace(/\s+/g, ' ');
    const contact = String(formData.contact || '').trim();
    const address = String(formData.address || '').trim();
    const invitedBy = String(formData.invitedBy || '').trim();

    if (!fullName) throw new Error('Please enter your full name.');
    if (!contact) throw new Error('Please enter your contact number.');
    if (normalizeContact(contact).length < 7) {
      throw new Error('Please enter a valid contact number.');
    }
    if (!address) throw new Error('Please enter your address.');
    if (!invitedBy) throw new Error('Please enter who invited you.');

    // Final duplicate check (inside the lock).
    const existing = checkExistingAttendee(fullName, contact);

    if (existing.exists) {
      return {
        success: false,
        alreadyExists: true,
        reason: existing.reason,
        id: existing.id,
        fullName: existing.fullName
      };
    }

    const id = generateNextId(registeredSheet);
    const now = new Date();

    // "Dannie Lyn Rodriguez" -> First: Dannie, Last: Lyn Rodriguez
    const nameParts = fullName.split(' ');
    const firstName = nameParts.shift() || '';
    const lastName = nameParts.join(' ');

    appendRowWithTextColumns_(
      registeredSheet,
      [id, fullName, firstName, lastName, contact, address, invitedBy, now],
      [1, 5] // ID and contact stay as text
    );

    appendRowWithTextColumns_(
      attendanceSheet,
      [now, id, fullName, contact, address, invitedBy, 'New Attendee'],
      [2, 4]
    );

    return {
      success: true,
      alreadyExists: false,
      id: normalizeId(id),
      fullName: fullName,
      timestamp: formatTimestamp(now)
    };

  } finally {
    lock.releaseLock();
  }
}


/**
 * Generate the next sequential ID (0001, 0002, ...).
 */
function generateNextId(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return '0001';
  }

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  let highest = 0;

  ids.forEach(function(id) {
    const number = parseInt(normalizeId(id), 10);
    if (!isNaN(number) && number > highest) {
      highest = number;
    }
  });

  return String(highest + 1).padStart(4, '0');
}


/* ============================================================
 * HELPERS
 * ============================================================ */

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error(name + ' sheet was not found.');
  }
  return sheet;
}


/**
 * Append a row, keeping the listed columns (1-based) as plain text.
 * Without this, Sheets turns "0001" into 1 and "09171234567"
 * into 9171234567.
 */
function appendRowWithTextColumns_(sheet, values, textColumns) {
  const row = sheet.getLastRow() + 1;

  textColumns.forEach(function(col) {
    sheet.getRange(row, col).setNumberFormat('@');
  });

  sheet.getRange(row, 1, 1, values.length).setValues([
    values.map(function(v, i) {
      return textColumns.indexOf(i + 1) !== -1 ? String(v) : v;
    })
  ]);
}


function dayKey_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}


/**
 * IDs of everyone who has checked in today.
 */
function getTodayPresentIds_() {
  const ids = new Set();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ATTENDANCE_SHEET);

  if (!sheet || sheet.getLastRow() < 2) {
    return ids;
  }

  const today = dayKey_(new Date());
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();

  rows.forEach(function(row) {
    if (row[0] instanceof Date && dayKey_(row[0]) === today) {
      ids.add(normalizeId(row[1]));
    }
  });

  return ids;
}


/**
 * Return today's check-in time for this ID, or null.
 */
function findTodayAttendance_(attendanceSheet, normalizedId) {
  const lastRow = attendanceSheet.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  const today = dayKey_(new Date());
  const rows = attendanceSheet.getRange(2, 1, lastRow - 1, 2).getValues();

  for (let i = 0; i < rows.length; i++) {
    const timestamp = rows[i][0];
    if (
      normalizeId(rows[i][1]) === normalizedId &&
      timestamp instanceof Date &&
      dayKey_(timestamp) === today
    ) {
      return timestamp;
    }
  }

  return null;
}


/**
 * Formatted on the server in the script's time zone, and shown
 * as-is by the page (no browser date parsing).
 */
function formatTimestamp(date) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'MMMM d, yyyy h:mm a'
  );
}
