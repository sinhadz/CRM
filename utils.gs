// ============================================================================
// Utils.gs
// Part 1A.1
// Foundation & Spreadsheet Engine
// Write4U CRM v2.0 Enterprise
// ============================================================================

'use strict';

// ============================================================================
// GLOBAL CACHE
// ============================================================================

const _CACHE = {
  spreadsheet: null,
  sheets: {},
  headers: {},
  metadata: {}
};

// ============================================================================
// SPREADSHEET ENGINE
// ============================================================================

/**
 * Returns Spreadsheet instance.
 * Singleton.
 *
 * @returns {Spreadsheet}
 */
function getSpreadsheet() {

  if (_CACHE.spreadsheet) {
    return _CACHE.spreadsheet;
  }

  _CACHE.spreadsheet =
    SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  return _CACHE.spreadsheet;

}

/**
 * Returns Spreadsheet Timezone.
 *
 * @returns {string}
 */
function getSpreadsheetTimeZone() {

  return getSpreadsheet().getSpreadsheetTimeZone();

}

/**
 * Returns Spreadsheet Name.
 *
 * @returns {string}
 */
function getSpreadsheetName() {

  return getSpreadsheet().getName();

}

/**
 * Returns Spreadsheet URL.
 *
 * @returns {string}
 */
function getSpreadsheetUrl() {

  return getSpreadsheet().getUrl();

}

/**
 * Flush pending writes.
 */
function flushSpreadsheet() {

  SpreadsheetApp.flush();

}

/**
 * Clears all in-memory cache.
 */
function clearMemoryCache() {

  _CACHE.spreadsheet = null;
  _CACHE.sheets = {};
  _CACHE.headers = {};
  _CACHE.metadata = {};

}

/**
 * Refresh spreadsheet cache.
 */
function reloadSpreadsheet() {

  clearMemoryCache();

  return getSpreadsheet();

}

// ============================================================================
// SHEET ENGINE
// ============================================================================

/**
 * Returns Sheet object.
 *
 * Cached automatically.
 *
 * @param {string} sheetName
 * @returns {Sheet}
 */
function getSheet(sheetName) {

  sheetName = safeString(sheetName);

  if (_CACHE.sheets[sheetName]) {
    return _CACHE.sheets[sheetName];
  }

  const sheet =
    getSpreadsheet().getSheetByName(sheetName);

  if (!sheet) {

    throw new Error(
      "Sheet not found : " + sheetName
    );

  }

  _CACHE.sheets[sheetName] = sheet;

  return sheet;

}

/**
 * Safe sheet getter.
 *
 * Returns null.
 *
 * Never throws.
 *
 * @param {string} sheetName
 * @returns {Sheet|null}
 */
function getSheetSafe(sheetName) {

  try {

    return getSheet(sheetName);

  } catch (e) {

    return null;

  }

}

/**
 * Returns TRUE if sheet exists.
 *
 * @param {string} sheetName
 * @returns {boolean}
 */
function sheetExists(sheetName) {

  return getSheetSafe(sheetName) !== null;

}

/**
 * Creates sheet if missing.
 *
 * @param {string} sheetName
 * @returns {Sheet}
 */
function createSheetIfMissing(sheetName) {

  const existing = getSheetSafe(sheetName);

  if (existing) {
    return existing;
  }

  const sheet =
    getSpreadsheet().insertSheet(sheetName);

  _CACHE.sheets[sheetName] = sheet;

  return sheet;

}

/**
 * Deletes sheet.
 *
 * @param {string} sheetName
 */
function deleteSheet(sheetName) {

  const sheet = getSheetSafe(sheetName);

  if (!sheet) return;

  getSpreadsheet().deleteSheet(sheet);

  delete _CACHE.sheets[sheetName];
  delete _CACHE.headers[sheetName];
  delete _CACHE.metadata[sheetName];

}
// ============================================================================
// Utils.gs
// Part 1A.2
// Metadata Engine
// ============================================================================

/**
 * Returns total rows.
 *
 * Cached.
 *
 * @param {string} sheetName
 * @returns {number}
 */
function getLastRow(sheetName) {

  const meta = getSheetMetadata(sheetName);

  return meta.lastRow;

}

/**
 * Returns total columns.
 *
 * Cached.
 *
 * @param {string} sheetName
 * @returns {number}
 */
function getLastColumn(sheetName) {

  const meta = getSheetMetadata(sheetName);

  return meta.lastColumn;

}

/**
 * Returns sheet metadata.
 *
 * Metadata is cached in memory.
 *
 * {
 *   name,
 *   sheet,
 *   lastRow,
 *   lastColumn,
 *   frozenRows,
 *   frozenColumns,
 *   maxRows,
 *   maxColumns
 * }
 *
 * @param {string} sheetName
 * @returns {Object}
 */
function getSheetMetadata(sheetName) {

  sheetName = safeString(sheetName);

  if (_CACHE.metadata[sheetName]) {
    return _CACHE.metadata[sheetName];
  }

  const sheet = getSheet(sheetName);

  const meta = {

    name: sheetName,

    sheet: sheet,

    lastRow: sheet.getLastRow(),

    lastColumn: sheet.getLastColumn(),

    frozenRows: sheet.getFrozenRows(),

    frozenColumns: sheet.getFrozenColumns(),

    maxRows: sheet.getMaxRows(),

    maxColumns: sheet.getMaxColumns()

  };

  _CACHE.metadata[sheetName] = meta;

  return meta;

}

/**
 * Refresh metadata cache.
 *
 * @param {string} sheetName
 */
function refreshSheetMetadata(sheetName) {

  delete _CACHE.metadata[sheetName];

  return getSheetMetadata(sheetName);

}

/**
 * Clears metadata cache.
 */
function clearMetadataCache() {

  _CACHE.metadata = {};

}

/**
 * Returns header row.
 *
 * Cached in memory.
 *
 * @param {string} sheetName
 * @returns {Array}
 */
function getHeaders(sheetName) {

  sheetName = safeString(sheetName);

  if (_CACHE.headers[sheetName]) {

    return _CACHE.headers[sheetName];

  }

  const headers = getSheet(sheetName)

    .getRange(

      1,

      1,

      1,

      getLastColumn(sheetName)

    )

    .getValues()[0];

  _CACHE.headers[sheetName] = headers;

  return headers;

}

/**
 * Refresh header cache.
 *
 * @param {string} sheetName
 * @returns {Array}
 */
function refreshHeaders(sheetName) {

  delete _CACHE.headers[sheetName];

  return getHeaders(sheetName);

}

/**
 * Clears all header cache.
 */
function clearHeaderCache() {

  _CACHE.headers = {};

}

/**
 * Returns
 *
 * {
 *   Header : Index
 * }
 *
 * Cached.
 *
 * @param {string} sheetName
 * @returns {Object}
 */
function getHeaderMap(sheetName) {

  sheetName = safeString(sheetName);

  const cacheKey = "__MAP__" + sheetName;

  if (_CACHE.headers[cacheKey]) {

    return _CACHE.headers[cacheKey];

  }

  const headers = getHeaders(sheetName);

  const map = {};

  headers.forEach(function(header, index) {

    map[safeString(header)] = index;

  });

  _CACHE.headers[cacheKey] = map;

  return map;

}

/**
 * Refresh header map.
 *
 * @param {string} sheetName
 * @returns {Object}
 */
function refreshHeaderMap(sheetName) {

  delete _CACHE.headers["__MAP__" + sheetName];

  return getHeaderMap(sheetName);

}

/**
 * Returns TRUE if column exists.
 *
 * @param {string} sheetName
 * @param {string} column
 * @returns {boolean}
 */
function hasColumn(sheetName, column) {

  const H = getHeaderMap(sheetName);

  return H.hasOwnProperty(column);

}
// ============================================================================
// Utils.gs
// Part 1A.3
// Data Engine
// ============================================================================

/**
 * Returns entire sheet data.
 *
 * @param {string} sheetName
 * @returns {Array}
 */
function getSheetData(sheetName) {

  const sheet = getSheet(sheetName);

  return sheet
    .getDataRange()
    .getValues();

}

/**
 * Returns values from any range.
 *
 * @param {string} sheetName
 * @param {number} row
 * @param {number} column
 * @param {number} numRows
 * @param {number} numColumns
 * @returns {Array}
 */
function getRangeValues(sheetName, row, column, numRows, numColumns) {

  return getSheet(sheetName)
    .getRange(row, column, numRows, numColumns)
    .getValues();

}

/**
 * Batch write.
 *
 * @param {string} sheetName
 * @param {number} row
 * @param {number} column
 * @param {Array} values
 */
function setRangeValues(sheetName, row, column, values) {

  if (!values || values.length === 0) return;

  getSheet(sheetName)
    .getRange(
      row,
      column,
      values.length,
      values[0].length
    )
    .setValues(values);

  refreshSheetMetadata(sheetName);

}

/**
 * Returns entire row.
 *
 * @param {string} sheetName
 * @param {number} row
 * @returns {Array}
 */
function getRow(sheetName, row) {

  return getSheet(sheetName)
    .getRange(
      row,
      1,
      1,
      getLastColumn(sheetName)
    )
    .getValues()[0];

}

/**
 * Returns row as object.
 *
 * Example:
 *
 * {
 *   "Employee ID":"Employee 1",
 *   "PIN":"1234"
 * }
 *
 * @param {string} sheetName
 * @param {number} row
 * @returns {Object}
 */
function getRowObject(sheetName, row) {

  const values = getRow(sheetName, row);
  const H = getHeaderMap(sheetName);

  const obj = {};

  Object.keys(H).forEach(function (key) {

    obj[key] = values[H[key]];

  });

  return obj;

}

/**
 * Updates complete row.
 *
 * @param {string} sheetName
 * @param {number} row
 * @param {Array} values
 */
function updateRow(sheetName, row, values) {

  getSheet(sheetName)
    .getRange(
      row,
      1,
      1,
      values.length
    )
    .setValues([values]);

}

/**
 * Appends new row.
 *
 * @param {string} sheetName
 * @param {Array} values
 */
function appendRowData(sheetName, values) {

  getSheet(sheetName)
    .appendRow(values);

  refreshSheetMetadata(sheetName);

}

/**
 * Deletes row.
 *
 * @param {string} sheetName
 * @param {number} row
 */
function deleteRowData(sheetName, row) {

  getSheet(sheetName)
    .deleteRow(row);

  refreshSheetMetadata(sheetName);

}

/**
 * Returns single cell.
 *
 * @param {string} sheetName
 * @param {number} row
 * @param {number} column
 * @returns {*}
 */
function getCell(sheetName, row, column) {

  return getSheet(sheetName)
    .getRange(row, column)
    .getValue();

}

/**
 * Sets single cell.
 *
 * @param {string} sheetName
 * @param {number} row
 * @param {number} column
 * @param {*} value
 */
function setCell(sheetName, row, column, value) {

  getSheet(sheetName)
    .getRange(row, column)
    .setValue(value);

}

/**
 * Updates multiple individual cells.
 *
 * Usage:
 *
 * updateCells(sheet,{
 *   A1:123,
 *   B4:"Done",
 *   C8:new Date()
 * });
 *
 * @param {string} sheetName
 * @param {Object} updates
 */
function updateCells(sheetName, updates) {

  const sheet = getSheet(sheetName);

  Object.keys(updates).forEach(function (cell) {

    sheet
      .getRange(cell)
      .setValue(updates[cell]);

  });

}

/**
 * Clears row.
 *
 * @param {string} sheetName
 * @param {number} row
 */
function clearRow(sheetName, row) {

  getSheet(sheetName)
    .getRange(
      row,
      1,
      1,
      getLastColumn(sheetName)
    )
    .clearContent();

}

/**
 * Clears sheet data except header.
 *
 * @param {string} sheetName
 */
function clearSheetData(sheetName) {

  const sheet = getSheet(sheetName);

  if (sheet.getLastRow() <= 1) return;

  sheet
    .getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      sheet.getLastColumn()
    )
    .clearContent();

  refreshSheetMetadata(sheetName);

}

/**
 * Finds first matching row.
 *
 * Generic search engine.
 *
 * @param {string} sheetName
 * @param {string} columnName
 * @param {*} value
 * @returns {number}
 */
function findRow(sheetName, columnName, value) {

  const H = getHeaderMap(sheetName);

  if (!(columnName in H)) {

    throw new Error(
      "Column not found : " + columnName
    );

  }

  const index = H[columnName];

  const data = getSheetData(sheetName);

  value = safeString(value);

  for (let i = 1; i < data.length; i++) {

    if (safeString(data[i][index]) === value) {

      return i + 1;

    }

  }

  return -1;

}
```javascript
// ============================================================================
// Utils.gs
// Part 2.1
// Validation & Response Engine
// ============================================================================

/**
 * Standard Success Response.
 *
 * @param {*} data
 * @param {string=} message
 * @returns {Object}
 */
function ok(data, message) {

  return {
    success: true,
    message: message || "Success",
    data: data || null,
    transactionId: generateTransactionId(),
    timestamp: new Date()
  };

}

/**
 * Standard Failure Response.
 *
 * @param {string} message
 * @param {*=} data
 * @returns {Object}
 */
function fail(message, data) {

  return {
    success: false,
    message: message || MESSAGE.SERVER_ERROR,
    data: data || null,
    transactionId: generateTransactionId(),
    timestamp: new Date()
  };

}

/**
 * Execute safely.
 *
 * Logs errors automatically.
 *
 * @param {Function} callback
 * @returns {Object}
 */
function executeSafe(callback) {

  try {

    return callback();

  } catch (error) {

    if (typeof logError === "function") {

      logError({
        functionName: callback.name || "Anonymous",
        message: error.message,
        stack: error.stack
      });

    } else {

      Logger.log(error);

    }

    return fail(error.message);

  }

}

/**
 * Executes with retry.
 *
 * Exponential Backoff
 */
function retry(callback, attempts) {

  attempts = attempts || 3;

  let delay = 300;

  for (let i = 0; i < attempts; i++) {

    try {

      return callback();

    } catch (err) {

      if (i === attempts - 1) {

        throw err;

      }

      Utilities.sleep(delay);

      delay *= 2;

    }

  }

}

/**
 * Require value.
 */
function requireValue(value, field) {

  if (
    value === null ||
    value === undefined ||
    String(value).trim() === ""
  ) {

    throw new Error(field + " is required.");

  }

}

/**
 * Require one of values.
 */
function requireOneOf(value, list, field) {

  if (list.indexOf(value) === -1) {

    throw new Error(
      field +
      " contains an invalid value."
    );

  }

}

/**
 * Validate email.
 */
function validateEmail(email) {

  email = safeString(email);

  const regex =
    /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  if (!regex.test(email)) {

    throw new Error(
      "Invalid email address."
    );

  }

  return true;

}

/**
 * Validate PIN.
 */
function validatePIN(pin) {

  pin = safeString(pin);

  if (!/^\d{4}$/.test(pin)) {

    throw new Error(
      "PIN must contain exactly 4 digits."
    );

  }

  return true;

}

/**
 * Validate Role.
 */
function validateRole(role) {

  requireOneOf(

    role,

    [
      ROLE.ADMIN,
      ROLE.SUPERVISOR,
      ROLE.EMPLOYEE
    ],

    "Role"

  );

}

/**
 * Validate Employee Status.
 */
function validateEmployeeStatus(status) {

  requireOneOf(

    status,

    [
      EMPLOYEE_STATUS.ACTIVE,
      EMPLOYEE_STATUS.INACTIVE
    ],

    "Employee Status"

  );

}

/**
 * Validate Lead Status.
 */
function validateLeadStatus(status) {

  requireOneOf(

    status,

    [

      LEAD_STATUS.PENDING,

      LEAD_STATUS.FOLLOW_UP,

      LEAD_STATUS.BOOKED_HOME,

      LEAD_STATUS.BOOKED_WAREHOUSE,

      LEAD_STATUS.DUPLICATE,

      LEAD_STATUS.INCORRECT_NUMBER,

      LEAD_STATUS.NOT_INTERESTED,

      LEAD_STATUS.UNATTENDED

    ],

    "Lead Status"

  );

}

/**
 * Validate Date.
 */
function validateDate(value, field) {

  if (!(value instanceof Date)) {

    throw new Error(
      field + " must be a valid date."
    );

  }

}

/**
 * Validate Required Columns.
 */
function validateColumns(sheetName, columns) {

  const H = getHeaderMap(sheetName);

  columns.forEach(function(col){

    if (!(col in H)) {

      throw new Error(

        "Missing required column : " + col

      );

    }

  });

}

/**
 * Validate Sheet Exists.
 */
function validateSheet(sheetName) {

  if (!sheetExists(sheetName)) {

    throw new Error(

      "Missing sheet : " + sheetName

    );

  }

}

/**
 * Validate Employee Exists.
 */
function validateEmployee(employeeId) {

  if (findEmployeeRow(employeeId) === -1) {

    throw new Error(

      "Employee not found."

    );

  }

}

/**
 * Validate Lead Exists.
 */
function validateLead(leadNumber) {

  if (findLeadRow(leadNumber) === -1) {

    throw new Error(

      "Lead not found."

    );

  }

}

/**
 * Validate Assignment.
 */
function validateLeadOwnership(employeeId, leadNumber) {

  if (

    !isLeadAssignedToEmployee(

      employeeId,

      leadNumber

    )

  ) {

    throw new Error(

      "Access denied."

    );

  }

}
// ============================================================================
// Utils.gs
// Part 2.2
// Sanitization, Conversion & Safe Helpers
// ============================================================================

/**
 * Safe String
 *
 * @param {*} value
 * @returns {string}
 */
function safeString(value) {

  if (value === null || value === undefined) {
    return "";
  }

  return String(value).trim();

}

/**
 * Safe Number
 *
 * @param {*} value
 * @returns {number}
 */
function safeNumber(value) {

  const number = Number(value);

  return isNaN(number) ? 0 : number;

}

/**
 * Safe Boolean
 *
 * @param {*} value
 * @returns {boolean}
 */
function safeBoolean(value) {

  if (value === true) return true;

  const v = safeString(value).toLowerCase();

  return (
    v === "true" ||
    v === "yes" ||
    v === "1"
  );

}

/**
 * Safe Date
 *
 * @param {*} value
 * @returns {Date|null}
 */
function safeDate(value) {

  if (!value) return null;

  if (value instanceof Date) {

    if (!isNaN(value.getTime())) {
      return value;
    }

  }

  const d = new Date(value);

  return isNaN(d.getTime()) ? null : d;

}

/**
 * Null → Empty
 */
function nullToEmpty(value) {

  return value === null || value === undefined
    ? ""
    : value;

}

/**
 * Empty Check
 */
function isEmpty(value) {

  return (
    value === "" ||
    value === null ||
    value === undefined
  );

}

/**
 * Not Empty
 */
function isNotEmpty(value) {

  return !isEmpty(value);

}

/**
 * Trim Object
 */
function trimObject(obj) {

  if (!obj) return {};

  const out = {};

  Object.keys(obj).forEach(function(key){

    const value = obj[key];

    out[key] =

      typeof value === "string"

      ? value.trim()

      : value;

  });

  return out;

}

/**
 * HTML Escape
 */
function escapeHTML(value){

  value = safeString(value);

  return value

    .replace(/&/g,"&amp;")

    .replace(/</g,"&lt;")

    .replace(/>/g,"&gt;")

    .replace(/"/g,"&quot;")

    .replace(/'/g,"&#39;");

}

/**
 * Strip HTML
 */
function stripHTML(value){

  return safeString(value)

      .replace(/<[^>]*>/g,"")

      .trim();

}

/**
 * Sanitize User Input
 */
function sanitize(value){

  value = stripHTML(value);

  value = value.replace(/[\r\n\t]/g," ");

  return value.trim();

}

/**
 * Sanitize Object
 */
function sanitizeObject(obj){

  if(!obj) return {};

  const clean = {};

  Object.keys(obj).forEach(function(key){

    const value = obj[key];

    clean[key] =

      typeof value === "string"

      ? sanitize(value)

      : value;

  });

  return clean;

}

/**
 * Normalize Employee ID
 *
 * Examples:
 *
 * 1
 * Employee1
 * employee 1
 *
 * =>
 *
 * Employee 1
 */
function normalizeEmployeeId(employeeId){

  employeeId = safeString(employeeId)

      .replace(/^employee/i,"")

      .trim();

  if(employeeId===""){

    return "";

  }

  return "Employee " + employeeId;

}

/**
 * Normalize Lead Number
 */
function normalizeLeadNumber(lead){

  return safeString(lead)

      .toUpperCase();

}

/**
 * Uppercase
 */
function upper(value){

  return safeString(value)

      .toUpperCase();

}

/**
 * Lowercase
 */
function lower(value){

  return safeString(value)

      .toLowerCase();

}

/**
 * Title Case
 */
function title(value){

  return safeString(value)

      .toLowerCase()

      .replace(/\b\w/g,function(x){

        return x.toUpperCase();

      });

}

/**
 * Left Pad
 */
function leftPad(value,length,char){

  char = char || "0";

  value = safeString(value);

  while(value.length < length){

    value = char + value;

  }

  return value;

}

/**
 * Right Pad
 */
function rightPad(value,length,char){

  char = char || " ";

  value = safeString(value);

  while(value.length < length){

    value += char;

  }

  return value;

}

/**
 * Deep Clone
 *
 * Preserves Dates.
 */
function deepClone(obj){

  if(obj===null) return null;

  if(obj instanceof Date){

    return new Date(obj.getTime());

  }

  if(Array.isArray(obj)){

    return obj.map(deepClone);

  }

  if(typeof obj==="object"){

    const copy={};

    Object.keys(obj).forEach(function(key){

      copy[key]=deepClone(obj[key]);

    });

    return copy;

  }

  return obj;

}

/**
 * Freeze Object
 */
function freeze(obj){

  return Object.freeze(obj);

}

/**
 * Merge Objects
 */
function merge(){

  const result={};

  Array.prototype.forEach.call(arguments,function(obj){

    if(!obj) return;

    Object.keys(obj).forEach(function(key){

      result[key]=obj[key];

    });

  });

  return result;

}

/**
 * Object Keys Count
 */
function objectSize(obj){

  return Object.keys(obj||{}).length;

}
// ============================================================================
// Utils.gs
// Part 2.3
// Enterprise Engine (Cache, Lock, Retry, Profiling)
// ============================================================================

// ============================================================================
// SCRIPT CACHE
// ============================================================================

/**
 * Returns Script Cache.
 *
 * @returns {GoogleAppsScript.Cache.Cache}
 */
function getCache() {
  return CacheService.getScriptCache();
}

/**
 * Returns User Cache.
 *
 * @returns {GoogleAppsScript.Cache.Cache}
 */
function getUserCache() {
  return CacheService.getUserCache();
}

/**
 * Returns Document Cache.
 *
 * @returns {GoogleAppsScript.Cache.Cache}
 */
function getDocumentCache() {
  return CacheService.getDocumentCache();
}

/**
 * Cache Put
 */
function cachePut(key, value, seconds) {

  if (value === undefined || value === null) return;

  if (typeof value === "object") {
    value = JSON.stringify(value);
  }

  getCache().put(
    safeString(key),
    String(value),
    seconds || SESSION.CACHE_SECONDS
  );

}

/**
 * Cache Get
 */
function cacheGet(key) {

  const value = getCache().get(safeString(key));

  if (value === null) return null;

  try {
    return JSON.parse(value);
  } catch (e) {
    return value;
  }

}

/**
 * Cache Remove
 */
function cacheRemove(key) {

  getCache().remove(safeString(key));

}

/**
 * Cache Exists
 */
function cacheExists(key) {

  return getCache().get(key) !== null;

}

/**
 * Clear Cache
 */
function clearCache(keys) {

  if (!Array.isArray(keys)) return;

  getCache().removeAll(keys);

}

// ============================================================================
// PROPERTIES
// ============================================================================

function getScriptProperties() {

  return PropertiesService.getScriptProperties();

}

function getProperty(key) {

  return getScriptProperties()

    .getProperty(key);

}

function setProperty(key, value) {

  getScriptProperties()

    .setProperty(

      safeString(key),

      safeString(value)

    );

}

function deleteProperty(key) {

  getScriptProperties()

    .deleteProperty(key);

}

function getProperties() {

  return getScriptProperties()

    .getProperties();

}

// ============================================================================
// LOCK ENGINE
// ============================================================================

/**
 * Execute inside Script Lock.
 *
 * @param {Function} callback
 * @param {number=} timeout
 */
function withLock(callback, timeout) {

  timeout = timeout || 30000;

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(timeout)) {

    throw new Error(
      "Unable to obtain lock."
    );

  }

  try {

    return callback();

  }

  finally {

    lock.releaseLock();

  }

}

/**
 * Execute inside Document Lock.
 */
function withDocumentLock(callback, timeout) {

  timeout = timeout || 30000;

  const lock = LockService.getDocumentLock();

  lock.waitLock(timeout);

  try {

    return callback();

  }

  finally {

    lock.releaseLock();

  }

}

// ============================================================================
// PERFORMANCE TIMER
// ============================================================================

/**
 * Start Timer
 *
 * const sw=startTimer();
 */
function startTimer() {

  return Date.now();

}

/**
 * Stop Timer
 */
function stopTimer(start) {

  return Date.now() - start;

}

/**
 * Profile Function
 */
function profile(name, callback) {

  const start = startTimer();

  const result = callback();

  const elapsed = stopTimer(start);

  if (typeof logInfo === "function") {

    logInfo({

      action: "PROFILE",

      operation: name,

      elapsed: elapsed + " ms"

    });

  }

  return result;

}

// ============================================================================
// RETRY ENGINE
// ============================================================================

/**
 * Enterprise Retry
 *
 * Exponential Backoff
 */
function retry(callback, maxRetry) {

  maxRetry = maxRetry || 3;

  let delay = 300;

  let error;

  for (let i = 0; i < maxRetry; i++) {

    try {

      return callback();

    }

    catch (e) {

      error = e;

      Utilities.sleep(delay);

      delay *= 2;

    }

  }

  throw error;

}

// ============================================================================
// TRANSACTION IDS
// ============================================================================

function generateUUID() {

  return Utilities.getUuid();

}

function generateTransactionId() {

  return Utilities.getUuid()

    .replace(/-/g, "")

    .substring(0, 16)

    .toUpperCase();

}

function generateRequestId() {

  return "REQ-" +

    generateTransactionId();

}

function generateAuditId() {

  return "AUDIT-" +

    generateTransactionId();

}

function generateLockId() {

  return "LOCK-" +

    generateTransactionId();

}

function generateSessionId() {

  return "SESSION-" +

    generateTransactionId();

}

// ============================================================================
// MEMORY CACHE
// ============================================================================

/**
 * Clear Everything
 */
function clearAllCache() {

  clearMemoryCache();

}

/**
 * Refresh Everything
 */
function refreshApplicationCache() {

  clearMemoryCache();

  Object.keys(CONFIG.SHEETS)

    .forEach(function(key){

      const sheetName = CONFIG.SHEETS[key];

      if(sheetExists(sheetName)){

        refreshSheetMetadata(sheetName);

        refreshHeaders(sheetName);

        refreshHeaderMap(sheetName);

      }

    });

}

/**
 * Health Check
 */
function systemHealth() {

  return ok({

    application: APP.NAME,

    version: APP.VERSION,

    build: APP.BUILD,

    timezone: getSpreadsheetTimeZone(),

    spreadsheet: getSpreadsheetName(),

    timestamp: new Date()

  });

}
// ============================================================================
// Utils.gs
// Part 2.4
// Generic Lookup Engine & Search Utilities
// ============================================================================

// ============================================================================
// GENERIC ROW FINDER
// ============================================================================

/**
 * Generic row finder.
 *
 * @param {string} sheetName
 * @param {string} columnName
 * @param {*} value
 * @param {boolean=} ignoreCase
 * @returns {number}
 */
function findRow(sheetName, columnName, value, ignoreCase) {

  ignoreCase = ignoreCase !== false;

  const data = getSheetData(sheetName);
  const H = getHeaderMap(sheetName);

  if (!(columnName in H)) {

    throw new Error(
      "Column not found : " + columnName
    );

  }

  const col = H[columnName];

  value = safeString(value);

  if (ignoreCase) {

    value = value.toLowerCase();

  }

  for (let i = 1; i < data.length; i++) {

    let cell = safeString(data[i][col]);

    if (ignoreCase) {

      cell = cell.toLowerCase();

    }

    if (cell === value) {

      return i + 1;

    }

  }

  return -1;

}

// ============================================================================
// MULTIPLE ROW FINDER
// ============================================================================

function findRows(sheetName, columnName, value) {

  const rows = [];

  const data = getSheetData(sheetName);

  const H = getHeaderMap(sheetName);

  const col = H[columnName];

  value = safeString(value);

  for (let i = 1; i < data.length; i++) {

    if (

      safeString(data[i][col]) === value

    ) {

      rows.push(i + 1);

    }

  }

  return rows;

}

// ============================================================================
// OBJECT FINDER
// ============================================================================

function findObject(sheetName, columnName, value) {

  const row = findRow(

    sheetName,

    columnName,

    value

  );

  if (row === -1) {

    return null;

  }

  return getRowObject(

    sheetName,

    row

  );

}

// ============================================================================
// OBJECT LIST
// ============================================================================

function getObjects(sheetName) {

  const data = getSheetData(sheetName);

  const H = getHeaderMap(sheetName);

  const list = [];

  for (let i = 1; i < data.length; i++) {

    const obj = {};

    Object.keys(H).forEach(function(key){

      obj[key] = data[i][H[key]];

    });

    list.push(obj);

  }

  return list;

}

// ============================================================================
// FILTER OBJECTS
// ============================================================================

function filterObjects(sheetName, callback){

  return getObjects(sheetName)

      .filter(callback);

}

// ============================================================================
// SORT OBJECTS
// ============================================================================

function sortObjects(list, field, ascending){

  ascending = ascending !== false;

  return list.sort(function(a,b){

    if(a[field] < b[field]){

      return ascending ? -1 : 1;

    }

    if(a[field] > b[field]){

      return ascending ? 1 : -1;

    }

    return 0;

  });

}

// ============================================================================
// UNIQUE VALUES
// ============================================================================

function getUniqueValues(sheetName,columnName){

  const H = getHeaderMap(sheetName);

  const col = H[columnName];

  const values = getSheetData(sheetName)

      .slice(1)

      .map(function(r){

        return safeString(r[col]);

      });

  return [...new Set(values)];

}

// ============================================================================
// EXISTS
// ============================================================================

function valueExists(sheetName,columnName,value){

  return findRow(

      sheetName,

      columnName,

      value

  ) !== -1;

}

// ============================================================================
// COUNT
// ============================================================================

function countWhere(sheetName,columnName,value){

  return findRows(

      sheetName,

      columnName,

      value

  ).length;

}

// ============================================================================
// INDEX BUILDER
// ============================================================================

/**
 * Creates dictionary.
 *
 * {
 * Lead Number :
 * row
 * }
 */
function buildIndex(sheetName,columnName){

  const H = getHeaderMap(sheetName);

  const col = H[columnName];

  const data = getSheetData(sheetName);

  const index = {};

  for(let i=1;i<data.length;i++){

    index[

      safeString(

        data[i][col]

      )

    ] = i+1;

  }

  return index;

}

// ============================================================================
// EMPLOYEE CACHE
// ============================================================================

let EMPLOYEE_INDEX = null;

function getEmployeeIndex(){

  if(EMPLOYEE_INDEX){

    return EMPLOYEE_INDEX;

  }

  EMPLOYEE_INDEX = buildIndex(

      CONFIG.SHEETS.EMPLOYEES,

      "Employee ID"

  );

  return EMPLOYEE_INDEX;

}

// ============================================================================
// LEAD CACHE
// ============================================================================

let LEAD_INDEX = null;

function getLeadIndex(){

  if(LEAD_INDEX){

    return LEAD_INDEX;

  }

  LEAD_INDEX = buildIndex(

      CONFIG.SHEETS.LEADS,

      "Lead Number"

  );

  return LEAD_INDEX;

}

// ============================================================================
// REFRESH INDEXES
// ============================================================================

function refreshIndexes(){

  EMPLOYEE_INDEX = null;

  LEAD_INDEX = null;

}

// ============================================================================
// FAST EMPLOYEE ROW
// ============================================================================

function findEmployeeRow(employeeId){

  employeeId = normalizeEmployeeId(employeeId);

  const index = getEmployeeIndex();

  return index[employeeId] || -1;

}

// ============================================================================
// FAST LEAD ROW
// ============================================================================

function findLeadRow(leadNumber){

  const index = getLeadIndex();

  return index[safeString(leadNumber)] || -1;

}

// ============================================================================
// GET EMPLOYEE
// ============================================================================

function getEmployee(employeeId){

  const row = findEmployeeRow(employeeId);

  if(row===-1){

    return null;

  }

  return getRowObject(

      CONFIG.SHEETS.EMPLOYEES,

      row

  );

}

// ============================================================================
// GET LEAD
// ============================================================================

function getLead(leadNumber){

  const row = findLeadRow(leadNumber);

  if(row===-1){

    return null;

  }

  return getRowObject(

      CONFIG.SHEETS.LEADS,

      row

  );

}
// ============================================================================
// Utils.gs
// Part 2.5 (FINAL)
// Status, Roles, Date Engine & Startup Validation
// ============================================================================

// ============================================================================
// ROLE HELPERS
// ============================================================================

function getEmployeeRole(employeeId) {

  const employee = getEmployee(employeeId);

  return employee

    ? safeString(employee.Role)

    : ROLE.EMPLOYEE;

}

function isAdmin(employeeId) {

  return getEmployeeRole(employeeId) === ROLE.ADMIN;

}

function isSupervisor(employeeId) {

  return getEmployeeRole(employeeId) === ROLE.SUPERVISOR;

}

function isEmployee(employeeId) {

  return getEmployeeRole(employeeId) === ROLE.EMPLOYEE;

}

function hasRole(employeeId, roles) {

  if (!Array.isArray(roles)) {

    roles = [roles];

  }

  return roles.indexOf(

    getEmployeeRole(employeeId)

  ) !== -1;

}

// ============================================================================
// EMPLOYEE HELPERS
// ============================================================================

function employeeExists(employeeId) {

  return findEmployeeRow(employeeId) !== -1;

}

function employeeActive(employeeId) {

  const emp = getEmployee(employeeId);

  if (!emp) return false;

  return safeString(emp.Status) === EMPLOYEE_STATUS.ACTIVE;

}

function employeeOnline(employeeId) {

  const emp = getEmployee(employeeId);

  if (!emp) return false;

  return safeString(emp["Login Status"]) === LOGIN_STATUS.ONLINE;

}

// ============================================================================
// LEAD HELPERS
// ============================================================================

function leadExists(leadNumber) {

  return findLeadRow(leadNumber) !== -1;

}

function isLeadAssignedToEmployee(employeeId, leadNumber) {

  const lead = getLead(leadNumber);

  if (!lead) return false;

  return normalizeEmployeeId(

    lead["Assigned To"]

  ) === normalizeEmployeeId(employeeId);

}

// ============================================================================
// STATUS HELPERS
// ============================================================================

function isPendingStatus(status) {

  status = safeString(status);

  return status === "" ||

         status === LEAD_STATUS.PENDING;

}

function isFollowUpStatus(status) {

  return safeString(status)

      === LEAD_STATUS.FOLLOW_UP;

}

function isBookingStatus(status) {

  return STATUS_GROUP.BOOKED

      .indexOf(status) !== -1;

}

function isDisposedStatus(status) {

  return STATUS_GROUP.DISPOSED

      .indexOf(status) !== -1;

}

// ============================================================================
// DATE ENGINE
// ============================================================================

function now() {

  return new Date();

}

function today() {

  const d = new Date();

  d.setHours(0,0,0,0);

  return d;

}

function addMinutes(date, minutes) {

  const d = new Date(date);

  d.setMinutes(

    d.getMinutes() + minutes

  );

  return d;

}

function addHours(date, hours) {

  return addMinutes(

    date,

    hours * 60

  );

}

function diffMinutes(start, end) {

  return Math.floor(

    (

      end.getTime() -

      start.getTime()

    ) / 60000

  );

}

function diffSeconds(start, end) {

  return Math.floor(

    (

      end.getTime() -

      start.getTime()

    ) / 1000

  );

}

function sameDate(a, b) {

  return formatDate(a)

      === formatDate(b);

}

// ============================================================================
// BUSINESS HELPERS
// ============================================================================

function requireBooking(status) {

  return isBookingStatus(status);

}

function requireFollowUp(status) {

  return isFollowUpStatus(status);

}

function requireRemarks(status) {

  return !isPendingStatus(status);

}

// ============================================================================
// CACHE INVALIDATION
// ============================================================================

function invalidateEmployeeCache() {

  EMPLOYEE_INDEX = null;

  cacheRemove(CACHE_KEY.EMPLOYEES);

}

function invalidateLeadCache() {

  LEAD_INDEX = null;

  cacheRemove(CACHE_KEY.LEADS);

}

function invalidateHeaderCache(sheetName) {

  refreshHeaders(sheetName);

  refreshHeaderMap(sheetName);

}

function invalidateAllCaches() {

  clearMemoryCache();

  refreshIndexes();

}

// ============================================================================
// STARTUP VALIDATION
// ============================================================================

function validateApplication() {

  Object.keys(CONFIG.SHEETS)

    .forEach(function(key){

      validateSheet(

        CONFIG.SHEETS[key]

      );

    });

  validateColumns(

    CONFIG.SHEETS.EMPLOYEES,

    REQUIRED_EMPLOYEE_COLUMNS

  );

  validateColumns(

    CONFIG.SHEETS.LEADS,

    REQUIRED_LEAD_COLUMNS

  );

  return true;

}

// ============================================================================
// APPLICATION INFO
// ============================================================================

function appInfo() {

  return ok({

    application: APP.NAME,

    version: APP.VERSION,

    build: APP.BUILD,

    company: APP.COMPANY,

    client: APP.CLIENT,

    spreadsheet: getSpreadsheetName(),

    timezone: getSpreadsheetTimeZone(),

    timestamp: new Date()

  });

}

// ============================================================================
// PING
// ============================================================================

function ping() {

  return appInfo();

}

// ============================================================================
// STARTUP
// ============================================================================

(function initializeUtils() {

  try {

    validateApplication();

    Logger.log(

      APP.NAME +

      " " +

      APP.VERSION +

      " initialized."

    );

  }

  catch (e) {

    Logger.log(

      "Initialization failed : " +

      e.message

    );

  }

})();