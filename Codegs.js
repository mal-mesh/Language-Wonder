// @ts-nocheck
/* ==========================================================
    Language Wonder Academy Suite
    Version : v1.0.0 - Bonjour
    File   : code.gs
    Module : Master Backend API Engine
========================================================== */

/* ==========================================================
    SERVER CACHE
    CacheService persists cached values between executions.
    Cache duration: 5 minutes.
========================================================== */

var CACHE_TTL_SECONDS_ = 300;

function getCachedJson_(key) {
    var cache = CacheService.getScriptCache();
    var value = cache.get(key);

    if (!value) {
        return null;
    }

    try {
        return JSON.parse(value);
    } catch (error) {
        Logger.log("Cache parse error for " + key + ": " + error.message);
        cache.remove(key);
        return null;
    }
}

function setCachedJson_(key, value) {
    try {
        var serialized = JSON.stringify(value);

        // Apps Script cache values have a size limit.
        // Do not cache very large datasets.
        if (serialized.length <= 95000) {
            CacheService
                .getScriptCache()
                .put(key, serialized, CACHE_TTL_SECONDS_);
        }
    } catch (error) {
        Logger.log("Cache write error for " + key + ": " + error.message);
    }
}

function invalidateCache_(keys) {
    var cache = CacheService.getScriptCache();

    var cacheKeys = {
        students: "LWA_STUDENTS",
        batches: "LWA_BATCHES"
    };

    var targets = keys || ["students", "batches"];

    targets.forEach(function(key) {
        if (cacheKeys[key]) {
            cache.remove(cacheKeys[key]);
        }
    });
}

/* ==========================================================
    STUDENTS API
========================================================== */
function getStudents() {
    requireAuth_();

    var cacheKey = "LWA_STUDENTS";
    var cachedStudents = getCachedJson_(cacheKey);

    if (cachedStudents !== null) {
        return cachedStudents;
    }

    var sheet = SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName("Students");

    if (!sheet) {
        return [];
    }

    var data = sheet.getDataRange().getDisplayValues();

    if (data.length <= 1) {
        setCachedJson_(cacheKey, []);
        return [];
    }

    data.shift();

    var students = data.map(function(row) {
        return {
            studentId     : row[0] || "",
            studentName   : row[1] || "",
            parentName    : row[2] || "",
            whatsapp      : row[3] || "",
            classType     : row[4] || "",
            batchId       : row[5] || "",
            classFromTime : row[6] || "",
            classToTime   : row[7] || "",
            status        : row[8] || "Active",
            notes         : row[9] || "",
            expectedFee   : row[10] ? Number(row[10]) : 0,
            paidUntil     : row[11] || ""
        };
    });

    setCachedJson_(cacheKey, students);

    return students;
}

function clearLanguageWonderCache() {
    CacheService.getScriptCache().removeAll([
        "LWA_STUDENTS",
        "LWA_BATCHES"
    ]);

    Logger.log("Language Wonder cache cleared.");
}

/* ==========================================================
    BATCHES API
========================================================== */
function getBatches() {
    requireAuth_();

    var cacheKey = "LWA_BATCHES";
    var cachedBatches = getCachedJson_(cacheKey);

    if (cachedBatches !== null) {
        return cachedBatches;
    }

    var sheet = SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName("Batches");

    if (!sheet) {
        return [];
    }

    var data = sheet.getDataRange().getDisplayValues();

    if (data.length <= 1) {
        setCachedJson_(cacheKey, []);
        return [];
    }

    data.shift();

    var batches = data.map(function(row) {
        return {
            batchId   : row[0] || "",
            batchName : row[1] || "",
            days      : row[2] || ""
        };
    });

    setCachedJson_(cacheKey, batches);

    return batches;
}

function doGet() {
  if (!isAuthorized_()) {
    return HtmlService
      .createHtmlOutput(
        '<div style="font-family:sans-serif;padding:40px;text-align:center;color:#374151;">' +
        '<h2>Access Restricted</h2>' +
        '<p>This app is only available to people added as editors on the underlying Google Sheet.</p>' +
        '<p>Ask the administrator to add your Google account under the Sheet\'s Share settings, then reload this page.</p>' +
        '</div>'
      )
      .setTitle("Language Wonder Academy Suite - Access Restricted");
  }

  return HtmlService
    .createTemplateFromFile("index") // Points to lowercase "index" to match sidebar exactly
    .evaluate()
    .setTitle("Language Wonder Academy Suite")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ==========================================================
    SECURITY: ACCESS CONTROL
    isAuthorized_() ties app access to "who can edit the
    underlying Google Sheet" -- the simplest model for a small
    business tool: manage access the same way you already manage
    the spreadsheet, via Share settings. No extra config needed.

    IMPORTANT DEPLOYMENT REQUIREMENT: for Session.getActiveUser()
    to return the real accessing user's email (rather than an
    empty string), this web app must be deployed with:
      Execute as:   User accessing the web app
      Who has access: Anyone with Google account within your org
                       (or "Anyone with a Google account")
    If deployed as "Execute as: Me", getActiveUser() will return
    an empty string for most users and everyone will be denied.
========================================================== */
function isAuthorized_() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (!email) return false; // No verified identity -> deny by default

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var owner = ss.getOwner();
    if (owner && owner.getEmail() === email) return true;

    var editors = ss.getEditors();
    for (var i = 0; i < editors.length; i++) {
      if (editors[i].getEmail() === email) return true;
    }
    return false;
  } catch (e) {
    // Fail closed: any error in the authorization check means "deny".
    return false;
  }
}

// Throws if the caller is not authorized. Call this as the first line of
// every server function that reads or writes student/attendance/fee data,
// since google.script.run can invoke these functions directly -- the
// doGet() gate above only protects the initial page load, not each
// subsequent server call.
function requireAuth_() {
  if (!isAuthorized_()) {
    throw new Error("Not authorized to access this application.");
  }
}

/* ==========================================================
    SECURITY / DATA INTEGRITY: SEQUENTIAL ID GENERATION
    The lock now wraps the entire function including the
    cold-start sheet scan, so two simultaneous first-time
    calls can never both read the same maxNum and generate
    duplicate IDs.
========================================================== */
function nextSequentialId_(sheet, prefix, idColumnIndex, propKey, padLength) {
    // Lock wraps everything including the cold-start scan
    // so two simultaneous calls can never get the same number
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        var props  = PropertiesService.getScriptProperties();
        var stored = props.getProperty(propKey);
        var maxNum = stored ? Number(stored) : 0;

        // Cold start — no counter stored yet
        // Scan the sheet to find the highest existing ID
        // so we never collide with legacy rows
        if (!stored) {
            var data = sheet.getDataRange().getDisplayValues();
            for (var i = 1; i < data.length; i++) {
                var idVal = String(data[i][idColumnIndex] || "");
                var match = idVal.match(/(\d+)$/);
                if (match) {
                    var n = Number(match[1]);
                    if (n > maxNum) maxNum = n;
                }
            }
        }

        var next = maxNum + 1;
        props.setProperty(propKey, String(next));
        return prefix + String(next).padStart(padLength, "0");
    } finally {
        lock.releaseLock();
    }
}

/* ==========================================================
    DATE UTILITIES
    Single source of truth for date formatting across all
    attendance functions. All dates are stored and compared
    in DD/MM/YYYY format to match Google Sheets display.
========================================================== */

function toDisplayDate_(dateInput) {
    if (!dateInput) return "";

    var str = String(dateInput).trim();

    // Already DD/MM/YYYY — return as is
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(str)) {
        return str;
    }

    // YYYY-MM-DD — convert to DD/MM/YYYY
    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
        var parts = str.split("-");
        return parts[2] + "/" + parts[1] + "/" + parts[0];
    }

    // Native Date object
    if (dateInput instanceof Date) {
        return Utilities.formatDate(
            dateInput,
            Session.getScriptTimeZone(),
            "dd/MM/yyyy"
        );
    }

    // Fallback — return as is and log for debugging
    Logger.log("toDisplayDate_: unrecognized format: " + str);
    return str;
}

function toIsoDate_(dateInput) {
    if (!dateInput) return "";

    var str = String(dateInput).trim();

    // Already YYYY-MM-DD — return as is
    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
        return str;
    }

    // DD/MM/YYYY — convert to YYYY-MM-DD
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(str)) {
        var parts = str.split("/");
        return parts[2] + "-" + parts[1] + "-" + parts[0];
    }

    // Native Date object
    if (dateInput instanceof Date) {
        return Utilities.formatDate(
            dateInput,
            Session.getScriptTimeZone(),
            "yyyy-MM-dd"
        );
    }

    // Fallback
    Logger.log("toIsoDate_: unrecognized format: " + str);
    return str;
}

/* ==========================================================
    MONTH UTILITIES
    Used to find the true latest paid month after a payment
    is deleted. Sorts by actual calendar order, not row order.
========================================================== */

var MONTH_ORDER_ = {
    "january"   : 1,
    "february"  : 2,
    "march"     : 3,
    "april"     : 4,
    "may"       : 5,
    "june"      : 6,
    "july"      : 7,
    "august"    : 8,
    "september" : 9,
    "october"   : 10,
    "november"  : 11,
    "december"  : 12
};

function parseMonthYear_(str) {
    // Expects "Month YYYY" e.g. "June 2025"
    if (!str) return null;

    var parts = str.trim().split(" ");
    if (parts.length !== 2) return null;

    var monthNum = MONTH_ORDER_[parts[0].toLowerCase()];
    var year     = Number(parts[1]);

    if (!monthNum || !year) return null;

    return {
        label : str.trim(),
        month : monthNum,
        year  : year
    };
}

function findLatestPaidMonth_(payments) {
    // payments: array of payment objects with a feeMonth string
    // feeMonth may contain multiple months e.g. "June 2025, July 2025"
    // Returns the label of the latest month or "" if none found

    var allMonths = [];

    payments.forEach(function(p) {
        if (!p.feeMonth) return;

        p.feeMonth.split(",").forEach(function(m) {
            var parsed = parseMonthYear_(m.trim());
            if (parsed) {
                allMonths.push(parsed);
            }
        });
    });

    if (allMonths.length === 0) return "";

    // Sort by year then month ascending
    allMonths.sort(function(a, b) {
        if (a.year !== b.year) return a.year - b.year;
        return a.month - b.month;
    });

    // Return the label of the last (latest) month
    return allMonths[allMonths.length - 1].label;
}

/* ==========================================================
    INCLUDE HTML COMPONENTS (WITH AUTOMATIC CASE FALLBACK)
========================================================== */
function include(filename){
  try {
    return HtmlService.createHtmlOutputFromFile(filename.toLowerCase()).getContent();
  } catch (e) {
    try {
      var capitalized = filename.charAt(0).toUpperCase() + filename.slice(1);
      return HtmlService.createHtmlOutputFromFile(capitalized).getContent();
    } catch (err) {
      return "<!-- Template Error: Could not find " + filename + " -->";
    }
  }
}

function saveStudent(student) {
    requireAuth_();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        const studentId = nextSequentialId_(sheet, "LWA", 0, "STUDENT_ID_SEQ", 3);

        sheet.appendRow([
            studentId,
            student.studentName,
            student.parentName,
            student.whatsapp,
            student.classType,
            student.batchId,
            student.classFromTime,
            student.classToTime,
            "Active",
            "", // Notes
            student.expectedFee,
            student.paidUntil
        ]);

        invalidateCache_(['students']); // ← ADDED
        return studentId;
    } finally {
        lock.releaseLock();
    }
}

function updateStudent(studentId, student) {
    requireAuth_();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
    if (!sheet) return false;

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        const data = sheet.getDataRange().getValues();

        for (let i = 1; i < data.length; i++) {
            if (data[i][0] === studentId) {
                sheet.getRange(i + 1, 2, 1, 7).setValues([[
                    student.studentName,
                    student.parentName,
                    student.whatsapp,
                    student.classType,
                    student.batchId,
                    student.classFromTime,
                    student.classToTime
                ]]);

                sheet.getRange(i + 1, 9, 1, 4).setValues([[
                    student.status || data[i][8],
                    data[i][9],
                    student.expectedFee,
                    student.paidUntil
                ]]);

                invalidateCache_(['students']); // ← ADDED
                return true;
            }
        }
        return false;
    } finally {
        lock.releaseLock();
    }
}

/* ==========================================================
    BATCHES API
========================================================== */

function getBatchDays(batchId){
  requireAuth_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Batches");
  if (!sheet) return "";
  
  const data = sheet.getDataRange().getValues();
  data.shift();

  const batch = data.find(row => row[0] == batchId);
  return batch ? batch[2] : "";
}

/* ==========================================================
    DASHBOARD & ATTENDANCE API
========================================================== */
function getDashboardSummary() {
    requireAuth_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Count only Active students (column index 8 = status)
    const studentSheet = ss.getSheetByName("Students");
    let totalStudents = 0;

    if (studentSheet && studentSheet.getLastRow() > 1) {
        const studentData = studentSheet.getDataRange().getDisplayValues();

        // slice(1) skips the header row
        totalStudents = studentData.slice(1).filter(function(row) {
            var status = (row[8] || "Active").trim().toLowerCase();
            return status === "active";
        }).length;
    }

    // Count how many batches have a class today
    const batchSheet = ss.getSheetByName("Batches");
    let todayClasses = 0;

    if (batchSheet) {
        const batchData = batchSheet.getDataRange().getDisplayValues();
        batchData.shift(); // Remove header

        const todayStr = Utilities.formatDate(
            new Date(),
            Session.getScriptTimeZone(),
            "EEEE"
        );

        todayClasses = batchData.filter(function(row) {
            return row[2] && row[2].toString().includes(todayStr);
        }).length;
    }

    // Calculate today's attendance percentage
    const attSheet = ss.getSheetByName("Attendance");
    let attendanceDisplay = "--";

    if (attSheet && attSheet.getLastRow() > 1) {
        const attData = attSheet.getDataRange().getDisplayValues();
        attData.shift(); // Remove header

        const todayDate = Utilities.formatDate(
            new Date(),
            Session.getScriptTimeZone(),
            "dd/MM/yyyy"
        );

        const todaysRecords = attData.filter(function(row) {
            return row[1] === todayDate;
        });

        if (todaysRecords.length > 0) {
            const presentCount = todaysRecords.filter(function(row) {
                return row[6] === "Present";
            }).length;

            attendanceDisplay = Math.round(
                (presentCount / todaysRecords.length) * 100
            ) + "%";
        }
    }

    return {
        totalStudents : totalStudents,
        todayClasses  : todayClasses,
        pendingFees   : "₹0",
        attendance    : attendanceDisplay
    };
}

function saveAttendance(record) {
    requireAuth_();
    const sheet = SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName("Attendance");

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        const attendanceId = nextSequentialId_(
            sheet, "ATT", 0, "ATTENDANCE_ID_SEQ", 4
        );

        // Always store as DD/MM/YYYY
        var formattedDate = toDisplayDate_(record.date);

        sheet.appendRow([
            attendanceId,
            formattedDate,
            record.studentId,
            record.studentName,
            record.batchName,
            record.time,
            record.status,
            record.remarks,
            record.markedBy,
            new Date()
        ]);

        return attendanceId;
    } finally {
        lock.releaseLock();
    }
}

function getTodayAttendance(){
  requireAuth_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  
  data.shift(); // Remove header

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  return data
    .filter(function(row){
      return row[1] === today;
    })
    .map(function(row){
      return {
        attendanceId : row[0],
        date         : row[1],
        studentId    : row[2],
        studentName  : row[3],
        batchName    : row[4],
        time         : row[5],
        status       : row[6],
        remarks      : row[7],
        markedBy     : row[8],
        timestamp    : row[9]
      };
    });
}

/* ==========================================================
    FINANCE & TRANSACTIONS API
========================================================== */
function savePayment(payment) {
    requireAuth_();
    const ss           = SpreadsheetApp.getActiveSpreadsheet();
    const feeSheet     = ss.getSheetByName("Fees");
    const studentSheet = ss.getSheetByName("Students");

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        const paymentId = nextSequentialId_(
            feeSheet, "LWPR", 0, "PAYMENT_ID_SEQ", 4
        );

        const monthsCoveredArray = Array.isArray(payment.monthsCovered)
            ? payment.monthsCovered
            : [payment.monthsCovered];

        feeSheet.appendRow([
            paymentId,
            payment.date,
            payment.studentId,
            payment.studentName,
            monthsCoveredArray.join(", "),
            payment.amount,
            payment.mode,
            payment.remarks || ""
        ]);

        if (payment.monthsCovered) {
            // Use findLatestPaidMonth_ so month order does not matter
            const latestMonth = findLatestPaidMonth_([{
                feeMonth: monthsCoveredArray.join(", ")
            }]);

            if (latestMonth) {
                const studentData = studentSheet.getDataRange().getValues();
                for (let i = 1; i < studentData.length; i++) {
                    if (String(studentData[i][0]).trim() ===
                        String(payment.studentId).trim()) {
                        studentSheet.getRange(i + 1, 12).setValue(latestMonth);
                        break;
                    }
                }
            }
        }

        invalidateCache_(["students"]);
        return paymentId;
    } finally {
        lock.releaseLock();
    }
}

function getPayments() {
  requireAuth_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fees");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  
  data.shift(); // Remove headers

  return data.map(function(row) {
    return {
      paymentId: row[0],
      date: row[1],
      studentId: row[2],
      studentName: row[3],
      feeMonth: row[4],
      amount: row[5],
      mode: row[6],
      remarks: row[7]
    };
  });
}

function deletePaymentRecord(query) {
    requireAuth_();
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const feeSheet     = ss.getSheetByName("Fees");
    const studentSheet = ss.getSheetByName("Students");

    if (!feeSheet || !studentSheet) {
        throw new Error("Required sheets (Fees or Students) not found.");
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        const feeData        = feeSheet.getDataRange().getDisplayValues();
        let   paymentRemoved = false;
        let   targetStudentId = String(query.studentId).trim();

        // ── 1. Delete the matching row from the Fees sheet ──────────────
        for (let i = feeData.length - 1; i >= 1; i--) {
            const row = feeData[i];

            const matchesId = query.paymentId &&
                String(row[0]).trim() === String(query.paymentId).trim();

            const matchesDetails =
                String(row[2]).trim() === targetStudentId &&
                String(row[1]).trim() === String(query.date).trim() &&
                String(row[5]).trim() === String(query.amount).trim();

            if (matchesId || matchesDetails) {
                feeSheet.deleteRow(i + 1);
                paymentRemoved = true;
                break;
            }
        }

        if (!paymentRemoved) {
            throw new Error(
                "Matching transaction ledger record could not be found."
            );
        }

        // ── 2. Recalculate paidUntil using true calendar ordering ────────
        try {
            // Get fresh payment list now that the row is deleted
            const remainingPayments = getPayments().filter(function(p) {
                return String(p.studentId).trim() === targetStudentId;
            });

            // findLatestPaidMonth_ sorts all months properly
            // instead of assuming the last row is the latest
            const newPaidUntil = findLatestPaidMonth_(remainingPayments);

            const studentData = studentSheet.getDataRange().getValues();
            for (let j = 1; j < studentData.length; j++) {
                if (String(studentData[j][0]).trim() === targetStudentId) {
                    studentSheet.getRange(j + 1, 12).setValue(newPaidUntil);
                    break;
                }
            }
        } catch (err) {
            Logger.log("Paid until update warning: " + err.message);
        }

        invalidateCache_(["students"]);
        return true;
    } finally {
        lock.releaseLock();
    }
}

function recordQuickAttendance(data) {
    var lock;
    try {
        requireAuth_();
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("Attendance");

        if (!sheet) {
            sheet = ss.insertSheet("Attendance");
            sheet.appendRow([
                "Attendance ID", "Date", "Student ID", "Student Name",
                "Batch", "Time", "Status", "Remarks", "Marked By", "Timestamp"
            ]);
        }

        lock = LockService.getScriptLock();
        lock.waitLock(30000);

        var attendanceId = nextSequentialId_(
            sheet, "ATT", 0, "ATTENDANCE_ID_SEQ", 4
        );

        var now = new Date();

        var formattedTimestamp = Utilities.formatDate(
            now,
            Session.getScriptTimeZone(),
            "M/d/yyyy H:mm:ss"
        );

        // Always store as DD/MM/YYYY regardless of what format arrives
        var formattedDate = toDisplayDate_(data.date);

        sheet.appendRow([
            attendanceId,
            formattedDate,
            data.studentId  || "",
            data.studentName || "",
            data.batch      || "",
            data.time       || "",
            data.status     || "Present",
            data.remarks    || "",
            Session.getActiveUser().getEmail(), // was hardcoded "Josephine"
            formattedTimestamp
        ]);

        return { success: true, attendanceId: attendanceId };
    } catch (err) {
        return { success: false, error: err.toString() };
    } finally {
        if (lock) lock.releaseLock();
    }
}

function getAttendanceForDate(dateStr) {
    try {
        requireAuth_();
        var sheet = SpreadsheetApp
            .getActiveSpreadsheet()
            .getSheetByName("Attendance");

        if (!sheet) return [];

        var data = sheet.getDataRange().getDisplayValues();
        if (data.length <= 1) return [];

        // Normalize the input date to DD/MM/YYYY
        // so it matches however the date was stored
        var targetDisplay = toDisplayDate_(dateStr);
        var records = [];

        for (var i = 1; i < data.length; i++) {
            // Normalize each sheet row date to DD/MM/YYYY
            var rowDate = toDisplayDate_(data[i][1]);

            if (rowDate === targetDisplay) {
                records.push({
                    attendanceId : data[i][0],
                    date         : rowDate,
                    studentId    : String(data[i][2]),
                    status       : data[i][6],
                    remarks      : data[i][7]
                });
            }
        }

        return records;
    } catch (err) {
        Logger.log("Error in getAttendanceForDate: " + err.message);
        return [];
    }
}

function getStudentAttendanceHistory(studentId) {
    requireAuth_();

    var sheet = SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName("Attendance");

    if (!sheet) return [];

    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return [];

    var records = [];

    // Column alignment:
    // [0] Attendance ID
    // [1] Date
    // [2] Student ID
    // [3] Student Name
    // [4] Batch
    // [5] Time
    // [6] Status
    // [7] Remarks

    for (var i = 1; i < data.length; i++) {
        if (String(data[i][2]).trim() === String(studentId).trim()) {

            // Normalize date to DD/MM/YYYY so it is consistent
            // regardless of how it was originally stored
            var normalizedDate = toDisplayDate_(data[i][1]);

            records.push({
                date    : normalizedDate,
                status  : String(data[i][6] || ""),
                remarks : String(data[i][7] || "")
            });
        }
    }

    // Sort records from most recent to oldest
    // so the student profile shows latest classes first
    records.sort(function(a, b) {
        var dateA = toIsoDate_(a.date);
        var dateB = toIsoDate_(b.date);
        if (dateA > dateB) return -1;
        if (dateA < dateB) return  1;
        return 0;
    });

    return records;
}

function cancelAllClassesForDate(dateIso, reason) {
    requireAuth_();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Attendance");
    if (!sheet) return { success: false };

    var students = getStudents();
    var batches  = getBatches();

    var batchMap = {};
    batches.forEach(function(b) {
        batchMap[b.batchId] = b.batchName;
    });

    // Use noon to avoid timezone/DST edge cases shifting the day
    var dateObj = new Date(dateIso + "T12:00:00");
    var dayName = Utilities.formatDate(
        dateObj,
        Session.getScriptTimeZone(),
        "EEEE"
    );

    var todayBatchIds = batches
        .filter(function(b) {
            return b.days && b.days.includes(dayName);
        })
        .map(function(b) { return b.batchId; });

    var affectedStudents = students.filter(function(s) {
        return todayBatchIds.includes(s.batchId);
    });

    if (affectedStudents.length === 0) {
        return { success: true, count: 0 };
    }

    // Always store as DD/MM/YYYY to match all other attendance records
    var displayDate = toDisplayDate_(dateIso);

    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        var rows = affectedStudents.map(function(s) {
            var attendanceId = nextSequentialId_(
                sheet, "ATT", 0, "ATTENDANCE_ID_SEQ", 4
            );
            return [
                attendanceId,
                displayDate, // ← was: dateIso (raw ISO string)
                s.studentId,
                s.studentName,
                batchMap[s.batchId] || s.batchId || "",
                (s.classFromTime || "") + " - " + (s.classToTime || ""),
                "Cancelled",
                reason,
                "Teacher Emergency Cancel",
                new Date()
            ];
        });

        sheet.getRange(
            sheet.getLastRow() + 1, 1,
            rows.length, rows[0].length
        ).setValues(rows);

        return { success: true, count: rows.length };
    } finally {
        lock.releaseLock();
    }
}

/* ==========================================================
    CUSTOM CLASSES API (Makeup / Extra Sessions)
    Ported from the pre-hardening Code.gs and brought in line
    with the rest of this file: requireAuth_() guard, a script
    lock around the write, and nextSequentialId_() for IDs.

    IMPORTANT: getDisplayValues() (not getValues()) is used when
    reading the sheet. The Date/FromTime/ToTime columns are
    Date-typed cells in Sheets; getValues() would return native
    JS Date objects, which serialize to full UTC timestamps over
    google.script.run (and can shift a calendar day depending on
    the spreadsheet's timezone vs UTC). getDisplayValues() returns
    the same plain strings you see in the sheet (e.g. "2026-08-03"),
    which is what the frontend's exact-match date comparison needs.
========================================================== */
function saveCustomClass(customClass) {
  requireAuth_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CustomClasses");
  if (!sheet) {
    return { success: false, error: "CustomClasses sheet not found" };
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var id = nextSequentialId_(sheet, "CC-", 0, "CUSTOM_CLASS_ID_SEQ", 6);

    sheet.appendRow([
      id,
      customClass.studentId,
      customClass.studentName,
      customClass.batchId || "",
      customClass.batchName || "",
      customClass.date,
      customClass.fromTime,
      customClass.toTime,
      customClass.reason || "",
      customClass.status || "Scheduled",
      customClass.attendanceStatus || "",
      customClass.attendanceRemarks || "",
      true,
      new Date().toISOString()
    ]);

    return { success: true, id: id };
  } catch (e) {
    Logger.log("Error saving custom class: " + e.message);
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function getCustomClasses() {
  requireAuth_();
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CustomClasses");
    if (!sheet) return [];

    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return [];

    var headers = data[0];
    var result = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var obj = {};
      headers.forEach(function (header, index) {
        obj[header] = row[index];
      });
      result.push(obj);
    }

    return result;
  } catch (e) {
    Logger.log(e);
    return [];
  }
}

/* ==========================================================
    INTERNAL (NO-AUTH) VARIANTS
    These are called only from already-authenticated contexts
    like getScheduleData and getScheduleDataLight.
    Never expose these to google.script.run directly.
    The public versions below keep their requireAuth_() guard.
========================================================== */
function getBatches_() {
    var cacheKey       = "LWA_BATCHES";
    var cachedBatches  = getCachedJson_(cacheKey);

    if (cachedBatches !== null) {
        return cachedBatches;
    }

    var sheet = SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName("Batches");

    if (!sheet) return [];

    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) {
        setCachedJson_(cacheKey, []);
        return [];
    }

    data.shift();

    var batches = data.map(function(row) {
        return {
            batchId   : row[0] || "",
            batchName : row[1] || "",
            days      : row[2] || ""
        };
    });

    setCachedJson_(cacheKey, batches);
    return batches;
}

function getStudents_() {
    var cacheKey        = "LWA_STUDENTS";
    var cachedStudents  = getCachedJson_(cacheKey);

    if (cachedStudents !== null) {
        return cachedStudents;
    }

    var sheet = SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName("Students");

    if (!sheet) return [];

    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) {
        setCachedJson_(cacheKey, []);
        return [];
    }

    data.shift();

    var students = data.map(function(row) {
        return {
            studentId     : row[0] || "",
            studentName   : row[1] || "",
            parentName    : row[2] || "",
            whatsapp      : row[3] || "",
            classType     : row[4] || "",
            batchId       : row[5] || "",
            classFromTime : row[6] || "",
            classToTime   : row[7] || "",
            status        : row[8] || "Active",
            notes         : row[9] || "",
            expectedFee   : row[10] ? Number(row[10]) : 0,
            paidUntil     : row[11] || ""
        };
    });

    setCachedJson_(cacheKey, students);
    return students;
}

function getAttendanceForDate_(dateStr) {
    try {
        var sheet = SpreadsheetApp
            .getActiveSpreadsheet()
            .getSheetByName("Attendance");

        if (!sheet) return [];

        var data = sheet.getDataRange().getDisplayValues();
        if (data.length <= 1) return [];

        var targetDisplay = toDisplayDate_(dateStr);
        var records       = [];

        for (var i = 1; i < data.length; i++) {
            var rowDate = toDisplayDate_(data[i][1]);

            if (rowDate === targetDisplay) {
                records.push({
                    attendanceId : data[i][0],
                    date         : rowDate,
                    studentId    : String(data[i][2]),
                    status       : data[i][6],
                    remarks      : data[i][7]
                });
            }
        }

        return records;
    } catch (err) {
        Logger.log("Error in getAttendanceForDate_: " + err.message);
        return [];
    }
}

function getCustomClassesForDate(dateString) {
    requireAuth_();
    return getCustomClassesForDate_(dateString);
}

function getCustomClassesForDate_(dateString) {
    try {
        var sheet = SpreadsheetApp
            .getActiveSpreadsheet()
            .getSheetByName("CustomClasses");

        if (!sheet) return [];

        var data = sheet.getDataRange().getDisplayValues();
        if (data.length <= 1) return [];

        var headers = data[0];
        var result  = [];

        for (var i = 1; i < data.length; i++) {
            var row = data[i];
            var obj = {};
            headers.forEach(function(header, index) {
                obj[header] = row[index];
            });

            if (obj.Date === dateString) {
                result.push(obj);
            }
        }

        return result;
    } catch (err) {
        Logger.log("Error in getCustomClassesForDate_: " + err.message);
        return [];
    }
}
/* ==========================================================
    COMBINED SCHEDULE LOADER
    Single round trip replaces 4 sequential google.script.run
    calls. Uses private no-auth variants since requireAuth_()
    is already called once at the top of this function.
========================================================== */
function getScheduleData(dateIso) {
    requireAuth_(); // ← single auth check for all four reads

    try {
        var batches       = getBatches_();
        var students      = getStudents_();
        var attendance    = getAttendanceForDate_(dateIso);
        var customClasses = getCustomClassesForDate_(dateIso);

        return {
            batches       : batches       || [],
            students      : students      || [],
            attendance    : attendance    || [],
            customClasses : customClasses || []
        };
    } catch (e) {
        Logger.log("Error in getScheduleData: " + e.message);
        return {
            batches       : [],
            students      : [],
            attendance    : [],
            customClasses : []
        };
    }
}

function updateCustomClassAttendance(customClassId, status, remarks) {
  requireAuth_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CustomClasses");
  if (!sheet) return { success: false, error: "Sheet not found" };

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var data = sheet.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(customClassId)) {
        sheet.getRange(i + 1, 11).setValue(status);
        sheet.getRange(i + 1, 12).setValue(remarks);
        return { success: true };
      }
    }
    return { success: false, error: "Custom class not found" };
  } catch (e) {
    Logger.log("Error updating custom class attendance: " + e.message);
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function getScheduleDataLight(dateIso) {
    requireAuth_(); // ← single auth check for both reads

    try {
        var attendance    = getAttendanceForDate_(dateIso);
        var customClasses = getCustomClassesForDate_(dateIso);

        return {
            attendance    : attendance    || [],
            customClasses : customClasses || []
        };
    } catch (e) {
        Logger.log("Error in getScheduleDataLight: " + e.message);
        return {
            attendance    : [],
            customClasses : []
        };
    }
}