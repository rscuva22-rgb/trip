/**
 * RSC (Uva) Sports & Welfare Society - Arugam Bay Tour 2026
 * Backend webhook handler with Real-time Read & Write + Admin Management
 */

// 1. GET: Sends all booked seats & payment statuses to all devices
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Seat Manifest") || ss.getSheets()[0];
    var values = sheet.getDataRange().getValues();
    var bookedSeats = {};

    // Read rows starting from row 2 (skipping header)
    for (var i = 1; i < values.length; i++) {
      var seatNum = parseInt(values[i][0], 10);
      var name = values[i][1];
      var category = values[i][2];
      var startPlace = values[i][3];
      var connectNo = values[i][4];
      var status = values[i][5];
      var amount = values[i][6];
      var paymentStatus = values[i][7] || (status === "Paid" ? "Paid" : "Pending");

      if (seatNum && (status === "Booked" || status === "Reserved" || status === "Paid" || (name && name !== ""))) {
        bookedSeats[seatNum] = {
          name: name || "Booked",
          type: category || "Member",
          startingPlace: startPlace || "",
          contactNo: connectNo || "",
          amount: amount || 0,
          status: status || "Booked",
          paymentStatus: paymentStatus
        };
      }
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      bookedSeats: bookedSeats
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 2. POST: Handles bookings, payment updates, and seat releases
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Seat Manifest") || ss.getSheets()[0];
    var data = JSON.parse(e.postData.contents);
    var targetSeat = parseInt(data.seatNumber, 10);
    var values = sheet.getDataRange().getValues();
    var seatFound = false;

    // Handle Admin: Release Seat
    if (data.action === "releaseSeat") {
      for (var i = 1; i < values.length; i++) {
        if (parseInt(values[i][0], 10) === targetSeat) {
          var row = i + 1;
          sheet.getRange(row, 2, 1, 7).clearContent();
          seatFound = true;
          break;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        action: "released",
        seat: targetSeat
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Handle Admin: Update Payment Status
    if (data.action === "updatePayment") {
      for (var i = 1; i < values.length; i++) {
        if (parseInt(values[i][0], 10) === targetSeat) {
          var row = i + 1;
          sheet.getRange(row, 8).setValue(data.paymentStatus);
          seatFound = true;
          break;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        action: "paymentUpdated",
        seat: targetSeat,
        paymentStatus: data.paymentStatus
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Default: New / Existing Seat Booking
    for (var i = 1; i < values.length; i++) {
      if (parseInt(values[i][0], 10) === targetSeat) {
        var row = i + 1;
        sheet.getRange(row, 2).setValue(data.passengerName);
        sheet.getRange(row, 3).setValue(data.category);
        sheet.getRange(row, 4).setValue(data.startingPlace);
        sheet.getRange(row, 5).setValue("'" + data.connectNo);
        sheet.getRange(row, 6).setValue("Booked");
        sheet.getRange(row, 7).setValue(Number(data.amount));
        sheet.getRange(row, 8).setValue(data.paymentStatus || "Pending");
        seatFound = true;
        break;
      }
    }

    if (!seatFound) {
      sheet.appendRow([
        targetSeat,
        data.passengerName,
        data.category,
        data.startingPlace,
        "'" + data.connectNo,
        "Booked",
        Number(data.amount),
        data.paymentStatus || "Pending",
        new Date()
      ]);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      seat: targetSeat
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
