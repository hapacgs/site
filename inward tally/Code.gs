function doGet(e) {
  try {
    var sheetId = e.parameter.sheetId;
    var sheetName = e.parameter.sheetName;
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName(sheetName);
    var action = e.parameter.action || "latest";

    if (action === "getRow") {
      var rakeId = e.parameter.rakeId;
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (var i = 0; i < ids.length; i++) {
          if (ids[i][0] == rakeId) {
            var rowVals = sheet.getRange(i + 2, 1, 1, 30).getValues()[0]; // A to AD
            var data = {
              rakeType: rowVals[1],                                 // B
              wagonType: rowVals[2],                                // C
              from: rowVals[3],                                     // D
              to: rowVals[4],                                       // E
              wagons: rowVals[5],                                   // F
              invoiceNo: rowVals[6],                                // G
              invoiceDate: formatCell(rowVals[7], "dd/MM/yyyy"),    // H
              actualWt: rowVals[8],                                 // I
              consigner: rowVals[9],                                // J
              consignee: rowVals[10],                               // K
              commodity: rowVals[11],                               // L
              wagonSummary: (function () {                         // M - WAGON SUMMARY JSON
                try {
                  return rowVals[12] ? JSON.parse(rowVals[12]) : [];
                } catch (e2) {
                  return [];
                }
              })(),
              rrNumber: rowVals[13],                                // N
              rrDate: formatCell(rowVals[14], "dd/MM/yyyy"),        // O
              chargeWeight: rowVals[15],                            // P
              freightExclGst: rowVals[16],                          // Q
              gst: rowVals[17],                                     // R
              totalFreightInclGst: rowVals[18],                     // S
              placementDate: formatCell(rowVals[19], "dd/MM/yyyy"), // T
              placementTime: formatCell(rowVals[20], "HH:mm"),      // U
              releaseDate: formatCell(rowVals[21], "dd/MM/yyyy"),   // V
              releaseTime: formatCell(rowVals[22], "HH:mm"),        // W
              totalDet: formatCell(rowVals[23], "H:mm"),            // X
              freeHrs: rowVals[24],                                 // Y
              nightInc: formatCell(rowVals[25], "H:mm"),            // Z
              netDet: formatCell(rowVals[26], "H:mm"),              // AA
              demmHrs: rowVals[27],                                 // AB
              demmAmt: rowVals[28],                                 // AC
              status: rowVals[29],                                  // AD
              suggestedClass: getClassForCommodity(ss, rowVals[11])
            };
            return ContentService.createTextOutput(JSON.stringify({status:"success", found:true, data:data}))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({status:"success", found:false}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Placement Memo ka "SELECT RAKE ID" dropdown yahi action use karta hain.
    // Status "DELIVERED" match karte hain (pehle yahan "Arrived" tha, jo galat
    // rakes dikha raha tha).
    if (action === "pending") {
      var lastRow3 = sheet.getLastRow();
      var idsList2 = [];
      if (lastRow3 >= 2) {
        var rows = sheet.getRange(2, 1, lastRow3 - 1, 30).getValues();
        for (var j = rows.length - 1; j >= 0; j--) {
          var st = rows[j][29];
          if (st === "DELIVERED") { idsList2.push(rows[j][0]); }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({status:"success", rakeIds: idsList2}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "done") {
      var lastRow4 = sheet.getLastRow();
      var idsList3 = [];
      if (lastRow4 >= 2) {
        var rows2 = sheet.getRange(2, 1, lastRow4 - 1, 30).getValues();
        for (var k = rows2.length - 1; k >= 0; k--) {
          var st2 = rows2[k][29];
          if (st2 === "RELEASED") { idsList3.push(rows2[k][0]); }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({status:"success", rakeIds: idsList3}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "lastDeliveryNo") {
      var type = (e.parameter.type || "").toString().trim().toUpperCase();
      var lastRow5 = sheet.getLastRow();
      var maxNo = 0;
      if (lastRow5 >= 2) {
        var rows3 = sheet.getRange(2, 31, lastRow5 - 1, 2).getValues();
        for (var m = 0; m < rows3.length; m++) {
          var lf = (rows3[m][0] || "").toString().trim().toUpperCase();
          if (lf === type) {
            var n = parseInt(rows3[m][1], 10);
            if (!isNaN(n) && n > maxNo) maxNo = n;
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({status:"success", nextDeliveryNo: maxNo + 1}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "getShuntingRate") {
      var rateProp1 = PropertiesService.getScriptProperties().getProperty("SHUNTING_RATE_PER_HR");
      var rateNum1 = rateProp1 ? parseFloat(rateProp1) : 10620;
      return ContentService.createTextOutput(JSON.stringify({status:"success", rate: rateNum1}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var lastRow2 = sheet.getLastRow();
    var idsList = [];
    if (lastRow2 >= 2) {
      var count = Math.min(5, lastRow2 - 1);
      var startRow = lastRow2 - count + 1;
      idsList = sheet.getRange(startRow, 1, count, 1).getValues().flat().reverse();
    }
    return ContentService.createTextOutput(JSON.stringify({status:"success", rakeIds: idsList}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status:"error", message: err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Date/Time cell agar Google Sheet ne asli Date object bana diya hain, to usse text format me convert karta hain
function formatCell(val, pattern) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Kolkata", pattern);
  }
  return val || "";
}

// CMDT_MASTER sheet me Commodity Code (A column) dhoondh kar uski Class (B column) return karta hain.
// Match na mile (ya CMDT_MASTER sheet hi na ho) to default "130A" return karta hain.
function getClassForCommodity(ss, commodityCode) {
  try {
    var cmdtSheet = ss.getSheetByName("CMDT_MASTER");
    if (!cmdtSheet) return "130A";
    var lastRow = cmdtSheet.getLastRow();
    if (lastRow < 1) return "130A";
    var rows = cmdtSheet.getRange(1, 1, lastRow, 2).getValues(); // A: code, B: class
    var code = (commodityCode || "").toString().trim().toUpperCase();
    for (var i = 0; i < rows.length; i++) {
      var cellCode = (rows[i][0] || "").toString().trim().toUpperCase();
      if (cellCode !== "" && cellCode === code) {
        var cls = rows[i][1];
        return (cls === "" || cls === null || typeof cls === "undefined") ? "130A" : cls;
      }
    }
    return "130A";
  } catch (err) {
    return "130A";
  }
}

// "H:MM" (jaise "2:30") ya seedha decimal number ("2.5") — dono ko decimal hours me convert karta hain
function parseHrsMinToDecimal(val) {
  if (val === "" || val === null || typeof val === "undefined") return 0;
  var s = val.toString().trim();
  if (s === "") return 0;
  if (s.indexOf(":") === -1) {
    var f = parseFloat(s);
    return isNaN(f) ? 0 : f;
  }
  var parts = s.split(":");
  var h = parseInt(parts[0], 10) || 0;
  var mi = parseInt(parts[1], 10) || 0;
  return h + (mi / 60);
}

// "H:MM" ya "HH:MM" string ko total minutes me convert karta hain
function parseTimeToMinutes(val) {
  if (val === "" || val === null || typeof val === "undefined") return 0;
  var s = val.toString().trim();
  if (s === "") return 0;
  var parts = s.split(":");
  if (parts.length !== 2) return 0;
  var h = parseInt(parts[0], 10) || 0;
  var m = parseInt(parts[1], 10) || 0;
  return h * 60 + m;
}

// Total minutes ko "H:MM" string me convert karta hain (negative ko 0 par clamp karta hain)
function minutesToHHMM(totalMinutes) {
  if (totalMinutes < 0) totalMinutes = 0;
  var h = Math.floor(totalMinutes / 60);
  var m = totalMinutes % 60;
  return h + ":" + (m < 10 ? "0" + m : m);
}

// NET DET. (AA) = TOTAL DET. (X) - FREE HRS. (Y) - NIGHT INCENTIVE (Z)
function computeNetDet(totalDetVal, freeHrsVal, nightIncVal) {
  var totalMin = parseTimeToMinutes(totalDetVal);
  var freeMin = (parseFloat(freeHrsVal) || 0) * 60;
  var nightMin = parseTimeToMinutes(nightIncVal);
  return minutesToHHMM(totalMin - freeMin - nightMin);
}

// DEMM. HRS. (AB) = NET DET. (AA) ko agle poore ghante me round off karo
// (koi bhi minute ho to +1 hr, jaise 05:50 => 6, 06:01 => 7, 06:00 => 6)
function computeDemmHrs(netDetStr) {
  var parts = (netDetStr || "0:00").split(":");
  var h = parseInt(parts[0], 10) || 0;
  var m = parseInt(parts[1], 10) || 0;
  return m > 0 ? h + 1 : h;
}

// DEMM. AMT. (AC) = No. of Wagons (F) x slab-wise DEMM. HRS. (AB) rate per wagon
var DEMM_SLABS = [
  { hrs: 6,          rate: 150 },    // UPTO 6 HRS.
  { hrs: 6,          rate: 165 },    // 7 TO 12 HRS.
  { hrs: 12,         rate: 187.5 },  // 13 TO 24 HRS.
  { hrs: 24,         rate: 225 },    // 25 TO 48 HRS.
  { hrs: 24,         rate: 300 },    // 49 TO 72 HRS.
  { hrs: Infinity,   rate: 450 }     // ABOVE 72 HRS.
];

function computeDemmAmount(wagonCount, demmHrs) {
  var remaining = demmHrs;
  var perWagon = 0;
  for (var i = 0; i < DEMM_SLABS.length && remaining > 0; i++) {
    var chargeable = Math.min(remaining, DEMM_SLABS[i].hrs);
    perWagon += chargeable * DEMM_SLABS[i].rate;
    remaining -= chargeable;
  }
  var wagons = parseInt(wagonCount, 10) || 0;
  return Math.round(perWagon * wagons * 100) / 100;
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var sheetId = body.sheetId;
    var sheetName = body.sheetName;
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName(sheetName);
    var action = body.action || "insert";

    if (action === "updateEntry") {
      var rakeId = body.rakeId;
      var data = body.data; // 11 values: N se X (last value = TOTAL DET.)
      var lastRow = sheet.getLastRow();
      var found = false;

      if (lastRow >= 2) {
        var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (var i = 0; i < ids.length; i++) {
          if (ids[i][0] == rakeId) {
            var rowNum = i + 2;

            // BUG FIX: X column (TOTAL DET., ek "H:MM" jaisi text string, e.g. "45:10") ko
            // Google Sheets khud-ba-khud Time/Duration value me convert kar deta tha, kyunki
            // usme colon (:) hain. Uske baad neeche computeNetDet() jab X ko wapas padhta tha
            // to woh ek Date object ban chuka hota tha, jisse parseTimeToMinutes() galat 0
            // nikaalta tha aur NET DET./DEMM. HRS./DEMM. AMT. sahi se calculate/save nahi hote
            // the. Fix: X (aur Z - NIGHT INC., jo bhi "H:MM" text hain) ko likhne se PEHLE
            // column ka number format "@" (Plain Text) force kar do, taaki Sheets isse text
            // hi rehne de aur auto-convert na kare.
            sheet.getRange(rowNum, 24).setNumberFormat("@"); // X - TOTAL DET.
            sheet.getRange(rowNum, 14, 1, data.length).setValues([data]); // N onwards

            // Y - FREE HRS.: sirf tab likhte hain jab yeh column pehle se KHALI ho
            // (normal flow me yeh Inward Tally.html submit ke time hi bhar jaata hain,
            // WTR.html isse kabhi overwrite nahi karta — sirf purani/missing entries
            // ke liye backfill karta hain jab frontend ne freeHrs bheja ho)
            if (typeof body.freeHrs !== "undefined" && body.freeHrs !== "") {
              var currentFreeHrs = sheet.getRange(rowNum, 25).getValue(); // Y column
              if (currentFreeHrs === "" || currentFreeHrs === null) {
                sheet.getRange(rowNum, 25).setValue(body.freeHrs); // Y column
              }
            }

            // Z - NIGHT INCENTIVE (frontend se calculate hokar aata hain) — isko bhi
            // Plain Text format me hi likho, upar wali wajah se.
            if (typeof body.nightIncentive !== "undefined") {
              var zCell = sheet.getRange(rowNum, 26); // Z column
              zCell.setNumberFormat("@");
              zCell.setValue(body.nightIncentive);
            }

            // AA - NET DET., AB - DEMM. HRS., AC - DEMM. AMT.
            // Sheet me abhi-abhi save hui X (TOTAL DET.), Y (FREE HRS.), Z (NIGHT INC.), F (NO. OF WAGONS)
            // se calculate karo. getDisplayValue() use kar rahe hain X/Z ke liye taaki formatted
            // text hi mile, kisi bhi tarah ka Date/Duration object nahi.
            var totalDetVal = sheet.getRange(rowNum, 24).getDisplayValue(); // X
            var freeHrsVal = sheet.getRange(rowNum, 25).getValue();         // Y
            var nightIncVal = sheet.getRange(rowNum, 26).getDisplayValue(); // Z
            var wagonCountVal = sheet.getRange(rowNum, 6).getValue();       // F

            var netDetStr = computeNetDet(totalDetVal, freeHrsVal, nightIncVal);
            var demmHrsVal = computeDemmHrs(netDetStr);
            var demmAmtVal = computeDemmAmount(wagonCountVal, demmHrsVal);

            var aaCell = sheet.getRange(rowNum, 27); // AA - NET DET.
            aaCell.setNumberFormat("@");
            aaCell.setValue(netDetStr);
            sheet.getRange(rowNum, 28).setValue(demmHrsVal);  // AB - DEMM. HRS.
            sheet.getRange(rowNum, 29).setValue(demmAmtVal);  // AC - DEMM. AMT.

            // AD - STATUS: WTR submit ka matlab hain saara data bhar gaya => RELEASED
            sheet.getRange(rowNum, 30).setValue("RELEASED"); // AD column

            found = true;
            break;
          }
        }
      }

      if (!found) {
        return ContentService.createTextOutput(JSON.stringify({status:"error", message:"Rake ID nahi mila"}))
          .setMimeType(ContentService.MimeType.JSON);
      }

      return ContentService.createTextOutput(JSON.stringify({status:"success"}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Delivery Book submit: AE-AJ columns bharo aur STATUS ko DELIVERED kar do
    if (action === "updateDelivery") {
      var rakeIdD = body.rakeId;
      var lastRowD = sheet.getLastRow();
      var foundD = false;

      if (lastRowD >= 2) {
        var idsD = sheet.getRange(2, 1, lastRowD - 1, 1).getValues();
        for (var p = 0; p < idsD.length; p++) {
          if (idsD[p][0] == rakeIdD) {
            var rowNumD = p + 2;

            var localForeign = body.localForeign;
            var deliveryNo = body.deliveryNo;
            var deliveryClass = body.deliveryClass;
            var wharfageAmount = (body.wharfageAmount === "" || body.wharfageAmount === null || typeof body.wharfageAmount === "undefined")
              ? 0 : (parseFloat(body.wharfageAmount) || 0);
            var shuntingHrsRaw = (body.shuntingHrs === "" || body.shuntingHrs === null || typeof body.shuntingHrs === "undefined")
              ? "0:00" : body.shuntingHrs;

            // SHUNTING AMOUNT hamesha server-side hi calculate karo (current stored rate se),
            // taaki rate ki galat client-side value par bharosa na karna pade
            var rateProp2 = PropertiesService.getScriptProperties().getProperty("SHUNTING_RATE_PER_HR");
            var rate2 = rateProp2 ? parseFloat(rateProp2) : 10620;
            var shuntingHrsDecimal = parseHrsMinToDecimal(shuntingHrsRaw);
            var shuntingAmountVal = Math.round(shuntingHrsDecimal * rate2 * 100) / 100;

            var deliveryRange = sheet.getRange(rowNumD, 31, 1, 6); // AE to AJ
            deliveryRange.setNumberFormat("@"); // shuntingHrs jaisa "H:MM" text auto-convert na ho, isliye plain text
            deliveryRange.setValues([[
              localForeign, deliveryNo, deliveryClass, wharfageAmount, shuntingHrsRaw, shuntingAmountVal
            ]]);

            sheet.getRange(rowNumD, 30).setValue("DELIVERED"); // AD - STATUS

            foundD = true;
            break;
          }
        }
      }

      if (!foundD) {
        return ContentService.createTextOutput(JSON.stringify({status:"error", message:"Rake ID nahi mila"}))
          .setMimeType(ContentService.MimeType.JSON);
      }

      return ContentService.createTextOutput(JSON.stringify({status:"success"}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Delivery Book: SHUNTING rate per hr. change karo — password (94976) required
    if (action === "setShuntingRate") {
      var pwd = (body.password || "").toString();
      if (pwd !== "94976") {
        return ContentService.createTextOutput(JSON.stringify({status:"error", message:"Galat password."}))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var newRate = parseFloat(body.rate);
      if (isNaN(newRate) || newRate < 0) {
        return ContentService.createTextOutput(JSON.stringify({status:"error", message:"Invalid rate."}))
          .setMimeType(ContentService.MimeType.JSON);
      }
      PropertiesService.getScriptProperties().setProperty("SHUNTING_RATE_PER_HR", newRate.toString());
      return ContentService.createTextOutput(JSON.stringify({status:"success", rate:newRate}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // default action: insert new row with duplicate check
    var data2 = body.data;
    var rakeId2 = data2[0];
    var lastRow2 = sheet.getLastRow();
    if (lastRow2 >= 2) {
      var existingIds = sheet.getRange(2, 1, lastRow2 - 1, 1).getValues().flat();
      if (existingIds.indexOf(rakeId2) !== -1) {
        return ContentService.createTextOutput(JSON.stringify({status:"duplicate"}))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    sheet.appendRow(data2);

    // Naye row ke liye STATUS set karo: Inward Tally submit hote hi => ARRIVED
    var newRow = sheet.getLastRow();
    sheet.getRange(newRow, 30).setValue("ARRIVED"); // AD column

    return ContentService.createTextOutput(JSON.stringify({status:"success"}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status:"error", message: err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}