var MASTER_SPREADSHEET_ID = (typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID)
  ? SPREADSHEET_ID
  : '1Y5yMDhW9Lou2VY0zgsPqo7DDih66Qa4sfupI3cNV-0Q';

var DATABASE_COL_INDEX = {
  ID: 0,
  NAME: 3,
  PHONE: 4,
  PERSON_COUNT: 6,
  CITY: 7,
  HOTEL: 8,
  HOTEL_CONFIRMATION: 9,
  CHECKIN_DATE: 10,
  CHECKOUT_DATE: 11,
  ROOM_COUNT: 12,
  ROOM_TYPE: 13,
  VIEW_TYPE: 14,
  MEAL_TYPE: 15,
  HOTEL_PRICE_EURO: 17,
  SELLIN_PRICE: 18,
  SELLIN_EURO_PRICE: 19,
  CURRENCY: 20,
  ARRIVED_AMOUNT: 24,
  ARRIVED_EURO_AMOUNT: 25,
  DEPOSIT_METHOD: 26,
  ARRIVED_AMOUNT_CURRENCY: 29,
  REMAINING_AMOUNT: 30,
  REMAINING_EURO_AMOUNT: 31,
  REMAINING_AMOUNT_CURRENCY: 32,
  PAYMENT_STATUS: 33,
  REMAINING_METHOD: 34,
  SERVICE_NAME: 36,
  SERVICE_PRICE: 37,
  SERVICE_EURO_PRICE: 38,
  SERVICE_SELLING_PRICE: 39,
  SERVICE_SELLING_EURO_PRICE: 40,
  NOTES: 41,
  SELLER: 2,
  SUPPLIER: 16,
  RESERVATION_STATUS: 48
};

function getCities() {
  var sheet = SpreadsheetApp.openById("1Y5yMDhW9Lou2VY0zgsPqo7DDih66Qa4sfupI3cNV-0Q").getSheetByName("CITIES");
  var cities = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  console.log(cities);
  return cities; // Returns an array of city names
}



function getColumnByName(columnName) {
  var sheet = SpreadsheetApp.openById("1Y5yMDhW9Lou2VY0zgsPqo7DDih66Qa4sfupI3cNV-0Q").getSheetByName("INFORMATION");
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  console.log(header);
  var columnIndex = header.indexOf(columnName) + 1; // +1 because getRange is 1-indexed
  var column = []; 

  if (columnIndex > 0) {
      var columnData = sheet.getRange(2, columnIndex, sheet.getLastRow() - 1, 1).getValues(); // Get hotel names below the city

      for (var i = 0; i < columnData.length; i++) {
          var columnN = columnData[i][0]; // Access the hotel name
          if (columnN) { // Check if the hotel name is not empty
              column.push(columnN); // Add non-empty hotel names to the array
          }
      }
  }
  console.log(column);
  return column; 
}

function getHotelsByCity(city) {
  var sheet = SpreadsheetApp.openById("1Y5yMDhW9Lou2VY0zgsPqo7DDih66Qa4sfupI3cNV-0Q").getSheetByName("CITIES");
  var cities = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var columnIndex = cities.indexOf(city) + 1; // +1 because getRange is 1-indexed
  var hotels = []; // Initialize an empty array to store hotel names

  if (columnIndex > 0) {
      var hotelData = sheet.getRange(2, columnIndex, sheet.getLastRow() - 1, 1).getValues(); // Get hotel names below the city

      for (var i = 0; i < hotelData.length; i++) {
          var hotelName = hotelData[i][0]; // Access the hotel name
          if (hotelName) { // Check if the hotel name is not empty
              hotels.push(hotelName); // Add non-empty hotel names to the array
          }
      }
  }

  return hotels;
}

  function getCustomerDataById(customerId) {
    const sheet = SpreadsheetApp.openById("1Y5yMDhW9Lou2VY0zgsPqo7DDih66Qa4sfupI3cNV-0Q").getSheetByName("DATABASE");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == customerId) {

        if (data[i][10] instanceof Date) {
          data[i][10] = Utilities.formatDate(data[i][10], Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        if (data[i][11] instanceof Date) {
          data[i][11] = Utilities.formatDate(data[i][11], Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        if (data[i][44] instanceof Date) {
          data[i][44] = Utilities.formatDate(data[i][44], Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        if (data[i][46] instanceof Date) {
          data[i][46] = Utilities.formatDate(data[i][46], Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        console.log(data[i]);
        return data[i];
      }
    }
    console.log("Nothing here");
    return [];  // Return an empty array if the customer is not found
  }

  function getCustomerMapData(customerId) {
    const sheet = SpreadsheetApp.openById("1Y5yMDhW9Lou2VY0zgsPqo7DDih66Qa4sfupI3cNV-0Q").getSheetByName("DATABASE");
    const data = sheet.getDataRange().getValues();
    let customerData = null;
    for (let i = 1; i < data.length; i++) { // Assuming the first row is headers
      if (data[i][0] == customerId) {  // Check if the first column matches customerId
        customerData = {
            seller: data[i][2],
            name: data[i][3],
            phone: data[i][4],
            person: data[i][6],
            city: data[i][7],
            hotel: data[i][8],
            hotelConfirmation: data[i][9],
            checkinDate: Utilities.formatDate(data[i][10], Session.getScriptTimeZone(), "yyyy-MM-dd"),
            checkoutDate: Utilities.formatDate(data[i][11], Session.getScriptTimeZone(), "yyyy-MM-dd"),
            roomCount: data[i][12],
            roomType: data[i][13],
            viewType: data[i][14],
            meals: data[i][15],
            supplier: data[i][16],
            hotelPriceEuro: data[i][17],
            sellinPrice: data[i][18],
            sellinEuroPrice: data[i][19],
            currency: data[i][20],
            arrivedAmount: data[i][24],
            arrivedEuroAmount: data[i][25],
            depositMethod: data[i][26],
            sendingCost: data[i][27],
            sendingEuroCost: data[i][28],
            arrivedAmountCurrency: data[i][29],
            remainingAmount: data[i][30],
            remainingEuroAmount: data[i][31],
            remainingAmountCurrency: data[i][32],
            paymentStatus: data[i][33],
            remainingMethod: data[i][34],
            service: data[i][36],
            servicePrice: data[i][37],
            serviceEuroPrice: data[i][38],
            serviceSellingPrice: data[i][39],
            serviceSellingEuroPrice: data[i][40],
            flowerGift: data[i][41],
            reservationStatus: data[i][48],
        };
        break;
      }
    }
    console.log(customerData);
    return customerData;
  }

function getCustomers(searchTerm, pageNumber, checkInDate, checkOutDate) {
  var sheet = SpreadsheetApp.openById("1Y5yMDhW9Lou2VY0zgsPqo7DDih66Qa4sfupI3cNV-0Q").getSheetByName("DATABASE");
  var data = sheet.getDataRange().getValues();
  var timezone = Session.getScriptTimeZone();
  var sanitizedSearch = (searchTerm || '').toString().trim().toLowerCase();
  var pageSize = PAGE_SIZE || 20;
  var page = parseInt(pageNumber, 10);
  if (isNaN(page) || page < 1) {
    page = 1;
  }

  var startDate = checkInDate ? new Date(checkInDate) : null;
  var endDate = checkOutDate ? new Date(checkOutDate) : null;
  if (startDate && isNaN(startDate.getTime())) {
    startDate = null;
  }
  if (endDate && isNaN(endDate.getTime())) {
    endDate = null;
  }

  var rows = [];
  for (var i = 1; i < data.length; i++) {
    rows.push(data[i]);
  }

  var filteredRows = rows.filter(function (row) {
    var checkinValue = parseSheetDate(row[10]);
    var checkoutValue = parseSheetDate(row[11]);
    var text = [
      row[0],
      row[3],
      row[4],
      row[7],
      row[8],
      row[9]
    ].join(' ').toString().toLowerCase();

    var matchesSearch = sanitizedSearch ? text.indexOf(sanitizedSearch) !== -1 : true;
    if (!matchesSearch) {
      return false;
    }

    if (startDate && (!checkinValue || checkinValue < startDate)) {
      return false;
    }
    if (endDate && (!checkoutValue || checkoutValue > endDate)) {
      return false;
    }

    return true;
  });

  var totalCount = filteredRows.length;
  var totalPages = Math.max(1, Math.ceil(totalCount / pageSize));
  if (page > totalPages) {
    page = totalPages;
  }
  var startIndex = (page - 1) * pageSize;
  var pageRows = filteredRows.slice(startIndex, startIndex + pageSize);

  var customers = pageRows.map(function (row) {
    var checkinValue = parseSheetDate(row[10]);
    var checkoutValue = parseSheetDate(row[11]);
    return {
      id: row[0],
      seller: row[2],
      name: row[3],
      phone: row[4],
      person: row[6],
      city: row[7],
      hotel: row[8],
      hotelConfirmation: row[9],
      checkinDate: checkinValue ? Utilities.formatDate(checkinValue, timezone, "yyyy-MM-dd") : '',
      checkoutDate: checkoutValue ? Utilities.formatDate(checkoutValue, timezone, "yyyy-MM-dd") : '',
      roomCount: row[12],
      roomType: row[13],
      viewType: row[14],
      meals: row[15],
      supplier: row[16],
      hotelPriceEuro: row[17],
      sellinPrice: row[18],
      sellinEuroPrice: row[19],
      currency: row[20],
      arrivedAmount: row[24],
      arrivedEuroAmount: row[25],
      depositMethod: row[26],
      sendingCost: row[27],
      sendingEuroCost: row[28],
      arrivedAmountCurrency: row[29],
      remainingAmount: row[30],
      remainingEuroAmount: row[31],
      remainingAmountCurrency: row[32],
      paymentStatus: row[33],
      remainingMethod: row[34],
      service: row[36],
      servicePrice: row[37],
      serviceEuroPrice: row[38],
      serviceSellingPrice: row[39],
      serviceSellingEuroPrice: row[40],
      flowerGift: row[41],
      reservationStatus: row[48]
    };
  });

  return {
    data: customers,
    pagination: {
      page: page,
      totalPages: totalPages,
      pageSize: pageSize,
      totalCount: totalCount
    }
  };
}

/**
 * يعيد الحجوزات المرتبطة بمورد محدد بناءً على طريقتي تحصيل العربون والمبلغ المتبقي.
 * @param {{supplier:string,startDate:string,endDate:string}} filterOptions
 * @returns {{
 *   supplier:string,
 *   filters:{startDate:string,endDate:string},
 *   receivables:Array<Object>,
 *   obligations:Array<Object>,
 *   totals:{
 *     receivablesCount:number,
 *     obligationsCount:number,
 *     receivablesTotal:number,
 *     obligationsTotal:number,
 *     receivablesByCurrency:Object,
 *     obligationsByCurrency:Object
 *   }
 * }}
 */
function getSupplierFinancials(filterOptions) {
  filterOptions = filterOptions || {};
  var supplierName = (filterOptions.supplier || "").toString().trim();
  var startDateInput = filterOptions.startDate;
  var endDateInput = filterOptions.endDate;
  var startDate = parseFilterDate(startDateInput, true);
  var endDate = parseFilterDate(endDateInput, false);
  var timezone = Session.getScriptTimeZone();
  var normalizedSupplier = supplierName.toLowerCase();
  var response = buildSupplierFinancialResponse_(supplierName, startDate, endDate);

  if (!supplierName) {
    return response;
  }

  var sheet = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID).getSheetByName("DATABASE");
  if (!sheet) {
    return response;
  }

  var values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) {
    return response;
  }

  var header = values.shift();
  var headerMap = buildHeaderIndexMap_(header);
  var indexes = {
    id: getHeaderIndexByName_(headerMap, ["booking id", "id"], DATABASE_COL_INDEX.ID),
    customerName: getHeaderIndexByName_(headerMap, ["name", "customer name", "client name"], DATABASE_COL_INDEX.NAME),
    seller: getHeaderIndexByName_(headerMap, ["seller", "agent", "agency"], -1),
    supplier: getHeaderIndexByName_(headerMap, ["hotel supplier", "supplier"], -1),
    checkin: getHeaderIndexByName_(headerMap, ["check-in date", "check in", "checkin", "chek in date"], DATABASE_COL_INDEX.CHECKIN_DATE),
    checkout: getHeaderIndexByName_(headerMap, ["check-out date", "checkout", "check out"], DATABASE_COL_INDEX.CHECKOUT_DATE),
    hotel: getHeaderIndexByName_(headerMap, ["hotel", "hotel name"], DATABASE_COL_INDEX.HOTEL),
    ra3bon: getHeaderIndexByName_(headerMap, ["ra3bon", "advance method", "deposit method", "طريقة تحصيل العربون"], -1),
    collection: getHeaderIndexByName_(headerMap, ["collection method", "remaining method", "طريقة التحصيل"], -1),
    depositAmount: getHeaderIndexByName_(headerMap, ["amount sent", "deposit amount", "prepayment", "arrived amount"], DATABASE_COL_INDEX.ARRIVED_AMOUNT),
    depositCurrency: getHeaderIndexByName_(headerMap, ["amount sent currency", "deposit currency", "prepayment currency", "arrived amount currency"], DATABASE_COL_INDEX.ARRIVED_AMOUNT_CURRENCY),
    remainingAmount: getHeaderIndexByName_(headerMap, ["remaining amount", "rest amount"], DATABASE_COL_INDEX.REMAINING_AMOUNT),
    remainingCurrency: getHeaderIndexByName_(headerMap, ["remaining amount currency", "amount remaining currency"], DATABASE_COL_INDEX.REMAINING_AMOUNT_CURRENCY),
    reservationState: getHeaderIndexByName_(headerMap, ["reservation state", "status"], -1),
    note: getHeaderIndexByName_(headerMap, ["notes", "note"], DATABASE_COL_INDEX.NOTES)
  };

  if (indexes.ra3bon === -1 && indexes.collection === -1) {
    return response;
  }

  var receivables = [];
  var obligations = [];
  var receivableTotals = {};
  var obligationTotals = {};
  var receivableSum = 0;
  var obligationSum = 0;

  values.forEach(function (row) {
    var checkInDate = indexes.checkin > -1 ? parseSheetDate(row[indexes.checkin]) : null;
    if (startDate && (!checkInDate || checkInDate < startDate)) {
      return;
    }
    if (endDate && (!checkInDate || checkInDate > endDate)) {
      return;
    }

    var depositMethod = indexes.ra3bon > -1 ? (row[indexes.ra3bon] || "").toString().trim() : "";
    var collectionMethod = indexes.collection > -1 ? (row[indexes.collection] || "").toString().trim() : "";

    var depositMatches = depositMethod && depositMethod.toLowerCase().indexOf(normalizedSupplier) !== -1;
    var collectionMatches = collectionMethod && collectionMethod.toLowerCase().indexOf(normalizedSupplier) !== -1;

    if (!depositMatches && !collectionMatches) {
      return;
    }

    var checkoutDate = indexes.checkout > -1 ? parseSheetDate(row[indexes.checkout]) : null;
    var recordBase = {
      id: indexes.id > -1 ? (row[indexes.id] || "").toString().trim() : "",
      customerName: indexes.customerName > -1 ? (row[indexes.customerName] || "").toString().trim() : "",
      seller: indexes.seller > -1 ? (row[indexes.seller] || "").toString().trim() : "",
      supplier: indexes.supplier > -1 ? (row[indexes.supplier] || "").toString().trim() : "",
      hotel: indexes.hotel > -1 ? (row[indexes.hotel] || "").toString().trim() : "",
      reservationState: indexes.reservationState > -1 ? (row[indexes.reservationState] || "").toString().trim() : "",
      note: indexes.note > -1 ? (row[indexes.note] || "").toString().trim() : "",
      checkIn: formatDateValue_(checkInDate, timezone),
      checkOut: formatDateValue_(checkoutDate, timezone),
      nights: calculateNights_(checkInDate, checkoutDate)
    };

    if (depositMatches) {
      var depositAmount = indexes.depositAmount > -1 ? sanitizeNumber_(row[indexes.depositAmount]) : 0;
      var depositCurrency = indexes.depositCurrency > -1 ? (row[indexes.depositCurrency] || "").toString().trim() : "";
      receivables.push(Object.assign({}, recordBase, {
        method: depositMethod || "غير محدد",
        amount: depositAmount,
        currency: depositCurrency
      }));
      receivableSum += depositAmount;
      accumulateCurrencyTotal_(receivableTotals, depositCurrency, depositAmount);
    }

    if (collectionMatches) {
      var remainingAmount = indexes.remainingAmount > -1 ? sanitizeNumber_(row[indexes.remainingAmount]) : 0;
      var remainingCurrency = indexes.remainingCurrency > -1 ? (row[indexes.remainingCurrency] || "").toString().trim() : "";
      obligations.push(Object.assign({}, recordBase, {
        method: collectionMethod || "غير محدد",
        amount: remainingAmount,
        currency: remainingCurrency
      }));
      obligationSum += remainingAmount;
      accumulateCurrencyTotal_(obligationTotals, remainingCurrency, remainingAmount);
    }
  });

  response.receivables = receivables;
  response.obligations = obligations;
  response.totals.receivablesCount = receivables.length;
  response.totals.obligationsCount = obligations.length;
  response.totals.receivablesTotal = receivableSum;
  response.totals.obligationsTotal = obligationSum;
  response.totals.receivablesByCurrency = receivableTotals;
  response.totals.obligationsByCurrency = obligationTotals;
  response.filters.startDate = formatDateValue_(startDate, timezone);
  response.filters.endDate = formatDateValue_(endDate, timezone);
  return response;
}

function clearInvoice(){
  var myGooglSheet = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  var invoiceSheet = myGooglSheet.getSheetByName("INVOICE");
  resetInvoiceTemplate_(invoiceSheet);
}

function clearVoucher(){
  var myGooglSheet = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  var voucherSheet = myGooglSheet.getSheetByName("VOUCHER");
  resetVoucherTemplate_(voucherSheet);
}

function resetInvoiceTemplate_(invoiceSheet) {
  if (!invoiceSheet) {
    throw new Error('تعذر العثور على ورقة INVOICE داخل ملف Google Sheets.');
  }
  invoiceSheet.getRange("D10").setValue("");
  invoiceSheet.getRange("D11").setValue("");
  invoiceSheet.getRange("C15:K24").clearContent();
  invoiceSheet.getRange("K26").clearContent();
  invoiceSheet.getRange("K28").clearContent();
  invoiceSheet.getRange("K30").clearContent();
  invoiceSheet.getRange("G34").clearContent();
  invoiceSheet.getRange("A34:Z37").clearContent();
  invoiceSheet.getRange("G46").clearContent();
}

function resetVoucherTemplate_(voucherSheet) {
  if (!voucherSheet) {
    throw new Error('تعذر العثور على ورقة VOUCHER داخل ملف Google Sheets.');
  }
  voucherSheet.getRange("G3").setValue("");
  voucherSheet.getRange("E11").setValue("");
  voucherSheet.getRange("E13").setValue("");
  voucherSheet.getRange("E15").setValue("");
  voucherSheet.getRange("D18:D28").clearContent();
  voucherSheet.getRange("G18:G28").clearContent();
  voucherSheet.getRange("C32").setValue("");
}

function getUrl() {
 var url = ScriptApp.getService().getUrl();
 return url;
}

function generateCustomUniqueId(prefix) {
  const date = new Date();
  date.setHours(date.getHours() + 3); // Add 3 hours to the current UTC time
  const year = String(date.getUTCFullYear()).slice(-2); // Last 2 digits of the year in UTC
  const month = String(date.getUTCMonth() + 1); // Month is 0-based
  const day = String(date.getUTCDate()); // Day of the month
  const hours = String(date.getUTCHours()); // Current hours in UTC
  const minutes = String(date.getUTCMinutes()); //
  // Manually pad month, day, and minutes with leading zeros if needed
  const paddedMonth = month.length < 2 ? '0' + month : month;
  const paddedDay = day.length < 2 ? '0' + day : day;
  const paddedHours = hours.length < 2 ? '0' + hours : hours;
  const paddedMinutes = minutes.length < 2 ? '0' + minutes : minutes;
  console.log(prefix + year + paddedMonth + paddedDay + paddedHours + paddedMinutes);

  // Construct the unique ID using concatenation
  return prefix + year + paddedMonth + paddedDay + paddedHours + paddedMinutes; // Regular concatenation
}

function include(fileName){
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

function submitManualVoucher(formData) {
  var payload = normalizeManualVoucherPayload_(formData || {});
  var spreadsheet = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  var voucherSheet = spreadsheet.getSheetByName("VOUCHER");
  if (!voucherSheet) {
    throw new Error("تعذر العثور على ورقة VOUCHER داخل ملف Google Sheets.");
  }
  logVoucherSubmission_(spreadsheet, payload, "manual", "");
  var pdfBlob = generateVoucherPdf_(voucherSheet, payload);
  sendVoucherEmail_(payload, pdfBlob);
  return "تم إرسال الإيصال";
}

function submitManualInvoice(formData) {
  var payload = normalizeManualInvoicePayload_(formData || {});
  var spreadsheet = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  var invoiceSheet = spreadsheet.getSheetByName("INVOICE");
  if (!invoiceSheet) {
    throw new Error("تعذر العثور على ورقة INVOICE داخل ملف Google Sheets.");
  }
  logInvoiceSubmission_(spreadsheet, payload, "manual", "");
  var pdfBlob = generateInvoicePdf_(invoiceSheet, payload);
  sendInvoiceEmail_(payload, pdfBlob);
  return "تم إرسال الفاتورة";
}

function submitVoucher(requestPayload) {
  var payload = requestPayload || {};
  var reservationId = payload.rowId || payload.reservationId;
  var email = payload.email;
  reservationId = reservationId !== null && typeof reservationId !== "undefined"
    ? String(reservationId).trim()
    : "";
  if (!reservationId) {
    throw new Error("الرجاء اختيار حجز صالح لإصدار الإيصال.");
  }
  if (!email) {
    throw new Error("الرجاء اختيار بريد إلكتروني لإرسال الإيصال.");
  }
  var spreadsheet = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  var voucherSheet = spreadsheet.getSheetByName("VOUCHER");
  if (!voucherSheet) {
    throw new Error("تعذر العثور على ورقة VOUCHER داخل ملف Google Sheets.");
  }
  var reservations = getReservationsByIds_([reservationId], spreadsheet);
  var row = reservations[reservationId];
  if (!row) {
    throw new Error("لم يتم العثور على بيانات الحجز رقم " + reservationId + ".");
  }
  var voucherPayload = buildVoucherPayloadFromReservationRow_(row, {
    email: email,
    refundType: payload.refundType,
    note: payload.note
  });
  logVoucherSubmission_(spreadsheet, voucherPayload, "automatic", reservationId);
  var pdfBlob = generateVoucherPdf_(voucherSheet, voucherPayload);
  sendVoucherEmail_(voucherPayload, pdfBlob);
  return "تم إرسال الإيصال";
}

function submitInvoice(reservationIds, targetEmail) {
  var normalizedIds = normalizeReservationIds_(reservationIds);
  if (!normalizedIds.length) {
    throw new Error("الرجاء اختيار حجز أو أكثر لإصدار الفاتورة.");
  }
  if (!targetEmail) {
    throw new Error("الرجاء اختيار بريد إلكتروني لإرسال الفاتورة.");
  }
  if (normalizedIds.length > 10) {
    throw new Error("لا يمكن معالجة أكثر من 10 حجوزات في عملية واحدة.");
  }
  var spreadsheet = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  var invoiceSheet = spreadsheet.getSheetByName("INVOICE");
  if (!invoiceSheet) {
    throw new Error("تعذر العثور على ورقة INVOICE داخل ملف Google Sheets.");
  }
  var reservations = getReservationsByIds_(normalizedIds, spreadsheet);
  var orderedRows = normalizedIds.map(function (reservationId) {
    var row = reservations[reservationId];
    if (!row) {
      throw new Error("لم يتم العثور على بيانات الحجز رقم " + reservationId + ".");
    }
    return row;
  });
  var invoicePayload = buildInvoiceDocumentPayloadFromReservations_(orderedRows, {
    email: targetEmail,
    reservationIds: normalizedIds
  });
  logInvoiceSubmission_(spreadsheet, invoicePayload, "automatic", invoicePayload.reservationId);
  var pdfBlob = generateInvoicePdf_(invoiceSheet, invoicePayload);
  sendInvoiceEmail_(invoicePayload, pdfBlob);
  return "تم إرسال فاتورة واحدة تشمل " + (invoicePayload.lineItems ? invoicePayload.lineItems.length : 1) + " حجوزات";
}

function generateVoucherPdf_(voucherSheet, payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    resetVoucherTemplate_(voucherSheet);
    writeVoucherTemplate_(voucherSheet, payload);
    SpreadsheetApp.flush();
    var spreadsheet = voucherSheet.getParent();
    var fileName = payload.fileName || buildFileSafeName_(payload.customerName || payload.reservationId, "voucher", payload.confirmationId || payload.reservationId);
    return exportSheetAsPdf_(spreadsheet, voucherSheet, fileName, { portrait: true });
  } finally {
    resetVoucherTemplate_(voucherSheet);
    lock.releaseLock();
  }
}

function generateInvoicePdf_(invoiceSheet, payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    resetInvoiceTemplate_(invoiceSheet);
    writeInvoiceTemplate_(invoiceSheet, payload);
    SpreadsheetApp.flush();
    var spreadsheet = invoiceSheet.getParent();
    var fileName = payload.fileName || buildFileSafeName_(payload.customerName || payload.reservationId, "invoice", payload.reservationId);
    return exportSheetAsPdf_(spreadsheet, invoiceSheet, fileName, { portrait: true });
  } finally {
    resetInvoiceTemplate_(invoiceSheet);
    lock.releaseLock();
  }
}

function writeVoucherTemplate_(voucherSheet, payload) {
  voucherSheet.getRange("G3").setValue(payload.reservationId || payload.confirmationId || "");
  voucherSheet.getRange("E11").setValue(payload.customerName || "");
  voucherSheet.getRange("E13").setValue(payload.confirmationId || payload.reservationId || "");
  voucherSheet.getRange("E15").setValue(payload.hotelName || "");
  voucherSheet.getRange("D18").setValue(payload.city || "");
  voucherSheet.getRange("G18").setValue(payload.customerPhone || "");
  voucherSheet.getRange("D20").setValue(payload.displayCheckin || "");
  voucherSheet.getRange("G20").setValue(payload.displayCheckout || "");
  voucherSheet.getRange("D22").setValue(payload.nights || "");
  voucherSheet.getRange("G22").setValue(payload.guests || "");
  voucherSheet.getRange("D24").setValue(payload.roomType || "");
  voucherSheet.getRange("G24").setValue(payload.rooms || "");
  voucherSheet.getRange("D26").setValue(payload.viewType || "");
  voucherSheet.getRange("G26").setValue(payload.mealType || "");
  voucherSheet.getRange("D28").setValue(payload.service || "");
  voucherSheet.getRange("G28").setValue(payload.refundType || "");
  voucherSheet.getRange("C32").setValue(payload.notes || "");
}

function writeInvoiceTemplate_(invoiceSheet, payload) {
  payload = payload || {};
  invoiceSheet.getRange("D10").setValue(payload.customerName || "");
  invoiceSheet.getRange("D11").setValue(payload.customerPhone || "");

  var lineItems = payload.lineItems || [];
  var startRow = 15;
  var maxRows = 10; // rows 15-24
  if (lineItems.length > maxRows) {
    lineItems = lineItems.slice(0, maxRows);
  }
  for (var i = 0; i < maxRows; i++) {
    var currentRow = startRow + i;
    var line = lineItems[i];
    invoiceSheet.getRange(currentRow, 3).setValue(line ? (line.hotel || "") : "");
    invoiceSheet.getRange(currentRow, 5).setValue(line ? (line.type || "") : "");
    invoiceSheet.getRange(currentRow, 6).setValue(line ? (line.checkin || "") : "");
    invoiceSheet.getRange(currentRow, 7).setValue(line ? (line.checkout || "") : "");
    invoiceSheet.getRange(currentRow, 8).setValue(line ? (line.rooms || "") : "");
    invoiceSheet.getRange(currentRow, 9).setValue(line ? (line.nightPrice || "") : "");
    invoiceSheet.getRange(currentRow, 10).setValue(line ? (line.arrived || 0) : "");
    invoiceSheet.getRange(currentRow, 11).setValue(line ? (line.total || 0) : "");
  }

  var calculatedTotal = lineItems.reduce(function (sum, item) {
    return sum + sanitizeNumber_(item.total);
  }, 0);
  var calculatedArrived = lineItems.reduce(function (sum, item) {
    return sum + sanitizeNumber_(item.arrived);
  }, 0);
  var totalSum = (typeof payload.total !== "undefined") ? payload.total : calculatedTotal;
  var arrivedSum = (typeof payload.prepay !== "undefined") ? payload.prepay : calculatedArrived;
  var remainingAmount = (typeof payload.remaining !== "undefined")
    ? payload.remaining
    : Math.max(totalSum - arrivedSum, 0);

  invoiceSheet.getRange("K26").setValue(totalSum);
  invoiceSheet.getRange("K28").setValue(arrivedSum);
  invoiceSheet.getRange("K30").setValue(remainingAmount);
  invoiceSheet.getRange("G34").setValue(remainingAmount);
  invoiceSheet.getRange("A34").setValue(payload.notes || "");
  invoiceSheet.getRange("G46").setValue(remainingAmount);
}

function buildVoucherDetailRows_(payload) {
  var rows = [
    ["رقم الجوال", payload.customerPhone || ""],
    ["البريد الإلكتروني", payload.email || ""],
    ["تاريخ الدخول", payload.displayCheckin || ""],
    ["تاريخ المغادرة", payload.displayCheckout || ""],
    ["عدد الليالي", payload.nights || ""],
    ["عدد الأشخاص", payload.guests || ""],
    ["نوع الغرفة", payload.roomType || ""],
    ["عدد الغرف", payload.rooms || ""],
    ["نوع الإطلالة", payload.viewType || ""],
    ["نوع الوجبة", payload.mealType || ""],
    ["قابلية الإرجاع", payload.refundType || ""]
  ];
  return rows.filter(function (row) {
    return row[1] !== null && row[1] !== "" && typeof row[1] !== "undefined";
  }).slice(0, 11);
}

function normalizeManualVoucherPayload_(raw) {
  var payload = {};
  payload.customerName = (raw.customerName || "").toString().trim();
  payload.confirmationId = (raw.confirmationId || "").toString().trim();
  payload.hotelName = (raw.hotelName || "").toString().trim();
  payload.city = (raw.city || "").toString().trim();
  payload.customerPhone = (raw.customerPhone || "").toString().trim();
  payload.roomType = (raw.roomType || "").toString().trim();
  payload.viewType = (raw.viewType || "").toString().trim();
  payload.service = (raw.service || "").toString().trim();
  payload.mealType = (raw.mealType || "").toString().trim();
  payload.notes = (raw.notes || "").toString().trim();
  payload.email = (raw.email || "").toString().trim();
  payload.refundType = (raw.refundType || "REFUNDABLE").toString().trim() || "REFUNDABLE";
  var checkinDate = parseDateInput_(raw.checkin);
  var checkoutDate = parseDateInput_(raw.checkout);
  if (!payload.customerName || !payload.hotelName || !payload.email || !checkinDate || !checkoutDate) {
    throw new Error("الرجاء تعبئة جميع الحقول المطلوبة لإصدار تأكيد الحجز.");
  }
  var nightsInput = sanitizeNumber_(raw.nights);
  var nights = nightsInput > 0 ? nightsInput : calculateNights_(checkinDate, checkoutDate);
  if (!nights || nights < 1) {
    nights = 1;
  }
  var guests = Math.max(0, Math.round(sanitizeNumber_(raw.guests)));
  var rooms = Math.max(1, Math.round(sanitizeNumber_(raw.rooms) || 1));
  payload.nights = nights;
  payload.guests = guests;
  payload.rooms = rooms;
  payload.checkinDate = checkinDate;
  payload.checkoutDate = checkoutDate;
  payload.displayCheckin = formatPrettyDate_(checkinDate);
  payload.displayCheckout = formatPrettyDate_(checkoutDate);
  payload.reservationId = "";
  if (!payload.confirmationId) {
    payload.confirmationId = generateCustomUniqueId("CN-");
  }
  payload.fileName = buildFileSafeName_(payload.customerName, "voucher", payload.confirmationId);
  payload.source = "manual";
  return payload;
}

function normalizeManualInvoicePayload_(raw) {
  var payload = {};
  payload.customerName = (raw.customerName || "").toString().trim();
  payload.customerPhone = (raw.customerPhone || "").toString().trim();
  payload.hotelName = (raw.hotelName || "").toString().trim();
  payload.type = (raw.type || "").toString().trim() || "حجز فندقي";
  payload.notes = (raw.notes || "").toString().trim();
  payload.email = (raw.email || "").toString().trim();
  var checkinDate = parseDateInput_(raw.checkin);
  var checkoutDate = parseDateInput_(raw.checkout);
  if (!payload.customerName || !payload.hotelName || !payload.email || !checkinDate || !checkoutDate) {
    throw new Error("الرجاء تعبئة جميع الحقول المطلوبة لإصدار الفاتورة اليدوية.");
  }
  var rooms = Math.max(1, Math.round(sanitizeNumber_(raw.rooms) || 1));
  var nightPrice = sanitizeNumber_(raw.nightPrice);
  if (!nightPrice) {
    throw new Error("قيمة الليلة مطلوبة.");
  }
  var nights = calculateNights_(checkinDate, checkoutDate);
  if (!nights || nights < 1) {
    nights = 1;
  }
  var total = nightPrice * rooms * nights;
  var prepay = sanitizeNumber_(raw.prepay);
  if (prepay > total) {
    prepay = total;
  }
  var remaining = Math.max(total - prepay, 0);
  payload.rooms = rooms;
  payload.nightPrice = nightPrice;
  payload.nights = nights;
  payload.total = total;
  payload.prepay = prepay;
  payload.remaining = remaining;
  payload.checkinDate = checkinDate;
  payload.checkoutDate = checkoutDate;
  payload.displayCheckin = formatPrettyDate_(checkinDate);
  payload.displayCheckout = formatPrettyDate_(checkoutDate);
  payload.description = buildInvoiceDescription_(payload.hotelName, payload.type);
  payload.currency = "";
  payload.fileName = buildFileSafeName_(payload.customerName, "invoice", generateCustomUniqueId("INV-"));
  payload.reservationId = "";
  payload.city = "";
  payload.source = "manual";
  payload.lineItems = [{
    hotel: payload.hotelName || payload.description || "",
    type: payload.type || "",
    checkin: payload.displayCheckin || "",
    checkout: payload.displayCheckout || "",
    rooms: rooms,
    nightPrice: nightPrice,
    arrived: prepay,
    total: total
  }];
  return payload;
}

function buildInvoiceLineFromReservation_(row) {
  var checkinDate = parseSheetDate(row[DATABASE_COL_INDEX.CHECKIN_DATE]);
  var checkoutDate = parseSheetDate(row[DATABASE_COL_INDEX.CHECKOUT_DATE]);
  var rooms = Math.max(1, Math.round(sanitizeNumber_(row[DATABASE_COL_INDEX.ROOM_COUNT]) || 1));
  var nights = calculateNights_(checkinDate, checkoutDate);
  if (!nights || nights < 1) {
    nights = 1;
  }
  var total = sanitizeNumber_(row[DATABASE_COL_INDEX.SELLIN_PRICE]);
  if (!total) {
    total = sanitizeNumber_(row[DATABASE_COL_INDEX.SERVICE_SELLING_PRICE]);
  }
  if (!total) {
    total = sanitizeNumber_(row[DATABASE_COL_INDEX.SELLIN_EURO_PRICE]);
  }
  if (!total) {
    total = sanitizeNumber_(row[DATABASE_COL_INDEX.SERVICE_SELLING_EURO_PRICE]);
  }
  var arrived = sanitizeNumber_(row[DATABASE_COL_INDEX.ARRIVED_AMOUNT]);
  var remaining = sanitizeNumber_(row[DATABASE_COL_INDEX.REMAINING_AMOUNT]);
  if (!total) {
    total = arrived + remaining;
  }
  if (!total) {
    total = nights * rooms * sanitizeNumber_(row[DATABASE_COL_INDEX.SELLIN_PRICE]);
  }
  if (!total) {
    total = 0;
  }
  var nightPrice = (rooms * nights) ? total / (rooms * nights) : total;
  return {
    hotel: (row[DATABASE_COL_INDEX.HOTEL] || "").toString().trim(),
    type: (row[DATABASE_COL_INDEX.ROOM_TYPE] || "حجز فندقي").toString().trim() || "حجز فندقي",
    checkin: formatPrettyDate_(checkinDate),
    checkout: formatPrettyDate_(checkoutDate),
    rooms: rooms,
    nightPrice: nightPrice,
    arrived: arrived,
    total: total
  };
}

function buildInvoiceDocumentPayloadFromReservations_(rows, context) {
  context = context || {};
  if (!rows || !rows.length) {
    throw new Error("لا توجد بيانات صالحة لإصدار الفاتورة.");
  }
  var lineItems = rows.map(function (row) {
    return buildInvoiceLineFromReservation_(row);
  });
  var headerRow = rows[0];
  var reservationIds = context.reservationIds || [];
  var notes = rows.map(function (row) {
    return buildAutoNotes_(row, context.note);
  }).filter(function (value) {
    return value;
  }).join(" | ");
  var total = lineItems.reduce(function (sum, item) {
    return sum + sanitizeNumber_(item.total);
  }, 0);
  var prepay = lineItems.reduce(function (sum, item) {
    return sum + sanitizeNumber_(item.arrived);
  }, 0);
  var remaining = Math.max(total - prepay, 0);

  return {
    reservationId: reservationIds.length
      ? reservationIds.join(",")
      : String(headerRow[DATABASE_COL_INDEX.ID] || ""),
    customerName: (headerRow[DATABASE_COL_INDEX.NAME] || "").toString().trim() || "عميل",
    customerPhone: (headerRow[DATABASE_COL_INDEX.PHONE] || "").toString().trim(),
    hotelName: rows.length === 1 ? (headerRow[DATABASE_COL_INDEX.HOTEL] || "").toString().trim() : "",
    type: rows.length === 1
      ? (headerRow[DATABASE_COL_INDEX.ROOM_TYPE] || "حجز فندقي").toString().trim() || "حجز فندقي"
      : "حجز فندقي",
    notes: notes,
    email: context.email,
    currency: (headerRow[DATABASE_COL_INDEX.CURRENCY] || "").toString().trim(),
    lineItems: lineItems,
    total: total,
    prepay: prepay,
    remaining: remaining,
    fileName: buildFileSafeName_(
      (headerRow[DATABASE_COL_INDEX.NAME] || "client").toString().trim() || "client",
      "invoice",
      reservationIds.length ? reservationIds.join("-") : headerRow[DATABASE_COL_INDEX.ID]
    ),
    source: "automatic"
  };
}

function buildVoucherPayloadFromReservationRow_(row, context) {
  context = context || {};
  var checkinDate = parseSheetDate(row[DATABASE_COL_INDEX.CHECKIN_DATE]);
  var checkoutDate = parseSheetDate(row[DATABASE_COL_INDEX.CHECKOUT_DATE]);
  var nights = calculateNights_(checkinDate, checkoutDate);
  if (!nights || nights < 1) {
    nights = 1;
  }
  var rooms = Math.max(1, Math.round(sanitizeNumber_(row[DATABASE_COL_INDEX.ROOM_COUNT]) || 1));
  var guests = Math.max(0, Math.round(sanitizeNumber_(row[DATABASE_COL_INDEX.PERSON_COUNT])));
  var payload = {
    reservationId: String(row[DATABASE_COL_INDEX.ID] || ""),
    confirmationId: (row[DATABASE_COL_INDEX.HOTEL_CONFIRMATION] || row[DATABASE_COL_INDEX.ID] || "").toString().trim(),
    customerName: (row[DATABASE_COL_INDEX.NAME] || "").toString().trim() || "عميل",
    customerPhone: (row[DATABASE_COL_INDEX.PHONE] || "").toString().trim(),
    hotelName: (row[DATABASE_COL_INDEX.HOTEL] || "").toString().trim(),
    city: (row[DATABASE_COL_INDEX.CITY] || "").toString().trim(),
    roomType: (row[DATABASE_COL_INDEX.ROOM_TYPE] || "").toString().trim(),
    viewType: (row[DATABASE_COL_INDEX.VIEW_TYPE] || "").toString().trim(),
    service: (row[DATABASE_COL_INDEX.SERVICE_NAME] || "").toString().trim(),
    mealType: (row[DATABASE_COL_INDEX.MEAL_TYPE] || "").toString().trim(),
    guests: guests,
    rooms: rooms,
    nights: nights,
    checkinDate: checkinDate || "",
    checkoutDate: checkoutDate || "",
    displayCheckin: formatPrettyDate_(checkinDate),
    displayCheckout: formatPrettyDate_(checkoutDate),
    email: context.email,
    notes: buildAutoNotes_(row, context.note),
    refundType: (context.refundType || "REFUNDABLE").toString().trim() || "REFUNDABLE",
    source: "automatic"
  };
  payload.fileName = buildFileSafeName_(payload.customerName || payload.reservationId, "voucher", payload.reservationId);
  return payload;
}

function normalizeReservationIds_(reservationIds) {
  var ids = [];
  if (Array.isArray(reservationIds)) {
    ids = reservationIds;
  } else if (reservationIds) {
    ids = [reservationIds];
  }
  var seen = {};
  var normalized = [];
  ids.forEach(function (value) {
    var id = value !== null && typeof value !== "undefined"
      ? String(value).trim()
      : "";
    if (!id || seen[id]) {
      return;
    }
    seen[id] = true;
    normalized.push(id);
  });
  return normalized;
}

function getReservationsByIds_(reservationIds, spreadsheet) {
  var normalized = reservationIds || [];
  if (!normalized.length) {
    return {};
  }
  var lookup = {};
  normalized.forEach(function (id) {
    lookup[id] = true;
  });
  var workbook = spreadsheet || SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  var sheet = workbook.getSheetByName("DATABASE");
  if (!sheet) {
    throw new Error("تعذر العثور على ورقة DATABASE داخل ملف Google Sheets.");
  }
  var data = sheet.getDataRange().getValues();
  var result = {};
  var found = 0;
  for (var i = 1; i < data.length; i++) {
    var rowId = data[i][DATABASE_COL_INDEX.ID];
    var key = rowId !== null && typeof rowId !== "undefined"
      ? String(rowId)
      : "";
    if (lookup[key]) {
      result[key] = data[i];
      found += 1;
      if (found === normalized.length) {
        break;
      }
    }
  }
  return result;
}

function exportSheetAsPdf_(spreadsheet, sheet, fileName, options) {
  options = options || {};
  var params = {
    format: "pdf",
    size: options.size || "A4",
    portrait: options.portrait !== false ? "true" : "false",
    fitw: "true",
    sheetnames: "false",
    printtitle: "false",
    pagenum: "UNDEFINED",
    gridlines: options.gridlines === true ? "true" : "false",
    fzr: "false",
    top_margin: "0.35",
    bottom_margin: "0.35",
    left_margin: "0.35",
    right_margin: "0.35",
    gid: sheet.getSheetId()
  };
  var query = Object.keys(params).map(function (key) {
    return key + "=" + encodeURIComponent(params[key]);
  }).join("&");
  var url = "https://docs.google.com/spreadsheets/d/" + spreadsheet.getId() + "/export?" + query;
  var response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    throw new Error("تعذر إنشاء ملف PDF: " + response.getContentText());
  }
  var safeName = fileName || "document";
  return response.getBlob().setName(safeName + ".pdf");
}

function ensureSheetWithHeaders_(spreadsheet, sheetName, headers) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
  return sheet;
}

function logVoucherSubmission_(spreadsheet, payload, mode, reservationId) {
  var logSheet = ensureSheetWithHeaders_(spreadsheet, "VOUCHER_LOG", [
    "timestamp",
    "mode",
    "reservationId",
    "customerName",
    "hotel",
    "city",
    "email",
    "confirmationId",
    "sentBy"
  ]);
  logSheet.appendRow([
    new Date(),
    mode,
    reservationId || payload.reservationId || "",
    payload.customerName || "",
    payload.hotelName || "",
    payload.city || "",
    payload.email || "",
    payload.confirmationId || "",
    getExecutionEmail_()
  ]);
}

function logInvoiceSubmission_(spreadsheet, payload, mode, reservationId) {
  var logSheet = ensureSheetWithHeaders_(spreadsheet, "INVOICE_LOG", [
    "timestamp",
    "mode",
    "reservationId",
    "customerName",
    "hotel",
    "city",
    "email",
    "total",
    "prepay",
    "remaining",
    "currency",
    "sentBy"
  ]);
  logSheet.appendRow([
    new Date(),
    mode,
    reservationId || payload.reservationId || "",
    payload.customerName || "",
    payload.hotelName || "",
    payload.city || "",
    payload.email || "",
    payload.total || 0,
    payload.prepay || 0,
    payload.remaining || 0,
    payload.currency || "",
    getExecutionEmail_()
  ]);
}

function sendVoucherEmail_(payload, pdfBlob) {
  MailApp.sendEmail({
    to: payload.email,
    subject: payload.customerName || "عميل",
    body: buildVoucherEmailBody_(payload),
    attachments: [pdfBlob]
  });
}

function sendInvoiceEmail_(payload, pdfBlob) {
  MailApp.sendEmail({
    to: payload.email,
    subject: payload.customerName || "عميل",
    body: buildInvoiceEmailBody_(payload),
    attachments: [pdfBlob]
  });
}

function buildVoucherEmailBody_(payload) {
  var lines = [
    "مرحباً،",
    "",
    "مرفق تأكيد حجز العميل " + (payload.customerName || ""),
    "الفندق: " + (payload.hotelName || "-"),
    "تاريخ الدخول: " + (payload.displayCheckin || "-"),
    "تاريخ المغادرة: " + (payload.displayCheckout || "-"),
    "قابلية الإرجاع: " + (payload.refundType || "-")
  ];
  if (payload.notes) {
    lines.push("");
    lines.push("ملاحظات:", payload.notes);
  }
  lines.push("");
  lines.push("تحيات فريق Ghada Tourism");
  return lines.join("\n");
}

function buildInvoiceEmailBody_(payload) {
  var bookingCount = payload.lineItems ? payload.lineItems.length : 1;
  var lines = [
    "مرحباً،",
    "",
    "مرفق فاتورة العميل " + (payload.customerName || ""),
    "عدد الحجوزات: " + bookingCount
  ];
  if (bookingCount === 1 && payload.lineItems && payload.lineItems[0]) {
    var single = payload.lineItems[0];
    lines.push("الفندق: " + (single.hotel || "-"));
    lines.push("تاريخ الدخول: " + (single.checkin || "-"));
    lines.push("تاريخ المغادرة: " + (single.checkout || "-"));
  }
  lines.push("إجمالي الفاتورة: " + formatMoneyString_(payload.total, payload.currency));
  lines.push("المبلغ المدفوع: " + formatMoneyString_(payload.prepay, payload.currency));
  lines.push("المتبقي: " + formatMoneyString_(payload.remaining, payload.currency));
  if (payload.notes) {
    lines.push("");
    lines.push("ملاحظات:", payload.notes);
  }
  lines.push("");
  lines.push("تحيات فريق Ghada Tourism");
  return lines.join("\n");
}

function buildInvoiceDescription_(hotelName, roomType) {
  var parts = [];
  if (hotelName) {
    parts.push("فندق " + hotelName);
  }
  if (roomType) {
    parts.push(roomType);
  }
  return parts.length ? parts.join(" - ") : "حجز فندقي";
}

function buildAutoNotes_(row, extraNote) {
  var notes = [];
  if (row[DATABASE_COL_INDEX.NOTES]) {
    notes.push(row[DATABASE_COL_INDEX.NOTES]);
  }
  if (extraNote) {
    notes.push(extraNote);
  }
  return notes.join(" | ");
}

function getExecutionEmail_() {
  try {
    var active = Session.getActiveUser().getEmail();
    if (active) {
      return active;
    }
  } catch (err) {}
  try {
    var effective = Session.getEffectiveUser().getEmail();
    if (effective) {
      return effective;
    }
  } catch (e) {}
  return "";
}

function parseDateInput_(value) {
  if (!value) {
    return null;
  }
  if (value instanceof Date) {
    return new Date(value.getTime());
  }
  var parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function formatPrettyDate_(value) {
  var date = value instanceof Date ? value : parseDateInput_(value);
  if (!date) {
    return "";
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

function calculateNights_(checkinDate, checkoutDate) {
  if (!(checkinDate instanceof Date) || isNaN(checkinDate.getTime())) {
    return 0;
  }
  if (!(checkoutDate instanceof Date) || isNaN(checkoutDate.getTime())) {
    return 0;
  }
  var diff = checkoutDate.getTime() - checkinDate.getTime();
  if (diff <= 0) {
    return 0;
  }
  var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  return Math.ceil(diff / MILLIS_PER_DAY);
}

function buildFileSafeName_(customerName, suffix, uniqueToken) {
  var base = (customerName || "client").toString().trim();
  if (!base) {
    base = "client";
  }
  base = base.replace(/\s+/g, "_").replace(/[^\w\u0600-\u06FF\-]/g, "");
  var parts = [base];
  if (suffix) {
    parts.push(suffix);
  }
  if (uniqueToken) {
    parts.push(String(uniqueToken).replace(/[^\w\u0600-\u06FF\-]/g, ""));
  }
  return parts.filter(function (part) { return part; }).join("_");
}

function sanitizeNumber_(value) {
  if (value === null || typeof value === "undefined" || value === "") {
    return 0;
  }
  if (typeof value === "number") {
    return value;
  }
  var parsed = parseFloat(value.toString().replace(/[^\d\.\-]/g, ""));
  return isNaN(parsed) ? 0 : parsed;
}

function formatMoneyString_(amount, currency) {
  if (amount === null || typeof amount === "undefined" || amount === "") {
    return "0" + (currency ? " " + currency : "");
  }
  var value = typeof amount === "number"
    ? Math.round(amount * 100) / 100
    : sanitizeNumber_(amount);
  var suffix = currency ? " " + currency : "";
  return value + suffix;
}

function buildHeaderIndexMap_(headerRow) {
  var map = {};
  (headerRow || []).forEach(function (cell, index) {
    if (cell === null || typeof cell === "undefined") {
      return;
    }
    var key = cell.toString().trim().toLowerCase();
    if (key) {
      map[key] = index;
    }
  });
  return map;
}

function getHeaderIndexByName_(headerMap, names, fallbackIndex) {
  var normalizedFallback = typeof fallbackIndex === "number" ? fallbackIndex : -1;
  if (!headerMap) {
    return normalizedFallback;
  }
  var lookupList = Array.isArray(names) ? names : [names];
  for (var i = 0; i < lookupList.length; i++) {
    var candidate = lookupList[i];
    if (candidate === null || typeof candidate === "undefined") {
      continue;
    }
    var normalized = candidate.toString().trim().toLowerCase();
    if (normalized && typeof headerMap[normalized] === "number") {
      return headerMap[normalized];
    }
  }
  return normalizedFallback;
}

function formatDateValue_(value, timezone) {
  if (!value) {
    return "";
  }
  var date = value instanceof Date ? value : parseDateInput_(value);
  if (!date || isNaN(date.getTime())) {
    return "";
  }
  return Utilities.formatDate(date, timezone || Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function buildSupplierFinancialResponse_(supplierName, startDate, endDate) {
  var timezone = Session.getScriptTimeZone();
  return {
    supplier: supplierName || "",
    filters: {
      startDate: formatDateValue_(startDate, timezone),
      endDate: formatDateValue_(endDate, timezone)
    },
    receivables: [],
    obligations: [],
    totals: {
      receivablesCount: 0,
      obligationsCount: 0,
      receivablesTotal: 0,
      obligationsTotal: 0,
      receivablesByCurrency: {},
      obligationsByCurrency: {}
    }
  };
}

function accumulateCurrencyTotal_(totalsMap, currency, amount) {
  if (!totalsMap) {
    return;
  }
  var key = (currency || "").toString().trim();
  if (!key) {
    key = "غير محدد";
  }
  totalsMap[key] = (totalsMap[key] || 0) + (amount || 0);
}

function updateReservationStatus(reservationId, newStatus) {
  var sheet = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID).getSheetByName("DATABASE");
  if (!sheet) {
    throw new Error("تعذر العثور على ورقة DATABASE");
  }
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][DATABASE_COL_INDEX.ID] == reservationId) {
      sheet.getRange(i + 1, DATABASE_COL_INDEX.RESERVATION_STATUS + 1).setValue(newStatus);
      return { success: true, message: "تم تحديث حالة الحجز بنجاح" };
    }
  }
  
  throw new Error("لم يتم العثور على الحجز رقم " + reservationId);
}

