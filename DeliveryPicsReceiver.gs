var DELIVERY_ROOT_FOLDER_ID = '1HCsD3malMvrlAydBit6wed0E2JziVRSW';
var DELIVERY_SPREADSHEET_ID = '1qzcnqszrm_R00Sn-Gd7GYqF38pfg8ATWIM9WpMai0Q8';
var DELIVERY_SHEET_NAME = 'Sheet1';
var DELIVERY_TIME_ZONE = 'Asia/Kolkata';
var DELIVERY_MAX_BYTES = 25 * 1024 * 1024;
var DELIVERY_HEADERS = [
  'Uploaded At (IST)', 'Upload ID', 'Staff ID / Name', 'Store Code', 'Hub Name',
  'City', 'State', 'Photo File Name', 'Drive File URL', 'Photo Type', 'Photo Size (MB)'
];

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'uploadStatus') return uploadStatusResponse_(e);
  return HtmlService.createHtmlOutput(
    'Cars24 delivery photo receiver is ready. Run setupDeliveryPhotoLogging() once before accepting uploads.'
  );
}

function setupDeliveryPhotoLogging() {
  var sheet = getDeliverySheet_();
  ensureDeliveryHeaders_(sheet);
  sheet.setFrozenRows(1);
  if (!sheet.getFilter()) {
    sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), DELIVERY_HEADERS.length).createFilter();
  }
  return 'Delivery photo tracking is ready: ' + sheet.getParent().getUrl() + ' (' + DELIVERY_SHEET_NAME + ')';
}

function doPost(e) {
  var payload = {};
  var savedFile = null;

  try {
    payload = JSON.parse((e && e.parameter && e.parameter.payload) || '{}');
    var image = validateAndDecode_(payload);
    var now = new Date();
    var day = Utilities.formatDate(now, DELIVERY_TIME_ZONE, 'yyyy-MM-dd');
    var time = Utilities.formatDate(now, DELIVERY_TIME_ZONE, 'HHmmss');

    var folderName = safeName_(payload.storeCode + ' ' + payload.hubName);
    var rootFolder = DriveApp.getFolderById(DELIVERY_ROOT_FOLDER_ID);
    var hubFolder = getOrCreateFolder_(rootFolder, folderName);
    var extension = image.mimeType === 'image/png' ? '.png' :
      image.mimeType === 'image/webp' ? '.webp' : '.jpg';
    var fileName = safeName_(payload.hubName) + '_' + day + '_' + time +
      '_' + Utilities.getUuid().split('-')[0] + extension;

    savedFile = hubFolder.createFile(
      Utilities.newBlob(image.bytes, image.mimeType, fileName)
    );
    savedFile.setDescription(
      'CARS24 buy-side delivery photo. Store: ' + payload.storeCode +
      '; staff: ' + payload.staffId +
      '; uploaded: ' + Utilities.formatDate(now, DELIVERY_TIME_ZONE, "yyyy-MM-dd'T'HH:mm:ssXXX")
    );

    var sheet;
    var logged = false;
    var logError = '';
    try {
      sheet = getDeliverySheet_();
      ensureDeliveryHeaders_(sheet);
      sheet.appendRow([
        Utilities.formatDate(now, DELIVERY_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss'),
        payload.requestId,
        String(payload.staffId).trim(),
        payload.storeCode,
        payload.hubName,
        payload.city || '',
        payload.state || '',
        fileName,
        savedFile.getUrl(),
        image.mimeType,
        Number((image.bytes.length / (1024 * 1024)).toFixed(2))
      ]);
      logged = true;
    } catch (sheetError) {
      logError = String(sheetError && sheetError.message || sheetError);
    }

    if (logged) {
      try {
        sheet.setFrozenRows(1);
        var filter = sheet.getFilter();
        if (filter) {
          filter.setRange(sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), DELIVERY_HEADERS.length));
        }
      } catch (formatError) {
        // Formatting errors must not undo a saved photo or tracking row.
      }
    }

    if (!logged) {
      return response_({
        ok: true,
        logged: false,
        requestId: payload.requestId,
        fileName: fileName,
        fileId: savedFile.getId(),
        fileUrl: savedFile.getUrl(),
        message: 'Photo saved in Drive, but sheet logging failed: ' + logError
      });
    }

    return response_({
      ok: true,
      logged: true,
      requestId: payload.requestId,
      fileName: fileName,
      fileId: savedFile.getId(),
      fileUrl: savedFile.getUrl(),
      message: 'Upload confirmed: the photo was saved in Google Drive and logged in the tracking sheet.'
    });
  } catch (error) {
    var message = String(error && error.message || error);

    // Never trash a created photo because sheet logging or formatting failed.
    if (savedFile) {
      return response_({
        ok: true,
        logged: false,
        requestId: payload.requestId || '',
        fileName: savedFile.getName(),
        fileId: savedFile.getId(),
        fileUrl: savedFile.getUrl(),
        message: 'Photo saved in Drive, but tracking could not be confirmed: ' + message
      });
    }

    return response_({
      ok: false,
      requestId: payload.requestId || '',
      error: message,
      message: 'Upload was not completed: ' + message
    });
  }
}

function validateAndDecode_(p) {
  if (p.service !== 'cars24-delivery-pics' || p.eventType !== 'delivery_photo_upload') {
    throw new Error('Unsupported request type.');
  }
  if (!/^[\w-]{8,80}$/.test(String(p.requestId || ''))) {
    throw new Error('Invalid request ID.');
  }
  if (p.intent !== 'buy' || p.code !== 'C1' || !/^C1\/[\w -]+$/.test(String(p.storeCode || ''))) {
    throw new Error('Only a C1 buy hub can upload here.');
  }
  if (!String(p.staffId || '').trim() || String(p.staffId).length > 80) {
    throw new Error('Staff ID or name is required.');
  }
  if (!String(p.hubName || '').trim() || String(p.hubName).length > 220) {
    throw new Error('Invalid hub name.');
  }

  var mime = String(p.mimeType || '').toLowerCase();
  if (['image/jpeg', 'image/png', 'image/webp'].indexOf(mime) < 0) {
    throw new Error('Use a JPG, PNG, or WebP image.');
  }
  var encoded = String(p.imageBase64 || '');
  if (!encoded || encoded.length > Math.ceil(DELIVERY_MAX_BYTES * 4 / 3) + 8) {
    throw new Error('Photo is empty or exceeds the 25 MB limit.');
  }
  var bytes = Utilities.base64Decode(encoded);
  if (!bytes.length || bytes.length > DELIVERY_MAX_BYTES) {
    throw new Error('Photo is empty or exceeds the 25 MB limit.');
  }

  var sig = bytes.slice(0, 12).map(function (b) { return (b + 256) % 256; });
  var isJpeg = sig[0] === 0xff && sig[1] === 0xd8 && sig[2] === 0xff;
  var isPng = sig[0] === 0x89 && sig[1] === 0x50 && sig[2] === 0x4e && sig[3] === 0x47;
  var isWebp = sig[0] === 0x52 && sig[1] === 0x49 && sig[2] === 0x46 && sig[3] === 0x46 &&
    sig[8] === 0x57 && sig[9] === 0x45 && sig[10] === 0x42 && sig[11] === 0x50;
  if (!(mime === 'image/jpeg' && isJpeg || mime === 'image/png' && isPng || mime === 'image/webp' && isWebp)) {
    throw new Error('The selected file does not match its image type.');
  }
  return { bytes: bytes, mimeType: mime };
}

function getOrCreateFolder_(parent, name) {
  var folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function getDeliverySheet_() {
  var spreadsheet = SpreadsheetApp.openById(DELIVERY_SPREADSHEET_ID);
  var sheet = spreadsheet.getSheetByName(DELIVERY_SHEET_NAME);
  if (!sheet) throw new Error('Tracking tab "' + DELIVERY_SHEET_NAME + '" was not found.');
  return sheet;
}

function ensureDeliveryHeaders_(sheet) {
  var current = sheet.getRange(1, 1, 1, DELIVERY_HEADERS.length).getDisplayValues()[0];
  var blank = current.every(function (value) { return !String(value || '').trim(); });
  if (blank) {
    sheet.getRange(1, 1, 1, DELIVERY_HEADERS.length).setValues([DELIVERY_HEADERS]);
    sheet.getRange(1, 1, 1, DELIVERY_HEADERS.length)
      .setFontWeight('bold').setBackground('#eef0f4');
  } else if (DELIVERY_HEADERS.some(function (value, index) {
    return current[index] !== value;
  })) {
    throw new Error('The tracking sheet header row does not match the expected delivery-photo columns.');
  }
}

function safeName_(value) {
  var safe = String(value || '').normalize('NFKD')
    .replace(/[^A-Za-z0-9_-]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 90);
  return safe || 'Cars24_Hub';
}

function response_(data) {
  data.service = 'cars24-delivery-pics';
  if (data.requestId) {
    try {
      CacheService.getScriptCache().put('delivery-upload:' + data.requestId, JSON.stringify(data), 21600);
    } catch (cacheError) {}
  }
  var json = JSON.stringify(data).replace(/</g, '\\u003c');
  var html = '<!doctype html><html><head><meta charset="utf-8"></head><body>' +
    '<script>(function(){var result=' + json + ';' +
    'try{window.top.postMessage(result,"*");}' +
    'catch(e){try{window.parent.postMessage(result,"*");}catch(ignored){}}' +
    '})();</script></body></html>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function uploadStatusResponse_(e) {
  var requestId = String(e.parameter.requestId || '');
  var callback = String(e.parameter.callback || '');
  if (!/^[\\w-]{8,80}$/.test(requestId) || !/^[A-Za-z_$][A-Za-z0-9_$]{0,80}$/.test(callback)) {
    return ContentService.createTextOutput('void 0;').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  var result = null;
  try {
    var cached = CacheService.getScriptCache().get('delivery-upload:' + requestId);
    if (cached) result = JSON.parse(cached);
  } catch (cacheReadError) {}
  if (!result) result = findUploadStatusInSheet_(requestId);
  if (!result) result = {service:'cars24-delivery-pics',requestId:requestId,pending:true,ok:null};
  result.service = 'cars24-delivery-pics';
  var json = JSON.stringify(result).replace(/</g, '\\u003c');
  return ContentService.createTextOutput(callback + '(' + json + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function findUploadStatusInSheet_(requestId) {
  try {
    var sheet = getDeliverySheet_(), lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    var match = sheet.getRange(2,2,lastRow-1,1).createTextFinder(requestId).matchEntireCell(true).findNext();
    if (!match) return null;
    var row = sheet.getRange(match.getRow(),1,1,DELIVERY_HEADERS.length).getValues()[0];
    return {
      service:'cars24-delivery-pics',requestId:requestId,ok:true,logged:true,
      fileName:String(row[7]||''),fileUrl:String(row[8]||''),
      message:'Upload confirmed: the photo was saved in Google Drive and logged in the tracking sheet.'
    };
  } catch (lookupError) { return null; }
}
