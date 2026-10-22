const DELIVERY_ROOT_FOLDER_ID = '1LCYjoCpnv7gKQcjVzUcPNizu4w-as6QC';
const DELIVERY_SPREADSHEET_ID = '1qzcnqszrm_R00Sn-Gd7GYqF38pfg8ATWIM9WpMai0Q8';
const DELIVERY_SHEET_NAME = 'Sheet1';
const DELIVERY_TIME_ZONE = 'Asia/Kolkata';
const DELIVERY_MAX_BYTES = 5 * 1024 * 1024;
const DELIVERY_ALLOWED_ORIGIN = 'https://ampcars24-cmyk.github.io';
const DELIVERY_HEADERS = [
  'Uploaded At (IST)', 'Upload ID', 'Staff ID / Name', 'Store Code', 'Hub Name',
  'City', 'State', 'Photo File Name', 'Drive File URL', 'Photo Type', 'Photo Size (MB)'
];

function doGet() {
  return HtmlService.createHtmlOutput('Cars24 delivery photo receiver is ready. Run setupDeliveryPhotoLogging() once before accepting uploads.');
}

function setupDeliveryPhotoLogging() {
  const sheet = getDeliverySheet_();
  ensureDeliveryHeaders_(sheet);
  sheet.setFrozenRows(1);
  if (!sheet.getFilter()) sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), DELIVERY_HEADERS.length).createFilter();
  return 'Delivery photo tracking is ready: ' + sheet.getParent().getUrl() + ' (' + DELIVERY_SHEET_NAME + ')';
}

function doPost(e) {
  let deliveryPicsPayload;
  try {
    deliveryPicsPayload = JSON.parse((e && e.parameter && e.parameter.payload) || '{}');
    const deliveryPicsImage = validateAndDecode_(deliveryPicsPayload);
    const deliveryPicsNow = new Date();
    const deliveryPicsDay = Utilities.formatDate(deliveryPicsNow, DELIVERY_TIME_ZONE, 'yyyy-MM-dd');
    const deliveryPicsTime = Utilities.formatDate(deliveryPicsNow, DELIVERY_TIME_ZONE, 'HHmmss');
    const deliveryPicsHubFolderName = safeName_(deliveryPicsPayload.storeCode + ' ' + deliveryPicsPayload.hubName);
    const deliveryPicsRootFolder = DriveApp.getFolderById(DELIVERY_ROOT_FOLDER_ID);
    const deliveryPicsHubFolder = getOrCreateFolder_(deliveryPicsRootFolder, deliveryPicsHubFolderName);
    const deliveryPicsExtension = deliveryPicsImage.mimeType === 'image/png' ? '.png' : deliveryPicsImage.mimeType === 'image/webp' ? '.webp' : '.jpg';
    const deliveryPicsRequestTag = Utilities.getUuid().split('-')[0];
    const deliveryPicsFileName = safeName_(deliveryPicsPayload.hubName) + '_' + deliveryPicsDay + '_' + deliveryPicsTime + '_' + deliveryPicsRequestTag + deliveryPicsExtension;
    const deliveryPicsSavedFile = deliveryPicsHubFolder.createFile(Utilities.newBlob(deliveryPicsImage.bytes, deliveryPicsImage.mimeType, deliveryPicsFileName));
    deliveryPicsSavedFile.setDescription('CARS24 buy-side delivery photo. Store: ' + deliveryPicsPayload.storeCode + '; staff: ' + deliveryPicsPayload.staffId + '; uploaded: ' + Utilities.formatDate(deliveryPicsNow, DELIVERY_TIME_ZONE, "yyyy-MM-dd'T'HH:mm:ssXXX"));
    try {
      const deliveryPicsTrackingSheet = getDeliverySheet_();
      ensureDeliveryHeaders_(deliveryPicsTrackingSheet);
      const deliveryPicsUploadTime = Utilities.formatDate(deliveryPicsNow, DELIVERY_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');
      deliveryPicsTrackingSheet.appendRow([
        deliveryPicsUploadTime, deliveryPicsPayload.requestId, String(deliveryPicsPayload.staffId).trim(), deliveryPicsPayload.storeCode,
        deliveryPicsPayload.hubName, deliveryPicsPayload.city || '', deliveryPicsPayload.state || '', deliveryPicsFileName,
        deliveryPicsSavedFile.getUrl(), deliveryPicsImage.mimeType, Number((deliveryPicsImage.bytes.length / (1024 * 1024)).toFixed(2))
      ]);
      deliveryPicsTrackingSheet.setFrozenRows(1);
      const deliveryPicsFilter = deliveryPicsTrackingSheet.getFilter();
      if (deliveryPicsFilter) deliveryPicsFilter.setRange(deliveryPicsTrackingSheet.getRange(1, 1, Math.max(deliveryPicsTrackingSheet.getLastRow(), 1), DELIVERY_HEADERS.length));
    } catch (deliveryPicsLogError) {
      deliveryPicsSavedFile.setTrashed(true);
      throw new Error('Could not log this upload to the tracking sheet, so the photo was not kept. ' + String(deliveryPicsLogError && deliveryPicsLogError.message || deliveryPicsLogError));
    }
    return response_({
      ok: true,
      requestId: deliveryPicsPayload.requestId,
      fileName: deliveryPicsFileName,
      fileId: deliveryPicsSavedFile.getId(),
      fileUrl: deliveryPicsSavedFile.getUrl(),
      message: 'Upload confirmed: the photo was saved in Google Drive and logged in the tracking sheet.'
    });
  } catch (deliveryPicsError) {
    const deliveryPicsErrorDetails = String(deliveryPicsError && deliveryPicsError.message || deliveryPicsError);
    return response_({ ok: false, requestId: deliveryPicsPayload && deliveryPicsPayload.requestId || '', error: deliveryPicsErrorDetails, message: 'Upload was not confirmed. ' + deliveryPicsErrorDetails });
  }
}

function validateAndDecode_(p) {
  if (p.service !== 'cars24-delivery-pics' || p.eventType !== 'delivery_photo_upload') throw new Error('Unsupported request type.');
  if (!/^[\w-]{8,80}$/.test(String(p.requestId || ''))) throw new Error('Invalid request ID.');
  if (p.intent !== 'buy' || p.code !== 'C1' || !/^C1\/[\w -]+$/.test(String(p.storeCode || ''))) throw new Error('Only a C1 buy hub can upload here.');
  if (!String(p.staffId || '').trim() || String(p.staffId).length > 80) throw new Error('Staff ID or name is required.');
  if (!String(p.hubName || '').trim() || String(p.hubName).length > 220) throw new Error('Invalid hub name.');
  const mime = String(p.mimeType || '').toLowerCase();
  if (['image/jpeg', 'image/png', 'image/webp'].indexOf(mime) < 0) throw new Error('Use a JPG, PNG, or WebP image.');
  const encoded = String(p.imageBase64 || '');
  if (!encoded || encoded.length > Math.ceil(DELIVERY_MAX_BYTES * 4 / 3) + 8) throw new Error('Photo is empty or exceeds the 5 MB limit.');
  const bytes = Utilities.base64Decode(encoded);
  if (!bytes.length || bytes.length > DELIVERY_MAX_BYTES) throw new Error('Photo is empty or exceeds the 5 MB limit.');
  const sig = bytes.slice(0, 12).map(b => (b + 256) % 256);
  const isJpeg = sig[0] === 0xff && sig[1] === 0xd8 && sig[2] === 0xff;
  const isPng = sig[0] === 0x89 && sig[1] === 0x50 && sig[2] === 0x4e && sig[3] === 0x47;
  const isWebp = sig[0] === 0x52 && sig[1] === 0x49 && sig[2] === 0x46 && sig[3] === 0x46 && sig[8] === 0x57 && sig[9] === 0x45 && sig[10] === 0x42 && sig[11] === 0x50;
  if (!(mime === 'image/jpeg' && isJpeg || mime === 'image/png' && isPng || mime === 'image/webp' && isWebp)) throw new Error('The selected file does not match its image type.');
  return { bytes: bytes, mimeType: mime };
}

function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function getDeliverySheet_() {
  const spreadsheet = SpreadsheetApp.openById(DELIVERY_SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(DELIVERY_SHEET_NAME);
  if (!sheet) throw new Error('Tracking tab "' + DELIVERY_SHEET_NAME + '" was not found.');
  return sheet;
}

function ensureDeliveryHeaders_(sheet) {
  const current = sheet.getRange(1, 1, 1, DELIVERY_HEADERS.length).getDisplayValues()[0];
  const blank = current.every(value => !String(value || '').trim());
  if (blank) {
    sheet.getRange(1, 1, 1, DELIVERY_HEADERS.length).setValues([DELIVERY_HEADERS]);
    sheet.getRange(1, 1, 1, DELIVERY_HEADERS.length).setFontWeight('bold').setBackground('#eef0f4');
  } else if (DELIVERY_HEADERS.some((value, index) => current[index] !== value)) {
    throw new Error('The tracking sheet header row does not match the expected delivery-photo columns.');
  }
}

function safeName_(value) {
  const safe = String(value || '').normalize('NFKD').replace(/[^A-Za-z0-9_-]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 90);
  return safe || 'Cars24_Hub';
}

function response_(data) {
  const json = JSON.stringify(data).replace(/</g, '\\u003c');
  const html = '<!doctype html><meta charset="utf-8"><script>parent.postMessage(Object.assign({service:"cars24-delivery-pics"},' + json + '),' + JSON.stringify(DELIVERY_ALLOWED_ORIGIN) + ');</script>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
