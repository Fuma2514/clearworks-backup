// ClearWorks backend cadangan — deploy manual oleh Fuma (no KYC):
// 1. Buka script.google.com (login Gmail lokomotifkeren@gmail.com)
// 2. Paste semua kode ini, lalu Deploy > New deployment > Web app
//    Execute as: Me | Who has access: Anyone
// 3. Salin Web app URL ke index.html (ganti __APPS_SCRIPT_URL__) dan ke pesan WA/IG

const SHEET_ID = ''; // kosong = buat sheet baru otomatis

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.create('ClearWorks Orders');
  let sh = ss.getSheetByName('orders') || ss.insertSheet('orders');
  if (sh.getLastRow() === 0) sh.appendRow(['timestamp','ref','name','email','service','brief']);
  const ref = 'CW-' + Math.floor(100000 + Math.random()*900000);
  sh.appendRow([new Date(), ref, d.name, d.email, d.service, d.brief]);
  GmailApp.sendEmail(Session.getEffectiveUser().getEmail(),
    'Order baru ' + ref + ' — ' + d.service,
    'Nama: ' + d.name + '\nEmail: ' + d.email + '\nLayanan: ' + d.service + '\nBrief: ' + d.brief);
  return ContentService.createTextOutput(JSON.stringify({ok:true, ref:ref})).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.create('ClearWorks Orders');
  const sh = ss.getSheetByName('orders') || ss.insertSheet('orders');
  const data = sh.getDataRange().getValues();
  return ContentService.createTextOutput(JSON.stringify({ok:true, orders:data.slice(1)})).setMimeType(ContentService.MimeType.JSON);
}
