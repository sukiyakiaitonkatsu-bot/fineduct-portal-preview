// Fineduct order submissions -> Google Sheet + email alert
// Deploy: Extensions > Apps Script in the sheet, paste, Deploy > New deployment > Web app
// (Execute as: Me, Who has access: Anyone). Copy the /exec URL into SCRIPT_URL in fineduct-order.html.
const SHEET_ID = '10_XIMV5rgaBbnPSsMMNC1AgdkDO_IeRJzJAWPnwqkQY';
const ALERT_TO = 'louis.lee@gtapcoltd.com';

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const t = d.totals || {};
    const id = d.orderId || ('FD-' + Date.now());
    ss.getSheetByName('Orders').appendRow([new Date(), id, d.client || '', d.company || '', d.contact || '', d.email || '',
      d.po || '', d.container || '', t.fill || '', t.pkg || 0, t.ctn || 0, t.gw || 0, t.cu || 0, d.notes || '']);
    const rows = (d.lines || []).map(l => [id, l.no || '(missing)', l.code, l.pkg, l.ctn, l.qty, l.cu, l.gw]);
    if (rows.length) { const sh = ss.getSheetByName('Lines'); sh.getRange(sh.getLastRow() + 1, 1, rows.length, 8).setValues(rows); }
    const body = [
      'New Fineduct order submitted', '',
      'Order ID: ' + id, 'Client: ' + (d.client || ''), 'Company: ' + (d.company || ''),
      'Contact: ' + (d.contact || ''), 'Email: ' + (d.email || ''), 'PO ref: ' + (d.po || ''), '',
      'Container: ' + (d.container || 'none fits') + (t.fill ? ' (' + t.fill + '% full)' : ''),
      'Packages: ' + t.pkg + ' | Cartons: ' + t.ctn + ' | Gross: ' + t.gw + ' kg | Volume: ' + t.cu + ' CUFT', '',
      'Lines:'].concat((d.lines || []).map(l => '  ' + l.code + ' (' + (l.no || 'no item no.') + ') x ' + l.pkg + ' pkg = ' + l.ctn + ' ctn, ' + l.cu + ' CUFT'))
      .concat(['', 'Notes: ' + (d.notes || '-'), '', 'Sheet: ' + ss.getUrl()]).join('\n');
    MailApp.sendEmail({ to: ALERT_TO, subject: 'New Fineduct order ' + id + ' - ' + (d.company || d.client || ''), body: body,
      replyTo: d.email && /@/.test(d.email) ? d.email : undefined });
    return ContentService.createTextOutput(JSON.stringify({ ok: true, id: id })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally { lock.releaseLock(); }
}

// Run once from the editor to grant permissions and test the email.
function testSubmit() {
  doPost({ postData: { contents: JSON.stringify({ orderId: 'TEST-1', client: 'Test', company: 'Test Co', email: '', po: 'PO-TEST',
    container: '20ft', totals: { fill: 60, pkg: 10, ctn: 10, gw: 36, cu: 19.87 },
    lines: [{ no: '52202-11', code: 'FA-80I', pkg: 10, ctn: 10, qty: 100, cu: 19.87, gw: 36 }] }) } });
}
