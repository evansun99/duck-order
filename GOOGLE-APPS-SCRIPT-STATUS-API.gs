// 俞记黑鸭官网：顾客订单状态与营业日期查询
// 绑定表格：俞记黑鸭订单生产与卤汤管理

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.status === '1') return json_(lookupOrderStatus_(p.orderNumber, p.phoneLast4));
    if (p.settings === '1') return json_(getBusinessSettings_());
    return json_({ ok: true, service: 'Yuji Black Duck customer status API' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function lookupOrderStatus_(orderNumber, phoneLast4) {
  const orderNo = String(orderNumber || '').trim().toUpperCase();
  const last4 = String(phoneLast4 || '').replace(/\D/g, '').slice(-4);
  if (!/^(YJ\d{6}|MAN\d{6})$/.test(orderNo) || last4.length !== 4) {
    return { ok: false, found: false, error: '请输入正确的订单号和手机号后四位' };
  }

  const book = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = book.getSheetByName('订单总表');
  if (!orderSheet) throw new Error('订单总表不存在');
  const lastRow = orderSheet.getLastRow();
  if (lastRow < 5) return { ok: true, found: false };
  const values = orderSheet.getRange(5, 1, lastRow - 4, 16).getDisplayValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    const storedNo = String(row[1] || '').trim().toUpperCase();
    const storedPhone = String(row[7] || '').replace(/\D/g, '');
    if (storedNo !== orderNo || storedPhone.slice(-4) !== last4) continue;

    const detailSheet = book.getSheetByName('订单商品明细');
    const items = [];
    if (detailSheet && detailSheet.getLastRow() >= 5) {
      const detailRows = detailSheet.getRange(5, 1, detailSheet.getLastRow() - 4, 13).getDisplayValues();
      detailRows.forEach(function (detail) {
        if (String(detail[1] || '').trim().toUpperCase() !== orderNo) return;
        items.push({
          name: String(detail[5] || ''),
          quantity: String(detail[6] || ''),
          unit: String(detail[7] || ''),
          source: String(detail[8] || ''),
          lineQuantity: String(detail[9] || ''),
          lineAmount: String(detail[10] || '')
        });
      });
    }

    return {
      ok: true,
      found: true,
      orderNumber: storedNo,
      status: String(row[2] || '已确认'),
      deliveryDate: String(row[4] || ''),
      fulfilmentMethod: String(row[5] || ''),
      packaging: String(row[8] || ''),
      itemSubtotal: String(row[10] || ''),
      deliveryFee: String(row[11] || ''),
      total: String(row[12] || ''),
      paymentMethod: String(row[13] || ''),
      orderDescription: String(row[15] || ''),
      items: items
    };
  }
  return { ok: true, found: false };
}

function getBusinessSettings_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('营业设置');
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, closedDates: [] };
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getDisplayValues();
  const closedDates = rows.filter(function (row) {
    return row[0] && row[1] === '打烊';
  }).map(function (row) {
    return { date: row[0], status: row[1], scope: row[2] || '全部', message: row[3] || '本日打烊' };
  });
  return { ok: true, closedDates: closedDates };
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
