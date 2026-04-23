// ═══════════════════════════════════════════════════════════════
// AeA FROTA – BACKEND Google Apps Script  v3
// Ayuda en Acción Moçambique · Gestão de Frota
// ═══════════════════════════════════════════════════════════════

const SHEET_NAME = 'AeA_Frota_Requisicoes';
const TAB_NAME   = 'Pedidos';

const HEADERS = [
  'id','created','status','nome','dep','tel','email',
  'dataSaida','dataRegresso','horaSaida','horaRegresso',
  'dias','partida','destino','npax','pax','motivo',
  'urgencia','projecto','carga',
  'viaturaId','viaturaNome','motorista','contactoMotorista',
  'dataAprovacao','obsGestor'
];

function getSheet() {
  let ss;
  const files = DriveApp.getFilesByName(SHEET_NAME);
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(SHEET_NAME);
    const sheet = ss.getActiveSheet();
    sheet.setName(TAB_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setBackground('#F26522')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  let sheet = ss.getSheetByName(TAB_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TAB_NAME);
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function rowToObj(row) {
  const obj = {};
  HEADERS.forEach((h, i) => obj[h] = row[i] !== undefined ? String(row[i]) : '');
  return obj;
}

function objToRow(obj) {
  return HEADERS.map(h => obj[h] !== undefined ? obj[h] : '');
}

// ── Resposta JSON com suporte JSONP (resolve CORS) ──
function jsonResponse(data, callback) {
  const json = JSON.stringify(data);
  if (callback) {
    // JSONP – contorna CORS para pedidos GET do browser
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET ──
function doGet(e) {
  try {
    const action   = (e.parameter && e.parameter.action)   ? e.parameter.action   : 'getAll';
    const callback = (e.parameter && e.parameter.callback) ? e.parameter.callback : null;

    if (action === 'ping') {
      const sheet = getSheet();
      const rows = Math.max(0, sheet.getLastRow() - 1);
      return jsonResponse({ status: 'ok', rows: rows, timestamp: new Date().toISOString() }, callback);
    }

    if (action === 'getAll') {
      const sheet = getSheet();
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return jsonResponse([], callback);
      const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
      const requests = data
        .filter(row => row[0] && String(row[0]).trim() !== '')
        .map(rowToObj);
      return jsonResponse(requests, callback);
    }

    return jsonResponse({ error: 'Acção desconhecida: ' + action }, callback);

  } catch(err) {
    Logger.log('doGet error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

// ── POST ──
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action || 'saveAll';

    if (action === 'addOne') {
      const req = body.data;
      if (!req || !req.id) throw new Error('Pedido inválido: falta id');
      const sheet = getSheet();
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String);
        if (ids.includes(String(req.id))) {
          const rowIdx = ids.indexOf(String(req.id));
          sheet.getRange(rowIdx + 2, 1, 1, HEADERS.length).setValues([objToRow(req)]);
          colorRow(sheet, rowIdx + 2, req.status);
          return jsonResponse({ status: 'ok', action: 'updated', id: req.id });
        }
      }
      sheet.appendRow(objToRow(req));
      colorRow(sheet, sheet.getLastRow(), req.status);
      return jsonResponse({ status: 'ok', action: 'added', id: req.id });
    }

    if (action === 'updateOne') {
      const req = body.data;
      const sheet = getSheet();
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return jsonResponse({ status: 'not_found' });
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String);
      const rowIdx = ids.indexOf(String(req.id));
      if (rowIdx === -1) return jsonResponse({ status: 'not_found', id: req.id });
      sheet.getRange(rowIdx + 2, 1, 1, HEADERS.length).setValues([objToRow(req)]);
      colorRow(sheet, rowIdx + 2, req.status);
      return jsonResponse({ status: 'ok', action: 'updated', id: req.id });
    }

    if (action === 'saveAll') {
      const requests = body.data;
      if (!Array.isArray(requests)) throw new Error('data deve ser um array');
      const sheet = getSheet();
      if (sheet.getLastRow() > 1) {
        sheet.deleteRows(2, sheet.getLastRow() - 1);
      }
      if (requests.length > 0) {
        const rows = requests.map(objToRow);
        sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
        requests.forEach((req, i) => colorRow(sheet, i + 2, req.status));
      }
      return jsonResponse({ status: 'ok', saved: requests.length });
    }

    return jsonResponse({ error: 'Acção desconhecida: ' + action });

  } catch(err) {
    Logger.log('doPost error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

function colorRow(sheet, rowNum, status) {
  let color = '#FFFFFF';
  if (status === 'Aprovado')  color = '#E8F7EE';
  if (status === 'Rejeitado') color = '#FDECEA';
  if (status === 'Pendente')  color = '#FFF5E0';
  sheet.getRange(rowNum, 1, 1, HEADERS.length).setBackground(color);
}

function teste() {
  const sheet = getSheet();
  Logger.log('✔ Linhas: ' + Math.max(0, sheet.getLastRow() - 1));
  const ping = JSON.parse(doGet({ parameter: { action: 'ping' } }).getContent());
  Logger.log('✔ Ping: ' + JSON.stringify(ping));
}
