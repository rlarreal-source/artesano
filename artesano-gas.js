// ════════════════════════════════════════════════════════
// ARTESANO — Google Apps Script v4
// INSTRUCCIONES:
// 1. script.google.com → abre tu proyecto
// 2. Borra TODO el código existente
// 3. Pega este código completo
// 4. Ctrl+S para guardar
// 5. Implementar → Administrar implementaciones → editar
//    → Nueva versión → Implementar
// ════════════════════════════════════════════════════════

var SHEET_NAME = "Artesano Base de Datos";

// Punto de entrada GET — también maneja preflight OPTIONS
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "load";
  var data;

  try {
    if (action === "test" || action === "setup") {
      var ss = getSheet();
      data = { success: true, message: "Conexion OK", sheetUrl: ss.getUrl() };
    } else {
      data = loadData();
    }
  } catch(err) {
    data = { success: false, error: err.message };
  }

  return respond(data);
}

// Punto de entrada POST — guarda datos
function doPost(e) {
  var data;
  try {
    var body = JSON.parse(e.postData.contents);
    data = saveData(body);
  } catch(err) {
    data = { success: false, error: err.message };
  }
  return respond(data);
}

// Respuesta con headers permisivos
function respond(data) {
  var output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ── Cargar datos ──
function loadData() {
  var ss = getSheet();
  var result = { success: true, orders: [], clients: {}, catalog: [] };

  var osh = ss.getSheetByName("Pedidos");
  if (osh && osh.getLastRow() > 1) {
    var rows = osh.getRange(2, 1, osh.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (rows[i][0]) {
        try { result.orders.push(JSON.parse(rows[i][0])); } catch(ex) {}
      }
    }
  }

  var csh = ss.getSheetByName("Clientes");
  if (csh && csh.getLastRow() > 1) {
    try { result.clients = JSON.parse(csh.getRange(2, 1).getValue()); } catch(ex) {}
  }

  var catsh = ss.getSheetByName("Catalogo");
  if (catsh && catsh.getLastRow() > 1) {
    try { result.catalog = JSON.parse(catsh.getRange(2, 1).getValue()); } catch(ex) {}
  }

  return result;
}

// ── Guardar datos ──
function saveData(body) {
  if (body.test) return { success: true, message: "test ok" };
  var ss = getSheet();

  // Pedidos
  if (body.orders !== undefined) {
    var sh = ss.getSheetByName("Pedidos") || ss.insertSheet("Pedidos");
    sh.clearContents();
    sh.appendRow(["JSON", "Fecha", "Cliente", "Total", "Estatus", "Pago"]);
    sh.getRange(1,1,1,6).setFontWeight("bold").setBackground("#c8e6c9");
    if (body.orders.length > 0) {
      var rows = body.orders.map(function(o) {
        return [JSON.stringify(o), o.d||"", o.c||"", o.t||0, o.s||"", o.pg||""];
      });
      sh.getRange(2, 1, rows.length, 6).setValues(rows);
    }
  }

  // Clientes
  if (body.clients !== undefined) {
    var csh = ss.getSheetByName("Clientes") || ss.insertSheet("Clientes");
    csh.clearContents();
    csh.appendRow(["JSON"]); csh.getRange(1,1).setFontWeight("bold").setBackground("#c8e6c9");
    csh.getRange(2,1).setValue(JSON.stringify(body.clients));

    // Vista legible
    var cv = ss.getSheetByName("Clientes_Vista") || ss.insertSheet("Clientes_Vista");
    cv.clearContents();
    cv.appendRow(["Nombre","Telefono","Total","Pedidos"]);
    cv.getRange(1,1,1,4).setFontWeight("bold").setBackground("#c8e6c9");
    var ns = Object.keys(body.clients);
    if (ns.length > 0) {
      var crows = ns.map(function(n){var d=body.clients[n];return[n,d.phone||"",d.total||0,d.orders||0];});
      crows.sort(function(a,b){return b[2]-a[2];});
      cv.getRange(2,1,crows.length,4).setValues(crows);
    }
  }

  // Catalogo
  if (body.catalog !== undefined) {
    var catsh = ss.getSheetByName("Catalogo") || ss.insertSheet("Catalogo");
    catsh.clearContents();
    catsh.appendRow(["JSON"]); catsh.getRange(1,1).setFontWeight("bold").setBackground("#c8e6c9");
    catsh.getRange(2,1).setValue(JSON.stringify(body.catalog));

    var catv = ss.getSheetByName("Catalogo_Vista") || ss.insertSheet("Catalogo_Vista");
    catv.clearContents();
    catv.appendRow(["Nombre","Categoria","Precio"]);
    catv.getRange(1,1,1,3).setFontWeight("bold").setBackground("#c8e6c9");
    if (body.catalog.length > 0) {
      var catr = body.catalog.map(function(p){return[p.n,p.cat||"",p.p||0];});
      catv.getRange(2,1,catr.length,3).setValues(catr);
    }
  }

  // Log
  var log = ss.getSheetByName("Log") || ss.insertSheet("Log");
  if (log.getLastRow() === 0) {
    log.appendRow(["Timestamp","Pedidos","Clientes","Catalogo"]);
    log.getRange(1,1,1,4).setFontWeight("bold").setBackground("#c8e6c9");
  }
  log.appendRow([new Date().toLocaleString("es-VE"), (body.orders||[]).length, Object.keys(body.clients||{}).length, (body.catalog||[]).length]);

  return { success: true, timestamp: new Date().toISOString() };
}

// ── Obtener o crear el spreadsheet ──
function getSheet() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("ART_ID");
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch(e) {}
  }
  var ss = SpreadsheetApp.create(SHEET_NAME);
  props.setProperty("ART_ID", ss.getId());
  var s = ss.getActiveSheet(); s.setName("Pedidos");
  s.appendRow(["JSON","Fecha","Cliente","Total","Estatus","Pago"]);
  s.getRange(1,1,1,6).setFontWeight("bold").setBackground("#c8e6c9");
  ["Clientes","Clientes_Vista","Catalogo","Catalogo_Vista","Log"].forEach(function(n){ss.insertSheet(n);});
  return ss;
}
