// ════════════════════════════════════════════════════════
// ARTESANO - Google Apps Script
// 1. Ve a script.google.com → Nuevo proyecto
// 2. Borra todo y pega este código
// 3. Implementar → Nueva implementación → Aplicación web
//    Ejecutar como: Yo | Acceso: Cualquier persona
// 4. Copia la URL y pégala en la app Artesano
// ════════════════════════════════════════════════════════

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "load";
  var output;
  
  try {
    if (action === "setup" || action === "test") {
      var ss = getOrCreate();
      output = JSON.stringify({
        success: true,
        url: ss.getUrl(),
        id: ss.getId(),
        message: "Artesano conectado correctamente"
      });
    } else {
      // Load data
      var ss = getOrCreate();
      var data = { success: true, orders: [], clients: {}, catalog: [] };
      
      var os = ss.getSheetByName("Pedidos");
      if (os && os.getLastRow() > 1) {
        var rows = os.getRange(2, 1, os.getLastRow() - 1, 1).getValues();
        rows.forEach(function(row) {
          if (row[0]) try { data.orders.push(JSON.parse(row[0])); } catch(e) {}
        });
      }
      
      var cs = ss.getSheetByName("Clientes");
      if (cs && cs.getLastRow() > 1) {
        try { data.clients = JSON.parse(cs.getRange(2, 1).getValue()); } catch(e) {}
      }
      
      var cats = ss.getSheetByName("Catalogo");
      if (cats && cats.getLastRow() > 1) {
        try { data.catalog = JSON.parse(cats.getRange(2, 1).getValue()); } catch(e) {}
      }
      
      output = JSON.stringify(data);
    }
  } catch(err) {
    output = JSON.stringify({ success: false, error: err.toString() });
  }
  
  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    
    // Skip test calls
    if (payload.test) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, message: "test ok" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var ss = getOrCreate();
    
    // ── Pedidos ──
    if (payload.orders !== undefined) {
      var sh = ss.getSheetByName("Pedidos") || ss.insertSheet("Pedidos");
      sh.clearContents();
      sh.appendRow(["JSON_Data", "Fecha", "Cliente", "Total", "Estatus"]);
      sh.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#c8e6c9");
      if (payload.orders.length > 0) {
        var rows = payload.orders.map(function(o) {
          return [JSON.stringify(o), o.d || "", o.c || "", o.t || 0, o.s || ""];
        });
        sh.getRange(2, 1, rows.length, 5).setValues(rows);
      }
    }
    
    // ── Clientes ──
    if (payload.clients !== undefined) {
      var sh = ss.getSheetByName("Clientes") || ss.insertSheet("Clientes");
      sh.clearContents();
      sh.appendRow(["JSON_Data"]);
      sh.getRange(1, 1).setFontWeight("bold").setBackground("#c8e6c9");
      sh.getRange(2, 1).setValue(JSON.stringify(payload.clients));
      
      // Human-readable tab
      var sh2 = ss.getSheetByName("Clientes_Vista") || ss.insertSheet("Clientes_Vista");
      sh2.clearContents();
      sh2.appendRow(["Nombre", "Teléfono", "Total $", "Pedidos"]);
      sh2.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#c8e6c9");
      var names = Object.keys(payload.clients);
      if (names.length > 0) {
        var crows = names.map(function(n) {
          var d = payload.clients[n];
          return [n, d.phone || "", (d.total || 0).toFixed(2), d.orders || 0];
        }).sort(function(a, b) { return parseFloat(b[2]) - parseFloat(a[2]); });
        sh2.getRange(2, 1, crows.length, 4).setValues(crows);
      }
    }
    
    // ── Catálogo ──
    if (payload.catalog !== undefined) {
      var sh = ss.getSheetByName("Catalogo") || ss.insertSheet("Catalogo");
      sh.clearContents();
      sh.appendRow(["JSON_Data"]);
      sh.getRange(1, 1).setFontWeight("bold").setBackground("#c8e6c9");
      sh.getRange(2, 1).setValue(JSON.stringify(payload.catalog));
      
      var sh2 = ss.getSheetByName("Catalogo_Vista") || ss.insertSheet("Catalogo_Vista");
      sh2.clearContents();
      sh2.appendRow(["Nombre", "Categoría", "Precio $"]);
      sh2.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#c8e6c9");
      if (payload.catalog.length > 0) {
        var catrows = payload.catalog.map(function(p) { return [p.n, p.cat || "", p.p || 0]; });
        sh2.getRange(2, 1, catrows.length, 3).setValues(catrows);
      }
    }
    
    // ── Log ──
    var log = ss.getSheetByName("Log") || ss.insertSheet("Log");
    if (log.getLastRow() === 0) {
      log.appendRow(["Timestamp", "Pedidos", "Clientes", "Productos"]);
      log.getRange(1,1,1,4).setFontWeight("bold").setBackground("#c8e6c9");
    }
    log.appendRow([
      new Date().toLocaleString("es-VE"),
      (payload.orders || []).length,
      Object.keys(payload.clients || {}).length,
      (payload.catalog || []).length
    ]);
    
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, timestamp: new Date().toISOString() }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreate() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty("ARTESANO_SS_ID");
  
  if (ssId) {
    try { return SpreadsheetApp.openById(ssId); } catch(e) {}
  }
  
  var ss = SpreadsheetApp.create("🌿 Artesano — Base de Datos");
  props.setProperty("ARTESANO_SS_ID", ss.getId());
  
  var sh = ss.getActiveSheet();
  sh.setName("Pedidos");
  sh.appendRow(["JSON_Data", "Fecha", "Cliente", "Total", "Estatus"]);
  sh.getRange(1,1,1,5).setFontWeight("bold").setBackground("#c8e6c9");
  
  ["Clientes","Clientes_Vista","Catalogo","Catalogo_Vista","Log"].forEach(function(name) {
    ss.insertSheet(name);
  });
  
  return ss;
}
