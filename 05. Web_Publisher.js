// ============================================================
//  MaKing Trips — Web Publisher v1
//  Genera un endpoint JSON público para el sitio web
//  Despliega como Web App: Ejecutar como "Yo", acceso "Cualquiera"
// ============================================================

const WEB_SHEET = 'PAQUETES_V3';
const WEB_ESTADO_PUBLICAR = 'Web';
const WEB_FILA_HEADER = 2;

// ── COLUMNAS PAQUETES_V3 (base 1) ────────────────────────────
const WP_CODIGO       = 1;   // A
const WP_PAIS         = 2;   // B — País
const WP_DESTINO      = 3;   // C
const WP_TIPO_PAQUETE = 4;   // D — Todo Incluido, Solo Hotel, etc.
const WP_ADULTOS      = 6;   // F
const WP_NINOS        = 7;   // G
const WP_PERSONAS     = 8;   // H
const WP_FECHA_SAL    = 9;   // I
const WP_FECHA_REG    = 10;  // J
const WP_NOCHES       = 11;  // K
const WP_DIAS         = 12;  // L
const WP_INCL_VUELOS  = 13;  // M
const WP_TIPO_VIAJE   = 14;  // N
const WP_AEROLINEA    = 15;  // O
const WP_INCL_HOTEL   = 19;  // S
const WP_HOTEL        = 20;  // T
const WP_TIPO_HOTEL   = 21;  // U
const WP_INCL_TRASLADO = 23; // W
const WP_INCL_ACT     = 27;  // AA
const WP_INCL_TOURS   = 31;  // AE
const WP_INCL_SEGURO  = 35;  // AI
const WP_PRECIO_PP    = 52;  // AZ — Precio por persona
const WP_PRECIO_TRANSF = 51; // AY — Precio transferencia 5%
const WP_PRECIO_TOT   = 50;  // AX — Precio total
const WP_TARJETA_1    = 53;  // BA
const WP_CUOTA_12     = 61;  // BI
const WP_VALIDEZ      = 65;  // BM
const WP_ESTADO       = 67;  // BO — Estado: Cotizando / Publicado / Web


// ── LIMPIAR NÚMERO ────────────────────────────────────────────
function toNumWP(val) {
  if (typeof val === 'number') return val;
  return Number(String(val).replace(/[^0-9.-]/g, '')) || 0;
}

// ── NORMALIZAR TEXTO PARA COMPARAR ───────────────────────────
function normalizarKey(texto) {
  return String(texto || '').toLowerCase()
    .replace(/á/g,'a').replace(/é/g,'e').replace(/í/g,'i')
    .replace(/ó/g,'o').replace(/ú/g,'u').replace(/ñ/g,'n')
    .replace(/[^a-z0-9]/g,'') // eliminar espacios y caracteres especiales
    .trim();
}

// ── LEER DICCIONARIO DE FOTOS DESDE SHEET ───────────────────
// Hoja: FOTOS_DESTINOS | Columna A: destino | Columna B: url_foto
function leerFotosDestinos() {
  var fotos = {};
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('FOTOS_DESTINOS');
    if (!sheet) return fotos;
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) { // saltar encabezado
      var destino = normalizarKey(data[i][0]);
      var url     = String(data[i][1] || '').trim();
      if (destino && url) fotos[destino] = url;
    }
  } catch(e) {
    Logger.log('Error leyendo FOTOS_DESTINOS: ' + e.message);
  }
  return fotos;
}

// ── FOTO DEL DESTINO ─────────────────────────────────────────
// Busca primero en FOTOS_DESTINOS del Sheet, luego en Drive
var _fotosCache = null; // cache para no leer el sheet múltiples veces por request

function fotoUrlPublica(destino) {
  // Cargar diccionario del Sheet (una sola vez por ejecución)
  if (!_fotosCache) _fotosCache = leerFotosDestinos();

  var key = normalizarKey(destino);

  // 1. Buscar en diccionario del Sheet (coincidencia parcial)
 // Primero buscar coincidencia exacta
  if (_fotosCache[key]) return _fotosCache[key];
  
  // Luego coincidencia parcial — más específico primero (más largo gana)
  var mejorKey = null;
  var mejorLen = 0;
  for (var k in _fotosCache) {
    if (key.indexOf(k) !== -1 || k.indexOf(key) !== -1) {
      if (k.length > mejorLen) {
        mejorKey = k;
        mejorLen = k.length;
      }
    }
  }
  if (mejorKey) return _fotosCache[mejorKey];

  // 2. Buscar archivo en Drive como fallback
  try {
    var carpetas = DriveApp.getFoldersByName('Fotos Destinos MaKing Trips');
    if (carpetas.hasNext()) {
      var carpeta = carpetas.next();
      var exts = ['.jpg','.jpeg','.png','.webp'];
      for (var i = 0; i < exts.length; i++) {
        var archivos = carpeta.getFilesByName(key + exts[i]);
        if (archivos.hasNext()) {
          var file = archivos.next();
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          return 'https://drive.google.com/uc?export=view&id=' + file.getId();
        }
      }
    }
  } catch(e) {}

  // 3. Sin foto — el widget usará el fallback de avión
  return '';
}




// ── ENDPOINT PRINCIPAL (doGet) ────────────────────────────────
// Este es el endpoint que llama la web para obtener los paquetes
function doGet(e) {
  var output = {};
  var callback = e && e.parameter && e.parameter.callback ? e.parameter.callback : null;
  _fotosCache = null; // Siempre leer FOTOS_DESTINOS fresco

  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(WEB_SHEET);
    var data  = sheet.getRange(1, 1, sheet.getLastRow(), 67).getValues();
    var paquetes = [];

    for (var i = WEB_FILA_HEADER; i < data.length; i++) {
      var row = data[i];

      // Solo filas con estado "Web"
      if (String(row[WP_ESTADO - 1]).trim() !== WEB_ESTADO_PUBLICAR) continue;
      // Solo filas con código
      if (!row[WP_CODIGO - 1]) continue;

      var destino    = String(row[WP_DESTINO - 1] || '');
      var precioTot  = toNumWP(row[WP_PRECIO_TOT - 1]);
      var precioPP   = toNumWP(row[WP_PRECIO_PP  - 1]);
      var pTransf    = toNumWP(row[WP_PRECIO_TRANSF - 1]);
      if (pTransf <= 0) pTransf = Math.floor(precioTot * 0.95 / 100) * 100 + 99;
      var tarjeta1   = toNumWP(row[WP_TARJETA_1 - 1]);
      if (tarjeta1 <= 0) tarjeta1 = precioTot;
      var ahorro     = Math.round(tarjeta1 - pTransf);
      var cuota12    = toNumWP(row[WP_CUOTA_12 - 1]);

      // Fechas
      var fechaSal = row[WP_FECHA_SAL - 1];
      var fechaReg = row[WP_FECHA_REG - 1];
      var fechaStr = '';
      var MESES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
      if (fechaSal instanceof Date && fechaReg instanceof Date) {
        var dSal = fechaSal.getDate();
        var mSal = MESES[fechaSal.getMonth()];
        var dReg = fechaReg.getDate();
        var mReg = MESES[fechaReg.getMonth()];
        if (mSal === mReg) {
          fechaStr = dSal + ' al ' + dReg + ' ' + mReg;
        } else {
          fechaStr = dSal + ' ' + mSal + ' al ' + dReg + ' ' + mReg;
        }
      }

      // Fecha validez
      var validezStr = '';
      var validez = row[WP_VALIDEZ - 1];
      if (validez instanceof Date) {
        var MESES_L = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
        validezStr = validez.getDate() + ' de ' + MESES_L[validez.getMonth()] + ' de ' + validez.getFullYear();
      }

      // Incluidos
      function siInc(val) {
        var v = String(val).toLowerCase().trim().replace(/í/g,'i');
        return v === 'si' || v === 'yes' || v === '1' || v === 'true';
      }
      var incluidos = [];
      if (siInc(row[WP_INCL_VUELOS  - 1])) incluidos.push({ tipo: 'vuelo',    texto: 'Vuelos ' + String(row[WP_TIPO_VIAJE - 1] || 'incluidos'), detalle: String(row[WP_AEROLINEA - 1] || '') });
      if (siInc(row[WP_INCL_HOTEL   - 1])) incluidos.push({ tipo: 'hotel',    texto: String(row[WP_HOTEL - 1] || 'Hotel'), detalle: String(row[WP_TIPO_HOTEL - 1] || '') });
      if (siInc(row[WP_INCL_TRASLADO- 1])) incluidos.push({ tipo: 'traslado', texto: 'Traslados incluidos', detalle: '' });
      if (siInc(row[WP_INCL_ACT     - 1])) incluidos.push({ tipo: 'actividad',texto: 'Actividades', detalle: '' });
      if (siInc(row[WP_INCL_TOURS   - 1])) incluidos.push({ tipo: 'tour',     texto: 'Tours incluidos', detalle: '' });
      if (siInc(row[WP_INCL_SEGURO  - 1])) incluidos.push({ tipo: 'seguro',   texto: 'Seguro de viaje', detalle: '' });

      paquetes.push({
        codigo      : String(row[WP_CODIGO - 1]),
        destino     : destino,
        tipoPaquete : String(row[WP_TIPO_PAQUETE - 1] || ''),
        personas    : toNumWP(row[WP_PERSONAS - 1]),
        noches      : toNumWP(row[WP_NOCHES - 1]),
        dias        : toNumWP(row[WP_DIAS - 1]),
        fechaViaje  : fechaStr,
        validez     : validezStr,
        aerolinea   : String(row[WP_AEROLINEA - 1] || ''),
        hotel       : String(row[WP_HOTEL - 1] || ''),
        tipoHotel   : String(row[WP_TIPO_HOTEL - 1] || ''),
        precioPP    : precioPP,
        precioTot   : precioTot,
        precioTransf: pTransf,
        ahorro      : ahorro,
        cuota12     : cuota12,
        incluidos   : incluidos,
        foto        : fotoUrlPublica(destino),
      });
    }

    output = { ok: true, paquetes: paquetes, total: paquetes.length, generado: new Date().toISOString() };

  } catch(err) {
    output = { ok: false, error: err.message };
  }

  var json = JSON.stringify(output);

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}


// ── FUNCIÓN DE PRUEBA (ejecutar manualmente) ──────────────────
// Corre esta función en el editor para ver el JSON en los logs
function probarEndpoint() {
  var resultado = doGet({});
  var json = JSON.parse(resultado.getContent());
  Logger.log('Paquetes encontrados: ' + json.total);
  Logger.log(JSON.stringify(json, null, 2));
}