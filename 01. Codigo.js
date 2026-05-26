// ============================================================
//  MaKing Trips — Botón "Enviar por WhatsApp"  v6
//  Actualizado para PAQUETES_V3
// ============================================================

// ── HOJA Y COLUMNAS (base 1) ──────────────────────────────────
const SHEET_PAQUETES    = 'PAQUETES_V3';
const COL_CODIGO        = 1;   // A
const COL_ESTADO        = 2;   // B
const COL_DESTINO       = 3;   // C
const COL_ADULTOS       = 6;   // F
const COL_NINOS         = 7;   // G
const COL_PERSONAS      = 8;   // H
const COL_FECHA_SAL     = 9;   // I
const COL_FECHA_REG     = 10;  // J
const COL_NOCHES        = 11;  // K
const COL_DIAS          = 12;  // L
const COL_TIPO_VIAJE    = 14;  // N
const COL_AEROLINEA     = 15;  // O
const COL_HORARIO_SAL   = 16;  // P
const COL_HORARIO_REG   = 17;  // Q
const COL_HOTEL         = 20;  // T
const COL_TIPO_HOTEL    = 21;  // U
const COL_INCL_TRASLADO = 23;  // W
const COL_PROV_TRASLADO = 24;  // X
const COL_TIPO_TRASLADO = 25;  // Y
const COL_INCL_ACT      = 27;  // AA
const COL_PROV_ACT      = 28;  // AB
const COL_DET_ACT       = 29;  // AC
const COL_INCL_TOURS    = 31;  // AE
const COL_PROV_TOURS    = 32;  // AF
const COL_DET_TOURS     = 33;  // AG
const COL_INCL_SEGURO   = 35;  // AI
const COL_PROV_SEGURO   = 36;  // AJ
const COL_DET_SEGURO    = 37;  // AK
const COL_INCL_OTROS    = 39;  // AM
const COL_PROV_OTROS    = 40;  // AN
const COL_DET_OTROS     = 41;  // AO
const COL_PRECIO_TOT    = 50;  // AX
const COL_PRECIO_TRANSF = 51;  // AY
const COL_PRECIO_PP     = 52;  // AZ
const COL_TARJETA_1     = 53;  // BA
const COL_12C_TOTAL     = 60;  // BH
const COL_CUOTA_12      = 61;  // BI
const COL_VALIDEZ       = 65;  // BM
const COL_TEXTO_PUB     = 66;  // BN
const COL_LINK_HOTEL    = 68;  // BP
const FILA_HEADER       = 2;// Fila del encabezado (base 1)
// URL de la Web App (directa, sin intermediario)
const WEBAPP_URL = 'https://script.google.com/macros/s/AKfycby9J92SfLi3PzUbDumFzshQf9kcmjyApJVDQIzJ_CTBxvsLdTCr9Noett-CiOva74DK/exec';
const BITLY_URL = 'https://bit.ly/makingtrips-';

// ── MAPA DE REEMPLAZOS DE EMOJIS ─────────────────────────────
const REEMPLAZOS = [
  ['\u2708\uFE0F', '[Vuelo]'],
  ['\u26A0\uFE0F', '(!)'],
  ['\u2708',       '[Vuelo]'],
  ['\u2714',       '\u00BB'],
  ['\u1F4B0',      '$'],
  ['\u26A0',       '(!)'],
  ['\uFE0F',       ''],
  ['\u1F44B',      ''],
  ['\u1F4B5',      '$'],
  ['\u1F4B3',      ''],
  ['\u23F3',       ''],
  ['\u1F60A',      ''],
  ['\u1F30D',      ''],
];

function limpiarEmojis(str) {
  let resultado = String(str);
  REEMPLAZOS.forEach(function(par) {
    while (resultado.indexOf(par[0]) !== -1) {
      resultado = resultado.replace(par[0], par[1]);
    }
  });
  return resultado;
}

// ── MENÚ ──────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🌎 MaKing Trips')
    .addItem('📱 Enviar WhatsApp',           'enviarPorWhatsApp')
    .addItem('📄 Ver cotización HTML',        'generarYAbrirHTML')
    .addItem('📧 Enviar cotización por mail', 'enviarPorCorreo')
    .addSeparator()
    .addItem('🎨 Generar artes para redes',   'generarArtes')
    .addSeparator()
    .addItem('🔗 Abrir Web App',              'abrirWebApp')
    .addToUi();
}

// ── FUNCIÓN PRINCIPAL ─────────────────────────────────────────
function enviarPorWhatsApp() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PAQUETES);
  const fila  = sheet.getActiveCell().getRow();
  const ui    = SpreadsheetApp.getUi();

  if (fila <= FILA_HEADER) {
    ui.alert('Selecciona una fila de paquete, no el encabezado.');
    return;
  }

  const datos = sheet.getRange(fila, 1, 1, COL_LINK_HOTEL).getValues()[0];

  const codigo       = datos[COL_CODIGO        - 1];
  const estado       = datos[COL_ESTADO        - 1];
  const destino      = datos[COL_DESTINO       - 1];
  const adultos      = datos[COL_ADULTOS       - 1];
  const ninos        = datos[COL_NINOS         - 1];
  const personas     = datos[COL_PERSONAS      - 1];
  const fechaSal     = datos[COL_FECHA_SAL     - 1];
  const fechaReg     = datos[COL_FECHA_REG     - 1];
  const noches       = datos[COL_NOCHES        - 1];
  const dias         = datos[COL_DIAS          - 1];
  const tipoViaje    = datos[COL_TIPO_VIAJE    - 1];
  const aerolinea    = datos[COL_AEROLINEA     - 1];
  const hotel        = datos[COL_HOTEL         - 1];
  const tipoHotel    = datos[COL_TIPO_HOTEL    - 1];
  const horarioSal   = datos[COL_HORARIO_SAL   - 1];
  const horarioReg   = datos[COL_HORARIO_REG   - 1];
  const inclTraslado = datos[COL_INCL_TRASLADO - 1];
  const provTraslado = datos[COL_PROV_TRASLADO - 1];
  const tipoTraslado = datos[COL_TIPO_TRASLADO - 1];
  const inclAct      = datos[COL_INCL_ACT      - 1];
  const provAct      = datos[COL_PROV_ACT      - 1];
  const detAct       = datos[COL_DET_ACT       - 1];
  const inclTours    = datos[COL_INCL_TOURS    - 1];
  const provTours    = datos[COL_PROV_TOURS    - 1];
  const detTours     = datos[COL_DET_TOURS     - 1];
  const inclSeguro   = datos[COL_INCL_SEGURO   - 1];
  const provSeguro   = datos[COL_PROV_SEGURO   - 1];
  const detSeguro    = datos[COL_DET_SEGURO    - 1];
  const inclOtros    = datos[COL_INCL_OTROS    - 1];
  const provOtros    = datos[COL_PROV_OTROS    - 1];
  const detOtros     = datos[COL_DET_OTROS     - 1];
  const precioTot    = datos[COL_PRECIO_TOT    - 1];
  const precioTransf = datos[COL_PRECIO_TRANSF - 1];
  const precioPP     = datos[COL_PRECIO_PP     - 1];
  const tarjeta1     = datos[COL_TARJETA_1     - 1];
  const cuota12      = datos[COL_CUOTA_12      - 1];
  const validez      = datos[COL_VALIDEZ       - 1];
  const textoPub     = datos[COL_TEXTO_PUB     - 1];
  const linkHotel    = datos[COL_LINK_HOTEL    - 1];

  if (!codigo || !textoPub) {
    ui.alert('La fila no tiene Codigo o Texto de Publicacion.');
    return;
  }

  // URL cotización — directa a la Web App
  var urlCotizacion = WEBAPP_URL + '?codigo=' + encodeURIComponent(codigo);

  const mensaje = construirMensaje(
    textoPub, adultos, ninos, personas,
    fechaSal, fechaReg, noches, dias,
    aerolinea, horarioSal, horarioReg,
    hotel, tipoHotel,
    inclTraslado, provTraslado, tipoTraslado,
    inclAct, provAct, detAct,
    inclTours, provTours, detTours,
    inclSeguro, provSeguro, detSeguro,
    inclOtros, provOtros, detOtros,
    precioTot, precioTransf, precioPP, tarjeta1, cuota12,
    validez, linkHotel, urlCotizacion
  );

  // Modal con preview y botones
  const urlWA = 'https://wa.me/50232270977?text=' + encodeURIComponent(mensaje);
  const html  = HtmlService
    .createHtmlOutput(construirHtmlPreview(destino, codigo, estado, mensaje, urlWA))
    .setWidth(520)
    .setHeight(600);
  ui.showModalDialog(html, '📱 WhatsApp — ' + destino);
}

// ── CONSTRUIR MENSAJE ─────────────────────────────────────────
function construirMensaje(
  textoPub, adultos, ninos, personas,
  fechaSal, fechaReg, noches, dias,
  aerolinea, horarioSal, horarioReg,
  hotel, tipoHotel,
  inclTraslado, provTraslado, tipoTraslado,
  inclAct, provAct, detAct,
  inclTours, provTours, detTours,
  inclSeguro, provSeguro, detSeguro,
  inclOtros, provOtros, detOtros,
  precioTot, precioTransf, precioPP, tarjeta1, cuota12,
  validez, linkHotel, urlCotizacion
) {
  const fmtQ = function(n) { return 'Q' + Math.round(Number(n)).toLocaleString('es-GT'); };
  const siInc = function(v) {
    var s = String(v||'').toLowerCase().trim().replace(/í/g,'i');
    return s === 'si' || s === 'yes' || s === 's' || s === '1' || s === 'true';
  };

  const MESES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  const MESES_CORTOS = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];

  const fmtFechaCorta = function(f) {
    if (!f || !(f instanceof Date)) return '';
    return f.getDate() + '/' + MESES_CORTOS[f.getMonth()] + '/' + String(f.getFullYear()).slice(-2);
  };
  const fmtFechaLarga = function(f) {
    if (!f || !(f instanceof Date)) return '';
    return f.getDate() + ' de ' + MESES[f.getMonth()] + ' de ' + f.getFullYear();
  };

  // Precios
  const pTransf  = (precioTransf && Number(precioTransf) > 0)
    ? Number(precioTransf)
    : Math.floor(Number(precioTot) * 0.95 / 100) * 100 + 99;
  const pTarjeta = (tarjeta1 && Number(tarjeta1) > 0) ? Number(tarjeta1) : Number(precioTot);
  const ahorro   = Math.round(pTarjeta - pTransf);

  // Pasajeros
  const numAdultos = Number(adultos) || 0;
  const numNinos   = Number(ninos)   || 0;
  const numTotal   = Number(personas) || (numAdultos + numNinos);

  let lineaPasajeros = '';
  if (numNinos > 0) {
    lineaPasajeros = '👨‍👩‍👧 ' + numAdultos + ' adulto' + (numAdultos !== 1 ? 's' : '') +
      ' + ' + numNinos + ' niño' + (numNinos !== 1 ? 's' : '') + ' (' + numTotal + ' pax)';
  } else {
    lineaPasajeros = '👨‍👩‍👧 ' + numTotal + ' persona' + (numTotal !== 1 ? 's' : '');
  }

  // Duración — días primero ☀️
  let lineaDuracion = (noches && dias) ? '☀️ ' + dias + 'D / ' + noches + 'N' : '';

  // Aerolíneas separadas por "/"
  let aeroIda = String(aerolinea || ''), aeroVuelta = String(aerolinea || '');
  if (aeroIda.indexOf('/') !== -1) {
    var partes = aeroIda.split('/');
    aeroIda    = partes[0].trim();
    aeroVuelta = partes[1].trim();
  }

  // Bloque de incluidos
  var incluye = [];

  if (fechaSal instanceof Date || fechaReg instanceof Date) {
    incluye.push('✈️ *Vuelos*');
    if (fechaSal instanceof Date) {
      var lineaSal = '   🛫 ' + fmtFechaCorta(fechaSal) + ' · ' + aeroIda;
      if (horarioSal) lineaSal += ' · ' + horarioSal;
      incluye.push(lineaSal);
    }
    if (fechaReg instanceof Date) {
      var lineaReg = '   🛬 ' + fmtFechaCorta(fechaReg) + ' · ' + aeroVuelta;
      if (horarioReg) lineaReg += ' · ' + horarioReg;
      incluye.push(lineaReg);
    }
  }

  var hotelNombre = String(hotel || '').trim();
  var hotelTipo   = String(tipoHotel || '').trim();
  if (hotelNombre) {
    var hotelLine = '🏨 *Hotel:* ' + hotelNombre;
    if (hotelTipo) hotelLine += ' · ' + hotelTipo;
    incluye.push(hotelLine);
  }

  if (siInc(inclTraslado)) {
    var det = [provTraslado, tipoTraslado].filter(function(x){ return x && String(x).trim(); }).join(' · ');
    incluye.push('🚌 *Traslados:* ' + det);
  }

  if (siInc(inclAct)) {
    var det = [provAct, detAct].filter(function(x){ return x && String(x).trim(); }).join(' · ');
    incluye.push('⭐ *Actividades:* ' + det);
  }

  if (siInc(inclTours)) {
    var det = [provTours, detTours].filter(function(x){ return x && String(x).trim(); }).join(' · ');
    incluye.push('🗺️ *Tours:* ' + det);
  }

  if (siInc(inclSeguro)) {
    var det = [provSeguro, detSeguro].filter(function(x){ return x && String(x).trim(); }).join(' · ');
    incluye.push('🛡️ *Seguro:* ' + det);
  }

  if (siInc(inclOtros)) {
    var det = [provOtros, detOtros].filter(function(x){ return x && String(x).trim(); }).join(' · ');
    incluye.push('➕ *Adicionales:* ' + det);
  }

  // Armar mensaje
  let msg = '👑 ¡Hola! Gracias por confiar en la familia MarroKing de MaKing Trips 👑\n';
  msg += 'Hemos trabajado en tu cotización y aquí te dejamos todos los detalles:\n\n';
  msg += '📋 ' + lineaPasajeros;
  if (lineaDuracion) msg += ' · ' + lineaDuracion;
  msg += '\n';

  if (incluye.length > 0) {
    msg += '\n✅ *Incluye:*\n';
    incluye.forEach(function(s){ msg += s + '\n'; });
  }

  msg += '\n💰 *Precios (' + numTotal + ' pax)*\n';
  msg += '   💳 Tarjeta: *' + fmtQ(pTarjeta) + '*\n';
  msg += '   🏦 Transferencia: *' + fmtQ(pTransf) + '* (ahorras ' + fmtQ(ahorro) + ')\n';
  msg += '   📆 12 cuotas de *' + fmtQ(cuota12) + '/mes*\n';

  if (validez && validez instanceof Date) {
    msg += '\n⏳ Válido hasta ' + fmtFechaLarga(validez) + '\n';
  }

  msg += '\n📄 *Cotización completa:* ' + urlCotizacion;
  msg += '\n\n¡Aquí para cualquier consulta! 🙌';
  msg += '\n\n📧 _Si el link no abre, revisa tu correo — ahí también está tu cotización._';

  return msg;
}

// ── HTML PREVIEW ──────────────────────────────────────────────
function construirHtmlPreview(destino, codigo, estado, mensaje, urlWhatsApp) {
  const esc = function(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  };
  const fmtWA = function(s) {
    return esc(s)
      .replace(/\*(.*?)\*/g, '<strong>$1</strong>')
      .replace(/_(.*?)_/g,   '<em>$1</em>')
      .replace(/\n/g, '<br>');
  };

  return '<!DOCTYPE html>' +
    '<html><head><meta charset="UTF-8"><style>' +
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:Arial,sans-serif;font-size:14px;background:#f5f7fa;padding:16px;color:#111}' +
    '.hdr{background:#1a1a2e;color:#C9922E;border-radius:10px;padding:14px 18px;margin-bottom:14px}' +
    '.hdr h2{font-size:15px;font-weight:600;margin-bottom:2px;color:#C9922E}' +
    '.hdr small{font-size:11px;opacity:.7;color:#fff}' +
    '.badge{display:inline-block;font-size:10px;padding:1px 7px;border-radius:20px;background:#2AABB5;color:#fff;margin-left:6px}' +
    '.lbl{font-size:10px;color:#666;text-transform:uppercase;letter-spacing:.05em;font-weight:600;margin-bottom:5px}' +
    '.bubble{background:#dcf8c6;border-radius:10px 10px 2px 10px;padding:13px 15px;font-size:13px;line-height:1.8;color:#111;max-height:340px;overflow-y:auto;border:1px solid #c5e8a3;word-break:break-word}' +
    '.btns{display:flex;gap:8px;margin-top:14px}' +
    '.btn{flex:1;padding:12px;font-size:13px;font-weight:700;text-align:center;border:none;border-radius:10px;cursor:pointer;text-decoration:none}' +
    '.btn-copy{background:#2AABB5;color:#fff}' +
    '.btn-wa{background:#25d366;color:#fff}' +
    '.copied{font-size:11px;color:#2AABB5;text-align:center;margin-top:6px;height:16px}' +
    '.note{font-size:11px;color:#999;text-align:center;margin-top:7px}' +
    '</style></head><body>' +
    '<div class="hdr">' +
    '<h2>' + esc(String(destino)) + ' <span class="badge">' + esc(String(codigo)) + '</span></h2>' +
    '<small>Estado: ' + esc(String(estado)) + '</small>' +
    '</div>' +
    '<p class="lbl">Vista previa del mensaje</p>' +
    '<div class="bubble" id="msg">' + fmtWA(mensaje) + '</div>' +
    '<div class="btns">' +
    '<button class="btn btn-copy" onclick="copiar()">📋 Copiar mensaje</button>' +
    '<a class="btn btn-wa" href="' + urlWhatsApp + '" target="_blank">💬 Abrir WhatsApp Web</a>' +
    '</div>' +
    '<div class="copied" id="aviso"></div>' +
    '<p class="note">El mensaje se abrirá listo para enviar en WhatsApp Web.</p>' +
    '<script>' +
    'function copiar(){' +
    'var txt=document.getElementById("msg").innerText;' +
    'if(navigator.clipboard){navigator.clipboard.writeText(txt).then(function(){ok();}).catch(function(){fb(txt);});}else{fb(txt);}' +
    '}' +
    'function fb(t){var ta=document.createElement("textarea");ta.value=t;document.body.appendChild(ta);ta.select();document.execCommand("copy");document.body.removeChild(ta);ok();}' +
    'function ok(){document.getElementById("aviso").textContent="✓ Mensaje copiado al portapapeles";setTimeout(function(){document.getElementById("aviso").textContent="";},2500);}' +
    '<\/script>' +
    '</body></html>';
}

// ── VER COTIZACIÓN HTML ───────────────────────────────────────
function generarYAbrirHTML() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_PAQUETES);
    const fila  = sheet.getActiveCell().getRow();
    if (fila <= FILA_HEADER) throw new Error('Selecciona una fila de paquete, no el encabezado.');
    const datos = sheet.getRange(fila, 1, 1, 1).getValues()[0];
    const codigo = String(datos[0] || '').trim().toUpperCase();
    if (!codigo) throw new Error('La fila no tiene código en columna A');
    const url = WEBAPP_URL + '?codigo=' + codigo;
    const html = HtmlService.createHtmlOutput(`
      <style>
        body{font-family:Arial,sans-serif;padding:20px;background:#f5f7fa;text-align:center}
        .btn{display:inline-block;margin-top:16px;padding:12px 28px;background:#2AABB5;
             color:#fff;font-size:14px;font-weight:700;border-radius:8px;text-decoration:none}
        .url{font-size:11px;color:#2AABB5;word-break:break-all;margin-top:12px}
        p{font-size:13px;color:#555;margin-top:8px}
      </style>
      <div style="font-size:22px;margin-bottom:8px;">📄</div>
      <strong style="font-size:15px;color:#1a1a2e">Cotización ${codigo}</strong>
      <p>Haz clic para abrir en el navegador</p>
      <a class="btn" href="${url}" target="_blank">🔗 Ver cotización</a>
      <div class="url">${url}</div>
    `).setWidth(400).setHeight(220);
    SpreadsheetApp.getUi().showModalDialog(html, `Cotización ${codigo}`);
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

// ── ABRIR WEB APP ─────────────────────────────────────────────
function abrirWebApp() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body{font-family:Arial,sans-serif;padding:20px;background:#f5f7fa;text-align:center}
      .btn{display:inline-block;margin-top:16px;padding:12px 28px;background:#2AABB5;
           color:#fff;font-size:14px;font-weight:700;border-radius:8px;text-decoration:none}
    </style>
    <div style="font-size:22px;margin-bottom:8px;">🌎</div>
    <strong style="font-size:15px;color:#1a1a2e">MaKing Trips — Web App</strong>
    <br><br>
    <a class="btn" href="${WEBAPP_URL}" target="_blank">🔗 Abrir Web App</a>
  `).setWidth(320).setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(html, 'Web App MaKing Trips');
}

// ── ALIAS para botones del sheet ──────────────────────────────
function enviarWhatsApp()   { enviarPorWhatsApp(); }
function getMensajeWA(fila) { return ''; }

// ── obtenerFilaActiva — para uso de otros .gs ─────────────────
function obtenerFilaActiva() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PAQUETES);
  if (!sheet) throw new Error('No se encontró la hoja PAQUETES_V3');
  const fila = sheet.getActiveCell().getRow();
  if (fila <= FILA_HEADER) throw new Error('Selecciona una fila de paquete, no el encabezado.');
  return sheet.getRange(fila, 1, 1, sheet.getLastColumn()).getValues()[0];
}

// ── Objeto WA — alias para compatibilidad con otros .gs ───────
const WA = {
  CODIGO:        COL_CODIGO        - 1,
  DESTINO:       COL_DESTINO       - 1,
  NOCHES:        COL_NOCHES        - 1,
  DIAS:          COL_DIAS          - 1,
  PERSONAS:      COL_PERSONAS      - 1,
  PRECIO_TRANSF: COL_PRECIO_TRANSF - 1,
  PRECIO_5PCT:   COL_PRECIO_TRANSF - 1,
  TARJETA_1:     COL_TARJETA_1     - 1,
  CUOTA_12:      COL_CUOTA_12      - 1,
  FECHA_VALIDEZ: COL_VALIDEZ       - 1,
  LINK_HOTEL:    COL_LINK_HOTEL    - 1,
  LINKS_ACT:     68,
  LINKS_TOURS:   69,
  LINKS_HOT_EX:  70,
};