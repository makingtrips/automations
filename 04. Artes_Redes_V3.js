// ============================================================
//  04_ArtesRedes_v3.gs — MaKing Trips
//  Generador de Artes para Redes Sociales
//  Fusión backup original + adaptación a PAQUETES_V3 catálogo
//  Autónomo — no depende de funciones de otros archivos
// ============================================================

// ── CONSTANTES ────────────────────────────────────────────────
const SHEET_PAQUETES_ARTES      = 'PAQUETES_V3';
const _FILA_HEADER        = 1;
const _TOTAL_COLS         = 72;
const _LOGO_FIGURA_ID     = '1w2Cj5cJ7Aux33aRhvCx-2aw4VQcIMcbr';
const CARPETA_ARTES       = 'Artes MaKing Trips';
const EMAIL_AGENCIA_ARTES = 'makingtripsgt@gmail.com';

// ── ÍNDICES 0-based — PAQUETES_V3 (catálogo A→BT) ─────────────
const COL_A = {
  CODIGO:        0,   // A
  DESTINO:       2,   // C
  TIPO_PAQUETE:  3,   // D
  PERSONAS:      7,   // H
  FECHA_SAL:     8,   // I
  FECHA_REG:     9,   // J
  NOCHES:       10,   // K
  DIAS:         11,   // L
  INC_VUELOS:   12,   // M
  TIPO_VIAJE:   13,   // N
  AEROLINEA:    14,   // O
  INC_HOTEL:    18,   // S
  HOTEL_REC:    19,   // T
  TIPO_HOTEL:   20,   // U
  INC_TRASL:    22,   // W
  TIPO_TRASL:   24,   // Y
  INC_ACT:      26,   // AA
  DET_ACT:      28,   // AC
  INC_TOURS:    30,   // AE
  DET_TOURS:    32,   // AG
  INC_SEGURO:   34,   // AI
  DET_SEGURO:   36,   // AK
  PRECIO_PP:    51,   // AZ
  TARJETA_1:    52,   // BA
  CUOTA_12:     60,   // BI
};

// ── siIncluye ─────────────────────────────────────────────────
function siIncluye(val) {
  var v = String(val||'').toLowerCase().trim()
    .replace(/í/g,'i');
  return v === 'si' || v === 'yes' || v === '1' || v === 'true';
}

// ── convertirGithubRawACdn ────────────────────────────────────
function convertirGithubRawACdn(url) {
  if (!url) return '';
  return String(url).trim().replace(
    'https://raw.githubusercontent.com/makingtrips/assets/main/',
    'https://cdn.jsdelivr.net/gh/makingtrips/assets@main/'
  );
}

// ── fotoBase64 — soporta Drive FileID y URLs (GitHub/CDN) ─────
function artesFotoBase64(fileIdOrUrl) {
  if (!fileIdOrUrl) return '';
  try {
    var src = convertirGithubRawACdn(String(fileIdOrUrl).trim());
    if (src.indexOf('http') === 0) {
      var response = UrlFetchApp.fetch(src);
      var bytes    = response.getContent();
      var mime     = (response.getHeaders()['Content-Type'] || 'image/jpeg').split(';')[0].trim();
      return 'data:' + mime + ';base64,' + Utilities.base64Encode(bytes);
    }
    var file  = DriveApp.getFileById(src);
    var bytes = file.getBlob().getBytes();
    return 'data:' + file.getMimeType() + ';base64,' + Utilities.base64Encode(bytes);
  } catch(e) { return fileIdOrUrl; }
}

// ── logoBase64 ────────────────────────────────────────────────
function artesLogoBase64(fileId) {
  if (!fileId) return '';
  try {
    var file  = DriveApp.getFileById(fileId);
    var bytes = file.getBlob().getBytes();
    return 'data:' + file.getMimeType() + ';base64,' + Utilities.base64Encode(bytes);
  } catch(e) { return ''; }
}

// ── leerFotoDestinoSheet ──────────────────────────────────────
function artesLeerFotoDestino(destino) {
  try {
    var key = String(destino).toLowerCase()
      .replace(/á/g,'a').replace(/é/g,'e').replace(/í/g,'i')
      .replace(/ó/g,'o').replace(/ú/g,'u').replace(/ñ/g,'n')
      .replace(/[^a-z0-9]/g,'').trim();
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('FOTOS_DESTINOS');
    if (!sheet) return '';
    var data  = sheet.getDataRange().getValues();
    var mejorLen = 0, mejorUrl = '';
    for (var i = 1; i < data.length; i++) {
      var k   = String(data[i][0]||'').toLowerCase()
        .replace(/á/g,'a').replace(/é/g,'e').replace(/í/g,'i')
        .replace(/ó/g,'o').replace(/ú/g,'u').replace(/ñ/g,'n')
        .replace(/[^a-z0-9]/g,'').trim();
      var url = String(data[i][1]||'').trim();
      if (!k || !url) continue;
      if (key === k) return url;
      if ((key.indexOf(k) !== -1 || k.indexOf(key) !== -1) && k.length > mejorLen) {
        mejorLen = k.length; mejorUrl = url;
      }
    }
    return mejorUrl;
  } catch(e) { return ''; }
}

// ── leerDatosPaquete — usa COL_A (índices reales) ─────────────
function artesLeerDatos(datos) {
  function n(v) { return Number(String(v||'').replace(/[^0-9.-]/g,'')) || 0; }
  return {
    codigo      : String(datos[COL_A.CODIGO]       || ''),
    destino     : String(datos[COL_A.DESTINO]      || ''),
    tipoPaquete : String(datos[COL_A.TIPO_PAQUETE] || ''),
    personas    : Number(datos[COL_A.PERSONAS]     || 2),
    fechaSal    : datos[COL_A.FECHA_SAL],
    fechaReg    : datos[COL_A.FECHA_REG],
    noches      : Number(datos[COL_A.NOCHES]       || 0),
    dias        : Number(datos[COL_A.DIAS]         || 0),
    aerolinea   : String(datos[COL_A.AEROLINEA]    || ''),
    tipoViaje   : String(datos[COL_A.TIPO_VIAJE]   || 'Redondo'),
    hotel       : String(datos[COL_A.HOTEL_REC]    || ''),
    tipoHotel   : String(datos[COL_A.TIPO_HOTEL]   || ''),
    tipoTraslado: String(datos[COL_A.TIPO_TRASL]   || ''),
    inclVuelos  : String(datos[COL_A.INC_VUELOS]   || ''),
    inclHotel   : String(datos[COL_A.INC_HOTEL]    || ''),
    inclTraslado: String(datos[COL_A.INC_TRASL]    || ''),
    inclAct     : String(datos[COL_A.INC_ACT]      || ''),
    inclTours   : String(datos[COL_A.INC_TOURS]    || ''),
    inclSeguro  : String(datos[COL_A.INC_SEGURO]   || ''),
    detAct      : String(datos[COL_A.DET_ACT]      || ''),
    detTours    : String(datos[COL_A.DET_TOURS]    || ''),
    detSeguro   : String(datos[COL_A.DET_SEGURO]   || ''),
    precioPP    : n(datos[COL_A.PRECIO_PP]),
    precioTot   : n(datos[COL_A.TARJETA_1]),
    tarjeta1    : n(datos[COL_A.TARJETA_1]),
    cuota12     : n(datos[COL_A.CUOTA_12]),
  };
}

// ── FUNCIÓN PRINCIPAL ─────────────────────────────────────────
function generarArtes() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PAQUETES);
  var fila  = sheet.getActiveCell().getRow();
  var ui    = SpreadsheetApp.getUi();

  if (fila <= _FILA_HEADER) {
    ui.alert('Selecciona una fila de paquete, no el encabezado.');
    return;
  }

  var datos = sheet.getRange(fila, 1, 1, _TOTAL_COLS).getValues()[0];
  var d     = artesLeerDatos(datos);

  if (!d.codigo || !d.destino) {
    ui.alert('La fila no tiene datos suficientes.');
    return;
  }

  var logoB64 = artesLogoBase64(_LOGO_FIGURA_ID);

  // Obtener foto — igual que el backup original
  var fotoUrl = artesLeerFotoDestino(d.destino);
  var fotoB64 = artesFotoBase64(fotoUrl);
  if (!fotoB64) fotoB64 = fotoUrl || '';  // fallback a URL directa

  var incluidos = [];
  if (siIncluye(d.inclVuelos))   incluidos.push('Vuelos incluidos');
  if (siIncluye(d.inclHotel))    incluidos.push(d.hotel ? d.hotel.split(' ').slice(0,2).join(' ') : 'Hotel incluido');
  if (siIncluye(d.inclTraslado)) incluidos.push('Traslados ' + (d.tipoTraslado || 'incluidos').toLowerCase());
  if (siIncluye(d.inclAct))      incluidos.push(d.detAct ? d.detAct.split('·')[0].trim() : 'Actividades');
  if (siIncluye(d.inclTours))    incluidos.push(d.detTours ? d.detTours.split('·')[0].trim() : 'Tours incluidos');
  if (siIncluye(d.inclSeguro))   incluidos.push('Seguro de viaje');

  var badgeTI = (d.tipoPaquete && d.tipoPaquete.trim()) ? d.tipoPaquete.toUpperCase() :
                (d.tipoHotel && d.tipoHotel.toLowerCase().indexOf('incluido') !== -1) ? d.tipoHotel.toUpperCase() : '';
  var precioStr  = 'Q' + Math.round(d.precioPP).toLocaleString('es-GT');
  var MESES_ES   = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  var fechaViaje = '';
  if (d.fechaSal instanceof Date && d.fechaReg instanceof Date) {
    fechaViaje = d.fechaSal.getDate() + ' ' + MESES_ES[d.fechaSal.getMonth()]
      + ' al ' + d.fechaReg.getDate() + ' ' + MESES_ES[d.fechaReg.getMonth()] + ' ' + d.fechaReg.getFullYear();
  }

  var carpetas   = DriveApp.getFoldersByName(CARPETA_ARTES);
  var carpeta    = carpetas.hasNext() ? carpetas.next() : DriveApp.createFolder(CARPETA_ARTES);
  var nombreBase = d.codigo + '_' + d.destino.replace(/[^a-zA-Z0-9]/g,'');

  var archivoPost    = guardarArte(carpeta, generarArtePost(d, logoB64, fotoB64, incluidos, precioStr, fechaViaje, badgeTI),             nombreBase + '_Post_1080x1080.html');
  var archivoHist    = guardarArte(carpeta, generarArteHistoria(d, logoB64, fotoB64, incluidos, precioStr, fechaViaje, badgeTI),          nombreBase + '_Historia_1080x1920.html');
  var archivoMod     = guardarArte(carpeta, generarArteModerno(d, logoB64, fotoB64, incluidos, precioStr, fechaViaje, badgeTI),           nombreBase + '_Post_Moderno_1080x1080.html');
  var archivoHistMod = guardarArte(carpeta, generarArteHistoriaModerna(d, logoB64, fotoB64, incluidos, precioStr, fechaViaje, badgeTI),   nombreBase + '_Historia_Moderna_1080x1920.html');

  ui.showModalDialog(
    HtmlService.createHtmlOutput(modalArtes(
      d.destino, d.codigo,
      'https://drive.google.com/uc?export=download&id=' + archivoPost.getId(),
      'https://drive.google.com/uc?export=download&id=' + archivoHist.getId(),
      'https://drive.google.com/uc?export=download&id=' + archivoMod.getId(),
      'https://drive.google.com/uc?export=download&id=' + archivoHistMod.getId()
    )).setWidth(500).setHeight(520),
    'Artes generados — ' + d.destino
  );
}

// ── GUARDAR ARTE EN DRIVE ─────────────────────────────────────
function guardarArte(carpeta, htmlContent, nombre) {
  var existentes = carpeta.getFilesByName(nombre);
  if (existentes.hasNext()) existentes.next().setTrashed(true);
  var archivo = carpeta.createFile(Utilities.newBlob(htmlContent, 'text/html', nombre));
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return archivo;
}

// ── ICONO SVG POR TIPO ────────────────────────────────────────
function getIconoArte(txt, size) {
  var s = size || 28;
  var t = txt.toLowerCase();
  if (t.indexOf('vuelo') !== -1 || t.indexOf('aereo') !== -1)
    return '<svg width="'+s+'" height="'+s+'" viewBox="0 0 24 24" fill="white"><path d="M21 16v-2l-8-5V3.5a1.5 1.5 0 0 0-3 0V9l-8 5v2l8-2.5V19l-2 1.5V22l3.5-1 3.5 1v-1.5L13 19v-5.5l8 2.5z"/></svg>';
  if (t.indexOf('hotel') !== -1 || t.indexOf('hosped') !== -1 || t.indexOf('resort') !== -1 || t.indexOf('tower') !== -1 || t.indexOf('beach') !== -1 || t.indexOf('fives') !== -1 || t.indexOf('riu') !== -1 || t.indexOf('dreams') !== -1 || t.indexOf('occidental') !== -1 || t.indexOf('americas') !== -1)
    return '<svg width="'+s+'" height="'+s+'" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><line x1="7" y1="4" x2="7" y2="20"/><line x1="17" y1="4" x2="17" y2="20"/><line x1="7" y1="12" x2="17" y2="12"/></svg>';
  if (t.indexOf('traslado') !== -1 || t.indexOf('transfer') !== -1 || t.indexOf('renta') !== -1)
    return '<svg width="'+s+'" height="'+s+'" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M2 17V9a2 2 0 0 1 2-2h11l4 4v6H2z"/><path d="M2 13h17"/><circle cx="6.5" cy="17.5" r="1.5"/><circle cx="16.5" cy="17.5" r="1.5"/></svg>';
  if (t.indexOf('tour') !== -1 || t.indexOf('excursion') !== -1 || t.indexOf('canal') !== -1)
    return '<svg width="'+s+'" height="'+s+'" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>';
  if (t.indexOf('seguro') !== -1)
    return '<svg width="'+s+'" height="'+s+'" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><polyline points="9 12 11 14 15 10"/></svg>';
  return '<svg width="'+s+'" height="'+s+'" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg>';
}

// ── ARTE POST INSTAGRAM 1080x1080 ─────────────────────────────
function generarArtePost(d, logoB64, fotoB64, incluidos, precioStr, fechaViaje, badgeTI) {
  var logoTag = logoB64
    ? '<img src="' + logoB64 + '" style="height:140px;object-fit:contain;filter:drop-shadow(0 2px 12px rgba(0,0,0,0.5));" alt="MaKing Trips">'
    : '<span style="font-size:20px;font-weight:700;color:#fff;">MaKing Trips</span>';
  var tagsHtml = incluidos.map(function(inc) {
    return '<div style="display:inline-block;background:rgba(42,171,181,0.35);border:1px solid rgba(42,171,181,0.7);color:#fff;font-size:18px;padding:6px 16px;border-radius:30px;margin:4px;">' + inc + '</div>';
  }).join('');
  return '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
  '<style>*{margin:0;padding:0;box-sizing:border-box}body{width:1080px;height:1080px;overflow:hidden;font-family:Arial,Helvetica,sans-serif}</style></head><body>' +
  '<div style="position:relative;width:1080px;height:1080px;overflow:hidden;">' +
    '<img src="' + fotoB64 + '" style="position:absolute;inset:0;width:100%;height:100%;object-fit:cover;">' +
    '<div style="position:absolute;inset:0;background:linear-gradient(180deg,rgba(0,0,0,0.05) 0%,rgba(0,0,0,0.08) 30%,rgba(0,0,0,0.60) 65%,rgba(0,0,0,0.82) 100%);"></div>' +
    '<div style="position:absolute;inset:0;display:flex;flex-direction:column;padding:60px 130px 55px 130px;">' +
      '<div>' + logoTag + '</div>' +
      '<div style="flex:1;"></div>' +
      '<div style="margin-bottom:28px;">' +
        '<div style="display:inline-block;background:rgba(42,171,181,0.4);border:1px solid rgba(42,171,181,0.8);color:#fff;font-size:16px;letter-spacing:.18em;text-transform:uppercase;padding:6px 20px;border-radius:30px;margin-bottom:16px;">Oferta especial</div>' +
        '<div style="font-size:80px;font-weight:700;color:#fff;line-height:1;letter-spacing:-2px;text-shadow:0 4px 20px rgba(0,0,0,0.5);margin-bottom:12px;">' + d.destino + '</div>' +
        (badgeTI ? '<div style="display:inline-block;background:linear-gradient(135deg,#C9922E,#f0c060);color:#fff;font-size:20px;font-weight:700;letter-spacing:.12em;padding:6px 20px;border-radius:8px;margin-bottom:16px;">&#11088; ' + badgeTI + '</div><br>' : '') +
        '<div style="margin-bottom:10px;"><div style="display:inline-block;background:rgba(42,171,181,0.5);border:1px solid rgba(42,171,181,0.8);color:#fff;font-size:19px;font-weight:600;padding:8px 20px;border-radius:30px;">' + d.dias + ' días &nbsp;·&nbsp; ' + d.noches + ' noches</div></div>' +
        (fechaViaje ? '<div style="background:rgba(0,0,0,0.35);border:1px solid rgba(255,255,255,0.2);color:rgba(255,255,255,0.9);font-size:18px;padding:7px 18px;border-radius:30px;display:inline-block;margin-bottom:20px;">' + fechaViaje + '</div>' : '<div style="margin-bottom:20px;"></div>') +
        '<div style="display:flex;align-items:stretch;gap:16px;margin-bottom:22px;flex-wrap:wrap;">' +
          '<div style="background:linear-gradient(135deg,#C9922E,#f0c060);border-radius:16px;padding:18px 32px;">' +
            '<div style="font-size:14px;letter-spacing:.2em;text-transform:uppercase;color:rgba(255,255,255,.9);margin-bottom:3px;">Desde</div>' +
            '<div style="font-size:68px;font-weight:700;color:#fff;line-height:1;">' + precioStr + '</div>' +
            '<div style="font-size:16px;color:rgba(255,255,255,.9);margin-top:3px;">por persona &nbsp;·&nbsp; ' + d.personas + ' personas</div>' +
            '<div style="font-size:14px;color:rgba(255,255,255,0.7);margin-top:5px;border-top:1px solid rgba(255,255,255,0.3);padding-top:6px;">Total paquete: <strong>Q' + Math.round(d.precioTot).toLocaleString('es-GT') + '</strong></div>' +
          '</div>' +
          '<div style="background:rgba(255,255,255,0.12);border:1px solid rgba(255,255,255,0.25);border-radius:16px;padding:18px 20px;text-align:center;display:flex;flex-direction:column;justify-content:center;">' +
            '<div style="font-size:13px;letter-spacing:.1em;text-transform:uppercase;color:rgba(255,255,255,.7);margin-bottom:4px;">O en</div>' +
            '<div style="font-size:36px;font-weight:700;color:#f0c060;line-height:1;">12 cuotas</div>' +
            '<div style="font-size:30px;font-weight:700;color:#fff;margin-top:4px;">Q' + Math.ceil(d.cuota12).toLocaleString('es-GT') + '</div>' +
            '<div style="font-size:14px;color:rgba(255,255,255,.65);margin-top:3px;">/mes</div>' +
          '</div>' +
        '</div>' +
        '<div style="margin-bottom:16px;padding:0 40px;">' + tagsHtml + '</div>' +
        '<div style="font-size:14px;color:rgba(255,255,255,0.55);line-height:1.5;">*Precios sujetos a disponibilidad al momento de la compra. Tarifas pueden variar sin previo aviso.</div>' +
      '</div>' +
    '</div>' +
  '</div></body></html>';
}

// ── ARTE HISTORIA 1080x1920 ───────────────────────────────────
function generarArteHistoria(d, logoB64, fotoB64, incluidos, precioStr, fechaViaje, badgeTI) {
  var logoTag = logoB64
    ? '<img src="' + logoB64 + '" style="height:130px;object-fit:contain;filter:drop-shadow(0 2px 12px rgba(0,0,0,0.4));" alt="MaKing Trips">'
    : '<span style="font-size:22px;font-weight:700;color:#fff;">MaKing Trips</span>';
  var checksHtml = incluidos.map(function(inc) {
    return '<div style="display:flex;align-items:center;gap:14px;background:rgba(255,255,255,0.1);border:1px solid rgba(255,255,255,0.2);border-radius:12px;padding:14px 20px;margin-bottom:10px;">' +
      '<div style="width:28px;height:28px;min-width:28px;background:linear-gradient(135deg,#2AABB5,#30C4E9);border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:16px;font-weight:700;color:#fff;">&#10003;</div>' +
      '<span style="font-size:26px;color:#fff;font-weight:500;">' + inc + '</span>' +
    '</div>';
  }).join('');
  return '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
  '<style>*{margin:0;padding:0;box-sizing:border-box}body{width:1080px;height:1920px;overflow:hidden;font-family:Arial,Helvetica,sans-serif}</style></head><body>' +
  '<div style="position:relative;width:1080px;height:1920px;overflow:hidden;">' +
    '<img src="' + fotoB64 + '" style="position:absolute;inset:0;width:100%;height:100%;object-fit:cover;">' +
    '<div style="position:absolute;inset:0;background:linear-gradient(180deg,rgba(13,74,85,0.45) 0%,rgba(0,0,0,0.1) 30%,rgba(0,0,0,0.72) 62%,rgba(10,30,45,0.9) 100%);"></div>' +
    '<div style="position:absolute;inset:0;display:flex;flex-direction:column;padding:270px 70px 360px 70px;">' +
      '<div style="text-align:center;margin-bottom:auto;">' + logoTag + '</div>' +
      '<div style="text-align:center;margin:auto 0;padding:40px 0;">' +
        '<div style="display:inline-block;background:rgba(255,255,255,0.15);border:1px solid rgba(255,255,255,0.35);color:#fff;font-size:18px;letter-spacing:.18em;text-transform:uppercase;padding:8px 24px;border-radius:30px;margin-bottom:24px;">Cotización especial</div>' +
        '<div style="font-size:88px;font-weight:700;color:#fff;line-height:1;letter-spacing:-2px;text-shadow:0 4px 24px rgba(0,0,0,0.5);margin-bottom:8px;">' + d.destino + '</div>' +
        (badgeTI ? '<div style="display:inline-block;background:linear-gradient(135deg,#C9922E,#f0c060);color:#fff;font-size:22px;font-weight:700;letter-spacing:.12em;padding:7px 22px;border-radius:8px;margin-bottom:10px;">&#11088; ' + badgeTI + '</div>' : '') +
        '<div style="font-size:26px;color:rgba(255,255,255,0.85);margin-bottom:28px;">' + d.dias + ' días &nbsp;·&nbsp; ' + d.noches + ' noches' + (fechaViaje ? ' &nbsp;·&nbsp; ' + fechaViaje : '') + '</div>' +
        '<div style="display:inline-block;background:linear-gradient(135deg,#C9922E,#f0c060);border-radius:20px;padding:24px 56px;">' +
          '<div style="font-size:17px;letter-spacing:.2em;text-transform:uppercase;color:rgba(255,255,255,.9);margin-bottom:6px;">Desde</div>' +
          '<div style="font-size:86px;font-weight:700;color:#fff;line-height:1;">' + precioStr + '</div>' +
          '<div style="font-size:20px;color:rgba(255,255,255,.9);margin-top:6px;">por persona &nbsp;·&nbsp; ocupación ' + (d.personas === 2 ? 'doble' : d.personas === 3 ? 'triple' : d.personas === 4 ? 'cuádruple' : d.personas + ' personas') + '</div>' +
        '</div>' +
      '</div>' +
      '<div style="margin-bottom:40px;">' + checksHtml + '</div>' +
      '<div style="text-align:center;">' +
        '<div style="display:inline-block;background:linear-gradient(135deg,#2AABB5,#30C4E9);color:#fff;font-size:28px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;padding:22px 64px;border-radius:60px;">Escríbenos y reserva tu lugar</div>' +
        '<div style="font-size:20px;color:rgba(255,255,255,0.6);margin-top:16px;">' + EMAIL_AGENCIA_ARTES + '</div>' +
      '</div>' +
    '</div>' +
  '</div></body></html>';
}

// ── ARTE MODERNO 1080x1080 ────────────────────────────────────
function generarArteModerno(d, logoB64, fotoB64, incluidos, precioStr, fechaViaje, badgeTI) {
  var logoTag = logoB64
    ? '<img src="' + logoB64 + '" style="height:150px;object-fit:contain;filter:drop-shadow(0 2px 16px rgba(0,0,0,0.6));" alt="MaKing Trips">'
    : '<span style="font-size:18px;font-weight:700;color:#fff;">MaKing Trips</span>';
  var cardW = Math.floor((740 - (incluidos.length - 1) * 16) / Math.max(incluidos.length, 1));
  var iconosHtml = incluidos.map(function(inc) {
    return '<div style="width:' + cardW + 'px;display:flex;flex-direction:column;align-items:center;gap:12px;background:rgba(0,0,0,0.35);border:1px solid rgba(255,255,255,0.25);border-radius:20px;padding:22px 12px;flex-shrink:0;">' +
      '<div style="width:64px;height:64px;background:linear-gradient(135deg,#2AABB5,#30C4E9);border-radius:50%;display:flex;align-items:center;justify-content:center;">' + getIconoArte(inc, 28) + '</div>' +
      '<span style="font-size:18px;color:#fff;font-weight:600;text-align:center;line-height:1.3;">' + inc + '</span>' +
    '</div>';
  }).join('');
  return '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
  '<style>*{margin:0;padding:0;box-sizing:border-box}body{width:1080px;height:1080px;overflow:hidden;font-family:Arial,Helvetica,sans-serif}</style></head><body>' +
  '<div style="position:relative;width:1080px;height:1080px;overflow:hidden;">' +
    '<img src="' + fotoB64 + '" style="position:absolute;inset:0;width:100%;height:100%;object-fit:cover;">' +
    '<div style="position:absolute;inset:0;background:linear-gradient(180deg,rgba(0,0,0,0.78) 0%,rgba(0,0,0,0.2) 30%,rgba(0,0,0,0.25) 60%,rgba(0,0,0,0.88) 100%);"></div>' +
    '<div style="position:absolute;inset:0;display:flex;flex-direction:column;padding:60px 130px 55px 130px;">' +
      '<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:20px;">' +
        '<div style="flex:1;">' +
          '<div style="font-size:82px;font-weight:700;color:#fff;line-height:1;letter-spacing:-2px;text-shadow:0 4px 24px rgba(0,0,0,0.7);">' + d.destino + '</div>' +
          (badgeTI ? '<div style="display:inline-block;background:linear-gradient(135deg,#C9922E,#f0c060);color:#fff;font-size:20px;font-weight:700;letter-spacing:.12em;padding:6px 20px;border-radius:8px;margin-top:10px;">&#11088; ' + badgeTI + '</div>' : '') +
          (fechaViaje ? '<div style="font-size:20px;color:rgba(255,255,255,0.85);margin-top:10px;margin-bottom:10px;">' + fechaViaje + '</div>' : '<div style="margin-bottom:10px;"></div>') +
          '<div style="display:inline-block;background:rgba(42,171,181,0.5);border:1px solid rgba(42,171,181,0.8);color:#fff;font-size:18px;font-weight:600;padding:7px 20px;border-radius:30px;">' + d.dias + ' días &nbsp;·&nbsp; ' + d.noches + ' noches</div>' +
        '</div>' +
        '<div style="flex-shrink:0;padding-left:16px;">' + logoTag + '</div>' +
      '</div>' +
      '<div style="display:flex;gap:16px;flex-wrap:nowrap;margin:auto 0;padding:24px 40px;">' + iconosHtml + '</div>' +
      '<div style="display:flex;justify-content:space-between;align-items:flex-end;padding-bottom:4px;">' +
        '<div style="max-width:480px;padding-bottom:8px;"><div style="font-size:14px;color:rgba(255,255,255,0.5);line-height:1.6;">*Precios sujetos a disponibilidad. Tarifas pueden variar sin previo aviso.</div></div>' +
        '<div style="text-align:right;padding-bottom:8px;">' +
          '<div style="font-size:19px;color:rgba(255,255,255,0.8);letter-spacing:.05em;margin-bottom:4px;">Desde</div>' +
          '<div style="font-size:82px;font-weight:700;color:#fff;line-height:1;text-shadow:0 4px 20px rgba(0,0,0,0.5);">' + precioStr + '</div>' +
          '<div style="font-size:18px;color:rgba(255,255,255,0.8);margin-top:5px;">por persona &nbsp;·&nbsp; ' + d.personas + ' personas</div>' +
          '<div style="font-size:16px;color:rgba(255,255,255,0.6);margin-top:3px;">Total: <strong style="color:rgba(255,255,255,0.85);">Q' + Math.round(d.precioTot).toLocaleString('es-GT') + '</strong></div>' +
          '<div style="font-size:16px;color:#f0c060;margin-top:6px;font-weight:700;">O en 12 cuotas de Q' + Math.ceil(d.cuota12).toLocaleString('es-GT') + '/mes</div>' +
        '</div>' +
      '</div>' +
    '</div>' +
  '</div></body></html>';
}

// ── ARTE HISTORIA MODERNA 1080x1920 ───────────────────────────
function generarArteHistoriaModerna(d, logoB64, fotoB64, incluidos, precioStr, fechaViaje, badgeTI) {
  var logoTag = logoB64
    ? '<img src="' + logoB64 + '" style="height:130px;object-fit:contain;filter:drop-shadow(0 2px 16px rgba(0,0,0,0.6));" alt="MaKing Trips">'
    : '<span style="font-size:20px;font-weight:700;color:#fff;">MaKing Trips</span>';
  var iconosHtml = incluidos.map(function(inc) {
    return '<div style="display:flex;align-items:center;gap:20px;background:rgba(0,0,0,0.35);border:1px solid rgba(255,255,255,0.2);border-radius:18px;padding:18px 24px;margin-bottom:12px;">' +
      '<div style="width:64px;height:64px;min-width:64px;background:linear-gradient(135deg,#2AABB5,#30C4E9);border-radius:50%;display:flex;align-items:center;justify-content:center;">' + getIconoArte(inc, 32) + '</div>' +
      '<span style="font-size:28px;color:#fff;font-weight:600;">' + inc + '</span>' +
    '</div>';
  }).join('');
  return '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
  '<style>*{margin:0;padding:0;box-sizing:border-box}body{width:1080px;height:1920px;overflow:hidden;font-family:Arial,Helvetica,sans-serif}</style></head><body>' +
  '<div style="position:relative;width:1080px;height:1920px;overflow:hidden;">' +
    '<img src="' + fotoB64 + '" style="position:absolute;inset:0;width:100%;height:100%;object-fit:cover;">' +
    '<div style="position:absolute;inset:0;background:linear-gradient(180deg,rgba(0,0,0,0.7) 0%,rgba(0,0,0,0.15) 25%,rgba(0,0,0,0.15) 55%,rgba(0,0,0,0.85) 100%);"></div>' +
    '<div style="position:absolute;inset:0;display:flex;flex-direction:column;padding:270px 70px 360px 70px;">' +
      '<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:30px;">' +
        '<div style="flex:1;padding-right:24px;">' +
          '<div style="font-size:92px;font-weight:700;color:#fff;line-height:1;letter-spacing:-2px;text-shadow:0 4px 28px rgba(0,0,0,0.7);">' + d.destino + '</div>' +
          (badgeTI ? '<div style="display:inline-block;background:linear-gradient(135deg,#C9922E,#f0c060);color:#fff;font-size:24px;font-weight:700;letter-spacing:.12em;padding:8px 24px;border-radius:8px;margin-top:10px;">&#11088; ' + badgeTI + '</div>' : '') +
          (fechaViaje ? '<div style="font-size:28px;color:rgba(255,255,255,0.85);margin-top:14px;margin-bottom:14px;">' + fechaViaje + '</div>' : '<div style="margin-bottom:14px;"></div>') +
          '<div style="display:inline-block;background:rgba(42,171,181,0.5);border:1px solid rgba(42,171,181,0.8);color:#fff;font-size:24px;font-weight:600;padding:10px 26px;border-radius:30px;">' + d.dias + ' días &nbsp;·&nbsp; ' + d.noches + ' noches</div>' +
        '</div>' +
        '<div style="flex-shrink:0;">' + logoTag + '</div>' +
      '</div>' +
      '<div style="margin:auto 0;padding:20px 0;">' + iconosHtml + '</div>' +
      '<div style="display:flex;justify-content:space-between;align-items:flex-end;padding-bottom:10px;">' +
        '<div style="max-width:480px;padding-bottom:10px;"><div style="font-size:20px;color:rgba(255,255,255,0.5);line-height:1.6;">*Precios sujetos a disponibilidad. Tarifas pueden variar sin previo aviso.</div></div>' +
        '<div style="text-align:right;padding-bottom:10px;">' +
          '<div style="font-size:24px;color:rgba(255,255,255,0.8);letter-spacing:.05em;margin-bottom:4px;">Desde</div>' +
          '<div style="font-size:96px;font-weight:700;color:#fff;line-height:1;text-shadow:0 4px 20px rgba(0,0,0,0.5);">' + precioStr + '</div>' +
          '<div style="font-size:22px;color:rgba(255,255,255,0.8);margin-top:6px;">por persona &nbsp;·&nbsp; ' + d.personas + ' personas</div>' +
          '<div style="font-size:20px;color:rgba(255,255,255,0.6);margin-top:4px;">Total: <strong style="color:rgba(255,255,255,0.85);">Q' + Math.round(d.precioTot).toLocaleString('es-GT') + '</strong></div>' +
          '<div style="font-size:22px;color:#f0c060;margin-top:8px;font-weight:700;">O en 12 cuotas de Q' + Math.ceil(d.cuota12).toLocaleString('es-GT') + '/mes</div>' +
        '</div>' +
      '</div>' +
    '</div>' +
  '</div></body></html>';
}

// ── MODAL ─────────────────────────────────────────────────────
function modalArtes(destino, codigo, urlPostDl, urlHistDl, urlModDl, urlHistModDl) {
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
  '*{box-sizing:border-box;margin:0;padding:0}body{font-family:Arial,sans-serif;font-size:13px;background:#f5f7fa;padding:20px;color:#111}' +
  '.hdr{background:linear-gradient(135deg,#2AABB5,#1d8fa0);color:#fff;border-radius:8px;padding:14px 18px;margin-bottom:14px;text-align:center}' +
  '.hdr h2{font-size:15px;font-weight:700}.card{background:#fff;border:1px solid #e0e0e0;border-radius:8px;padding:14px 16px;margin-bottom:10px}' +
  '.card-title{font-size:11px;font-weight:700;color:#2AABB5;letter-spacing:.08em;text-transform:uppercase;margin-bottom:6px}' +
  '.card-desc{font-size:12px;color:#666;line-height:1.5;margin-bottom:10px}' +
  '.btn{display:block;width:100%;padding:10px;font-size:12px;font-weight:700;text-align:center;border-radius:6px;cursor:pointer;border:none;margin-bottom:6px;text-decoration:none}' +
  '.btn-teal{background:#2AABB5;color:#fff}.btn-gold{background:linear-gradient(135deg,#C9922E,#f0c060);color:#fff}' +
  '.btn-mod{background:linear-gradient(135deg,#1a1a2e,#2d3561);color:#f0c060;border:1px solid #f0c060}' +
  '.btn-out{background:#fff;color:#999;border:1px solid #ddd;font-weight:400}' +
  '.steps{background:#f0fafb;border-radius:8px;padding:12px 14px;margin-bottom:10px;font-size:12px;color:#444;line-height:1.8}.steps strong{color:#2AABB5}' +
  '</style></head><body>' +
  '<div class="hdr"><div style="font-size:22px;margin-bottom:4px;">&#127912;</div><h2>Artes generados</h2><small>' + destino + ' &nbsp;&middot;&nbsp; ' + codigo + '</small></div>' +
  '<div class="steps"><strong>Cómo convertir a PNG:</strong><br>1. Abre el arte en Chrome<br>2. Instala la extensión <strong>GoFullPage</strong> (gratis)<br>3. Haz clic en el ícono de GoFullPage<br>4. Descarga el PNG — listo para publicar</div>' +
  '<div class="card"><div class="card-title">Post Instagram — 1080 x 1080</div><div class="card-desc">Formato cuadrado para feed de Instagram y Facebook.</div><a class="btn btn-teal" href="' + urlPostDl + '" target="_blank" download>Descargar Post HTML</a></div>' +
  '<div class="card"><div class="card-title">Historia — 1080 x 1920</div><div class="card-desc">Formato vertical para historias y estado de WhatsApp.</div><a class="btn btn-gold" href="' + urlHistDl + '" target="_blank" download>Descargar Historia HTML</a></div>' +
  '<div class="card"><div class="card-title">Post Estilo Moderno — 1080 x 1080</div><div class="card-desc">Ciudad arriba, iconos en el centro, precio abajo derecha.</div><a class="btn btn-mod" href="' + urlModDl + '" target="_blank" download>Descargar Post Moderno</a></div>' +
  '<div class="card"><div class="card-title">Historia Moderna — 1080 x 1920</div><div class="card-desc">Mismo estilo moderno en vertical para Stories.</div><a class="btn btn-mod" href="' + urlHistModDl + '" target="_blank" download>Descargar Historia Moderna</a></div>' +
  '<button class="btn btn-out" onclick="google.script.host.close()" style="margin-top:4px;">Cerrar</button>' +
  '</body></html>';
}