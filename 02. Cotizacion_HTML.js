// ============================================================
//  Cotizacion_HTML.gs — MaKing Trips
//  Diseño HTML premium basado en backup original v5
//  Índices verificados contra catálogo PAQUETES_V3 (72 cols A→BT)
// ============================================================

// ── LOGOS ─────────────────────────────────────────────────────
const LOGO_HORIZONTAL_ID = '1GOJPtjESKQJnRnNllLMktK3XeyD7G5KE';
const LOGO_CUADRADO_ID   = '1s-6O08GeA-nfY5LWSpDNVzYkoMHrsa53';

// ── CONTACTO ──────────────────────────────────────────────────
const BANCO_NOMBRE  = 'Banco Industrial';
const CTA_Q         = '0690225383';
const CTA_USD       = '0180135675';
const EMAIL_COT     = 'makingtripsgt@gmail.com';

// ── ÍNDICES 0-based — PAQUETES_V3 (catálogo A→BT) ─────────────
const COL = {
  CODIGO:        0,  // A
  DESTINO:       2,  // C
  TIPO_PAQUETE:  3,  // D
  PERSONAS:      7,  // H
  FECHA_SAL:     8,  // I
  FECHA_REG:     9,  // J
  NOCHES:       10,  // K
  DIAS:         11,  // L
  INC_VUELOS:   12,  // M
  TIPO_VIAJE:   13,  // N
  AEROLINEA:    14,  // O
  INC_HOTEL:    18,  // S
  HOTEL_REC:    19,  // T  Hotel Recomendado
  TIPO_HOTEL:   20,  // U
  INC_TRASL:    22,  // W
  TIPO_TRASL:   24,  // Y
  INC_ACT:      26,  // AA
  DET_ACT:      28,  // AC
  INC_TOURS:    30,  // AE
  DET_TOURS:    32,  // AG
  INC_SEGURO:   34,  // AI
  DET_SEGURO:   36,  // AK
  PRECIO_TOT:   49,  // AX
  PRECIO_TRANSF:50,  // AY ← Precio Transferencia 5%
  PRECIO_PP:    51,  // AZ
  TARJETA_1:    52,  // BA
  CUOTA_3:      54,  // BC
  CUOTA_6:      56,  // BE
  CUOTA_10:     58,  // BG
  CUOTA_12:     60,  // BI
  CUOTA_18:     62,  // BK
  FECHA_VALIDEZ:64,  // BM
  LINK_HOTEL:   67,  // BP  (existente)
  LINKS_ACT:    68,  // BQ  ← nueva
  LINKS_TOURS:  69,  // BR  ← nueva
  LINKS_HOT_EX: 70,  // BS  ← nueva
};

const MESES_HTML = ['enero','febrero','marzo','abril','mayo','junio',
                    'julio','agosto','septiembre','octubre','noviembre','diciembre'];
const MESES_CORTOS = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];

// ─────────────────────────────────────────────
//  Utilidades
// ─────────────────────────────────────────────
function htmlToNum(val) {
  if (typeof val === 'number') return val;
  return Number(String(val).replace(/[^0-9.-]/g,'')) || 0;
}
function htmlFmtQ(n) { return 'Q' + Math.round(Number(n)).toLocaleString('es-GT'); }
function htmlIncl(val) {
  var v = String(val||'').toLowerCase().trim().replace(/í/g,'i');
  return v === 'si' || v === 'sí' || v === 'yes' || v === '1' || v === 'true';
}
function htmlFmtFecha(valor) {
  if (!valor) return '';
  const d = new Date(valor);
  if (isNaN(d.getTime())) return String(valor);
  return `${d.getDate()} de ${MESES_HTML[d.getMonth()]} de ${d.getFullYear()}`;
}
function htmlFmtFechaCorta(valor) {
  if (!valor) return '';
  const d = new Date(valor);
  if (isNaN(d.getTime())) return String(valor);
  return `${d.getDate()} ${MESES_CORTOS[d.getMonth()]} ${d.getFullYear()}`;
}

// ─────────────────────────────────────────────
//  parseLinks — convierte BQ/BR/BS en array {nombre, url}
//  Formato: Nombre|https://url.com  (una por línea)
// ─────────────────────────────────────────────
function parseLinks(cellValue) {
  if (!cellValue || String(cellValue).trim() === '') return [];
  return String(cellValue).split('\n')
    .map(l => l.trim()).filter(l => l.includes('|'))
    .map(l => { const p = l.split('|'); return { nombre:(p[0]||'').trim(), url:(p[1]||'').trim() }; })
    .filter(i => i.nombre && i.url);
}

// ─────────────────────────────────────────────
//  renderLinkButtons — botones por categoría
// ─────────────────────────────────────────────
function renderLinkButtons(links, tipo) {
  if (!links || links.length === 0) return '';
  const colores = { actividad:'#2AABB5', tour:'#C9922E', hotel:'#1a1a2e' };
  const iconos  = { actividad:'🎯', tour:'🗺️', hotel:'🏨' };
  const bg = colores[tipo] || colores.actividad;
  const ic = iconos[tipo]  || iconos.actividad;
  return links.map(lk =>
    `<a href="${lk.url}" target="_blank" rel="noopener"
        style="display:inline-flex;align-items:center;gap:6px;
               background:${bg};color:#ffffff;padding:8px 16px;
               border-radius:8px;font-family:'Lato',Arial,sans-serif;
               font-size:12px;font-weight:700;text-decoration:none;margin:4px;">
       ${ic} ${lk.nombre}
     </a>`
  ).join('');
}

// ─────────────────────────────────────────────
//  obtenerFotoDestino — busca en FOTOS_DESTINOS
// ─────────────────────────────────────────────
function obtenerFotoDestino(destino) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FOTOS_DESTINOS');
    if (!sheet) return '';
    const datos = sheet.getDataRange().getValues();
    const dest  = String(destino).toLowerCase().trim()
      .replace(/á/g,'a').replace(/é/g,'e').replace(/í/g,'i')
      .replace(/ó/g,'o').replace(/ú/g,'u').replace(/ñ/g,'n');
    let mejorLen = 0, mejorUrl = '';
    for (let i = 1; i < datos.length; i++) {
      const k   = String(datos[i][0]||'').toLowerCase().trim()
        .replace(/á/g,'a').replace(/é/g,'e').replace(/í/g,'i')
        .replace(/ó/g,'o').replace(/ú/g,'u').replace(/ñ/g,'n');
      const url = String(datos[i][1]||'').trim();
      if (!k || !url) continue;
      if (dest === k) return url;
      if ((dest.indexOf(k) !== -1 || k.indexOf(dest) !== -1) && k.length > mejorLen) {
        mejorLen = k.length; mejorUrl = url;
      }
    }
    return mejorUrl;
  } catch(e) { return ''; }
}

// ─────────────────────────────────────────────
//  logoBase64 — para el adjunto de mail
// ─────────────────────────────────────────────
function logoBase64(fileId) {
  if (!fileId) return '';
  try {
    const file  = DriveApp.getFileById(fileId);
    const bytes = file.getBlob().getBytes();
    return 'data:' + file.getMimeType() + ';base64,' + Utilities.base64Encode(bytes);
  } catch(e) { return ''; }
}

// ─────────────────────────────────────────────
//  iconoSvg — íconos SVG blancos para las cards
// ─────────────────────────────────────────────
function iconoSvg(tipo) {
  const iconos = {
    vuelo:     '<svg width="16" height="16" viewBox="0 0 24 24" fill="white"><path d="M21 16v-2l-8-5V3.5a1.5 1.5 0 0 0-3 0V9l-8 5v2l8-2.5V19l-2 1.5V22l3.5-1 3.5 1v-1.5L13 19v-5.5z"/></svg>',
    hotel:     '<svg width="16" height="16" viewBox="0 0 24 24" fill="white"><path d="M3 21V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v16H3zm5-2h8v-5H8v5zm0-7h3V9H8v3zm5 0h3V9h-3v3z"/></svg>',
    traslado:  '<svg width="16" height="16" viewBox="0 0 24 24" fill="white"><path d="M23 11l-2-6H3L1 11v2h1v7h4v-2h10v2h4v-7h1v-2zM5.5 16a1.5 1.5 0 1 1 0-3 1.5 1.5 0 0 1 0 3zm13 0a1.5 1.5 0 1 1 0-3 1.5 1.5 0 0 1 0 3zM3 11l1.5-4.5h15L21 11H3z"/></svg>',
    actividad: '<svg width="16" height="16" viewBox="0 0 24 24" fill="white"><path d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01z"/></svg>',
    tours:     '<svg width="16" height="16" viewBox="0 0 24 24" fill="white"><path d="M12 2C8.13 2 5 5.13 5 9c0 5.25 7 13 7 13s7-7.75 7-13c0-3.87-3.13-7-7-7zm0 9.5c-1.38 0-2.5-1.12-2.5-2.5s1.12-2.5 2.5-2.5 2.5 1.12 2.5 2.5-1.12 2.5-2.5 2.5z"/></svg>',
    seguro:    '<svg width="16" height="16" viewBox="0 0 24 24" fill="white"><path d="M12 1L3 5v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V5l-9-4zm-2 16l-4-4 1.41-1.41L10 14.17l6.59-6.59L18 9l-8 8z"/></svg>',
  };
  return iconos[tipo] || iconos.actividad;
}

// ─────────────────────────────────────────────
//  secLabel — título de sección con línea dorada
// ─────────────────────────────────────────────
function secLabel(txt) {
  return '<div style="display:flex;align-items:center;gap:8px;margin-bottom:14px;">' +
    '<div style="width:3px;height:18px;background:#C9922E;border-radius:2px;flex-shrink:0;"></div>' +
    '<div style="font-size:10px;letter-spacing:.18em;text-transform:uppercase;color:#C9922E;font-weight:600;">' + txt + '</div>' +
    '</div>';
}

// ─────────────────────────────────────────────
//  generarCotizacionHTML — función principal
//  Recibe fila de PAQUETES_V3, devuelve HTML completo
// ─────────────────────────────────────────────
function generarCotizacionHTML(fila) {

  // ── Datos ──
  const codigo      = fila[COL.CODIGO]       || '';
  const destino     = fila[COL.DESTINO]      || '';
  const tipoPaquete = fila[COL.TIPO_PAQUETE] || '';
  const personas    = Number(fila[COL.PERSONAS]  || 2);
  const noches      = Number(fila[COL.NOCHES]    || 0);
  const dias        = Number(fila[COL.DIAS]      || 0);
  const aerolinea   = fila[COL.AEROLINEA]    || '';
  const tipoViaje   = fila[COL.TIPO_VIAJE]   || 'Redondo';
  const hotel       = fila[COL.HOTEL_REC]    || '';
  const tipoHotel   = fila[COL.TIPO_HOTEL]   || '';
  const tipoTrasl   = fila[COL.TIPO_TRASL]   || '';
  const detAct      = fila[COL.DET_ACT]      || '';
  const detTours    = fila[COL.DET_TOURS]    || '';
  const detSeguro   = fila[COL.DET_SEGURO]   || '';

  // ── Toggles ──
  const inclVuelos  = htmlIncl(fila[COL.INC_VUELOS]);
  const inclHotel   = htmlIncl(fila[COL.INC_HOTEL]);
  const inclTrasl   = htmlIncl(fila[COL.INC_TRASL]);
  const inclAct     = htmlIncl(fila[COL.INC_ACT]);
  const inclTours   = htmlIncl(fila[COL.INC_TOURS]);
  const inclSeguro  = htmlIncl(fila[COL.INC_SEGURO]);

  // ── Precios ──
  const precioTot   = htmlToNum(fila[COL.PRECIO_TOT]);
  const pTransf     = htmlToNum(fila[COL.PRECIO_TRANSF]);
  const precioPP    = htmlToNum(fila[COL.PRECIO_PP]);
  const tarjeta1    = htmlToNum(fila[COL.TARJETA_1]);
  const cuota3      = htmlToNum(fila[COL.CUOTA_3]);
  const cuota6      = htmlToNum(fila[COL.CUOTA_6]);
  const cuota10     = htmlToNum(fila[COL.CUOTA_10]);
  const cuota12     = htmlToNum(fila[COL.CUOTA_12]);
  const cuota18     = htmlToNum(fila[COL.CUOTA_18]);
  const pTarjeta    = tarjeta1 > 0 ? tarjeta1 : precioTot;
  const ahorro      = Math.round(pTarjeta - pTransf);

  // ── Links ──
  const linkHotel   = fila[COL.LINK_HOTEL]   ? String(fila[COL.LINK_HOTEL]).trim()   : '';
  const linksAct    = parseLinks(fila[COL.LINKS_ACT]);
  const linksTours  = parseLinks(fila[COL.LINKS_TOURS]);
  const linksHotEx  = parseLinks(fila[COL.LINKS_HOT_EX]);

  // ── Foto, logos, fechas ──
  const fotoSrc     = obtenerFotoDestino(destino);
  const logoHero    = logoBase64(LOGO_HORIZONTAL_ID);
  const logoFooter  = logoBase64(LOGO_CUADRADO_ID);
  const fechaSal    = htmlFmtFechaCorta(fila[COL.FECHA_SAL]);
  const fechaReg    = htmlFmtFechaCorta(fila[COL.FECHA_REG]);
  const fechaViaje  = (fechaSal && fechaReg) ? `${fechaSal} al ${fechaReg}` : '';
  const validezTxt  = htmlFmtFecha(fila[COL.FECHA_VALIDEZ]);
  const ocupTipo    = personas <= 1 ? 'individual' : personas === 2 ? 'doble' : personas === 3 ? 'triple' : personas === 4 ? 'cuádruple' : personas + ' personas';

  // ── Tags de logo e imagen ──
  const logoHeroTag   = logoHero
    ? `<img src="${logoHero}" alt="MaKing Trips" height="70" style="display:block;margin:0 auto 6px;object-fit:contain;">`
    : '<div style="font-size:22px;font-weight:700;color:#1a1a2e;margin-bottom:6px;">MaKing Trips</div>';
  const logoFooterTag = logoFooter
    ? `<img src="${logoFooter}" alt="MaKing Trips" height="32" style="display:block;margin:0 auto 8px;object-fit:contain;">` : '';
  const fotoTag = fotoSrc
    ? `<img src="${fotoSrc}" alt="${destino}" width="560" height="280" style="width:100%;height:280px;object-fit:cover;display:block;">`
    : '<div style="width:100%;height:280px;background:linear-gradient(135deg,#2AABB5,#1d8fa0);display:block;"></div>';

  // ── Sección horizontal premium de servicios incluidos ──
  const infoResumen = [];

  if (inclVuelos) {
    infoResumen.push({ icon:'vuelo', titulo:'Tipo de viaje', detalle:tipoViaje || 'Redondo' });
  }

  if (inclTrasl) {
    infoResumen.push({ icon:'traslado', titulo:'Traslados', detalle:tipoTrasl || 'Incluidos' });
  }

  if (inclSeguro) {
    infoResumen.push({ icon:'seguro', titulo:'Seguro de viaje', detalle:detSeguro || 'Incluido' });
  }

  function renderMiniInfo(items) {
    if (!items || items.length === 0) return '';
    return '<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:12px;"><tr>' +
      items.map(function(it) {
        return '<td style="padding:4px;vertical-align:top;">' +
          '<div style="background:#F7F5F0;border:1px solid #ece8e0;border-radius:12px;padding:12px;min-height:64px;">' +
            '<table cellpadding="0" cellspacing="0" width="100%"><tr>' +
              '<td style="width:34px;vertical-align:top;padding-right:9px;">' +
                '<div style="width:32px;height:32px;background:#2AABB5;border-radius:9px;text-align:center;line-height:32px;">' + iconoSvg(it.icon) + '</div>' +
              '</td>' +
              '<td style="vertical-align:top;">' +
                '<div style="font-size:11px;color:#1A3A52;font-weight:700;line-height:1.25;margin-bottom:3px;">' + it.titulo + '</div>' +
                '<div style="font-size:10.5px;color:#1d8fa0;font-weight:600;line-height:1.25;">' + it.detalle + '</div>' +
              '</td>' +
            '</tr></table>' +
          '</div>' +
        '</td>';
      }).join('') +
    '</tr></table>';
  }

  function renderPillButtons(links, tipo) {
    if (!links || links.length === 0) return '';
    const iconos = { hotel:'🏨', actividad:'🎯', tour:'🗺️' };
    const ic = iconos[tipo] || '🔗';
    return links.map(function(lk) {
      return '<a href="' + lk.url + '" target="_blank" rel="noopener" ' +
        'style="display:inline-block;background:#ffffff;border:1px solid #d8d2c8;border-radius:10px;' +
        'padding:10px 16px;margin:0 7px 8px 0;text-decoration:none;color:#1A3A52;' +
        'font-weight:700;font-size:13px;font-family:Lato,Arial,sans-serif;line-height:1.2;">' +
        '🔗 ' + lk.nombre + '</a>';
    }).join('');
  }

  function renderPremiumSection(icon, title, body, marginBottom) {
    if (!body) return '';
    return '<div style="background:#F7F5F0;border-radius:13px;padding:15px;margin-bottom:' + (marginBottom ? '12px' : '0') + ';border:1px solid rgba(201,146,46,0.12);">' +
      '<div style="font-size:14px;font-weight:700;color:#1A3A52;margin-bottom:11px;font-family:Lato,Arial,sans-serif;line-height:1.2;">' + icon + ' ' + title + '</div>' +
      body +
    '</div>';
  }

  let hotelBody = '';
  if (inclHotel) {
    if (linkHotel) {
      hotelBody += '<a href="' + linkHotel + '" target="_blank" rel="noopener" ' +
        'style="display:inline-block;background:#ffffff;border:1px solid #d8d2c8;border-radius:10px;' +
        'padding:10px 18px;margin:0 7px 8px 0;text-decoration:none;color:#1A3A52;' +
        'font-weight:700;font-size:13px;font-family:Lato,Arial,sans-serif;line-height:1.2;">🔗 Ver hotel</a>';
    }

    if (linksHotEx.length > 0) {
      hotelBody += renderPillButtons(linksHotEx, 'hotel');
    }

    if (!hotelBody) {
      hotelBody = '<div style="font-size:12px;color:#1d8fa0;font-weight:700;line-height:1.35;">' + (hotel || 'Hotel incluido') + '</div>';
    }

    if (hotel) {
      hotelBody += '<div style="font-size:11px;color:#68707a;margin-top:2px;line-height:1.35;">' + hotel + '</div>';
    }
    if (tipoHotel) {
      hotelBody += '<div style="font-size:10.5px;color:#9a8f80;margin-top:3px;line-height:1.35;">' + tipoHotel + '</div>';
    }
  }

  let actBody = '';
  if (inclAct) {
    if (linksAct.length > 0) {
      actBody += renderPillButtons(linksAct, 'actividad');
    }
    if (detAct && detAct.trim()) {
      actBody += '<div style="font-size:11px;color:#68707a;margin-top:2px;line-height:1.35;">' + detAct.trim() + '</div>';
    }
    if (!actBody) actBody = '<div style="font-size:12px;color:#1d8fa0;font-weight:700;">Actividades incluidas</div>';
  }

  let toursBody = '';
  if (inclTours) {
    if (linksTours.length > 0) {
      toursBody += renderPillButtons(linksTours, 'tour');
    }
    if (detTours && detTours.trim()) {
      toursBody += '<div style="font-size:11px;color:#68707a;margin-top:2px;line-height:1.35;">' + detTours.trim() + '</div>';
    }
    if (!toursBody) toursBody = '<div style="font-size:12px;color:#1d8fa0;font-weight:700;">Tours incluidos</div>';
  }

  let filasIncluidos =
  '<tr><td style="padding:0;">' +
    renderMiniInfo(infoResumen) +
    '<div style="background:linear-gradient(135deg,#2AABB5 0%, #2397a0 100%);border-radius:16px;padding:18px;border:1px solid rgba(255,255,255,0.18);">' +
      renderPremiumSection('🏨', 'Hotel', hotelBody, (actBody || toursBody)) +
      renderPremiumSection('🎯', 'Actividades incluidas', actBody, toursBody) +
      renderPremiumSection('🗺️', 'Tours incluidos', toursBody, false) +
      ((hotelBody || actBody || toursBody)
        ? '<div style="text-align:center;color:#ffffff;opacity:.72;font-size:10.5px;margin-top:14px;font-family:Lato,Arial,sans-serif;">Cada botón abre el detalle de la experiencia incluida</div>'
        : '') +
    '</div>' +
  '</td></tr>';

  // ── Cuotas ──
  const cuotas = [
    { n:'1 pago',    m:htmlFmtQ(tarjeta1), pop:false },
    { n:'3 cuotas',  m:htmlFmtQ(cuota3),   pop:false },
    { n:'6 cuotas',  m:htmlFmtQ(cuota6),   pop:false },
    { n:'12 cuotas', m:htmlFmtQ(cuota12),  pop:true  },
    { n:'18 cuotas', m:htmlFmtQ(cuota18),  pop:false },
  ];
  let celdas = '';
  cuotas.forEach(c => {
    const bg    = c.pop ? 'background:linear-gradient(135deg,#C9922E,#f0c060);' : 'background:#f7f5f0;border:1.5px solid #ece8e0;';
    const cn    = c.pop ? 'color:rgba(255,255,255,0.8);' : 'color:#aaa;';
    const cm    = c.pop ? 'color:#fff;' : 'color:#1a1a2e;';
    const badge = c.pop ? '<div style="font-size:7px;letter-spacing:.1em;text-transform:uppercase;background:rgba(255,255,255,0.3);color:#fff;padding:2px 6px;border-radius:10px;display:inline-block;margin-bottom:4px;">Popular</div><br>' : '';
    celdas += `<td style="width:20%;padding:3px;"><div style="border-radius:12px;padding:12px 4px;text-align:center;${bg}">${badge}` +
      `<div style="font-size:8px;letter-spacing:.08em;text-transform:uppercase;margin-bottom:5px;${cn}">${c.n}</div>` +
      `<div style="font-size:13px;font-weight:700;${cm}">${c.m}</div>` +
      '</div></td>';
  });

  // ── HTML completo ──
  return `<!DOCTYPE html><html lang="es"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Cotización ${destino} — MaKing Trips</title>
<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:'Lato',Arial,sans-serif;background:#f0ede6;padding:8px}a{text-decoration:none}@import url('https://fonts.googleapis.com/css2?family=Lato:wght@300;400;700&display=swap');</style>
</head><body>
<table width="100%" cellpadding="0" cellspacing="0"><tr><td align="center">
<table width="560" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:16px;overflow:hidden;border:1px solid #e8e4dc;max-width:560px;width:100%;">

<!-- LOGO HEADER -->
<tr><td style="background:#fff;padding:28px 28px 22px;text-align:center;border-bottom:1px solid #f0ebe0;">
  ${logoHeroTag}
  <div style="width:48px;height:2px;background:linear-gradient(90deg,#C9922E,#f0c060);border-radius:2px;margin:8px auto 0;"></div>
</td></tr>

<!-- FOTO HERO -->
<tr><td style="padding:0;">
  <div style="position:relative;height:280px;overflow:hidden;">
    ${fotoTag}
    <div style="position:absolute;top:0;left:0;right:0;bottom:0;background:linear-gradient(180deg,rgba(10,20,35,0.45) 0%,rgba(10,20,35,0.75) 100%);"></div>
    <div style="position:absolute;top:0;left:0;right:0;bottom:0;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:24px;text-align:center;">
      <div style="display:inline-block;background:rgba(255,255,255,0.15);border:1px solid rgba(255,255,255,0.35);color:#fff;font-size:10px;letter-spacing:.2em;text-transform:uppercase;padding:4px 16px;border-radius:20px;margin-bottom:12px;">Cotización de viaje</div>
      <div style="font-size:40px;font-weight:700;color:#fff;margin-bottom:6px;letter-spacing:-.5px;text-shadow:0 2px 12px rgba(0,0,0,0.4);">${destino}</div>
      ${tipoPaquete ? `<div style="display:inline-block;background:linear-gradient(135deg,#C9922E,#f0c060);color:#fff;font-size:13px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;padding:5px 18px;border-radius:20px;margin-bottom:14px;">⭐ ${tipoPaquete}</div>` : ''}
      <div style="font-size:13px;color:rgba(255,255,255,0.85);margin-bottom:20px;">${dias} días &nbsp;·&nbsp; ${noches} noches${fechaViaje ? ' &nbsp;·&nbsp; ' + fechaViaje : ''}</div>
      <div style="display:inline-block;background:linear-gradient(135deg,#C9922E,#f0c060);border-radius:14px;padding:15px 34px;">
        <div style="font-size:9px;letter-spacing:.2em;text-transform:uppercase;color:rgba(255,255,255,0.88);margin-bottom:3px;">Desde</div>
        <div style="font-size:42px;font-weight:700;color:#fff;line-height:1;">${htmlFmtQ(precioPP)}</div>
        <div style="font-size:11px;color:rgba(255,255,255,0.88);margin-top:4px;">por persona &nbsp;·&nbsp; ocupación ${ocupTipo}</div>
      </div>
    </div>
  </div>
</td></tr>

<!-- BARRA DE DATOS RÁPIDOS -->
<tr><td style="background:#f7f5f0;border-bottom:1px solid #ece8e0;">
  <table width="100%" cellpadding="0" cellspacing="0"><tr>
    <td style="width:25%;padding:12px 6px;text-align:center;border-right:1px solid #ece8e0;">
      <div style="font-size:18px;color:#2AABB5;margin-bottom:3px;">✈️</div>
      <div style="font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:#aaa;">Aerolínea</div>
      <div style="font-size:12px;font-weight:600;color:#333;margin-top:2px;">${aerolinea}</div>
    </td>
    <td style="width:25%;padding:12px 6px;text-align:center;border-right:1px solid #ece8e0;">
      <div style="font-size:18px;color:#2AABB5;margin-bottom:3px;">👥</div>
      <div style="font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:#aaa;">Personas</div>
      <div style="font-size:12px;font-weight:600;color:#333;margin-top:2px;">${personas}</div>
    </td>
    <td style="width:25%;padding:12px 6px;text-align:center;border-right:1px solid #ece8e0;">
      <div style="font-size:18px;color:#2AABB5;margin-bottom:3px;">☀️</div>
      <div style="font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:#aaa;">Días</div>
      <div style="font-size:12px;font-weight:600;color:#333;margin-top:2px;">${dias}</div>
    </td>
    <td style="width:25%;padding:12px 6px;text-align:center;">
      <div style="font-size:18px;color:#2AABB5;margin-bottom:3px;">🏨</div>
      <div style="font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:#aaa;">Hotel</div>
      <div style="font-size:10px;font-weight:600;color:#333;margin-top:2px;line-height:1.3;">
        ${linkHotel ? `<a href="${linkHotel}" target="_blank" style="color:#2AABB5;text-decoration:underline;">${hotel}</a>` : hotel}
      </div>
    </td>
  </tr></table>
</td></tr>

<!-- CUERPO PRINCIPAL -->
<tr><td style="padding:24px;">

  ${secLabel('Qué incluye tu paquete')}
  <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:20px;">${filasIncluidos}</table>

  ${secLabel('Precio por Transferencia')}
  <div style="background:linear-gradient(135deg,#2AABB5,#1d8fa0);border-radius:14px;padding:20px 22px;margin-bottom:10px;">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
      <div>
        <div style="font-size:11px;letter-spacing:.1em;text-transform:uppercase;color:rgba(255,255,255,0.75);margin-bottom:4px;">Precio por transferencia</div>
        <div style="font-size:36px;font-weight:700;color:#fff;line-height:1;">${htmlFmtQ(pTransf)}</div>
        <div style="font-size:12px;color:rgba(255,255,255,0.85);margin-top:4px;">(${personas} personas)</div>
      </div>
      <div style="text-align:right;">
        <div style="background:rgba(255,255,255,0.2);border-radius:10px;padding:10px 14px;">
          <div style="font-size:10px;color:rgba(255,255,255,0.8);margin-bottom:3px;">💰 Ahorras</div>
          <div style="font-size:22px;font-weight:700;color:#fff;">- ${htmlFmtQ(ahorro)}</div>
          <div style="font-size:10px;color:rgba(255,255,255,0.7);margin-top:2px;">vs pago con tarjeta</div>
        </div>
      </div>
    </div>
    <div style="border-top:1px solid rgba(255,255,255,0.2);padding-top:10px;display:flex;justify-content:space-between;align-items:center;">
      <span style="font-size:11px;color:rgba(255,255,255,0.75);">Precio con tarjeta</span>
      <span style="font-size:13px;color:rgba(255,255,255,0.75);text-decoration:line-through;">${htmlFmtQ(pTarjeta)}</span>
    </div>
  </div>

  <div style="background:#f7f5f0;border-radius:14px;padding:16px 18px;margin-bottom:20px;border:1px solid #ece8e0;">
    <div style="font-size:10px;letter-spacing:.15em;text-transform:uppercase;color:#2AABB5;font-weight:600;margin-bottom:12px;">🏦 ${BANCO_NOMBRE} — Datos para transferir</div>
    <div style="display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid #ece8e0;">
      <span style="font-size:11px;color:#888;">Cuenta Monetaria Q (Quetzales)</span>
      <span style="font-size:14px;font-weight:700;color:#1a1a2e;letter-spacing:.04em;">${CTA_Q}</span>
    </div>
    <div style="display:flex;justify-content:space-between;align-items:center;padding:8px 0;">
      <span style="font-size:11px;color:#888;">Cuenta Monetaria $ (Dólares)</span>
      <span style="font-size:14px;font-weight:700;color:#1a1a2e;letter-spacing:.04em;">${CTA_USD}</span>
    </div>
  </div>

  ${secLabel('Opciones de pago con tarjeta')}
  <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:20px;"><tr>${celdas}</tr></table>

  <div style="margin-bottom:20px;">
    <p style="font-size:12px;color:#555;margin-bottom:12px;">Aceptamos pagos con tarjetas <strong>Visa</strong> y <strong>Mastercard</strong>, ya sea al contado o en cuotas con bancos participantes.</p>
    <table width="100%" cellpadding="0" cellspacing="0" style="border-radius:10px;overflow:hidden;border:1px solid #ece8e0;margin-bottom:10px;">
      <tr>
        <td style="width:50%;background:#f0faf4;padding:14px 16px;vertical-align:top;border-right:1px solid #ece8e0;">
          <div style="font-size:11px;font-weight:700;color:#1a7a3c;margin-bottom:8px;">✅ Aplican para pago en cuotas</div>
          <div style="font-size:11px;color:#333;line-height:2;">Banco Industrial<br>G&amp;T Continental<br>Promerica<br>Bantrab<br>BAM<br>Ficohsa<br>Banrural Cuotas</div>
        </td>
        <td style="width:50%;background:#fffbf0;padding:14px 16px;vertical-align:top;">
          <div style="font-size:11px;font-weight:700;color:#b8860b;margin-bottom:8px;">⚠️ Solo aplican al contado</div>
          <div style="font-size:11px;color:#333;line-height:2;">BAC<br>Banrural<br>Cuscatlán<br>Micoope<br>Banco Azteca<br>Banco Antigua</div>
        </td>
      </tr>
    </table>
    <div style="font-size:10px;color:#aaa;line-height:1.7;">*Las cuotas aplican únicamente para pagos en Quetzales (GTQ).<br>*La disponibilidad de cuotas puede depender del banco emisor y de la aprobación de la transacción.<br>*American Express no aplica actualmente.</div>
  </div>

  ${validezTxt ? `<div style="background:#fffbf0;border-radius:12px;padding:12px 16px;display:flex;gap:12px;align-items:center;border:1px solid rgba(201,146,46,0.3);margin-bottom:20px;"><span style="color:#C9922E;font-size:18px;flex-shrink:0;">⏰</span><span style="font-size:12px;color:#7a6020;line-height:1.55;">Cotización válida hasta el <strong>${validezTxt}</strong>. Tarifas sujetas a disponibilidad al momento de reservar.</span></div>` : ''}

  <div style="text-align:center;margin-bottom:8px;">
    <a href="https://wa.me/50232270977?text=${encodeURIComponent('Hola MarroKing, revisé la propuesta y ')}" target="_blank"
       style="display:inline-block;background:linear-gradient(135deg,#25d366,#1ebe57);color:#fff;font-size:13px;font-weight:700;letter-spacing:.06em;padding:15px 48px;border-radius:50px;text-decoration:none;">
      💬 Confirmar mi Viaje
    </a>
  </div>
  <div style="font-size:11px;color:#bbb;text-align:center;margin-top:10px;margin-bottom:4px;">MaKing Trips &nbsp;·&nbsp; ${EMAIL_COT}</div>

</td></tr>

<!-- FOOTER -->
<tr><td style="background:#f7f5f0;padding:18px 28px;text-align:center;border-top:1px solid #ece8e0;">
  ${logoFooterTag}
  <div style="font-size:10px;color:#bbb;margin-top:6px;">Cotización ${codigo} &nbsp;·&nbsp; Precios en Quetzales (GTQ)</div>
</td></tr>

</table></td></tr></table>
</body></html>`;
}

// ─────────────────────────────────────────────
//  doGet — Web App endpoint (?codigo=PAN-005)
// ─────────────────────────────────────────────
function doGet(e) {
  try {
    const codigo = (e && e.parameter && e.parameter.codigo)
      ? String(e.parameter.codigo).trim().toUpperCase() : '';

    if (!codigo) {
      return HtmlService.createHtmlOutput(
        '<p style="font-family:sans-serif;padding:2rem;color:#888;text-align:center">' +
        '🌎 MaKing Trips — No se especificó un código de cotización.</p>'
      ).setTitle('MaKing Trips');
    }

    const ss    = SpreadsheetApp.openById('1srfMKl_ARNwyoEwnK0Hu6zq18TvOzFbhguKr9G6626U');
    const sheet = ss.getSheetByName('PAQUETES_V3');
    const datos = sheet.getDataRange().getValues();

    let fila = null;
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).trim().toUpperCase() === codigo) {
        fila = datos[i];
        break;
      }
    }

    if (!fila) {
      return HtmlService.createHtmlOutput(
        `<p style="font-family:sans-serif;padding:2rem;color:#c00;text-align:center">Cotización <strong>${codigo}</strong> no encontrada.</p>`
      ).setTitle('MaKing Trips');
    }

    return HtmlService
      .createHtmlOutput(generarCotizacionHTML(fila))
      .setTitle(`Cotización ${fila[COL.DESTINO]} — MaKing Trips`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch(err) {
    return HtmlService.createHtmlOutput(
      `<p style="font-family:sans-serif;padding:2rem;color:#c00">Error al cargar la cotización: ${err.message}</p>`
    ).setTitle('MaKing Trips');
  }
}