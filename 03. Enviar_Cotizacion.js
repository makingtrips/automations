// ============================================================
//  Makingtrips_enviar_cotizacion.gs — MaKing Trips
//  Body del correo: diseño original turquesa/dorado/beige
//  Adjunto: generarCotizacionHTML() de Cotizacion_HTML.gs
//  Índices: catálogo PAQUETES_V3 (72 cols A→BT)
// ============================================================

const REMITENTE_NOMBRE = 'MaKing Trips';
const EMAIL_AGENCIA    = 'makingtripsgt@gmail.com';
const LOGO_HORIZONTAL_TRANSP_ID = '1ZHR83YLXGTBV2QnEm1z_feAIOK9XXh8d';

function logoUrl_(fileId) {
  return 'https://drive.google.com/uc?export=view&id=' + fileId;
}

// ─────────────────────────────────────────────
//  obtenerFotoUrlCorreo
//  Busca la URL de foto en FOTOS_DESTINOS para el correo
// ─────────────────────────────────────────────
function obtenerFotoUrlCorreo(destino) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FOTOS_DESTINOS');
    if (!sheet) return '';
    const datos = sheet.getDataRange().getValues();
    const dest  = String(destino).toLowerCase().trim();
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).toLowerCase().trim() === dest) return String(datos[i][1]).trim();
    }
    return '';
  } catch(e) { return ''; }
}

// ─────────────────────────────────────────────
//  enviarCotizacionMail — función principal
// ─────────────────────────────────────────────
function enviarCotizacionMail() {
  const ui = SpreadsheetApp.getUi();
  try {
    const fila    = obtenerFilaActiva();  // Codigo.gs
    const codigo  = String(fila[WA.CODIGO]  || '').trim().toUpperCase();
    const destino = String(fila[WA.DESTINO] || '').trim();
    if (!codigo) throw new Error('La fila seleccionada no tiene código (columna A).');

    // ── Pedir correo ──
    const respEmail = ui.prompt(
      '📧 Enviar cotización por correo',
      `Destino: ${destino} (${codigo})\n\nCorreo del cliente:`,
      ui.ButtonSet.OK_CANCEL
    );
    if (respEmail.getSelectedButton() !== ui.Button.OK) return;
    const correoCliente = respEmail.getResponseText().trim();
    if (!correoCliente || !correoCliente.includes('@')) {
      ui.alert('⚠️ Correo inválido. Verifica el formato.');
      return;
    }

    // ── Pedir nombre ──
    const respNombre = ui.prompt(
      '👤 Nombre del cliente',
      'Ingresa el nombre para personalizar el correo (opcional):',
      ui.ButtonSet.OK_CANCEL
    );
    const nombreCliente = (respNombre.getSelectedButton() === ui.Button.OK)
      ? respNombre.getResponseText().trim() : '';

    // ── Datos para el body ──
    const noches   = fila[WA.NOCHES]    || '';
    const dias     = fila[WA.DIAS]      || '';
    const fotoSrc  = obtenerFotoUrlCorreo(destino);
    const urlCot   = `https://script.google.com/macros/s/AKfycbxs2O0fjan9-R1lhXCNLBIvzG0r4Ci2QJ8RwjwsW_XVR_rPwWkfFa6HMRHeRCfxEMah/exec?codigo=${codigo}`;
    const asunto   = `Cotizacion de viaje a ${destino} — MaKing Trips`;
    const saludo   = nombreCliente ? `Hola, ${nombreCliente} ✈️` : 'Hola ✈️';
    const duracion = (dias && noches) ? `${dias}D / ${noches}N` : '';

    // ── HTML body del correo (diseño original) ──
    const htmlCompleto =
      '<div style="font-family:\'Lato\',Arial,Helvetica,sans-serif;max-width:560px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;border:1px solid #e8e4dc;">' +

      // HEADER turquesa
      '<div style="background:linear-gradient(135deg,#2AABB5,#1d8fa0);padding:32px 32px 28px;text-align:center;">' +
        '<div style="font-size:11px;letter-spacing:.25em;text-transform:uppercase;color:rgba(255,255,255,0.7);margin-bottom:10px;">MaKing Trips</div>' +
        '<div style="font-size:28px;font-weight:700;color:#fff;letter-spacing:-.5px;margin-bottom:6px;">' + destino + '</div>' +
        (duracion ? '<div style="font-size:13px;color:rgba(255,255,255,.75);margin-bottom:8px;">' + duracion + '</div>' : '') +
        '<div style="width:40px;height:2px;background:linear-gradient(90deg,#C9922E,#f0c060);border-radius:2px;margin:0 auto;"></div>' +
      '</div>' +

      // FOTO del destino
      (fotoSrc ? '<div style="height:180px;overflow:hidden;"><img src="' + fotoSrc + '" style="width:100%;height:180px;object-fit:cover;display:block;"></div>' : '') +

      // CUERPO
      '<div style="padding:28px 32px;">' +

        '<p style="font-size:15px;font-weight:600;color:#1a1a2e;margin-bottom:16px;">' + saludo + '</p>' +

        '<p style="font-size:13px;color:#555;line-height:1.8;margin-bottom:16px;">' +
          'Estuve revisando opciones, comparando fechas y buscando la mejor combinación para hacer de tu viaje a <strong style="color:#1a1a2e;">' + destino + '</strong> una experiencia que valga cada quetzal.' +
        '</p>' +
        '<p style="font-size:13px;color:#555;line-height:1.8;margin-bottom:24px;">' +
          'En el archivo adjunto encontrarás la propuesta completa — con los detalles del paquete, precios y opciones de pago. La preparé pensando en lo que me comentaste, así que te pido que la revises con calma.' +
        '</p>' +

        // Card beige con borde dorado
        '<div style="background:linear-gradient(135deg,#f7f5f0,#f0ede6);border-radius:12px;padding:20px 22px;margin-bottom:24px;border-left:4px solid #C9922E;">' +
          '<div style="font-size:10px;letter-spacing:.18em;text-transform:uppercase;color:#C9922E;font-weight:600;margin-bottom:6px;">Lo que encontrarás adentro</div>' +
          '<div style="font-size:13px;color:#444;line-height:1.9;">' +
            '&#9992;&nbsp; Detalles completos del paquete<br>' +
            '&#127963;&nbsp; Hotel y servicios incluidos<br>' +
            '&#128176;&nbsp; Precio especial por transferencia<br>' +
            '&#128179;&nbsp; Opciones de pago en cuotas' +
          '</div>' +
        '</div>' +

       
        '<p style="font-size:13px;color:#555;line-height:1.8;margin-bottom:28px;">' +
          'Cualquier pregunta, ajuste o detalle adicional — estoy a la orden. Con gusto hacemos de este viaje exactamente lo que tienes en mente.' +
        '</p>' +

        // Firma
        '<div style="border-top:1px solid #f0ede6;padding-top:20px;">' +
          '<div style="font-size:13px;font-weight:700;color:#1a1a2e;">MaKing Trips</div>' +
          '<div style="font-size:11px;color:#aaa;margin-top:2px;font-style:italic;">Creating journeys with purpose</div>' +
          '<div style="font-size:11px;color:#2AABB5;margin-top:4px;">' + EMAIL_AGENCIA + '</div>' +
        '</div>' +

      '</div>' +
      '</div>';

    // ── Adjunto: HTML premium de cotización ──
    const htmlAdj  = generarCotizacionHTML(fila);  // Cotizacion_HTML.gs
    const nombreAdj = `Cotizacion_${codigo}_${destino.replace(/[^a-zA-Z0-9]/g,'')}.html`;
    const blobAdj   = Utilities.newBlob(htmlAdj, 'text/html', nombreAdj);

    // ── Enviar — string vacío como body, HTML en opciones ──
    GmailApp.sendEmail(correoCliente, asunto, '', {
      htmlBody:    htmlCompleto,
      name:        REMITENTE_NOMBRE,
      replyTo:     EMAIL_AGENCIA,
      attachments: [blobAdj],
    });

    // ── Guardar en Drive ──
    try { guardarEnDrive(blobAdj, codigo); } catch(ed) { Logger.log('Drive: ' + ed.message); }

    // ── Modal de éxito ──
    ui.showModalDialog(
      HtmlService.createHtmlOutput(modalExitoCorreo(
        'Correo enviado',
        `La cotización fue enviada a <strong>${correoCliente}</strong>.`,
        '📧'
      )).setWidth(400).setHeight(220),
      'Correo enviado'
    );

  } catch(e) {
    ui.alert('❌ Error al enviar: ' + e.message);
    Logger.log(e.stack);
  }
}

// ─────────────────────────────────────────────
//  guardarEnDrive
// ─────────────────────────────────────────────
function guardarEnDrive(blob, codigo) {
  const CARPETA = 'COTIZACIONES MaKing Trips';
  const iter    = DriveApp.getFoldersByName(CARPETA);
  const carpeta = iter.hasNext() ? iter.next() : DriveApp.createFolder(CARPETA);
  const archivo = carpeta.createFile(blob);
  Logger.log(`Drive: ${archivo.getName()} — ${archivo.getId()}`);
  return archivo;
}

// ─────────────────────────────────────────────
//  modalExitoCorreo — modal bonito de confirmación
// ─────────────────────────────────────────────
function modalExitoCorreo(titulo, mensaje, emoji) {
  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:Arial,sans-serif;background:#f5f7fa;padding:24px;text-align:center;color:#111}
    .ico{font-size:36px;margin-bottom:12px}
    .titulo{font-size:16px;font-weight:700;margin-bottom:10px;color:#1a1a2e}
    .msg{font-size:13px;color:#555;line-height:1.6;margin-bottom:20px}
    .btn{display:inline-block;padding:10px 28px;background:#2AABB5;color:#fff;font-weight:700;font-size:13px;border-radius:50px;border:none;cursor:pointer}
  </style></head><body>
  <div class="ico">${emoji}</div>
  <div class="titulo">${titulo}</div>
  <div class="msg">${mensaje}</div>
  <button class="btn" onclick="google.script.host.close()">Cerrar</button>
  </body></html>`;
}

// ─────────────────────────────────────────────
//  Alias para el botón del sheet
// ─────────────────────────────────────────────
function enviarPorCorreo() {
  enviarCotizacionMail();
}