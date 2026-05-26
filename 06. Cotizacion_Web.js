// ============================================================
//  06_CotizacionWeb.gs — MaKing Trips
//  Web App endpoint — redirige al doGet de Cotizacion_HTML.gs
// ============================================================

const COTIZACION_SHEET_ID = '1srfMKl_ARNwyoEwnK0Hu6zq18TvOzFbhguKr9G6626U';

function doGet(e) {
  var codigo = (e && e.parameter && e.parameter.codigo)
    ? String(e.parameter.codigo).trim().toUpperCase() : '';

  if (!codigo) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:Arial;padding:40px;text-align:center;color:#888;">🌎 MaKing Trips — No se especificó un código de cotización.</div>'
    ).setTitle('MaKing Trips');
  }

  try {
    var ss    = SpreadsheetApp.openById(COTIZACION_SHEET_ID);
    var sheet = ss.getSheetByName('PAQUETES_V3');
    var datos = sheet.getDataRange().getValues();

    var fila = null;
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).trim().toUpperCase() === codigo) {
        fila = datos[i];
        break;
      }
    }

    if (!fila) {
      return HtmlService.createHtmlOutput(
        '<div style="font-family:Arial;padding:40px;text-align:center;color:#888;">No se encontró la cotización <strong>' + codigo + '</strong>.</div>'
      ).setTitle('MaKing Trips');
    }

    var htmlContent = generarCotizacionHTML(fila);

    return HtmlService
      .createHtmlOutput(htmlContent)
      .setTitle('Cotización ' + fila[2] + ' — MaKing Trips')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch(err) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:Arial;padding:40px;text-align:center;color:#c00;">Error: ' + err.message + '</div>'
    ).setTitle('MaKing Trips');
  }
}