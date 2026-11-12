/**
 * AVAKIAN · Conector Drive → sitio web
 *
 * Entrega al sitio el contenido de esta hoja y la lista de fotos de cada carpeta de propiedad.
 * Se instala UNA sola vez:
 *   1. En la hoja «Propiedades Avakian»: Extensiones → Apps Script.
 *   2. Borra lo que haya, pega este archivo completo y guarda.
 *   3. Implementar → Nueva implementación → tipo «Aplicación web».
 *      Ejecutar como: Yo.   Quién tiene acceso: Cualquier usuario.
 *   4. Autoriza los permisos y copia la «URL de la aplicación web».
 *
 * Estructura esperada en Drive (la hoja vive dentro de la carpeta «aram avakian»):
 *   aram avakian/
 *     Propiedades Avakian     esta hoja
 *     PORTADA/                fotos de la portada del sitio
 *     landmark 630k/          fotos de una propiedad (una carpeta por propiedad)
 * El nombre de cada carpeta se escribe en la columna «Carpeta de fotos» de su fila.
 */

var SEGUNDOS_CACHE = 60;

function doGet() {
  var cache = CacheService.getScriptCache();
  var salida = cache.get('datos');
  if (!salida) {
    salida = JSON.stringify(construir());
    try { cache.put('datos', salida, SEGUNDOS_CACHE); } catch (e) { /* más de 100 KB: sin caché */ }
  }
  return ContentService.createTextOutput(salida).setMimeType(ContentService.MimeType.JSON);
}

function construir() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojas = {};
  libro.getSheets().forEach(function (h) {
    var nombre = h.getName();
    if (/^instrucciones/i.test(nombre)) return;
    if (/^(propiedades|desarrollos)/i.test(nombre)) { try { ubicar(h); } catch (e) { /* sin mapa, el resto sigue */ } }
    hojas[nombre] =h.getDataRange().getValues().map(function (fila) {
      return fila.map(function (v) { return v instanceof Date ? Utilities.formatDate(v, libro.getSpreadsheetTimeZone(), 'dd/MM/yyyy') : v; });
    });
  });

  // Cada subcarpeta de la carpeta donde vive esta hoja es una carpeta de fotos.
  var fotos = {};
  var sinPermiso = [];
  var padres = DriveApp.getFileById(libro.getId()).getParents();
  if (padres.hasNext()) {
    var carpetas = padres.next().getFolders();
    while (carpetas.hasNext()) {
      var c = carpetas.next();
      // Las fotos deben poder verse sin iniciar sesión para que el sitio las muestre.
      // Solo se comparten las carpetas de fotos, nunca la hoja.
      try {
        if (c.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK) {
          c.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        }
      } catch (e) { sinPermiso.push(c.getName()); }
      var lista = [];
      var archivos = c.getFiles();
      while (archivos.hasNext()) {
        var a = archivos.next();
        if (String(a.getMimeType()).indexOf('image/') === 0) lista.push({ n: a.getName(), id: a.getId() });
      }
      lista.sort(function (x, y) { return x.n.localeCompare(y.n, undefined, { numeric: true }); });
      fotos[c.getName().trim()] = lista.map(function (x) { return x.id; });
    }
  }
  return { actualizado: new Date().toISOString(), hojas: hojas, fotos: fotos, sinPermiso: sinPermiso };
}

/**
 * Completa la columna «Ubicación» con coordenadas:
 *   - si tiene un link de Google Maps (también los cortos de «Compartir»), lo convierte;
 *   - si está vacía pero hay «Dirección», busca la dirección en Google Maps.
 * El link original se conserva como nota de la celda.
 */
var RE_COORDS = [/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/, /@(-?\d+\.\d+),(-?\d+\.\d+)/, /[?&](?:q|ll|query|center)=(-?\d+\.\d+)(?:,|%2C)(-?\d+\.\d+)/];

function leerCoords(url) {
  for (var i = 0; i < RE_COORDS.length; i++) { var m = url.match(RE_COORDS[i]); if (m) return m[1] + ', ' + m[2]; }
  return '';
}

function coordsDeLink(url) {
  for (var i = 0; i < 5; i++) {
    var c = leerCoords(url);
    if (c) return c;
    var r = UrlFetchApp.fetch(url, { followRedirects: false, muteHttpExceptions: true });
    var cab = r.getHeaders();
    var sig = cab['Location'] || cab['location'];
    if (!sig) return '';
    url = sig;
  }
  return '';
}

function ubicar(hoja) {
  var datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return;
  var cU = -1, cD = -1;
  datos[0].forEach(function (t, i) {
    t = String(t).toLowerCase();
    if (t.indexOf('ubicaci') === 0) cU = i;
    if (t.indexOf('direcci') === 0) cD = i;
  });
  if (cU < 0) return;
  var cache = CacheService.getScriptCache();
  for (var f = 1; f < datos.length; f++) {
    var ubi = String(datos[f][cU]).trim();
    var dir = cD < 0 ? '' : String(datos[f][cD]).trim();
    var esLink = /^https?:\/\//i.test(ubi);
    if (ubi && !esLink) continue;                         // ya tiene coordenadas
    var origen = esLink ? ubi : dir;
    if (!origen) continue;
    var clave = 'sinmapa:' + Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, origen));
    if (cache.get(clave)) continue;                       // ya se intentó hace poco y no se encontró
    var coords = '';
    try { if (esLink) coords = coordsDeLink(ubi); } catch (e) {}
    if (!coords && dir && !/^https?:/i.test(dir)) {
      try {
        var texto = (/tijuana/i.test(dir) ? dir : dir + ', Tijuana, Baja California') + ', México';
        var r = Maps.newGeocoder().setRegion('mx').setLanguage('es').geocode(texto);
        if (r.status === 'OK' && r.results.length) {
          var p = r.results[0].geometry.location;
          coords = p.lat.toFixed(6) + ', ' + p.lng.toFixed(6);
        }
      } catch (e) {}
    }
    if (coords) {
      var celda = hoja.getRange(f + 1, cU + 1);
      if (esLink) celda.setNote(ubi);
      celda.setValue(coords);
    } else {
      cache.put(clave, '1', 21600);                       // reintenta en 6 horas
    }
  }
  SpreadsheetApp.flush();
}

/** Ejecuta esta función desde el editor para revisar qué recibirá el sitio. */
function probar() {
  var d = construir();
  Logger.log('Hojas: ' + Object.keys(d.hojas).join(', '));
  Object.keys(d.fotos).forEach(function (k) { Logger.log('Carpeta «' + k + '»: ' + d.fotos[k].length + ' fotos'); });
  if (d.sinPermiso.length) Logger.log('NO se pudieron compartir públicamente: ' + d.sinPermiso.join(', '));
}
