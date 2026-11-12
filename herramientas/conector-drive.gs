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
    hojas[nombre] = h.getDataRange().getValues().map(function (fila) {
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

/** Ejecuta esta función desde el editor para revisar qué recibirá el sitio. */
function probar() {
  var d = construir();
  Logger.log('Hojas: ' + Object.keys(d.hojas).join(', '));
  Object.keys(d.fotos).forEach(function (k) { Logger.log('Carpeta «' + k + '»: ' + d.fotos[k].length + ' fotos'); });
  if (d.sinPermiso.length) Logger.log('NO se pudieron compartir públicamente: ' + d.sinPermiso.join(', '));
}
