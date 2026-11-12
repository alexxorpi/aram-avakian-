/**
 * AVAKIAN · Conector Drive → sitio web
 *
 * Entrega al sitio el contenido de esta hoja y la lista de fotos de la carpeta «Fotos».
 * Se instala UNA sola vez:
 *   1. En la hoja «Propiedades Avakian»: Extensiones → Apps Script.
 *   2. Borra lo que haya, pega este archivo completo y guarda.
 *   3. Implementar → Nueva implementación → tipo «Aplicación web».
 *      Ejecutar como: Yo.   Quién tiene acceso: Cualquier usuario.
 *   4. Autoriza los permisos y copia la «URL de la aplicación web».
 *
 * Estructura esperada en Drive (misma carpeta que la hoja):
 *   Fotos/
 *     PORTADA/      fotos de la portada del sitio
 *     AV-001/       fotos de la propiedad con ID AV-001
 *     DES-01/       fotos del desarrollo con ID DES-01
 */

var CARPETA_FOTOS = 'Fotos';
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

  var fotos = {};
  var padres = DriveApp.getFileById(libro.getId()).getParents();
  if (padres.hasNext()) {
    var raices = padres.next().getFoldersByName(CARPETA_FOTOS);
    if (raices.hasNext()) {
      var raiz = raices.next();
      // Las fotos deben poder verse sin iniciar sesión para que el sitio las muestre.
      try {
        if (raiz.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK) {
          raiz.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        }
      } catch (e) { /* la organización no permite compartir públicamente */ }
      var carpetas = raiz.getFolders();
      while (carpetas.hasNext()) {
        var c = carpetas.next();
        var lista = [];
        var archivos = c.getFiles();
        while (archivos.hasNext()) {
          var a = archivos.next();
          if (String(a.getMimeType()).indexOf('image/') === 0) lista.push({ n: a.getName(), id: a.getId() });
        }
        lista.sort(function (x, y) { return x.n.localeCompare(y.n, undefined, { numeric: true }); });
        fotos[c.getName().trim().toUpperCase()] = lista.map(function (x) { return x.id; });
      }
    }
  }
  return { actualizado: new Date().toISOString(), hojas: hojas, fotos: fotos };
}

/** Ejecuta esta función desde el editor para revisar qué recibirá el sitio. */
function probar() {
  var d = construir();
  Logger.log('Hojas: ' + Object.keys(d.hojas).join(', '));
  Logger.log('Carpetas de fotos: ' + JSON.stringify(Object.keys(d.fotos)));
}
