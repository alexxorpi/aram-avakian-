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
  var avisos = [];
  try { migrar(libro); } catch (e) { avisos.push('migrar: ' + e); }
  libro.getSheets().forEach(function (h) {
    var nombre = h.getName();
    if (/^instrucciones/i.test(nombre)) return;
    if (/^(propiedades|desarrollos)/i.test(nombre)) { try { ubicar(h, avisos); } catch (e) { avisos.push('ubicar ' + nombre + ': ' + e); } }
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
  return { version: 3, actualizado: new Date().toISOString(), hojas: hojas, fotos: fotos, sinPermiso: sinPermiso, avisos: avisos };
}

/**
 * Ajusta la hoja una sola vez: encabezados más claros y las instrucciones al día.
 * (El sitio reconoce tanto los encabezados viejos como los nuevos.)
 */
function migrar(libro) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('hoja_v3')) return;
  var nuevos = [
    [/^im[aá]genes/i, 'Portada (número de la foto que va primero)',
      'Escribe solo el número de la foto que quieres de portada, contando las fotos de la carpeta en orden de nombre (1 = la primera). Vacío = la primera.'],
    [/^zona/i, 'Zona (colonia)', 'Solo la colonia o zona. Ej. Calete, Zona Río, Playas de Tijuana.'],
    [/^direcci/i, 'Dirección (calle y número, en texto)', 'Texto que se muestra en la ficha. Ej. Lerdo de Tejada 13, Calete. Aquí NO va el link de Google Maps.'],
    [/^ubicaci/i, 'Ubicación (link de Google Maps)',
      'Pega aquí el link de Google Maps de la propiedad (botón Compartir → Copiar vínculo). En uno o dos minutos se convierte solo en coordenadas y aparece el pin. Si lo dejas vacío, se busca la Dirección.']
  ];
  libro.getSheets().forEach(function (h) {
    if (!/^(propiedades|desarrollos)/i.test(h.getName()) || h.getLastColumn() < 1) return;
    var fila = h.getRange(1, 1, 1, h.getLastColumn());
    var enc = fila.getValues()[0];
    enc.forEach(function (t, i) {
      nuevos.forEach(function (n) {
        if (n[0].test(String(t))) { var c = h.getRange(1, i + 1); c.setValue(n[1]); c.setNote(n[2]); }
      });
    });
  });
  var ins = libro.getSheets().filter(function (h) { return /^instrucciones/i.test(h.getName()); })[0];
  if (ins && ins.getLastRow() > 0) {
    var col = ins.getRange(1, 1, ins.getLastRow(), 1);
    var textos = col.getValues().map(function (r) {
      var t = String(r[0]);
      if (/^3\.\s+UBICACI/i.test(t)) return ['3.  UBICACIÓN: en Google Maps abre la propiedad, toca «Compartir» y copia el vínculo. Pégalo en la columna «Ubicación». En uno o dos minutos se convierte solo en coordenadas y aparece el pin en el mapa. En «Dirección» escribe la calle y el número en texto, y en «Zona» solo la colonia.'];
      if (/^4\.\s+FOTOS/i.test(t)) return ['4.  FOTOS: en la carpeta de Drive «aram avakian» crea una carpeta para la propiedad y sube ahí las fotos. Escribe el nombre de esa carpeta en la última columna, «Carpeta de fotos». En «Portada» escribe el número de la foto que quieres que salga primero (1 = la primera por nombre).'];
      return [t];
    });
    col.setValues(textos);
  }
  props.setProperty('hoja_v3', '1');
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

function ubicar(hoja, avisos) {
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
    if (!ubi && /^https?:\/\//i.test(dir)) ubi = dir;       // el link quedó en «Dirección»
    var esLink = /^https?:\/\//i.test(ubi);
    if (ubi && !esLink) continue;                         // ya tiene coordenadas
    var origen = esLink ? ubi : dir;
    if (!origen) continue;
    var clave = 'sinmapa3:' + Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, origen));
    if (cache.get(clave)) continue;                       // ya se intentó hace poco y no se encontró
    var coords = '';
    try { if (esLink) coords = coordsDeLink(ubi); } catch (e) { avisos.push('link fila ' + (f + 1) + ': ' + e); }
    if (!coords && dir && !/^https?:/i.test(dir)) {
      try {
        var texto = (/tijuana/i.test(dir) ? dir : dir + ', Tijuana, Baja California') + ', México';
        var r = Maps.newGeocoder().setRegion('mx').setLanguage('es').geocode(texto);
        if (r.status === 'OK' && r.results.length) {
          var p = r.results[0].geometry.location;
          coords = p.lat.toFixed(6) + ', ' + p.lng.toFixed(6);
        }
      } catch (e) { avisos.push('direccion fila ' + (f + 1) + ': ' + e); }
    }
    if (coords) {
      var celda = hoja.getRange(f + 1, cU + 1);
      if (esLink) celda.setNote(ubi);
      celda.setValue(coords);
    } else {
      avisos.push('sin ubicacion fila ' + (f + 1));
      cache.put(clave, '1', 600);                         // reintenta en 10 minutos
    }
  }
  SpreadsheetApp.flush();
}

/** Ejecuta esta función desde el editor para revisar qué recibirá el sitio. */
function probar() {
  var d = construir();
  Logger.log('Hojas: ' + Object.keys(d.hojas).join(', '));
  Object.keys(d.fotos).forEach(function (k) { Logger.log('Carpeta «' + k + '»: ' + d.fotos[k].length + ' fotos'); });
  Logger.log('Avisos: ' + JSON.stringify(d.avisos));
  if (d.sinPermiso.length) Logger.log('NO se pudieron compartir públicamente: ' + d.sinPermiso.join(', '));
}
