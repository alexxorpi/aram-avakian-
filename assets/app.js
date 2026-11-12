/* =========================================================
   AVAKIAN · Grupo Inmobiliario
   Lee la hoja de Google Drive (FUENTE) o, si no hay, datos/propiedades.xlsx.
   ========================================================= */
(() => {
'use strict';

// Dirección de la aplicación web de Apps Script que entrega la hoja de Drive y sus fotos.
// Vacío = el sitio usa el Excel local de la carpeta «datos».
const FUENTE = 'https://script.google.com/macros/s/AKfycbyhGSfFcSMqT1Snrr73Qe_KukNqW5hrfk3xFzDAXYCG6ELqaeomdBuIpoiLRL6NdaMD0w/exec';
const EXCEL = 'datos/propiedades.xlsx';
const IMG_DIR = 'imagenes/';
const TIJUANA = [-117.0382, 32.5149];           // [lng, lat]
const HERO_END = { center: [-117.0218, 32.5262], zoom: 14.7, pitch: 60, bearing: 24 };

const SAT_STYLE = {
  version: 8,
  sources: { sat: { type: 'raster', tileSize: 256, maxzoom: 18,
    tiles: ['https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}'],
    attribution: 'Imágenes © Esri, Maxar, Earthstar Geographics' } },
  layers: [{ id: 'sat', type: 'raster', source: 'sat' }]
};
const STREET_STYLE = 'https://tiles.openfreemap.org/styles/positron';

const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const reduced = matchMedia('(prefers-reduced-motion: reduce)').matches;

/* ---------- utilidades ---------- */
const norm = s => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const num = v => { if (typeof v === 'number') return v; const n = parseFloat(String(v).replace(/[^0-9.\-]/g, '')); return isFinite(n) ? n : null; };
const list = v => String(v ?? '').split(/[,\n;]+/).map(s => s.trim()).filter(Boolean);
const yes = v => !norm(v).startsWith('n');     // vacío = sí
const safeUrl = u => /^https?:\/\//i.test(String(u).trim()) ? String(u).trim() : '';

function field(row, ...keys) {
  const hs = Object.keys(row);
  for (const k of keys) { const h = hs.find(h => h === k || h.startsWith(k + ' ')); if (h && row[h] !== '') return row[h]; }
  return '';
}
function parseLoc(v) {
  const s = String(v ?? '').trim(); if (!s) return null;
  const m = s.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/) || s.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/) ||
            s.match(/[?&](?:q|ll|query|center)=(-?\d+\.\d+)(?:,|%2C)\s*(-?\d+\.\d+)/i) ||
            s.match(/^\(?\s*(-?\d+(?:\.\d+)?)\s*[,;\s]\s*(-?\d+(?:\.\d+)?)\s*\)?$/);
  if (!m) return null;
  let lat = +m[1], lng = +m[2];
  if (Math.abs(lat) > 90 && Math.abs(lng) <= 90) [lat, lng] = [lng, lat];
  if (Math.abs(lat) > 90 || Math.abs(lng) > 180) return null;
  return [lng, lat];
}
const driveImg = id => `https://drive.google.com/thumbnail?id=${id}&sz=w1600`;
const imgSrc = s => {
  const d = s.match(/drive\.google\.com\/(?:file\/d\/|open\?id=|uc\?(?:[\w=&]*&)?id=|thumbnail\?id=)([\w-]{20,})/);
  if (d) return driveImg(d[1]);
  return /^(https?:)?\/\//i.test(s) || s.startsWith('/') || s.startsWith('data:') ? s : IMG_DIR + encodeURI(s);
};
// Si la miniatura de Drive falla se intenta el servidor alterno de Google y, al final, el marcador.
window.__imgErr = im => {
  const d = (im.currentSrc || im.src).match(/drive\.google\.com\/thumbnail\?id=([\w-]+)/);
  if (d && !im.dataset.alt) { im.dataset.alt = 1; im.src = `https://lh3.googleusercontent.com/d/${d[1]}=w1600`; return; }
  im.onerror = null; im.src = PLACEHOLDER;
};
const PLACEHOLDER = 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 300"><rect width="400" height="300" fill="#E7DFD2"/><text x="200" y="158" text-anchor="middle" font-family="Arial" font-weight="700" font-size="26" letter-spacing="2" fill="#B9AE9C">AVAKIAN</text></svg>');

function money(n, cur) {
  if (n == null) return 'Precio a consultar';
  return '$' + Math.round(n).toLocaleString('es-MX') + ' ' + cur;
}
function moneyShort(n) {
  if (n == null) return '•';
  if (n >= 1e6) return '$' + (n / 1e6).toFixed(n % 1e6 === 0 ? 0 : (n >= 1e7 ? 1 : 2)).replace(/\.?0+$/, '') + 'M';
  if (n >= 1e3) return '$' + Math.round(n / 1e3) + 'k';
  return '$' + n;
}

/* ---------- lectura de datos (hoja de Drive o Excel) ---------- */
function sheetRows(sheets, name) {
  const sn = Object.keys(sheets).find(n => norm(n).startsWith(name)); if (!sn) return [];
  const [head, ...rows] = sheets[sn] || []; if (!head) return [];
  const keys = head.map(norm);
  return rows.map(r => { const o = {}; keys.forEach((k, i) => { if (k) { const v = r[i]; o[k] = typeof v === 'string' ? v.trim() : (v ?? ''); } }); return o; });
}
function parseWorkbook(buf) {
  const wb = XLSX.read(buf, { type: 'array' }), sheets = {};
  wb.SheetNames.forEach(n => { sheets[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: '' }); });
  return parseSheets(sheets, {});
}
function parseSheets(wb, fotos) {
  // Las carpetas de fotos se buscan sin distinguir mayúsculas, acentos ni espacios de más.
  const byName = {}; Object.keys(fotos).forEach(k => { byName[norm(k)] = fotos[k]; });
  const folder = (...names) => { for (const n of names) { const f = n && byName[norm(n)]; if (f && f.length) return f.map(driveImg); } return []; };
  const pics = (v, p) => { const l = list(v).map(imgSrc); return l.length ? l : folder(field(p.raw, 'carpeta'), p.id, p.title); };
  const props = sheetRows(wb, 'propiedades').filter(r => field(r, 'titulo') && yes(field(r, 'publicar'))).map((r, i) => {
    const op = norm(field(r, 'operacion')).startsWith('r') ? 'renta' : 'venta';
    const estatus = String(field(r, 'estatus') || 'Disponible');
    // Si faltan columnas, se deducen del texto de la descripción (ej. «3 recámaras y 3 baños»).
    const txt = norm(field(r, 'titulo') + ' ' + field(r, 'descripcion'));
    const guess = re => { const m = txt.match(re); return m ? +m[1] : null; };
    const tipoTxt = [['Departamento', /\b(departamento|depa|penthouse|loft)\b/], ['Casa', /\b(casa|residencia)\b/], ['Terreno', /\b(terreno|lote)\b/],
      ['Oficina', /\boficinas?\b/], ['Local', /\blocal\b/], ['Bodega', /\bbodega\b/]].find(t => t[1].test(txt));
    return {
      kind: 'prop', id: String(field(r, 'id') || 'P' + (i + 1)), title: String(field(r, 'titulo')), op,
      tipo: String(field(r, 'tipo') || (tipoTxt ? tipoTxt[0] : 'Propiedad')), price: num(field(r, 'precio')), cur: String(field(r, 'moneda') || 'MXN').toUpperCase(),
      zone: String(field(r, 'zona')), address: String(field(r, 'direccion')), loc: parseLoc(field(r, 'ubicacion')),
      beds: num(field(r, 'recamaras')) ?? guess(/\b(\d{1,2}) (?:recamaras?|habitaciones|habitacion|cuartos)\b/),
      baths: num(field(r, 'banos')) ?? guess(/\b(\d{1,2}) banos?\b/), park: num(field(r, 'estacionamientos')),
      m2c: num(field(r, 'construccion')), m2t: num(field(r, 'terreno')), desc: String(field(r, 'descripcion')),
      amen: list(field(r, 'amenidades')), imgs: [], video: safeUrl(field(r, 'video')),
      feat: norm(field(r, 'destacada')).startsWith('s'), estatus, off: /vendid|rentad/.test(norm(estatus)), raw: r
    };
  });
  const builds = sheetRows(wb, 'desarrollos').filter(r => field(r, 'nombre') && yes(field(r, 'publicar'))).map((r, i) => ({
    kind: 'build', id: String(field(r, 'id') || 'D' + (i + 1)), title: String(field(r, 'nombre')), etapa: String(field(r, 'etapa') || 'Desarrollo'),
    avance: num(field(r, 'avance')), entrega: String(field(r, 'entrega')), zone: String(field(r, 'zona')), loc: parseLoc(field(r, 'ubicacion')),
    units: num(field(r, 'unidades')), price: num(field(r, 'precio')), cur: String(field(r, 'moneda') || 'MXN').toUpperCase(),
    desc: String(field(r, 'descripcion')), amen: list(field(r, 'amenidades')), imgs: [], raw: r, brochure: safeUrl(field(r, 'brochure'))
  }));
  const cfg = {};
  sheetRows(wb, 'configuracion').forEach(r => { const k = norm(field(r, 'clave')); if (k) cfg[k] = String(field(r, 'valor')).trim(); });
  [...props, ...builds].forEach(p => { p.imgs = pics(field(p.raw, 'imagenes'), p); delete p.raw; });
  if (!cfg['imagenes de portada'] && folder('PORTADA').length) cfg['imagenes de portada'] = folder('PORTADA').join(',');
  return { props, builds, cfg };
}

/* ---------- estado ---------- */
let DATA = { props: [], builds: [], cfg: {} };
const F = { op: '', tipo: '', rec: 0, sort: '', q: '' };
let map, markers = new Map(), popup, miniMap, useHeroVideo = () => {}, heroVideoSet = '';

const priceLabel = p => p.kind === 'build'
  ? (p.price != null ? `Desde ${money(p.price, p.cur)}` : 'Precio a consultar')
  : money(p.price, p.cur) + (p.op === 'renta' && p.price != null ? ' <small>/ mes</small>' : '');
const specsHtml = p => [
  p.beds ? `<span><b>${p.beds}</b> rec.</span>` : '', p.baths ? `<span><b>${p.baths}</b> baños</span>` : '',
  p.m2c ? `<span><b>${p.m2c.toLocaleString('es-MX')}</b> m²</span>` : (p.m2t ? `<span><b>${p.m2t.toLocaleString('es-MX')}</b> m² terreno</span>` : '')
].join('');
const cover = p => p.imgs[0] || PLACEHOLDER;
const img = (src, alt, eager) => `<img src="${esc(src)}" alt="${esc(alt)}" ${eager ? '' : 'loading="lazy"'} onerror="__imgErr(this)">`;

/* ---------- render ---------- */
function applyConfig() {
  const c = DATA.cfg;
  $$('[data-cfg]').forEach(el => { const v = c[el.dataset.cfg]; if (v) el.textContent = v; });
  if (c.nombre) document.title = `${c.nombre} · ${c.eslogan || 'Grupo Inmobiliario'} en Tijuana`;
  const wa = (c.whatsapp || '').replace(/\D/g, '');
  const items = [
    wa && ['WhatsApp', `<a href="https://wa.me/${wa}" target="_blank" rel="noopener">+${esc(wa)}</a>`],
    c.telefono && ['Teléfono', `<a href="tel:${esc(c.telefono.replace(/[^\d+]/g, ''))}">${esc(c.telefono)}</a>`],
    c.correo && ['Correo', `<a href="mailto:${esc(c.correo)}">${esc(c.correo)}</a>`],
    safeUrl(c.instagram) && ['Instagram', `<a href="${esc(safeUrl(c.instagram))}" target="_blank" rel="noopener">Ver perfil</a>`],
    safeUrl(c.facebook) && ['Facebook', `<a href="${esc(safeUrl(c.facebook))}" target="_blank" rel="noopener">Ver página</a>`],
    c.direccion && ['Oficina', `<span>${esc(c.direccion)}</span>`],
    c.horario && ['Horario', `<span>${esc(c.horario)}</span>`]
  ].filter(Boolean);
  const hv = c['video de portada'];
  const hi = list(c['imagenes de portada']).map(imgSrc);
  $$('#heroScenes .scene img').forEach((im, i) => { if (hi[i]) { im.src = hi[i]; customHero = true; } });
  if (hv && hv !== heroVideoSet) { heroVideoSet = hv; useHeroVideo(/^https?:\/\//i.test(hv) || hv.includes('/') ? hv : 'assets/' + encodeURI(hv)); }
  $('#contactData').innerHTML = items.map(([k, v]) => `<li><small>${k}</small>${v}</li>`).join('');
}

function fillOptions() {
  const tipos = [...new Set(DATA.props.map(p => p.tipo))].sort();
  const opts = tipos.map(t => `<option value="${esc(norm(t))}">${esc(t)}</option>`).join('');
  $('#hsTipo').innerHTML = '<option value="">Cualquiera</option>' + opts;
  $('#fTipo').innerHTML = '<option value="">Tipo</option>' + opts;
  $('#zonas').innerHTML = [...new Set(DATA.props.map(p => p.zone).filter(Boolean))].sort().map(z => `<option value="${esc(z)}">`).join('');
}

function renderFeatured() {
  let feats = DATA.props.filter(p => p.feat && !p.off);
  if (!feats.length) feats = DATA.props.slice(0, 4);
  $('#destacadas').hidden = !feats.length;
  $('#rail').innerHTML = feats.map(p => `
    <button class="feat" data-open="${esc(p.id)}">
      <div class="feat__img">${img(cover(p), p.title)}
        <div class="feat__over">
          <span class="tag tag--accent">${p.op === 'renta' ? 'En renta' : 'En venta'}</span>
          <h3>${esc(p.title)}</h3>
          <p class="feat__price">${priceLabel(p)}</p>
          <p class="specs">${p.zone ? `<span>${esc(p.zone)}</span>` : ''}${specsHtml(p)}</p>
        </div>
      </div>
    </button>`).join('');
}

function filtered() {
  const q = norm(F.q);
  let r = DATA.props.filter(p =>
    (!F.op || p.op === F.op) && (!F.tipo || norm(p.tipo) === F.tipo) && (!F.rec || (p.beds || 0) >= F.rec) &&
    (!q || norm([p.title, p.zone, p.address, p.tipo, p.desc, p.amen.join(' ')].join(' ')).includes(q)));
  if (F.sort) r = [...r].sort((a, b) => ((a.price ?? Infinity) - (b.price ?? Infinity)) * (F.sort === 'asc' ? 1 : -1));
  return r;
}

function renderCards() {
  const r = filtered();
  $('#count').textContent = r.length === 1 ? '1 propiedad' : `${r.length} propiedades`;
  $('#cards').innerHTML = r.length ? r.map(p => `
    <button class="card ${p.off ? 'is-off' : ''}" data-open="${esc(p.id)}" data-id="${esc(p.id)}">
      <div class="card__img">${img(cover(p), p.title)}
        <div class="card__tags"><span class="tag">${p.op === 'renta' ? 'Renta' : 'Venta'}</span>${p.estatus && norm(p.estatus) !== 'disponible' ? `<span class="tag tag--dark">${esc(p.estatus)}</span>` : ''}</div>
      </div>
      <div class="card__body">
        <p class="card__zone">${esc([p.tipo, p.zone].filter(Boolean).join(' · '))}</p>
        <h3 class="card__title">${esc(p.title)}</h3>
        <p class="card__price">${priceLabel(p)}</p>
        <p class="specs">${specsHtml(p)}</p>
      </div>
    </button>`).join('') : `<p class="empty">${DATA.props.length ? 'No hay propiedades con esos filtros. Prueba quitando alguno.' : 'Muy pronto publicaremos aquí nuestras propiedades.'}</p>`;
  syncMarkers(r);
}

function renderBuilds() {
  $('#desarrollos').hidden = !DATA.builds.length;
  $('#builds').innerHTML = DATA.builds.map(b => `
    <button class="build" data-open="${esc(b.id)}" data-reveal>
      <div class="build__img">${img(cover(b), b.title)}<span class="tag tag--accent">${esc(b.etapa)}</span></div>
      <div class="build__info">
        <p class="build__zone">${esc(b.zone)}</p>
        <h3>${esc(b.title)}</h3>
        <p>${esc(b.desc.split('\n')[0])}</p>
        ${progressHtml(b)}
        <dl class="facts">
          ${b.price != null ? `<div><dt>Desde</dt><dd>${money(b.price, b.cur)}</dd></div>` : ''}
          ${b.units ? `<div><dt>Unidades</dt><dd>${b.units}</dd></div>` : ''}
          ${b.entrega ? `<div><dt>Entrega</dt><dd>${esc(b.entrega)}</dd></div>` : ''}
        </dl>
        <span class="link">Ver desarrollo →</span>
      </div>
    </button>`).join('');
  observeReveals();
}
const progressHtml = b => b.avance == null ? '' : `
  <div class="progress"><div class="progress__top"><span>Avance de obra</span><span>${Math.round(b.avance)}%</span></div>
  <div class="progress__bar"><i data-w="${Math.max(0, Math.min(100, b.avance))}"></i></div></div>`;

/* ---------- mapa de propiedades ---------- */
function initMap() {
  if (map || !window.maplibregl) return;
  try {
    map = new maplibregl.Map({ container: 'map', style: STREET_STYLE, center: TIJUANA, zoom: 11.3, attributionControl: { compact: true }, cooperativeGestures: true,
      locale: { 'CooperativeGesturesHandler.WindowsHelpText': 'Usa Ctrl + scroll para acercar el mapa', 'CooperativeGesturesHandler.MacHelpText': 'Usa ⌘ + scroll para acercar el mapa', 'CooperativeGesturesHandler.MobileHelpText': 'Usa dos dedos para mover el mapa' } });
  } catch (e) { $('#map').innerHTML = '<p style="padding:24px">No se pudo cargar el mapa en este navegador.</p>'; return; }
  map.addControl(new maplibregl.NavigationControl({ showCompass: false }), 'bottom-right');
  popup = new maplibregl.Popup({ offset: 22, maxWidth: 'none' });
  syncMarkers(filtered(), true);
}
function syncMarkers(props, fit) {
  if (!map) return;
  markers.forEach(m => m.remove()); markers.clear();
  const items = [...props, ...DATA.builds].filter(p => p.loc);
  items.forEach(p => {
    const el = document.createElement('button');
    el.className = 'pin' + (p.kind === 'build' ? ' pin--build' : '') + (p.off ? ' pin--off' : '');
    el.textContent = p.kind === 'build' ? p.title : moneyShort(p.price);
    el.setAttribute('aria-label', p.title);
    el.addEventListener('click', e => { e.stopPropagation(); showPopup(p); });
    el.addEventListener('mouseenter', () => hot(p.id, true)); el.addEventListener('mouseleave', () => hot(p.id, false));
    markers.set(p.id, new maplibregl.Marker({ element: el, anchor: 'bottom' }).setLngLat(p.loc).addTo(map));
  });
  if (items.length) {
    const b = new maplibregl.LngLatBounds(); items.forEach(p => b.extend(p.loc));
    map.fitBounds(b, { padding: { top: 80, bottom: 60, left: 60, right: 60 }, maxZoom: 14.5, duration: fit ? 0 : 900 });
  }
}
function showPopup(p) {
  const el = document.createElement('button'); el.className = 'pop'; el.dataset.open = p.id;
  el.innerHTML = `${img(cover(p), p.title, true)}<div><small>${esc(p.kind === 'build' ? p.etapa : [p.tipo, p.zone].filter(Boolean).join(' · '))}</small><span>${esc(p.title)}</span><b>${priceLabel(p)}</b></div>`;
  popup.setLngLat(p.loc).setDOMContent(el).addTo(map);
  map.easeTo({ center: p.loc, offset: [0, 110], duration: 600 });
}
function hot(id, on) {
  markers.get(id)?.getElement().classList.toggle('is-hot', on);
  $$('.card').find(c => c.dataset.id === id)?.classList.toggle('is-hot', on);
}

/* ---------- ficha de detalle ---------- */
const sheet = $('#sheet');
function openSheet(id, push = true) {
  const p = [...DATA.props, ...DATA.builds].find(x => x.id === id); if (!p) return;
  const imgs = p.imgs.length ? p.imgs : [PLACEHOLDER];
  const isB = p.kind === 'build';
  const facts = isB
    ? [[p.units, 'Unidades'], [p.entrega, 'Entrega'], [p.etapa, 'Etapa']]
    : [[p.beds, 'Recámaras'], [p.baths, 'Baños'], [p.park, 'Estac.'], [p.m2c && p.m2c.toLocaleString('es-MX') + ' m²', 'Construcción'], [p.m2t && p.m2t.toLocaleString('es-MX') + ' m²', 'Terreno']];
  $('#sheetBody').innerHTML = `
    <div class="gal">
      <div class="gal__main">${img(imgs[0], p.title, true)}
        ${imgs.length > 1 ? `<button class="gal__nav" data-d="-1" aria-label="Foto anterior">←</button><button class="gal__nav" data-d="1" aria-label="Foto siguiente">→</button><span class="gal__n">1 / ${imgs.length}</span>` : ''}
      </div>
      ${imgs.length > 1 ? `<div class="gal__thumbs">${imgs.map((s, i) => `<img src="${esc(s)}" alt="" data-i="${i}" class="${i ? '' : 'is-on'}" onerror="this.remove()">`).join('')}</div>` : ''}
    </div>
    <div class="info">
      <div class="info__tags">
        <span class="tag tag--accent">${esc(isB ? p.etapa : (p.op === 'renta' ? 'En renta' : 'En venta'))}</span>
        ${!isB ? `<span class="tag tag--dark">${esc(p.tipo)}</span>` : ''}
        ${!isB && norm(p.estatus) !== 'disponible' ? `<span class="tag tag--dark">${esc(p.estatus)}</span>` : ''}
      </div>
      <p class="info__zone">${esc([p.zone, p.address].filter(Boolean).join(' · '))}</p>
      <h3>${esc(p.title)}</h3>
      <p class="info__price">${priceLabel(p)}</p>
      ${isB ? progressHtml(p) : ''}
      <div class="grid4">${facts.filter(f => f[0]).map(f => `<div><b>${esc(f[0])}</b><small>${f[1]}</small></div>`).join('')}</div>
      ${p.desc ? `<p class="info__desc">${esc(p.desc)}</p>` : ''}
      ${p.amen.length ? `<ul class="chips">${p.amen.map(a => `<li>${esc(a)}</li>`).join('')}</ul>` : ''}
      ${p.loc ? '<div class="info__map" id="miniMap"></div>' : ''}
      <div class="info__cta">
        <button class="btn btn--accent" data-ask="${esc(p.id)}">${isB ? 'Solicitar información' : 'Agendar visita'}</button>
        ${p.video ? `<a class="btn btn--ghost" href="${esc(p.video)}" target="_blank" rel="noopener">Ver video</a>` : ''}
        ${p.brochure ? `<a class="btn btn--ghost" href="${esc(p.brochure)}" target="_blank" rel="noopener">Brochure</a>` : ''}
        <button class="btn btn--ghost" data-share="${esc(p.id)}">Compartir</button>
      </div>
      <p style="font-size:.74rem;color:var(--mute)">Clave: ${esc(p.id)}</p>
    </div>`;
  let cur = 0;
  const go = i => {
    cur = (i + imgs.length) % imgs.length;
    $('.gal__main img', sheet).src = imgs[cur];
    const n = $('.gal__n', sheet); if (n) n.textContent = `${cur + 1} / ${imgs.length}`;
    $$('.gal__thumbs img', sheet).forEach(t => t.classList.toggle('is-on', +t.dataset.i === cur));
  };
  sheet.onclick = e => {
    const nav = e.target.closest('.gal__nav'); if (nav) return go(cur + +nav.dataset.d);
    const th = e.target.closest('.gal__thumbs img'); if (th) return go(+th.dataset.i);
    if (e.target === sheet) sheet.close();
  };
  sheet.onkeydown = e => { if (e.key === 'ArrowRight') go(cur + 1); if (e.key === 'ArrowLeft') go(cur - 1); };
  if (!sheet.open) sheet.showModal();
  $$('.progress__bar i', sheet).forEach(i => requestAnimationFrame(() => i.style.width = i.dataset.w + '%'));
  if (push) history.replaceState(null, '', '#p=' + encodeURIComponent(p.id));
  if (p.loc && window.maplibregl) {
    try {
      miniMap?.remove();
      miniMap = new maplibregl.Map({ container: 'miniMap', style: STREET_STYLE, center: p.loc, zoom: 14.2, interactive: false, attributionControl: { compact: true } });
      const dot = document.createElement('div'); dot.className = 'pin'; dot.textContent = isB ? p.title : moneyShort(p.price);
      new maplibregl.Marker({ element: dot, anchor: 'bottom' }).setLngLat(p.loc).addTo(miniMap);
    } catch (e) { /* sin WebGL: se omite el mini mapa */ }
  }
}
sheet.addEventListener('close', () => { miniMap?.remove(); miniMap = null; if (location.hash.startsWith('#p=')) history.replaceState(null, '', location.pathname + location.search); });
$('#sheetClose').addEventListener('click', () => sheet.close());

/* ---------- contacto ---------- */
function toast(msg) { const t = $('#toast'); t.textContent = msg; t.classList.add('is-on'); clearTimeout(toast.t); toast.t = setTimeout(() => t.classList.remove('is-on'), 4200); }
function sendMessage(text) {
  const wa = (DATA.cfg.whatsapp || '').replace(/\D/g, ''), mail = DATA.cfg.correo;
  if (wa) { window.open(`https://wa.me/${wa}?text=${encodeURIComponent(text)}`, '_blank', 'noopener'); return true; }
  if (mail) { location.href = `mailto:${mail}?subject=${encodeURIComponent('Contacto desde el sitio web')}&body=${encodeURIComponent(text)}`; return true; }
  return false;
}
function ask(id) {
  const p = [...DATA.props, ...DATA.builds].find(x => x.id === id); if (!p) return;
  const text = `Hola, me interesa "${p.title}" (clave ${p.id}). ¿Me pueden dar más información?`;
  if (!sendMessage(text)) { sheet.close(); $('#cfMsg').value = text; $('#contacto').scrollIntoView(); $('#contactForm [name=nombre]').focus({ preventScroll: true }); }
}
$('#contactForm').addEventListener('submit', e => {
  e.preventDefault();
  const d = new FormData(e.target);
  const text = `Hola, soy ${d.get('nombre')}.${d.get('telefono') ? ' Mi teléfono es ' + d.get('telefono') + '.' : ''}\nMe interesa: ${d.get('interes')}.\n${d.get('mensaje') || ''}`.trim();
  $('#cfNote').textContent = sendMessage(text)
    ? 'Se abrió tu aplicación de mensajes para enviar la solicitud.'
    : 'Falta configurar el WhatsApp o el correo en la hoja «Configuración» del Excel.';
});

/* ---------- eventos globales ---------- */
document.addEventListener('click', e => {
  const o = e.target.closest('[data-open]'); if (o) return openSheet(o.dataset.open);
  const a = e.target.closest('[data-ask]'); if (a) return ask(a.dataset.ask);
  const s = e.target.closest('[data-share]');
  if (s) {
    const url = location.origin + location.pathname + '#p=' + encodeURIComponent(s.dataset.share);
    if (navigator.share) navigator.share({ title: document.title, url }).catch(() => {});
    else navigator.clipboard?.writeText(url).then(() => toast('Link copiado al portapapeles'), () => toast(url));
  }
});
document.addEventListener('mouseover', e => { const c = e.target.closest('.card'); if (c) hot(c.dataset.id, true); });
document.addEventListener('mouseout', e => { const c = e.target.closest('.card'); if (c) hot(c.dataset.id, false); });

function setSeg(root, v) { $$('button', root).forEach(b => b.classList.toggle('is-on', (b.dataset.v ?? b.dataset.op) === v)); }
$('#fOp').addEventListener('click', e => { const b = e.target.closest('button'); if (!b) return; F.op = b.dataset.v; setSeg($('#fOp'), F.op); renderCards(); });
$('#fTipo').addEventListener('change', e => { F.tipo = e.target.value; renderCards(); });
$('#fRec').addEventListener('change', e => { F.rec = +e.target.value; renderCards(); });
$('#fSort').addEventListener('change', e => { F.sort = e.target.value; renderCards(); });
$('#fQ').addEventListener('input', e => { F.q = e.target.value; renderCards(); });
let heroOp = '';
$('.search__tabs').addEventListener('click', e => { const b = e.target.closest('button'); if (!b) return; heroOp = b.dataset.op; setSeg($('.search__tabs'), heroOp); });
$('#heroSearch').addEventListener('submit', e => {
  e.preventDefault();
  F.op = heroOp; F.tipo = $('#hsTipo').value; F.q = $('#hsZona').value;
  setSeg($('#fOp'), F.op); $('#fTipo').value = F.tipo; $('#fQ').value = F.q;
  renderCards(); $('#mapa').scrollIntoView();
});
$('#mapToggle').addEventListener('click', e => {
  const b = e.target.closest('button'); if (!b || !map) return;
  $$('#mapToggle button').forEach(x => x.classList.toggle('is-on', x === b));
  map.setStyle(b.dataset.s === 'sat' ? SAT_STYLE : STREET_STYLE);
});
$('#railPrev').addEventListener('click', () => $('#rail').scrollBy({ left: -600, behavior: 'smooth' }));
$('#railNext').addEventListener('click', () => $('#rail').scrollBy({ left: 600, behavior: 'smooth' }));

const nav = $('#nav'), burger = $('#burger'), links = $('#navLinks');
const onScroll = () => nav.classList.toggle('is-solid', scrollY > innerHeight * 0.7 || links.classList.contains('is-open'));
addEventListener('scroll', onScroll, { passive: true });
burger.addEventListener('click', () => { const o = links.classList.toggle('is-open'); burger.setAttribute('aria-expanded', o); onScroll(); });
links.addEventListener('click', e => { if (e.target.closest('a')) { links.classList.remove('is-open'); burger.setAttribute('aria-expanded', false); onScroll(); } });
$('#year').textContent = new Date().getFullYear();

let io;
function observeReveals() {
  io ??= new IntersectionObserver(es => es.forEach(en => {
    if (!en.isIntersecting) return;
    en.target.classList.add('is-in'); io.unobserve(en.target);
    $$('.progress__bar i', en.target).forEach(i => i.style.width = i.dataset.w + '%');
  }), { threshold: 0.15 });
  $$('[data-reveal]:not(.is-in)').forEach(el => io.observe(el));
}

/* ---------- vista previa arrastrando el Excel ---------- */
const dz = $('#dropzone'); let dragN = 0;
const hasFile = e => [...(e.dataTransfer?.types || [])].includes('Files');
addEventListener('dragenter', e => { if (hasFile(e)) { dragN++; dz.classList.add('is-on'); } });
addEventListener('dragleave', () => { if (--dragN <= 0) { dragN = 0; dz.classList.remove('is-on'); } });
addEventListener('dragover', e => { if (hasFile(e)) e.preventDefault(); });
addEventListener('drop', async e => {
  if (!hasFile(e)) return; e.preventDefault(); dragN = 0; dz.classList.remove('is-on');
  const f = e.dataTransfer.files[0]; if (!f || !/\.xlsx?$/i.test(f.name)) return toast('Arrastra un archivo de Excel (.xlsx)');
  try { DATA = parseWorkbook(await f.arrayBuffer()); renderAll(); toast(`Vista previa: ${DATA.props.length} propiedades y ${DATA.builds.length} desarrollos. Solo tú la ves hasta subir el archivo.`); }
  catch (err) { toast('No se pudo leer el Excel.'); }
});

function renderAll() {
  applyConfig(); fillOptions(); renderFeatured(); renderBuilds(); renderCards(); if (map) syncMarkers(filtered());
  // los enlaces del menú a secciones vacías se ocultan
  $$('#navLinks a').forEach(l => { const t = $(l.getAttribute('href')); l.hidden = !!(t && t.hidden); });
}

/* ---------- portada por escenas ---------- */
const SCENES = [
  { eyebrow: 'Tijuana, Baja California', l1: 'Conoce Tijuana', l2: 'desde <em>otro nivel.</em>', ms: 13000 },
  { eyebrow: 'Torres y desarrollos', l1: 'El nuevo horizonte', l2: 'de <em>Tijuana.</em>', ms: 8000, credit: 'Foto: Urbaner44 / Wikimedia Commons · CC BY-SA 4.0' },
  { eyebrow: 'Bulevar Agua Caliente', l1: 'Invierte donde', l2: 'la ciudad <em>crece.</em>', ms: 8000, credit: 'Foto: Urbaner44 / Wikimedia Commons · CC BY-SA 4.0' }
];
let customHero = false, sceneIdx = 0, sceneTimer, scenesOn = false, heroInView = true, heroMapCtl = () => {}, swapT;
function scheduleScene() {
  clearTimeout(sceneTimer);
  const dots = $('#heroDots'); dots.classList.remove('is-run');
  if (!scenesOn || reduced || !heroInView || document.hidden) return;
  dots.style.setProperty('--d', SCENES[sceneIdx].ms + 'ms');
  void dots.offsetWidth; dots.classList.add('is-run');
  sceneTimer = setTimeout(() => setScene((sceneIdx + 1) % SCENES.length), SCENES[sceneIdx].ms);
}
function setScene(i) {
  if (i === sceneIdx) return;
  const els = $$('#heroScenes .scene'), prev = els[sceneIdx], next = els[i], sc = SCENES[i];
  els.forEach(e => e.classList.remove('is-prev'));
  prev.classList.remove('is-on'); prev.classList.add('is-prev'); next.classList.add('is-on');
  setTimeout(() => { if (!prev.classList.contains('is-on')) prev.classList.remove('is-prev'); }, 1350);
  sceneIdx = i;
  $$('#heroDots button').forEach(b => b.classList.toggle('is-on', +b.dataset.i === i));
  $('#heroAlt').style.visibility = i === 0 ? '' : 'hidden';
  const cr = $('#heroCredit'); cr.textContent = (!customHero && sc.credit) || ''; cr.hidden = !cr.textContent;
  const title = $('#heroTitle'), spans = $$('.line>span', title);
  clearTimeout(swapT); title.classList.remove('is-pre'); title.classList.add('is-out');
  swapT = setTimeout(() => {
    spans[0].innerHTML = sc.l1; spans[1].innerHTML = sc.l2; $('#heroEyebrow').textContent = sc.eyebrow;
    title.classList.remove('is-out'); title.classList.add('is-pre'); void title.offsetWidth; title.classList.remove('is-pre');
  }, reduced ? 0 : 520);
  heroMapCtl(); scheduleScene();
}
$('#heroDots').addEventListener('click', e => { const b = e.target.closest('button'); if (b) setScene(+b.dataset.i); });
document.addEventListener('visibilitychange', scheduleScene);
new IntersectionObserver(([en]) => { heroInView = en.isIntersecting; heroMapCtl(); scheduleScene(); }, { threshold: 0.25 }).observe($('#inicio'));

/* ---------- intro: el logo se abre y la cámara baja sobre Tijuana ---------- */
function intro() {
  const introEl = $('#intro'), bar = $('.intro__bar i'), alt = $('#heroAlt'), cut = $('#introCut'), fill = $('#introFill');
  let hero, done = false, started = false, flown = false, orbiting = false, t0 = performance.now();
  const MIN = reduced ? 300 : 2300, MAX = 6500;
  const altitude = z => Math.round(40075016 * Math.cos(32.5 * Math.PI / 180) / Math.pow(2, z + 1) * 1.2 / 10) * 10;   // aprox. en metros
  const showAlt = () => { if (hero) alt.textContent = 'Altitud ' + altitude(hero.getZoom()).toLocaleString('es-MX') + ' m'; };

  const canOrbit = () => hero && started && !reduced && heroInView && sceneIdx === 0;
  const orbit = () => { if (!canOrbit() || orbiting) return; orbiting = true; hero.easeTo({ bearing: hero.getBearing() + 90, duration: 150000, easing: t => t, essential: true }); };
  heroMapCtl = () => {
    if (!hero || !started) return;
    if (canOrbit()) return orbit();
    if (!flown) { flown = true; hero.jumpTo(HERO_END); } else hero.stop();
    orbiting = false;
  };
  const fly = () => {
    started = true; if (!hero) return;
    if (reduced) { hero.jumpTo(HERO_END); showAlt(); return; }
    hero.flyTo({ ...HERO_END, duration: 9500, curve: 1.1, essential: true, easing: t => 1 - Math.pow(1 - t, 3) });
  };
  const ready = () => { document.body.classList.remove('is-loading'); document.body.classList.add('is-ready'); onScroll(); };
  const finish = () => { introEl.classList.add('is-done'); scenesOn = true; scheduleScene(); };

  const reveal = () => {
    if (done) return; done = true; bar.style.width = '100%';
    let ox, oy;
    try { const r = cut.getExtentOfChar(1); ox = r.x + r.width / 2; oy = r.y + r.height * 0.42; } catch (e) { ox = null; }
    if (reduced || ox == null || !isFinite(ox)) { ready(); fly(); finish(); return; }
    // 1) las letras se vuelven ventana  2) la cámara atraviesa la «V»
    introEl.classList.add('is-open');
    const A = 520, B = 1650, s0 = performance.now(); let flew = false, shown = false;
    const step = now => {
      const t = now - s0, k = Math.min(1, t / A);
      cut.setAttribute('fill-opacity', k); fill.style.opacity = 1 - k;
      if (t >= A) {
        if (!flew) { flew = true; fly(); }
        const z = Math.min(1, (t - A) / B), sc = Math.pow(160, z * z);
        cut.setAttribute('transform', `translate(${ox} ${oy}) scale(${sc}) translate(${-ox} ${-oy})`);
        if (z > 0.5 && !shown) { shown = true; ready(); }
        if (z > 0.8) introEl.style.opacity = 1 - (z - 0.8) / 0.2;
        if (z >= 1) return finish();
      }
      requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  };
  const tryReveal = () => setTimeout(reveal, Math.max(0, MIN - (performance.now() - t0)));

  // Toma de dron propia: se activa con «Video de portada» en la hoja Configuración del Excel.
  useHeroVideo = src => {
    const video = $('#heroVideo');
    video.addEventListener('loadeddata', () => { video.hidden = false; $('#heroMap').hidden = true; alt.hidden = true; hero?.remove(); hero = null; }, { once: true });
    video.src = src; video.play().catch(() => {});
  };

  try {
    hero = new maplibregl.Map({ container: 'heroMap', style: SAT_STYLE, center: [-117.09, 32.47], zoom: 10.4, pitch: 0, bearing: 0,
      interactive: false, attributionControl: { compact: false }, maxPitch: 70, fadeDuration: 400 });
    hero.on('move', showAlt); showAlt();
    hero.once('idle', tryReveal);
    hero.on('moveend', () => { if (!started) return; flown = true; orbiting = false; orbit(); });
  } catch (e) { hero = null; alt.hidden = true; tryReveal(); }

  const tick = setInterval(() => { const p = Math.min(92, (performance.now() - t0) / MAX * 100); if (!done) bar.style.width = p + '%'; else clearInterval(tick); }, 120);
  setTimeout(reveal, MAX);
  introEl.addEventListener('click', reveal);
}

/* ---------- arranque ---------- */
async function boot() {
  intro();
  observeReveals();
  let ok = false;
  if (FUENTE) {
    const use = j => { DATA = parseSheets(j.hojas, j.fotos || {}); ok = true; };
    try { const c = JSON.parse(localStorage.getItem('avk-datos')); if (c && c.hojas) { use(c); renderAll(); } } catch (e) {}
    try {
      const j = await (await fetch(FUENTE, { redirect: 'follow' })).json();
      if (j && j.hojas) { use(j); try { localStorage.setItem('avk-datos', JSON.stringify(j)); } catch (e) {} }
    } catch (e) {}
  }
  if (!ok) try {
    const r = await fetch(EXCEL + '?v=' + Date.now(), { cache: 'no-store' });
    if (!r.ok) throw new Error(r.status);
    DATA = parseWorkbook(await r.arrayBuffer());
  } catch (e) {
    $('#cards').innerHTML = `<p class="empty">No se pudieron cargar las propiedades. ${location.protocol === 'file:' ? 'Abre el sitio desde un servidor o hosting (no con doble clic).' : 'Intenta recargar la página.'}</p>`;
  }
  renderAll();
  // el mapa se crea cuando la sección se acerca a la pantalla
  new IntersectionObserver((es, o) => { if (es[0].isIntersecting) { o.disconnect(); initMap(); } }, { rootMargin: '600px' }).observe($('#mapa'));
  const m = location.hash.match(/^#p=(.+)/); if (m) openSheet(decodeURIComponent(m[1]), false);
}
boot();
})();
