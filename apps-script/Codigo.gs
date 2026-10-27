/**
 * E33 Estudio — Apps Script para la hoja "E33 Estudio — Control landing"
 * Pegar en: Extensiones > Apps Script > Código.gs. Luego ejecutar instalar() una vez.
 *
 * Qué hace
 *  - doGet  ?action=datos  → JSON con tarifas, producción, equipos, rider, media, agenda ocupada y TRM (lo lee la landing).
 *  - doPost action=solicitud → registra la solicitud, crea un evento TENTATIVO en el calendario y avisa por correo.
 *  - actualizarTRM()     → trae la TRM oficial (datos.gov.co, Superfinanciera) todos los días a las 06:00.
 *  - sincronizarMedia()  → lee la carpeta de Drive de fotos y videos y actualiza la pestaña Media.
 *  - alCambiarEstado()   → al cambiar "estado" en Solicitudes: Confirmada / Cancelada / Realizada actualiza el calendario.
 *  - resumenDiario()     → correo a las 19:00 con las sesiones del día siguiente.
 */

const TZ = 'America/Bogota';
const FRANJAS = [ { id:'m', label:'10:00–14:00', ini:10, fin:14 },
                  { id:'t', label:'14:00–18:00', ini:14, fin:18 },
                  { id:'n', label:'18:00–22:00', ini:18, fin:22 } ];
const ESTADOS = ['Nueva', 'Tentativa', 'Anticipo recibido', 'Confirmada', 'Realizada', 'Cancelada'];

/* ---------- Menú ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('E33 Estudio')
    .addItem('Actualizar TRM ahora', 'actualizarTRM')
    .addItem('Sincronizar fotos y videos', 'sincronizarMedia')
    .addItem('Limpiar caché de la landing', 'limpiarCache')
    .addSeparator()
    .addItem('Instalar automatizaciones', 'instalar')
    .addToUi();
}

/* ---------- Utilidades ---------- */
function ss_() { return SpreadsheetApp.getActive(); }
function filas_(nombre) {
  const v = ss_().getSheetByName(nombre).getDataRange().getValues();
  const h = v.shift();
  return v.filter(r => r.join('') !== '').map(r => Object.fromEntries(h.map((k, i) => [k, r[i]])));
}
function config_() {
  const c = {}; filas_('Config').forEach(r => c[r.clave] = r.valor); return c;
}
function setConfig_(clave, valor) {
  const sh = ss_().getSheetByName('Config'), v = sh.getRange('A:A').getValues();
  const i = v.findIndex(r => r[0] === clave);
  if (i >= 0) sh.getRange(i + 1, 2).setValue(valor); else sh.appendRow([clave, valor, '']);
}
function activo_(r) { return r.activo === true || String(r.activo).toUpperCase() === 'TRUE' || r.activo === ''; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function limpiarCache() { CacheService.getScriptCache().remove('datos'); }

/* ---------- API para la landing ---------- */
function doGet(e) {
  if (e && e.parameter && e.parameter.app === 'admin') {
    return HtmlService.createHtmlOutputFromFile('Admin').setTitle('E33 · Gestión del estudio')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
  }
  const cache = CacheService.getScriptCache();
  const hit = cache.get('datos');
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  const out = JSON.stringify(datos_());
  cache.put('datos', out, 300); // 5 min
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

function datos_() {
  const c = config_();
  const tarifas = {};
  filas_('Tarifas').filter(activo_).forEach(r => {
    tarifas[r.id] = { label:r.label, label_en:r.label_en || '', horas:+r.horas, min:+r.min || undefined, max:+r.max || undefined,
      sin:+r.sin_ingeniero_usd, con:+r.con_ingeniero_usd,
      extraSin:+r.extra_sin_usd || undefined, extraCon:+r.extra_con_usd || undefined, desc:r.desc, desc_en:r.desc_en || '' };
  });
  const produccion = filas_('Produccion').filter(activo_).map(r => ({ id:r.id, label:r.label, label_en:r.label_en || '', usd:+r.usd }));
  const equipos = filas_('Equipos')
    .filter(r => (r.en_configurador === true || String(r.en_configurador).toUpperCase() === 'TRUE') && ['Nuevo', 'Operativo'].includes(r.estado))
    .map(r => ({ id:r.id, cat:r.cat, cat_en:r.cat_en || '', short:r.short || '', name:r.name, name_en:r.name_en || '', spec:r.spec, chain:r.chain, qty:+r.qty || 1 }));
  const grupos = {}, gEn = {};
  filas_('Rider').sort((a, b) => a.orden - b.orden).forEach(r => {
    (grupos[r.grupo] = grupos[r.grupo] || []).push([r.item, r.detalle, r.item_en || '', r.detalle_en || '']);
    if (r.grupo_en) gEn[r.grupo] = r.grupo_en;
  });
  const rider = Object.entries(grupos).map(([g, items]) => ({ g, g_en:gEn[g] || '', items }));
  const media = filas_('Media').filter(r => activo_(r) && r.src).sort((a, b) => a.orden - b.orden)
    .map(r => ({ k:r.k, cap:r.cap, cap_en:r.cap_en || '', tipo:r.tipo || 'foto', src:r.src }));
  return { tarifas, produccion, equipos, rider, media, nocturno:+c.nocturno || 0.25,
           capacidad:{ comoda:+c.capacidad_comoda || 6, max:+c.capacidad_max || 10 },
           ocupado: ocupado_(c.calendar_id), trm:+c.trm, trmFecha:fmtFecha_(c.trm_fecha),
           whatsapp:String(c.whatsapp || ''), email:String(c.email || '') };
}

function fmtFecha_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'dd/MM/yyyy') : String(v || ''); }

/** Devuelve ["AAAA-MM-DD|m", ...] para las franjas ocupadas en los próximos 63 días. */
function ocupado_(calId) {
  if (!calId) return [];
  const cal = CalendarApp.getCalendarById(calId); if (!cal) return [];
  const desde = new Date(), hasta = new Date(Date.now() + 63 * 864e5), out = new Set();
  cal.getEvents(desde, hasta).forEach(ev => {
    if (/^CANCELADA/i.test(ev.getTitle())) return;
    for (let d = new Date(ev.getStartTime()); d < ev.getEndTime(); d = new Date(d.getTime() + 36e5)) {
      const h = +Utilities.formatDate(d, TZ, 'H'), dia = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
      FRANJAS.forEach(f => { if (h >= f.ini && h < f.fin) out.add(dia + '|' + f.id); });
    }
  });
  return [...out];
}

/* ---------- Solicitudes desde la landing ---------- */
function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action !== 'solicitud') return json_({ ok:false, error:'acción desconocida' });
    if (!b.name || !b.phone || !b.mail) return json_({ ok:false, error:'faltan datos' });
    const c = config_(), lock = LockService.getScriptLock(); lock.waitLock(10000);
    const trm = +c.trm || +b.trm, cop = Math.round(b.totalUSD * trm / 1000) * 1000;
    let evId = '';
    if (c.calendar_id && b.fecha) evId = crearEvento_(c.calendar_id, b);
    ss_().getSheetByName('Solicitudes').appendRow([
      new Date(), b.id, 'Nueva', b.name, b.phone, b.mail, b.modo, b.ingeniero ? 'Sí' : 'No', b.nocturno ? 'Sí' : 'No',
      b.horas, b.fecha, b.franja, (b.equipos || []).join(', '), JSON.stringify(b.produccion || {}),
      b.totalUSD, trm, cop, b.notes || '', evId, b.resumen || '', b.rider || '', b.faltante || '',
      b.personas || '', JSON.stringify(b.crew || {}), b.lang || 'es' ]);
    lock.releaseLock();
    limpiarCache();
    avisar_(c, 'Nueva solicitud ' + b.id + ' · ' + b.name + (b.faltante === 'Cotizarlo con E33' ? ' · cotizar equipo' : ''),
      (b.resumen || '') + '\n\nCOP ref: $' + cop.toLocaleString('es-CO') + '\n\nResponder en máximo 24 h hábiles.');
    acuse_(c, b);
    return json_({ ok:true, id:b.id });
  } catch (err) {
    return json_({ ok:false, error:String(err) });
  }
}

function crearEvento_(calId, b) {
  const cal = CalendarApp.getCalendarById(calId); if (!cal) return '';
  const [y, m, d] = b.fecha.split('-').map(Number);
  let ini, fin;
  if (b.franja === 'dia' || b.modo === 'jornada') { ini = new Date(y, m - 1, d, 10); fin = new Date(y, m - 1, d, 18); }
  else { const f = FRANJAS.find(x => x.id === b.franja) || FRANJAS[0]; ini = new Date(y, m - 1, d, f.ini); fin = new Date(y, m - 1, d, f.fin + (+b.horas || 0) * (b.modo === 'hora' ? 0 : 1)); }
  if (b.modo === 'hora') fin = new Date(ini.getTime() + (+b.horas || 2) * 36e5);
  const ev = cal.createEvent('TENTATIVO · ' + b.name + ' · ' + b.id, ini, fin,
    { description: (b.resumen || '') + '\n\nWhatsApp: ' + b.phone + '\nCorreo: ' + b.mail });
  ev.setColor(CalendarApp.EventColor.YELLOW);
  return ev.getId();
}

function avisar_(c, asunto, cuerpo) {
  const to = String(c.avisar_a || '').trim(); if (!to) return;
  MailApp.sendEmail({ to, subject: asunto, body: cuerpo, name: 'Estudio E33' });
}

/** Correo automático al artista: confirma que la solicitud llegó y que debe esperar respuesta. */
function acuse_(c, b) {
  if (!b.mail) return;
  if (b.lang === 'en') {
    const body = 'Hi ' + b.name + ',\n\n' +
      'We received your request ' + b.id + ' for Estudio E33.\n\n' +
      'Your slot is on hold and not booked yet. We will review the schedule, engineer, space and the rider you sent, ' +
      'and reply by WhatsApp and email within 24 business hours with the final price and the details for the 50% deposit. ' +
      'Your booking is confirmed once we receive the deposit.\n\n' +
      'Request summary:\n' + (b.resumen || '') + '\n\nEstudio E33 · Entr3t3nimiento';
    const o = { to: b.mail, subject: 'We received your request ' + b.id + ' · Estudio E33', body, name: 'Estudio E33' };
    if (c.email) o.replyTo = String(c.email);
    return MailApp.sendEmail(o);
  }
  const cuerpo = 'Hola ' + b.name + ',\n\n' +
    'Recibimos tu solicitud ' + b.id + ' para el Estudio E33.\n\n' +
    'Tu franja queda en espera: todavía no está reservada. Vamos a revisar agenda, ingeniero y el rider que nos enviaste, ' +
    'y te respondemos por WhatsApp y correo en máximo 24 horas hábiles con el valor final y los datos para el anticipo del 50 %. ' +
    'La reserva queda confirmada cuando recibimos el anticipo.\n\n' +
    'Resumen de tu solicitud:\n' + (b.resumen || '') + '\n\nEstudio E33 · Entr3t3nimiento';
  const opts = { to: b.mail, subject: 'Recibimos tu solicitud ' + b.id + ' · Estudio E33', body: cuerpo, name: 'Estudio E33' };
  if (c.email) opts.replyTo = String(c.email);
  MailApp.sendEmail(opts);
}

/* ---------- Cambios de estado en Solicitudes (disparador instalable) ---------- */
function alCambiarEstado(e) {
  const sh = e.range.getSheet(); if (sh.getName() !== 'Solicitudes' || e.range.getColumn() !== 3 || e.range.getRow() < 2) return;
  const fila = sh.getRange(e.range.getRow(), 1, 1, 25).getValues()[0];
  const estado = e.value, evId = fila[18], artista = fila[3], id = fila[1], c = config_();
  if (!evId || !c.calendar_id) return;
  const ev = CalendarApp.getCalendarById(c.calendar_id).getEventById(evId); if (!ev) return;
  if (estado === 'Confirmada' || estado === 'Anticipo recibido') { ev.setTitle('CONFIRMADA · ' + artista + ' · ' + id); ev.setColor(CalendarApp.EventColor.RED); }
  if (estado === 'Tentativa') { ev.setTitle('TENTATIVO · ' + artista + ' · ' + id); ev.setColor(CalendarApp.EventColor.YELLOW); }
  if (estado === 'Realizada') ev.setColor(CalendarApp.EventColor.GRAY);
  if (estado === 'Cancelada') ev.deleteEvent();
  limpiarCache();
}

/* ---------- TRM oficial ---------- */
function actualizarTRM() {
  const url = 'https://www.datos.gov.co/resource/32sa-8pi3.json?$order=vigenciadesde%20DESC&$limit=1';
  const r = JSON.parse(UrlFetchApp.fetch(url, { muteHttpExceptions:true }).getContentText())[0];
  if (!r || !r.valor) throw new Error('No se pudo leer la TRM');
  const valor = +r.valor, vig = Utilities.formatDate(new Date(r.vigenciadesde), TZ, 'dd/MM/yyyy');
  setConfig_('trm', valor); setConfig_('trm_fecha', vig);
  ss_().getSheetByName('TRM_Historial').appendRow([new Date(), valor, vig]);
  limpiarCache();
}

/* ---------- Fotos y videos desde Drive ---------- */
function sincronizarMedia() {
  const c = config_(); if (!c.drive_media_folder_id) throw new Error('Falta drive_media_folder_id en Config');
  const sh = ss_().getSheetByName('Media');
  const prev = {}; filas_('Media').forEach(r => prev[r.k] = r);
  const files = DriveApp.getFolderById(c.drive_media_folder_id).getFiles(), rows = [];
  let i = 0;
  while (files.hasNext()) {
    const f = files.next(), mt = f.getMimeType();
    if (!/^image\/|^video\//.test(mt)) continue;
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const id = f.getId(), tipo = mt.startsWith('video/') ? 'video' : 'foto';
    const src = tipo === 'foto' ? 'https://lh3.googleusercontent.com/d/' + id + '=w1600'
                                : 'https://drive.google.com/file/d/' + id + '/preview';
    const p = prev[id] || {};
    rows.push([id, p.cap || f.getName().replace(/\.[^.]+$/, ''), tipo, src, p.orden || ++i, p.activo === '' ? true : (p.activo ?? true)]);
  }
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 6).setValues(rows);
  limpiarCache();
}

/* ---------- Resumen diario ---------- */
function resumenDiario() {
  const c = config_(); if (!c.calendar_id) return;
  const man = new Date(Date.now() + 864e5); man.setHours(0, 0, 0, 0);
  const evs = CalendarApp.getCalendarById(c.calendar_id).getEvents(man, new Date(man.getTime() + 864e5));
  if (!evs.length) return;
  const txt = evs.map(e => Utilities.formatDate(e.getStartTime(), TZ, 'HH:mm') + '–' + Utilities.formatDate(e.getEndTime(), TZ, 'HH:mm') + '  ' + e.getTitle()).join('\n');
  avisar_(c, 'Sesiones de mañana · ' + Utilities.formatDate(man, TZ, 'dd/MM/yyyy'), txt + '\n\nRevisar saldo pendiente y checklist de equipos.');
}

/* ---------- Instalación ---------- */
function instalar() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('actualizarTRM').timeBased().atHour(6).everyDays(1).inTimezone(TZ).create();
  ScriptApp.newTrigger('resumenDiario').timeBased().atHour(19).everyDays(1).inTimezone(TZ).create();
  ScriptApp.newTrigger('alCambiarEstado').forSpreadsheet(ss_()).onEdit().create();
  const sh = ss_().getSheetByName('Solicitudes');
  sh.getRange('C2:C1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).build());
  actualizarTRM();
}
