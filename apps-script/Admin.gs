/**
 * E33 Estudio — App de gestión (Admin.gs)
 * Va en el MISMO proyecto de Apps Script que Codigo.gs, junto al archivo HTML "Admin".
 *
 * Roles (pestaña Usuarios):
 *   admin        → todo, incluido aprobar compras (Pepe)
 *   operaciones  → reservas, inventario, compras (solicitar y registrar), ingresos (Isis)
 *   local        → ver todo, marcar sesiones realizadas, estado de equipos, solicitar compras (Santi)
 */

const PERMISOS = {
  admin:       ['*'],
  operaciones: ['reservas', 'inventario', 'inventario.estado', 'compras.solicitar', 'compras.registrar', 'ingresos', 'reservas.sesion'],
  local:       ['reservas.sesion', 'inventario.estado', 'compras.solicitar']
};
const PREFIJOS = { 'Monitoreo':'MON', 'Micrófonos':'MIC', 'Audífonos':'AUD', 'Instrumentos':'INS', 'Cables':'CAB', 'Adaptadores':'ADP',
                   'Soportes':'SOP', 'Accesorios':'ACC', 'Energía':'ENE', 'Red':'RED', 'Mobiliario':'MOB', 'Hospitalidad':'HOS' };
const CAMPOS_LOCAL_INV = ['estado', 'ubicacion', 'ultima_revision', 'observaciones'];

/* ---------- Utilidades de hoja ---------- */
function hoja_(n) { return SpreadsheetApp.getActive().getSheetByName(n); }
function serial_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, "yyyy-MM-dd'T'HH:mm") : v; }
function leer_(n) {
  const sh = hoja_(n), v = sh.getDataRange().getValues(), h = v.shift();
  return v.map((r, i) => {
    if (r.join('') === '') return null;
    const o = { _row: i + 2 }; h.forEach((k, j) => o[k] = serial_(r[j])); return o;
  }).filter(Boolean);
}
function cabeceras_(n) { return hoja_(n).getRange(1, 1, 1, hoja_(n).getLastColumn()).getValues()[0]; }
function guardar_(n, clave, obj) {
  const sh = hoja_(n), h = cabeceras_(n), filas = leer_(n);
  const prev = filas.find(f => String(f[clave]) === String(obj[clave]));
  const fila = h.map(k => (k in obj) ? obj[k] : (prev ? prev[k] : ''));
  if (prev) sh.getRange(prev._row, 1, 1, h.length).setValues([fila]);
  else sh.appendRow(fila);
  return Object.fromEntries(h.map((k, i) => [k, serial_(fila[i])]));
}

/* ---------- Usuario y permisos ---------- */
function usuario_() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const u = leer_('Usuarios').find(r => String(r.email).toLowerCase() === email && email);
  if (!u) throw new Error('Tu cuenta (' + (email || 'sin correo') + ') no tiene acceso. Pide que la agreguen en la pestaña Usuarios.');
  return { email, nombre: u.nombre, rol: u.rol };
}
function puede_(u, accion) {
  const p = PERMISOS[u.rol] || [];
  return p.includes('*') || p.includes(accion);
}
function exigir_(u, accion) { if (!puede_(u, accion)) throw new Error('Tu rol (' + u.rol + ') no permite esta acción.'); }

/* ---------- Carga inicial ---------- */
function apiInicio() {
  const u = usuario_(), c = config_();
  const desde = new Date(), hasta = new Date(Date.now() + 45 * 864e5);
  let eventos = [];
  if (c.calendar_id) {
    const cal = CalendarApp.getCalendarById(c.calendar_id);
    if (cal) eventos = cal.getEvents(desde, hasta).map(e => ({ id: e.getId(), titulo: e.getTitle(), ini: serial_(e.getStartTime()), fin: serial_(e.getEndTime()) }));
  }
  return {
    usuario: u, permisos: PERMISOS[u.rol] || [],
    config: { trm: +c.trm, trmFecha: fmtFecha_(c.trm_fecha), capacidad: +c.capacidad_comoda || 6 },
    solicitudes: leer_('Solicitudes'), inventario: leer_('Inventario'),
    compras: leer_('Compras'), ingresos: leer_('Ingresos'), eventos, prefijos: PREFIJOS
  };
}

/* ---------- Reservas ---------- */
function apiReservaEstado(id, estado) {
  const u = usuario_();
  exigir_(u, (estado === 'Realizada') ? 'reservas.sesion' : 'reservas');
  if (!ESTADOS.includes(estado)) throw new Error('Estado no válido');
  const f = leer_('Solicitudes').find(r => r.id === id); if (!f) throw new Error('No encontré la solicitud ' + id);
  hoja_('Solicitudes').getRange(f._row, 3).setValue(estado);
  actualizarEvento_(f.evento_calendar, estado, f.artista, id);
  limpiarCache();
  return leer_('Solicitudes').find(r => r.id === id);
}
function actualizarEvento_(evId, estado, artista, id) {
  const c = config_(); if (!evId || !c.calendar_id) return;
  const ev = CalendarApp.getCalendarById(c.calendar_id).getEventById(evId); if (!ev) return;
  if (estado === 'Confirmada' || estado === 'Anticipo recibido') { ev.setTitle('CONFIRMADA · ' + artista + ' · ' + id); ev.setColor(CalendarApp.EventColor.RED); }
  if (estado === 'Tentativa' || estado === 'Nueva') { ev.setTitle('TENTATIVO · ' + artista + ' · ' + id); ev.setColor(CalendarApp.EventColor.YELLOW); }
  if (estado === 'Realizada') ev.setColor(CalendarApp.EventColor.GRAY);
  if (estado === 'Cancelada') ev.deleteEvent();
}

/* ---------- Inventario ---------- */
function apiInventarioGuardar(item) {
  const u = usuario_();
  const existe = item.codigo && leer_('Inventario').some(r => r.codigo === item.codigo);
  if (!existe || !puede_(u, 'inventario')) {
    if (!existe) exigir_(u, 'inventario');
    exigir_(u, 'inventario.estado');
    Object.keys(item).forEach(k => { if (k !== 'codigo' && !CAMPOS_LOCAL_INV.includes(k)) delete item[k]; });
  }
  if (!existe) item.codigo = siguienteCodigo_(item.categoria);
  delete item._row;
  const r = guardar_('Inventario', 'codigo', item);
  limpiarCache();
  return r;
}
function siguienteCodigo_(cat) {
  const p = PREFIJOS[cat]; if (!p) throw new Error('Categoría sin prefijo: ' + cat);
  const n = leer_('Inventario').map(r => String(r.codigo)).filter(c => c.startsWith(p + '-')).map(c => +c.split('-')[1] || 0);
  return p + '-' + String((n.length ? Math.max(...n) : 0) + 1).padStart(3, '0');
}

/* ---------- Compras ---------- */
function apiCompraGuardar(c) {
  const u = usuario_(), hoy = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  delete c._row;
  if (!c.id) {
    exigir_(u, 'compras.solicitar');
    c.id = 'C-' + Utilities.formatDate(new Date(), TZ, 'yyMMdd-HHmmss');
    c.fecha = hoy; c.estado = 'Solicitada'; c.solicitado_por = u.nombre;
    const r = guardar_('Compras', 'id', c);
    avisar_(config_(), 'Compra por aprobar · ' + c.item, u.nombre + ' solicitó: ' + c.item + '\nMonto estimado: ' + c.monto + ' ' + c.moneda + '\n' + (c.notas || ''));
    return r;
  }
  const prev = leer_('Compras').find(r => r.id === c.id); if (!prev) throw new Error('No encontré la compra');
  if (c.estado === 'Aprobada' || c.estado === 'Rechazada') { exigir_(u, 'compras.aprobar'); c.aprobado_por = u.nombre; }
  else if (c.estado === 'Comprada') { exigir_(u, 'compras.registrar'); if (prev.estado !== 'Aprobada') throw new Error('Primero debe estar aprobada'); }
  else exigir_(u, 'compras.registrar');
  return guardar_('Compras', 'id', c);
}

/* ---------- Ingresos ---------- */
function apiIngresoGuardar(i) {
  const u = usuario_(); exigir_(u, 'ingresos');
  delete i._row;
  if (!i.id) { i.id = 'I-' + Utilities.formatDate(new Date(), TZ, 'yyMMdd-HHmmss'); i.registrado_por = u.nombre; }
  const r = guardar_('Ingresos', 'id', i);
  if (i.reserva_id && i.es_anticipo) {
    const f = leer_('Solicitudes').find(s => s.id === i.reserva_id);
    if (f && ['Nueva', 'Tentativa'].includes(f.estado)) apiReservaEstado(i.reserva_id, 'Anticipo recibido');
  }
  return r;
}
