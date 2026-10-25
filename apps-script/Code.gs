/**
 * CONSTRUBLOCK · Servidor de la app (Google Apps Script)
 * Va pegado en la hoja CONSTRUBLOCK 1.0 → Extensiones → Apps Script.
 * Archivos del proyecto: Code.gs (este) e Index.html (la app).
 *
 * La app dentro de Google (/exec) usa google.script.run.
 * La app instalable (GitHub Pages) manda {fn, args} por doPost y recibe {ok, r}.
 */

const CONFIG = {
  NEGOCIO: 'Grupo Construblock',
  ZONA: 'America/Mazatlan',
  CARPETA: 'CONSTRUBLOCK · Archivos',
  SESION_SEG: 21600, // 6 horas sin usar la app y pide PIN otra vez
};

const HOJAS = {
  EMPRESA: ['CLAVE', 'VALOR'],
  USUARIOS_APP: ['NOMBRE', 'PIN', 'ROL', 'ACTIVO'],
  CLIENTES: ['ID_CLIENTE', 'NOMBRE', 'TELEFONO', 'CORREO', 'RFC', 'RAZON_SOCIAL', 'REGIMEN', 'USO_CFDI', 'CP', 'DIRECCION', 'LAT', 'LNG', 'MAPS', 'CONSTANCIA_ID', 'CONSTANCIA_NOMBRE', 'NOTAS', 'ALTA', 'ACTIVO'],
  INVENTARIO: ['ID_ARTICULO', 'NOMBRE', 'CATEGORIA', 'UNIDAD', 'COSTO', 'PRECIO', 'MINIMO', 'FOTO_ID', 'ACTIVO'],
  MOVIMIENTOS: ['ID', 'FECHA', 'ID_ARTICULO', 'ARTICULO', 'CANTIDAD', 'TIPO', 'REFERENCIA', 'NOTA', 'QUIEN'],
  RECETAS: ['ID_BLOCK', 'BLOCK', 'PIEZAS_POR_MEZCLA', 'MANO_OBRA_PZA', 'OTROS_PZA', 'MATERIALES'],
  VENTAS: ['ID_VENTA', 'FOLIO', 'FECHA', 'ID_CLIENTE', 'CLIENTE', 'TOTAL', 'METODO', 'PAGADO', 'ENTREGA_PROGRAMADA', 'FECHA_ENTREGA', 'DIRECCION_ENTREGA', 'ESTATUS_ENTREGA', 'ENTREGADA_EL', 'NOTA', 'VENDEDOR'],
  VENTAS_DETALLE: ['ID_VENTA', 'ID_ARTICULO', 'PRODUCTO', 'UNIDAD', 'CANTIDAD', 'PRECIO', 'IMPORTE'],
  ABONOS: ['ID', 'FECHA', 'ID_VENTA', 'MONTO', 'METODO', 'QUIEN'],
  PRODUCCION: ['ID', 'FECHA', 'ID_BLOCK', 'BLOCK', 'PIEZAS', 'NOTA', 'QUIEN'],
  AJUSTES: ['ID', 'FECHA', 'ID_ARTICULO', 'ARTICULO', 'ANTES', 'CONTEO', 'DIFERENCIA', 'MOTIVO', 'QUIEN'],
  GASTOS: ['ID', 'FECHA', 'PROVEEDOR', 'RFC', 'CONCEPTO', 'CATEGORIA', 'METODO', 'TOTAL', 'ES_COMPRA', 'FOTO_ID', 'QUIEN'],
  COMPRAS_DETALLE: ['ID_GASTO', 'ID_ARTICULO', 'ARTICULO', 'CANTIDAD', 'COSTO', 'IMPORTE'],
  FACTURAS: ['ID', 'TIPO', 'FECHA', 'FOLIO', 'UUID', 'RFC', 'NOMBRE', 'ID_CLIENTE', 'TOTAL', 'ARCHIVOS_ID', 'ARCHIVOS_NOMBRE'],
  CUENTAS: ['ID', 'BANCO', 'TITULAR', 'CLABE', 'CUENTA', 'TARJETA', 'NOTA'],
  DUPLICADOS_IGNORADOS: ['PAR'],
};

const CAT_A_APP = { 'MATERIA PRIMA': 'material', 'BLOCK': 'block', 'FERRETERIA': 'ferreteria', 'FERRETERÍA': 'ferreteria' };
const CAT_A_HOJA = { material: 'MATERIA PRIMA', block: 'BLOCK', ferreteria: 'FERRETERIA' };
const EMPRESA_CLAVES = { nombre: 'NOMBRE', slogan: 'ESLOGAN', rfc: 'RFC', razon: 'RAZON_SOCIAL', dir: 'DIRECCION', tel: 'TELEFONO', email: 'CORREO', contadorEmail: 'CONTADOR_CORREO', contadorTel: 'CONTADOR_TELEFONO', leyenda: 'LEYENDA' };

// ======================= ENTRADA WEB =======================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(CONFIG.NEGOCIO)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** La app instalable manda {fn, args} y recibe {ok, r} o {ok:false, error}. Solo funciones api*. */
function doPost(e) {
  let out;
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = String(b.fn || '');
    const f = /^api[A-Z]\w*$/.test(fn) ? globalThis[fn] : null;
    if (typeof f !== 'function') throw new Error('Función no válida');
    out = { ok: true, r: f.apply(null, Array.isArray(b.args) ? b.args : []) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ======================= UTILIDADES =======================
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function hoja_(n) {
  let h = ss_().getSheetByName(n);
  if (!h) { h = ss_().insertSheet(n); h.getRange(1, 1, 1, HOJAS[n].length).setValues([HOJAS[n]]).setFontWeight('bold'); h.setFrozenRows(1); }
  return h;
}
function norm_(s) { return String(s == null ? '' : s).trim(); }
function up_(s) { return norm_(s).toUpperCase(); }
function num_(v) { const n = Number(String(v == null ? '' : v).replace(/[$,\s]/g, '')); return isFinite(n) ? n : 0; }
function r2_(n) { return Math.round(n * 100) / 100; }
function r3_(n) { return Math.round(n * 1000) / 1000; }
function sum_(a, f) { return a.reduce((t, x) => t + num_(f(x)), 0); }
function id_(pre) { let s; do { s = Utilities.getUuid().replace(/-/g, '').slice(0, 8); } while (/^\d*e?\d*$/i.test(s)); return (pre || '') + s.toUpperCase(); }
function ahora_() { return Utilities.formatDate(new Date(), CONFIG.ZONA, "yyyy-MM-dd'T'HH:mm"); }
function hoy_() { return ahora_().slice(0, 10); }
function fmt_(d) {
  const s = Utilities.formatDate(d, CONFIG.ZONA, "yyyy-MM-dd'T'HH:mm");
  return s.slice(11) === '00:00' ? s.slice(0, 10) : s;
}
function celda_(v) { return v instanceof Date ? fmt_(v) : v; }
function txt_(v) { return norm_(celda_(v)); }
function conLock_(fn) { const l = LockService.getScriptLock(); l.waitLock(25000); try { return fn(); } finally { l.releaseLock(); } }

function ultimaFila_(h) {
  const n = h.getLastRow(); if (n < 2) return n;
  const col = h.getRange(1, 1, n, 1).getValues();
  for (let i = n - 1; i >= 0; i--) if (col[i][0] !== '' && col[i][0] !== null) return i + 1;
  return 1;
}
function leer_(n) {
  const h = hoja_(n), enc = HOJAS[n], lr = ultimaFila_(h);
  if (lr < 2) return [];
  const vals = h.getRange(2, 1, lr - 1, enc.length).getValues();
  const out = [];
  vals.forEach((r, i) => { if (r[0] === '' || r[0] === null) return; const o = { _fila: i + 2 }; enc.forEach((k, j) => o[k] = celda_(r[j])); out.push(o); });
  return out;
}
function fila_(n, o) { return HOJAS[n].map(k => (o[k] == null ? '' : o[k])); }
function escribirMuchos_(n, objs) {
  if (!objs.length) return;
  const h = hoja_(n), r = ultimaFila_(h) + 1;
  h.getRange(r, 1, objs.length, HOJAS[n].length).setValues(objs.map(o => fila_(n, o)));
}
function escribir_(n, o) { escribirMuchos_(n, [o]); }
function actualizar_(n, o) { hoja_(n).getRange(o._fila, 1, 1, HOJAS[n].length).setValues([fila_(n, o)]); }
function borrarFilas_(n, pred) {
  const h = hoja_(n);
  leer_(n).filter(pred).map(o => o._fila).sort((a, b) => b - a).forEach(f => h.deleteRow(f));
}
function buscar_(n, col, val) { return leer_(n).find(o => norm_(o[col]) === norm_(val)); }
function reemplazarEnColumna_(n, col, de, a) {
  const h = hoja_(n), lr = ultimaFila_(h); if (lr < 2) return;
  const c = HOJAS[n].indexOf(col) + 1, rg = h.getRange(2, c, lr - 1, 1), v = rg.getValues();
  let cambio = false; v.forEach(r => { if (norm_(r[0]) === de) { r[0] = a; cambio = true; } });
  if (cambio) rg.setValues(v);
}

// ======================= SESIÓN =======================
function sesion_(token, roles) {
  const c = CacheService.getScriptCache(), v = c.get('s_' + norm_(token));
  if (!v) throw new Error('SESION');
  c.put('s_' + norm_(token), v, CONFIG.SESION_SEG);
  const u = JSON.parse(v);
  if (roles && u.rol !== 'DUENO' && roles.indexOf(u.rol) < 0) throw new Error('Tu usuario no tiene permiso para esto. Pídeselo al dueño.');
  return u;
}
function apiLogin(pin) {
  pin = norm_(pin);
  const c = CacheService.getScriptCache(), n = num_(c.get('intentos'));
  if (n >= 15) throw new Error('Demasiados intentos con PIN equivocado. Espera 10 minutos.');
  const u = leer_('USUARIOS_APP').find(x => norm_(x.PIN) === pin && up_(x.ACTIVO) !== 'NO');
  if (!u) { c.put('intentos', String(n + 1), 600); throw new Error('PIN incorrecto'); }
  const t = Utilities.getUuid(), s = { nombre: norm_(u.NOMBRE), rol: up_(u.ROL) };
  c.put('s_' + t, JSON.stringify(s), CONFIG.SESION_SEG);
  return { token: t, d: datos_(s) };
}
function apiLogout(token) { CacheService.getScriptCache().remove('s_' + norm_(token)); return true; }
function apiDatos(token) { return { d: datos_(sesion_(token)) }; }

// ======================= ARCHIVOS EN DRIVE =======================
function carpetaRaiz_() {
  const p = PropertiesService.getScriptProperties(); const id = p.getProperty('carpeta');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(CONFIG.CARPETA); p.setProperty('carpeta', f.getId()); return f;
}
function subcarpeta_(ruta) {
  let f = carpetaRaiz_();
  ruta.split('/').forEach(n => { const it = f.getFoldersByName(n); f = it.hasNext() ? it.next() : f.createFolder(n); });
  return f;
}
function guardarArchivo_(ruta, a, publico) {
  const b64 = String(a.data || '').replace(/^data:[^,]*,/, '');
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), a.type || 'application/octet-stream', a.name || 'archivo');
  const f = subcarpeta_(ruta).createFile(blob);
  if (publico) { try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {} }
  return f.getId();
}
function tirar_(id) { if (!norm_(id)) return; try { DriveApp.getFileById(norm_(id)).setTrashed(true); } catch (e) {} }
function enCarpeta_(file) {
  const raiz = carpetaRaiz_().getId(); let nivel = [file.getParents()];
  for (let n = 0; n < 6 && nivel.length; n++) {
    const sig = [];
    for (const it of nivel) while (it.hasNext()) { const p = it.next(); if (p.getId() === raiz) return true; sig.push(p.getParents()); }
    nivel = sig;
  }
  return false;
}
function fotoUrl_(id) { return id ? 'https://lh3.googleusercontent.com/d/' + id + '=w600' : ''; }
function logoData_(id) {
  id = norm_(id); if (!id) return '';
  const c = CacheService.getScriptCache(), k = 'logo_' + id, v = c.get(k); if (v) return v;
  try { const b = DriveApp.getFileById(id).getBlob(); const d = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()); if (d.length < 95000) c.put(k, d, 21600); return d; } catch (e) { return ''; }
}
/** Abre un archivo guardado por la app (constancia, factura, ticket). */
function apiArchivo(token, fileId) {
  sesion_(token);
  const f = DriveApp.getFileById(norm_(fileId));
  if (!enCarpeta_(f)) throw new Error('Ese archivo no es de la app');
  const b = f.getBlob();
  return { name: f.getName(), type: b.getContentType(), data: Utilities.base64Encode(b.getBytes()) };
}

// ======================= DATOS COMPLETOS =======================
function datos_(u) {
  const E = {}; leer_('EMPRESA').forEach(r => E[up_(r.CLAVE)] = txt_(r.VALOR));
  const empresa = {}; Object.keys(EMPRESA_CLAVES).forEach(k => empresa[k] = E[EMPRESA_CLAVES[k]] || '');
  empresa.logo = logoData_(E.LOGO_ID);

  const movsH = leer_('MOVIMIENTOS'), stock = {};
  movsH.forEach(m => { const k = txt_(m.ID_ARTICULO); stock[k] = r3_((stock[k] || 0) + num_(m.CANTIDAD)); });
  const movs = movsH.map(m => ({ f: txt_(m.FECHA), id: txt_(m.ID_ARTICULO), c: num_(m.CANTIDAD), t: txt_(m.TIPO), r: txt_(m.REFERENCIA), n: txt_(m.NOTA) }));

  const inventario = leer_('INVENTARIO').filter(i => up_(i.ACTIVO) !== 'NO').map(i => {
    const id = txt_(i.ID_ARTICULO);
    return { id, nombre: txt_(i.NOMBRE), cat: CAT_A_APP[up_(i.CATEGORIA)] || 'material', unidad: txt_(i.UNIDAD), stock: stock[id] || 0, costo: num_(i.COSTO), precio: num_(i.PRECIO), min: num_(i.MINIMO), foto: fotoUrl_(txt_(i.FOTO_ID)) };
  });

  const recetas = {};
  leer_('RECETAS').forEach(r => {
    recetas[txt_(r.ID_BLOCK)] = {
      piezasMezcla: num_(r.PIEZAS_POR_MEZCLA) || 1, mo: num_(r.MANO_OBRA_PZA), otros: num_(r.OTROS_PZA),
      items: txt_(r.MATERIALES).split(';').map(x => x.split(':')).filter(x => norm_(x[0])).map(x => ({ matId: norm_(x[0]), cant: num_(x[1]) })),
    };
  });

  const clientes = leer_('CLIENTES').filter(c => up_(c.ACTIVO) !== 'NO').map(c => ({
    id: txt_(c.ID_CLIENTE), nombre: txt_(c.NOMBRE), tel: txt_(c.TELEFONO), email: txt_(c.CORREO), rfc: txt_(c.RFC), razon: txt_(c.RAZON_SOCIAL),
    regimen: txt_(c.REGIMEN), usoCfdi: txt_(c.USO_CFDI), cp: txt_(c.CP), dir: txt_(c.DIRECCION),
    lat: norm_(c.LAT) ? num_(c.LAT) : null, lng: norm_(c.LNG) ? num_(c.LNG) : null, maps: txt_(c.MAPS),
    constancia: txt_(c.CONSTANCIA_ID) ? { id: txt_(c.CONSTANCIA_ID), name: txt_(c.CONSTANCIA_NOMBRE) || 'Constancia' } : null, notas: txt_(c.NOTAS),
  })).sort((a, b) => a.nombre.localeCompare(b.nombre));

  const det = {}; leer_('VENTAS_DETALLE').forEach(d => { const k = txt_(d.ID_VENTA); (det[k] = det[k] || []).push({ prodId: txt_(d.ID_ARTICULO), nombre: txt_(d.PRODUCTO), unidad: txt_(d.UNIDAD), cant: num_(d.CANTIDAD), precio: num_(d.PRECIO) }); });
  const abo = {}; leer_('ABONOS').forEach(a => { const k = txt_(a.ID_VENTA); (abo[k] = abo[k] || []).push({ fecha: txt_(a.FECHA), monto: num_(a.MONTO), metodo: txt_(a.METODO) }); });
  const ventas = leer_('VENTAS').map(v => {
    const id = txt_(v.ID_VENTA), prog = up_(v.ENTREGA_PROGRAMADA) === 'SI';
    return { id, folio: num_(v.FOLIO), fecha: txt_(v.FECHA), clienteId: txt_(v.ID_CLIENTE), clienteNombre: txt_(v.CLIENTE), items: det[id] || [], total: num_(v.TOTAL),
      metodo: txt_(v.METODO), pagado: num_(v.PAGADO), abonos: abo[id] || [], nota: txt_(v.NOTA), vendedor: txt_(v.VENDEDOR),
      entrega: prog ? { programada: true, fecha: txt_(v.FECHA_ENTREGA), dir: txt_(v.DIRECCION_ENTREGA), estado: up_(v.ESTATUS_ENTREGA) === 'ENTREGADA' ? 'entregada' : 'pendiente', entregadaEl: txt_(v.ENTREGADA_EL) } : null };
  }).sort((a, b) => b.folio - a.folio);

  const consumos = {}; movs.forEach(m => { if (m.t === 'CONSUMO') (consumos[m.r] = consumos[m.r] || []).push({ matId: m.id, cant: -m.c }); });
  const producciones = leer_('PRODUCCION').map(p => ({ id: txt_(p.ID), fecha: txt_(p.FECHA).slice(0, 10), blockId: txt_(p.ID_BLOCK), cant: num_(p.PIEZAS), nota: txt_(p.NOTA), consumos: consumos[txt_(p.ID)] || [] }))
    .sort((a, b) => b.fecha.localeCompare(a.fecha));
  const ajustes = leer_('AJUSTES').map(a => ({ id: txt_(a.ID), fecha: txt_(a.FECHA), matId: txt_(a.ID_ARTICULO), antes: num_(a.ANTES), despues: num_(a.CONTEO), motivo: txt_(a.MOTIVO) })).reverse();

  const cd = {}; leer_('COMPRAS_DETALLE').forEach(d => { const k = txt_(d.ID_GASTO); (cd[k] = cd[k] || []).push({ matId: txt_(d.ID_ARTICULO), cant: num_(d.CANTIDAD), costo: num_(d.COSTO) }); });
  const gastos = leer_('GASTOS').map(g => ({ id: txt_(g.ID), fecha: txt_(g.FECHA).slice(0, 10), proveedor: txt_(g.PROVEEDOR), rfc: txt_(g.RFC), concepto: txt_(g.CONCEPTO), cat: txt_(g.CATEGORIA), metodo: txt_(g.METODO), total: num_(g.TOTAL), items: cd[txt_(g.ID)] || [], fotoId: txt_(g.FOTO_ID) }))
    .sort((a, b) => b.fecha.localeCompare(a.fecha));

  const facturas = leer_('FACTURAS').map(f => {
    const ids = txt_(f.ARCHIVOS_ID).split('|').filter(String), nom = txt_(f.ARCHIVOS_NOMBRE).split('|');
    return { id: txt_(f.ID), tipo: up_(f.TIPO) === 'RECIBIDA' ? 'recibida' : 'emitida', fecha: txt_(f.FECHA).slice(0, 10), folio: txt_(f.FOLIO), uuid: txt_(f.UUID), rfc: txt_(f.RFC), nombre: txt_(f.NOMBRE), clienteId: txt_(f.ID_CLIENTE), total: num_(f.TOTAL), files: ids.map((id, i) => ({ id, name: nom[i] || 'archivo' })) };
  });
  const cuentas = leer_('CUENTAS').map(k => ({ id: txt_(k.ID), banco: txt_(k.BANCO), titular: txt_(k.TITULAR), clabe: txt_(k.CLABE), cuenta: txt_(k.CUENTA), tarjeta: txt_(k.TARJETA), nota: txt_(k.NOTA) }));

  return {
    usuario: u, empresa, inventario, recetas, clientes, ventas, producciones, ajustes, gastos, facturas, cuentas, movs,
    config: { folio: ventas.reduce((m, v) => Math.max(m, v.folio), 0), ignorados: leer_('DUPLICADOS_IGNORADOS').map(x => txt_(x.PAR)) },
    servidor: ahora_(),
  };
}
function mapaInv_() { const m = {}; leer_('INVENTARIO').forEach(i => m[txt_(i.ID_ARTICULO)] = i); return m; }
function stockDe_(id) { return r3_(sum_(leer_('MOVIMIENTOS').filter(m => txt_(m.ID_ARTICULO) === id), m => m.CANTIDAD)); }
function mov_(u, id, nombre, cant, tipo, ref, nota, fecha) {
  return { ID: id_('M'), FECHA: fecha || ahora_(), ID_ARTICULO: id, ARTICULO: nombre, CANTIDAD: r3_(cant), TIPO: tipo, REFERENCIA: ref || '', NOTA: nota || '', QUIEN: u.nombre };
}

// ======================= VENTAS Y ENTREGAS =======================
function apiVenta(token, v) {
  const u = sesion_(token, ['MOSTRADOR']);
  return conLock_(() => {
    const inv = mapaInv_();
    const items = (v.items || []).filter(i => norm_(i.prodId) && num_(i.cant) > 0);
    if (!items.length) throw new Error('La venta no tiene productos');
    items.forEach(i => { if (!inv[i.prodId]) throw new Error('Un producto ya no existe en el inventario'); });
    const total = r2_(sum_(items, i => num_(i.cant) * num_(i.precio)));
    const metodo = ['Efectivo', 'Transferencia', 'Tarjeta', 'Fiado'].indexOf(v.metodo) >= 0 ? v.metodo : 'Efectivo';
    const pagado = metodo === 'Fiado' ? Math.min(total, Math.max(0, num_(v.pagado))) : total;
    const folio = leer_('VENTAS').reduce((m, x) => Math.max(m, num_(x.FOLIO)), 0) + 1;
    const id = id_('V'), f = ahora_(), c = v.clienteId ? buscar_('CLIENTES', 'ID_CLIENTE', v.clienteId) : null;
    const ent = v.entrega && v.entrega.fecha;
    escribir_('VENTAS', { ID_VENTA: id, FOLIO: folio, FECHA: f, ID_CLIENTE: c ? txt_(c.ID_CLIENTE) : '', CLIENTE: c ? txt_(c.NOMBRE) : 'Público en general', TOTAL: total, METODO: metodo, PAGADO: pagado,
      ENTREGA_PROGRAMADA: ent ? 'SI' : 'NO', FECHA_ENTREGA: ent ? norm_(v.entrega.fecha) : '', DIRECCION_ENTREGA: ent ? norm_(v.entrega.dir) : '', ESTATUS_ENTREGA: ent ? 'PENDIENTE' : '', ENTREGADA_EL: '', NOTA: norm_(v.nota), VENDEDOR: u.nombre });
    escribirMuchos_('VENTAS_DETALLE', items.map(i => ({ ID_VENTA: id, ID_ARTICULO: i.prodId, PRODUCTO: txt_(inv[i.prodId].NOMBRE), UNIDAD: txt_(inv[i.prodId].UNIDAD), CANTIDAD: num_(i.cant), PRECIO: num_(i.precio), IMPORTE: r2_(num_(i.cant) * num_(i.precio)) })));
    escribirMuchos_('MOVIMIENTOS', items.map(i => mov_(u, i.prodId, txt_(inv[i.prodId].NOMBRE), -num_(i.cant), 'VENTA', id, 'Folio ' + folio, f)));
    return { id, d: datos_(u) };
  });
}
function venta_(id) { const v = buscar_('VENTAS', 'ID_VENTA', id); if (!v) throw new Error('No encontré esa venta'); return v; }
function apiAbono(token, ventaId, monto, metodo) {
  const u = sesion_(token, ['MOSTRADOR']);
  return conLock_(() => {
    const v = venta_(ventaId), saldo = r2_(num_(v.TOTAL) - num_(v.PAGADO)), m = Math.min(saldo, num_(monto));
    if (m <= 0) throw new Error('Escribe un monto mayor a cero');
    v.PAGADO = r2_(num_(v.PAGADO) + m); actualizar_('VENTAS', v);
    escribir_('ABONOS', { ID: id_('A'), FECHA: ahora_(), ID_VENTA: txt_(v.ID_VENTA), MONTO: m, METODO: norm_(metodo) || 'Efectivo', QUIEN: u.nombre });
    return { d: datos_(u) };
  });
}
function apiEntrega(token, ventaId, fecha, dir) {
  const u = sesion_(token, ['MOSTRADOR']);
  return conLock_(() => {
    const v = venta_(ventaId); if (!norm_(fecha)) throw new Error('Elige la fecha de entrega');
    Object.assign(v, { ENTREGA_PROGRAMADA: 'SI', FECHA_ENTREGA: norm_(fecha), DIRECCION_ENTREGA: norm_(dir), ESTATUS_ENTREGA: 'PENDIENTE', ENTREGADA_EL: '' });
    actualizar_('VENTAS', v); return { d: datos_(u) };
  });
}
function apiEntregada(token, ventaId) {
  const u = sesion_(token, ['MOSTRADOR']);
  return conLock_(() => { const v = venta_(ventaId); v.ESTATUS_ENTREGA = 'ENTREGADA'; v.ENTREGADA_EL = ahora_(); actualizar_('VENTAS', v); return { d: datos_(u) }; });
}
function apiBorrarVenta(token, ventaId) {
  const u = sesion_(token, []);
  return conLock_(() => {
    venta_(ventaId);
    borrarFilas_('MOVIMIENTOS', m => txt_(m.REFERENCIA) === ventaId);
    borrarFilas_('VENTAS_DETALLE', d => txt_(d.ID_VENTA) === ventaId);
    borrarFilas_('ABONOS', a => txt_(a.ID_VENTA) === ventaId);
    borrarFilas_('VENTAS', v => txt_(v.ID_VENTA) === ventaId);
    return { d: datos_(u) };
  });
}

// ======================= CLIENTES =======================
function apiCliente(token, c) {
  const u = sesion_(token, ['MOSTRADOR']);
  return conLock_(() => {
    if (!norm_(c.nombre)) throw new Error('Escribe el nombre del cliente');
    let row = c.id ? buscar_('CLIENTES', 'ID_CLIENTE', c.id) : null;
    const nuevo = !row; if (nuevo) row = { ID_CLIENTE: id_('C'), ALTA: ahora_(), ACTIVO: 'SI' };
    if (c.constanciaNueva && c.constanciaNueva.data) {
      tirar_(row.CONSTANCIA_ID);
      row.CONSTANCIA_ID = guardarArchivo_('Constancias', { name: (norm_(c.rfc) || norm_(c.nombre)) + ' - ' + c.constanciaNueva.name, type: c.constanciaNueva.type, data: c.constanciaNueva.data });
      row.CONSTANCIA_NOMBRE = c.constanciaNueva.name;
    }
    Object.assign(row, { NOMBRE: norm_(c.nombre), TELEFONO: norm_(c.tel), CORREO: norm_(c.email), RFC: up_(c.rfc), RAZON_SOCIAL: norm_(c.razon), REGIMEN: norm_(c.regimen), USO_CFDI: norm_(c.usoCfdi),
      CP: norm_(c.cp), DIRECCION: norm_(c.dir), LAT: c.lat == null ? '' : String(c.lat), LNG: c.lng == null ? '' : String(c.lng), MAPS: norm_(c.maps), NOTAS: norm_(c.notas) });
    if (nuevo) escribir_('CLIENTES', row); else actualizar_('CLIENTES', row);
    return { id: txt_(row.ID_CLIENTE), d: datos_(u) };
  });
}
function apiBorrarCliente(token, id) {
  const u = sesion_(token, []);
  return conLock_(() => { const c = buscar_('CLIENTES', 'ID_CLIENTE', id); if (c) { c.ACTIVO = 'NO'; actualizar_('CLIENTES', c); } return { d: datos_(u) }; });
}

// ======================= PRODUCCIÓN Y AJUSTES =======================
function apiProduccion(token, p) {
  const u = sesion_(token, ['PRODUCCION']);
  return conLock_(() => {
    const inv = mapaInv_(), b = inv[p.blockId], cant = num_(p.cant);
    if (!b) throw new Error('Elige el tipo de block');
    if (cant <= 0) throw new Error('Escribe las piezas producidas');
    const id = id_('P'), fecha = norm_(p.fecha) || hoy_();
    escribir_('PRODUCCION', { ID: id, FECHA: fecha, ID_BLOCK: p.blockId, BLOCK: txt_(b.NOMBRE), PIEZAS: cant, NOTA: norm_(p.nota), QUIEN: u.nombre });
    const ms = [mov_(u, p.blockId, txt_(b.NOMBRE), cant, 'PRODUCCION', id, '', fecha)];
    (p.consumos || []).filter(c => inv[c.matId] && num_(c.cant) > 0).forEach(c => ms.push(mov_(u, c.matId, txt_(inv[c.matId].NOMBRE), -num_(c.cant), 'CONSUMO', id, 'Producción de ' + txt_(b.NOMBRE), fecha)));
    escribirMuchos_('MOVIMIENTOS', ms);
    return { d: datos_(u) };
  });
}
function apiBorrarProduccion(token, id) {
  const u = sesion_(token, []);
  return conLock_(() => { borrarFilas_('MOVIMIENTOS', m => txt_(m.REFERENCIA) === id); borrarFilas_('PRODUCCION', p => txt_(p.ID) === id); return { d: datos_(u) }; });
}
function ajuste_(u, art, real, motivo) {
  const id = txt_(art.ID_ARTICULO), antes = stockDe_(id), dif = r3_(num_(real) - antes);
  const aid = id_('J');
  escribir_('AJUSTES', { ID: aid, FECHA: hoy_(), ID_ARTICULO: id, ARTICULO: txt_(art.NOMBRE), ANTES: antes, CONTEO: num_(real), DIFERENCIA: dif, MOTIVO: norm_(motivo) || 'Conteo físico', QUIEN: u.nombre });
  if (dif !== 0) escribir_('MOVIMIENTOS', mov_(u, id, txt_(art.NOMBRE), dif, 'AJUSTE', aid, norm_(motivo)));
}
function apiAjuste(token, artId, real, motivo) {
  const u = sesion_(token, ['PRODUCCION', 'MOSTRADOR']);
  return conLock_(() => { const a = mapaInv_()[artId]; if (!a) throw new Error('No encontré ese artículo'); ajuste_(u, a, real, motivo); return { d: datos_(u) }; });
}

// ======================= INVENTARIO =======================
function apiArticulo(token, a) {
  const u = sesion_(token, ['MOSTRADOR', 'PRODUCCION']);
  return conLock_(() => {
    if (!norm_(a.nombre)) throw new Error('Escribe el nombre');
    let row = a.id ? buscar_('INVENTARIO', 'ID_ARTICULO', a.id) : null;
    const nuevo = !row;
    if (nuevo) row = { ID_ARTICULO: ({ material: 'MP-', block: 'BL-', ferreteria: 'FE-' }[a.cat] || 'AR-') + id_().slice(0, 6), ACTIVO: 'SI', FOTO_ID: '' };
    if (a.quitarFoto) { tirar_(row.FOTO_ID); row.FOTO_ID = ''; }
    if (a.fotoNueva) { tirar_(row.FOTO_ID); row.FOTO_ID = guardarArchivo_('Fotos de inventario', { name: norm_(a.nombre) + '.jpg', type: 'image/jpeg', data: a.fotoNueva }, true); }
    Object.assign(row, { NOMBRE: norm_(a.nombre), CATEGORIA: CAT_A_HOJA[a.cat] || 'MATERIA PRIMA', UNIDAD: norm_(a.unidad) || 'pza', COSTO: num_(a.costo), PRECIO: num_(a.precio), MINIMO: num_(a.min) });
    if (nuevo) { escribir_('INVENTARIO', row); if (num_(a.stock)) escribir_('MOVIMIENTOS', mov_(u, row.ID_ARTICULO, row.NOMBRE, num_(a.stock), 'INICIAL', '', 'Existencia inicial')); }
    else { actualizar_('INVENTARIO', row); if (a.stock != null && r3_(num_(a.stock)) !== stockDe_(txt_(row.ID_ARTICULO))) ajuste_(u, row, a.stock, 'Edición manual'); }
    return { id: txt_(row.ID_ARTICULO), d: datos_(u) };
  });
}
function apiBorrarArticulo(token, id) {
  const u = sesion_(token, []);
  return conLock_(() => { const a = buscar_('INVENTARIO', 'ID_ARTICULO', id); if (a) { a.ACTIVO = 'NO'; actualizar_('INVENTARIO', a); } return { d: datos_(u) }; });
}
/** Junta el artículo B dentro del A: suma existencias y pasa todo el historial. */
function apiFusionar(token, aId, bId) {
  const u = sesion_(token, []);
  return conLock_(() => {
    const inv = mapaInv_(), a = inv[aId], b = inv[bId]; if (!a || !b) throw new Error('No encontré los artículos');
    reemplazarEnColumna_('MOVIMIENTOS', 'ID_ARTICULO', bId, aId);
    reemplazarEnColumna_('VENTAS_DETALLE', 'ID_ARTICULO', bId, aId);
    reemplazarEnColumna_('COMPRAS_DETALLE', 'ID_ARTICULO', bId, aId);
    reemplazarEnColumna_('AJUSTES', 'ID_ARTICULO', bId, aId);
    reemplazarEnColumna_('PRODUCCION', 'ID_BLOCK', bId, aId);
    leer_('RECETAS').forEach(r => {
      const m = txt_(r.MATERIALES).split(';').map(s => s.trim()).filter(String).map(s => { const [k, v] = s.split(':'); return [norm_(k) === bId ? aId : norm_(k), v]; });
      const nuevo = m.map(x => x[0] + ':' + norm_(x[1])).join('; ');
      if (nuevo !== txt_(r.MATERIALES) || txt_(r.ID_BLOCK) === bId) { r.MATERIALES = nuevo; if (txt_(r.ID_BLOCK) === bId) r.ID_BLOCK = aId; actualizar_('RECETAS', r); }
    });
    if (!txt_(a.FOTO_ID) && txt_(b.FOTO_ID)) { a.FOTO_ID = b.FOTO_ID; actualizar_('INVENTARIO', a); }
    b.ACTIVO = 'NO'; b.NOMBRE = txt_(b.NOMBRE) + ' (fusionado en ' + txt_(a.NOMBRE) + ')'; actualizar_('INVENTARIO', b);
    return { d: datos_(u) };
  });
}
function apiIgnorarDuplicado(token, par) {
  const u = sesion_(token, []);
  return conLock_(() => { escribir_('DUPLICADOS_IGNORADOS', { PAR: norm_(par) }); return { d: datos_(u) }; });
}

// ======================= COSTO DE PRODUCCIÓN =======================
function apiReceta(token, blockId, r, precio) {
  const u = sesion_(token, []);
  return conLock_(() => {
    const inv = mapaInv_(), b = inv[blockId]; if (!b) throw new Error('No encontré ese block');
    const mats = (r.items || []).filter(i => norm_(i.matId)).map(i => norm_(i.matId) + ':' + num_(i.cant)).join('; ');
    const datos = { ID_BLOCK: blockId, BLOCK: txt_(b.NOMBRE), PIEZAS_POR_MEZCLA: num_(r.piezasMezcla) || 1, MANO_OBRA_PZA: num_(r.mo), OTROS_PZA: num_(r.otros), MATERIALES: mats };
    const row = buscar_('RECETAS', 'ID_BLOCK', blockId);
    if (row) actualizar_('RECETAS', Object.assign(row, datos)); else escribir_('RECETAS', datos);
    if (precio != null && num_(precio) !== num_(b.PRECIO)) { b.PRECIO = num_(precio); actualizar_('INVENTARIO', b); }
    return { d: datos_(u) };
  });
}

// ======================= GASTOS Y COMPRAS =======================
function apiGasto(token, g) {
  const u = sesion_(token, ['MOSTRADOR']);
  return conLock_(() => {
    const base = { FECHA: norm_(g.fecha) || hoy_(), PROVEEDOR: norm_(g.proveedor), RFC: up_(g.rfc), CONCEPTO: norm_(g.concepto), CATEGORIA: norm_(g.cat) || 'Otro', METODO: norm_(g.metodo) || 'Efectivo', TOTAL: r2_(num_(g.total)) };
    if (g.id) { const row = buscar_('GASTOS', 'ID', g.id); if (!row) throw new Error('No encontré ese gasto'); Object.assign(row, base); actualizar_('GASTOS', row); return { id: g.id, d: datos_(u) }; }
    const id = id_('G'), inv = mapaInv_();
    const items = (g.items || []).filter(i => inv[i.matId] && num_(i.cant) > 0);
    const foto = g.fotoNueva ? guardarArchivo_('Tickets/' + base.FECHA.slice(0, 7), { name: base.FECHA + ' ' + (base.PROVEEDOR || 'ticket') + '.jpg', type: 'image/jpeg', data: g.fotoNueva }) : '';
    escribir_('GASTOS', Object.assign({ ID: id, ES_COMPRA: items.length ? 'SI' : 'NO', FOTO_ID: foto, QUIEN: u.nombre }, base));
    if (items.length) {
      escribirMuchos_('COMPRAS_DETALLE', items.map(i => ({ ID_GASTO: id, ID_ARTICULO: i.matId, ARTICULO: txt_(inv[i.matId].NOMBRE), CANTIDAD: num_(i.cant), COSTO: num_(i.costo), IMPORTE: r2_(num_(i.cant) * num_(i.costo)) })));
      escribirMuchos_('MOVIMIENTOS', items.map(i => mov_(u, i.matId, txt_(inv[i.matId].NOMBRE), num_(i.cant), 'COMPRA', id, base.PROVEEDOR, base.FECHA)));
      items.forEach(i => { if (num_(i.costo) > 0) { const a = inv[i.matId]; a.COSTO = num_(i.costo); actualizar_('INVENTARIO', a); } });
    }
    return { id, d: datos_(u) };
  });
}
function apiBorrarGasto(token, id) {
  const u = sesion_(token, []);
  return conLock_(() => {
    const g = buscar_('GASTOS', 'ID', id); if (g) tirar_(g.FOTO_ID);
    borrarFilas_('MOVIMIENTOS', m => txt_(m.REFERENCIA) === id);
    borrarFilas_('COMPRAS_DETALLE', d => txt_(d.ID_GASTO) === id);
    borrarFilas_('GASTOS', x => txt_(x.ID) === id);
    return { d: datos_(u) };
  });
}

// ======================= FACTURAS =======================
function apiFactura(token, f) {
  const u = sesion_(token, ['CONTADOR']);
  return conLock_(() => {
    const fecha = norm_(f.fecha) || hoy_(), tipo = f.tipo === 'recibida' ? 'RECIBIDA' : 'EMITIDA';
    const ruta = 'Facturas/' + fecha.slice(0, 7) + '/' + (tipo === 'EMITIDA' ? 'Emitidas' : 'Recibidas');
    const ids = [], noms = [];
    (f.files || []).forEach(a => { ids.push(guardarArchivo_(ruta, a)); noms.push(norm_(a.name).replace(/\|/g, '-')); });
    escribir_('FACTURAS', { ID: id_('F'), TIPO: tipo, FECHA: fecha, FOLIO: norm_(f.folio), UUID: up_(f.uuid), RFC: up_(f.rfc), NOMBRE: norm_(f.nombre), ID_CLIENTE: norm_(f.clienteId), TOTAL: r2_(num_(f.total)), ARCHIVOS_ID: ids.join('|'), ARCHIVOS_NOMBRE: noms.join('|') });
    return { d: datos_(u) };
  });
}
function apiBorrarFactura(token, id) {
  const u = sesion_(token, ['CONTADOR']);
  return conLock_(() => {
    const f = buscar_('FACTURAS', 'ID', id); if (f) txt_(f.ARCHIVOS_ID).split('|').forEach(tirar_);
    borrarFilas_('FACTURAS', x => txt_(x.ID) === id);
    return { d: datos_(u) };
  });
}
/** Arma el ZIP del mes (Emitidas/, Recibidas/, resumen.csv), lo guarda en Drive y lo regresa para compartir. */
function apiFacturasZip(token, mes) {
  sesion_(token, ['CONTADOR']);
  mes = norm_(mes);
  const L = leer_('FACTURAS').filter(f => txt_(f.FECHA).slice(0, 7) === mes);
  if (!L.length) throw new Error('No hay facturas en ese mes');
  const blobs = [], usados = {};
  const csv = [['Tipo', 'Fecha', 'Folio', 'UUID', 'RFC', 'Nombre', 'Total', 'Archivos']];
  L.forEach(f => {
    const dir = up_(f.TIPO) === 'RECIBIDA' ? 'Recibidas' : 'Emitidas';
    txt_(f.ARCHIVOS_ID).split('|').filter(String).forEach(id => {
      try {
        const file = DriveApp.getFileById(id); let n = dir + '/' + file.getName();
        while (usados[n]) n = n.replace(/(\.\w+)?$/, '_1$1'); usados[n] = 1;
        blobs.push(file.getBlob().setName(n));
      } catch (e) {}
    });
    csv.push([up_(f.TIPO), txt_(f.FECHA), txt_(f.FOLIO), txt_(f.UUID), txt_(f.RFC), txt_(f.NOMBRE), num_(f.TOTAL).toFixed(2), txt_(f.ARCHIVOS_NOMBRE).replace(/\|/g, ' / ')]);
  });
  blobs.push(Utilities.newBlob('﻿' + csv.map(r => r.map(c => '"' + String(c).replace(/"/g, '""') + '"').join(',')).join('\n'), 'text/csv', 'resumen.csv'));
  const E = {}; leer_('EMPRESA').forEach(r => E[up_(r.CLAVE)] = txt_(r.VALOR));
  const name = 'Facturas_' + (E.NOMBRE || 'empresa').replace(/\W+/g, '_') + '_' + mes + '.zip';
  const zip = Utilities.zip(blobs, name);
  const file = subcarpeta_('Paquetes para el contador').createFile(zip);
  return { name, data: Utilities.base64Encode(zip.getBytes()), url: file.getUrl(), emitidas: L.filter(f => up_(f.TIPO) !== 'RECIBIDA').length, recibidas: L.filter(f => up_(f.TIPO) === 'RECIBIDA').length };
}

// ======================= CUENTAS Y EMPRESA =======================
function apiCuenta(token, k) {
  const u = sesion_(token, []);
  return conLock_(() => {
    const datos = { BANCO: norm_(k.banco), TITULAR: norm_(k.titular), CLABE: String(k.clabe || '').replace(/\D/g, ''), CUENTA: norm_(k.cuenta), TARJETA: String(k.tarjeta || '').replace(/\D/g, ''), NOTA: norm_(k.nota) };
    if (!datos.BANCO) throw new Error('Escribe el banco');
    if (datos.CLABE && datos.CLABE.length !== 18) throw new Error('La CLABE debe tener 18 dígitos');
    const row = k.id ? buscar_('CUENTAS', 'ID', k.id) : null;
    if (row) actualizar_('CUENTAS', Object.assign(row, datos)); else escribir_('CUENTAS', Object.assign({ ID: id_('K') }, datos));
    return { d: datos_(u) };
  });
}
function apiBorrarCuenta(token, id) {
  const u = sesion_(token, []);
  return conLock_(() => { borrarFilas_('CUENTAS', x => txt_(x.ID) === id); return { d: datos_(u) }; });
}
function apiEmpresa(token, e, logoNuevo) {
  const u = sesion_(token, []);
  return conLock_(() => {
    const filas = leer_('EMPRESA'), get = k => filas.find(r => up_(r.CLAVE) === k);
    const poner = (k, v) => { const r = get(k); if (r) { r.VALOR = v; actualizar_('EMPRESA', r); } else { escribir_('EMPRESA', { CLAVE: k, VALOR: v }); filas.push(leer_('EMPRESA').find(x => up_(x.CLAVE) === k)); } };
    Object.keys(EMPRESA_CLAVES).forEach(k => { if (e[k] != null) poner(EMPRESA_CLAVES[k], k === 'rfc' ? up_(e[k]) : norm_(e[k])); });
    if (logoNuevo === '') { const r = get('LOGO_ID'); if (r) tirar_(r.VALOR); poner('LOGO_ID', ''); }
    else if (logoNuevo) { const r = get('LOGO_ID'); if (r) tirar_(r.VALOR); poner('LOGO_ID', guardarArchivo_('Logo', { name: 'logo.png', type: (String(logoNuevo).match(/^data:([^;]+)/) || [])[1] || 'image/png', data: logoNuevo })); }
    return { d: datos_(u) };
  });
}

// ======================= INSTALACIÓN =======================
/** Ejecuta esta función UNA vez desde el editor (botón Ejecutar). Revisa hojas, crea la carpeta y pide permisos. */
function instalar() {
  Object.keys(HOJAS).forEach(n => {
    const h = hoja_(n), enc = h.getRange(1, 1, 1, HOJAS[n].length).getValues()[0].map(norm_);
    if (enc.join('|') !== HOJAS[n].join('|')) h.getRange(1, 1, 1, HOJAS[n].length).setValues([HOJAS[n]]).setFontWeight('bold');
  });
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('secreto')) props.setProperty('secreto', Utilities.getUuid());
  const carpeta = carpetaRaiz_();
  ['Constancias', 'Facturas', 'Tickets', 'Fotos de inventario', 'Logo', 'Paquetes para el contador'].forEach(subcarpeta_);
  let us = leer_('USUARIOS_APP');
  if (!us.length) {
    const pin = () => String(Math.floor(100000 + Math.random() * 900000));
    escribir_('USUARIOS_APP', { NOMBRE: 'DUEÑO', PIN: pin(), ROL: 'DUENO', ACTIVO: 'SI' });
    us = leer_('USUARIOS_APP');
  }
  Logger.log('Listo. Carpeta de archivos: ' + carpeta.getUrl());
  Logger.log('Usuarios y PIN: ' + us.map(u => u.NOMBRE + ' (' + u.ROL + '): ' + u.PIN).join(' · '));
  return 'Instalación completa';
}
