/* ============================================================================
   SISTEMA DE GESTIÓN DOCUMENTAL (SGD) — script compartido
   ----------------------------------------------------------------------------
   Estas pantallas están pensadas para EMBEBERSE dentro de index.html (padre).
   Todas las operaciones de SharePoint se piden al padre por postMessage
   (puente SP_BRIDGE), igual que Cotizador / BPM Dashboard. En modo standalone
   (abierto suelto para pruebas) el puente no existe y se usan datos de respaldo.
   ============================================================================ */

window.SGD = window.SGD || {};

/* ---------- Cliente del puente SP (postMessage al padre index.html) ---------- */
// Protocolo:  { type:'SP_BRIDGE_REQUEST',  callId, op, args }
//         →   { type:'SP_BRIDGE_RESPONSE', callId, error, result }
window.SGD.embebido = function () { return window.parent && window.parent !== window; };

window.SGD.sp = function (op, args) {
  return new Promise(function (resolve, reject) {
    if (!window.SGD.embebido()) { reject(new Error('standalone')); return; }
    var callId = 'sgd_' + Date.now() + '_' + Math.random().toString(36).slice(2);
    function onMsg(e) {
      if (!e.data || e.data.type !== 'SP_BRIDGE_RESPONSE' || e.data.callId !== callId) return;
      window.removeEventListener('message', onMsg);
      if (e.data.error) reject(new Error(e.data.error)); else resolve(e.data.result);
    }
    window.addEventListener('message', onMsg);
    window.parent.postMessage({ type: 'SP_BRIDGE_REQUEST', callId: callId, op: op, args: (args || {}) }, '*');
    setTimeout(function () { window.removeEventListener('message', onMsg); reject(new Error('timeout')); }, 15000);
  });
};

/* ---------- Handshake con el padre (index.html) al embeber ----------
   Cada pantalla del SGD, al cargar, pide al padre el usuario y el nivel de
   acceso. El padre responde con un mensaje { type:'SGD_AUTH', user, nivel }.
   Así funciona aunque el iframe navegue internamente (solicitud, atención…). */
window.addEventListener('message', function (e) {
  var d = e.data;
  if (!d || d.type !== 'SGD_AUTH') return;
  if (d.user) window.SGD.setUsuario(d.user);
  window.SGD.setNivelAtencion(d.nivel || '');
  // Aviso a la pantalla activa (p.ej. seguimiento) de que ya hay usuario en sesión
  if (typeof window.SGD._onAuth === 'function') { try { window.SGD._onAuth(); } catch (_) {} }
});
function sgdPedirAuthAlPadre() {
  if (window.SGD.embebido()) {
    try { window.parent.postMessage({ type: 'SGD_READY' }, '*'); } catch (e) {}
  }
}

/* ---------- Usuario en sesión (lo inyecta el padre al embeber) ---------- */
// El padre (index.html) debe llamar:  window.SGD.setUsuario({ nombre:'...', correo:'...' })
window.SGD._usuario = window.SGD._usuario || null;
window.SGD.setUsuario = function (u) {
  window.SGD._usuario = u || null;
  try { if (typeof sgdRellenarUsuario === 'function') sgdRellenarUsuario(); } catch (e) {}
};
window.SGD.getUsuario = function () {
  if (window.SGD._usuario) return window.SGD._usuario;
  // Respaldo standalone (solo pruebas): permite fijar por URL ?nombre=&correo=
  try {
    var q = new URLSearchParams(location.search);
    var n = q.get('nombre'), c = q.get('correo');
    if (n || c) return { nombre: n || '', correo: c || '' };
  } catch (e) {}
  return null;
};

/* ---------- Trazabilidad / auditoría (norma) ----------
   Estampa en el objeto de campos QUIÉN y CUÁNDO realiza el cambio.
   Se usa en TODAS las escrituras del SGD para dejar rastro auditable.
   Columnas SP requeridas en cada lista:  Usuario_Modifica (texto),  Ultima_Modificacion (fecha/hora o texto). */
window.SGD.stampAuditoria = function (fields) {
  fields = fields || {};
  var u = (window.SGD.getUsuario && window.SGD.getUsuario()) || null;
  var quien = u ? (u.nombre || u.correo || '') : '';
  if (u && u.nombre && u.correo) quien = u.nombre + ' (' + u.correo + ')';
  fields.Usuario_Modifica = quien || 'Desconocido';
  fields.Ultima_Modificacion = new Date().toISOString();
  return fields;
};

/* ---------- Nombres de las listas de SharePoint (dar de alta en SP) ---------- */
//  SGD_TiposDocumento : columna Title = tipo de documento (Procedimiento, Política, Formato, Diagrama, ...)
//  SGD_Solicitudes    : columnas para las solicitudes de documentación (ver documentación aparte)
window.SGD.LISTA_TIPOS = 'SGD_TiposDocumento';
window.SGD.LISTA_SOLICITUDES = 'SGD_Solicitudes';
//  SGD_NotificacionesCalidad : columnas Title = nombre del analista, Correo = correo electrónico
window.SGD.LISTA_NOTIF_CALIDAD = 'SGD_NotificacionesCalidad';
//  SGD_Bitacora : registro de auditoría (un renglón por acción). Columnas:
//    Title (texto), Accion (texto), Entidad (texto), Entidad_Id (texto),
//    Detalle (varias líneas), Usuario (texto), Correo (texto),
//    Pantalla (texto), Fecha_Evento (fecha y hora)
window.SGD.LISTA_BITACORA = 'SGD_Bitacora';

/* ---------- Bitácora / trazabilidad global ----------
   Escribe un renglón de auditoría en SGD_Bitacora. Es "a prueba de fallos":
   nunca interrumpe el flujo principal (captura errores y no lanza).
   Uso:  window.SGD.bitacora('Registrar solicitud', 'SGD_Solicitudes', 'Procedimiento X', { id: 12, estatus:'Abierta' }); */
window.SGD.bitacora = function (accion, entidad, detalle, extra) {
  try {
    extra = extra || {};
    var u = (window.SGD.getUsuario && window.SGD.getUsuario()) || null;
    var nombre = u ? (u.nombre || '') : '';
    var correo = u ? (u.correo || '') : '';
    var pantalla = extra.pantalla || (document && document.title) || (location && location.pathname) || '';
    var fields = {
      Title: (accion || 'Acción') + (entidad ? (' · ' + entidad) : ''),
      Accion: accion || '',
      Entidad: entidad || '',
      Entidad_Id: (extra.id !== undefined && extra.id !== null) ? String(extra.id) : '',
      Detalle: detalle || '',
      Usuario: nombre,
      Correo: correo,
      Pantalla: pantalla,
      Fecha_Evento: new Date().toISOString()
    };
    if (!window.SGD.embebido()) {
      // Standalone (pruebas): sin SP; solo dejamos rastro en consola
      try { console.log('[SGD_Bitacora]', fields); } catch (_) {}
      return Promise.resolve();
    }
    return window.SGD.sp('create', { lista: window.SGD.LISTA_BITACORA, fields: fields })
      .catch(function (e) { try { console.warn('[SGD_Bitacora] no se pudo registrar:', e && e.message); } catch (_) {} });
  } catch (e) {
    try { console.warn('[SGD_Bitacora] error:', e && e.message); } catch (_) {}
    return Promise.resolve();
  }
};

// Tipos por defecto (respaldo cuando no hay conexión con SP / standalone)
window.SGD.TIPOS_DEFAULT = ['Procedimiento', 'Política', 'Formato', 'Diagrama', 'Instructivo', 'Manual'];

/* ============================================================================
   PANTALLA "actualizar" (ya existente)
   ============================================================================ */
function abrirActualizar(procedimiento) {
  localStorage.setItem("procedimiento", procedimiento);
  window.location.href = "actualizar.html";
}

/* ============================================================================
   PERMISOS / ROL — botón "ATENCIÓN A SOLICITUDES DE DOCUMENTOS"
   Solo visible para Calidad (acceso Completo o Lectura). El padre inyecta el
   nivel con window.SGD.setNivelAtencion('completo'|'lectura'|'').
   ============================================================================ */
window.SGD.getNivelAtencion = function () {
  if (window.SGD._nivelAtencion) return window.SGD._nivelAtencion;
  try {
    var r = (new URLSearchParams(location.search).get('rol') || '').toLowerCase();
    if (r === 'completo' || r === 'lectura') return r;
    if (r === 'calidad') return 'completo';
  } catch (e) {}
  return '';
};
window.SGD.setNivelAtencion = function (nivel) {
  window.SGD._nivelAtencion = (nivel || '').toLowerCase();
  sgdAplicarPermisos();
};
function sgdPuedeAtender() {
  // El botón "ATENCIÓN A SOLICITUDES" solo se muestra con acceso COMPLETO.
  // Con acceso de solo Lectura NO se muestra.
  return window.SGD.getNivelAtencion() === 'completo';
}
function sgdAplicarPermisos() {
  var btn = document.getElementById('btn-atencion-solicitudes');
  if (btn) btn.style.display = sgdPuedeAtender() ? '' : 'none';
}

/* ============================================================================
   PANTALLA PRINCIPAL — buscador, visor de PDF y descarga de copia no controlada
   ============================================================================ */

// Filtra la tabla por el NOMBRE del documento conforme se escribe
function sgdFiltrarTabla(txt) {
  var q = (txt || '').toLowerCase().trim();
  var filas = document.querySelectorAll('.tabla1 tbody tr');
  filas.forEach(function (tr) {
    var celdaNombre = tr.querySelector('.nombre');
    var nombre = celdaNombre ? (celdaNombre.textContent || '').toLowerCase() : '';
    tr.style.display = (!q || nombre.indexOf(q) >= 0) ? '' : 'none';
  });
}

// Abre el visor de PDF (solo lectura; se ocultan las barras de descarga/impresión del visor)
function sgdVerDocumento(url, titulo) {
  if (!url) { alert('El documento aún no está disponible.'); return; }
  var modal = document.getElementById('sgd-visor-modal');
  if (!modal) return;
  var frame = document.getElementById('sgd-visor-frame');
  var tit = document.getElementById('sgd-visor-titulo');
  if (tit) tit.textContent = titulo || 'Documento';
  // #toolbar=0&navpanes=0 oculta los controles (descargar/imprimir) del visor nativo del navegador
  var sep = url.indexOf('#') >= 0 ? '&' : '#';
  if (frame) frame.src = url + sep + 'toolbar=0&navpanes=0&statusbar=0&view=FitH';
  modal.style.display = 'flex';
  document.body.style.overflow = 'hidden';
  if (window.SGD.bitacora) window.SGD.bitacora('Ver documento', 'Documento', titulo || '', {});
}
function sgdCerrarVisor() {
  var modal = document.getElementById('sgd-visor-modal');
  var frame = document.getElementById('sgd-visor-frame');
  if (frame) frame.src = 'about:blank';
  if (modal) modal.style.display = 'none';
  document.body.style.overflow = '';
}

// Descarga la "Copia NO controlada" del documento
function sgdDescargarCopia(url, nombre) {
  if (!url) { alert('El documento aún no está disponible para descargar.'); return; }
  var a = document.createElement('a');
  a.href = url;
  a.download = ((nombre || 'documento').replace(/[^\w\-. áéíóúÁÉÍÓÚñÑ]/g, '_')) + ' (COPIA NO CONTROLADA).pdf';
  a.target = '_blank';
  a.rel = 'noopener';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  if (window.SGD.bitacora) window.SGD.bitacora('Descargar copia NO controlada', 'Documento', nombre || '', {});
}

/* ============================================================================
   PANTALLA "solicitud" — tipos de documento, usuario y registro en SharePoint
   ============================================================================ */

// Rellena el nombre y correo del solicitante con el usuario en sesión
function sgdRellenarUsuario() {
  var u = window.SGD.getUsuario();
  var inSolic = document.getElementById('sol-solicitante');
  var inCorreo = document.getElementById('sol-correo');
  if (inSolic) inSolic.value = u && u.nombre ? u.nombre : '';
  if (inCorreo) inCorreo.value = u && u.correo ? u.correo : '';
}

// Carga los tipos de documento desde SP (o usa los de respaldo) al abrir la solicitud
async function sgdCargarTipos() {
  var sel = document.getElementById('sol-tipo');
  if (!sel) return;
  var tipos = null;
  try {
    var r = await window.SGD.sp('list', { lista: window.SGD.LISTA_TIPOS, queryParams: '$expand=fields&$top=200' });
    var items = (r && r.value) ? r.value : [];
    tipos = items
      .map(function (i) { var f = i.fields || i; return (f.Title || f.Tipo || '').toString().trim(); })
      .filter(Boolean);
  } catch (e) { /* standalone / sin conexión → respaldo */ }
  if (!tipos || !tipos.length) tipos = window.SGD.TIPOS_DEFAULT.slice();
  tipos.sort(function (a, b) { return a.localeCompare(b); });
  sel.innerHTML = '<option value="">Seleccione un tipo de documento</option>' +
    tipos.map(function (t) { return '<option>' + t + '</option>'; }).join('');
}

// Habilita/deshabilita la agenda según el check "Se requiere revisión presencial"
function sgdToggleRevisionPresencial() {
  var chk = document.getElementById('sol-revision-presencial');
  var req = !!(chk && chk.checked);
  var fecha = document.getElementById('sol-fecha');
  var hora = document.getElementById('sol-hora');
  var agenda = document.getElementById('sol-agenda');
  if (fecha) { fecha.disabled = !req; if (!req) fecha.value = ''; }
  if (hora) { hora.disabled = !req; if (!req) hora.value = ''; }
  if (agenda) { if (req) agenda.classList.remove('deshabilitada'); else agenda.classList.add('deshabilitada'); }
}

// Registra la solicitud de documentación
async function sgdRegistrarSolicitud(ev) {
  if (ev) ev.preventDefault();
  var err = document.getElementById('sol-error');
  if (err) err.textContent = '';

  var val = function (id) { var el = document.getElementById(id); return el ? el.value.trim() : ''; };
  var chkPres = document.getElementById('sol-revision-presencial');
  var requierePresencial = !!(chkPres && chkPres.checked);
  var datos = {
    titulo: val('sol-titulo'),
    tipo: val('sol-tipo'),
    solicitante: val('sol-solicitante'),
    correo: val('sol-correo'),
    departamento: val('sol-departamento'),
    objetivo: val('sol-objetivo'),
    alcance: val('sol-alcance'),
    responsables: val('sol-responsables'),
    descripcion: val('sol-descripcion'),
    fecha: val('sol-fecha'),
    hora: val('sol-hora')
  };

  // Validación de obligatorios (la fecha/hora solo son obligatorias si se requiere revisión presencial)
  var obligatorios = [
    ['titulo', 'Título del Documento'],
    ['tipo', 'Tipo de documento'],
    ['objetivo', 'Objetivo del Documento'],
    ['alcance', 'Alcance del Documento'],
    ['responsables', 'Responsables del Documento'],
    ['descripcion', 'Descripción General del Documento']
  ];
  if (requierePresencial) {
    obligatorios.push(['fecha', 'Fecha estimada para revisión']);
    obligatorios.push(['hora', 'Hora estimada para revisión']);
  }
  var faltan = obligatorios.filter(function (c) { return !datos[c[0]]; }).map(function (c) { return c[1]; });
  if (faltan.length) {
    if (err) err.textContent = 'Completa los campos obligatorios: ' + faltan.join(', ') + '.';
    else alert('Completa los campos obligatorios: ' + faltan.join(', '));
    return;
  }

  // Estatus de arranque para solicitudes de Nuevo Ingreso:
  //  - Con revisión presencial  -> "Pendiente" (queda a la espera de confirmar la cita)
  //  - Sin revisión presencial  -> "Abierta" (lista para atenderse)
  var estatusInicial = requierePresencial ? 'Pendiente' : 'Abierta';

  var btn = document.getElementById('sol-registrar');
  if (btn) { btn.disabled = true; btn.textContent = 'Registrando…'; }

  // Campos para la lista SGD_Solicitudes en SharePoint
  var fields = {
    Title: datos.titulo,
    Tipo_Documento: datos.tipo,
    Tipo_Modificacion: 'Nuevo Ingreso',
    Solicitante: datos.solicitante,
    Correo: datos.correo,
    Objetivo: datos.objetivo,
    Alcance: datos.alcance,
    Responsables: datos.responsables,
    Descripcion: datos.descripcion,
    Departamento: datos.departamento,
    Requiere_Revision: requierePresencial ? 'Sí' : 'No',
    Fecha_Revision: (requierePresencial && datos.fecha) ? (datos.fecha + 'T00:00:00Z') : '',
    Hora_Revision: requierePresencial ? (datos.hora || '') : '',
    Estatus: estatusInicial,
    Fecha_Solicitud: new Date().toISOString()
  };
  // Trazabilidad: quién y cuándo registró la solicitud
  window.SGD.stampAuditoria(fields);

  var nuevoId = null;
  try {
    var cr = await window.SGD.sp('create', { lista: window.SGD.LISTA_SOLICITUDES, fields: fields });
    nuevoId = (cr && cr.id !== undefined) ? cr.id : null;
  } catch (e) {
    if (!window.SGD.embebido()) {
      // Standalone (pruebas): no hay SP; guardamos localmente y continuamos el flujo
      try {
        var pend = JSON.parse(localStorage.getItem('sgd_solicitudes_local') || '[]');
        pend.push(fields); localStorage.setItem('sgd_solicitudes_local', JSON.stringify(pend));
      } catch (_) {}
    } else {
      if (btn) { btn.disabled = false; btn.textContent = 'Registrar'; }
      if (err) err.textContent = 'No se pudo registrar la solicitud. Intenta de nuevo.';
      return;
    }
  }

  // Bitácora de auditoría
  if (window.SGD.bitacora) {
    window.SGD.bitacora('Registrar solicitud', 'SGD_Solicitudes', datos.titulo,
      { id: nuevoId, pantalla: 'Solicitud de documentación' });
  }

  sgdMostrarModalOK();
  if (btn) { btn.disabled = false; btn.textContent = 'Registrar'; }
}

function sgdMostrarModalOK() {
  var m = document.getElementById('sol-modal-ok');
  if (m) { m.style.display = 'flex'; document.body.style.overflow = 'hidden'; }
  else { alert('Su solicitud ha sido registrada. La confirmación de la visita solicitada le será confirmada vía correo electrónico'); sgdIrAPrincipal(); }
}
function sgdCerrarModalOK() {
  var m = document.getElementById('sol-modal-ok');
  if (m) m.style.display = 'none';
  document.body.style.overflow = '';
  sgdIrAPrincipal();
}
function sgdIrAPrincipal() {
  window.location.href = 'Sistema_Gestion_Documental.html';
}

/* ---------- Arranque por pantalla ---------- */
document.addEventListener('DOMContentLoaded', function () {
  sgdPedirAuthAlPadre();   // pide usuario + nivel al padre (index.html) si está embebido
  sgdAplicarPermisos();
  // Fecha mínima para revisión = mañana (no puede ser el mismo día)
  var fEl = document.getElementById('sol-fecha');
  if (fEl) {
    var m = new Date(); m.setDate(m.getDate() + 1);
    fEl.min = m.getFullYear() + '-' + String(m.getMonth() + 1).padStart(2, '0') + '-' + String(m.getDate()).padStart(2, '0');
  }
  if (document.getElementById('sol-tipo')) sgdCargarTipos();
  if (document.getElementById('sol-solicitante')) sgdRellenarUsuario();
});
