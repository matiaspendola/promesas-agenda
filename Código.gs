// ============================================================
// AGENDA PROMESAS CHILE · Google Apps Script Backend
// ============================================================
// INSTRUCCIONES:
// 1. Ve a script.google.com → Nuevo proyecto
// 2. Borra el contenido y pega TODO este código
// 3. Actualiza CONFIG con los correos reales
// 4. Crea un Google Sheets y pega su ID en sheetId
// 5. Implementar → Nueva implementación → Aplicación web
//    · Ejecutar como: "Yo"
//    · Quién tiene acceso: "Cualquier persona"
// 6. Copia la URL generada y pégala en el sistema web
// ============================================================

// ── CONFIGURACIÓN ────────────────────────────────────────────
// IMPORTANTE (modelo de calendarios):
//   El evento se crea en el calendario del profesional. Para que funcione,
//   cada profesional debe COMPARTIR su Google Calendar con la cuenta que
//   ejecuta este script (la cuenta del sistema), con permiso
//   "Hacer cambios en los eventos". El calendarId es su propio email.
//   Si se deja calendarId vacío, se usa el email del profesional.
const CONFIG = {
  profesionales: {
    kine: {
      nombre: 'Kin. Katalina Camino',
      email:  'kine.katalinacamino@gmail.com',   // ← email real (fuente: hoja Usuarios)
      calendarId: '',                            // ← vacío = usa su email como calendario
    },
    psico: {
      nombre: 'Ps. Mario Pidal',
      email:  'pidalmario@gmail.com',        // ← email real
      calendarId: '',
    },
    nutri: {
      nombre: 'Nut. Josefina Enríquez',
      email:  'josefina.enri.sch@gmail.com', // ← email real
      calendarId: '',
    },
    medico: {
      nombre: 'Dr. Juan Manuel Guzmán',
      email:  'juanma.guzmanh@gmail.com',      // ← email real
      calendarId: '',
    },
    profis: {
      nombre: 'Prof. Matías Péndola',
      email:  'matias.pendola@gmail.com',     // ← email real
      calendarId: '3a79537bf1bafcc39e1a21feed51e2946785ee960f6e4f4abda4dd7b22ca552a@group.calendar.google.com', // Calendario "Promesas Chile PF"
    },
  },
  // Crear Google Sheets en drive.google.com
  // El ID está en la URL: .../spreadsheets/d/[ID_AQUI]/edit
  sheetId: '1NzsiSKMu5gTipwfhc1dDP3qAwPEV0L9zZyo5uvjUumA',

  // Correo del administrador (recibe copia de errores)
  adminEmail: 'admin.promesaschile@gmail.com',

  // Nombre que aparece en los emails enviados
  nombrePrograma: 'Promesas Chile · CAR Náutico Valdivia · IND Los Ríos',

  // Remitente de los correos. Dejar VACÍO para enviar desde la cuenta que
  // ejecuta el script (recomendado tras migrar a la cuenta del sistema).
  // Si se ejecuta bajo una cuenta personal y se quiere usar un alias verificado
  // ("Enviar como" en Gmail), poner aquí esa dirección.
  remitente: '',

  // Ubicación que aparece en los eventos de Calendar
  lugar: 'CAR Náutico Valdivia, Valdivia, Los Ríos',

  // Duración fija de cada atención en minutos
  duracionMinutos: 60,
};

// Client ID de Google (debe coincidir con el data-client_id del frontend).
// Se usa para verificar que el id_token fue emitido para ESTA aplicación.
const CLIENT_ID = '1011451823200-sntlql8cvr0l1i2fs5h15b3e5sndjp3i.apps.googleusercontent.com';
const SESSION_TTL_MS = 30 * 24 * 60 * 60 * 1000; // 30 días

// ── FUNCIÓN DE PRUEBA (ejecutar manualmente desde editor para autorizar MailApp) ──
function probarCorreo() {
  const dest = 'matias.pendola@gmail.com';
  MailApp.sendEmail(dest, 'Prueba Promesas Chile - autorizacion correo', 'Permiso de envio activo desde promesaschilelosrios@gmail.com');
  return { ok: true, msg: 'Correo enviado a ' + dest };
}

// ── AUTORIZAR TODOS LOS PERMISOS ─────────────────────────────
// EJECUTAR ESTA FUNCIÓN UNA VEZ desde el editor (como la cuenta del sistema).
// Toca cada servicio para que Google pida y conceda TODOS los permisos,
// incluido UrlFetchApp (necesario para verificar el login con Google).
function autorizarPermisos() {
  SpreadsheetApp.openById(CONFIG.sheetId).getName();            // Sheets
  CalendarApp.getDefaultCalendar().getName();                  // Calendar
  MailApp.getRemainingDailyQuota();                            // Mail
  UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=x', { muteHttpExceptions: true }); // External request
  Logger.log('Todos los permisos autorizados correctamente.');
  return { ok: true, msg: 'Permisos autorizados' };
}

// ── SEGURIDAD: VERIFICACIÓN DE IDENTIDAD Y SESIÓN ────────────
// Secreto para firmar los tokens de sesión. Se autogenera una vez y queda
// guardado en las propiedades del script (no en el código).
function _sessionSecret() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty('SESSION_SECRET');
  if (!s) {
    s = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('SESSION_SECRET', s);
  }
  return s;
}

function _b64url(bytes) {
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

// Verifica el id_token de Google contra el endpoint oficial de Google.
// Devuelve el email verificado (en minúsculas) o null si no es válido.
function verificarIdToken(idToken) {
  if (!idToken) return null;
  try {
    const resp = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    const info = JSON.parse(resp.getContentText());
    if (info.aud !== CLIENT_ID) return null;                       // emitido para otra app
    if (info.email_verified !== 'true' && info.email_verified !== true) return null;
    return String(info.email || '').toLowerCase();
  } catch (e) {
    return null;
  }
}

// Emite un token de sesión firmado (HMAC). Formato: base64url(payload).base64url(firma)
function emitirToken(email) {
  const payload = JSON.stringify({ email: String(email).toLowerCase(), exp: Date.now() + SESSION_TTL_MS });
  const payloadB64 = _b64url(Utilities.newBlob(payload).getBytes());
  const sig = Utilities.computeHmacSha256Signature(payloadB64, _sessionSecret());
  return payloadB64 + '.' + _b64url(sig);
}

// Verifica un token de sesión. Devuelve el email si la firma y la expiración
// son válidas; null en caso contrario.
function verificarToken(token) {
  if (!token || String(token).indexOf('.') < 0) return null;
  const parts = String(token).split('.');
  const payloadB64 = parts[0], sigB64 = parts[1];
  const expected = _b64url(Utilities.computeHmacSha256Signature(payloadB64, _sessionSecret()));
  if (expected !== sigB64) return null;                            // firma inválida
  try {
    const json = Utilities.newBlob(Utilities.base64DecodeWebSafe(payloadB64)).getDataAsString();
    const payload = JSON.parse(json);
    if (!payload.exp || Date.now() > payload.exp) return null;     // expirado
    return String(payload.email).toLowerCase();
  } catch (e) {
    return null;
  }
}

// True si el email tiene rol admin en cualquiera de sus filas de Usuarios.
function esAdmin(email) {
  return buscarTodosUsuarios(email).some(u => u.rol === 'admin');
}

// ── ENDPOINT PRINCIPAL (POST) ─────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // TEMPORAL: verifica todos los calendarios sin token.
    if (data.tipo === 'check_cals') {
      const todos = CalendarApp.getAllCalendars().map(c => c.getId() + ' → ' + c.getName());
      const psico = getProfesional('psico');
      let psicoAccesible = false;
      try { psicoAccesible = !!CalendarApp.getCalendarById(psico.email); } catch(e) {}
      return ok({ todos, psicoEmail: psico.email, psicoAccesible });
    }

    // El login es el único endpoint sin token previo (verifica el id_token de Google).
    if (data.tipo === 'validar_usuario') return ok(validarUsuario(data));

    // Todos los demás endpoints exigen un token de sesión válido (no spoofeable).
    const email = verificarToken(data.token);
    if (!email) return err('Sesión no válida o expirada. Vuelve a iniciar sesión.');
    const usuario = buscarUsuario(email);
    if (!usuario) return err('No autorizado');
    data._email = email;          // email verificado por el servidor (no viene del cliente)

    // Restricciones por rol (admin se evalúa entre todos los roles del usuario)
    const _admin = esAdmin(email);
    if (data.tipo === 'reagendar_cita' && !_admin) return err('Solo el administrador puede reagendar');
    if (data.tipo === 'suspender_cita' && usuario.rol === 'tecnico') return err('Los técnicos no pueden suspender citas');
    if (data.tipo === 'guardar_disponibilidad' && usuario.rol === 'tecnico') return err('Los técnicos no pueden modificar disponibilidad');
    if (data.tipo === 'completar_cita' && usuario.rol === 'tecnico') return err('Los técnicos no pueden cerrar citas');
    // Endpoints solo-admin (datos sensibles / operaciones destructivas)
    if (['leer_hoja', 'reset_citas', 'verificar_calendarios'].indexOf(data.tipo) !== -1 && !_admin) {
      return err('Solo el administrador puede realizar esta acción');
    }

    switch (data.tipo) {
      case 'nueva_cita':          return ok(agendarCita(data));
      case 'suspender_cita':      return ok(suspenderCita(data));
      case 'reagendar_cita':      return ok(reagendarCita(data));
      case 'guardar_disponibilidad': return ok(guardarDisponibilidad(data));
      case 'get_disponibilidad':  return ok(getDisponibilidad(data));
      case 'get_disponibilidad_config': return ok(getDisponibilidadConfig(data));
      case 'setup_usuarios':      return ok(setupUsuarios());
      case 'leer_hoja':           return ok(leerHoja(data));
      case 'get_citas':           return ok(getCitas(data));
      case 'completar_cita':      return ok(completarCita(data));
      case 'reset_citas':         return ok(resetCitas(data));
      case 'verificar_calendarios': return ok(verificarCalendarios(data));
      default: return err('Tipo desconocido: ' + data.tipo);
    }
  } catch (e) {
    return err(e.toString());
  }
}

// ── ENDPOINT GET (para consultas) ─────────────────────────────
function doGet(e) {
  // Sin datos sensibles por GET (no hay token). Solo señal de vida.
  return ok({ msg: 'API Agenda Promesas Chile activa ✅', version: '2.0' });
}

// ── AGENDAR CITA ─────────────────────────────────────────────
function agendarCita(data) {
  const prof = getProfesional(data.profesionalKey);
  if (!prof) throw new Error('Profesional no encontrado: ' + data.profesionalKey);

  const inicio = new Date(data.fechaHoraInicio);
  const fin    = new Date(inicio.getTime() + CONFIG.duracionMinutos * 60 * 1000);

  // Validar fecha/hora
  if (isNaN(inicio.getTime())) throw new Error('Fecha/hora inválida');
  if (inicio.getTime() < Date.now() - 60 * 1000) throw new Error('No se puede agendar en el pasado');

  // Validar contra disponibilidad declarada del profesional
  const fechaSolo = data.fechaHoraInicio.split('T')[0];
  const horaSolo  = (data.fechaHoraInicio.split('T')[1] || '').slice(0, 5);
  const disp = getDisponibilidad({ profKey: data.profesionalKey, fecha: fechaSolo });
  if (!disp.trabaja) throw new Error('El profesional no tiene disponibilidad configurada para ese día');

  // El horario debe caer dentro de alguna franja válida (libre)
  if (disp.slotsLibres.indexOf(horaSolo) === -1) {
    // Si el slot existe pero está ocupado, mensaje específico; si no existe, fuera de franja
    const existe = (disp.slots || []).some(s => s.hora === horaSolo);
    throw new Error(existe ? 'Ese horario ya está ocupado. Actualiza y elige otro.' : 'Ese horario está fuera de la disponibilidad del profesional');
  }

  const ss = SpreadsheetApp.openById(CONFIG.sheetId);

  // 1. Crear evento en el Google Calendar del profesional (su email)
  const cal = _calendarioDe(prof);

  const titulo = `[Promesas Chile] ${data.atletaNombre} — ${data.profesionalLabel}`;
  const descripcion = [
    `Atleta: ${data.atletaNombre}`,
    `Polo deportivo: ${data.polo}`,
    `Técnico: ${data.tecnicoNombre} (${data.tecnicoEmail})`,
    `Apoderado: ${data.apoderadoNombre} (${data.apoderadoEmail})`,
    `Motivo: ${data.motivo || 'Sin especificar'}`,
    `Tipo de bloque: ${data.tipoBloque === 'libre' ? '🟢 Libre (solicitado por técnico)' : '🟣 Reservado (prof/admin)'}`,
    ``,
    `Registrado: ${new Date().toLocaleString('es-CL')}`,
  ].join('\n');

  const evento = cal.createEvent(titulo, inicio, fin, {
    description: descripcion,
    location: CONFIG.lugar,
    guests: data.apoderadoEmail, // invitar al apoderado opcionalmenye
    sendInvites: false,          // cambiar a true si quieres invitar
  });

  // 2. Email al técnico (solo aviso de solicitud registrada)
  if (data.tecnicoEmail) enviarCorreo(
    data.tecnicoEmail,
    `[Promesas Chile] Solicitud de hora registrada — ${data.atletaNombre}`,
    emailTecnicoHtml(data, prof)
  );

  // 3. Email al apoderado (confirmación completa)
  if (data.apoderadoEmail) enviarCorreo(
    data.apoderadoEmail,
    `[Promesas Chile] Confirmación de cita para ${data.atletaNombre}`,
    emailApoderadoHtml(data, prof, fin)
  );

  // 4. Registrar en Google Sheets
  const eventoId = evento.getId();
  registrarEnSheet({
    ...data,
    eventoId,
    estado: 'Confirmada',
    fechaRegistro: new Date().toISOString(),
  });

  return { ok: true, eventoId, msg: 'Cita agendada correctamente' };
}

// ── SUSPENDER CITA ───────────────────────────────────────────
function suspenderCita(data) {
  // 1. Eliminar evento del calendario
  if (data.eventoId) {
    try {
      const prof = getProfesional(data.profesionalKey);
      if (prof) {
        const cal = _calendarioDe(prof);
        const evento = cal.getEventById(data.eventoId);
        if (evento) evento.deleteEvent();
      }
    } catch (e) { /* evento ya no existe */ }
  }

  // 2. Email al técnico — aviso de suspensión
  if (data.tecnicoEmail) enviarCorreo(
    data.tecnicoEmail,
    `[Promesas Chile] Cita suspendida — ${data.atletaNombre}`,
    emailSuspensionTecnicoHtml(data)
  );

  // 3. Email al apoderado — notificación con disculpas
  if (data.apoderadoEmail) enviarCorreo(
    data.apoderadoEmail,
    `[Promesas Chile] Aviso de suspensión — ${data.atletaNombre}`,
    emailSuspensionApoderadoHtml(data)
  );

  // 4. Actualizar estado en Google Sheets
  actualizarEstadoSheet(data.eventoId, 'Suspendida', data.motivoSuspension);

  return { ok: true, msg: 'Cita suspendida y notificaciones enviadas' };
}

// ── REAGENDAR CITA (solo admin) ──────────────────────────────
function reagendarCita(data) {
  // Suspende la antigua sin notificar de nuevo (ya se notificó)
  if (data.eventoId) {
    try {
      const prof = getProfesional(data.profesionalKey);
      if (prof) {
        const cal = _calendarioDe(prof);
        const evento = cal.getEventById(data.eventoId);
        if (evento) evento.deleteEvent();
      }
    } catch (e) {}
  }
  // Agenda la nueva con la nueva fecha/hora
  const nuevaData = { ...data, fechaHoraInicio: data.nuevaFechaHoraInicio };
  return agendarCita(nuevaData);
}

// ── DISPONIBILIDAD SEMANAL ────────────────────────────────────
// Modelo: cada profesional define su horario laboral por día de la semana.
// Hoja "Disponibilidad": Profesional | DíaSemana(0-6) | Hora Inicio | Hora Fin | Activo | Fecha Registro
// DíaSemana: 0=Domingo, 1=Lunes ... 6=Sábado

function _hojaDisponibilidad(ss) {
  let sheet = ss.getSheetByName('Disponibilidad');
  if (!sheet) {
    sheet = ss.insertSheet('Disponibilidad');
    sheet.appendRow(['Profesional', 'Día Semana', 'Hora Inicio', 'Hora Fin', 'Activo', 'Fecha Registro']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#003087').setFontColor('white');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function guardarDisponibilidad(data) {
  // data: { profKey, dias: [ {dia:1, inicio:'09:00', fin:'13:00', activo:true}, ... ] }
  if (!data.profKey) throw new Error('Falta el identificador del profesional (profKey)');
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = _hojaDisponibilidad(ss);

  // Borrar filas previas de este profesional (se sobrescribe toda su config)
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][0] === data.profKey) sheet.deleteRow(i + 1);
  }

  // Insertar solo los días activos. Las horas se guardan como MINUTOS (entero)
  // para que Google Sheets no las convierta en valores de tiempo con desfase horario.
  const now = new Date().toISOString();
  (data.dias || []).forEach(d => {
    if (d.activo) sheet.appendRow([data.profKey, d.dia, _toMin(d.inicio), _toMin(d.fin), 'si', now]);
  });
  return { ok: true, msg: 'Disponibilidad guardada' };
}

function getDisponibilidadConfig(data) {
  // Devuelve la config semanal del profesional (para el editor)
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('Disponibilidad');
  if (!sheet) return { ok: true, dias: [] };
  const rows = sheet.getDataRange().getValues();
  const dias = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.profKey) {
      dias.push({
        dia: Number(rows[i][1]),
        inicio: _minToHHMM(rows[i][2]),
        fin: _minToHHMM(rows[i][3]),
        activo: String(rows[i][4]).toLowerCase() === 'si',
      });
    }
  }
  return { ok: true, dias };
}

function getDisponibilidad(data) {
  // Calcula los slots de 30 min de un profesional en una fecha concreta.
  // Soporta múltiples franjas por día (ej: 06:00-08:00 y 16:00-18:00).
  // data: { profKey, fecha:'YYYY-MM-DD' }
  if (!data.fecha) return { ok: true, trabaja: false, slots: [], slotsLibres: [], ocupados: [] };
  const dow = new Date(data.fecha + 'T00:00:00').getDay(); // 0=Dom .. 6=Sáb

  const cfg = getDisponibilidadConfig({ profKey: data.profKey });
  const franjas = cfg.dias.filter(d => d.dia === dow && d.activo);
  if (!franjas.length) return { ok: true, trabaja: false, slots: [], slotsLibres: [], ocupados: [] };

  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  const ocupados = _citasOcupadas(ss, data.profKey, data.fecha);
  const bloqueados = new Set();
  ocupados.forEach(h => { bloqueados.add(h); bloqueados.add(_addMin(h, 30)); });

  // Generar slots de 30 min de todas las franjas, ordenados y sin duplicar
  franjas.sort((a, b) => a.inicio < b.inicio ? -1 : 1);
  const vistos = {};
  const slots = [];
  franjas.forEach(fr => {
    _slots30(fr.inicio, fr.fin).forEach(h => {
      if (_addMin(h, 60) > fr.fin) return;   // el bloque de 60 min no cabe en la franja
      if (vistos[h]) return;                 // evitar duplicado si las franjas se solapan
      vistos[h] = true;
      const libre = !bloqueados.has(h) && !bloqueados.has(_addMin(h, 30));
      slots.push({ hora: h, fin: _addMin(h, 60), libre: libre });
    });
  });
  const slotsLibres = slots.filter(s => s.libre).map(s => s.hora);

  return {
    ok: true,
    trabaja: true,
    slots: slots,
    slotsLibres: slotsLibres,
    ocupados: ocupados,
  };
}

// Horas ocupadas por citas vigentes (no suspendidas) de un profesional en una fecha
function _citasOcupadas(ss, profKey, fecha) {
  const sheet = ss.getSheetByName('Citas');
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  const horas = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const estado = String(r[13] || '').toLowerCase();
    if (r[10] === profKey && _fechaTxt(ss, r[1]) === String(fecha) && estado !== 'suspendida' && r[2]) {
      horas.push(_horaTxt(ss, r[2]));
    }
  }
  return horas;
}

function _slots30(inicio, fin) {
  const out = [];
  let t = inicio;
  while (t < fin) { out.push(t); t = _addMin(t, 30); }
  return out;
}

function _addMin(t, m) {
  const [h, mn] = String(t).split(':').map(Number);
  const tot = h * 60 + mn + m;
  return String(Math.floor(tot / 60)).padStart(2, '0') + ':' + String(tot % 60).padStart(2, '0');
}

// Conversión hora <-> minutos (almacenamiento a prueba de zona horaria)
function _toMin(hhmm) {
  const [h, m] = String(hhmm).split(':').map(Number);
  return (h || 0) * 60 + (m || 0);
}
function _minToHHMM(min) {
  min = Number(min) || 0;
  return String(Math.floor(min / 60)).padStart(2, '0') + ':' + String(min % 60).padStart(2, '0');
}

// Normaliza la hora de una celda de Citas a 'HH:MM'. Si Sheets la guardó como
// valor de tiempo (Date), la formatea usando la zona horaria de la planilla.
function _horaTxt(ss, v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, ss.getSpreadsheetTimeZone(), 'HH:mm');
  }
  return String(v).slice(0, 5);
}

// Normaliza la fecha de una celda a 'YYYY-MM-DD'
function _fechaTxt(ss, v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  }
  return String(v).slice(0, 10);
}

// ── GOOGLE SHEETS: REGISTRAR CITA ────────────────────────────
function registrarEnSheet(data) {
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  let sheet = ss.getSheetByName('Citas');
  if (!sheet) {
    sheet = ss.insertSheet('Citas');
    sheet.appendRow([
      'ID Evento','Fecha','Hora','Atleta','Polo',
      'Técnico','Email Técnico','Apoderado','Email Apoderado',
      'Profesional','Disciplina','Tipo Bloque','Motivo','Estado','Fecha Registro'
    ]);
    sheet.getRange(1,1,1,15).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  // Forzar Fecha (B) y Hora (C) como TEXTO para evitar que la zona horaria
  // de la planilla (por defecto Pacific) las desplace al leerlas.
  sheet.getRange('B:C').setNumberFormat('@');
  SpreadsheetApp.flush();
  const fecha = data.fechaHoraInicio ? data.fechaHoraInicio.split('T')[0] : '';
  const hora  = data.fechaHoraInicio ? (data.fechaHoraInicio.split('T')[1]||'').slice(0,5) : '';
  const fila = sheet.getLastRow() + 1;
  sheet.getRange(fila, 1, 1, 15).setValues([[
    data.eventoId, fecha, hora,
    data.atletaNombre, data.polo,
    data.tecnicoNombre, data.tecnicoEmail,
    data.apoderadoNombre, data.apoderadoEmail,
    data.profesionalLabel, data.profesionalKey,
    data.tipoBloque, data.motivo || '',
    data.estado || 'Confirmada', data.fechaRegistro,
  ]]);
}

function actualizarEstadoSheet(eventoId, nuevoEstado, motivo) {
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('Citas');
  if (!sheet) return;
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === eventoId) {
      sheet.getRange(i + 1, 14).setValue(nuevoEstado);
      if (motivo) sheet.getRange(i + 1, 13).setValue(motivo);
      break;
    }
  }
}

// ── CERRAR CITA (marcar asistencia) ──────────────────────────
// data: { userEmail, eventoId, estado: 'Completada' | 'No asistió' }
function completarCita(data) {
  const estado = data.estado === 'No asistió' ? 'No asistió' : 'Completada';
  actualizarEstadoSheet(data.eventoId, estado, '');
  return { ok: true, msg: 'Cita marcada como ' + estado };
}

// ── PLANTILLAS DE EMAIL ───────────────────────────────────────
function emailTecnicoHtml(data, prof) {
  return `
<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;">
  <div style="background:#003087;padding:16px 20px;border-radius:8px 8px 0 0;">
    <h2 style="color:white;margin:0;font-size:16px;">✉ Solicitud registrada — Promesas Chile</h2>
  </div>
  <div style="padding:20px;border:1px solid #E2E6EF;border-top:none;border-radius:0 0 8px 8px;background:#fff;">
    <p>Estimado/a <b>${data.tecnicoNombre}</b>,</p>
    <p style="margin-top:10px;">La solicitud de hora para <b>${data.atletaNombre}</b> fue <b>registrada exitosamente</b> en el sistema.</p>
    <div style="background:#F5F6FA;border-radius:8px;padding:14px;margin:14px 0;border-left:4px solid #003087;">
      <p><b>⚕ Profesional:</b> ${prof.nombre}</p>
      <p><b>• Fecha y hora:</b> ${data.fechaHoraInicio}</p>
      <p><b>• Atleta:</b> ${data.atletaNombre} · ${data.polo}</p>
      <p><b>• Motivo:</b> ${data.motivo || 'Sin especificar'}</p>
    </div>
    <p style="color:#6B7280;font-size:12px;">La confirmación de asistencia fue enviada directamente al apoderado. Te notificaremos ante cualquier cambio de horario.</p>
    <p style="margin-top:16px;">Saludos,<br><b>${CONFIG.nombrePrograma}</b></p>
  </div>
</div>`;
}

function emailApoderadoHtml(data, prof, fin) {
  const horaFin = fin.toTimeString().slice(0, 5);
  const fechaStr = data.fechaHoraInicio.split('T')[0];
  const horaStr  = (data.fechaHoraInicio.split('T')[1] || '').slice(0, 5);
  return `
<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;">
  <div style="background:#003087;padding:16px 20px;border-radius:8px 8px 0 0;">
    <h2 style="color:white;margin:0;font-size:16px;">✅ Confirmación de Cita — Promesas Chile</h2>
  </div>
  <div style="padding:20px;border:1px solid #E2E6EF;border-top:none;border-radius:0 0 8px 8px;background:#fff;">
    <p>Estimado/a <b>${data.apoderadoNombre}</b>,</p>
    <p style="margin-top:10px;">Confirmamos la siguiente cita para <b>${data.atletaNombre}</b>:</p>
    <div style="background:#F0FDF4;border-radius:8px;padding:14px;margin:14px 0;border-left:4px solid #16A34A;">
      <p><b>• Fecha:</b> ${fechaStr}</p>
      <p><b>• Hora:</b> ${horaStr} a ${horaFin}</p>
      <p><b>⚕ Profesional:</b> ${prof.nombre}</p>
      <p><b>• Lugar:</b> ${CONFIG.lugar}</p>
      <p><b>• Motivo:</b> ${data.motivo || 'Sin especificar'}</p>
    </div>
    <p>Por favor confirme asistencia respondiendo este correo. En caso de no poder asistir, avise con al menos <b>24 horas de anticipación</b>.</p>
    <p style="margin-top:16px;">Saludos cordiales,<br><b>${CONFIG.nombrePrograma}</b></p>
  </div>
</div>`;
}

function emailSuspensionTecnicoHtml(data) {
  return `
<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;">
  <div style="background:#C8102E;padding:16px 20px;border-radius:8px 8px 0 0;">
    <h2 style="color:white;margin:0;font-size:16px;">⚠️ Cita Suspendida — Promesas Chile</h2>
  </div>
  <div style="padding:20px;border:1px solid #FECACA;border-top:none;border-radius:0 0 8px 8px;background:#FEF2F2;">
    <p>Estimado/a <b>${data.tecnicoNombre}</b>,</p>
    <p style="margin-top:10px;">La cita de <b>${data.atletaNombre}</b> programada para el <b>${data.fechaHoraInicio}</b> fue <b>suspendida</b>.</p>
    <p style="margin-top:8px;"><b>Motivo:</b> ${data.motivoSuspension || 'No especificado'}</p>
    <p style="margin-top:10px;">El horario fue liberado en el calendario del profesional. Puedes solicitar un nuevo horario desde el sistema de agendamiento.</p>
    <p style="margin-top:16px;">Saludos,<br><b>${CONFIG.nombrePrograma}</b></p>
  </div>
</div>`;
}

function emailSuspensionApoderadoHtml(data) {
  return `
<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;">
  <div style="background:#C8102E;padding:16px 20px;border-radius:8px 8px 0 0;">
    <h2 style="color:white;margin:0;font-size:16px;">⚠️ Aviso de Suspensión — Promesas Chile</h2>
  </div>
  <div style="padding:20px;border:1px solid #E2E6EF;border-top:none;border-radius:0 0 8px 8px;background:#fff;">
    <p>Estimado/a <b>${data.apoderadoNombre}</b>,</p>
    <p style="margin-top:10px;">Lamentamos informarle que la cita de <b>${data.atletaNombre}</b> programada para el <b>${data.fechaHoraInicio}</b> fue <b>suspendida</b>.</p>
    <p style="margin-top:8px;"><b>Motivo:</b> ${data.motivoSuspension || 'No especificado'}</p>
    <p style="margin-top:10px;">Nos pondremos en contacto a la brevedad para coordinar un nuevo horario. Disculpe los inconvenientes ocasionados.</p>
    <p style="margin-top:16px;">Saludos cordiales,<br><b>${CONFIG.nombrePrograma}</b></p>
  </div>
</div>`;
}

// ── HELPERS ───────────────────────────────────────────────────
function stripHtml(html) {
  return html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

// Devuelve el calendario del profesional (su email). Si no está accesible
// —porque aún no compartió su calendario con la cuenta del sistema—, cae al
// calendario por defecto para no bloquear el agendamiento.
function _calendarioDe(prof) {
  const id = prof.calendarId || prof.email;
  if (id) {
    try {
      const cal = CalendarApp.getCalendarById(id);
      if (cal) return cal;
    } catch (e) { /* sin acceso al calendario */ }
  }
  return CalendarApp.getDefaultCalendar();
}

// Envío centralizado de correos. Usa el remitente configurado (alias) si existe;
// si no, envía desde la cuenta que ejecuta el script.
function enviarCorreo(to, subject, htmlBody) {
  try {
    const opts = { htmlBody: htmlBody, name: CONFIG.nombrePrograma };
    if (CONFIG.remitente) opts.from = CONFIG.remitente; // alias verificado opcional
    MailApp.sendEmail(to, subject, stripHtml(htmlBody), opts);
    return true;
  } catch (e) {
    // No abortar la operación si falla el correo (p.ej. permiso no autorizado aún)
    console.warn('No se pudo enviar correo a ' + to + ': ' + e);
    return false;
  }
}

// Diagnóstico: indica, por profesional, si la cuenta del sistema puede
// escribir en su calendario. Úsalo para verificar que compartieron bien.
function verificarCalendarios(data) {
  // El acceso admin ya fue validado en doPost; aquí solo se ejecuta el diagnóstico.
  let ejecutaComo = '';
  try { ejecutaComo = Session.getEffectiveUser().getEmail(); } catch (e) { ejecutaComo = '(no disponible)'; }
  const resultado = {};
  Object.keys(CONFIG.profesionales).forEach(key => {
    const prof = getProfesional(key) || {};
    const id = prof.calendarId || prof.email || '';
    let estado;
    if (!id) { estado = 'SIN EMAIL en la hoja Usuarios'; }
    else {
      try {
        const cal = CalendarApp.getCalendarById(id);
        estado = cal ? 'OK (accesible)' : 'SIN ACCESO (no compartido con el sistema)';
      } catch (e) { estado = 'ERROR: ' + e.message; }
    }
    resultado[key] = { calendario: id || '(vacío)', estado: estado };
  });
  return { ok: true, ejecutaComo: ejecutaComo, calendarios: resultado };
}

// ── VALIDACIÓN DE USUARIO DESDE SHEETS ───────────────────────
function validarUsuario(data) {
  // Identidad verificada: id_token de Google (login nuevo) o token de sesión (restaurar).
  let email = null;
  if (data.idToken)      email = verificarIdToken(data.idToken);   // login con Google
  else if (data.token)   email = verificarToken(data.token);       // restaurar sesión
  if (!email) return { autorizado: false, error: 'No se pudo verificar la identidad. Inicia sesión con Google.' };

  const usuarios = buscarTodosUsuarios(email);
  if (!usuarios.length) return { autorizado: false, error: 'Usuario no autorizado' };
  // Si tiene múltiples roles, el nombre preferido es el del profesional (más específico)
  const prof = usuarios.find(u => u.rol === 'profesional');
  const admin = usuarios.find(u => u.rol === 'admin');
  const primary = prof || admin || usuarios[0];
  return {
    autorizado: true,
    token: emitirToken(email),      // token de sesión firmado para las siguientes llamadas
    email: email,
    roles: usuarios,                // todos los roles del usuario
    rol: admin ? 'admin' : primary.rol,
    nombre: primary.nombre,
    key: primary.key,
  };
}

function buscarUsuario(email) {
  const todos = buscarTodosUsuarios(email);
  return todos.length ? todos[0] : null;
}

// Datos del profesional por su key (kine/psico/nutri/medico/profis).
// Fuente de verdad: hoja Usuarios (rol=profesional). Cae a CONFIG si no está.
function getProfesional(key) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.sheetId);
    const sheet = ss.getSheetByName('Usuarios');
    if (sheet) {
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const [rol, nombre, correo, k] = rows[i];
        if (rol === 'profesional' && String(k).trim() === key) {
          const email = String(correo || '').trim();
          const base = CONFIG.profesionales[key] || {};
          return { nombre: nombre || base.nombre || key, email: email, calendarId: base.calendarId || email };
        }
      }
    }
  } catch (e) { /* usar fallback */ }
  return CONFIG.profesionales[key] || null;
}

function buscarTodosUsuarios(email) {
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('Usuarios');
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  const matches = [];
  for (let i = 1; i < rows.length; i++) {
    const [rol, nombre, correo, key, estado] = rows[i];
    if (correo && correo.toLowerCase().trim() === email.toLowerCase().trim() && estado !== 'inactivo') {
      matches.push({ rol, nombre, key });
    }
  }
  return matches;
}

// ── MIGRACIÓN: ADMIN A CUENTA DE PROMESAS ────────────────────
// Ejecutar UNA VEZ desde el editor. Reasigna el rol admin a la cuenta del
// sistema y deja a matias.pendola solo como profesional (profis).
function migrarAdminAPromesas() {
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('Usuarios');
  if (!sheet) return { ok: false, error: 'No existe la hoja Usuarios' };
  const rows = sheet.getDataRange().getValues();
  let cambios = 0;
  for (let i = 1; i < rows.length; i++) {
    const rol = String(rows[i][0]).trim().toLowerCase();
    if (rol === 'admin') {
      sheet.getRange(i + 1, 2).setValue('Promesas Chile (Admin)');          // NOMBRE
      sheet.getRange(i + 1, 3).setValue('promesaschilelosrios@gmail.com');  // CORREO
      sheet.getRange(i + 1, 5).setValue('confirmado');                      // ESTADO
      cambios++;
    }
  }
  return { ok: true, msg: 'Filas admin actualizadas: ' + cambios + '. Ahora el admin es promesaschilelosrios@gmail.com y matias.pendola queda solo como profesional.' };
}

// ── OBTENER CITAS DESDE SHEETS ───────────────────────────────
function getCitas(data) {
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('Citas');
  if (!sheet) return { ok: true, citas: [] };
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { ok: true, citas: [] };
  // Headers: ID Evento | Fecha | Hora | Atleta | Polo | Técnico | Email Técnico |
  //          Apoderado | Email Apoderado | Profesional | Disciplina | Tipo Bloque | Motivo | Estado | Fecha Registro
  const citas = rows.slice(1).map(r => ({
    eventoId:    r[0],
    fecha:       _fechaTxt(ss, r[1]),
    hora:        _horaTxt(ss, r[2]),
    atleta:      r[3],
    polo:        r[4],
    tecNom:      r[5],
    tecEmail:    r[6],
    apodNom:     r[7],
    apodEmail:   r[8],
    profNom:     r[9],
    prof:        r[10],
    tipoBloque:  r[11],
    motivo:      r[12],
    estado:      (r[13] || 'confirmada').toLowerCase(),
    fechaReg:    r[14],
  })).filter(c => c.atleta); // ignorar filas vacías
  // Filtro opcional por profesional
  if (data.profKey) return { ok: true, citas: citas.filter(c => c.prof === data.profKey) };
  return { ok: true, citas };
}

// ── RESET CITAS (solo admin) ─────────────────────────────────
// Borra todas las filas de datos de la hoja Citas (mantiene encabezados).
function resetCitas(data) {
  // El acceso admin ya fue validado en doPost.
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('Citas');
  if (!sheet) return { ok: true, msg: 'No hay hoja Citas' };
  const last = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last - 1);
  return { ok: true, msg: 'Citas reseteadas' };
}

// ── LEER HOJA GENÉRICA ───────────────────────────────────────
function leerHoja(data) {
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName(data.hoja);
  if (!sheet) return { ok: false, error: 'Hoja no encontrada: ' + data.hoja };
  const rows = sheet.getDataRange().getValues();
  return { ok: true, rows: rows };
}

// ── SETUP: HOJA USUARIOS ─────────────────────────────────────
function setupUsuarios() {
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  let sheet = ss.getSheetByName('Usuarios');
  if (sheet) return { ok: true, msg: 'La hoja Usuarios ya existe' };

  sheet = ss.insertSheet('Usuarios');

  // Encabezados
  const headers = [
    ['ROL', 'NOMBRE', 'CORREO', 'KEY / POLO', 'ESTADO'],
  ];

  // Datos iniciales: profesionales
  const profesionales = [
    ['profesional', 'Kin. Katalina Camino',        'kine.katalinacamino@gmail.com',  'kine',   'confirmado'],
    ['profesional', 'Ps. Mario Pidal',              'pidalmario@gmail.com',           'psico',  'confirmado'],
    ['profesional', 'Nut. Josefina Henríquez',      'josefina.enri.sch@gmail.com',    'nutri',  'confirmado'],
    ['profesional', 'Dr. Juan Manuel Guzmán',       '',                               'medico', 'PENDIENTE'],
    ['profesional', 'Prof. Matías Péndola',         'matias.pendola@gmail.com',       'profis', 'confirmado'],
  ];

  // Técnicos/entrenadores (correos pendientes de confirmar)
  const tecnicos = [
    ['tecnico', 'Cesar Barria Vargas',                   '', 'Basketball', 'PENDIENTE'],
    ['tecnico', 'Miguel Cerda',                          '', 'Remo',       'PENDIENTE'],
    ['tecnico', 'Angela Parra',                          '', 'Natación',   'PENDIENTE'],
    ['tecnico', 'Gustavo Miranda Rivera',                '', 'Remo',       'PENDIENTE'],
    ['tecnico', 'Hernán Alejandro Chavarría Cuevas',     '', 'Remo',       'PENDIENTE'],
    ['tecnico', 'Israel Eliseo Saez Molina',             '', 'Remo',       'PENDIENTE'],
    ['tecnico', 'Patricio Barrera Alarcon',              '', 'Atletismo',  'PENDIENTE'],
    ['tecnico', 'Nelson Morales Poblete',                '', 'Atletismo',  'PENDIENTE'],
    ['tecnico', 'Cristian Marcelo Astete Miranda',       '', 'Judo',       'PENDIENTE'],
    ['tecnico', 'Ignacio Andrés Silva Becerra',          '', 'Judo',       'PENDIENTE'],
  ];

  // Admin
  const admins = [
    ['admin', 'Matías Péndola (Admin)', 'matias.pendola@gmail.com', 'admin', 'confirmado'],
  ];

  const allRows = headers.concat(admins).concat(profesionales).concat(tecnicos);
  sheet.getRange(1, 1, allRows.length, 5).setValues(allRows);

  // Formato encabezado
  const headerRange = sheet.getRange(1, 1, 1, 5);
  headerRange.setBackground('#003087').setFontColor('white').setFontWeight('bold');

  // Ancho de columnas
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 260);
  sheet.setColumnWidth(3, 280);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 110);

  sheet.setFrozenRows(1);

  return { ok: true, msg: 'Hoja Usuarios creada con ' + (allRows.length - 1) + ' registros' };
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function err(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
