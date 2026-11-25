/* =====================================================================
   MI DÍA · Script de Google del equipo
   ---------------------------------------------------------------------
   Hace cuatro cosas:
   1. Da a la app las tareas del día (del calendario del equipo).
   2. Da a la app los premios activos (de la hoja de Google).
   3. Recibe el registro de lo que ha hecho y lo guarda en la hoja.
   4. Manda cada mañana un correo con el resumen del día anterior.

   IMPORTANTE: este archivo está en un repositorio público.
   No escribas aquí nombres ni datos personales. Rellena la
   configuración solo en la copia que pegues en script.google.com.
   ===================================================================== */

/* ---------- CONFIGURACIÓN: rellena solo esto ---------- */
const CONFIG = {
  // ID del calendario. Está en Google Calendar > Configuración del calendario > "Integrar el calendario" > "ID de calendario"
  CALENDARIO_ID: 'PEGA_AQUI_EL_ID_DEL_CALENDARIO',

  // Correo (o correos separados por comas) que recibe el resumen. Puede ser de Outlook.
  CORREO_RESUMEN: 'PEGA_AQUI_EL_CORREO',

  // Hora a la que llega el resumen del día anterior (hora de Madrid)
  HORA_RESUMEN: 8,

  // false = todos los días. true = de lunes a viernes, y el lunes resume viernes, sábado y domingo.
  SOLO_LABORABLES: true
};
/* ------------------------------------------------------ */

const ZONA = 'Europe/Madrid';
const P = PropertiesService.getScriptProperties();
const COLS_REGISTRO = ['Clave','Fecha','Tarea','Categoría','Estado','Inicio','Fin','Min. trabajando','Min. en pausa','Pausas','Min. consulta','Veces pospuesta','Salió de la app','Hora que indica (si la marcó a mano)','Comentario','Foto','Estrellas','Actualizado'];
const COLS_PREMIOS = ['Premio','Estrellas','Activo','Desde','Hasta','Tipo'];
const COLS_CANJES = ['Clave','Fecha','Premio','Estrellas'];

/* =====================================================================
   1. INSTALAR: ejecútalo una vez a mano (botón "Ejecutar")
   ===================================================================== */
function instalar(){
  if (Session.getScriptTimeZone() !== ZONA){
    throw new Error('La zona horaria del proyecto no es Europe/Madrid. Ve a "Configuración del proyecto" (rueda dentada) y cámbiala, o revisa el archivo appsscript.json.');
  }
  if (CONFIG.CALENDARIO_ID.indexOf('PEGA_AQUI') === 0) throw new Error('Falta poner el ID del calendario en CONFIG.');
  if (CONFIG.CORREO_RESUMEN.indexOf('PEGA_AQUI') === 0) throw new Error('Falta poner el correo del resumen en CONFIG.');
  if (!CalendarApp.getCalendarById(CONFIG.CALENDARIO_ID)) throw new Error('No encuentro el calendario. Revisa el ID y que esta cuenta tenga acceso.');

  // Hoja de cálculo
  let ss = P.getProperty('SHEET_ID') ? abrirHoja_() : null;
  if (!ss){
    ss = SpreadsheetApp.create('Mi día · registro y premios');
    const premios = ss.getSheets()[0].setName('Premios');
    premios.getRange(1,1,1,COLS_PREMIOS.length).setValues([COLS_PREMIOS]).setFontWeight('bold');
    premios.getRange(2,1,5,COLS_PREMIOS.length).setValues([
      ['Elegir la cena un día', 15, 'Sí', '', '', 'Fijo'],
      ['30 minutos extra de móvil', 20, 'Sí', '', '', 'Fijo'],
      ['Elegir la peli del finde', 25, 'Sí', '', '', 'Fijo'],
      ['Salida especial con el equipo', 60, 'Sí', '', '', 'Fijo'],
      ['Ir a la playa', 30, 'Sí', '2027-06-01', '2027-09-15', 'Temporal']
    ]);
    const si = SpreadsheetApp.newDataValidation().requireValueInList(['Sí','No'], true).build();
    const tipo = SpreadsheetApp.newDataValidation().requireValueInList(['Fijo','Temporal'], true).build();
    premios.getRange('C2:C200').setDataValidation(si);
    premios.getRange('F2:F200').setDataValidation(tipo);
    premios.getRange('D2:E200').setNumberFormat('yyyy-mm-dd');
    premios.setFrozenRows(1); premios.autoResizeColumns(1, COLS_PREMIOS.length);

    const reg = ss.insertSheet('Registro');
    reg.getRange(1,1,1,COLS_REGISTRO.length).setValues([COLS_REGISTRO]).setFontWeight('bold');
    // Fechas y horas como texto, para que la hoja no las convierta a su manera
    ['B2:B','F2:G','N2:N'].forEach(r => reg.getRange(r).setNumberFormat('@'));
    reg.setFrozenRows(1); reg.hideColumns(1);

    const can = ss.insertSheet('Canjes');
    can.getRange(1,1,1,COLS_CANJES.length).setValues([COLS_CANJES]).setFontWeight('bold');
    can.getRange('B2:B').setNumberFormat('@');
    can.setFrozenRows(1); can.hideColumns(1);

    P.setProperty('SHEET_ID', ss.getId());
  }

  // Carpeta de fotos
  if (!P.getProperty('FOLDER_ID')) P.setProperty('FOLDER_ID', DriveApp.createFolder('Mi día · fotos de tareas').getId());

  // Token secreto
  if (!P.getProperty('TOKEN')) P.setProperty('TOKEN', Utilities.getUuid().replace(/-/g,''));

  // Disparador del resumen diario (se borra el anterior para no duplicarlo)
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'resumenDiario').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('resumenDiario').timeBased().everyDays(1).atHour(CONFIG.HORA_RESUMEN).nearMinute(0).inTimezone(ZONA).create();

  Logger.log('Instalado.');
  Logger.log('Hoja: ' + ss.getUrl());
  Logger.log('Carpeta de fotos: ' + DriveApp.getFolderById(P.getProperty('FOLDER_ID')).getUrl());
  Logger.log('Ahora publica el script (Implementar > Nueva implementación) y después ejecuta "verEnlace".');
}

/* Te dice el enlace completo que hay que pegar en la app */
function verEnlace(){
  const url = ScriptApp.getService().getUrl();
  if (!url) { Logger.log('Todavía no has publicado el script. Ve a Implementar > Nueva implementación.'); return; }
  Logger.log('Pega este enlace en la app (Ajustes > Zona del equipo > Enlace del calendario):');
  Logger.log(url + '?token=' + P.getProperty('TOKEN'));
}

/* Te dice cuántos correos puede mandar todavía hoy esta cuenta */
function verCuotaCorreo(){
  Logger.log('Correos que se pueden mandar todavía hoy: ' + MailApp.getRemainingDailyQuota());
}

/* Prueba el resumen sin esperar a mañana */
function probarResumen(){ resumenDiario(true); }

/* =====================================================================
   2. LA APP PIDE TAREAS Y PREMIOS (GET)
   ===================================================================== */
function doGet(e){
  const q = (e && e.parameter) || {};
  if (!tokenOk_(q.token)) return json_({error:'token'});
  const dia = q.d && /^\d{4}-\d{2}-\d{2}$/.test(q.d) ? fechaDe_(q.d) : new Date();
  return json_({events: eventosDelDia_(dia), rewards: premiosActivos_(dia)});
}

function eventosDelDia_(dia){
  const cal = CalendarApp.getCalendarById(CONFIG.CALENDARIO_ID);
  if (!cal) return [];
  return cal.getEventsForDay(dia).map(ev => ({
    id: idEvento_(ev),
    title: ev.getTitle(),
    start: ev.getStartTime().toISOString(),
    end: ev.getEndTime().toISOString(),
    allDay: ev.isAllDayEvent(),
    description: ev.getDescription() || ''
  }));
}
/* Los eventos que se repiten comparten ID, así que se añade la hora de inicio */
function idEvento_(ev){ return ev.getId() + '@' + ev.getStartTime().getTime(); }

function premiosActivos_(dia){
  const sh = abrirHoja_().getSheetByName('Premios');
  const filas = sh.getLastRow() > 1 ? sh.getRange(2,1,sh.getLastRow()-1,COLS_PREMIOS.length).getValues() : [];
  const hoy = texto_(dia, 'yyyy-MM-dd');
  return filas.filter(f => {
    const nombre = String(f[0]).trim(), est = Number(f[1]);
    if (!nombre || !(est > 0)) return false;
    if (!/^s[ií]$/i.test(String(f[2]).trim())) return false;
    const desde = f[3] ? texto_(f[3] instanceof Date ? f[3] : new Date(f[3]), 'yyyy-MM-dd') : '';
    const hasta = f[4] ? texto_(f[4] instanceof Date ? f[4] : new Date(f[4]), 'yyyy-MM-dd') : '';
    if (desde && hoy < desde) return false;
    if (hasta && hoy > hasta) return false;
    return true;
  }).map(f => ({id: slug_(f[0]), name: String(f[0]).trim(), cost: Math.round(Number(f[1])), type: String(f[5] || '').trim()}));
}

/* =====================================================================
   3. LA APP ENVÍA EL REGISTRO (POST)
   ===================================================================== */
function doPost(e){
  let body;
  try { body = JSON.parse(e.postData.contents); } catch(err){ return json_({error:'formato'}); }
  if (!tokenOk_(body.token)) return json_({error:'token'});

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    const ss = abrirHoja_();
    const ok = [];
    (body.items || []).forEach(it => { guardarTarea_(ss.getSheetByName('Registro'), it); ok.push(it.key); });
    (body.canjes || []).forEach(c => { guardarCanje_(ss.getSheetByName('Canjes'), c); ok.push(c.key); });
    return json_({ok: ok});
  } finally { lock.releaseLock(); }
}

function guardarTarea_(sh, it){
  const fila = buscarFila_(sh, it.key);
  let foto = fila ? sh.getRange(fila, 16).getValue() : '';
  if (it.photo && /^data:image\/jpeg;base64,/.test(it.photo)){
    const datos = Utilities.base64Decode(it.photo.split(',')[1]);
    const nombre = it.fecha + '_' + slug_(it.tarea).slice(0,40) + '.jpg';
    const archivo = DriveApp.getFolderById(P.getProperty('FOLDER_ID')).createFile(Utilities.newBlob(datos, 'image/jpeg', nombre));
    foto = archivo.getUrl();
  }
  const valores = [[it.key, it.fecha, it.tarea, it.categoria, it.estado, it.inicio || '', it.fin || '',
    num_(it.minTrabajando), num_(it.minPausa), it.pausas || '', num_(it.minConsulta), num_(it.pospuesta),
    it.salio === true ? 'Sí' : it.salio === false ? 'No' : '', it.horaManual || '', it.comentario || '', foto,
    num_(it.estrellas), texto_(new Date(), 'yyyy-MM-dd HH:mm')]];
  if (fila) sh.getRange(fila, 1, 1, COLS_REGISTRO.length).setValues(valores);
  else sh.appendRow(valores[0]);
}

function guardarCanje_(sh, c){
  if (buscarFila_(sh, c.key)) return;
  sh.appendRow([c.key, c.fecha, c.premio, num_(c.estrellas)]);
}

/* =====================================================================
   4. RESUMEN DIARIO POR CORREO
   ===================================================================== */
function resumenDiario(esPrueba){
  const hoy = new Date();
  const dow = Number(texto_(hoy, 'u')); // 1 = lunes ... 7 = domingo
  if (CONFIG.SOLO_LABORABLES && dow >= 6 && esPrueba !== true) return;

  const atras = CONFIG.SOLO_LABORABLES && dow === 1 ? [3,2,1] : [1];
  const dias = atras.map(n => { const d = new Date(hoy); d.setDate(d.getDate() - n); return d; });

  const ss = abrirHoja_();
  const reg = leer_(ss.getSheetByName('Registro'), COLS_REGISTRO.length);
  const can = leer_(ss.getSheetByName('Canjes'), COLS_CANJES.length);

  const bloques = dias.map(dia => bloqueDia_(dia, reg, can)).join('');
  const titulo = dias.length > 1 ? 'Resumen del fin de semana' : 'Resumen del ' + fechaLarga_(dias[0]);
  const html = `<div style="font-family:Arial,sans-serif;font-size:14px;color:#26213A;max-width:640px">
    <h2 style="margin:0 0 4px">Mi día · ${esc_(titulo)}</h2>
    <p style="color:#6E6886;margin:0 0 18px">Correo automático. El detalle completo está en la <a href="${ss.getUrl()}">hoja del equipo</a>.</p>
    ${bloques}</div>`;

  MailApp.sendEmail({to: CONFIG.CORREO_RESUMEN, subject: 'Mi día · ' + titulo, htmlBody: html, name: 'Mi día'});
}

function bloqueDia_(dia, reg, can){
  const f = texto_(dia, 'yyyy-MM-dd');
  const eventos = eventosDelDia_(dia);
  const filas = reg.filter(r => texto_(r[1] instanceof Date ? r[1] : new Date(r[1]), 'yyyy-MM-dd') === f || String(r[1]) === f);
  const porClave = {}; filas.forEach(r => porClave[String(r[0]).split('|')[1]] = r);

  const hechas = [], pendientes = [];
  eventos.forEach(ev => {
    const r = porClave[ev.id];
    if (r && /^Hecha/.test(r[4])) hechas.push(r); else pendientes.push({ev, r});
  });
  // Tareas registradas que ya no están en el calendario
  filas.forEach(r => { if (/^Hecha/.test(r[4]) && !hechas.includes(r)) hechas.push(r); });

  const estrellas = filas.reduce((s, r) => s + (Number(r[16]) || 0), 0);
  const canjes = can.filter(c => String(c[1]) === f || (c[1] instanceof Date && texto_(c[1], 'yyyy-MM-dd') === f));

  const li = s => `<li style="margin:0 0 8px">${s}</li>`;
  const detalleHecha = r => {
    const p = [];
    if (r[5]) p.push(`de ${r[5]} a ${r[6]}`);
    if (r[7] !== '') p.push(`${r[7]} min trabajando`);
    if (r[13]) p.push(`marcada a mano, dice que a las ${esc_(r[13])}`);
    if (Number(r[8]) > 0) p.push(`pausas: ${esc_(r[9])}`);
    if (Number(r[10]) > 0) p.push(`${r[10]} min de consulta`);
    if (Number(r[11]) > 0) p.push(`pospuesta ${r[11]} ${Number(r[11]) === 1 ? 'vez' : 'veces'}`);
    if (r[12] === 'Sí') p.push('salió de la app');
    let s = `<b>${esc_(r[2])}</b><br><span style="color:#6E6886">${p.join(' · ')}</span>`;
    if (r[14]) s += `<br>Comentario: “${esc_(r[14])}”`;
    if (r[15]) s += `<br><a href="${r[15]}">Ver foto</a>`;
    return li(s);
  };
  const detallePendiente = x => {
    const t = x.ev.title.replace(/^(casa|insti|instituto|clase)\s*[:\-]\s*/i, '');
    const hora = x.ev.allDay ? 'todo el día' : texto_(new Date(x.ev.start), 'HH:mm');
    let estado = 'sin empezar';
    if (x.r){ estado = String(x.r[4]).toLowerCase(); if (Number(x.r[11]) > 0) estado += `, pospuesta ${x.r[11]} ${Number(x.r[11]) === 1 ? 'vez' : 'veces'}`; }
    return li(`<b>${esc_(t)}</b> (${hora})<br><span style="color:#6E6886">${esc_(estado)}</span>`);
  };

  return `<h3 style="margin:18px 0 6px;border-bottom:2px solid #E4E1EE;padding-bottom:4px">${esc_(fechaLarga_(dia))}</h3>
    <p style="margin:0 0 10px"><b>${hechas.length}</b> hechas de ${Math.max(eventos.length, hechas.length)} · <b>${estrellas}</b> estrellas ganadas</p>
    ${hechas.length ? `<p style="margin:10px 0 4px"><b>Hechas</b></p><ul style="padding-left:18px;margin:0">${hechas.map(detalleHecha).join('')}</ul>` : ''}
    ${pendientes.length ? `<p style="margin:10px 0 4px"><b>Sin hacer</b></p><ul style="padding-left:18px;margin:0">${pendientes.map(detallePendiente).join('')}</ul>` : ''}
    ${canjes.length ? `<p style="margin:10px 0 4px"><b>Premios canjeados</b></p><ul style="padding-left:18px;margin:0">${canjes.map(c => li(esc_(c[2]) + ' (' + c[3] + ' estrellas)')).join('')}</ul>` : ''}
    ${!eventos.length && !filas.length ? '<p style="color:#6E6886">No había tareas en el calendario.</p>' : ''}`;
}

/* =====================================================================
   Utilidades
   ===================================================================== */
function abrirHoja_(){
  const id = P.getProperty('SHEET_ID');
  if (!id) throw new Error('Primero ejecuta "instalar".');
  return SpreadsheetApp.openById(id);
}
function tokenOk_(t){ const ok = P.getProperty('TOKEN'); return !!ok && t === ok; }
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function texto_(d, f){ return Utilities.formatDate(d, ZONA, f); }
function fechaDe_(s){ const p = s.split('-').map(Number); return new Date(p[0], p[1]-1, p[2], 12); }
function fechaLarga_(d){
  const dias = ['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];
  const meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  const x = fechaDe_(texto_(d, 'yyyy-MM-dd'));
  return dias[x.getDay()] + ' ' + x.getDate() + ' de ' + meses[x.getMonth()];
}
function num_(n){ return (n === undefined || n === null || n === '') ? '' : Number(n); }
function slug_(s){ return String(s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,''); }
function esc_(s){ return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }
function leer_(sh, n){ return sh.getLastRow() > 1 ? sh.getRange(2,1,sh.getLastRow()-1,n).getValues() : []; }
function buscarFila_(sh, key){
  if (sh.getLastRow() < 2) return 0;
  const claves = sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  for (let i = 0; i < claves.length; i++) if (claves[i][0] === key) return i + 2;
  return 0;
}
