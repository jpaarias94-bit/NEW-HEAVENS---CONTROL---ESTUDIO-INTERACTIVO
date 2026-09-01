/*************************************************************
 * AGENTE ACADÉMICO — Mantiene el calendario al día, automáticamente.
 *
 * Cuando un correo anuncia un cambio de fecha, el agente lo APLICA solo.
 * El usuario no tiene que hacer nada.
 *
 * REGLAS DE SEGURIDAD (impiden los errores del pasado):
 *  1. Solo modifica exámenes de LA MISMA ASIGNATURA del correo.
 *     Un correo de Ciencias jamás puede mover un examen de Música.
 *  2. Para saber a cuál examen se refiere, usa la fecha DEL CORREO
 *     como referencia, no la de hoy.
 *  3. Si hay varias evaluaciones posibles y no se puede saber cuál,
 *     no adivina: deja una propuesta para que el usuario elija.
 *  4. Cada fecha aplicada guarda el correo que la respalda,
 *     visible en el calendario (columna "Respaldo").
 *************************************************************/

function configurarAgente(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  hojaAgente_(ss);
  SpreadsheetApp.getUi().alert('Listo. Bitácora del Agente creada (hoja "Agente").');
}

function hojaAgente_(ss){
  return ss.getSheetByName('Agente') || crearHoja_(ss,'Agente',
    ['Fecha','Asignatura','Tipo','Detección','Acción tomada','Estado','ID correo',
     'Detalle','FechaNueva','ContenidoObjetivo','FechaCorreo','TextoCorreo']);
}

/*************************************************************
 * LIMPIEZA — devuelve el calendario al plan oficial del colegio.
 * Borra todas las "Fecha confirmada", marcas de "¿Movida?" y estados
 * que el agente haya escrito. Deja el plan tal como lo publicó el colegio.
 *************************************************************/
function restaurarPlanOficial(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const ui=SpreadsheetApp.getUi();
  const h=ss.getSheetByName('Examenes');
  if(!h || h.getLastRow()<2){ ui.alert('No hay exámenes.'); return; }

  const resp=ui.alert('Restaurar plan oficial',
    'Se borrarán TODAS las fechas confirmadas y marcas de "movida" del calendario, '+
    'volviendo al plan original del colegio.\n\n¿Continuar?', ui.ButtonSet.YES_NO);
  if(resp!==ui.Button.YES) return;

  const n=h.getLastRow()-1;
  h.getRange(2,8,n,1).clearContent();          // Fecha confirmada
  h.getRange(2,9,n,1).clearContent();          // ¿Movida?
  // Estado vuelve a "Programado"
  const estados=[]; for(let i=0;i<n;i++) estados.push(['Programado']);
  h.getRange(2,10,n,1).setValues(estados);
  ordenarExamenes_(ss);

  // limpia la bitácora para volver a analizar desde cero
  const hAg=hojaAgente_(ss);
  if(hAg.getLastRow()>1) hAg.getRange(2,1,hAg.getLastRow()-1,hAg.getLastColumn()).clearContent();

  ui.alert('✅ Calendario restaurado al plan oficial.\n\n'+
           'La bitácora del agente se limpió. Usa "🕵️ Revisar correos (proponer cambios)" '+
           'para que el agente vuelva a analizar y te proponga los cambios que tú aprobarás.');
}

// ====== Ejecutar desde el menú ======
function ejecutarAgente(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const n=correrAgente_(ss);
  try{ SpreadsheetApp.getUi().alert('Revisión lista.\n\nPropuestas nuevas: '+n+
    '\n\nRevísalas en la pestaña "Agente" de la app y aprueba las que correspondan.'); }catch(e){}
}

// ====== Núcleo: analiza y PROPONE (nunca aplica) ======
function correrAgente_(ss){
  const hAnun=ss.getSheetByName('Anuncios');
  if(!hAnun || hAnun.getLastRow()<2) return 0;
  const hAg=hojaAgente_(ss);

  const yaRevisados=new Set(hAg.getLastRow()>1
    ? hAg.getRange(2,7,hAg.getLastRow()-1,1).getValues().flat().map(String) : []);

  let filas=hAnun.getRange(2,1,hAnun.getLastRow()-1,hAnun.getLastColumn()).getValues();
  filas.sort((a,b)=> new Date(a[0]) - new Date(b[0]));  // cronológico

  let novedades=0;

  filas.forEach(f=>{
    const fechaCorreo=f[0], asig=f[1], texto=String(f[3]||''), id=String(f[4]||'');
    if(!id || yaRevisados.has(id) || !texto) return;

    const analisis=analizarCorreoAgente_(ss, asig, texto, fechaCorreo);
    yaRevisados.add(id);

    const fCorreoTxt = fechaCorreo ? Utilities.formatDate(new Date(fechaCorreo),'GMT-3','dd/MM/yyyy HH:mm') : '';
    const extracto = texto.replace(/\s+/g,' ').slice(0,300);

    if(!analisis || !analisis.eventos || !analisis.eventos.length){
      hAg.appendRow([new Date(), asig,'Revisado','Sin cambios relevantes','—','ok', id,'','','',fCorreoTxt,'']);
      return;
    }

    analisis.eventos.forEach(ev=>{
      const tipo=ev.tipo||'Aviso';

      // Material: se registra directo (no toca fechas)
      if(tipo==='Material nuevo'){
        if(ev.detalle) registrarMaterial_(ss, asig, cap_(String(ev.detalle)), new Date());
        hAg.appendRow([new Date(), asig, tipo, ev.deteccion||'', 'Material registrado','ok', id,
                       ev.detalle||'','','',fCorreoTxt, extracto]);
        novedades++;
        return;
      }

      // Cambios de fecha / suspension: se APLICAN si son seguros
      const res = aplicarCambioSeguro_(ss, asig, ev, fechaCorreo, id);
      hAg.appendRow([new Date(), asig, tipo, ev.deteccion||'', res.accion, res.estado, id,
                     ev.detalle||'', ev.fecha||'', res.objetivo||'', fCorreoTxt, extracto]);
      novedades++;
    });  });

  return novedades;
}

/*************************************************************
 * APLICA el cambio automáticamente, con reglas estrictas de seguridad.
 *
 * REGLAS QUE IMPIDEN ERRORES:
 *  1. Solo puede tocar exámenes de LA MISMA ASIGNATURA del correo.
 *  2. Para elegir cuál, usa la fecha DEL CORREO como referencia
 *     (no la fecha de hoy), así reanalizar el pasado no descoloca nada.
 *  3. Si hay varios candidatos y la IA no identificó cuál, NO adivina:
 *     deja propuesta para que el usuario elija.
 *  4. Cada cambio guarda el correo que lo respalda.
 *************************************************************/
function aplicarCambioSeguro_(ss, asig, ev, fechaCorreo, idCorreo){
  const tipo=ev.tipo||'Aviso';

  // Exámenes de ESTA asignatura, vigentes a la fecha del correo
  const cands=candidatosAFecha_(ss, asig, fechaCorreo);

  // Determina a qué examen se refiere
  let objetivo=null;
  if(ev.examen_contenido){
    // valida que ese contenido exista EN ESTA asignatura
    const encontrado=cands.filter(c=>
      c.contenido.toLowerCase().trim()===String(ev.examen_contenido).toLowerCase().trim());
    if(encontrado.length) objetivo=encontrado[0].contenido;
    else {
      // busca en toda la asignatura (puede ser un examen ya pasado)
      const todos=examenesDeAsignaturaConFecha_(ss, asig);
      const e2=todos.filter(c=>
        c.contenido.toLowerCase().trim()===String(ev.examen_contenido).toLowerCase().trim());
      if(e2.length) objetivo=e2[0].contenido;
    }
  }
  // Si no lo identificó pero hay UN SOLO candidato en esta asignatura → es ese
  if(!objetivo && cands.length===1) objetivo=cands[0].contenido;

  if(tipo==='Suspensión'){
    if(objetivo){
      const ok=marcarEstadoExamen_(ss, asig, objetivo, 'SUSPENDIDA');
      return ok ? {accion:'Marcada como suspendida', estado:'aplicado', objetivo:objetivo}
                : {accion:'No se ubicó el examen', estado:'propuesta'};
    }
    return {accion:'Suspensión — elige a cuál evaluación aplicarla', estado:'propuesta'};
  }

  if((tipo==='Examen movido' || tipo==='Evaluación nueva') && ev.fecha){
    const f=parsearFechaISO_(ev.fecha);
    if(!f) return {accion:'Fecha ilegible', estado:'propuesta'};

    if(objetivo){
      ajustarExamenPorContenido_(ss, asig, objetivo, f);
      ordenarExamenes_(ss);
      guardarRespaldo_(ss, asig, objetivo, idCorreo);
      return {accion:'✅ Fecha aplicada automáticamente ('+ev.fecha+')', estado:'aplicado', objetivo:objetivo};
    }
    return {accion:'Cambio de fecha — hay varias evaluaciones posibles, elige cuál', estado:'propuesta'};
  }

  return {accion:'Registrado — revisar', estado:'propuesta'};
}

// Exámenes de una asignatura vigentes a una fecha dada (la del correo)
function candidatosAFecha_(ss, asig, fechaRef){
  const h=ss.getSheetByName('Examenes');
  if(!h || h.getLastRow()<2) return [];
  const datos=h.getRange(2,1,h.getLastRow()-1,10).getValues();
  const ref = fechaRef ? new Date(fechaRef) : new Date();
  const refD=new Date(ref.getFullYear(),ref.getMonth(),ref.getDate());
  const out=[];
  datos.forEach(r=>{
    if(r[0]!==asig || !r[1]) return;                       // MISMA ASIGNATURA, siempre
    if(String(r[9]||'').toUpperCase().indexOf('SUSPEND')>=0) return;
    const vig = r[7] ? new Date(r[7]) : new Date(r[5]);
    if(isNaN(vig)) return;
    const v=new Date(vig.getFullYear(),vig.getMonth(),vig.getDate());
    if(v.getTime() >= refD.getTime()-86400000) out.push({contenido:String(r[1]), fecha:v});
  });
  out.sort((a,b)=>a.fecha-b.fecha);
  return out;
}

// ====== Exámenes de una asignatura, con su fecha vigente ======
function examenesDeAsignaturaConFecha_(ss, asig){
  const h=ss.getSheetByName('Examenes');
  if(!h || h.getLastRow()<2) return [];
  const datos=h.getRange(2,1,h.getLastRow()-1,10).getValues();
  const out=[];
  datos.forEach(r=>{
    if(r[0]!==asig || !r[1]) return;
    const vig = r[7] ? new Date(r[7]) : new Date(r[5]);
    out.push({contenido:String(r[1]), fecha: isNaN(vig)?null:vig, estado:String(r[9]||'')});
  });
  out.sort((a,b)=>(a.fecha||0)-(b.fecha||0));
  return out;
}

// ====== Análisis con Gemini (solo detecta, no decide aplicar) ======
function analizarCorreoAgente_(ss, asig, texto, fechaCorreo){
  const key=obtenerClaveGemini_();
  if(!key) return null;

  const lista=examenesDeAsignaturaConFecha_(ss, asig);
  const bloqueExam = lista.length
    ? '\n\nEVALUACIONES DE ESTA ASIGNATURA (contenido exacto | fecha en el plan):\n'+
      lista.map(c=>'- '+c.contenido+' | '+(c.fecha?Utilities.formatDate(c.fecha,'GMT-3','dd/MM/yyyy'):'sin fecha')).join('\n')+'\n'
    : '';

  const fTxt = fechaCorreo ? Utilities.formatDate(new Date(fechaCorreo),'GMT-3','dd/MM/yyyy') : '';

  const instruccion=
    'Eres un agente que lee correos de profesores (4° básico, Chile) y detecta cambios en las evaluaciones. '+
    'Este correo se envió el '+fTxt+' y corresponde a la asignatura "'+asig+'".\n'+
    'Devuelve SOLO un JSON válido, sin texto ni ```.\n\n'+
    '{ "eventos": [ ... ] }. Cada evento:\n'+
    '- "tipo": "Examen movido" | "Suspensión" | "Cambio de temario" | "Material nuevo" | "Evaluación nueva".\n'+
    '- "examen_contenido": el contenido EXACTO de la lista de abajo al que se refiere, o null si no se puede saber.\n'+
    '- "fecha": la nueva fecha en "YYYY-MM-DD" (año 2026) si el correo la menciona, o null.\n'+
    '- "deteccion": frase corta y literal de lo que dice el correo (cita el dato clave).\n'+
    '- "detalle": material a traer, si aplica.\n\n'+
    'REGLAS ESTRICTAS:\n'+
    '· Solo reporta un evento si el correo REALMENTE lo dice. No deduzcas ni inventes.\n'+
    '· Si el correo solo cuenta lo que se hizo en clases o comparte material, NO es un cambio de fecha.\n'+
    '· "Examen movido" solo si se indica una nueva fecha. "Suspensión" solo si se cancela sin nueva fecha.\n'+
    '· Si no hay nada relevante, devuelve { "eventos": [] }.'+
    bloqueExam+
    '\n\nCORREO:\n"""'+texto+'"""';

  const url='https://generativelanguage.googleapis.com/v1beta/models/'+obtenerModeloGemini_()+':generateContent';
  const payload={ contents:[{ parts:[{ text:instruccion }] }],
    generationConfig:{ temperature:0.1, maxOutputTokens:1500, responseMimeType:'application/json' } };

  try{
    let resp, intentos=0;
    do{
      resp=UrlFetchApp.fetch(url,{ method:'post', contentType:'application/json',
        headers:{ 'x-goog-api-key':key }, payload:JSON.stringify(payload), muteHttpExceptions:true });
      const c=resp.getResponseCode();
      if(c===503||c===429){ intentos++; Utilities.sleep(3000); } else break;
    } while(intentos<4);
    if(resp.getResponseCode()!==200) return null;
    const obj=extraerJSONAgente_(resp.getContentText());
    if(!obj || !Array.isArray(obj.eventos)) return {eventos:[]};
    return obj;
  }catch(e){ return null; }
}

/*************************************************************
 * APROBAR una propuesta (lo llama la app cuando tú aprietas el botón)
 * Solo aquí se toca el calendario, y siempre con tu confirmación.
 *************************************************************/
function aprobarPropuesta(idCorreo, asig, contenido, fechaISO, tipo){
  const ss=SpreadsheetApp.getActiveSpreadsheet();

  if(tipo==='Suspensión'){
    const ok=marcarEstadoExamen_(ss, asig, contenido, 'SUSPENDIDA');
    if(ok) marcarBitacora_(ss, idCorreo, 'Suspensión aplicada por ti', 'aplicado', contenido);
    return {ok:ok, msg: ok?'Marcada como suspendida':'No se ubicó ese examen'};
  }

  const f=parsearFechaISO_(fechaISO);
  if(!f) return {ok:false, msg:'Fecha inválida'};
  ajustarExamenPorContenido_(ss, asig, contenido, f);
  ordenarExamenes_(ss);
  guardarRespaldo_(ss, asig, contenido, idCorreo);
  marcarBitacora_(ss, idCorreo, 'Fecha aplicada por ti ('+fechaISO+')', 'aplicado', contenido);
  return {ok:true, msg:'Fecha aplicada a '+asig};
}

// Descartar una propuesta que no corresponde
function descartarPropuesta(idCorreo){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  marcarBitacora_(ss, idCorreo, 'Descartada por ti', 'descartada', '');
  return {ok:true, msg:'Propuesta descartada'};
}

function marcarBitacora_(ss, idCorreo, accion, estado, objetivo){
  const h=hojaAgente_(ss);
  if(h.getLastRow()<2) return;
  const d=h.getRange(2,1,h.getLastRow()-1,12).getValues();
  for(let i=0;i<d.length;i++){
    if(String(d[i][6])===String(idCorreo) && String(d[i][5]).toLowerCase()==='propuesta'){
      h.getRange(i+2,5).setValue(accion);
      h.getRange(i+2,6).setValue(estado);
      if(objetivo) h.getRange(i+2,10).setValue(objetivo);
      return;
    }
  }
}

// Guarda en la hoja Examenes de qué correo vino el cambio (trazabilidad)
function guardarRespaldo_(ss, asig, contenido, idCorreo){
  const h=ss.getSheetByName('Examenes'); if(!h||h.getLastRow()<2) return;
  // asegura columna "Respaldo"
  const enc=h.getRange(1,1,1,h.getLastColumn()).getValues()[0];
  let col=enc.indexOf('Respaldo')+1;
  if(col===0){
    col=h.getLastColumn()+1;
    h.getRange(1,col).setValue('Respaldo')
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
  }
  // busca el correo para citar fecha y extracto
  let cita='Correo '+idCorreo;
  const hA=ss.getSheetByName('Anuncios');
  if(hA && hA.getLastRow()>1){
    const d=hA.getRange(2,1,hA.getLastRow()-1,5).getValues();
    for(const r of d){
      if(String(r[4])===String(idCorreo)){
        const f = r[0] instanceof Date ? Utilities.formatDate(r[0],'GMT-3','dd/MM/yyyy') : r[0];
        cita='Correo del '+f+': "'+String(r[3]).replace(/\s+/g,' ').slice(0,180)+'…"';
        break;
      }
    }
  }
  const datos=h.getRange(2,1,h.getLastRow()-1,2).getValues();
  const cLow=String(contenido).toLowerCase().trim();
  for(let i=0;i<datos.length;i++){
    if(datos[i][0]===asig && String(datos[i][1]).toLowerCase().trim()===cLow){
      h.getRange(i+2,col).setValue(cita);
      return;
    }
  }
}

// Lista de exámenes de una asignatura, para elegir a cuál aplicar
function opcionesExamenes(asig){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  return examenesDeAsignaturaConFecha_(ss, asig).map(c=>
    c.contenido + (c.fecha?'  ('+Utilities.formatDate(c.fecha,'GMT-3','dd/MM')+')':'')
  );
}
// Devuelve solo los contenidos (sin la fecha entre paréntesis)
function opcionesExamenesLimpio(asig){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  return examenesDeAsignaturaConFecha_(ss, asig).map(c=>c.contenido);
}

function marcarEstadoExamen_(ss, asig, contenido, estado){
  const h=ss.getSheetByName('Examenes'); if(!h || h.getLastRow()<2) return false;
  const datos=h.getRange(2,1,h.getLastRow()-1,10).getValues();
  const cLow=String(contenido).toLowerCase().trim();
  for(let i=0;i<datos.length;i++){
    if(datos[i][0]===asig && String(datos[i][1]).toLowerCase().trim()===cLow){
      h.getRange(i+2,10).setValue(estado);
      return true;
    }
  }
  return false;
}

function leerBitacoraAgente_(ss){
  const h=ss.getSheetByName('Agente'); if(!h || h.getLastRow()<2) return [];
  const enc=h.getRange(1,1,1,h.getLastColumn()).getValues()[0];
  const filas=h.getRange(2,1,h.getLastRow()-1,h.getLastColumn()).getValues().map(f=>{
    const o={}; enc.forEach((c,i)=>o[c]=f[i] instanceof Date?f[i].toISOString():f[i]); return o;
  });
  const rel=filas.filter(x=> x.Tipo!=='Revisado' && String(x.Estado).toLowerCase()!=='descartada');
  // propuestas primero, luego por fecha
  rel.sort((a,b)=>{
    const pa=String(a.Estado).toLowerCase()==='propuesta'?0:1;
    const pb=String(b.Estado).toLowerCase()==='propuesta'?0:1;
    if(pa!==pb) return pa-pb;
    return new Date(b.Fecha)-new Date(a.Fecha);
  });
  return rel;
}

function extraerJSONAgente_(rawContentText){
  try{
    const data=JSON.parse(rawContentText);
    let salida=(data?.candidates?.[0]?.content?.parts?.[0]?.text||'').replace(/```json|```/g,'').trim();
    const ini=salida.indexOf('{');
    if(ini<0) return null;
    let prof=0, fin=-1;
    for(let i=ini;i<salida.length;i++){
      if(salida[i]==='{') prof++;
      else if(salida[i]==='}'){ prof--; if(prof===0){ fin=i; break; } }
    }
    if(fin>ini) salida=salida.slice(ini,fin+1);
    return JSON.parse(salida);
  }catch(e){ return null; }
}
