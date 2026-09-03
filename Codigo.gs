/*************************************************************
 * SISTEMA DE CONTROL ACADÉMICO - EMMANUEL · v2
 * Fuente: API de Google Classroom (texto limpio del Tablón)
 * Se ejecuta desde la cuenta de Emmanuel.
 *
 * REQUISITO: activar el servicio "Classroom API" en el editor:
 * Servicios (＋) → Google Classroom API → Agregar.
 *************************************************************/

const ANIO = 2026;
const MESES = {enero:0,febrero:1,marzo:2,abril:3,mayo:4,junio:5,julio:6,
               agosto:7,septiembre:8,octubre:9,noviembre:10,diciembre:11};

// ====== INSTALACIÓN (ejecutar UNA vez) ======
function configurarSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  crearHoja_(ss,'Examenes',
    ['Asignatura','Contenido','Tipo','Instrumento','Porcentaje',
     'Fecha plan inicio','Fecha plan término','Fecha confirmada','¿Movida?','Estado']);
  crearHoja_(ss,'Asignaturas',['Asignatura','Tema actual','Última actualización','Fuente']);
  crearHoja_(ss,'Materiales',['Asignatura','Qué traer','Para cuándo','Detectado el','Estado']);
  crearHoja_(ss,'Avisos',['Fecha','Asignatura','Aviso','Tipo']);
  crearHoja_(ss,'Anuncios',['Fecha','Asignatura','Profesor','Texto','ID','Adjuntos','Link','ExamenFecha','ExamenContenido']);
  crearHoja_(ss,'Config',['Clave','Valor']);
  prepararConfig_(ss);
  cargarCalendarioBase_();
  crearMenu_();
  SpreadsheetApp.getUi().alert('Listo. Ahora: Servicios ＋ → Classroom API → Agregar. Luego usa "🔄 Sincronizar Classroom".');
}

function onOpen(){ crearMenu_(); }
function crearMenu_(){
  SpreadsheetApp.getUi().createMenu('📚 Control Académico')
    .addItem('🔄 Sincronizar Classroom','sincronizarClassroom')
    .addItem('🕵️ Revisar correos (proponer cambios)','ejecutarAgente')
    .addItem('♻️ Restaurar plan oficial del colegio','restaurarPlanOficial')
    .addItem('♻️ Reprocesar temas y materiales','reprocesarAnuncios')
    .addItem('📎 Descargar documentos de Classroom','descargarDocumentosClassroom')
    .addItem('🧠 Resumir documentos pendientes','resumirDocumentosPendientes')
    .addItem('📄 Procesar PDFs de asignaturas','procesarPDFs')
    .addItem('🗂️ Crear carpetas por examen (Drive)','crearCarpetasPorExamen')
    .addItem('🧹 Limpiar anuncios duplicados','limpiarDuplicados')
    .addItem('🤖 Probar conexión Gemini','probarGemini')
    .addItem('📅 Recargar calendario base','cargarCalendarioBase_')
    .addItem('🎯 Ver contenido del banco de juegos','resumenBanco')
    .addItem('🌐 Ver interfaz (URL)','mostrarURL')
    .addSeparator()
    .addItem('⚙️ Activar actualización automática','instalarAutomatizacion')
    .addItem('⏸️ Desactivar actualización automática','desactivarAutomatizacion')
    .addToUi();
}

// Reaplica la extracción mejorada a los anuncios YA guardados (sin volver a Classroom).
// Limpia Asignaturas y Materiales y los reconstruye desde la pestaña Anuncios.
function reprocesarAnuncios(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const hAnun=ss.getSheetByName('Anuncios');
  if(!hAnun || hAnun.getLastRow()<2){
    SpreadsheetApp.getUi().alert('No hay anuncios guardados. Ejecuta primero "Sincronizar Classroom".');
    return;
  }
  // limpia lo derivado (no toca Anuncios ni Examenes)
  ['Asignaturas','Materiales','Avisos'].forEach(n=>{
    const h=ss.getSheetByName(n);
    if(h && h.getLastRow()>1) h.getRange(2,1,h.getLastRow()-1,h.getLastColumn()).clearContent();
  });

  const filas=hAnun.getRange(2,1,hAnun.getLastRow()-1,4).getValues(); // Fecha,Asignatura,Profesor,Texto
  let n=0;
  filas.forEach(([fecha,asig,prof,texto])=>{
    if(texto){ procesarTextoIA_(ss,asig,String(texto),fecha); n++; }
  });
  ordenarExamenes_(ss);
  SpreadsheetApp.getUi().alert('Reprocesados '+n+' anuncios con la extracción mejorada.');
}

function crearHoja_(ss,nombre,enc){
  let h=ss.getSheetByName(nombre); if(!h) h=ss.insertSheet(nombre);
  if(h.getLastRow()===0){
    h.getRange(1,1,1,enc.length).setValues([enc])
     .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
    h.setFrozenRows(1);
  }
  return h;
}

// Deja lista la fila para pegar la API key de Gemini
function prepararConfig_(ss){
  const h=ss.getSheetByName('Config');
  const datos=h.getLastRow()>1?h.getRange(2,1,h.getLastRow()-1,1).getValues().flat():[];
  if(datos.indexOf('gemini_api_key')===-1){
    h.appendRow(['gemini_api_key','PEGA_AQUI_TU_CLAVE']);
  }
}

// ====== SINCRONIZACIÓN CON CLASSROOM ======
function sincronizarClassroom(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const hAnun=ss.getSheetByName('Anuncios');
  const yaVistos=new Set(hAnun.getLastRow()>1
    ? hAnun.getRange(2,5,hAnun.getLastRow()-1,1).getValues().flat().map(String):[]);

  const cursos=(Classroom.Courses.list({courseStates:['ACTIVE']}).courses)||[];
  let nuevos=0;

  cursos.forEach(curso=>{
    const asig=mapearAsignatura_(curso.name);
    let anuncios=[];
    try {
      let pageToken=null, vueltas=0;
      do{
        const resp=Classroom.Courses.Announcements.list(curso.id,{
          pageSize:100,
          orderBy:'updateTime desc',
          pageToken:pageToken||undefined
        });
        (resp.announcements||[]).forEach(a=>anuncios.push(a));
        pageToken=resp.nextPageToken;
        vueltas++;
      } while(pageToken && vueltas<5);
    } catch(e){ return; }

    anuncios.forEach(a=>{
      const id=String(a.id);
      if(!a.text || yaVistos.has(id)) return;
      yaVistos.add(id); // ← marca YA en esta misma corrida para no repetir

      const fecha=new Date(a.creationTime);
      if(fecha < new Date(ANIO,6,1)) return; // solo desde julio 2026

      const texto=a.text.trim();
      const numAdj=(a.materials||[]).length;
      const link=a.alternateLink||'';

      // IA analiza: tema, materiales y SUGERENCIA de examen (no ajusta la fecha sola)
      const sug=procesarTextoIA_(ss,asig,texto,fecha);
      const exFecha = (sug && sug.fecha) ? sug.fecha : '';
      const exCont  = (sug && sug.examenContenido) ? sug.examenContenido : '';

      hAnun.appendRow([fecha,asig,curso.name,texto,id,numAdj,link,exFecha,exCont]);
      nuevos++;

      // P2: descarga a Drive + resumen IA de los adjuntos de este anuncio
      try{ if(typeof procesarMaterialesAnuncio_==='function') procesarMaterialesAnuncio_(ss, asig, a, fecha); }catch(e){}
    });
  });

  ordenarExamenes_(ss);
  // El agente revisa los correos nuevos y aplica/registra cambios detectados
  try{ if(typeof correrAgente_==='function') correrAgente_(ss); }catch(e){}
  try{ SpreadsheetApp.getUi().alert('Sincronización lista. Anuncios nuevos: '+nuevos); }
  catch(e){ /* corre en activador automático, sin UI: no pasa nada */ }
}

// ====== PROCESAMIENTO INTELIGENTE DEL TEXTO ======
function procesarTexto_(ss,asig,texto,fecha){
  if(!asig) return;
  const t=texto.toLowerCase();

  const tema=extraerTema_(texto, asig);
  if(tema) actualizarAsignatura_(ss,asig,tema,fecha);

  const mats=extraerMateriales_(texto);
  mats.forEach(m=>registrarMaterial_(ss,asig,m,fecha));

  const info=extraerEventoFecha_(t);
  if(info && info.fecha){
    ajustarExamen_(ss,asig,info.fecha);
    const tipo = info.reagenda ? 'Reagendado' : (info.evaluacion?'Evaluación':'Aviso');
    ss.getSheetByName('Avisos').appendRow([fecha,asig,texto,tipo]);
  } else if(/no se realizará|suspend|cancel/.test(t)){
    ss.getSheetByName('Avisos').appendRow([fecha,asig,texto,'Suspensión']);
  }
}

function extraerTema_(texto, asig){
  const t=texto.replace(/\s+/g,' ');
  const tl=t.toLowerCase();

  // 1) PATRONES EXPLÍCITOS de tema (lo más confiable)
  const patrones=[
    /contenidos?\s+(?:de|del|sobre)\s+(?:la\s+|el\s+|los\s+|las\s+)?(?:unidad\s+)?["“']?([^.,;:"”'()]{3,60})/i,
    /(?:la\s+)?unidad\s+["“']?([^.,;:"”'()]{3,60})/i,
    /estamos\s+(?:viendo|trabajando|revisando)\s+(?:la\s+|el\s+|los\s+|las\s+)?([^.,;:"”'()]{3,60})/i,
    /(?:evaluaci[óo]n|prueba|trabajo)\s+(?:de|sobre)\s+(?:la\s+|el\s+)?(?:unidad\s+)?["“']?([^.,;:"”'()]{3,60})/i,
    /temario\s+(?:con\s+el\s+)?contenido\s+(?:de|del)\s+([^.,;:"”'()]{3,60})/i
  ];
  for(const re of patrones){
    const m=t.match(re);
    if(m && m[1]){
      let tema=m[1].trim().replace(/\s+(que|el cual|la cual|realizad[ao]s?).*$/i,'').trim();
      if(tema.length>2 && !/^(guías?|clases?|actividades?)$/i.test(tema))
        return 'Estamos viendo: '+cap_(tema);
    }
  }

  // 2) CRUCE CON EL CALENDARIO
  const temaPlan=cruzarConCalendario_(tl, asig);
  if(temaPlan) return 'Estamos viendo: '+temaPlan;

  // 3) RESPALDO: primera frase con contenido
  let limpio=t
    .replace(/^(buenas?\s+(tardes?|d[ií]as?)|buen\s+d[ií]a)[^.]*\.?\s*/i,'')
    .replace(/estimad[oa]s?\s+y?\s*estimad[oa]s?[,.:]?\s*/i,'')
    .replace(/estimad[oa]s?\([^)]*\)[.,:]?\s*/i,'')
    .replace(/estimad[oa]s?[,.:]?\s*/i,'')
    .replace(/esperando que se encuentren?\s+bien[,.]?\s*/i,'')
    .replace(/les\s+(comparto|informo|comunico|escribo|recuerdo)\s*/i,'');
  const frase=limpio.split(/(?<=\.)\s/)[0];
  const r=recorte_(frase.trim(),120);
  return r && r.length>4 ? r : null;
}

// Busca en el calendario base si algún contenido de esa asignatura aparece en el texto
function cruzarConCalendario_(textoLower, asig){
  try{
    const h=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Examenes');
    if(!h || h.getLastRow()<2) return null;
    const datos=h.getRange(2,1,h.getLastRow()-1,2).getValues(); // [Asignatura, Contenido]
    for(const [a,contenido] of datos){
      if(a!==asig || !contenido) continue;
      const palabras=String(contenido).toLowerCase().split(/[\s,()]+/).filter(p=>p.length>4);
      for(const p of palabras){
        if(textoLower.includes(p)) return contenido; // nombre oficial del contenido
      }
    }
  }catch(e){}
  return null;
}

function extraerMateriales_(texto){
  const out=[]; const t=texto.replace(/\s+/g,' ');
  const re=/(?:deber[áa]n?\s+)?tra(?:er|igan|erán|erá|iga)\s+([^.,;]{3,90})/gi;
  let m;
  while((m=re.exec(t))){
    let item=m[1].trim()
      .replace(/^(su|sus|el|la|los|las|un|una|los siguientes|el siguiente)\s+/i,'')
      .replace(/\s+(para|de regreso|igualmente|indicad[oa]s|a clases?|para clases?).*/i,'')
      .trim();
    if(item.length>2 && !/^(materiales?|lo siguiente|esto)$/i.test(item))
      out.push(cap_(item));
  }
  return [...new Set(out)];
}

function extraerEventoFecha_(t){
  const reagenda=/(se reagendar|reagendad|quedar[áa] para|se realizar[áa] el|pasa para|se traslada|nueva fecha)/.test(t);
  const evaluacion=/(evaluaci[óo]n|prueba|test|control|disertaci[óo]n|trabajo pr[áa]ctico)/.test(t);
  let fecha=null;
  let m=t.match(/(\d{1,2})\s+de\s+(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)/);
  if(m) fecha=new Date(ANIO,MESES[m[2]],parseInt(m[1]));
  if(!fecha){
    m=t.match(/(\d{1,2})\/(\d{1,2})/);
    if(m) fecha=new Date(ANIO,parseInt(m[2])-1,parseInt(m[1]));
  }
  return {fecha,reagenda,evaluacion};
}

function actualizarAsignatura_(ss,asig,tema,fecha){
  const h=ss.getSheetByName('Asignaturas');
  const datos=h.getLastRow()>1?h.getRange(2,1,h.getLastRow()-1,1).getValues().flat():[];
  const i=datos.indexOf(asig);
  if(i===-1) h.appendRow([asig,tema,fecha,'Classroom']);
  else h.getRange(i+2,2,1,3).setValues([[tema,fecha,'Classroom']]);
}

function registrarMaterial_(ss,asig,item,fecha){
  const h=ss.getSheetByName('Materiales');
  const dup=h.getLastRow()>1 && h.getRange(2,1,h.getLastRow()-1,2).getValues()
    .some(r=>r[0]===asig && String(r[1]).toLowerCase()===item.toLowerCase());
  if(!dup) h.appendRow([asig,item,'',fecha,'Pendiente']);
}

function ajustarExamen_(ss,asig,fecha){
  const h=ss.getSheetByName('Examenes'); if(h.getLastRow()<2) return;
  const datos=h.getRange(2,1,h.getLastRow()-1,10).getValues();
  for(let i=0;i<datos.length;i++){
    if(datos[i][0]===asig){
      const plan=new Date(datos[i][5]);
      const movida=Math.abs(fecha-plan)>86400000;
      h.getRange(i+2,8).setValue(fecha);
      h.getRange(i+2,9).setValue(movida?'SÍ':'No');
      return;
    }
  }
}

// Ajusta el examen que coincide por asignatura Y contenido (más preciso)
function ajustarExamenPorContenido_(ss,asig,contenido,fecha){
  const h=ss.getSheetByName('Examenes'); if(h.getLastRow()<2) return;
  const datos=h.getRange(2,1,h.getLastRow()-1,10).getValues();
  const cLow=String(contenido).toLowerCase().trim();
  // 1) match exacto de contenido
  for(let i=0;i<datos.length;i++){
    if(datos[i][0]===asig && String(datos[i][1]).toLowerCase().trim()===cLow){
      return escribirAjuste_(h,i,datos[i],fecha);
    }
  }
  // 2) match parcial (contiene)
  for(let i=0;i<datos.length;i++){
    if(datos[i][0]===asig && String(datos[i][1]).toLowerCase().includes(cLow.slice(0,15))){
      return escribirAjuste_(h,i,datos[i],fecha);
    }
  }
  // 3) respaldo: primer examen de la asignatura
  ajustarExamen_(ss,asig,fecha);
}

function escribirAjuste_(h,i,fila,fecha){
  const plan=new Date(fila[5]);
  const movida=Math.abs(fecha-plan)>86400000;
  h.getRange(i+2,8).setValue(fecha);
  h.getRange(i+2,9).setValue(movida?'SÍ':'No');
}

function ordenarExamenes_(ss){
  const h=ss.getSheetByName('Examenes'); if(h.getLastRow()<3) return;
  h.getRange(2,1,h.getLastRow()-1,10).sort({column:6,ascending:true});
}

function recorte_(s,n){ s=String(s).replace(/\s+/g,' ').trim(); return s.length>n?s.slice(0,n)+'…':s; }
function cap_(s){ return s.charAt(0).toUpperCase()+s.slice(1); }

function mapearAsignatura_(nombreCurso){
  const t=(nombreCurso||'').toLowerCase();
  if(/matem/.test(t)) return 'Matemática';
  if(/lenguaje|lengua y com/.test(t)) return 'Lenguaje y Comunicación';
  if(/ciencias natural|naturales/.test(t)) return 'Ciencias Naturales';
  if(/historia|geograf|sociales/.test(t)) return 'Historia, Geografía y Ciencias Sociales';
  if(/ingl[ée]s|english/.test(t)) return 'Inglés';
  if(/religi/.test(t)) return 'Religión';
  if(/f[íi]sica|ed\.?\s*f[íi]sica/.test(t)) return 'Educación física';
  if(/arte/.test(t)) return 'Artes visuales';
  if(/m[úu]sica/.test(t)) return 'Música';
  if(/tecnolog/.test(t)) return 'Tecnología';
  if(/originarios|ancestral|lengua ind/.test(t)) return 'Lengua y cultura de los pueblos originarios';
  return nombreCurso;
}

function doGet(){
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Control Académico - Emmanuel')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}
function obtenerDatos(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  return {
    examenes:leerHoja_(ss,'Examenes'),
    anuncios:leerAnunciosLigero_(ss),
    documentos:(typeof leerDocumentos_==='function'?leerDocumentos_(ss):[]),
    agente:(typeof leerBitacoraAgente_==='function'?leerBitacoraAgente_(ss):[]),
    auto:(function(){ try{ return (typeof estadoAutomatizacion==='function')?estadoAutomatizacion():null; }catch(e){ return null; } })(),
    hoy:new Date().toISOString()
  };
}

// Anuncios livianos: solo los más recientes y con el texto recortado,
// para que la app cargue rápido sin colgarse.
function leerAnunciosLigero_(ss){
  const h=ss.getSheetByName('Anuncios');
  if(!h || h.getLastRow()<2) return [];
  const enc=h.getRange(1,1,1,h.getLastColumn()).getValues()[0];
  let filas=h.getRange(2,1,h.getLastRow()-1,h.getLastColumn()).getValues();
  // más recientes primero y limita a 150 (varios meses de correos)
  filas.sort((a,b)=> new Date(b[0]) - new Date(a[0]));
  filas=filas.slice(0,150);
  return filas.map(f=>{
    const o={};
    enc.forEach((c,i)=>{
      let v=f[i];
      if(v instanceof Date) v=v.toISOString();
      if(c==='Texto' && typeof v==='string' && v.length>600) v=v.slice(0,600)+'…';
      if(c!=='ID') o[c]=v; // el ID no lo necesita la interfaz
    });
    return o;
  });
}
function leerHoja_(ss,nombre){
  const h=ss.getSheetByName(nombre); if(!h||h.getLastRow()<2) return [];
  const enc=h.getRange(1,1,1,h.getLastColumn()).getValues()[0];
  return h.getRange(2,1,h.getLastRow()-1,h.getLastColumn()).getValues().map(f=>{
    const o={}; enc.forEach((c,i)=>o[c]=f[i] instanceof Date?f[i].toISOString():f[i]); return o;
  });
}
// Igual que leerHoja_ pero omite columnas pesadas/no usadas
function leerHojaLigera_(ss,nombre,omitir){
  const h=ss.getSheetByName(nombre); if(!h||h.getLastRow()<2) return [];
  const enc=h.getRange(1,1,1,h.getLastColumn()).getValues()[0];
  const skip=new Set(omitir||[]);
  return h.getRange(2,1,h.getLastRow()-1,h.getLastColumn()).getValues().map(f=>{
    const o={}; enc.forEach((c,i)=>{ if(!skip.has(c)) o[c]=f[i] instanceof Date?f[i].toISOString():f[i]; }); return o;
  });
}
// La web llama a esto al presionar "Aplicar fecha". Ajusta el examen y confirma.
function aplicarFechaExamen(asig, contenido, fechaISO){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const f=parsearFechaISO_(fechaISO);
  if(!f) return {ok:false, msg:'Fecha inválida'};
  if(contenido) ajustarExamenPorContenido_(ss, asig, contenido, f);
  else ajustarExamen_(ss, asig, f);
  ordenarExamenes_(ss);
  return {ok:true, msg:'Fecha aplicada a '+asig};
}

function mostrarURL(){
  const url=ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert('URL de tu interfaz:\n\n'+(url||'Primero implementa como Aplicación web.'));
}

// Elimina filas de Anuncios con ID repetido (deja la primera). Respaldo manual.
function limpiarDuplicados(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const h=ss.getSheetByName('Anuncios');
  if(!h||h.getLastRow()<2){SpreadsheetApp.getUi().alert('Sin anuncios.');return;}
  const datos=h.getRange(2,1,h.getLastRow()-1,h.getLastColumn()).getValues();
  const vistos=new Set(), conservar=[];
  datos.forEach(f=>{ const id=String(f[4]); if(id && !vistos.has(id)){ vistos.add(id); conservar.push(f); } });
  h.getRange(2,1,datos.length,h.getLastColumn()).clearContent();
  if(conservar.length) h.getRange(2,1,conservar.length,h.getLastColumn()).setValues(conservar);
  SpreadsheetApp.getUi().alert('Duplicados eliminados. Quedaron '+conservar.length+' anuncios únicos.');
}
