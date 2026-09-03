/*************************************************************
 * P2 — DOCUMENTOS
 * Cuando un profe sube material a Classroom, el sistema:
 *   1) descarga (copia) el archivo a la carpeta de la asignatura en Drive,
 *   2) la IA lee de qué trata (resumen + temas),
 *   3) lo guarda en la hoja "Documentos" (la base que luego alimenta los juegos).
 *
 * La descarga real ocurre durante "🔄 Sincronizar Classroom" (ahí llega el
 * objeto con los adjuntos). El botón "📎 Descargar documentos de Classroom"
 * vuelve a recorrer los cursos para bajar lo que aún no esté en la base.
 *
 * REQUISITO: Drive API activada (Servicios ＋ → Drive API) y las carpetas por
 * asignatura creadas (menú u hoja "Carpetas"). Si no existen, se crean solas.
 *************************************************************/

function hojaDocumentos_(ss){
  return ss.getSheetByName('Documentos') || crearHoja_(ss,'Documentos',
    ['Fecha','Asignatura','Examen','Archivo','Tipo','Link','Resumen','Temas','FileId','MaterialId','Estado']);
}

// asig -> carpeta Drive. Usa la hoja "Carpetas"; si no está, crea la estructura.
function carpetaDeAsignatura_(ss, asig){
  const h=ss.getSheetByName('Carpetas');
  if(h && h.getLastRow()>1){
    const d=h.getRange(2,1,h.getLastRow()-1,2).getValues();
    for(const [a,id] of d){
      if(String(a).trim()===asig && id){ try{ return DriveApp.getFolderById(String(id).trim()); }catch(e){} }
    }
  }
  const it=DriveApp.getFoldersByName('EMMANUEL - Materiales');
  const madre = it.hasNext()?it.next():DriveApp.createFolder('EMMANUEL - Materiales');
  const sub=madre.getFoldersByName(asig);
  return sub.hasNext()?sub.next():madre.createFolder(asig);
}

// Empareja un documento con un examen de la asignatura (por palabras del contenido).
// Devuelve el "contenido" del examen (para enlazarlo con examen_id/juegos) o ''.
function examenParaDocumento_(ss, asig, titulo, resumen){
  const h=ss.getSheetByName('Examenes'); if(!h || h.getLastRow()<2) return '';
  const datos=h.getRange(2,1,h.getLastRow()-1,2).getValues();
  const txt=((titulo||'')+' '+(resumen||'')).toLowerCase();
  let mejor='', pts=0;
  for(const [a,cont] of datos){
    if(a!==asig || !cont) continue;
    const palabras=String(cont).toLowerCase().split(/[\s,()]+/).filter(p=>p.length>4);
    let n=0; palabras.forEach(p=>{ if(txt.indexOf(p)>=0) n++; });
    if(n>pts){ pts=n; mejor=String(cont); }
  }
  return pts>0 ? mejor : '';
}

/*************************************************************
 * NÚCLEO P2 — procesa los materiales de UN anuncio de Classroom.
 * Se llama desde sincronizarClassroom con el objeto "announcement".
 * Devuelve cuántos documentos nuevos guardó.
 *************************************************************/
function procesarMaterialesAnuncio_(ss, asig, announcement, fecha){
  const materiales=(announcement && announcement.materials) || [];
  if(!materiales.length) return 0;

  const h=hojaDocumentos_(ss);
  const vistos=new Set(h.getLastRow()>1
    ? h.getRange(2,10,h.getLastRow()-1,1).getValues().flat().map(String) : []);  // col 10 = MaterialId
  let n=0, carpeta=null;

  materiales.forEach(m=>{
    // ---- 1) Archivo de Drive: se DESCARGA (copia) a la carpeta de la asignatura ----
    if(m.driveFile && m.driveFile.driveFile){
      const df=m.driveFile.driveFile;
      const matId='drive:'+df.id;
      if(vistos.has(matId)) return; vistos.add(matId);

      let titulo=df.title||'documento', link=df.alternateLink||'', resumen='', temas='', fileId='';
      try{
        if(!carpeta) carpeta=carpetaDeAsignatura_(ss, asig);
        const copia=DriveApp.getFileById(df.id).makeCopy(titulo, carpeta);
        fileId=copia.getId(); link=copia.getUrl();
        const r=analizarArchivoConIA_(asig, copia, copia.getMimeType());  // PDF, Docs, Slides, .docx, imágenes
        if(r){ resumen=r.resumen||''; temas=(r.temas||[]).join(', '); }
      }catch(e){ resumen='(no se pudo descargar: '+e.message+')'; }

      const examen=examenParaDocumento_(ss, asig, titulo, resumen);
      h.appendRow([fecha, asig, examen, titulo, 'Archivo', link, resumen, temas, fileId, matId,
                   resumen && resumen.indexOf('no se pudo')<0 ? 'Procesado' : 'Descargado']);
      n++;
    }
    // ---- 2) Enlace externo: se registra el link (no se descarga) ----
    else if(m.link && m.link.url){
      const matId='link:'+m.link.url;
      if(vistos.has(matId)) return; vistos.add(matId);
      const titulo=m.link.title||m.link.url;
      h.appendRow([fecha, asig, examenParaDocumento_(ss,asig,titulo,''), titulo, 'Enlace', m.link.url,
                   '', '', '', matId, 'Enlace']);
      n++;
    }
    // ---- 3) YouTube: registro simple ----
    else if(m.youtubeVideo && m.youtubeVideo.alternateLink){
      const matId='yt:'+m.youtubeVideo.id;
      if(vistos.has(matId)) return; vistos.add(matId);
      h.appendRow([fecha, asig, '', m.youtubeVideo.title||'Video', 'Video',
                   m.youtubeVideo.alternateLink, '', '', '', matId, 'Video']);
      n++;
    }
  });
  return n;
}

/*************************************************************
 * DIAGNÓSTICO — revisa 1 documento pendiente y muestra el error exacto.
 * Ejecútala desde el editor (▶) o desde el menú. NO cambia nada.
 *************************************************************/
function diagnosticarDocumento(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const h=ss.getSheetByName('Documentos');
  if(!h || h.getLastRow()<2){ SpreadsheetApp.getUi().alert('No hay documentos.'); return; }
  const enc=h.getRange(1,1,1,h.getLastColumn()).getValues()[0];
  const col=n=>enc.indexOf(n);
  const filas=h.getRange(2,1,h.getLastRow()-1,h.getLastColumn()).getValues();

  let fila=null;
  for(const r of filas){
    if(r[col('Tipo')]==='Archivo' && r[col('FileId')] && !String(r[col('Resumen')]||'').trim()){ fila=r; break; }
  }
  if(!fila){ SpreadsheetApp.getUi().alert('No hay archivos pendientes de resumen.'); return; }

  let msg='ARCHIVO: '+fila[col('Archivo')]+'\nAsignatura: '+fila[col('Asignatura')]+'\n';
  let file;
  try{ file=DriveApp.getFileById(String(fila[col('FileId')])); }
  catch(e){ SpreadsheetApp.getUi().alert(msg+'\n❌ No se pudo abrir el archivo: '+e.message); return; }
  const mime=file.getMimeType();
  msg+='MIME: '+mime+'\n';

  // 1) ¿Clave Gemini?
  const key=obtenerClaveGemini_();
  msg+='Clave Gemini: '+(key?'OK':'❌ FALTA')+'\n';
  msg+='Modelo: '+obtenerModeloGemini_()+'\n\n';

  // 2) ¿Se puede convertir/leer el archivo?
  let parte=null, convOk='';
  try{
    if(mime.indexOf('image/')===0){ parte={inlineData:{mimeType:mime,data:Utilities.base64Encode(file.getBlob().getBytes())}}; convOk='imagen directa OK'; }
    else if(mime==='application/pdf'){ parte={inlineData:{mimeType:'application/pdf',data:Utilities.base64Encode(file.getBlob().getBytes())}}; convOk='PDF directo OK'; }
    else { const pdf=file.getAs('application/pdf'); parte={inlineData:{mimeType:'application/pdf',data:Utilities.base64Encode(pdf.getBytes())}}; convOk='convertido a PDF OK'; }
  }catch(e){ SpreadsheetApp.getUi().alert(msg+'❌ FALLA EN LA CONVERSIÓN A PDF:\n'+e.message); return; }
  msg+='Conversión: '+convOk+'\n\n';

  // 3) Llamada a Gemini
  const payload={ contents:[{ parts:[ parte, { text:'Resume en JSON {"resumen":"...","temas":[]} en español.' } ] }],
    generationConfig:{ temperature:0.2, maxOutputTokens:800, responseMimeType:'application/json' } };
  const url='https://generativelanguage.googleapis.com/v1beta/models/'+obtenerModeloGemini_()+':generateContent';
  try{
    const resp=UrlFetchApp.fetch(url,{ method:'post', contentType:'application/json',
      headers:{ 'x-goog-api-key':key }, payload:JSON.stringify(payload), muteHttpExceptions:true });
    msg+='Gemini HTTP: '+resp.getResponseCode()+'\n';
    msg+='Respuesta: '+resp.getContentText().slice(0,400);
  }catch(e){ msg+='❌ Excepción llamando a Gemini: '+e.message; }
  SpreadsheetApp.getUi().alert(msg);
}

/*************************************************************
 * RESUMEN IA GENERAL — lee PDF, Google Docs/Slides, .docx e imágenes.
 * Convierte lo que puede a PDF y se lo pasa a Gemini; las imágenes van
 * directo. Devuelve { resumen, temas:[...] } o null si no se pudo leer.
 *************************************************************/
function analizarArchivoConIA_(asig, file, mime){
  const key=obtenerClaveGemini_();
  if(!key) return null;
  mime = mime || (function(){ try{ return file.getMimeType(); }catch(e){ return ''; } })();

  // Arma la "parte" que entiende Gemini (imagen directa, o PDF)
  let parte=null;
  try{
    if(mime && mime.indexOf('image/')===0){
      parte={ inlineData:{ mimeType:mime, data:Utilities.base64Encode(file.getBlob().getBytes()) } };
    } else if(mime==='application/pdf'){
      parte={ inlineData:{ mimeType:'application/pdf', data:Utilities.base64Encode(file.getBlob().getBytes()) } };
    } else {
      // Google Docs/Slides y Office (.docx/.pptx): exportar a PDF
      const pdf=file.getAs('application/pdf');
      parte={ inlineData:{ mimeType:'application/pdf', data:Utilities.base64Encode(pdf.getBytes()) } };
    }
  }catch(e){ _ultimoErrorGemini='No se pudo convertir a PDF: '+e.message; return null; }

  const instruccion=
    'Eres un asistente educativo. Este es material de la asignatura "'+asig+'" (4° básico). '+
    'Responde SOLO con JSON válido, sin texto ni ```:\n'+
    '- "resumen": 2-3 frases claras de qué contenido educativo cubre.\n'+
    '- "temas": lista de los temas o conceptos principales.\n'+
    'Escribe en español, para que un apoderado entienda qué estudia su hijo.';

  const payload={ contents:[{ parts:[ parte, { text:instruccion } ] }],
    generationConfig:{ temperature:0.2, maxOutputTokens:1000, responseMimeType:'application/json' } };
  const url='https://generativelanguage.googleapis.com/v1beta/models/'+obtenerModeloGemini_()+':generateContent';

  try{
    let resp, intentos=0;
    do{
      resp=UrlFetchApp.fetch(url,{ method:'post', contentType:'application/json',
        headers:{ 'x-goog-api-key':key }, payload:JSON.stringify(payload), muteHttpExceptions:true });
      const c=resp.getResponseCode();
      if(c===503||c===429){ intentos++; Utilities.sleep(3000); } else break;
    } while(intentos<3);
    if(resp.getResponseCode()!==200){ _ultimoErrorGemini='HTTP '+resp.getResponseCode(); return null; }
    const obj=extraerJSONAgente_(resp.getContentText());
    if(!obj) return null;
    return { resumen:obj.resumen||'', temas:Array.isArray(obj.temas)?obj.temas:[] };
  }catch(e){ _ultimoErrorGemini='Excepción: '+e.message; return null; }
}

/*************************************************************
 * BOTÓN DE MENÚ — resume los documentos ya descargados que aún no
 * tienen resumen (los que quedaron en estado "Descargado").
 *************************************************************/
function resumirDocumentosPendientes(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const h=ss.getSheetByName('Documentos');
  if(!h || h.getLastRow()<2){ try{ SpreadsheetApp.getUi().alert('No hay documentos.'); }catch(e){}; return 0; }

  const enc=h.getRange(1,1,1,h.getLastColumn()).getValues()[0];
  const col=n=>enc.indexOf(n);
  const cAsig=col('Asignatura'), cExamen=col('Examen'), cArchivo=col('Archivo'),
        cTipo=col('Tipo'), cResumen=col('Resumen'), cTemas=col('Temas'),
        cFileId=col('FileId'), cEstado=col('Estado');

  const filas=h.getRange(2,1,h.getLastRow()-1,h.getLastColumn()).getValues();
  let n=0, err=0;
  for(let i=0;i<filas.length;i++){
    const r=filas[i];
    if(String(r[cResumen]||'').trim()) continue;                 // ya tiene resumen
    if(r[cTipo]!=='Archivo' || !r[cFileId]) continue;            // enlaces/videos no se resumen
    let file; try{ file=DriveApp.getFileById(String(r[cFileId])); }catch(e){ err++; continue; }
    const res=analizarArchivoConIA_(r[cAsig], file, file.getMimeType());
    if(res && res.resumen){
      h.getRange(i+2,cResumen+1).setValue(res.resumen);
      h.getRange(i+2,cTemas+1).setValue((res.temas||[]).join(', '));
      const ex=examenParaDocumento_(ss, r[cAsig], r[cArchivo], res.resumen);
      if(ex && !String(r[cExamen]||'').trim()) h.getRange(i+2,cExamen+1).setValue(ex);
      h.getRange(i+2,cEstado+1).setValue('Procesado');
      n++;
    } else err++;
    Utilities.sleep(1200);                                        // respeta cuota Gemini
  }
  try{ SpreadsheetApp.getUi().alert('Documentos resumidos: '+n+(err?('\nNo se pudieron leer: '+err):'')+
    '\n\nSi quedaron pendientes (límite de tiempo), vuelve a ejecutar y continúa donde quedó.'); }catch(e){}
  return n;
}

/*************************************************************
 * BOTÓN DE MENÚ — recorre Classroom y baja los documentos que falten.
 * Útil para traer lo de anuncios ya existentes (la descarga necesita
 * volver a Classroom porque el archivo vive en el Drive del profe).
 *************************************************************/
function descargarDocumentosClassroom(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let cursos;
  try{ cursos=(Classroom.Courses.list({courseStates:['ACTIVE']}).courses)||[]; }
  catch(e){ try{ SpreadsheetApp.getUi().alert('Activa el servicio Classroom API (Servicios ＋).'); }catch(_){}; return 0; }

  let total=0;
  cursos.forEach(curso=>{
    const asig=mapearAsignatura_(curso.name);
    let anuncios=[];
    try{
      let pageToken=null, vueltas=0;
      do{
        const resp=Classroom.Courses.Announcements.list(curso.id,{
          pageSize:100, orderBy:'updateTime desc', pageToken:pageToken||undefined });
        (resp.announcements||[]).forEach(a=>anuncios.push(a));
        pageToken=resp.nextPageToken; vueltas++;
      } while(pageToken && vueltas<5);
    }catch(e){ return; }

    anuncios.forEach(a=>{
      if(new Date(a.creationTime) < new Date(ANIO,6,1)) return;   // solo desde julio 2026
      total += procesarMaterialesAnuncio_(ss, asig, a, new Date(a.creationTime));
    });
  });

  try{ SpreadsheetApp.getUi().alert('Documentos nuevos descargados y resumidos: '+total); }catch(e){}
  return total;
}

// Lectura para la app (pestaña "Documentos descargables")
function leerDocumentos_(ss){
  const h=ss.getSheetByName('Documentos'); if(!h || h.getLastRow()<2) return [];
  const enc=h.getRange(1,1,1,h.getLastColumn()).getValues()[0];
  let filas=h.getRange(2,1,h.getLastRow()-1,h.getLastColumn()).getValues();
  filas.sort((a,b)=> new Date(b[0]) - new Date(a[0]));         // más recientes primero
  return filas.map(f=>{
    const o={}; enc.forEach((c,i)=>{ if(c!=='FileId' && c!=='MaterialId') o[c]=f[i] instanceof Date?f[i].toISOString():f[i]; });
    return o;
  });
}
