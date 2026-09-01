/*************************************************************
 *  MÓDULO PDF — Lee los temarios/guías de cada asignatura
 *  Las carpetas se crean en el Drive de la MISMA cuenta que
 *  corre el script (Emmanuel), así no hay problemas de permisos.
 *
 *  REQUISITO: Drive API activada (Servicios ＋ → Drive API).
 *
 *  USO:
 *   1) Ejecuta crearCarpetasAsignaturas() UNA vez → crea la
 *      estructura y guarda los IDs en la pestaña "Carpetas".
 *   2) Arrastra cada PDF a su carpeta.
 *   3) Menú → 📄 Procesar PDFs de asignaturas.
 *************************************************************/

const ASIGNATURAS_PDF = [
  'Matemática','Lenguaje y Comunicación','Ciencias Naturales',
  'Historia, Geografía y Ciencias Sociales','Inglés','Religión',
  'Educación física','Artes visuales','Música','Tecnología',
  'Lengua y cultura de los pueblos originarios'
];

// Crea la carpeta madre y las subcarpetas en el Drive de Emmanuel, y guarda IDs en "Carpetas"
function crearCarpetasAsignaturas(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const hC=crearHoja_(ss,'Carpetas',['Asignatura','CarpetaId','Link']);

  // ¿ya existe la carpeta madre? (evita duplicar)
  let madre;
  const existentes=DriveApp.getFoldersByName('EMMANUEL - Materiales');
  madre = existentes.hasNext() ? existentes.next()
                               : DriveApp.createFolder('EMMANUEL - Materiales');

  // limpia y reescribe el mapeo
  if(hC.getLastRow()>1) hC.getRange(2,1,hC.getLastRow()-1,3).clearContent();

  ASIGNATURAS_PDF.forEach(asig=>{
    let sub;
    const ya=madre.getFoldersByName(asig);
    sub = ya.hasNext() ? ya.next() : madre.createFolder(asig);
    hC.appendRow([asig, sub.getId(), sub.getUrl()]);
  });

  SpreadsheetApp.getUi().alert(
    'Carpetas listas en el Drive de Emmanuel dentro de "EMMANUEL - Materiales".\n\n'+
    'Abre esa carpeta, arrastra cada PDF a su asignatura, y luego usa '+
    '"📄 Procesar PDFs de asignaturas".\n\nEnlace madre:\n'+madre.getUrl());
}

// Lee el mapeo asignatura→carpeta desde la pestaña "Carpetas"
function obtenerMapeoCarpetas_(){
  const h=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Carpetas');
  const mapa={};
  if(h && h.getLastRow()>1){
    h.getRange(2,1,h.getLastRow()-1,2).getValues().forEach(([asig,id])=>{
      if(asig && id) mapa[id]=asig;
    });
  }
  return mapa;
}

// Diagnóstico profundo: intenta procesar el primer PDF que encuentre y muestra el error exacto
function diagnosticarPDF(){
  const mapa=obtenerMapeoCarpetas_();
  for(const folderId in mapa){
    const asig=mapa[folderId];
    let carpeta;
    try{ carpeta=DriveApp.getFolderById(folderId); }catch(e){ continue; }
    const it=carpeta.getFiles();
    while(it.hasNext()){
      const f=it.next();
      const tipo=f.getBlob().getContentType();
      if(tipo==='application/pdf' || /\.pdf$/i.test(f.getName())){
        const tam=Math.round(f.getSize()/1024);
        _ultimoErrorGemini='';
        const r=analizarPDFconIA_(asig, f);
        SpreadsheetApp.getUi().alert(
          'Archivo: '+f.getName()+'\nAsignatura: '+asig+
          '\nTipo: '+tipo+'\nTamaño: '+tam+' KB\n\n'+
          (r ? '✅ Resumen:\n'+r.resumen : '❌ Error:\n'+(_ultimoErrorGemini||'desconocido')));
        return;
      }
    }
  }
  SpreadsheetApp.getUi().alert('No encontré ningún PDF en las carpetas.');
}

// Diagnóstico: muestra qué archivos ve en cada carpeta y su tipo real
function diagnosticarCarpetas(){
  const mapa=obtenerMapeoCarpetas_();
  if(!Object.keys(mapa).length){
    SpreadsheetApp.getUi().alert('No hay carpetas. Ejecuta primero crearCarpetasAsignaturas().');
    return;
  }
  let msg='';
  for(const folderId in mapa){
    const asig=mapa[folderId];
    let carpeta;
    try{ carpeta=DriveApp.getFolderById(folderId); }
    catch(e){ msg+='❌ '+asig+': '+e.message+'\n'; continue; }
    try{
      const it=carpeta.getFiles(); let lista=[];
      while(it.hasNext()){ const f=it.next(); lista.push(f.getName()); }
      msg+='📁 '+asig+' ('+lista.length+')'+(lista.length?': '+lista.join(', '):'')+'\n';
    }catch(e){ msg+='⚠️ '+asig+': '+e.message+'\n'; }
  }
  SpreadsheetApp.getUi().alert(msg);
}
function procesarPDFs(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const hMat=crearHoja_(ss,'ContenidoPDF',
    ['Asignatura','Archivo','Resumen del contenido','Objetivos / temas','Procesado el','FileId']);
  const yaVistos=new Set(hMat.getLastRow()>1
    ? hMat.getRange(2,6,hMat.getLastRow()-1,1).getValues().flat():[]);

  let nuevos=0, errores=0;
  const mapa=obtenerMapeoCarpetas_();
  if(!Object.keys(mapa).length){
    SpreadsheetApp.getUi().alert('No hay carpetas. Ejecuta primero crearCarpetasAsignaturas().');
    return;
  }
  for(const folderId in mapa){
    const asig=mapa[folderId];
    let carpeta;
    try{ carpeta=DriveApp.getFolderById(folderId); }catch(e){ continue; }
    const archivos=carpeta.getFiles();   // todos, luego filtramos por PDF

    while(archivos.hasNext()){
      const f=archivos.next();
      const tipo=f.getBlob().getContentType();
      const esPdf = tipo==='application/pdf' || /\.pdf$/i.test(f.getName());
      if(!esPdf) continue;
      const id=f.getId();
      if(yaVistos.has(id)) continue;

      const resumen=analizarPDFconIA_(asig, f);
      if(resumen){
        hMat.appendRow([asig, f.getName(), resumen.resumen||'',
                        (resumen.temas||[]).join(', '), new Date(), id]);
        // enlaza el contenido al tema de la asignatura
        enriquecerTema_(ss, asig, resumen.resumen);
        nuevos++;
      } else {
        errores++;
      }
      Utilities.sleep(1500); // respeta cuota
    }
  }
  SpreadsheetApp.getUi().alert('PDFs procesados: '+nuevos+
    (errores?('\nNo se pudieron leer: '+errores):''));
}

// Envía el PDF a Gemini (que lee documentos con imágenes) y pide resumen estructurado
function analizarPDFconIA_(asig, file){
  const key=obtenerClaveGemini_();
  if(!key) return null;

  let base64;
  try{ base64=Utilities.base64Encode(file.getBlob().getBytes()); }
  catch(e){ _ultimoErrorGemini='No se pudo leer el archivo: '+e.message; return null; }

  const instruccion=
    'Eres un asistente educativo. Lee este documento (temario/guía de la asignatura "'+asig+
    '", 4° básico) y responde SOLO con JSON válido:\n'+
    '- "resumen": 2-3 frases claras de qué contenido educativo cubre el documento.\n'+
    '- "temas": lista de los temas o conceptos principales que aparecen.\n'+
    'Escribe en español, para que un apoderado entienda qué está estudiando su hijo. '+
    'IMPORTANTE: responde únicamente el objeto JSON, sin texto antes ni después.';

  const payload={
    contents:[{ parts:[
      { inlineData:{ mimeType:'application/pdf', data:base64 } },
      { text:instruccion }
    ]}],
    generationConfig:{ temperature:0.2, maxOutputTokens:1000, responseMimeType:'application/json' }
  };
  const url='https://generativelanguage.googleapis.com/v1beta/models/'+
            obtenerModeloGemini_()+':generateContent';

  try{
    let resp, intentos=0;
    do{
      resp=UrlFetchApp.fetch(url,{ method:'post', contentType:'application/json',
        headers:{ 'x-goog-api-key':key }, payload:JSON.stringify(payload),
        muteHttpExceptions:true });
      const c=resp.getResponseCode();
      if(c===503||c===429){ intentos++; Utilities.sleep(3000); } else break;
    } while(intentos<3);

    if(resp.getResponseCode()!==200){
      _ultimoErrorGemini='HTTP '+resp.getResponseCode()+': '+resp.getContentText().slice(0,200);
      return null;
    }
    const data=JSON.parse(resp.getContentText());
    let salida=(data?.candidates?.[0]?.content?.parts?.[0]?.text||'').replace(/```json|```/g,'').trim();
    // extrae solo el primer bloque { ... } balanceado, ignorando texto sobrante
    const ini=salida.indexOf('{');
    if(ini>=0){
      let prof=0, fin=-1;
      for(let i=ini;i<salida.length;i++){
        if(salida[i]==='{') prof++;
        else if(salida[i]==='}'){ prof--; if(prof===0){ fin=i; break; } }
      }
      if(fin>ini) salida=salida.slice(ini,fin+1);
    }
    const obj=JSON.parse(salida);
    return { resumen:obj.resumen||'', temas:Array.isArray(obj.temas)?obj.temas:[] };
  }catch(e){
    _ultimoErrorGemini='Excepción PDF: '+e.message;
    return null;
  }
}

// Añade el contenido del PDF a la nota de la asignatura (columna extra)
function enriquecerTema_(ss, asig, resumen){
  if(!resumen) return;
  const h=ss.getSheetByName('Asignaturas');
  if(!h) return;
  // asegura columna "Contenido del material"
  const enc=h.getRange(1,1,1,h.getLastColumn()).getValues()[0];
  let col=enc.indexOf('Contenido del material')+1;
  if(col===0){ col=h.getLastColumn()+1; h.getRange(1,col).setValue('Contenido del material')
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff'); }
  const datos=h.getLastRow()>1?h.getRange(2,1,h.getLastRow()-1,1).getValues().flat():[];
  const i=datos.indexOf(asig);
  if(i>=0) h.getRange(i+2,col).setValue(recorte_(resumen,300));
}
