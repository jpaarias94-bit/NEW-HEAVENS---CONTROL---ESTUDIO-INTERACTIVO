/*************************************************************
 *  MÓDULO IA (Gemini) — Entiende los anuncios de Classroom
 *
 *  REQUISITO: pega tu API key de Gemini en la pestaña "Config"
 *  de la hoja: fila con Clave = "gemini_api_key", Valor = tu clave.
 *  (Se saca gratis en https://aistudio.google.com → Get API key)
 *************************************************************/

const GEMINI_MODELO_DEFECTO = 'gemini-3.6-flash';  // modelo vigente (2.0 fue retirado por Google)

function obtenerModeloGemini_(){
  const h=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if(h && h.getLastRow()>1){
    const datos=h.getRange(2,1,h.getLastRow()-1,2).getValues();
    for(const [c,v] of datos){
      if(String(c).trim()==='gemini_modelo' && v) return String(v).trim();
    }
  }
  return GEMINI_MODELO_DEFECTO;
}

// Lee la API key desde la pestaña Config
function obtenerClaveGemini_(){
  const h=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if(!h || h.getLastRow()<2) return null;
  const datos=h.getRange(2,1,h.getLastRow()-1,2).getValues();
  for(const [clave,valor] of datos){
    if(String(clave).trim()==='gemini_api_key' && valor) return String(valor).trim();
  }
  return null;
}

/**
 * Envía un anuncio a Gemini y devuelve un objeto estructurado:
 * { tema, materiales:[...], fecha:'YYYY-MM-DD'|null, reagenda:bool, tipo }
 * Si la IA falla o no hay clave, devuelve null (el sistema usa las reglas de respaldo).
 */
// Devuelve los contenidos de exámenes de una asignatura (para que la IA empareje)
function examenesDeAsignatura_(asig){
  const h=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Examenes');
  if(!h || h.getLastRow()<2) return [];
  const datos=h.getRange(2,1,h.getLastRow()-1,2).getValues();
  return datos.filter(r=>r[0]===asig && r[1]).map(r=>String(r[1]));
}

function analizarConIA_(asignatura, textoAnuncio){
  const key=obtenerClaveGemini_();
  if(!key) return null;

  const listaExamenes=examenesDeAsignatura_(asignatura);
  const bloqueExam = listaExamenes.length
    ? '\n\nExámenes planificados de esta asignatura (por si el anuncio mueve alguno):\n- '+
      listaExamenes.join('\n- ')+'\n'
    : '';

  const instruccion =
    'Eres un asistente que lee anuncios de profesores de un colegio (4° básico, Chile) '+
    'y extrae información para un panel de control académico. '+
    'Responde SOLO con un JSON válido, sin texto adicional ni ```.\n\n'+
    'Del siguiente anuncio de la asignatura "'+asignatura+'", extrae:\n'+
    '- "tema": el contenido o materia que se está viendo, en una frase corta y clara '+
    '(ej: "Fracciones y decimales"). Si no se menciona un tema, usa null.\n'+
    '- "materiales": lista de materiales que los estudiantes deben traer '+
    '(ej: ["libro de ciencias","plasticina"]). Lista vacía si no hay.\n'+
    '- "fecha": si se menciona una fecha de evaluación, prueba o entrega, en formato '+
    '"YYYY-MM-DD" (año 2026). null si no hay fecha.\n'+
    '- "reagenda": true si el anuncio CAMBIA o mueve una fecha ya existente '+
    '(reagenda, reprograma, "quedará para", "se traslada"). false si no.\n'+
    '- "examen_contenido": si el anuncio se refiere a un examen de la lista de abajo, '+
    'copia EXACTAMENTE el contenido de ese examen (tal como aparece en la lista). '+
    'Si no corresponde a ninguno, usa null.\n'+
    '- "tipo": una de estas palabras: "Tema", "Evaluación", "Reagendado", '+
    '"Suspensión", "Material", "Aviso".'+
    bloqueExam+
    '\n\nANUNCIO:\n"""'+textoAnuncio+'"""';

  const url='https://generativelanguage.googleapis.com/v1beta/models/'+
            obtenerModeloGemini_()+':generateContent';
  const payload={ contents:[{ parts:[{ text:instruccion }] }],
                  generationConfig:{
                    temperature:0.1,
                    maxOutputTokens:800,
                    responseMimeType:'application/json'   // obliga a devolver SOLO JSON
                  } };

  try{
    let resp, intentos=0;
    do{
      resp=UrlFetchApp.fetch(url,{
        method:'post', contentType:'application/json',
        headers:{ 'x-goog-api-key': key },
        payload:JSON.stringify(payload), muteHttpExceptions:true
      });
      const c=resp.getResponseCode();
      if(c===503 || c===429){ intentos++; Utilities.sleep(2500); }
      else break;
    } while(intentos<3);

    if(resp.getResponseCode()!==200){
      _ultimoErrorGemini = 'HTTP '+resp.getResponseCode()+': '+resp.getContentText().slice(0,300);
      return null;
    }
    const raw=resp.getContentText();
    let data;
    try{ data=JSON.parse(raw); }
    catch(e){ _ultimoErrorGemini='Respuesta no-JSON: '+raw.slice(0,300); return null; }

    let salida=data?.candidates?.[0]?.content?.parts?.[0]?.text || '';
    if(!salida){ _ultimoErrorGemini='Respuesta sin texto: '+raw.slice(0,300); return null; }

    // limpia fences y extrae el bloque { ... } aunque venga con texto alrededor
    salida=salida.replace(/```json|```/g,'').trim();
    const ini=salida.indexOf('{'), fin=salida.lastIndexOf('}');
    if(ini>=0 && fin>ini) salida=salida.slice(ini,fin+1);

    let obj;
    try{ obj=JSON.parse(salida); }
    catch(e){ _ultimoErrorGemini='JSON del modelo inválido: '+salida.slice(0,200); return null; }

    return {
      tema: obj.tema || null,
      materiales: Array.isArray(obj.materiales)?obj.materiales:[],
      fecha: obj.fecha || null,
      reagenda: !!obj.reagenda,
      examenContenido: obj.examen_contenido || null,
      tipo: obj.tipo || 'Aviso'
    };
  }catch(e){
    _ultimoErrorGemini = 'Excepción: '+e.message;
    return null;
  }
}

var _ultimoErrorGemini = '';

/**
 * Versión IA de procesarTexto_. Intenta con Gemini; si no hay clave o falla,
 * cae automáticamente a la extracción por reglas (procesarTexto_).
 */
function procesarTextoIA_(ss, asig, texto, fecha){
  const ia=analizarConIA_(asig, texto);
  if(!ia){ procesarTexto_(ss, asig, texto, fecha); return null; }  // respaldo

  // TEMA
  if(ia.tema) actualizarAsignatura_(ss, asig, 'Estamos viendo: '+ia.tema, fecha);

  // MATERIALES
  (ia.materiales||[]).forEach(m=>{
    if(m && String(m).length>2) registrarMaterial_(ss, asig, cap_(String(m)), fecha);
  });

  // AVISO (registro), pero la fecha NO se ajusta sola: se sugiere y el usuario confirma
  if(ia.fecha || ia.tipo==='Suspensión' || ia.reagenda){
    ss.getSheetByName('Avisos').appendRow([fecha, asig, texto, ia.tipo||'Aviso']);
  }

  // devuelve la sugerencia para guardarla en Anuncios
  return { fecha: ia.fecha||'', examenContenido: ia.examenContenido||'' };
}

function parsearFechaISO_(s){
  const m=String(s).match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if(!m) return null;
  return new Date(parseInt(m[1]), parseInt(m[2])-1, parseInt(m[3]));
}

// Lista los modelos que TU clave puede usar. Ejecuta y mira el recuadro.
function listarModelosGemini(){
  const key=obtenerClaveGemini_();
  if(!key){ SpreadsheetApp.getUi().alert('Falta la clave en Config.'); return; }
  try{
    const resp=UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models',
      { method:'get', headers:{ 'x-goog-api-key': key }, muteHttpExceptions:true });
    if(resp.getResponseCode()!==200){
      SpreadsheetApp.getUi().alert('Error '+resp.getResponseCode()+': '+resp.getContentText().slice(0,300));
      return;
    }
    const data=JSON.parse(resp.getContentText());
    const nombres=(data.models||[])
      .filter(m=>(m.supportedGenerationMethods||[]).indexOf('generateContent')>=0)
      .map(m=>m.name.replace('models/',''))
      .filter(n=>n.indexOf('flash')>=0 || n.indexOf('pro')>=0);
    SpreadsheetApp.getUi().alert('Modelos disponibles para tu clave:\n\n'+
      (nombres.length?nombres.join('\n'):'(ninguno con generateContent)')+
      '\n\nCopia uno y ponlo en Config → gemini_modelo.');
  }catch(e){
    SpreadsheetApp.getUi().alert('Excepción: '+e.message);
  }
}

// Prueba rápida: ejecuta esto tras poner la clave para verificar que Gemini responde
function probarGemini(){
  _ultimoErrorGemini='';
  const r=analizarConIA_('Matemática',
    'Buenas tardes estimados. Les comparto las guías con el contenido de fracciones. '+
    'La evaluación quedará para el 14 de agosto. Favor traer su libro de matemáticas.');
  if(r){
    SpreadsheetApp.getUi().alert('✅ Gemini funciona:\n\n'+JSON.stringify(r,null,2));
  } else {
    SpreadsheetApp.getUi().alert('❌ No respondió.\n\nDetalle: '+
      (_ultimoErrorGemini||'sin detalle')+
      '\n\nRevisa que la clave esté en Config (gemini_api_key).');
  }
}
