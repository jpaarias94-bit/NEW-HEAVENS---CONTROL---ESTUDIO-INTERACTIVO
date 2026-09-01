/*************************************************************
 * BANCO DE JUEGOS — contenido de estudio FIJO por examen.
 * NO usa IA. Tú y Claude arman las preguntas; aquí solo se guardan
 * y se sirven al azar para que cada partida se sienta distinta.
 *
 * Estructura de la hoja "BancoJuegos" (una fila por ítem):
 *   Asignatura | Contenido | Tipo | Dato1 | Dato2 | Dato3 | Dato4 | Correcta | Extra
 *
 * Según el Tipo, las columnas significan:
 *  - quiz:       Dato1=pregunta, Dato2..Dato5=4 opciones, Correcta=índice 0-3
 *  - flash:      Dato1=frente, Dato2=reverso
 *  - vf:         Dato1=afirmación, Correcta=VERDADERO/FALSO, Extra=porqué
 *  - completar:  Dato1=frase (con ___), Correcta=palabra
 *  - adivinanza: Dato1=pista, Correcta=respuesta
 *
 * "Contenido" debe coincidir con el Contenido del examen en la hoja Examenes
 * (así el botón Estudiar encuentra su banco).
 *************************************************************/

function configurarBancoJuegos(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  crearHoja_(ss,'BancoJuegos',
    ['Asignatura','Contenido','Tipo','Dato1','Dato2','Dato3','Dato4','Dato5','Correcta','Extra']);
  SpreadsheetApp.getUi().alert('Listo. Se creó la hoja "BancoJuegos".\n\nAquí se guardan las preguntas fijas de cada examen (sin IA). Claude te entregará el contenido listo para pegar.');
}

// La interfaz llama a esto: devuelve un set de juego mezclado al azar para un examen.
// No gasta IA: solo lee y baraja el banco guardado.
function obtenerRepasoBanco(asig, contenido, tipo){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const h=ss.getSheetByName('BancoJuegos');
  if(!h || h.getLastRow()<2) return {ok:false, msg:'Aún no hay banco de juegos cargado.'};

  const datos=h.getRange(2,1,h.getLastRow()-1,10).getValues();
  const cLow=String(contenido||'').toLowerCase().trim();

  // filtra por asignatura + contenido + tipo
  const filas=datos.filter(r=>
    r[0]===asig &&
    String(r[1]).toLowerCase().trim()===cLow &&
    String(r[2]).toLowerCase().trim()===tipo
  );
  if(!filas.length) return {ok:false, msg:'Todavía no hay actividades de "'+tipo+'" para esta prueba. Pronto las agregamos.'};

  // baraja y toma un subconjunto para que cada partida sea distinta
  const CANT={quiz:8, flash:10, vf:8, completar:6, adivinanza:5}[tipo] || 8;
  const mezcladas=filas.sort(()=>Math.random()-0.5).slice(0, CANT);

  const items=mezcladas.map(r=>armarItem_(tipo, r));
  return {ok:true, tipo:tipo, total:filas.length, items:items};
}

function armarItem_(tipo, r){
  // r = [Asig,Cont,Tipo,Dato1,Dato2,Dato3,Dato4,Dato5,Correcta,Extra]
  if(tipo==='quiz'){
    return { pregunta:r[3], opciones:[r[4],r[5],r[6],r[7]].filter(x=>x!==''), correcta:parseInt(r[8])||0 };
  }
  if(tipo==='flash'){
    return { frente:r[3], reverso:r[4] };
  }
  if(tipo==='vf'){
    return { afirmacion:r[3], esVerdadero:String(r[8]).toUpperCase().indexOf('V')===0, porque:r[9]||'' };
  }
  if(tipo==='completar'){
    return { frase:r[3], respuesta:r[8] };
  }
  if(tipo==='adivinanza'){
    return { pista:r[3], respuesta:r[8] };
  }
  return {};
}

// Cuántas actividades hay cargadas por examen (para saber qué falta llenar)
function resumenBanco(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const h=ss.getSheetByName('BancoJuegos');
  if(!h || h.getLastRow()<2){ SpreadsheetApp.getUi().alert('El banco está vacío.'); return; }
  const datos=h.getRange(2,1,h.getLastRow()-1,3).getValues();
  const cont={};
  datos.forEach(r=>{
    const k=r[0]+' › '+r[1];
    cont[k]=cont[k]||{}; cont[k][r[2]]=(cont[k][r[2]]||0)+1;
  });
  let txt='📊 Contenido del banco:\n\n';
  Object.keys(cont).sort().forEach(k=>{
    txt+='• '+k+'\n   '+Object.entries(cont[k]).map(([t,n])=>t+':'+n).join('  ')+'\n';
  });
  SpreadsheetApp.getUi().alert(txt.slice(0,1400));
}
