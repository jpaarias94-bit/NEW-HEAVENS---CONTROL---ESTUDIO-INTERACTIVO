/*************************************************************
 * CARPETAS POR EXAMEN — crea en Drive una subcarpeta por cada
 * examen del calendario, dentro de la carpeta de su asignatura.
 *
 * Requiere la hoja "Carpetas" con: Asignatura (col A) | ID carpeta (col B).
 * Guarda los IDs de las subcarpetas creadas en una hoja nueva
 * "CarpetasExamen": Asignatura | Contenido | ID subcarpeta | Link.
 *
 * Se ejecuta desde la cuenta de Emmanuel (dueña de las carpetas).
 *************************************************************/

function crearCarpetasPorExamen(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();

  // 1) Mapa asignatura -> ID carpeta madre (desde la hoja Carpetas)
  const hCar=ss.getSheetByName('Carpetas');
  if(!hCar || hCar.getLastRow()<2){
    SpreadsheetApp.getUi().alert('No encontré la hoja "Carpetas" con los IDs de asignatura.');
    return;
  }
  const carpetas={};
  hCar.getRange(2,1,hCar.getLastRow()-1,2).getValues().forEach(r=>{
    if(r[0] && r[1]) carpetas[String(r[0]).trim()]=String(r[1]).trim();
  });

  // 2) Lista de exámenes (asignatura + contenido) desde la hoja Examenes
  const hEx=ss.getSheetByName('Examenes');
  if(!hEx || hEx.getLastRow()<2){
    SpreadsheetApp.getUi().alert('No hay exámenes en la hoja "Examenes".');
    return;
  }
  const examenes=hEx.getRange(2,1,hEx.getLastRow()-1,2).getValues(); // [Asignatura, Contenido]

  // 3) Hoja destino donde guardamos los IDs de subcarpeta
  const hDest=crearHoja_(ss,'CarpetasExamen',['Asignatura','Contenido','ID subcarpeta','Link']);
  const yaCreadas=new Set(
    hDest.getLastRow()>1
      ? hDest.getRange(2,1,hDest.getLastRow()-1,2).getValues().map(r=>r[0]+'||'+r[1])
      : []
  );

  let creadas=0, saltadas=0, sinMadre=0;

  examenes.forEach(([asig, contenido])=>{
    if(!asig || !contenido) return;
    const clave=asig+'||'+contenido;
    if(yaCreadas.has(clave)){ saltadas++; return; }

    const idMadre=carpetas[String(asig).trim()];
    if(!idMadre){ sinMadre++; return; }

    try{
      const madre=DriveApp.getFolderById(idMadre);
      // nombre de subcarpeta = contenido del examen (recortado y limpio)
      const nombre=String(contenido).replace(/[\\/:*?"<>|]/g,' ').trim().slice(0,80);

      // evita duplicar si ya existe una carpeta con ese nombre
      let sub=null;
      const it=madre.getFoldersByName(nombre);
      sub = it.hasNext() ? it.next() : madre.createFolder(nombre);

      hDest.appendRow([asig, contenido, sub.getId(), sub.getUrl()]);
      yaCreadas.add(clave);
      creadas++;
    }catch(e){
      // si el ID madre falla (permiso/cuenta), lo cuenta como sin madre
      sinMadre++;
    }
  });

  let msg='Carpetas por examen:\n\n'+
    '✅ Creadas: '+creadas+'\n'+
    '↔️ Ya existían (saltadas): '+saltadas+'\n';
  if(sinMadre>0) msg+='⚠️ Sin carpeta de asignatura o sin acceso: '+sinMadre+'\n'+
    '(revisa que esas asignaturas tengan su ID en la hoja "Carpetas")';
  SpreadsheetApp.getUi().alert(msg);
}

// Devuelve el ID de la subcarpeta de un examen (para otros módulos, ej. leer material)
function idSubcarpetaExamen_(ss, asig, contenido){
  const h=ss.getSheetByName('CarpetasExamen');
  if(!h || h.getLastRow()<2) return null;
  const d=h.getRange(2,1,h.getLastRow()-1,3).getValues();
  const cLow=String(contenido).toLowerCase().trim();
  for(const r of d){
    if(r[0]===asig && String(r[1]).toLowerCase().trim()===cLow) return r[2];
  }
  return null;
}
