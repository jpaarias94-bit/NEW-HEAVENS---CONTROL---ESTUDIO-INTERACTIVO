/*************************************************************
 *  CALENDARIO BASE - Plan oficial 4° Básico B (II Semestre 2026)
 *  Fuente: Calendario de evaluaciones New Heaven High School
 *  Estas fechas son el PLAN. Los correos las confirman o mueven.
 *************************************************************/

function cargarCalendarioBase_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.getSheetByName('Examenes') || crearHoja_(ss,'Examenes',
    ['Asignatura','Contenido','Tipo','Instrumento','Porcentaje',
     'Fecha plan inicio','Fecha plan término','Fecha confirmada','¿Movida?','Estado']);

  // Limpia solo datos (conserva encabezado)
  if (h.getLastRow() > 1) h.getRange(2,1,h.getLastRow()-1,10).clearContent();

  // [Asignatura, Contenido, Tipo, Instrumento, %, inicio(mes,día), término(mes,día)]
  const plan = [
    ['Lenguaje y Comunicación','Textos instructivos y cómic','Sumativa n°1','Prueba escrita',15,[7,27],[7,27]],
    ['Matemática','Operaciones básicas','Sumativa n°1','Prueba escrita',15,[7,28],[7,28]],
    ['Ciencias Naturales','Placas tectónicas, sismos y volcanes','Sumativa n°1','Prueba escrita',20,[7,31],[7,31]],
    ['Historia, Geografía y Ciencias Sociales','Civilización inca','Sumativa n°1','Prueba escrita',20,[8,6],[8,6]],
    ['Inglés','Vocabulario de planetas + adjetivos descriptivos','Sumativa n°1','Lista de cotejo',25,[7,31],[8,7]],
    ['Religión','Dios nos Perdona','Sumativa n°1','Prueba escrita',25,[8,3],[8,7]],
    ['Educación física','Danza folclórica: La Cueca','Sumativa n°1','Rúbrica',15,[8,10],[8,10]],
    ['Lenguaje y Comunicación','Biografía, infografía y caligrama','Sumativa n°2','Prueba escrita',15,[8,19],[8,19]],
    ['Matemática','Fracciones y decimales','Sumativa n°2','Prueba escrita',15,[8,27],[8,27]],
    ['Inglés','Vocabulario minibeast + have/don\'t have + can/can\'t + verbos de acción','Sumativa n°2','Prueba escrita',25,[9,4],[9,4]],
    ['Lengua y cultura de los pueblos originarios','Atacameños','Sumativa n°1','Prueba escrita',25,[9,4],[9,4]],
    ['Educación física','Danza folclórica: La Guaracha','Sumativa n°2','Rúbrica',15,[9,8],[9,8]],
    ['Lenguaje y Comunicación','Género lírico (figuras literarias, poemas, caligramas)','Sumativa n°3','Prueba escrita',15,[9,9],[9,9]],
    ['Ciencias Naturales','Propiedades de la materia, masa, volumen e instrumentos de medición','Sumativa n°2','Prueba escrita',30,[9,28],[9,28]],
    ['Artes visuales','Simetría','Sumativa n°1','Pauta de evaluación',30,[9,28],[9,28]],
    ['Música','Canción en metalófono','Sumativa n°1','Rúbrica',30,[10,5],[10,5]],
    ['Tecnología','Área y perímetro','Sumativa n°1','Pauta de evaluación',30,[10,7],[10,7]],
    ['Matemática','Geometría (área, perímetro y volumen)','Sumativa n°3','Prueba escrita',15,[10,8],[10,8]],
    ['Religión','La Biblia','Sumativa n°2','Prueba escrita',25,[10,5],[10,9]],
    ['Educación física','Juegos pre-deportivos 1','Sumativa n°3','Rúbrica',15,[10,20],[10,20]],
    ['Lenguaje y Comunicación','Género dramático','Sumativa n°4','Prueba escrita',15,[10,21],[10,21]],
    ['Historia, Geografía y Ciencias Sociales','Democracia y derechos','Sumativa n°2','Prueba escrita',30,[10,22],[10,22]],
    ['Matemática','Simetría y transformaciones isométricas','Sumativa n°4','Prueba escrita',15,[10,29],[10,29]],
    ['Lengua y cultura de los pueblos originarios','Mapuches','Sumativa n°2','Prueba escrita',25,[10,30],[10,30]],
    ['Ciencias Naturales','Estados de la materia y características de la fuerza','Sumativa n°3','Prueba escrita',30,[11,16],[11,16]],
    ['Música','Canción en metalófono (navideña)','Sumativa n°2','Rúbrica',40,[11,16],[11,16]],
    ['Lenguaje y Comunicación','Género narrativo','Sumativa n°5','Prueba escrita',10,[11,18],[11,18]],
    ['Religión','Las promesas de Dios a Josué. La obediencia y la fidelidad a Dios','Sumativa n°3','Prueba escrita',25,[11,16],[11,20]],
    ['Artes visuales','Maqueta rapa nui','Sumativa n°2','Rúbrica',40,[11,23],[11,23]],
    ['Matemática','Ecuaciones e inecuaciones','Sumativa n°5','Prueba escrita',10,[11,24],[11,24]],
    ['Educación física','Juegos pre-deportivos 2','Sumativa n°4','Rúbrica',25,[11,24],[11,24]],
    ['Tecnología','Maqueta rapa nui','Sumativa n°2','Rúbrica',40,[11,25],[11,25]],
    ['Historia, Geografía y Ciencias Sociales','Organización política en Chile','Sumativa n°3','Prueba escrita',30,[11,26],[11,26]],
    ['Inglés','Vocabulario personajes de fantasía + adjetivos + acciones','Sumativa n°3','Lista de cotejo',30,[11,20],[11,27]],
    ['Lengua y cultura de los pueblos originarios','Maqueta rapa nui','Sumativa n°3','Rúbrica',25,[11,27],[11,27]]
  ];

  const filas = plan.map(p => [
    p[0], p[1], p[2], p[3], p[4],
    new Date(ANIO, p[5][0]-1, p[5][1]),
    new Date(ANIO, p[6][0]-1, p[6][1]),
    '', '', 'Programado'
  ]);

  h.getRange(2,1,filas.length,10).setValues(filas);
  h.getRange(2,6,filas.length,3).setNumberFormat('dd/mm/yyyy');
}
