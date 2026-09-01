# Proyecto: Control Académico — Emmanuel (4° Básico B)

## Qué es
Sistema en Google Apps Script que lee Google Classroom y muestra exámenes y avisos
en una app web. Backend en Apps Script + Google Sheets como base de datos.

## Archivos del proyecto (pegar en la sesión)
- Codigo.gs — backend principal, sincronización Classroom, lectura de hojas
- GeminiIA.gs — integración con Gemini para analizar anuncios
- CalendarioBase.gs — plan oficial: 35 evaluaciones del II Semestre 2026
- MaterialPDF.gs — procesa PDFs de temarios
- Agente.gs — agente que lee correos y detecta cambios de fecha de examen
- Index.html — interfaz web
- (extra) CarpetasExamen.gs — carpetas por examen en Drive
- (extra) BancoJuegos.gs — banco de estudio fijo, sin IA

## Problema a resolver (lo central)
El agente debe mover automáticamente la fecha de un examen cuando un correo lo
anuncia, y NO lo está haciendo.
Caso real: Ciencias Naturales pasa del 31/7 al 14/8 y el agente no lo aplica.

## Lo que quiero de Cowork
1. Reproducir la lógica del agente (analizarCorreoAgente_ + aplicarCambioSeguro_ +
   candidatosAFecha_) en un entorno ejecutable, ya que en Apps Script no se puede
   depurar paso a paso.
2. Pasarle los correos de prueba de abajo y encontrar el punto EXACTO donde falla
   el caso de Ciencias.
3. Correr la batería completa y confirmar que la corrección no rompe los otros casos.
4. Entregar el código corregido listo para pegar en Apps Script.

## Reglas de seguridad que NO se pueden violar (errores del pasado)
- El agente solo puede mover exámenes de LA MISMA ASIGNATURA del correo.
- Debe usar la fecha DEL CORREO como referencia temporal, no la fecha de hoy.
- Si hay ambigüedad, NO adivina: deja propuesta para que el usuario apruebe.
- Debe existir siempre restaurarPlanOficial como red de seguridad.

## Correos de prueba

### Caso 1 — Reagenda CLARA (debe aplicarse solo)
Asignatura: Ciencias Naturales
Fecha del correo: 2026-07-24
Texto:
"Estimados apoderados, junto con saludar les informo que la evaluación de
Ciencias Naturales sobre placas tectónicas, sismos y volcanes, programada para
el 31 de julio, se reagendará para el 14 de agosto. Saludos cordiales."
Resultado esperado: el examen "Placas tectónicas, sismos y volcanes" queda en
2026-08-14, marcado como movido, con respaldo del correo.

### Caso 2 — Reagenda AMBIGUA (debe quedar como propuesta, no aplicar solo)
Asignatura: Matemática
Fecha del correo: 2026-08-20
Texto:
"Estimados, la prueba se moverá para la próxima semana. Atenta a la fecha exacta."
Resultado esperado: propuesta para que el usuario elija (no hay fecha clara ni
se sabe cuál de las pruebas de Matemática).

### Caso 3 — Suspensión (sin nueva fecha)
Asignatura: Educación física
Fecha del correo: 2026-08-08
Texto:
"Estimados, la evaluación de danza folclórica La Cueca no se realizará esta
semana por actividad del colegio. Se avisará nueva fecha más adelante."
Resultado esperado: examen marcado como suspendido, sin fecha nueva.

### Caso 4 — Solo comparte material (NO debe mover nada)
Asignatura: Ciencias Naturales
Fecha del correo: 2026-07-28
Texto:
"Estimados, les comparto la guía de estudio para la prueba de placas tectónicas.
Favor traer su libro de ciencias la próxima clase."
Resultado esperado: registra material, NO cambia ninguna fecha.

## Nota
Cowork prueba la lógica en un entorno simulado. El código final se pega y se
ejecuta en el Apps Script de la cuenta de Emmanuel; Cowork no se conecta a
Classroom real.
