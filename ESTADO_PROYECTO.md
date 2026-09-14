# ESTADO DEL PROYECTO — Control Académico Emmanuel

_Handoff para continuar en una sesión nueva. Rama de trabajo: `claude/gas-project-review-o399mb`._

## Qué es
App Google Apps Script (Sheets + Index.html) que mantiene al día el calendario de
exámenes de 4° básico leyendo Classroom y los correos de profesores, descarga el
material y (a futuro) genera juegos de estudio. Backend en .gs, base en Google Sheets.

## Objetivo (3 pilares)
1. **P1 Fechas** — el correo corrige el calendario solo; lo dudoso se aprueba en 1 clic.
2. **P2 Documentos** — auto-descarga de Classroom a Drive + resumen IA en la base.
3. **P3 Juegos** — preguntas generadas del propio material descargado.

Todo se conecta por `examen_id`. Recomendaciones ya aprobadas por el usuario:
fuente de verdad = el correo; automatización = solo alta confianza + bandeja;
match por número de candidato; datos en Sheets por ahora.

## ✅ P1 — HECHO y validado
- `Agente.gs` reescrito: la IA elige el examen por NÚMERO de candidato (no por texto
  exacto) y decide por confianza. Aplica lo claro, propone lo dudoso con "antes → después".
- Validado 5/5 casos (incluido Ciencias 31/7→14/8 y mate agosto→noviembre).

## ✅ P2 — Backend HECHO, resúmenes completándose
- Nuevo `Documentos.gs`: al sincronizar Classroom copia adjuntos a Drive, la IA los
  resume (PDF, Google Docs/Slides, .docx, imágenes) y los guarda en la hoja `Documentos`
  (Archivo, Tipo, Link, Resumen, Temas, Examen, Estado…).
- Menú: "📎 Descargar documentos de Classroom" y "🧠 Resumir documentos pendientes".
- Trigger de tiempo cada 6 h corre `resumirDocumentosPendientes` (idempotente).
- Modelo Gemini correcto: **gemini-3.6-flash** (2.0 fue retirado; en Config gemini_modelo).
- Estado actual: 67 documentos bajados; ~22 en "Procesado", resto "Descargado" (se van
  completando solos). Enlace/Video no se resumen (normal).

## ⏳ PENDIENTE (por dónde seguir)
1. **UI de P2**: que la pestaña "Documentos descargables" muestre el resumen ("de qué
   trata") + botón de descarga directa de Drive (hoy solo enlaza a Classroom).
   Backend ya expone los datos vía `leerDocumentos_` en `obtenerDatos`.
2. **P3 Juegos**: generar preguntas desde `Documentos` + tema del examen → hoja
   `BancoJuegos`, con aprobación 1ª vez por examen.
3. Menores: destacar "antes → después" en la bandeja del Agente; borrar 1 trigger
   `tareaAutomatica` duplicado; respaldar `Automatizacion.gs` (existe en Apps Script,
   no está en el repo).

## Cómo se prueba / despliega
- Pegar los .gs en el editor Apps Script → Guardar. Para la app web (Index): Implementar
  › Administrar implementaciones › Nueva versión.
- Requisitos activos: Classroom API, Drive API, clave Gemini en Config.

## Notas
- GitHub: rama `claude/gas-project-review-o399mb` (main sigue con lo viejo).
- Este repo (Emmanuel) NO tiene nada del proyecto contable "cierre-cuentas";
  aparecen juntos solo porque ambos están conectados a la sesión.
