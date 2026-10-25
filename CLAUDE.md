# Control Académico — Emmanuel (4° Básico B)

## Cierre de sesión pedido por Gestión de proyectos (autorizado por Juan, 25-10-2026)
- Juan autoriza que su chat «Gestión de proyectos» (y su revisión nocturna) pida cerrar esta sesión cuando el contexto se llena.
- Al recibir ese pedido, ejecuta solo el cierre con el protocolo propio de este proyecto (o, si no hay, `RESUMEN SECCIÓN.md` + `HISTORICO CLAUDE.md`), sin volver a preguntarle. Nada más que el cierre.

## ⚠️ AL INICIAR CADA SESIÓN, LEE PRIMERO
`bitacora/ULTIMA-SESION.md` — resumen de lo ÚLTIMO trabajado en la sesión anterior.
Con eso retomas el hilo. Si necesitas el detalle completo de TODAS las sesiones,
está en `bitacora/HISTORICO.md`.

## Preferencias del usuario (Juan)
- Responde en **español**.
- **NO NARRES** pasos internos ni herramientas; entrega solo el resultado final y las
  acciones concretas que Juan debe ejecutar.
- Antes de cambios grandes, da opciones y valida en cada paso.

## Dos frentes de trabajo
1. **Estudio de Emmanuel** — artefactos interactivos de estudio + lectura de correos del
   colegio (conector Gmail de `emmanuel.arias@newheavenhs.cl`) + Carta Gantt de pruebas.
   Estado al día en `bitacora/ULTIMA-SESION.md`.
2. **App Apps Script** (este repo) — agente de fechas (P1), documentos (P2), juegos (P3).
   Estado técnico en `ESTADO_PROYECTO.md`; arranque y casos de prueba en `ARRANQUE_COWORK.md`.

## Bitácora de sesiones (cómo funciona)
- `bitacora/ULTIMA-SESION.md` — se **sobrescribe** al cerrar cada sesión; es lo que lee la siguiente.
- `bitacora/HISTORICO.md` — se **acumula**; registro de todas las sesiones, nunca se borra.
- `bitacora/sesiones/sesion-NN.md` — copia **congelada** de cada sesión.

Al **cerrar** una sesión N: 1) guardar `bitacora/sesiones/sesion-NN.md`, 2) agregar su
resumen al final de `bitacora/HISTORICO.md`, 3) sobrescribir `bitacora/ULTIMA-SESION.md`
con el resumen de la sesión N.

## Rama de trabajo
`claude/gas-project-review-o399mb`
