# Reglas del proyecto

## Reglas generales

- NUNCA implementes cambios, ediciones o acciones sobre el código sin que el
  usuario lo ordene explícitamente. Solo ejecuta lo que se te pida de forma
  concreta; la exploración y verificación no cuentan como implementación.

## Apps Script (clasp)

- Crear una nueva versión (`clasp version "descripción breve del cambio"`) **solo
  cuando el usuario lo ordene explícitamente**, nunca de forma automática tras un
  `clasp push`.
- Desplegar la versión en la aplicación web para que los cambios queden
  visibles en la URL `/exec` **solo cuando el usuario lo ordene**.
- El despliegue `@HEAD` es de solo lectura y no se puede modificar.
