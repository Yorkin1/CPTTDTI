# 📇 Gestor de Clientes y Citas — Plantilla de Google Sheets + Apps Script

Plantilla **gratuita** y **genérica** para administrar **clientes**, **citas** (con eventos en **Google Calendar**), **historial** y **reservas en línea**. Se distribuye como una **hoja de cálculo de Google con un Apps Script vinculado**, lista para **copiar y personalizar sin escribir código**.

- 📊 **Google Sheets** como base de datos (hoja vinculada, no se crea otra).
- 📅 **Google Calendar** para generar los eventos de las citas.
- 🌐 **Aplicación web** moderna, responsiva y **100 % en español**.
- 🎨 **Asistente de personalización** en el primer uso (nombre del negocio, etiquetas, color, tema, etc.).
- 🗓️ **Reservas en línea**: sus clientes reservan su propio horario desde una página pública, sin iniciar sesión.
- 🔐 **Usuarios con roles** (admin/editor), registro abierto y contraseñas protegidas.

> Cada persona que copia la plantilla obtiene **su propia base de datos**: los datos de una copia **nunca**
> se mezclan con la plantilla maestra ni con otras copias. Además, la **primera ejecución** de
> `Client Manager → Iniciar configuración` **vacía los datos de ejemplo** que la copia hereda de la maestra,
> de modo que **cada copia comienza limpia**.

---

## 🚀 Flujo para el usuario de la plantilla (recomendado)

Siga estos pasos **en orden**:

1. **Hacer una copia** de la plantilla maestra de Google Sheets
   (menú `Archivo → Hacer una copia`).
2. **Abrir la hoja copiada** (su propia copia, en su Google Drive).
3. Hacer clic en **`Client Manager` → `Iniciar configuración`** (menú superior de la hoja).
   **Este paso es obligatorio**: vincula la copia a su propia base de datos y,
   en la primera ejecución, elimina los datos que la copia pudo heredar de la
   plantilla maestra (clientes, citas, historial, usuarios y configuración de
   ejemplo) para que cada copia nazca vacía.
4. **Aprobar los permisos** solicitados de Google Sheets y Google Calendar.
5. **Desplegar el Apps Script vinculado como aplicación web**, ejecutando **como el usuario**
   (`Implementar → Nueva implementación → Aplicación web`, *Ejecutar como: Yo*).
6. **Abrir la URL `/exec`** generada por la implementación.
7. **Completar el asistente** de personalización inicial (nombre del negocio, etiquetas, color, duración de citas…).
8. **Registrar el primer usuario** (el primero en registrarse se convierte en **administrador**) e iniciar sesión.

Al completar el asistente y entrar con un usuario se carga el **dashboard** ya personalizado.

> 💡 El menú **`Client Manager`** aparece automáticamente al abrir la hoja copiada (función `onOpen`).
> Incluye: **Iniciar configuración**, **Verificar base de datos**, **Limpiar datos**, **Proteger/Desproteger base de datos** y **Abrir dashboard**.

---

## 🇬🇧 English summary

This is a **copyable Google Sheets template with a bound Apps Script**. The end-user flow is:

1. **Make a copy** of the master Google Sheets template (`File → Make a copy`).
2. **Open the copied sheet**.
3. Click **`Client Manager` → `Iniciar configuración`** (custom menu).
   **This step is required**: it binds the copy to its own database and, on
   first run, wipes any data inherited from the master template so every copy
   starts clean.
4. **Approve** the requested Google Sheets & Calendar permissions.
5. **Deploy the bound Apps Script as a Web App**, *executing as the user*
   (`Deploy → New deployment → Web app`, *Execute as: Me*).
6. **Open the generated `/exec` URL**.
7. **Complete the first-run setup wizard** (business name, appointment labels, primary color, default duration).
8. **Register the first user** (becomes the **administrator**) and log in.

It also includes **online reservations**: a public page (`…/exec?v=reservar`) where customers book their own slot. Every copy stores **its own spreadsheet ID** (`ID_HOJA_CALCULO` in `PropertiesService`), so data written from one copy never reaches the master template. All UI text is in Spanish.

---

## ✅ Requisitos previos

- Una **cuenta de Google** (personal o de trabajo). El plan gratuito es suficiente.
- Un navegador web moderno (Chrome, Edge, Firefox, etc.).
- Aproximadamente **10 minutos** para la configuración inicial.

---

## 🧭 Detalle de los pasos

### 1) Copiar la plantilla maestra
Abra la plantilla maestra y use `Archivo → Hacer una copia`. Trabajará siempre sobre **su copia**.

### 2) Abrir la hoja copiada
Ábrala desde su Google Drive. Verá un menú nuevo llamado **`Client Manager`** en la barra superior.
### 3) `Client Manager → Iniciar configuración`

**Paso obligatorio.** Esta acción ejecuta `configurarPlantilla()`, que:
- Toma la **hoja de cálculo actual** (`SpreadsheetApp.getActiveSpreadsheet()`).
- Guarda **su propio ID** en `PropertiesService` bajo la clave **`ID_HOJA_CALCULO`**.
- En la **primera configuración de cada copia** (cuando ese ID aún está vacío),
  **vacía los datos copiados de la maestra** (Clientes, Citas, Historial y
  Usuarios quedan solo con encabezados) y **restablece `Configuracion`** a los
  valores predeterminados (`CONFIGURADA = NO`). Así cada copia comienza limpia.
- Usa **`LockService`** para evitar ejecuciones simultáneas.
- Verifica/crea las hojas **Clientes, Citas, Historial** y **Configuracion**.
- **Repara los encabezados sin borrar, mover ni sobrescribir** los registros existentes.
- Crea los **valores de configuración predeterminados** que falten.

> ⚠️ Si abre la URL `/exec` **antes** de ejecutar este paso, la aplicación
> mostrará una pantalla indicando que la copia **no está configurada** y cómo
> hacerlo. No podrá usarse hasta completar el paso 3.

### 4) Aprobar permisos
Al ejecutar por primera vez, Google pedirá autorización para **Sheets** y **Calendar**.
Si aparece “Google no ha verificado esta aplicación”, es normal (la app es suya):
`Configuración avanzada → Ir a … (no seguro) → Permitir`.

### 5) Desplegar como aplicación web
`Implementar → Nueva implementación → Aplicación web`:
- **Ejecutar como:** **`Yo`** (importante: así los eventos de Calendar y los datos quedan en **su** cuenta).
- **Quién tiene acceso:** `Cualquier usuario` o `Solo yo`, según prefiera.
- Copie la **URL `/exec`** que se genera.

### 6) Abrir la URL `/exec`
Ábrala en el navegador. La primera vez se mostrará el **asistente de personalización**.

### 7) Completar el asistente
Ingrese: **nombre del negocio**, **mensaje de bienvenida**,
**etiqueta de cita**, **color primario** y **duración predeterminada de cita**.
Al guardar, `CONFIGURADA` pasa a **`SI`**.

### 8) Registrar el primer usuario e iniciar sesión
La pantalla de login ofrece **Registrarse**. El **primer usuario** en registrarse se convierte en
**administrador** (único que ve las pestañas *Configuración* y *Usuarios*). Con la sesión iniciada
se carga el **dashboard** personalizado. Las contraseñas se guardan cifradas en la hoja `Usuarios`
y las sesiones usan tokens que vencen a las **12 horas**.

> 🔑 ¿Olvidó su contraseña? Use *"¿Olvidó su contraseña?"* en el login para recibir una **temporal** por correo.

> 🔄 **Reejecutar la configuración** (menú → *Iniciar configuración*) es **seguro**: repara la estructura
> **sin eliminar datos** existentes.

---

## 🧪 Lista de verificación de pruebas (checklist)

Use esta lista para comprobar que la plantilla funciona correctamente como plantilla copiable:

- [ ] **ID propio por copia:** en una copia recién hecha, ejecutar *Iniciar configuración* y confirmar que
      `ID_HOJA_CALCULO` (en *Configuración del proyecto → Propiedades del script*, o vía log) es el **ID de la copia**,
      **no** el de la plantilla maestra.
- [ ] **Copia limpia:** tras copiar la plantilla y ejecutar *Iniciar configuración*, las hojas **Clientes, Citas,
      Historial y Usuarios** quedan solo con encabezados y `CONFIGURADA = NO` (sin datos heredados de la maestra).
- [ ] **Aislamiento de datos:** agregar un cliente/cita en la copia y verificar que en la
      **plantilla maestra no aparece ningún dato nuevo** (las hojas de la maestra quedan intactas).
- [ ] **Calendar del desplegador:** agendar una cita desde la app web y confirmar que el evento se crea
      en el **Google Calendar de la cuenta que desplegó** la aplicación web (Ejecutar como: Yo).
- [ ] **Reejecución no destructiva:** con datos ya cargados, volver a ejecutar *Iniciar configuración* y
      confirmar que **no se eliminan ni alteran** los registros existentes (solo se reparan encabezados/estructura si hace falta).
- [ ] **Asistente solo la primera vez:** con `CONFIGURADA = SI`, al abrir la URL `/exec` se muestra
      directamente el dashboard (no el asistente).
- [ ] **Personalización aplicada:** el nombre del negocio, la etiqueta de cita, el color primario y
      la duración predeterminada se reflejan en toda la interfaz.
- [ ] **`Verificar base de datos`** (menú) informa que todas las hojas y encabezados están correctos.
- [ ] **Roles/Admin:** el primer usuario registrado aparece como **admin** y ve las pestañas "Configuración" y "Usuarios"; los editores no.
- [ ] **Sesión:** al expirar el token (12 h) o cerrar la pestaña, la app vuelve al login.
- [ ] **Recuperación:** "¿Olvidó su contraseña?" envía una contraseña temporal por correo.
- [ ] **Reservas en línea:** con las reservas habilitadas, desde una ventana anónima se abre `…/exec?v=reservar`, se elige horario y se confirma; la cita aparece en el dashboard y en Calendar.
- [ ] **Reserva de un cliente ya registrado:** si el correo/teléfono coincide con una ficha, la página avisa y la reserva queda vinculada a ese cliente.
- [ ] **Días cerrados:** un día sin horas en el horario de atención no muestra horarios disponibles.

---

## 🛠️ Solución de problemas

| Problema | Solución |
|----------|----------|
| **No veo el menú `Client Manager`** | Cierre y vuelva a abrir la hoja copiada; el menú se crea en `onOpen`. Si no aparece, ejecute `onOpen` manualmente desde el editor una vez. |
| **“Se requiere autorización” y no avanza** | Ejecute *Iniciar configuración*, elija su cuenta y `Configuración avanzada → Ir a … (no seguro) → Permitir`. |
| **La app carga en blanco** | Verifique que todos los archivos HTML tengan el **nombre exacto** (sin `.html`) y que `Code.gs` esté completo. |
| **“No hay hoja de cálculo configurada”** | Ejecute **`Client Manager → Iniciar configuración`** dentro de su copia. |
| **La app muestra “Falta configurar esta copia”** | Es lo esperado antes del paso 3: ejecute **`Client Manager → Iniciar configuración`** en su copia y pulse **Reintentar**. |
| **No se crea el evento en Calendar** | Asegúrese de haber desplegado la app **Ejecutar como: Yo** y de haber aprobado el permiso de Calendar. |
| **El botón “Abrir dashboard” dice que no hay URL** | Debe **desplegar** la app web primero (Paso 5). Luego reintente. |
| **La página de reservas dice “Sistema sin configurar” o “Reservas cerradas”** | Ejecute *Iniciar configuración* y despliegue la app web. Si ya está desplegada, verifique que **Habilitar reservas** esté activo en Clientes → Reservas en línea. |
| **No hay horarios disponibles para un día** | Ese día está **cerrado** en el horario de atención, o todos los horarios ya fueron reservados. |
| **Se me olvidó la contraseña** | Use *“¿Olvidó su contraseña?”* en el login → recibirá una **contraseña temporal** por correo. |
| **No puedo ver Configuración/Usuarios** | Solo el **administrador** (primer usuario registrado) ve esas pestañas. |
| **Cambié el código pero la app no se actualiza** | Cree una **nueva versión** de la implementación (`Gestionar implementaciones → ✏️ → Nueva versión`). |

---

## 📂 Estructura del proyecto

```
client-manager-gas/
├── appsscript.json      # Manifiesto: scopes y config. de la app web (executeAs=USER_DEPLOYING)
├── Code.gs              # Backend: onOpen, doGet, configurarPlantilla, API config, CRUD, Calendar, usuarios, reservas públicas
├── Index.html           # Estructura principal: asistente de configuración + login + dashboard (SPA)
├── HojaEstilos.html     # Estilos CSS (colores configurables, tema claro/oscuro, responsivo)
├── JavaScript.html      # Lógica del cliente (asistente, personalización dinámica, formularios, tablas, sesión)
├── Clientes.html        # Vista: alta/listado de clientes + ajustes de reservas en línea
├── Citas.html           # Vista: agendar citas (crea evento en Calendar)
├── Historial.html       # Vista: historial por cliente
├── Configuracion.html   # Vista (solo admin): claves de configuración general
├── Usuarios.html        # Vista (solo admin): gestión de usuarios y roles
├── Reserva.html         # Página pública de reservas en línea (4 pasos, sin login)
└── README.md            # Esta guía
```

### Hojas de la base de datos

| Hoja | Columnas |
|------|----------|
| **Clientes** | ID_Cliente, Nombre, Apellido, Telefono, Email, Direccion, Notas, Fecha_Registro, Foto |
| **Citas** | ID_Cita, ID_Cliente, Titulo, Fecha, Hora, Duracion_Mins, Descripcion, ID_Evento_Calendar, Estado |
| **Historial** | ID_Registro, ID_Cliente, ID_Cita, Fecha, Descripcion, Resultado |
| **Usuarios** | ID_Usuario, Nombre, Email, Salt, Hash, Rol, Activo, Fecha_Registro |
| **Configuracion** | Clave, Valor |

### Claves de la hoja `Configuracion`

| Clave | Valor por defecto |
|-------|-------------------|
| `CONFIGURADA` | `NO` |
| `NOMBRE_NEGOCIO` | `Mi Negocio` |
| `MENSAJE_BIENVENIDA` | `Bienvenido a nuestro sistema de gestión` |
| `ETIQUETA_CITA` | `Cita` |
| `COLOR_PRIMARIO` | `#4285F4` |
| `COLOR_SECUNDARIO` | `#16a34a` |
| `TEMA` | `claro` |
| `DURACION_CITA_PREDETERMINADA` | `60` |
| `HABILITAR_RESERVAS` | `SI` |
| `HORARIO_ATENCION` | JSON por día (Lun–Vie 09:00–17:00, Sáb 09:00–13:00, Dom cerrado) |
| `PASO_RESERVA_MIN` | `30` |

---

## 🗓️ Reservas en línea

El sistema incluye una **página pública de reservas** (`…/exec?v=reservar`) donde sus clientes eligen su propio horario **sin iniciar sesión**:

1. En la vista **Servicios → Reservas en línea** active el switch **"Habilitar reservas en línea"**.
2. Configure el **horario de atención** por día (los días sin hora quedan cerrados) y el **paso entre horarios** (min).
3. Pulse **"Copiar enlace de reservas"** y compártalo con sus clientes (o **"Abrir página de reservas"** para verla).

Cada reserva registra una **cita** normal (con su evento en Google Calendar). Si el cliente ya está en su base (por correo o teléfono), la reserva se **vincula automáticamente a su ficha**. La disponibilidad se refresca en vivo mientras el cliente elige horario, y la página se muestra con el color y el nombre de su negocio.

---

## 🔐 Permisos (OAuth scopes)

El manifiesto (`appsscript.json`) solicita únicamente lo necesario:

- `https://www.googleapis.com/auth/spreadsheets` — leer/escribir la base de datos (incluye `openById`).
- `https://www.googleapis.com/auth/calendar` — crear los eventos de las citas.
- `https://www.googleapis.com/auth/drive` — guardar las fotos de los clientes (carpeta `ClientManager_Imagenes`).
- `https://www.googleapis.com/auth/script.send_mail` — confirmaciones por correo y recuperación de contraseña.
- `https://www.googleapis.com/auth/script.container.ui` — menú "Client Manager" y diálogos en la hoja.
- `https://www.googleapis.com/auth/userinfo.email` — identificar al desplegador y proteger las hojas.

---

## Actualizaciones vía GitHub Actions

Este repositorio incluye un sistema para enviar actualizaciones a múltiples
copias de clientes de forma automatizada.

### Cómo funciona

1. Editas el código en `src/` (localmente o en GitHub)
2. Haces push a GitHub
3. Disparas el workflow "Actualizar copias de clientes"
4. El sistema envía los cambios a cada cliente activo automáticamente

### Clientes activos

La lista de clientes está en `deploy/clientes.json`. Cada cliente tiene:

- `slug`: nombre corto
- `scriptId`: ID del proyecto Apps Script
- `deploymentId`: ID de la implementación desplegada
- `spreadsheetId`: ID de la hoja de cálculo
- `activo`: true/false

### Agregar un cliente nuevo

1. Crear la copia del Apps Script
2. Compartir con la cuenta de actualizaciones como Editor
3. GitHub → Actions → "Agregar cliente" → llenar slug + URL

### Publicar cambios

1. GitHub → Actions → "Actualizar copias de clientes" → Run workflow
2. Ver resultado en `deploy/REPORT.md`

### Pausar un cliente

Cambiar `"activo": false` en `deploy/clientes.json`

---

¡Listo! Copie la plantilla, ejecute la configuración, despliegue la app web y personalícela. 🚀
