# Inventario Legislativo

PWA local para control de inventario, resguardos, evidencias y conciliación de listados de la Cámara de Diputados.

## Cómo abrirla

En Windows, haz doble clic en `Iniciar_Inventario.cmd`. Abrirá el navegador en una dirección local (`http://127.0.0.1:4173/` o el siguiente puerto libre); deja abierta la ventana negra mientras trabajas y usa `Ctrl+C` para cerrarla. No requiere Node.js ni una conexión a Internet.

También puedes servir esta carpeta desde cualquier servidor estático. No abras `index.html` con `file://`, porque la cámara, el lector QR y el modo offline requieren un origen web. Por ejemplo, desde una terminal con Node.js:

```powershell
npx serve .
```

Después abre la URL local mostrada por el servidor e ingresa un nombre para la sesión local. La versión pública no contiene directorios de personal ni autenticación institucional.

## Capacidades incluidas

- Importación local de exportaciones Oracle HTML (`.xls`), libros OOXML (`.xlsx`/`.xlsm`) y CSV; detecta área, tipo de libro y claves decimales. Corrige de forma segura los artefactos binarios de Excel en claves de 5–6 decimales para evitar falsos duplicados.
- Tabla de inventario de alta densidad con búsqueda instantánea, filtros, selección masiva, asignación de resguardante/ubicación, re-etiquetado, notas, evidencia fotográfica y deshacer en la sesión.
- Gestión de resguardantes y nomenclatura de ubicaciones consecutivas globales.
- Croquis local por resguardante: pines reubicables para cada ubicación, resumen de bienes asignados y carga opcional de un plano de referencia PNG/JPEG/WebP. Las posiciones viajan en el respaldo; la imagen permanece en el almacenamiento local por privacidad.
- Altas de bienes adicionales, claves `CD-ÁREA-NNN`, detección de series duplicadas y perfiles de autollenado por expresión regular.
- Reportes imprimibles, conciliación con vista previa de altas/bajas/modificados, exportación Excel compatible y CSV de re-etiquetado.
- Respaldos ZIP, restauración controlada, fusión de sesiones y datos/fotos persistidos localmente en IndexedDB.
- Escáner QR con la API nativa del navegador cuando está disponible; solicita permiso de cámara solo al activarlo.

## Nota de producción

La app es deliberadamente local-first: no transmite inventario ni fotografías a un servidor. Para operación multiusuario real habría que añadir autenticación institucional, API con control de roles, sincronización/concurrencia y una política de auditoría centralizada.

## Publicación en GitHub Pages

El repositorio incluye un flujo de GitHub Actions para publicar la aplicación en GitHub Pages al hacer `push` a `main`. La publicación es estática: la sesión, inventario, fotos y respaldos que cree cada persona permanecen únicamente en su navegador.
