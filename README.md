# Mi día

App web para organizar tareas de casa e instituto, con temporizador de concentración, estrellas y premios.

- `index.html`, `sw.js`, `manifest.webmanifest` e iconos: la app (se publica con GitHub Pages).
- `google-apps-script/`: el script que se pega en script.google.com, en la cuenta de Google del equipo.

## Privacidad

Este repositorio es público. Aquí no puede aparecer ningún nombre ni dato personal.
El ID del calendario va solo en la copia del script dentro de la cuenta de Google del equipo.
El token va solo en el enlace que se pega en la app.
El nombre que se escribe en Ajustes se guarda solo en el móvil y no se envía nunca.

## Qué hace el script de Google

1. Da a la app las tareas del día del calendario del equipo.
2. Da a la app los premios activos de la hoja "Mi día · registro y premios", pestaña Premios.
3. Guarda en la pestaña Registro lo que se hace en la app, y las fotos en una carpeta de Drive.
4. Manda de lunes a viernes, sobre las 8:00, un correo con el resumen del día anterior. El lunes resume viernes, sábado y domingo.

## Premios

En la pestaña Premios de la hoja: nombre, estrellas, Activo (Sí/No), Desde y Hasta (opcional) y Tipo (Fijo/Temporal).
Solo salen en la app los activos y dentro de sus fechas. La app avisa cuando aparece uno nuevo.

## Antes de pasar del móvil de prueba al definitivo

- En el móvil de prueba: Ajustes > Zona del equipo > "Borrar todo y empezar de cero" (o simplemente no volver a usarlo).
- En la hoja, borrar las filas de prueba de Registro y Canjes, y las fotos de prueba de la carpeta.
- En el móvil definitivo: instalar la app, pegar el enlace en Zona del equipo y cambiar el código 1234.
