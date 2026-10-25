# Grupo Construblock · App

Ventas, entregas, clientes, block producido, inventario, gastos, facturas, costo de producción y cuentas de la ferretería y bloquera.

- `index.html`, `manifest.json`, `sw.js`, `icon.svg`: la app instalable que se publica en GitHub Pages.
- `apps-script/`: el servidor en Google Apps Script, pegado en la hoja **CONSTRUBLOCK 1.0**. Los datos viven en esa hoja y los archivos en la carpeta **CONSTRUBLOCK · Archivos** de Drive.

La app habla con el servidor por `doPost` (ver `apps-script/Code.gs`). La versión dentro de Google (`/exec`) funciona igual.

## Usuarios
Los PIN están en la pestaña `USUARIOS_APP` de la hoja. Roles: `DUENO` (todo), `MOSTRADOR` (ventas, entregas, clientes, inventario, gastos), `PRODUCCION` (block producido, ajustes, inventario), `CONTADOR` (facturas, gastos, clientes, ventas).
