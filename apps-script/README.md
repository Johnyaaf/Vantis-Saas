# Apps Script legado

Esta carpeta guarda los modulos originales del ERP en Google Apps Script antes de migrarlos al backend/frontend del SaaS.

## Orden de archivos segun el proyecto original

1. `src/M0_Setup.gs`
2. `src/M1_Sesion.gs`
3. `src/M2_Transacciones.gs`
4. `src/M3_Maestros.gs`
5. `src/M4_Auditoria.gs`
6. `src/M5_Batch.gs`
7. `src/M6_CalcMotores.gs`
8. `src/M7_Utils.gs`
9. `src/M8_Dashboard.gs`
10. `src/vantis_app.html`

## Criterio de migracion

- Los archivos `.gs` quedan aqui como referencia fuente.
- La logica de negocio se migra al backend por modulo.
- La interfaz HTML se migra al frontend cuando se defina la estructura final.
