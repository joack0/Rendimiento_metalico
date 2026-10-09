# Palanquillas vs Stock

Página única (`index.html`) que lee el programa de laminación (`.xlsm`, hoja *Prog LAM REN*) y el reporte de stock
de palanquilla (`.xlsx` de SAP) y compara lo que pide el programa con lo que hay, por código, calidad y largo.

- Se abre en el navegador y se cargan los archivos nuevos arrastrándolos a cada recuadro. Trae precargados
  `10_Prog_03_Lam-Renca_OCT_2026 - V0.1.xlsm` y `Stock_palanquilla_09_10_26.XLSX`.
- `plantilla.html` + `logica.js` son las fuentes; `index.html` es la plantilla con la lógica y los datos de ejemplo incrustados.

## Reglas

- Requerido = columna **TON** del programa; N° de palanquillas = **C/U**.
- Pendiente = órdenes cuyo **INICIO** es igual o posterior a la fecha del stock (se toma del nombre del archivo, `dd_mm_aa`)
  y sin producción **REAL** (o descontando lo real, según el selector).
- Stock disponible = **Libre utilización** (opcionalmente + inspección de calidad). Bloqueado se muestra aparte.
- Cruce por código de material; si no existe, por sección + calidad + largo de la descripción.
- *Se agota en*: primera orden, en orden de inicio, que ya no se cubre con el stock de ese código.
- *Alternativas*: otros largos de la misma calidad con saldo libre después de cubrir su propio uso en el programa.
