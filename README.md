# Rendimiento metálico

Dashboard con la carta de control del rendimiento metálico (Rend. Met = kg producidos / kg consumidos).

## Actualizar

```bash
pip install openpyxl        # para leer el maestro de productos
python3 scripts/preparar_datos.py [datos/Consumo_lotes_de_produccion.csv] [datos/Datos_productos.xlsx]
```

Genera `dashboard/index.html`, un archivo único que se abre en el navegador.

```bash
python3 scripts/exportar_excel.py
```

Genera `dashboard/Rendimiento_metalico.xlsx` con las hojas Resumen, Ordenes (base completa), Carta mensual,
Carta semanal, Carta diaria, Grupo x mes, Grupo x semana, Grupo x día, Pareto semanal, Ranking productos, A corregir y Productos. Los resúmenes son fórmulas sobre la hoja
Ordenes: si se corrige una orden ahí, todo se recalcula.

## Reglas

- Las órdenes con rendimiento > 100 % se excluyen de los cálculos y aparecen en la tabla "Órdenes a corregir".
- El rendimiento de un período o producto es ponderado: Σ kg producidos / Σ kg consumidos.
- Carta de control I-MR: LC = rendimiento ponderado, σ = MR̄ / 1,128, límites = LC ± 3σ.
- Meta = columna `Rend_Met` del CSV, ponderada por kg consumidos.
- Grupo y largo vienen del maestro (`Tipo`, `CORTE T4`). Lo que no está en el maestro se deduce del nombre.
- Barra traspaso de carga pertenece a Redondo liso, y los productos SEMI se unifican con su equivalente.
- Diámetros, anchos y espesores en mm (las pulgadas se convierten ×25,4) y largos en m.
- La fecha de cada orden es su `Fecha_inicio`.
- Las semanas son ISO 8601: de lunes a domingo, y la semana 1 es la que contiene el primer jueves del año.

## Pareto semanal

- Mide los **kg faltantes** para cumplir la meta: kg consumidos × (meta − rendimiento).
- Los ítems que suman el 80 % de los kg faltantes de la semana son los prioritarios.
- **Actuar ya:** está dentro del 80 % y tiene una brecha ≥ 2 pp o estuvo bajo meta en ≥ 2 de las 4 semanas previas.
- **Plan de mejora:** está dentro del 80 %, pero no cumple lo anterior.
- **Monitorear:** está fuera del 80 %.

## Timeline de planta

```bash
python3 scripts/preparar_timeline.py [archivo_o_carpeta ...]
```

Genera `dashboard/timeline.html`: una fila por día del mes (00:00 → 24:00) con lo que pasó en el tren.
Sin argumentos incluye todos los `.csv` de `datos/interrupciones/` (en Windows, doble clic en `actualizar_timeline.bat`
regenera y abre la página).

### Alimentarlo siempre con el CSV

- **Vincular CSV…** (Chrome o Edge): se elige el CSV una vez. La página lo vuelve a leer cada vez que se abre y cada
  minuto mientras está abierta; basta con sobrescribir ese archivo con la nueva exportación (mismo nombre y carpeta).
  Al reabrir el navegador puede pedir un clic en **Reconectar** para confirmar el acceso.
- **Cargar CSV…** o arrastrar archivos: se agregan y quedan guardados en ese navegador, así se pueden juntar varios meses.
- **Carpeta `datos/interrupciones/`**: dejar ahí cada exportación y correr el script (o el `.bat`).
- Si dos archivos cubren las mismas fechas, manda el último cargado (o el de fecha de modificación más nueva); así no se
  duplican declaraciones al volver a exportar el mes. La tabla "Datos cargados" muestra qué cubre cada archivo y permite quitar cualquiera (también el incluido en la página),
  por ejemplo para ver solo el CSV del mes pasado.

### Reglas

- Cada declaración ocupa de `Fecha_Inicio` a `Fecha_Inicio + DuracionMinutos`; si cruza la medianoche se reparte entre ambos días.
- Sin declaración ni minutos = tren disponible / en producción (verde).
- Tipo: Parada programada (azul), Parada padrón (morado) y Libre (gris) según `Disciplina`; Retraso (naranjo) si `Zona`,
  `Equipo` u `Ocurrencia` dicen "retraso"; el resto es Interrupción/falla (rojo). Con "Colorear por: Disciplina" cada
  falla toma el color de su disciplina (Mecánica, Electrónica, Operación…).
- % prod. de cada día = minutos de producción / minutos con programa (excluye Libre y Sin registro) dentro de la ventana horaria.
- Lo que no cubre ningún archivo (antes de su primer día o después de su última declaración) queda como "Sin registro".
- Ver por **Mes**, **Semana** (ISO, lunes a domingo) o **Rango** de fechas; con ‹ › se pasa al período anterior o siguiente.
  Se puede cargar el HCA de todo el año en un solo CSV.
- Filtros por Tipo, Disciplina, Zona, Equipo (se encadenan: al elegir una disciplina solo quedan sus zonas y equipos) y
  búsqueda de texto en ocurrencia, hecho, causa, acción y producto. Lo que no cumple el filtro se ve como "Fuera del filtro".
- Ventana horaria (ej. 17:00 → 23:00) y L-V / Sáb-Dom. Clic en un día muestra su detalle; sin día elegido se listan las
  25 interrupciones más largas del período. Clic en la leyenda resalta una categoría.
- **Resumen** por mes o por semana de todo lo cargado: % prod., horas por tipo, N° de fallas y la disciplina con más falla.
  Clic en una fila la abre en el gráfico.
