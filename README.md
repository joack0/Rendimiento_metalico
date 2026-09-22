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
Carta diaria, Grupo x mes, Ranking productos, A corregir y Productos. Los resúmenes son fórmulas sobre la hoja
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
