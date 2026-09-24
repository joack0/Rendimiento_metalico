# Guía para controlar el rendimiento metálico

Esta guía acompaña a `plantilla/Plantilla_control_rendimiento.xlsx`. La plantilla sigue el formato del
"Análisis de Rendimiento metálico 2010" (pérdidas T1 a T4, fuera de medida, corto T4, oxidación, barras
perdidas, pérdidas SAP, delta y Hecho / Causa / Acción) y le agrega resúmenes, una carta de control, un
Pareto y un plan de acción.

## 1. El indicador

**Rendimiento metálico = kg producidos / kg consumidos**

- Los kg producidos son las barras buenas. Los kg consumidos son la palanquilla cargada al horno.
- Siempre se pondera: Σ producción / Σ consumo. Si se promedian los porcentajes de cada orden, una orden chica
  pesa lo mismo que una campaña grande.
- Cada punto porcentual es mucho acero: en 10.000 t consumidas al mes, 1 pp son 100 t que vuelven como chatarra
  o se pierden como laminilla.

## 2. El balance de pérdidas

```
Pérdidas SAP = Consumo − Producción
             = T1 + T2 + T3 + T4 + Fuera de medida + Corto T4 + Cobles/barras perdidas + Oxidación + Delta
```

| Pérdida | Dónde se produce | Causas típicas |
|---|---|---|
| Oxidación / laminilla | Horno y descascarillado | Temperatura alta, tiempo de residencia largo, esperas con material caliente. Normalmente del 1 al 2,5 % del consumo (en 2010 se estimaba en 2 %). |
| T1, T2, T3 | Despuntes de las cizallas | Palanquilla con defectos en los extremos (hojas, rechupe), cabezas frías, largo de despunte mal calibrado, fotoceldas o HMD con fallas. |
| T4 | Corte a largo comercial | Sobrante de la barra madre por no ser múltiplo exacto del largo comercial. Depende del largo y peso de la palanquilla (padrón). |
| Corto T4 / fuera de medida | Mesa de enfriamiento y control dimensional | Barras cortas, dimensión fuera de tolerancia por desgaste de cilindros o guías, puesta a punto en cambios de medida. |
| Cobles / barras perdidas | Tren de laminación | Atascos, salidas de línea, fallas de guías o de señales. Cada coble puede costar cientos de kg. |
| Delta no explicado | — | Lo que nadie pesó o registró, o balanzas con error. |

**El delta es el primer indicador de la calidad del dato.** Si supera la tolerancia (1 % del consumo por
defecto), la orden aparece como "Revisar balance": antes de analizar el proceso hay que completar el registro.

## 3. Rutina de control

| Frecuencia | Qué se revisa | Dónde |
|---|---|---|
| Cada orden o turno | Rendimiento vs meta, pérdidas por etapa, cobles, estado del balance. Toda orden "Bajo meta" lleva Hecho, Causa y Acción. | Registro |
| Semanal (15 min) | Carta de control, Pareto de la semana o del mes, acciones nuevas para los ítems "Prioridad (80 %)". | Carta semanal, Pareto, Plan de acción |
| Mensual | Tendencia por grupo de producto, kg/t de pérdidas, eficacia de las acciones cerradas, ajuste de metas. | Resumen mensual, Plan de acción |

Antes de analizar, corrija los datos: una orden con rendimiento mayor a 100 % o con el delta fuera de tolerancia
distorsiona todos los resúmenes.

## 4. Cómo leer la carta de control

- **LC** es el nivel actual del proceso (rendimiento ponderado del año). **LSC y LIC** (LC ± 3σ, con
  σ = MR̄ / 1,128) marcan la variación normal. La **meta** es adonde se quiere llegar, y es otra cosa.
- **Un punto bajo LIC** indica que algo pasó esa semana (causa especial): busque en el Registro qué órdenes lo bajaron.
- **8 semanas seguidas bajo LC** indica que el proceso cambió de nivel: desgaste, otro proveedor de palanquilla,
  un cambio de práctica.
- **LC bajo la meta sin señales** significa que el proceso es estable pero no alcanza la meta. No sirve reaccionar
  a cada semana: hace falta un proyecto de mejora que cambie el proceso.
- Cuando una mejora se consolida, recalcule los límites (cambie el año en Parámetros o use una fecha de corte).

## 5. Cómo analizar una desviación

1. **Hecho**: lo que se observó, con datos. Ej.: "Mayor pérdida en T1 y T2, 2 barras perdidas".
2. **Causa**: pregunte "¿por qué?" hasta llegar a algo que se pueda corregir (5 porqués) y clasifíquela en
   una categoría de la lista, para que el Pareto por causa funcione.
3. **Acción**: contención (hoy), correctiva (elimina la causa) y preventiva (evita que se repita en otros productos).
4. **Eficacia**: compare el rendimiento del mismo producto antes y después. Si no mejoró, márquela
   "Cerrada - no eficaz" y vuelva a analizar.

Ejemplos del archivo 2010:

| Hecho | Causa | Acción |
|---|---|---|
| Cobles (trozo de coble y barra de sacrificio) | HMD A5-A6 pierde señal | Revisar señales de fotoceldas y HMD en IBA |
| Mayor pérdida en T1 y T2 | Palanquilla con defecto (hojas); largo de corte irregular en T2 | Alargar despunte de cizallas T1 y T2; corregir stand falso en A1 |
| Sobraba cola de 5 a 8 m en T4 | Largo de palanquilla | Revisar padrón |
| Pérdidas altas en T4 | Se usó palanquilla de 3,1 m en vez de la estándar de 3,2 m | Respetar el largo estandarizado |

## 6. Palancas típicas de mejora

- **Palanquilla**: largo y peso optimizados por producto para reducir el sobrante en T4; controlar defectos en la
  recepción; usar siempre el largo estandarizado.
- **Horno**: controlar temperatura y tiempo de residencia para reducir la laminilla; evitar esperas con material caliente.
- **Cizallas**: calibrar el despunte con fotoceldas / HMD y revisar señales cuando el despunte varía.
- **Laminación**: cambiar cilindros y guías según el tonelaje, hacer prueba de primera barra y estandarizar la puesta
  a punto para reducir fuera de medida y cobles.
- **Planificación**: campañas más largas y una secuencia de medidas con menos cambios.
- **Registro**: pesar y anotar todas las pérdidas en el turno, para que el delta tienda a cero.

## 7. Supuestos de la plantilla

- Meta general 95,8 % y metas por grupo: `Rend_Met` del CSV de consumo 2026, ponderado por kg consumidos.
  Estrella no tiene órdenes en el CSV y usa la meta general.
- Oxidación estimada: 2 % del consumo (como en 2010), salvo que se ingrese la oxidación medida.
- Tolerancia del delta: 1 % del consumo.
- "Barras perdidas" se registra en número y en kg; solo los kg entran en el balance.
- Todo se puede cambiar en la hoja Parámetros.
