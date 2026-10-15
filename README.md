# Control mensual — The Bubble Tea Bar (Centro)

Todo el control mensual del local, en un solo lugar: **Web (HTML + CSS + JS) + Python puro + macros de Excel**.
No hay servidor ni instalación: abrís las páginas con doble clic.

| Pieza | Archivo | Para qué sirve |
|---|---|---|
| **Tablero «Análisis mensual»** (web) | `index.html` | KPIs, 6 gráficos, filtros en cascada y tabla exploratoria del mes. Se abre solo, sin servidor |
| **Módulo de regresión** (web) | `js/*.js` + `css/app.css` | Motor de regresión lineal simple o múltiple con métricas, diagnóstico y predicción. **Motor listo; falta montar su pestaña en `index.html`** |
| **Regresión por consola** | `analisis_regresion.py` | El **mismo motor** que el módulo web, en Python de biblioteca estándar (sin `pip install`) |
| **Consolidados en Excel** | `Consolidados.bas` + `*.xlsm` | Macros que arman `ventas`, `caja`, `gastos`, `insumos`, `vasos` y el análisis del reporte |
| **Informe gerencial** | `generar_informe.py` + `INFORME_GERENCIAL_ANALITICA.docx` | Arma el informe del mes a partir de las cifras del tablero |
| **Datos** | `control_mensual_sept*.csv`, `sabores_septiembre.csv`, `ReporteVariablesProductos_….xlsx` | Insumos de las piezas anteriores (septiembre 2026) |

> **Las páginas necesitan internet** para los CDN: PapaParse 5.4.1, Plotly 2.35.2 y SheetJS 0.18.5.
> Si la máquina está sin internet, bajá esos tres archivos y cambiá los `src` del `<head>`.

---

## Índice

- [1. Empezar rápido](#1-empezar-rápido)
- [2. El tablero `index.html`](#2-el-tablero-indexhtml)
- [3. El módulo de regresión (`js/*.js`)](#3-el-módulo-de-regresión-jsjs)
- [4. Regresión por consola `analisis_regresion.py`](#4-regresión-por-consola-analisis_regresionpy)
- [5. Excel: macros y libros](#5-excel-macros-y-libros)
- [6. Cómo se calcula (motor estadístico)](#6-cómo-se-calcula-motor-estadístico)
- [7. Gráficos](#7-gráficos)
- [8. Archivos del repositorio](#8-archivos-del-repositorio)
- [9. Verificación](#9-verificación)
- [10. Limitaciones y pendientes](#10-limitaciones-y-pendientes)

---

## 1. Empezar rápido

```bash
# 1) El tablero del mes: abrí el archivo con doble clic
index.html

# 2) Lo mismo que el módulo de regresión, pero por consola y sin navegador
python analisis_regresion.py --listar
python analisis_regresion.py --y monto_total --x cantidad_grandes,cantidad_medianos,ticket_promedio
```

En el tablero el flujo es siempre el mismo: **arrastrá el archivo a la zona punteada** y listo.

---

## 2. El tablero `index.html`

Un solo archivo (1.286 líneas) pensado para ver el mes de un vistazo. Todo corre en el navegador,
sin Python ni servidor.

- **Menú naranja** con ilustraciones de bubble tea en **SVG embebido** (no depende de imágenes externas).
- **Dos zonas de arrastrar y soltar**: una para el **CSV del mes** y otra para el **`.xlsx` del reporte**
  (se busca sola la hoja `Detalle`; también acepta un CSV con las columnas `sabor` y `unidades`).
- **Filtros en cascada**: departamento → segmento (sección) → ítem, más rango de fechas y chips de día
  de la semana. Los KPIs, los gráficos y la tabla reaccionan a los filtros.
- **Ocho KPIs**: vendido, unidades, promedio por día, ticket por unidad, mejor día, ítem más vendido,
  segmento más vendido e ítems distintos.
- **Seis gráficos**, elegibles con botones del menú:

| Botón | Qué muestra |
|---|---|
| **Ventas por día** | monto de cada día (barras) + línea del promedio diario |
| **Vasos grandes / medianos** | unidades de GRANDES y MEDIANOS (barras) y su monto (eje derecho) |
| **Sabores más vendidos** | ranking por unidades, tomado del **detalle del reporte** (`Sabor` / `Sabor 1` / `Sabor 2`) o, si no se cargó, de los nombres de ítem del CSV |
| **Mejores días** | días ordenados de mayor a menor, con el promedio marcado |
| **Por departamento** | participación del monto (dona): Bebidas, Complementos, Pastelería, Merchandise |
| **Por sección** | monto de cada sección del archivo, de mayor a menor |

- **Tabla exploratoria**: una fila por ítem con departamento, segmento, unidades, monto y % del total
  (con barra), ordenable por cualquier columna y con buscador. Muestra hasta 100 ítems por vez.

La página detecta las columnas solo, así que también acepta variantes tipo
`Producto / Monto Total de la Venta / Cantidad`. Expone `window.AnalisisMensual` para las pruebas.

**Departamentos** (agrupan las secciones del archivo en dos niveles):
Bebidas (`GRANDES`, `MEDIANOS`) · Complementos (`TOPPINGS`, `EXTRAS`, `BUBBLEBATH`) ·
Pastelería (`PASTELERIA`) · Merchandise (`MERCHANDISE`).

---

## 3. El módulo de regresión (`js/*.js`)

Es el módulo de **regresión lineal simple o múltiple** con métricas, diagnóstico y predicción.
Su **motor está hecho y verificado**, pero su pestaña **todavía no está montada** en el `index.html`
actual (que hoy es el tablero del punto 2). Para usarlo desde el navegador hay que montar el
markup `#reg-…` y la barra de pestañas `[Tablero] [Regresión]` — ver la Fase 3 de
[`PLAN_REGRESION.md`](PLAN_REGRESION.md).

| Archivo | Qué es |
|---|---|
| `js/estadistica.js` | Motor OLS (QR/Householder), distribuciones t y F, métricas, VIF, validación cruzada y predicción |
| `js/datos.js` | Lectura CSV/XLSX, reconocimiento de variables y pivote día / ítem-día |
| `js/graficos.js` | Los 7 gráficos con Plotly |
| `js/app.js` | Estado de la UI, pestañas, selección Y/X, resultados, predicción y escenarios |
| `css/app.css` | Estilos, barra de pestañas y layout del módulo |

**Sin módulos ES:** los `js/*.js` se cargan con `<script src>` clásico y exponen objetos globales
(`window.Estadistica`, `window.Datos`, `window.Graficos`), porque `type="module"` falla bajo `file://`.

### 3.1 Pestaña Regresión: 5 pasos en una pantalla

1. **Cargar** — arrastrá el archivo (o hacé clic). Acepta tres formatos:
   - **CSV transaccional**: `fecha, dia_semana, seccion, item, cantidad, precio, monto` (el `control_mensual_sept.csv`).
   - **CSV / hoja matriz diaria**: un concepto por fila y un día por columna (como salen `ventas`, `caja`, `gastos`, `insumos`).
   - **`.xlsx` / `.xlsm` directo**: se lee con SheetJS; si el libro tiene varias hojas, elegís la hoja. Las macros se ignoran: solo se leen valores (y fórmulas ya calculadas).
2. **Unidad + variables** — elegí la unidad de análisis, la hoja (si aplica), la **Y** (una sola) y una o más **X**. El tipo de cada X se puede alternar (numérica ↔ categórica) con el enlace de tipo; las categóricas generan variables *dummy*.
3. **Ajustar** — marcá/desmarcá intercepto, elegí el `k` de la validación cruzada y tocá **Ajustar modelo**. Con 1 X se comporta como regresión **simple**; con 2 o más, como **múltiple**.
4. **Ver** — ecuación, coeficientes (error estándar, t, p, IC 95 %, β estandarizado), métricas de ajuste y de error, diagnóstico (VIF, Durbin-Watson, outliers de Cook) y los 7 gráficos.
5. **Predecir** — completá los valores de las X y obtené **Ŷ** con IC 95 % para la media y para una observación nueva. Podés cargar varios **escenarios** y exportarlos a CSV.

### 3.2 Unidades de análisis

| Unidad | Filas | Y típica | Predictoras típicas |
|---|---|---|---|
| **Día** (por defecto) | 28 | `monto_total` | `cantidad_grandes`, `cantidad_medianos`, `ticket_promedio`, `dia_mes`, `fin_de_semana`, `dia_semana`… |
| **Ítem-día** | 330 | `monto` de la línea | `cantidad`, `precio`, `seccion`, `item`, `dia_semana` |

### 3.3 Variables que el módulo construye solo (nivel día)

`monto_total`, `unidades_total`, `cantidad_grandes`, `cantidad_medianos`, `monto_grandes`,
`monto_medianos`, `monto_toppings`, `monto_pasteleria`, `monto_bubblebath`, `monto_merchandise`,
`monto_extras`, `monto_delivery`, `pedidos_ya`, `unidades_promo`, `ticket_promedio`,
`precio_promedio_vaso`, `items_distintos`, `dia_mes` (tendencia), `dia_semana` (categórica) y
`fin_de_semana` (0/1).

Si cargás una hoja matriz (por ejemplo `caja`), cada concepto aparece como una variable más
(`efectivo`, `qr_neto`, `tarjeta_neto`, `total_gastos`, …) y se suma a la matriz diaria.

---

## 4. Regresión por consola `analisis_regresion.py`

Replica en Python (solo biblioteca estándar) el análisis del módulo web: **lee los mismos
archivos, arma las mismas variables del día y ajusta el mismo modelo OLS**. Útil para dejar un
informe reproducible sin abrir el navegador.

```bash
python analisis_regresion.py                          # modelo por defecto
python analisis_regresion.py --listar                 # ver todas las variables disponibles
python analisis_regresion.py --y monto_total --x cantidad_grandes
python analisis_regresion.py --x cantidad_grandes,cantidad_medianos \
                             --predecir cantidad_grandes=5,cantidad_medianos=8
python analisis_regresion.py --archivo "260901 CONTROL MENSUAL CENTRO.xlsm" --hoja caja
python analisis_regresion.py --unidad item-dia --y monto --x cantidad,seccion --cat seccion
```

| Opción | Qué hace |
|---|---|
| `--archivo`, `--hoja` | Archivo a leer y hoja del libro (por ej. `ventas`, `caja`, `gastos`) |
| `--concepto COLUMNA` | En formato matriz, cuál columna es el concepto (la hoja `ventas` tiene `SECCION` e `ITEM`) |
| `--unidad dia\|item-dia` | Una fila por día (28) o por línea de venta (330) |
| `--y`, `--x`, `--cat` | Dependiente, predictoras y cuáles se tratan como categóricas |
| `--sin-intercepto`, `--kfold N` | Ajuste sin intercepto y folds de la validación cruzada (0 = deshabilitada) |
| `--predecir VAR=VALOR,…` | Predicción (repetible) con IC 95 % y aviso de extrapolación |
| `--escenarios-csv`, `--json` | Guardar escenarios de predicción / toda la salida en JSON |
| `--listar`, `--version` | Listar variables o mostrar la versión |

**No escribe ningún archivo salvo que se lo pidas** con `--json` o `--escenarios-csv`.
Las hojas diarias `dd.mm.aa` se **rechazan a propósito** (no tienen formato de tabla): de un libro
se leen los consolidados, para no armar un modelo con datos mal leídos.

---

## 5. Excel: macros y libros

El módulo **`Consolidados.bas`** recorre las pestañas diarias (`01.09.26`, `02.09.26`, …) y rehace
los consolidados del mes. Detecta las hojas por el formato del nombre (`dd.mm.yy`), así que no
depende de las fechas.

| Botón del panel | Hoja que genera |
|---|---|
| 1) GENERAR MES | Crea las hojas diarias `01.mm.aa`… copiando `MODELO DIA` |
| 2) Consolidar VENTAS | `ventas` |
| 3) Consolidar CAJA | `caja` |
| 4) Consolidar GASTOS | `gastos` |
| 5) Consolidar INSUMOS | `insumos` |
| 6) Consolidar VASOS | `vasos` |
| 7) ANALIZAR REPORTE VARIABLES | `ANALISIS VARIABLES` (a partir del reporte del sistema) |
| CONSOLIDAR TODO | Las cinco de una |

Los consolidados salen como **matriz**: un concepto por fila y **cada día en una columna**, más una
columna `TOTAL`.

En **`ventas`** cada fila es un *ítem + precio*: además de la columna `PRECIO U` (columna C),
las cantidades quedan **separadas por precio** — un ítem vendido a 28 y a 26 sale en dos filas,
una por precio, con su `TOTAL CANT` y su `TOTAL $` de ese precio (no de todo el mes).

**Guía completa paso a paso** (guardar como `.xlsm`, importar el módulo, crear el panel, arrancar un
mes nuevo, qué hace el botón 7 y cómo ajustar el módulo si cambia el diseño del arqueo):
ver **[`COMO_USAR.md`](COMO_USAR.md)**.

---

## 6. Cómo se calcula (motor estadístico)

Mismo motor en las dos piezas (`js/estadistica.js` y `analisis_regresion.py`), en código propio,
sin librerías de cálculo:

- **Mínimos cuadrados ordinarios resueltos por descomposición QR (Householder)** en vez de invertir `X'X` a mano → más estable numéricamente.
- **Ajuste**: R², R² ajustado, n, k, grados de libertad, σ (error estándar de la estimación) y F global con p-valor.
- **Error**: MAE, RMSE, MAPE, error medio (sesgo) y **validación cruzada k-fold** (R², RMSE y MAE fuera de muestra).
- **Por coeficiente**: β, error estándar, t, p-valor, IC 95 % y β estandarizado.
- **Diagnóstico**: VIF y tolerancia, Durbin-Watson, apalancamiento, residuo estudentizado, distancia de Cook, outliers, asimetría y curtosis.
- **Distribuciones** t de Student y F implementadas con la función beta incompleta (Lanczos + fracción continua), como en libros de recetas numéricas.
- **Predicción**: `Ŷ`, IC 95 % para la media y para una observación nueva (`σ²·(1 + x₀'(X'X)⁻¹x₀)`), con aviso de **extrapolación** cuando un valor cae fuera del rango observado.

### Controles de realismo

- Con menos de ~10 observaciones por predictora el módulo **avisa** (n = 28 en el nivel día → conviene usar 2–3 X como máximo).
- Si un VIF supera 10, lo marca como colineal.
- Las variables de caja (efectivo/QR/tarjeta) son **consecuencia** de las ventas, no causa: se señalan como tales.
- Priorizá siempre el **R² de validación cruzada**, no el R² del ajuste.

---

## 7. Gráficos

En el módulo de regresión (`js/graficos.js`, con Plotly):

1. **Dispersión X vs Y + recta ajustada** con bandas de confianza y predicción 95 % (regresión simple; es el *gráfico representativo del modelo*).
2. Y real vs Ŷ con recta de 45° y R² anotado (múltiple).
3. Residuos vs valores ajustados.
4. QQ-plot de residuos.
5. Coeficientes estandarizados ± IC 95 % (importancia).
6. Regresión parcial (*added-variable*) por cada X, seleccionable.
7. Serie temporal: Y real vs predicho día a día.

El gráfico 1 se marca como representativo en el encabezado de resultados, con la ecuación escrita
(`Y = b0 + b1·X1 + …`) y el efecto marginal de cada variable.

El tablero `index.html` tiene sus propios 6 gráficos (ver el punto 2).

---

## 8. Archivos del repositorio

### La web

```
index.html              1286 líneas  Tablero «Análisis mensual» (menú naranja, KPIs, 6 gráficos, tabla exploratoria).
css/app.css              131 líneas  Estilos del módulo de regresión (barra de pestañas y layout).
js/estadistica.js        449 líneas  Motor OLS + métricas + distribuciones + VIF + validación cruzada + predicción.
js/datos.js              506 líneas  Carga CSV/XLSX, reconocimiento de variables y pivote día / ítem-día.
js/graficos.js           167 líneas  Los 7 gráficos con Plotly.
js/app.js                558 líneas  Estado de la UI, pestañas, selección Y/X, resultados, predicción y escenarios.
                                    (Necesita el markup #reg-… en el HTML: hoy no está montado.)
analisis_regresion.py   1791 líneas  La misma regresión por consola, en Python puro.
```

### Datos

| Archivo | Contenido |
|---|---|
| `control_mensual_sept.csv` | Septiembre **transaccional**: 330 filas, 28 días, 37 ítems, 1.587 unidades, **$30.016,8** |
| `control_mensual_septiembre.csv` | El mismo mes guardado como `df` de trabajo (idéntico al anterior) |
| `sabores_septiembre.csv` | 1.635 filas, una por (comanda, sabor): `fecha, dia_semana, producto, cantidad, posicion, sabor, monto` |
| `ReporteVariablesProductos_….xlsx` | El reporte del sistema, con las hojas `Detalle` y `Resumen` |

### Excel y macros

| Archivo | Qué es |
|---|---|
| `Consolidados.bas` | Módulo VBA de consolidados (incluye el botón 7 de análisis del reporte) |
| `ConsolidadosMira.bas` | Variante del módulo para el local Mira |
| `MODELO CONTROL MENSUAL CENTRO.xlsm` | Plantilla mensual vacía: `PANEL`, hojas de salida y `MODELO DIA` |
| `260901 CONTROL MENSUAL CENTRO.xlsm` | Septiembre **con datos** (35 hojas) + hoja `ANALISIS` con fórmulas y 4 gráficos |
| `septiembre.xlsm` | Septiembre con el panel de 8 botones, el reporte pegado y `ANALISIS VARIABLES` ya generada |
| `construir_modelo.py` | Regenera el `MODELO …xlsm` desde la plantilla e incrusta el módulo y los botones (usa Excel real por COM) |
| `generar_informe.py` | Genera `INFORME_GERENCIAL_ANALITICA.docx` desde las cifras del tablero |
| `260901 … .xlsx`, `… (respaldo pre-analisis).xlsm`, `… (respaldo pre-precio unitario).xlsm`, `MODELO … (respaldo).xlsm` | Copias históricas, sin uso diario |
| `260109 CONTROL MENSUAL MIRA.xlsx` | Libro del local Mira (fuera del alcance de la regresión de Centro) |
| `COMO_USAR.md` | Guía de las macros y del libro de Excel |
| `PLAN_REGRESION.md` | Plan y decisiones del módulo de regresión, con su estado |

### Verificación

```
_val/validar.js                    Motor OLS en JS (genera salida_js.json)
_val/validar.py                    El mismo ajuste con statsmodels (genera salida_py.json)
_val/validar_regresion_py.py       Compara el motor JS, el de Python puro y statsmodels
_val/validar_datos.js              Criterios de aceptación con los js/*.js reales
_val/validar_analisis_mensual.js   Lógica del tablero index.html sin navegador
_val/smoke_mensual.js              End-to-end del tablero en Chrome headless (CDP)
_val/smoke.js                      End-to-end del módulo de regresión en Chrome headless (requiere su pestaña montada)
_val/validar_analisis_variables.py Compara la macro del botón 7 contra el reporte (requiere Excel)
_val/matriz_prueba.csv             Matriz diaria mínima para probar el pivote
_val/_gen_mira.py                  Genera/valida material del local Mira
_val/salida_js.json, salida_py.json, py_simple.json   Salidas de referencia cruzada JS ↔ Python
```

---

## 9. Verificación

Todo es opcional: las páginas no lo necesitan para funcionar. Requiere **Node**, **Python** y **Chrome**.

```bash
cd _val
node validar.js && python validar.py    # motor OLS vs statsmodels
node validar_datos.js                   # criterios de aceptación con los JS reales
node validar_analisis_mensual.js        # lógica del tablero
node smoke_mensual.js                   # end-to-end del tablero en Chrome headless
cd .. && python _val/validar_regresion_py.py   # motor JS vs Python puro vs statsmodels
```

Resultados de la última corrida (**exit 0** en los tres que no dependen del navegador):

| Prueba | Resultado |
|---|---|
| `validar.js` | `n=28`, `total=30016.8`, R² simple **0,4310**, R² múltiple **0,9605**, F = 194,62 |
| `validar_datos.js` | **Todos los criterios OK** · detecta el CSV transaccional (28 días, 330 filas, 37 ítems, $30.016,8), ajusta ítem-día (R²=0,7383, n=330), detecta formato matriz, diferencia máxima vs Python **5,8e-13**, VIF vs statsmodels **4e-4**, avisa n bajo |
| `validar_analisis_mensual.js` | **Todo OK** · totales, 28 barras que suman el total, vasos grandes/medianos, ranking de sabores, orden de «mejores días», tabla exploratoria y KPIs con y sin filtro |
| `smoke.js` (regresión) | ⏳ **no se puede correr** hasta montar la pestaña Regresión en `index.html` (Fase 3 del plan) |

Números de referencia del modelo por defecto (`monto_total ~ cantidad_grandes + cantidad_medianos + ticket_promedio`):

- R² = **0,9605** · R² ajustado = 0,9556 · σ = 70,56 · F = 194,62 (p < 1e-15).
- Predicción para `(5, 8, 30)`: **Ŷ = 415,7429**, IC 95 % de la media `[294,2; 537,3]` y de una observación nueva `[226,1; 605,4]` — idénticos a `statsmodels`.

Los scripts que dependen de **Excel instalado** (`_val/validar_analisis_variables.py` y
`construir_modelo.py`) no forman parte de esta corrida: se ejecutan contra Excel por COM.

---

## 10. Limitaciones y pendientes

- **La pestaña Regresión no está montada.** El motor (fases 1, 2 y 4) está hecho y verificado, pero `js/app.js` necesita el markup `#reg-…` y la barra de pestañas que hoy no existen en `index.html`. Es el pendiente principal (Fase 3 del plan). Mientras tanto, la misma regresión se obtiene con `analisis_regresion.py`.
- **n = 28 días.** Con muchas X el R² sube artificialmente. Regla práctica: **máximo 2–3 predictores**.
- La multicolinealidad es esperable (grandes y medianos comparten estructura con las unidades totales).
- Septiembre tiene promociones puntuales (Promo 50 %, stickers, Mediano Promo premium) que actúan como *outliers*; la distancia de Cook los señala.
- Correlación no es causalidad: los resultados describen el mes, no son una ley.
- El ranking de sabores del tablero **junta variantes de mayúsculas** (`Lavanda` 117 + `LAVANDA` 9 = 126); la macro del Excel las lista separadas.
- Las **combinaciones** `Sabor 1 + Sabor 2` están en la hoja `ANALISIS VARIABLES` del Excel, no en el tablero web.
- **Pendiente (Fase 5 del plan):** la macro `ExportarDatosAnalisis` en `Consolidados.bas`, que escribirá `datos/datos_diarios.csv` con fecha, ventas por sección, caja, gastos e insumos (una fila por día). Mientras tanto, ese mismo contenido se carga arrastrando el `.xlsm` y eligiendo la hoja (`caja`, `gastos`, `insumos`).
- Las páginas **necesitan internet** para los tres CDN (PapaParse, Plotly, SheetJS).
