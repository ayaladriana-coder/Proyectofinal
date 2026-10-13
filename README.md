# Proyectofinal
# Control mensual — The Bubble Tea Bar (Centro)

Todo el control mensual del local, en un solo lugar: **Web (HTML + CSS + JS) + Python puro + macros de Excel**.
No hay servidor ni instalación: abrís la página con doble clic.

| Pieza | Archivo | Para qué sirve |
|---|---|---|
| **Tablero + Regresión** (web) | `index.html` | El tablero del mes (KPIs y 4 gráficos) y una **app de regresión lineal simple o múltiple** con métricas, diagnóstico y predicción |
| **Análisis del mes** (web) | `analisis_mensual.html` | Página suelta con menú naranja: ventas por día, vasos grandes/medianos, sabores más vendidos y mejores días |
| **Regresión por consola** | `analisis_regresion.py` | El **mismo motor** que la web, en Python de biblioteca estándar (sin `pip install`) |
| **Consolidados en Excel** | `Consolidados.bas` + `*.xlsm` | Macros que arman `ventas`, `caja`, `gastos`, `insumos`, `vasos` y el análisis del reporte |
| **Datos** | `control_mensual_sept*.csv`, `sabores_septiembre.csv` | Insumos de las tres piezas anteriores (septiembre 2026) |

> **Las páginas necesitan internet** para los CDN: PapaParse 5.4.1, Plotly 2.35.2 y SheetJS 0.18.5.
> Si la máquina está sin internet, bajá esos tres archivos y cambiá los `src` del `<head>`.

---

## Índice

- [1. Empezar rápido](#1-empezar-rápido)
- [2. La app `index.html`](#2-la-app-indexhtml)
- [3. Página de análisis `analisis_mensual.html`](#3-página-de-análisis-analisis_mensualhtml)
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
# 1) La app (tablero + regresión): abrí el archivo con doble clic
index.html

# 2) La página de análisis del mes
analisis_mensual.html

# 3) Lo mismo que la app, pero por consola y sin navegador
python analisis_regresion.py --listar
python analisis_regresion.py --y monto_total --x cantidad_grandes,cantidad_medianos,ticket_promedio
```

En las dos páginas, el flujo es el mismo: **arrastrá el archivo a la zona punteada** y listo.

---

## 2. La app `index.html`

Dos pestañas. La de **Tablero** es el tablero que ya existía (KPIs y 4 gráficos) y conserva su
lógica original; la de **Regresión** es la app nueva.

### 2.1 Pestaña Regresión: 5 pasos en una pantalla

1. **Cargar** — arrastrá el archivo (o hacé clic). Acepta tres formatos:
   - **CSV transaccional**: `fecha, dia_semana, seccion, item, cantidad, precio, monto` (el `control_mensual_sept.csv`).
   - **CSV / hoja matriz diaria**: un concepto por fila y un día por columna (como salen `ventas`, `caja`, `gastos`, `insumos`).
   - **`.xlsx` / `.xlsm` directo**: se lee con SheetJS; si el libro tiene varias hojas, elegís la hoja. Las macros se ignoran: solo se leen valores (y fórmulas ya calculadas).
2. **Unidad + variables** — elegí la unidad de análisis, la hoja (si aplica), la **Y** (una sola) y una o más **X**. El tipo de cada X se puede alternar (numérica ↔ categórica) con el enlace de tipo; las categóricas generan variables *dummy*.
3. **Ajustar** — marcá/desmarcá intercepto, elegí el `k` de la validación cruzada y tocá **Ajustar modelo**. Con 1 X se comporta como regresión **simple**; con 2 o más, como **múltiple**.
4. **Ver** — ecuación, coeficientes (error estándar, t, p, IC 95 %, β estandarizado), métricas de ajuste y de error, diagnóstico (VIF, Durbin-Watson, outliers de Cook) y los 7 gráficos.
5. **Predecir** — completá los valores de las X y obtené **Ŷ** con IC 95 % para la media y para una observación nueva. Podés cargar varios **escenarios** y exportarlos a CSV.

### 2.2 Unidades de análisis

| Unidad | Filas | Y típica | Predictoras típicas |
|---|---|---|---|
| **Día** (por defecto) | 28 | `monto_total` | `cantidad_grandes`, `cantidad_medianos`, `ticket_promedio`, `dia_mes`, `fin_de_semana`, `dia_semana`… |
| **Ítem-día** | 330 | `monto` de la línea | `cantidad`, `precio`, `seccion`, `item`, `dia_semana` |

### 2.3 Variables que la app construye sola (nivel día)

`monto_total`, `unidades_total`, `cantidad_grandes`, `cantidad_medianos`, `monto_grandes`,
`monto_medianos`, `monto_toppings`, `monto_pasteleria`, `monto_bubblebath`, `monto_merchandise`,
`monto_extras`, `pedidos_ya`, `unidades_promo`, `ticket_promedio`, `precio_promedio_vaso`,
`items_distintos`, `dia_mes` (tendencia), `dia_semana` (categórica) y `fin_de_semana` (0/1).

Si cargás una hoja matriz (por ejemplo `caja`), cada concepto aparece como una variable más
(`efectivo`, `qr_neto`, `tarjeta_neto`, `total_gastos`, …) y se suma a la matriz diaria.

---

## 3. Página de análisis `analisis_mensual.html`

Un solo archivo, pensado para ver el mes de un vistazo:

- **Menú naranja** con ilustraciones de bubble tea en **SVG embebido** (no depende de imágenes externas).
- **Dos zonas de arrastrar y soltar**: una para el **CSV del mes** y otra para el **`.xlsx` del reporte** (o un CSV con columna `sabor`).
- Cuatro gráficos, elegibles con botones:

| Botón | Qué muestra |
|---|---|
| **Ventas por día** | monto de cada día (barras) + línea del promedio diario |
| **Vasos grandes / medianos** | unidades de GRANDES y MEDIANOS (barras) y su monto (eje derecho) |
| **Sabores más vendidos** | ranking por unidades, tomado del **detalle del reporte** (variable `Sabor` / `Sabor 1` / `Sabor 2`) |
| **Mejores días** | días ordenados de mayor a menor, con el promedio marcado |

Además muestra 5 tarjetas de resumen (vendido, unidades, promedio diario, mejor día, ítem más
vendido) y detecta las columnas solo, así que también acepta variantes tipo
`Producto / Monto Total de la Venta / Cantidad`.

---

## 4. Regresión por consola `analisis_regresion.py`

Replica en Python (solo biblioteca estándar) el análisis de la pestaña Regresión: **lee los mismos
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

**Guía completa paso a paso** (guardar como `.xlsm`, importar el módulo, crear el panel, arrancar un
mes nuevo, qué hace el botón 7 y cómo ajustar el módulo si cambia el diseño del arqueo):
ver **[`COMO_USAR.md`](COMO_USAR.md)**.

---

## 6. Cómo se calcula (motor estadístico)

Mismo motor en las tres piezas (`js/estadistica.js` y `analisis_regresion.py`), en código propio,
sin librerías de cálculo:

- **Mínimos cuadrados ordinarios resueltos por descomposición QR (Householder)** en vez de invertir `X'X` a mano → más estable numéricamente.
- **Ajuste**: R², R² ajustado, n, k, grados de libertad, σ (error estándar de la estimación) y F global con p-valor.
- **Error**: MAE, RMSE, MAPE, error medio (sesgo) y **validación cruzada k-fold** (R², RMSE y MAE fuera de muestra).
- **Por coeficiente**: β, error estándar, t, p-valor, IC 95 % y β estandarizado.
- **Diagnóstico**: VIF y tolerancia, Durbin-Watson, apalancamiento, residuo estudentizado, distancia de Cook, outliers, asimetría y curtosis.
- **Distribuciones** t de Student y F implementadas con la función beta incompleta (Lanczos + fracción continua), como en libros de recetas numéricas.
- **Predicción**: `Ŷ`, IC 95 % para la media y para una observación nueva (`σ²·(1 + x₀'(X'X)⁻¹x₀)`), con aviso de **extrapolación** cuando un valor cae fuera del rango observado.

### Controles de realismo

- Con menos de ~10 observaciones por predictora la app **avisa** (n = 28 en el nivel día → conviene usar 2–3 X como máximo).
- Si un VIF supera 10, lo marca como colineal.
- Las variables de caja (efectivo/QR/tarjeta) son **consecuencia** de las ventas, no causa: se señalan como tales.
- Priorizá siempre el **R² de validación cruzada**, no el R² del ajuste.

---

## 7. Gráficos

En la app (`js/graficos.js`, con Plotly):

1. **Dispersión X vs Y + recta ajustada** con bandas de confianza y predicción 95 % (regresión simple; es el *gráfico representativo del modelo*).
2. Y real vs Ŷ con recta de 45° y R² anotado (múltiple).
3. Residuos vs valores ajustados.
4. QQ-plot de residuos.
5. Coeficientes estandarizados ± IC 95 % (importancia).
6. Regresión parcial (*added-variable*) por cada X, seleccionable.
7. Serie temporal: Y real vs predicho día a día.

El gráfico 1 se marca como representativo en el encabezado de resultados, con la ecuación escrita
(`Y = b0 + b1·X1 + …`) y el efecto marginal de cada variable.

---

## 8. Archivos del repositorio

### La app

```
index.html               465 líneas  Pestañas [Tablero] [Regresión]. El tablero conserva su lógica original.
css/app.css              131 líneas  Barra de pestañas y estilos de la app de regresión.
js/estadistica.js        449 líneas  Motor OLS + métricas + distribuciones + VIF + validación cruzada + predicción.
js/datos.js              506 líneas  Carga CSV/XLSX, reconocimiento de variables y pivote día / ítem-día.
js/graficos.js           167 líneas  Los 7 gráficos con Plotly.
js/app.js                558 líneas  Estado de la UI, pestañas, selección Y/X, resultados, predicción y escenarios.
analisis_mensual.html    840 líneas  Página de análisis del mes (menú naranja, 4 gráficos).
analisis_regresion.py   1791 líneas  La misma regresión por consola, en Python puro.
```

**Sin módulos ES:** los `js/*.js` se cargan con `<script src>` clásico y exponen objetos globales
(`window.Estadistica`, `window.Datos`, `window.Graficos`), porque `type="module"` falla bajo `file://`.

### Datos

| Archivo | Contenido |
|---|---|
| `control_mensual_sept.csv` | Septiembre **transaccional**: 330 filas, 28 días, 37 ítems, 1.587 unidades, **$30.016,8** |
| `control_mensual_septiembre.csv` | El mismo mes guardado como `df` de trabajo (idéntico al anterior) |
| `sabores_septiembre.csv` | 1.635 filas, una por (comanda, sabor): `fecha, dia_semana, producto, cantidad, posicion, sabor, monto` |
| `ReporteVariablesProductos_…xlsx` | El reporte del sistema, con las hojas `Detalle` y `Resumen` |

### Excel y macros

| Archivo | Qué es |
|---|---|
| `Consolidados.bas` | Módulo VBA de consolidados (incluye el botón 7 de análisis del reporte) |
| `MODELO CONTROL MENSUAL CENTRO.xlsm` | Plantilla mensual vacía: `PANEL`, hojas de salida y `MODELO DIA` |
| `260901 CONTROL MENSUAL CENTRO.xlsm` | Septiembre **con datos** (35 hojas) + hoja `ANALISIS` con fórmulas y 4 gráficos |
| `septiembre.xlsm` | Septiembre con el panel de 8 botones, el reporte pegado y `ANALISIS VARIABLES` ya generada |
| `construir_modelo.py` | Regenera el `MODELO …xlsm` desde la plantilla e incrusta el módulo y los botones (usa Excel real por COM) |
| `260901 … .xlsx`, `… (respaldo pre-analisis).xlsm`, `MODELO … (respaldo).xlsm` | Copias históricas, sin uso diario |
| `COMO_USAR.md` | Guía de las macros y del libro de Excel |
| `PLAN_REGRESION.md` | Plan y decisiones del aplicativo de regresión, con su estado |

### Verificación

```
_val/validar.js                 Motor OLS en JS (genera salida_js.json)
_val/validar.py                 El mismo ajuste con statsmodels (genera salida_py.json)
_val/validar_regresion_py.py    Compara el motor JS, el de Python puro y statsmodels
_val/validar_datos.js           Criterios de aceptación con los js/*.js reales
_val/validar_analisis_mensual.js  Lógica de analisis_mensual.html sin navegador
_val/smoke.js                   End-to-end en Chrome headless (CDP): drag&drop, ajuste, gráficos, predicción
_val/validar_analisis_variables.py  Compara la macro del botón 7 contra el reporte (requiere Excel)
_val/matriz_prueba.csv          Matriz diaria mínima para probar el pivote
_val/salida_js.json             Salidas generadas por la suite (referencia cruzada JS ↔ Python)
_val/salida_py.json
_val/py_simple.json
```

---

## 9. Verificación

Todo es opcional: la app no lo necesita para funcionar. Requiere **Node**, **Python** y **Chrome**.

```bash
cd _val
node validar.js && python validar.py    # motor OLS vs statsmodels
node validar_datos.js                   # criterios de aceptación con los JS reales
node validar_analisis_mensual.js        # lógica de la página de análisis
node smoke.js                           # end-to-end en Chrome headless
cd .. && python _val/validar_regresion_py.py   # motor JS vs Python puro vs statsmodels
```

Resultados de la última corrida (todos con **exit 0**):

| Prueba | Resultado |
|---|---|
| `validar_regresion_py.py` | **30 OK, 0 fallas** · diferencia máxima vs `statsmodels` **9,7e-13** (criterio 1e-6); vs motor JS **3,8e-13**; predicción 2,3e-13; validación cruzada 8,5e-14 |
| `validar.js` | `n=28`, `total=30016.8`, R² simple **0,4310**, R² múltiple **0,9605**, F = 194,62 |
| `validar_datos.js` | **28 criterios OK, 0 fallas** · detecta el CSV transaccional: 28 días, 330 filas, 37 ítems, $30.016,8 |
| `validar_analisis_mensual.js` | **35 cheques OK, 0 fallas** · totales, 28 barras que suman el total, vasos grandes/medianos, ranking de sabores y orden de «mejores días» |
| `smoke.js` | **19 pasos OK, 0 fallas** · carga por arrastrar y soltar, cambio de pestañas, modelo simple y múltiple, los 7 gráficos, predicción con IC, aviso de extrapolación, escenarios y export — **sin excepciones ni errores de consola** |

Números de referencia del modelo por defecto (`monto_total ~ cantidad_grandes + cantidad_medianos + ticket_promedio`):

- R² = **0,9605** · R² ajustado = 0,9556 · σ = 70,56 · F = 194,62 (p < 1e-15).
- Predicción para `(5, 8, 30)`: **Ŷ = 415,7429**, IC 95 % de la media `[294,2; 537,3]` y de una observación nueva `[226,1; 605,4]` — idénticos a `statsmodels`.

Los dos scripts que dependen de **Excel instalado** (`_val/validar_analisis_variables.py` y
`construir_modelo.py`) no forman parte de esta corrida: se ejecutan contra Excel por COM.

---

## 10. Limitaciones y pendientes

- **n = 28 días.** Con muchas X el R² sube artificialmente. Regla práctica: **máximo 2–3 predictores**.
- La multicolinealidad es esperable (grandes y medianos comparten estructura con las unidades totales).
- Septiembre tiene promociones puntuales (Promo 50 %, stickers, Mediano Promo premium) que actúan como *outliers*; la distancia de Cook los señala.
- Correlación no es causalidad: los resultados describen el mes, no son una ley.
- El ranking de sabores de la página **junta variantes de mayúsculas** (`Lavanda` 117 + `LAVANDA` 9 = 126); la macro del Excel las lista separadas.
- Las **combinaciones** `Sabor 1 + Sabor 2` están en la hoja `ANALISIS VARIABLES` del Excel, no en la página web.
- **Pendiente (Fase 5 del plan):** la macro `ExportarDatosAnalisis` en `Consolidados.bas`, que escribirá `datos/datos_diarios.csv` con fecha, ventas por sección, caja, gastos e insumos (una fila por día). Mientras tanto, ese mismo contenido se carga arrastrando el `.xlsm` y eligiendo la hoja (`caja`, `gastos`, `insumos`).
- Las páginas **necesitan internet** para los tres CDN (PapaParse, Plotly, SheetJS).
