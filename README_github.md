# Analizador Comparativo de Series Financieras

Herramienta de análisis técnico que compara **dos series de precios cualesquiera** (acciones, ETFs, fondos, índices, criptos, divisas...) y calcula las métricas estadísticas estándar de la industria: **Beta, Alpha de Jensen, Correlación, R², Volatilidad, Sharpe, Sortino, Tracking Error, Information Ratio, Max Drawdown, CAGR** y captura up/down.

Funciona contra cualquier ticker de Yahoo Finance, archivos locales `.xlsx/.csv/.txt`, o cualquier combinación de ambos. Genera reporte Excel completo + dashboard HTML interactivo con Plotly.

> *Comparative analysis tool for financial time series. Computes Beta, Alpha (Jensen), R², Sharpe, Sortino, Tracking Error, Information Ratio, Max Drawdown and more between any two assets — fetched from Yahoo Finance or loaded from local files. Outputs Excel reports and interactive HTML dashboards.*

---

## Tres formas de usarla

| Forma | Cuándo | Archivo |
|---|---|---|
| Atajos `.bat` (doble click) | Lo más rápido, sin escribir nada en consola | `comparar_dos_activos.bat`, `correr.bat`, `correr_personalizado.bat`, `informe_desde_excel.bat` |
| Script Python | Máxima flexibilidad: cualquier ticker, archivo, periodo, benchmark | `analizador.py`, `informe_html.py`, `analizar_gld.py` |
| Plantilla Excel | Calculadora viva: pegas dos series y todas las métricas se calculan al instante | `Plantilla_Analisis_Comparativa.xlsx` |

---

## Demo rápida

```bash
# Comparar oro contra el SP500 (10 años por defecto)
python analizador.py GLD

# Comparar Apple contra Microsoft
python analizador.py AAPL --benchmark MSFT

# Comparar varios ETFs contra el Nasdaq
python analizador.py SPY EFA EEM --benchmark ^IXIC

# Tu propio archivo Excel contra Bitcoin
python analizador.py mi_serie.xlsx --benchmark BTC-USD
```

Genera dos archivos:

- `<base>_analisis_<timestamp>.xlsx` — reporte con 4 hojas (Resumen, Precios alineados, Retornos diarios, NAV base 100).
- `<base>_dashboard_<timestamp>.html` — dashboard interactivo con 6 paneles (evolución base 100, drawdown, dispersión de retornos, volatilidad rolling, correlación rolling, beta rolling).

---

## Instalación

### Requisitos

- Python 3.10+ (recomendado vía [Anaconda](https://www.anaconda.com/download) o [Miniconda](https://docs.conda.io/en/latest/miniconda.html))
- Conexión a internet (solo si vas a usar tickers de Yahoo; los archivos locales no la necesitan)

### Setup

```bash
# 1. Clonar
git clone https://github.com/<tu_usuario>/<este_repo>.git
cd <este_repo>

# 2. (Opcional) entorno aislado
conda create -n analisis python=3.11 -y
conda activate analisis

# 3. Instalar dependencias
pip install -r requirements.txt
```

### Para los atajos `.bat` en Windows

Una sola vez, abre **Anaconda Prompt** y ejecuta:

```cmd
conda init cmd.exe
```

Cierra todas las ventanas. A partir de ahí los `.bat` funcionan con doble click.

---

## Uso del script Python

El primer argumento (y los siguientes) puede ser un **archivo local** o un **ticker de Yahoo Finance**. El script detecta automáticamente cuál es cuál.

### Casos de uso

```bash
# --- Solo tickers ---
python analizador.py GLD                            # GLD vs SP500
python analizador.py GLD TLT QQQ                    # cada uno vs SP500
python analizador.py GLD --benchmark QQQ            # GLD vs QQQ
python analizador.py AAPL MSFT --benchmark ^IXIC    # vs Nasdaq
python analizador.py BTC-USD --benchmark ETH-USD    # cripto vs cripto

# --- Atajos ---
python analizar_gld.py                              # GLD vs SP500, 10 años, sin parámetros

# --- Archivos locales ---
python analizador.py serie1.csv                     # archivo vs SP500
python analizador.py serie1.csv --benchmark TLT     # archivo vs TLT
python analizador.py serie1.csv serie2.xlsx GLD     # mezclar archivos y tickers

# --- Parámetros ---
python analizador.py GLD --rf 0.045                          # tasa libre de riesgo
python analizador.py GLD --start 2020-01-01 --end 2025-12-31 # periodo manual
python analizador.py GLD --out ./reportes                    # carpeta de salida
```

### Formatos de archivo aceptados

- `.xlsx`, `.xls`, `.xlsm`
- `.csv`, `.txt`, `.tsv` (separador autodetectado: `,` `;` `\t` `|`)

Estructura esperada: primera columna = fecha, columnas siguientes = precios.

---

## Convertir un Excel rellenado a HTML interactivo

```bash
python informe_html.py "Plantilla_Analisis_Comparativa.xlsx"
python informe_html.py reporte_previo.xlsx --rf 0.05
```

Reconoce automáticamente tres formatos:

| Caso | Hoja | Estructura |
|---|---|---|
| Plantilla Comparativa | `Datos` | Fecha · Activo 1 · Activo 2 (nombres editables en B2/C2) |
| Plantilla legacy vs SP500 | `Datos` | Fecha · Mi Serie · SP500 |
| Reporte de `analizador.py` | `Precios_alineados` | Fecha + N series + benchmark |

---

## Plantilla Excel (sin instalar Python)

Abre `Plantilla_Analisis_Comparativa.xlsx`:

1. Ve a la hoja **Datos**.
2. En las celdas **B2 y C2** escribe los nombres de los dos activos (ej. "GLD" y "SP500"). Se reflejan automáticamente en Dashboard y Gráfico.
3. En las columnas:
   - **A** → fechas
   - **B** → precios del Activo 1
   - **C** → precios del Activo 2 (benchmark)
4. Abre **Dashboard**: todas las métricas se calculan al instante con fórmulas nativas de Excel.
5. Abre **Gráfico**: evolución base 100 y drawdown comparado.

**Parámetros editables (Dashboard):**
- `B4` — tasa libre de riesgo anual (default 4%)
- `B5` — días bursátiles/año (default 252)

Soporta hasta **5,000 filas** de datos diarios (~20 años). Compatible con Excel 2010+ y LibreOffice Calc 7+.

---

## Métricas calculadas

Convención: 252 días bursátiles/año. Tasa libre de riesgo configurable.

| Métrica | Fórmula | Interpretación |
|---|---|---|
| **Beta** | `Cov(rA, rB) / Var(rB)` (excesos) | Sensibilidad del Activo 1 al Activo 2: >1 más volátil, <1 menos, =1 igual |
| **Alpha (Jensen)** | `(retA − rf) − β·(retB − rf)` anualizado | Sobre/sub-rendimiento ajustado por riesgo |
| **Correlación** | Pearson sobre retornos | Co-movimiento (−1 a +1) |
| **R²** | Correlación al cuadrado | Varianza explicada por el benchmark |
| **Volatilidad** | `σ(ret) · √252` | Desviación estándar anualizada |
| **Sharpe** | `(ret − rf) / σ · √252` | Retorno por unidad de riesgo total |
| **Sortino** | Sharpe con sólo volatilidad a la baja | Penaliza únicamente las caídas |
| **Tracking Error** | `σ(retA − retB) · √252` | Volatilidad del diferencial |
| **Information Ratio** | `(retA − retB)·252 / TE` | Retorno activo por unidad de TE |
| **Max Drawdown** | `min(Px / Px_max − 1)` | Mayor caída pico-valle |
| **CAGR** | `(Px_fin / Px_ini)^(252/n) − 1` | Tasa anual compuesta |
| **Up / Down Capture** | `mean(retA\|mkt↑) / mean(retB\|mkt↑)` | Captura de subidas/bajadas |

---

## Datos de ejemplo

`datos_ejemplo.xlsx` y `datos_ejemplo.csv` contienen 500 días bursátiles sintéticos generados con Beta objetivo = 1.30 y Alpha = 3% anual contra un benchmark con volatilidad 15%. Útiles para validar la herramienta antes de usarla con datos reales.

Resultado esperado con `rf = 4%`:

- Beta = 1.261, R² = 0.85, Correlación = 0.922
- Vol Activo 1 = 20.2%, Vol Activo 2 = 14.7%
- Sharpe = 0.78 / 0.37 · CAGR = 19.4% / 8.8%
- Tracking Error = 8.7%, Information Ratio = 1.18
- Max DD = −21.1% / −13.4%

Las cifras coinciden entre script Python y plantilla Excel (< 0.5% de diferencia por redondeo).

---

## Estructura del proyecto

```
.
├── analizador.py                      # script principal (archivos + tickers)
├── informe_html.py                    # Excel rellenado → HTML interactivo
├── analizar_gld.py                    # atajo: GLD vs SP500
├── build_template_v2.py               # genera la plantilla comparativa
│
├── correr.bat                         # doble click: GLD vs SP500 (10 años)
├── correr_personalizado.bat           # doble click: tickers vs SP500
├── comparar_dos_activos.bat           # doble click: compara dos activos cualesquiera
├── informe_desde_excel.bat            # doble click: Excel → HTML interactivo
│
├── Plantilla_Analisis_Comparativa.xlsx   # plantilla genérica (recomendada)
├── Plantilla_Analisis_vs_SP500.xlsx      # plantilla legacy (sigue funcionando)
│
├── datos_ejemplo.xlsx / .csv          # 500 días sintéticos para pruebas
├── requirements.txt                   # dependencias pip
└── README.md
```

---

## Tickers de Yahoo Finance soportados

Cualquier símbolo válido en Yahoo Finance funciona como argumento o benchmark:

- **Acciones US**: `AAPL`, `MSFT`, `TSLA`, `NVDA`
- **ETFs**: `GLD`, `TLT`, `QQQ`, `SPY`, `VTI`, `EEM`
- **Índices**: `^GSPC` (SP500), `^IXIC` (Nasdaq), `^DJI` (Dow Jones), `^STOXX50E` (Eurostoxx), `^MXX` (IPC México)
- **Fondos Morningstar/ISIN**: `0P0001DC0B.F`, `0P0001J4N8`, etc.
- **Divisas**: `EURUSD=X`, `MXN=X`, `JPY=X`
- **Commodities y cripto**: `GC=F` (oro futuros), `CL=F` (petróleo), `BTC-USD`, `ETH-USD`

---

## Notas técnicas y limitaciones

- **Frecuencia**: asume datos *diarios*. Para frecuencia semanal o mensual cambia `Días bursátiles/año` en la plantilla (52 ó 12) y/o la constante `PERIODOS_ANIO` en el script.
- **Fechas faltantes**: script y plantilla usan la intersección de fechas; no necesitas calendarios idénticos.
- **Monedas distintas**: convierte ambas a la misma divisa antes; Beta es robusta al nivel pero no al tipo de cambio.
- **Mínimo recomendado**: 60 observaciones para Beta confiable; 250+ para que el rolling 60d tenga sentido.
- **Compatibilidad Excel**: las fórmulas usan `LOOKUP`, `INDEX/MATCH` y `SUMPRODUCT`, compatibles con Excel 2010+ y LibreOffice Calc 7+.

---

## Stack técnico

- [pandas](https://pandas.pydata.org/) — manipulación de series temporales
- [numpy](https://numpy.org/) — cálculo de covarianzas y regresiones
- [yfinance](https://github.com/ranaroussi/yfinance) — descarga de precios de Yahoo Finance
- [openpyxl](https://openpyxl.readthedocs.io/) — generación de reportes Excel con fórmulas y gráficos nativos
- [plotly](https://plotly.com/python/) — dashboards HTML interactivos

---

## Licencia

MIT — úsalo, modifícalo, redistribúyelo libremente.

## Contribuciones

Pull requests y sugerencias bienvenidas. Issues con casos de uso interesantes (nuevos tickers, métricas adicionales, integración con otros data providers) especialmente apreciados.
