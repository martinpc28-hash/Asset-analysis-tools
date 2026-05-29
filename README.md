# Herramienta de Análisis Técnico Comparativo

Compara cualquier serie de tiempo (activo, fondo, ETF, índice, cripto, divisa…) contra
**cualquier otra** que uses como benchmark, y entrega
**Beta, Alpha de Jensen, Correlación, R², Volatilidad, Sharpe, Sortino,
Tracking Error, Information Ratio, Max Drawdown y CAGR**.

Aunque inicialmente la herramienta se construyó pensando en S&P 500, ahora es
**genérica**: el benchmark es configurable (default `^GSPC` por costumbre, pero
puede ser `GLD`, `QQQ`, `^MXX`, `BTC-USD`, otro fondo Morningstar, lo que sea).

## Tres formas de usarla

| Forma | Cuándo | Archivo |
|---|---|---|
| Atajos `.bat` (doble click) | Lo más rápido, sin escribir nada | `comparar_dos_activos.bat` · `correr.bat` · `correr_personalizado.bat` · `informe_desde_excel.bat` |
| Script Python | Máxima flexibilidad: cualquier ticker, archivo, periodo, benchmark | `analizador.py` · `informe_html.py` · `analizar_gld.py` |
| Plantilla Excel | Calculadora viva: pegas dos series y todo se calcula | `Plantilla_Analisis_Comparativa.xlsx` |

---

## 1. Atajos `.bat` (doble click)

Después de instalar dependencias una vez (sección 3), basta con doble click.

### `comparar_dos_activos.bat` — comparar dos cosas sin pensar
Doble click. Pregunta:
```
Activo 1 (el que quieres analizar):
Activo 2 (benchmark / referencia):
```
Escribes dos tickers de Yahoo (por ejemplo `GLD` y `QQQ`, o `0P0001DC0B.F` y `^GSPC`).
Genera Excel + HTML interactivo con todas las métricas.

### `correr.bat` — GLD vs SP500, 10 años, cero configuración
Doble click ejecuta el ejemplo clásico (GLD vs S&P 500) sin parámetros.

### `correr_personalizado.bat` — tickers contra SP500
Doble click; pregunta los tickers que quieras y los compara contra SP500.

### `informe_desde_excel.bat` — convierte un Excel en informe HTML
Doble click; arrastra cualquier Excel (la plantilla rellenada o un reporte
generado por `analizador.py`) y produce el dashboard HTML interactivo.

**Tickers válidos en cualquier `.bat`** — todo lo que admita Yahoo Finance:
- Acciones US: `AAPL`, `MSFT`, `TSLA`, `NVDA`
- ETFs: `GLD`, `TLT`, `QQQ`, `SPY`, `VTI`, `EEM`
- Índices: `^GSPC` (SP500), `^IXIC` (Nasdaq), `^DJI` (Dow), `^MXX` (IPC México), `^STOXX50E`
- Fondos Morningstar/ISIN: `0P0001DC0B.F`, `0P0001J4N8`, etc.
- Divisas: `MXN=X`, `EURUSD=X`
- Commodities y cripto: `GC=F` (oro futuros), `BTC-USD`, `ETH-USD`

---

## 2. Script Python — máxima flexibilidad

### Comparar dos activos directamente
```bash
python analizador.py GLD                            # GLD vs SP500 (default)
python analizador.py GLD --benchmark QQQ            # GLD vs QQQ
python analizador.py 0P0001DC0B.F --benchmark GLD   # fondo Morningstar vs oro
python analizador.py AAPL --benchmark MSFT          # acción vs acción
python analizador.py BTC-USD --benchmark ETH-USD    # cripto vs cripto
```

### Varios activos contra un mismo benchmark
```bash
python analizador.py GLD TLT QQQ                    # cada uno vs SP500
python analizador.py GLD TLT QQQ --benchmark ^IXIC  # cada uno vs Nasdaq
```

### Con archivos locales
```bash
python analizador.py mi_serie.xlsx                  # archivo vs SP500
python analizador.py mi_serie.xlsx --benchmark TLT  # archivo vs TLT
python analizador.py serie1.csv serie2.xlsx GLD     # mezclar archivos y tickers
```

### Parámetros comunes
```bash
python analizador.py GLD --rf 0.045                              # tasa libre de riesgo
python analizador.py GLD --start 2020-01-01 --end 2025-12-31     # periodo manual
python analizador.py GLD --out C:/Reportes                       # carpeta de salida
```

### Salidas
- `<base>_analisis_<timestamp>.xlsx` — reporte con 4 hojas: Resumen, Precios_alineados, Retornos_diarios, NAV_base100.
- `<base>_dashboard_<timestamp>.html` — dashboard Plotly con 6 paneles: Evolución base 100, Drawdown, Dispersión de retornos, Volatilidad rolling 60d, Correlación rolling 60d, Beta rolling 60d.

### Si solo tienes el Excel y quieres el HTML
```bash
python informe_html.py "Plantilla_Analisis_Comparativa.xlsx"
python informe_html.py "0P0001DC0B.F_analisis_20260522_1315.xlsx"
```
Reconoce automáticamente:
- La **plantilla nueva** (`Plantilla_Analisis_Comparativa.xlsx`).
- La **plantilla vieja** (`Plantilla_Analisis_vs_SP500.xlsx`).
- Cualquier **reporte de `analizador.py`** (hoja `Precios_alineados`).

---

## 3. Instalación en Anaconda (una sola vez)

1. Menú Inicio → **Anaconda Prompt** (NO el cmd normal).
2. Crea entorno (opcional pero recomendado):
   ```
   conda create -n analisis python=3.11 -y
   conda activate analisis
   ```
3. Posiciónate en la carpeta del proyecto y instala:
   ```
   cd "U:\ADD_AM\32_Addenda_Seleccion_Activos\Herramientas analisis tecnico"
   pip install -r requirements.txt
   ```
4. **Activa conda en cmd normal** (clave para que los `.bat` funcionen al doble click):
   ```
   conda init cmd.exe
   ```
   Cierra todas las ventanas. A partir de aquí, doble click a cualquier `.bat`.

### Errores típicos

| Error | Causa | Solución |
|---|---|---|
| `'conda' no se reconoce` | cmd no inicializado | Anaconda Prompt → `conda init cmd.exe` → cerrar todo |
| `ModuleNotFoundError: plotly` | Falta librería | `pip install plotly` en el entorno activo |
| `FileNotFoundError` | No estás en la carpeta correcta | `cd /d "U:\ADD_AM\...\Herramientas analisis tecnico"` |
| `No se pudo descargar ^GSPC` | Sin internet o firewall bloquea Yahoo | Verifica red; `pip install --upgrade yfinance` |
| `.bat` se cierra al instante | Conda no encuentra el entorno | Edita el `.bat` o reinstala dependencias en `base` |
| Python 3.14 de Microsoft Store en lugar de Anaconda | PATH equivocado | Usa la ruta completa: `"C:\Users\<tu>\anaconda3\python.exe" analizador.py ...` |

---

## 4. Plantilla Excel (sin instalar nada)

Abre `Plantilla_Analisis_Comparativa.xlsx`:

1. Ve a la hoja **Datos**.
2. En **B2 y C2** escribe los nombres de los dos activos (ej. "GLD" y "SP500", o "Fondo X" y "Indice Y"). Estos nombres se reflejan automáticamente en el Dashboard y el Gráfico.
3. En las columnas:
   - **A** → pega tus *fechas*.
   - **B** → pega los *precios del Activo 1*.
   - **C** → pega los *precios del Activo 2 (benchmark)*.
4. Abre **Dashboard**: todas las métricas calculadas.
5. Abre **Gráfico**: evolución base 100 y drawdown comparado.

**Parámetros editables (Dashboard):**
- `B4` — Tasa libre de riesgo anual (default 4%).
- `B5` — Días bursátiles/año (default 252).

**Convenciones de color:**
- Amarillo = inputs editables.
- Verde claro = KPI (no editar; son fórmulas).
- Rojo/Verde en Beta = resaltado si Beta > 1.1 o < 0.9.

Soporta hasta **5,000 filas** (~20 años de datos diarios).

La plantilla anterior (`Plantilla_Analisis_vs_SP500.xlsx`) sigue funcionando si la prefieres y el script la sigue leyendo.

---

## 5. Métricas calculadas — referencia rápida

Convención: 252 días bursátiles/año. Tasa libre de riesgo configurable.

| Métrica | Fórmula | Interpretación |
|---|---|---|
| **Beta** | `Cov(rA,rB) / Var(rB)` (excesos) | Sensibilidad del Activo 1 al Activo 2: >1 más volátil, <1 menos, =1 igual |
| **Alpha (Jensen)** | `(retA − rf) − β·(retB − rf)` anualizado | Sobre/sub-rendimiento ajustado por riesgo |
| **Correlación** | Pearson sobre retornos | Co-movimiento (−1 a +1) |
| **R²** | Correlación al cuadrado | % de la varianza del Activo 1 explicada por el Activo 2 |
| **Volatilidad** | `σ(ret) · √252` | Desviación estándar anualizada |
| **Sharpe** | `(ret − rf) / σ · √252` | Retorno por unidad de riesgo |
| **Sortino** | Sharpe con sólo volatilidad a la baja | Penaliza únicamente las caídas |
| **Tracking Error** | `σ(retA − retB) · √252` | Volatilidad del diferencial |
| **Information Ratio** | `(retA − retB)·252 / TE` | Retorno activo por unidad de TE |
| **Max Drawdown** | `min(Px / Px_max − 1)` | Mayor caída pico-valle |
| **CAGR** | `(Px_fin / Px_ini)^(252/n) − 1` | Tasa anual compuesta |
| **Up / Down Capture** | `mean(retA\|mkt↑) / mean(retB\|mkt↑)` | Captura de subidas/bajadas |

---

## 6. Datos de ejemplo

`datos_ejemplo.xlsx` y `datos_ejemplo.csv` contienen 500 días bursátiles
sintéticos con Beta objetivo = 1.30 y Alpha = 3% anual contra un benchmark
con vol 15%. Útiles para probar la herramienta. La plantilla rellenada con
estos datos produce:

- Beta = 1.261, R² = 0.85, Correlación = 0.922
- Vol Activo 1 = 20.2%, Vol Activo 2 = 14.7%
- Sharpe = 0.78 / 0.37 · CAGR = 19.4% / 8.8%
- Tracking Error = 8.7%, Information Ratio = 1.18
- Max DD = −21.1% / −13.4%

Las cifras coinciden entre script y plantilla (diferencias < 0.5% por redondeo).

---

## 7. Tips y limitaciones

- **Frecuencia**: la herramienta asume datos *diarios*. Para frecuencia semanal o mensual cambia `Días bursátiles/año` en Dashboard (52 o 12) y/o `PERIODOS_ANIO` en el script.
- **Fechas faltantes**: el script y la plantilla usan la intersección de fechas; no necesitas que ambas series tengan el mismo calendario exacto.
- **Monedas distintas**: convierte ambas a la misma divisa antes; Beta es robusta al nivel pero no al tipo de cambio.
- **Mínimo recomendado**: 60 observaciones para Beta confiable; 250+ para que el rolling 60d tenga sentido.
- **LibreOffice / Excel viejo**: las fórmulas usan `LOOKUP`, `INDEX/MATCH` y `SUMPRODUCT`, compatibles con Excel 2010+ y LibreOffice Calc 7+.

---

## 8. Mapa de archivos

```
Herramientas analisis tecnico/
├── comparar_dos_activos.bat               ← doble click: compara dos activos cualesquiera
├── correr.bat                              ← doble click: GLD vs SP500 (10 años)
├── correr_personalizado.bat                ← doble click: tickers vs SP500
├── informe_desde_excel.bat                 ← doble click: Excel → HTML interactivo
│
├── analizador.py                           ← script principal (archivos + tickers)
├── analizar_gld.py                         ← atajo GLD vs SP500
├── informe_html.py                         ← Excel rellenado → HTML
├── build_template_v2.py                    ← genera la plantilla comparativa
│
├── Plantilla_Analisis_Comparativa.xlsx     ← plantilla genérica (la nueva)
├── Plantilla_Analisis_vs_SP500.xlsx        ← plantilla antigua (sigue funcionando)
│
├── requirements.txt                        ← dependencias pip
├── datos_ejemplo.xlsx / .csv               ← 500 días sintéticos para pruebas
└── README.md                               ← este archivo
```
