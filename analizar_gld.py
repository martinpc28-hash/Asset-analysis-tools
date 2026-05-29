"""
analizar_gld.py
================
Atajo: analiza GLD (SPDR Gold Shares) vs SP500 en los ultimos 10 anios
sin necesidad de archivos de entrada. Descarga todo desde Yahoo Finance.

Uso:
    python analizar_gld.py

Equivalente a:
    python analizador.py GLD --benchmark ^GSPC --start <hoy-10y>
"""
from datetime import datetime
import pandas as pd
import sys

# Reusa todo el motor de calculo del script principal
from analizador import (
    cargar_benchmark,
    alinear,
    calcular_metricas,
    exportar_excel,
    exportar_html,
)

TICKER_SERIE = "GLD"
TICKER_BENCH = "^GSPC"
LOOKBACK_ANIOS = 10
RF_ANUAL = 0.045  # 4.5%


def main():
    end = datetime.now()
    start = end - pd.DateOffset(years=LOOKBACK_ANIOS)
    s_start = start.strftime("%Y-%m-%d")
    s_end = end.strftime("%Y-%m-%d")

    print("\n=== ANALIZADOR  GLD vs SP500  (10 anios) ===")
    print(f"Periodo: {s_start}  ->  {s_end}")
    print(f"Tasa libre de riesgo: {RF_ANUAL*100:.2f}%\n")

    # 1) Descarga ambas series desde Yahoo
    serie = cargar_benchmark(TICKER_SERIE, s_start, s_end).to_frame()
    bench = cargar_benchmark(TICKER_BENCH, s_start, s_end)

    # 2) Alinea por fechas comunes
    combinado = alinear(serie, bench)
    print(f"  -> Datos alineados: {combinado.shape[0]} filas comunes "
          f"({combinado.index.min().date()}  ->  {combinado.index.max().date()})\n")

    # 3) Calcula metricas
    bench_col = bench.name
    resultados = {}
    for c in [col for col in combinado.columns if col != bench_col]:
        resultados[c] = calcular_metricas(combinado[c], combinado[bench_col], RF_ANUAL)

    # 4) Resumen en pantalla
    print("=== RESUMEN ===")
    for nombre, m in resultados.items():
        print(f"\n[{nombre}]")
        for k, v in m.items():
            if isinstance(v, float):
                print(f"  {k:.<35s} {v:>10.4f}")
            else:
                print(f"  {k:.<35s} {v}")

    # 5) Exporta
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    excel_out = f"GLD_vs_SP500_{stamp}.xlsx"
    html_out = f"GLD_vs_SP500_{stamp}.html"
    exportar_excel(combinado, resultados, excel_out)
    exportar_html(combinado, resultados, bench_col, html_out)

    print(f"\n[OK] Listo. Archivos generados:")
    print(f"  - {excel_out}")
    print(f"  - {html_out}")


if __name__ == "__main__":
    main()
