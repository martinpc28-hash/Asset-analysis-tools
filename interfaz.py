"""
interfaz.py
===========
Interfaz grafica (Tkinter) para el Analizador Comparativo de Activos.

Modos:
  1. Solo tickers Yahoo
  2. Solo archivo Excel/CSV
  3. Combinado: archivo(s) + tickers

Si el campo Benchmark esta VACIO -> modo STANDALONE: compara los activos
entre si (NAV base 100, drawdown, vol rolling, matriz de correlacion y
tabla comparativa) sin necesidad de un benchmark externo.

Si tiene valor -> modo COMPARATIVO: usa el benchmark indicado y calcula
Beta, Alpha, R^2, etc. ademas del dashboard interactivo completo.
"""
import os
import sys
import threading
import traceback
from datetime import datetime, timedelta
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from analizador import (
    cargar_serie, cargar_benchmark, alinear,
    calcular_metricas, exportar_excel, exportar_html,
    es_archivo,
)

COLOR_HDR = "#1F4E78"
COLOR_HDR2 = "#CCE0F0"


class AnalizadorUI:
    def __init__(self, root):
        self.root = root
        root.title("Analizador Comparativo de Activos")
        root.geometry("780x740")
        root.minsize(720, 660)
        root.configure(bg="#FFFFFF")
        self._build_ui()

    # ===================================================================
    # LAYOUT
    # ===================================================================
    def _build_ui(self):
        header = tk.Frame(self.root, bg=COLOR_HDR, padx=20, pady=12)
        header.pack(fill="x")
        tk.Label(header, text="Analizador Comparativo de Activos",
                 bg=COLOR_HDR, fg="white",
                 font=("Arial", 16, "bold")).pack(anchor="w")
        tk.Label(header,
                 text="Compara activos vs un benchmark, o entre si si dejas el "
                      "benchmark vacio. Acepta tickers Yahoo y archivos locales.",
                 bg=COLOR_HDR, fg=COLOR_HDR2,
                 font=("Arial", 9)).pack(anchor="w")

        content = tk.Frame(self.root, bg="#FFFFFF", padx=20, pady=12)
        content.pack(fill="both", expand=True)

        # ----- Modo -----
        modo_frame = tk.LabelFrame(content, text=" Modo de analisis ",
                                   padx=10, pady=6, bg="#FFFFFF",
                                   font=("Arial", 10, "bold"))
        modo_frame.pack(fill="x", pady=(0, 10))
        self.modo = tk.StringVar(value="tickers")
        for val, label in [
            ("tickers", "Solo tickers Yahoo  (GLD, AAPL, ^GSPC, BTC-USD, etc.)"),
            ("archivo", "Solo archivo Excel/CSV"),
            ("mixto",   "Combinado: archivo(s) + tickers"),
        ]:
            tk.Radiobutton(modo_frame, text=label, variable=self.modo, value=val,
                           font=("Arial", 10), bg="#FFFFFF",
                           command=self._on_modo_changed).pack(anchor="w")

        # ----- Inputs -----
        ipt = tk.LabelFrame(content, text=" Configuracion ",
                            padx=10, pady=10, bg="#FFFFFF",
                            font=("Arial", 10, "bold"))
        ipt.pack(fill="x", pady=(0, 10))

        tk.Label(ipt, text="Activos a analizar:", bg="#FFFFFF",
                 font=("Arial", 10)).grid(row=0, column=0, sticky="w", pady=3)
        self.activos_entry = tk.Entry(ipt, font=("Consolas", 10))
        self.activos_entry.grid(row=0, column=1, columnspan=2, sticky="ew", padx=8, pady=3)
        tk.Label(ipt, text="separados por espacio, ej:  GLD TLT QQQ",
                 bg="#FFFFFF", fg="#666",
                 font=("Arial", 8, "italic")).grid(row=1, column=1, sticky="w", padx=8)

        tk.Label(ipt, text="Archivo:", bg="#FFFFFF",
                 font=("Arial", 10)).grid(row=2, column=0, sticky="w", pady=3)
        self.archivo_entry = tk.Entry(ipt, font=("Consolas", 9))
        self.archivo_entry.grid(row=2, column=1, sticky="ew", padx=8, pady=3)
        self.archivo_btn = tk.Button(ipt, text="Examinar...",
                                     command=self._browse_file, font=("Arial", 9))
        self.archivo_btn.grid(row=2, column=2, padx=4)

        tk.Label(ipt, text="Benchmark:", bg="#FFFFFF",
                 font=("Arial", 10)).grid(row=3, column=0, sticky="w", pady=3)
        self.bench_entry = tk.Entry(ipt, font=("Consolas", 10))
        self.bench_entry.insert(0, "^GSPC")
        self.bench_entry.grid(row=3, column=1, columnspan=2, sticky="ew", padx=8, pady=3)
        tk.Label(ipt, text="ej: ^GSPC, SPY, QQQ, GLD, BALANCED   |   "
                           "VACIO = comparar entre los activos sin benchmark",
                 bg="#FFFFFF", fg="#666",
                 font=("Arial", 8, "italic")).grid(row=4, column=1, sticky="w", padx=8)

        date_row = tk.Frame(ipt, bg="#FFFFFF")
        date_row.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(8, 3))
        tk.Label(date_row, text="Fecha inicio:", bg="#FFFFFF",
                 font=("Arial", 10)).pack(side="left")
        self.start_entry = tk.Entry(date_row, font=("Consolas", 10), width=12)
        self.start_entry.insert(0,
            (datetime.now() - timedelta(days=365 * 10)).strftime("%Y-%m-%d"))
        self.start_entry.pack(side="left", padx=(8, 20))
        tk.Label(date_row, text="Fecha fin:", bg="#FFFFFF",
                 font=("Arial", 10)).pack(side="left")
        self.end_entry = tk.Entry(date_row, font=("Consolas", 10), width=12)
        self.end_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.end_entry.pack(side="left", padx=(8, 20))
        tk.Label(date_row, text="Tasa libre de riesgo (%):", bg="#FFFFFF",
                 font=("Arial", 10)).pack(side="left")
        self.rf_entry = tk.Entry(date_row, font=("Consolas", 10), width=6)
        self.rf_entry.insert(0, "4.5")
        self.rf_entry.pack(side="left", padx=(8, 0))

        tk.Label(ipt, text="Carpeta de salida:", bg="#FFFFFF",
                 font=("Arial", 10)).grid(row=6, column=0, sticky="w", pady=3)
        self.out_entry = tk.Entry(ipt, font=("Consolas", 9))
        self.out_entry.insert(0, str(Path(__file__).parent))
        self.out_entry.grid(row=6, column=1, sticky="ew", padx=8, pady=3)
        tk.Button(ipt, text="Examinar...", command=self._browse_dir,
                  font=("Arial", 9)).grid(row=6, column=2, padx=4)
        ipt.columnconfigure(1, weight=1)

        # ----- Botones -----
        btn_frame = tk.Frame(content, bg="#FFFFFF")
        btn_frame.pack(fill="x", pady=(0, 10))
        self.run_btn = tk.Button(btn_frame, text="Ejecutar analisis",
                                 bg=COLOR_HDR, fg="white",
                                 font=("Arial", 11, "bold"), padx=22, pady=8,
                                 command=self._run_analysis_threaded,
                                 activebackground="#2E75B6", activeforeground="white")
        self.run_btn.pack(side="left", padx=(0, 8))
        tk.Button(btn_frame, text="Abrir carpeta de salida",
                  font=("Arial", 10), padx=15, pady=8,
                  command=self._open_out_folder).pack(side="left", padx=(0, 8))
        tk.Button(btn_frame, text="Limpiar log",
                  font=("Arial", 10), padx=15, pady=8,
                  command=lambda: self.log.delete("1.0", "end")).pack(side="left")

        # ----- Log -----
        log_frame = tk.LabelFrame(content, text=" Registro de ejecucion ",
                                   padx=5, pady=5, bg="#FFFFFF",
                                   font=("Arial", 10, "bold"))
        log_frame.pack(fill="both", expand=True)
        self.log = scrolledtext.ScrolledText(log_frame, font=("Consolas", 9),
                                             bg="#F8F8F8", fg="#222", height=12,
                                             wrap="word")
        self.log.pack(fill="both", expand=True)
        self._log("Listo. Configura los parametros y pulsa 'Ejecutar analisis'.\n")
        self._on_modo_changed()

    # ===================================================================
    # HANDLERS
    # ===================================================================
    def _on_modo_changed(self):
        state = "disabled" if self.modo.get() == "tickers" else "normal"
        self.archivo_entry.config(state=state)
        self.archivo_btn.config(state=state)

    def _browse_file(self):
        f = filedialog.askopenfilename(
            title="Selecciona el archivo de datos",
            filetypes=[("Excel y CSV", "*.xlsx *.xls *.csv *.txt"),
                       ("Todos", "*.*")])
        if f:
            self.archivo_entry.delete(0, "end")
            self.archivo_entry.insert(0, f)

    def _browse_dir(self):
        d = filedialog.askdirectory(title="Carpeta de salida")
        if d:
            self.out_entry.delete(0, "end")
            self.out_entry.insert(0, d)

    def _open_out_folder(self):
        d = self.out_entry.get().strip()
        if d and os.path.isdir(d):
            os.startfile(d)
        else:
            messagebox.showwarning("Aviso", "La carpeta no existe todavia.")

    def _log(self, msg):
        self.log.insert("end", msg)
        self.log.see("end")
        self.log.update_idletasks()

    def _run_analysis_threaded(self):
        threading.Thread(target=self._do_analysis, daemon=True).start()

    # ===================================================================
    # ANALYSIS LOGIC
    # ===================================================================
    def _do_analysis(self):
        self.run_btn.config(state="disabled", text="Procesando...")
        try:
            modo = self.modo.get()
            activos_str = self.activos_entry.get().strip()
            archivo = self.archivo_entry.get().strip()
            bench = self.bench_entry.get().strip()  # vacio => standalone
            start = self.start_entry.get().strip()
            end = self.end_entry.get().strip()
            try:
                rf = float(self.rf_entry.get()) / 100.0
            except ValueError:
                raise ValueError("La tasa libre de riesgo debe ser un numero (ej: 4.5).")
            out_dir = self.out_entry.get().strip() or str(Path(__file__).parent)

            fuentes = []
            if modo in ("archivo", "mixto"):
                if not archivo:
                    raise ValueError("Selecciona un archivo para este modo.")
                if not os.path.exists(archivo):
                    raise FileNotFoundError(f"No existe: {archivo}")
                fuentes.append(archivo)
            if modo in ("tickers", "mixto") and activos_str:
                fuentes.extend(activos_str.split())
            if not fuentes:
                raise ValueError("No has indicado activos ni archivo.")

            standalone = (bench == "")
            self._log("\n=== ANALIZADOR " + ("STANDALONE" if standalone else "COMPARATIVO") + " ===\n")
            self._log(f"Modo:        {modo}\n")
            self._log(f"Fuentes:     {fuentes}\n")
            self._log(f"Benchmark:   {bench if bench else '(vacio -> comparar entre activos)'}\n")
            self._log(f"Periodo:     {start}  ->  {end}\n")
            self._log(f"Rf anual:    {rf*100:.2f}%\n\n")

            # Carga cada fuente
            dataframes = []
            for f in fuentes:
                if es_archivo(f):
                    self._log(f"-> Cargando archivo: {f}\n")
                    df_f = cargar_serie(f)
                else:
                    self._log(f"-> Descargando ticker Yahoo: {f}\n")
                    df_f = cargar_benchmark(f, start, end).to_frame()
                dataframes.append(df_f)
                self._log(f"   {df_f.shape[0]} filas, columnas: {list(df_f.columns)}\n")

            df_user = pd.concat(dataframes, axis=1, join="outer").sort_index()

            out_path = Path(out_dir)
            out_path.mkdir(parents=True, exist_ok=True)
            primera = fuentes[0]
            base = Path(primera).stem if es_archivo(primera) else primera.replace("^", "")
            stamp = datetime.now().strftime("%Y%m%d_%H%M")

            # ============================================================
            # MODO STANDALONE
            # ============================================================
            if standalone:
                df_clean = df_user.dropna(how="any")
                if df_clean.shape[0] < 5:
                    raise ValueError("Datos insuficientes despues de alinear fechas.")
                self._log(f"-> Datos alineados: {df_clean.shape[0]} filas "
                          f"({df_clean.index.min().date()} -> {df_clean.index.max().date()})\n\n")

                P = 252
                resumen = {}
                for c in df_clean.columns:
                    r = df_clean[c].pct_change().dropna()
                    rf_diario = (1 + rf) ** (1 / P) - 1
                    vol = r.std(ddof=1) * np.sqrt(P)
                    sharpe = ((r.mean() - rf_diario) / r.std(ddof=1) * np.sqrt(P)
                              if r.std(ddof=1) > 0 else np.nan)
                    down = np.minimum(r - rf_diario, 0)
                    dd_dev = np.sqrt((down ** 2).mean()) * np.sqrt(P)
                    sortino = ((r.mean() - rf_diario) * P / dd_dev
                               if dd_dev > 0 else np.nan)
                    nav = (1 + r).prod() - 1
                    n_anios = len(r) / P
                    cagr = (1 + nav) ** (1 / n_anios) - 1 if n_anios > 0 else np.nan
                    eq = (1 + r).cumprod()
                    mdd = (eq / eq.cummax() - 1).min()
                    resumen[c] = {
                        "Observaciones": len(r),
                        "Retorno acumulado": nav,
                        "CAGR (anualizado)": cagr,
                        "Volatilidad anual": vol,
                        "Sharpe": sharpe,
                        "Sortino": sortino,
                        "Max Drawdown": mdd,
                    }
                    self._log(
                        f"-> {c}: CAGR={cagr*100:.2f}%  Vol={vol*100:.2f}%  "
                        f"Sharpe={sharpe:.3f}  Sortino={sortino:.3f}  "
                        f"MaxDD={mdd*100:.2f}%\n"
                    )

                rets_df = df_clean.pct_change().dropna()
                corr_matrix = rets_df.corr()
                self._log("\n--- Correlacion entre activos ---\n")
                self._log(corr_matrix.round(3).to_string() + "\n")

                tabla = pd.DataFrame(resumen)  # metricas en filas, activos en cols

                excel_out = out_path / f"{base}_standalone_{stamp}.xlsx"
                with pd.ExcelWriter(str(excel_out), engine="openpyxl") as xw:
                    tabla.to_excel(xw, sheet_name="Comparativa")
                    corr_matrix.to_excel(xw, sheet_name="Correlacion")
                    df_clean.to_excel(xw, sheet_name="Precios")
                    rets_df.to_excel(xw, sheet_name="Retornos_diarios")
                self._log(f"\n[OK] Excel: {excel_out.name}\n")

                html_out = out_path / f"{base}_standalone_{stamp}.html"
                self._generar_html_standalone(df_clean, rets_df, resumen,
                                              corr_matrix, str(html_out))
                self._log(f"[OK] HTML: {html_out.name}\n")

                if messagebox.askyesno(
                    "Analisis standalone completado",
                    f"Comparativa entre {len(df_clean.columns)} activos generada:\n"
                    f"  - {excel_out.name}\n  - {html_out.name}\n\n"
                    f"¿Abrir el dashboard HTML ahora?",
                ):
                    os.startfile(str(html_out))
                return

            # ============================================================
            # MODO COMPARATIVO (con benchmark)
            # ============================================================
            self._log(f"\n-> Descargando benchmark: {bench}\n")
            bench_series = cargar_benchmark(bench, start, end)
            combinado = alinear(df_user, bench_series)
            self._log(f"-> Datos alineados: {combinado.shape[0]} filas "
                      f"({combinado.index.min().date()} -> {combinado.index.max().date()})\n\n")

            benchmark_col = bench_series.name
            resultados = {}
            for c in [col for col in combinado.columns if col != benchmark_col]:
                m = calcular_metricas(combinado[c], combinado[benchmark_col], rf)
                resultados[c] = m
                self._log(
                    f"-> {c}: Beta={m['Beta']:.3f}  R2={m['R_cuadrado']:.3f}  "
                    f"Sharpe={m['Sharpe_serie']:.3f}  Sortino={m['Sortino_serie']:.3f}\n"
                )

            excel_out = out_path / f"{base}_analisis_{stamp}.xlsx"
            html_out = out_path / f"{base}_dashboard_{stamp}.html"
            self._log(f"\n-> Generando reporte Excel...\n")
            exportar_excel(combinado, resultados, str(excel_out))
            self._log(f"-> Generando dashboard HTML...\n")
            exportar_html(combinado, resultados, benchmark_col, str(html_out))
            self._log(f"\n[OK] Generados:\n  - {excel_out.name}\n  - {html_out.name}\n")

            if messagebox.askyesno(
                "Analisis completado",
                f"Reportes generados:\n  - {excel_out.name}\n  - {html_out.name}\n\n"
                f"¿Abrir el dashboard HTML ahora?",
            ):
                os.startfile(str(html_out))

        except Exception as e:
            self._log(f"\n[ERROR] {e}\n{traceback.format_exc()}")
            messagebox.showerror("Error", str(e))
        finally:
            self.run_btn.config(state="normal", text="Ejecutar analisis")

    # ===================================================================
    # HTML STANDALONE COMPARATIVO
    # ===================================================================
    def _generar_html_standalone(self, df_clean, rets_df, resumen,
                                  corr_matrix, salida):
        """Genera HTML comparativo entre activos sin benchmark.

        Incluye: NAV base 100, drawdown, vol rolling, heatmap de correlacion
        y tabla comparativa de metricas (activos como columnas).
        """
        try:
            import plotly.graph_objects as go
            from plotly.subplots import make_subplots
        except ImportError:
            self._log("[AVISO] plotly no instalado; omito HTML.\n")
            return

        activos = list(df_clean.columns)
        PALETA = ["#1F4E78", "#C00000", "#00B050", "#7030A0", "#ED7D31", "#2E75B6"]
        color_de = {c: PALETA[i % len(PALETA)] for i, c in enumerate(activos)}

        nav = df_clean.divide(df_clean.iloc[0]).multiply(100)
        P = 252
        titulo = f"Comparativa entre {len(activos)} activos: " + " · ".join(activos)

        fig = make_subplots(
            rows=2, cols=2,
            subplot_titles=("Evolucion Base 100", "Drawdown (%)",
                            "Volatilidad rolling 60d (anualizada)",
                            "Matriz de correlacion"),
            specs=[[{}, {}], [{}, {"type": "heatmap"}]],
            vertical_spacing=0.14, horizontal_spacing=0.12,
        )

        for c in activos:
            fig.add_trace(go.Scatter(
                x=nav.index, y=nav[c], name=c, legendgroup=c,
                mode="lines", line=dict(color=color_de[c], width=2),
            ), row=1, col=1)

        for c in activos:
            cum = df_clean[c] / df_clean[c].iloc[0]
            dd = (cum / cum.cummax() - 1) * 100
            fig.add_trace(go.Scatter(
                x=dd.index, y=dd, name=c, legendgroup=c,
                mode="lines", fill="tozeroy", opacity=0.4,
                line=dict(color=color_de[c]), showlegend=False,
            ), row=1, col=2)

        vol = rets_df.rolling(60).std() * np.sqrt(P) * 100
        for c in activos:
            fig.add_trace(go.Scatter(
                x=vol.index, y=vol[c], name=c, legendgroup=c,
                mode="lines", line=dict(color=color_de[c]),
                showlegend=False,
            ), row=2, col=1)

        fig.add_trace(go.Heatmap(
            z=corr_matrix.values,
            x=corr_matrix.columns.tolist(), y=corr_matrix.index.tolist(),
            colorscale="RdBu_r", zmid=0, zmin=-1, zmax=1,
            text=corr_matrix.round(2).values, texttemplate="%{text}",
            textfont=dict(size=11), showscale=True,
            colorbar=dict(thickness=10, len=0.4, y=0.22, x=1.02),
        ), row=2, col=2)

        fig.update_layout(
            height=900, template="plotly_white", hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02,
                        xanchor="left", x=0, font=dict(size=10)),
            margin=dict(t=60, b=40, l=60, r=40),
        )
        fig.update_yaxes(title_text="Base 100", row=1, col=1)
        fig.update_yaxes(title_text="Drawdown (%)", row=1, col=2)
        fig.update_yaxes(title_text="Vol anual (%)", row=2, col=1)

        pct_metrics = {"Retorno acumulado", "CAGR (anualizado)",
                       "Volatilidad anual", "Max Drawdown"}
        metricas_keys = list(next(iter(resumen.values())).keys())

        tabla = "<h2 style='font-family:Arial;color:#1F4E78;margin-top:30px'>Tabla comparativa</h2>"
        tabla += ("<table style='border-collapse:collapse;font-family:Arial;"
                  "font-size:13px;border:1px solid #BFBFBF'>")
        tabla += "<tr style='background:#1F4E78;color:white'>"
        tabla += "<th style='padding:8px 14px;text-align:left'>Metrica</th>"
        for a in activos:
            tabla += f"<th style='padding:8px 14px;text-align:center'>{a}</th>"
        tabla += "</tr>"
        for k in metricas_keys:
            tabla += ("<tr><td style='padding:6px 14px;background:#F2F2F2;"
                      f"font-weight:bold;border-bottom:1px solid #E0E0E0'>{k}</td>")
            for a in activos:
                v = resumen[a].get(k, "")
                if isinstance(v, (int, float)) and not pd.isna(v):
                    if k in pct_metrics:
                        txt = f"{v*100:,.2f}%"
                    elif k == "Observaciones":
                        txt = f"{int(v):,}"
                    else:
                        txt = f"{v:,.3f}"
                else:
                    txt = str(v)
                tabla += ("<td style='padding:6px 14px;text-align:center;"
                          f"border-bottom:1px solid #E0E0E0'>{txt}</td>")
            tabla += "</tr>"
        tabla += "</table>"

        html_chart = fig.to_html(include_plotlyjs="cdn", full_html=False)
        full = (
            "<!doctype html><html><head><meta charset='utf-8'>"
            f"<title>{titulo}</title></head>"
            "<body style='font-family:Arial;margin:20px;color:#222'>"
            f"<h1 style='color:#1F4E78;margin-bottom:4px'>{titulo}</h1>"
            f"<p style='color:#888;font-size:12px;margin-top:0'>"
            f"Generado el {datetime.now().strftime('%Y-%m-%d %H:%M')} — modo standalone</p>"
            f"{html_chart}{tabla}"
            "</body></html>"
        )
        with open(salida, "w", encoding="utf-8") as f:
            f.write(full)


def main():
    root = tk.Tk()
    AnalizadorUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
