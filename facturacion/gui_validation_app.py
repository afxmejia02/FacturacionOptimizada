"""
Aplicación GUI para validación de perfiles, servicios y equipos.

La aplicación compara información extraída desde PDF contra un Excel histórico
y muestra solo las diferencias. Las filas con discrepancias se resaltan en rojo.
"""

from __future__ import annotations

import datetime as dt
import os
import re
import threading
import traceback
import unicodedata
from collections import Counter

import pandas as pd
import pdfplumber

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except Exception:  # pragma: no cover - allows headless import on Streamlit Cloud
    tk = None
    filedialog = None
    messagebox = None
    ttk = None


class ServicesValidationApp:
    """Interfaz principal para validar perfiles, servicios y equipos."""

    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Validación de Perfiles, Servicios y Equipos")
        self.root.geometry("1400x800")
        self.root.configure(bg="#f0f0f0")

        self.excel_path = tk.StringVar()
        self.pdf_path = tk.StringVar()
        self.tipo_extraccion = tk.StringVar(value="equipos")
        self.df_resultado = None
        self.debug_mode = os.environ.get("VALIDATION_DEBUG", "1") == "1"

        self._create_ui()

    def _debug_print(self, message):
        if self.debug_mode:
            print(f"[DEBUG][ServicesValidationApp] {message}")

    def _create_ui(self):
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Label(
            header_frame,
            text="Sistema de Validación de Perfiles, Servicios y Equipos",
            font=("Arial", 16, "bold"),
        ).pack(side=tk.LEFT)

        type_frame = ttk.LabelFrame(self.root, text="Seleccionar tipo de datos", padding=10)
        type_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(type_frame, text="¿Qué deseas validar?").pack(side=tk.LEFT, padx=5)
        for tipo in ("perfiles", "equipos", "servicios"):
            ttk.Radiobutton(
                type_frame,
                text=tipo.upper(),
                variable=self.tipo_extraccion,
                value=tipo,
            ).pack(side=tk.LEFT, padx=10)

        path_frame = ttk.LabelFrame(self.root, text="Seleccionar archivos", padding=10)
        path_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(path_frame, text="Informe PDF:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(path_frame, textvariable=self.pdf_path, width=80).grid(row=0, column=1, padx=5)
        ttk.Button(path_frame, text="Buscar", command=self._select_pdf_file).grid(row=0, column=2)

        ttk.Label(path_frame, text="Datos históricos de Excel:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(path_frame, textvariable=self.excel_path, width=80).grid(row=1, column=1, padx=5)
        ttk.Button(path_frame, text="Buscar", command=self._select_excel_file).grid(row=1, column=2)

        control_frame = ttk.Frame(self.root, padding=10)
        control_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(control_frame, text="Validar archivos", command=self._process_files, width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Exportar a CSV", command=self._export_to_csv, width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Limpiar resultados", command=self._clear_results, width=20).pack(side=tk.LEFT, padx=5)

        self.status_label = ttk.Label(control_frame, text="Listo", foreground="blue")
        self.status_label.pack(side=tk.LEFT, padx=20)

        results_frame = ttk.LabelFrame(self.root, text="Resultados", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.tree = ttk.Treeview(
            results_frame,
            columns=("Fecha", "Servicio", "PDF", "Excel", "Diferencia"),
            height=20,
            show="headings",
        )
        self.tree.heading("Fecha", text="Fecha")
        self.tree.heading("Servicio", text="Descripción")
        self.tree.heading("PDF", text="Cantidad PDF")
        self.tree.heading("Excel", text="Cantidad Excel")
        self.tree.heading("Diferencia", text="Diferencia")

        self.tree.column("Fecha", width=120, anchor=tk.CENTER)
        self.tree.column("Servicio", width=400, anchor=tk.W)
        self.tree.column("PDF", width=140, anchor=tk.CENTER)
        self.tree.column("Excel", width=140, anchor=tk.CENTER)
        self.tree.column("Diferencia", width=140, anchor=tk.CENTER)

        vsb = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(results_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscroll=vsb.set, xscroll=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        


        results_frame.grid_rowconfigure(0, weight=1)
        results_frame.grid_columnconfigure(0, weight=1)

        self.tree.tag_configure("error", background="#ffcccc", foreground="darkred")
        self.tree.tag_configure("ok", background="#ccffcc", foreground="darkgreen")

    def _select_pdf_file(self):
        file = filedialog.askopenfilename(
            title="Selecciona el informe PDF",
            filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")],
        )
        if file:
            self.pdf_path.set(file)

    def _select_excel_file(self):
        file = filedialog.askopenfilename(
            title="Selecciona los datos históricos de Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")],
        )
        if file:
            self.excel_path.set(file)

    def _process_files(self):
        if not self.pdf_path.get():
            messagebox.showerror("Error", "Por favor selecciona un archivo PDF")
            return
        if not self.excel_path.get():
            messagebox.showerror("Error", "Por favor selecciona un archivo de Excel")
            return

        self.status_label.config(text="Procesando... por favor espera.", foreground="orange")
        self.root.update()

        thread = threading.Thread(target=self._process_files_thread, daemon=True)
        thread.start()

    def _process_files_thread(self):
        try:
            tipo = self.tipo_extraccion.get().lower()
            self._debug_print(f"Inicio de procesamiento. tipo={tipo}, pdf={self.pdf_path.get()}, excel={self.excel_path.get()}")

            if tipo == "perfiles":
                conteo_pdf, fecha_reporte = self._extraer_perfiles_pdf(self.pdf_path.get())
                self._debug_print(f"Perfiles extraidos del PDF: {len(conteo_pdf)} perfiles, fecha_reporte={fecha_reporte}")
                if not conteo_pdf:
                    self.root.after(
                        0,
                        lambda: messagebox.showerror(
                            "Error",
                            "No se encontraron perfiles en el PDF.\nRevisa la estructura del archivo.",
                        ),
                    )
                    self.root.after(0, lambda: self.status_label.config(text="Procesamiento fallido", foreground="red"))
                    return

                conteo_excel = self._extraer_conteo_excel_perfiles(self.excel_path.get(), fecha_reporte)
                self._debug_print(f"Perfiles extraidos del Excel para fecha {fecha_reporte}: {len(conteo_excel)}")
                self.df_resultado = self._comparar_conteos_perfiles(conteo_pdf, conteo_excel)
                self._debug_print(f"Diferencias encontradas (perfiles): {len(self.df_resultado)}")

                if self.df_resultado.empty:
                    self.root.after(0, lambda: self._show_custom_ok_message(
                        "Validación correcta ✅",
                        "No se encontraron diferencias entre el PDF y el Excel.",
                    ))
                    self.root.after(0, lambda: self.status_label.config(text="Validación completa. ¡Todo correcto!", foreground="green"))
                else:
                    self.root.after(0, self._display_profile_results)
                    self.root.after(0, lambda: self.status_label.config(
                        text=f"Validación completa. Se encontraron {len(self.df_resultado)} diferencias.",
                        foreground="red",
                    ))
                return

            df_pdf = self._extraer_conteo_pdf(self.pdf_path.get(), tipo)
            self._debug_print(f"Registros PDF extraidos para tipo {tipo}: {0 if df_pdf is None else len(df_pdf)}")
            if df_pdf is None or df_pdf.empty:
                self.root.after(
                    0,
                    lambda: messagebox.showerror(
                        "Error",
                        f"No se encontraron elementos válidos de tipo {tipo} en el PDF.",
                    ),
                )
                self.root.after(0, lambda: self.status_label.config(text="Procesamiento fallido", foreground="red"))
                return

            self.df_resultado = self._comparar_conteos(df_pdf, self.excel_path.get())
            self._debug_print(f"Diferencias encontradas ({tipo}): {len(self.df_resultado)}")
            if self.df_resultado.empty or self.df_resultado["Diferencia"].sum() == 0:
                self.root.after(0, self._show_all_ok_message)
                self.root.after(0, lambda: self.status_label.config(text="Validación completa. ¡Todo correcto!", foreground="green"))
            else:
                self.root.after(0, self._display_results)
                self.root.after(0, lambda: self.status_label.config(
                    text=f"Validación completa. Se encontraron {len(self.df_resultado)} diferencias.",
                    foreground="red",
                ))
        except Exception as exc:
            print("[ERROR][ServicesValidationApp] Procesamiento fallido")
            traceback.print_exc()
            self.root.after(0, lambda: messagebox.showerror("Error", f"Procesamiento fallido:\n{exc}"))
            self.root.after(0, lambda: self.status_label.config(text="Procesamiento fallido", foreground="red"))

    def _normalizar_texto_equipo(self, texto):
        if not isinstance(texto, str):
            return texto
        texto = texto.replace("\n", " ")
        texto = re.sub(r"\s+", " ", texto)
        return texto.strip()

    def _clave_equipo(self, texto):
        """Clave robusta para emparejar equipos/servicios entre PDF y Excel.

        Pliega acentos y mayúsculas, descarta comillas/paréntesis/comas y otros
        signos, y elimina las conjunciones sueltas (y/o/e/u). Así
        ``Camperos y camionetas ... (10 Horas)`` (PDF) empareja con
        ``Camperos o camionetas ... (10 Horas)`` (Excel), pero sigue siendo
        distinto de la variante ``(24 Horas)`` porque conserva los dígitos.
        """
        if texto is None:
            return ""
        plano = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode().lower()
        plano = re.sub(r"[^a-z0-9]+", " ", plano)
        plano = re.sub(r"\b[yoeu]\b", " ", plano)
        return re.sub(r"\s+", " ", plano).strip()

    def _leer_excel_facturacion(self, path_hist):
        """Lee el Excel detectando la fila de encabezado real.

        Algunos cuadros traen filas de título antes de la cabecera
        ``COD. TAR. | DESCRIPCION TARIFA | ...``; se localiza esa fila y se usa
        como encabezado en vez de asumir la primera fila.
        """
        crudo = pd.read_excel(path_hist, header=None)
        fila_encabezado = 0
        for i in range(min(15, len(crudo))):
            celdas = {self._normalizar_busqueda(v) for v in crudo.iloc[i].tolist()}
            if "descripcion tarifa" in celdas:
                fila_encabezado = i
                break
        return pd.read_excel(path_hist, header=fila_encabezado)

    def _normalizar_perfil(self, valor):
        if not isinstance(valor, str):
            return valor
        texto = valor.strip()
        texto = texto.replace("Nivel", "").replace("Perfil", "")
        return texto.replace("/", "").strip()

    def _es_celda_vacia(self, valor):
        if valor is None:
            return True
        try:
            if pd.isna(valor):
                return True
        except Exception:
            pass

        texto_raw = str(valor)
        texto = unicodedata.normalize("NFKC", texto_raw)

        # Replace a broad set of invisible/whitespace characters with a single space
        texto = re.sub(r"[\s\u00A0\u2007\u202F\u200B\uFEFF\u2060\u200C\u200D]+", " ", texto)
        texto = texto.strip()

        # Remove surrounding matching quote characters repeatedly (handles '"   "')
        QUOTES = '"\'\u201C\u201D\u201E\u201F\u00AB\u00BB\u2039\u203A'
        while len(texto) >= 2 and texto[0] in QUOTES and texto[-1] in QUOTES:
            texto = texto[1:-1].strip()

        # Remove remaining quote characters and collapse interior whitespace
        texto = texto.replace('"', "").replace("'", "").replace("\u0000", "")
        texto = re.sub(r"\s+", "", texto).lower()

        if not texto:
            if self.debug_mode:
                self._debug_print(f"_es_celda_vacia: raw={repr(texto_raw)} -> normalized empty string")
            return True
        if self.debug_mode:
            self._debug_print(f"_es_celda_vacia: raw={repr(texto_raw)} -> normalized={repr(texto)}")
        return texto in {"nan", "none", "null"}

    def _normalizar_fecha(self, valor):
        meses = {
            "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
            "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
            "septiembre": 9, "setiembre": 9, "octubre": 10,
            "noviembre": 11, "diciembre": 12,
        }

        if valor is None:
            return None

        texto = str(valor).strip().lower()
        fecha = pd.to_datetime(texto, errors="coerce", dayfirst=True)
        if not pd.isna(fecha):
            return fecha.normalize()

        match = re.search(r"(\d{1,2})\s+de\s+([a-záéíóúñ]+)\s+de\s+(\d{4})", texto)
        if match:
            dia = int(match.group(1))
            mes = meses.get(match.group(2))
            anio = int(match.group(3))
            if mes is not None:
                try:
                    return pd.Timestamp(anio, mes, dia)
                except ValueError:
                    return None
        return None

    def _normalizar_busqueda(self, texto):
        if texto is None:
            return ""
        texto = unicodedata.normalize("NFKD", str(texto).strip().lower())
        texto = "".join(c for c in texto if not unicodedata.combining(c))
        return re.sub(r"\s+", " ", texto)

    def _buscar_indice_columna(self, row, opciones):
        opciones_norm = tuple(self._normalizar_busqueda(opcion) for opcion in opciones)
        for idx, cell in enumerate(row):
            cell_norm = self._normalizar_busqueda(cell)
            if any(opcion in cell_norm for opcion in opciones_norm):
                return idx
        return None

    def _parsear_cantidad(self, valor):
        if valor is None:
            return None

        match = re.search(r"\d{1,3}(?:\.\d{3})*(?:,\d+)?|\d+(?:[.,]\d+)?", str(valor))
        if not match:
            return None

        try:
            return float(match.group().replace(".", "").replace(",", "."))
        except ValueError:
            return None

    def _extraer_valor_etiqueta(self, tablas, etiquetas_objetivo):
        etiquetas_norm = tuple(self._normalizar_busqueda(etiqueta).replace(" ", "") for etiqueta in etiquetas_objetivo)

        for tabla in tablas:
            for row in tabla:
                for i, cell in enumerate(row):
                    cell_norm = self._normalizar_busqueda(cell).replace(" ", "")
                    if any(etiqueta in cell_norm for etiqueta in etiquetas_norm):
                        for next_cell in row[i + 1 :]:
                            if next_cell and str(next_cell).strip():
                                return self._normalizar_texto_equipo(next_cell)
        return None

    def _extraer_fecha_reporte(self, page_text, tablas):
        patron_fecha = (
            r"(\d{1,2}\s+de\s+[a-záéíóúñ]+\s+de\s+\d{4}|"
            r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|"
            r"\d{4}-\d{1,2}-\d{1,2})"
        )

        fecha_match = re.search(rf"fecha\s*:?.*?{patron_fecha}", page_text, flags=re.IGNORECASE)
        if fecha_match:
            return self._normalizar_fecha(fecha_match.group(1))

        fecha_match = re.search(patron_fecha, page_text, flags=re.IGNORECASE)
        if fecha_match:
            return self._normalizar_fecha(fecha_match.group(1))

        for tabla in tablas:
            for row in tabla:
                for cell in row:
                    if not cell:
                        continue
                    cell_text = str(cell)
                    if "fecha" not in self._normalizar_busqueda(cell_text):
                        continue
                    fecha_match = re.search(patron_fecha, cell_text, flags=re.IGNORECASE)
                    if fecha_match:
                        return self._normalizar_fecha(fecha_match.group(1))

        return None

    def _extraer_perfiles_pdf(self, path_planilla):
        conteo = Counter()
        fecha_reporte = None

        with pdfplumber.open(path_planilla) as pdf:
            for page in pdf.pages:
                for tabla in page.extract_tables() or []:
                    if not tabla or len(tabla) <= 7:
                        continue

                    header = tabla[6]
                    
                    header_norm = [self._normalizar_busqueda(celda).replace(" ", "") if celda else "" for celda in header]
                    if "nivel/perfil" not in header_norm:
                        continue

                    idx_perfil = header_norm.index("nivel/perfil")
                    if header:
                        fecha_detectada = self._normalizar_fecha(header[-1])
                        if fecha_detectada is not None:
                            fecha_reporte = fecha_detectada
                    tabla_info = str(tabla[4][2])
                    for row in tabla[7:]:
                        if len(row) <= idx_perfil:
                            continue
                        perfil = row[idx_perfil]
                        observacion = row[-1]
                        
                        tabla_info_upper = self._normalizar_busqueda(tabla_info).upper()
                        if "GLOBAL" in tabla_info_upper or "NO FACTURABLE" in tabla_info_upper:
                            continue

                        cantidad = 1 / 3 if "24" in tabla_info else 1

                        celda_validacion = row[7] if len(row) > 7 else None
                        if self._es_celda_vacia(celda_validacion):
                            self._debug_print(
                                f"Fila omitida por row[7] vacio: {repr(celda_validacion)} | perfil={repr(perfil)} | observacion={repr(observacion)}"
                            )
                            continue

                        if isinstance(perfil, str) and self._es_celda_vacia(observacion):
                            perfil = perfil.strip()
                            if perfil:
                                conteo[self._normalizar_perfil(perfil)] += cantidad
                                
                        if not self._es_celda_vacia(observacion):
                            perfil = str(observacion).split()[-1]
                            if perfil:
                                conteo[self._normalizar_perfil(perfil)] += cantidad
                                             
        return conteo, fecha_reporte

    def _extraer_conteo_pdf_detallado(self, path_planilla, tipo_formato):
        """Extrae registros de PDF de equipos o servicios según la estructura del formato."""
        registros = []
        tipo_formato = str(tipo_formato).lower()

        with pdfplumber.open(path_planilla) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                tablas = page.extract_tables() or []

                if tipo_formato == "equipos":
                    tipo_equipo = self._extraer_valor_etiqueta(tablas, ("equipo", "tipo de equipo", "tipo equipo"))
                    if not tipo_equipo:
                        continue

                    for tabla in tablas:
                        if not tabla or len(tabla) < 3:
                            continue

                        header_idx = None
                        idx_fecha = None
                        idx_cantidad = None

                        for i, row in enumerate(tabla[:12]):
                            idx_fecha = self._buscar_indice_columna(row, ("fecha", "dia"))
                            idx_cantidad = self._buscar_indice_columna(row, ("cant", "cantidad"))
                            if idx_fecha is not None and idx_cantidad is not None:
                                header_idx = i
                                break

                        if header_idx is None:
                            continue

                        for row in tabla[header_idx + 1 :]:
                            if len(row) <= max(idx_fecha, idx_cantidad):
                                continue

                            fecha = self._normalizar_fecha(row[idx_fecha])
                            cantidad = self._parsear_cantidad(row[idx_cantidad])
                            if not fecha or cantidad is None:
                                continue

                            registros.append(
                                {
                                    "FECHA": fecha,
                                    "TIPO DE EQUIPO": tipo_equipo,
                                    "CANTIDAD": cantidad,
                                }
                            )
                    continue

                if tipo_formato == "servicios":
                    fecha_reporte = self._extraer_fecha_reporte(page_text, tablas)
                    if fecha_reporte is None:
                        continue

                    for tabla in tablas:
                        if not tabla or len(tabla) < 3:
                            continue

                        header_idx = None
                        idx_tipo = None
                        idx_cantidad = None

                        for i, row in enumerate(tabla[:12]):
                            idx_tipo = self._buscar_indice_columna(row, ("tipo de equipo", "tipo equipo", "servicio"))
                            idx_cantidad = self._buscar_indice_columna(row, ("cant", "cantidad"))
                            if idx_tipo is not None and idx_cantidad is not None:
                                header_idx = i
                                break

                        if header_idx is None:
                            continue

                        for row in tabla[header_idx + 1 :]:
                            if len(row) <= max(idx_tipo, idx_cantidad):
                                continue

                            tipo = self._normalizar_texto_equipo(row[idx_tipo])
                            cantidad = self._parsear_cantidad(row[idx_cantidad])
                            if not isinstance(tipo, str) or not tipo.strip() or cantidad is None:
                                continue

                            registros.append(
                                {
                                    "FECHA": fecha_reporte,
                                    "TIPO DE EQUIPO": tipo,
                                    "CANTIDAD": cantidad,
                                }
                            )
                    continue

                raise ValueError(f"Unknown extraction type: {tipo_formato}")

        df = pd.DataFrame(registros)
        if df.empty:
            return df

        return df.groupby(["FECHA", "TIPO DE EQUIPO"], as_index=False)["CANTIDAD"].sum()

    def _extraer_equipos_pdf(self, path_planilla):
        return self._extraer_conteo_pdf_detallado(path_planilla, "equipos")

    def _extraer_servicios_pdf(self, path_planilla):
        return self._extraer_conteo_pdf_detallado(path_planilla, "servicios")

    def _extraer_conteo_pdf(self, path_planilla, tipo_extraccion="equipos"):
        if tipo_extraccion.lower() == "perfiles":
            return self._extraer_perfiles_pdf(path_planilla)
        if tipo_extraccion.lower() == "equipos":
            return self._extraer_equipos_pdf(path_planilla)
        if tipo_extraccion.lower() == "servicios":
            return self._extraer_servicios_pdf(path_planilla)
        raise ValueError(f"Unknown extraction type: {tipo_extraccion}")

    def _extraer_conteo_excel_perfiles(self, path_hist, fecha_reporte):
        if fecha_reporte is None:
            raise ValueError("No fue posible detectar la fecha de referencia en el PDF.")

        df_hist = self._leer_excel_facturacion(path_hist)
        if "DESCRIPCION TARIFA" not in df_hist.columns:
            raise KeyError("El archivo Excel no contiene la columna 'DESCRIPCION TARIFA'.")

        df_niveles = df_hist[df_hist["DESCRIPCION TARIFA"].notna()].copy()
        df_niveles = df_niveles[df_niveles["DESCRIPCION TARIFA"].astype(str).str.contains("Nivel|Perfil", na=False)].copy()

        cols_fecha = [col for col in df_niveles.columns if isinstance(col, (pd.Timestamp, dt.datetime))]
        if not cols_fecha:
            raise ValueError("No se detectaron columnas de fecha en el archivo Excel.")

        cols_id = [col for col in df_niveles.columns if col not in cols_fecha]
        df_largo = df_niveles.melt(id_vars=cols_id, value_vars=cols_fecha, var_name="FECHA", value_name="VALOR")
        df_largo["FECHA"] = pd.to_datetime(df_largo["FECHA"], errors="coerce").dt.normalize()

        df_fecha = df_largo[df_largo["FECHA"] == fecha_reporte].copy()
        df_fecha = df_fecha[df_fecha["VALOR"].notna()]
        df_fecha = df_fecha.groupby(["DESCRIPCION TARIFA"], as_index=False)["VALOR"].sum()
        df_fecha = df_fecha[df_fecha["VALOR"] != 0]
        df_fecha["PERFIL_NORM"] = df_fecha["DESCRIPCION TARIFA"].apply(self._normalizar_perfil)
        return df_fecha.set_index("PERFIL_NORM")["VALOR"].to_dict()

    def _comparar_conteos_perfiles(self, conteo_pdf, conteo_excel):
        diferencias = []
        todos_perfiles = set(conteo_pdf.keys()).union(conteo_excel.keys())

        for perfil in sorted(todos_perfiles):
            pdf_cnt = conteo_pdf.get(perfil, 0)
            excel_cnt = conteo_excel.get(perfil, 0)
            if isinstance(excel_cnt, float) and excel_cnt.is_integer():
                excel_cnt = int(excel_cnt)
            if isinstance(pdf_cnt, float) and pdf_cnt.is_integer():
                pdf_cnt = int(pdf_cnt)

            if pdf_cnt != excel_cnt:
                diferencias.append({
                    "Nivel/Perfil": perfil,
                    "PDF": pdf_cnt,
                    "Excel": excel_cnt,
                    "Diferencia": abs(excel_cnt - pdf_cnt),
                })

        return pd.DataFrame([
            d for d in diferencias
            if d["Nivel/Perfil"] not in {"None", "OBSERVACIONES", "INCAPACIDAD"}
        ])

    def _extraer_conteo_excel(self, path_hist, fecha_reporte):
        if fecha_reporte is None:
            raise ValueError("No reference date detected in PDF.")

        df_hist = self._leer_excel_facturacion(path_hist)
        if "DESCRIPCION TARIFA" not in df_hist.columns:
            raise KeyError("Excel file missing 'DESCRIPCION TARIFA' column.")

        df_niveles = df_hist[df_hist["DESCRIPCION TARIFA"].notna()].copy()
        cols_fecha = [col for col in df_niveles.columns if isinstance(col, (pd.Timestamp, dt.datetime))]
        if not cols_fecha:
            raise ValueError("No date columns detected in Excel file.")

        cols_id = [col for col in df_niveles.columns if col not in cols_fecha]
        df_largo = df_niveles.melt(id_vars=cols_id, value_vars=cols_fecha, var_name="FECHA", value_name="VALOR")
        df_largo["FECHA"] = pd.to_datetime(df_largo["FECHA"], errors="coerce").dt.normalize()

        df_fecha = df_largo[df_largo["FECHA"] == fecha_reporte].copy()
        df_fecha = df_fecha[df_fecha["VALOR"].notna()]
        # Clave robusta para que el emparejamiento tolere comillas/conjunciones.
        df_fecha["PERFIL_NORM"] = df_fecha["DESCRIPCION TARIFA"].apply(self._clave_equipo)
        df_fecha = df_fecha.groupby(["PERFIL_NORM"], as_index=False)["VALOR"].sum()
        df_fecha = df_fecha[df_fecha["VALOR"] != 0]
        return df_fecha.set_index("PERFIL_NORM")["VALOR"].to_dict()

    def _comparar_conteos(self, df_pdf, path_excel):
        if df_pdf is None or df_pdf.empty:
            return pd.DataFrame(columns=["Fecha", "Servicio", "PDF", "Excel", "Diferencia"])

        pdf_agg = df_pdf.groupby(["FECHA", "TIPO DE EQUIPO"], as_index=False)["CANTIDAD"].sum()
        diferencias = []

        for fecha in sorted(pdf_agg["FECHA"].dropna().unique()):
            conteo_excel = self._extraer_conteo_excel(path_excel, fecha)
            for _, row in pdf_agg[pdf_agg["FECHA"] == fecha].iterrows():
                servicio = row["TIPO DE EQUIPO"]
                pdf_cnt = row["CANTIDAD"]
                excel_cnt = conteo_excel.get(self._clave_equipo(servicio), 0)
                diferencias.append({
                    "Fecha": fecha,
                    "Servicio": servicio,
                    "PDF": pdf_cnt,
                    "Excel": excel_cnt,
                    "Diferencia": abs(excel_cnt - pdf_cnt),
                })

        diferencias = pd.DataFrame(diferencias)
        if diferencias.empty:
            return diferencias

        diferencias = diferencias[~diferencias["Servicio"].isin({"None", "OBSERVACIONES", "INCAPACIDAD"})].copy()
        if diferencias["Diferencia"].sum() == 0:
            return diferencias.sort_values(["Fecha", "Servicio"]).reset_index(drop=True)
        return diferencias[diferencias["Diferencia"] > 0].sort_values(["Fecha", "Servicio"]).reset_index(drop=True)

    def _show_all_ok_message(self):
        messagebox.showinfo(
            "Validación exitosa ✅",
            "No se encontraron diferencias.\n\nTodas las cantidades de servicios y equipos coinciden entre el PDF y el Excel.",
        )
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _show_custom_ok_message(self, titulo, mensaje):
        messagebox.showinfo(titulo, mensaje)
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _display_profile_results(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        if self.df_resultado is None or self.df_resultado.empty:
            return

        for _, row in self.df_resultado.iterrows():
            self.tree.insert(
                "",
                tk.END,
                values=("", str(row["Nivel/Perfil"]), self._format_number(row["PDF"]), self._format_number(row["Excel"]), self._format_number(row["Diferencia"])),
                tags=("error",),
            )

    def _display_results(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        if self.df_resultado is None or self.df_resultado.empty:
            return

        for _, row in self.df_resultado.iterrows():
            fecha_str = row["Fecha"].strftime("%Y-%m-%d") if pd.notna(row["Fecha"]) else ""
            self.tree.insert(
                "",
                tk.END,
                values=(
                    fecha_str,
                    str(row["Servicio"])[:70],
                    self._format_number(row["PDF"]),
                    self._format_number(row["Excel"]),
                    self._format_number(row["Diferencia"]),
                ),
                tags=("error",),
            )

    def _format_number(self, valor):
        if pd.isna(valor):
            return ""
        try:
            if float(valor) == int(float(valor)):
                return str(int(float(valor)))
            return f"{float(valor):.2f}"
        except (ValueError, TypeError):
            return str(valor)

    def _export_to_csv(self):
        if self.df_resultado is None or self.df_resultado.empty:
            messagebox.showwarning("Aviso", "No hay resultados para exportar. Primero valida los archivos.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")],
        )
        if not file_path:
            return

        try:
            self.df_resultado.to_csv(file_path, index=False)
            messagebox.showinfo("Éxito", f"Resultados exportados en:\n{file_path}")
        except Exception as exc:
            messagebox.showerror("Error", f"La exportación falló:\n{exc}")

    def _clear_results(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.df_resultado = None
        self.status_label.config(text="Resultados limpiados", foreground="blue")


def main():
    if tk is None:
        raise RuntimeError("Tkinter is not available in this environment.")
    root = tk.Tk()
    ServicesValidationApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
