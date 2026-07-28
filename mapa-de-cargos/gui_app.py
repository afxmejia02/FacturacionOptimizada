"""
Aplicación GUI para conciliación de datos de nómina.

Esta aplicación procesa documentos de nómina y concilia desprendibles
con registros de transferencias bancarias para verificar la consistencia de los pagos.

Características:
- Selección de carpetas con archivos PDF de desprendibles y transferencias
- Extracción automática y cruce de registros de pago
- Resultados mostrados en una tabla interactiva con colores
- Resaltado en rojo para documentos faltantes o no encontrados
- Exportación de resultados a CSV

Autor: Sistema de Conciliación de Nómina
Versión: 1.0
"""

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except Exception:  # pragma: no cover - allows headless import on Streamlit Cloud
    tk = None
    filedialog = None
    messagebox = None
    ttk = None
import pandas as pd
import pdfplumber
import re

import os
import unicodedata
from collections import Counter
from pathlib import Path
import threading


# --- Detección robusta de renglones de soporte bancario (transferencias) -----
# La etiqueta de destino real ("PAGO NOMINA BCA") varía entre soportes y puede
# llegar con ruido de OCR. Solo se usa como señal de que la línea es un pago de
# nómina; el cruce no depende de su forma exacta. Tolera confusiones típicas de
# OCR: O/0, I/1, A/4.
_RE_NOMINA = re.compile(r"N[O0]M[I1]N[A4]", re.IGNORECASE)

# Importe monetario: admite miles con punto o coma y decimales opcionales.
#   1.234.567,89 | 1,234,567.89 | 5156483 | 913,143.00
_RE_IMPORTE = re.compile(r"\d{1,3}(?:[.,]\d{3})+(?:[.,]\d{1,2})?|\d+(?:[.,]\d{1,2})?")

# Documento del beneficiario (cédula/NIT): grupo de 5 a 15 dígitos.
_RE_DOCUMENTO = re.compile(r"\b(\d{5,15})\b")

# Cuenta / número de producto (opcional): grupo largo de 9+ dígitos.
_RE_PRODUCTO = re.compile(r"\b(\d{9,})\b")

# Número de factura: token (YYMMDD o YYYYMMDD) inmediatamente antes de "PAGO".
# Es la referencia de la quincena del soporte (p. ej. 250415 -> 2025-04-15), no
# la fecha real de consignación. Se usa para filtrar transferencias por periodo.
_RE_FACTURA = re.compile(r"(\d{6,8})\s+PAG", re.IGNORECASE)


class PayrollReconciliationApp:
    """
    Aplicación principal GUI para la conciliación de datos de nómina.
    
    Attributes:
        root (tk.Tk): Main application window
        desprendibles_path (tk.StringVar): Path to payslips folder
        transferencia_path (tk.StringVar): Path to transfers folder
        df_resultado (pd.DataFrame): Final reconciliation results
    """
    
    def __init__(self, root, mode="both"):
        """Inicializa la aplicación."""
        self.root = root
        self.root.title("Sistema de Conciliación de Nómina")
        self.root.geometry("1400x800")
        
        # Variables para almacenar las rutas seleccionadas
        self.desprendibles_path = tk.StringVar()
        self.transferencia_path = tk.StringVar()
        self.seguridad_path = tk.StringVar()
        
        # DataFrame de resultados
        self.df_resultado = None
        # mode can be 'both', 'transfers', or 'seguridad'
        self.mode = mode
        
        # Configurar la interfaz
        self._create_ui()
        
    def _create_ui(self):
        """Crea los componentes de la interfaz de usuario."""
        
        # ===== HEADER FRAME =====
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill=tk.X, padx=10, pady=10)
        
        title_label = ttk.Label(
            header_frame,
            text="Sistema de Conciliación de Nómina",
            font=("Arial", 16, "bold")
        )
        title_label.pack(side=tk.LEFT)
        
        # ===== PATH SELECTION FRAME =====
        path_frame = ttk.LabelFrame(self.root, text="Seleccionar carpetas", padding=10)
        path_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Payslips folder
        ttk.Label(path_frame, text="Carpeta de desprendibles:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(path_frame, textvariable=self.desprendibles_path, width=80).grid(row=0, column=1, padx=5)
        ttk.Button(path_frame, text="Buscar", command=self._select_desprendibles_folder).grid(row=0, column=2)
        
        # Transfers folder
        ttk.Label(path_frame, text="Carpeta de transferencias:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(path_frame, textvariable=self.transferencia_path, width=80).grid(row=1, column=1, padx=5)
        ttk.Button(path_frame, text="Buscar", command=self._select_transferencia_folder).grid(row=1, column=2)

        # Seguridad social folder
        ttk.Label(path_frame, text="Carpeta seguridad social (IBC):").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(path_frame, textvariable=self.seguridad_path, width=80).grid(row=2, column=1, padx=5)
        ttk.Button(path_frame, text="Buscar", command=self._select_seguridad_folder).grid(row=2, column=2)
        
        # ===== CONTROL BUTTONS FRAME =====
        control_frame = ttk.Frame(self.root, padding=10)
        control_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(
            control_frame,
            text="Procesar PDFs",
            command=self._process_files,
            width=20
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            control_frame,
            text="Exportar a CSV",
            command=self._export_to_csv,
            width=20
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            control_frame,
            text="Limpiar resultados",
            command=self._clear_results,
            width=20
        ).pack(side=tk.LEFT, padx=5)
        
        # Status label
        self.status_label = ttk.Label(control_frame, text="Listo", foreground="blue")
        self.status_label.pack(side=tk.LEFT, padx=20)
        
        # ===== RESULTS FRAME =====
        results_frame = ttk.LabelFrame(self.root, text="Resultados", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Create frames and treeviews depending on mode
        results_frame.grid_columnconfigure(0, weight=1)

        self.tree_transfers = None
        self.tree_seguridad = None

        if self.mode in ("both", "transfers"):
            transfers_frame = ttk.LabelFrame(results_frame, text="Revisión Transferencias", padding=5)
            transfers_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
            results_frame.grid_rowconfigure(0, weight=1)

            self.tree_transfers = ttk.Treeview(
                transfers_frame,
                columns=("Identificación", "Cuenta", "Estado", "Neto_desprendibles", "Valores_transferencia"),
                height=10,
                show="headings"
            )

            self.tree_transfers.heading("Identificación", text="Identificación")
            self.tree_transfers.heading("Cuenta", text="Cuenta")
            self.tree_transfers.heading("Estado", text="Estado")
            self.tree_transfers.heading("Neto_desprendibles", text="Valores desprendibles")
            self.tree_transfers.heading("Valores_transferencia", text="Valores de transferencia")

            self.tree_transfers.column("Identificación", width=150, anchor=tk.CENTER)
            self.tree_transfers.column("Cuenta", width=150, anchor=tk.CENTER)
            self.tree_transfers.column("Estado", width=200, anchor=tk.CENTER)
            self.tree_transfers.column("Neto_desprendibles", width=300, anchor=tk.W)
            self.tree_transfers.column("Valores_transferencia", width=300, anchor=tk.W)

            vsb_t = ttk.Scrollbar(transfers_frame, orient=tk.VERTICAL, command=self.tree_transfers.yview)
            hsb_t = ttk.Scrollbar(transfers_frame, orient=tk.HORIZONTAL, command=self.tree_transfers.xview)
            self.tree_transfers.configure(yscroll=vsb_t.set, xscroll=hsb_t.set)
            self.tree_transfers.grid(row=0, column=0, sticky="nsew")
            vsb_t.grid(row=0, column=1, sticky="ns")
            hsb_t.grid(row=1, column=0, sticky="ew")
            transfers_frame.grid_rowconfigure(0, weight=1)
            transfers_frame.grid_columnconfigure(0, weight=1)

        if self.mode in ("both", "seguridad"):
            seguridad_frame = ttk.LabelFrame(results_frame, text="Revisión Seguridad Social (IBC)", padding=5)
            row_idx = 1 if self.mode == "both" else 0
            seguridad_frame.grid(row=row_idx, column=0, sticky="nsew", padx=5, pady=5)
            results_frame.grid_rowconfigure(row_idx, weight=1)

            self.tree_seguridad = ttk.Treeview(
                seguridad_frame,
                columns=("Identificación", "Estado", "Devengado", "IBC"),
                height=10,
                show="headings"
            )

            self.tree_seguridad.heading("Identificación", text="Identificación")
            self.tree_seguridad.heading("Estado", text="Estado")
            self.tree_seguridad.heading("Devengado", text="Devengado")
            self.tree_seguridad.heading("IBC", text="IBC")

            self.tree_seguridad.column("Identificación", width=150, anchor=tk.CENTER)
            self.tree_seguridad.column("Estado", width=250, anchor=tk.CENTER)
            self.tree_seguridad.column("Devengado", width=200, anchor=tk.W)
            self.tree_seguridad.column("IBC", width=300, anchor=tk.W)

            vsb_s = ttk.Scrollbar(seguridad_frame, orient=tk.VERTICAL, command=self.tree_seguridad.yview)
            hsb_s = ttk.Scrollbar(seguridad_frame, orient=tk.HORIZONTAL, command=self.tree_seguridad.xview)
            self.tree_seguridad.configure(yscroll=vsb_s.set, xscroll=hsb_s.set)
            self.tree_seguridad.grid(row=0, column=0, sticky="nsew")
            vsb_s.grid(row=0, column=1, sticky="ns")
            hsb_s.grid(row=1, column=0, sticky="ew")
            seguridad_frame.grid_rowconfigure(0, weight=1)
            seguridad_frame.grid_columnconfigure(0, weight=1)

        # Configurar etiquetas de color para ambas tablas (si existen)
        for t in (self.tree_transfers, self.tree_seguridad):
            if t is None:
                continue
            t.tag_configure("error", background="#ffcccc", foreground="darkred")
            t.tag_configure("ok", background="#ccffcc", foreground="darkgreen")
            t.tag_configure("warning", background="#ffffcc", foreground="darkorange")
        
    def _select_desprendibles_folder(self):
        """Open folder selection dialog for payslips."""
        folder = filedialog.askdirectory(title="Select Payslips Folder (Desprendibles)")
        if folder:
            self.desprendibles_path.set(folder)
            
    def _select_transferencia_folder(self):
        """Open folder selection dialog for transfers."""
        folder = filedialog.askdirectory(title="Select Transfers Folder (Transferencias)")
        if folder:
            self.transferencia_path.set(folder)

    def _select_seguridad_folder(self):
        """Open folder selection dialog for seguridad social PDFs (IBC)."""
        folder = filedialog.askdirectory(title="Select Seguridad Social Folder (IBC)")
        if folder:
            self.seguridad_path.set(folder)
    
    def _process_files(self):
        """Process PDF files from selected folders."""
        # Validate folder selections
        if not self.desprendibles_path.get():
            messagebox.showerror("Error", "Please select a Payslips folder")
            return
        if not self.transferencia_path.get():
            messagebox.showerror("Error", "Please select a Transfers folder")
            return
        
        # Update status
        self.status_label.config(text="Processing... Please wait.", foreground="orange")
        self.root.update()
        
        # Ejecutar el procesamiento en un hilo separado para no bloquear la interfaz
        thread = threading.Thread(target=self._process_files_thread)
        thread.daemon = True
        thread.start()
    
    def _process_files_thread(self):
        """Procesa los archivos en segundo plano."""
        try:
            # Procesar desprendibles
            df_desprendibles = self._process_desprendibles(self.desprendibles_path.get())
            
            # Procesar transferencias
            df_transferencia = self._process_transferencia(self.transferencia_path.get())
            
            # Procesar seguridad social (buscar IBCs) si se indicó carpeta
            if self.seguridad_path.get():
                try:
                    df_seguridad = self.procesar_seguridad_social(self.seguridad_path.get())
                except Exception:
                    df_seguridad = pd.DataFrame()
            else:
                df_seguridad = pd.DataFrame()
            
            # Conciliar datos (incluye Devengado e IBC si están disponibles)
            df_transfers, df_seguridad_res = self._reconcile_data(df_desprendibles, df_transferencia, df_seguridad)
            self.df_resultado_transfers = df_transfers
            self.df_resultado_seguridad = df_seguridad_res

            # Actualizar la interfaz en el hilo principal
            self.root.after(0, self._display_results)
            total = 0
            try:
                total = len(self.df_resultado_transfers) + len(self.df_resultado_seguridad)
            except Exception:
                total = 0
            self.root.after(0, lambda: self.status_label.config(
                text=f"Proceso completado. Se encontraron {total} registros.",
                foreground="green"
            ))
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"El procesamiento falló:\n{str(e)}"))
            self.root.after(0, lambda: self.status_label.config(
                text="Procesamiento fallido",
                foreground="red"
            ))
    
    def _log(self, mensaje):
        """Log de depuración (visible salvo que ``VALIDATION_DEBUG`` sea ``"0"``).

        No depende del estado de instancia, por lo que funciona también cuando la
        app se crea con ``__new__`` (como hace la web, sin abrir ventanas).
        """
        if os.environ.get("VALIDATION_DEBUG", "1") != "0":
            print(f"[DEBUG][PayrollReconciliationApp] {mensaje}")

    def _process_desprendibles(self, folder_path, formato="tabarca"):
        """
        Extrae datos de desprendibles desde archivos PDF.

        Args:
            folder_path (str): Ruta de la carpeta con PDFs de desprendibles
            formato (str): Formato del desprendible a interpretar ("tabarca" o "italco")

        Returns:
            pd.DataFrame: DataFrame con las columnas [Identificacion, Neto, Devengado, Cuenta]
        """
        formato = (formato or "tabarca").strip().lower()
        if formato == "italco":
            return self._process_desprendibles_italco(folder_path)
        return self._process_desprendibles_tabarca(folder_path)

    def _process_desprendibles_tabarca(self, folder_path):
        registros = []
        
        # Buscar PDFs de desprendibles (normalmente nombrados por mes/año)
        for filename in os.listdir(folder_path):
            if filename.endswith(".pdf"):
                path = os.path.join(folder_path, filename)
                
                with pdfplumber.open(path) as pdf:
                    for page in pdf.pages:
                        texto = page.extract_text()
                        if not texto:
                            continue
                        
                        # Separar por el marcador del bloque del desprendible
                        bloques = texto.split("Comprobante de Nómina")
                        
                        for bloque in bloques:
                            # Extraer identificación (formato con puntos: 123.456.789)
                            id_match = re.search(r"\b\d{1,3}(?:[.,]\d{3}){1,3}\b", bloque)
                            cuenta_match = re.search(r"Cuenta No\s*(\d{6,})\b", bloque)
                            
                            # Extraer el valor neto a pagar
                            # Extraer el valor neto a pagar (más flexible, case-insensitive)
                            neto_match = re.search(
                                r"Neto(?:\s+a\s+pagar)?[:\s].*?\$\s*([\d\.,]+)",
                                bloque,
                                re.IGNORECASE | re.DOTALL
                            )

                            # Extraer Devengado: buscar 'Devengado' primero, si no, intentar 'TOTALES'
                            #si en la columna SALDOS hay

                            deven_match = re.search(
                                    r"TOTALES[:\s].*?\$\s*([\d\.,]+)",
                                    bloque,
                                    re.IGNORECASE | re.DOTALL
                                )
                            
                            aux_match = re.search(r"AUXILIO DE LOCALIZACION\s+([\d.,]+)\s+([\d.,]+)", bloque)
                            pri_match = re.search(r"PRIMA LEGAL DE SERVICIOS JUNIO\s+([\d.,]+)\s+([\d.,]+)", bloque)
                            vac_match = re.search(r"VACACIONES INDEMNIZADAS\s+([\d.,]+)\s+([\d.,]+)", bloque)
                            ces_match = re.search(r"CESANTIAS\s+([\d.,]+)\s+([\d.,]+)", bloque)
                            int_ces_match= re.search(r"INTERESES DE CESANTIAS\s+([\d.,]+)\s+([\d.,]+)", bloque)
                            

                            
                            if id_match and neto_match:
                                identificacion = id_match.group()
                                identificacion = identificacion.replace(".", "").replace(",", "")
                                neto = self._limpiar_numero(neto_match.group(1))
                                devengado = self._limpiar_numero(deven_match.group(1)) if deven_match else None
                                dtos= []
                                for m in [aux_match, pri_match, vac_match, ces_match, int_ces_match]:
                                    if m:
                                        valor = self._limpiar_numero(m.group(2))
                                        if valor is not None:
                                            dtos.append(valor)
                                if devengado is not None and sum(dtos) is not None:
                                    devengado -= sum(dtos)
                                cuenta = cuenta_match.group(1) if cuenta_match else None
                                
                                registros.append({
                                    "Identificacion": identificacion,
                                    "Neto": neto,
                                    "Devengado": devengado,
                                    "Cuenta": cuenta
                                })
        
        df = pd.DataFrame(registros)
        
        if not df.empty:
            # Eliminar puntos de la identificación
            df["Identificacion"] = df["Identificacion"].str.replace(".", "", regex=False)
            # Convertir Neto a entero
            df["Neto"] = df["Neto"].fillna(0).astype("int64")
            # Convertir Devengado a entero
            if "Devengado" in df.columns:
                df["Devengado"] = df["Devengado"].fillna(0).astype("int64")
        return df

    def _process_desprendibles_italco(self, folder_path):
        """
        Extrae datos de los comprobantes de pago de nómina en formato ITALCO.

        A diferencia del formato TABARCA, cada página es un comprobante con la
        cédula tras ``CC:`` y el neto tras ``Total Neto:`` (sin símbolo ``$``).

        Returns:
            pd.DataFrame: DataFrame con las columnas [Identificacion, Neto, Devengado, Cuenta]
        """
        registros = []

        for filename in os.listdir(folder_path):
            if not filename.lower().endswith(".pdf"):
                continue

            path = os.path.join(folder_path, filename)
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:
                    texto = page.extract_text() or ""
                    if not texto:
                        continue

                    texto_plano = re.sub(r"\s+", " ", texto)
                    # Documento: preferir "CC: <num>"; si el CC viene vacío (algunas
                    # plantillas lo dejan en blanco), usar "Documento <num>" de la
                    # línea del trabajador.
                    cc_match = re.search(r"CC:\s*(\d[\d.,]*)", texto_plano, re.IGNORECASE)
                    doc_match = re.search(r"\bDocumento\b\s*:?\s*(\d[\d.,]*)", texto_plano, re.IGNORECASE)
                    fuente_doc = cc_match or doc_match
                    # Los valores pueden traer símbolo "$" (p. ej. "Total Neto: $3,675,627").
                    neto_match = re.search(r"Total Neto:\s*\$?\s*([\d,\.]+)", texto_plano, re.IGNORECASE)
                    cuenta_match = re.search(r"CUENTA:\s*(\d+)", texto_plano, re.IGNORECASE)
                    # El devengado en ITALCO es el TOTAL INGRESOS (no el Total Neto).
                    dev_match = re.search(r"TOTAL INGRESOS\s*\$?\s*([\d,\.]+)", texto_plano, re.IGNORECASE)
                    # Periodo de la quincena: "Periodo: 2025-04-01 al 2025-04-15".
                    # Se usa para descartar transferencias de otras quincenas/meses.
                    periodo_match = re.search(
                        r"Periodo:?\s*(\d{4}-\d{1,2}-\d{1,2})\s*al\s*(\d{4}-\d{1,2}-\d{1,2})",
                        texto_plano,
                        re.IGNORECASE,
                    )

                    if not (fuente_doc and neto_match):
                        continue

                    identificacion = re.sub(r"[^\d]", "", fuente_doc.group(1))
                    neto = self._limpiar_numero(neto_match.group(1))
                    devengado = self._limpiar_numero(dev_match.group(1)) if dev_match else None
                    cuenta = cuenta_match.group(1) if cuenta_match else None
                    periodo_inicio = (
                        pd.to_datetime(periodo_match.group(1), errors="coerce")
                        if periodo_match else pd.NaT
                    )
                    periodo_fin = (
                        pd.to_datetime(periodo_match.group(2), errors="coerce")
                        if periodo_match else pd.NaT
                    )

                    if not identificacion:
                        continue

                    registros.append({
                        "Identificacion": identificacion,
                        "Neto": neto,
                        "Devengado": devengado,
                        "Cuenta": cuenta,
                        "PeriodoInicio": periodo_inicio,
                        "PeriodoFin": periodo_fin,
                    })

        df = pd.DataFrame(registros)

        if not df.empty:
            df["Neto"] = df["Neto"].fillna(0).astype("int64")
            df["Devengado"] = df["Devengado"].fillna(0).astype("int64")
        return df

    def _process_transferencia(self, folder_path, formato="tabarca"):
        """
        Extract transfer data from PDF files.

        Args:
            folder_path (str): Path to folder containing transfer PDFs
            formato (str): Transfer layout to parse ("tabarca" or "italco")

        Returns:
            pd.DataFrame: DataFrame with transfer information
        """
        formato = (formato or "tabarca").strip().lower()
        if formato == "italco":
            return self._process_transferencia_italco(folder_path)
        return self._process_transferencia_tabarca(folder_path)

    def _process_transferencia_tabarca(self, folder_path):
        registros = []

        # Iterate PDF files in the folder and parse transfer lines
        for filename in os.listdir(folder_path):
            if not filename.lower().endswith(".pdf"):
                continue
            path = os.path.join(folder_path, filename)
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:
                    texto = page.extract_text()
                    if not texto:
                        continue

                    for linea in texto.split("\n"):
                        linea = linea.strip()
                        # Detect lines with account numbers (10+ digits)
                        if re.match(r"\d{10,}", linea):
                            data = self._parsear_linea(linea)
                            if data:
                                registros.append(data)

        df = pd.DataFrame(registros)

        if not df.empty and "Valor" in df.columns:
            # Normalize and convert Valor to integer using the existing cleaner
            def _to_int(v):
                try:
                    if v is None:
                        return None
                    num = self._limpiar_numero(str(v))
                    return int(num) if num is not None else None
                except Exception:
                    return None

            df["Valor"] = df["Valor"].apply(_to_int)
            df["Valor"] = df["Valor"].fillna(0).astype("int64")

        return df

    def _normalizar_linea_ocr(self, linea):
        """Normaliza una línea para un matching tolerante a OCR.

        Pliega acentos a ASCII, reemplaza espacios duros (NBSP) y colapsa
        cualquier secuencia de espacios en uno solo. No altera los dígitos, solo
        homogeneiza el texto para que las búsquedas no fallen por tildes o por
        espaciado irregular del PDF/OCR.
        """
        texto = unicodedata.normalize("NFKD", str(linea or ""))
        texto = texto.encode("ascii", "ignore").decode()
        texto = texto.replace(" ", " ")
        return re.sub(r"\s+", " ", texto).strip()

    def _lineas_desde_palabras(self, page, x_tolerance=1.5, y_tolerance=3.0):
        """Reconstruye las líneas de la página a partir de las palabras y sus
        coordenadas (x, y), no de ``extract_text()``.

        Es necesario porque en algunos soportes columnas contiguas (p. ej. la
        columna "Productos" y el documento, o el documento y el número de
        producto) quedan pegadas o entrelazadas en ``extract_text()``, lo que
        rompe la lectura del documento. Agrupando por línea (``top``) y ordenando
        por ``x0`` se recuperan los campos como tokens separados.

        Si no hay palabras (página sin capa de texto), cae a ``extract_text()``.
        """
        try:
            words = page.extract_words(x_tolerance=x_tolerance, use_text_flow=False)
        except Exception:
            words = []
        if not words:
            return (page.extract_text() or "").split("\n")

        words = sorted(words, key=lambda w: (w["top"], w["x0"]))
        lineas = []
        actual = []
        base_top = None
        for w in words:
            if base_top is not None and abs(w["top"] - base_top) > y_tolerance:
                lineas.append(" ".join(c["text"] for c in sorted(actual, key=lambda c: c["x0"])))
                actual = []
                base_top = None
            actual.append(w)
            if base_top is None:
                base_top = w["top"]
        if actual:
            lineas.append(" ".join(c["text"] for c in sorted(actual, key=lambda c: c["x0"])))
        return lineas

    def _match_linea_transferencia(self, linea):
        """Extrae ``{Documento, Cuenta, Nombre, Valor}`` de una línea de soporte.

        Diseño robusto y genérico (no atado a un layout concreto). Tolera:

        - **ausencia/presencia de la columna de fecha** (el bug original: el
          patrón exigía una fecha de 8 dígitos que muchos soportes no traen);
        - cualquier **formato monetario** (miles con punto o coma, con/sin
          decimales) vía :func:`_limpiar_numero`;
        - **espacios y OCR ruidoso** (se normaliza la línea antes de leerla);
        - **variantes de la etiqueta de destino** ("PAGO NOMINA BCA",
          "PAGO DE NOMINA", "PAGONOMINA"...): basta con que aparezca "NOMINA".

        Estrategia por proximidad de campos en el renglón:
          ``[productos] <nombre> <documento> [cuenta] [fechaPago] <factura> <...NOMINA...> <valor> [ods]``
          - documento = primer grupo de 5-15 dígitos (cédula/NIT del beneficiario);
          - valor = primer importe DESPUÉS de la etiqueta de destino (evita tomar
            la columna "ods"/consecutivo que algunos soportes ponen al final);
          - cuenta = primer grupo largo (9+ dígitos) tras el documento (opcional);
          - fecha = el número de factura (YYMMDD) justo antes de "PAGO", que indica
            la quincena del soporte (se usa luego para filtrar por periodo).

        Hay un segundo layout (la "consulta de pagos a terceros" del banco) cuyos
        renglones **no traen la etiqueta NÓMINA** ni la fecha-factura de quincena
        (solo la fecha de consignación, que puede ser de otro mes). Esos renglones
        se devuelven como **candidatos** (``EsNomina=False``): se extraen documento
        y valor, pero solo se aceptan en la conciliación si su valor coincide con
        un neto del desprendible (los que no, pueden ser de otra quincena).

        Devuelve ``None`` si la línea no parece un renglón de pago.
        """
        plano = self._normalizar_linea_ocr(linea)
        if not plano:
            return None

        # 1) Documento del beneficiario: primer grupo de 5-15 dígitos.
        m_doc = _RE_DOCUMENTO.search(plano)
        if not m_doc:
            return None
        documento = m_doc.group(1)
        nombre = plano[: m_doc.start()].strip(" -")
        # Quitar un eventual número de "Productos" al inicio del renglón.
        nombre = re.sub(r"^[\d\s]+", "", nombre).strip(" -") or None

        # 2) ¿Trae la etiqueta NÓMINA? (señal tolerante a OCR). Si la trae, es una
        #    transferencia confiable; si no, es un candidato a validar por valor.
        m_nomina = _RE_NOMINA.search(plano)
        es_nomina = bool(m_nomina)

        # 3) Valor.
        if m_nomina:
            # Confiable: primer importe DESPUÉS de la etiqueta de destino. Así no se
            # confunde con la columna "ods" (1-3 dígitos) al final del renglón.
            importes_post = _RE_IMPORTE.findall(plano[m_nomina.end():])
            if importes_post:
                valor = self._limpiar_numero(importes_post[0])
            else:
                importes = _RE_IMPORTE.findall(plano)
                valor = self._limpiar_numero(importes[-1]) if importes else None
        else:
            # Candidato: último importe monetario (con separador de miles/decimales)
            # tras el documento, que es la columna "Vr. Pago" en este layout. Exigir
            # el separador evita tomar identificadores/consecutivos como valor.
            importes = [t for t in _RE_IMPORTE.findall(plano[m_doc.end():]) if re.search(r"[.,]", t)]
            valor = self._limpiar_numero(importes[-1]) if importes else None
        if valor is None:
            return None

        # 4) Cuenta/producto (opcional): primer grupo largo tras el documento.
        cola = plano[m_doc.end():]
        m_cta = _RE_PRODUCTO.search(cola)
        cuenta = m_cta.group(1) if m_cta else None

        # 5) Fecha de la factura (referencia de la quincena): token antes de "PAGO".
        fecha = None
        m_fact = _RE_FACTURA.search(plano)
        if m_fact:
            factura = m_fact.group(1)
            formato = "%Y%m%d" if len(factura) == 8 else "%y%m%d"
            parsed = pd.to_datetime(factura, format=formato, errors="coerce")
            fecha = None if pd.isna(parsed) else parsed

        return {
            "Cuenta": cuenta,
            "Tipo": "ITALCO",
            "Documento": documento,
            "Nombre": nombre,
            "Valor": valor,
            "Fecha": fecha,
            "EsNomina": es_nomina,
        }

    def _process_transferencia_italco(self, folder_path):
        registros = []
        total_lineas_pago = 0

        for filename in os.listdir(folder_path):
            if not filename.lower().endswith(".pdf"):
                continue

            path = os.path.join(folder_path, filename)
            filas_archivo = 0

            with pdfplumber.open(path) as pdf:
                for page_num, page in enumerate(pdf.pages, start=1):
                    texto = page.extract_text() or ""
                    if not texto:
                        self._log(f"[transfer][{filename}] página {page_num} sin texto extraíble (¿escaneo sin OCR?).")
                        continue

                    # Reconstruir líneas por palabras/coordenadas: separa columnas
                    # contiguas que extract_text() pega (p. ej. documento+producto).
                    page_rows = []
                    for linea in self._lineas_desde_palabras(page):
                        data = self._match_linea_transferencia(linea)
                        if data:
                            page_rows.append(data)

                    if page_rows:
                        registros.extend(page_rows)
                        filas_archivo += len(page_rows)
                        total_lineas_pago += len(page_rows)
                        continue

                    # Fallback: soporte tipo desprendible (una persona por página),
                    # con la cédula tras "CC:" y el neto tras "Total Neto:".
                    texto_plano = re.sub(r"\s+", " ", texto)
                    cc_match = re.search(r"CC:\s*([\d.]+)", texto_plano, re.IGNORECASE)
                    neto_match = re.search(r"Total Neto:\s*([\d,\.]+)", texto_plano, re.IGNORECASE)
                    if cc_match and neto_match:
                        cc = re.sub(r"[^\d]", "", cc_match.group(1))
                        neto = self._limpiar_numero(neto_match.group(1))
                        registros.append(
                            {
                                "Cuenta": None,
                                "Tipo": "ITALCO",
                                "Documento": cc,
                                "Nombre": None,
                                "Valor": neto,
                                "Fecha": None,
                            }
                        )
                        filas_archivo += 1
                        total_lineas_pago += 1
                    else:
                        self._log(
                            f"[transfer][{filename}] página {page_num}: sin renglones de "
                            f"nómina ni patrón de respaldo (CC/Total Neto)."
                        )

            self._log(f"[transfer][{filename}] {filas_archivo} valor(es) de transferencia extraído(s).")

        df = pd.DataFrame(registros)
        self._log(f"[transfer] Total transferencias extraídas: {total_lineas_pago} en {len(df)} registros.")

        if not df.empty and "Valor" in df.columns:
            def _to_int(v):
                try:
                    if v is None:
                        return None
                    num = self._limpiar_numero(str(v))
                    return int(num) if num is not None else None
                except Exception:
                    return None

            df["Valor"] = df["Valor"].apply(_to_int)
            df["Valor"] = df["Valor"].fillna(0).astype("int64")

        if not df.empty:
            dedupe_cols = [col for col in ("Documento", "Cuenta", "Valor", "Fecha") if col in df.columns]
            if dedupe_cols:
                df = df.drop_duplicates(subset=dedupe_cols)

        return df

    def procesar_seguridad_social(self, folder_path, formato="tabarca"):
        """Procesa PDFs de seguridad social (IBC) según el formato indicado.

        Args:
            folder_path (str): ruta carpeta PDFs
            formato (str): "tabarca" o "italco"

        Returns:
            pd.DataFrame con columnas [archivo, cc, ibc]
        """
        formato = (formato or "tabarca").strip().lower()
        if formato == "italco":
            return self._procesar_seguridad_social_italco(folder_path)
        return self._procesar_seguridad_social_tabarca(folder_path)

    def _procesar_seguridad_social_tabarca(self, folder_path):
        """
        Procesa PDFs de seguridad social y extrae:
        - CC
        - IBC únicos (solo el primer valor de la columna IBC por página)

        Args:
            folder_path (str): ruta carpeta PDFs

        Returns:
            pd.DataFrame
        """

        registros = []

        def _extraer_ibc_desde_tabla(tabla):
            """Obtiene todos los valores distintos de la columna IBC dentro de una tabla extraída."""
            if not tabla:
                return []

            columna_ibc = None
            fila_inicio = None

            for idx_fila, fila in enumerate(tabla):
                if not fila:
                    continue
                for idx_col, celda in enumerate(fila):
                    if celda and re.search(r"\bIBC\b", str(celda), re.IGNORECASE):
                        columna_ibc = idx_col
                        fila_inicio = idx_fila
                        break
                if columna_ibc is not None:
                    break

            if columna_ibc is None:
                return []

            ibc_encontrados = []
            vistos = set()

            for fila in tabla[fila_inicio + 1:]:
                if not fila or columna_ibc >= len(fila):
                    continue

                celda = fila[columna_ibc]
                if not celda:
                    continue

                match_ibc = re.search(r"\$?\s*[\d\.,]+", str(celda))
                if not match_ibc:
                    continue

                ibc = re.sub(r"[^\d]", "", match_ibc.group(0))
                if ibc and ibc not in vistos:
                    vistos.add(ibc)
                    ibc_encontrados.append(ibc)

            return ibc_encontrados

        for filename in os.listdir(folder_path):
            if not filename.lower().endswith(".pdf"):
                continue

            path = os.path.join(folder_path, filename)

            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:

                    texto = page.extract_text()
                    if not texto:
                        continue

                    # =========================
                    # 1. EXTRAER CC (más flexible: permite puntos/comas)
                    # =========================
                    match_cc = re.search(r'CC\s*[:\.-]?\s*([\d\.,]+)', texto, re.IGNORECASE)
                    if not match_cc:
                        # fallback: any 6+ digit number
                        match_cc = re.search(r'\b(\d{6,})\b', texto)
                    if not match_cc:
                        continue

                    cc_raw = match_cc.group(1)
                    cc = re.sub(r"[^\d]", "", cc_raw)

                    # =========================
                    # 2. EXTRAER IBC DESDE LA POSICIÓN DE LA COLUMNA EN LA TABLA
                    # =========================
                    ibc_set = set()

                    tablas = []
                    try:
                        tablas = page.extract_tables() or []
                    except Exception:
                        tablas = []

                    for tabla in tablas:
                        for ibc in _extraer_ibc_desde_tabla(tabla):
                            ibc_set.add(ibc)

                    # Fallback: si no se pudo leer la tabla, usar el texto plano
                    if not ibc_set:
                        lines = texto.splitlines()
                        for idx, line in enumerate(lines):
                            if re.search(r'\bIBC\b', line, re.IGNORECASE):
                                for follow in lines[idx: idx + 12]:
                                    if not follow or follow.strip() == "":
                                        continue
                                    match_ibc = re.search(r'\$?\s*[\d\.,]+', follow)
                                    if match_ibc:
                                        valor = match_ibc.group(0)
                                        ibc = re.sub(r"[^\d]", "", valor)
                                        if ibc:
                                            ibc_set.add(ibc)
                                

                    # =========================
                    # 4. GUARDAR RESULTADOS
                    # =========================
                    for ibc in ibc_set:
                        registros.append({
                            "archivo": filename,
                            "cc": cc,
                            "ibc": ibc
                        })

                    

        df = pd.DataFrame(registros)

        if not df.empty:
            df = df.drop_duplicates(subset=["cc", "ibc"])

        return df

    def _procesar_seguridad_social_italco(self, folder_path):
        """
        Procesa la "Planilla Resumen" (aportes en línea) en formato ITALCO y
        extrae el documento y el IBC de pensión por fila.

        La columna del IBC depende del layout de la página: en la primera página
        las filas de personas empiezan en el índice 13 y el IBC está en la
        columna 26; en las páginas siguientes las filas válidas tienen 43 celdas
        y el IBC está en la columna 27.

        Returns:
            pd.DataFrame con columnas [archivo, cc, ibc]
        """
        registros = []

        for filename in os.listdir(folder_path):
            if not filename.lower().endswith(".pdf"):
                continue

            path = os.path.join(folder_path, filename)
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:
                    tabla = page.extract_table()
                    if not tabla:
                        continue

                    if page.page_number == 1:
                        filas = tabla[13:]
                        ibc_idx = 26
                    else:
                        filas = [fila for fila in tabla if len(fila) == 43]
                        ibc_idx = 27

                    for fila in filas:
                        if len(fila) <= ibc_idx:
                            continue

                        doc_raw = fila[2]
                        ibc_raw = fila[ibc_idx]
                        if not doc_raw or not ibc_raw:
                            continue

                        cc = re.sub(r"[^\d]", "", str(doc_raw))
                        ibc = self._limpiar_numero(ibc_raw)
                        if not cc or ibc is None:
                            continue

                        registros.append({
                            "archivo": filename,
                            "cc": cc,
                            "ibc": int(ibc),
                        })

        df = pd.DataFrame(registros)

        if not df.empty:
            df = df.drop_duplicates(subset=["cc", "ibc"])

        return df

    def _parsear_linea(self, linea):
        """
        Parse a single line from transfer PDF.
        
        Args:
            linea (str): Line from transfer PDF
            
        Returns:
            dict: Parsed transfer data or None if parsing fails
        """
        partes = linea.split()
        
        if len(partes) < 4:
            return None
        
        cuenta = partes[0]
        tipo_cuenta = partes[1]
        documento = partes[2]
        
        # Find monetary value (formatted as 1,234,567.89)
        valor_idx = None
        for i, p in enumerate(partes):
            if re.match(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?", p):
                valor_idx = i
                break
        
        if valor_idx is None:
            return None
        
        valor = partes[valor_idx]
        
        # Extract name (between document and value)
        nombre = " ".join(partes[3:valor_idx]) if valor_idx > 3 else ""
        
        # Find date (DD-MM-YYYY format)
        fecha = None
        for p in reversed(partes):
            if re.match(r"\d{2}-\d{2}-\d{4}", p):
                fecha = p
                break
        
        return {
            "Cuenta": cuenta,
            "Tipo": tipo_cuenta,
            "Documento": documento,
            "Nombre": nombre,
            "Valor": valor,
            "Fecha": fecha
        }
    
    def _limpiar_numero(self, valor):
        """
        Limpia y convierte una cadena numérica a float.
        
        Args:
            valor (str): Representación en texto del número
            
        Returns:
            float: Valor numérico limpio
        """
        if not valor: 
            return None
        valor = valor.replace("$", "").strip()
        if "," in valor and "." in valor:
            if valor.rfind(",") > valor.rfind("."):
                valor = valor.replace(".", "").replace(",", ".")
            else:
                valor = valor.replace(",", "")
        elif "," in valor:
            partes = valor.split(",")
            if len(partes[-1]) == 2:
                valor = valor.replace(".", "").replace(",", ".")
            else:
                valor = valor.replace(",", "")
        elif "." in valor:
            partes = valor.split(".")
            if len(partes[-1]) == 2:
                valor = valor.replace(",", "")
            elif valor.count(".") > 1:
                valor = valor.replace(".", "")
        valor = valor.strip()
        try:
            return float(valor)
        except:
            return None
    
    def formatear_valores(self, cadena):
        """
        Formatea valores numéricos al formato de moneda colombiana (COP).
        
        Convierte representaciones de texto de números en cadenas monetarias formateadas
        con símbolo $, separador de miles (coma) y separador decimal (punto).
        Los múltiples valores se separan con el carácter barra vertical (|).
        
        Args:
            cadena (str or None): Cadena de entrada con números.
                Puede contener corchetes, varios números o None.
                Formato: "[1000000 2000000]" o "1000000 2000000"
        
        Returns:
            str: Cadena formateada en COP.
                Ejemplo: "$1.000.000,00 | $2.000.000,00"
                Devuelve una cadena vacía si la entrada es None o está vacía.
        
        Proceso:
            1. Valida la entrada (devuelve "" si está vacía)
            2. Elimina corchetes y espacios extra
            3. Divide múltiples valores por espacios
            4. Convierte cada valor a float y lo formatea como moneda
            5. Une múltiples valores con el separador " | "
        
        Ejemplo:
            >>> app.formatear_valores("[1000000 2500000]")
            "$1.000.000,00 | $2.500.000,00"
            
            >>> app.formatear_valores(None)
            ""
        """
        # Paso 1: validar la entrada - devolver cadena vacía si está vacía
        if not cadena:
            return ""
        
        
        # Paso 2: limpiar la cadena de entrada - quitar corchetes y espacios
        cadena = str(cadena).replace("[", "").replace("]", "").strip()
        
        # Paso 3: dividir en números individuales
        numeros = cadena.split()
        
        # Paso 4: convertir cada número a moneda formateada
        resultado = []
        for num in numeros:
            num = num.replace(",", "").strip()  # Eliminar comas existentes

            try:
                # Convertir la cadena a float
                valor = float(num)
                
                # Formatear como moneda: ${value:,.2f}
                # Ejemplo: $1000000.00 -> $1,000,000.00
                formato = "${:,.2f}".format(valor)
                
                # Convertir al formato colombiano: reemplazar comas temporalmente,
                # luego intercambiar coma y punto para obtener el formato COP
                # $1,000,000.00 → $1.000.000,00 (COP format)
                formato = formato.replace(",", "X")      # $1X000X000.00
                formato = formato.replace(".", ",")      # $1X000X000,00
                formato = formato.replace("X", ".")      # $1.000.000,00
                
                # Agregar el valor formateado al resultado
                resultado.append(formato)
                
            except (ValueError, TypeError):
                # Omitir valores que no se puedan convertir a float
                continue
        
        # Paso 5: unir múltiples valores con separador de barra vertical
        # Example: ["$1.000.000,00", "$2.000.000,00"] → "$1.000.000,00 | $2.000.000,00"
        return " | ".join(resultado)
    
    def _limpiar_doc(self, col):
        """
        Limpia números de documento o identificación.
        
        Args:
            col (pd.Series): Serie de Pandas para limpiar
            
        Returns:
            pd.Series: Números de documento limpios
        """
        return (
            col.astype(str)
            .str.replace(r"\.0$", "", regex=True)  # Eliminar .0
            .str.replace(r"[^\d]", "", regex=True)  # Dejar solo dígitos
            .str.lstrip("0")  # Eliminar ceros a la izquierda
            .str.strip()
        )

    def _reconcile_data(self, df_desprendibles, df_transferencia, df_seguridad=None):
        """
        Concilia los datos de desprendibles y transferencias e incorpora Devengado e IBC.
        """
        def _normalizar_numero(valor):
            if valor is None:
                return None
            try:
                if pd.isna(valor):
                    return None
            except Exception:
                pass

            try:
                numero = float(valor)
            except Exception:
                return None

            if numero.is_integer():
                return int(numero)
            return numero

        def _normalizar_lista(valores):
            normalizados = []
            vistos = set()
            for valor in valores or []:
                numero = _normalizar_numero(valor)
                if numero is None or numero in vistos:
                    continue
                vistos.add(numero)
                normalizados.append(numero)
            return normalizados

        def _normalizar_lista_completa(valores):
            """Como ``_normalizar_lista`` pero **sin deduplicar**.

            Para seguridad social los devengados/IBC pueden repetirse entre
            quincenas (dos quincenas con el mismo valor) y deben contarse ambas.
            """
            return [
                numero
                for numero in (_normalizar_numero(v) for v in (valores or []))
                if numero is not None
            ]

        resultados = []
        resultados_transfers = []
        resultados_seguridad = []

        # Limpiar números de documento
        if df_desprendibles is None or df_desprendibles.empty:
            return pd.DataFrame(resultados)

        def _clave_cuenta(valor):
            """Clave de cuenta robusta: solo dígitos y sin ceros a la izquierda.

            Permite cruzar la cuenta del desprendible contra el número de producto
            del soporte aunque difieran en ceros de relleno o en formato.
            """
            if valor is None:
                return ""
            try:
                if pd.isna(valor):
                    return ""
            except (TypeError, ValueError):
                pass
            return re.sub(r"\D", "", str(valor)).lstrip("0")

        df_desprendibles["Identificacion"] = self._limpiar_doc(df_desprendibles["Identificacion"])

        hay_transferencias = df_transferencia is not None and not df_transferencia.empty
        if hay_transferencias:
            df_transferencia["Documento"] = self._limpiar_doc(df_transferencia["Documento"])
            # Clave de cuenta normalizada para el cruce por producto/cuenta.
            df_transferencia["_cuenta_key"] = (
                df_transferencia["Cuenta"].map(_clave_cuenta)
                if "Cuenta" in df_transferencia.columns
                else ""
            )
            self._log(
                f"[reconcile] {len(df_transferencia)} transferencias disponibles; "
                f"{df_transferencia['Documento'].nunique()} documento(s) distinto(s)."
            )
        else:
            self._log("[reconcile] No se recibieron transferencias para cruzar.")

        if df_seguridad is not None and not df_seguridad.empty:
            df_seguridad["cc"] = (
                df_seguridad["cc"].astype(str).str.replace(r"[^\d]", "", regex=True).str.lstrip("0")
            )

        # Agrupar desprendibles por identificación y conciliar con transferencias
        for doc, grupo_despr in df_desprendibles.groupby("Identificacion"):
            netos = _normalizar_lista(grupo_despr["Neto"].dropna().tolist())
            suma_netos = _normalizar_numero(sum(netos)) if netos else None
            # Devengado: sumar TODOS los valores SIN deduplicar. Dos quincenas con
            # el mismo devengado deben contarse ambas (a diferencia del cruce de
            # transferencias, donde sí se deduplica).
            devs = _normalizar_lista_completa(grupo_despr["Devengado"].dropna().tolist()) if "Devengado" in grupo_despr else []
            sum_devs = _normalizar_numero(sum(devs)) if devs else None

            cta = grupo_despr["Cuenta"].iloc[0] if "Cuenta" in grupo_despr.columns else None
            cta_key = _clave_cuenta(cta)

            # Ventana de periodo de esta persona, tomada de SUS desprendibles
            # ("Periodo: inicio al fin"). Sirve para descartar transferencias de
            # otras quincenas/meses que comparten el mismo documento.
            win_ini = win_fin = None
            if "PeriodoInicio" in grupo_despr.columns and "PeriodoFin" in grupo_despr.columns:
                inis = pd.to_datetime(grupo_despr["PeriodoInicio"], errors="coerce").dropna()
                fins = pd.to_datetime(grupo_despr["PeriodoFin"], errors="coerce").dropna()
                if not inis.empty and not fins.empty:
                    win_ini, win_fin = inis.min(), fins.max()

            # Buscar transferencias coincidentes (por documento o por cuenta) y
            # comparar por suma. El cruce por documento es el principal; el de
            # cuenta es un respaldo para soportes donde el documento difiera.
            grupo_trans = pd.DataFrame()
            estado_trans = "Transferencia no encontrada"
            valores_trans = []
            suma_trans = None
            if hay_transferencias:
                por_documento = df_transferencia["Documento"] == doc
                por_cuenta = (
                    (df_transferencia["_cuenta_key"] == cta_key) & (cta_key != "")
                    if "_cuenta_key" in df_transferencia.columns
                    else False
                )
                grupo_trans = df_transferencia[por_documento | por_cuenta]

            # Separar transferencias confiables (con etiqueta NÓMINA) de las
            # candidatas (otro layout sin etiqueta ni fecha-factura de quincena).
            # Si no hay columna 'EsNomina' (formato TABARCA o datos antiguos), se
            # tratan todas como confiables -> comportamiento previo intacto.
            if not grupo_trans.empty and "EsNomina" in grupo_trans.columns:
                es_nom = grupo_trans["EsNomina"].fillna(True).astype(bool)
                trans_confiables = grupo_trans[es_nom]
                trans_candidatas = grupo_trans[~es_nom]
            else:
                trans_confiables = grupo_trans
                trans_candidatas = grupo_trans.iloc[0:0]

            # Filtrar por periodo SOLO las confiables: conservar las que caen dentro
            # de la ventana de los desprendibles (las sin fecha legible se conservan).
            # Evita sumar quincenas/meses ajenos al soporte.
            if (
                not trans_confiables.empty
                and win_ini is not None
                and win_fin is not None
                and "Fecha" in trans_confiables.columns
            ):
                fechas = pd.to_datetime(trans_confiables["Fecha"], errors="coerce")
                en_ventana = fechas.isna() | ((fechas >= win_ini) & (fechas <= win_fin))
                descartadas = int((~en_ventana).sum())
                if descartadas:
                    self._log(
                        f"[reconcile] doc={doc}: {descartadas} transferencia(s) fuera del "
                        f"periodo [{win_ini.date()}..{win_fin.date()}] descartada(s)."
                    )
                trans_confiables = trans_confiables[en_ventana]

            # Candidatas: solo valen si su valor coincide con un neto del desprendible
            # que aún no haya sido cubierto por una transferencia confiable. Así se
            # rescata el pago real (mismo valor que el neto) aunque el renglón no
            # traiga etiqueta/fecha, y se descartan importes de otra quincena.
            candidatas_validas = trans_candidatas.iloc[0:0]
            if not trans_candidatas.empty:
                val_conf = _normalizar_lista(trans_confiables["Valor"].dropna().tolist()) if not trans_confiables.empty else []
                pendientes = Counter(int(round(float(n))) for n in netos if pd.notna(n))
                pendientes.subtract(Counter(int(round(float(v))) for v in val_conf))
                idx_keep = []
                for idx_c, val_c in trans_candidatas["Valor"].dropna().items():
                    clave = int(round(float(val_c)))
                    if pendientes.get(clave, 0) > 0:
                        idx_keep.append(idx_c)
                        pendientes[clave] -= 1
                    else:
                        self._log(
                            f"[reconcile] doc={doc}: transferencia candidata {val_c} sin neto "
                            f"que la respalde -> descartada (posible otra quincena)."
                        )
                candidatas_validas = trans_candidatas.loc[idx_keep]

            grupo_trans = pd.concat([trans_confiables, candidatas_validas])

            if not grupo_trans.empty:
                valores_trans = _normalizar_lista(grupo_trans["Valor"].dropna().tolist())
                # Sumar valores y comparar con la suma de netos.
                try:
                    suma_trans = _normalizar_numero(sum(valores_trans)) if valores_trans else None
                except Exception:
                    suma_trans = None

                if suma_netos is not None and suma_trans is not None and suma_netos == suma_trans:
                    estado_trans = "OK"
                else:
                    estado_trans = "Valor no coincide"
                    self._log(
                        f"[reconcile] doc={doc}: transferencia encontrada pero la suma no "
                        f"coincide (netos={suma_netos} vs transferencias={suma_trans})."
                    )
            else:
                valores_trans = None
                if hay_transferencias:
                    # Log diagnóstico: la transferencia existe en el universo pero no
                    # cruzó por documento ni por cuenta para esta persona.
                    self._log(
                        f"[reconcile] doc={doc} (cuenta={cta_key or '—'}): sin transferencia "
                        f"que cruce por documento ni por cuenta -> 'Transferencia no encontrada'."
                    )

            # Obtener IBCs desde df_seguridad si está disponible. NO se deduplican:
            # se suman todos (igual que el cruce de transferencias).
            ibc_vals = []
            if df_seguridad is not None and not df_seguridad.empty:
                matches = df_seguridad[df_seguridad["cc"] == doc]
                if not matches.empty:
                    ibc_vals = _normalizar_lista_completa(matches["ibc"].dropna().tolist())
            suma_ibc = _normalizar_numero(sum(ibc_vals)) if ibc_vals else None

            # Estado seguridad social: comparar SUMA de devengados vs SUMA de IBC.
            if sum_devs is None:
                estado_seg = "Devengado no encontrado"
            elif suma_ibc is not None and sum_devs == suma_ibc:
                estado_seg = "OK"
            else:
                estado_seg = "Devengado no coincide"

            # Diferencia = devengado(desprendible) - IBC(seguridad). Lo ausente cuenta como 0.
            diferencia_seg = (
                _normalizar_numero((sum_devs or 0) - (suma_ibc or 0))
                if sum_devs is not None else None
            )
            # Diferencia = neto(desprendible) - transferencia. Lo ausente cuenta como 0.
            diferencia_trans = (
                _normalizar_numero((suma_netos or 0) - (suma_trans or 0))
                if (suma_netos is not None or suma_trans is not None) else None
            )

            # Agregar resultado para transferencias (usa suma para OK)
            resultados_transfers.append({
                "Identificación": doc,
                "Cuenta": cta,
                "Estado": estado_trans,
                "Neto_desprendibles": list(netos),
                "Valores_transferencia": list(valores_trans) if valores_trans is not None else None,
                "Diferencia": diferencia_trans,
            })

            # Agregar resultado para seguridad social (IBC)
            resultados_seguridad.append({
                "Identificación": doc,
                "Estado": estado_seg,
                "Devengado": sum_devs,
                "IBC": ibc_vals if ibc_vals else None,
                "Diferencia": diferencia_seg,
            })
        df_t = pd.DataFrame(resultados_transfers)
        df_s = pd.DataFrame(resultados_seguridad)

        return df_t, df_s

    def _display_results(self):
        """
        Muestra los resultados de la conciliación en la tabla de la interfaz (Treeview).
        """
        # Limpiar ambas tablas
        for t in (self.tree_transfers, self.tree_seguridad):
            for item in t.get_children():
                t.delete(item)

        # Validar que existan resultados
        if (not hasattr(self, 'df_resultado_transfers') or self.df_resultado_transfers is None or self.df_resultado_transfers.empty) and (
            not hasattr(self, 'df_resultado_seguridad') or self.df_resultado_seguridad is None or self.df_resultado_seguridad.empty):
            messagebox.showinfo("Información", "No hay resultados para mostrar")
            return

        # Insertar filas en la tabla de transferencias
        if hasattr(self, 'df_resultado_transfers') and self.df_resultado_transfers is not None:
            for idx, row in self.df_resultado_transfers.iterrows():
                estado = str(row["Estado"] or "")
                estado_norm = estado.lower()
                if "transferencia no encontrada" in estado_norm or "valor no coincide" in estado_norm:
                    tag = "error"
                elif estado_norm == "ibc sin soporte":
                    tag = "warning"
                elif estado_norm.strip() == "ok":
                    tag = "ok"
                else:
                    tag = "error"

                neto_display = self.formatear_valores(row.get("Neto_desprendibles", ""))
                if row.get("Valores_transferencia") is None:
                    valor_display = ""
                else:
                    valor_display = self.formatear_valores(row.get("Valores_transferencia"))

                self.tree_transfers.insert(
                    "",
                    tk.END,
                    values=(
                        row.get("Identificación"),
                        row.get("Cuenta"),
                        row.get("Estado"),
                        neto_display,
                        valor_display,
                    ),
                    tags=(tag,)
                )

        # Insertar filas en la tabla de seguridad social (IBC)
        if hasattr(self, 'df_resultado_seguridad') and self.df_resultado_seguridad is not None:
            for idx, row in self.df_resultado_seguridad.iterrows():
                estado = str(row["Estado"] or "")
                estado_norm = estado.lower()
                if "devengado no encontrado" in estado_norm or "devengado no coincide" in estado_norm:
                    tag = "error"
                elif estado_norm == "ibc sin soporte":
                    tag = "warning"
                elif estado_norm.strip() == "ok":
                    tag = "ok"
                else:
                    tag = "error"

                deveng_display = self.formatear_valores(row.get("Devengado", ""))
                ibc_display = self.formatear_valores(row.get("IBC", ""))

                self.tree_seguridad.insert(
                    "",
                    tk.END,
                    values=(
                        row.get("Identificación"),
                        row.get("Estado"),
                        deveng_display,
                        ibc_display,
                    ),
                    tags=(tag,)
                )
    
    def _export_to_csv(self):
        """Exporta los resultados a un archivo CSV."""
        df_t = getattr(self, 'df_resultado_transfers', None)
        df_s = getattr(self, 'df_resultado_seguridad', None)

        if (df_t is None or df_t.empty) and (df_s is None or df_s.empty):
            messagebox.showwarning("Aviso", "No hay resultados para exportar. Primero procesa los PDFs.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")]
        )

        if not file_path:
            return

        try:
            parts = []
            if df_t is not None and not df_t.empty:
                df_tmp = df_t.copy()
                df_tmp["Devengado"] = None
                df_tmp["IBC"] = None
                df_tmp["Tipo"] = "Transferencia"
                parts.append(df_tmp)
            if df_s is not None and not df_s.empty:
                df_tmp2 = df_s.copy()
                df_tmp2["Cuenta"] = None
                df_tmp2["Neto_desprendibles"] = None
                df_tmp2["Valores_transferencia"] = None
                df_tmp2["Tipo"] = "Seguridad"
                parts.append(df_tmp2)

            df_out = pd.concat(parts, ignore_index=True, sort=False)
            df_out.to_csv(file_path, index=False)
            messagebox.showinfo("Éxito", f"Resultados exportados en:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"La exportación falló:\n{str(e)}")
    
    def _clear_results(self):
        """Limpia todos los resultados."""
        for t in (getattr(self, 'tree_transfers', None), getattr(self, 'tree_seguridad', None)):
            if t is None:
                continue
            for item in t.get_children():
                t.delete(item)
        self.df_resultado = None
        self.df_resultado_transfers = None
        self.df_resultado_seguridad = None
        self.status_label.config(text="Resultados limpiados", foreground="blue")


def main():
    """Inicializa y ejecuta la aplicación."""
    if tk is None:
        raise RuntimeError("Tkinter is not available in this environment.")
    root = tk.Tk()
    app = PayrollReconciliationApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
