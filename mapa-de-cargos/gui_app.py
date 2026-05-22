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

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import pdfplumber
import re
import os
from pathlib import Path
import threading


class PayrollReconciliationApp:
    """
    Aplicación principal GUI para la conciliación de datos de nómina.
    
    Attributes:
        root (tk.Tk): Main application window
        desprendibles_path (tk.StringVar): Path to payslips folder
        transferencia_path (tk.StringVar): Path to transfers folder
        df_resultado (pd.DataFrame): Final reconciliation results
    """
    
    def __init__(self, root):
        """Inicializa la aplicación."""
        self.root = root
        self.root.title("Sistema de Conciliación de Nómina")
        self.root.geometry("1400x800")
        self.root.configure(bg="#f0f0f0")
        
        # Variables para almacenar las rutas seleccionadas
        self.desprendibles_path = tk.StringVar()
        self.transferencia_path = tk.StringVar()
        self.seguridad_path = tk.StringVar()
        
        # DataFrame de resultados
        self.df_resultado = None
        
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
        
        # Crear Treeview para mostrar resultados
        self.tree = ttk.Treeview(
            results_frame,
            columns=("Identificación", "Cuenta", "Estado", "Neto_desprendibles", "Valores_transferencia", "Devengado", "IBC"),
            height=20,
            show="headings"
        )
        
        # Define column headings and widths
        self.tree.heading("Identificación", text="Identificación")
        self.tree.heading("Cuenta", text="Cuenta")
        self.tree.heading("Estado", text="Estado")
        self.tree.heading("Neto_desprendibles", text="Valores desprendibles")
        self.tree.heading("Valores_transferencia", text="Valores de transferencia")
        self.tree.heading("Devengado", text="Devengado")
        self.tree.heading("IBC", text="IBC")
        
        self.tree.column("Identificación", width=150, anchor=tk.CENTER)
        self.tree.column("Cuenta", width=150, anchor=tk.CENTER)
        self.tree.column("Estado", width=200, anchor=tk.CENTER)
        self.tree.column("Neto_desprendibles", width=250, anchor=tk.W)
        self.tree.column("Valores_transferencia", width=250, anchor=tk.W)
        self.tree.column("Devengado", width=200, anchor=tk.W)
        self.tree.column("IBC", width=200, anchor=tk.W)
        
        # Agregar barras de desplazamiento
        vsb = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(results_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscroll=vsb.set, xscroll=hsb.set)
        
        # Distribución del Treeview y las barras
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        results_frame.grid_rowconfigure(0, weight=1)
        results_frame.grid_columnconfigure(0, weight=1)
        
        # Configurar etiquetas de color
        self.tree.tag_configure("error", background="#ffcccc", foreground="darkred")
        self.tree.tag_configure("ok", background="#ccffcc", foreground="darkgreen")
        self.tree.tag_configure("warning", background="#ffffcc", foreground="darkorange")
        
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
            self.df_resultado = self._reconcile_data(df_desprendibles, df_transferencia, df_seguridad)
            
            # Actualizar la interfaz en el hilo principal
            self.root.after(0, self._display_results)
            self.root.after(0, lambda: self.status_label.config(
                text=f"Proceso completado. Se encontraron {len(self.df_resultado)} registros.",
                foreground="green"
            ))
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"El procesamiento falló:\n{str(e)}"))
            self.root.after(0, lambda: self.status_label.config(
                text="Procesamiento fallido",
                foreground="red"
            ))
    
    def _process_desprendibles(self, folder_path):
        """
        Extrae datos de desprendibles desde archivos PDF.
        
        Args:
            folder_path (str): Ruta de la carpeta con PDFs de desprendibles
            
        Returns:
            pd.DataFrame: DataFrame con las columnas [Identificacion, Neto, Cuenta]
        """
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
                            deven_match = re.search(
                                r"Devengad[o|os]?[:\s].*?\$\s*([\d\.,]+)",
                                bloque,
                                re.IGNORECASE | re.DOTALL
                            )
                            if not deven_match:
                                deven_match = re.search(
                                    r"TOTALES[:\s].*?\$\s*([\d\.,]+)",
                                    bloque,
                                    re.IGNORECASE | re.DOTALL
                                )

                            
                            if id_match and neto_match:
                                identificacion = id_match.group()
                                identificacion = identificacion.replace(".", "").replace(",", "")
                                neto = self._limpiar_numero(neto_match.group(1))
                                devengado = self._limpiar_numero(deven_match.group(1)) if deven_match else None
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
    
    def _process_transferencia(self, folder_path):
        """
        Extract transfer data from PDF files.
        
        Args:
            folder_path (str): Path to folder containing transfer PDFs
            
        Returns:
            pd.DataFrame: DataFrame with transfer information
        """
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

    def procesar_seguridad_social(self, folder_path):
        """
        Procesa PDFs de seguridad social y extrae:
        - CC
        - IBC únicos (solo el primer valor monetario por fila)

        Args:
            folder_path (str): ruta carpeta PDFs

        Returns:
            pd.DataFrame
        """

        registros = []

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
                    # 2. UBICAR LÍNEAS DONDE ESTÁN LOS IBC
                    # =========================
                    lines = texto.splitlines()
                    ibc_set = set()
                    for idx, line in enumerate(lines):
                        if re.search(r'\bIBC\b', line, re.IGNORECASE):
                            # scan following few lines for monetary values
                            for follow in lines[idx: idx + 12]:
                                if not follow or follow.strip() == "":
                                    continue
                                match_ibc = re.search(r'\$?\s*[\d\.,]+', follow)
                                if match_ibc:
                                    valor = match_ibc.group(0)
                                    numero = self._limpiar_numero(valor)
                                    if numero and numero > 0:
                                        ibc_set.add(int(numero))

                    # =========================
                    # 4. GUARDAR RESULTADOS
                    # =========================
                    for ibc in ibc_set:
                        registros.append({
                            "archivo": filename,
                            "cc": cc,
                            "ibc": ibc
                        })
                        print(cc, ibc)

                    

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
        resultados = []

        # Limpiar números de documento
        if df_desprendibles is None or df_desprendibles.empty:
            return pd.DataFrame(resultados)

        df_desprendibles["Identificacion"] = self._limpiar_doc(df_desprendibles["Identificacion"])

        if df_transferencia is not None and not df_transferencia.empty:
            df_transferencia["Documento"] = self._limpiar_doc(df_transferencia["Documento"])

        if df_seguridad is not None and not df_seguridad.empty:
            df_seguridad["cc"] = (
                df_seguridad["cc"].astype(str).str.replace(r"[^\d]", "", regex=True).str.lstrip("0")
            )

        # Agrupar desprendibles por identificación y conciliar con transferencias
        for doc, grupo_despr in df_desprendibles.groupby("Identificacion"):
            netos = set(grupo_despr["Neto"])
            # Devengado: tomar valores únicos (puede haber varios)
            devs = list(grupo_despr["Devengado"].dropna().unique()) if "Devengado" in grupo_despr else []
            cta = grupo_despr["Cuenta"].iloc[0] if "Cuenta" in grupo_despr.columns else None

            # Buscar transferencias coincidentes
            grupo_trans = pd.DataFrame()
            if df_transferencia is not None and not df_transferencia.empty:
                grupo_trans = df_transferencia[
                    (df_transferencia["Documento"] == doc) |
                    (df_transferencia["Cuenta"] == cta)
                ]

            # Caso 1: documento no encontrado
            if grupo_trans.empty:
                valores_trans = None
                estado = "Documento o Cuenta no encontrado"
            else:
                valores_trans = set(grupo_trans["Valor"])
                if len(netos.intersection(valores_trans)) > 0:
                    estado = "OK"
                else:
                    estado = "Valor no coincide"

            # Obtener IBCs desde df_seguridad si está disponible
            ibc_vals = None
            if df_seguridad is not None and not df_seguridad.empty:
                matches = df_seguridad[df_seguridad["cc"] == doc]
                if not matches.empty:
                    ibc_vals = list(matches["ibc"].unique())

            resultados.append({
                "Identificación": doc,
                "Cuenta": cta,
                "Estado": estado,
                "Neto_desprendibles": list(netos),
                "Valores_transferencia": list(valores_trans) if valores_trans is not None else None,
                "Devengado": devs,
                "IBC": ibc_vals,
            })

        return pd.DataFrame(resultados)

    def _display_results(self):
        """
        Muestra los resultados de la conciliación en la tabla de la interfaz (Treeview).
        """
        # Paso 1: limpiar los elementos existentes de la tabla
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Paso 2: validar que existan resultados
        if self.df_resultado is None or self.df_resultado.empty:
            messagebox.showinfo("Información", "No hay resultados para mostrar")
            return

        # Paso 3: insertar cada fila en la tabla
        for idx, row in self.df_resultado.iterrows():
            # Determinar la etiqueta de color según el estado
            if row["Estado"] == "Documento o Cuenta no encontrado":
                tag = "error"
            elif row["Estado"] == "OK":
                tag = "ok"
            else:
                tag = "warning"

            # Formatear los valores para mostrar
            neto_display = self.formatear_valores(row["Neto_desprendibles"])
            if row["Valores_transferencia"] is None:
                valor_display = ""
            else:
                valor_display = self.formatear_valores(row["Valores_transferencia"])

            # Devengado and IBC
            deveng_display = self.formatear_valores(row.get("Devengado", ""))
            ibc_display = self.formatear_valores(row.get("IBC", ""))

            # Insertar la fila en la tabla con datos formateados
            self.tree.insert(
                "",
                tk.END,
                values=(
                    row["Identificación"],
                    row["Cuenta"],
                    row["Estado"],
                    neto_display,
                    valor_display,
                    deveng_display,
                    ibc_display,
                ),
                tags=(tag,)
            )
    
    def _export_to_csv(self):
        """Exporta los resultados a un archivo CSV."""
        if self.df_resultado is None or self.df_resultado.empty:
            messagebox.showwarning("Aviso", "No hay resultados para exportar. Primero procesa los PDFs.")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")]
        )
        
        if file_path:
            try:
                self.df_resultado.to_csv(file_path, index=False)
                messagebox.showinfo("Éxito", f"Resultados exportados en:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"La exportación falló:\n{str(e)}")
    
    def _clear_results(self):
        """Limpia todos los resultados."""
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.df_resultado = None
        self.status_label.config(text="Resultados limpiados", foreground="blue")


def main():
    """Inicializa y ejecuta la aplicación."""
    root = tk.Tk()
    app = PayrollReconciliationApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
