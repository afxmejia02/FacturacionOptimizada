"""
TECHNICAL DOCUMENTATION
Payroll Data Reconciliation System - GUI Application

================================================================================
TABLE OF CONTENTS
================================================================================
1. Architecture Overview
2. Class Structure
3. Data Processing Pipeline
4. Algorithm Details
5. GUI Components
6. Error Handling
7. Performance Considerations
8. Extension Points

================================================================================
1. ARCHITECTURE OVERVIEW
================================================================================

The application follows a Model-View-Controller (MVC) pattern:

┌─────────────────────────────────────────────────────────────────┐
│                         USER INTERFACE (View)                    │
│                  TkInter GUI with buttons, fields, table          │
└────────────────────────┬────────────────────────────────────────┘
                         │ User Interactions
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│                    APPLICATION LOGIC (Controller)                │
│              PayrollReconciliationApp class                       │
│     - Handles user input                                         │
│     - Coordinates file processing                                │
│     - Manages data flow                                          │
└────────────────────────┬────────────────────────────────────────┘
                         │ Process & Query
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│                         DATA (Model)                              │
│     - PDF files on disk                                          │
│     - DataFrame objects                                          │
│     - Reconciliation results                                     │
└─────────────────────────────────────────────────────────────────┘


Data Flow Diagram:
─────────────────

Payslips PDF ──┐
               ├──► Extract & Parse ──┐
Transfers PDF ─┤                      ├──► Clean Data ──┐
               │                      │                 │
               └──────────────────────┘                 ├──► Reconcile ──► Results Table
                                                        │
                                                        └──► Export CSV


================================================================================
2. CLASS STRUCTURE
================================================================================

PayrollReconciliationApp
    └── Main Application Controller

    Attributes:
        root: tk.Tk
            - Main window instance
            - Window configuration and event loop
        
        desprendibles_path: tk.StringVar
            - Stores path to payslips folder
            - Updates GUI when changed
        
        transferencia_path: tk.StringVar
            - Stores path to transfers folder
            - Updates GUI when changed
        
        df_resultado: pd.DataFrame
            - Final reconciliation results
            - Displayed in Treeview widget
        
        tree: ttk.Treeview
            - Table widget for displaying results
            - Supports color-coded rows

    Methods:
        __init__(root)
            Purpose: Initialize application
            Parameters: root (tk.Tk) - Main window
            Process:
                1. Set window properties
                2. Initialize variables
                3. Create UI elements
                4. Configure event handlers
        
        _create_ui()
            Purpose: Build all GUI components
            Components:
                - Header frame with title
                - Path selection section
                - Control buttons
                - Results table (Treeview)
                - Scrollbars
            Color tags:
                - "error": Red background for missing documents
                - "ok": Green background for matches
                - "warning": Yellow background for mismatches
        
        _select_desprendibles_folder()
            Purpose: Open folder browser for payslips
            Returns: Sets desprendibles_path variable
            
        _select_transferencia_folder()
            Purpose: Open folder browser for transfers
            Returns: Sets transferencia_path variable
        
        _process_files()
            Purpose: Orchestrate PDF processing
            Validation:
                - Check both paths are selected
                - Update status bar
            Execution: Spawns background thread
        
        _process_files_thread()
            Purpose: Background thread for processing
            Steps:
                1. Extract payslip data
                2. Extract transfer data
                3. Reconcile records
                4. Update UI with results
            Error handling: Try-except with messagebox
        
        _process_desprendibles(folder_path)
            Purpose: Extract data from payslip PDFs
            Input: Path to folder containing PDF files
            Process:
                1. Iterate through PDF files
                2. Extract text from each page
                3. Parse identification and net amount
                4. Clean and validate data
            Returns: pd.DataFrame with columns:
                - Identificacion (str): ID number without dots
                - Neto (int): Net payment amount
            
        _process_transferencia(folder_path)
            Purpose: Extract data from transfer PDFs
            Input: Path to folder containing PDF files
            Process:
                1. Iterate through PDF files
                2. Extract text from each page
                3. Parse lines starting with account numbers
                4. Extract transfer details
                5. Convert amounts to numeric format
            Returns: pd.DataFrame with columns:
                - Cuenta: Account number
                - Tipo: Account type
                - Documento: Document/ID number
                - Nombre: Beneficiary name
                - Valor: Transfer amount
                - Fecha: Transfer date
            
        _parsear_linea(linea)
            Purpose: Parse a single transfer record line
            Input: Line of text from PDF
            Logic:
                1. Split line by whitespace
                2. Extract account number (first element)
                3. Extract account type (second element)
                4. Extract document (third element)
                5. Find monetary value using regex
                6. Extract name (between doc and value)
                7. Find date in DD-MM-YYYY format
            Returns: dict with transfer data or None
            
        _limpiar_numero(valor)
            Purpose: Convert string number to float
            Process:
                1. Remove $ symbol
                2. Remove thousand separators (.)
                3. Convert comma to decimal point
                4. Cast to float
            Returns: float
            
        _limpiar_doc(col)
            Purpose: Normalize document numbers
            Process (on each value in series):
                1. Convert to string
                2. Remove trailing ".0"
                3. Keep only digits
                4. Remove leading zeros
                5. Strip whitespace
            Returns: pd.Series with cleaned values
            
        _reconcile_data(df_desprendibles, df_transferencia)
            Purpose: Match and compare payslip vs transfer data
            Algorithm:
                1. Clean document numbers in both DataFrames
                2. FOR EACH unique identification in payslips:
                   a. Collect all payment amounts for this ID
                   b. Find matching transfers by document number
                   c. IF no transfer found:
                      Status = "Documento no encontrado"
                   d. ELSE collect transfer amounts
                   e. IF any amount matches:
                      Status = "OK"
                   f. ELSE:
                      Status = "Valor no coincide"
                3. Build result record
            Returns: pd.DataFrame with columns:
                - Identificación: ID number
                - Estado: Status string
                - Neto_desprendibles: List of payslip amounts
                - Valores_transferencia: List of transfer amounts
        
        _display_results()
            Purpose: Render results in GUI table
            Process:
                1. Clear existing Treeview items
                2. FOR EACH result row:
                   a. Determine color tag based on status
                   b. Format data for display
                   c. Insert into Treeview
                   d. Apply color tag
            Color mapping:
                - "Documento no encontrado" → "error" (red)
                - "OK" → "ok" (green)
                - "Valor no coincide" → "warning" (yellow)
        
        _export_to_csv()
            Purpose: Save results to CSV file
            Process:
                1. Check if results exist
                2. Open save dialog
                3. Export DataFrame to CSV
                4. Show success message
        
        _clear_results()
            Purpose: Clear all displayed results
            Process:
                1. Remove all Treeview items
                2. Reset df_resultado to None
                3. Update status label


================================================================================
3. DATA PROCESSING PIPELINE
================================================================================

Stage 1: Extract Payslip Data
─────────────────────────────

Input: PDF file from desprendibles folder

Process:
    1. Open PDF with pdfplumber
    2. Iterate through pages
    3. Extract text from page
    4. Split text by "Comprobante de Nómina" marker
    5. For each block:
        a. Find identification with regex: \b\d{1,3}(?:\.\d{3}){1,3}\b
           Matches: 123.456.789, 12.345.678, 1.234.567
        b. Find "Neto a Pagar" followed by amount
           Regex: Neto a Pagar.*?\$\s*([\d\.,]+)
        c. If both found, add to records list

Output: List of dicts with:
    - Identificacion: "123456789"
    - Neto: 2500000

Transformation:
    - Remove dots from identification
    - Convert Neto to integer


Stage 2: Extract Transfer Data
──────────────────────────────

Input: PDF file from transferencia folder

Process:
    1. Open PDF with pdfplumber
    2. Iterate through pages
    3. Extract text from page
    4. For each line in text:
        a. Check if line starts with 10+ digits (account number)
        b. If yes, parse line:
            - Split by whitespace
            - Extract account, type, document, values, dates
            - Find monetary value with regex: \d{1,3}(?:,\d{3})+(?:\.\d+)?
            - Extract name between document and value
            - Find date with regex: \d{2}-\d{2}-\d{4}

Output: List of dicts with:
    - Cuenta: "1234567890"
    - Tipo: "Checking"
    - Documento: "987654321"
    - Nombre: "John Doe"
    - Valor: "2,500,000.00"
    - Fecha: "15-03-2025"

Transformation:
    - Convert Valor to numeric (remove commas, convert to int)


Stage 3: Clean Data
───────────────────

Payslips cleaning:
    1. Remove dots from identification
       "123.456.789" → "123456789"
    2. Convert Neto to integer
       "2500000.0" → 2500000

Transfers cleaning:
    1. Convert document to numeric
       "0987654321" → "987654321" (remove leading zeros)
    2. Convert Valor to integer
       "2,500,000" → 2500000
    3. Keep only digits in document numbers
       "98765-4321" → "987654321"


Stage 4: Reconciliation
───────────────────────

Matching algorithm:
    FOR EACH unique document in payslips:
        1. Get set of net amounts for this document
        2. Find all transfers with matching document number
        3. Get set of transfer amounts
        
        IF no transfers found:
            Status = "Documento no encontrado"
            Values_transfer = None
        ELSE:
            Get transfer amounts
            IF any payslip amount in transfer amounts:
                Status = "OK"
            ELSE:
                Status = "Valor no coincide"

Result: DataFrame with status for each person


Stage 5: Display & Export
─────────────────────────

Display:
    1. Iterate through results
    2. Determine color tag based on status
    3. Insert row into Treeview with tag
    4. GUI renders with color highlighting

Export:
    1. Save DataFrame to CSV
    2. Include all columns: ID, Status, Amounts, Transfer Amounts


================================================================================
4. ALGORITHM DETAILS
================================================================================

4.1 Identification Extraction (Regex Pattern Analysis)
─────────────────────────────────────────────────────

Pattern: \b\d{1,3}(?:\.\d{3}){1,3}\b

Breakdown:
    \b          - Word boundary (start of number)
    \d{1,3}     - 1-3 digits (first group)
    (?:         - Non-capturing group (can repeat)
        \.\d{3} - Literal dot followed by exactly 3 digits
    ){1,3}      - Repeat 1-3 times
    \b          - Word boundary (end of number)

Examples:
    ✓ Matches: 1.234.567, 12.345.678, 123.456.789
    ✗ Doesn't match: 1234567 (no dots), 1.23.45 (wrong format)


4.2 Amount Extraction (Regex Pattern Analysis)
──────────────────────────────────────────────

Pattern: Neto a Pagar.*?\$\s*([\d\.,]+)

Breakdown:
    Neto a Pagar    - Literal text to find
    .*?             - Any characters (non-greedy)
    \$              - Literal dollar sign
    \s*             - Zero or more whitespace
    ([\d\.,]+)      - Captured: digits, dots, commas (amount)

Examples:
    Input: "Neto a Pagar          $ 2.500.000"
    Captured: "2.500.000"
    
    Input: "Neto a Pagar: $   1,234,567.50"
    Captured: "1,234,567.50"


4.3 Transfer Line Parsing
─────────────────────────

Logic for parsing lines:
    1. Split by whitespace
    2. Index 0: Account number (10+ digits)
    3. Index 1: Account type (checking/savings, etc.)
    4. Index 2: Document (beneficiary ID)
    5. Indices 3-N: Parse carefully
        - Find monetary value: regex \d{1,3}(?:,\d{3})+(?:\.\d+)?
        - Elements before value: part of name
        - Elements after value: entity, status, date
        - Last element matching DD-MM-YYYY: date

Example line parsing:
    Input: "1234567890 Cta.Cte. 98765432 John Doe Lopez 2,500,000 ACME Corp TX 01-15-2025"
    
    Parsed:
        Cuenta: "1234567890"
        Tipo: "Cta.Cte."
        Documento: "98765432"
        Nombre: "John Doe Lopez"
        Valor: "2,500,000"
        Entidad: "ACME Corp"
        Estado: "TX"
        Fecha: "01-15-2025"


4.4 Reconciliation Matching
──────────────────────────

Two-level matching:
    
    Level 1 - Identification Matching:
        Payslip ID == Transfer Document Number
        (after cleaning both to pure numbers)
    
    Level 2 - Amount Matching (if ID found):
        IF any(payslip_amounts) IN transfer_amounts:
            Match found
        ELSE:
            Amounts differ


4.5 Data Type Conversions
──────────────────────────

Identification:
    "123.456.789" → Remove dots → "123456789" → Keep as string

Net Amount:
    "2.500.000" → Replace "." with "" → "2500000" → int(2500000)

Transfer Amount:
    "2,500,000" → Replace "," with "" → "2500000" → int(2500000)


================================================================================
5. GUI COMPONENTS
================================================================================

5.1 Main Window Structure
─────────────────────────

root: tk.Tk
    ├── header_frame (ttk.Frame)
    │   └── title_label (ttk.Label)
    │
    ├── path_frame (ttk.LabelFrame)
    │   ├── Row 0: Payslips path entry and browse button
    │   └── Row 1: Transfers path entry and browse button
    │
    ├── control_frame (ttk.Frame)
    │   ├── Process PDFs button
    │   ├── Export to CSV button
    │   ├── Clear Results button
    │   └── status_label (ttk.Label)
    │
    └── results_frame (ttk.LabelFrame)
        ├── tree (ttk.Treeview)
        ├── vsb (ttk.Scrollbar - vertical)
        └── hsb (ttk.Scrollbar - horizontal)


5.2 Treeview Configuration
──────────────────────────

Columns:
    1. "Identificación" (150px) - ID number
    2. "Estado" (200px) - Status
    3. "Neto_desprendibles" (250px) - Payslip amounts
    4. "Valores_transferencia" (250px) - Transfer amounts

Color Tags:
    "error"    → Background: #ffcccc (light red), Text: darkred
    "ok"       → Background: #ccffcc (light green), Text: darkgreen
    "warning"  → Background: #ffffcc (light yellow), Text: darkorange


5.3 Event Handlers
──────────────────

Button Clicks:
    Browse (Payslips) → _select_desprendibles_folder()
    Browse (Transfers) → _select_transferencia_folder()
    Process PDFs → _process_files() → spawn thread _process_files_thread()
    Export to CSV → _export_to_csv()
    Clear Results → _clear_results()

Window Events:
    Close (X button) → root.destroy()


5.4 Status Updates
──────────────────

Status transitions:
    "Ready" (blue) → Initial state
    "Processing..." (orange) → When processing starts
    "Processing complete. Found X records." (green) → On success
    "Processing failed" (red) → On error
    "Results cleared" (blue) → After clearing


================================================================================
6. ERROR HANDLING
================================================================================

6.1 User Input Validation
─────────────────────────

Before processing:
    ✓ Check desprendibles_path is set
    ✓ Check transferencia_path is set
    
If validation fails:
    → Show messagebox with error message
    → Return without processing

Example:
    if not self.desprendibles_path.get():
        messagebox.showerror("Error", "Please select a Payslips folder")
        return


6.2 PDF Processing Errors
─────────────────────────

Try-except blocks:
    - File not found: Caught by pdfplumber
    - PDF reading error: Caught by pdfplumber
    - Regex parsing error: Returns None, skipped
    - Data type conversion error: Caught and logged
    
Recovery:
    - Continue with next file on error
    - Report to user if entire operation fails
    - Show last successfully processed count


6.3 Graceful Degradation
────────────────────────

If a PDF file is corrupted:
    - Skip that file, continue with others
    - Display partial results

If a line can't be parsed:
    - Return None from _parsear_linea()
    - Skip line, continue with next

If regex doesn't match:
    - Variable stays None
    - Conditional check prevents processing


6.4 Threading Errors
──────────────────────

Main thread → Background thread:
    - UI remains responsive
    - Errors caught in background thread
    - Results sent back to main thread via root.after()

Example:
    try:
        # Processing in background
        self.df_resultado = self._reconcile_data(...)
    except Exception as e:
        self.root.after(0, lambda: messagebox.showerror(...))


================================================================================
7. PERFORMANCE CONSIDERATIONS
================================================================================

7.1 Time Complexity
───────────────────

PDF Extraction: O(p × l)
    p = number of pages
    l = lines per page
    Linear scan through text

Reconciliation: O(n × m)
    n = unique IDs in payslips
    m = average transfer records per ID
    For each ID, search transfers (optimizable with indexing)

Total expected time:
    100 payslips: ~5-10 seconds
    50 transfer PDFs: ~10-15 seconds
    Reconciliation: ~2-3 seconds


7.2 Memory Usage
────────────────

DataFrame storage: O(n)
    n = total records

Typical memory profile:
    1000 payslips: ~5-10 MB
    1000 transfers: ~10-15 MB
    Results: ~5-10 MB
    Total: ~30 MB

Optimization: Use dtypes efficiently
    - Identificacion: string (no numeric ID needed)
    - Neto/Valor: int64 instead of float (exact amounts)


7.3 I/O Optimization
─────────────────────

Current approach: Sequential file reading
    - Open PDF 1, read all, close
    - Open PDF 2, read all, close
    - Etc.

For large batches:
    - Consider parallel processing
    - Use multiprocessing for PDF extraction
    - Cache parsed results


7.4 Responsive UI
─────────────────

Threading strategy:
    - Main thread: UI only
    - Background thread: Heavy processing
    
No blocking:
    - User can interact while processing
    - Status updates in real-time
    - Can close app anytime


================================================================================
8. EXTENSION POINTS
================================================================================

8.1 Adding New PDF Formats
──────────────────────────

To support new payslip format:
    1. Create new method: _process_desprendibles_format_v2()
    2. Implement format-specific extraction
    3. Return same DataFrame structure
    4. Call from _process_desprendibles() based on filename pattern

Example:
    if "Hoja" in filename:
        df = self._process_desprendibles_format_v1(path)
    elif "Boleta" in filename:
        df = self._process_desprendibles_format_v2(path)


8.2 Database Integration
────────────────────────

Instead of CSV export:
    1. Add database connection method
    2. Store results in SQL database
    3. Enable historical tracking
    4. Support queries and reports

Example:
    def _export_to_database(self, db_connection):
        self.df_resultado.to_sql('reconciliation_results', db_connection)


8.3 Configuration File
──────────────────────

Add config.ini or config.json for:
    - Default folders
    - Regex patterns
    - Column mappings
    - Color schemes
    - Processing options


8.4 Advanced Filtering
──────────────────────

Add filter options:
    - Filter by status
    - Filter by ID range
    - Filter by amount range
    - Filter by date range

Example:
    def _apply_filters(self, status=None, min_amount=None, max_amount=None):
        filtered = self.df_resultado
        if status:
            filtered = filtered[filtered['Estado'] == status]
        return filtered


8.5 Report Generation
─────────────────────

Generate detailed reports:
    - Summary statistics
    - Error analysis
    - Trend analysis
    - PDF export

Example:
    def _generate_report(self):
        summary = {
            'total_records': len(self.df_resultado),
            'ok_count': len(self.df_resultado[self.df_resultado['Estado'] == 'OK']),
            'error_count': len(self.df_resultado[...]),
        }


================================================================================
END OF TECHNICAL DOCUMENTATION
================================================================================
"""

# This is a documentation file - import gui_app module to use the application
if __name__ == "__main__":
    print(__doc__)
