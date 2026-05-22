"""
TECHNICAL DOCUMENTATION
Services & Equipment Validation System - GUI Application

================================================================================
ARCHITECTURE OVERVIEW
================================================================================

The application follows Model-View-Controller (MVC) pattern:

┌─────────────────────────────────────┐
│    USER INTERFACE (View)            │
│    TkInter GUI with file selection  │
└────────────────┬────────────────────┘
                 │ User interactions
                 ▼
┌─────────────────────────────────────┐
│  APPLICATION LOGIC (Controller)     │
│  ServicesValidationApp class        │
│  - File selection                   │
│  - Data processing coordination     │
│  - Results display                  │
└────────────────┬────────────────────┘
                 │ Process & analyze
                 ▼
┌─────────────────────────────────────┐
│    DATA (Model)                     │
│    PDF files, Excel files           │
│    DataFrames, Results              │
└─────────────────────────────────────┘


Data Flow:

PDF File ──┐
           ├──► Extract Tables ──┐
           │                     ├──► Parse Data ──┐
           │                     │                 │
           └─────────────────────┘                 ├──► Normalize ──┐
                                                   │                │
Excel File ────────────────────────────────────┐   │                ├──► Compare ──► Results
                                              │   │                │
                                              └──► Extract Data ──┘

================================================================================
CLASS STRUCTURE
================================================================================

ServicesValidationApp
│
├── Attributes:
│   ├── root: tk.Tk (Main window)
│   ├── excel_path: tk.StringVar (Excel file path)
│   ├── pdf_path: tk.StringVar (PDF file path)
│   ├── df_resultado: pd.DataFrame (Results)
│   ├── tree: ttk.Treeview (Results table)
│   └── status_label: ttk.Label (Status display)
│
├── UI Methods:
│   ├── _create_ui() - Build GUI
│   ├── _select_pdf_file() - Browse PDF
│   ├── _select_excel_file() - Browse Excel
│   ├── _display_results() - Show results table
│   ├── _show_all_ok_message() - Success dialog
│   └── _clear_results() - Clear display
│
├── Processing Methods:
│   ├── _process_files() - Main processing
│   ├── _process_files_thread() - Background thread
│   ├── _extraer_conteo_pdf() - Extract PDF data
│   ├── _extraer_conteo_excel() - Extract Excel data
│   └── _comparar_conteos() - Compare data
│
├── Utility Methods:
│   ├── _normalizar_texto_equipo() - Clean text
│   ├── _normalizar_fecha() - Parse date
│   ├── _format_number() - Format number
│   └── main() - Entry point


================================================================================
DETAILED METHOD DOCUMENTATION
================================================================================

__init__(self, root)
Purpose: Initialize the application
Parameters:
    root (tk.Tk): Main window instance
Process:
    1. Set window properties
    2. Initialize path variables
    3. Initialize results dataframe
    4. Create UI components
Returns: None


_create_ui(self)
Purpose: Build all GUI components
Components created:
    - Header with title
    - Path selection frame
    - Control buttons
    - Results Treeview with scrollbars
    - Color tags for highlighting
Returns: None


_select_pdf_file(self)
Purpose: Open file browser for PDF selection
Returns: None (updates self.pdf_path)


_select_excel_file(self)
Purpose: Open file browser for Excel selection
Returns: None (updates self.excel_path)


_process_files(self)
Purpose: Validate file selections and start processing
Validation:
    - Checks both paths are set
    - Updates status bar
    - Spawns background thread
Returns: None


_process_files_thread(self)
Purpose: Background thread for file processing
Steps:
    1. Extract PDF data
    2. Check for valid PDF data
    3. Compare with Excel
    4. Update UI with results
Error handling:
    - Try-except with error messagebox
    - Reports status to UI
Returns: None


_normalizar_texto_equipo(self, texto)
Purpose: Clean equipment/service text
Process:
    1. Check if string type
    2. Replace newlines with spaces
    3. Collapse multiple spaces
    4. Strip whitespace
Args:
    texto (str): Raw text
Returns:
    str: Normalized text


_normalizar_fecha(self, valor)
Purpose: Parse various date formats
Supports:
    - Spanish: "6 de abril de 2026"
    - ISO: "2026-04-06"
    - EU: "06/04/2026"
    - Timestamp objects
Algorithm:
    1. Convert to string and lowercase
    2. Try pandas default parsing
    3. Try Spanish regex pattern
    4. Extract day, month (from dict), year
    5. Create Timestamp
Args:
    valor (str/Timestamp): Date value
Returns:
    pd.Timestamp: Normalized date or None


_extraer_conteo_pdf(self, path_planilla)
Purpose: Extract service counts from PDF
Algorithm:
    1. Open PDF with pdfplumber
    2. For each page:
        a. Extract all tables
        b. For each table:
           - Extract date from header
           - Find header row (contains "Tipo Equipo", "Cantidad")
           - Extract service and quantity for each row
           - Add to records list
    3. Group by date and service
    4. Sum quantities
Args:
    path_planilla (str): PDF file path
Returns:
    pd.DataFrame: [FECHA, TIPO DE EQUIPO, CANTIDAD]


_extraer_conteo_excel(self, path_hist, fecha_reporte)
Purpose: Extract equipment counts from Excel for specific date
Algorithm:
    1. Read Excel file
    2. Identify date columns (Timestamp type)
    3. Reshape from wide to long format
    4. Filter by report date
    5. Group by service description
    6. Remove zero values
    7. Normalize service names
    8. Return as dictionary
Args:
    path_hist (str): Excel file path
    fecha_reporte (pd.Timestamp): Filter date
Returns:
    dict: {service_name: quantity}


_comparar_conteos(self, df_pdf, path_excel)
Purpose: Compare PDF vs Excel data
Algorithm:
    1. For each unique date in PDF:
        a. Extract Excel data for that date
        b. For each service in PDF:
           - Get PDF quantity
           - Get Excel quantity
           - Calculate absolute difference
           - If difference > 0: add to results
    2. Filter out excluded services
    3. Return DataFrame with discrepancies
Args:
    df_pdf (pd.DataFrame): PDF extracted data
    path_excel (str): Excel file path
Returns:
    pd.DataFrame: [Fecha, Servicio, PDF, Excel, Diferencia]


_format_number(self, valor)
Purpose: Format numbers for display
Logic:
    1. Check if NaN
    2. If whole number: return as int
    3. Else: return with 2 decimals
Args:
    valor (float): Number to format
Returns:
    str: Formatted string


_display_results(self)
Purpose: Render results in Treeview table
Process:
    1. Clear existing rows
    2. For each result row:
        a. Format date as YYYY-MM-DD
        b. Truncate long service names
        c. Format all numbers
        d. Insert into tree with "error" tag
Tags applied:
    "error": Red background, darkred text
Returns: None


_show_all_ok_message(self)
Purpose: Display success message when validation passes
Shows:
    - "Validation Successful ✅" dialog
    - Message about no discrepancies
    - Clears result table
Returns: None


_export_to_csv(self)
Purpose: Save validation results to CSV file
Process:
    1. Check results exist
    2. Open save dialog
    3. Export DataFrame to CSV
    4. Show success message
Returns: None


_clear_results(self)
Purpose: Clear displayed results and reset state
Process:
    1. Delete all table rows
    2. Reset df_resultado
    3. Update status label
Returns: None


================================================================================
DATA STRUCTURES
================================================================================

PDF Extracted Data (df_pdf):
┌─────────────┬─────────────────┬──────────┐
│ FECHA       │ TIPO DE EQUIPO  │ CANTIDAD │
├─────────────┼─────────────────┼──────────┤
│ 2026-04-15  │ Service A       │ 100      │
│ 2026-04-15  │ Service B       │ 50       │
│ 2026-05-20  │ Service A       │ 105      │
└─────────────┴─────────────────┴──────────┘


Excel Extracted Data (for specific date):
{
    "Service A": 100,
    "Service B": 55,
    "Service C": 200
}


Comparison Results (df_resultado):
┌─────────────┬──────────┬─────┬───────┬─────────────┐
│ Fecha       │ Servicio │ PDF │ Excel │ Diferencia  │
├─────────────┼──────────┼─────┼───────┼─────────────┤
│ 2026-04-15  │ Service B│ 50  │ 55    │ 5           │
│ 2026-05-20  │ Service A│ 105 │ 100   │ 5           │
└─────────────┴──────────┴─────┴───────┴─────────────┘


================================================================================
ERROR HANDLING
================================================================================

Input Validation:
    - Check file paths selected
    - Verify files exist and readable
    - Validate file formats
    - Check for required columns

PDF Processing Errors:
    - Handle missing tables
    - Skip malformed rows
    - Handle date parsing failures
    - Handle quantity parsing failures

Excel Processing Errors:
    - Check for required columns
    - Handle missing date columns
    - Handle missing data
    - Handle format mismatches

Comparison Errors:
    - Handle empty DataFrames
    - Handle None values
    - Handle type mismatches

Recovery Strategy:
    - Return empty DataFrame
    - Show informative error message
    - Allow user to try different files
    - Log errors for debugging

================================================================================
REGEX PATTERNS
================================================================================

Equipment Header Detection:
    Pattern: "tipo" in text AND "equipo" in text
    Detects: "Tipo Equipo", "TIPO EQUIPO", "Tipo de Equipo"
    Purpose: Find equipment column

Quantity Header Detection:
    Pattern: "cant" in text
    Detects: "Cantidad", "CANTIDAD", "Cant", "cant"
    Purpose: Find quantity column

Spanish Date Pattern:
    Pattern: (\d{1,2})\s+de\s+([a-záéíóúñ]+)\s+de\s+(\d{4})
    Example: "6 de abril de 2026"
    Captures: day, month (text), year

Colombian Number Format:
    Pattern: \d{1,3}(?:\.\d{3})*(?:,\d+)?
    Examples: "1.000,00", "100", "1.234.567,89"
    Process: Replace "." with "", replace "," with "."

================================================================================
PERFORMANCE CONSIDERATIONS
================================================================================

Time Complexity:
    - PDF extraction: O(p × t × r) where:
        p = number of pages
        t = tables per page
        r = rows per table
    - Excel extraction: O(r × c) where:
        r = rows
        c = columns
    - Comparison: O(n × m) where:
        n = unique dates
        m = services per date

Space Complexity:
    - Temporary storage: O(n) for DataFrames
    - Final results: O(d) where d = discrepancies

Optimization Strategies:
    - Use pandas groupby for aggregation
    - Filter early to reduce data size
    - Use sets for unique value tracking
    - Threading for responsive UI

Memory Optimization:
    - Delete temporary DataFrames
    - Use appropriate dtypes
    - Avoid creating copies unnecessarily

================================================================================
THREADING IMPLEMENTATION
================================================================================

Main Thread (UI Thread):
    - Runs Tkinter event loop
    - Handles user interactions
    - Stays responsive

Background Thread:
    - Runs _process_files_thread()
    - Heavy processing (PDF/Excel reading)
    - Doesn't block UI

Thread Communication:
    - Uses root.after() for safe UI updates
    - Passes results via instance variables
    - Messagebox for error reporting

Daemon Thread:
    - Set as daemon (thread.daemon = True)
    - Dies when main thread exits
    - No cleanup needed

================================================================================
EXTENSION POINTS
================================================================================

Adding New File Format Support:

1. PDF variants (PowerPoint, Word):
    - Create new extraction method
    - Follow _extraer_conteo_pdf() pattern
    - Return same DataFrame structure

2. Database source:
    - Create _extraer_conteo_database()
    - Query database for records
    - Convert to DataFrame
    - Return same format

3. API source:
    - Create _extraer_conteo_api()
    - Call external API
    - Parse JSON response
    - Convert to DataFrame

Adding New Validation Rules:

1. Tolerance-based matching:
    - Add tolerance parameter
    - Calculate percentage difference
    - Compare against tolerance

2. Date fuzzy matching:
    - Allow date range instead of exact
    - Match closest date if exact not found

3. Partial name matching:
    - Use fuzzy string matching
    - Handle typos and abbreviations

Adding New Output Formats:

1. Excel export:
    - Use openpyxl
    - Add formatting/colors
    - Include charts

2. PDF report:
    - Use reportlab
    - Generate formatted report
    - Add charts/graphs

3. Email notification:
    - Use smtplib
    - Send results via email
    - Include summary statistics

================================================================================
TESTING STRATEGY
================================================================================

Unit Tests:
    - Test normalization functions
    - Test date parsing
    - Test number formatting
    - Test DataFrame operations

Integration Tests:
    - Test full workflow
    - Test with sample files
    - Test error conditions
    - Test edge cases

Test Files Needed:
    - Small valid PDF
    - Small valid Excel
    - Invalid PDF (no tables)
    - Invalid Excel (missing column)
    - Edge case: Unicode characters
    - Edge case: Very large numbers

================================================================================
END OF TECHNICAL DOCUMENTATION
================================================================================
"""

if __name__ == "__main__":
    print(__doc__)
