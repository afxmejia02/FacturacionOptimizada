"""
Quick Start Guide for Payroll Reconciliation GUI Application

This script demonstrates how to run the GUI application.
"""

# ============================================================================
# INSTALLATION
# ============================================================================
# 1. Make sure you're in your virtual environment
# 2. Install required packages (if not already installed):
#    pip install pandas pdfplumber
#
# 3. Run this script or run directly:
#    python gui_app.py

# ============================================================================
# USAGE
# ============================================================================
# 1. Open your terminal/PowerShell
# 2. Navigate to the mapa-de-cargos directory:
#    cd c:\Users\andres.mejia\venv1.2\mapa-de-cargos
#
# 3. Activate your virtual environment:
#    On Windows: ..\Scripts\activate
#    On macOS/Linux: source ../bin/activate
#
# 4. Run the application:
#    python gui_app.py
#
# 5. The GUI window will open
# 6. Follow the on-screen instructions

# ============================================================================
# EXPECTED FOLDER STRUCTURE
# ============================================================================
# Your folder structure should be something like:
#
# docs/
#   ├── 3.Desprendibles/
#   │   └── desprendibles octubre 2025.pdf
#   │   └── desprendibles noviembre 2025.pdf
#   │   └── (other payslip PDFs)
#   │
#   └── 1.Transferencia Bancaria/
#       └── transferencias octubre 2025.pdf
#       └── transferencias noviembre 2025.pdf
#       └── (other transfer PDFs)

# ============================================================================
# STEP-BY-STEP GUIDE
# ============================================================================
#
# STEP 1: Launch Application
#   $ python gui_app.py
#
# STEP 2: Select Payslips Folder
#   - Click "Browse" button next to "Payslips Folder (Desprendibles)"
#   - Navigate to: docs/3.Desprendibles/
#   - Click "Select Folder"
#
# STEP 3: Select Transfers Folder
#   - Click "Browse" button next to "Transfers Folder (Transferencias)"
#   - Navigate to: docs/1.Transferencia Bancaria/
#   - Click "Select Folder"
#
# STEP 4: Process PDFs
#   - Click "Process PDFs" button
#   - Wait for processing to complete (check status bar)
#   - Results will appear in the table
#
# STEP 5: Review Results
#   - Green rows: ✓ OK - document matched and values correct
#   - Red rows: ✗ Documento no encontrado - document missing in transfers
#   - Yellow rows: ⚠ Valor no coincide - document found but amounts differ
#
# STEP 6: Export Results (Optional)
#   - Click "Export to CSV"
#   - Choose a location and filename
#   - File will be saved with all reconciliation details

# ============================================================================
# UNDERSTANDING THE OUTPUT
# ============================================================================
#
# Identification: The document/ID number of the person
#
# Status:
#   - "OK": Payment successfully transferred
#   - "Documento no encontrado": No transfer found for this person
#   - "Valor no coincide": Transfer found but amount doesn't match
#
# Payslip Values: Amounts shown in the payslip(s)
# Transfer Values: Amounts shown in the transfer record(s)

# ============================================================================
# TROUBLESHOOTING
# ============================================================================
#
# Q: Application doesn't start
# A: Make sure you have Python installed and all packages:
#    pip install tkinter pandas pdfplumber
#
# Q: "Folder not found" error
# A: Make sure you've selected valid folders with PDF files
#
# Q: No results after processing
# A: Verify PDFs are in the correct format and location
#
# Q: Application is slow
# A: Normal for large PDF files. Processing happens in background.
#    Wait for status to show "Processing complete"
#
# Q: All results show "Documento no encontrado"
# A: Check that:
#    - Transfer folder is correct
#    - Transfer PDFs are in the folder
#    - Document numbers in PDFs match format

# ============================================================================
# APPLICATION FEATURES
# ============================================================================
#
# ✓ Automatic PDF extraction
# ✓ Intelligent data matching
# ✓ Color-coded results
# ✓ Real-time processing status
# ✓ Export to CSV
# ✓ User-friendly interface
# ✓ Comprehensive documentation
# ✓ Error handling and validation

if __name__ == "__main__":
    # Import and run the GUI application
    from gui_app import main
    main()
