"""
USER GUIDE - Payroll Data Reconciliation System
Complete instructions with examples and troubleshooting
"""

# ============================================================================
# TABLE OF CONTENTS
# ============================================================================
# 1. Getting Started
# 2. Installing Required Software
# 3. Using the Application
# 4. Understanding Results
# 5. Exporting Data
# 6. Troubleshooting
# 7. Tips and Tricks
# 8. Frequently Asked Questions

# ============================================================================
# 1. GETTING STARTED
# ============================================================================

"""
WHAT DOES THIS APPLICATION DO?

The Payroll Data Reconciliation System verifies that employee payments from
payroll (desprendibles) match the corresponding bank transfer records.

WHY IS THIS IMPORTANT?

1. Verification: Ensures all payments were correctly transferred
2. Accuracy: Detects payment amount mismatches
3. Compliance: Identifies missing or unmatched records
4. Efficiency: Automates manual verification process

HOW DOES IT WORK?

1. You provide two folders:
   - Payslips folder (where your desprendibles PDFs are stored)
   - Transfers folder (where your bank transfer PDFs are stored)

2. The application reads all PDFs and extracts:
   - From payslips: Employee ID and net payment amount
   - From transfers: Document/ID number and transfer amount

3. The system matches records by ID and compares amounts

4. Results are displayed with color coding:
   - Green: Perfect match ✓
   - Red: Document not found ✗
   - Yellow: Document found but amounts differ ⚠

EXAMPLE SCENARIO:

Employee: Juan García
ID: 12.345.678

In Payslips (Desprendibles):
  - Net payment: $2,500,000

In Bank Transfers (Transferencias):
  - Transfer to ID 12345678: $2,500,000

Result: ✓ OK (Green) - Everything matches!

ANOTHER EXAMPLE:

Employee: Maria López
ID: 98.765.432

In Payslips:
  - Net payment: $3,000,000

In Bank Transfers:
  - No transfer found for ID 98765432

Result: ✗ Documento no encontrado (Red) - Missing transfer!
"""

# ============================================================================
# 2. INSTALLING REQUIRED SOFTWARE
# ============================================================================

"""
STEP 1: VERIFY PYTHON INSTALLATION

Open PowerShell and type:
  python --version

You should see something like:
  Python 3.9.0

If you get "command not found", you need to install Python first.

STEP 2: ACTIVATE VIRTUAL ENVIRONMENT

Open PowerShell in your project folder and type:
  ..\Scripts\activate

You should see (venv1.2) before your prompt.

STEP 3: INSTALL REQUIRED PACKAGES

Type the following command:
  pip install -r requirements.txt

Or manually install:
  pip install pandas pdfplumber

You should see something like:
  Successfully installed pandas-1.5.2 pdfplumber-0.7.0

STEP 4: VERIFY INSTALLATION

Type:
  python -c "import pandas; import pdfplumber; print('All packages installed!')"

If successful, you'll see:
  All packages installed!

If you see an error, repeat STEP 3.
"""

# ============================================================================
# 3. USING THE APPLICATION
# ============================================================================

"""
LAUNCHING THE APPLICATION

Method 1: From PowerShell
  1. Navigate to your project folder:
     cd c:\Users\andres.mejia\venv1.2\mapa-de-cargos
  
  2. Make sure virtual environment is activated:
     ..\Scripts\activate
  
  3. Run the application:
     python gui_app.py

Method 2: From Python directly
  python gui_app.py

Method 3: Using the quick start script
  python QUICKSTART.py

The GUI window should appear within a few seconds.


STEP-BY-STEP WALKTHROUGH

┌─────────────────────────────────────────────────────────────┐
│  Step 1: Select Payslips Folder                             │
└─────────────────────────────────────────────────────────────┘

1. Look for the button labeled "Browse" next to "Payslips Folder"
2. Click that button
3. A folder dialog will open
4. Navigate to your payslips folder
   Typical path: c:\Users\...\docs\3.Desprendibles\
5. Click "Select Folder" or double-click the folder name
6. The path should now appear in the text field

Example folder contents:
  📁 3.Desprendibles\
     📄 desprendibles_octubre_2025.pdf
     📄 desprendibles_noviembre_2025.pdf


┌─────────────────────────────────────────────────────────────┐
│  Step 2: Select Transfers Folder                            │
└─────────────────────────────────────────────────────────────┘

1. Look for the button labeled "Browse" next to "Transfers Folder"
2. Click that button
3. Navigate to your transfers folder
   Typical path: c:\Users\...\docs\1.Transferencia Bancaria\
4. Click "Select Folder"
5. The path should now appear in the text field

Example folder contents:
  📁 1.Transferencia Bancaria\
     📄 transferencias_octubre_2025.pdf
     📄 transferencias_noviembre_2025.pdf


┌─────────────────────────────────────────────────────────────┐
│  Step 3: Process PDFs                                       │
└─────────────────────────────────────────────────────────────┘

1. Click the "Process PDFs" button
2. The status bar will show "Processing... Please wait."
3. The application is now:
   - Reading all PDF files from both folders
   - Extracting employee IDs and amounts
   - Matching payslips with transfers
   - Analyzing discrepancies

4. Wait for the status to change to:
   "Processing complete. Found XXX records."

⏱️ Typical processing times:
   - Small batch (10-20 PDFs): 10-30 seconds
   - Medium batch (50-100 PDFs): 30-60 seconds
   - Large batch (100+ PDFs): 1-3 minutes

⚠️ The application may appear unresponsive - this is normal!
   The UI is processing in the background.


┌─────────────────────────────────────────────────────────────┐
│  Step 4: Review Results                                     │
└─────────────────────────────────────────────────────────────┘

Once processing is complete, you'll see a table with results.

TABLE COLUMNS:

Column 1: Identification
  - The employee's ID number
  - Example: 12345678, 98765432

Column 2: Status
  - OK: Record successfully matched
  - Documento no encontrado: Not found in transfers
  - Valor no coincide: Found but amounts don't match

Column 3: Payslip Values
  - Amount(s) from the payslip
  - Example: [2500000], [3000000]

Column 4: Transfer Values
  - Amount(s) from the transfer
  - Example: [2500000], [1500500]


COLOR MEANINGS:

🟢 GREEN (OK)
  ✓ Employee found in transfers
  ✓ Amount matches exactly
  ✓ No action needed

🔴 RED (Documento no encontrado)
  ✗ Employee NOT found in transfers
  ✗ Possible missing transfer or wrong ID format
  ⚠️ Action needed: Investigate why transfer wasn't recorded

🟡 YELLOW (Valor no coincide)
  ⚠️ Employee found BUT amounts don't match
  ✗ Payslip amount ≠ Transfer amount
  ⚠️ Action needed: Verify if amounts should match or if there's an error


INTERPRETING RESULTS - EXAMPLES

Example 1: GREEN ROW
  ID: 12345678
  Status: OK
  Payslip: [2500000]
  Transfer: [2500000]
  ✓ Perfect match - payslip amount matches transfer

Example 2: RED ROW
  ID: 87654321
  Status: Documento no encontrado
  Payslip: [3000000]
  Transfer: (empty)
  ✗ No transfer found for this employee

Example 3: YELLOW ROW
  ID: 11111111
  Status: Valor no coincide
  Payslip: [1500000]
  Transfer: [1500500]
  ⚠️ Amounts differ by 500 - investigate why

Example 4: Multiple amounts
  ID: 22222222
  Status: OK
  Payslip: [1000000, 2000000]
  Transfer: [2000000, 1000000]
  ✓ Multiple payments, all matched (order doesn't matter)


┌─────────────────────────────────────────────────────────────┐
│  Step 5: Export Results (Optional)                          │
└─────────────────────────────────────────────────────────────┘

1. Click "Export to CSV" button
2. A file save dialog will open
3. Choose a location (recommend your Documents folder)
4. Enter a filename (e.g., "reconciliation_results_2025.csv")
5. Click "Save"
6. You'll see: "Results exported to: [path]"

The CSV file will contain:
  - Identification
  - Estado (Status)
  - Neto_desprendibles (Payslip amounts)
  - Valores_transferencia (Transfer amounts)


┌─────────────────────────────────────────────────────────────┐
│  Step 6: Clear Results (Optional)                           │
└─────────────────────────────────────────────────────────────┘

If you want to process different folders:
  1. Click "Clear Results" button
  2. The table will be emptied
  3. You can select new folders and process again
"""

# ============================================================================
# 4. UNDERSTANDING RESULTS
# ============================================================================

"""
WHAT DO THE STATUS VALUES MEAN?

✓ OK
  Meaning: Everything looks good!
  - Employee ID found in transfer records
  - Amount in payslip matches transfer amount
  - No discrepancy
  Action: None needed - this is expected

✗ Documento no encontrado (Document not found)
  Meaning: Problem - employee not found in transfers!
  - Payslip shows a payment
  - No corresponding transfer record
  - Possible causes:
    * Transfer hasn't been made yet
    * Employee ID in payslip is wrong format
    * Transfer uses different ID
    * Transfer is in different file/period
  Action: Investigate - find the missing transfer

⚠️ Valor no coincide (Value doesn't match)
  Meaning: Warning - amounts don't match!
  - Employee found in transfers
  - BUT: Payslip amount ≠ Transfer amount
  - Possible causes:
    * Transfer amount was modified
    * Payslip amount was corrected
    * Partial transfer was made
    * Deductions applied differently
  Action: Review and verify which amount is correct


ANALYZING YOUR RESULTS

Healthy reconciliation looks like:
  - 85%+ records showing "OK"
  - <10% showing "Documento no encontrado"
  - <5% showing "Valor no coincide"

If you see mostly errors:
  - Check that folders are correct
  - Verify PDF files are in the right location
  - Ensure PDFs are readable (not corrupted)
  - Try processing one folder at a time


STATISTICAL SUMMARY

After processing, check:
  1. Total records: How many employees?
  2. OK count: How many matched perfectly?
  3. Missing count: How many not found?
  4. Mismatch count: How many differ?

Calculate:
  Match rate = OK count / Total × 100%

Example:
  Total: 100 employees
  OK: 95
  Missing: 3
  Mismatched: 2
  
  Match rate: 95% ✓ (Excellent)

Target match rates:
  > 95%: Excellent
  90-95%: Good
  80-90%: Needs review
  < 80%: Major issues - investigate
"""

# ============================================================================
# 5. EXPORTING DATA
# ============================================================================

"""
EXPORTING TO CSV

What is CSV?
  - CSV = Comma Separated Values
  - Plain text format that opens in Excel, Google Sheets, etc.
  - Easy to analyze and share

Steps to export:
  1. Click "Export to CSV" button
  2. Choose save location
  3. Name the file (e.g., "results_2025.csv")
  4. Click "Save"
  5. A confirmation message appears

Where to find the file:
  - It will be saved wherever you chose to save it
  - Usually in Documents or Desktop folder

Opening in Excel:
  1. Open Excel
  2. File → Open
  3. Navigate to the CSV file
  4. Click "Open"
  5. Data appears in Excel format

Excel features you can use:
  - Sort by Status column
  - Filter to show only errors
  - Create charts and pivot tables
  - Format cells and columns
  - Add calculations

Example Excel analysis:
  - Use AutoFilter to show only "Documento no encontrado" rows
  - Sort by ID
  - Export to PDF for reporting


FURTHER ANALYSIS IN EXCEL

After opening the CSV:

1. Sort by Status
   - Select all data
   - Data → Sort
   - Sort by "Estado" column
   - See all errors grouped together

2. Filter errors
   - Click on header row
   - Data → AutoFilter
   - Click dropdown on "Estado" column
   - Uncheck "OK"
   - See only problem records

3. Create pivot table
   - Data → Pivot Table
   - Summarize status counts
   - Show count of each status type

4. Add formulas
   - Count OK: =COUNTIF(B:B,"OK")
   - Count errors: =COUNTIF(B:B,"Documento no encontrado")
   - Match percentage: =COUNTIF(B:B,"OK")/COUNTA(B:B)
"""

# ============================================================================
# 6. TROUBLESHOOTING
# ============================================================================

"""
COMMON ISSUES AND SOLUTIONS

┌─────────────────────────────────────────────────────────────┐
│ ISSUE: Application won't start                              │
└─────────────────────────────────────────────────────────────┘

Error messages:
  "ModuleNotFoundError: No module named 'tkinter'"
  "ModuleNotFoundError: No module named 'pandas'"

Solution:
  1. Make sure your virtual environment is activated
     See the (venv1.2) prefix in PowerShell
  
  2. Install required packages:
     pip install pandas pdfplumber
  
  3. Try again:
     python gui_app.py


┌─────────────────────────────────────────────────────────────┐
│ ISSUE: "Please select a folder" error                       │
└─────────────────────────────────────────────────────────────┘

Cause:
  - You forgot to select one or both folders
  - Or selected folders but they show as empty

Solution:
  1. Click "Browse" button
  2. Navigate to folder
  3. Make sure folder is selected (highlighted)
  4. Click "Select Folder" button
  5. Verify the path appears in the text field
  6. Try again


┌─────────────────────────────────────────────────────────────┐
│ ISSUE: Processing starts but no results appear               │
└─────────────────────────────────────────────────────────────┘

Possible causes:
  1. PDF files are in wrong format
  2. PDF files are corrupted
  3. Folder is empty
  4. Processing is still running (status bar shows "Processing...")

Solution:
  1. Wait 1-2 minutes (it might still be processing)
  2. Check folder contents:
     - Right-click folder
     - Click "Open in Explorer" or "Show Files"
     - Verify PDF files exist
  
  3. Verify PDF format:
     - Open one PDF file manually
     - Look for expected data (IDs and amounts)
     - If unreadable, file might be corrupted
  
  4. Try with a smaller folder first:
     - Test with just 1-2 PDF files
     - If that works, issue is with file size or complexity


┌─────────────────────────────────────────────────────────────┐
│ ISSUE: All results show "Documento no encontrado"            │
└─────────────────────────────────────────────────────────────┘

Cause:
  - Transfer folder is wrong
  - Transfer PDFs are empty
  - Employee IDs don't match between payslips and transfers

Solution:
  1. Clear Results
  2. Check transfer folder:
     - Is it the correct "Transferencia Bancaria" folder?
     - Does it contain PDF files?
     - Open a PDF manually and check for data
  
  3. Check ID format:
     - Payslip ID: "12.345.678"
     - Transfer ID: "12345678"
     - They should match after removing dots
  
  4. Verify both folders are selected correctly
  5. Try processing with just one transfer PDF
     - Copy 1 transfer PDF to a test folder
     - Select test folder as transfer folder
     - Process again


┌─────────────────────────────────────────────────────────────┐
│ ISSUE: Application is slow / taking a long time              │
└─────────────────────────────────────────────────────────────┘

Cause:
  - Large number of PDF files
  - Large PDF file size
  - Slow disk drive

Expected times:
  - 10 PDFs: 5-10 seconds
  - 50 PDFs: 30-60 seconds
  - 100+ PDFs: 1-3 minutes

Solution:
  1. Be patient - let it finish
  2. For faster processing:
     - Use fewer PDFs
     - Break into smaller batches
     - Use a faster disk
  
  3. You can close and restart:
     - Close the application (×button)
     - Try with a smaller subset of files


┌─────────────────────────────────────────────────────────────┐
│ ISSUE: Export to CSV doesn't work                            │
└─────────────────────────────────────────────────────────────┘

Error:
  "No results to export. Process PDFs first."

Solution:
  1. Run "Process PDFs" first
  2. Wait for processing to complete
  3. Results should appear in table
  4. Then click "Export to CSV"

If still doesn't work:
  1. Check that results are showing in table
  2. Try a different save location
  3. Make sure you have write permissions to that location


┌─────────────────────────────────────────────────────────────┐
│ ISSUE: CSV file is empty or corrupted                        │
└─────────────────────────────────────────────────────────────┘

Solution:
  1. Delete the file
  2. Try exporting again
  3. Choose a different location
  4. Try different filename


┌─────────────────────────────────────────────────────────────┐
│ ISSUE: Can't find the application file                       │
└─────────────────────────────────────────────────────────────┘

Solution:
  1. Open File Explorer
  2. Navigate to: c:\Users\andres.mejia\venv1.2\mapa-de-cargos\
  3. You should see: gui_app.py
  4. Open PowerShell in this folder:
     - Click address bar
     - Type: powershell
     - Press Enter
  5. Run: python gui_app.py
"""

# ============================================================================
# 7. TIPS AND TRICKS
# ============================================================================

"""
PRO TIPS FOR BEST RESULTS

Tip 1: Start with a small sample
  - Test with just 1-2 payslip PDFs and 1-2 transfer PDFs
  - Verify results are correct
  - Then process all files

Tip 2: Keep PDFs organized
  - All payslip PDFs in one folder
  - All transfer PDFs in another folder
  - No other files mixed in
  - Clear folder names (easier to select)

Tip 3: Use consistent data
  - Same date range for both sets
  - Same employee ID format
  - Same currency

Tip 4: Review RED entries first
  - Sort by "Documento no encontrado"
  - These are priority issues
  - Investigate why transfers are missing

Tip 5: Export regularly
  - Keep historical records
  - Compare month-to-month trends
  - Name files with dates: "results_2025_01.csv"

Tip 6: Check the amounts carefully
  - Yellow entries might be legitimate
  - Example: Partial transfer + separate deductions
  - Document any intentional mismatches

Tip 7: Use Excel for analysis
  - Export to CSV
  - Use pivot tables
  - Create charts
  - Generate reports

Tip 8: Batch processing
  - If you have many months of data
  - Process one month at a time
  - Export each to separate CSV
  - Compare trends


KEYBOARD SHORTCUTS

Once the GUI is open:
  Tab             - Move between fields
  Enter/Return    - Activate button (if focused)
  Ctrl+A (in fields) - Select all text
  Ctrl+C (selected text) - Copy


KEYBOARD NAVIGATION

Tab through interface:
  1. Payslips path field
  2. Payslips browse button
  3. Transfers path field
  4. Transfers browse button
  5. Process PDFs button
  6. Export to CSV button
  7. Clear Results button
  8. Results table
"""

# ============================================================================
# 8. FREQUENTLY ASKED QUESTIONS
# ============================================================================

"""
FAQ

Q: What if I need to process more folders?
A: Click "Clear Results" and select new folders. You can process as many
   times as needed.

Q: Can I keep the same folders and process again?
A: Yes. Just click "Process PDFs" button again. It will overwrite previous
   results.

Q: What if PDFs have different formats?
A: The application is optimized for standard Colombian payroll formats. If
   PDFs have very different layouts, they might not extract correctly.

Q: Can I edit the results in the table?
A: No, results are view-only. Export to CSV if you need to edit.

Q: Does this work with Excel files too?
A: Not yet. Only PDF files are supported currently.

Q: What if I have very large PDF files?
A: Processing might be slow. Try breaking into smaller batches.

Q: Is there a limit on number of records?
A: No hard limit, but processing slows with very large files (1000+ records).

Q: Can I process PDFs from different dates?
A: Yes, but results will be mixed. Best to keep same date range.

Q: What does "Neto a Pagar" mean?
A: "Net to Pay" - the final payment amount after deductions.

Q: What if amounts are in different currencies?
A: Results might not match. Ensure both PDFs use same currency.

Q: Can I undo processing?
A: Click "Clear Results" to clear current results.
   Original files are never modified.

Q: Is my data secure?
A: Yes. Application runs locally on your computer.
   No data is sent to internet or cloud.

Q: Can I run this on Mac?
A: Yes. Requires Python and same packages.
   File paths will be different (/Users/... instead of C:\Users\...)

Q: Can I schedule this to run automatically?
A: Yes. You can create a scheduled task in Windows.
   Contact IT if you need help.

Q: What if I want to modify the code?
A: Code is well-documented. See TECHNICAL_DOCS.py for details.
   Python knowledge required.

Q: Who do I contact for support?
A: Check TECHNICAL_DOCS.py or README_GUI.md for detailed documentation.
   For bugs, review error messages and troubleshooting section.
"""

# ============================================================================
# END OF USER GUIDE
# ============================================================================

if __name__ == "__main__":
    print(__doc__)
    print("\n" + "="*70)
    print("For more information, see:")
    print("  - README_GUI.md (User-friendly overview)")
    print("  - TECHNICAL_DOCS.py (Developer documentation)")
    print("  - gui_app.py (Source code with full documentation)")
    print("="*70)
