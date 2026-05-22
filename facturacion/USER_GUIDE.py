"""
COMPREHENSIVE USER GUIDE
Services & Equipment Validation System GUI Application

================================================================================
TABLE OF CONTENTS
================================================================================
1. Getting Started
2. Installation
3. Step-by-Step Usage
4. Understanding Results
5. Exporting Data
6. Troubleshooting
7. Tips and Tricks
8. FAQ

================================================================================
1. GETTING STARTED
================================================================================

WHAT DOES THIS APPLICATION DO?

The Services & Equipment Validation System verifies that service/equipment
quantities reported in PDF documents match the historical data stored in Excel.

WHY IS THIS IMPORTANT?

1. Data Accuracy: Ensures reporting consistency
2. Error Detection: Identifies discrepancies quickly
3. Reconciliation: Automates manual verification
4. Audit Trail: Documents validation results
5. Efficiency: Saves time vs manual checking

HOW IT WORKS:

1. You provide two files:
   - PDF report file (new data to validate)
   - Excel historical data file (reference data)

2. The application:
   - Extracts service quantities from PDF
   - Retrieves historical quantities from Excel
   - Matches records by service and date
   - Compares quantities

3. Results display:
   - "All OK" if everything matches
   - Error table if discrepancies found
   - Can export results to CSV

EXAMPLE SCENARIO 1 - SUCCESS:

PDF Report (April 2026):
  Service A: 100 units
  Service B: 50 units

Excel Historical (April 2026):
  Service A: 100 units
  Service B: 50 units

Result: ✓ All OK - Perfect match!


EXAMPLE SCENARIO 2 - DISCREPANCY:

PDF Report (April 2026):
  Service A: 105 units
  Service B: 50 units

Excel Historical (April 2026):
  Service A: 100 units
  Service B: 50 units

Result: ✗ Discrepancy found!
  Service A differs by 5 units (PDF: 105, Excel: 100)

================================================================================
2. INSTALLATION
================================================================================

STEP 1: VERIFY PYTHON INSTALLATION

Open PowerShell and type:
  python --version

You should see: Python 3.9.0 (or similar)

If error, install Python first.

STEP 2: ACTIVATE VIRTUAL ENVIRONMENT

Open PowerShell in project folder:
  ..\Scripts\activate

You should see (venv1.2) before your prompt.

STEP 3: INSTALL REQUIRED PACKAGES

Type:
  pip install -r requirements.txt

Or manually:
  pip install pandas pdfplumber openpyxl

You should see: Successfully installed...

STEP 4: VERIFY INSTALLATION

Type:
  python -c "import pandas; import pdfplumber; print('OK')"

Should show: OK

================================================================================
3. STEP-BY-STEP USAGE
================================================================================

LAUNCHING THE APPLICATION

Method 1: From PowerShell
  1. Navigate to folder:
     cd c:\Users\andres.mejia\venv1.2\facturacion
  
  2. Activate environment:
     ..\Scripts\activate
  
  3. Run app:
     python gui_validation_app.py

Method 2: Direct execution
  python gui_validation_app.py

Method 3: Using quick start
  python QUICKSTART.py

The GUI window will appear.


DETAILED WALKTHROUGH

┌─────────────────────────────────────────────────────────────┐
│  Step 1: Select PDF Report File                             │
└─────────────────────────────────────────────────────────────┘

1. Look for button labeled "Browse" next to "PDF Report:"
2. Click that button
3. File dialog opens
4. Navigate to your PDF report file
   Example location: C:\Reports\services_april_2026.pdf
5. Click "Open" or double-click filename
6. Path appears in text field

PDF File Requirements:
  - Must contain tables
  - Tables need columns for:
    * Equipment/service type
    * Quantity
    * Date
  - Supported formats: .pdf


┌─────────────────────────────────────────────────────────────┐
│  Step 2: Select Excel Historical Data File                  │
└─────────────────────────────────────────────────────────────┘

1. Look for button labeled "Browse" next to "Excel Historical Data:"
2. Click that button
3. File dialog opens
4. Navigate to Excel file
   Example: C:\Data\historical_services.xlsx
5. Click "Open"
6. Path appears in text field

Excel File Requirements:
  - Column named "DESCRIPCION TARIFA"
  - Date columns (Timestamp type, not text)
  - Service descriptions
  - Quantity values
  - Supported formats: .xlsx, .xls


┌─────────────────────────────────────────────────────────────┐
│  Step 3: Validate Files                                      │
└─────────────────────────────────────────────────────────────┘

1. Click "Validate Files" button
2. Status bar shows: "Processing... Please wait."
3. Application is now:
   - Reading PDF file
   - Scanning for tables
   - Extracting service data
   - Reading Excel file
   - Comparing quantities
   - Analyzing discrepancies

4. Wait for completion - status bar shows:
   "All OK!" or "Found X discrepancies"

⏱️ Expected Times:
   - Simple files: 5-10 seconds
   - Medium files: 15-30 seconds
   - Large files: 1-3 minutes

⚠️ UI may appear unresponsive - this is normal!
   Processing happens in background.


┌─────────────────────────────────────────────────────────────┐
│  Step 4: Review Results                                      │
└─────────────────────────────────────────────────────────────┘

RESULT 1 - All OK:

A dialog box appears:
  ✓ Validation Successful
  ✓ No discrepancies found!
  ✓ All services match between PDF and Excel

Action: Click OK to close dialog
Result: Table is empty (no errors to show)
Next: Optionally export empty results for documentation


RESULT 2 - Discrepancies Found:

Table shows error rows (red background):

┌────────────┬──────────────────┬──────┬───────┬──────────┐
│ Date       │ Service          │ PDF  │ Excel │ Diff     │
├────────────┼──────────────────┼──────┼───────┼──────────┤
│ 2026-04-15 │ Service A        │ 105  │ 100   │ 5        │
│ 2026-04-15 │ Equipment B      │ 50   │ 55    │ 5        │
│ 2026-05-20 │ Special Service  │ 200  │ 195   │ 5        │
└────────────┴──────────────────┴──────┴───────┴──────────┘

Understanding Each Column:
  Date: When service was recorded
  Service: Type of service/equipment
  PDF: Quantity reported in PDF
  Excel: Quantity in historical data
  Diff: Absolute difference between the two

What to do:
  1. Check if discrepancy is expected
  2. Verify service names are identical
  3. Look for data entry errors
  4. Compare dates carefully
  5. Investigate source of difference
  6. Update file if correction needed


┌─────────────────────────────────────────────────────────────┐
│  Step 5: Export Results (Optional)                          │
└─────────────────────────────────────────────────────────────┘

1. Click "Export to CSV" button
2. File save dialog appears
3. Choose location:
   Example: C:\Results\validation_2026_04.csv
4. Enter filename
5. Click "Save"
6. Success message appears with file location

CSV File Contents:
  - Column headers: Date, Service, PDF, Excel, Difference
  - One row per discrepancy
  - Can be opened in Excel or text editor
  - Useful for archiving and reporting

How to use CSV:
  - Import to Excel for analysis
  - Create charts from data
  - Email to stakeholders
  - Archive for audit trail
  - Process for further analysis

================================================================================
4. UNDERSTANDING RESULTS
================================================================================

READING THE TABLE

When discrepancies are shown:

🔴 RED ROW COLOR MEANS:
   - This service has a quantity mismatch
   - PDF quantity ≠ Excel quantity
   - Action may be needed

INTERPRETING EACH COLUMN:

Date Column:
  - Format: YYYY-MM-DD (2026-04-15)
  - Matches the date from PDF report
  - Used to link with Excel historical data
  - Helps identify which period the error occurred

Service Column:
  - Equipment or service type name
  - From PDF (may be truncated if very long)
  - Should match between PDF and Excel
  - If names don't match, no comparison possible

PDF Column:
  - Quantity reported in the PDF document
  - Extracted from table in PDF
  - This is the "new" or "reported" data
  - Starting point for comparison

Excel Column:
  - Historical quantity in Excel file
  - Reference/expected quantity
  - This is the "known" or "baseline" data
  - What the PDF should match

Difference Column:
  - Absolute difference: |PDF - Excel|
  - Always positive number
  - Magnitude of discrepancy
  - Helps prioritize by severity

Example Row Analysis:

Row: 2026-04-15 | Service A | 105 | 100 | 5

Interpretation:
  "For Service A on April 15, 2026:
   - PDF shows 105 units
   - Excel shows 100 units
   - Difference of 5 units
   - PDF is 5% higher than expected"

Possible causes:
  - Data entry error in PDF
  - Excel data not yet updated
  - Service delivered but not recorded
  - Counting error
  - System sync issue


SEVERITY LEVELS

Small Differences (1-5):
  Likely: Rounding errors, data entry typo
  Action: Double-check entries

Medium Differences (6-50):
  Likely: Missing update or error
  Action: Investigate source

Large Differences (50+):
  Likely: Significant discrepancy
  Action: Immediate investigation required

================================================================================
5. EXPORTING DATA
================================================================================

CSV EXPORT PROCESS

Steps:
  1. Validate files (get results)
  2. Click "Export to CSV"
  3. Choose save location
  4. Name the file
  5. Click "Save"

Save Location Tips:
  - Desktop: Easy access
  - Documents: Better organization
  - Project folder: Keep with source files
  - Date-stamped name: Better tracking

Example Filenames:
  validation_2026_04.csv
  services_discrepancies_april.csv
  reconciliation_report_04_15.csv


USING THE CSV FILE

In Excel:
  1. Open Excel
  2. File → Open
  3. Select CSV file
  4. Data appears in columns

Advanced Excel Actions:
  - Sort by Service
  - Filter by Difference > 10
  - Create pivot table
  - Add formulas for analysis
  - Create charts
  - Generate reports

In Text Editor (Notepad, VSCode, etc):
  - View raw data
  - Edit manually if needed
  - Email as attachment
  - Version control

In Python:
  ```
  import pandas as pd
  df = pd.read_csv('validation_2026_04.csv')
  # Further analysis
  ```

Sharing Results:
  - Email CSV to stakeholders
  - Include in reports
  - Archive for audit trail
  - Use for compliance documentation

================================================================================
6. TROUBLESHOOTING
================================================================================

ISSUE: Application won't start

Error: "ModuleNotFoundError: No module named 'pdfplumber'"

Solution:
  1. Activate virtual environment
  2. Install packages:
     pip install pandas pdfplumber openpyxl
  3. Try again:
     python gui_validation_app.py


ISSUE: File selection error - "Please select a file"

Cause: One or both files not selected

Solution:
  1. Click Browse button for PDF
  2. Select file and click Open
  3. Verify path appears in field
  4. Repeat for Excel file
  5. Try Validate again


ISSUE: "No valid services found in PDF"

Cause: PDF format not recognized

Solution:
  1. Verify PDF contains tables
  2. Check for columns:
     - "Tipo Equipo" (equipment type)
     - "Cantidad" (quantity)
  3. Ensure date is present
  4. Try different PDF
  5. Check file isn't corrupted


ISSUE: "No date columns detected in Excel"

Cause: Excel structure doesn't match expected format

Solution:
  1. Verify Excel has date columns
  2. Dates must be Timestamp type (not text)
  3. Look for "DESCRIPCION TARIFA" column
  4. Check Excel structure:
     - Row 1: Headers
     - Column 1: Service descriptions
     - Other columns: Dates
  5. Recreate Excel if needed


ISSUE: Processing takes very long

Cause: Large file size or many pages

Solution:
  1. Be patient - let it complete
  2. For faster processing:
     - Use smaller file subsets
     - Close other applications
     - Use faster storage (SSD)
  3. For recurring processing:
     - Split large files
     - Create smaller batches


ISSUE: Results show all services as errors

Cause: Service names don't match

Solution:
  1. Check capitalization
  2. Look for extra spaces
  3. Compare spelling carefully
  4. Check for special characters
  5. Normalize service names in source files


ISSUE: CSV export fails

Cause: File permission or path issue

Solution:
  1. Choose different location
  2. Ensure write permissions
  3. Close any open Excel files
  4. Try different filename
  5. Check disk space


ISSUE: Table shows but values look wrong

Cause: Number formatting issue

Solution:
  1. Check decimal places
  2. Verify number format
  3. Review original PDF/Excel
  4. Export to CSV and check
  5. Contact support if persists

================================================================================
7. TIPS AND TRICKS
================================================================================

PRO TIPS FOR BEST RESULTS

Tip 1: Use consistent file names
  Good: services_april_2026.pdf, services_april_2026.xlsx
  Bad: report.pdf, data.xlsx
  Why: Easier to track and organize

Tip 2: Keep files organized
  - All PDFs in one folder
  - All Excel files in another
  - Makes selection faster

Tip 3: Validate one date at a time initially
  - Start with single month PDF
  - Verify Excel file works
  - Then process full periods

Tip 4: Document discrepancies
  - Note expected differences
  - Comment on known issues
  - Track with version numbers

Tip 5: Use CSV for archiving
  - Save validation results
  - Keep monthly records
  - Create audit trail
  - Reference for future issues

Tip 6: Regular validation
  - Run weekly or monthly
  - Catch errors early
  - Maintain data quality

Tip 7: Cross-check manually
  - For large discrepancies
  - Before making corrections
  - Verify data source


KEYBOARD SHORTCUTS

Tab             - Move between fields
Shift+Tab       - Move backwards
Enter           - Activate button
Escape          - Close dialogs
Ctrl+C (table)  - Copy selected rows


WORKFLOW FOR MONTHLY VALIDATION

1. Get new PDF report
2. Launch application
3. Select PDF file
4. Select Excel file
5. Validate
6. Review results
7. Export to CSV (with date stamp)
8. Archive CSV
9. If errors: Investigate and correct
10. Document findings

================================================================================
8. FREQUENTLY ASKED QUESTIONS
================================================================================

Q: Can I validate multiple files at once?
A: Currently validates one PDF + one Excel pair. For multiple sets,
   repeat the process for each pair.

Q: What if PDF has multiple dates?
A: Application handles this! It compares data for each date separately.

Q: Can I edit results in the table?
A: No, results are view-only. Export to CSV to edit.

Q: Is there a limit on file size?
A: No hard limit, but very large files (1000+ pages) may be slow.

Q: What date formats are supported?
A: Spanish ("6 de abril de 2026"), ISO (2026-04-06), and others.

Q: Can I run this on Mac or Linux?
A: Yes! Python works on all platforms. File paths will be different.

Q: Does this work with Google Sheets?
A: Currently only standard Excel files. Export Google Sheet to .xlsx first.

Q: Can I schedule validation automatically?
A: Not built-in, but you can create a batch script or scheduled task.

Q: Is my data secure?
A: Yes! Application runs locally on your computer. No cloud upload.

Q: What if service names differ between files?
A: They must match to compare. Normalize names in source files first.

Q: Can I change the comparison tolerance?
A: Current version looks for exact matches. Edit code to add tolerance.

Q: How do I report bugs?
A: Check TECHNICAL_DOCS.py for error details or review error messages.

Q: Can I use this for other validations?
A: Yes! Code is modular and can be adapted for other PDF/Excel comparisons.

Q: What's the maximum number of services?
A: No limit - handles hundreds or thousands of services.

Q: Can I get help with Excel formatting?
A: See "EXCEL REQUIREMENTS" section of README_GUI.md

Q: Does it support PDF protection?
A: Not currently. Unprotect PDFs first if needed.

Q: How do I update the application?
A: Redownload latest version or update source files manually.

================================================================================
END OF USER GUIDE
================================================================================

For more information, see:
- README_GUI.md (Overview and features)
- TECHNICAL_DOCS.py (Developer documentation)
- QUICKSTART.py (Quick reference)
- gui_validation_app.py (Source code)
"""

if __name__ == "__main__":
    print(__doc__)
