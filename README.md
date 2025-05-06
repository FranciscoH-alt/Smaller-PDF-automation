# Smaller-PDF-automation
This is the smaller version of the automation of the pdf to excel

 # Monthly Surcharge PDF Matching

This notebook extracts and matches PDF content against descriptions from a monthly surcharge Excel sheet.

  ## How to Use

1. Upload the monthly Excel file (e.g., `June_2025_Surcharges.xlsx`).
2. Upload PDFs into a folder (e.g., `June_PDFs`).
3. Update variables in the notebook:
   - `pdf_folder`
   - `excel_path`
   - `output_excel_path`
   - `sheet_name`
   - `desc_col`
4. Run the notebook.
5. Check `Updated_<month>_Surcharges.xlsx` for output.

  ## Dependencies

- pandas
- PyMuPDF (`pip install pymupdf`)
