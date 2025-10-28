# Payslip Generator Excel Template Setup Instructions

## How to Create Your Excel Template:

1. **Open Excel** and create a new workbook
2. **Copy the CSV data** from `employee_template.csv` into Excel
3. **Format the headers** (make them bold, add background color)
4. **Set column widths** appropriately
5. **Add data validation** for better user experience

## Excel Template Features:

### Column Structure:
- **Employee Name**: Text (required)
- **Employee ID**: Text (required, unique)
- **Hourly Rate**: Number with 2 decimal places (required)
- **Hours Worked**: Number (required)
- **Department**: Text (optional)
- **Position**: Text (optional)
- **Tax Exempt**: Yes/No dropdown (optional)
- **Insurance Plan**: Dropdown: Standard, Premium, Basic (optional)
- **Notes**: Text (optional)

### Recommended Excel Formatting:

1. **Freeze Header Row**: View → Freeze Panes → Freeze Top Row
2. **Data Validation**:
   - Tax Exempt: List with "Yes,No"
   - Insurance Plan: List with "Standard,Premium,Basic"
   - Hourly Rate: Decimal number, minimum 0
   - Hours Worked: Decimal number, minimum 0, maximum 168

3. **Conditional Formatting**:
   - Highlight overtime hours (if Hours Worked > 40)
   - Highlight high hourly rates (if Hourly Rate > 30)

4. **Column Widths**:
   - Employee Name: 150px
   - Employee ID: 100px
   - Hourly Rate: 100px
   - Hours Worked: 100px
   - Department: 120px
   - Position: 150px
   - Tax Exempt: 100px
   - Insurance Plan: 120px
   - Notes: 200px

## Converting CSV to Excel:

1. Open the `employee_template.csv` file
2. In Excel: File → Save As → Excel Workbook (*.xlsx)
3. Apply the formatting above
4. Save as `employee_template.xlsx`

## Usage with GUI:

1. Run `python gui_payslip.py`
2. Click "Browse" and select your Excel file
3. Click "Load Employee Data"
4. Double-click any employee to preview their payslip
5. Click "Generate Payslips" to create all payslips

## Future Excel Integration:

To fully support Excel files (.xlsx), install:
```
pip install openpyxl
```

Then the GUI will automatically detect and read Excel files.